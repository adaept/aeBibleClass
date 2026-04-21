"""
check_i18n.py — i18n completeness scan for aeBibleClass VBA source files.

Scans all src/*.bas and src/*.cls files for string literals that look like
user-visible UI text and are NOT references to basUIStrings constants.

Violations = inline string literals that should be constants.
False positives = technical strings (control IDs, format strings, debug output, etc.)

Run:
    python3 py/check_i18n.py          # from project root
    python3 py/check_i18n.py --strict # treat any match as exit code 1
"""

import re
import sys
import pathlib

# ---------------------------------------------------------------------------
# Known-clean patterns: strings that are intentionally inline and not i18n
# violations. Add to this list when the scanner produces known false positives.
# ---------------------------------------------------------------------------
KNOWN_CLEAN = {
    # Debug / logging output — not user-visible
    r'"basBibleRibbonSetup: AutoExec at "',
    r'">> RibbonOnLoad at "',
    r'">> OnPrevButtonClick at "',
    r'">> OnNextButtonClick at "',
    r'"hh:nn:ss"',
    r'"hh:mm:ss"',
    # Error handler boilerplate (module/procedure names)
    r'"Erl="',
    r'" Error "',
    r'" \(.*\) in procedure "',
    # VBA attribute lines (top of each file)
    r'Attribute VB_Name',
    # Format strings for internal use
    r'"\{[0-9]\}"',         # placeholder tokens used inside FormatMsg
    r'"00:00:0[0-9]"',      # OnTime delay literals
    # Already-constant status bar strings assigned via SB_* or FormatMsg
    # (These appear as the RHS of Application.StatusBar = ... and are caught
    #  only if someone accidentally inlines the text. The scanner catches the
    #  literal text, not the constant reference.)
}

# Regex: a double-quoted VBA string literal that is at least 3 characters,
# contains at least one space OR starts with an uppercase letter followed by
# a lowercase letter (heuristic for human-readable text).
STRING_LITERAL = re.compile(r'"([^"]{3,})"')

# Strings that are clearly technical identifiers, not UI text.
# Checked against the captured group (without quotes).
TECHNICAL_PATTERNS = [
    re.compile(r'^[a-z][a-zA-Z0-9]+$'),          # camelCase identifier
    re.compile(r'^[A-Z][a-zA-Z0-9]+$'),          # PascalCase identifier
    re.compile(r'^[A-Z_]{2,}$'),                  # ALL_CAPS constant ref (shouldn't appear as literal)
    re.compile(r'^\d'),                           # starts with digit
    re.compile(r'^[,.\[\]<>{}()+\-*/=!@#$%^&*]+$'),  # punctuation-only
    re.compile(r'^Erl='),                         # error handler boilerplate
    re.compile(r'^\d{2}:\d{2}:\d{2}$'),          # time format string
    re.compile(r'^\d{2}:\d{2}:\d{1,2}$'),        # OnTime delay
    re.compile(r'^[A-Za-z0-9_]+\.[A-Za-z0-9_]+\.[A-Za-z0-9_]+$'),  # dotted name (module.sub)
    re.compile(r'^>>\s'),                         # debug prefix
    re.compile(r'in procedure '),                 # error handler boilerplate
    re.compile(r' of (Class|Module) '),           # error handler boilerplate
    re.compile(r'Error \d'),                      # error handler boilerplate
    re.compile(r'basRibbonDeferred\.'),           # OnTime target suffix (technical)
    re.compile(r'^Heading [1-9]$'),               # Word built-in style name (not i18n)
    re.compile(r'^yyyy-'),                        # date/time format string
    re.compile(r'\\rpt\\'),                       # file path fragment
    re.compile(r'\[0-9\]'),                       # Like pattern / regex fragment
    re.compile(r'^sessionID,'),                   # CSV header (log output, not UI)
    re.compile(r',H1\['),                         # CSV data fragment (log output)
    re.compile(r'\s&\s'),                         # VBA concatenation fragment (" & var & ")
    re.compile(r'^\s*&\s'),                       # starts with & (fragment after closing quote)
    re.compile(r'marker$'),                       # Word style name suffix (not i18n)
]

# Lines containing these VBA keywords are structural — not UI strings.
LINE_SKIP_PATTERNS = [
    re.compile(r'^\s*Attribute\s+VB_'),
    re.compile(r'^\s*\''),                        # comment line
    re.compile(r'Debug\.Print'),
    re.compile(r'MsgBox\s+"Erl='),               # standard error MsgBox
    re.compile(r'^\s*Public\s+Const\s+\w+\s+As\s+String\s*='),  # constant definition line (basUIStrings)
    re.compile(r'^\s*Private\s+Const\s+\w+\s+As\s+String\s*='),
]

# Known basUIStrings constants (prefix check — any string that starts with one
# of these prefixes and is used in context is OK; we flag literals only).
UISTRINGS_PREFIXES = ("KT_", "LBL_", "SB_", "CTRL_")


def is_technical(text: str) -> bool:
    for pat in TECHNICAL_PATTERNS:
        if pat.search(text):
            return True
    return False


def should_skip_line(line: str) -> bool:
    for pat in LINE_SKIP_PATTERNS:
        if pat.search(line):
            return True
    return False


def scan_file(path: pathlib.Path) -> list[tuple[int, str, str]]:
    """Return list of (line_no, literal, line) for suspected violations."""
    violations = []
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except Exception as e:
        print(f"  ERROR reading {path}: {e}")
        return violations

    for lineno, line in enumerate(text.splitlines(), 1):
        if should_skip_line(line):
            continue
        for m in STRING_LITERAL.finditer(line):
            literal = m.group(1)
            if is_technical(literal):
                continue
            violations.append((lineno, literal, line.rstrip()))
    return violations


def main():
    strict = "--strict" in sys.argv
    root = pathlib.Path(__file__).parent.parent
    src = root / "src"

    if not src.exists():
        print(f"ERROR: src/ directory not found at {src}")
        sys.exit(1)

    files = sorted(src.glob("*.bas")) + sorted(src.glob("*.cls"))
    if not files:
        print("No .bas or .cls files found in src/")
        sys.exit(1)

    total_violations = 0
    for path in files:
        hits = scan_file(path)
        if hits:
            print(f"\n{path.name}  ({len(hits)} potential violation(s))")
            for lineno, literal, line in hits:
                print(f"  {lineno:4d}  \"{literal}\"")
                print(f"        {line[:120]}")
            total_violations += len(hits)

    print()
    if total_violations == 0:
        print("OK — no inline UI string literals detected.")
        sys.exit(0)
    else:
        print(f"REVIEW — {total_violations} potential inline literal(s) flagged.")
        print("Check each: add to basUIStrings.bas if user-visible, or add to KNOWN_CLEAN if intentional.")
        if strict:
            sys.exit(1)
        else:
            sys.exit(0)


if __name__ == "__main__":
    main()
