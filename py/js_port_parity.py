#!/usr/bin/env python3
"""Parity check: aeBibleClass.cls VBA Test/Case numbers vs. aeBibleAddin JS port status.

Built per adaept5tudio-word-addin-conversion-plan.md Section 15.11a's proposed
mechanism: enumerate the VBA side's actual Case-N set from aeBibleClass.cls's four
dispatch switches, enumerate what aeBibleAddin has ported/tested/live-wired for each
Case number (found via "Case N" / "Test N" annotations in source comments, linked to
the nearest function definition below them), and report one row per VBA Case.

This is a review tool, not an authority: it surfaces candidates by pattern-matching
comments, which can over- or under-attribute in edge cases. Read the per-case
evidence (file:line, function name) before trusting a status, especially for
MENTIONED and NOT PORTED rows.

Usage: python3 py/js_port_parity.py [--case N] [--js-root PATH]
"""
import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
VBA_FILE = REPO_ROOT / "src" / "aeBibleClass.cls"
DEFAULT_JS_ROOT = REPO_ROOT.parent / "aeBibleAddin"

SWITCH_NAMES = ["GetTestDescription", "GetPassFail", "RunTest", "OutputTestReport"]

HEADER_RE = re.compile(r'^\s*(Public\s+|Private\s+)?(Function|Sub|Property)\s+', re.IGNORECASE)
CASE_LINE_RE = re.compile(r'^\s*Case\s+(.+?)\s*$', re.IGNORECASE)
FUNC_DEF_RE = re.compile(r'^(export\s+)?(async\s+)?function\s+(\w+)\s*\(')

# Matches "Case N", "Cases N", "Test N", "Tests N" followed by a digit/separator run
# (comma, slash, hyphen/en-dash, or the word "to"/"To" for VBA-style ranges).
NUM_LIST_RE = re.compile(
    r'\b(?:Case|Cases|Test|Tests)\s+(\d+(?:\s*(?:,|/|-|\u2013|[Tt]o)\s*\d+)*)'
)


def fail(msg):
    print(f"ERROR: {msg}", file=sys.stderr)
    sys.exit(1)


# ---------------------------------------------------------------------------
# VBA side
# ---------------------------------------------------------------------------

def extract_vba_block(lines, func_name):
    start = None
    pat = re.compile(rf'^\s*(Public\s+|Private\s+)?(Function|Sub)\s+{re.escape(func_name)}\b', re.IGNORECASE)
    for i, line in enumerate(lines):
        if pat.match(line):
            start = i
            break
    if start is None:
        fail(f"could not find VBA routine {func_name} in {VBA_FILE}")
    end = len(lines)
    for j in range(start + 1, len(lines)):
        if HEADER_RE.match(lines[j]):
            end = j
            break
    return lines[start:end]


def parse_vba_case_numbers(block_lines):
    nums = set()
    for line in block_lines:
        m = CASE_LINE_RE.match(line)
        if not m:
            continue
        spec = m.group(1).strip()
        if spec.lower().startswith("else"):
            continue
        for part in re.split(r'\s*,\s*', spec):
            to_m = re.match(r'^(\d+)\s+To\s+(\d+)$', part, re.IGNORECASE)
            if to_m:
                nums.update(range(int(to_m.group(1)), int(to_m.group(2)) + 1))
            elif part.isdigit():
                nums.add(int(part))
    return nums


def load_vba_cases():
    text = VBA_FILE.read_text(encoding="utf-8", errors="replace")
    lines = text.splitlines()
    m = re.search(r'Private\s+Const\s+MaxTests\s*=\s*(\d+)', text)
    if not m:
        fail("could not find 'Private Const MaxTests = N' in aeBibleClass.cls")
    max_tests = int(m.group(1))

    switch_cases = {}
    for name in SWITCH_NAMES:
        block = extract_vba_block(lines, name)
        switch_cases[name] = parse_vba_case_numbers(block)

    expected = set(range(1, max_tests + 1))
    consistency_issues = []
    for name, nums in switch_cases.items():
        missing = sorted(expected - nums)
        extra = sorted(nums - expected)
        if missing:
            consistency_issues.append(f"{name}: missing Case(s) {missing}")
        if extra:
            consistency_issues.append(f"{name}: out-of-range Case(s) {extra}")

    return max_tests, expected, consistency_issues


# ---------------------------------------------------------------------------
# JS side
# ---------------------------------------------------------------------------

def expand_prose_numbers(spec):
    tokens = re.split(r'\s*(?:,|/|[Tt]o)\s*', spec)
    nums = set()
    for tok in tokens:
        tok = tok.strip()
        if not tok:
            continue
        m = re.match(r'^(\d+)\s*[-\u2013]\s*(\d+)$', tok)
        if m:
            nums.update(range(int(m.group(1)), int(m.group(2)) + 1))
        elif tok.isdigit():
            nums.add(int(tok))
    return nums


def find_case_refs(text):
    found = set()
    for m in NUM_LIST_RE.finditer(text):
        found |= expand_prose_numbers(m.group(1))
    return found


def iter_ts_files(js_root, pattern="*.ts"):
    for p in js_root.rglob(pattern):
        parts = p.parts
        if "node_modules" in parts or "dist" in parts or ".git" in parts:
            continue
        yield p


def is_comment_or_blank(line):
    s = line.strip()
    return (
        s == ""
        or s.startswith("//")
        or s.startswith("*")
        or s.startswith("/**")
        or s.startswith("/*")
        or s.endswith("*/")
    )


def extract_function_case_map(path, max_lines=80, max_noncomment_gap=3):
    """Map function name -> set of Case/Test numbers mentioned in the comment block
    immediately preceding its definition. Walks upward from the function, tolerating
    up to `max_noncomment_gap` consecutive non-comment lines in a row (e.g. a lone
    `const X = ...` sitting between a doc-comment and the function it describes)
    before giving up -- this stops the walk from bleeding into an unrelated, distant
    file-header comment while still crossing small intervening declarations. Also
    truncated at the previous function's own definition line and at `max_lines`."""
    try:
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError:
        return {}
    func_positions = []
    for i, line in enumerate(lines):
        m = FUNC_DEF_RE.match(line.strip())
        if m:
            func_positions.append((i, m.group(3)))
    mapping = {}
    for idx, (i, func_name) in enumerate(func_positions):
        prev_end = func_positions[idx - 1][0] + 1 if idx > 0 else 0
        collected = []
        noncomment_streak = 0
        j = i - 1
        while j >= prev_end and (i - j) <= max_lines:
            line = lines[j]
            if is_comment_or_blank(line):
                noncomment_streak = 0
            else:
                noncomment_streak += 1
                if noncomment_streak > max_noncomment_gap:
                    break
            collected.append(line)
            j -= 1
        chunk = "\n".join(reversed(collected))
        nums = find_case_refs(chunk)
        if nums:
            mapping.setdefault(func_name, set()).update(nums)
    return mapping


def get_import_line_ranges(lines):
    ranges = []
    i = 0
    while i < len(lines):
        if re.match(r'^\s*import\s*\{', lines[i]):
            start = i
            while i < len(lines) and "from" not in lines[i]:
                i += 1
            ranges.append((start, i))
        i += 1
    return ranges


def is_in_ranges(i, ranges):
    return any(a <= i <= b for a, b in ranges)


def is_called_in(file_path, func_name):
    try:
        lines = file_path.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError:
        return False
    import_ranges = get_import_line_ranges(lines)
    call_re = re.compile(rf'\b{re.escape(func_name)}\s*\(')
    for i, line in enumerate(lines):
        if is_in_ranges(i, import_ranges):
            continue
        if call_re.search(line):
            return True
    return False


def count_occurrences(text, func_name):
    return len(re.findall(rf'\b{re.escape(func_name)}\s*\(', text))


def load_js_status(js_root):
    """Returns case_number -> dict(status, evidence=[(file, func_name), ...])."""
    all_ts = list(iter_ts_files(js_root))
    source_files = [p for p in all_ts if not p.name.endswith(".test.ts")]
    test_files = [p for p in all_ts if p.name.endswith(".test.ts")]
    taskpane_candidates = [p for p in source_files if p.name == "taskpane.ts"]

    test_texts = {}
    for p in test_files:
        try:
            test_texts[p] = p.read_text(encoding="utf-8", errors="replace")
        except OSError:
            pass

    case_to_funcs = {}  # case_number -> set of (file, func_name)
    for p in source_files:
        fmap = extract_function_case_map(p)
        for func_name, nums in fmap.items():
            for n in nums:
                case_to_funcs.setdefault(n, set()).add((p, func_name))

    # Also catch bare mentions in prose with no associated function (e.g. docstring
    # notes, out-of-scope callouts) so they still surface as "MENTIONED".
    bare_mentions = {}  # case_number -> set of files
    for p in source_files:
        try:
            text = p.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
        for n in find_case_refs(text):
            bare_mentions.setdefault(n, set()).add(p)

    result = {}
    all_case_nums = set(case_to_funcs) | set(bare_mentions)
    for n in all_case_nums:
        funcs = case_to_funcs.get(n, set())
        evidence = []
        status = "MENTIONED"
        live = False
        tested = False
        for (file_path, func_name) in funcs:
            is_live = any(is_called_in(tp, func_name) for tp in taskpane_candidates)
            is_tested = any(
                count_occurrences(txt, func_name) >= 2 for txt in test_texts.values()
            )
            live = live or is_live
            tested = tested or is_tested
            evidence.append((str(file_path.relative_to(js_root)), func_name, is_live, is_tested))
        if funcs:
            status = "LIVE-WIRED" if live else ("TESTED-ONLY" if tested else "PORTED-NOT-WIRED")
        else:
            status = "MENTIONED"
            evidence = [(str(f.relative_to(js_root)), "(no associated function found)", False, False)
                        for f in bare_mentions.get(n, set())]
        result[n] = {"status": status, "evidence": evidence}
    return result


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

def build_report(args, max_tests, expected, consistency_issues, js_status):
    lines = []
    out = lines.append
    out(f"Generated: {datetime.now().isoformat(timespec='seconds')}")
    out(f"Command: python3 py/js_port_parity.py{' --case ' + str(args.case) if args.case else ''}")
    out("This file is overwritten on every run -- diff it in git to see what changed.")
    out("")
    out(f"VBA side: aeBibleClass.cls MaxTests = {max_tests} (Cases 1..{max_tests})")
    if consistency_issues:
        out("VBA dispatch-switch consistency issues found:")
        for issue in consistency_issues:
            out(f"  - {issue}")
    else:
        out("VBA dispatch-switch consistency: OK (all four switches cover 1..MaxTests identically)")
    out("")

    cases_to_show = [args.case] if args.case else sorted(expected)

    counts = {"NOT PORTED": 0, "MENTIONED": 0, "PORTED-NOT-WIRED": 0, "TESTED-ONLY": 0, "LIVE-WIRED": 0}
    header = f"{'Case':>4}  {'Status':<17}  Evidence"
    out(header)
    out("-" * len(header))
    for n in cases_to_show:
        info = js_status.get(n)
        status = info["status"] if info else "NOT PORTED"
        counts[status] += 1
        evidence = info["evidence"] if info else []
        if not evidence:
            out(f"{n:>4}  {status:<17}  (no mention found anywhere in {args.js_root.name})")
        else:
            first = True
            for (fpath, fname, is_live, is_tested) in evidence:
                tag = "LIVE" if is_live else ("TESTED" if is_tested else "-")
                prefix = f"{n:>4}  {status:<17}" if first else f"{'':>4}  {'':<17}"
                out(f"{prefix}  {fpath} :: {fname} [{tag}]")
                first = False

    if not args.case:
        out("")
        out("Summary:")
        for status in ["NOT PORTED", "MENTIONED", "PORTED-NOT-WIRED", "TESTED-ONLY", "LIVE-WIRED"]:
            out(f"  {status:<17} {counts[status]}")
        not_ported = [n for n in cases_to_show if (js_status.get(n) or {}).get("status", "NOT PORTED") in ("NOT PORTED", "MENTIONED")]
        out("")
        out(f"Candidates still needing JS-port work (NOT PORTED or MENTIONED-only): {not_ported}")
        out("")
        out("Evidence is a best-effort heuristic (comment-proximity matching) -- trust NOT")
        out("PORTED as unambiguous (nothing was found at all); spot-check LIVE/TESTED tags")
        out("and function-name attributions before relying on them, especially for older")
        out("Cases where multiple functions sit close together in the source.")

    return "\n".join(lines) + "\n"


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--case", type=int, help="show detail for a single Case number only")
    ap.add_argument("--js-root", type=Path, default=DEFAULT_JS_ROOT)
    ap.add_argument(
        "--output", type=Path, default=None,
        help=f"write report to this file (default: <js-root>/rpt/js_port_parity_report.txt). "
             "Pass '-' to skip writing a file."
    )
    args = ap.parse_args()

    max_tests, expected, consistency_issues = load_vba_cases()
    js_status = load_js_status(args.js_root)
    report = build_report(args, max_tests, expected, consistency_issues, js_status)

    print(report, end="")

    output_path = args.output
    if output_path is None:
        output_path = args.js_root / "rpt" / "js_port_parity_report.txt"
    if str(output_path) != "-":
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(report, encoding="utf-8")
        print(f"\nReport written to {output_path}")


if __name__ == "__main__":
    main()
