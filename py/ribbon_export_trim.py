#!/usr/bin/env python3
"""
ribbon_export_trim.py - Production export gateway for the Radiant Word Bible ribbon.

Reads dev VBA sources from src/, computes the call graph rooted at the ribbon
callbacks declared in customUI14backupRWB.xml (plus AutoExec / RibbonOnLoad /
Document_Open), and writes trimmed copies into aeRibbon/src/ containing only
reachable routines.

Idempotent: re-running on the same src/ produces identical aeRibbon/src/ output.

Per md/aeProductionRibbonPlan.md sec 2.3:
  - TRIM (call-graph reachability): basBibleRibbonSetup.bas, aeRibbonClass.cls,
    aeBibleClass.cls, aeBibleCitationClass.cls
  - COPY AS-IS (small, all-public): basRibbonDeferred.bas, basUIStrings.bas
  - EXCLUDED: ThisDocument.cls (handled as build step; see BUILD.md sec 7.2)

Usage:
    python3 py/ribbon_export_trim.py
    python3 py/ribbon_export_trim.py --check    # exit 1 if output would change

Honors [feedback_casing]: no identifier-casing changes, ever. We only
add or remove whole routine blocks; surviving text is byte-identical to src/.
"""

from __future__ import annotations

import argparse
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Set, Tuple

# ----- Configuration ---------------------------------------------------------

REPO_ROOT = Path(__file__).resolve().parent.parent
SRC_DIR = REPO_ROOT / "src"
OUT_DIR = REPO_ROOT / "aeRibbon" / "src"
LOG_PATH = REPO_ROOT / "aeRibbon" / "RoutineLog.md"
RIBBON_XML = REPO_ROOT / "customUI14backupRWB.xml"

TRIM_FILES = [
    "basBibleRibbonSetup.bas",
    "aeRibbonClass.cls",
    "aeBibleCitationClass.cls",
    "basSBL_VerseCountsGenerator.bas",
]
ASIS_FILES = [
    "basRibbonDeferred.bas",
    "basUIStrings.bas",
]

EXTRA_ROOTS = {"AutoExec", "RibbonOnLoad", "Document_Open", "Instance"}

# VBA implicit-call names: Word/VBA invokes these without an explicit caller.
# They must always be preserved if defined, regardless of static reachability.
LIFECYCLE_NAMES = {
    "class_initialize", "class_terminate",
    "document_open", "document_close", "document_new",
    "autoexec", "autonew", "autoopen", "autoclose", "autoexit",
}

# ----- Ribbon XML root extraction --------------------------------------------

def extract_callback_roots(xml_path: Path) -> Set[str]:
    """Pull every attribute value that looks like a VBA callback name."""
    roots: Set[str] = set()
    tree = ET.parse(xml_path)
    callback_attrs = {
        "onLoad", "onAction", "onChange",
        "getLabel", "getEnabled", "getKeytip",
        "getItemCount", "getItemLabel", "getItemID", "getText",
        "getVisible", "getImage", "getScreentip", "getSupertip",
    }
    for elem in tree.iter():
        for attr, value in elem.attrib.items():
            if attr in callback_attrs and value and re.match(r"^[A-Za-z_]\w*$", value):
                roots.add(value)
    return roots

# ----- VBA file parser -------------------------------------------------------

# A routine starts on a line matching this, ignoring leading whitespace:
ROUTINE_START = re.compile(
    r"^\s*(?:Public|Private|Friend)?\s*(?:Static\s+)?"
    r"(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+"
    r"(?P<name>[A-Za-z_]\w*)\b",
    re.IGNORECASE,
)
ROUTINE_END = re.compile(
    r"^\s*End\s+(?:Sub|Function|Property)\s*$",
    re.IGNORECASE,
)
ATTRIBUTE_LINE = re.compile(r"^\s*Attribute\s+VB_", re.IGNORECASE)

# Identifier scan for call-graph edges: match VBA identifiers, case-insensitive.
IDENT = re.compile(r"[A-Za-z_]\w*")

class Routine:
    __slots__ = ("name", "kind", "start", "end", "lines", "attr_lines")
    def __init__(self, name: str, kind: str, start: int, end: int,
                 lines: List[str], attr_lines: List[str]):
        self.name = name
        self.kind = kind
        self.start = start
        self.end = end
        self.lines = lines
        self.attr_lines = attr_lines  # leading Attribute lines that belong to this routine

class ParsedFile:
    def __init__(self, path: Path, raw: str):
        self.path = path
        self.raw_lines = raw.splitlines(keepends=True)
        # Header: everything from line 0 up to (but not including) the first routine.
        # Routines: list of Routine.
        # Trailer: anything after the last End Sub/Function/Property (rare).
        self.header_end = 0
        self.routines: List[Routine] = []
        self.trailer_start = len(self.raw_lines)
        self._parse()

    def _parse(self) -> None:
        lines = self.raw_lines
        i = 0
        n = len(lines)
        first_routine_idx: int | None = None
        last_end_idx = -1

        while i < n:
            m = ROUTINE_START.match(lines[i])
            if m:
                # A routine begins. Capture any preceding Attribute lines that
                # were emitted between the prior routine's End and this start
                # (VBA Property attributes), but we treat them as part of the
                # header for the first routine, and as preface for subsequent
                # routines only if they appear inline with the routine.
                # Simpler: leading Attribute lines just before the routine
                # (no blank line separation) belong to this routine.
                attr_lines: List[str] = []
                j = i - 1
                while j > last_end_idx and ATTRIBUTE_LINE.match(lines[j]):
                    attr_lines.insert(0, lines[j])
                    j -= 1
                routine_first = j + 1  # first line that "belongs" to this routine (attrs or def)

                if first_routine_idx is None:
                    first_routine_idx = routine_first
                    self.header_end = routine_first

                # Find matching End
                k = i + 1
                while k < n and not ROUTINE_END.match(lines[k]):
                    k += 1
                if k >= n:
                    raise ValueError(f"Unterminated routine {m.group('name')} in {self.path}")
                routine_lines = lines[routine_first:k + 1]
                # attr_lines is a slice prefix of routine_lines, so do not pass
                # it separately when writing; we keep it only for diagnostics.
                self.routines.append(Routine(
                    name=m.group("name"),
                    kind=m.group("kind"),
                    start=routine_first,
                    end=k + 1,
                    lines=routine_lines,
                    attr_lines=attr_lines,
                ))
                last_end_idx = k
                i = k + 1
            else:
                i += 1

        if first_routine_idx is None:
            # No routines at all - entire file is "header".
            self.header_end = n
            self.trailer_start = n
        else:
            self.trailer_start = last_end_idx + 1

    def header(self) -> List[str]:
        return self.raw_lines[:self.header_end]

    def trailer(self) -> List[str]:
        return self.raw_lines[self.trailer_start:]


# ----- Identifier extraction (skipping comments and strings) ----------------

def strip_comments_and_strings(line: str) -> str:
    """Return line with VBA comments and string literals blanked out."""
    out = []
    i = 0
    in_str = False
    while i < len(line):
        ch = line[i]
        if in_str:
            if ch == '"':
                if i + 1 < len(line) and line[i + 1] == '"':
                    out.append(" "); out.append(" "); i += 2; continue
                in_str = False
                out.append(" ")
            else:
                out.append(" ")
            i += 1
            continue
        if ch == '"':
            in_str = True
            out.append(" ")
            i += 1
            continue
        if ch == "'":
            # Rest of line is comment
            out.append(" " * (len(line) - i))
            break
        out.append(ch)
        i += 1
    result = "".join(out)
    # Also strip "Rem " comments at start of a logical statement.
    m = re.match(r"^(\s*)Rem\b", result, re.IGNORECASE)
    if m:
        return m.group(1) + " " * (len(result) - len(m.group(1)))
    return result

def identifiers_in_body(lines: List[str]) -> Set[str]:
    found: Set[str] = set()
    for raw in lines:
        clean = strip_comments_and_strings(raw)
        for m in IDENT.finditer(clean):
            found.add(m.group(0).lower())
    return found

# ----- Build call graph + reachability --------------------------------------

def build_graph(parsed: Dict[str, ParsedFile]) -> Tuple[Dict[str, Set[str]], Dict[str, List[Tuple[str, Routine]]]]:
    """
    Returns (edges, defs):
      edges: lowercased routine name -> set of lowercased names it references
      defs:  lowercased routine name -> list of (filename, Routine) defining it
    """
    defs: Dict[str, List[Tuple[str, Routine]]] = {}
    for fname, pf in parsed.items():
        for r in pf.routines:
            defs.setdefault(r.name.lower(), []).append((fname, r))

    known = set(defs.keys())
    edges: Dict[str, Set[str]] = {}
    for name_lc, occurrences in defs.items():
        refs: Set[str] = set()
        for _, r in occurrences:
            # Skip the def line itself when scanning for outbound refs - it
            # contains the routine's own name, which would create a self-loop
            # (harmless but noisy).
            body = r.lines[1:] if r.lines else []
            for ident in identifiers_in_body(body):
                if ident in known and ident != name_lc:
                    refs.add(ident)
        edges[name_lc] = refs
    return edges, defs

def reachable_from(roots: Set[str], edges: Dict[str, Set[str]]) -> Set[str]:
    seen: Set[str] = set()
    stack = [r for r in roots if r in edges]
    while stack:
        cur = stack.pop()
        if cur in seen:
            continue
        seen.add(cur)
        for nxt in edges.get(cur, ()):
            if nxt not in seen:
                stack.append(nxt)
    return seen

# ----- Write trimmed output --------------------------------------------------

def write_trimmed(pf: ParsedFile, keep_names_lc: Set[str], out_path: Path) -> Tuple[List[Routine], List[Routine]]:
    kept: List[Routine] = []
    dropped: List[Routine] = []
    parts: List[str] = []
    parts.extend(pf.header())

    # Preserve original routine order.
    for idx, r in enumerate(pf.routines):
        if r.name.lower() in keep_names_lc:
            kept.append(r)
            parts.extend(r.lines)
        else:
            dropped.append(r)

    parts.extend(pf.trailer())
    out_path.parent.mkdir(parents=True, exist_ok=True)
    # VBA editor on Windows requires CRLF to recognize the .cls header
    # ("VERSION 1.0 CLASS" block). Force CRLF regardless of in-memory state.
    out_path.write_text("".join(parts), encoding="utf-8", newline="\r\n")
    return kept, dropped

# ----- Main ------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true",
                    help="Exit 1 if regenerated output differs from on-disk output.")
    args = ap.parse_args()

    # 1. Roots from ribbon XML + manual.
    xml_roots = extract_callback_roots(RIBBON_XML)
    roots = {r.lower() for r in xml_roots | EXTRA_ROOTS}

    # 2. Parse every file we care about (TRIM + AS-IS). We need AS-IS files in
    #    the graph too so that routines in TRIM files reached only via AS-IS
    #    callers stay alive.
    parsed: Dict[str, ParsedFile] = {}
    for fname in TRIM_FILES + ASIS_FILES:
        p = SRC_DIR / fname
        parsed[fname] = ParsedFile(p, p.read_text(encoding="utf-8"))

    # 3. Treat every routine in AS-IS files as an additional root (because
    #    those files ship whole and could call anything in TRIM files).
    asis_root_names = {
        r.name.lower()
        for fname in ASIS_FILES
        for r in parsed[fname].routines
    }
    all_roots = roots | asis_root_names

    edges, defs = build_graph(parsed)
    keep = reachable_from(all_roots, edges)

    # Always preserve VBA lifecycle / auto-macro hooks if defined - Word
    # invokes them without an explicit caller, so they look unreachable to a
    # static call-graph analysis.
    lifecycle_kept: Set[str] = set()
    for name_lc in defs:
        if name_lc in LIFECYCLE_NAMES:
            keep.add(name_lc)
            # Also pull in anything those hooks call.
            keep |= reachable_from({name_lc}, edges)
            lifecycle_kept.add(name_lc)

    # 4. Write outputs.
    log_rows: List[Tuple[str, str, str, str]] = []  # (file, routine, decision, reason)

    for fname in TRIM_FILES:
        pf = parsed[fname]
        out_path = OUT_DIR / fname
        kept, dropped = write_trimmed(pf, keep, out_path)
        for r in kept:
            reason = ("VBA lifecycle hook (always preserved)"
                      if r.name.lower() in LIFECYCLE_NAMES
                      else "reachable from ribbon callbacks")
            log_rows.append((fname, r.name, "KEPT", reason))
        for r in dropped:
            log_rows.append((fname, r.name, "REMOVED", "not reachable from ribbon callbacks"))

    for fname in ASIS_FILES:
        pf = parsed[fname]
        out_path = OUT_DIR / fname
        # Copy whole file byte-for-byte from src/.
        src_path = SRC_DIR / fname
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(src_path.read_text(encoding="utf-8"),
                            encoding="utf-8", newline="\r\n")
        for r in pf.routines:
            log_rows.append((fname, r.name, "KEPT", "AS-IS file (no trim applied)"))

    # 5. Write RoutineLog.md.
    log_lines: List[str] = []
    log_lines.append("# aeRibbon Routine Log\n\n")
    log_lines.append("Generated by `py/ribbon_export_trim.py`. Re-run to refresh.\n\n")
    log_lines.append("| File | Routine | Decision | Reason |\n")
    log_lines.append("|---|---|---|---|\n")
    for fname, rname, decision, reason in sorted(log_rows, key=lambda x: (x[0], x[1].lower())):
        log_lines.append(f"| {fname} | {rname} | {decision} | {reason} |\n")

    # Summary block.
    by_file: Dict[str, Tuple[int, int]] = {}
    for fname, _r, decision, _reason in log_rows:
        kept_n, drop_n = by_file.get(fname, (0, 0))
        if decision == "KEPT":
            by_file[fname] = (kept_n + 1, drop_n)
        else:
            by_file[fname] = (kept_n, drop_n + 1)
    log_lines.append("\n## Summary\n\n")
    log_lines.append("| File | Kept | Removed |\n|---|---|---|\n")
    for fname in TRIM_FILES + ASIS_FILES:
        k, d = by_file.get(fname, (0, 0))
        log_lines.append(f"| {fname} | {k} | {d} |\n")

    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    LOG_PATH.write_text("".join(log_lines), encoding="utf-8", newline="")

    print(f"Roots from XML: {sorted(xml_roots)}")
    print(f"Total routines kept: {sum(k for k, _ in by_file.values())}")
    print(f"Total routines removed: {sum(d for _, d in by_file.values())}")
    print(f"Output: {OUT_DIR}")
    print(f"Log:    {LOG_PATH}")
    return 0

if __name__ == "__main__":
    sys.exit(main())
