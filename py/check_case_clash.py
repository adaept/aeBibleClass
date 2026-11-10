"""List identifiers declared in one VBA module that clash in case with a different spelling used anywhere else in src/.

The VBE re-cases every use of an identifier to the spelling of the most recent declaration, project-wide, so a local named
`saved` turns Word's `.Saved` into `.saved` in every other module (found 2026-11-10 in aeRibbonClass.cls).
Usage: python3 -I case_clash.py <module.bas> <src folder>
"""
import re
import sys
from pathlib import Path

module = Path(sys.argv[1])
src = Path(sys.argv[2])
text = module.read_text(encoding="utf-8", errors="replace")

declared = set()
for line in text.splitlines():
    code = line.split("'")[0]
    # Dim / Const / ReDim names, and Sub / Function names, and parameters (ByVal x As ..., ByRef x As ...)
    for m in re.finditer(r'\b(?:Dim|Const|ReDim|Static)\s+([A-Za-z_][A-Za-z0-9_]*)', code, flags=re.I):
        declared.add(m.group(1))
    for m in re.finditer(r'\b(?:Private|Public)?\s*(?:Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)', code, flags=re.I):
        declared.add(m.group(1))
    for m in re.finditer(r'\b(?:ByVal|ByRef)\s+([A-Za-z_][A-Za-z0-9_]*)', code, flags=re.I):
        declared.add(m.group(1))
    # a second Dim on the same line: Dim a As Long, b As Long
    for m in re.finditer(r',\s*([A-Za-z_][A-Za-z0-9_]*)\s+As\b', code, flags=re.I):
        declared.add(m.group(1))

others = {}
for p in sorted(src.rglob("*")):
    if p.suffix.lower() not in {".bas", ".cls", ".frm"} or p.resolve() == module.resolve():
        continue
    t = p.read_text(encoding="utf-8", errors="replace")
    for line in t.splitlines():
        code = line.split("'")[0]
        for tok in re.findall(r'(?<![A-Za-z0-9_])[A-Za-z_][A-Za-z0-9_]*', code):
            others.setdefault(tok.lower(), {}).setdefault(tok, set()).add(p.name)

own = {}
for tok in re.findall(r'(?<![A-Za-z0-9_])[A-Za-z_][A-Za-z0-9_]*', "\n".join(l.split("'")[0] for l in text.splitlines())):
    own.setdefault(tok.lower(), set()).add(tok)

clashes = 0
for name in sorted(declared, key=str.lower):
    spellings = others.get(name.lower(), {})
    differing = {s: sorted(files) for s, files in spellings.items() if s != name}
    own_other = {s for s in own.get(name.lower(), set()) if s != name}
    if differing or own_other:
        clashes += 1
        print(f"CLASH  declared as '{name}' here; elsewhere: " +
              "; ".join(f"{s} ({', '.join(f[:3])}{'...' if len(f) > 3 else ''})" for s, f in differing.items()) +
              (f"; inside this module also: {sorted(own_other)}" if own_other else ""))
print(f"{len(declared)} declared names checked, {clashes} with a case clash")
