# docm_text_diff.py <old document.xml> <new document.xml>: the paragraphs whose text differs, as blocks (difflib). Part of the RWB docm tools.
import re, difflib, sys
def paras(path):
    x = open(path, encoding="utf-8").read()
    out = []
    for p in re.findall(r'<w:p[ >].*?</w:p>', x, flags=re.S):
        out.append("".join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', p, flags=re.S)))
    return out
a = paras(sys.argv[1]); b = paras(sys.argv[2])
print("paragraphs", len(a), len(b))
sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
n = 0
for tag, i1, i2, j1, j2 in sm.get_opcodes():
    if tag == "equal": continue
    n += 1
    if n > 12: continue
    print(f"--- {tag} old[{i1}:{i2}] new[{j1}:{j2}]")
    for s in a[i1:i2][:4]: print("  OLD:", repr(s[:200]), "(len", len(s), ")")
    for s in b[j1:j2][:4]: print("  NEW:", repr(s[:200]), "(len", len(s), ")")
print("changed blocks:", n)
