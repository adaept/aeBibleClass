# docm_rpr_signatures.py <flat-OPC or document.xml>: the distinct run-property signatures of the VerseText paragraphs, with counts and an example.
# Part of the RWB docm tools (tools/word-com/README.md). Run in WSL: wsl.exe --exec python3 -I py/docm_rpr_signatures.py <file>.
import re, sys, collections
x = open(sys.argv[1], encoding="utf-8").read()
paras = re.findall(r'<w:p[ >].*?</w:p>', x, flags=re.S)
sig = collections.Counter(); ex = {}
for p in paras:
    m = re.search(r'<w:pStyle w:val="([^"]+)"', p)
    if not m or m.group(1) != "VerseText": continue
    for r in re.finditer(r'<w:r[ >](.*?)</w:r>', p, flags=re.S):
        body = r.group(1)
        rp = re.search(r'<w:rPr>(.*?)</w:rPr>', body, flags=re.S)
        tags = []
        if rp:
            for t in re.finditer(r'<w:([A-Za-z]+)([^>]*?)/?>', rp.group(1)):
                n, a = t.group(1), t.group(2)
                v = re.search(r'w:val="([^"]*)"', a)
                tags.append(n + ("=" + v.group(1) if v and n in ("rStyle","b","bCs","i","iCs","sz","szCs","color","u","lang","vertAlign","highlight","caps","smallCaps","strike") else ""))
        s = "|".join(sorted(tags)) or "(none)"
        sig[s] += 1
        if s not in ex:
            t = re.search(r'<w:t[^>]*>(.*?)</w:t>', body, flags=re.S)
            ex[s] = (t.group(1)[:30] if t else "")
print("verse paragraphs:", sum(1 for p in paras if '<w:pStyle w:val="VerseText"' in p))
for s, n in sig.most_common(40):
    print(n, s, "  e.g.", repr(ex[s]))
