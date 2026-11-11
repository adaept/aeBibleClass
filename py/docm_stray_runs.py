# docm_stray_runs.py <flat-OPC or document.xml>: every VerseText run whose properties name anything except rStyle, lang, rFonts, noProof (Test 98 rule, decision A of FU-32;
# the style definition at the end needs a flat-OPC package that includes styles.xml), with book chapter:verse. Part of the RWB docm tools (tools/word-com/README.md).
import re, sys
x = open(sys.argv[1], encoding="utf-8").read()
paras = re.findall(r'<w:p[ >].*?</w:p>', x, flags=re.S)
def txt(s): return "".join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', s, flags=re.S))
book = ""
for p in paras:
    m = re.search(r'<w:pStyle w:val="([^"]+)"', p)
    st = m.group(1) if m else ""
    if st == "Heading1": book = txt(p)[:30]
    if st != "VerseText": continue
    ch = vs = ""
    for r in re.finditer(r'<w:r[ >](.*?)</w:r>', p, flags=re.S):
        b = r.group(1)
        rs = re.search(r'<w:rStyle w:val="([^"]+)"', b)
        t = txt(b)
        if rs and rs.group(1) == "ChapterVersemarker": ch = re.sub(r'\D', '', t) or ch
        if rs and rs.group(1) == "Versemarker": vs = re.sub(r'\D', '', t) or vs
        rp = re.search(r'<w:rPr>(.*?)</w:rPr>', b, flags=re.S)
        if not rp: continue
        names = re.findall(r'<w:([A-Za-z]+)[ />]', rp.group(1))
        bad = [n for n in names if n not in ("rStyle", "lang", "rFonts", "noProof")]
        if bad:
            print(book, ch + ":" + vs, "| rPr:", re.sub(r'\s+', ' ', rp.group(1))[:120], "| text:", repr(t[:40]))
print("VerseText style definition:")
m = re.search(r'<w:style [^>]*w:styleId="VerseText".*?</w:style>', x, flags=re.S)
print(re.sub(r'\s+', ' ', m.group(0))[:600] if m else "not found")
