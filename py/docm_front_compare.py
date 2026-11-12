# docm_front_compare.py <docm A> <docm B>: compares the first 18 paragraphs of two SAVED docm files: style, wording, italic spans (RWBFrontItalic), justification,
# whether the paragraph ends a section, and any other direct paragraph property. Used to prove that the insert routine reproduces the owner's hand edits
# (RT-22a, 2026-11-11: 18 paragraphs, 0 differences). Reads the files' XML only (no Word). Part of the RWB docm tools.
# Run: wsl.exe --exec python3 -I py/docm_front_compare.py "<A.docm>" "<B.docm>"
import html
import re
import sys
import zipfile

PARAGRAPHS = 18


def load(path, count=PARAGRAPHS):
    x = zipfile.ZipFile(path).read("word/document.xml").decode("utf8")
    body = x[x.index("<w:body>"):]
    out = []
    for p in re.findall(r"<w:p[ >].*?</w:p>", body[:300000], re.S)[:count]:
        style = re.search(r'<w:pStyle w:val="([^"]*)"', p)
        ppr = re.search(r"<w:pPr>.*?</w:pPr>", p, re.S)
        inner = re.sub(r"<w:sectPr.*?</w:sectPr>", "", ppr.group(0), flags=re.S) if ppr else ""
        inner = re.sub(r"<w:pStyle[^>]*/>", "", inner)
        jc = re.search(r'<w:jc w:val="([^"]*)"', inner)
        text, italic = "", []
        for r in re.findall(r"<w:r[ >].*?</w:r>", p, re.S):
            t = html.unescape("".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", r)))
            if not t:
                continue
            if 'w:rStyle w:val="RWBFrontItalic"' in r:
                italic.append((len(text), len(text) + len(t)))
            text += t
        spans = []
        for a, b in italic:
            if spans and spans[-1][1] == a:
                spans[-1] = (spans[-1][0], b)
            else:
                spans.append((a, b))
        others = sorted(set(re.findall(r"<w:(\w+)[ >/]", inner)) - {"pPr", "jc", "rPr", "lang", "noProof"})
        out.append(dict(style=style.group(1) if style else "-", text=text, italics=spans, jc=jc.group(1) if jc else "", sect="<w:sectPr" in p, other=others))
    return out


def main():
    a, b = load(sys.argv[1]), load(sys.argv[2])
    print("A =", sys.argv[1].replace("\\", "/").split("/")[-1], "| B =", sys.argv[2].replace("\\", "/").split("/")[-1])
    bad = 0
    for n, (x, y) in enumerate(zip(a, b), 1):
        diff = [k for k in x if x[k] != y[k]]
        if diff:
            bad += 1
            print("DIFF", n, diff, "| A:", {k: x[k] for k in diff}, "| B:", {k: y[k] for k in diff})
    print("paragraphs compared", min(len(a), len(b)), "| differences", bad)
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
