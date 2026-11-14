# docm_front_verify.py <docm or docx> [front-matter text file]: checks that the RWB front matter at the top of a SAVED docm or exported docx is exactly
# the text file's: per paragraph the style, the wording, the italic spans (character style RWBFrontItalic), the justification ({justify}), the page break before
# (the @pagebreak line: direct w:pageBreakBefore on that paragraph and on no other front paragraph), the section
# each paragraph is in, and that no Times New Roman, direct italic, page break before or empty paragraph is among them; then the page setup, vertical
# alignment and first footers of the first six sections. Reads the file's XML only (no Word). Part of the RWB docm tools; RT-22a, 2026-11-11.
# Run: wsl.exe --exec python3 -I py/docm_front_verify.py "<path>" [frontmatter/rwb-front-matter.txt]
import html
import re
import sys
import zipfile
from pathlib import Path

DEFAULT_TEXT = Path(__file__).resolve().parents[1] / "frontmatter" / "rwb-front-matter.txt"
SECTION_SCAN = 6


def run_text(run_xml):
    """Text of one run in document order; an optional (soft) hyphen element <w:softHyphen/> is returned as U+00AD (2026-11-12, owner's hand hyphenation)."""
    return "".join(html.unescape(m.group(1)) if m.group(1) is not None else "­"
                   for m in re.finditer(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>|<w:softHyphen\s*/>", run_xml))


def read_spec(path):
    want, section, brk = [], -1, False
    for line in Path(path).read_text(encoding="utf8").split("\n"):
        if not line or line.startswith("#"):
            continue
        if line.strip() == "@pagebreak":
            brk = True
            continue
        if line.startswith("@section"):
            section += 1
            continue
        style, text = line.split("\t", 1)
        justify = text.startswith("{justify}")
        if justify:
            text = text[len("{justify}"):]
        spans, plain = [], ""
        while True:
            opened = text.find("{i}")
            if opened < 0:
                plain += text
                break
            closed = text.index("{/i}", opened)
            plain += text[:opened]
            start = len(plain)
            plain += text[opened + 3:closed]
            spans.append((start, len(plain)))
            text = text[closed + 4:]
        want.append(dict(style=style, text=plain, spans=spans, justify=justify, section=section, pagebreak=brk))
        brk = False
    return want


def read_front(docx_path, count):
    x = zipfile.ZipFile(docx_path).read("word/document.xml").decode("utf8")
    body = x[x.index("<w:body>"):]
    found, section = [], 0
    for m in re.finditer(r"<w:p[ >].*?</w:p>", body[:400000], re.S):
        p = m.group(0)
        if len(found) >= count:
            break
        style = re.search(r'<w:pStyle w:val="([^"]*)"', p)
        text, italic = "", []
        for r in re.findall(r"<w:r[ >].*?</w:r>", p, re.S):
            t = run_text(r)
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
        justify = bool(re.search(r'<w:pPr>.*?<w:jc w:val="both"', re.sub(r"<w:sectPr.*?</w:sectPr>", "", p, flags=re.S), re.S))
        # direct page break before: <w:pageBreakBefore/> in the paragraph's own pPr (a val of 0 or false is "off")
        ppr = re.search(r"<w:pPr>.*?</w:pPr>", re.sub(r"<w:sectPr.*?</w:sectPr>", "", p, flags=re.S), re.S)
        brk_on = bool(ppr and re.search(r"<w:pageBreakBefore(?![^>]*w:val=\"(?:0|false)\")[^>]*/?>", ppr.group(0)))
        found.append(dict(style=style.group(1) if style else "-", text=text, spans=spans, justify=justify, section=section, pagebreak=brk_on, xml=p))
        if "<w:sectPr" in p:
            section += 1
    return found, x


def main():
    docx = sys.argv[1]
    spec = Path(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_TEXT
    want = read_spec(spec)
    got, x = read_front(docx, len(want))
    bad = 0
    for n, (w, g) in enumerate(zip(want, got), 1):
        for key in ("style", "text", "spans", "justify", "section", "pagebreak"):
            if w[key] != g[key]:
                bad += 1
                print("MISMATCH paragraph", n, key, "| want", w[key] if key != "text" else w[key][:40], "| got", g[key] if key != "text" else g[key][:40])
    print("paragraphs checked", len(got), "of", len(want), "| mismatches", bad)
    front = "".join(g["xml"] for g in got)
    print("Times New Roman:", front.count("Times New Roman"), "| direct w:i:", len(re.findall(r"<w:i[ />]", front)), "| page break before:", sum(1 for g in got if g["pagebreak"]), "(wanted", sum(1 for w in want if w["pagebreak"]), ")",
          "| empty paragraphs among them:", sum(1 for g in got if not g["text"]))
    sections = [m.group(0) for m in re.finditer(r"<w:sectPr[ >].*?</w:sectPr>", x, re.S)]
    print("sections in the file:", len(sections))
    for i, s in enumerate(sections[:SECTION_SCAN], 1):
        def pick(pattern):
            m = re.search(pattern, s)
            return m.group(1) if m else "-"
        print("  ", i, "type", pick(r'<w:type w:val="([^"]*)"'), "| page", pick(r'<w:pgSz w:w="(\d+)"'), "x", pick(r'<w:pgSz [^>]*w:h="(\d+)"'),
              "| top/left margin", pick(r'<w:pgMar [^>]*w:top="(\d+)"'), pick(r'<w:pgMar [^>]*w:left="(\d+)"'), "| vAlign", pick(r'<w:vAlign w:val="([^"]*)"'),
              "| different first page", "titlePg" in s, "| page numbering", pick(r"<w:pgNumType ([^/]*)/"))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
