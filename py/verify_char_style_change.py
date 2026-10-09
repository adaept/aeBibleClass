"""
verify_char_style_change.py
---------------------------
Offline OOXML check for RadiantWordBible plan items B2/B3 (2026-10-08):
EmphasisBlack character style removed (text takes VerseText), EmphasisRed
replaced by "Words of Jesus" - and NOTHING else changed.

Compares a baseline .docx (e.g. Bible/v0.0.RadiantWordBible.docx) with a
new export. Checks:
  1. Paragraph style sequence identical.
  2. Paragraph text identical, excluding the Bible Index paragraphs
     (BibleIndexEyebrow / BibleIndex / BibleIndexList): their page numbers
     legitimately change with pagination.
  3. Characters covered per character style (run counts are unreliable -
     runs merge on save):
       EmphasisBlack -> 0, EmphasisRed -> 0,
       WordsofJesus_after == WordsofJesus_before + EmphasisRed_before,
       every other character style unchanged.

  4. Optional (--verse-align left|both|center|right): the VerseText style
     alignment in styles.xml equals the expected value ("left" also accepts
     "start" or an absent jc) and no VerseText paragraph carries a direct
     alignment override that differs (VerseText left-alignment, v1.0 task).

  5. Always: every two-column (book body) section has a non-empty effective
     default header (its own, else inherited from the previous section) -
     catches a book losing its running header (B4, Revelation, 2026-10-08).

  6. Always: every paragraph/character style used in the baseline's header
     and footer parts (word/header*.xml, word/footer*.xml) is still used in
     the new export, and the style counts are reported. document.xml-only
     checks missed the style purge deleting TheHeaders/TheFooters while 132
     header + 1 footer paragraphs used them (2026-10-09).

Usage:
    python3 -I py/verify_char_style_change.py baseline.docx new.docx [--verse-align left]

Exit code 0 = all checks pass, 1 = at least one failure, 2 = usage/error.
Read-only: never modifies either file.
"""

import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
INDEX_STYLES = {'BibleIndexEyebrow', 'BibleIndex', 'BibleIndexList'}


def read(path):
    """Return (paragraph style list, paragraph text list excl. index rows,
    Counter of chars covered per character style id)."""
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read('word/document.xml'))
    styles = []
    texts = []
    chars = Counter()
    for p in root.iter(W + 'p'):
        ppr = p.find(W + 'pPr')
        ps = ''
        if ppr is not None:
            st = ppr.find(W + 'pStyle')
            if st is not None:
                ps = st.get(W + 'val', '')
        styles.append(ps)
        buf = []
        for r in p.iter(W + 'r'):
            rpr = r.find(W + 'rPr')
            rs = ''
            if rpr is not None:
                rst = rpr.find(W + 'rStyle')
                if rst is not None:
                    rs = rst.get(W + 'val', '')
            for child in r:
                if child.tag == W + 't':
                    t = child.text or ''
                    buf.append(t)
                    if rs:
                        chars[rs] += len(t)
                elif child.tag == W + 'tab':
                    buf.append('\t')
        if ps not in INDEX_STYLES:
            texts.append(''.join(buf))
    return styles, texts, chars


def header_footer_styles(path):
    """Return Counter of ('p'|'r', style id) used in word/header*.xml and
    word/footer*.xml."""
    import re
    c = Counter()
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if re.match(r'word/(header|footer)\d*\.xml$', name):
                x = z.read(name).decode('utf-8')
                for m in re.findall(r'<w:pStyle w:val="([^"]+)"', x):
                    c[('p', m)] += 1
                for m in re.findall(r'<w:rStyle w:val="([^"]+)"', x):
                    c[('r', m)] += 1
    return c


def verse_alignment(path):
    """Return (style jc or None, Counter of direct jc overrides on VerseText paragraphs)."""
    with zipfile.ZipFile(path) as z:
        sroot = ET.fromstring(z.read('word/styles.xml'))
        droot = ET.fromstring(z.read('word/document.xml'))
    style_jc = None
    for st in sroot.iter(W + 'style'):
        if st.get(W + 'styleId') == 'VerseText':
            jc = st.find(W + 'pPr/' + W + 'jc')
            style_jc = jc.get(W + 'val') if jc is not None else None
    overrides = Counter()
    for p in droot.iter(W + 'p'):
        ppr = p.find(W + 'pPr')
        if ppr is None:
            continue
        st = ppr.find(W + 'pStyle')
        if st is None or st.get(W + 'val') != 'VerseText':
            continue
        jc = ppr.find(W + 'jc')
        if jc is not None:
            overrides[jc.get(W + 'val')] += 1
    return style_jc, overrides


def body_section_headers(path):
    """Return a list of (section number, header text) for every two-column
    section whose effective default header text is empty."""
    import re
    with zipfile.ZipFile(path) as z:
        doc = z.read('word/document.xml').decode('utf-8')
        rels_xml = z.read('word/_rels/document.xml.rels').decode('utf-8')
        rels = {}
        for m in re.finditer(r'<Relationship [^>]*>', rels_xml):
            rid = re.search(r'Id="([^"]+)"', m.group(0))
            tgt = re.search(r'Target="([^"]+)"', m.group(0))
            if rid and tgt:
                rels[rid.group(1)] = tgt.group(1)
        empty = []
        effective = ''
        for n, m in enumerate(re.finditer(r'<w:sectPr[ >].*?</w:sectPr>', doc, re.S), 1):
            sp = m.group(0)
            ref = re.search(r'<w:headerReference w:type="default" r:id="([^"]+)"', sp)
            if ref:
                x = z.read('word/' + rels[ref.group(1)]).decode('utf-8')
                effective = ''.join(re.findall(r'<w:t[ >][^<]*|<w:t>[^<]*', x))
                effective = re.sub(r'<w:t[^>]*>', '', effective).strip()
            if re.search(r'<w:cols [^>]*w:num="2"', sp) and not effective:
                empty.append(n)
        return empty


def main(argv):
    expect_align = None
    if '--verse-align' in argv:
        i = argv.index('--verse-align')
        if i + 1 >= len(argv):
            print(__doc__)
            return 2
        expect_align = argv[i + 1]
        argv = argv[:i] + argv[i + 2:]
    if len(argv) != 3:
        print(__doc__)
        return 2
    b_styles, b_texts, b_chars = read(argv[1])
    n_styles, n_texts, n_chars = read(argv[2])
    failures = []

    if b_styles != n_styles:
        failures.append('paragraph style sequence differs (%d vs %d paragraphs)'
                        % (len(b_styles), len(n_styles)))
    if b_texts != n_texts:
        diff = sum(1 for a, b in zip(b_texts, n_texts) if a != b) + abs(len(b_texts) - len(n_texts))
        failures.append('paragraph text differs in %d paragraph(s) (index rows excluded)' % diff)

    for sid in ('EmphasisBlack', 'EmphasisRed'):
        if n_chars.get(sid, 0) != 0:
            failures.append('%s still covers %d chars, expected 0' % (sid, n_chars[sid]))
    want = b_chars.get('WordsofJesus', 0) + b_chars.get('EmphasisRed', 0)
    if n_chars.get('WordsofJesus', 0) != want:
        failures.append('WordsofJesus covers %d chars, expected %d'
                        % (n_chars.get('WordsofJesus', 0), want))
    others = (set(b_chars) | set(n_chars)) - {'EmphasisBlack', 'EmphasisRed', 'WordsofJesus'}
    for sid in sorted(others):
        if b_chars.get(sid, 0) != n_chars.get(sid, 0):
            failures.append('%s changed: %d -> %d chars'
                            % (sid, b_chars.get(sid, 0), n_chars.get(sid, 0)))

    if expect_align is not None:
        norm = {'left': {'left', 'start', None}}.get(expect_align, {expect_align})
        style_jc, overrides = verse_alignment(argv[2])
        if style_jc not in norm:
            failures.append('VerseText style alignment is %r, expected %s' % (style_jc, expect_align))
        bad = sum(n for v, n in overrides.items() if v not in norm)
        if bad:
            failures.append('%d VerseText paragraph(s) override alignment (expected %s)' % (bad, expect_align))
        print('VerseText style jc=%r, direct overrides=%s' % (style_jc, dict(overrides)))

    empty_hdr = body_section_headers(argv[2])
    if empty_hdr:
        failures.append('two-column section(s) with an empty header: %s' % empty_hdr)
    print('two-column sections with empty header:', empty_hdr)

    b_hf = header_footer_styles(argv[1])
    n_hf = header_footer_styles(argv[2])
    lost = sorted(k for k in b_hf if n_hf.get(k, 0) == 0)
    if lost:
        failures.append('style(s) used in baseline headers/footers but absent from the new '
                        'headers/footers: %s' % [('%s:%s' % k) for k in lost])
    print('header/footer styles baseline:', {('%s:%s' % k): v for k, v in b_hf.items()})
    print('header/footer styles new     :', {('%s:%s' % k): v for k, v in n_hf.items()})

    print('baseline chars by style:', dict(b_chars))
    print('new      chars by style:', dict(n_chars))
    if failures:
        for f in failures:
            print('FAIL', f)
        return 1
    print('PASS: only EmphasisBlack/EmphasisRed changed; all else identical.')
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
