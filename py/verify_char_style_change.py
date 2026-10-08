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

Usage:
    python3 -I py/verify_char_style_change.py baseline.docx new.docx

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


def main(argv):
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
