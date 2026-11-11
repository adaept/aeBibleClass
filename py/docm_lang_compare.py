# docm_lang_compare.py <older document.xml> <newer document.xml>: the paragraphs with a de-DE run or paragraph mark in each, and the differences
# (how the German tags came back, FU-30). Part of the RWB docm tools (tools/word-com/README.md).
import sys
import xml.etree.ElementTree as ET
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def langs(path):
    root = ET.parse(path).getroot()
    out = []
    for i, p in enumerate(root.iter(W + 'p')):
        style = p.find(W + 'pPr/' + W + 'pStyle')
        style = style.get(W + 'val') if style is not None else 'Normal'
        mark = p.find(W + 'pPr/' + W + 'rPr/' + W + 'lang')
        runs = []
        for r in p.findall(W + 'r'):
            l = r.find(W + 'rPr/' + W + 'lang')
            if l is not None and l.get(W + 'val') == 'de-DE':
                t = ''.join((x.text or '') for x in r.iter(W + 't'))[:14]
                runs.append(t)
        m = mark.get(W + 'val') if mark is not None else None
        if m == 'de-DE' or runs:
            out.append((i, style, m, tuple(runs)))
    return out
a = langs(sys.argv[1])  # the older document.xml
b = langs(sys.argv[2])  # the newer document.xml
print('paragraphs with a de-DE run or mark: 18:50 =', len(a), ' 20:45 =', len(b))
sa = {x[0]: x for x in a}; sb = {x[0]: x for x in b}
by_style = {}
for x in b: by_style[x[1]] = by_style.get(x[1], 0) + 1
print('20:45 by paragraph style:', by_style)
print('only in 20:45:', [(k, sb[k][1], sb[k][2], sb[k][3]) for k in sb if k not in sa][:20])
print('only in 18:50:', [(k, sa[k][1], sa[k][2], sa[k][3]) for k in sa if k not in sb][:10])
print('in both but different:', [(k, sa[k][2:], sb[k][2:]) for k in sa if k in sb and sa[k] != sb[k]][:10])
