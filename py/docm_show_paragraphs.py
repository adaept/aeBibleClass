# docm_show_paragraphs.py <document.xml> <index> [index ...]: the XML of the paragraphs with these indexes (document order, all w:p), shortened.
# Part of the RWB docm tools (tools/word-com/README.md).
import sys
import xml.etree.ElementTree as ET
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
root = ET.parse(sys.argv[1]).getroot()
paras = list(root.iter(W + 'p'))
print('paragraphs', len(paras))
def short(el, depth=0):
    tag = el.tag.replace(W, 'w:')
    attrs = ' '.join(f'{k.replace(W, "w:")}="{v}"' for k, v in el.attrib.items())
    txt = (el.text or '').strip()
    kids = ''.join(short(c, depth + 1) for c in el)
    return f'<{tag}{" " + attrs if attrs else ""}>{txt[:40]}{kids}</{tag}>' if (kids or txt) else f'<{tag}{" " + attrs if attrs else ""}/>'
for i in map(int, sys.argv[2:]):
    p = paras[i]
    print(f'--- paragraph {i}: {short(p)[:1100]}')
