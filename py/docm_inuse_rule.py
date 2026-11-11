# docm_inuse_rule.py <slots-probe.txt> <styles.xml>: which styles.xml attributes predict Word's Style.InUse, checked against the list a VBA probe printed
# (233 paragraph and character styles - 56 linked "Char" twins - 1 hidden internal style = 176). Part of the RWB docm tools.
import re, sys
import xml.etree.ElementTree as ET
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
probe = open(sys.argv[1], encoding='utf-8', errors='replace').read().splitlines()
inuse = set()
for line in probe[3:]:
    parts = line.split(' | ')
    if len(parts) >= 8 and parts[1] in ('1', '2'):
        inuse.add(parts[0].strip())
print('InUse paragraph/character styles in the VBA probe:', len(inuse))
root = ET.parse(sys.argv[2]).getroot()
rows = []
for s in root.iter(W + 'style'):
    t = s.get(W + 'type')
    if t not in ('paragraph', 'character'): continue
    name = s.find(W + 'name').get(W + 'val')
    flags = {
        'semiHidden': s.find(W + 'semiHidden') is not None,
        'unhideWhenUsed': s.find(W + 'unhideWhenUsed') is not None,
        'link': s.find(W + 'link') is not None,
        'custom': s.get(W + 'customStyle') == '1',
        'uiPriority': s.find(W + 'uiPriority').get(W + 'val') if s.find(W + 'uiPriority') is not None else None,
    }
    rows.append((name, t, flags))
print('paragraph/character styles in styles.xml:', len(rows))
# the names in the probe are Word's NameLocal (e.g. 'Heading 1'); xml names are lowercase for built-ins
low = {n.lower() for n in inuse}
def is_in(name): return name.lower() in low
tab = {}
for name, t, f in rows:
    key = (is_in(name), f['semiHidden'], f['link'] and t == 'character')
    tab[key] = tab.get(key, 0) + 1
print('(InUse, semiHidden, linked-char-twin): count')
for k in sorted(tab): print(' ', k, tab[k])
print('InUse but semiHidden examples:', [n for n, t, f in rows if is_in(n) and f['semiHidden']][:8])
print('not InUse and not semiHidden examples:', [n for n, t, f in rows if not is_in(n) and not f['semiHidden']][:25])
print('names in probe missing from styles.xml:', sorted(n for n in inuse if n.lower() not in {r[0].lower() for r in rows})[:10])
