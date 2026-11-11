# docm_paragraph_containers.py <document.xml>: the w:p elements grouped by container (table cell, text box in mc:Choice or mc:Fallback, content control)
# with an example each. It found the 100 extra paragraphs of FU-28 (50 text boxes stored twice). Part of the RWB docm tools.
import sys
import collections
import xml.etree.ElementTree as ET
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
MC = '{http://schemas.openxmlformats.org/markup-compatibility/2006}'
root = ET.parse(sys.argv[1]).getroot()
parent = {c: p for p in root.iter() for c in p}
def ancestors(el):
    out = []
    while el in parent:
        el = parent[el]; out.append(el.tag.split('}')[1])
    return out
total = 0; groups = collections.Counter(); samples = {}
for i, p in enumerate(root.iter(W + 'p')):
    total += 1
    anc = ancestors(p)
    key = tuple(a for a in anc if a in ('txbxContent', 'tbl', 'tc', 'sdtContent', 'Fallback', 'Choice', 'AlternateContent', 'pict', 'drawing', 'ins', 'del', 'moveFrom'))
    groups[key] += 1
    if key not in samples:
        samples[key] = (i, ''.join((t.text or '') for t in p.iter(W + 't'))[:50])
print('w:p total', total)
for k, n in groups.most_common(): print(f'{n:6d}  {k}  e.g. paragraph {samples[k][0]} "{samples[k][1]}"')
