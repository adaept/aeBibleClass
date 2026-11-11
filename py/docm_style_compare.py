# docm_style_compare.py <old styles.xml> <new styles.xml>: styles added or removed and the styles whose character properties differ. Part of the RWB docm tools.
import sys
import xml.etree.ElementTree as ET
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def load(path):
    root = ET.parse(path).getroot()
    d = {}
    for s in root.iter(W + 'style'):
        rpr = s.find(W + 'rPr')
        props = []
        if rpr is not None:
            for c in rpr:
                props.append((c.tag.replace(W, ''), tuple(sorted((k.replace(W, ''), v) for k, v in c.attrib.items()))))
        d[s.get(W + 'styleId')] = (s.find(W + 'name').get(W + 'val') if s.find(W + 'name') is not None else '', tuple(sorted(props)))
    return d
a, b = load(sys.argv[1]), load(sys.argv[2])
print("styles", len(a), len(b))
only_a = sorted(set(a) - set(b)); only_b = sorted(set(b) - set(a))
print("only in 14:13:", only_a[:10], "| only now:", only_b[:10])
diff = [k for k in a if k in b and a[k] != b[k]]
print("styles whose character properties differ:", len(diff), diff[:10])
for k in diff[:5]:
    print(" ", k, "\n    14:13", a[k][1][:6], "\n    now  ", b[k][1][:6])
