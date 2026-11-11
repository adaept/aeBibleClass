# docm_text_compare.py <document.xml> [more.xml ...]: character count and SHA-256 of the w:t text (the identity check: a clean-up must not change it) and
# the number of de-DE lang elements. Part of the RWB docm tools (tools/word-com/README.md).
import sys, hashlib
import xml.etree.ElementTree as ET
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
for path in sys.argv[1:]:
    root = ET.parse(path).getroot()
    text = "".join((t.text or "") for t in root.iter(W + "t"))
    langs = sum(1 for l in root.iter(W + "lang") if l.get(W + "val") == "de-DE")
    print(path.split("/")[-1], "w:t characters", len(text), "sha", hashlib.sha256(text.encode()).hexdigest()[:16], "| lang de-DE elements", langs)
