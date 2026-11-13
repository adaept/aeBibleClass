"""
check_embedded_fonts.py
------------------------
Halting check for RT-31 (fonts embedded in the exported scripture-only .docx).

Reads word/fontTable.xml and word/_rels/fontTable.xml.rels of the saved .docx and fails (exit 1) unless:
  - at least one font is embedded (the EmbedTrueTypeFonts switch really took effect),
  - every font that carries an embedded face is in the allowed family set (ALLOWED_FAMILIES),
  - every embedded face has a relationship and the part that relationship names exists,
  - no word/fonts/ part is left without an entry (orphan).

It does not read the font bytes: the independent check of the font files (valid font, fsType, subset size)
is RT-32a (Node, aeRWB). This one only guards the export against embedding a Microsoft font by mistake.

Usage:
    python3 -I py/check_embedded_fonts.py path/to/RadiantWordBible.docx
"""

import sys
import zipfile
import xml.etree.ElementTree as ET

ALLOWED_FAMILIES = {'Carlito', 'Noto Sans', 'Liberation Serif'}

NS_W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS_R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
EMBED_TAGS = ('embedRegular', 'embedBold', 'embedItalic', 'embedBoldItalic')
FONT_TABLE = 'word/fontTable.xml'
FONT_RELS = 'word/_rels/fontTable.xml.rels'
FONT_FOLDER = 'word/fonts/'


def read_relationships(zin: zipfile.ZipFile) -> dict:
    """Relationship id -> package part name, for the font table."""
    if FONT_RELS not in zin.namelist():
        return {}
    targets = {}
    for rel in ET.fromstring(zin.read(FONT_RELS)):
        target = rel.get('Target', '')
        targets[rel.get('Id')] = 'word/' + target.lstrip('./')
    return targets


def check(docx_path: str) -> list:
    """Return the list of problems (empty means the embedded set is as allowed)."""
    problems = []
    with zipfile.ZipFile(docx_path) as zin:
        names = set(zin.namelist())
        if FONT_TABLE not in names:
            return [f'{FONT_TABLE} is missing']
        targets = read_relationships(zin)
        referenced = set()
        embedded_faces = 0
        for font in ET.fromstring(zin.read(FONT_TABLE)).iter(NS_W + 'font'):
            family = font.get(NS_W + 'name')
            for tag in EMBED_TAGS:
                for embed in font.findall(NS_W + tag):
                    embedded_faces += 1
                    if family not in ALLOWED_FAMILIES:
                        problems.append(f'{family}: {tag} embedded but the family is not allowed')
                    part = targets.get(embed.get(NS_R + 'id'))
                    if part is None:
                        problems.append(f'{family}: {tag} has no relationship')
                    elif part not in names:
                        problems.append(f'{family}: {tag} names {part}, which is not in the package')
                    else:
                        referenced.add(part)
        if embedded_faces == 0:
            problems.append('no font is embedded (EmbedTrueTypeFonts did not take effect)')
        for name in sorted(n for n in names if n.startswith(FONT_FOLDER) and not n.endswith('/')):
            if name not in referenced:
                problems.append(f'{name} is in the package but no font entry uses it (orphan)')
    return problems


def main(argv: list) -> int:
    if len(argv) != 2:
        print('Usage: python3 -I py/check_embedded_fonts.py path/to/file.docx')
        return 2
    problems = check(argv[1])
    for problem in problems:
        print('PROBLEM: ' + problem)
    if problems:
        print(f'FAILED: {len(problems)} problem(s) in the embedded fonts of {argv[1]}')
        return 1
    print('OK: embedded fonts are within the allowed families (' + ', '.join(sorted(ALLOWED_FAMILIES)) + ')')
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
