"""
strip_ribbon.py
----------------
Removes the customUI/customUI14.xml Ribbon customization from a .docx --
the inverse of inject_ribbon.py's REPLACE mode.

Why this exists: inject_ribbon.py deliberately embeds the ribbon XML
directly into the dev .docm files (a convenience so the RWB tab shows
without attaching aeRibbon.dotm as a global template during development).
Word's own File -> Save As -> .docx strips vbaProject.bin (the macro
security half) but has no reason to strip an unrelated OOXML part like
customUI/customUI14.xml -- it isn't part of that operation. So a .docx
produced by Save-As from one of these docm files always carries an
orphaned ribbon customization with no macros behind it, and Word pops
"the macro can't be found" repeatedly on every open (confirmed the hard
way: aeBibleAddin's Bible.docx test fixture, 2026-09-12). This script is
the production-docx pipeline's fix for that, run right after Save-As.

Removes:
  - customUI/customUI14.xml
  - customUI/images/adaept.png
  - customUI/_rels/customUI14.xml.rels
  - the ribbon Relationship entry in _rels/.rels

Deliberately leaves [Content_Types].xml untouched -- its generic
`<Default Extension="png" .../>` declaration is shared by any other PNG
part in the package (e.g. inline images in the Bible content itself);
removing it would break those too, not just the ribbon's icon.

Usage:
    python py/strip_ribbon.py path/to/Word-Bible.docx

Verify first (should print the same three lines this script removes):
    unzip -l path/to/Word-Bible.docx | grep -i customUI

The target file must be closed in Word before running. Exits non-zero
(no changes made) if the target has no customUI part to strip, or if it
still carries vbaProject.bin (this script is for the production docx,
not a dev .docm/.dotm -- stripping the ribbon from those would be wrong).
"""

import os
import re
import sys
import zipfile
from pathlib import Path

ENTRY_XML       = 'customUI/customUI14.xml'
ENTRY_IMAGE     = 'customUI/images/adaept.png'
ENTRY_UI_RELS   = 'customUI/_rels/customUI14.xml.rels'
ENTRY_ROOT_RELS = '_rels/.rels'
ENTRY_VBA       = 'word/vbaProject.bin'

STRIPPED_ENTRIES = {ENTRY_XML, ENTRY_IMAGE, ENTRY_UI_RELS}

# Same Type URI inject_ribbon.py writes -- match it exactly so this is a
# true inverse, not a guess at what a customUI relationship looks like.
RIBBON_REL_TYPE = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'


def strip_root_rels(rels_xml: bytes) -> bytes:
    """Remove the customUI Relationship element from _rels/.rels."""
    text = rels_xml.decode('utf-8')
    pattern = re.compile(r'<Relationship\b[^>]*Type="' + re.escape(RIBBON_REL_TYPE) + r'"[^>]*/>')
    patched, count = pattern.subn('', text)
    if count == 0:
        raise RuntimeError(f'{ENTRY_ROOT_RELS}: no ribbon Relationship found to remove')
    if count > 1:
        raise RuntimeError(f'{ENTRY_ROOT_RELS}: expected exactly one ribbon Relationship, found {count}')
    return patched.encode('utf-8')


def strip(target_path: Path) -> None:
    if not target_path.exists():
        print(f'ERROR: Target not found: {target_path}')
        sys.exit(1)

    tmp_path = target_path.with_suffix(target_path.suffix + '.tmp')

    with zipfile.ZipFile(target_path, 'r') as zin:
        names = {item.filename for item in zin.infolist()}

        if ENTRY_XML not in names:
            print(f'Nothing to strip: {target_path.name} has no {ENTRY_XML}.')
            sys.exit(1)

        if ENTRY_VBA in names:
            print(f'ERROR: {target_path.name} still has {ENTRY_VBA} -- this looks like a dev '
                  '.docm/.dotm, not a production .docx. Refusing to strip its ribbon.')
            sys.exit(1)

        with zipfile.ZipFile(tmp_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in STRIPPED_ENTRIES:
                    print(f'REMOVED   {item.filename}')
                    continue
                if item.filename == ENTRY_ROOT_RELS:
                    patched = strip_root_rels(zin.read(item.filename))
                    zout.writestr(item, patched)
                    print(f'PATCHED   {ENTRY_ROOT_RELS}')
                else:
                    zout.writestr(item, zin.read(item.filename))

    try:
        os.replace(tmp_path, target_path)
    except PermissionError:
        tmp_path.unlink(missing_ok=True)
        print(f'ERROR: {target_path.name} is locked - close it in Word first.')
        sys.exit(1)

    print(f'Done. {target_path.name} no longer carries a Ribbon customization.')


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print('Usage: python py/strip_ribbon.py path/to/file.docx')
        sys.exit(1)
    strip(Path(sys.argv[1]))
