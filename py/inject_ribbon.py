"""
inject_ribbon.py
----------------
Replaces customUI/customUI14.xml inside the target .docm with the contents of
customUI14backupRWB.xml.  All other parts of the file are untouched.

Usage:
    python py/inject_ribbon.py [path/to/target.docm]

The .docm must be closed in Word before running.
Default target: Blank Bible Copy.docm in the project root.
"""

import os
import sys
import zipfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
XML_PATH     = PROJECT_ROOT / 'customUI14backupRWB.xml'
DEFAULT_DOCM = PROJECT_ROOT / 'Blank Bible Copy.docm'
ENTRY        = 'customUI/customUI14.xml'


def inject(docm_path: Path, xml_path: Path) -> None:
    if not xml_path.exists():
        print(f'ERROR: XML source not found: {xml_path}')
        sys.exit(1)
    if not docm_path.exists():
        print(f'ERROR: Target .docm not found: {docm_path}')
        sys.exit(1)

    # Stale lock files are common after a Word crash — test writability directly.
    try:
        with open(docm_path, 'r+b'):
            pass
    except PermissionError:
        print(f'ERROR: {docm_path.name} is locked — close it in Word first.')
        sys.exit(1)

    with open(xml_path, 'r', encoding='utf-8') as f:
        new_xml = f.read().encode('utf-8')

    tmp_path = docm_path.with_suffix('.docm.tmp')

    with zipfile.ZipFile(docm_path, 'r') as zin:
        with zipfile.ZipFile(tmp_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            replaced = False
            for item in zin.infolist():
                if item.filename == ENTRY:
                    zout.writestr(item, new_xml)
                    replaced = True
                    print(f'REPLACED  {ENTRY}')
                else:
                    zout.writestr(item, zin.read(item.filename))

    if not replaced:
        tmp_path.unlink()
        print(f'ERROR: Entry "{ENTRY}" not found in {docm_path.name}.')
        print('The file may not have a customUI part yet — use RibbonX Editor to add one first.')
        sys.exit(1)

    os.replace(tmp_path, docm_path)
    print(f'Done.  {docm_path.name} updated.')


if __name__ == '__main__':
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_DOCM
    inject(target, XML_PATH)
