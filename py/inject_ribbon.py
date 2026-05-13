"""
inject_ribbon.py
----------------
Embeds customUI/customUI14.xml inside the target .docm or .dotm.

Two modes (auto-detected):
  REPLACE: target already has a customUI part - swap it in place
           (all other parts untouched). This is the dev-side flow used
           for the existing Bible .docm files.
  BOOTSTRAP: target has no customUI part yet (e.g. a freshly-created
             aeRibbon.dotm). Adds:
               - customUI/customUI14.xml
               - customUI/images/adaept.png
               - customUI/_rels/customUI14.xml.rels
             and patches _rels/.rels to include the ribbon relationship.

Usage:
    python py/inject_ribbon.py [path/to/target]

The target file must be closed in Word before running.
Default target: Blank Bible Copy.docm in the project root.
"""

import os
import re
import sys
import zipfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
XML_PATH     = PROJECT_ROOT / 'customUI14backupRWB.xml'
IMAGE_PATH   = PROJECT_ROOT / 'aeRibbon' / 'template' / 'images' / 'adaept.png'
DEFAULT_DOCM = PROJECT_ROOT / 'Blank Bible Copy.docm'

ENTRY_XML        = 'customUI/customUI14.xml'
ENTRY_IMAGE      = 'customUI/images/adaept.png'
ENTRY_UI_RELS    = 'customUI/_rels/customUI14.xml.rels'
ENTRY_ROOT_RELS  = '_rels/.rels'
ENTRY_CONTENT_TY = '[Content_Types].xml'

# Same Type URI as the existing dev .docm files.
RIBBON_REL_TYPE = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'
RIBBON_REL_ID   = 'R50ef49dfd3874c72'   # mirrors the dev .docm so Word treats it the same

UI_RELS_BODY = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="adaept" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
    'Target="images/adaept.png"/></Relationships>'
)


def patch_content_types(ct_xml: bytes) -> bytes:
    """Ensure [Content_Types].xml declares image/png. Idempotent."""
    text = ct_xml.decode('utf-8')
    if 'Extension="png"' in text:
        return ct_xml
    new_default = '<Default Extension="png" ContentType="image/png"/>'
    # Insert just before the closing </Types>.
    patched = re.sub(r'</Types>\s*$', new_default + '</Types>', text)
    if patched == text:
        raise RuntimeError('Could not patch [Content_Types].xml - closing tag not found')
    return patched.encode('utf-8')


def patch_root_rels(rels_xml: bytes) -> bytes:
    """Insert the customUI relationship into _rels/.rels if absent."""
    text = rels_xml.decode('utf-8')
    if RIBBON_REL_TYPE in text:
        return rels_xml  # already wired
    new_rel = (
        f'<Relationship Id="{RIBBON_REL_ID}" '
        f'Type="{RIBBON_REL_TYPE}" '
        f'Target="customUI/customUI14.xml"/>'
    )
    patched = re.sub(r'</Relationships>\s*$', new_rel + '</Relationships>', text)
    if patched == text:
        raise RuntimeError('Could not patch _rels/.rels - closing tag not found')
    return patched.encode('utf-8')


def inject(target_path: Path, xml_path: Path) -> None:
    if not xml_path.exists():
        print(f'ERROR: XML source not found: {xml_path}')
        sys.exit(1)
    if not target_path.exists():
        print(f'ERROR: Target not found: {target_path}')
        sys.exit(1)

    new_xml = xml_path.read_bytes()

    tmp_path = target_path.with_suffix(target_path.suffix + '.tmp')

    with zipfile.ZipFile(target_path, 'r') as zin:
        existing = {item.filename for item in zin.infolist()}
        replace_mode = ENTRY_XML in existing

        if not replace_mode:
            if not IMAGE_PATH.exists():
                print(f'ERROR: ribbon image not found: {IMAGE_PATH}')
                print('Bootstrap requires aeRibbon/template/images/adaept.png.')
                sys.exit(1)
            image_bytes = IMAGE_PATH.read_bytes()
        else:
            image_bytes = None

        with zipfile.ZipFile(tmp_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            wrote_xml = False
            for item in zin.infolist():
                if item.filename == ENTRY_XML:
                    zout.writestr(item, new_xml)
                    wrote_xml = True
                    print(f'REPLACED  {ENTRY_XML}')
                elif item.filename == ENTRY_ROOT_RELS and not replace_mode:
                    patched = patch_root_rels(zin.read(item.filename))
                    zout.writestr(item, patched)
                    print(f'PATCHED   {ENTRY_ROOT_RELS}')
                elif item.filename == ENTRY_CONTENT_TY:
                    original = zin.read(item.filename)
                    patched = patch_content_types(original)
                    zout.writestr(item, patched)
                    if patched != original:
                        print(f'PATCHED   {ENTRY_CONTENT_TY} (added png default)')
                else:
                    zout.writestr(item, zin.read(item.filename))

            if not replace_mode:
                # Bootstrap path: add the three customUI parts.
                zout.writestr(ENTRY_XML, new_xml)
                zout.writestr(ENTRY_IMAGE, image_bytes)
                zout.writestr(ENTRY_UI_RELS, UI_RELS_BODY)
                print(f'ADDED     {ENTRY_XML}')
                print(f'ADDED     {ENTRY_IMAGE}')
                print(f'ADDED     {ENTRY_UI_RELS}')
                wrote_xml = True

    if not wrote_xml:
        tmp_path.unlink()
        print(f'ERROR: failed to write {ENTRY_XML} into {target_path.name}.')
        sys.exit(1)

    try:
        os.replace(tmp_path, target_path)
    except PermissionError:
        tmp_path.unlink(missing_ok=True)
        print(f'ERROR: {target_path.name} is locked - close it in Word first.')
        sys.exit(1)

    mode = 'replace' if replace_mode else 'bootstrap'
    print(f'Done ({mode}).  {target_path.name} updated.')


if __name__ == '__main__':
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_DOCM
    inject(target, XML_PATH)
