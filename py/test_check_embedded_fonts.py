"""Tests for check_embedded_fonts.py (synthetic packages, no Word, no real fonts).

Run: wsl.exe --exec python3 -I py/test_check_embedded_fonts.py
"""

import os
import sys
import tempfile
import unittest
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import check_embedded_fonts as cef  # noqa: E402

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def font_entry(name, ids):
    embeds = ''.join(f'<w:embedRegular r:id="{i}" w:fontKey="{{00000000-0000-0000-0000-000000000000}}"/>' for i in ids)
    return f'<w:font w:name="{name}">{embeds}</w:font>'


def make_docx(entries, rels, parts):
    table = f'<w:fonts xmlns:w="{W}" xmlns:r="{R}">' + ''.join(entries) + '</w:fonts>'
    rel_xml = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
               ''.join(f'<Relationship Id="{i}" Target="{t}"/>' for i, t in rels) + '</Relationships>')
    handle, path = tempfile.mkstemp(suffix='.docx')
    os.close(handle)
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr(cef.FONT_TABLE, table)
        z.writestr(cef.FONT_RELS, rel_xml)
        for part in parts:
            z.writestr(part, b'x')
    return path


class CheckEmbeddedFonts(unittest.TestCase):
    def run_check(self, *args):
        path = make_docx(*args)
        try:
            return cef.check(path)
        finally:
            os.remove(path)

    def test_allowed_family_passes(self):
        self.assertEqual(self.run_check([font_entry('Carlito', ['rId1'])], [('rId1', 'fonts/font1.odttf')],
                                        ['word/fonts/font1.odttf']), [])

    def test_microsoft_family_fails(self):
        problems = self.run_check([font_entry('Calibri', ['rId1'])], [('rId1', 'fonts/font1.odttf')],
                                  ['word/fonts/font1.odttf'])
        self.assertTrue(any('not allowed' in p for p in problems))

    def test_nothing_embedded_fails(self):
        problems = self.run_check(['<w:font w:name="Carlito"/>'], [], [])
        self.assertTrue(any('no font is embedded' in p for p in problems))

    def test_missing_part_fails(self):
        problems = self.run_check([font_entry('Carlito', ['rId1'])], [('rId1', 'fonts/font1.odttf')], [])
        self.assertTrue(any('not in the package' in p for p in problems))

    def test_missing_relationship_fails(self):
        problems = self.run_check([font_entry('Carlito', ['rId9'])], [], [])
        self.assertTrue(any('no relationship' in p for p in problems))

    def test_orphan_part_fails(self):
        problems = self.run_check([font_entry('Carlito', ['rId1'])], [('rId1', 'fonts/font1.odttf')],
                                  ['word/fonts/font1.odttf', 'word/fonts/font2.odttf'])
        self.assertTrue(any('orphan' in p for p in problems))


if __name__ == '__main__':
    unittest.main()
