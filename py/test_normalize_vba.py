#!/usr/bin/env python3
"""Self-test for py/normalize_vba.py (added 2026-11-12).

The casing rules must change CODE only: a word in a comment or a string literal stays as written (the VBE never changes those), and
`.style` is put back to `.Style` (the VBE demotes it). A second run changes nothing, and CRLF line ends are kept.

Run (WSL, like every Python tool here):  python3 -I py/test_normalize_vba.py
"""
import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

sys.dont_write_bytecode = True  # do not leave py/__pycache__ behind

HERE = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location("normalize_vba", HERE / "normalize_vba.py")
norm = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(norm)


def run(text: str) -> str:
    """Normalize `text` as a .bas file and return the result."""
    with tempfile.TemporaryDirectory() as folder:
        path = Path(folder) / "m.bas"
        path.write_bytes(text.encode("utf-8"))
        norm.normalize_file(path)
        return path.read_bytes().decode("utf-8")


class CodeOnly(unittest.TestCase):
    def test_code_is_normalized(self):
        self.assertEqual(run("For Each note In oRng.Footnotes\r\n"), "For Each Note In oRng.Footnotes\r\n")

    def test_comment_is_left_alone(self):
        text = "' any note, field, result or items here\r\nx = 1 ' and a note after code\r\n"
        self.assertEqual(run(text), text)

    def test_string_literal_is_left_alone(self):
        text = 'Debug.Print "the result is a note"\r\n'
        self.assertEqual(run(text), text)

    def test_escaped_quote_does_not_end_the_string(self):
        text = 'Debug.Print "say ""result"" now"\r\n'
        self.assertEqual(run(text), text)

    def test_code_after_a_string_on_the_same_line_is_normalized(self):
        self.assertEqual(run('x = "a result" & note\r\n'), 'x = "a result" & Note\r\n')

    def test_apostrophe_inside_a_string_is_not_a_comment(self):
        self.assertEqual(run('x = "it\'s" & note\r\n'), 'x = "it\'s" & Note\r\n')

    def test_rem_line_is_a_comment(self):
        text = "Rem the note and the result\r\n"
        self.assertEqual(run(text), text)


class StyleRule(unittest.TestCase):
    def test_lowercase_style_property_becomes_style(self):
        self.assertEqual(run("If oPara.style.NameLocal = s Then\r\n"), "If oPara.Style.NameLocal = s Then\r\n")
        self.assertEqual(run("    .style = oStyle\r\n"), "    .Style = oStyle\r\n")

    def test_other_members_that_start_with_style_are_untouched(self):
        text = "x = oDoc.Styles(n)\r\ny = o.StyleName\r\nz = o.styleCount\r\n"
        self.assertEqual(run(text), text)

    def test_style_in_a_comment_or_string_is_untouched(self):
        text = "' use oPara.style here\r\nDebug.Print \"oPara.style\"\r\n"
        self.assertEqual(run(text), text)


class ExecRule(unittest.TestCase):
    def test_lowercase_exec_method_becomes_exec(self):
        self.assertEqual(run("Set oExec = wsh.exec(shellCmd)\r\n"), "Set oExec = wsh.Exec(shellCmd)\r\n")
        self.assertEqual(run("Set e = CreateObject(\"WScript.Shell\").exec(\"cmd\")\r\n"),
                         "Set e = CreateObject(\"WScript.Shell\").Exec(\"cmd\")\r\n")

    def test_other_members_and_names_that_contain_exec_are_untouched(self):
        text = "o.ExecuteItem i\r\nx = rs.Execute\r\ny = o.exec2\r\nSet oExec = Nothing\r\nDim exec As Long\r\n"
        self.assertEqual(run(text), text)

    def test_exec_in_a_comment_or_string_is_untouched(self):
        text = "' call wsh.exec here\r\nDebug.Print \"wsh.exec\"\r\n"
        self.assertEqual(run(text), text)


class Stability(unittest.TestCase):
    SAMPLE = ("For Each note In r.Footnotes\r\n    x = o.style.NameLocal ' a note\r\nNext note\r\n"
              "s = Space(2) & \"the result\"\r\n\r\n\r\n")

    def test_second_run_changes_nothing(self):
        once = run(self.SAMPLE)
        self.assertEqual(run(once), once)

    def test_crlf_is_kept_and_trailing_blank_lines_are_still_collapsed(self):
        once = run(self.SAMPLE)
        self.assertNotIn("\r\n\r\n\r\n", once)
        self.assertEqual(once.count("\n"), once.count("\r\n"))

    def test_code_spans(self):
        line = 'a = "x" & b \' c'
        spans = norm.code_spans(line)
        self.assertEqual([line[s:e] for s, e in spans], ["a = ", " & b "])


if __name__ == "__main__":
    unittest.main()
