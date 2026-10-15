"""
count_explicit_default_paragraph_font.py
----------------------------------------
Release-gate guard (decided 2026-10-15): counts EXPLICIT applications of the
"Default Paragraph Font" character style in a saved Word file.

Background: "Default Paragraph Font" is Word's "no character style" placeholder.
Text with no character style carries NO marker in the file (implicit - normal,
and indistinguishable from an explicit one through Word's object model). An
EXPLICIT application is a run whose properties contain
    <w:rStyle w:val="DefaultParagraphFont"/>
i.e. a stray style left by a paste or a tool. It does nothing useful and can
interfere with later Find/Replace and style-normalisation runs, so the release
requires the count to be 0 in the final .docx (and in the archived .docm).
Word never writes this marker for ordinary text; both the dev docm and the
v1.0 RadiantWordBible.docx counted 0 on 2026-10-15.

Scans every word/*.xml part (document, headers, footers, footnotes, endnotes,
comments, numbering) EXCEPT styles.xml (style definitions, not applications).

Usage:
    python3 -I py/count_explicit_default_paragraph_font.py file.docx [file.docm ...]

Exit code 0 = every file has 0 explicit applications, 1 = at least one found,
2 = usage/error. Read-only: never modifies a file.
"""

import re
import sys
import zipfile

PATTERN = re.compile(r'<w:rStyle\b[^>]*\bw:val="DefaultParagraphFont"')


def count_in_file(path):
    """Returns (total, {part_name: count}) for one .docx/.docm."""
    per_part = {}
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not (name.startswith("word/") and name.endswith(".xml")):
                continue
            if name == "word/styles.xml":
                continue
            n = len(PATTERN.findall(z.read(name).decode("utf8", "ignore")))
            if n:
                per_part[name] = n
    return sum(per_part.values()), per_part


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 2
    failed = False
    for path in argv[1:]:
        try:
            total, per_part = count_in_file(path)
        except (OSError, zipfile.BadZipFile) as exc:
            print("ERROR  %s: %s" % (path, exc))
            return 2
        status = "PASS" if total == 0 else "FAIL"
        print("%s  %s: explicit DefaultParagraphFont applications = %d" % (status, path, total))
        for part, n in sorted(per_part.items()):
            print("         %s: %d" % (part, n))
        if total:
            failed = True
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
