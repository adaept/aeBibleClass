"""
fix_times_font.py
=================
Replaces all occurrences of the font name "Times" with "Times New Roman"
inside word/styles.xml and word/document.xml of a .docm or .docx file,
leaving all other parts of the package untouched.

Usage
-----
    python fix_times_font.py <input.docm> <output.docm>
"""

import sys
import zipfile
import os

def fix_times_font(input_path, output_path):
    print(f"Reading: {input_path}")

    # Read all ZIP members
    members = {}
    with zipfile.ZipFile(input_path, "r") as z:
        for name in z.namelist():
            members[name] = z.read(name)

    # Process both styles.xml and document.xml
    keys_to_fix = [
        "word/styles.xml",
        "word/document.xml",
        "word/fontTable.xml",
        "word/footer1.xml",
        "word/theme/theme1.xml",
        "word/vbaData.xml",
    ]
    total_replaced = 0

    for fix_key in keys_to_fix:
        if fix_key not in members:
            print(f"WARNING: {fix_key} not found — skipping.")
            continue

        xml = members[fix_key].decode("utf-8")

        # Protect "Times New Roman" with a placeholder then replace bare "Times"
        xml_fixed   = xml.replace("Times New Roman", "##PLACEHOLDER##")
        count_bare  = xml_fixed.count("Times")
        print(f"{fix_key}: bare 'Times' found: {count_bare}")

        if count_bare > 0:
            xml_fixed    = xml_fixed.replace("Times", "Times New Roman")
            xml_fixed    = xml_fixed.replace("##PLACEHOLDER##", "Times New Roman")
            members[fix_key] = xml_fixed.encode("utf-8")
            total_replaced += count_bare
            print(f"{fix_key}: {count_bare} replacement(s) made.")
        else:
            print(f"{fix_key}: nothing to replace.")

    if total_replaced == 0:
        print("\nNothing to replace. Output file not written.")
        sys.exit(0)

    print(f"\nWriting: {output_path}")
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for name, data in members.items():
            z.writestr(name, data)

    print(f"Done. {total_replaced} total replacement(s) made.")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python fix_times_font.py <input.docm> <output.docm>")
        sys.exit(1)

    input_path  = sys.argv[1]
    output_path = sys.argv[2]

    if not os.path.exists(input_path):
        print(f"ERROR: Input file not found: {input_path}")
        sys.exit(1)

    if os.path.abspath(input_path) == os.path.abspath(output_path):
        print("ERROR: Input and output paths must be different.")
        sys.exit(1)

    fix_times_font(input_path, output_path)
