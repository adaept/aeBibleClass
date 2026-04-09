import re
import sys
from pathlib import Path

# =============================================================================
# VBA Casing Normalizer
# Fixes identifier casing that Word VBA IDE corrupts due to missing globals.
# Run after export to src/ before committing.
# =============================================================================

# Each entry: (pattern, replacement, description)
# Patterns use word boundaries and are case-insensitive.
# Order matters — more specific patterns should come first.
NORMALIZATIONS = [
    (r'(?i)\bisEmpty\b',        'IsEmpty',          'IsEmpty built-in function casing'),
    (r'(?i)\bShell\b',          'Shell',            'Shell built-in function casing'),
    (r'(?i)\.Range\b',          '.Range',           '.Range property on Word Paragraph/Selection/Section'),
    (r'(?i)\bRange(?=:=)',      'Range',            'Range named argument in VBA method calls (Fields.Add, Bookmarks.Add, etc.)'),
    (r'(?i)\.Paragraphs\b',     '.Paragraphs',      '.Paragraphs collection property on Document/Range'),
    (r'(?i)\.PageSetup\b',      '.PageSetup',       '.PageSetup property on Section/Document'),
    (r'(?i)\.TopMargin\b',      '.TopMargin',       '.TopMargin property on PageSetup'),
    (r'(?i)\.BottomMargin\b',   '.BottomMargin',    '.BottomMargin property on PageSetup'),
    (r'(?i)\.PageHeight\b',     '.PageHeight',      '.PageHeight property on PageSetup'),
    (r'(?i)\.Orientation\b',    '.Orientation',     '.Orientation property on PageSetup'),
    (r'(?i)\bmid\$?\(',          'Mid$(',            'Mid$( string function — normalizes Mid( to Mid$( and fixes casing'),
    (r'(?i)\.Path\b',           '.Path',            '.Path property on Document/FileSystemObject'),
    (r'(?i)\.Item\b',           '.Item',            '.Item method on Collection'),
    (r'(?i)\.Count\b',          '.Count',           '.Count property on Collection/object'),
    (r'(?i)\.Font\b',           '.Font',            '.Font property on Range/Style/object'),
    (r'(?i)\.Keys\b',           '.Keys',            '.Keys property on Dictionary/object'),
    (r'(?i)\.Text\b',           '.Text',            '.Text property on Range/Field/object'),
    (r'(?i)\.Code\b',           '.Code',            '.Code property on Field object'),
    (r'(?i)\.Name\b',           '.Name',            '.Name property on VBProject/Style/Document/object'),
    (r'(?i)\bAs\s+(?:Word\.)?Range\b',      'As Word.Range',      'As Word.Range type declaration'),
    (r'(?i)\bAs\s+(?:Word\.)?Paragraph\b',  'As Word.Paragraph',  'As Word.Paragraph type declaration'),
    (r'(?i)\bAs\s+(?:Word\.)?Paragraphs\b', 'As Word.Paragraphs', 'As Word.Paragraphs type declaration'),
    (r'(?i)\bAs\s+(?:Word\.)?Section\b',    'As Word.Section',    'As Word.Section type declaration'),
    (r'(?i)\bAs\s+PageSetup\b',             'As PageSetup',       'As PageSetup type declaration'),
    (r'(?i)\bnote\b',                        'Note',               'Note loop variable (Footnote collection iteration)'),
    (r'(?i)\bitems\b',                       'Items',              'Items variable casing (Collection iteration)'),
]

EXTENSIONS = {'.bas', '.cls', '.frm'}

def normalize_file(path: Path) -> tuple[int, list[str]]:
    """Normalize a single file. Returns (change_count, list_of_change_descriptions)."""
    with open(path, 'r', encoding='utf-8', errors='replace', newline='') as f:
        original = f.read()
    result = original
    changes = []

    for pattern, replacement, description in NORMALIZATIONS:
        normalized = re.sub(pattern, replacement, result, flags=re.IGNORECASE)
        count = len(re.findall(pattern, result, flags=re.IGNORECASE))
        if normalized != result:
            changes.append(f'  {description}: {count} replacement(s)')
            result = normalized

    if result != original:
        with open(path, 'w', encoding='utf-8', errors='replace', newline='') as f:
            f.write(result)

    return len(changes), changes


def normalize_folder(src_folder: Path) -> None:
    print(f'Normalizing VBA source files in: {src_folder}')
    print('-' * 60)

    total_files = 0
    total_changed = 0

    for path in sorted(src_folder.rglob('*')):
        if path.suffix.lower() not in EXTENSIONS:
            continue

        total_files += 1
        change_count, changes = normalize_file(path)

        if changes:
            total_changed += 1
            print(f'CHANGED  {path.name}')
            for c in changes:
                print(c)
        else:
            print(f'ok       {path.name}')

    print('-' * 60)
    print(f'Done. {total_changed} of {total_files} files updated.')


if __name__ == '__main__':
    # Default to ./src relative to this script, or pass a path as argument
    folder = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / 'src'

    if not folder.exists():
        print(f'ERROR: Folder not found: {folder}')
        sys.exit(1)

    normalize_folder(folder)