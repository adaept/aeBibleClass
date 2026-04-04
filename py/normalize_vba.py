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
    (r'(?i)\.Range\b',          '.Range',           '.Range property access'),
    (r'(?i)\.Paragraphs\b',     '.Paragraphs',      '.Paragraphs property access'),
    (r'(?i)\.PageSetup\b',      '.PageSetup',       '.PageSetup property access'),
    (r'(?i)\.TopMargin\b',      '.TopMargin',       '.TopMargin on PageSetup'),
    (r'(?i)\.BottomMargin\b',   '.BottomMargin',    '.BottomMargin on PageSetup'),
    (r'(?i)\.PageHeight\b',     '.PageHeight',      '.PageHeight on PageSetup'),
    (r'(?i)\.Orientation\b',    '.Orientation',     '.Orientation on PageSetup'),
    (r'(?i)\bmid\$?\(',           'Mid$(',            'Mid$( built-in function casing (includes Mid( -> Mid$()'),
    (r'(?i)\.Path\b',           '.Path',            '.Path property on Document/ActiveDocument'),
    (r'(?i)\.Item\b',           '.Item',            '.Item method on Collection'),
    (r'(?i)\.Keys\b',           '.Keys',            '.Keys property on Dictionary/Object'),
    (r'(?i)\.Text\b',           '.Text',            '.Text property on Range'),
    (r'(?i)\bAs\s+(?:Word\.)?Range\b',      'As Word.Range',      'As Word.Range declaration'),
    (r'(?i)\bAs\s+(?:Word\.)?Paragraph\b',  'As Word.Paragraph',  'As Word.Paragraph declaration'),
    (r'(?i)\bAs\s+(?:Word\.)?Paragraphs\b', 'As Word.Paragraphs', 'As Word.Paragraphs declaration'),
    (r'(?i)\bnote\b',                        'Note',               'Note loop variable (Footnote iteration)'),
    (r'(?i)\bitems\b',                       'Items',              'Items variable casing (Collection)'),
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