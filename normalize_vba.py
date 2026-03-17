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
    # Word.Range — fully qualified
    (r'\bWord\.Range\b',        'Word.Range',   'Word.Range qualified'),

    # As Range — type declarations
    (r'\bAs\s+Range\b',         'As Range',     'As Range declaration'),

    # .Range — property/method access (e.g. ActiveDocument.Range, para.Range)
    (r'\.Range\b',              '.Range',       '.Range property access'),
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