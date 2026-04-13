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
    (r'(?i)\.content\b',         '.Content',         '.Content property on Word.Document (canonical uppercase C)'),
    (r'(?i)\.Paragraphs\b',     '.Paragraphs',      '.Paragraphs collection property on Document/Range'),
    (r'(?i)\.PageSetup\b',      '.PageSetup',       '.PageSetup property on Section/Document'),
    (r'(?i)\.TopMargin\b',      '.TopMargin',       '.TopMargin property on PageSetup'),
    (r'(?i)\.BottomMargin\b',   '.BottomMargin',    '.BottomMargin property on PageSetup'),
    (r'(?i)\.PageHeight\b',     '.PageHeight',      '.PageHeight property on PageSetup'),
    (r'(?i)\.Orientation\b',    '.Orientation',     '.Orientation property on PageSetup'),
    (r'(?i)\bmid\$?\(',          'Mid$(',            'Mid$( string function — normalizes Mid( to Mid$( and fixes casing'),
    (r'(?i)\.Path\b',           '.Path',            '.Path property on Document/FileSystemObject'),
    (r'(?i)\.Item\b',           '.Item',            '.Item method on Collection'),
    (r'(?i)\bCount(?=:=)',       'Count',            'Count named argument in VBA method calls (e.g. MoveDown Count:=)'),
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
    # --- Long-process framework identifiers (added 2026-04-09) ---
    (r'(?i)\bStyleName\b',                   'StyleName',          'StyleName property on aeUpdateCharStyleClass'),
    (r'(?i)\bBatchSize\b',                   'BatchSize',          'BatchSize property on aeLongProcessClass'),
    (r'(?i)\bPauseMs\b',                     'PauseMs',            'PauseMs property on aeLongProcessClass'),
    (r'(?i)\bTaskName\b',                    'TaskName',           'TaskName property on IaeLongProcessClass'),
    (r'(?i)\bItemCount\b',                   'ItemCount',          'ItemCount property on IaeLongProcessClass'),
    (r'(?i)\bExecuteItem\b',                 'ExecuteItem',        'ExecuteItem method on IaeLongProcessClass'),
    (r'(?i)\bStartOrResume\b',               'StartOrResume',      'StartOrResume entry point in basLongProcess'),
    (r'(?i)\bStopTask\b',                    'StopTask',           'StopTask entry point in basLongProcess / aeLongProcessClass'),
    (r'(?i)\bResetTask\b',                   'ResetTask',          'ResetTask entry point in basLongProcess / aeLongProcessClass'),
    (r'(?i)\bLog_Init\b',                    'Log_Init',           'Log_Init method on aeLoggerClass'),
    (r'(?i)\bLog_Write\b',                   'Log_Write',          'Log_Write method on aeLoggerClass'),
    (r'(?i)\bLog_Close\b',                   'Log_Close',          'Log_Close method on aeLoggerClass'),
    (r'(?i)\bLog_UnicodeDetail\b',           'Log_UnicodeDetail',  'Log_UnicodeDetail method on aeLoggerClass'),
    (r'(?i)\bSetLogger\b',                   'SetLogger',          'SetLogger method on aeAssertClass'),
    # --- Ribbon keyboard navigation identifiers (added 2026-04-13) ---
    (r'(?i)\bGetPrevBkEnabled\b',            'GetPrevBkEnabled',   'GetPrevBkEnabled method on aeRibbonClass'),
    (r'(?i)\bGetNextBkEnabled\b',            'GetNextBkEnabled',   'GetNextBkEnabled method on aeRibbonClass'),
    (r'(?i)\bKT_BOOK\b',                     'KT_BOOK',            'KT_BOOK keytip constant in basRibbonStrings'),
    (r'(?i)\bKT_CHAPTER\b',                  'KT_CHAPTER',         'KT_CHAPTER keytip constant in basRibbonStrings'),
    (r'(?i)\bKT_VERSE\b',                    'KT_VERSE',           'KT_VERSE keytip constant in basRibbonStrings'),
    (r'(?i)\bKT_PREV_BOOK\b',               'KT_PREV_BOOK',        'KT_PREV_BOOK keytip constant in basRibbonStrings'),
    (r'(?i)\bKT_NEXT_BOOK\b',               'KT_NEXT_BOOK',        'KT_NEXT_BOOK keytip constant in basRibbonStrings'),
    (r'(?i)\bKT_PREV_CHAPTER\b',            'KT_PREV_CHAPTER',     'KT_PREV_CHAPTER keytip constant in basRibbonStrings'),
    (r'(?i)\bKT_NEXT_CHAPTER\b',            'KT_NEXT_CHAPTER',     'KT_NEXT_CHAPTER keytip constant in basRibbonStrings'),
    (r'(?i)\bKT_PREV_VERSE\b',              'KT_PREV_VERSE',       'KT_PREV_VERSE keytip constant in basRibbonStrings'),
    (r'(?i)\bKT_NEXT_VERSE\b',              'KT_NEXT_VERSE',       'KT_NEXT_VERSE keytip constant in basRibbonStrings'),
    (r'(?i)\bKT_NEW_SEARCH\b',              'KT_NEW_SEARCH',       'KT_NEW_SEARCH keytip constant in basRibbonStrings'),
    (r'(?i)\bKT_ABOUT\b',                   'KT_ABOUT',            'KT_ABOUT keytip constant in basRibbonStrings'),
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