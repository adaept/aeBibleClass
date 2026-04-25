# 02 — Editing process

Each editing stage maps to one or more routines in
`src/basStyleInspector.bas` (and friends). Run them in order; skip
stages that don't apply to your change.

## Stage 1 — Style design

When introducing a new style or changing an existing one.

1. Pick the style name. Convention: PascalCase, no spaces if possible.
   Position in the [book order](04-qa-workflow.md) determines its
   priority number.
2. Add a `Define<Name>Style` routine to `src/basFixDocxRoutines.bas`
   that creates / updates the style with explicit property values
   (no inheritance — `BaseStyle = ""`).
3. Add the name to the `approved` array in
   `src/basTEST_aeBibleConfig.bas`, in the position the style first
   appears in the document.
4. Add expected property values to `RUN_TAXONOMY_STYLES` so the QA
   suite can audit it.
5. Run `WordEditingConfig` to repromote priorities — every approved
   style gets a fresh priority, everything else falls to 99.

## Stage 2 — Apply the style in the document

Manual. Word UI, Styles pane (Alt+Ctrl+Shift+S), or the dropdown.
Indent values come off the **ruler** (WRIST principle); never guess
or copy from another style without verifying on the ruler.

For paragraph styles, just place the cursor in the paragraph. For
character styles, select the run.

## Stage 3 — Single-style audit

After applying a new or changed style:

```vba
DumpStyleProperties "MyNewStyle"             ' Immediate window only
DumpStyleProperties "MyNewStyle", True       ' also writes rpt\Styles\style_MyNewStyle.txt
```

Output is paste-ready into a `Define<Name>Style` routine. Verify each
of the four QA-checklist properties (see [01-styles](01-styles.md)
§ QA checklist).

## Stage 4 — Bulk audit + orphan cleanup

After multiple style changes, or after a rename:

```vba
WordEditingConfig            ' repromote priorities first
DumpAllApprovedStyles        ' dumps every approved style + orphan check
```

`DumpAllApprovedStyles` writes one file per current approved style and
then scans `rpt/Styles/` for orphans (files left behind by renames).
A single MsgBox prompts to delete them — yes deletes all listed, no
skips.

## Stage 5 — Order verification

Whenever the document content has shifted (new pages, new styles,
reordered sections):

```vba
ListApprovedStylesByBookOrder           ' Immediate only
ListApprovedStylesByBookOrder True      ' also writes rpt\Styles\styles_book_order.txt
```

The output is the canonical order for the `approved` array. If the
two disagree, the array is wrong — reorder the array, re-run
`WordEditingConfig`, re-run the order check until they match. Detail
in [04-qa-workflow](04-qa-workflow.md).

## Stage 6 — Header / footer changes

When adding or modifying header / footer content:

```vba
DumpHeaderFooterStyles      ' writes rpt\Styles\header_footer_audit.txt
```

Walks every section x every header/footer slot (Even, Primary,
FirstPage). The audit file shows which sections own unlinked
headers/footers vs which are Linked-to-Previous. Detail in
[05-headers-footers](05-headers-footers.md).

## Stage 7 — Pre-commit gate

`SUPER_TEST_RUNS` (when implemented; see
[07-super-test-runs](07-super-test-runs.md)). The single command that
runs every QA suite and produces a master pass/fail report. Currently
deferred until the taxonomy is stable.

## Anti-patterns / gotchas

- **Don't apply direct formatting** to paragraphs that use a style
  with `AutomaticallyUpdate = True`. The style definition gets
  silently rewritten for every other paragraph using it. The QA
  checklist enforces `AutomaticallyUpdate = False`; verify before
  hand-editing.
- **Don't reorder `approved` manually** without immediately re-running
  `ListApprovedStylesByBookOrder` to verify alignment with book order.
- **Don't delete `rpt/Styles/` files by hand**. Let
  `DumpAllApprovedStyles`'s orphan-cleanup prompt do it; it knows
  exactly which files are stale.
- **Don't skip `WordEditingConfig`** before `DumpAllApprovedStyles` if
  you've changed the array — stale priorities will show up as
  unexpected entries.
- **Don't add VBA project references** to support a style operation.
  Late binding only — `As Object` + `CreateObject`.

## Per-routine quick reference

| Routine | Output | When to run |
|---|---|---|
| `DumpStyleProperties "X"` | Immediate window | After applying / changing style X |
| `DumpStyleProperties "X", True` | + `rpt/Styles/style_X.txt` | When you want a versioned snapshot |
| `DumpAllApprovedStyles` | All `style_*.txt` + orphan cleanup | After bulk style work or renames |
| `ListApprovedStylesByBookOrder` | Immediate window | Whenever content order may have shifted |
| `ListApprovedStylesByBookOrder True` | + `rpt/Styles/styles_book_order.txt` | When you want a committed snapshot |
| `DumpHeaderFooterStyles` | `rpt/Styles/header_footer_audit.txt` | After header / footer changes |
| `WordEditingConfig` | Immediate window | Before any audit; resets priorities |
| `StartTimer` / `EndTimer` | Immediate window | Bracket long-running custom work |

Full signatures and examples in [03-inspection-tools](03-inspection-tools.md).
