# 10 — Word's `List Paragraph` numbering-engine bug

## Symptom

Modifying any paragraph style that inherits from `List Paragraph`
(or any style that holds a `LinkToListTemplate`) in a large Word
document — typically via the **Modify Style** dialog — causes Word
to hang. The status bar shows nothing useful; eventually the title
bar reads **(Not Responding)** and the only recovery is to kill the
application from Task Manager. Unsaved work is lost.

The hang scales with document size. In a 33,857-paragraph
Bible-class document this is reliably reproducible. In a
10-paragraph test document it never appears.

## Project policy

> **No style in the `approved` array may inherit from `List Paragraph`
> or hold a `LinkToListTemplate`. `BaseStyle = ""` is mandatory for
> every approved style.**

This rule is already in the [01-styles QA checklist](01-styles.md);
this page documents the **why**, the **history**, and the **migration
recipe** so the rule isn't silently relaxed by anyone who doesn't
know the cost.

## Root cause read

Word's list engine recomputes every paragraph that references a
`ListTemplate` whenever a style touching list numbering is changed.
The recompute is O(N) in paragraph count and runs synchronously on
the UI thread. In short documents it finishes in milliseconds; in
long documents it hangs until the user kills the process.

Inheritance from `List Paragraph` carries the implicit list-template
reference even on styles you don't think of as "list" styles. The
recompute fires on every property edit — font name, size, indent —
whether or not the edit has anything to do with numbering.

## Status with Microsoft

The bug is well-known and long-lived (observable across Word 2010,
2013, 2016, 2019, 365). It will almost certainly **not be fixed**:

- Backward-compatibility constraints reach back to Word 2003 / .doc
  format support, where the list engine's data model was already in
  place. A fix would be a re-implementation, not a patch.
- The user base hitting the bug is a minority — long-document authors,
  technical writers, academic / theological / legal editors. The
  larger user base writing short Office documents never sees it.
- Microsoft's official guidance, where it exists, is workarounds
  rather than fixes: "create your styles in a blank document and
  import," "don't use the Modify Style dialog," "use VBA."

## Common bad advice (do not follow)

When this bug is reported, automated assistants and search results
typically suggest one of:

- "Your document is corrupt. Save as new and re-import."
- "Your styles are corrupt. Delete and recreate via the dialog."
- "Your `.docx` is too large. Split it."
- "User error — you're modifying the wrong style."
- "Re-register Word / clear template cache / reset normal.dotm."

**None of these address the bug.** They burn hours and produce no
fix. Several actively make things worse (style recreation via the
dialog re-triggers the hang).

The only intervention that reliably works is the structural one:
**detach from `List Paragraph` and stop using `LinkToListTemplate`.**

## Migration recipe (five steps, all VBA, never the dialog)

### Step 0 — Diagnostic

Inventory which approved styles inherit from `List Paragraph` or
carry a list-template link.

```vba
Public Sub AuditListStyleRisk()
    Dim s As Word.Style
    Debug.Print "Name", "BaseStyle", "HasListTemplate"
    For Each s In ActiveDocument.Styles
        If s.Type = wdStyleTypeParagraph Then
            On Error Resume Next
            Debug.Print s.NameLocal, s.BaseStyle, _
                        Not (s.LinkToListTemplate Is Nothing)
            On Error GoTo 0
        End If
    Next s
End Sub
```

Currently expected to flag `ListItem`, `ListItemBody`, `ListItemTab`
in this project. Confirm before proceeding.

### Step 1 — Define replacements in a blank `.docm`

Open a fresh empty Word document. Define each replacement style
there, with `BaseStyle = ""` set **before** any other property —
this is the critical isolation. The blank document has no
list-engine state to recompute against, so the cascade doesn't fire.

> **File extension matters.** Save as `.docm` (macro-enabled
> document), not `.docx` (macro-free). Word strips VBA modules on
> save in `.docx` — the file looks fine, but the style-creation
> macro is silently lost. The macro-enabled distinction is a
> security boundary introduced in Office 2007 so tooling (Outlook,
> antivirus, group policy) can identify executable content from
> the file extension alone. Same logic applies to templates:
> `.dotx` strips macros, `.dotm` keeps them.

```vba
Dim s As Word.Style
Set s = ActiveDocument.Styles.Add("AuthorListItem", wdStyleTypeParagraph)
s.BaseStyle = ""                       ' must come first
s.AutomaticallyUpdate = False
s.QuickStyle = False
s.Font.Name = "Carlito"
s.Font.Size = 11
With s.ParagraphFormat
    .Alignment = wdAlignParagraphLeft
    .LeftIndent = 18                    ' adjust to design intent
    .FirstLineIndent = -18              ' the hanging-indent visual
    .SpaceBefore = 0
    .SpaceAfter = 0
    .LineSpacingRule = wdLineSpaceSingle
End With
' Critically absent: any LinkToListTemplate call.
```

Save the blank doc as a holding file (e.g. `tools/style_holding.docm`).

### Step 2 — Transport to the live document

Two safe options. Pick one. **Never use the Modify Style dialog.**

(a) Organizer:
```vba
ActiveDocument.CopyStylesFromTemplate "C:\path\style_holding.dotm"
```

(b) Direct read-and-write:
```vba
Dim src As Word.Style, dst As Word.Style
Set src = Documents("style_holding.docm").Styles("AuthorListItem")
Set dst = ActiveDocument.Styles.Add("AuthorListItem", wdStyleTypeParagraph)
dst.BaseStyle = ""
dst.AutomaticallyUpdate = src.AutomaticallyUpdate
dst.QuickStyle = src.QuickStyle
dst.Font.Name = src.Font.Name
' ... copy each property explicitly. Don't shortcut via assignment.
```

### Step 3 — Re-apply to existing paragraphs (batched)

```vba
Public Sub MigrateParagraphs(oldName As String, newName As String)
    Application.ScreenUpdating = False
    Dim para As Word.Paragraph
    Dim n As Long
    For Each para In ActiveDocument.Paragraphs
        If para.Style.NameLocal = oldName Then
            para.Style = ActiveDocument.Styles(newName)
            n = n + 1
        End If
    Next para
    Application.ScreenUpdating = True
    Debug.Print n & " paragraphs migrated from " & oldName & " to " & newName
End Sub
```

With `ScreenUpdating = False` and no list-template link on the new
style, the migration completes in seconds even on large documents.
The list-engine recompute that causes the original hang doesn't
fire because the new style isn't part of the list-engine graph.

### Step 4 — Decommission the old styles

Either delete the old style or set `Priority = 99` and remove it
from the `approved` array in `src/basTEST_aeBibleConfig.bas`. Keep
one copy alive briefly for rollback in case any paragraph migration
was missed; remove fully after one clean
[`RUN_TAXONOMY_STYLES`](04-qa-workflow.md) pass.

### Step 5 — Update the taxonomy audit

The renamed style needs its own `AuditOneStyle` entry in
`RUN_TAXONOMY_STYLES`. Encode the descriptive spec from the
replacement definition (Step 1).

## What we lose by detaching

- **Word's auto-numbering** for these styles (no `1. 2. 3.` from the
  list engine). If numbering is needed, render markers in the body
  text or apply a separate dedicated list pattern via VBA.
- **Outline view / Navigation pane numbering** integration weakens
  for these blocks. Heading 1 / 2 are unaffected.
- **Cross-references to numbered items** stop resolving. Anything
  that says "see item 4(b)" depends on the list engine.

## What we don't lose, in this project

- Heading 1 / 2 — already standalone styles, untouched.
- Verse markers, chapter markers — character styles, not affected.
- Footnote handling — independent of the list engine.
- Body text rendering, page layout, headers/footers — all
  unaffected.

For a Bible-class document where list usage is structural
(`ListItem`, `ListItemBody`, `ListItemTab` for the front-matter
list blocks), and where verses are marked rather than
auto-numbered, none of the lost features are in active use.

## Cost-now vs cost-later

| Phase | Total cost | Notes |
|---|---|---|
| Cost-now (one-off refactor) | ~2-4 hours | Bounded; one sitting, one branch |
| Cost-later (defer) | unbounded, recurring | Hang risk on every style edit; compounds with translated/longer docs |

The probability that a refactor is eventually mandatory is 1. The
question is whether it happens on a planned schedule or in an
emergency after a multi-hour lost-work incident.

**Canary indicator.** If the next style modification on a
`List Paragraph`-derived style takes more than 10 seconds via the
dialog, or shows "Not Responding", refactor is overdue. The
canary has already chirped on this project.

## i18n implications

Refactor first, then i18n. Two reasons:

- Translated documents are typically longer (English → German
  averages +30%). Bigger documents hit the hang harder. Translating
  into the bug makes the bug worse.
- Word's `NameLocal` aliasing for `List Paragraph` differs by
  Office UI language. A document built around inheritance can
  behave differently on a Word installation with a different UI
  language. Standalone styles round-trip cleanly across locales.

The cost of localising list markers manually (after detachment) is
small in this project — verses are not auto-numbered, list usage is
structural and finite, and any RTL handling is cleaner in our own
code than in Word's RTL list engine.

## Cross-references

- [01-styles.md](01-styles.md) — QA checklist row `BaseStyle = ""`
  (this page is the *why*).
- [02-editing-process.md](02-editing-process.md) — Stage 1 style
  design must follow this rule for any list-shaped style.
- [04-qa-workflow.md](04-qa-workflow.md) — `RUN_TAXONOMY_STYLES`
  enforces the rule indirectly via the audit; the diagnostic at
  Step 0 above is the direct check.

## History

This page added 2026-04-29. Decision rationale recorded in
`rvw/Code_review 2026-04-25.md` under the dated section
"`List Paragraph` numbering-engine bug — analysis and migration
recipe."
