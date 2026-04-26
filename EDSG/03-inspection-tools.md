# 03 — Inspection tools

Reference for `src/basStyleInspector.bas`. Five public entry points,
several private helpers. All routines are read-only on the document
unless explicitly noted (the orphan-cleanup prompt deletes files in
`rpt/Styles/` after user confirmation).

## `DumpStyleProperties`

Print one named style's properties in paste-ready VBA form.

```vba
Public Sub DumpStyleProperties(ByVal sStyleName As String, _
                               Optional ByVal bWriteFile As Boolean = False)
```

- `sStyleName` — style name, exact case. Errors with a clear message
  if not found.
- `bWriteFile` — `True` also writes `rpt\Styles\style_<name>.txt`
  (spaces and slashes in the name are replaced with underscores).

Output covers `BaseStyle`, `QuickStyle`, `Font.*`, and (paragraph
styles only) `NextParagraphStyle`, `AutomaticallyUpdate`, every
`ParagraphFormat.*`. Character styles get only the universally-valid
properties — accessing `AutomaticallyUpdate` on a character style
raises run-time error 5900.

```
DumpStyleProperties "FrontPageBodyText", True
```

Sample output (raw values; symbolic decoding is manual — see the
"Possible refinements" note in `rvw/Code_review 2026-04-21.md`):

```
'--- FrontPageBodyText  (Type=Paragraph, Priority=8) ---
.BaseStyle = ""
.QuickStyle = False
.Font.Name = "Liberation Serif"
.Font.Size = 11
...
.ParagraphFormat.Alignment = 1
.ParagraphFormat.LineSpacingRule = 0
...
```

## `DumpAllApprovedStyles`

Dump every approved (Priority ≠ 99) paragraph or character style.
Includes orphan-file cleanup.

```vba
Public Sub DumpAllApprovedStyles()
```

Steps:

1. Collect every paragraph / character style with `Priority <> 99`.
2. Sort by priority ascending.
3. For each, call `DumpStyleProperties name, True`. Errors are caught
   per-style and logged as `!! FAILED`; the batch continues.
4. Scan `rpt/Styles/style_*.txt` and identify orphans (files with no
   corresponding current approved style — typically left by a
   rename like `ContentsCPBB` → `Contents`).
5. List orphans to Immediate; single MsgBox prompts to delete them
   all or skip.

Source of truth is the **live document** (priorities set by
`PromoteApprovedStyles`), not a duplicate list in this module.

## `ListApprovedStylesByBookOrder`

For each approved style, find the page of its FIRST occurrence
anywhere in the document (main body, headers, footers, footnotes,
endnotes), and list them in `(Page, Priority)` ascending order.

```vba
Public Sub ListApprovedStylesByBookOrder( _
                Optional ByVal bWriteFile As Boolean = False)
```

- `bWriteFile = True` also writes
  `rpt\Styles\styles_book_order.txt`.

Search strategy:

- **Main body, footnotes, endnotes, etc.** — `For Each StoryRanges`
  walk with `Range.Find` and `Style` filter.
  Header/footer story types (6–11) are **skipped** here; the
  Sections walk handles them deterministically.
- **Headers / footers** — explicit `For Each oSection / Headers(1..3)
  / Footers(1..3)` walk, using paragraph iteration (not Find) because
  Find is unreliable for tab-only header content.
- **Page-1 fallback** — when a paragraph match is found in a header
  / footer story (Word's `Information(wdActiveEndPageNumber)` returns
  a misleading section-anchor page for headers), the routine returns
  page 1 — accurate for headers that tile from the start of the
  document.

Output:

```
Approved styles in book order (by page of first occurrence)
 Page | Prio | Style
------+------+-----------------------------
    1 |    1 | TheHeaders
    1 |    2 | BodyText
    1 |    3 | TheFooters
    2 |    4 | FrontPageTopLine
    ...
    - |   N  | <unused style>  [not used]
```

The output IS the canonical priority order for the `approved` array.
Not a hint, not a suggestion — see [04-qa-workflow](04-qa-workflow.md).

## `DumpHeaderFooterStyles`

Read-only diagnostic. Walks every section x header/footer slot and
records:

- Section number
- Slot kind (`Header(1)` / `Header(2)` / `Header(3)` /
  `Footer(1)` / `Footer(2)` / `Footer(3)`)
- `LinkToPrevious` flag (shown as `linked`)
- Paragraph count
- First paragraph's `Style.NameLocal`
- First 50 chars of text (tabs shown as `<tab>`)

Output to `rpt\Styles\header_footer_audit.txt`. With 148 sections
that's ~888 lines — too long for the Immediate window, which only
shows a summary.

```vba
Public Sub DumpHeaderFooterStyles()
```

Use to debug when a header/footer style appears `[not used]` in
`ListApprovedStylesByBookOrder` output. Detail in
[05-headers-footers](05-headers-footers.md).

## `StartTimer` / `EndTimer`

Bracket long-running routines for "expected (last run) vs actual"
feedback in the Immediate window.

```vba
Public Sub StartTimer(ByVal sName As String, ByRef startTime As Double)
Public Sub EndTimer  (ByVal sName As String, ByVal startTime As Double)
```

Storage is a module-level `Scripting.Dictionary` — session-scoped.
First run after Word restart prints "first run this session, no prior
timing"; subsequent runs print "expected ~X.XX sec (last run)".

Wired into `DumpAllApprovedStyles`, `ListApprovedStylesByBookOrder`,
and `DumpHeaderFooterStyles`. To time custom code:

```vba
Dim t As Double
StartTimer "MyRoutine", t
' ... work ...
EndTimer "MyRoutine", t
```

## Private helpers (informational)

Not callable from outside the module but useful to know:

- `FindStylePage(oRng, oStyle, ByRef bFoundAnywhere)` — Find within a
  range; returns page or -1; sets `bFoundAnywhere` on success.
- `FirstPageInParagraphs(oRng, oStyle, ByRef bFoundAnywhere)` —
  paragraph-iteration alternative for header/footer ranges.
- `FirstPageForStyle(oDoc, oStyle)` — orchestrates the StoryRanges +
  Sections walks; applies the page-1 fallback.
- `WriteStyleDump`, `WriteBookOrderFile`, `WriteHeaderFooterAuditFile`
  — file output via `Scripting.FileSystemObject`, ASCII encoding.
- `SafeFileName(s)` — replaces spaces and slashes with underscores
  for filesystem-safe basenames.
- `CleanupOrphanStyleDumps(arr, nCount)` — orphan detection and
  deletion prompt; called from `DumpAllApprovedStyles`.

## Module conventions

- Late binding throughout — every COM access is `As Object` +
  `CreateObject`. No project references.
- ASCII hyphens only in comments — no em-dashes.
- `Option Private Module` — public subs callable from within the
  project / Immediate window only.
- Identifier casing preserved exactly — the git commit normalizer
  depends on it.
