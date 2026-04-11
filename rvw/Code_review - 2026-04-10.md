# Code Review - 2026-04-10
## 3-Button Navigation Ribbon: GoTo Book, GoTo Chapter, GoTo Verse

---

## § 1 — Project Development Overview (Expert Review)

The following is a chronological synthesis of the `rvw/` review files. It provides
context for anyone starting work on this codebase.

---

### 2026-03-15 to 2026-03-17 — Foundation: aeBibleClass + Initial Ribbon

The project began as a QA and navigation tool for the Radiant Word Bible (RWB), a
full-Bible `.docm` in Microsoft Word 365. The early review files establish:

- `aeBibleClass.cls` — the core QA harness running 72+ style and content tests
- `aeRibbonClass.cls` — the ribbon controller (singleton pattern via
  `basBibleRibbonSetup.bas`)
- The `headingData(1 To 66, 0 To 1)` array: captured once at ribbon load, stores
  Heading 1 (book) text and character position for all 66 books
- First implementation of GoToH1 (GoTo Book), Next Book, and Prev Book ribbon buttons
- Casing normalizer `py/normalize_vba.py` established to prevent VBA IDE from
  lowercasing exported identifiers

### 2026-03-23 to 2026-03-27 — Citation Parser: aeBibleCitationClass

- `aeBibleCitationClass.cls` — a Deterministic Structural Parser (DSP) for SBL-style
  Bible references (`John 3:16`, `1 John 1:1-3`, `Rev 22`)
- Private `ChaptersInBook(bookName)` and `VersesInChapter(bookName, chapter)` — look
  up valid chapter/verse ranges from the packed `GetChapterVerseMap()` data
- `aeSBL_Citation_Class` interface and test harness `basTEST_aeBibleCitationClass.bas`
- **Key constraint:** citation class methods are Private — not callable from ribbon

### 2026-04-01 to 2026-04-04 — SBL Tests + GoToH1 Performance Problem

- `Run_All_SBL_Tests` wired and passing
- GoToH1 first identified as causing a 37-second double-block when navigating to
  far books (Genesis → Revelation). Root cause: Word's layout engine synchronously
  builds the page layout table for the full document on each navigation to an unvisited
  position.

### 2026-04-06 — GoToH1 Fix Attempts (28 iterations)

- Extensive investigation of the double-block. 28 separate approaches tried.
- None fully successful within the single session.

### 2026-04-08 — GoToH1 RESOLVED

- Root cause confirmed: `InvalidateControl` after `ScreenUpdating = True` triggered a
  second layout pass.
- Final fix (Option B): defer GoToH1 via `Application.OnTime` (1-second delay) to
  escape the ribbon callback post-processing; pre-warm the layout cache at startup
  via `WarmLayoutCacheDeferred` (scheduled 5 seconds after ribbon load).
- Both blocks eliminated in testing.

### 2026-04-09 — Long-Running Process Framework

- `IaeLongProcessClass.cls` — interface contract for parameterizable tasks
- `aeLongProcessClass.cls` — runner: batch loop, DoEvents, progress sidecar, logging
- `aeUpdateCharStyleClass.cls` — first task implementation (re-applies character style)
- `aeLoggerClass.cls` — converted from `basLogger.bas`; UTF-8 string-buffer approach
- `basLongProcess.bas` — thin public skeleton (renamed from `XLongRunningProcessCode.bas`)
- Several bugs found and fixed during live testing: Escape key unreliable in Word VBA,
  `Options.Pagination` fails when Find dialog open (runner must be state-neutral),
  IDE Stop resets singleton (lazy re-capture guard added to GoToH1)
- Normalizer updated with 15 new framework identifiers

---

## § 2 — Current Ribbon State

Three buttons are currently defined in the ribbon:

| Button | Callback | Status |
|--------|----------|--------|
| GoTo Book | `OnGoToH1ButtonClick` → `GoToH1` | Working |
| GoTo Next | `OnNextButtonClick` → `NextButton` | Working |
| GoTo Prev | `OnPrevButtonClick` → `PrevButton` | Working |

The GoTo Verse button (`OnGoToVerseSblClick` → `GoToVerseSBL`) is a stub:
```vba
Private Sub GoToVerseSBL()
    MsgBox "GoToVerseSBL - Parser not yet implemented."
End Sub
```

---

## § 3 — Plan: 3-Button Navigation Ribbon

### Goal

Replace the current navigation ribbon with three primary action buttons:

| Button | Action |
|--------|--------|
| **GoTo Book** | Navigate to a Bible book (Heading 1) — extends current GoToH1 |
| **GoTo Chapter** | Navigate to a chapter within the current/selected book (Heading 2) |
| **GoTo Verse** | Navigate to a verse within the current chapter |

Next / Prev buttons remain unchanged.

---

## § 4 — GoTo Book (Extend Existing)

### Current behaviour

`GoToH1` accepts an abbreviation and matches it against all 66 entries in
`headingData` using `Like "*" & UCase(pattern) & "*"`. The user enters `g` and
finds `GENESIS`; `Rev` finds `REVELATION OF JOHN` (or however it appears in the
document).

### Issue: Numbered books not found by number+letter abbreviation

Current input `"1j"` does NOT find `"1 JOHN"` because the `Like` pattern compares
`"*1J*"` against `"1 JOHN"`. The `"1"` and `"J"` are adjacent in the input but
separated by a space in the heading text.

### Fix: Input normalization

Before the `Like` match, apply a normalizer to the user input:

```vba
' Normalize: insert a space between a leading digit and the first letter
' "1j" -> "1 J", "2CO" -> "2 CO", "3Jn" -> "3 JN"
Private Function NormalizeBookInput(ByVal raw As String) As String
    Dim s As String
    s = Trim(UCase(raw))
    If Len(s) >= 2 Then
        If s Like "[0-9][A-Z]*" Then
            s = Left$(s, 1) & " " & Mid$(s, 2)
        End If
    End If
    NormalizeBookInput = s
End Function
```

Replace the `Like` pattern in `GoToH1`:

```vba
' Before:
If CStr(headingData(i, 0)) Like "*" & UCase(pattern) & "*" Then

' After:
If CStr(headingData(i, 0)) Like "*" & NormalizeBookInput(pattern) & "*" Then
```

**Examples after normalization:**

| Input | Normalized | Matches heading |
|-------|-----------|-----------------|
| `g` | `G` | `GENESIS` |
| `Rev` | `REV` | `REVELATION...` |
| `1j` | `1 J` | `1 JOHN` |
| `2Co` | `2 CO` | `2 CORINTHIANS` |
| `3Jn` | `3 JN` | `3 JOHN` |
| `1John` | `1 JOHN` | `1 JOHN` |

**NOTE:** The `NormalizeBookInput` function must be kept in `aeRibbonClass.cls`
(or a dedicated input module), completely separate from the SBL citation parser
in `aeBibleCitationClass.cls`. These are two different concerns:
- SBL parser: parse a full citation string like `"1Jn 3:16"` into structured components
- Ribbon input: normalize an abbreviation typed by a user for interactive search

Mixing them would couple the UI layer to the parser internals.

### State tracking after GoTo Book

`GoToH1` should record which book was selected (book index `i` and `foundPos`) in
instance variables so GoTo Chapter can use them without re-searching:

```vba
Private m_currentBookIndex As Long   ' 1..66 — set by GoToH1, GoTo Chapter
Private m_currentBookPos   As Long   ' character position of current book H1
```

---

## § 5 — GoTo Chapter (New)

### Document structure

The document uses **Heading 2** for chapter markers. This is confirmed by
`basAddHeaderFooter.bas` which walks sections and distinguishes `Heading 1`
(book title) from `Heading 2` (first chapter).

### Design

GoTo Chapter proceeds as follows:

1. **Determine current book.** Use `m_currentBookIndex` / `m_currentBookPos` set by
   GoTo Book. If not set (user jumped directly to GoTo Chapter without using GoTo Book
   first), infer from the current selection position: find the nearest preceding Heading 1.

2. **Determine chapter range.** The book name is known from `headingData(m_currentBookIndex, 0)`.
   The number of chapters is not directly available from the heading array — it must be
   looked up. Two options:

   **Option A — Capture Heading 2 positions at load time.**
   Extend `CaptureHeading1s` (or add a separate `CaptureHeading2s`) to capture all
   1189 Heading 2 paragraphs into a chapter array:
   ```vba
   Private chapterData(1 To 66, 1 To 150) As Long  ' chapterData(bookIdx, chNum) = charPos
   ```
   This array stores the character position of each chapter heading. The chapter range
   for a book is then implicit in which entries are non-zero.

   **Option B — Count Heading 2s at runtime by Find.**
   After navigating to the book's H1 position, count forward Heading 2 finds up to
   the next Heading 1. Slower but requires no pre-capture array.

   **Recommendation: Option A.** The capture runs once per session (same pattern as
   `CaptureHeading1s`). The array size 66 × 150 × 8 bytes ≈ 79 KB — acceptable.
   Max chapters in Bible = 150 (Psalms). Runtime look-up is then O(1).

3. **Prompt user.** InputBox: `"Enter chapter number (1-N):"`. Pre-fill `N` with the
   chapter count for the current book.

4. **Validate.** Chapter must be in range 1..N. Show error if out of range.

5. **Navigate.** Use the position from `chapterData(bookIdx, chNum)`.
   Update `m_btnPrevEnabled` / `m_btnNextEnabled` based on whether prev/next chapters
   exist.

6. **Record state.** Set `m_currentChapter = chNum` for use by GoTo Verse.

### Chapter heading format

The exact text of Heading 2 paragraphs needs to be confirmed. Likely formats:
- `"Chapter 1"`, `"Chapter 2"`, ... — if so, position-based navigation via
  `chapterData` array is correct (we navigate by position, not by matching text).
- A plain number: `"1"`, `"2"`, ... — same approach applies.

In either case, GoTo Chapter navigates **by array position** (book index + chapter
number), not by text matching. This is faster and unambiguous.

---

## § 6 — GoTo Verse (New)

### Document structure

Verse paragraphs are body text (not a heading style). Each verse is one paragraph
(the document was formatted with `OneVersePerPara = True` for the main branch).
The `"Chapter Verse marker"` character style marks verse reference numbers within
each paragraph.

### Design

GoTo Verse proceeds as follows:

1. **Determine current book and chapter.** Use `m_currentBookIndex` and
   `m_currentChapter`. Both should be set by a prior GoTo Book + GoTo Chapter
   sequence. If not set, infer from the current selection.

2. **Determine verse range.** The number of verses is available from
   `aeBibleCitationClass.VersesInChapter` — but that method is `Private`.

   **Required change:** Make `ChaptersInBook` and `VersesInChapter` Public in
   `aeBibleCitationClass.cls`. These are pure data look-ups with no side effects;
   there is no reason for them to be Private.

   **NOTE:** Use these only for range validation (max verse number). Do not use
   `aeBibleCitationClass` for the navigation itself — navigation uses document
   positions, not citation parsing.

3. **Prompt user.** InputBox: `"Enter verse number (1-N):"`. Pre-fill `N` with the
   verse count for the current book and chapter.

4. **Validate.** Verse must be in range 1..N.

5. **Navigate.** This is the hardest step. Options:

   **Option A — Capture verse positions at load time.**
   A `verseData` array would store character positions for all 31,102 verses. At
   8 bytes each ≈ 249 KB. Feasible but the capture scan would take several seconds.

   **Option B — Find forward from chapter start at runtime.**
   From `chapterData(bookIdx, chNum)`, use `Selection.Find` (style = Heading 2,
   forward = True, count = 1 to confirm chapter start), then count forward N body
   paragraphs to reach verse N.

   **Option C — Use the verse reference marker.**
   The `"Chapter Verse marker"` character style marks verse numbers in the text.
   A targeted `Find` looking for the verse number string with that character style
   would navigate directly to the verse. This is fragile if verse markers are
   inconsistently formatted.

   **Recommendation: Defer Option A/C selection until document structure is confirmed.**
   Option B is the lowest-risk starting point and avoids a large pre-capture scan.
   If performance is acceptable (< 1 second for most navigations), Option B is
   sufficient.

---

## § 7 — Combo-Box InputBox (Drop-Down Suggestion)

### Current InputBox limitation

VBA's built-in `InputBox` is a plain text entry field. It does not support drop-down
lists, autocomplete, or dynamic suggestion filtering.

### Options

**Option A — UserForm with ComboBox.**
Create a `.frm` (UserForm) with a ComboBox control populated with all 66 book names.
The ComboBox filters matching entries as the user types. This is the correct solution
for a polished UX.

**Downsides:**
- Requires creating and maintaining a `.frm` file
- The UserForm must be imported alongside the `.cls`/`.bas` files
- More complex to test

**Option B — InputBox with sorted hint in prompt.**
Keep the basic `InputBox` but add a hint string listing all matching suggestions
dynamically. This is clunky but zero new UI infrastructure.

**Option C — Accept basic InputBox for now; plan UserForm for later.**
The current `Like "*pattern*"` matching plus the `"1j" → "1 J"` normalization gives
a usable search experience. A UserForm is a separate feature request and can be
approved and scoped independently.

**Recommendation: Option C for this plan.** The normalization fix and the new
GoTo Chapter / GoTo Verse InputBoxes are deliverable without a UserForm.
Record the UserForm as a future enhancement requiring explicit approval before work
begins.

---

## § 8 — OLD_CODE Module

### Approach

Any code that is superseded but not yet ready to delete should be moved to a module
named `basOLD_CODE.bas` (or `OLD_CODE` as a module name in the VBA project). This is
a staging area under user discretion.

### Current candidates

| Code | Location | Reason |
|------|----------|--------|
| `UpdateCharacterStyle` legacy stub | `basLongProcess.bas` | Superseded by `aeUpdateCharStyleClass` |
| `GoToVerseSBL` stub | `aeRibbonClass.cls` | Will be replaced by GoTo Verse implementation |
| `basBibleRibbon_OLD.bas` | Already exists | Contains prior ribbon setup code |

### Rule

Do NOT delete any code in this plan — move it to `basOLD_CODE.bas` with a comment
noting what it was superseded by and when. The user decides when to permanently delete.

---

## § 9 — Additional Issues Identified

### Issue 1 — `CaptureHeading1s` uses `Static hasRun` — blocks refresh

`CaptureHeading1s` uses a `Static hasRun As Boolean` flag to ensure it runs only
once per session. This is correct for performance but means the heading data is never
refreshed if the document is edited (headings added, removed, or repositioned) within
a session. If GoTo Chapter requires a parallel `CaptureHeading2s`, the same pattern
applies. For a read-only Bible navigation document this is acceptable; for an editing
workflow it would be a limitation.

### Issue 2 — `m_currentBookIndex` not set when user navigates manually

If a user scrolls to Revelation manually (not via GoTo Book) and then clicks GoTo
Chapter, `m_currentBookIndex` is 0. The design must handle this gracefully:
infer the current book from `Selection.Range.Start` by scanning `headingData`
for the nearest preceding book position.

### Issue 3 — Heading 2 chapter data requires 1189-entry capture

1189 Heading 2 paragraphs in the full Bible. The capture scan iterates all document
paragraphs (~31,000+). This will take a few seconds if done at ribbon load. Consider
whether to add it to `EnableButtonsRoutine` (runs at load alongside `CaptureHeading1s`)
or defer to first GoTo Chapter call (lazy pattern).

Recommendation: defer to first GoTo Chapter call (lazy). The scan cost is paid once,
only when the feature is first used.

### Issue 4 — `aeBibleCitationClass.ChaptersInBook` and `VersesInChapter` are Private

These are pure data look-ups that the ribbon needs for range validation. They should
be made Public. No behavioural changes are needed — only access modifier change.
This is a one-line change per function.

### Issue 5 — Next/Prev buttons operate on books only

After GoTo Chapter navigates to a chapter, pressing Next navigates to the next
**book** (Heading 1 forward search), not the next chapter. This may confuse users who
expect Next/Prev to follow the most recent navigation context. This is a pre-existing
limitation to note; fixing it is a separate plan item.

---

## § 10 — Implementation Steps (in order)

These steps are proposed. No code will be written until each step is approved.

### Step 1 — Input normalization for GoTo Book

Add `NormalizeBookInput` private function to `aeRibbonClass.cls`. Apply it in
`GoToH1` before the `Like` comparison. No ribbon XML changes needed.

Adds state variables: `m_currentBookIndex As Long`, `m_currentBookPos As Long`.
Set both in `GoToH1` when a match is found.

### Step 2 — CaptureHeading2s (lazy, first-call)

Add `Private chapterData(1 To 66, 1 To 150) As Long` to `aeRibbonClass.cls`.
Add `Private Sub CaptureHeading2s()` with `Static hasRun As Boolean` guard.
Call from `GoToChapter` on first use (lazy).

### Step 3 — GoToChapter implementation

Add `Private Sub GoToChapter()` to `aeRibbonClass.cls`.
Add ribbon callback `Public Sub OnGoToChapterButtonClick(control As IRibbonControl)`.

### Step 4 — Expose ChaptersInBook and VersesInChapter as Public

Change access modifier in `aeBibleCitationClass.cls`.

### Step 5 — GoToVerse implementation

Add `Private Sub GoToVerse()` to `aeRibbonClass.cls`.
Navigation via Option B (Find forward from chapter position, count paragraphs).

### Step 6 — Ribbon XML update

Add the two new buttons to the `.docm` ribbon XML. Update `basBibleRibbonSetup.bas`
with the new callback stubs. This requires manual editing of the `.docm`'s Custom UI
XML (via the Custom UI Editor or direct XML edit).

### Step 7 — Move OLD_CODE

Move `UpdateCharacterStyle` stub and `GoToVerseSBL` stub to `basOLD_CODE.bas`.

### Step 8 — normalize_vba.py update

Add new identifiers: `NormalizeBookInput`, `CaptureHeading2s`, `GoToChapter`,
`GoToVerse`, `ChaptersInBook`, `VersesInChapter`, `chapterData`.

---

## § 11 — Open Questions for User Decision

1. **GoTo Chapter navigation method:** Option A (pre-capture `chapterData` array,
   lazy) is recommended. Confirm.

2. **GoTo Verse navigation method:** Option B (Find forward from chapter, count
   paragraphs) is recommended as starting point. Confirm.

3. **Combo-box UserForm:** Deferred (Option C). Confirm.

4. **Next/Prev after chapter navigation:** Pre-existing limitation — do not fix in
   this plan unless explicitly requested.

5. **Heading 2 text format:** What is the actual text of chapter headings in the
   document? (e.g. `"Chapter 1"`, `"1"`, `"1 GENESIS 1"`, etc.) This affects how
   `CaptureHeading2s` should verify it is reading the right paragraphs.

6. **Ribbon XML edit process:** Confirm the current process for updating the `.docm`
   ribbon XML (Custom UI Editor tool, direct XML, or other).

---

## § 13 — Sample Document: JUDE.docx (2026-04-10)

A minimal single-book sample file (`JUDE.docx`) will be created as a reference for
items that need clarification before implementation begins. Additional content can
be added to the sample as needed to support further design decisions.

**Why Jude:** single chapter, 25 verses — small enough to keep the file minimal but
complete enough to exercise all three navigation levels (book, chapter, verse).

**Open questions the sample will answer (§ 11):**

- Q5: Exact text format of Heading 2 chapter headings
  (e.g. `"Chapter 1"`, `"1"`, `"1 JUDE 1"`, etc.)
- Verse paragraph structure and how `"Chapter Verse marker"` character style
  appears in context
- Section break pattern: Heading 1 → Heading 2 → body paragraphs
- Any other structural details needed to finalize the GoTo Chapter / GoTo Verse design

**Status:** Reviewed 2026-04-10. Findings in § 14 and § 15.

---

## § 14 — JUDE.docx Document Structure Findings (2026-04-10)

### Heading styles

| Style | Example text | Purpose |
|-------|-------------|---------|
| `Heading1` | `JUDE`, `REVELATION` | Book title |
| `CustomParaAfterH1` | `THE GENERAL EPISTLE OF JUDE` | Book subtitle |
| `Heading2` | `CHAPTER 1`, `CHAPTER 2` | Chapter heading |
| `Brief`, `DatAuthRef` | background text | Introductory matter between H1 and H2 |

**Answer to § 11, Question 5:** Heading 2 text is `"CHAPTER N"` (e.g. `"CHAPTER 1"`).
`CaptureHeading2s` can verify a Heading 2 is a chapter by checking
`UCase(Left(text, 7)) = "CHAPTER"`.

### Verse paragraph run structure

Every verse paragraph has style `(Normal)` and its runs always open with exactly
two marker runs, in this fixed order:

```
Run 1:  ChapterVersemarker  — chapter number digit(s), e.g. "1"
Run 2:  Versemarker         — verse number digit(s) + narrow no-break space U+202F, e.g. "1 "
Run 3+: body text           — default / EmphasisBlack / WordsofJesus / etc.
```

Examples from the file (paragraph → runs):

```
P178  Run1 ChapterVersemarker "1"   Run2 Versemarker "1 "   Run3 default "Jude, a servant..."
P179  Run1 ChapterVersemarker "1"   Run2 Versemarker "2 "   Run3 default "May mercy, peace..."
P181  Run1 ChapterVersemarker "1"   Run2 Versemarker "4 "   Run3 EmphasisBlack "For there are..."
```

**VBA style name mapping** (XML attribute → VBA `rng.style`):
- `ChapterVersemarker` → `"Chapter Verse marker"`
- `Versemarker` → `"Verse marker"`

### Non-verse paragraphs

Introductory paragraphs (`Brief`, `DatAuthRef`, `PlainText`, `ListParagraph`, etc.)
carry no marker runs at all. A character scan of these paragraphs exits immediately
on the first character — confirming the short-circuit optimization is safe for the
entire document.

---

## § 15 — Optimization Task: aeUpdateCharStyleClass (2026-04-10)

### Current behaviour (inefficient)

`IaeLongProcessClass_ExecuteItem` scans **every character** in every paragraph
looking for `"Chapter Verse marker"`. For a verse paragraph of 100+ characters, this
means 100+ style comparisons to find and re-apply 2–4 marker characters. Across
33,857 paragraphs the total character checks are in the millions.

Only one style (`StyleName = "Chapter Verse marker"`) is currently processed.
`"Verse marker"` is not refreshed.

### Optimization: short-circuit loop + dual-style pass

The marker runs are **always the first two runs** of every verse paragraph. Once the
loop encounters any character that carries neither marker style, all remaining
characters in that paragraph are body text — no further scanning is needed.

**Proposed change to `IaeLongProcessClass_ExecuteItem`:**

```vba
Private Function IaeLongProcessClass_ExecuteItem(itemIndex As Long) As Boolean
    On Error GoTo PROC_ERR
    Dim para As Word.Paragraph
    Dim rng  As Word.Range
    Set para = ActiveDocument.Paragraphs(itemIndex)
    For Each rng In para.Range.Characters
        If rng.style = "Chapter Verse marker" Then
            rng.style = "Chapter Verse marker"
        ElseIf rng.style = "Verse marker" Then
            rng.style = "Verse marker"
        Else
            Exit For   ' no more marker characters in this paragraph
        End If
    Next rng
    IaeLongProcessClass_ExecuteItem = True
PROC_EXIT:
    Exit Function
PROC_ERR:
    IaeLongProcessClass_ExecuteItem = False
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ExecuteItem (item " & itemIndex & ") of Class aeUpdateCharStyleClass"
    Resume PROC_EXIT
End Function
```

### Changes from current code

1. **Short-circuit:** `Exit For` as soon as a non-marker character is found.
   Verse paragraphs: 2–4 character checks instead of 100+.
   Non-verse paragraphs: exit on first character check.

2. **Dual style:** both `"Chapter Verse marker"` and `"Verse marker"` refreshed in
   the same single pass. The `StyleName` public property becomes redundant for this
   specific task — remove or keep as a legacy parameter.

3. **No `StyleName` property needed:** both style names are now hardcoded in
   `ExecuteItem`. If generalization is wanted later, a second `StyleName2` property
   can be added, but for the current use case this is unnecessary complexity.

### Expected performance improvement

| Scenario | Characters checked per paragraph | Total for 33,857 paragraphs |
|----------|----------------------------------|------------------------------|
| Current | ~100 (full scan) | ~3,400,000 |
| Optimized | ~3 (2 markers + 1 exit check) | ~102,000 |

Approximately **33x fewer character comparisons**. The batch pause time (678 × 1
second = ~11 minutes) dominates the total run time; reducing character work
compresses the active processing time within each batch significantly. Overall
elapsed time will be shorter but the pause structure is unchanged unless `PauseMs`
is also reduced.

### Additional note: `StyleName` property

The `StyleName` public property on `aeUpdateCharStyleClass` was designed for
generalization. With the dual-style short-circuit approach it is no longer used
in `ExecuteItem`. Options:

- **Keep it** as a public property for potential future single-style tasks that
  use the same class.
- **Remove it** since the class is now specific to the two Bible marker styles.

Recommendation: keep it for now, add a comment that `ExecuteItem` ignores it
in favour of the hardcoded dual-style logic. Revisit if a second task class is
needed.

**ACCEPTED 2026-04-10.** Implementation follows.

---

## § 12 — Long-Running Process: Step-Through of Main Code Events (2026-04-10)

### Completion result

```
UpdateCharacterStyle progress: 99.39%
...
UpdateCharacterStyle progress: 100.00%
UpdateCharacterStyle: complete - 33857 Items processed
```

### The Three Layers

```
basLongProcess.bas           -- public entry points (Immediate Window)
aeLongProcessClass.cls       -- runner: batching, DoEvents, progress, logging
aeUpdateCharStyleClass.cls   -- task: what to do with each item
```

---

### Step-Through

**1. `TestUpdateCharStyle` typed in Immediate Window**

`basLongProcess.bas:25`
```vba
Public Sub TestUpdateCharStyle()
    Dim t As New aeUpdateCharStyleClass
    StartOrResume t
End Sub
```
Creates a task object `t`. Calls `StartOrResume`.

**2. StartOrResume ensures a runner exists**

`basLongProcess.bas:38`
```vba
Public Sub StartOrResume(task As IaeLongProcessClass)
    If s_runner Is Nothing Then Set s_runner = New aeLongProcessClass
    s_runner.Run task
End Sub
```
`s_runner` is a module-level variable — it persists across calls in the same session.
If this is a resume, the same runner instance is reused. Then `Run` is called.

**3. Run loads previous progress**

`aeLongProcessClass.cls:46`
```vba
m_continueProcessing = True
LoadProgress task.TaskName
```
`LoadProgress` looks for `rpt/LongRunningProgress_UpdateCharacterStyle.txt`. If found
(a previous partial run), it reads `LastProcessedItem` and sets `m_lastProcessedItem`.
If the file does not exist, `m_lastProcessedItem` stays 0.

**4. Run opens the log and asks the task for the item count**

```vba
log.Log_Init ActiveDocument.Path & "\rpt\LongProcess_UpdateCharacterStyle.txt"
totalItems = task.ItemCount   ' -> IaeLongProcessClass_ItemCount
```
`ItemCount` returns `ActiveDocument.Paragraphs.Count` — the 33,857 shown in the
completion message.

**5. Run determines the starting paragraph**

```vba
If m_lastProcessedItem = 0 Then m_lastProcessedItem = 1
```
Fresh run: start at paragraph 1. Resume (e.g. from 14.6% after an IDE Stop): start
at the saved paragraph index (4901).

**6. The batch loop**

`aeLongProcessClass.cls:77`
```vba
For startIndex = m_lastProcessedItem To totalItems Step BatchSize  ' BatchSize = 50
    endIndex = startIndex + BatchSize - 1
    If endIndex > totalItems Then endIndex = totalItems

    For i = startIndex To endIndex
        If Not m_continueProcessing Then  '...save and exit...
        If Not task.ExecuteItem(i) Then   '...save and exit...
        DoEvents
    Next i

    ' after each batch of 50:
    SaveProgress task.TaskName
    PauseWithDoEvents PauseMs    ' 1000 ms = 1 second
Next startIndex
```
Each outer iteration processes 50 paragraphs. At 33,857 paragraphs that is 678 batches
with a 1-second pause between each — approximately 11 minutes of clock time at minimum
plus processing time per batch.

**7. ExecuteItem: what happens to each paragraph**

`aeUpdateCharStyleClass.cls:53`
```vba
Private Function IaeLongProcessClass_ExecuteItem(itemIndex As Long) As Boolean
    Dim para As Word.Paragraph
    Set para = ActiveDocument.Paragraphs(itemIndex)
    For Each rng In para.Range.Characters
        If rng.style = StyleName Then   ' "Chapter Verse marker"
            rng.style = StyleName       ' re-apply same style
        End If
    Next rng
    IaeLongProcessClass_ExecuteItem = True  ' continue
End Function
```
For each of the 33,857 paragraphs: walk every character, find ones carrying
`"Chapter Verse marker"`, re-apply the same style. This forces Word to rebuild its
internal style data for those characters — a repair operation with no visible content
change.

**8. DoEvents between items**

After each `ExecuteItem` call, `DoEvents` runs. This yields control to the Windows
message loop, allowing Word to repaint, process ribbon callbacks, and keep the UI
responsive. Without it, Word would be "not responding" for the entire duration.

**9. SaveProgress after each batch**

`aeLongProcessClass.cls:157` writes to `rpt/LongRunningProgress_UpdateCharacterStyle.txt`:
```
LastProcessedItem=4951
ProgressPercentage=14.62
```
This is what enabled resume from 14.6% after the earlier IDE Stop.

**10. PauseWithDoEvents between batches**

```vba
Private Sub PauseWithDoEvents(ByVal milliseconds As Long)
    Dim startTime As Single
    startTime = Timer
    Do While Timer < startTime + (milliseconds / 1000)
        DoEvents
        Sleep 10
    Loop
End Sub
```
Sleeps 1 second between batches in 10 ms slices, calling `DoEvents` each time. This
is the source of reduced responsiveness during the run — Word is processing but
yielding briefly every 10 ms.

**11. Completion**

After the last batch, `m_lastProcessedItem` is reset to 0 and the sidecar is updated:
```
LastProcessedItem=0
ProgressPercentage=100
```
The log writes `"Complete: 33857 Items processed"`, the Immediate Window prints the
completion line, and `log.Log_Close` flushes `rpt/LongProcess_UpdateCharacterStyle.txt`.

---

### Key Design Properties

| Property | How it works |
|----------|-------------|
| **Resume** | `LoadProgress` reads the sidecar; outer loop starts at saved item |
| **Stop** | IDE Stop button only (VBA single-threaded; Escape unreliable in Word) |
| **State-neutral runner** | No `ScreenUpdating` or `Options.Pagination` — safe if user opens dialogs during `DoEvents` |
| **Task is pluggable** | Any class implementing `IaeLongProcessClass` drops into the runner unchanged |
| **Log** | `rpt/LongProcess_UpdateCharacterStyle.txt` — UTF-8, one line per batch |
| **Progress** | `rpt/LongRunningProgress_UpdateCharacterStyle.txt` — key=value, updated after every 50 items |

---
