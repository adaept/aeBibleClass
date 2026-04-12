# Code Review - 2026-04-10a

## Carry-Forward: 3-Button Ribbon Plan + Open Items

Continues from `rvw/Code_review - 2026-04-10.md`.

---

## § 1 — Status of Previous Session (2026-04-10)

### Completed

| Item | Detail |
|------|--------|
| Long-running process framework | Fully implemented and validated (33,857 paragraphs, 100%) |
| aeUpdateCharStyleClass optimization | Short-circuit `Exit For` + dual-style pass implemented and tested |
| JUDE sample review | Document structure confirmed (§ 14-15 of prior file) |
| Style names confirmed | `"Chapter Verse marker"`, `"Verse marker"` — both correct |
| Heading 2 format confirmed | `"CHAPTER N"` (e.g. `"CHAPTER 1"`) |
| Diagnostic Debug.Print | Added to diagnose runner exit; removed after confirmation |

### Pending import

Two clean files are ready (diagnostic prints removed, CRLF converted):

- `src/aeLongProcessClass.cls`
- `src/aeUpdateCharStyleClass.cls`

Import both, then run `TestUpdateCharStyle` to resume from the saved sidecar position.

---

## § 2 — Open Questions (from Code_review - 2026-04-10.md § 11)

| # | Question | Status |
|---|----------|--------|
| Q1 | GoTo Chapter: Option A (lazy `chapterData` array) vs Option B (Find at runtime) | **Resolved — see § 6** |
| Q2 | GoTo Verse: Option B (Find forward from chapter, count paragraphs) as starting point | **Pending approval** |
| Q3 | Combo-box UserForm: deferred | Confirmed deferred |
| Q4 | Next/Prev operate on books only after chapter navigation | Noted as pre-existing limitation |
| Q5 | Heading 2 text format | **Resolved:** `"CHAPTER N"` |
| Q6 | Ribbon XML edit process | **Pending:** confirm Custom UI Editor or direct XML |

---

## § 3 — 3-Button Ribbon Implementation Plan (carry-forward)

Full design detail is in `rvw/Code_review - 2026-04-10.md` §§ 3-9.
Steps are listed here for tracking.

### Step 1 — GoTo Book: input normalization (NEXT)

Add `NormalizeBookInput` to `aeRibbonClass.cls`.
Add `m_currentBookIndex As Long` and `m_currentBookPos As Long` instance variables.
Set both when a match is found in `GoToH1`.

**Rule:** `NormalizeBookInput` is kept in `aeRibbonClass.cls`, separate from
`aeBibleCitationClass`. It is a UI input helper, not a citation parser.

Normalization logic:

```vba
Private Function NormalizeBookInput(ByVal raw As String) As String
    Dim s As String
    s = Trim(UCase(raw))
    If Len(s) >= 2 Then
        If s Like "[0-9][A-Z]*" Then s = Left$(s, 1) & " " & Mid$(s, 2)
    End If
    NormalizeBookInput = s
End Function
```

Apply in `GoToH1`:

```vba
If CStr(headingData(i, 0)) Like "*" & NormalizeBookInput(pattern) & "*" Then
```

---

### Step 2 — ELIMINATED: CaptureHeading2s not needed

The original plan (Option A) called for a `CaptureHeading2s` scan that would build a
`chapterData(1 To 66, 1 To 150)` array of character positions for all 1,189 chapter
headings at session start. This pre-scan was intended to solve two problems:

| Problem | Old approach | New approach |
|---------|-------------|--------------|
| Chapter count (validation) | Count non-zero entries in `chapterData(bookIdx, *)` | `aeBibleCitationClass.ChaptersInBook(bookName)` — no scan |
| Chapter position (navigation) | Array lookup `chapterData(bookIdx, chNum)` | Runtime Find (N steps forward from book H1) |

**Why `CaptureHeading2s` is not needed:**

Problem 1 — "How many chapters does Genesis have?" — is now answered by
`aeBibleCitationClass.ChaptersInBook("GENESIS")` → 50. This is Step 4 (make the
function Public). No document scan required.

Problem 2 — "Where in the document is Genesis Chapter 3?" — is answered at runtime
by Word's Find:

```text
1. Start at Genesis H1 position (known from headingData)
2. Find next Heading 2 → "CHAPTER 1"
3. Find next Heading 2 → "CHAPTER 2"
4. Find next Heading 2 → "CHAPTER 3"  ← arrived
```

Each Find call takes milliseconds. Worst case is Psalms chapter 150 — 150 Find
executions, still well under one second. The pre-scan built a lookup table for all
1,189 chapters up front to avoid this cost. Runtime Find pays the same cost for the
one chapter actually requested.

**Why `ExpectedChapterCounts` in `basSBL_VerseCountsGenerator.bas` cannot be used
directly:** it is `Private` and the module has `Option Private Module`. Both barriers
make it inaccessible from `aeRibbonClass`. `aeBibleCitationClass.ChaptersInBook` is
the correct path — already in the plan (Step 4).

**Net change to state variables:** `chapterData` array removed. Only
`m_currentChapter As Long` is needed (to track the chapter selected for use by GoTo
Verse).

---

### Step 3 — GoToChapter implementation

Add `Private Sub GoToChapter()` and `Public Sub OnGoToChapterButtonClick`.

Logic:

1. If `m_currentBookIndex = 0` → infer book from cursor (scan `headingData` backward)
2. Get chapter count: `aeBibleCitationClass.ChaptersInBook(headingData(m_currentBookIndex, 0))`
3. InputBox: `"Enter chapter number (1-N):"`
4. Validate range (1 to N)
5. Navigate: set cursor to book H1 position, then Find Heading 2 forward N times
6. Set `m_currentChapter = chNum`; update Prev/Next enabled state

---

### Step 4 — Expose ChaptersInBook and VersesInChapter as Public

In `aeBibleCitationClass.cls`: change `Private Function ChaptersInBook` and
`Private Function VersesInChapter` to `Public Function`.

These are pure data look-ups. No behavioural change. Used by GoTo Verse for range
validation only — not for navigation.

---

### Step 5 — GoToVerse implementation

Add `Private Sub GoToVerse()` and `Public Sub OnGoToVerseButtonClick`.

Range validation via `aeBibleCitationClass.VersesInChapter`.

Navigation method depends on document type — see § 6 for full design.

---

### Step 6 — Ribbon XML update

Add two new buttons to the `.docm` Custom UI XML.
Add callback stubs in `basBibleRibbonSetup.bas`.

**Pending Q6:** confirm edit process (Custom UI Editor or direct XML).

---

### Step 7 — Move OLD_CODE

Move superseded stubs to `basOLD_CODE.bas`:

- `UpdateCharacterStyle` legacy sub in `basLongProcess.bas`
- `GoToVerseSBL` stub in `aeRibbonClass.cls`

---

### Step 8 — normalize_vba.py update

Add new identifiers:
`NormalizeBookInput`, `GoToChapter`, `GoToVerse`, `ChaptersInBook`,
`VersesInChapter`, `m_currentBookIndex`, `m_currentChapter`,
`GoToVerseByCount`, `GoToVerseByScan`, `ValidateVerseMarkers`.

---

## § 6 — GoTo Verse: Design (Q2 Resolution) (2026-04-10)

### Document type detection

The document exists in two forms:

| Form | Paragraph count | Verse layout |
|------|----------------|-------------|
| Study version (current) | > 30,000 (confirmed: 33,857) | One verse per paragraph |
| Print candidate | ≤ 30,000 | Multiple verses per paragraph |

Detection:

```vba
Private Function IsStudyVersion() As Boolean
    IsStudyVersion = (ActiveDocument.Paragraphs.Count > 30000)
End Function
```

`GoToVerse` calls `IsStudyVersion` and delegates to the appropriate navigation method.

---

### Method 1 — Paragraph count (study version)

Each verse is exactly one paragraph. From the chapter's Heading 2 position, the Nth
verse is the Nth body paragraph below that heading.

```vba
' Navigate to chapter first (already done by GoToChapter or inferred)
' Then move down N paragraphs
Selection.SetRange chapterPos, chapterPos
Selection.MoveDown Unit:=wdParagraph, Count:=verseNum
Selection.Collapse Direction:=wdCollapseStart
```

**Assumption:** paragraphs between the chapter H2 and the next H2 (or H1) are all
verse paragraphs. This holds if the document is correctly formatted. Validation
(see § 7) confirms this assumption.

---

### Method 2 — Verse marker scan (print candidate)

Multiple verses share a paragraph. Navigate by scanning for the Nth occurrence of a
run with `"Verse marker"` character style, starting from the chapter's Heading 2
position.

```vba
Dim found As Boolean
Dim count As Long
Selection.SetRange chapterPos, chapterPos
Selection.Find.ClearFormatting
Selection.Find.Text = ""
Selection.Find.style = ActiveDocument.Styles("Verse marker")
Selection.Find.Forward = True
Selection.Find.Wrap = wdFindStop
count = 0
Do
    found = Selection.Find.Execute
    If Not found Then Exit Do
    count = count + 1
Loop Until count = verseNum
If found Then Selection.Collapse Direction:=wdCollapseStart
```

Each Find execution locates the next "Verse marker" run. For verse N, the loop
executes N times.

---

### Worst-case timing (to be measured)

| Scenario | Operations | Expected |
|----------|-----------|---------|
| GoTo Chapter: Psalm 150 | 150 × Heading 2 Find from Psalms H1 | < 1 second |
| GoTo Verse: Psalm 119:176 (count method) | 119 × H2 Find + 176 × MoveDown | < 1 second |
| GoTo Verse: Psalm 119:176 (scan method) | 119 × H2 Find + 176 × Verse marker Find | < 2 seconds |

**Timing test to confirm:** Psalm 119 is the longest chapter (176 verses). Time both
methods for verse 176 using `Timer` before and after navigation:

```vba
Dim t As Single
t = Timer
' ...navigation code...
Debug.Print "GoToVerse elapsed: " & Format(Timer - t, "0.000") & "s"
```

Record results and add to this review before committing the implementation.

---

## § 7 — Verse Marker Validation: Long Process Improvement (2026-04-10)

### Purpose

The current `aeUpdateCharStyleClass` repairs "Chapter Verse marker" and "Verse marker"
by re-applying them. It does not validate that they are **correct** — i.e. that the
chapter number and verse number encoded in those runs match what is expected.

A new `IaeLongProcessClass` task, `aeValidateVerseMarkersClass`, will validate marker
correctness book by book and report PASS / FAIL.

---

### What to validate per verse paragraph

Each verse paragraph begins with exactly two marker runs (confirmed from JUDE sample):

```
Run 1:  "Chapter Verse marker"  — text = chapter number (e.g. "1", "3", "119")
Run 2:  "Verse marker"          — text = verse number + narrow no-break space (e.g. "1 ", "176 ")
```

Validation checks per paragraph:

1. First character has style `"Chapter Verse marker"` and its text = expected chapter number
2. Second run has style `"Verse marker"` and its numeric text = expected verse number (in sequence)
3. Verse numbers increment correctly from 1 to `VersesInChapter(book, chapter)`
4. No extra or missing marker runs at the paragraph start

---

### PASS / FAIL behaviour

- **PASS** per book: all verse paragraphs in the book have correct, sequential markers
- **FAIL**: stop at the first error; log the book name, chapter number, verse number,
  expected value, and actual value; allow the user to fix the document before re-running

Output goes to `rpt/LongProcess_ValidateVerseMarkers.txt` via `aeLoggerClass`.
Progress is persisted to `rpt/LongRunningProgress_ValidateVerseMarkers.txt` so a
partial validation can be resumed after a fix.

---

### Entry point

```vba
' basLongProcess.bas
Public Sub TestValidateVerseMarkers()
    Dim t As New aeValidateVerseMarkersClass
    StartOrResume t
End Sub
```

---

### Document type parameter

The validator needs to know whether to expect one verse per paragraph (study version)
or multiple verses per paragraph (print candidate). Use `IsStudyVersion()` detection
(§ 6) or allow caller to override:

```vba
' aeValidateVerseMarkersClass
Public ForceStudyMode As Boolean   ' True = one verse per para; False = scan mode
                                   ' default: auto-detect via paragraph count
```

---

### Relationship to aeUpdateCharStyleClass

These are two separate tasks with different purposes:

| Task | Purpose | Action on error |
|------|---------|----------------|
| `aeUpdateCharStyleClass` | Repair — re-apply marker styles | Continues (style re-application) |
| `aeValidateVerseMarkersClass` | Validate — check marker correctness | Stops with FAIL + location |

The repair task should be run first; the validation task confirms the repair was
complete and correct.

---

## § 4 — Additional Issues (carry-forward from Code_review - 2026-04-10.md § 9)

| Issue | Detail | Action |
|-------|--------|--------|
| `CaptureHeading1s` Static flag blocks refresh | Heading changes within a session are not picked up | Accept for read-only Bible navigation |
| `m_currentBookIndex = 0` when user navigates manually | GoTo Chapter must infer book from cursor | Handled in Step 3 design |
| 1189 Heading 2 capture cost | Pre-scan no longer needed — runtime Find + ChaptersInBook replace it | Resolved: Step 2 eliminated |
| `ChaptersInBook` / `VersesInChapter` are Private | Cannot call from ribbon | Fixed in Step 4 |
| Next/Prev operate on books only | After chapter navigation, Next goes to next book | Pre-existing limitation; not in scope |

---

## § 8 — Ribbon Layout Design: Progressive Lock (2026-04-10)

### Confirmed layout

Three stacked button columns, each a self-contained navigation level:

```text
Stack 1              Stack 2              Stack 3              Separate
-----------          -----------          -----------          -----------
GoTo Book            GoTo Chapter         GoTo Verse           New Search
Prev Book            Prev Chapter         Prev Verse           About
Next Book            Next Chapter         Next Verse
```

### Progressive enable/disable state

| State | Stack 1 | Stack 2 | Stack 3 | New Search |
|-------|---------|---------|---------|------------|
| Default (open) | enabled | disabled | disabled | disabled |
| After GoTo Book | enabled | enabled | disabled | enabled |
| After GoTo Chapter | disabled | enabled | disabled | enabled |
| After GoTo Verse | disabled | disabled | enabled | enabled |

**New Search** clears `m_currentBookIndex` and `m_currentChapter`, re-enables Stack 1,
disables Stacks 2 and 3, and invalidates the ribbon. The name communicates user intent
(start a new navigation sequence) rather than a technical operation.

### Why this design was chosen over context-aware Prev/Next

The alternative considered was a 2-stack layout with a single context-aware Prev/Next
pair (6 buttons total). **Rejected** because:

- The active navigation level is invisible — pressing Prev/Next with no visual cue as
  to whether it moves by book, chapter, or verse is confusing.
- The 3-stack layout makes the active level explicit: only one column is enabled at a
  time, so the user always knows what Prev/Next will do.
- The extra 3 buttons are justified by the clarity they provide.

### Ribbon XML (updated)

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="RibbonOnLoad">
  <ribbon startFromScratch="false">
    <tabs>
      <tab id="RWB" label="Radiant Word Bible">
        <group id="TestGroup" label="Bible Class Group">
          <button id="GoToH1Button"       imageMso="NotebookNew"                 size="normal" showLabel="false" onAction="OnGoToH1ButtonClick"       screentip="Go To Book"       getEnabled="GetBookEnabled"/>
          <button id="GoToPrevButton"     imageMso="HeaderFooterPreviousSection" size="normal" showLabel="false" onAction="OnPrevButtonClick"          screentip="Previous Book"    getEnabled="GetPrevEnabled"/>
          <button id="GoToNextButton"     imageMso="HeaderFooterNextSection"     size="normal" showLabel="false" onAction="OnNextButtonClick"          screentip="Next Book"        getEnabled="GetNextEnabled"/>
          <separator id="sep1"/>
          <button id="GoToChapterButton"  imageMso="GoToPage"                    size="normal" showLabel="false" onAction="OnGoToChapterButtonClick"   screentip="Go To Chapter"    getEnabled="GetChapterEnabled"/>
          <button id="GoToPrevChButton"   imageMso="HeaderFooterPreviousSection" size="normal" showLabel="false" onAction="OnPrevChapterButtonClick"   screentip="Previous Chapter" getEnabled="GetPrevChEnabled"/>
          <button id="GoToNextChButton"   imageMso="HeaderFooterNextSection"     size="normal" showLabel="false" onAction="OnNextChapterButtonClick"   screentip="Next Chapter"     getEnabled="GetNextChEnabled"/>
          <separator id="sep2"/>
          <button id="GoToVerseButton"    imageMso="FormatNumberDefault"         size="normal" showLabel="false" onAction="OnGoToVerseButtonClick"     screentip="Go To Verse"      getEnabled="GetVerseEnabled"/>
          <button id="GoToPrevVerseButton" imageMso="HeaderFooterPreviousSection" size="normal" showLabel="false" onAction="OnPrevVerseButtonClick"   screentip="Previous Verse"   getEnabled="GetPrevVerseEnabled"/>
          <button id="GoToNextVerseButton" imageMso="HeaderFooterNextSection"    size="normal" showLabel="false" onAction="OnNextVerseButtonClick"    screentip="Next Verse"       getEnabled="GetNextVerseEnabled"/>
          <separator id="sep3"/>
          <button id="NewSearchButton"    imageMso="Undo"                        size="normal" showLabel="false" onAction="OnNewSearchButtonClick"     screentip="New Search"       getEnabled="GetNewSearchEnabled"/>
          <separator id="sep4"/>
          <button id="adaeptButton"       label="About" image="adaept"           size="large"                   onAction="OnAdaeptAboutClick"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

**Note:** `imageMso="GoToPage"` is a placeholder for the Chapter button — confirm the
correct icon name from the Office icon gallery before finalising.

### Open question: WarmLayoutCache startup block

Every document open currently freezes Word for ~50 seconds while `WarmLayoutCache`
navigates to Revelation to pre-build the layout table. File Explorer comes to the
foreground during the freeze.

**Recommendation:** remove `WarmLayoutCache` entirely. The first GoTo Book navigation
will take ~12 seconds once per session (layout built on demand). Every document open
will be immediate. This is better UX than a guaranteed 50-second freeze at startup.

Decision required before any ribbon implementation begins.

---

## § 9 — Ribbon State Matrix Design (2026-04-10)

### Pattern

Nine buttons across three levels represented as a 3×3 Boolean matrix. Four
pre-defined states replace scattered `InvalidateControl` calls throughout GoTo/Prev/Next
procedures.

```text
              Prev    GoTo    Next
Book:         [1,1]   [1,2]   [1,3]
Chapter:      [2,1]   [2,2]   [2,3]
Verse:        [3,1]   [3,2]   [3,3]
```

**Column order — Prev / GoTo / Next** mirrors the universal UI convention (previous ←
anchor → next), matching browser history buttons, media controls, and Word's own
section navigation. GoTo occupies the centre column (2), which is the natural anchor.

| State | Prev Bk | GoTo Bk | Next Bk | Prev Ch | GoTo Ch | Next Ch | Prev Vs | GoTo Vs | Next Vs |
|-------|---------|---------|---------|---------|---------|---------|---------|---------|---------|
| Default | OFF | ON | OFF | OFF | OFF | OFF | OFF | OFF | OFF |
| Book selected | ON | ON | ON | OFF | ON | OFF | OFF | OFF | OFF |
| Chapter selected | OFF | OFF | OFF | ON | ON | ON | OFF | ON | OFF |
| Verse selected | OFF | OFF | OFF | OFF | OFF | OFF | ON | ON | ON |

`SetNavState` compares the desired matrix to `m_navState`, calls `InvalidateControl`
only for cells that changed, then updates `m_navState`. `GetEnabled` callbacks read
directly from the matrix. `New Search` calls `SetNavState STATE_DEFAULT` and clears
`m_currentBookIndex` / `m_currentChapter` in one operation.

**`Option Base 1` required.** Any module that declares the `m_navState` array must
include `Option Base 1` at the top, otherwise VBA allocates indices 0-3 while the
design uses 1-3, wasting row/column 0 and risking silent off-by-one errors if a
developer forgets the convention.

Index constants (mitigate VBA 2D array verbosity). The numeric values are the
canonical representation; the names are readable aliases — both are needed:

```vba
Private Const NAV_BOOK    As Long = 1
Private Const NAV_CHAPTER As Long = 2
Private Const NAV_VERSE   As Long = 3
Private Const BTN_PREV    As Long = 1
Private Const BTN_GOTO    As Long = 2
Private Const BTN_NEXT    As Long = 3
```

### Boundary condition overlay

The matrix handles level activation. Prev/Next boundary conditions (first/last item)
are a secondary check in the `GetEnabled` callbacks:

```vba
Public Function GetPrevEnabled(control As IRibbonControl) As Boolean
    GetPrevEnabled = m_navState(NAV_BOOK, BTN_PREV) And m_btnPrevEnabled
End Function

Public Function GetPrevChEnabled(control As IRibbonControl) As Boolean
    GetPrevChEnabled = m_navState(NAV_CHAPTER, BTN_PREV) And (m_currentChapter > 1)
End Function
```

The matrix says whether a level is active; the boundary flag says whether there is a
previous item within that level.

### Pros

1. **Single source of truth** — all valid states declared in one place; no state logic
   scattered across GoTo/Prev/Next procedures
2. **Minimal InvalidateControl calls** — only changed buttons are invalidated
3. **VBA reset resilience** — `SetNavState STATE_DEFAULT` restores known state in one call
4. **Auditable** — matrix can be printed at any point to verify ribbon state
5. **Extensible** — adding a button or state is a matrix row/column change
6. **Self-documenting transitions** — `SetNavState STATE_BOOK_SELECTED` at the end of
   `GoToH1` is immediately readable

### Cons

> **NOTE: The following cons require further discussion before implementation begins.**

1. **Prev/Next boundary conditions not captured in the matrix** — first/last item
   detection is a secondary layer on top. The interaction between the two layers needs
   to be designed carefully to avoid conflicts.

2. **State drift after VBA reset** — `m_navState` is an instance variable destroyed by
   the IDE Stop button. The ribbon UI retains its last visual state but `m_navState`
   resets to all-False. The next `GetEnabled` callback returns False for everything,
   effectively disabling all buttons. A rehydration strategy is needed.

3. **Matrix reflects code state, not cursor position** — if a user manually scrolls to
   a different location, `m_navState` still reflects the last navigation action. The
   matrix is authoritative for button state only; it does not track document position.

4. **Four clean states may not cover all transitions** — e.g. GoTo Book followed by
   GoTo Book again (same or different book) needs to decide whether to re-enter
   STATE_BOOK_SELECTED cleanly or preserve chapter state. Edge cases need enumeration
   before coding.

---

## § 5 — Pending Decisions Before Step 1 Begins

Decisions needed before any code is written:

1. **Q1:** Resolved — `CaptureHeading2s` eliminated; runtime Find + `ChaptersInBook`.
2. **Q2:** Resolved — paragraph count (study) / Verse marker scan (print); see § 6.
3. **Q6:** Resolved — Office RibbonX Editor; see § 10.
4. **WarmLayoutCache:** Resolved — entry points commented out in `EnableButtonsRoutine` and `WarmLayoutCacheDeferred`; `WarmLayoutCache` itself preserved for future use.
5. **Ribbon layout:** 3-stack progressive lock confirmed (§ 8). imageMso for Chapter button to confirm.

Steps 1 and 2 can begin without Q6 (ribbon XML is Step 6). Steps 1-5 only touch
`aeRibbonClass.cls` and `aeBibleCitationClass.cls`.

---

## § 10 — Q6 Resolved: Ribbon XML Edit Process (2026-04-11)

### Tool decision: Office RibbonX Editor

**Use the Office RibbonX Editor** (`fernandreu/office-ribbonx-editor` on GitHub).

The original Microsoft "Custom UI Editor for Office" has been abandoned since ~2010
and lacks support for the Office 2010 schema (`customUI14`). The GitHub fork is a
complete WPF rewrite, not a patch. It is the de-facto modern replacement.

| Feature | Old Microsoft tool | Office RibbonX Editor |
|---------|-------------------|----------------------|
| Office 2010 schema (`customUI14`) | No | Yes |
| XML validation | No | Yes |
| Schema-aware IntelliSense | No | Yes |
| Image import/export | No | Yes |
| Active maintenance | No (dead since ~2010) | Yes |
| Latest release | n/a | v2.0.0 (Nov 16, 2025) |
| Dark mode | No | Yes |

**Primary workflow:** open `.docm` directly in RibbonX Editor, edit
`customUI14.xml` with IntelliSense, validate, save. Icon names (`imageMso`) are
resolvable directly in the editor without guessing.

---

### Can a .docx file include a ribbon?

**Yes — technically. No — not usefully, for this project.**

The Office Open XML format allows `customUI` markup in any document type, including
`.docx`. The RibbonX Editor will open a `.docx` and add ribbon XML to it without
complaint. The ribbon buttons will appear in Word.

However, ribbon callbacks (`onAction`, `getEnabled`, etc.) must resolve to a
procedure at runtime. In a `.docx` there is no VBA project — the file format
explicitly prohibits it. Clicking a button produces:
`"Cannot run the macro 'OnGoToH1ButtonClick'. The macro may not be available..."`

**The document must remain `.docm`** because the callbacks live in the VBA project
embedded in the same file. Word will warn and strip macros if you attempt to save
`.docm` as `.docx`.

Exception — **in scope as a future deliverable:** a COM add-in can host callbacks
for a `.docx` ribbon, delivering the full navigation and citation interface without
requiring the user to open a macro-enabled document. See § 11 for distribution
requirements and development path considerations.

---

### Extract XML process (fallback / audit)

When you need to inspect or recover the ribbon XML outside RibbonX Editor — for
example to diff against the backup in `customUI14backupRWB.xml` or to diagnose a
corruption — a `.docm` is a ZIP archive and can be unpacked directly.

**Manual steps:**

```
1. Close the document in Word.
2. Copy MyDoc.docm → MyDoc_work.docm
3. Rename MyDoc_work.docm → MyDoc_work.zip
4. Open the ZIP; navigate to customUI/
5. Extract customUI14.xml
6. Edit in any text editor
7. Replace customUI14.xml in the ZIP
8. Rename back to .docm
```

**Automation — preferred: WSL bash**

The development system has WSL installed. Bash is already the shell used for the
Python normalizer (`normalize_vba.py`) and is the preferred automation tool for this
project. The standard `unzip` / `zip` utilities handle `.docm` files directly since
they are ZIP archives. Windows paths are accessed via `/mnt/c/...`.

Extract ribbon XML:

```bash
#!/usr/bin/env bash
DOCM="/mnt/c/adaept/aeBibleClass/rpt/MyDoc.docm"
OUT="/mnt/c/adaept/aeBibleClass/rpt/customUI14_extract.xml"

unzip -p "$DOCM" "customUI/customUI14.xml" > "$OUT"
echo "Extracted to $OUT"
```

Replace ribbon XML (document must be closed in Word):

```bash
#!/usr/bin/env bash
DOCM="/mnt/c/adaept/aeBibleClass/rpt/MyDoc.docm"
XML="/mnt/c/adaept/aeBibleClass/rpt/customUI14_extract.xml"

# zip -j replaces the file in-place without extracting the whole archive
zip "$DOCM" -j "$XML" --archive-name "customUI/customUI14.xml"
echo "Replaced customUI14.xml in $DOCM"
```

Diff against the committed backup:

```bash
diff <(unzip -p "$DOCM" "customUI/customUI14.xml") \
     /mnt/c/adaept/aeBibleClass/customUI14backupRWB.xml
```

**PowerShell alternative** (if WSL is unavailable):

```powershell
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip   = [System.IO.Compression.ZipFile]::OpenRead($DocmPath)
$entry = $zip.Entries | Where-Object { $_.FullName -eq "customUI/customUI14.xml" }
$entry.Open() | % { (New-Object System.IO.StreamReader($_)).ReadToEnd() } |
    Set-Content -Path $OutPath -Encoding UTF8
$zip.Dispose()
```

**When to use the extract process vs RibbonX Editor:**

| Situation | Use |
|-----------|-----|
| Normal editing / icon selection | RibbonX Editor |
| Diff ribbon XML against backup in git | WSL bash extract → diff |
| Diagnose a file that RibbonX Editor won't open | Manual extraction |
| Automated ribbon injection / CI | WSL bash replace script |
| Recover from a corrupt customUI save | Extract → repair → replace |

---

## § 11 — Future Distribution: COM Add-in Requirements (2026-04-11)

### Context

The current `.docm` + VBA approach is the **development vehicle**. The target
delivery mechanism for end-user distribution is a COM add-in, enabling the full
navigation and citation interface to be delivered to users who open the Study Bible
as a plain `.docx` — no macro prompts, no Trust Center configuration required.

No implementation timeline has been set.

---

### Requirements

| # | Requirement |
|---|-------------|
| 1 | Distribute the Study Bible to end users, including i18n audiences |
| 2 | Word 365 support only (this version) |
| 3 | Ribbon interface includes navigation (Book / Chapter / Verse) and citation block lookup, using the existing citation block verification code with minor adjustments |
| 4 | Add-in available via Microsoft Store |
| 5 | Development process must accommodate the Store publication path from the outset |
| 6 | Code signed; secure distribution |
| 7 | No current implementation timeline |

---

### Technology path

**VSTO (Visual Studio Tools for Office) packaged as MSIX** is the recommended path
for a COM add-in targeting Word 365 on Windows, distributed via the Microsoft Store.

| Layer | Technology |
|-------|------------|
| Add-in host | VSTO — **VB.NET preferred** (see note below), compiled to a COM-visible DLL |
| Ribbon XML | Reuse `customUI14.xml` from the `.docm` directly — no redesign |
| Callbacks | Port VBA subs to .NET methods; signatures are identical in structure |
| Packaging | MSIX (Windows App Installer) wrapping the VSTO installer |
| Store submission | Microsoft Partner Center → Microsoft Store for Business / consumer |
| Code signing | EV (Extended Validation) code-signing certificate required for Store submission |

**VB.NET vs C#:** Both are fully supported for VSTO. VB.NET is the natural choice
here for two reasons: (1) the existing codebase is VBA — VB.NET shares the same
language lineage, so identifier names, control flow, and Office object model calls
port with minimal syntactic friction; (2) VB.NET retains optional parameters,
`With` blocks, and late binding via `Object` in a way that mirrors VBA idioms.
C# is equally capable but requires more mechanical translation. Either language
produces an identical MSIX/Store deliverable — the choice affects only the porting
effort, not the output.

**Alternative — Office JS Add-in (web-based):** Microsoft's strategic direction for
Office extensibility is the JavaScript/TypeScript API hosted in a browser frame.
It is cross-platform (Windows, Mac, web). However, the Office JS API does not yet
expose the full paragraph-level navigation and character-style inspection that the
current VBA code relies on. VSTO is the appropriate choice for this feature set.

---

### Development process considerations

To avoid a costly rewrite when moving from VBA to VSTO, the VBA code should be
structured so that logic is easy to port:

1. **Separation of concerns** — keep ribbon callback stubs thin; all logic in class
   methods (`aeRibbonClass`, `aeBibleCitationClass`). This maps directly to .NET
   classes in a VSTO project.
2. **No Word object model shortcuts** — use explicit `ActiveDocument` /
   `Application` references rather than implicit globals. VSTO requires explicit
   references; VBA that already uses them ports without change.
3. **Ribbon XML is reusable as-is** — VSTO loads the same `customUI14.xml`; callback
   attribute names map directly to .NET method names.
4. **i18n** — string literals used in ribbon `screentip`, `label`, and `MsgBox`
   calls should be centralised in a single resource location (a `bas` module now,
   a `.resx` file in .NET). Avoid embedding UI strings inline in logic procedures.
5. **Code-signing discipline starts now** — the VBA project should be digitally
   signed with the same certificate intended for the VSTO add-in. This establishes
   the signing workflow and trust chain before the Store submission process begins.

---

### Open questions (no timeline)

- Certificate provider for EV code signing (DigiCert, Sectigo, or equivalent)
- Microsoft Partner Center account setup
- Whether the `.docx` Study Bible will be distributed via the Store alongside the
  add-in, or separately
- i18n scope for the first Store release (languages to support)

---

## § 13 — Boundary Condition Design (2026-04-11)

### The two-layer problem

The state matrix answers one question: **is this navigation level active?**

It does not know where you are *within* an active level. That is a second, separate
question: **is there a previous / next item from the current position?**

Both must be true for a Prev or Next button to be enabled. The `GetEnabled` callback
ANDs them:

```vba
Public Function GetPrevChEnabled(control As IRibbonControl) As Boolean
    GetPrevChEnabled = m_navState(NAV_CHAPTER, BTN_PREV) And (m_currentChapter > 1)
End Function
```

The matrix supplies the left side. The position check supplies the right side. When
either is False the button is disabled.

---

### Where it breaks down

**Scenario 1 — stale boundary flag on level change**

User navigates to Genesis (first book). Prev Book correctly disabled —
`m_currentBookIndex = 1`. User types "Exodus" in the Book comboBox and confirms.
Matrix still says Book level active. If `m_currentBookIndex` is not updated before
the next `GetEnabled` callback fires, Prev Book stays disabled even though Exodus
has a predecessor. Matrix transitioned correctly; position variable did not keep up.

**Scenario 2 — New Search resets matrix but not position**

User navigates to Revelation (last book). Next Book correctly disabled. User clicks
New Search. Matrix resets to `STATE_DEFAULT`. User types "Revelation" again. Matrix
sets Book row active, Prev and Next both ON (per the pre-defined state). But
`m_currentBookIndex` still holds 66. Next Book should be OFF at the boundary — the
matrix says ON. The `And` gives the correct answer only if the position variable was
current when the callback fired.

**Scenario 3 — inconsistent boundary patterns**

The current design uses `m_btnPrevEnabled` (a separate Boolean flag) for Book
Prev/Next, but inline expressions (`m_currentChapter > 1`) for Chapter and Verse.
These are two different patterns for the same problem. The flag can go stale
independently; the inline expression derives from authoritative position state. They
must be unified.

---

### Three questions to resolve

1. Who updates each position variable, and when?
2. Are position variables reset when `SetNavState STATE_DEFAULT` is called?
3. Is `m_btnPrevEnabled` kept, or replaced by an inline expression consistent with
   Chapter and Verse?

---

### Proposed solutions

**Q1 — Who updates position variables, and when?**

Each navigation procedure owns its own position variable. Update occurs as the last
step before ribbon invalidation — never before, so the variable is current when the
next `GetEnabled` callback fires:

```
GoToBook    → sets m_currentBookIndex, m_currentChapter = 0, m_currentVerse = 0
GoToChapter → sets m_currentChapter, m_currentVerse = 0
GoToVerse   → sets m_currentVerse
PrevBook    → sets m_currentBookIndex (decremented), m_currentChapter = 0, m_currentVerse = 0
NextBook    → sets m_currentBookIndex (incremented), m_currentChapter = 0, m_currentVerse = 0
PrevChapter → sets m_currentChapter (decremented), m_currentVerse = 0
NextChapter → sets m_currentChapter (incremented), m_currentVerse = 0
PrevVerse   → sets m_currentVerse (decremented)
NextVerse   → sets m_currentVerse (incremented)
```

Downstream variables are zeroed on every upward level change. A zero value means
"not yet set at this level" and renders the comboBox blank.

**Q2 — Are position variables reset on New Search?**

Yes — all three zeroed explicitly in `OnNewSearchButtonClick`, before
`SetNavState STATE_DEFAULT`. This ensures that when the user re-enters Book
navigation, all boundary expressions evaluate from a clean baseline:

```vba
Public Sub OnNewSearchButtonClick(control As IRibbonControl)
    m_currentBookIndex = 0
    m_currentChapter   = 0
    m_currentVerse     = 0
    SetNavState STATE_DEFAULT
End Sub
```

**Q3 — Retire `m_btnPrevEnabled`; use inline expressions throughout**

Replace the separate Boolean flag with the same pattern used for Chapter and Verse.
All six boundary expressions then derive directly from position variables — one
pattern, no secondary state to maintain:

```vba
' Book row
Public Function GetPrevBkEnabled(control As IRibbonControl) As Boolean
    GetPrevBkEnabled = m_navState(NAV_BOOK, BTN_PREV) And (m_currentBookIndex > 1)
End Function
Public Function GetNextBkEnabled(control As IRibbonControl) As Boolean
    GetNextBkEnabled = m_navState(NAV_BOOK, BTN_NEXT) And (m_currentBookIndex < BOOK_COUNT)
End Function

' Chapter row
Public Function GetPrevChEnabled(control As IRibbonControl) As Boolean
    GetPrevChEnabled = m_navState(NAV_CHAPTER, BTN_PREV) And (m_currentChapter > 1)
End Function
Public Function GetNextChEnabled(control As IRibbonControl) As Boolean
    GetNextChEnabled = m_navState(NAV_CHAPTER, BTN_NEXT) And _
        (m_currentChapter < aeBibleCitationClass.ChaptersInBook(m_currentBookIndex))
End Function

' Verse row
Public Function GetPrevVsEnabled(control As IRibbonControl) As Boolean
    GetPrevVsEnabled = m_navState(NAV_VERSE, BTN_PREV) And (m_currentVerse > 1)
End Function
Public Function GetNextVsEnabled(control As IRibbonControl) As Boolean
    GetNextVsEnabled = m_navState(NAV_VERSE, BTN_NEXT) And _
        (m_currentVerse < aeBibleCitationClass.VersesInChapter( _
            m_currentBookIndex, m_currentChapter))
End Function
```

`BOOK_COUNT = 66` is a named constant. `ChaptersInBook` and `VersesInChapter` are
the Public functions from Step 4. No secondary flags; no stale state possible.

---

### Items 5, 6, 7 — Resolution (2026-04-11)

**Decision: Option A — ribbon reflects last navigation action only.**

The ribbon behaves like a search box: it is a navigation tool, not a position
tracker. When the user manually scrolls or clicks outside a navigation sequence,
the ribbon resets to default (all comboBoxes blank, Book row active, Chapter and
Verse rows disabled). This is consistent with standard search bar behaviour — the
bar goes empty once the search is complete.

---

**Item 5 — State drift after VBA reset: resolved by Option A.**

On VBA runtime reset all instance variables are destroyed. Under Option A the ribbon
simply reverts to its default state — blank comboBoxes, Book row active. This is
identical to a fresh document open. No rehydration strategy is needed. The user
re-enters a reference and continues.

---

**Item 6 — Matrix reflects code state, not cursor position: accepted by design.**

Manual navigation outside the ribbon does not update the ribbon state. The ribbon
resets to default. This is the correct behaviour for a navigation tool. The browser
address bar analogy applies: it reflects the last navigation, not the current scroll
position.

**Option B considered and rejected — permanently, including for the Store release.**

`Document_SelectionChange` fires on every cursor movement. Each event requires a
backward scan through `headingData`, a forward Heading 2 scan, and ribbon
invalidation. On a 33,857-paragraph document this runs continuously during reading.
The cost is structural — not fixable by optimisation. A guard against repeated
events helps for stationary cursors but does not reduce cost during genuine reading
navigation (e.g. holding the down arrow through Psalm 119 fires 176 events in
seconds).

More fundamentally, Option B solves a problem Study Bible readers do not have. The
reading pattern is: navigate to a passage via ribbon, read, navigate again. The user
controls position through the ribbon; the ribbon does not need to track them. A
browser address bar does not update as you scroll — no one considers this a defect.

The actual user need Option B addresses ("how do I get back to John 3?") is answered
better by the history list (see below).

---

**Item 7 — Four clean states / same-level re-entry: resolved by Q1 rule.**

Navigating from one book to another (STATE_BOOK_SELECTED → STATE_BOOK_SELECTED)
always resets downstream variables (`m_currentChapter = 0`, `m_currentVerse = 0`)
and re-enters STATE_BOOK_SELECTED cleanly. Chapter and Verse comboBoxes go blank.
No fifth state is needed.

---

### History list — last N searches

The comboBox dropdown is the natural home for a navigation history. After each
confirmed navigation the full reference (e.g. `"John 3"`, `"Psalm 23:1"`) is
prepended to a fixed-length MRU list (suggested N = 10). The Book comboBox dropdown
shows the history list when no text has been typed; typing filters to book names as
normal.

Benefits:
- Answers the "where was I?" need without any position tracking
- Reaches any recent location, not just the current one
- Fully compatible with Option A — no event overhead
- Persists naturally to a document custom property or sidecar file across sessions
- i18n-neutral: stored references use canonical SBL form, displayed as entered

Implementation is deferred — not required for the current development phase.

---

## § 12 — ComboBox Navigation Design (2026-04-11)

### Decision

Replace the three GoTo buttons (column 2 of the state matrix) with `<comboBox>`
controls. The state matrix, row/column 1-based indices, and progressive lock model
are all unchanged. The comboBox occupies `[row, 2]` in each row.

### Layout

```text
              Prev    GoTo (comboBox)         Next
Book:    [1,1] ◀    [1,2] Genesis          ▼  [1,3] ▶
Chapter: [2,1] ◀    [2,2] 1               ▼  [2,3] ▶
Verse:   [3,1] ◀    [3,2] 1               ▼  [3,3] ▶
```

Each row is a horizontal box: `<button> | <comboBox> | <button>`. The state matrix
controls enable/disable the entire row as before. No imageMso needed anywhere in the
navigation group. No screentips needed on comboBox controls — the current value is
self-describing, and removing screentips eliminates one i18n translation surface.

---

### OT / NT separator

A blank item (`""`) returned from `getItemLabel` renders as an empty line in the
dropdown, cleanly dividing Old Testament from New Testament. The empty string is
language-neutral — no translation required in any locale, ever.

Detection uses `getItemID`, not the label text, so it is robust against any future
label change:

```vba
Public Function GetBookItemID(control As IRibbonControl, index As Long) As String
    If index = OT_NT_SEPARATOR_INDEX Then
        GetBookItemID = "SEP"
    Else
        GetBookItemID = CStr(index)
    End If
End Function

Public Sub OnBookChanged(control As IRibbonControl, _
                         selectedId As String, selectedIndex As Long)
    If selectedId = "SEP" Then Exit Sub   ' blank separator row — ignore
    ' ... navigation logic
End Sub
```

`OT_NT_SEPARATOR_INDEX` is a named constant (index 39, after Malachi, before
Matthew). Total item count = 66 books + 1 separator = 67.

---

### Parser integration (Book comboBox)

The Book comboBox feeds directly into the existing Stage-based SBL parser:

- Typing `"Jn"`, `"1 Cor"`, or `"Genesis"` expands via Stage 13 (shorthand) and
  Stage 14 (canonical compression)
- Dropdown selection sets the canonical book name and updates `m_currentBookIndex`
- Navigating with Prev/Next updates the comboBox display via ribbon invalidation

`onChange` fires on every keystroke. Navigate only on `Enter` or confirmed dropdown
selection; discard partial input that does not resolve to a valid book.

---

### Data source

`GetBookLabel` delegates to `aeBibleCitationClass` — the authoritative book list.
No parallel `BookList` array is needed in `aeRibbonClass`. Chapter and Verse item
counts delegate to `ChaptersInBook` and `VersesInChapter` (Step 4, already planned).

---

### Invalidation sequence

Changing the Book comboBox invalidates Chapter, then Verse — in that order — to
avoid circular callbacks. Never invalidate Book from within a Chapter or Verse
callback.

```
Book changed → invalidate Chapter → invalidate Verse
Chapter changed → invalidate Verse
Verse changed → (nothing downstream)
```

---

### Corrections to the design input

| Item | Original | Correction |
|------|----------|-----------|
| Control type | `<dropDown>` shown in sample XML | Use `<comboBox>` — supports free-text input |
| Book data source | `BookList` module-level `Variant` array | Delegate to `aeBibleCitationClass` |
| `CurrentBookIndex` | Separate module variable | Consolidate into `m_currentBookIndex` in `aeRibbonClass` |
| Separator label | `"── New Testament ──"` | Empty string `""` — language-neutral, i18n-free |
| Separator detection | String comparison on label | `getItemID` returning `"SEP"` — robust, never needs changing |
| ComboBox width | Not specified | Add `sizeString="2 Thessalonians"` to reserve width for longest book name |

---

## § 14 — Implementation Steps: Status and Revised Scope (2026-04-11)

### Step status

| Step | Description | Status |
|------|-------------|--------|
| 1 | `NormalizeBookInput` + `m_currentBookIndex` / `m_currentBookPos` in `aeRibbonClass.cls` | **NEXT — scope revised, see below** |
| 2 | `CaptureHeading2s` | Eliminated |
| 3 | `GoToChapter` implementation | Pending |
| 4 | Expose `ChaptersInBook` / `VersesInChapter` as Public | Pending |
| 5 | `GoToVerse` implementation | Pending |
| 6 | Ribbon XML update | Pending — requires full rewrite for `<comboBox>` row layout, removal of screentips and imageMso |
| 7 | Move OLD_CODE | Pending |
| 8 | `normalize_vba.py` update | Pending |

### Note on scope change

Steps 1–5 were designed around the original GoTo-button-plus-InputBox model.
The comboBox design (§ 12) changes the callback signatures and removes the need
for InputBox dialogs entirely. Step 1 in particular needs revisiting:
`NormalizeBookInput` is still valid but it now feeds `OnBookChanged` rather than
a standalone `GoToH1` sub. Revised Step 1 scope to be confirmed before any code
is written.

---

### Revised Step 1 — confirmed scope (2026-04-11)

**What survives from the original plan:**

`NormalizeBookInput` is unchanged in purpose and logic — it cleans raw text input
before matching against book names. It remains a private helper in `aeRibbonClass.cls`.

```vba
Private Function NormalizeBookInput(ByVal raw As String) As String
    Dim s As String
    s = Trim(UCase(raw))
    If Len(s) >= 2 Then
        If s Like "[0-9][A-Z]*" Then s = Left$(s, 1) & " " & Mid$(s, 2)
    End If
    NormalizeBookInput = s
End Function
```

**What changes:**

The original plan called for a `GoToH1` sub that called an InputBox, normalised the
input, scanned `headingData`, and navigated. That sub is replaced by the comboBox
callback chain:

```
User types / selects in Book comboBox
    → OnBookChanged fires
        → NormalizeBookInput cleans the text
        → matched against headingData
        → m_currentBookIndex set
        → m_currentChapter = 0, m_currentVerse = 0
        → document navigates to book H1
        → SetNavState STATE_BOOK_SELECTED
        → invalidate Chapter and Verse comboBoxes
```

**New instance variables:**

| Variable | Type | Purpose |
|----------|------|---------|
| `m_currentBookIndex` | `Long` | 1-based index into headingData; 0 = not set |
| `m_currentBookPos` | `Long` | Character position of book H1 in document |
| `m_currentChapter` | `Long` | Current chapter number; 0 = not set |
| `m_currentVerse` | `Long` | Current verse number; 0 = not set |

`m_currentChapter` and `m_currentVerse` are declared here even though not used until
Steps 3 and 5 — they are part of the same reset block and must exist before
`OnBookChanged` can zero them.

**New public callbacks (ribbon-facing):**

| Callback | Purpose |
|----------|---------|
| `OnBookChanged` | Fires on comboBox text change or selection |
| `GetBookText` | Returns current book name for comboBox display |
| `GetBookCount` | Returns 67 (66 books + 1 separator) |
| `GetBookItemLabel` | Returns book name or `""` for separator |
| `GetBookItemID` | Returns index string or `"SEP"` for separator |

**`GoToH1` / `GoToH1Direct` / `GoToH1Deferred`:**

These become internal — called from `OnBookChanged` rather than directly from a
ribbon button. Their signatures do not change; only the call site changes.

---

## § 15 — Test Plan: basTEST_aeRibbonClass (2026-04-11)

### Module

`src\basTEST_aeRibbonClass.bas`

Follows the same pattern as `basTEST_aeBibleCitationClass.bas`:
- `Option Private Module` — only the runner appears in Alt+F8
- Shared `aeAssert` public variable
- Log written to `rpt\Ribbon_Tests.UTF8.txt`
- Tests added to runner as each step is implemented

### Entry point

```vba
Public Sub Run_All_Ribbon_Tests()
```

### Test groups

Tests are divided into two tiers. **Headless** tests require no open document and
run against logic only. **Document** tests require the Bible `.docm` to be open and
navigate the live document.

---

#### Group 1 — NormalizeBookInput (headless, Step 1)

Added to runner after Step 1.

| Test | Input | Expected |
|------|-------|----------|
| Already normalised | `"GENESIS"` | `"GENESIS"` |
| Lowercase | `"genesis"` | `"GENESIS"` |
| Prefix no space | `"1John"` | `"1 JOHN"` |
| Prefix already spaced | `"1 John"` | `"1 JOHN"` |
| Leading/trailing spaces | `"  Mark  "` | `"MARK"` |
| Single character | `"J"` | `"J"` |
| Two-digit prefix guard | `"2Tim"` | `"2 TIM"` |

---

#### Group 2 — Book item list (headless, Step 1)

Added to runner after Step 1.

| Test | Expected |
|------|----------|
| `GetBookCount` | 67 (66 books + 1 separator) |
| Item at index 1 | `"Genesis"` |
| Item at `OT_NT_SEPARATOR_INDEX` label | `""` |
| Item at `OT_NT_SEPARATOR_INDEX` ID | `"SEP"` |
| Item at index 67 (last) | `"Revelation"` |
| No item label is `"SEP"` (separator ID only) | True |

---

#### Group 3 — State matrix transitions (headless, Step 1)

Added to runner after Step 1.

| Test | Action | Expected |
|------|--------|----------|
| Default state | Initial | Book GoTo ON; all others OFF |
| STATE_BOOK_SELECTED | `SetNavState` | Book all ON; Chapter GoTo ON; Verse all OFF |
| STATE_CHAPTER_SELECTED | `SetNavState` | Book all OFF; Chapter all ON; Verse GoTo ON |
| STATE_VERSE_SELECTED | `SetNavState` | Chapter/Book all OFF; Verse all ON |
| New Search reset | `OnNewSearchButtonClick` | All position vars = 0; STATE_DEFAULT |

---

#### Group 4 — Boundary expressions (headless, Steps 1–5)

Tests added incrementally as each step introduces its position variable.

| Test | Condition | Expected |
|------|-----------|----------|
| Prev Book at Genesis | `m_currentBookIndex = 1` | `GetPrevBkEnabled` = False |
| Next Book at Revelation | `m_currentBookIndex = 66` | `GetNextBkEnabled` = False |
| Prev/Next Book mid-range | `m_currentBookIndex = 33` | Both True |
| Prev Chapter at 1 | `m_currentChapter = 1` | `GetPrevChEnabled` = False |
| Next Chapter at max | `m_currentChapter = ChaptersInBook` | `GetNextChEnabled` = False |
| Prev Verse at 1 | `m_currentVerse = 1` | `GetPrevVsEnabled` = False |
| Next Verse at max | `m_currentVerse = VersesInChapter` | `GetNextVsEnabled` = False |

---

#### Group 5 — Book navigation (document, Step 1)

Added to runner after Step 1. Requires Bible `.docm` open.

| Test | Input | Expected |
|------|-------|----------|
| Full name | `"Genesis"` | `m_currentBookIndex = 1`; cursor at Genesis H1 |
| Abbreviation | `"Jn"` | `m_currentBookIndex` = John's index |
| Shorthand prefix | `"1cor"` | `m_currentBookIndex` = 1 Corinthians index |
| Invalid input | `"Zzz"` | No navigation; position unchanged |
| Separator selected | ID = `"SEP"` | No navigation; position unchanged |
| Book re-entry resets chapter | GoTo Genesis, then GoTo Exodus | `m_currentChapter = 0` |

---

#### Group 6 — Chapter navigation (document, Step 3)

Added to runner after Step 3.

| Test | Input | Expected |
|------|-------|----------|
| Valid chapter | Book = Genesis, Chapter = 3 | `m_currentChapter = 3`; cursor at H2 |
| Chapter 1 | Book = Genesis, Chapter = 1 | `m_currentChapter = 1` |
| Max chapter | Book = Psalms, Chapter = 150 | `m_currentChapter = 150` |
| Over max | Book = Genesis, Chapter = 51 | No navigation; position unchanged |
| Chapter re-entry resets verse | GoTo Ch 3 then GoTo Ch 5 | `m_currentVerse = 0` |

---

#### Group 7 — Verse navigation (document, Step 5)

Added to runner after Step 5.

| Test | Input | Expected |
|------|-------|----------|
| Study version — paragraph count | Book = Jude, Ch = 1, Verse = 3 | Cursor at verse 3 paragraph |
| Worst case | Book = Psalms, Ch = 119, Verse = 176 | Cursor at verse 176; elapsed logged |
| Over max | Book = Jude, Ch = 1, Verse = 26 | No navigation; position unchanged |

---

### Runner skeleton

```vba
Public Sub Run_All_Ribbon_Tests()
    On Error GoTo PROC_ERR

    Dim log As New aeLoggerClass
    log.Log_Init ActiveDocument.Path & "\rpt\Ribbon_Tests.UTF8.txt"

    Set aeAssert = New aeAssertClass
    aeAssert.SetLogger log
    aeAssert.Initialize

    ' --- Headless (no document required) ---
    Test_NormalizeBookInput        ' Step 1
    Test_BookItemList              ' Step 1
    Test_StateMatrixTransitions    ' Step 1
    Test_BoundaryExpressions       ' Steps 1-5

    ' --- Document required ---
    Test_BookNavigation            ' Step 1
    Test_ChapterNavigation         ' Step 3
    Test_VerseNavigation           ' Step 5

    aeAssert.Terminate
    Set aeAssert = Nothing
    log.Log_Close
    Set log = Nothing

PROC_EXIT:
    Exit Sub
PROC_ERR:
    If Not log Is Nothing Then log.Log_Close
    MsgBox "Erl=" & Erl & " Error " & Err.Number & _
           " (" & Err.Description & ") in Run_All_Ribbon_Tests"
    Resume PROC_EXIT
End Sub
```

Document tests are stubbed with `Debug.Print "SKIP: ..."` until the implementation
step that enables them is complete. This keeps the runner green at every step.

---

### VSTO / i18n / Code-signing Alignment

#### Preparation steps

These are disciplines and structural decisions applied during VBA development now.
None require new technology. Each is listed with its cost to apply now versus the
cost to retrofit later.

---

**P1 — String resource module**

Create `basRibbonStrings.bas`. All user-facing strings used by the ribbon — status
bar messages, MsgBox text, error messages, `screentip` values if any remain — are
declared as named constants or returned from a single function here. Logic
procedures never contain inline string literals for UI text.

In VBA: a `bas` module with `Public Const` declarations.
In VB.NET VSTO: a `.resx` resource file. The port is a file replacement, not a
code audit.

i18n benefit: a translator works on one file. No code changes required.

| | Cost |
|-|------|
| Now | 1–2 hours to create the module and establish the discipline during Step 1 |
| Later | Full audit of all modules for inline strings, regression test of every message path |
| Risk if deferred | Strings missed in audit, UI text untranslated in one or more locales |

Test addition to Group 1 (headless): `Test_StringResourceCoverage` — asserts that
no `MsgBox`, `Application.StatusBar`, or `Debug.Print` call in `aeRibbonClass.cls`
contains a string literal directly. Implemented as a grep-style scan of the module
text at runtime via `VBComponent.CodeModule.Lines`.

---

**P2 — Callback signature discipline**

All ribbon callback subs and functions must use `Long` not `Integer` for index
parameters. VBA `Integer` is 16-bit; VB.NET `Integer` is 32-bit (= VBA `Long`).
Using `Long` throughout means the port is type-name-compatible without change.

```vba
' Correct — ports directly to VB.NET Integer
Public Function GetBookItemLabel(control As IRibbonControl, index As Long) As String

' Wrong — requires type change on port
Public Function GetBookItemLabel(control As IRibbonControl, index As Integer) As String
```

VSTO ribbon callback signatures are identical in structure to VBA. A correctly
typed VBA callback is a copy-paste with namespace changes only.

| | Cost |
|-|------|
| Now | Zero — a naming rule applied during writing |
| Later | Audit of all 20+ callbacks; risk of runtime errors on large indices at port |
| Risk if deferred | Silent overflow on book/chapter/verse indices > 32767 (not a real risk for 66 books, but a correctness issue in principle) |

Test addition to runner: `Test_CallbackSignatures` — a checklist sub that prints
each callback name and its parameter types. Not an automated assertion; a
confirmation print reviewed before each Step is marked complete.

---

**P3 — Separation of concerns verification**

Each ribbon callback sub must contain no navigation logic directly — it calls a
private method on `aeRibbonClass` and returns. The private method contains the
logic. This maps directly to the VSTO pattern where the ribbon callback class
delegates to a service class.

Test addition (headless): `Test_CallbackDelegation` — for each public callback,
assert that its line count is ≤ 5 lines (stub + one call + error handler).
Implemented via `VBComponent.CodeModule` line count inspection.

| | Cost |
|-|------|
| Now | Zero — a structural rule applied during writing |
| Later | Full refactor of callback bodies into private methods; high regression risk |
| Risk if deferred | VSTO port requires rewriting every callback, not just retyping |

---

**P4 — i18n: book name source discipline**

Book names displayed in the comboBox must always come from `aeBibleCitationClass`
via `GetBookItemLabel`. No book name string literal anywhere in `aeRibbonClass.cls`
or `basTEST_aeRibbonClass.bas`. Test cases that reference a specific book use the
index constant (`NAV_BOOK_GENESIS = 1`) not the string `"Genesis"`.

i18n benefit: localising book names requires updating `aeBibleCitationClass` only.
The ribbon, the tests, and the navigation logic are all index-based and require no
change.

| | Cost |
|-|------|
| Now | Near zero — a naming discipline during test writing |
| Later | Test suite audit; risk of tests passing in English but failing in localised builds |

---

**P5 — Numeric canonical form for navigation state**

Navigation state is always stored and passed as indices (`m_currentBookIndex`,
`m_currentChapter`, `m_currentVerse`). String forms are only produced at the
display layer (`GetBookText`, `GetChapterText`, `GetVerseText`). This is already
the design; this preparation step makes it an explicit tested constraint.

Test addition (headless): assert that after every navigation action, all three
position variables hold numeric values > 0 (or exactly 0 for unset). No string
state exists anywhere in the navigation model.

---

#### Impact summary

| Preparation | Implementation cost now | Retrofit cost later | Deferred risk |
|-------------|------------------------|--------------------|-|
| P1 String resources | Low (1–2 hrs) | High (audit + regression) | Untranslated UI strings |
| P2 Callback signatures | Zero (naming rule) | Low–medium (type audit) | Port friction |
| P3 Separation of concerns | Zero (structural rule) | High (full refactor) | VSTO callbacks need rewrite |
| P4 Book name discipline | Near zero | Low (test audit) | Tests fail in localised builds |
| P5 Numeric canonical form | Zero (already designed) | Medium (state refactor) | i18n breakage in display layer |

The combined cost of all five preparations applied now is **2–3 hours** above the
baseline implementation time. The combined retrofit cost later is estimated at
**3–5 days** of refactoring, audit, and regression testing, with non-trivial risk
of introducing defects into working code.

---

#### Code signing — when it becomes useful

| Stage | What to sign | When | Why |
|-------|-------------|------|-----|
| **Now — development** | VBA project in `.docm` | Immediately | Prevents "Macros disabled" prompts on the development machine; `Application.OnTime` callbacks resolve without Trust Center intervention |
| **First external distribution** | VBA project in `.docm` | Before sharing with any tester outside the development machine | Word shows a security warning for unsigned macros; testers will be blocked or alarmed without a signature |
| **VSTO development begins** | VSTO assembly (`.dll`) | When the first VSTO build is produced | Windows SmartScreen and Office add-in registration both check for a valid Authenticode signature; unsigned VSTO add-ins require manual Trust Center override on each machine |
| **Store submission** | MSIX package | Mandatory before submission | Microsoft Store requires a valid code-signing certificate; EV (Extended Validation) certificate required for kernel-mode drivers but standard OV (Organisation Validation) is sufficient for MSIX Office add-ins |

**Certificate recommendation:** obtain a standard OV Authenticode certificate now
(DigiCert, Sectigo, or equivalent, ~USD 200–400/year). Use it to sign the VBA
project immediately. The same certificate signs the VSTO assembly and the MSIX
package when those stages are reached. Establishing the signing workflow early
means no disruption to the distribution pipeline later.

An EV certificate (~USD 400–700/year, requires hardware token) is not required
unless kernel-mode code is involved. It is not needed for this project.

---

## § 16 — Font Licensing and Management (2026-04-11)

### Current state — what the audit found

Font code exists but is inconsistent and scattered across four modules:

| Module | What it does | Problem |
|--------|-------------|---------|
| `basTEST_aeBibleFonts.bas` | Checks availability of 5 Google fonts; per-style audit + redefine subs; Arial Unicode MS scan | `IsFontInstalled` creates a full document to test each font — extremely expensive. No fallback chain. Per-style subs are one-offs with no shared logic. |
| `basAuditDocument.bas` | `FindFontUsage` searches paragraphs + styles for a target font | Standalone utility; no connection to font management strategy |
| `aeBibleClass.cls` | Hardcoded `"Arial Black"`, `"Calibri"`, `"Liberation Sans Narrow"` inline in style application logic | Font names embedded in logic — no single source of truth |
| `basWordRepairRunner.bas` | `RGB(255, 165, 0)` / `RGB(80, 200, 120)` inline for marker colour detection | Colour values coupled to font presentation, not font identity |

---

### Licensing audit — fonts currently in use

| Font | Owner | License | Distributable? |
|------|-------|---------|---------------|
| Arial Black | Monotype / Microsoft | Proprietary | No — bundled with Windows; not for redistribution |
| Calibri | Microsoft | Proprietary | No — Office/Windows license only; embedding restricted |
| Times New Roman | Monotype / Microsoft | Proprietary | No — same restrictions |
| Arial Unicode MS | Microsoft | Proprietary | No — Office license only |
| Liberation Sans Narrow | Red Hat | SIL OFL | Yes — free for all use including distribution |
| Noto Sans | Google | SIL OFL | Yes — already targeted as replacement in existing code |

**Distribution risk:** any `.docm` or `.docx` distributed to users with embedded
proprietary fonts is in violation of the font vendor's EULA under an Office 365
Individual/Family licence. Print use is separately licensed under corporate
agreements that do not extend to digital distribution.

---

### Recommended free font replacements

| Role | Current font | Primary replacement | Fallback |
|------|-------------|--------------------|-|
| Body text | Times New Roman | **Gentium Plus** (SIL OFL) | Linux Libertine G |
| Headings | Arial Black | **Source Serif 4** (SIL OFL) | EB Garamond |
| UI / captions / footnotes | Calibri | **Lato** (SIL OFL — see Calibri note) | Nunito Sans |
| Verse / Chapter markers | Arial Black | **Noto Sans** (SIL OFL — already targeted) | Liberation Sans Narrow |
| Biblical Greek / Hebrew | (none currently) | **SBL Greek / SBL Hebrew** (SBL free licence) | Gentium Plus (Greek); Noto Serif Hebrew |

**Gentium Plus** is the primary body font recommendation. It was designed
specifically for scholarly and biblical text, has broad diacritic coverage for
i18n, and is a direct visual replacement for Times New Roman at matching point
sizes. SIL OFL permits embedding in distributed documents without restriction.

**Noto Sans** is already the target in existing code for footnote and caption
styles. Extending it to verse and chapter marker roles achieves consistency.

**SBL fonts** are free for non-commercial use under the SBL licence. If commercial
distribution is planned, confirm licence terms at the point of Store submission.
For the current development phase they are unrestricted.

#### Calibri — no metrically identical free replacement exists

This is a hard constraint that affects layout planning.

Calibri was designed for Microsoft's ClearType rendering system and is protected by
licensing that prevents metric-compatible clones. No open-licensed font replicates
its exact glyph widths, kerning pairs, or line metrics. Any substitution will cause
document reflow — pagination, line breaks, table widths, and UI alignment will all
shift.

The closest free alternatives are visually similar but not metrically identical:

| Font | Character | License |
|------|-----------|---------|
| **Lato** | Clean, modern, humanist sans-serif | SIL OFL |
| **Nunito Sans** | Rounded, friendly, highly readable | SIL OFL |
| **Work Sans** | Modern, simple, good for UI | SIL OFL |
| Open Sans, Source Sans Pro, PT Sans | Common Calibri-adjacent choices | SIL OFL |

**Practical guidance for this project:**

- **UI roles** (ribbon labels, captions, footnotes): reflow is not a concern —
  use **Lato** as primary. Visual similarity is sufficient; pixel-perfect layout
  is not required for these elements.
- **Body text**: Calibri is not used for body text in this document (Times New
  Roman / Gentium Plus are the body fonts). No reflow risk here.
- **Print candidate version**: if any styles use Calibri for body or table content,
  the print layout *will* reflow on substitution. A layout review pass is required
  after font substitution before the print candidate is finalised. This is a known,
  accepted cost — it is better to discover and fix layout shifts during development
  than after distribution.
- **Study version** (33,857 paragraphs, one verse per paragraph): reflow is
  functionally irrelevant — the document is read on screen and navigation is
  position-based, not page-based.

The font manager's `frUI` stack reflects this:

```vba
Case frUI
    stack = Array("Lato", "Nunito Sans", "Work Sans", "Calibri")
```

Calibri remains as the last-resort fallback so the document degrades gracefully on
machines where none of the free alternatives are installed. It is never the
preferred choice.

---

### aeFontManagerClass — design

A new class `aeFontManagerClass.cls` centralises all font decisions.

**Responsibilities:**
- Check font availability using `Application.FontNames` (not document creation —
  the current `IsFontInstalled` approach of opening a temp document is discarded)
- Define named font stacks (primary + fallback chain per role)
- Return the best available font for a given role at runtime
- Apply font stacks to document styles in a single batch operation
- Report which fonts are in use and their licence status

**Font stack pattern:**

```vba
' aeFontManagerClass
Public Enum FontRole
    frBody = 1
    frHeading = 2
    frUI = 3
    frVerseMarker = 4
    frBiblicalGreek = 5
    frBiblicalHebrew = 6
End Enum

Public Function BestAvailable(ByVal role As FontRole) As String
    Dim stack As Variant
    Select Case role
        Case frBody
            stack = Array("Gentium Plus", "Linux Libertine G", "Times New Roman")
        Case frHeading
            stack = Array("Source Serif 4", "EB Garamond", "Arial Black")
        Case frUI
            stack = Array("Noto Sans", "Liberation Sans", "Calibri")
        Case frVerseMarker
            stack = Array("Noto Sans", "Liberation Sans Narrow", "Arial Black")
        Case frBiblicalGreek
            stack = Array("SBL Greek", "Gentium Plus", "Times New Roman")
        Case frBiblicalHebrew
            stack = Array("SBL Hebrew", "Noto Serif Hebrew", "Arial Unicode MS")
    End Select
    Dim i As Long
    For i = LBound(stack) To UBound(stack)
        If IsFontAvailable(CStr(stack(i))) Then
            BestAvailable = CStr(stack(i))
            Exit Function
        End If
    Next i
    BestAvailable = ""   ' no font in stack is installed — caller must handle
End Function

Private Function IsFontAvailable(ByVal fontName As String) As Boolean
    Dim f As Variant
    For Each f In Application.FontNames
        If StrComp(CStr(f), fontName, vbTextCompare) = 0 Then
            IsFontAvailable = True
            Exit Function
        End If
    Next f
End Function
```

**ApplyToStyles** iterates all document styles and replaces any font name not in
the free-font list with `BestAvailable` for the appropriate role. This replaces the
scattered per-style `Redefine*` subs in `basTEST_aeBibleFonts.bas`.

---

### Consolidation of existing font code

| Existing sub | Action |
|-------------|--------|
| `IsFontInstalled` | Replace with `aeFontManagerClass.IsFontAvailable` using `Application.FontNames` |
| `CheckOpenFontsWithDownloads` | Replace with `aeFontManagerClass.ReportAvailability` — logs all stacks and best available for each role |
| Per-style `Redefine*` subs (×3) | Replace with `aeFontManagerClass.ApplyToStyles` — one call covers all styles |
| Per-style `AuditStyleUsage_*` subs (×3) | Consolidate into one `AuditFontUsage` sub using `FindFontUsage` pattern from `basAuditDocument` |
| `Identify_ArialUnicodeMS_Paragraphs` | Generalise to `AuditInlineFontOverrides` — reports all paragraphs where inline font differs from style definition, for any target font |
| `CreateEmphasisBlackStyle` | Update to call `BestAvailable(frHeading)` instead of hardcoding `"Arial Black"` |
| Inline `"Arial Black"` / `"Calibri"` in `aeBibleClass.cls` | Replace with `aeFontManagerClass.BestAvailable(role)` calls |

---

### Test additions to basTEST_aeRibbonClass

**Group 8 — Font manager (headless)**

| Test | Expected |
|------|----------|
| `IsFontAvailable` on installed font | True |
| `IsFontAvailable` on non-existent font | False |
| `BestAvailable(frUI)` | Returns `"Noto Sans"` if installed; `"Liberation Sans"` if not; never `""` unless all fail |
| `BestAvailable` with empty stack | Returns `""` — caller handles gracefully |
| No style in document uses a proprietary font after `ApplyToStyles` | True |
| `ApplyToStyles` is idempotent — running twice produces same result | True |

---

### i18n implications

`BestAvailable` is locale-unaware by design — it returns the best installed font
for a role regardless of language. For i18n builds:

- Latin-script languages: Gentium Plus and Noto Sans both have broad diacritic
  coverage — no change to the font stack required
- Right-to-left scripts (Hebrew, Arabic): `frBiblicalHebrew` stack already includes
  Noto Serif Hebrew; extend with Noto Naskh Arabic if Arabic support is needed
- East Asian scripts: add Noto Serif CJK / Noto Sans CJK to a new `frCJK` role

The font manager is the single point of change for any new script. No other code
requires modification.

---

### Cost-benefit: now versus later

| Work item | Cost now | Cost later | Risk if deferred |
|-----------|----------|-----------|-----------------|
| Create `aeFontManagerClass` | Medium (4–6 hrs) | High (refactor all scattered font code under time pressure before distribution) | Proprietary fonts shipped to users; licence violation |
| Replace `IsFontInstalled` | Low (30 min) | Low | Performance issue only — each call opens a Word document |
| Consolidate per-style audit/redefine subs | Low (1–2 hrs) | Low–medium | Duplicate maintenance burden grows with each new style |
| Replace inline font literals in `aeBibleClass.cls` | Low (1 hr) | Medium (more inline literals added during ribbon development) | Hardcoded proprietary fonts persist into distributed build |
| Add font tests to runner | Low (1 hr alongside class creation) | Medium (tests must be written from scratch after the fact) | No automated verification before distribution |

**Combined cost now: ~7–9 hours.**
**Combined cost later: estimated 2–3 days plus legal exposure if distribution
proceeds before the work is done.**

**Recommended timing:** implement `aeFontManagerClass` and consolidate existing
font code as a discrete task before ribbon Step 1 begins. It is a clean,
self-contained unit of work with no dependency on the ribbon implementation. Doing
it first also means the ribbon's own style application calls use the font manager
from the start — no retrofit needed there.

---
