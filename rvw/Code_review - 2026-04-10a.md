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
