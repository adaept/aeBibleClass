# Code Review: Interactive Citation Block Repair Procedure

**Date:** 2026-04-04

---

## Overview

A Word VBA procedure that lets the user repair a citation block in-place within a
document paragraph. The procedure:

1. Prompts the user to confirm they want to run the repair on the current paragraph
2. Selects all text in the paragraph and passes it to an interactive block validation test
3. Validates the block, displaying all errors to the user
4. Prompts the user to manually fix errors and re-run validation until the block is clean
5. Sorts the validated block into canonical order
6. Renders the result with en-dash range separators and copies it to the clipboard
7. Prompts the user to replace the original selection with the corrected version

Line feeds for page formatting are left to the user to insert manually after replacement.

---

## Entry Point

```vb
Public Sub RepairCitationBlockInParagraph()
```

Placed in a standard module (e.g. `basRepairCitationBlock.bas`). Called from a toolbar
button or keyboard shortcut while the cursor is anywhere inside the target paragraph.

---

## Task 1 — Confirm Intent (Default No)

### Problem

The procedure modifies document content. An accidental trigger must not silently alter
the paragraph.

### Solution

```vb
Dim answer As VbMsgBoxResult
answer = MsgBox("Repair citation block in the current paragraph?", _
                vbYesNo + vbDefaultButton2 + vbQuestion, _
                "Repair Citation Block")
If answer <> vbYes Then Exit Sub
```

`vbDefaultButton2` makes **No** the default so pressing Enter does nothing. The user
must explicitly click **Yes** to proceed.

---

## Task 2 — Capture the Paragraph Text

### Problem

The procedure needs the full raw text of the paragraph that contains the cursor, without
the trailing paragraph mark, to pass to the parser.

### Solution

```vb
Dim para As Paragraph
Set para = Selection.Paragraphs(1)
Dim rawBlock As String
rawBlock = para.Range.Text
' Strip trailing paragraph mark (Chr(13)) if present
If Right$(rawBlock, 1) = Chr(13) Then
    rawBlock = Left$(rawBlock, Len(rawBlock) - 1)
End If
```

The paragraph range is saved for the replacement step (Task 6).

---

## Task 3 — Interactive Validation Loop

### Problem

`ParseCitationBlock` may find errors. The user must see them and be able to fix the
source text before the procedure continues.

### Solution

Call `VerifyCitationBlock` in a loop. On each iteration, display the full validation
report in a message box. If errors remain, prompt the user to fix the paragraph manually
and try again. The loop exits only when all tokens pass or the user cancels.

```vb
Dim verified As Boolean
verified = False
Do
    ' Re-read the paragraph text on each iteration (user may have edited it)
    rawBlock = para.Range.Text
    If Right$(rawBlock, 1) = Chr(13) Then
        rawBlock = Left$(rawBlock, Len(rawBlock) - 1)
    End If

    Dim report As String
    Dim passCount As Long, failCount As Long
    report = aeBibleCitationClass.VerifyCitationBlock(rawBlock, passCount, failCount)

    If failCount = 0 Then
        verified = True
        Exit Do
    End If

    Dim retry As VbMsgBoxResult
    retry = MsgBox(report & vbCrLf & vbCrLf & _
                   "Fix the errors above in the paragraph, then click Retry." & vbCrLf & _
                   "Click Cancel to abort.", _
                   vbRetryCancel + vbExclamation, _
                   "Citation Block Errors (" & failCount & " failed)")
    If retry <> vbRetry Then Exit Sub
Loop
```

`VerifyCitationBlock` must accept `rawBlock` as input and return the formatted report
string plus `passCount`/`failCount` by reference. If the current signature differs,
a thin wrapper is acceptable.

---

## Task 4 — Sort into Canonical Order

### Problem

After validation succeeds, the block may still be in thematic or arbitrary order.
Study Bible output requires canonical book order (Gen=1 … Rev=66), then chapter, then
start verse.

### Solution

```vb
Dim items As Collection
Set items = aeBibleCitationClass.SortCitationBlock( _
    aeBibleCitationClass.ParseCitationBlock(rawBlock))
```

`SortCitationBlock` is already implemented (see Code_review - 2026-04-04.md, Task 2).

---

## Task 5 — Render En-Dash and Copy to Clipboard

### Problem

Canonical strings use ASCII hyphen for ranges. The Study Bible document requires en-dash
(`–`, ChrW(8211)) for range separators.

### Solution

Build the result string by calling `RenderEnDash` on each item, joined by `"; "`,
then copy to the clipboard via a temporary `DataObject`.

```vb
Dim result As String
Dim item As Variant
For Each item In items
    If Len(result) > 0 Then result = result & "; "
    result = result & aeBibleCitationClass.RenderEnDash(CStr(item))
Next item

' Copy to clipboard (late binding — no added reference required)
Dim dataObj As Object
Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
dataObj.SetText result
dataObj.PutInClipboard
Set dataObj = Nothing
```

`RenderEnDash` is already implemented (see Code_review - 2026-04-04.md, Task 1).

**Note:** The CLSID `{1C3B4210-F441-11CE-B9EA-00AA006B1A69}` is the `DataObject` class
from Microsoft Forms 2.0. `CreateObject("new:{...}")` instantiates it without adding a
project reference.

---

## Task 6 — Prompt to Replace Selection

### Problem

The corrected block is on the clipboard. The procedure must not silently overwrite the
paragraph; the user must confirm before the original text is replaced.

### Solution

```vb
Dim replace As VbMsgBoxResult
replace = MsgBox("Corrected block copied to clipboard:" & vbCrLf & vbCrLf & _
                 result & vbCrLf & vbCrLf & _
                 "Replace the original paragraph text with the corrected version?", _
                 vbYesNo + vbDefaultButton1 + vbQuestion, _
                 "Replace Citation Block")
If replace = vbYes Then
    para.Range.Text = result
    ' Reposition cursor to end of replaced range
    para.Range.Select
    Selection.Collapse wdCollapseEnd
End If
```

If the user clicks **No**, the corrected text remains on the clipboard and the paragraph
is unchanged. The user can paste manually at their discretion.

**Note on line feeds:** `para.Range.Text` replaces the full paragraph text. Any line
feed characters needed for page formatting must be inserted by the user after replacement;
the procedure does not add them.

---

## Implementation Notes

- Standard `On Error GoTo PROC_ERR` / `PROC_EXIT` / `PROC_ERR` / `MsgBox` / `Resume PROC_EXIT`
  handler wraps the entire procedure.
- `para.Range` is captured once before the validation loop. On each retry, only `rawBlock`
  is re-read from the paragraph; the range reference remains valid as long as the
  paragraph exists.
- If `para` becomes invalid during the loop (e.g. user deletes the paragraph), the error
  handler fires and the procedure exits cleanly.
- No new class methods are required. All parsing, sorting, and rendering are delegated to
  `aeBibleCitationClass`.

---

## Implementation Order

1. Create `basRepairCitationBlock.bas` with `RepairCitationBlockInParagraph`
2. Verify `VerifyCitationBlock` signature — add `passCount`/`failCount` ref params if needed
3. Assign procedure to a toolbar button or shortcut key
4. Manual test: cursor in a paragraph with a known-good citation block — verify Yes/No
   prompt, validation pass, sort, en-dash render, clipboard copy, replacement prompt
5. Manual test: cursor in a paragraph with a known-bad citation block — verify error report
   displays, retry loop works, cancel exits cleanly

---

## Implementation Complete — 2026-04-04

### Files changed

**`src/basTEST_aeBibleCitationBlock.bas`**

- `VerifyCitationBlockReport` (new Public, end of file) — same validation logic as
  `VerifyCitationBlock` but builds and returns a String report; `passCount` and
  `failCount` returned ByRef; suitable for MsgBox display in `RepairCitationBlockInParagraph`
- `RepairCitationBlockInParagraph` — (new) full interactive repair procedure; see tasks below

---

## Goal State After Implementation

| Item | Purpose |
|---|---|
| `RepairCitationBlockInParagraph` | Full interactive repair procedure |

### Dependencies (existing, no changes required)

| Item | Used for |
|---|---|
| `aeBibleCitationClass.VerifyCitationBlock` | Validation and error reporting |
| `aeBibleCitationClass.ParseCitationBlock` | Parse validated raw block |
| `aeBibleCitationClass.SortCitationBlock` | Sort into canonical order |
| `aeBibleCitationClass.RenderEnDash` | En-dash rendering |
| `aeBibleCitationClass.ToSBLShortForm` | Canonical name → SBL abbreviation |

---

## Post-Implementation Fixes — 2026-04-04

### Fix — `Chr(11)` forced line break not normalized (`NormalizeRawInput`)

**Symptom:** Psalms, Jeremiah, Romans, and 2 Peter references were attributed to the
wrong book (1 Chronicles, Isaiah, John, 1 Peter respectively). References like
`Ps 19:1-2` were parsed as `1 Chronicles 29:1-2`.

**Cause:** The paragraph contained Word forced line breaks (`Chr(11)`, Shift+Enter).
`NormalizeRawInput` handles `vbCr`/`vbLf`/`vbCrLf` but not `Chr(11)`. After
`Split(normalized, ";")`, segments beginning after a line break position started with
`Chr(11)` as a prefix on the book abbreviation. `Trim$` does not strip `Chr(11)`, so
`parts(0)` became e.g. `Chr(11) & "Ps"`. The `Like "[A-Za-z]*"` test in Case 2 of
`ParseCitationBlock` failed, falling through to Case 3 (no new book), leaving the book
context unchanged from the previous segment.

**Fix:** Added `s = Replace(s, Chr(11), " ")` to `NormalizeRawInput` alongside the
existing line-break replacements.

---

### Fix — Output used full canonical book names instead of SBL abbreviations

**Symptom:** Clipboard/replacement output showed `Psalms 19:1–2` instead of `Ps 19:1–2`.

**Cause:** `RenderEnDash` operates on the canonical string as-is; no abbreviation step
existed.

**Fix:** Added `ToSBLShortForm` to `aeBibleCitationClass.cls` (after `SortCitationBlock`).
Maps BookID → SBL abbreviation via `Select Case` (all 66 books); falls back to canonical
name if BookID is unrecognised. Task 5 now calls
`RenderEnDash(ToSBLShortForm(CStr(item)))`.

---

### Fix — Repeated book names in output

**Symptom:** Output listed `Ps 19:1–2; Ps 23:1; Ps 28:7; ...` — book name repeated for
every entry of the same book.

**Fix:** Task 5 loop now tracks `prevBook` (canonical book name). When the book is the
same as the previous item, only the numeric part (`ch:verse[-end]`) is emitted; the SBL
abbreviation is emitted only on book change.

---

### Fix — Paragraph mark deleted on replacement (`para.Range.Text = Result`)

**Symptom:** Pasting the corrected block deleted the paragraph mark, merging the
paragraph with the next and losing the next paragraph's formatting.

**Cause:** `para.Range` includes the trailing `Chr(13)` paragraph mark. Assigning
`para.Range.Text = Result` replaced the mark with the new text (which has no mark).

**Fix:** Changed Task 2 to capture a `workRng As Object` before the confirm dialog,
branching on `Selection.Type`:
- `wdSelectionNormal` (text selected) → `workRng = Selection.Range` (selection never
  includes the paragraph mark)
- Cursor only → `workRng = Selection.Paragraphs(1).Range` with
  `workRng.End = workRng.End - 1` to exclude the mark

Task 6 replacement is now simply `workRng.Text = Result`. This also fixes Error 424
that occurred when the user ran the procedure with text selected rather than just a
cursor position.

---

### Fix — Comma-separated new book silently lost; single-chapter book wrong chapter (`ParseCitationBlock`)

**Test input:** `Gen 3:1–5; Job 1:6; Isa 14:12–15, Ezek 28:12–19; 2 Pet 2:4; Jude 6; Rev 12:9`

**Symptom 1 — `FAIL: Isaiah 14:0` / Ezekiel missing**

The input had a **comma** between `Isa 14:12–15` and `Ezek 28:12–19` instead of a
semicolon. The parser treats `,` as a verse sub-list separator (same as `Ps 145:8–9,17`),
so `Ezek 28:12–19` was parsed as a verse spec inside Isaiah 14. `Ezek 28:12` is not
numeric → malformed range → sentinel `Isaiah 14:0` emitted; the Ezekiel reference was
lost entirely.

**Resolution:** User data error — the comma must be a semicolon:
`Isa 14:12–15; Ezek 28:12–19`. No code change required.

---

**Symptom 2 — `FAIL: Jude 2:6` → `PASS: Jude 0:6` → `PASS: Jude 1:6`**

After parsing `2 Pet 2:4`, `ctxChapter = 2`. When `Jude 6` was recognised as a new
book, `refPart = "6"` (no colon), so no chapter was parsed and `ctxChapter` remained 2
from the previous book. The parser emitted `Jude 2:6`. Jude has only one chapter, so
chapter 2 fails validation.

**Cause:** `ctxChapter` was never reset when the book context switched.

**Fix 1:** Added `If newBook Then ctxChapter = 0` to `ParseCitationBlock` between
Case 3 and the chapter-parse block. This reset `ctxChapter` on every book switch.

**Result of Fix 1:** `Jude 6` now passed validation (because `ValidateSBLReference`
internally promotes chapter 0 → 1 for single-chapter books), but the canonical string
stored in the Collection was `Jude 0:6` — the chapter 0 was never written back to
`ctxChapter` before building the `canon` string.

**Fix 2:** Added a single-chapter promotion directly in `ParseCitationBlock`, after
the `colonPos` block:

```vb
If ctxChapter = 0 And GetMaxChapter(ctxBookID) = 1 Then ctxChapter = 1
```

`ctxChapter` is promoted to 1 before `canon` is assembled. The Collection now contains
`Jude 1:6`. The guard `ctxChapter = 0` ensures explicitly-chaptered references
(e.g. `Jude 1:6` written in full) are unaffected.

---

### Fix — Single-chapter books output chapter number in SBL shorthand (`ToSBLShortForm`)

**Symptom:** Output showed `Jude 1:6` instead of the SBL shorthand `Jude 6`. SBL style
omits the chapter number for single-chapter books (Obad, Phlm, 2 John, 3 John, Jude,
etc.) — only the verse number is shown.

**Cause:** `ToSBLShortForm` assembled `abbr & " " & numPart` directly. For a canonical
string `Jude 1:6`, `numPart = "1:6"` — the chapter was included unchanged.

**Fix:** After the `Select Case` in `ToSBLShortForm`, strip the chapter prefix for
single-chapter books:

```vb
If GetMaxChapter(bID) = 1 Then
    Dim cpPos As Long
    cpPos = InStr(numPart, ":")
    If cpPos > 0 Then numPart = Mid$(numPart, cpPos + 1)
End If
```

`Jude 1:6` → `numPart = "6"` → output `"Jude 6"`.
Ranges also handled: `Jude 1:3-7` → `numPart = "3-7"` → `RenderEnDash` → `"Jude 3–7"`.


---

### Fix — Same-chapter verses used semicolon instead of comma (`RepairCitationBlockInParagraph`)

**Symptom:** Input `Gen 2:17, 25; 3:6–11` produced output `Gen 2:17; 2:25; 3:6–11`. Verses within the same chapter should be comma-separated with the verse number only; a chapter change within the same book should be semicolon-separated.

**Cause:** Task 5 of `RepairCitationBlockInParagraph` tracked only `prevBook`. When the book was the same but the chapter was also the same, it emitted `ch:verse` with a `"; "` separator instead of `verse` with `", "`.

**Fix:** Added `prevChap` tracking alongside `prevBook`. Task 5 now uses three cases:

| Condition | Separator | Output |
|---|---|---|
| Same book, same chapter | `, ` | verse only |
| Same book, different chapter | `; ` | `ch:verse` |
| New book | `; ` | full SBL short form |

`Gen 2:17, 25; 3:6–11` is now produced correctly from canonical input
`["Genesis 2:17", "Genesis 2:25", "Genesis 3:6-11"]`.


---

### Fix — Validation retry loop does not work (`RepairCitationBlockInParagraph`)

**Symptom:** The `Do...Loop` with `vbRetryCancel` MsgBox cannot work as designed. A
MsgBox blocks the Word UI, so the user cannot edit the paragraph while the box is open.
Clicking Retry re-reads the same unchanged text and loops forever.

**Cause:** Design error — `vbRetryCancel` implies the user can act between clicks, but
Word's document is inaccessible while a modal MsgBox is displayed.

**Fix:** Replaced the loop with a single validation pass:

- If `failCount > 0`: show the error report with `vbOKOnly` and the message
  “Fix the errors above in the paragraph, then run the command again.”, then `Exit Sub`.
- If `failCount = 0`: proceed directly to sort and render.

The user corrects the paragraph in the normal editing environment, then re-invokes
the command. The confirm prompt (Task 2) acts as the natural re-entry point.


---

### Fix — Whole-chapter references not supported (`ParseCitationBlock`, `ValidateSBLReference`, `SortCitationBlock`, `VerifyCitationBlockReport`, `RepairCitationBlockInParagraph`)

**Symptom:** Input `Gen 6:6; Ezek 16; Luke 15:4–32` produced
`FAIL: Ezekiel 0:16 (start verse failed)`. A reference to an entire chapter
(e.g. `Ezek 16`) was misread as verse 16 of chapter 0.

**Cause:** In `ParseCitationBlock`, when a new-book segment had no colon and
`ctxChapter = 0` (just reset on book switch), the token was treated as a verse number
rather than a chapter number. For multi-chapter books this produced `Ezekiel 0:16`
instead of `Ezekiel 16`. Five functions were affected:

1. `ParseCitationBlock` — emitted `ch:verse` sentinel instead of whole-chapter form
2. `ValidateSBLReference` — Rule 5 rejected empty `VerseSpec` after chapter promotion
3. `SortCitationBlock` — no-colon `numPart` was parsed as verse not chapter (wrong sort key)
4. `VerifyCitationBlockReport` — `CLng(Left$(numPart, cpPos - 1))` crashed when `cpPos = 0`
5. Task 5 of `RepairCitationBlockInParagraph` — `thisChap = ""` caused wrong separator logic

**Fix:**

1. **`ParseCitationBlock`**: After single-chapter promotion, added:
   ```vb
   If ctxChapter = 0 And colonPos = 0 And IsNumeric(vsStr) Then
       ctxChapter = CLng(vsStr)
       Result.Add ctxCanon & " " & ctxChapter
       GoTo NEXT_SEG
   End If
   ```
   Emits `"Ezekiel 16"` (no colon) and skips the verse loop.

2. **`ValidateSBLReference`**: Rule 5 changed from reject to accept on empty `VerseSpec`:
   ```vb
   If Len(VerseSpec) = 0 Then
       ValidateSBLReference = True
       Exit Function
   End If
   ```

3. **`SortCitationBlock`**: Added `cpPos = 0` branch — `ch = CLng(numPart)`, `sV = 0`.
   Whole-chapter sorts before any verse reference in the same chapter.

4. **`VerifyCitationBlockReport`**: Added `cpPos = 0` branch — `ch = CLng(numPart)`,
   `vPart = ""`. Whole-chapter items call `ValidateSBLReference` with empty `VerseSpec`
   and bypass the verse parse block entirely.

5. **Task 5 `RepairCitationBlockInParagraph`**: `colonPos = 0` now sets
   `thisChap = numPart` (chapter number) instead of `thisChap = ""`, so the
   same-book/same-chapter separator logic works correctly.

**Tests added:** `Test_WholeChapterReference` in `basTEST_aeBibleCitationBlock.bas`,
covering parse, mixed-block, verify report, sort order, and `ToSBLShortForm`.
Added to `Run_Extra_Tests`.


---

### Fix — Raw error MsgBox shown for non-citation content in selection (`RepairCitationBlockInParagraph`, `VerifyCitationBlockReport`)

**Symptom:** Cursor placed in a paragraph containing non-citation text (e.g.
“—resolution explained in EDP pp. 153–162]”) produced a raw technical MsgBox:
“Error -2147220501 (Unrecognised token (non-ASCII character): “—resolution”)”.
A second MsgBox followed because `failCount` remained 0 after the error, so
`RepairCitationBlockInParagraph` proceeded to Task 4 and raised the same error again.

**Cause:** `ParseCitationBlock` correctly raised `vbObjectError + 1003` for the
em-dash token. This propagated to `VerifyCitationBlockReport`’s PROC_ERR handler,
which showed the raw error MsgBox and resumed to PROC_EXIT with `failCount = 0`.
The caller then interpreted 0 failures as a clean block and continued processing.

**Fix:**

1. **`VerifyCitationBlockReport` PROC_ERR**: parse errors 1002 (block too long) and
   1003 (non-ASCII token) are now caught silently — `failCount` is set to `-1` as a
   sentinel and the function returns without showing a MsgBox:
   ```vb
   If Err.Number = vbObjectError + 1002 Or Err.Number = vbObjectError + 1003 Then
       failCount = -1
       Resume PROC_EXIT
   End If
   ```

2. **`RepairCitationBlockInParagraph`**: added `failCount = -1` check immediately
   after the `VerifyCitationBlockReport` call:
   ```vb
   If failCount = -1 Then
       MsgBox "The selected text contains non-citation content." & vbCrLf & vbCrLf & _
              "Select only the citations to validate, or insert the cursor in a " & _
              "paragraph that contains only citations.", _
              vbOKOnly + vbExclamation, "Invalid Selection"
       Exit Sub
   End If
   ```

The user now sees a single user-friendly message and the procedure exits cleanly.
Both error 1002 and 1003 use the same friendly message path.
