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
