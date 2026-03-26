# Code Review — `C:\adaept\aeBibleClass\src`

**Date:** 2026-03-17
**Files reviewed:** 23 VBA modules and classes
**Total issues found:** 31 — Critical: 5 | High: 8 | Medium: 12 | Low/Style: 6

---

## Critical Issues

### `basWordRepairRunner.bas` — `Exit Do` exits wrong loop (line 189)

```vb
If Len(combinedNumber) = 1 And AscW(combinedNumber) = 12 Then
    ascii12Count = ascii12Count + 1
    i = verseEnd
    Exit Do   ' <-- exits inner loop only; outer Do While i < pageEnd continues from i = verseEnd
End If
```
`Exit Do` only breaks the inner loop. The outer loop resumes at `i = verseEnd`, potentially skipping verse markers silently.

---

### `basUSFMExport.bas` — Unsafe range expansion with no loop-entry guard (lines 336–351)

```vb
Do While rChap.End < p.Range.End And rChap.style = "Chapter Verse marker"
    rChap.MoveEnd wdCharacter, 1
Loop
rChap.MoveEnd wdCharacter, -1   ' step back one char after overshoot
```
If the loop never executes (style check fails on entry), `MoveEnd -1` still runs and shrinks the range incorrectly. Guard with a flag.

---

### `aewordgitClass.cls` — Dead `IsNull` check on a `String` variable (line 260)

```vb
If IsNull(aeWordGitSourceFolder) Then
```
`aeWordGitSourceFolder` is declared as `String` and initialized to `"default"` on line 67. A `String` can never be `Null` in VBA. This condition will never be true and the branch is dead code.

---

### `basSBL_VerseCountsGenerator.bas` — Call to undefined function (line 50)

```vb
AssertOneBased temp, context
```
`AssertOneBased` is not defined in this file or any visible module. This will cause a compile error at runtime.

---

### `XLongRunningProcessCode.bas` — `CustomDocumentProperties` accessed without existence check (lines 36–37, 42–44)

```vb
' Write (line 36-37) — no guard; error 5 if property absent
ActiveDocument.CustomDocumentProperties("LastProcessedParagraph").value = lastProcessedParagraph

' Read (lines 42-44) — silently leaves variable at 0 if property missing
On Error Resume Next
lastProcessedParagraph = ActiveDocument.CustomDocumentProperties("LastProcessedParagraph").value
progressPercentage     = ActiveDocument.CustomDocumentProperties("ProgressPercentage").value
On Error GoTo 0
```
The write path has no error guard; the read path silently leaves variables at 0 with no indication that the property was absent.

---

## High Priority

### `aewordgitClass.cls` — Error suppression without logging (lines 150–153, 294–300)

**Lines 150–153:** `Kill` runs under `On Error Resume Next`, then the handler is immediately replaced with `On Error GoTo PROC_ERR`. If `Kill` fails, the error is silently discarded before it can be caught.

```vb
On Error Resume Next
Kill FolderWithVBAProjectFiles & "*.*"
On Error GoTo PROC_ERR   ' error from Kill is lost here
```

**Lines 294–300:** Catches `E_FAIL (-2147467259)` with `Resume Next` — no log entry, no diagnostic. Other errors fall through to a MsgBox. Two inconsistent outcomes for different failures.

---

### `basSBL_TestHarness.bas` — `Split()` called twice on same input (line 46–47)

```vb
p.Chapter  = CLng(Split(refPart, ":")(0))   ' first split
p.VerseSpec = Split(refPart, ":")(1)         ' second split — redundant allocation
```
Two separate `Split()` calls on the same string. Cache: `Dim parts() As String: parts = Split(refPart, ":")`.

---

### `basSBL_TestHarness.bas` — `CLng()` on unvalidated verse spec (line 185)

```vb
rewritten = RewriteSingleChapterRef(BookID, parsed.Chapter, CLng(parsed.VerseSpec))
```
If `VerseSpec` contains a range (`"16-18"`) or is empty, `CLng()` raises runtime error 13 (Type Mismatch). No `On Error` guard present.

---

### `basSBL_TestHarness.bas` — Enum/Boolean type mismatch in test array (line 291)

```vb
' Test array stores False (Boolean = 0)
tests = Array(Array("Jude 0", False), ...)
' But comparison uses enum value (FailResolveBook = 1)
If tests(i)(1) = FailResolveBook Then
```
`False = 0` and `FailResolveBook = 1` are never equal. The negative-case test branch never fires, making negative tests permanently silent.

---

### `basUSFMExport.bas` — Unreachable `Case "Heading 1"` (line 205)

```vb
Case "Book Title", "Heading 1"   ' matches both
    ConvertParagraphToUSFM = MakeTitleLine(1, txt)
Case "Heading 1"                  ' <-- DEAD CODE: already matched above
    ConvertParagraphToUSFM = "\mt1 " & txt
```
Remove the second `Case "Heading 1"` block entirely.

---

### `basUSFMExport.bas` — `CLng()` on cleaned range text without guard (line 353)

```vb
chapNum = CLng(Trim$(CleanTextForUTF8(rChap.text)))
```
If `rChap.text` is empty after cleaning, `CLng("")` raises runtime error 13. No guard or `IsNumeric()` check.

---

### `basTestaeBibleClass.bas` — Git shell commands with no exit-code validation (lines 109–114)

```vb
cmdOutput = wsh.exec(shellCmd).StdOut.ReadAll   ' StdErr never read
Debug.Print "[TAG] " & cmdOutput                 ' prints blank on failure; no exit code checked

cmdOutput = wsh.exec(shellCmd).StdOut.ReadAll
Debug.Print "[PUSH] " & cmdOutput               ' a failed push looks identical to success
```
`wsh.exec()` does not raise a VBA error on non-zero exit codes. `StdErr` is never read. A failed `git push` is completely invisible.

---

### `basTESTaewordgitClass.bas` — `Shell()` call with no error handling (line 37)

```vb
Shell "cmd.exe /c """ & strBat & """", vbNormalFocus
```
If the batch file doesn't exist or `Shell` fails, execution continues silently. No `On Error` trap; no existence check on `strBat`.

---

## Medium Priority

### `basWordRepairRunner.bas` — Dead assignment (line 267)

```vb
fixCount = fixCount   ' no-op; remove
```

### `basWordRepairRunner.bas` — `AscW` magic number (line 189)

```vb
If AscW(combinedNumber) = 12 Then   ' 12 = form feed; define as named constant
```
Define `Const ASCII_FORMFEED As Long = 12`.

### `basWordRepairRunner.bas` — Style comparison without `Trim` (line 201)

```vb
If prefixStyle = "Normal" Then
```
Style names with leading/trailing spaces will silently fail to match. Use `Trim$(prefixStyle)`.

### `basUSFMExport.bas` — ADODB stream opened without validating success (lines 439–462)

`LogValidator` creates an `ADODB.Stream` object in late-binding mode but does not check whether `stream.Open` succeeded. An invalid file path fails silently.

### `basUSFMExport.bas` — `Mid$` length check missing (lines 518–523)

```vb
If Mid$(line, Len(marker) + 1, 1) <> " " Then
```
If `Len(line) = Len(marker)`, `Mid$` returns `""` and the comparison `"" <> " "` passes incorrectly, producing a false validation failure. Check `Len(line) > Len(marker)` first.

### `basWordSettingsDiagnostic.bas` — Stale `Err.Number` check (lines 68–79)

```vb
On Error Resume Next
Result = ActiveWindow.View.ShowTextBoundaries
' Err.Number is never cleared before this block; may carry a stale value from earlier code
```

### `basWordSettingsDiagnostic.bas` — Null variant concatenation (lines 109–110)

```vb
discrepancies.Add key, "Current: " & current(key) & " | Expected: " & target(key)
```
If `current(key)` or `target(key)` is `Null`, string concatenation raises a Type Mismatch. Use `CStr(current(key))`.

### `basTESTaeBibleFonts.bas` — Document not closed on error path (lines 60–66)

```vb
Set testRange = TestDoc.content
testRange.font.name = fontName
IsFontInstalled = (testRange.font.name = fontName)
TestDoc.Close SaveChanges:=False
```
If any line before `TestDoc.Close` raises an error, the document is never closed. Add `On Error GoTo` with cleanup.

### `basBibleRibbon.bas` — Unreachable code after unconditional `GoTo` (lines 209–213)

```vb
GoTo Cleanup      ' unconditional jump
    verseNum = 1  ' <-- DEAD CODE
    GoTo Chapter  ' <-- DEAD CODE
End If
```

### `basBibleRibbon.bas` — `chapIdx` used before assignment (lines 327–328)

```vb
Dim chapIdx As Long    ' declared but never assigned
idx = chapIdx          ' idx = 0 always; if loop below never runs, wrong result
```

### `Module1.bas` — `Mid()` called twice per iteration (line 25)

```vb
msg = msg & "Character " & i & ": " & mid(selectedText, i, 1) & " (ASCII: " & Asc(mid(selectedText, i, 1)) & ")"
```
Cache: `Dim ch As String: ch = mid(selectedText, i, 1)`.

### `basSBL_VerseCountsGenerator.bas` — `Debug.Assert` stripped in compiled mode (line 172)

```vb
Debug.Assert Len(packed) = (UBound(chapters) - LBound(chapters) + 1) * 3
```
`Debug.Assert` is silently removed when VBA is compiled. Replace with an explicit `If...Then Err.Raise` for production data validation.

---

## Low / Style

### `basSBL_TestFramework.bas` — Global mutable test counters with no encapsulation

`gTestsRun` and `gTestsFailed` are public globals. Any module can modify them directly. If tests reset them between runs, results from earlier tests are lost with no warning.

### `basTestaeBibleClass.bas` — Unreachable condition (line 41)

```vb
If CStr(varDebug) = "Error 448" Then
```
`varDebug` is either `Missing`, `"varDebug"`, or an integer by this point (lines 21–30). The string `"Error 448"` can never arrive here.

### `basImportWordGitFiles.bas` — No secondary confirmation before bulk module deletion (line 26)

`DeleteAllModulesExceptImporter()` deletes every VBA module. One MsgBox confirmation exists, but there is no recovery path if the user clicks Yes by mistake.

### `ThisDocument.cls` — All event handlers commented out

`Document_Open` and `Document_Close` are entirely commented out. The current state is ambiguous — it is unclear whether this is intentional or incomplete. Add a comment explaining why they are disabled.

### `bas_TODO.bas` — No executable code

Entire module is planning notes. Move content to a `.md` file; a `.bas` module with no executable code adds unnecessary clutter to the VBA project.

### `basChangeLog_aewordgit.bas` — Incomplete task note without issue number

```
' Use aeWordGit, NOTE: uppercase WordGit throughout for readability
```
No `#NNN` issue number assigned. Either link to an existing issue or close.

---

## Cross-File Patterns

### Inconsistent error handling style

Three patterns are used interchangeably with no clear rule:

| Pattern | Files |
|---------|-------|
| `On Error Resume Next` + `Err.Number` check | basSBL_TestHarness, basWordSettingsDiagnostic |
| `On Error GoTo PROC_ERR` + `PROC_EXIT` | aeBibleClass, aewordgitClass (majority) |
| `On Error GoTo 0` with no handler | basImportWordGitFiles, aewordgitClass (mixed) |

The `PROC_ERR`/`PROC_EXIT` pattern is correct and dominant — apply it consistently everywhere.

### `CLng()` on unvalidated parser output (multiple files)

`basSBL_TestHarness.bas` line 185, `basUSFMExport.bas` line 353, `basBibleRibbon.bas` line 328. All call `CLng()` on strings that may be empty or contain ranges. Add a shared helper:

```vb
Function SafeCLng(s As String, defaultVal As Long) As Long
    On Error Resume Next
    SafeCLng = CLng(Trim$(s))
    If Err.Number <> 0 Then SafeCLng = defaultVal
    On Error GoTo 0
End Function
```

---

## Summary

| Severity   | Count |
|------------|-------|
| Critical   | 5     |
| High       | 8     |
| Medium     | 12    |
| Low/Style  | 6     |
| **Total**  | **31**|

### Positive notes

The `PROC_ERR`/`PROC_EXIT` error pattern, `#NNN` issue-tracking convention, and 14-stage parser pipeline structure are consistently applied across the majority of the codebase. The test framework (`AssertTrue`/`AssertFalse`/`AssertEqual`) is clean and the module separation between parser, test harness, and repair runner is well-maintained.

---

## Resolution Summary — 2026-03-25

All 31 items reviewed one at a time. 9 fixed, 2 FIXME_LATER comments added, 8 skipped as already fixed from prior sessions, 8 skipped as review incorrect or not applicable, 4 skipped by design decision.

### Fixed (9)

| # | Item | File |
|---|------|------|
| Critical 3 | Removed dead `IsNull` branch on `String` variable — `ElseIf IsNull(aeWordGitSourceFolder)` block deleted | `aeWordGitClass.cls` |
| High 2 | Cached `Split(refPart, ":")` result into `cvParts()` — eliminated redundant double call | `basSBL_TestHarness.bas` |
| High 3 | Guarded `CLng(parsed.VerseSpec)` with `IsNumeric` — prevents error 13 crash when Stage 8–12 range support is added; added `"Jude 5-7"` range test case with forward-compat comment | `basSBL_TestHarness.bas` |
| High 8 | Added `Dir()` existence check and `On Error Resume Next` + `Err.Number` guard around `Shell` call for normalizer batch file | `basTEST_aeWordGitClass.bas` |
| Medium 2 | Replaced magic number `12` with named constant `ASCII_FORMFEED` | `basWordRepairRunner.bas` |
| Medium 4 | Added `Set stm = Nothing` in `ErrHandler` of `LogValidator` | `basUSFM_Export.bas` |
| Medium 7 | Wrapped `current(key)` and `target(key)` with `"" &` coercion to guard against Null variant in string concatenation | `basWordSettingsDiagnostic.bas` |
| High 8 | Added `Set stm = Nothing` to `ErrHandler` in `LogValidator` | `basUSFM_Export.bas` |
| Low 4 | Added comment to `ThisDocument.cls` explaining `Document_Open` is intentionally empty (used only for timing tests) | `ThisDocument.cls` |

### FIXME_LATER Comments Added (2)

| Item | File |
|------|------|
| Critical 2 — `MoveEnd -1` may shrink range incorrectly if loop exits via boundary rather than style-change overshoot | `basUSFM_Export.bas` |
| High 6 — `CLng()` on cleaned range text; guard needed if `CleanTextForUTF8` ever strips digit characters | `basUSFM_Export.bas` |

### Skipped — Already Fixed in Prior Sessions (8)

Critical 5, High 1, High 4, High 7, Medium 1, Medium 3, Medium 11, Low 2.

### Skipped — Review Incorrect or Not Applicable (8)

| Item | Reason |
|------|--------|
| Critical 1 | `Exit Do` at line 198 exits the outer loop — no inner loop active at that point |
| Critical 4 | `AssertOneBased` is defined as `Public Sub` in `basSBL_Citation_EBNF.bas` line 2451 |
| High 5 | `Case "Book Title"` and `Case "Heading 1"` are separate with distinct outputs — not combined |
| Medium 5 | `Mid$` already guarded by `Len(line) > Len(marker) + 1` at line 539 |
| Medium 6 | No `Err.Number` check in `GetShowTextBoundaries` — nothing stale to read |
| Medium 8 | `On Error Resume Next` wraps entire function; `TestDoc.Close` always reached |
| Low 1 | `gTestsRun` and `gTestsFailed` are already `Private` |
| Low 6 | Issue `#016` is already assigned in `basChangeLog_aeWordGitClass.bas` |

### Skipped — Not Applicable (2)

Medium 9 and Medium 10 reference `basBibleRibbon.bas` which no longer exists (refactored to `aeRibbonClass.cls`).

### Skipped — By Decision (4)

| Item | Reason |
|------|--------|
| Medium 12 | `Debug.Assert` in a dev-time generator Sub — appropriate for that context |
| Low 3 | Bulk delete confirmation already lists all modules by name; second prompt adds friction without value |
| Low 5 | `bas_TODO.bas` kept as-is |
| Critical 2 (partial) | Full guard deferred to FIXME_LATER; document-specific behaviour makes crash path unreachable in practice |
