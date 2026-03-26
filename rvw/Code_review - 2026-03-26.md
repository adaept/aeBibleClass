# Code Review — `C:\adaept\aeBibleClass\src`

**Date:** 2026-03-26
**Reviewer:** Claude Code (claude-sonnet-4-6)
**Files reviewed:** 27 VBA modules and classes
**Total issues found:** 13 — Critical: 2 | High: 3 | Medium: 5 | Low: 3

---

## Critical Issues

### `aeBibleClass.cls` — Bare `End` statement in `PROC_ERR` terminates application (line 1024)

```vb
PROC_ERR:
    answer = MsgBox("Err = " & Err.Number & " ... Do you want to continue?", ...)
    If answer = vbYes Then
        Resume
    Else
        ' commented-out Exit Sub and Stop
    End If
    Debug.Print "!!! Error in Test num = " & num ...
    Stop
    End       ' <-- terminates entire VBA application / destroys all module state
End Function
```

When the user answers "No" to the continue dialog, execution falls through past the commented-out code, hits `Stop` (IDE breakpoint), then `End`. `End` in VBA unconditionally terminates the entire running application — all module-level variables are reset, all object references destroyed. Any pending document operations are abandoned. Should be replaced with `Exit Function` or `Resume PROC_EXIT`.

---

### `XbasTESTaeBibleDOCVARIABLE.bas` — `ErrorHandler:` label unreachable; `Err.Raise` propagates unhandled (line 148)

```vb
Sub VerifyBookNameFromDocVariable(docVar As String, theTextOfH1 As String)
    ' ... no On Error GoTo ErrorHandler anywhere ...
RetrySearch:
    textFoundHere = FindNextHeading1OnVisiblePage(...)
    If textFoundHere Then
        ...
    Else
        Err.Raise 1000, "...", "..."   ' <-- no active error handler; propagates to caller
    End If
    Exit Sub

ErrorHandler:          ' <-- unreachable; no On Error GoTo ErrorHandler in scope
    MsgBox Err.Description...
    Resume RetrySearch ' <-- also unreachable
End Sub
```

`Err.Raise` at line 148 fires with no `On Error GoTo ErrorHandler` statement in the procedure. The `ErrorHandler:` label with `Resume RetrySearch` is dead code — it can never be reached. The raised error propagates unhandled to whatever called `VerifyBookNameFromDocVariable`. The intended retry-on-wrong-page logic never runs. Fix: add `On Error GoTo ErrorHandler` at the top of the Sub.

---

## High Priority

### `XbasTESTaeBibleDOCVARIABLE.bas` — `Replace()` called with `Word.Range` object instead of `String` (lines 519–572)

```vb
' In TestPageNumbers (New Testament section, lines 519-572):
VerifyBookNameFromDocVariable "Matt", "Matthew"
Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")   ' error 13 — lastFoundLocation is Word.Range
```

`lastFoundLocation` is declared `As Word.Range`. `Replace()` expects a `String` as its first argument; `Word.Range` has no default property and cannot be auto-coerced. Every `Debug.Print` line in the New Testament loop raises error 13 (Type Mismatch) at runtime.

Note: The same pattern in lines 434–510 (Old Testament) is unreachable dead code due to `GoTo NewTestament` at line 430. The fix from a prior session correctly used `.Text` at line 45 (`FindNextHeading1OnVisiblePage`), but the 28 calls in `TestPageNumbers` were not updated.

Fix: replace all `Replace(lastFoundLocation, vbCr, "")` with `Replace(lastFoundLocation.Text, vbCr, "")` in `TestPageNumbers`.

---

### `basUSFM_Export.bas` — Three hardcoded absolute paths (lines 46–48)

```vb
Private Const LOG_FILE      As String = "C:\adaept\aeBibleClass\rpt\USFM_Export_Log.txt"
Private Const OUTPUT_FILE   As String = "C:\adaept\aeBibleClass\rpt\ExportedBible.usfm"
Private Const VALIDATOR_LOG As String = "C:\adaept\aeBibleClass\rpt\USFM_Validator_Log.txt"
```

All three constants are hardcoded to a machine-specific absolute path. If the repository is cloned to a different location or run on another machine, all export and validation operations fail silently or with a path error. The rest of the codebase uses `ActiveDocument.Path` to derive paths at runtime (e.g., `aeRibbonClass.cls` line 254). These constants cannot be changed at runtime; they should be converted to module-level variables initialised from `ActiveDocument.Path`.

---

### `aeRibbonClass.cls` — `ScreenUpdating = False` with no error handler in `GoToH1` (line 131)

```vb
Private Sub GoToH1()
    ' ... no On Error handler ...
    Application.ScreenUpdating = False

    For Each para In ActiveDocument.Paragraphs   ' iterates 800+ pages
        If para.style = "Heading 1" Then
            ...
        End If
    Next para

    Application.ScreenUpdating = True   ' only reached if no error
End Sub
```

If any error occurs during the 800-page paragraph scan (COM error, object invalidated, etc.), `ScreenUpdating` is never restored to `True`. Word's screen stays frozen until the user manually invokes the application. Fix: add `On Error GoTo Cleanup` with a `Cleanup:` label that restores `Application.ScreenUpdating = True` before `Exit Sub`.

---

## Medium Priority

### `basAddHeaderFooter.bas` — Dead assignment immediately overwritten (line 157)

```vb
Set oFooter = oSection.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary)  ' <-- HEADER, not footer
Set oFooter = oSection.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)  ' overwrites immediately
```

Line 157 sets `oFooter` to the primary **header**, then line 158 immediately overwrites it with the primary **footer**. The first assignment is a dead no-op (likely a copy-paste leftover from the header code above). Remove line 157.

---

### `basAddHeaderFooter.bas` — `oSearch` and `oFound` not cleaned up at Sub exit (lines 64–67)

```vb
ElseIf oPara.style = oDoc.Styles("Heading 2") Then
    Dim oSearch As Word.Range
    Set oSearch = oDoc.Range(0, oSection.Range.Start)
    Dim oFound  As Word.Range
    Set oFound = Nothing
    ...
End If
' End of loop — no cleanup of oSearch or oFound
' Set oRange/oPara/oHeader/oSection/oSections/oDoc cleaned up, but oSearch and oFound are not
```

All other object variables declared in `AddBookNameHeaders` are explicitly set to `Nothing` at lines 98–103. `oSearch` and `oFound` are the only exceptions. Add `Set oSearch = Nothing` and `Set oFound = Nothing` before `End Sub`.

---

### `XbasTESTaeBibleClass_SLOW.bas` — `pageBreakParagraphs` declared and reported but never incremented (line 205)

```vb
Dim pageBreakParagraphs As Long   ' declared
' ... no code ever increments pageBreakParagraphs ...
Debug.Print "Paragraphs with Page Break: " & pageBreakParagraphs   ' always prints 0
```

`pageBreakParagraphs` is reported alongside `columnBreakParagraphs` and `textWrappingBreakParagraphs` (which are correctly incremented), but is never updated. The `^p` pattern check (page break in Find) is missing. The output line gives a false 0.

---

### `XbasTESTaeBibleDOCVARIABLE.bas` — Large unreachable block in `TestPageNumbers` (lines 430–510)

```vb
Sub TestPageNumbers()
    GoTo NewTestament   ' <-- unconditional jump

    ' Old Testament (lines 433-510):
    VerifyBookNameFromDocVariable "Gen", "Genesis"   ' dead
    ...
    ' 37 more dead calls
```

`GoTo NewTestament` at line 430 renders 78 lines of Old Testament verification calls permanently unreachable. These contain the same `Replace(lastFoundLocation, vbCr, "")` type-mismatch bug. Either remove the dead block or replace `GoTo NewTestament` with a comment explaining the skip.

---

### `basUSFM_Export.bas` — Public entry point `ExportUSFM_PageRange` has no error handler

```vb
Public Sub ExportUSFM_PageRange(ByVal startPage As Long, ByVal endPage As Long)
    ' No On Error GoTo / PROC_ERR handler
    ...
    Dim rng As Word.Range
    Set rng = GetRangeForPages(startPage, endPage)
    ...
    usfm = ConvertRangeToUSFM(rng)   ' can fail on large documents
    WriteTextFile OUTPUT_FILE, usfm  ' can fail if path invalid
```

The public entry point has no `PROC_ERR`/`PROC_EXIT` handler — the pattern used throughout the rest of the codebase. Any runtime error (range failure, file write error, style access error during conversion) surfaces as an unhandled VBA error dialog rather than a clean log entry.

---

## Low / Style

### `aeBibleClass.cls` — `Stop` before `End` in `PROC_ERR` (line 1023)

```vb
    Stop    ' development breakpoint — suspends in IDE
    End     ' terminates application (Critical issue above)
```

Once `End` is replaced with `Exit Function` (Critical fix above), `Stop` at line 1023 will remain. `Stop` is a development-time breakpoint that should not be in production error handlers. During test runs it will trap in the IDE unexpectedly. Remove `Stop` or convert to `Debug.Print`.

---

### `basBibleRibbon_OLD.bas` — Legacy module with public variable declarations (lines 7–19)

```vb
Public headingData(1 To 66, 0 To 1) As Variant  ' duplicates aeRibbonClass name
Public ribbonUI As IRibbonUI                      ' duplicates aeRibbonClass name
Public ribbonIsReady As Boolean                   ' duplicates aeRibbonClass name
Public BtnNextEnabled As Boolean                  ' duplicates aeRibbonClass name
Dim bookmarkIndex As Long                         ' implicit Public (no Private keyword)
```

The `_OLD` suffix indicates this module is superseded by `aeRibbonClass.cls`. Its public variable names shadow the class's private encapsulated state. `Dim bookmarkIndex` at module level without `Private` defaults to public scope in a standard module (contained by `Option Private Module`, but still inconsistent). Add a clear deprecation comment at the top, or mark all declarations `Private` to prevent accidental use.

---

### `basSBL_TestHarness.bas` — Comment-only lines as VBA line continuations (lines 116–119)

```vb
Array("Jude 5-7", 65, False, False) _  ' Range spec: ...
                                        ' When Stage 8-12 range support is added...
                                        ' Without the IsNumeric guard...
    )
```

VBA allows a comment after `_` on the same physical line, but comment-only lines (`'...`) as continuation targets between `_` and the next statement may not be supported consistently across VBA versions. If the closing `)` on line 119 is parsed as a standalone token rather than completing the `Array(...)` call, a compile error results. Verify this compiles cleanly; if not, move the multi-line comment block outside the array declaration.

---

## Summary

| Severity   | Count |
|------------|-------|
| Critical   | 2     |
| High       | 3     |
| Medium     | 5     |
| Low/Style  | 3     |
| **Total**  | **13**|

### Files with issues

| File | Severity |
|------|----------|
| `aeBibleClass.cls` | Critical, Low |
| `XbasTESTaeBibleDOCVARIABLE.bas` | Critical, High, Medium |
| `basUSFM_Export.bas` | High, Medium |
| `aeRibbonClass.cls` | High |
| `basAddHeaderFooter.bas` | Medium × 2 |
| `XbasTESTaeBibleClass_SLOW.bas` | Medium |
| `basBibleRibbon_OLD.bas` | Low |
| `basSBL_TestHarness.bas` | Low |

### Positive notes

The `PROC_ERR`/`PROC_EXIT` error pattern, `Option Explicit`, `#NNN` issue tracking, and object cleanup (`Set x = Nothing`) are consistently applied across the majority of the codebase. The 14-stage parser pipeline in `basSBL_Citation_EBNF.bas` is well-structured with clear stage contracts. `PtrSafe` API declarations, early binding for `Word.*` types, and the `X-prefix` deferred module convention are all correctly applied. `basAddHeaderFooter.bas`, `basImportWordGitFiles.bas`, and `basBibleRibbonSetup.bas` are clean with no issues found.

---

## Fixes Applied — 2026-03-26

All 13 issues resolved in session with Claude Code (claude-sonnet-4-6).

| # | Severity | File | Fix |
|---|----------|------|-----|
| 1 | Critical | `aeBibleClass.cls` | Replaced `Stop` + `End` with `Resume PROC_EXIT` in `PROC_ERR` |
| 2 | Critical | `XbasTESTaeBibleDOCVARIABLE.bas` | Added `On Error GoTo ErrorHandler` to `VerifyBookNameFromDocVariable` |
| 3 | High | `XbasTESTaeBibleDOCVARIABLE.bas` | Replaced all `Replace(lastFoundLocation, ...)` with `Replace(lastFoundLocation.Text, ...)` (77 occurrences) |
| 4 | High | `basUSFM_Export.bas` | Converted 3 hardcoded path `Const` declarations to module variables; added `InitPaths` initialised from `ActiveDocument.Path`; called from `ExportUSFM_PageRange` |
| 5 | High | `aeRibbonClass.cls` | Added `On Error GoTo Cleanup` and `Cleanup:` label restoring `ScreenUpdating = True` in `GoToH1` |
| 6 | Medium | `basAddHeaderFooter.bas` | Removed dead header assignment immediately overwritten by footer assignment |
| 7 | Medium | `basAddHeaderFooter.bas` | Added `Set oSearch = Nothing` and `Set oFound = Nothing` to cleanup block |
| 8 | Medium | `XbasTESTaeBibleClass_SLOW.bas` | Added `^p` page break check in `CountParagraphsTypes` loop; `pageBreakParagraphs` now correctly incremented and indexed |
| 9 | Medium | `XbasTESTaeBibleDOCVARIABLE.bas` | Replaced `GoTo NewTestament` with inline comment noting NT-only check is intentional |
| 10 | Medium | `basUSFM_Export.bas` | Added `PROC_ERR`/`PROC_EXIT` handler to `ExportUSFM_PageRange`; errors logged and reported via MsgBox |
| 11 | Low | `aeBibleClass.cls` | `Stop` removed as part of Issue 1 fix |
| 12 | Low | `basBibleRibbon_OLD.bas` | Added deprecation comment at top of module |
| 13 | Low | `basSBL_TestHarness.bas` | Moved `"Jude 5-7"` comment outside `Array(...)` declaration; prefixed with `"Jude 5-7":` to clarify scope |
