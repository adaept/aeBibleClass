# Code Review — `C:\adaept\aeBibleClass\src`

**Date:** 2026-03-15
**Files reviewed:** 23 VBA modules (~16,289 lines total)

---

## High Priority

**1. `aewordgitClass.cls` — Error suppression (~line 293)**
The error handler uses `If Err = -2147467259 Then Resume Next` — this silently swallows a specific error class. Callers cannot detect failure. The magic number should be a named constant at minimum.

**2. `basUSFMExport.bas` — Unreachable `Case` (~line 201)**
```vba
Case "Book Title", "Heading 1"
    ...
Case "Heading 1"   ' ← NEVER REACHED — already matched above
    ...
```
The second `"Heading 1"` case is dead code. Remove it or restructure the `Select`.

**3. `basWordRepairRunner.bas` — Loop counter modified via `GoTo` (~line 189)**
Inside a `Do While` loop, `i = verseEnd` is set then execution jumps to `SkipLogging:` via `GoTo`. This bypasses normal loop structure and risks incorrect iteration or an infinite loop. Restructure with a flag variable or `Exit Do`.

---

## Medium Priority

**4. `aeBibleClass.cls` — Uninitialized return in `TheBibleClassTests` (~line 103)**
The `Else` branch (unexpected parameter type) prints a debug message but never assigns a return value. Caller receives an uninitialized integer. Assign an explicit failure code.

**5. `basWordRepairRunner.bas` — Last-page edge case (~line 109)**
`GoTo(What:=wdGoToPage, name:=CStr(pageNum + 1))` returns the last page when `pageNum + 1` exceeds the document length. This makes `pageEnd` incorrect for the final page. Validate against `ActiveDocument.Pages.Count`.

**6. `basImportWordGitFiles.bas` — Unguarded `VBProject` access (lines 28, 57)**
`ThisDocument.VBProject.VBComponents` has no `On Error GoTo` wrapper. If the VBA project is protected or access is denied, it crashes unhandled. (The existing `basTestaeBibleClass.cls` already handles error 6068 — apply the same pattern here.)

**7. `XLongRunningProcessCode.bas` — `CustomDocumentProperties` assumed to exist (lines 36, 42)**
Reading/writing `CustomDocumentProperties("LastProcessedParagraph")` without checking existence first will crash if the property hasn't been created yet.

**8. `basBibleRibbon.bas` — Uninitialized public globals (lines 16–19)**
`ribbonUI`, `ribbonIsReady`, `btnNextEnabled` rely on `RibbonOnLoad` being called first. If a ribbon callback fires before load completes, `ribbonUI` is `Nothing` and any `ribbonUI.InvalidateControl` call crashes.

**9. `basWordSettingsDiagnostic.bas` — Silent `On Error Resume Next` blocks (lines 224, 260, 275)**
Multiple loops run under `On Error Resume Next` with no `Err.Number` check after. Callers can't distinguish "no match found" from "error occurred."

**10. `basSBL_TestHarness.bas` — Error check after `On Error GoTo 0` (lines 140–149)**
```vba
On Error Resume Next
bookName = ResolveAlias(...)
On Error GoTo 0
If Err.Number <> 0 Then  ' ← error state already cleared above
```
`Err.Number` should be checked *before* `On Error GoTo 0`. This pattern works by accident because `Err` isn't reset until the next operation — but it's fragile.

---

## Low Priority / Code Quality

**11. `basWordRepairRunner.bas` — Dead variable `chInfo` (lines 169–170)**
`chInfo` is declared and assigned but never read. Remove it.

**12. `basUSFMExport.bas` — Implicit empty-string return (line 218)**
`If Right$(txt, 1) = ":" Then ConvertParagraphToUSFM = "\is2 " & txt` — if the condition is false, the function returns `""` with no comment explaining why that's intentional.

**13. `basWordRepairRunner.bas` — O(n²) string concatenation inside loop (~lines 96–107)**
`logBuffer = logBuffer & ...` inside a `Do While` loop. For large documents, use an array and `Join()` once at the end.

**14. `Module1.bas` — No module-level header**
At 1,303 lines this is the largest utility module but has no documentation of its purpose or the relationships between its functions.

**15. `XLongRunningProcessCode.bas`, `XbasTESTaeBibleClass_SLOW.bas`, `XbasTESTaeBibleDOCVARIABLE.bas` — No module headers**
The `X`-prefix convention (marking deferred/slow modules) is not documented within the files themselves.

---

## Positive Notes

The overall architecture is sound: the 14-stage parser pipeline is well-structured, the test framework (`AssertTrue`/`AssertFalse`/`AssertEqual` with counters) is clean, error handling follows a consistent `PROC_ERR`/`PROC_EXIT` pattern in the majority of modules, and the `#NNN` issue-tracking system keeps changelog entries well-connected to the codebase.

---

## Issue Count Summary

| Severity | Count |
|----------|-------|
| High     | 3     |
| Medium   | 7     |
| Low      | 5     |
| **Total**| **15**|
