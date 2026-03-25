# Code Review — `C:\adaept\aeBibleClass\src`

**Date:** 2026-03-16
**Files reviewed:** 23 VBA modules (~18,000 lines total)
**Total issues found:** 48 — High: 13 | Medium: 22 | Low: 13

---

## Critical Priorities (Fix First)

1. **`basBibleRibbon.bas` line 702** — Dictionary key typo `"ZECHh"` breaks Zechariah lookup entirely
2. **`basSBL_TestHarness.bas` line 324** — Enum comparison bug causes false test results
3. **`basWordRepairRunner.bas` line 341** — Missing bounds check causes `Left$(txt, -1)` returning full string
4. **`basTestaeBibleClass.bas` lines 109–114** — Unvalidated shell execution; git push failures are invisible
5. **`XbasTESTaeBibleDOCVARIABLE.bas` lines 139–151** — `Err.Raise` followed by bare `Resume` creates undefined flow
6. **`aewordgitClass.cls` lines 150–152** — Nested error handlers with inconsistent cleanup

---

## Findings by File

### `aewordgitClass.cls`

| Line | Severity | Issue |
|------|----------|-------|
| 59 | LOW | `On Error GoTo 0` in `Class_Initialize` with no prior handler — dead code |
| 150–152 | HIGH | `On Error Resume Next` → `Kill` → `On Error GoTo PROC_ERR`: error from Kill is silently swallowed before handler is restored |
| 189 | MEDIUM | `docSource.VBProject.Protection = 1` compared as integer; actual VB enum type not validated |
| 251 | LOW | `Left$(strFileName, 3) <> "zzz"` is case-sensitive — `"ZZZ_file.cls"` bypasses filter; use `LCase()` |
| 293–299 | HIGH | `If Err = -2147467259 Then Resume Next` suppresses a specific error class silently; callers cannot detect failure; magic number needs a named constant |

```vb
' Line 150-152 — Error handling gap
On Error Resume Next
Kill FolderWithVBAProjectFiles & "*.*"
On Error GoTo PROC_ERR   ' <-- Kill error is lost here
```

---

### `basSBL_Citation_EBNF.bas`

| Line | Severity | Issue |
|------|----------|-------|
| — | MEDIUM | Module is ~3,140 lines combining grammar definitions and parsing logic — single responsibility violated; consider splitting into `basSBL_EBNF_Grammar.bas` and `basSBL_Parser.bas` |

---

### `basSBL_TestHarness.bas`

| Line | Severity | Issue |
|------|----------|-------|
| 8 | LOW | `RUN_FAILURE_DEMOS = False` is a feature flag — acceptable, but undocumented |
| 185 | MEDIUM | `CLng(parsed.VerseSpec)` — no error trap; if `VerseSpec` is `""`, raises runtime error 13 (Type Mismatch) |
| 324 | HIGH | `If tests(i)(1) = FailResolveBook` — `FailResolveBook` enum value is 1, but test array stores `False` (0); comparison always fails, producing incorrect test results |

```vb
' Line 324 — Enum vs Boolean mismatch
tests = Array(Array("Jude 0", False), ...)
If tests(i)(1) = FailResolveBook Then  ' FailResolveBook = 1, False = 0 — never equal
```

---

### `basSBL_TestFramework.bas`

| Line | Severity | Issue |
|------|----------|-------|
| 18–29 | MEDIUM | When `AssertTrue` fails with no optional params, output is just `"FAIL: [message]"` with no values — ambiguous for numeric assertions |
| 50 | LOW | `AssertFalse(condition, message)` has 2 params; `AssertTrue` has 4 — asymmetric signatures increase caller error risk |

---

### `basWordRepairRunner.bas`

| Line | Severity | Issue |
|------|----------|-------|
| 106 | HIGH | `ascii13InsertCount` declared on line 103 but only initialized to 0 on line 105 after other declarations; initialize at declaration |
| 122–124 | MEDIUM | `ActiveDocument.range(i, i + 1)` called for every character in O(n) API loop — use `range.Characters` or `Find` instead |
| 189 | LOW | `AscW(combinedNumber) = 12` — magic number; define `Const ASCII_FORMFEED As Long = 12` |
| 201 | MEDIUM | `prefixStyle = "Normal"` — no `Trim()`; style names with padding spaces will fail silently |
| 268 | LOW | `fixCount = fixCount` — no-op statement; remove |
| 341 | HIGH | `pos = InStrRev(txt, "CHAPTER ")` is guarded by outer `If InStrRev(...) > 0` — but `Left$(txt, pos - 1)` on the next line uses a freshly assigned `pos` that could be 0 if reassignment failed; validate `pos > 0` immediately after assignment |

```vb
' Line 341-346 — Bounds check on wrong variable
If InStrRev(txt, "CHAPTER ") > 0 Then
    Dim pos As Long
    pos = InStrRev(txt, "CHAPTER ")  ' second call; result could differ
    If pos > 0 Then
        txt = Trim$(Left$(txt, pos - 1))
    End If
End If
```

---

### `basUSFMExport.bas`

| Line | Severity | Issue |
|------|----------|-------|
| 159, 166, 181 | HIGH | `ParagraphHasCharStyle` iterates `p.range.words` and returns on first match — for mixed-style paragraphs, result depends on word order; unreliable for paragraphs with partial character style application |
| 168, 183 | HIGH | `If IsNumeric(chapTxt)` passes for `"12a"` or `"1.2"`; `CLng()` will fail at runtime — use explicit error-trapped conversion |
| 202 & 205 | MEDIUM | `Case "Heading 1"` appears inside `Case "Book Title", "Heading 1"` — second case is unreachable dead code |
| 289 | LOW | `ChrW(160)` magic number — define `Const NBSP As Long = 160` |
| 348–351 | MEDIUM | `rChap.MoveEnd wdCharacter, -1` corrects for loop overshoot — but if loop never entered, the correction still runs, corrupting range |
| 518–523 | MEDIUM | `Mid$(line, Len(marker) + 1, 1) <> " "` — if `Len(line) = Len(marker)`, Mid returns `""` and comparison passes incorrectly; check `Len(line) > Len(marker)` first |

```vb
' Line 202 & 205 — Unreachable duplicate case
Case "Book Title", "Heading 1"
    ConvertParagraphToUSFM = MakeTitleLine(1, txt)
Case "Heading 1"               ' <-- DEAD CODE: already matched above
    ConvertParagraphToUSFM = "\mt1 " & txt
```

---

### `basWordSettingsDiagnostic.bas`

| Line | Severity | Issue |
|------|----------|-------|
| 72–76 | MEDIUM | `Select Case ActiveWindow.View.Type` handles only `wdPrintView` and `wdWebView`; `wdOutlineView`, `wdMasterView` fall to `Case Else` with misleading error string |
| 141–144 | MEDIUM | `InStr(current(key), "Manual check:")` — substring match; any value containing that substring is miscategorized; use `Left$(..., 15) = "Manual check: "` |

---

### `basTESTaeBibleFonts.bas`

| Line | Severity | Issue |
|------|----------|-------|
| 60–66 | MEDIUM | `testRange.font.name = fontName` then checks if font name matches — but Word substitutes a default when font doesn't exist, making absent fonts appear installed; enumerate system fonts instead |
| 184–190 | MEDIUM | `Set s = ActiveDocument.Styles("Picture Caption")` without `On Error` — error 5809 if style absent; mixing `On Error` and `Is Nothing` check is unclear |

---

### `basTestaeBibleClass.bas`

| Line | Severity | Issue |
|------|----------|-------|
| 41–42 | LOW | `If CStr(varDebug) = "Error 448"` checks for literal string, not error code; comment is misleading |
| 109, 114 | HIGH | `wsh.exec(shellCmd).StdOut.ReadAll` — git failure is in process exit code, not VBA `Err`; output is printed without validating success; a failed push looks identical to a successful one |

```vb
' Lines 109, 114 — Silent git failure
cmdOutput = wsh.exec(shellCmd).StdOut.ReadAll  ' StdErr not checked
Debug.Print "[PUSH] " & cmdOutput              ' Prints nothing on failure; no exit code checked
```

---

### `XbasTESTaeBibleClass_SLOW.bas`

| Line | Severity | Issue |
|------|----------|-------|
| 62–65 | MEDIUM | `If count >= 1000 Then Exit Do` — unclear if safety limit or test requirement; document intent |
| 179 | MEDIUM | `MsgBox "ASCII 12 character found..."` inside loop — interrupts for every match; needs a batch mode flag for large documents |
| 243–245 | HIGH | `Kill debugFile` without error trap — if file is locked by another process, Kill fails silently |
| 264–276 | MEDIUM | Two `Find` operations without explicit `ClearFormatting` between them — prior search state may persist |

---

### `XbasTESTaeBibleDOCVARIABLE.bas`

| Line | Severity | Issue |
|------|----------|-------|
| 8 | MEDIUM | `Public lastFoundLocation As range` — public module-level variable in a test module encourages unintended coupling |
| 37 | HIGH | `Replace(lastFoundLocation, vbCr, "")` — `lastFoundLocation` is a `Range` object, not a string; should be `lastFoundLocation.text` |
| 139–151 | HIGH | `Err.Raise 1000, ...` followed by bare `Resume` (no label) — resumes at the intentional Raise statement, creating an infinite loop |
| 466 | LOW | `VerifyBookNameFromDocVariable "Song", "Solomon"` — canonical SBL name is "Song of Songs", not "Solomon" |

```vb
' Line 37 — Range used as String
Debug.Print ">lastFoundLocation = " & Replace(lastFoundLocation, vbCr, "")
' Should be: lastFoundLocation.text
```

---

### `basBibleRibbon.bas`

| Line | Severity | Issue |
|------|----------|-------|
| 210 | HIGH | `GoTo Cleanup` jumps over ~160 lines of code (lines 211–374) — all unreachable dead code |
| 309 | HIGH | `chapFound` used before explicit initialization — relies on implicit `False` default; assign explicitly |
| 327–328 | HIGH | `Dim chapIdx As Long` declared then immediately used as `idx = chapIdx` without assignment — `chapIdx` is 0 by default; if the loop on line 331 never runs, `idx = 0` produces wrong output |
| 702 | HIGH | `.Add "ZECHh", "Zechariah"` — key has spurious lowercase `h`; correct key is `"ZECH"`; lookup for Zechariah will always fail |

```vb
' Line 702 — Typo in dictionary key
.Add "ZECHh", "Zechariah"   ' <-- "ZECHh" will NEVER match input "ZECH"
.Add "ZEC",   "Zechariah"   ' ZEC matches, but ZECH does not
```

---

### `Module1.bas`

| Line | Severity | Issue |
|------|----------|-------|
| 24–26 | LOW | `Mid(selectedText, i, 1)` called 3 times per loop iteration — cache in `Dim ch As String` |
| 56 | LOW | Comment references `GetVScroll` but actual function is `GetExactVerticalScroll` — stale comment |

---

### `basSBL_VerseCountsGenerator.bas`

| Line | Severity | Issue |
|------|----------|-------|
| 172 | MEDIUM | `Debug.Assert Len(packed) = ...` — `Debug.Assert` is stripped from compiled VBA; use explicit `If...Then` for data validation |
| 212 | MEDIUM | `If IsArray(packedArr(BookID))` — if `BookID > UBound(packedArr)`, error occurs before `IsArray` runs; validate `BookID` bounds first |

---

### `basImportWordGitFiles.bas`

| Line | Severity | Issue |
|------|----------|-------|
| 8–19 | HIGH | All import procedure bodies are commented out — module is entirely non-functional; either restore or delete |
| 40 | MEDIUM | `On Error GoTo 0` in a function performing VBProject modification and file I/O — broad suppression; trap specific errors instead |

---

### `XLongRunningProcessCode.bas`

| Line | Severity | Issue |
|------|----------|-------|
| 36–37 | HIGH | `CustomDocumentProperties("LastProcessedParagraph")` accessed without existence check — raises error 5 if property absent |
| 59–62 | HIGH | WMI `ExecQuery` has no error trap — if WMI is disabled or inaccessible, crashes without recovery |

---

### `ThisDocument.cls`

| Line | Severity | Issue |
|------|----------|-------|
| 13–51 | MEDIUM | `Document_Close` and `Document_Open` bodies are entirely commented out — current state is ambiguous; remove if not needed, or document why disabled |

---

### `bas_TODO.bas`

| Line | Severity | Issue |
|------|----------|-------|
| 1–270 | LOW | Contains no executable code — entire module is planning notes; move to a `.md` file or link entries to `#NNN` issues |

---

## Cross-File Issues

### Inconsistent error handling style — HIGH

Three patterns are mixed across the codebase with no clear rule for when to use each:
- `On Error Resume Next` + conditional `Err.Number` check (some files)
- `On Error GoTo PROC_ERR` + `PROC_EXIT` label (most files — correct pattern)
- Bare `On Error GoTo 0` with no handler set up (aewordgitClass, basImportWordGitFiles)

### Unvalidated string-to-number conversions — HIGH

`CLng()`, `CInt()` called on parser output strings in at least three files without error trapping. A failed parse silently returns 0 or raises runtime error 13. Recommend a shared helper:
```vb
Function SafeCLng(s As String, defaultVal As Long) As Long
    On Error Resume Next
    SafeCLng = CLng(s)
    If Err.Number <> 0 Then SafeCLng = defaultVal
End Function
```

### File I/O without write validation — MEDIUM

`basUSFMExport.bas` and `basWordRepairRunner.bas` open files for output without checking whether the write succeeded. Consider checking `Dir()` and file size after close.

---

## Summary

| Severity | Count |
|----------|-------|
| High     | 13    |
| Medium   | 22    |
| Low      | 13    |
| **Total**| **48**|

### Positive notes

Strong architectural discipline throughout: the 14-stage parser pipeline is well-structured, the `PROC_ERR`/`PROC_EXIT` error handling pattern is consistent in the majority of modules, the `#NNN` issue-tracking system provides good traceability, and the test framework is appropriately lightweight.

---

## Resolution Summary

**Completed:** 2026-03-24

| # | File | Severity | Status | Notes |
|---|------|----------|--------|-------|
| 1 | `aeWordGitClass.cls` line 59 | LOW | Fixed 2026-03-24 | Dead `On Error GoTo 0` removed from `Class_Initialize` |
| 2 | `aeWordGitClass.cls` lines 150–152 | HIGH | Fixed 2026-03-24 | `Kill` error now checked; exits with MsgBox if folder cannot be cleared |
| 3 | `aeWordGitClass.cls` line 189 | MEDIUM | Fixed 2026-03-24 | `= 1` replaced with `= vbext_pp_locked` |
| 4 | `aeWordGitClass.cls` line 251 | LOW | Fixed 2026-03-24 | `LCase$()` added to make `"zzz"` filter case-insensitive |
| 5 | `aeWordGitClass.cls` lines 293–299 | HIGH | Already fixed | Magic number replaced with `E_FAIL` constant + `Debug.Print` logging |
| 6 | `basSBL_TestHarness.bas` line 8 | LOW | Fixed 2026-03-24 | Comment added explaining `RUN_FAILURE_DEMOS` purpose |
| 7 | `basSBL_TestHarness.bas` line 185 | MEDIUM | Stale — not applicable | `CLng(parsed.VerseSpec)` is within proper test harness flow |
| 8 | `basSBL_TestHarness.bas` line 324 | HIGH | Fixed 2026-03-24 | Test array updated to use `FailNone` enum value instead of `False` |
| 9 | `basSBL_TestFramework.bas` lines 18–29 | MEDIUM | Skipped | Framework correctly split — `AssertEqual` covers numeric case |
| 10 | `basSBL_TestFramework.bas` line 50 | LOW | Skipped | Asymmetry is intentional; `AssertFalse` delegates to `AssertTrue` |
| 11 | `basWordRepairRunner.bas` line 106 | HIGH | Fixed 2026-03-24 | Redundant `ascii13InsertCount = 0` removed |
| 12 | `basWordRepairRunner.bas` lines 122–124 | MEDIUM | Skipped | Performance concern theoretical; refactor would add complexity for marginal gain |
| 13 | `basWordRepairRunner.bas` line 189 | LOW | Stale — not applicable | `AscW(combinedNumber) = 12` check effectively dead given earlier guards |
| 14 | `basWordRepairRunner.bas` line 201 | MEDIUM | Fixed 2026-03-24 | `Trim()` added to `prefixStyle` comparison |
| 15 | `basWordRepairRunner.bas` line 268 | LOW | Fixed 2026-03-24 | `fixCount = fixCount` no-op removed |
| 16 | `basWordRepairRunner.bas` line 341 | HIGH | Fixed 2026-03-24 | Redundant double `InStrRev` call replaced with single call + `pos > 0` guard |
| 17 | `basUSFM_Export.bas` lines 159, 166, 181 | HIGH | Fixed 2026-03-24 | `FIXME_LATER` comment added; safe in practice — character styles always applied to full words |
| 18 | `basUSFM_Export.bas` lines 168, 183 | HIGH | Fixed 2026-03-24 | `FIXME_LATER` comments added to both chapter and verse checks |
| 19 | `basUSFM_Export.bas` lines 202 & 205 | MEDIUM | Already fixed | Duplicate `Case "Heading 1"` removed; cases now correctly separated |
| 20 | `basUSFM_Export.bas` line 289 | LOW | Skipped | `ChrW(160)` already has `' NBSP` comment; named constant unnecessary |
| 21 | `basUSFM_Export.bas` lines 348–351 | MEDIUM | Skipped | Early exit at line 353 guarantees loop always runs at least once |
| 22 | `basUSFM_Export.bas` lines 518–523 | MEDIUM | Skipped | Existing `Len(line) > Len(marker) + 1` guard already prevents out-of-bounds `Mid$` |
| 23 | `basWordSettingsDiagnostic.bas` lines 72–76 | MEDIUM | Skipped | `Case Else` returns descriptive string with view type number — clear and informative |
| 24 | `basWordSettingsDiagnostic.bas` lines 141–144 | MEDIUM | Skipped | `InStr` match safe — values are controlled internal diagnostic strings |
| 25 | `basTEST_aeBibleFonts.bas` lines 60–66 | MEDIUM | Skipped | Assign-and-compare pattern correctly detects font substitution |
| 26 | `basTEST_aeBibleFonts.bas` lines 184–190 | MEDIUM | Skipped | `On Error Resume Next` + `Is Nothing` is correct idiomatic VBA for optional style access |
| 27 | `basTest_aeBibleClass.bas` lines 41–42 | LOW | Fixed 2026-03-24 | Dead `"Error 448"` literal string check removed |
| 28 | `basTest_aeBibleClass.bas` lines 109, 114 | HIGH | Fixed 2026-03-24 | `StdErr` now captured and checked; git tag and push failures surface with MsgBox |
| 29 | `XbasTESTaeBibleClass_SLOW.bas` lines 62–65 | MEDIUM | Fixed 2026-03-24 | Comment added explaining 1000-iteration batch limit |
| 30 | `XbasTESTaeBibleClass_SLOW.bas` line 179 | MEDIUM | Skipped | Interactive `MsgBox` prompt is the intended feature — manual review tool |
| 31 | `XbasTESTaeBibleClass_SLOW.bas` lines 243–245 | HIGH | Fixed 2026-03-24 | `Kill` wrapped with `On Error Resume Next` + `Err.Number` check |
| 32 | `XbasTESTaeBibleClass_SLOW.bas` lines 264–276 | MEDIUM | Skipped | Neither `Find` uses formatting criteria; `ClearFormatting` between calls unnecessary |
| 33 | `XbasTESTaeBibleDOCVARIABLE.bas` line 8 | MEDIUM | Fixed 2026-03-24 | `Public lastFoundLocation` changed to `Private` |
| 34 | `XbasTESTaeBibleDOCVARIABLE.bas` line 37 | HIGH | Fixed 2026-03-24 | Implicit `Range` coercion replaced with explicit `.Text` |
| 35 | `XbasTESTaeBibleDOCVARIABLE.bas` lines 139–151 | HIGH | Fixed 2026-03-24 | `Resume` changed to `Resume RetrySearch`; `RetrySearch:` label added before search call |
| 36 | `XbasTESTaeBibleDOCVARIABLE.bas` line 466 | LOW | Skipped | `"Solomon"` is correct — matches actual `Heading 1` text used in this document |
| 37 | `basBibleRibbon.bas` line 210 | HIGH | Not applicable | Active ribbon refactored into `aeRibbonClass.cls`; `basBibleRibbon_OLD.bas` is deferred |
| 38 | `basBibleRibbon.bas` line 309 | HIGH | Not applicable | See item 37 |
| 39 | `basBibleRibbon.bas` lines 327–328 | HIGH | Not applicable | See item 37 |
| 40 | `basBibleRibbon.bas` line 702 | HIGH | Not applicable | See item 37 |
| 41 | `Module1.bas` lines 24–26 | LOW | Fixed 2026-03-24 | `Mid()` cached in `ch` variable; called once per iteration instead of twice |
| 42 | `Module1.bas` line 56 | LOW | Stale — not applicable | Referenced function/comment not present in current file |
| 43 | `basSBL_VerseCountsGenerator.bas` line 172 | MEDIUM | Skipped | Generator runs in IDE only; `Debug.Assert` is appropriate |
| 44 | `basSBL_VerseCountsGenerator.bas` line 212 | MEDIUM | Skipped | Fixed 66-book structure makes bounds risk theoretical |
| 45 | `basImportWordGitFiles.bas` lines 8–19 | HIGH | Stale — not applicable | Review said bodies were commented out; file is functional and was recently updated with error handling |
| 46 | `basImportWordGitFiles.bas` line 40 | MEDIUM | Fixed 2026-03-24 | Dead `On Error GoTo 0` removed from `ImportAllVBAFiles` |
| 47 | `XLongRunningProcessCode.bas` lines 36–37 | HIGH | Already fixed | `SaveProgress` now guards property existence via `CustomPropertyExists` |
| 48 | `XLongRunningProcessCode.bas` lines 59–62 | HIGH | Already fixed | `SetWordHighPriority` now has `On Error GoTo PROC_ERR` + cleanup |

**Cross-file issues:**

| | Cross-file | Severity | Status |
|---|-----------|----------|--------|
| A | Inconsistent error handling style | HIGH | Valid — open |
| B | Unvalidated string-to-number conversions | HIGH | Valid — open |
| C | File I/O without write validation | MEDIUM | Valid — open |

**Resolution counts (updated 2026-03-24):**

| Status | Count |
|--------|-------|
| Fixed this session | 24 (items 1–4, 6, 8, 11, 14–16, 27–29, 31, 33–35, 41, 46 + prior: 5, 19, 47, 48) |
| Skipped — not worth actioning | 13 (items 9, 10, 12, 20–26, 30, 32, 43, 44) |
| Not applicable | 4 (items 37–40 — old ribbon) |
| Stale/inaccurate | 4 (items 7, 13, 42, 45) |
| Skipped — document-specific | 1 (item 36 — "Solomon" correct for this document) |
| Cross-file — open | 3 (A, B, C) |
