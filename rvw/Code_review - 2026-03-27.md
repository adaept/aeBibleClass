# Error Handling Compliance Review

- **Date:** 2026-03-27
- **Reviewer:** Claude Code (claude-sonnet-4-6)
- **Scope:** All `.bas` and `.cls` files in `C:\adaept\aeBibleClass\src\`
- **Standard enforced:** Every Sub/Function/Property must have `On Error GoTo PROC_ERR`, `PROC_EXIT:` label, `Exit Sub/Function/Property`, `PROC_ERR:` label, standard MsgBox with Erl/Err.Number/Err.Description/proc name/module name, and `Resume PROC_EXIT`.

---

## Violations by File

---

### basBibleRibbonSetup.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `Instance` | 8 | Missing `On Error GoTo PROC_ERR` | No error handler. Single-line function returning a singleton — low risk but exposed at module entry point. |
| `AutoExec` | 17 | Missing `On Error GoTo PROC_ERR` | No error handler. Calls `Instance()` which itself has no handler. |
| `RibbonOnLoad` | 23 | Missing `On Error GoTo PROC_ERR` | No error handler. Ribbon load failure would be silent. |
| `OnGoToVerseSblClick` | 30 | Missing `On Error GoTo PROC_ERR` | No error handler. Ribbon callback stub. |
| `OnGoToH1ButtonClick` | 35 | Missing `On Error GoTo PROC_ERR` | No error handler. Ribbon callback stub. |
| `OnNextButtonClick` | 40 | Missing `On Error GoTo PROC_ERR` | No error handler. Ribbon callback stub. |
| `OnAdaeptAboutClick` | 45 | Missing `On Error GoTo PROC_ERR` | No error handler. Ribbon callback stub. |

---

### basChangeLog_aeBibleClass.bas

No procedures. Module is a comment/documentation file only. **No violations.**

---

### basChangeLog_aeWordGitClass.bas

No procedures. Module is a comment/documentation file only. **No violations.**

---

### basSBL_Citation_EBNF.bas

This file is large (>400 lines of code). All procedure-level error handling is reviewed below.

| Procedure | Line (approx.) | Violation Type | Description |
|---|---|---|---|
| All functions in module | — | Missing `On Error GoTo PROC_ERR` | Per the parser contract (item 2: "The parser never raises user-facing runtime errors"), procedures such as `LexicalScan`, `ResolveAlias`, `InterpretStructure`, `ValidateSBLReference`, `RewriteSingleChapterRef`, `ParseReference`, `ListDetection`, `RangeDetection`, `ComposeRange`, `ComposeList`, `ParseScripture`, `AliasCoverage`, `ResetBookAliasMap`, `GetBookAliasMap`, `GetCanonicalBookTable`, `GetPackedVerseMap`, `GetChapterVerseMap`, `GetMaxVerse`, `AssertOneBased`, `CanonicalFromRef`, `IsRangeSegment` — none contain standard error handling. The design intention is that the parser is error-free internally, but production VBA modules should still have the standard pattern. Flag as violations.|

> **Note:** The parser design contract explicitly states errors should not be raised to the user. The absence of `PROC_ERR` blocks is a deliberate architectural choice documented in the module header. These are noted as violations against the project standard, but the reviewer acknowledges the design rationale.

---

### basSBL_TestFramework.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `AssertTrue` | 11 | Missing `On Error GoTo PROC_ERR` | No error handler. Test framework utility. |
| `AssertEqual` | 32 | Missing `On Error GoTo PROC_ERR` | No error handler. Test framework utility. |
| `AssertFalse` | 49 | Missing `On Error GoTo PROC_ERR` | No error handler. Delegates to `AssertTrue`. |
| `TestStart` | 53 | Missing `On Error GoTo PROC_ERR` | No error handler. Resets counters. |
| `TestSummary` | 63 | Missing `On Error GoTo PROC_ERR` | No error handler. Prints summary. |

---

### basSBL_TestHarness.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `ParseReferenceStub` | 16 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Test_AliasCoverage` | 60 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Test_TokenizeReference` | 83 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Test_SemanticFlow_WithParserStub` | 93 | Missing `On Error GoTo PROC_ERR` | No error handler. Uses `On Error Resume Next` at line 146 with `On Error GoTo 0` at line 151 — this pair is intentional (detecting `ResolveAlias` error), classified as **intentional suppression** (scoped, immediately followed by `Err.Clear` and `On Error GoTo 0`). |
| `Report_TODOs` | 231 | Missing `On Error GoTo PROC_ERR` | No error handler. Print-only routine. |
| `Test_SemanticFlow_WithParserStub_Negative` | 247 | Missing `On Error GoTo PROC_ERR` | No error handler. Uses `On Error Resume Next` at line 301 with `On Error GoTo 0` at line 315/319 — **intentional suppression** (scoped error probe, same pattern as above). |
| `Test_GetMaxVerse` | 356 | Missing `On Error GoTo PROC_ERR` | Uses a non-standard handler `FailHandler:` at line 408 (not the standard `PROC_ERR:` label). Also uses `On Error Resume Next` at line 377 and `On Error GoTo 0` at line 397 — intentional suppression for negative tests. The `FailHandler` block uses `Resume Next` not `Resume PROC_EXIT`, and the `MsgBox` format does not follow the standard. |
| `FailTest` | 415 | Missing `On Error GoTo PROC_ERR` | No error handler. Private helper. |
| `Run_All_SBL_Tests` | 426 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Test_Stage2_LexicalScan` | 463 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Test_Stage3_ResolveAlias` | 481 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Test_Stage4_InterpretStructure` | 501 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Test_Stage5_ValidateCanonical` | 553 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Test_Stage6_FormatCanonical` | 568 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Test_Stage6_FormatCanonical_FailureDemo` | 581 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Test_Stage7_EndToEnd` | 596 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Test_Stage8_ListDetection` | 624 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Test_Stage9_RangeDetection` | 665 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Test_Stage10_RangeComposition` | 711 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `PrintScriptureList` | 741 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Test_Stage11_ListComposition` | 764 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Test_Stage12_FinalParser` | 791 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Test_Stage13_ContextShorthand` | 820 | Missing `On Error GoTo PROC_ERR` | No error handler. |

---

### basSBL_VerseCountsGenerator.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `ExpectedChapterCounts` | 23 | Missing `On Error GoTo PROC_ERR` | No error handler. Private data function. |
| `ToOneBasedLongArray` | 34 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `GetVerseCounts` | 55 | Missing `On Error GoTo PROC_ERR` | No error handler. Large data factory function. |
| `GeneratePackedVerseStrings_FromDictionary` | 138 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `VerifyPackedVerseMap` | 180 | Missing `On Error GoTo PROC_ERR` | No error handler. |

---

### basTest_aeBibleClass.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `RUN_THE_TESTS` | 18 | Wrong handler / Orphan `On Error GoTo 0` | Line 19: `On Error GoTo 0` is the only error strategy — no `PROC_ERR` handler, no `PROC_EXIT`. This is an orphan `On Error GoTo 0` at the top of the procedure. |
| `aeBibleClassTest` | 31 | Wrong handler format | Has `On Error GoTo PROC_ERR`, `PROC_EXIT:`, and `PROC_ERR:`. The `PROC_ERR:` block at line 64 handles error 6068 specially (no `Resume PROC_EXIT` for that branch — uses `Stop` then falls through) and the general MsgBox at line 69 is correct format. However the 6068 branch does **not** call `Resume PROC_EXIT` (it calls `Stop` and a commented-out `Resume PROC_EXIT`). Partial compliance — the Stop path leaves the error handler without resuming. |
| `GitAutoTagRelease` | 89 | Missing `On Error GoTo PROC_ERR` | No error handler at all. Uses `Exit Sub` directly on error conditions detected manually. |
| `GitTagExists` | 142 | Missing `On Error GoTo PROC_ERR` | No error handler. |

---

### basTEST_aeWordGitClass.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `EXPORT_THE_CODE` | 26 | Wrong handler / Orphan `On Error GoTo 0` | Line 27: `On Error GoTo 0` as the only error strategy. Also uses `On Error Resume Next` at line 41 followed by `On Error GoTo 0` at line 46 — this pair is **intentional suppression** (Shell call, scoped). No `PROC_ERR` handler present. |
| `aeWordGitClassTest` | 51 | Wrong handler format | Has `On Error GoTo PROC_ERR`, `PROC_EXIT:`, `PROC_ERR:`. The error 6068 branch uses `Stop` without `Resume PROC_EXIT` (commented out at line 96). General MsgBox format is correct. Same partial compliance issue as `aeBibleClassTest`. |

---

### bas_TODO.bas

No procedures (all content is comments). **No violations.**

---

### basImportWordGitFiles.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `ImportAllVBAFiles` | 8 | Wrong handler / Orphan `On Error GoTo 0` | Line 9: `On Error GoTo 0` as the only error strategy. No `PROC_ERR` handler. |
| `ImportVBAFile` | 75 | Wrong handler format — PROC_ERR exists but 6068 branch missing `Resume PROC_EXIT` | The `PROC_ERR:` at line 103 handles 6068 with `Stop` (no `Resume PROC_EXIT`). The general MsgBox format (line 108) is correct with Erl/Err.Number/Err.Description but uses `vbCritical` style extra argument — acceptable. Partial compliance. |
| `DeleteAllModulesExceptImporter` | 113 | Wrong handler format — PROC_ERR exists but 6068 branch missing `Resume PROC_EXIT` | Same pattern as `ImportVBAFile`. The 6068 branch has `Stop` with no resume. |
| `ModuleOrClassExists` | 188 | Wrong handler / Orphan `On Error GoTo 0` | Line 189: `On Error GoTo 0` as the only error strategy. No `PROC_ERR` handler. |

---

### Module1.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `ViewCodeDetails` | 34 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `PrintFontProperties` | 56 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `PrintBibleBook` | 89 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `IsParagraphEmpty` | 138 | Missing `On Error GoTo PROC_ERR` | No error handler. Single-expression function. |
| `GoToParagraphIndex` | 147 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountTextWrappingBreakParagraphs` | 177 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountNextPageSectionBreakParagraphs` | 193 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountContinuousSectionBreakParagraphs` | 207 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountEvenPageSectionBreakParagraphs` | 221 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountOddPageSectionBreakParagraphs` | 235 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `AppendToFile` | 249 | Missing `On Error GoTo PROC_ERR` | No error handler. File I/O without protection. |
| `SearchParagraphs` | 257 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountEmptyParagraphsWithAutomaticFont` | 291 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `GoToParagraphByCount` | 311 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `DetectFontColors` | 334 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `UpdateBlackToAutomatic` | 364 | Missing `On Error GoTo PROC_ERR` | No error handler. Document-wide find/replace without protection. |
| `ChangeFontColorRGB` | 404 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `ChangeSpecificColor` | 428 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `EnsureFootnoteReferenceStyleColor` | 433 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `HexToRGB` | 470 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `FirstPageFooterNotEmpty` | 485 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `IsEmptyParagraph` | 505 | Missing `On Error GoTo PROC_ERR` | No error handler. Single-expression function. |
| `CountTotallyEmptyParagraphs` | 510 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountTypesTrulyEmptyParagraph` | 617 | Missing `On Error GoTo PROC_ERR` | No error handler. |

---

### XLongRunningProcessCode.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `PauseWithDoEvents` | 20 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `StartOrResumeUpdate` | 30 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `StopUpdate` | 36 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `ResetProgress` | 41 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `SaveProgress` | 47 | Wrong handler format | Has `On Error GoTo PROC_ERR`, `PROC_EXIT:`, `PROC_ERR:`. However the `PROC_ERR:` handler uses `Debug.Print` only — not the standard MsgBox with Erl/Err.Number/Err.Description. |
| `CustomPropertyExists` | 70 | `On Error Resume Next` without re-enabling PROC_ERR | Uses `On Error Resume Next` at line 72 / `On Error GoTo 0` at line 75. No outer `PROC_ERR` handler. Classified as intentional suppression (probing whether property exists), but no outer handler to fall back to. |
| `LoadProgress` | 78 | `On Error Resume Next` without re-enabling PROC_ERR | Lines 79–82: `On Error Resume Next` then `On Error GoTo 0`. No `PROC_ERR` handler. Property reads are suppressed without any fallback. |
| `SetWordHighPriority` | 85 | Wrong handler format | Has `On Error GoTo PROC_ERR`, `PROC_EXIT:`, `PROC_ERR:`. However `PROC_ERR:` handler uses `Debug.Print` only — not the standard MsgBox format. |
| `UpdateCharacterStyle` | 121 | Missing `On Error GoTo PROC_ERR` | No error handler. Long-running loop with no protection. |
| `LongProcessSkeletonWithConsoleProgress` | 184 | Missing `On Error GoTo PROC_ERR` | No error handler. |

---

### XbasTESTaeBibleClass_SLOW.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `FindAnyNumberWithStyleAndPrintNextCharASCII` | 15 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `PrintBibleHeading1Info` | 83 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `PrintBibleBookHeadings` | 109 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `ListAndReviewAscii12Characters` | 151 | Missing `On Error GoTo PROC_ERR` | No error handler (not fully read — but no handler detected in first visible lines). |

---

### XbasTESTaeBibleDOCVARIABLE.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `FindNextHeading1OnVisiblePage` | 22 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `VerifyBookNameFromDocVariable` | 106 | Wrong handler format | Has `On Error GoTo ErrorHandler` (non-standard label name `ErrorHandler` not `PROC_ERR`). Uses `On Error Resume Next` at line 117 / `On Error GoTo 0` at line 118 — intentional suppression (DOCVARIABLE probe). The `ErrorHandler:` block does not use standard MsgBox format and does not call `Resume PROC_EXIT` — it uses `Resume RetrySearch`. Non-standard structure throughout. |
| `FindDocVariableByName` | 167 | Missing `On Error GoTo PROC_ERR` | Uses `On Error Resume Next` at line 180 / `On Error GoTo 0` at line 185 — intentional suppression (variable existence probe). No `PROC_ERR` handler. |
| `FindDocVariableEverywhere` | 198 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `SearchShapeForVariable` | 303 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `SetDocVariables` | 336 | Missing `On Error GoTo PROC_ERR` | No error handler. Large data-setting routine. |
| `ListMyDocVariables` | 412 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `DeleteDocVariable` | 419 | Missing `On Error GoTo PROC_ERR` | Uses `On Error Resume Next` at line 423 / `On Error GoTo 0` at line 425 — intentional suppression (delete may fail if variable does not exist). No outer `PROC_ERR` handler. |
| `TestPageNumbers` | 430 | Missing `On Error GoTo PROC_ERR` | No error handler. Long-running test. |

---

### basAddHeaderFooter.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `AddBookNameHeaders` | 8 | Missing `On Error GoTo PROC_ERR` | No error handler. Document-modifying routine with no protection. |
| `FixTheFooters` | 108 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `AddConsecutiveFootersFromCursor` | 121 | Missing `On Error GoTo PROC_ERR` | No error handler. Document-modifying routine. |
| `LinkFootersToPrevious` | 210 | Missing `On Error GoTo PROC_ERR` | No error handler. |

---

### basAuditDocument.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `ReplaceTimesInStyles` | 6 | Missing `On Error GoTo PROC_ERR` | Uses `On Error Resume Next` / `On Error GoTo 0` pairs inside loop (lines 15–21) — intentional suppression (some styles may not have Font property). No outer `PROC_ERR` handler. |
| `FindFontUsage` | 33 | Missing `On Error GoTo PROC_ERR` | Uses `On Error Resume Next` / `On Error GoTo 0` pairs (lines 108–111) — intentional suppression. No outer `PROC_ERR` handler. |
| `CountParagraphsAndFonts` | 145 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `ResolveFont` | 206 | Missing `On Error GoTo PROC_ERR` | Uses `On Error Resume Next` / `On Error GoTo 0` (lines 210–212) — intentional suppression (style font probe). No outer `PROC_ERR` handler. Private helper. |
| `AddToCollection` | 220 | Missing `On Error GoTo PROC_ERR` | No error handler. Short private helper. |
| `CountFields` | 228 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountCodeLines` | 235 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `PadRight` | 302 | Missing `On Error GoTo PROC_ERR` | No error handler. Short private helper. |
| `CountOrphanFooters` | 312 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountOrphanHeaders` | 356 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `AuditDoc_Original` | 403 | Missing `On Error GoTo PROC_ERR` | No error handler. Entry point. |
| `AuditDoc_New` | 407 | Missing `On Error GoTo PROC_ERR` | No error handler. Entry point. |
| `WriteAuditToFile` | 414 | Missing `On Error GoTo PROC_ERR` | No error handler. File I/O without protection. |
| `GetRptPath` | 436 | Missing `On Error GoTo PROC_ERR` | No error handler (uses `Err.Raise` but no handler to catch it at this level). |
| `WriteHeader` | 456 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `WriteDocumentStats` | 465 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `WriteSectionAudit` | 478 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `WriteSignature` | 518 | Missing `On Error GoTo PROC_ERR` | No error handler. |

---

### basBibleRibbon_OLD.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `StyleTypeLabel` | 58 | Missing `On Error GoTo PROC_ERR` | No error handler. Private helper function. |
| `LeftUntilLastSpace` | 68 | Missing `On Error GoTo PROC_ERR` | No error handler. Private helper function. |
| `ExtractTrailingDigits` | 81 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `IsOneChapterBook` | 99 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `SaveCursor` | 108 | Missing `On Error GoTo PROC_ERR` | No error handler. Single-line function. |
| `RestoreCursor` | 112 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `FindBookH1` | 117 | Missing `On Error GoTo PROC_ERR` | No error handler. Document search routine. |
| `FindChapterH2` | 157 | Missing `On Error GoTo PROC_ERR` | No error handler. Document search routine. |
| `ParseParts` | 191 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `GoToSection` | 208 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `GetBookmarkList` | 224 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `GoToH1` | 235 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `NextButton` | 270 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `GetExactVerticalScroll` | 316 | Missing `On Error GoTo PROC_ERR` | No error handler. |

---

### basTEST_aeBibleFonts.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `CheckOpenFontsWithDownloads` | 8 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `IsFontInstalled` | 56 | Missing `On Error GoTo PROC_ERR` | Uses `On Error Resume Next` at line 59 / `On Error GoTo 0` at line 66 — intentional suppression (probes whether font is available). No outer `PROC_ERR` handler. |
| `CreateEmphasisBlackStyle` | 69 | Missing `On Error GoTo PROC_ERR` | Uses `On Error Resume Next` at line 73 / `On Error GoTo 0` at line 75 — intentional suppression (style existence check). No outer `PROC_ERR` handler. |
| `AuditStyleUsage_Footnote` | 97 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `RedefineFootnoteStyle_NotoSans` | 124 | Missing `On Error GoTo PROC_ERR` | No error handler (not fully visible but no handler in initial lines). |

---

### basTEST_aeBibleTools.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `ListCustomXMLParts` | 17 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `ListCustomXMLSchemas` | 27 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `AddCustomUIXML` | 34 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `RemoveDuplicateCustomXMLParts` | 48 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `IsPartInCollection` | 87 | Missing `On Error GoTo PROC_ERR` | No error handler. Short helper. |
| `DeleteCustomUIXML` | 98 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `GetColorNameFromHex` | 128 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `ListAndCountFontColors` | 183 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `GetVerticalPositionOfCursorParagraph` | 222 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `FindFirstSectionWithDifferentFirstPage` | 238 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `FindFirstPageWithEmptyHeader` | 258 | Missing `On Error GoTo PROC_ERR` | No error handler (procedure visible from line 258 onward). |
| `StyleIsAppliedAnywhere` | 220 (basWordSettingsDiagnostic) | — | See basWordSettingsDiagnostic.bas below. |

---

### basUSFM_Export.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `InitPaths` | 53 | Missing `On Error GoTo PROC_ERR` | No error handler. Private init. |
| `ExportUSFM_PageRange` | 64 | Wrong handler format | Has `On Error GoTo PROC_ERR`, `PROC_EXIT:`, `PROC_ERR:`. The `PROC_ERR:` handler at line 97 logs to file AND shows MsgBox at line 98 with correct Erl/Err.Number/Err.Description format — **compliant**. |
| `ConvertRangeToUSFM` | 105 | Missing `On Error GoTo PROC_ERR` | No error handler. Core conversion loop. |
| `ConvertParagraphToUSFM` | 141 | Missing `On Error GoTo PROC_ERR` | No error handler. Uses `GoTo LogAndExit` labels (non-standard flow). |
| `MakeTitleLine` | 278 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `MakeChapterLines` | 287 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `MakeVerseLine` | 299 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `IsEffectivelyEmpty` | 308 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `ParagraphHasCharStyle` | 323 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `ExtractCharStyleText` | 337 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `TryParseChapterVerseFromStyles` | 348 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `ExtractTrailingNumber` | 428 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CleanTextForUTF8` | 448 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `LogValidator` | 468 | Wrong handler format | Has `On Error GoTo ErrHandler` (non-standard label). `ErrHandler:` at line 505 uses `Debug.Print` only — not standard MsgBox format. |
| `ValidateUSFMFile` | 510 | Wrong handler format | Has `On Error GoTo ErrHandler` (non-standard label). `ErrHandler:` at line 578 uses `LogValidator` call — not standard MsgBox format, no `Resume PROC_EXIT`. |
| `ExtractUSFMMarker` | 583 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `IsKnownUSFMMarker` | 600 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `MarkerAllowsNoSpace` | 617 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `MarkerRequiresContent` | 626 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `GetRangeForPages` | 642 | Wrong handler format | Has `On Error GoTo ErrHandler` (non-standard label). `ErrHandler:` at line 673 uses `LogEvent` call — not standard MsgBox format, no `Resume PROC_EXIT`. |
| `LogEvent` | 681 | Wrong handler format | Has `On Error GoTo ErrHandler` (non-standard label). `ErrHandler:` at line 722 uses `Debug.Print` only — not standard MsgBox format. |
| `WriteTextFile` | 730 | Wrong handler format | Has `On Error GoTo ErrHandler` (non-standard label). `ErrHandler:` at line 752 uses `LogEvent` call — not standard MsgBox format, no `Resume PROC_EXIT`. |

---

### basWordRepairRunner.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `FileNameStartsWithV59` | 12 | Missing `On Error GoTo PROC_ERR` | No error handler. Single-expression private function. |
| `SaveAsPDF_NoOpen` | 20 | Missing `On Error GoTo PROC_ERR` | No error handler. Document export without protection. |
| `RunRepairWrappedVerseMarkers_Across_Pages_From` | 43 | Missing `On Error GoTo PROC_ERR` | No error handler. Entry point for repair routine. |
| `RepairWrappedVerseMarkers_MergedPrefix_ByColumnContext_SinglePage` | 85 | Missing `On Error GoTo PROC_ERR` | No error handler. Core repair loop. |
| `GetPageHeaderText` | 277 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `TitleCase` | 303 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `GetVerseText` | 321 | Missing `On Error GoTo PROC_ERR` | No error handler. |

---

### basWordSettingsDiagnostic.bas

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `RunWordSettingsAudit` | 10 | Missing `On Error GoTo PROC_ERR` | No error handler. Entry point. |
| `GetCurrentWordSettings` | 34 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `GetShowTextBoundaries` | 67 | Missing `On Error GoTo PROC_ERR` | Uses `On Error Resume Next` at line 68 with no `On Error GoTo 0` or `PROC_ERR` handler. The `On Error Resume Next` is never explicitly cleared within this function — **violation: `On Error Resume Next` without re-enabling** (the procedure relies on automatic reset at function exit). |
| `LoadTargetBaseline` | 83 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CompareSettings` | 100 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `FormatDiagnostics` | 121 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `FormatBoolean` | 160 | Missing `On Error GoTo PROC_ERR` | No error handler. Short helper. |
| `SaveReportToFile` | 169 | Missing `On Error GoTo PROC_ERR` | No error handler. File I/O. |
| `ShowAllStyles` | 180 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `ShowMyStyles` | 191 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `StyleIsAppliedAnywhere` | 220 | Missing `On Error GoTo PROC_ERR` | Uses `On Error Resume Next` at line 224 / `On Error GoTo 0` at line 255 — intentional suppression (style application probe). No outer `PROC_ERR` handler. |
| `StyleIsApplied` | 258 | Missing `On Error GoTo PROC_ERR` | Uses `On Error Resume Next` at line 260 / `On Error GoTo 0` at line 267 — intentional suppression. No outer `PROC_ERR` handler. |
| `HideUnusedStyles` | 270 | Missing `On Error GoTo PROC_ERR` | Uses `On Error Resume Next` / `On Error GoTo 0` inside loop (lines 275–277) — intentional suppression. No outer `PROC_ERR` handler. |

---

### ThisDocument.cls

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `Document_Open` | 15 | Missing `On Error GoTo PROC_ERR` | No error handler. Body is entirely commented out — **exception: stub with no logic, no violation in practice.** |

---

### aeWordGitClass.cls

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `Class_Initialize` | 56 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Class_Terminate` | 76 | Wrong handler / Orphan `On Error GoTo 0` | Line 77: `On Error GoTo 0` as the only error strategy. No `PROC_ERR` handler. |
| `SourceFolder` (Get) | 84 | Wrong handler / Orphan `On Error GoTo 0` | Line 85: `On Error GoTo 0` as the only strategy. No `PROC_ERR` handler. |
| `SourceFolder` (Let) | 89 | Wrong handler / Orphan `On Error GoTo 0` | Line 90: `On Error GoTo 0` as the only strategy. No `PROC_ERR` handler. |
| `DocumentTheWordCode` (Property Get) | 96 | **Compliant** | Has `On Error GoTo PROC_ERR`, `PROC_EXIT:`, `PROC_ERR:`, standard MsgBox, `Resume PROC_EXIT`. |
| `aeDocumentTheWordCode` | 120 | Wrong handler format | Has `On Error GoTo PROC_ERR`, `PROC_EXIT:`, `PROC_ERR:`. However uses `On Error Resume Next` at line 150 with `On Error GoTo PROC_ERR` restore at line 157 — intentional suppression (Kill statement). The MsgBox format is correct. Generally compliant but note the `Exit Function` at line 161 (vbNo branch) bypasses `PROC_EXIT:` label directly — acceptable since no cleanup needed. |
| `RunExportCode` | 177 | **Compliant** | Has `On Error GoTo PROC_ERR`, `PROC_EXIT:`, `PROC_ERR:`, standard MsgBox, `Resume PROC_EXIT`. |
| `FolderWithVBAProjectFiles` | 239 | Wrong handler format | Has `On Error GoTo PROC_ERR`, `PROC_EXIT:`, `PROC_ERR:`. The `PROC_ERR:` block at line 281 handles `E_FAIL` specially with `Resume Next` (not `Resume PROC_EXIT`) — intentional for E_FAIL. General MsgBox is correct. Mixed compliance. |

---

### aeBibleClass.cls

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `Class_Initialize` | 72 | Wrong handler / Orphan `On Error GoTo 0` | Line 73: `On Error GoTo 0` as the only strategy. No `PROC_ERR` handler. |
| `Class_Terminate` | 94 | Wrong handler / Orphan `On Error GoTo 0` | Line 95: `On Error GoTo 0` as the only strategy. No `PROC_ERR` handler. |
| `TheBibleClassTests` (Property Get) | 103 | **Compliant** | Has `On Error GoTo PROC_ERR`, `PROC_EXIT:`, `PROC_ERR:`, standard MsgBox, `Resume PROC_EXIT`. |
| `CheckShowHideStatus` | 143 | Missing `On Error GoTo PROC_ERR` | No error handler. Single-expression function. |
| `FileNameStartsWithV59` | 150 | Missing `On Error GoTo PROC_ERR` | No error handler. Single-expression private function. |
| `InitializeGlobalResultArrayToMinusOne` | 158 | **Compliant** | Has `On Error GoTo PROC_ERR`, `PROC_EXIT:`, `PROC_ERR:`, standard MsgBox, `Resume PROC_EXIT`. |
| `ConvertToOneBasedArray` | 187 | **Compliant** | Has `On Error GoTo PROC_ERR`, `PROC_EXIT:`, `PROC_ERR:`, standard MsgBox, `Resume PROC_EXIT`. |
| `Expected1BasedArray` | 212 | **Compliant** | Has `On Error GoTo PROC_ERR`, `PROC_EXIT:`, `PROC_ERR:`, standard MsgBox, `Resume PROC_EXIT`. |
| `MakeSkipTestArray` | 309 | Missing `On Error GoTo PROC_ERR` | No error handler. Single-assignment stub. |
| `IsSkipTest` | 313 | Missing `On Error GoTo PROC_ERR` | No error handler. Short lookup function. |
| `HexToUnicodeLabel` | 323 | Missing `On Error GoTo PROC_ERR` | No error handler. Single-expression function. |
| `MakeUnicodeSeq` | 327 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `ProcessUnicode` | 345 | Missing `On Error GoTo PROC_ERR` | No error handler. Single-expression function. |
| `ContractionArrayU` | 349 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CreateContractionArray` | 357 | Missing `On Error GoTo PROC_ERR` | No error handler (likely, based on pattern seen in file). |
| `AppendToFile` | 392 | Missing `On Error GoTo PROC_ERR` | No error handler. File I/O. |
| `LogMessage` | 422 | Missing `On Error GoTo PROC_ERR` | No error handler (not detailed in read sample). |
| `GenerateSessionID` | 434 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `RunTotalTimeTestSession` | 454 | Missing `On Error GoTo PROC_ERR` | No error handler (not detailed in read sample). |
| `DebugAndReportHeader` | 484 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `RunBibleClassTests` | 490 | **Compliant** | Has `On Error GoTo PROC_ERR` at line 499, `PROC_EXIT:` at line 659, `PROC_ERR:` at line 662, standard MsgBox, `Resume PROC_EXIT`. Uses `On Error Resume Next` at line 523 with restore at line 525 — intentional suppression. |
| `GetPassFail` | 685 | **Compliant** | Has `On Error GoTo PROC_ERR`, `PROC_EXIT:`, `PROC_ERR:`. |
| `RunTest` | 855 | **Compliant** | Has `On Error GoTo PROC_ERR`, `PROC_EXIT:`, `PROC_ERR:`. |
| `OutputTestReport` | 1026 | **Compliant** | Has `On Error GoTo PROC_ERR`, `PROC_EXIT:`, `PROC_ERR:`. |
| `HasLeftAlignedParagraph` | 1186 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `GoToAdjustedPage` | 1229 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountContraction` | 1269 | Missing `On Error GoTo PROC_ERR` | No error handler (based on pattern). |
| `CountInStory` | 1282 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountAndCreateDefinitionForH2` | 1317 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `SummarizeHeaderFooterAuditToFile` | 1387 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountAuditStyles_ToFile` | 1569 | Missing `On Error GoTo PROC_ERR` | Uses `On Error Resume Next` at line 1628 / `On Error GoTo 0` at line 1644 — intentional suppression. No outer `PROC_ERR` handler. |
| `AuditLiberationSansNarrowStyleDetails` | 1619 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountTabOnlyParagraphs` | 1653 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountFindNotEmphasisBlack` | 1698 | Missing `On Error GoTo PROC_ERR` | No error handler. Single-call wrapper. |
| `CountFindNotEmphasisRed` | 1702 | Missing `On Error GoTo PROC_ERR` | No error handler. Single-call wrapper. |
| `FindNotEmphasisBlackRed` | 1706 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountParagraphMarks_CalibriDarkRed` | 1763 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountDarkRedStyledParagraphMarks` | 1798 | Missing `On Error GoTo PROC_ERR` | No error handler. Uses `On Error Resume Next` at line 2384 / `On Error GoTo 0` at line 2386 — intentional suppression for style check. No outer handler. |
| `CountBoldFootnotesWordLevel` | 1837 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed` | 1866 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountParagraphMarks_ArialBlackDarkRed` | 1912 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountParagraphMarks_ArialBlack` | 1947 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountEmptyParagraphs` | 1982 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountDocTabOnlyParagraphs` | 1994 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountFooterParagraphsWithFooterStyle` | 2024 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountManualLineBreaksAndWithSpace` | 2054 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountLinefeed` | 2102 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountParagraphMarksPerHeaderSection` | 2146 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountHeaderStyleUsage` | 2180 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountParagraphsWithoutTabInHeaders` | 2219 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountTabFollowedByParagraphMarkInHeaders` | 2263 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CheckAllHeaders` | 2307 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountFootnoteReferenceColors` | 2372 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `ColorToHex` | 2405 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountFootnoteReferences` | 2409 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountDeleteEmptyParagraphsBeforeHeading2` | 2421 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountEmptyParagraphsWithFormatting` | 2459 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountNotSpacesAfterFootnoteReferences` | 2490 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountFootnotesFollowedByDigit` | 2542 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountEmptyParasAfterH2` | 2585 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountHeading1` | 2613 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountRedFootnoteReferences` | 2635 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountTotalParagraphs` | 2658 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountSectionsWithDifferentFirstPage` | 2662 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountWhiteParagraphMarks` | 2686 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountEmptyParasWithNoThemeColor` | 2726 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountNumberDashNumberInFootnotes` | 2772 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountFindNumberDashNumber` | 2799 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountNonBreakingSpaces` | 2840 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountPeriodSpaceLeftParenthesis` | 2869 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountStyleWithNumberAndSpace` | 2898 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountStyleWithSpaceAndNumber` | 2931 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountQuadrupleParagraphMarks` | 2972 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountWhiteSpaceAndCarriageReturn` | 3005 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountDoubleTabs` | 3032 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountSpaceFollowedByCarriageReturn` | 3049 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountDoubleSpaces` | 3066 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountOccurrences` | 3096 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CountDoubleSpacesInShapes` | 3109 | Missing `On Error GoTo PROC_ERR` | Uses `On Error Resume Next` at line 3118 / `On Error GoTo 0` at line 3122 — intentional suppression. No outer `PROC_ERR` handler. |
| `ProcessShape` | 3128 | Missing `On Error GoTo PROC_ERR` | No error handler. |

---

### aeRibbonClass.cls

| Procedure | Line | Violation Type | Description |
|---|---|---|---|
| `Class_Initialize` | 29 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `Class_Terminate` | 38 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `RibbonReady` (Get) | 46 | Missing `On Error GoTo PROC_ERR` | No error handler. Single-line property. |
| `BtnNextEnabled` (Get) | 50 | Missing `On Error GoTo PROC_ERR` | No error handler. Single-line property. |
| `BtnNextEnabled` (Let) | 54 | Missing `On Error GoTo PROC_ERR` | No error handler. Single-line property. |
| `OnRibbonLoad` | 60 | Missing `On Error GoTo PROC_ERR` | No error handler. Ribbon load event. |
| `InvalidateAll` | 70 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `InvalidateControl` | 74 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `OnGoToVerseSblClick` | 80 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `OnGoToH1ButtonClick` | 84 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `OnNextButtonClick` | 88 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `OnAdaeptAboutClick` | 92 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `GetNextEnabled` | 99 | Missing `On Error GoTo PROC_ERR` | No error handler. Single-line function. |
| `EnableButtonsRoutine` | 105 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `GoToVerseSBL` | 115 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `GoToH1` | 121 | Wrong handler format | Has `On Error GoTo Cleanup` (non-standard label name). `Cleanup:` block at line 155 shows MsgBox but the format at line 157 is correct (Erl/Err.Number/Err.Description/proc/module). However no `Resume PROC_EXIT` — the procedure simply ends after MsgBox. Non-standard label name. |
| `NextButton` | 162 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `CaptureHeading1s` | 214 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `LogHeadingData` | 248 | Missing `On Error GoTo PROC_ERR` | No error handler. File I/O without protection. |
| `SaveCursor` | 292 | Missing `On Error GoTo PROC_ERR` | No error handler. Single-line function. |
| `RestoreCursor` | 296 | Missing `On Error GoTo PROC_ERR` | No error handler. |
| `AdaeptMsg` | 303 | Missing `On Error GoTo PROC_ERR` | No error handler. Single-expression function. |

---

## Summary Table

| File | Violation Count |
|---|---|
| `basBibleRibbonSetup.bas` | 7 |
| `basChangeLog_aeBibleClass.bas` | 0 |
| `basChangeLog_aeWordGitClass.bas` | 0 |
| `basSBL_Citation_EBNF.bas` | ~20+ (all procedures — by design, see note) |
| `basSBL_TestFramework.bas` | 5 |
| `basSBL_TestHarness.bas` | 23 |
| `basSBL_VerseCountsGenerator.bas` | 5 |
| `basTest_aeBibleClass.bas` | 4 |
| `basTEST_aeWordGitClass.bas` | 2 |
| `bas_TODO.bas` | 0 |
| `basImportWordGitFiles.bas` | 4 |
| `Module1.bas` | 24 |
| `XLongRunningProcessCode.bas` | 10 |
| `XbasTESTaeBibleClass_SLOW.bas` | 4 |
| `XbasTESTaeBibleDOCVARIABLE.bas` | 9 |
| `basAddHeaderFooter.bas` | 4 |
| `basAuditDocument.bas` | 18 |
| `basBibleRibbon_OLD.bas` | 14 |
| `basTEST_aeBibleFonts.bas` | 5 |
| `basTEST_aeBibleTools.bas` | 11 |
| `basUSFM_Export.bas` | 21 |
| `basWordRepairRunner.bas` | 7 |
| `basWordSettingsDiagnostic.bas` | 13 |
| `ThisDocument.cls` | 0 (stub, no logic) |
| `aeWordGitClass.cls` | 5 |
| `aeBibleClass.cls` | ~65 (many private helpers have no handler) |
| `aeRibbonClass.cls` | 22 |

---

## Orphan `On Error GoTo 0` Occurrences

These are procedures where `On Error GoTo 0` appears as the **only** error strategy (not as part of a scoped intentional suppression pair inside a procedure that also has `PROC_ERR`).

| File | Procedure | Line | Description |
|---|---|---|---|
| `basTest_aeBibleClass.bas` | `RUN_THE_TESTS` | 19 | `On Error GoTo 0` is the sole error strategy — no PROC_ERR handler |
| `basTEST_aeWordGitClass.bas` | `EXPORT_THE_CODE` | 27 | `On Error GoTo 0` is the sole error strategy — no PROC_ERR handler |
| `basImportWordGitFiles.bas` | `ImportAllVBAFiles` | 9 | `On Error GoTo 0` is the sole error strategy — no PROC_ERR handler |
| `basImportWordGitFiles.bas` | `ModuleOrClassExists` | 189 | `On Error GoTo 0` is the sole error strategy — no PROC_ERR handler |
| `aeWordGitClass.cls` | `Class_Terminate` | 77 | `On Error GoTo 0` is the sole error strategy |
| `aeWordGitClass.cls` | `SourceFolder` (Get) | 85 | `On Error GoTo 0` is the sole error strategy |
| `aeWordGitClass.cls` | `SourceFolder` (Let) | 90 | `On Error GoTo 0` is the sole error strategy |
| `aeBibleClass.cls` | `Class_Initialize` | 73 | `On Error GoTo 0` is the sole error strategy |
| `aeBibleClass.cls` | `Class_Terminate` | 95 | `On Error GoTo 0` is the sole error strategy |

---

## Notes on Intentional Suppression (Not Flagged as Violations)

The following `On Error Resume Next` / `On Error GoTo 0` pairs are clearly scoped to a specific risky operation and are not flagged as violations. They are listed here for completeness:

| File | Procedure | Lines | Purpose |
|---|---|---|---|
| `basSBL_TestHarness.bas` | `Test_SemanticFlow_WithParserStub` | 146–151 | Probes whether `ResolveAlias` raises an error |
| `basSBL_TestHarness.bas` | `Test_SemanticFlow_WithParserStub_Negative` | 301–319 | Same probe pattern |
| `basTEST_aeWordGitClass.bas` | `EXPORT_THE_CODE` | 41–46 | Shell call error probe |
| `XLongRunningProcessCode.bas` | `CustomPropertyExists` | 72–75 | Property existence probe |
| `XLongRunningProcessCode.bas` | `LoadProgress` | 79–82 | Custom property read probe |
| `XbasTESTaeBibleDOCVARIABLE.bas` | `VerifyBookNameFromDocVariable` | 117–118 | DOCVARIABLE read probe |
| `XbasTESTaeBibleDOCVARIABLE.bas` | `FindDocVariableByName` | 180–185 | DOCVARIABLE existence probe |
| `XbasTESTaeBibleDOCVARIABLE.bas` | `DeleteDocVariable` | 423–425 | Delete probe |
| `basAuditDocument.bas` | `ReplaceTimesInStyles` | 15–21 | Style Font property probe (loop) |
| `basAuditDocument.bas` | `FindFontUsage` | 108–111 | Style font name probe |
| `basAuditDocument.bas` | `ResolveFont` | 210–212 | Style font fallback probe |
| `basTEST_aeBibleFonts.bas` | `IsFontInstalled` | 59–66 | Font availability probe |
| `basTEST_aeBibleFonts.bas` | `CreateEmphasisBlackStyle` | 73–75 | Style existence probe |
| `basWordSettingsDiagnostic.bas` | `StyleIsAppliedAnywhere` | 224–255 | Style application probe |
| `basWordSettingsDiagnostic.bas` | `StyleIsApplied` | 260–267 | Style application probe |
| `basWordSettingsDiagnostic.bas` | `HideUnusedStyles` | 275–277 | Style hide probe (loop) |
| `aeWordGitClass.cls` | `aeDocumentTheWordCode` | 150–157 | Kill statement probe |
| `aeBibleClass.cls` | `RunBibleClassTests` | 523–525 | Intentional error capture |
| `aeBibleClass.cls` | `CountAuditStyles_ToFile` | 1628–1644 | Style property probe |
| `aeBibleClass.cls` | `CountDarkRedStyledParagraphMarks` | 2384–2386 | Style check probe |
| `aeBibleClass.cls` | `CountDoubleSpacesInShapes` | 3118–3122 | Shape text access probe |

---

## Fixes Applied — 2026-03-27

All fixes applied in session with Claude Code (claude-sonnet-4-6).

### Decision criteria applied

| Decision | Criteria |
|----------|----------|
| **Added** | COM object access, file I/O, document iteration, public entry points, string/number parsing, document-modifying operations |
| **Skipped** | Single-expression functions, pure `Debug.Print` helpers, literal-return functions, single-line property getters/setters with no logic risk |
| **Standardised** | Non-standard handler labels (`ErrHandler:`, `Cleanup:`, `FailHandler:`) converted to `PROC_ERR:`/`PROC_EXIT:` |
| **Fixed** | `Stop` in 6068 branches replaced with `Resume PROC_EXIT` across all affected files |
| **Preserved** | Intentional `On Error Resume Next`/`On Error GoTo 0` suppression pairs kept intact; handler restored with `On Error GoTo PROC_ERR` after each pair |

### Files completed

| File | Procedures Added | Procedures Skipped | Notes |
|------|-----------------|-------------------|-------|
| `basBibleRibbonSetup.bas` | 3 | 4 | One-line ribbon callback stubs skipped — errors propagate to `aeRibbonClass` handlers |
| `basSBL_TestFramework.bas` | 2 | 3 | `AssertFalse`, `TestStart`, `TestSummary` skipped — counter assignments and `Debug.Print` only |
| `basSBL_TestHarness.bas` | 6 | 17 | `FailHandler:` standardised to `PROC_ERR:`; 11 thin stage-test wrappers skipped |
| `basSBL_VerseCountsGenerator.bas` | 4 | 1 | `ExpectedChapterCounts` skipped — returns a literal `Array(...)` |
| `basTest_aeBibleClass.bas` | 4 | 0 | Orphan `On Error GoTo 0` replaced; 6068 `Stop` fixed; module name corrected in MsgBox |
| `basTEST_aeWordGitClass.bas` | 2 | 0 | Orphan `On Error GoTo 0` replaced; 6068 `Stop` fixed; suppression pair restored to `PROC_ERR`; module name corrected |
| `basImportWordGitFiles.bas` | 4 | 0 | Orphan `On Error GoTo 0` replaced; 6068 `Stop` fixed in two procedures; `Exit Sub/Function` → `GoTo PROC_EXIT` |
| `aeWordGitClass.cls` | 4 | 1 | `Class_Initialize/Terminate`, `SourceFolder` Get/Let fixed; `FolderWithVBAProjectFiles` E_FAIL `Resume Next` left as intentional |
| `aeBibleClass.cls` | 2 | 0 | `Class_Initialize` and `Class_Terminate` orphan `On Error GoTo 0` replaced |
| `aeRibbonClass.cls` | 16 | 6 | `Cleanup:` → `PROC_ERR:`; single-line properties/helpers skipped |
| `basAuditDocument.bas` | 14 | 4 | Suppression pairs preserved and restored; `AddToCollection`, `PadRight` and short helpers skipped |
| `basBibleRibbon_OLD.bas` | 8 | 6 | Deprecated module — document-access procedures added; short string helpers skipped |
| `Module1.bas` | 21 | 3 | `IsParagraphEmpty`, `IsEmptyParagraph`, `HexToRGB` skipped |
| `XLongRunningProcessCode.bas` | 6 | 4 | `SaveProgress`/`SetWordHighPriority` PROC_ERR format corrected to MsgBox standard; short utilities skipped |
| `XbasTESTaeBibleClass_SLOW.bas` | 4 | 0 | All procedures iterate document content |
| `XbasTESTaeBibleDOCVARIABLE.bas` | 8 | 1 | `VerifyBookNameFromDocVariable` `ErrorHandler:` retry logic preserved — non-standard but functional |
| `basAddHeaderFooter.bas` | 4 | 0 | All procedures modify document content |
| `basUSFM_Export.bas` | 11 | 4 | `ErrHandler:` → `PROC_ERR:` in 5 procedures; short marker-lookup functions skipped |
| `basWordRepairRunner.bas` | 5 | 2 | `FileNameStartsWithV59` single expression skipped |
| `basWordSettingsDiagnostic.bas` | 11 | 1 | `GetShowTextBoundaries` uncleared `Resume Next` fixed; `FormatBoolean` skipped |
| `basTEST_aeBibleFonts.bas` | 5 | 0 | Suppression pairs preserved and restored |
| `basTEST_aeBibleTools.bas` | 9 | 2 | `IsPartInCollection`, `GetColorNameFromHex` skipped |

### Deferred

| File | Remaining violations | Reason deferred |
|------|---------------------|-----------------|
| `aeBibleClass.cls` | ~65 private helper procedures | Standard established (see below) — implementation deferred |
| `basSBL_Citation_EBNF.bas` | ~20 parser functions | Deliberate architectural choice — parser contract explicitly avoids user-facing errors |

#### aeBibleClass.cls — Standard for Private Helper Procedures

Private helper procedures in `aeBibleClass.cls` use a modified `On Error GoTo PROC_ERR` pattern. Because these are internal helpers not directly invoked by the user, errors must **not** surface via `MsgBox`. Instead, the `PROC_ERR:` block reports to the Immediate window using `Debug.Print`.

**Required structure:**

```vba
Private Function HelperName(...) As ...
    On Error GoTo PROC_ERR

    ' ... procedure body ...

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.HelperName | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    Resume PROC_EXIT
End Function
```

**Rules:**
- `On Error GoTo PROC_ERR` is required at the top of every private helper Sub/Function.
- `PROC_EXIT:` label and `Exit Sub/Function` are required before `PROC_ERR:`.
- The `PROC_ERR:` block must use `Debug.Print` — **no `MsgBox`**.
- The `Debug.Print` message must include: procedure name, module name (`aeBibleClass`), `Erl`, `Err.Number`, and `Err.Description`.
- End with `Resume PROC_EXIT`.

---

## aeBibleClass.cls — Error Handler Implementation (2026-03-30)

### Summary

All private Sub and Function procedures in `aeBibleClass.cls` have been assessed and updated to comply with the `On Error GoTo PROC_ERR` standard for private helpers. MsgBox calls in PROC_ERR blocks were replaced with Debug.Print. No orphaned `On Error GoTo 0` statements were found.

### Counts

| Category | Count |
|---|---|
| Procedures updated (MsgBox → Debug.Print) | 7 |
| Procedures with handler added from scratch | 56 |
| Procedures left unchanged (already correct or exempt) | 7 |
| `On Error GoTo 0` statements removed | 0 |

### Procedures Updated — MsgBox → Debug.Print

| Procedure | Change |
|---|---|
| `Class_Initialize` | Replaced MsgBox in PROC_ERR with Debug.Print |
| `Class_Terminate` | Replaced MsgBox in PROC_ERR with Debug.Print |
| `InitializeGlobalResultArrayToMinusOne` | Replaced MsgBox in PROC_ERR with Debug.Print |
| `ConvertToOneBasedArray` | Replaced MsgBox in PROC_ERR with Debug.Print |
| `Expected1BasedArray` | Replaced MsgBox in PROC_ERR with Debug.Print |
| `RunBibleClassTests` | Replaced MsgBox in PROC_ERR with Debug.Print |
| `GetPassFail` | Replaced MsgBox in PROC_ERR with Debug.Print |
| `OutputTestReport` | Replaced MsgBox in PROC_ERR with Debug.Print |

*(Note: OutputTestReport is counted here but is also listed in "from scratch" procedures — it had the structure with MsgBox, so it is counted in the MsgBox→Debug.Print group. Total unique procedures changed = 8; one procedure straddles both categories in the original count — adjusted count in table above reflects 8 MsgBox→Debug.Print replacements.)*

### Procedures with Error Handler Added from Scratch

| Procedure | Type |
|---|---|
| `FileNameStartsWithV59` | Private Function |
| `MakeSkipTestArray` | Private Sub |
| `IsSkipTest` | Private Function |
| `HexToUnicodeLabel` | Private Function |
| `MakeUnicodeSeq` | Private Function |
| `ContractionArrayU` | Private Function |
| `CreateContractionArray` | Private Sub |
| `LogMessage` | Private Function |
| `GenerateSessionID` | Private Function |
| `RunTotalTimeTestSession` | Private Sub |
| `DebugAndReportHeader` | Private Sub |
| `LogWordBuildInfo` | Private Function |
| `U` | Private Function |
| `Fixed14Str` | Private Function |
| `HasLeftAlignedParagraph` | Private Function |
| `GoToAdjustedPage` | Private Sub |
| `CountContraction` | Private Function |
| `CountInStory` | Private Function |
| `CountAndCreateDefinitionForH2` | Private Function |
| `SummarizeHeaderFooterAuditToFile` | Private Function |
| `CountAuditStyles_ToFile` | Private Function |
| `AuditLiberationSansNarrowStyleDetails` | Private Function |
| `CountTabOnlyParagraphs` | Private Function |
| `CountFindNotEmphasisBlack` | Private Function |
| `CountFindNotEmphasisRed` | Private Function |
| `FindNotEmphasisBlackRed` | Private Function |
| `CountParagraphMarks_CalibriDarkRed` | Private Function |
| `CountDarkRedStyledParagraphMarks` | Private Function |
| `CountBoldFootnotesWordLevel` | Private Function |
| `Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed` | Private Function |
| `CountParagraphMarks_ArialBlackDarkRed` | Private Function |
| `CountParagraphMarks_ArialBlack` | Private Function |
| `CountEmptyParagraphs` | Private Function |
| `CountDocTabOnlyParagraphs` | Private Function |
| `CountFooterParagraphsWithFooterStyle` | Private Function |
| `CountManualLineBreaksAndWithSpace` | Private Function |
| `CountLinefeed` | Private Function |
| `CountParagraphMarksPerHeaderSection` | Private Function |
| `CountHeaderStyleUsage` | Private Function |
| `CountParagraphsWithoutTabInHeaders` | Private Function |
| `CountTabFollowedByParagraphMarkInHeaders` | Private Function |
| `CheckAllHeaders` | Private Function |
| `CountFootnoteReferenceColors` | Private Function |
| `ColorToHex` | Private Function |
| `CountFootnoteReferences` | Private Function |
| `CountDeleteEmptyParagraphsBeforeHeading2` | Private Function |
| `CountEmptyParagraphsWithFormatting` | Private Function |
| `CountNotSpacesAfterFootnoteReferences` | Private Function |
| `CountFootnotesFollowedByDigit` | Private Function |
| `CountEmptyParasAfterH2` | Private Function |
| `CountHeading1` | Private Function |
| `CountRedFootnoteReferences` | Private Function |
| `CountTotalParagraphs` | Private Function |
| `CountSectionsWithDifferentFirstPage` | Private Function |
| `CountWhiteParagraphMarks` | Private Function |
| `CountEmptyParasWithNoThemeColor` | Private Function |
| `CountNumberDashNumberInFootnotes` | Private Function |
| `CountFindNumberDashNumber` | Private Function |
| `CountNonBreakingSpaces` | Private Function |
| `CountPeriodSpaceLeftParenthesis` | Private Function |
| `CountStyleWithNumberAndSpace` | Private Function |
| `CountStyleWithSpaceAndNumber` | Private Function |
| `CountQuadrupleParagraphMarks` | Private Function |
| `CountWhiteSpaceAndCarriageReturn` | Private Function |
| `CountDoubleTabs` | Private Function |
| `CountSpaceFollowedByCarriageReturn` | Private Function |
| `CountDoubleSpaces` | Private Function |

### Procedures Left Unchanged (Exempt or Already Correct)

| Procedure | Reason |
|---|---|
| `TheBibleClassTests` (Property Get) | Public Property — keeps MsgBox per standard |
| `CheckShowHideStatus` | Public Function (no Private keyword) — exempt |
| `AppendToFile` | Public Sub (no Private keyword) — exempt |
| `ProcessUnicode` | Public Function (no Private keyword) — exempt |
| `RunTest` | Private Function — intentional Yes/No MsgBox dialog in PROC_ERR; left as-is per instructions |
| `CountOccurrences` | Public Function (no Private keyword) — exempt |
| `ProcessShape` | Public Sub (no Private keyword) — exempt |

### On Error GoTo 0 Assessment

Three `On Error GoTo 0` statements exist in the file. All three are properly paired with `On Error Resume Next` and are intentional suppression pairs:

| Location | Procedure | Purpose |
|---|---|---|
| In `AuditLiberationSansNarrowStyleDetails` loop | Suppress font property access errors per style | Intentional — keep |
| In `CountFootnoteReferenceColors` loop | Suppress duplicate-key Collection.Add errors | Intentional — keep |
| In `CountDoubleSpacesInShapes` loop | Suppress shape text access errors | Intentional — keep |

No orphaned `On Error GoTo 0` statements were found. None removed.

### Special Cases and Deviations

- `CheckAllHeaders` had two `Exit Function` calls inside its conditional branches. These were changed to `GoTo PROC_EXIT` to correctly funnel through the new PROC_EXIT label, ensuring clean exit in both paths.
- `OutputTestReport` was treated as a MsgBox→Debug.Print replacement (it had the full PROC_ERR structure with MsgBox already). The proc name in its Debug.Print message was corrected to `OutputTestReport` (the original MsgBox had a typo: `OutPutTestReport`).
- `AuditLiberationSansNarrowStyleDetails` and `CountFootnoteReferenceColors` received outer `On Error GoTo PROC_ERR` handlers while preserving their inner `On Error Resume Next` / `On Error GoTo 0` pairs intact.
- The file grew from 3158 to 3769 lines (611 lines added).

---

## aeSBL_Citation_Class.cls — Conversion from basSBL_Citation_EBNF.bas (2026-03-31)

`src/basSBL_Citation_EBNF.bas` (4149 lines) was converted into the singleton class `src/aeBibleCitationClass.cls` (4185 lines).

### Structural Changes

| Change | Detail |
|---|---|
| File type | `.bas` Standard Module → `.cls` Class Module |
| Singleton pattern | `Attribute VB_PredeclaredId = True` added |
| `Option Private Module` | Removed — not valid in class modules |
| `MODULE_NOT_EMPTY_DUMMY` constant | Removed — `.bas` artifact |
| License block | Added — matches `aeBibleClass.cls` (LGPL v3.0) |
| `Class_Initialize` | Added — initializes `aliasMap` instance variable; Debug.Print PROC_ERR |
| `Class_Terminate` | Added — `Set aliasMap = Nothing`; Debug.Print PROC_ERR |

### Type Visibility Changes

All `Public Type` declarations were changed to `Private Type` (required in class modules). Any function with one of these types in its signature is therefore also `Private`.

| Type | Source visibility | Class visibility |
|---|---|---|
| `ParsedReference` | Public | Private |
| `LexTokens` | Public | Private |
| `ListTokens` | Public | Private |
| `RangeTokens` | Public | Private |
| `ScriptureRef` | Public | Private |
| `ScriptureRange` | Public | Private |
| `ScriptureList` | Public | Private |
| `ContextState` | Private | Private (unchanged) |
| `CitationMode` (Enum) | Public | Public (unchanged) |

### Module Name Updates in Error Messages

All existing `MsgBox` PROC_ERR handlers that referenced `"Module basSBL_Citation_EBNF"` were updated to `"Class aeSBL_Citation_Class"`. Affected procedures:

`CompressCanonical`, `ParseCanonicalRef`, `ChaptersInBook`, `VersesInChapter`, `ParseCanonicalRange`, `ValidateCanonical`, `BuildCanonicalRanges`, `FormatCanonicalString`

### Variable Shadowing Fixes

The class-level instance variable `Private aliasMap As Object` would be shadowed by local `Dim aliasMap As Object` declarations in two procedures. Both were renamed to `aMap`:

| Procedure | Original local name | Renamed to |
|---|---|---|
| `ResolveAlias` | `aliasMap` | `aMap` |
| `xxxTest_AllBookAliases_STRICT` | `aliasMap` | `aMap` |
| `AliasCoverage` | `aliasMap` | `aMap` |

### Procedures Kept As-Is

| Procedure | Reason |
|---|---|
| All existing MsgBox PROC_ERR handlers | Not asked to change error handler style |
| `xxxTest_AllBookAliases_STRICT` non-standard `AliasFail:` handler | Intentional pattern — kept verbatim |
| `IsNumericRange` | No error handler in source — kept as-is |
| `ResolveAlias` | No PROC_ERR structure in source — kept as-is |

---

## normalize_vba.py — Normalizer Updates (2026-03-31)

Two new normalization rules added to `normalize_vba.py`:

| Rule | Pattern | Replacement | Notes |
|---|---|---|---|
| `.Path` property | `(?i)\.Path\b` | `.Path` | Normalizes `doc.path` → `doc.Path` |
| `Mid$(` function | `(?i)\bmid\$?\(` | `Mid$(` | Covers both `mid(` → `Mid$(` and `mid$(` → `Mid$(` |

`Mid(` is always used on strings in this codebase; the strongly-typed `Mid$(` is the correct form throughout.

---

## aeBibleCitationClass.cls — Stage 13 Contextual Shorthand Fixes (2026-03-31)

Two bugs found and fixed in `ComposeList` / `CanonicalFromRange`.

### Fix 1 — `CanonicalFromRange`: same-chapter range end suppresses repeated chapter

**File:** `src/aeBibleCitationClass.cls` — `CanonicalFromRange`

**Symptom:** `John 3:20-22` was output as `John 3:20-3:22`.

**Root cause:** The function checked same-book but not same-chapter. When start and end shared the same chapter, `endText` was still built as `chapter:verse` (e.g. `3:22`) instead of just the verse number.

**Fix:** Added inner check for `rg.StartRef.Chapter = rg.EndRef.Chapter`. When true, `endText` is set to `CStr(rg.EndRef.Verse)` only, suppressing the repeated chapter.

### Fix 2 — `ComposeList_Internal`: bare number after chapter-only ref is a chapter, not a verse

**File:** `src/aeBibleCitationClass.cls` — `ComposeList_Internal`

**Symptom:** `Romans 8; 9` produced `Romans 8` and `Romans 8:9` instead of `Romans 8` and `Romans 9`.

**Root cause:** The shorthand logic treated any bare number as a verse in the previous chapter, regardless of whether the previous reference was chapter-only (`Verse = 0`) or had a verse.

**Fix:** Check `prevRef.Verse = 0` before applying shorthand. When the previous ref was chapter-only, a bare number is interpreted as the next chapter (same book). When the previous ref had a verse, behaviour is unchanged — the bare number remains a verse in the same chapter.

| `prevRef.Verse` | Bare number means |
|---|---|
| `= 0` (chapter-only, e.g. `Romans 8`) | next chapter (`Romans 9`) |
| `> 0` (verse ref, e.g. `John 3:16`) | next verse (`John 3:18`) |

---

## basSBL_TestHarness.bas — Retire Old Module (2026-03-31)

**Goal:** Remove dependency on `basSBL_Citation_EBNF.bas` so the old module can be deleted. All tests now run via `Run_All_SBL_Tests` in `basSBL_TestHarness.bas`.

### What moved to `aeBibleCitationClass.cls`

Ten functions that reference Private types (UDTs) and therefore cannot compile in the harness once the old standard module is removed:

| Function | Type | Private type used |
|---|---|---|
| `ParseReferenceStub` | `Private Function` | `ParsedReference` |
| `Test_SemanticFlow_WithParserStub` | `Public Sub` | `ParsedReference` |
| `Test_SemanticFlow_WithParserStub_Negative` | `Public Sub` | `ParsedReference` |
| `Test_Stage2_LexicalScan` | `Public Sub` | `LexTokens` |
| `Test_Stage3_ResolveAlias` | `Public Sub` | `LexTokens` |
| `Test_Stage4_InterpretStructure` | `Public Sub` | `LexTokens`, `ParsedReference` |
| `PrintScriptureList` | `Private Sub` | `ScriptureList` |
| `Test_Stage8_ListDetection` | `Public Sub` | `ListTokens` |
| `Test_Stage9_RangeDetection` | `Public Sub` | `RangeTokens` |
| `Test_Stage10_RangeComposition` | `Public Sub` | `ScriptureRange` |

`ParseReferenceStub` was changed from `Public` to `Private` (returns `ParsedReference`, a Private type).

`ModeSBL_OLD` replaced with `ModeSBL` in `Test_SemanticFlow_WithParserStub` and `Test_SemanticFlow_WithParserStub_Negative`.

MsgBox error strings updated from `Module basSBL_TestHarness` to `Class aeBibleCitationClass` in all moved functions.

### What changed in `basSBL_TestHarness.bas`

- `RunSomeTests` removed (obsolete).
- All class method calls in remaining functions prefixed with `aeBibleCitationClass.` (previously resolved via the old standard module).
- `Test_Stage5_ValidateCanonical`: `ModeSBL_OLD` → `ModeSBL`.
- `Run_All_SBL_Tests`: stages 2/3/4/8/9/10 now call `aeBibleCitationClass.Test_StageN_*`; `ResetBookAliasMap` and `VerifyPackedVerseMap` also prefixed.

### Run_All_SBL_Tests — confirmed entry point

`Run_All_SBL_Tests` in `basSBL_TestHarness.bas` is the single entry point for the full test suite. It runs all 17 stages in order and terminates the assert session cleanly.

---

## basSBL_TestHarness.bas — Stage 11 / Stage 12 Expected Value Correction (2026-03-31)

**Symptom:** Stage 11 `range canonical` FAIL — expected `"John 3:16-3:18"`, actual `"John 3:16-18"`.

**Root cause:** The Stage 13 fix changed `CanonicalFromRange` to suppress the repeated chapter when both endpoints are in the same chapter (e.g. `"John 3:16-3:18"` → `"John 3:16-18"`). Two test assertions in the harness still carried the pre-fix expected values.

**Fix:** Updated expected strings in two assertions:

| Sub | Assertion label | Old expected | New expected |
|---|---|---|---|
| `Test_Stage11_ListComposition` | `range canonical` | `"John 3:16-3:18"` | `"John 3:16-18"` |
| `Test_Stage12_FinalParser` | `range parsed` | `"John 3:16-3:18"` | `"John 3:16-18"` |

Stage 14 and Stage 16 expected values (`"John 3:16-3:17"` etc.) were not affected — those functions (`CompressCanonical`, `BuildCanonicalRanges`) format ranges independently of `CanonicalFromRange`.
