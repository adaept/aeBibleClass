# Code Review — Access Modifier Audit
**Date:** 2026-04-01
**Scope:** All `.bas` and `.cls` files in `src/`
**Rule:** Every `Sub`, `Function`, and `Property` must be explicitly marked `Public` or `Private`. Omitting the modifier defaults to `Public` in VBA, which unintentionally exposes internal helpers.

---

## Summary

| File | Non-compliant procedures | Priority |
|---|---|---|
| `aeBibleClass.cls` | 5 | High |
| `basTest_aeBibleClass.bas` | 2 | High |
| `basBibleRibbon_OLD.bas` | 1 | Low (legacy) |
| `basTEST_aeBibleFonts.bas` | 10 | Medium |
| `basTEST_aeBibleTools.bas` | 63 | Medium |
| `basWordSettingsDiagnostic.bas` | 13 | Medium |
| `Module1.bas` | 38 | Low (scratch) |
| `XLongRunningProcessCode.bas` | 9 | Low (scratch) |
| `XbasTESTaeBibleClass_SLOW.bas` | 4 | Low (scratch) |
| `XbasTESTaeBibleDOCVARIABLE.bas` | 9 | Low (scratch) |

Files prefixed `X` are treated as archived/scratch — low priority. Files prefixed `bas` or `ae` in active use are high/medium priority.

---

## High Priority — Active Production Files

### `aeBibleClass.cls`

| Line | Declaration | Recommended |
|---|---|---|
| 152 | `Function CheckShowHideStatus() As Boolean` | `Private` — internal state check |
| 398 | `Function ProcessUnicode(s As String) As String` | `Private` — internal string transform |
| 463 | `Sub AppendToFile(filePath As String, text As String)` | `Private` — internal file helper |
| 3704 | `Function CountOccurrences(ByVal text As String, ...) As Long` | `Private` — internal string utility |
| 3736 | `Sub ProcessShape(ByVal shp As Shape, ByRef doubleSpaceCount As Long)` | `Private` — internal shape walker |

### `basTest_aeBibleClass.bas`

| Line | Declaration | Recommended |
|---|---|---|
| 90 | `Sub GitAutoTagRelease()` | `Public` — standalone macro entry point |
| 149 | `Function GitTagExists(sRepoPath As String, sTag As String) As Boolean` | `Private` — called only by `GitAutoTagRelease` |

---

## Medium Priority — Test and Tool Modules

### `basTEST_aeBibleFonts.bas`
All 10 procedures missing access modifiers. Recommended defaults:

| Line | Declaration | Recommended |
|---|---|---|
| 8 | `Sub CheckOpenFontsWithDownloads()` | `Public` |
| 63 | `Function IsFontInstalled(fontName As String) As Boolean` | `Private` — helper |
| 84 | `Sub CreateEmphasisBlackStyle()` | `Public` |
| 120 | `Sub AuditStyleUsage_Footnote()` | `Public` |
| 154 | `Sub RedefineFootnoteStyle_NotoSans()` | `Public` |
| 177 | `Sub AuditStyleUsage_FootnoteNormal()` | `Public` |
| 204 | `Sub RedefineFootnoteNormalStyle_NotoSans()` | `Public` |
| 227 | `Sub AuditStyleUsage_PictureCaption()` | `Public` |
| 265 | `Sub RedefinePictureCaptionStyle_NotoSans()` | `Public` |
| 297 | `Sub Identify_ArialUnicodeMS_Paragraphs()` | `Public` |

### `basWordSettingsDiagnostic.bas`
All 13 procedures missing access modifiers. Recommended defaults:

| Line | Declaration | Recommended |
|---|---|---|
| 10 | `Sub RunWordSettingsAudit(...)` | `Public` — entry point |
| 41 | `Function GetCurrentWordSettings() As Object` | `Private` — called by `RunWordSettingsAudit` |
| 81 | `Function GetShowTextBoundaries() As Variant` | `Private` — internal helper |
| 106 | `Function LoadTargetBaseline() As Object` | `Private` — called by `RunWordSettingsAudit` |
| 130 | `Function CompareSettings(...) As Object` | `Private` — called by `RunWordSettingsAudit` |
| 158 | `Function FormatDiagnostics(...) As String` | `Private` — called by `RunWordSettingsAudit` |
| 204 | `Function FormatBoolean(value As Variant) As String` | `Private` — called by `FormatDiagnostics` |
| 213 | `Sub SaveReportToFile(reportText As String, fileName As String)` | `Private` — called by `RunWordSettingsAudit` |
| 231 | `Sub ShowAllStyles()` | `Public` |
| 249 | `Sub ShowMyStyles()` | `Public` |
| 285 | `Function StyleIsAppliedAnywhere(sName As String) As Boolean` | `Private` — helper |
| 329 | `Function StyleIsApplied(sName As String) As Boolean` | `Private` — helper |
| 349 | `Sub HideUnusedStyles()` | `Public` |

### `basTEST_aeBibleTools.bas`
63 procedures missing access modifiers. Representative sample — confirm before applying:

| Line | Declaration | Recommended |
|---|---|---|
| 115 | `Function IsPartInCollection(...) As Boolean` | `Private` — helper |
| 170 | `Function GetColorNameFromHex(hexColor As String) As String` | `Private` — helper |
| 360 | `Function HeaderTypeName(hdrType As Variant) As String` | `Private` — helper |
| 739 | `Function IsCorrectFootnoteFormat(...) As Boolean` | `Private` — helper |
| 896 | `Function RGBToString(rgbVal As Long) As String` | `Private` — helper |
| 2076 | `Function ShouldFlag(...) As ...` | `Private` — helper |
| 2090 | `Function FlagLabel(codeLine As String) As String` | `Private` — helper |
| 2112 | `Function IsPrimitiveType(...) As Boolean` | `Private` — helper |
| 2129 | `Function IsWordNative(...) As Boolean` | `Private` — helper |
| 2146 | `Function IsEnumType(...) As Boolean` | `Private` — helper |
| All standalone `Sub ...()` | Entry-point macros | `Public` |

---

## Low Priority — Legacy / Scratch Files

### `basBibleRibbon_OLD.bas`

| Line | Declaration | Recommended |
|---|---|---|
| 232 | `Sub GoToSection()` | `Public` |

### `Module1.bas`
38 procedures, all missing modifiers. File appears to be a scratch/diagnostic module. Recommend either:
- Marking all entry-point `Sub`s `Public` and all helper `Function`s `Private`, or
- Retiring the file if it has been superseded.

### `XLongRunningProcessCode.bas`, `XbasTESTaeBibleClass_SLOW.bas`, `XbasTESTaeBibleDOCVARIABLE.bas`
Archived files (`X` prefix). Address only if reactivated.

---

## Compliant Files (no action needed)

The following active files already use explicit access modifiers on all procedures:
- `aeBibleCitationClass.cls`
- `basTEST_aeBibleClass.bas` (except 2 noted above)
- `basSBL_VerseCountsGenerator.bas`
- All `aeAssert*` and `aeWord*` class files
