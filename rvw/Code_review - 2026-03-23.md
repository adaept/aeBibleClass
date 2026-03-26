# Code Review — `/src` VBA Modules

**Date:** 2026-03-23
**Reviewer:** Claude Code (claude-sonnet-4-6)
**Scope:** All `.bas` and `.cls` files in `\src\`

---

## HIGH Severity

| # | File | Line | Issue | Fix |
|---|------|------|-------|-----|
| 1 | `aeRibbonClass.cls` | 258 | `isEmpty(...)` called — not a VBA built-in; will cause runtime error | Change to `IsEmpty(...)` |
| 2 | `basSBL_Citation_EBNF.bas` | 2487 | Same `isEmpty(...)` bug | Change to `IsEmpty(...)` |
| 3 | `basUSFM_Export.bas` | 539–541 | `ADODB.Stream` object (`stm`) not set to `Nothing` in error handler — leaks on failure | Add `Set stm = Nothing` in `ErrHandler` |
| 4 | `XLongRunningProcessCode.bas` | 1 | Missing `Option Explicit` — entire file has no compile-time type checking | Add `Option Explicit` at top |

---

## MEDIUM Severity

| # | File | Line | Issue | Fix |
|---|------|------|-------|-----|
| 5 | `XLongRunningProcessCode.bas` | 93 | `IsMissing(pageNumber)` on a non-`Optional` parameter — always evaluates False, dead check | Declare `Optional ByVal pageNumber As Integer` |
| 6 | `XLongRunningProcessCode.bas` | 59–72 | WMI objects (`objWMIService`, `colProcesses`) not cleaned up on error path | Add `On Error GoTo PROC_ERR` + cleanup in handler |
| 7 | `basSBL_Citation_EBNF.bas` | 1484 | `Option Explicit` appears halfway through a 3140-line file — first 1483 lines unprotected | Move to top, immediately after `Attribute` declarations |
| 8 | `aeWordGitClass.cls` | 296 | `Resume Next` in `PROC_ERR` handler skips the faulting line — risky for non-trivial errors | Change to `Resume PROC_EXIT` or add logging before resuming |
| 9 | `aeWordGitClass.cls` | 238–243 | `WshShell`, `fso` declared as generic `Object` — loses Early Binding benefits | Declare as `WScript.Shell` / `Scripting.FileSystemObject` |
| 10 | `aeRibbonClass.cls` | 254 | Hardcoded path `C:\adaept\aeBibleClass\rpt\HeadingLog.txt` — breaks if project moves | Derive from `ActiveDocument.Path` |
| 11 | `basAuditDocument.bas` | 15–20, 108, 210 | `On Error Resume Next` blocks suppress errors silently with no logging | Add `If Err.Number <> 0 Then` check before `On Error GoTo 0` |
| 12 | `aeBibleClass.cls` | 1586–1607 | `fso` / `outFile` objects not explicitly set to `Nothing` on exit | Add cleanup before `End Function` |
| 13 | `aeBibleClass.cls` | 1400–1534 | Five `Dictionary` objects never set to `Nothing` | Add `Set dictX = Nothing` before `End Function` |

---

## LOW Severity

| # | File | Line | Issue | Fix |
|---|------|------|-------|-----|
| 14 | `aeBibleClass.cls` | 541 | `space(n)` is deprecated VB6 syntax | Replace with `String(n, " ")` |
| 15 | `Module1.bas` | 96 | `GoTo EmptyPara` for control flow — anti-pattern | Refactor to `If/End If` |
| 16 | `Module1.bas` | 125 | `InputBox` result (String) compared to Long without `CLng()` — implicit coercion | Use `CLng(InputBox(...))` with error handling |

---

## Summary

| Severity | Count | Files Affected |
|----------|-------|----------------|
| HIGH | 4 | `aeRibbonClass.cls`, `basSBL_Citation_EBNF.bas`, `basUSFM_Export.bas`, `XLongRunningProcessCode.bas` |
| MEDIUM | 9 | `XLongRunningProcessCode.bas`, `basSBL_Citation_EBNF.bas`, `aeWordGitClass.cls`, `aeRibbonClass.cls`, `basAuditDocument.bas`, `aeBibleClass.cls` |
| LOW | 3 | `aeBibleClass.cls`, `Module1.bas` |

---

## Priority Fix Order

1. **Items 1 & 2** — `isEmpty` → `IsEmpty` runtime crashes (two files)
2. **Item 3** — `ADODB.Stream` leak in error path (`basUSFM_Export.bas`)
3. **Item 4** — Add `Option Explicit` to `XLongRunningProcessCode.bas`
4. **Item 5** — Fix dead `IsMissing()` check on non-Optional parameter
5. **Item 7** — Move `Option Explicit` to top of `basSBL_Citation_EBNF.bas`
6. **Items 8–11** — Error handling and object lifetime issues
7. **Items 12–13** — Dictionary/stream object cleanup in `aeBibleClass.cls`
8. **Items 14–16** — Low-risk style and pattern improvements

---

## Positive Findings

- `Option Explicit` / `Option Compare Text` present in most modules
- `PROC_ERR` / `PROC_EXIT` error-handler pattern consistently applied
- Early Binding used for `Word.Document`, `Word.Range`, `Word.Paragraph` throughout
- Code is well-commented with clear intent

---

## Resolution Summary — 2026-03-26

All 16 items reviewed one at a time. 2 fixed, 12 skipped as already fixed or review incorrect, 2 skipped by design decision.

### Fixed (2)

| # | Item | File |
|---|------|------|
| 5 | Replaced dead `IsMissing(pageNumber)` with `pageNumber = 0` — typed `Optional` parameters can never be missing; guard now actually fires | `XLongRunningProcessCode.bas` |
| 9 | Removed dead `WshShell` and `SpecialPath` variables left over from the Critical 3 `IsNull` branch removal; fixed `fullPath` declaration from implicit `Variant` to explicit `String` | `aeWordGitClass.cls` |

### Skipped — Already Fixed in Prior Sessions (10)

Items 1, 2, 3, 4, 6, 7, 10, 12, 13, 16 — all verified present and correct in current source.

### Skipped — Review Incorrect (4)

| # | Reason |
|---|--------|
| 8 | `Resume Next` for `E_FAIL` is intentional — known soft COM error; `Debug.Print` logging was already added in prior session |
| 11 | `On Error Resume Next` blocks in `basAuditDocument.bas` are the standard VBA style/font access pattern — silently swallowing COM errors on incompatible style types is correct, not a logging gap |
| 14 | `Space()` is not deprecated in VBA — fully supported in Word 365; `String(n, " ")` is an alternative with no functional benefit |
| 15 | `GoTo EmptyPara` is the standard VBA `continue` idiom; refactor to `If/End If` is cosmetic only |
