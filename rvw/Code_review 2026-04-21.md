# Code Review - 2026-04-21

## Carry-Forward from 2026-04-20

Continues from `rvw/Code_review 2026-04-20.md`.

---

## § 1 — Status of Previous Session (2026-04-20/21) Carry-Forward

### Completed items (closed this session)

| Item | Detail | Status |
|------|--------|--------|
| Step 4 — first-hit hint infrastructure | `m_lastHint`, `m_HintArray(1 To MaxTests)` added; initialized in `InitializeGlobalResultArrayToMinusOne`; captured in `GetPassFail`; printed in `RunTest` | **CLOSED — 2026-04-21** |
| Step 4 — Python text-mode bug | `\r\n` in replacement strings produced `\r\r\n` corruption; fixed to `\n` throughout `py/step4_hints.py` | **CLOSED — 2026-04-21** |
| Step 5 — first-hit capture in Count functions | `m_lastHint = "..."` added in all FAILing Count functions; `m_lastHint = ""` reset in `GetPassFail` before each call | **CLOSED — 2026-04-21** |
| Hint write to TestReport.txt | `BufAppend` call added in `OutputTestReport` for FAIL + non-empty hint | **CLOSED — 2026-04-21** |
| `basAddHeaderFooter` → `basFixDocxRoutines` rename | `git mv`; `Attribute VB_Name` updated; all 4 PROC_ERR strings updated; `README.md` updated | **CLOSED — 2026-04-21** |
| `AddBookNameHeaders` — Psalms bug | Root cause: `Paragraphs(1)` check fails when blank paragraph precedes H1/H2. Fix: full paragraph scan (`For Each oClassPara In oSection.Range.Paragraphs`). `sBookName` captured in H1 branch directly — backward search eliminated | **CLOSED — 2026-04-21** |
| `CountHeaderStyleUsage` rewrite | Was searching `doc.Content` (excludes header stories) for style "Header" — always returned 0. Rewritten to iterate `sec.Headers(wdHeaderFooterPrimary)` across all sections; counts paragraphs where `style.NameLocal <> "TheHeaders"` | **CLOSED — 2026-04-21** |
| `Stop` removed from `CountManualLineBreaksAndWithSpace` | `ElseIf prevChar <> " " Then Stop` — always-true branch, broke full suite; removed during Step 5 | **CLOSED — 2026-04-21** |
| Test 49 expected updated | 16 → 15 after `CustomParaAfterH1-2nd` (4) consolidated into `CustomParaAfterH1` (66 books, one style) | **CLOSED — 2026-04-21** |
| Tests 50 and 51 unblocked | Test 50 (`SummarizeHeaderFooterAuditToFile`) was returning -1 (silent crash) — now 147; Test 51 (`CountAndCreateDefinitionForH2`) was SKIPped — now 1189, matching expected | **CLOSED — 2026-04-21** |
| Test 30 — 7 Header violations resolved | Root cause: Section 1 `LinkToPrevious=True` inherits `Header` style from Normal template; Sections 2–7 chain from it. Fixed manually: Section 1 `LinkToPrevious` broken, `TheHeaders` + `vbTab` applied. Sections 2–7 inherit corrected style. Result: 0 violations | **CLOSED — 2026-04-21** |
| Style consolidation — Group C | `CustomParaAfterH1-2nd` (4) merged into `CustomParaAfterH1` (now 66). 10 pt vertical shift accepted | **CLOSED — 2026-04-21** |
| Style consolidation — Group D | `Footer` and `TheFooters` deduplicated; footer story now linked-to-previous; consistent across document | **CLOSED — 2026-04-21** |

### Open items (carry-forward)

| Item | Detail | Status |
|------|--------|--------|
| Bug #597 | New Search should focus `cmbBook` — Option A/B/C documented; awaiting decision | **OPEN** |
| Bug 16 | Keytip badges end-to-end test — re-test after `GetGoKeytip` injection | **PENDING** |
| Bug 22 / 23a | First-nav layout delay (~6–17s) | **KNOWN LIMITATION** |
| Step 7 | OLD_CODE cleanup — dead stubs in `aeRibbonClass.cls` | **PENDING** |
| WarmLayoutCache rewrite | Replace `Range.Select` with `ScrollIntoView`; re-enable deferred warm | **FUTURE** |
| Search tracking reset | Test `Selection.SetRange` from `OnTime` context | **FUTURE** |
| Import modules | `aeRibbonClass.cls`, `basBibleRibbonSetup.bas`, `basRibbonDeferred.bas`, `basUIStrings.bas` all modified — must be imported into VBA project | **PENDING** |
| Commit pending changes | `src/aeBibleClass.cls`, `src/basFixDocxRoutines.bas`, `rvw/Code_review 2026-04-20.md` — uncommitted | **PENDING** |
| Group B style fixes | `Plain Text:26`, `List Paragraph:82`, `Paragraph Continuation:158` — investigate contexts, write fix routines in `basFixDocxRoutines`, add tests | **OPEN** |

---

## § 2 — Test Suite Baseline — 2026-04-21

Test 30 now passes. Test 49 = 15 (expected 15). Tests 50 and 51 now running correctly.

| Test | Function | Result | Expected | Status | Notes |
|------|----------|--------|----------|--------|-------|
| 30 | `CountHeaderStyleUsage` | 0 | 0 | **PASS** | Rewritten; Section 1 front-matter fixed manually |
| 49 | `CountAuditStyles_ToFile` | 15 | 15 | **PASS** | Expected updated from 16; CustomParaAfterH1-2nd consolidated |
| 50 | `SummarizeHeaderFooterAuditToFile` | 147 | 147 | **PASS** | Was -1 (silent crash); now resolved |
| 51 | `CountAndCreateDefinitionForH2` | 1189 | 1189 | **PASS** | Was SKIPped; now running |

**Remaining style violations (target: Test 49 = 12):**

| Style | Count | Action |
|-------|-------|--------|
| `Plain Text` | 26 | Investigate contexts → fix routine in `basFixDocxRoutines` → test |
| `List Paragraph` | 82 | Investigate contexts → fix routine in `basFixDocxRoutines` → test |
| `Paragraph Continuation` | 158 | Investigate contexts → fix routine in `basFixDocxRoutines` → test |
| `Title` | 1 | Tolerated (artifact — Section 1 front-matter) |

Test 49 expected will move from 15 → 12 once Group B styles are resolved.

---
