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

## § 4 — Style taxonomy — approved 2026-04-21

### Approved taxonomy

| Word Style | USFM marker | Semantic role | Replaces |
|-----------|-------------|---------------|----------|
| `Heading 1` | `\mt1` | Book title | *(reserved — keep)* |
| `Heading 2` | `\c` | Chapter heading | *(reserved — keep)* |
| `BodyText` | `\p` | Standard body paragraph | `Normal`, `Plain Text` |
| `BodyTextIndent` | `\pi` | Indented body paragraph (quoted / subordinate) | direct formatting |
| `BodyTextContinuation` | `\m` | Continuation paragraph (no indent) | `Paragraph Continuation` |
| `ListItem` | `\li1` | List / bullet item | `List Paragraph` |
| `AppendixTitle` | `\imt` | Appendix section title | `Normal` (Concordance heading) |
| `AppendixBody` | `\ip` | Appendix body text | `Plain Text` (Concordance body) |
| `CustomParaAfterH1` | `\p` (first) | First para after book title | *(keep — already semantic)* |
| `DatAuthRef` | `\d` | Descriptive / date-author reference | *(keep — already semantic)* |
| `Brief` | `\is` | Book-level introduction / brief | *(keep — intentional, 66 instances)* |
| `Psalms BOOK` | `\ms` | Psalms book division (Books I–V) | *(keep — intentional, 5 instances)* |
| `Lamentation` | `\q1` | Lamentations acrostic / verse text | *(keep — intentional, 152 instances)* |
| `TheHeaders` | — | Running header | *(keep)* |
| `TheFooters` | — | Running footer / page number | *(keep)* |
| `Footnote Text` | `\f` | Footnote | *(keep — built-in Word name with space)* |
| `Title` | — | Document title (artifact — Section 1) | *(tolerated — 1 instance)* |

Note: USFM markers for `Brief`, `Psalms BOOK`, and `Lamentation` are best-fit
suggestions — confirm against the USFM export plan when that work begins.

### Font context

| Font | Status |
|------|--------|
| Calibri | Replaced globally by Carlito (metrically identical, free) — **DONE** |
| Times New Roman | Awaiting substitution — **PENDING** |
| Allowed fonts, fallback fonts, CJK prep | In queue as part of i18n work — **FUTURE** |

### Key finding: Normal = Bible body text

`Normal` (Carlito 9pt) is the style the author used for all Bible text paragraphs.
Renaming `Normal` → `BodyText` (or replacing all `Normal` paragraphs with `BodyText`)
is the single highest-impact style fix — it addresses the bulk of unintentional `Normal`
usage in one operation and establishes the semantic foundation for the USFM `\p` export.

### Plain Text items 1–8 — root cause confirmed

These are accidental. Word has a known spacing bug: when a paragraph with space-before
is the first paragraph on a page, Word ignores the space-before setting. The workaround
is to insert an empty paragraph (with a tab as marker) ahead of the heading. Those empty
paragraphs ended up styled as `Plain Text` rather than `BodyText`. Once `BodyText` is
defined, they are all straightforward replacements.

### Implementation plan (step-by-step)

**Step 1 — Define `BodyText` style**
Create the `BodyText` style in the document: Carlito 9pt (matching current `Normal`
body text formatting), `\p` semantic. This must be done in Word before any VBA fix
routine runs.

**Step 2 — Replace `Normal` → `BodyText` (Bible text, ~16 000+ paragraphs)**
Largest operation. All `Normal` paragraphs become `BodyText`. This is the Bible text
fix. A fix routine in `basFixDocxRoutines` will iterate all paragraphs and replace.
Test: add test for `Normal` count = 0.

**Step 3 — Replace `Plain Text` → `BodyText` / `AppendixBody` (26 paragraphs)**
Items 1–8 (front matter spacers) → `BodyText`.
Items 9–26 (concordance) → `AppendixBody` (requires `AppendixBody` style defined first).
Test: update Test 49 expected once complete.

**Step 4 — Replace `Paragraph Continuation` → `BodyTextContinuation` (158 paragraphs)**
Requires `BodyTextContinuation` style defined in Word first. Investigate 158 instances
before fix to confirm all are continuation paragraphs (no misuse).

**Step 5 — Replace `List Paragraph` → `ListItem` (82 paragraphs)**
Requires `ListItem` style defined in Word. Check for nested lists → may need `ListItem2`.

**Step 6 — Define and apply `AppendixTitle` / `AppendixBody`**
Concordance section title ("Bible Concordance") → `AppendixTitle`.
Concordance body paragraphs → `AppendixBody`.

**Step 7 — Times New Roman substitution**
Separate from style fixes; tracked under i18n / font work.

### Normal style formatting (captured 2026-04-22)

VBA values for `Normal` — used verbatim in `DefineBodyTextStyle`:

| Property | VBA value | Unit | UI equivalent |
|----------|-----------|------|---------------|
| `Alignment` | 0 | enum | Left (`wdAlignParagraphLeft`) |
| `SpaceBefore` | 0 | points | 0pt |
| `SpaceAfter` | 0 | points | 0pt |
| `LineSpacingRule` | 0 | enum | Single (`wdLineSpaceSingle`) |
| `LineSpacing` | 12 | points | Computed — 9pt font at Single = 12pt |
| `FirstLineIndent` | 14.4 | points | 0.2" (÷ 72) |
| `LeftIndent` | 0 | points | 0" |
| `Font` | Carlito | — | Carlito |
| `Size` | 9 | points | 9pt |
| `Bold` | 0 | bool | False |
| `Italic` | 0 | bool | False |

Note: `LineSpacing: 12` is the computed auto value for 9pt Carlito at Single spacing.
When `LineSpacingRule = wdLineSpaceSingle`, Word calculates line height from font size —
the point value is informational, not prescriptive.

### Step 1 result — 2026-04-22

`DefineBodyTextStyle` run successfully. `BodyText` style created in document:
Carlito 9pt, Left aligned, Single spacing, 0.2" first-line indent, 0pt space before/after.
No cascade dependency on `Normal`.

### Step 2 result — 2026-04-22

`ReplaceNormalWithBodyText` run successfully.

```
ReplaceNormalWithBodyText: 31846 replaced, 0 remaining.
```

31,846 paragraphs converted from `Normal` → `BodyText`. This confirms `Normal` was
used for the entire Bible text body — the single largest style fix in the document.
`Normal` count in body story is now 0.

### Target state after Steps 1–6

Test 49 expected = 17 (only intentional styles remain):
`Heading 1`, `Heading 2`, `BodyText`, `BodyTextIndent`, `BodyTextContinuation`, `ListItem`,
`AppendixTitle`, `AppendixBody`, `CustomParaAfterH1`, `DatAuthRef`,
`Brief`, `Psalms BOOK`, `Lamentation`,
`TheHeaders`, `TheFooters`, `FootnoteText`, `Title`

Styles to eliminate (currently unintentional):
`Normal` (113 — non-body stories), `Plain Text` (26), `List Paragraph` (82),
`Paragraph Continuation` (158), `Header` (2 — investigate), `Footer` (3 — investigate)

---

## § 3 — Plain Text style investigation — 2026-04-21

### Diagnostic results

26 paragraphs use `Plain Text`. Two distinct locations:

**Items 1–8** (positions 1856–26595) — front matter, all blank/whitespace:

| Group | Items | Context pattern | Interpretation |
|-------|-------|----------------|----------------|
| A | 1–4 | `Normal → Plain Text → Normal` | Isolated blank spacer in front matter |
| B | 5–6 | `Normal → Plain Text → Plain Text → Heading 1` | Pre-book blank spacers |
| C | 7–8 | `DatAuthRef → Plain Text → Plain Text → Heading 2` | Pre-chapter blank spacers |

**Items 9–26** (positions 4180954+) — Concordance appendix at end of document:

```
Normal       |              ← blank
Normal       |              ← blank
Normal       | Bible Concordance    ← section title
Normal       |              ← blank
Normal       |              ← blank
Plain Text   |              ← blank
Plain Text   | A
Plain Text   | written concordance has been omitted...
Plain Text   | (body text paragraphs, software links, etc.)
```

### Two style problems in the concordance section

| Element | Current style | Should be |
|---------|--------------|-----------|
| "Bible Concordance" title | `Normal` | ? — `Heading 1` or section title style |
| Body text paragraphs | `Plain Text` | ? — document standard body text style |
| Blank spacers | `Normal` / `Plain Text` | depends on surrounding styles |

Note: `Normal` paragraphs surrounding `Plain Text` in both locations are also
unintentional — the front-matter and concordance sections have a broader style
problem, not just `Plain Text` in isolation.

### Open questions (required before fix routine)

1. What is the standard body text style in this document? (Check what style
   Genesis chapter text uses — likely `CustomBody` or similar.)
2. Should "Bible Concordance" be treated as `Heading 1` so it appears in the
   heading structure and gets a running header, or is it a standalone appendix
   with a different style?

### Recommended fix approach (pending answers above)

- Replace `Plain Text` body paragraphs in concordance → standard body text style
- Replace `Normal` "Bible Concordance" title → `Heading 1` (or appendix heading style)
- Blank `Plain Text` / `Normal` spacers before headings → review whether spacing
  should be handled by paragraph space-before on the heading style instead

---

## § 5 — Style taxonomy test suite (RUN_TAXONOMY_STYLES) — proposed 2026-04-22

### Concept

One reusable routine `AuditOneStyle` takes a style name and a record of expected
property values. `RUN_TAXONOMY_STYLES` calls it once per approved style (~18 calls)
and writes a structured report to `rpt\StyleTaxonomyAudit.txt`.

Proposed module: `basTEST_aeBibleConfig` (already exists; config tests live there).

### What AuditOneStyle checks per style

| Property | Check |
|----------|-------|
| Style exists | Yes/No |
| Font name | exact match |
| Font size | exact match |
| Alignment | enum match |
| FirstLineIndent | point value match |
| LineSpacingRule | enum match |
| SpaceBefore / SpaceAfter | point value match |
| BaseStyle | confirm not `Normal` (cascade guard) |

### Pros

- Single source of truth for the style specification — expected values live in
  `RUN_TAXONOMY_STYLES` constants, not scattered across the document
- Catches style drift silently introduced by Word (template bleed, format painter misuse)
- rpt file is a versioned snapshot — compare across sessions to detect regressions
- Rerun-safe and fast (no paragraph iteration — style object access only)
- Documents the intended spec in executable form; replaces comments that go stale

### Cons

- Expected values in code must be updated whenever a style is intentionally changed;
  a stale expected value gives a false failure
- 18 styles × 8 properties = 144 checks — report can be verbose; needs compact format
- Does not verify that paragraphs USE the style correctly — only that the style
  definition is correct; a paragraph with direct-format override will not be caught

### Suggestions

- Output format: one line per style — `PASS StyleName` or `FAIL StyleName: Font
  expected Carlito got TimesNewRoman` — compact, grep-friendly
- Group output by PASS / FAIL with a summary line at end: `17 PASS  1 FAIL`
- Store expected values as a Type or as inline constants at the top of
  `RUN_TAXONOMY_STYLES` — not hardcoded inside `AuditOneStyle`
- `AuditOneStyle` should be `Private` — only called by `RUN_TAXONOMY_STYLES`

### Results — 2026-04-22

| Run | PASS | FAIL | Notes |
|-----|------|------|-------|
| 1st (wrong document) | 10 | 7 | Ran against copy; `BodyText`/`BodyTextIndent` not found; `FootnoteText` name wrong |
| 2nd (correct doc) | 12 | 5 | `Footnote Text` fixed; `BodyTextIndent` still missing — `DefineBodyTextIndentStyle` not yet run |
| 3rd (baseline) | 13 | 4 | `BodyTextIndent` created; all 4 FAILs are expected not-yet-created styles |

**Baseline: 13 PASS / 4 FAIL** — the 4 failing styles (`BodyTextContinuation`, `ListItem`,
`AppendixTitle`, `AppendixBody`) will each convert to PASS as their `Define*` routine is run.

Fixes made during first run: `FootnoteText` → `Footnote Text` (built-in Word name has space).

### Status

**IMPLEMENTED — 2026-04-22.** Module: `basTEST_aeBibleConfig`.
Baseline confirmed: 13 PASS / 4 FAIL.

---

## § 6 — SUPER_TEST_RUNS global verification command — proposed 2026-04-22

### Concept

A single entry point that runs all verification suites in sequence and produces a
master report in `rpt\SuperTestReport.txt`. Each suite runs independently — a
failure in one suite does not abort the rest.

Proposed location: `basTest_aeBibleClass.bas` alongside `RUN_THE_TESTS`, or a
dedicated `basVerificationSuite.bas` if the orchestration grows large.

### Proposed suite sequence

| Order | Suite | Entry point | Output |
|-------|-------|-------------|--------|
| 1 | Style taxonomy | `RUN_TAXONOMY_STYLES` | `rpt\StyleTaxonomyAudit.txt` |
| 2 | Document diagnostics | `RUN_THE_TESTS` | `rpt\TestReport.txt` |
| 3 | Font audit | (existing font test routines) | `rpt\FontAudit.txt` |
| 4 | Header/footer audit | `SummarizeHeaderFooterAuditToFile` | `rpt\HeaderFooterAudit.txt` |
| 5 | Scripture parser | `basSBL_TestHarness` entry point | Immediate Window |

Master report: timestamp, one summary line per suite (PASS / FAIL / count), then
links to individual rpt files for drill-down.

### Pros

- Prevents "works in isolation, broken globally" scenarios — the most common
  regression pattern in a large document automation project
- Single command before any commit is a forcing function for quality
- Master report gives a health snapshot across all dimensions of the document
- Failure lines include the suite name and function — directly actionable without
  manually re-running individual tests to find the source
- Establishes a CI-like discipline without requiring external tooling

### Cons

- Runtime will be long — `RUN_THE_TESTS` alone is several minutes; full suite
  may be 10–20 minutes; not suitable for running after every small change
- Mixed test frameworks (aeBibleClass assertion arrays, parser harness, style audit)
  produce different output shapes — unifying into a master report requires adapters
- If a suite hangs or crashes at the VBA level, the orchestrator may not recover
  gracefully; needs per-suite `On Error` isolation
- Maintenance: adding a new test suite requires updating `Super_Test_Runs` manually

### Suggestions

- Add a `Quick` mode flag: `Super_Test_Runs(quickMode:=True)` skips slow suites
  (marked with X prefix convention already in use) for pre-commit checks
- Structure output as sections with clear delimiters so the file is easy to scan:
  `=== SUITE 1: Style Taxonomy === PASS (18/18)` then detail below
- Each suite call wrapped in `On Error Resume Next` with a caught-error line in
  the master report — ensures one crashing suite does not silence the rest
- Parser tests (suite 5) currently write to Immediate Window only — before
  including in Super_Test_Runs, the harness should be updated to write to a file
- Consider a `rpt\SuperTestReport.txt` that accumulates runs with timestamps so
  trend analysis is possible (append mode, not overwrite)

### Decision

- Name: `SUPER_TEST_RUNS` (caps, consistent with `RUN_THE_TESTS`)
- Location: **Option B** — new module `basVerificationSuite.bas`
- Rationale: keeps orchestration separate from individual test logic;
  scales cleanly as more suites are added

### Status

**DEFERRED — implement after taxonomy is working correctly and current
major editing work is complete.** All design decisions recorded above.

---
