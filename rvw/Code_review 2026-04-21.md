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
| `List Paragraph` | 82 | **DEFERRED** — cross-reference lookup table needs separate taxonomy design pass; three sub-types identified: `ListItem`, `ListItemBody`, `BookRef` |
| `Paragraph Continuation` | 158 | Investigate contexts → fix routine in `basFixDocxRoutines` → test |
| `Title` | 1 | Tolerated (artifact — Section 1 front-matter) |

Test 49 expected will move from 15 → 12 once Group B styles are resolved.

---

## § 4 — Style taxonomy — approved 2026-04-21

### Approved taxonomy

| Word Style | USFM marker | Semantic role | Replaces |
|-----------|-------------|---------------|----------|
| `Book Title` | `\mt1` | "Holy Bible" — document title (appears once) | *(keep — intentional)* |
| `Heading 1` | `\mt2` | Individual book title (66 — Genesis, Exodus, etc.) | *(reserved — keep)* |
| `Heading 2` | `\c` | Chapter heading | *(reserved — keep)* |
| `BodyText` | `\p` | Standard body paragraph | `Normal`, `Plain Text` |
| `BodyTextIndent` | `\pi` | Indented body paragraph (quoted / subordinate) | direct formatting |
| `BodyTextContinuation` | `\m` | Continuation paragraph (no indent) | `Paragraph Continuation` |
| `ListItem` | `\li1` | Study list item (Carlito 11pt Bold Italic, hanging indent) | `List Paragraph` Type A |
| `ListItemBody` | `\li1` | Study list continuation paragraph | `List Paragraph` Type A continuation |
| `BookRef` | `\xt` | Cross-reference to named book (Carlito 11pt Bold, tab-leader) | `List Paragraph` Type B |
| `AppendixTitle` | `\imt` | Appendix section title | `Normal` (Concordance heading) |
| `AppendixBody` | `\ip` | Appendix body text | `Plain Text` (Concordance body) |
| `CustomParaAfterH1` | `\p` (first) | First para after book title | *(keep — already semantic)* |
| `DatAuthRef` | `\d` | Descriptive / date-author reference | *(keep — already semantic)* |
| `BookIntro` | `\ip` | Centered book introduction summary (follows DatAuthRef) | direct formatting on `BodyText` |
| `Brief` | `\is` | Book-level introduction / brief | *(keep — intentional, 66 instances)* |
| `Psalms BOOK` | `\ms` | Psalms book division (Books I–V) | *(keep — intentional, 5 instances)* |
| `Lamentation` | `\q1` | Lamentations acrostic / verse text | *(keep — intentional, 152 instances)* |
| `TheHeaders` | — | Running header | *(keep)* |
| `TheFooters` | — | Running footer / page number | *(keep)* |
| `Footnote Text` | `\f` | Footnote | *(keep — built-in Word name with space)* |
| `Title` | — | Document title (artifact — Section 1) | *(tolerated — 1 instance)* |

Note: USFM markers for `Brief`, `Psalms BOOK`, and `Lamentation` are best-fit
suggestions — confirm against the USFM export plan when that work begins.

### Taxonomy reconciliation — 2026-04-22

Source of truth: `PromoteApprovedStyles` in `basTEST_aeBibleConfig.bas`.
Styles added to taxonomy after reconciliation:

| Word Style | USFM marker | Semantic role | Notes |
|-----------|-------------|---------------|-------|
| `Words of Jesus` | `\wj` | Words spoken by Jesus (red text) | *(keep — intentional)* |
| `EmphasisRed` | `\em` | Red emphasis | *(keep — intentional)* |
| `EmphasisBlack` | `\em` | Black emphasis | *(keep — intentional)* |
| `Chapter Verse marker` | `\c` / `\v` | Chapter/verse marker | *(keep — intentional)* |
| `Verse marker` | `\v` | Verse marker | *(keep — intentional)* |
| `Book Title` | `\mt2` | Book title (clarify vs `Heading 1`) | *(pending clarification)* |
| `Footnote Reference` | `\fr` | Footnote reference mark | *(keep — intentional)* |

`CustomParaAfterH1-2nd` — confirmed 0 paragraphs; removed from `PromoteApprovedStyles`.
`Body Text` (built-in Word style with space) — confirmed 0 paragraphs; removed.
`FargleBlargle` — intentional diagnostic dummy; always expected missing.

### PromoteApprovedStyles — updated 2026-04-22

Added all new styles from this session. Current approved list (in priority order):

```txt
Normal, Heading 1, Heading 2,
BodyText, BodyTextIndent, BodyTextContinuation,
CustomParaAfterH1, DatAuthRef, BookIntro,
Brief, Psalms BOOK, Lamentation,
AppendixTitle, AppendixBody,
ListItem,
Chapter Verse marker, Verse marker,
EmphasisBlack, EmphasisRed,
Words of Jesus,
AuthorBodyText, AuthorSectionHead,
AuthorQuote, AuthorRef,
TheHeaders, TheFooters,
Title, Book Title,
Footnote Reference, Footnote Text,
FargleBlargle
```

### PromoteApprovedStyles — run result — 2026-04-22

```
WARNING: 7 styles NOT found:
  BodyTextIndent, BodyTextContinuation, BookIntro,
  AppendixTitle, AppendixBody, ListItem, FargleBlargle
```

All 7 WARNs are expected:

- `FargleBlargle` — intentional diagnostic dummy; always missing
- Remaining 6 — not yet created in the DOCM; will be created by their `Define*` routines

**24 styles promoted** with correct priority order. Priority gaps (5, 6, 9, 13, 14, 15)
are placeholders for the missing styles — they close when `Define*` routines are run
and `PromoteApprovedStyles` is re-run.

| Priority | Style |
|----------|-------|
| 1 | Normal |
| 2 | Heading 1 |
| 3 | Heading 2 |
| 4 | BodyText |
| 7 | CustomParaAfterH1 |
| 8 | DatAuthRef |
| 10 | Brief |
| 11 | Psalms BOOK |
| 12 | Lamentation |
| 16 | Chapter Verse marker |
| 17 | Verse marker |
| 18 | EmphasisBlack |
| 19 | EmphasisRed |
| 20 | Words of Jesus |
| 21 | AuthorBodyText |
| 22 | AuthorSectionHead |
| 23 | AuthorQuote |
| 24 | AuthorRef |
| 25 | TheHeaders |
| 26 | TheFooters |
| 27 | Title |
| 28 | Book Title |
| 29 | Footnote Reference |
| 30 | Footnote Text |

### Critical bug — ReplaceNormalWithBodyText — 2026-04-22

**Root cause:** `ReplaceNormalWithBodyText` used Word Find/Replace with
`.Style = oDoc.Styles("Normal")`. Word's Find/Replace matches ALL paragraphs
styled with styles based on Normal — including `Words of Jesus`, `EmphasisRed`,
`EmphasisBlack`, `Chapter Verse marker`, `Verse marker`, etc. This replaced
31,846 paragraphs including many that were not true `Normal` paragraphs,
destroying their semantic style assignments.

**Recovery:** Document closed without saving; reverted to 955-page backup.

**Fix:** `ReplaceNormalWithBodyText` must use paragraph iteration with exact
`NameLocal = "Normal"` matching — never Find/Replace for style replacement.
This ensures child styles (styles based on Normal) are never affected.

**Status:** Routine rewritten and re-run on clean document. COMPLETE.

### ReplaceNormalWithBodyText — result — 2026-04-22

**Run:** 31,846 Normal paragraphs replaced with BodyText using exact iteration.
Child styles (Words of Jesus, EmphasisRed, EmphasisBlack, etc.) confirmed unaffected.

**Outcome:**

| Section | Pages | Result |
|---------|-------|--------|
| Bible text (Genesis–Revelation) | ~905 | Correct — BodyText Exactly 10pt, justified |
| Front matter (author text) | first 18 | Formatting affected — Times New Roman |
| Back matter (author text) | last 35 | Formatting affected — Times New Roman |
| **Total affected** | **53 / 960** | **~5.5% — tolerated** |

**Root cause of front/back matter issues:** Those sections used `Normal` style
but with Times New Roman font applied as direct formatting (or via a style based
on Normal that inherits TNR). Replacing Normal → BodyText (Carlito 9pt Exactly 10pt)
changed the font and line spacing in those sections.

**Assessment: WIN.** Times New Roman substitution is already planned as a separate
work item. The front/back matter font fix is absorbed into that task. Bible text —
the primary content — is correctly converted.

**Next:** Times New Roman substitution (front/back matter font fix) tracked under
i18n / font work. Current page count: 960.

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

```txt
ReplaceNormalWithBodyText: 31846 replaced, 0 remaining.
```

31,846 paragraphs converted from `Normal` → `BodyText`. This confirms `Normal` was
used for the entire Bible text body — the single largest style fix in the document.
`Normal` count in body story is now 0.

### Target state after Steps 1–6

Test 49 expected = 18 (only intentional styles remain):
`Heading 1`, `Heading 2`, `BodyText`, `BodyTextIndent`, `BodyTextContinuation`, `ListItem`,
`AppendixTitle`, `AppendixBody`, `BookIntro`, `CustomParaAfterH1`, `DatAuthRef`,
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

```txt
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

## § 7 — Author text styles — Times New Roman substitution — 2026-04-22

### Context

Front matter (first 18 pages) and back matter (last 35 pages) are author-written
text — biographical notes, diagrams, and appendix content — originally formatted
in Times New Roman 12pt. After `ReplaceNormalWithBodyText`, these sections lost
their TNR formatting (replaced by Carlito 9pt Exactly 10pt). A dedicated style set
is needed to restore correct formatting before manual cleanup of these 53 pages.

### Font choice — Liberation Serif

Times New Roman is a proprietary font (Monotype). Liberation Serif is the
metrically compatible free replacement (same relationship as Carlito/Calibri):
identical metrics ensure no reflowing when the document is opened on a system
without TNR.

### Style decisions

| Aspect | Decision |
|--------|----------|
| Author body text font | Liberation Serif 12pt (replaces TNR 12pt) |
| Author section heads font | Liberation Serif 14pt (replaces TNR 14pt) |
| Section head style | No bold/italic in the style definition — individual words alternate bold/italic (applied as direct formatting word by word) |
| Inline quotes of Jesus | `AuthorQuote` character style — Italic + Red (wdColorRed) |
| Inline book references | `AuthorRef` character style — Bold |
| Character style preference | Character styles over direct formatting — auditable, reversible |

### Styles defined — DefineAuthorStyles

Four styles added in one routine in `basFixDocxRoutines.bas`:

**`AuthorBodyText`** (paragraph style):

- Liberation Serif 12pt
- Justified (`wdAlignParagraphJustify`)
- FirstLineIndent = 23.76pt (0.33" × 72 = 23.76)
- SpaceAfter = 12pt
- `wdLineSpaceSingle`
- WidowControl = True
- BaseStyle = "" (no cascade from Normal)

**`AuthorSectionHead`** (paragraph style):

- Liberation Serif 14pt, plain (no Bold/Italic in style)
- Left aligned
- SpaceBefore = 12pt, SpaceAfter = 6pt
- `wdLineSpaceSingle`
- WidowControl = True
- Individual words styled bold or italic directly in the document

**`AuthorQuote`** (character style):

- wdStyleTypeCharacter
- Italic = True
- Color = wdColorRed
- Semantic: inline quotes of Jesus in author commentary

**`AuthorRef`** (character style):

- wdStyleTypeCharacter
- Bold = True
- Semantic: references to named book sections in author text

All four styles are rerun-safe — the routine skips styles that already exist.

### USFM mapping

| Word Style | USFM marker | Notes |
|-----------|-------------|-------|
| `AuthorBodyText` | `\ip` | Author introductory paragraph |
| `AuthorSectionHead` | `\is` | Author introductory section heading |
| `AuthorQuote` | `\wj` | Words of Jesus (inline character) |
| `AuthorRef` | `\bd` | Bold inline reference |

### Application plan

After importing `basFixDocxRoutines` into the `.DOCM` and running `DefineAuthorStyles`:

1. Manually select front matter body paragraphs → apply `AuthorBodyText`
2. Manually select section headers → apply `AuthorSectionHead`
3. Select red italic quote runs → apply `AuthorQuote` character style
4. Select bold book reference runs → apply `AuthorRef` character style
5. Run `RUN_TAXONOMY_STYLES` to verify style definitions match spec

### Taxonomy additions

Add to approved taxonomy table:

| Word Style | USFM marker | Semantic role | Type |
|-----------|-------------|---------------|------|
| `AuthorBodyText` | `\ip` | Author body text (Liberation Serif 12pt, 0.33" indent) | Paragraph |
| `AuthorSectionHead` | `\is` | Author section heading (Liberation Serif 14pt, mixed bold/italic) | Paragraph |
| `AuthorQuote` | `\wj` | Inline quote of Jesus (Italic, Red) | Character |
| `AuthorRef` | `\bd` | Inline book section reference (Bold) | Character |

### Style code fixes — confirmed 2026-04-23

| Style | WidowControl | PageBreakBefore | Status |
|-------|-------------|-----------------|--------|
| `AuthorBodyText` | True | False | **CONFIRMED** — matches DOCM |
| `AuthorSectionHead` | False | True | **FIXED in src** — `PageBreakBefore = True` added |

`AuthorRef` renamed to `AuthBookRef` and applied — **DONE — 2026-04-23**.
`AuthorQuote` — not used in back matter. Status in front matter
still undecided. No action until front matter work resumes.

### Front matter page structure — corrected 2026-04-22

Two distinct pages with different mechanisms:

**"Books of the Bible" page** — the 66-book page-number listing

- Physical layout: 4 grouped text boxes (OT col 1, OT col 2, NT col 1, NT col 2)
- Each entry: book name + SBL abbreviation + `{ DOCVARIABLE }` field for page number
- Example: `{ DOCVARIABLE 1Sam }` already set up and visible via Alt+F9
- Variables defined in `SetDocVariables` (`XbasTESTaeBibleDOCVARIABLE`) but value not
  yet populated — trigger code (page number scan) is not yet wired up
- Standard Word TOC engine is NOT used here — too slow, too rigid for 66 entries
  in text boxes

**"Contents" page** — front/back matter section listing

- Lists major sections only: OT, NT, Maps, Concordance, etc. (~10 entries)
- Standard Word TOC is acceptable at this scale (fast for small entry counts)
- Or DOCVARIABLE fields for consistency with the Books of the Bible page
- This page carries the `TitleEyebrow` + `Title` heading

### Navigation pane vs Contents page — 2026-04-22

Word's navigation pane and TOC are driven independently:

- **Navigation pane / Outline view** — shows paragraphs with outline level 1–9.
  Outline level "Body Text" (0) removes a style from the pane entirely.
- **TOC** — can map any named style to a TOC level via the `\t` switch,
  independent of outline level.

A paragraph can appear in the TOC without appearing in the navigation pane.

**For this document:**

- Ribbon navigation covers only the 66 canonical books (`Heading 1` positions).
  `Title` / `TitleEyebrow` are outside this scope — no ribbon change needed.
- The Books of the Bible page uses DOCVARIABLE — no TOC involvement at all.
- The Contents page (~10 major sections) uses standard Word TOC or DOCVARIABLE.
- `TitleEyebrow` / `Title` heading on the Contents page: outline level Body Text,
  not in nav pane, optionally in the Contents TOC via `\t "Title,1"`.

### Two-line display title — TitleEyebrow + Title — 2026-04-22

The heading "The / HOLY BIBLE" (eyebrow + main title) cannot use `Heading 1`
— reserved for the 66 book titles. Reusable across front matter display pages.

**Recommended: Option B — two styles**

| Style | Role | Outline | TOC |
|-------|------|---------|-----|
| `TitleEyebrow` | "The" (preceding line), small centered | Body Text | none |
| `Title` | "HOLY BIBLE" (main line), large display centered | Body Text | Level 1 via `\t` |

`TitleEyebrow.NextParagraphStyle = Title`. `Title` already exists (1 instance);
needs formal definition. `TitleEyebrow` is new.

**Status:** Design approved — implementation pending.

### DOCVARIABLE — chosen approach for all page number references — 2026-04-22

**Decision: DOCVARIABLE for both pages.**

| Page | Variables | Notes |
|------|-----------|-------|
| Books of the Bible | 66 (one per canonical book) | OT/NT text boxes |
| Contents | ≤ 10 (major sections) | OT, NT, Maps, Concordance, etc. |
| **Total** | **≤ 76** | One methodology, one updater, one button |

**Rationale:**

- One methodology — no `\t` TOC switch manipulation, no TOC field options dialog
- 66 variables already planned; adding ≤ 10 more is negligible setup overhead
- One `UpdatePageNumbers` call updates everything in both pages in one pass
- Wire to `Document_BeforePrint` → set-it-and-forget-it

**Cons:**

| # | Con | Mitigation |
|---|-----|-----------|
| 1 | Values go stale silently after any edit that shifts page breaks | Wire to `Document_BeforePrint` in `ThisDocument.cls` — fires automatically before every print/export |
| 2 | Document must be fully paginated — cold open gives wrong numbers | `BeforePrint` fires after Word has paginated; also runs correctly after warm cache |
| 3 | Mismatched variable name in field code shows blank silently | Validation loop in updater — warn if any variable written as 0 or unchanged |
| 4 | Three-way sync: field code in doc + `SetDocVariables` + `SBLVarName` | One-time setup cost; convention: SBL abbreviation with spaces stripped = variable name (`1 Sam` → `1Sam`). `SBLVarName` must apply `Replace(sAbbrev, " ", "")` |
| 5 | `Fields.Update` may not reach fields inside grouped text boxes | Iterate `ActiveDocument.Shapes` explicitly as a safety net — one-time verification needed |

Cons 1 and 2 are fully resolved by the `BeforePrint` hook.
Con 3 resolved by a validation pass in the updater.
Con 4 is one-time setup, not ongoing burden.
Con 5 needs a single test after implementation.

### DOCVARIABLE trigger code design — 2026-04-22

What `XbasTESTaeBibleDOCVARIABLE` has:

- `SetDocVariables` — defines the 66 variable names and SBL abbreviation mapping
- One live `{ DOCVARIABLE 1Sam }` field confirmed via Alt+F9; value not yet populated

What is missing — `UpdatePageNumbers` (covers both pages in one call):

```vba
Public Sub UpdatePageNumbers()
    ' Pass 1: 66 canonical books from Heading 1 paragraphs
    Dim oPara As Word.Paragraph
    Dim sVar  As String
    Dim lPage As Long
    For Each oPara In ActiveDocument.Content.Paragraphs
        If oPara.Style.NameLocal = "Heading 1" Then
            sVar = SBLVarName(oPara.Range.Text)   ' SBL abbrev spaces stripped e.g. "1 Sam" -> "1Sam"
            lPage = oPara.Range.Information(wdActiveEndPageNumber)
            ActiveDocument.Variables(sVar).Value = CStr(lPage)
        End If
    Next oPara

    ' Pass 2: Contents page sections (loop over ~10 section variables)
    ' ... similar pattern for major section openers ...

    ' Refresh fields in body story
    ActiveDocument.Fields.Update
    ' Refresh fields inside text boxes (grouped shapes)
    Dim oShp As Shape
    For Each oShp In ActiveDocument.Shapes
        If oShp.TextFrame.HasText Then
            oShp.TextFrame.TextRange.Fields.Update
        End If
    Next oShp

    Debug.Print "UpdatePageNumbers: Done."
End Sub
```

Wire to `Document_BeforePrint` in `ThisDocument.cls`:

```vba
Private Sub Document_BeforePrint(Cancel As Boolean)
    UpdatePageNumbers
End Sub
```

`SBLVarName` — refactor the existing book→variable mapping in `SetDocVariables`
into a callable `Function`. Must return the space-stripped form of the SBL
abbreviation (`Replace(sAbbrev, " ", "")`) since DOCVARIABLE names cannot
contain spaces. SBL `1 Sam` → variable `1Sam`; SBL `Song` → variable `Song`.

**Status:** Design complete. Implementation deferred — promote from
`XbasTESTaeBibleDOCVARIABLE` when front matter work resumes.

### Status

`DefineAuthorStyles` **IMPLEMENTED** in `basFixDocxRoutines.bas`.
Author styles applied to back matter: **DONE — 2026-04-22**.
Author styles applied to front matter: **PENDING**.
`AuthorSectionHead` — `PageBreakBefore = True` added to src — **DONE — 2026-04-23**.
`AuthorRef` renamed to `AuthBookRef` and applied — **DONE — 2026-04-23**.
`AuthorQuote` — deferred; not used in back matter; front matter TBD.
`TitleEyebrow` style definition: **PENDING**.
`Title` style formalization: **PENDING**.
RUN_TAXONOMY_STYLES additions for AuthorBodyText/AuthorSectionHead: **PENDING**.

---

## § 8 — Pending work as of 2026-04-23

### Completed

| Task | Status |
|------|--------|
| Import `basFixDocxRoutines` into .DOCM | **DONE** |
| Run `DefineAuthorStyles` | **DONE** |
| Run `DefineBodyTextIndentStyle` | **DONE — 2026-04-22** |
| Run `DefineAppendixBodyStyle` + `DefineAppendixTitleStyle` | **DONE — 2026-04-22** |
| Apply author styles to back matter | **DONE — 2026-04-22** |
| `AuthorSectionHead` — `PageBreakBefore = True` in src | **DONE — 2026-04-23** |

### Next — List Paragraph → ListItem + ListItemBody — 2026-04-23

**ON HOLD:** `DefineBookIntroStyle`, `ReplacePlainTextStyles`, `ApplyBookIntroAfterDatAuthRef`,
`DefineAppendixTitleStyle`, `DefineAppendixBodyStyle`

Reason: `List Paragraph` overlap with Concordance — define list styles first,
then decide whether AppendixBody/AppendixTitle are redundant.

#### List Paragraph — confirmed spec

| Style | Font | Size | Bold | Italic | LeftIndent | SpaceAfter | Align | WidowControl | Next |
|-------|------|------|------|--------|-----------|------------|-------|-------------|------|
| `ListItem` | Carlito | 11pt | Yes | Yes | 36pt (0.5") | 0 | Left | True | `ListItemBody` |
| `ListItemBody` | Carlito | 11pt | No | No | 36pt (0.5") | 11pt | Left | True | `ListItem` |

USFM: `ListItem` → `\li1`, `ListItemBody` → `\lim1`

Base style: none (no cascade from Normal or List Paragraph built-in).

**Run order:** `DefineListItemBodyStyle` first, then `DefineListItemStyle`
(ListItem references ListItemBody as NextParagraphStyle).

#### Concordance

Currently uses bullet points (`List Paragraph` with bullets). Once `ListItem` /
`ListItemBody` are applied, the Concordance entries use the same pair —
no separate `AppendixBody` needed for Concordance body text.
`AppendixTitle` may still be needed for the "Bible Concordance" section heading
— decision deferred until styles are applied.

#### PromoteApprovedStyles + RUN_TAXONOMY_STYLES — updated 2026-04-23

`ListItemBody` added to `PromoteApprovedStyles` array (after `ListItem`).

`RUN_TAXONOMY_STYLES` updated:

- `ListItem` and `ListItemBody` moved from expected FAIL → existence-verified section
- Full spec checks added: Carlito 11pt, Left align (0), SpaceAfter 0 / 11pt

```txt
AuditOneStyle "ListItem",     "Carlito", 11, 0, 0, -1, -999, 0, 0
AuditOneStyle "ListItemBody", "Carlito", 11, 0, 0, -1, -999, 0, 11
```

Expected FAIL section now contains: `BodyTextContinuation`, `AppendixTitle`, `AppendixBody` (3).

#### Status

`DefineListItemBodyStyle` + `DefineListItemStyle` — **IMPLEMENTED — 2026-04-23**
in `basFixDocxRoutines.bas`.
`PromoteApprovedStyles` + `RUN_TAXONOMY_STYLES` — **UPDATED — 2026-04-23**
in `basTEST_aeBibleConfig.bas`.
Pending: import and run in DOCM, apply styles manually, then decide on AppendixTitle/AppendixBody.

### Deferred

| Task | Reason |
|------|--------|
| `DefineBookIntroStyle` + `ApplyBookIntroAfterDatAuthRef` | ON HOLD — pending List Paragraph investigation |
| `ReplacePlainTextStyles` | ON HOLD — pending List Paragraph investigation |
| `Paragraph Continuation` → `BodyTextContinuation` (158 paragraphs) | Investigate first |
| `AuthorQuote` | Not used in back matter; front matter TBD (`AuthorRef` renamed to `AuthBookRef` and applied 2026-04-23) |
| `TitleEyebrow` + `Title` formalization | Front matter work — deferred |
| DOCVARIABLE `UpdatePageNumbers` implementation | Front matter work — deferred |
| `SUPER_TEST_RUNS` | Deferred until taxonomy stable |
| Allowed fonts / fallback fonts / CJK prep | i18n queue |
| Add author styles to `RUN_TAXONOMY_STYLES` | After front matter styles settled |

---

## § 9 — Normalizer: Word.Style added — 2026-04-22

### Problem

`Word.Style` was not in `NORMALIZATIONS`. The VBA IDE lowercases type qualifiers
when a reference is missing (`Word.Style` → `Word.style`). Since the rule was absent,
exports silently retained the wrong casing.

Three instances in `basFixDocxRoutines.bas` were affected:

- Line 549 — `ReplaceNormalWithBodyText`
- Line 774 — `ReplacePlainTextStyles`
- Line 956 — `ApplyBookIntroAfterDatAuthRef`

Additionally, `DefineAuthorStyles` had a duplicate `Dim oCheck As Word.Style`
declaration (two `Dim` lines for the same variable in one procedure — VBA compile
error). Fix: remove the duplicate `Dim` in `src/basFixDocxRoutines.bas` and
reimport — the DOCM is always overwritten from `src/`.

### Fix

1. `py/normalize_vba.py` — added rule after `As Word.Section`:

   ```python
   (r'(?i)\bAs\s+(?:Word\.)?Style\b', 'As Word.Style', 'As Word.Style type declaration — added 2026-04-22'),
   ```

2. `src/basFixDocxRoutines.bas` — all 3 occurrences corrected (`replace_all`).
3. DOCM — duplicate `Dim` in `DefineAuthorStyles` must be removed manually.

### Coverage gap pattern

The normalizer covers: `Word.Range`, `Word.Paragraph`, `Word.Paragraphs`,
`Word.Section`, **`Word.Style`** (new).
Still not covered: `Word.Document`, `Word.Field`, `Word.Table`, `Word.Row`,
`Word.Cell`, `Word.HeaderFooter` — add if/when the IDE downcases them on export.

### Status

**FIXED — 2026-04-22.**

---

## § Front page styles — approach decided 2026-04-23

The Bible front page has a distinct layout: different title treatment and some
text differences from the rest of the document.

### Options considered

1. **Separate semantic styles** per element (e.g. `FrontTitle`, `FrontSubtitle`,
   `FrontBodyText`) — adds a few entries to the taxonomy, each greppable and
   reversible, consistent with the Author\* family.
2. **Override existing styles** with direct formatting on the front page —
   minimal style count but loses semantic meaning and breaks the auditability
   principle the rest of the taxonomy relies on.
3. **One base "front" style overridden per element** — sounds economical but
   defeats the point: once overridden, the style no longer describes what the
   text *is*.

### Decision

Go with **option 1 — separate semantic styles**. Consistent with `TitleEyebrow`
and `Title` formalization already PENDING, and with the direction set by the
Author\* style family (semantic, USFM-mappable, auditable).

Scope is bounded — likely 3–5 styles total for the front page.

### Next steps

1. User formats the front page manually to establish the visual target.
2. Then we name the styles in a consistent manner (naming scheme to align with
   the `TitleEyebrow` / `Title` / `Author*` conventions already in use).
3. Add `DefineFrontPageStyles` (or equivalent) to `basFixDocxRoutines.bas`
   following the same pattern as `DefineAuthorStyles`.
4. Update `RUN_TAXONOMY_STYLES` to include the new styles.

### Status

**PENDING** — awaiting front page formatting pass by user.

---

## § Style inspector — added 2026-04-23

A reusable utility to dump a style's properties in paste-ready VBA form, so
manually-formatted paragraphs can feed the `Define<Style>` routines. Also
intended to become the raw input for a future markdown style guide keyed by
each style's Priority number.

### Style numbering context

Priority numbering now reflects the order styles appear in the book:

| Priority | Style |
|---------:|-------|
| 1 | `FrontPageTopLine` |
| 2 | `TitleEyebrow` |
| 3 | `Title` |
| 4 | `TitleVersion` |
| 5 | `FrontPageBodyText` |
| 6 | `Normal` |
| 7 | `Heading 1` |
| 8 | `Heading 2` |
| 9 | `BodyText` |
| 10 | `BodyTextIndent` |
| 12 | `CustomParaAfterH1` |
| 13 | `DatAuthRef` |
| 15 | `Brief` |
| 16 | `Psalms BOOK` |
| 17 | `Lamentation` |
| 20 | `ListItem` |
| 21 | `ListItemBody` |
| 22 | `Chapter Verse marker` |
| 23 | `Verse marker` |
| 24 | `EmphasisBlack` |
| 25 | `EmphasisRed` |
| 26 | `Words of Jesus` |
| 27 | `AuthorBodyText` |
| 28 | `AuthorSectionHead` |
| 29 | `AuthorQuote` |
| 30 | `AuthorBookRef` |
| 31 | `TheHeaders` |
| 32 | `TheFooters` |
| 33 | `Book Title` |
| 34 | `Footnote Reference` |
| 35 | `Footnote Text` |

Priorities 1–5 are front-page styles. Gaps (11, 14, 18–19) reserved for future
insertions without wholesale renumbering. `AuthorRef` → `AuthorBookRef` rename
is visible at priority 30.

### New module — `basStyleInspector.bas`

`Public Sub DumpStyleProperties(sStyleName As String, Optional bWriteFile As Boolean = False)`

Prints a named style's properties to the Immediate window in a form that can
be pasted into a `Define<Style>` routine. With `bWriteFile:=True`, also writes
`rpt\style_<name>.txt`.

Output includes:

- Style header line — name, type, priority
- `.BaseStyle`, `.NextParagraphStyle`, `.AutomaticallyUpdate`, `.QuickStyle`
- `.Font.*` — Name, Size, Bold, Italic, Underline, Color, SmallCaps, AllCaps
- `.ParagraphFormat.*` — Alignment, LeftIndent, RightIndent, FirstLineIndent,
  SpaceBefore, SpaceAfter, LineSpacing, LineSpacingRule, WidowControl,
  KeepTogether, KeepWithNext, PageBreakBefore, OutlineLevel
  (paragraph-type styles only)

Late binding, no added references, `Option Private Module` — matches project
conventions.

### Usage

From the Immediate window:

```vba
DumpStyleProperties "FrontPageBodyText"
DumpStyleProperties "FrontPageBodyText", True
```

### Design — pros

- **Paste-ready output** — lines are valid VBA `With`-block content; copy
  straight into a new `Define*` routine.
- **Non-mutating** — safe to run on any document.
- **Two output sinks** — Immediate window for a quick look, `rpt\` file for
  diffing and style-guide generation.
- **Handles paragraph + character styles** — skips `ParagraphFormat` when it
  doesn't apply.

### Design — cons (deliberate YAGNI scope)

- **Dumps everything, not just divergence** — noisy when chasing one property.
- **Raw enum integers** — `Alignment = 1` not `wdAlignParagraphCenter`.
- **Not exhaustive** — skips `Tabs`, `Shading`, `Borders`, `Frame`,
  `NoProofing`, `LanguageID`. Add per-style if a new style actually needs one.
- **Type detection** uses hard-coded numeric constants (1–4).
- **No guard** for unsaved documents — `ActiveDocument.Path` is empty and the
  `rpt\` file write will fail. Add a guard only if encountered.

### Use cases

- **UX → code** — format a paragraph visually, dump, paste into
  `DefineFrontPageBodyText` (or equivalent).
- **Style-guide generation** — `rpt\style_*.txt` files become the raw input
  for a markdown report ordered by Priority.
- **Regression diffing** — dump before/after a style change, diff the text
  files.
- **Audit** — compare runtime style state against the `Define*` routine's
  intent.

### Possible refinements (flagged, not scheduled)

- **Lean mode** — emit only properties that differ from a baseline style.
- **Symbolic enum output** — translate `Alignment`, `LineSpacingRule`,
  `Underline`, etc. to their `wd*` constants as trailing comments.
- **Batch dump** — accept an array of names, or dump all styles whose Priority
  falls in a range (e.g., 1–5 for front page).
- **Full-skeleton mode** — wrap output in a complete
  `Sub DefineXxx(oDoc As Object)` boilerplate.

### Status

**IMPLEMENTED — 2026-04-23** in `src\basStyleInspector.bas`.

### First run — `FrontPageBodyText` dump, 2026-04-23

Initial dump (`rpt\style_FrontPageBodyText.txt`) revealed two inconsistencies
with the front-page-style decision, both fixed in the UI and re-dumped:

| Property | Before | After | Reason |
|----------|--------|-------|--------|
| `.BaseStyle` | `"AuthorBodyText"` | `""` | Front-page styles must be isolated — no cascading from body styles |
| `.AutomaticallyUpdate` | `True` | `False` | Prevents direct formatting from silently rewriting the style definition |

Side effect of the UI edits (intentional, captured on re-dump):

- `.QuickStyle` — `True` → `False`

Line spacing was briefly observed as `13.9` / `wdLineSpaceMultiple` in an
intermediate dump, then reset to Single in the final dump.

### Clean dump — raw inspector output (paste-ready)

```
'--- FrontPageBodyText  (Type=Paragraph, Priority=5) ---
.BaseStyle = ""
.NextParagraphStyle = "FrontPageBodyText"
.AutomaticallyUpdate = False
.QuickStyle = False
.Font.Name = "Liberation Serif"
.Font.Size = 11
.Font.Bold = 0
.Font.Italic = 0
.Font.Underline = 0
.Font.Color = -16777216
.Font.SmallCaps = 0
.Font.AllCaps = 0
.ParagraphFormat.Alignment = 1
.ParagraphFormat.LeftIndent = 0
.ParagraphFormat.RightIndent = 0
.ParagraphFormat.FirstLineIndent = 0
.ParagraphFormat.SpaceBefore = 0
.ParagraphFormat.SpaceAfter = 11
.ParagraphFormat.LineSpacing = 12
.ParagraphFormat.LineSpacingRule = 0
.ParagraphFormat.WidowControl = -1
.ParagraphFormat.KeepTogether = 0
.ParagraphFormat.KeepWithNext = 0
.ParagraphFormat.PageBreakBefore = 0
.ParagraphFormat.OutlineLevel = 10
```

### Symbolic-constant decoding (manual, not emitted by inspector)

The inspector emits raw values only. Symbolic-name decoding remains a
*Possible refinement* (not yet implemented). Manual decoding for reference
when writing `DefineFrontPageBodyText`:

| Raw | Symbolic |
|-----|----------|
| `Font.Color = -16777216` | `wdColorAutomatic` |
| `Alignment = 1` | `wdAlignParagraphCenter` |
| `LineSpacing = 12`, `LineSpacingRule = 0` | Single (12pt) — `wdLineSpaceSingle` |
| `WidowControl = -1` | `True` |
| `OutlineLevel = 10` | `wdOutlineLevelBodyText` |

### Validation benefit observed

The inspector immediately exposed the two drift points (`BaseStyle`,
`AutomaticallyUpdate`) before a `Define*` routine was written — exactly the
"audit" use case the tool was built for. Worth running a dump on every new
front-page style before treating it as the source of truth.

### QA checklist — approved style files

When manually reviewing each dumped `rpt\style_<name>.txt`, every approved
style should default to the following, with documented exceptions only:

| # | Property | Expected | UI equivalent | Applies to |
|---|----------|----------|---------------|------------|
| 1 | `.BaseStyle` | `""` | Style based on: **(no style)** | All styles |
| 2 | `.AutomaticallyUpdate` | `False` | "Automatically update" checkbox **NOT** selected | Paragraph only |
| 3 | `.QuickStyle` | `False` | Style does **not** appear in the Quick Styles gallery | All styles |
| 4 | `.ParagraphFormat.LineSpacingRule` | `0` (`wdLineSpaceSingle`) | Line spacing: **Single** | Paragraph only |

`.AutomaticallyUpdate` and `.ParagraphFormat.*` are not valid on character
styles (accessing them raises run-time error 5900). Those fields are simply
absent from character-style dumps, so QA items 2 and 4 apply only where the
dump emits them.

Expected exceptions (documented as encountered):

- `ListItem` / `ListItemBody` — may require inter-style dependencies; TBD
  when those are reviewed.
- Other styles identified during the QA pass will be recorded here.

### Notes on `QuickStyle = False` default

- Quick Styles gallery visibility is largely useless for this project —
  depends on font size and name length, and with 30+ styles the gallery
  becomes noise rather than signal.
- Styles 1–5 are front-page only; no reason to expose them in the gallery.
- There is no large audience of editors. Onboarded users are expected to
  know how to use the Styles dropdown / Styles pane.
- The final per-style decision (show in gallery or not) is deferred to the
  enum-comment pass, by which time usage patterns will be established.

### Two-phase validation workflow

**Phase 1 — bulk dump**

Dump every approved style to `rpt\style_<name>.txt`. Planned driver:
`DumpAllApprovedStyles` (batch helper to be added to `basStyleInspector.bas`)
— iterates the Priority-ordered style list and calls
`DumpStyleProperties name, True` for each.

**Phase 2 — QA pass**

Read through the `rpt\` files and verify each against the four-point checklist
above. Record any intentional exceptions in the QA exceptions list.

**Phase 3 — symbolic decoration (`OutputAllApprovedStyles`)**

Once the QA pass is clean, implement `OutputAllApprovedStyles` — a successor
to `DumpStyleProperties` that emits the same data with trailing
`' wd*` comments next to each enum value (`Alignment`, `LineSpacingRule`,
`Underline`, `OutlineLevel`, `Color`, etc.). Run once across all approved
styles to produce the canonical decorated `rpt\` files that feed the markdown
style guide.

Order matters: decorate after QA, not before — the decoration pass is noisy
and easier to diff once the raw values are known-good.

### Status

QA checklist: **DEFINED — 2026-04-23**.
`DumpAllApprovedStyles` batch driver: **IMPLEMENTED — 2026-04-23** in
`src\basStyleInspector.bas`. Source of truth is the live document
(`Priority <> 99`), no duplicate approved-list — consistent with
`PromoteApprovedStyles` / `DumpPrioritiesSorted` convention.

First-run bug fixed 2026-04-23: accessing `.AutomaticallyUpdate` on the first
character style in priority order (`Chapter Verse marker`, priority 22)
raised run-time error 5900 and aborted the batch. Fix:

- `NextParagraphStyle` and `AutomaticallyUpdate` moved inside the
  `If oStyle.Type = 1 Then` (wdStyleTypeParagraph) guard in
  `DumpStyleProperties`.
- Per-iteration `On Error Resume Next` added to `DumpAllApprovedStyles`
  so a future single-style failure logs `!! FAILED` and the batch continues.
  Summary line shows `N succeeded, M failed`.

### Reconciliation with `DumpPrioritiesSorted` - 2026-04-24

Initial 31-vs-27 discrepancy between `DumpAllApprovedStyles` (31 styles
found) and `DumpPrioritiesSorted` (27 listed) was caused by stale priorities
on four front-page styles left over from an earlier `approved` array. The
array in `src\basTEST_aeBibleConfig.bas` was missing:

- `FrontPageTopLine`
- `TitleEyebrow`
- `TitleVersion`
- `FrontPageBodyText`

(`Title` was already present.)

Fixed by adding the four missing front-page names to the head of the
`approved` array. Both reports now return 31 styles in matching priority
order. Gaps at priorities 6, 11, 14, 18-19 are intentional placeholders for
styles not yet defined in the document (`BodyTextContinuation`, `BookIntro`,
`AppendixTitle`, `AppendixBody`) plus the deliberate `FargleBlargle`
missing-style canary at the end of the array.

`OutputAllApprovedStyles` symbolic-decoration pass: **PENDING** — after QA.

### Rationale for the QA workflow

Scanning `rpt\style_*.txt` files against the 4-point QA checklist is
dramatically faster than clicking through Modify Style dialogs in Word for
30+ styles. The checklist pulls out the drift points up front — bad
`BaseStyle`, `AutomaticallyUpdate = True`, `QuickStyle = True`,
non-single line spacing — without ever opening the menus.

---
