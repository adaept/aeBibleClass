# Code review - 2026-05-16 carry-forward

This file opens a fresh review arc on 2026-05-16. The previous arc
[`rvw/Code_review 2026-05-15.md`](Code_review%202026-05-15.md) is
now **closed for new work**; that file remains the authoritative
dated history for everything between 2026-05-15 and 2026-05-16,
including:

- **Test 5 added** — `CountApprovedStylesInGallery`. Enforces the
  editorial policy that approved styles must not appear in the
  Styles ribbon gallery. Closes the QuickStyle audit gap for the
  approved cohort (Test 45 covers the non-approved cohort).
- **LineSpacingRule prescriptive pass closed.** `CustomParaAfterH1`
  fixed to Single (rule 0, LineSpacing 12); `Footnote Text`
  retained at Exactly 8 as a known exception (i18n-flagged).
  Taxonomy in `basTEST_aeBibleConfig.bas` updated.
- **Test 22 split into 22 / 38 / 74.** The bundled
  `CountEmptyParagraphsWithFormatting` was decomposed into three
  disjoint detectors after a step-back analysis surfaced that
  Test 38 already covered the bare-empty case as a subset:
  - Test 38 (`CountEmptyParagraphs`) - bare empty, expected
    rebaselined 153 -> 216.
  - Test 22 (`CountWhitespacePaddedEmptyParagraphs`) - spaces
    around the pilcrow only, expected 0.
  - Test 74 (`CountEmptyParagraphsWithInlineContent`, new) -
    integrity check for visually-empty paragraphs carrying
    InlineShapes / Fields / Bookmarks, expected 0.
  Combined runtime 7.6s vs prior single-test 390s.
- **Slot 5 retire + Slot 6 upgrade.** Empty-paragraph discipline
  tightened via path-(b) paragraph walk with three break-exception
  predicates. Operator-verification trail confirmed
  `wdActiveEndAdjustedPageNumber` is the right page index for
  hint emission.

Status tag legend (continued):

- **OPEN** - actively pending, all known prerequisites met.
- **PARTIAL** - partially complete; specific remaining work listed.
- **DEFERRED** - not started, waiting on a specific trigger.
- **FUTURE** - speculative; revisit only when conditions warrant.
- **RECOVERED** - surfaced from a prior arc where it was dropped
  off the carry-forward chain.

## Open carry-forward (priority order)

### 1. Run aeRibbon Gates G1-G8 and ship v1.0.0 (HIGH) - OPEN

The production export gateway is in place; nothing has been built
or gated yet. **Next active release-track item** and gates the
hand-off to the author for comments-only review.

Original analysis and gate definitions: see prior arcs via the
2026-05-15 carry-forward.

### 2. Item 13 remaining work - built-in hide-sweep + test wiring (MEDIUM) - CLOSED 2026-05-17

Pass 1 closed 2026-05-14 (`AuditNonPaletteStyleColors` added,
custom-style anomaly count brought to 0).

**2.1 Hide-sweep wired into `WordEditingConfig` 2026-05-17.**
Investigation found that `HideUnapprovedBuiltInStyles`
(`basStyleInspector.bas:1459`) was already implemented and
correct - three-property pattern (`Priority = 99`,
`QuickStyle = False`, `UnhideWhenUsed = False`), idempotent,
handles locked styles. The missing piece was operator
ergonomics: the sweep relied on memory each run.

**Operator-side verification (2026-05-17):**
- `HideUnapprovedBuiltInStyles` run: 1 newly hidden
  (`Default Paragraph Font`), 115 already hidden, 251 skipped
  (locked).
- Test 45 (`CountUnapprovedVisibleStyles`): PASS at 0.
- `Hyperlink` direct check: `99 / False / False` - properly
  hidden post-BookHyperlink refactor.
- `Followed Hyperlink`: not instantiated in the document
  (`Styles("Followed Hyperlink")` raises 5941). Word lazy-
  instantiates this built-in; if and when it materializes, the
  next `WordEditingConfig` run will hide it.

**Lock-in change.** `WordEditingConfig` now calls
`HideUnapprovedBuiltInStyles` between `PromoteApprovedStyles`
and `DumpPrioritiesSorted`. Placement after `PromoteApprovedStyles`
means the approved set is well-defined before the sweep runs;
both are idempotent so re-runs are safe.

**Follow-up (logged, not blocking):** the 251 "skipped (locked)"
count exceeds the document's ~176-style population. Likely
cause: `Priority = 99` reorders the Styles collection mid-`For
Each`, re-visiting some entries. Doesn't affect correctness
(Test 45 verifies the end state) but the report's count
arithmetic is misleading. Open as § 11.

Full prior analysis: see § 2 in
[`Code_review 2026-05-15.md`](Code_review%202026-05-15.md).

### 3. TestReport.txt - per-slot one-line descriptions (MEDIUM) - CLOSED 2026-05-17

All 74 slots populated in `GetTestDescription` across nine
authoring batches (A-I). Every PASS/FAIL row in
`rpt/TestReport.txt` now carries an inline rule statement or
baseline note.

**Authoring conventions adopted during the pass:**

- Rule-style descriptions (expected = 0) phrased as
  "Rule: <statement> - <Function> should return 0".
- Baseline-style descriptions (expected > 0) phrased as
  "<what it counts> - <why the baseline is what it is> -
  <Function> should match the expected baseline". Hard-coded
  baseline numbers were intentionally removed from description
  text after an operator rebaseline (index 16: 33783 -> 33822;
  index 28: 77 -> 82) made embedded numbers stale within hours.
  The expected value remains in the test report itself; the
  description explains intent only.
- Contraction tests (52-65) and Unicode-sequence tests (66-71)
  each get a per-slot description naming the specific
  contraction or codepoint sequence, rather than a shared
  generic line - the specificity is the signal.

**Side findings logged during authoring:**

- **Test 30 source-comment vs expected mismatch.**
  `CountHeaderStyleUsage` function comment says "Expected = 0"
  but the `values` array baseline is 70. Either the rule was
  relaxed and the comment is stale, or the expected is wrong.
  Description notes the stale comment; full reconciliation
  deferred. Tracked as part of § 12.
- **Test 47 baseline correction.** Earlier batch C mistakenly
  applied "baseline 147" to slot 47; the actual array value at
  position 47 is 3 (position 50 holds 147). Corrected during
  batch E.

Full design rationale and emission shape: see § 10 in
[`Code_review 2026-05-15.md`](Code_review%202026-05-15.md).

### 12. Revisit failed tests and verify status / code / performance (MEDIUM) - OPEN 2026-05-17

Surfaced repeatedly during the § 3 authoring pass. Several
slots merit individual revisit beyond what description-writing
addresses:

- **Status reconciliation.** Where a test's source comment or
  prior documentation disagrees with the current expected
  value (e.g. Test 30: comment says expected 0, array says 70),
  decide whether the rule was relaxed (update comment) or the
  expected drifted (rebaseline or restore). Same audit should
  catch any other slot whose intent has quietly diverged from
  its baseline.
- **Code review per failing slot.** For any slot that returns
  FAIL on a current run, walk the function once before
  rebaselining - the FAIL may be a real regression masked by a
  habit of "just bump the expected." Especially relevant for
  count-baseline tests (24, 27, 29, 30, 32-35, 37, 47, 49, 50,
  51) where drift could be either editorial or accidental.
- **Performance.** Several slots are slow (Test 22 was 390s
  before this session's perf work; others may have similar
  Range.Text materialization patterns or per-paragraph
  scoped-collection access - see § 11 for the
  HideUnapprovedBuiltInStyles example). Revisit slow slots
  with the same lens used on Test 22: cheap length guard,
  iterate the small collection rather than the large one,
  avoid Range.Text when Len(Range.Text) will do.

**Trigger:** next time a slot FAILs, do this revisit before
adjusting the expected. Over time the per-slot revisits will
also surface candidates for the same kind of split applied to
Test 22 (one bundled signal -> two or three disjoint ones).

### 4. Taxonomy audit - full-coverage final-state goal (LOW-MEDIUM, ASPIRATIONAL) - CLOSED 2026-05-17 (state check)

State-check pass complete. Recounted directly from
`RUN_TAXONOMY_STYLES` in `basTEST_aeBibleConfig.bas`:

| Bucket | Count |
|---|---:|
| Fully specified | 47 |
| Existence-verified | 0 |
| Not yet created (expected FAIL) | 3 |
| Tab-stop audits | 9 |
| **Total distinct AuditOneStyle entries** | **50** |
| **Total checks (AuditOneStyle + AuditStyleTabs)** | **59** |

50 of 52 `GetApprovedStyles()` entries have a taxonomy entry.
Gaps: `Normal` (anchor, intentionally unaudited - see § 17) and
`FargleBlargle` (canary, deliberately missing from document).

**Docstring drift:** the `RUN_TAXONOMY_STYLES` header comment
says 49 styles / 46 fully specified / 58 checks. Actuals are
50 / 47 / 59. BookHyperlink (added 2026-05-15) was not
propagated to the docstring. Tracked as part of § 17.

**Loophole analysis surfaced 4 high/medium-value follow-ups** -
opened as § 13-16. A fifth bundle (§ 17) collects the
lower-value items not worth elevating individually.

Original umbrella per-style decisions list: see § 4 in
[`Code_review 2026-05-15.md`](Code_review%202026-05-15.md).

### 5. EDSG documentation refresh - VerseText-aware (LOW) - CLOSED 2026-05-17

Investigation surfaced wider drift than the original carry-forward
summary suggested: `GetApprovedStyles()` had grown from ~44 to 52
entries, with VerseText at priority 33 (not 31 as the prior
summary said), plus `BibleIndexList`, `AuthorBookSections`,
`BookHyperlink`, `ParallelHeader`, `ParallelText`,
`SpeakerLabel`, `BodyTextContinuation`, `AppendixTitle`,
`AppendixBody` all added; `BodyTextIndent`, `AuthorQuote`,
`BookIntro` removed.

**Refreshed (one-shot rebuild from the SSOT in
`basTEST_aeBibleConfig.bas`):**

- `EDSG/01-styles.md` — priority table rebuilt as a single
  unified 52-row table with notes column; "Pending re-validation"
  / "Reserved gaps" framing retired (gaps were filled); category
  prose updated (`Body text`, `Author commentary`, `Anchor`,
  `Special book treatments`).
- `EDSG/04-qa-workflow.md` — "Current state" rewritten and dated
  2026-05-17; reflects 52-entry array, lists additions and
  removals since the 2026-04-26 snapshot.
- `EDSG/06-i18n.md` — `VerseText` added as the primary
  translation target; `BodyTextIndent` line removed.

Mechanical-dump script (option 2 from the investigation report)
was not built; reassess if the array starts churning again.

### 6. EDSG/02-editing-process.md - AuthorListItem* as canonical BaseStyle="" example (LOW) - WONTFIX 2026-05-17

Closed as obsolete after verification. `BaseStyle = ""` is a
universal QA-checklist rule for every approved paragraph style
(rationale: the Word List Paragraph numbering-engine hang in
large documents; bug is alive and Microsoft will almost
certainly never fix it - see `EDSG/10-list-paragraph-bug.md`).
A "canonical example" framing would misleadingly suggest the
rule is *especially* about list-item styles.

`10-list-paragraph-bug.md` already uses `AuthorListItem` as the
running worked example through its five-step migration recipe
(Step 0 diagnostic targets, Step 1 template, Steps 2-5
migration). `02-editing-process.md` correctly defers via
cross-reference. Adding a duplicate callout would split the
worked example across two pages and weaken the single source
of truth.

### 7. Finding 5 (ribbon nav) - umbrella OPEN (DEFERRED, WORD LIMITATION) - RECOVERED

Word-side limitation; no action available. Remains in the
register for awareness.

### 8. SoftHyphenSweep_FootnotesOnly sister routine (DEFERRED)

Surfaced during the 2026-05-08 SHA build; waiting on a footnote-
specific trigger before implementation.

### 9. SHA_ReplaceHard i18n consideration (FUTURE)

Speculative; revisit when a non-English target translation
materialises.

### 10. Architecture rule - class encapsulation + module/class as casual-coder safety boundary (RULE, 2026-05-15)

Codified as a feedback memory and documented in the 2026-05-15
arc. Standing rule, not an action item - listed here so it
remains visible during slot-by-slot review work.

Full rule and worked examples: see § 9 in
[`Code_review 2026-05-15.md`](Code_review%202026-05-15.md).

### 11. HideUnapprovedBuiltInStyles - skipped-count arithmetic (LOW) - CLOSED 2026-05-17 (false alarm)

**Original concern:** `skipped (locked): 251` on a document
believed to have ~176 styles - hypothesized re-visit inflation
from `Priority = 99` reordering the Styles collection
mid-`For Each`.

**Investigation finding:** the arithmetic was always correct.
Both the original report (1 + 115 + 251) and the post-fix
report (0 + 116 + 251) sum to **367 = candidates**, with no
duplication. The "~176 styles" reference came from
`AuditNonPaletteStyleColors`, which deliberately excludes
`wdStyleTypeTable` and `wdStyleTypeList`. The hide-sweep
includes both. Word ships dozens of built-in table styles, list
styles, and locale variants - 367 distinct built-in
non-approved names in the collection is genuine, and 251 of
them are locked against property writes (Word locks
table/list/system styles by design).

**Code change kept anyway.** The function now snapshots target
names into a dictionary before mutating, then iterates the
snapshot. Not needed for correctness, but adds two defensive
properties:

- Proves one-visit-per-name structurally rather than by
  inspection.
- Emits a new `candidates: N` line in the report, which makes
  the sum check (`candidates == newly + already + skipped`)
  trivially visible.

Originated 2026-05-17 during the § 2.1 hide-sweep close;
disproved 2026-05-17 by the post-fix run reproducing the same
totals.

### 13. Audit BaseStyle = "" and LinkToListTemplate on approved styles (HIGH) - CLOSED 2026-05-17

**Implemented as Test 75:**
`CountApprovedStylesWithListParagraphRisk` in `aeBibleClass.cls`.
Walks `GetApprovedStyles()`, and for each existing **paragraph**
style asserts:

- `BaseStyle = ""` (read directly).
- No `ListTemplate.ListLevels(n).LinkedStyle` matches the
  style's `NameLocal`. Pre-builds a dictionary by iterating
  `ActiveDocument.ListTemplates` once, then checks each approved
  style's name against it.

Character styles are skipped (they legitimately base on
`Default Paragraph Font`). Each property failure counts
separately, so a style failing both reports 2. Expected 0.

**Implementation note - LinkToListTemplate is write-only.**
First draft used `Not (s.LinkToListTemplate Is Nothing)`
following the example in `EDSG/10-list-paragraph-bug.md` Step 0.
That raises *"argument not optional"* at compile time -
`Style.LinkToListTemplate` is a write-side method that takes a
`ListTemplate` argument and cannot be read back as a property.
The correct read-side detection uses the
`ListTemplates -> ListLevels -> LinkedStyle` graph. The EDSG
Step 0 snippet is conceptual rather than working code; worth
fixing on a future EDSG pass.

**Verification:** Test 75 PASS at 0 in 0.03s on the live
document. `MaxTests` bumped 74 -> 75, `values` array extended
with `0`, all four Case-dispatch sites wired.

**Coverage closed:** the silent-failure mode that put built-in
`Hyperlink` back in the gallery (before the BookHyperlink work)
and the multi-hour-hang risk from accidental list-template
attachment are now both gated by automated tests on every
`RUN_THE_TESTS` run.

### 14. Audit AutomaticallyUpdate = False on approved styles (MEDIUM) - CLOSED 2026-05-17

**Implemented as Test 76:**
`CountApprovedStylesWithAutoUpdateOn` in `aeBibleClass.cls`.
Walks `GetApprovedStyles()` and flags any paragraph style whose
`AutomaticallyUpdate` is True. Character styles skipped (the
property is paragraph-scope in Word's object model). Returns
total violations; expected 0.

**Design choice - separate test rather than folding into Test 75.**
Considered piggybacking on the LP-bug walk for one combined
property audit, but kept the tests disjoint: Test 75 is the
LP-hang guard (specific failure mode + specific EDSG page);
Test 76 is the silent-drift guard (different mechanism, different
recovery path). A combined test would muddy the FAIL signal.

**Verification:** Test 76 PASS at 0 in 0.02s. `MaxTests` bumped
75 -> 76, `values` array extended with `0`, all four
Case-dispatch sites wired.

### 15. Audit UnhideWhenUsed = False on approved cohort (MEDIUM) - OPEN 2026-05-17

Surfaced by the § 4 taxonomy loophole analysis. Test 45
(`CountUnapprovedVisibleStyles`) verifies this property only on
the **non-approved** cohort. The approved cohort is uncovered.

This is exactly the silent re-surface mode that put built-in
`Hyperlink` back in the gallery before the 2026-05-15
BookHyperlink work - and the same risk applies to any approved
paragraph style modified via the dialog (Word can flip
UnhideWhenUsed back to True on certain operations).

Same shape as § 13/14; trivially folds into the same combined
audit walk.

### 16. Decide on three persistently-missing placeholders (LOW) - OPEN 2026-05-17

`BodyTextContinuation`, `AppendixTitle`, `AppendixBody` have
been in the `not yet created (expected FAIL)` bucket since
2026-05-07 (10+ days). `RUN_TAXONOMY_STYLES` reports three
FAILs every run as a result.

**Cost:** desensitisation. An operator who sees three FAILs on
every run learns to skip the FAIL summary, which masks any real
FAIL that later joins them.

**Decision needed per placeholder:**
- Define + populate via a `Define*` routine + promote to bucket
  1, or
- Remove from `GetApprovedStyles()` and from the not-yet-created
  taxonomy bucket.

Either way the bucket goes empty and the false-FAIL noise stops.

### 17. Lower-value taxonomy hardening (LOW, bundled) - OPEN 2026-05-17

Four small items surfaced by § 4 that aren't worth individual
follow-ups but should be tracked together to avoid loss:

- **Normal style** - intentionally unaudited (anchor). Worth at
  least an existence-verified entry with font/color pinned so
  drift in the underlying Normal style is detected.
- **Tab-stop count contract** - `AuditStyleTabs` checks expected
  stops exist but doesn't flag *extra* stops. A fourth stop
  added to `AuthorListItemTab` would pass silently. Consider
  asserting `Style.ParagraphFormat.TabStops.Count = expected`.
- **Priority drift** - `PromoteApprovedStyles` assigns priorities
  but nothing verifies they stay assigned. A user-drag in the
  gallery or `CopyStylesFromTemplate` import could reorder
  silently.
- **RUN_TAXONOMY_STYLES docstring count drift** - hand-maintained
  totals (49 / 46 / 58) already stale by 1 / 1 / 1 within 2 days
  of BookHyperlink. Either remove the literal counts from the
  docstring or auto-emit them.

## Pointer back to the closed arc

Full dated history of the work that produced this carry-forward
state is in [`rvw/Code_review 2026-05-15.md`](Code_review%202026-05-15.md).
That file (and the arcs it points back to) covers:

- The BookHyperlink design, implementation, and verification.
- The Test 5 / Test 22 split / Test 74 add sequence.
- The Slot 5 retire + Slot 6 upgrade arc with operator-verification
  trail.
- The LineSpacingRule prescriptive pass closure.
