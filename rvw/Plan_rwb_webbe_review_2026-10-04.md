# R16 plan: WEBBE ledger/review pass - 2026-10-04

**This is a plan document only - nothing here is built or executed yet.**
Per the operator's explicit framing: capture the plan now, implement
later. Follow-on to R15 (`Plan_rwb_webu_review_checklist_2026-09-28.md`),
which closed WEBU's `other` bucket to 0 pending/1,092 accepted on
2026-10-03/04 (Addendum 13). This plan covers the next body of work -
the WEBBE (World English Bible, British Edition) ledger/review pass -
plus two related but separable concerns the operator raised alongside it:
spelling-substitution timing, and VBA work needed for a UK-targeted print
layout (left-aligned, no hyphenation).

## 0. Baseline, confirmed live (2026-10-04)

- `aeRWB`'s R14 categorizer (`npm run docm.categorize`) already supports
  WEBBE and already has a working US/UK spelling-variant classifier rule
  (`docm-webu-webbe-categorize.mjs`'s `isSpellingVariantHunk`,
  `EXACT_SPELLING_PAIRS` + `SPELLING_SUFFIX_RULES`, built 2026-09-26) -
  this already tags and excludes pure spelling-only hunks from the WEBBE
  `other` bucket, same mechanism as every other classifier rule
  (divine-names, formatting, kjv-word-choice, etc.).
- Fresh run, 2026-10-04: **WEBBE `other` 1,550 of 14,169 changed verses
  (10.9%)**, no `rwb-webbe-review.txt`/`rwb-webbe-accepted.txt` exist yet
  - the full, untouched backlog. (WEBU, for comparison, is fully closed:
    `operator-accepted` 1,091 of 12,123, 0 `other`.)
- R15's `generate-review-checklist.mjs` already has a working `webbe`
  target wired (`npm run docm.review-checklist -- webbe`) - built in the
  original R15 design, never exercised. No new tool-building is needed to
  start; this is a continuation of existing infrastructure, not a new
  build.

## 1. docm fixes already carry over to WEBBE for free - no separate action needed

Every real defect fixed during the WEBU review (Rules 2/3/4/7 sweeps,
R15's ~50+ verse-level fixes across all sessions, including this
session's Revelation 3:2/3:9/16:16/21:9 and the 2 Corinthians 11:2 typo)
was a fix **to `docm-verses.txt`/the docm itself** - not to a
WEBU-specific copy. `docm-webu-webbe-categorize.mjs` reads the same
`docm-verses.txt` for both the WEBU and WEBBE comparisons. This means the
1,550-verse WEBBE `other` bucket above **already reflects every WEBU-era
fix** - there is no backlog of "WEBU fixes that still need porting to
WEBBE." The only work the WEBBE pass adds is reviewing the verses where
docm diverges from WEBBE specifically (mostly British vocabulary/spelling
differences beyond the classifier's current pair list, plus whatever
WEBBE-only judgment calls turn up, analogous to WEBU's 21).

## 2. Spelling-substitution prep: timing question, not a build question

Two separate things exist today under this name, and the "when" question
is really about which one runs where in the pipeline, not whether either
needs building (both already exist):

- **Comparison-time classification** (`aeRWB`, read-only): already live,
  see §0. Confirmed 2026-09-26 that `amongst`->`among` alone accounted for
  937 hunks - by far the largest single pattern in the WEBBE diff. This
  already runs *before* a human ever sees the WEBBE review checklist (it's
  baked into `classifyVerse`), so the review-checklist pass inherits this
  for free - the pending list should NOT be swamped with pure
  US/UK-spelling-only verses when R16 actually starts.
- **Production-time mutation** (`aeBibleAddin`, `applySpellingVariant` +
  `spelling-variants.ts`): built and tested 2026-09-18, deliberately never
  wired into a live document-mutation feature - this is the item the
  operator flagged 2026-09-25 as **required, high priority, explicitly
  deferred**, rationale still not supplied. This is the thing that would
  actually rewrite "color" -> "colour" etc. in a real UK-targeted
  docx/docm for print.

**Open sequencing question for the operator to resolve when
implementation starts** (not resolved here): does production-time
spelling substitution run *before* the WEBBE content-review pass (so
reviewers are looking at already-UK-spelled text side by side with
WEBBE) or *after* it (review docm's US-spelled text against WEBBE first,
confirming content correctness, then apply the spelling transform once as
a late pre-print step on a working copy)? Running it after is the lower-
risk default - it keeps the content-correctness review and the
spelling-locale transform as two independently-verifiable steps, and
avoids re-reviewing content that spelling substitution might
accidentally disturb (e.g. a substitution rule misfiring inside a proper
noun or an already-decided judgment call's wording). No decision made
here; flagging the question is the deliverable.

## 3. Hyphenation is out of scope for the WEBBE *review* pass

The operator confirmed the docm has substantial hyphenation already
applied through its print layout (observed at p170 of ~900 pages) via
Word's optional-hyphen character (`Chr(31)`, "soft hyphen") - existing
VBA tooling already handles this at the print-layout level
(`basWordRepairRunner.bas`'s `SoftHyphen_CalibrateColumns` /
`SoftHyphenSweep_ByColumnContext_SinglePage`, page-by-page, Active/Stray/
OutsideBody classification). This is a **print-layout concern**, entirely
separate from the WEBBE content-review pass: `ExportDocmVersesToRWBFormat`
exports plain verse text for comparison, and hyphenation is a rendering
artifact of the two-column print layout, not part of the exported text
`docm-verses.txt`/the WEBBE diff tooling ever sees. **Confirmed: no
hyphenation-related noise reaches the review-checklist tool at all** -
this is a reason to NOT worry about it when R16 starts, not a blocker.

## 4. Priority / timeline: not urgent

RWB's US edition prints before the UK edition - the operator stated
there's no deadline pressure driving WEBBE's review pass. This plan sits
behind the existing backlog (see §6c of the WEBU plan's own carried-over
list: `applySpellingVariant` live-check, the scholarly-grounding Genesis 1
spike, the `Jeremiah 37:910` malformed-reference bug, the 2 Samuel 21:8
confirmation) at the operator's discretion, not a fixed position - this
plan does not assert WEBBE must happen next, only that it's the next
*candidate* body of work in the R15/R16 review-tool lineage specifically.

## 5. VBA work needed for UK/i18n print layout (separate, later, print-side task)

Confirmed from `rpt/Styles/style_VerseText.txt`: `VerseText`'s
`ParagraphFormat.Alignment = 3` (`wdAlignParagraphJustify`) - the US
print convention, which is *why* hyphenation is "substantial" (justified
text needs hyphenation to avoid ragged word-spacing). A UK/British-print
convention of left-aligned, non-hyphenated `VerseText` needs:

1. **Alignment**: change `VerseText`'s `ParagraphFormat.Alignment` from
   justify (3) to left (0). This is a single style-definition edit *if*
   applied to a UK-specific working copy of the docm - **not** to the
   shared master style, since the master still needs to serve the
   US edition's justified-print requirement. This is the single most
   important architectural point for implementation time: **the US and
   UK print layouts cannot both be served by one shared `VerseText`
   style definition** - a UK build needs either a parallel style, a
   build-time style override applied to a working copy, or a separate
   per-locale document, analogous to how `ExportDocmVersesToRWBFormat`
   already treats "export a working copy, don't mutate the master" as
   the standing pattern for RWB-vs-docm work generally.
2. **Hyphenation removal**: this is NOT a simple document-level
   `AutoHyphenation = False` toggle - the operator confirmed substantial
   hyphenation is *already baked into the text* as literal `Chr(31)`
   optional-hyphen characters (not just a live auto-hyphenation setting
   that would stop applying if toggled off). Removing it means a bulk
   sweep across ~900 pages. **This is not starting from scratch**:
   `basWordRepairRunner.bas` already has working, page-scoped machinery
   for exactly this character (`SoftHyphenSweep_ByColumnContext_SinglePage`,
   with Active/Stray/OutsideBody classification and a prompt-each safety
   mode) - built for a different purpose (repairing misplaced hyphens in
   the two-column US layout) but directly reusable as the basis for a
   "remove every optional hyphen, unconditionally" UK-targeted variant,
   rather than a new tool.
3. Both (1) and (2) only need to happen on a UK-targeted **working copy**
   of the docm, never the shared master - consistent with this project's
   established convention (same posture as `ExportDocmVersesToRWBFormat`/
   the review-checklist tool, which only ever read from the docm, never
   mutate it in place for a locale-specific purpose).

**Not scoped in this plan**: which mechanism produces that working copy
(a one-time `SaveAs` + script pass? a maintained parallel UK `.docm`? a
build step triggered at print time?) - that's an implementation-time
decision, deliberately left open here.

## 6. Pros / cons / benefits / risks

**Pros / benefits of starting WEBBE's review now (whenever "now" is):**
- Reuses 100% of R15's tooling and workflow unchanged - no new design
  risk, same two-file pending/ledger split, same classifier, same
  "present/wait-for-decision" human workflow already proven across 5+
  sessions and ~1,100+ verses.
- The spelling-variant classifier rule (§0/§2) means WEBBE's `other`
  bucket is *already* filtered of the single largest predictable noise
  source before a human ever looks at it - the pass should feel more
  like WEBU's later sessions (judgment calls) than WEBU's very first
  session (mass pattern-mining).
- Every WEBU-era docm fix already carries over for free (§1) - the
  starting-point `other` count (1,550) is already lower than it would
  have been had WEBU's fixes not already landed in the shared docm.

**Cons / risks:**
- Still a large first pass - 1,550 pending verses is closer to WEBU's
  very first R15 session (1,226 pending) than to a routine continuation;
  expect multiple sessions, not one.
- New risk class WEBU's review never had to deal with: a WEBBE-side
  "defect" might actually be a **spelling-variant classifier gap**
  (a real UK/US word pair the classifier's `EXACT_SPELLING_PAIRS`/
  `SPELLING_SUFFIX_RULES` tables don't yet recognize) rather than a
  genuine docm defect or a translation judgment call - a third category
  WEBU's review never needed to distinguish. Recommend checking any
  single-word-swap "defect" candidate against the spelling-variant
  tables first, before concluding it's a real docm issue (mirrors the
  "check KJV before concluding a docm/WEBU divergence is a defect"
  lesson from the WEBU review, same shape of mistake, different source).
- The spelling-substitution sequencing question (§2) and the VBA
  hyphenation/alignment work (§5) are both genuinely unresolved
  dependencies for the *eventual UK print edition*, but - per §3 - neither
  blocks the WEBBE *review* pass itself. Risk is conflating "WEBBE review
  can start" with "WEBBE print edition is ready," which are different
  milestones on different timelines.
- No deadline pressure (§4) is also a soft risk: without urgency, this
  pass could stall indefinitely between sessions the way several other
  "not blocking" backlog items already have (scholarly-grounding spike,
  ASV work-list item) - worth deciding explicitly when to prioritize it
  over those, rather than by default inertia.

## 7. Development-timeline note (recorded here, mirrored into memory)

- **RWB US print edition precedes the UK edition** - this is the first
  time that ordering has been stated explicitly anywhere in this
  project's docs/memory. It reframes WEBBE's review pass and the UK
  print-layout VBA work (§5) as **pre-UK-print, not pre-US-print**
  dependencies - neither blocks the US edition's own path to print.
- Three genuinely separate milestones now exist on this timeline, not one
  "WEBBE work" blob: (a) WEBBE **content review** (this plan's main
  subject, tooling-ready, not urgent), (b) **spelling-substitution
  production wiring** (`applySpellingVariant`, already built, blocked on
  operator rationale + a live-check, §2), (c) **UK print-layout VBA**
  (alignment + hyphenation removal, §5, not started, needs its own
  working-copy mechanism decided first). All three are UK-print
  prerequisites; none are on the critical path to the US print edition.
