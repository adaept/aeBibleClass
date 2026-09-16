# Plan - Phase 4: `rwb.txt` punctuation sync & content audit - 2026-09-15

Follow-on to `rvw/Plan_engwebu_baseline_sync_2026-09-14.md` (Tests 70/71,
now both closed - see that doc for full history/tooling). That doc's Phase 4
("Sync `rwb.txt` content") and its "New goal - generate `rwb.txt` from the
docm" section are the design context this plan executes against. Structure
below follows the operator's 6-item outline (2026-09-15), each expanded with
scope, data already gathered, and open questions - status tracked with the
usual legend (✅🟡🔴⚪).

## 0. Baseline data (gathered before writing this plan, not yet acted on)

Current census totals (`aeRWB`, `npm run web.census`), for context on how
much work each pass actually involves:

| Pattern | web.txt | rwb.txt | WEBU | docm (now, post Test 70/71) |
|---|---:|---:|---:|---:|
| Test 70 (open triplet) | 0 | **0** | 75 | 75 |
| Test 71 (close triplet) | 35 | **35** | 63 | 63 |

Notable: for Test 71, `rwb.txt`'s 35 hits have an **identical per-book
breakdown** to `web.txt`'s 35 (Genesis 1, Exodus 4, Numbers 1, 1 Samuel 1,
2 Samuel 3, 1 Kings 4, 2 Kings 8, 1 Chronicles 2, 2 Chronicles 3, Isaiah 3,
Ezekiel 1, Zechariah 2, Matthew 2) - strong evidence `rwb.txt` inherited
these unchanged from the 2013 WEB baseline at these exact verses, not
independently edited. For Test 70, `rwb.txt` has **zero** - the pattern
doesn't exist there at all yet.

## ✅ 1. First pass - bring `rwb.txt` up to the docm's (now WEBU-matching) Test 70 punctuation - Done 2026-09-15, aeRWB `e42be6d`

**Scope:** all 75 Test 70 verses. Since `rwb.txt` currently has 0 hits, this
is not a small patch - every one of the 75 verses likely needs the docm's
opening-triplet punctuation ported in (pending per-verse confirmation, not
assumed).

**Tooling gap to close first:** R3's existing diff register compares
`web.txt` vs `rwb.txt`. This pass needs a **`docm-verses.txt` vs `rwb.txt`**
diff instead - a new mode (or new register) in `aeRWB/tools/web-diff`,
reusing `parseBible`/`diffBibles`/`wordDiff` unchanged (they're already
source-agnostic - see `lib.mjs`'s design). Scope the diff to just the 75
Test 70 refs (from `census/201C-2018-201C-worklist.md`'s docm hit list),
not the whole Bible, to keep the review reviewable.

**Process (per R3/R4's existing discipline - reviewed, not bulk):**
1. Generate the scoped docm-vs-rwb diff for the 75 refs.
2. For each: confirm whether `rwb.txt`'s wording differs from the docm's
   (expected - RWB has its own edits) and, if so, port **only the
   punctuation**, not the wording, into `rwb.txt` - mirroring the "leave
   RWB's own choices alone" rule already used throughout Test 70/71 editing.
3. Re-run the Test 70 census against `rwb.txt` specifically to confirm 75/75
   after the pass.

**Suggest a spike first** (per the confirmed-good pattern from Test 70/71):
1-2 verses through this new diff-and-patch pipeline before scaling to 75,
same reasoning as before - prove the mechanism, not just the data.

**Result:** the "port only the punctuation" framing above turned out to be
wrong once actual data was checked - **0 of the 75 verses differed on
punctuation alone**; every one had real wording drift too (`rwb.txt` reads
like a stale, never-updated layer against the actively-edited docm - e.g.
"shall"/"will", un-decontracted phrasing, restructured clauses). Operator
decision: **full verse-text replacement** from the docm for all 75 refs,
not a punctuation-only patch - this also matches the plan's own "generate
`rwb.txt` from the docm" goal more directly than surgical patching would
have. Built two new `aeRWB` tools instead of one: `docm-rwb-diff.mjs`
(read-only scoped diff, R8) and `apply-docm-rwb-sync.mjs` (the only tool
that writes `rwb.txt`). Caught and fixed a real bug before finalizing: the
docm export's leading character after each tab is `U+202F` (narrow no-break
space, a VBA-export artifact), not a regular space - an early version of
the sync script missed this and would have written the stray character
into all 75 `rwb.txt` lines; caught by spot-checking byte-level output,
reverted, fixed, redone. Verified: `docm-rwb-diff` 75/75 identical, `aeRWB`
census 75/75 matching WEBU, 24/24 unit tests pass, file line-count/encoding
unchanged.

**Confirmed/denied (operator's framing of the benefit, 2026-09-15):**
confirmed this aligns `rwb.txt` (not docm, which was already aligned) more
closely with WEBU, and confirmed `web.txt` stops being an active editing
reference; denied that `web.txt` itself "sunsets" - it remains the
permanent baseline for R3's diff-register provenance/transparency mechanism
regardless of how `rwb.txt` is produced. See the conversation record for
the full reasoning; not duplicated here.

## ⚪ 2. Second pass - bring `rwb.txt` up to the docm's Test 71 punctuation

**Scope:** the 63 Test 71 verses, but **not from zero** this time - 35 of
`rwb.txt`'s current hits are presumably already correct (inherited from
`web.txt`, and `web.txt`'s own 35 already happen to match WEBU's convention
at those verses - unconfirmed, needs the same per-verse diff to verify
rather than assumed). Likely a smaller true edit count than 63, but **must
be measured, not estimated** - a verse could coincidentally show a pattern
hit while still differing from the docm/WEBU in nesting depth (the same
"presence ≠ equality" blind spot documented in the Test 70/71 plan).

**Process:** same docm-vs-rwb diff tool from Pass 1, scoped to the 63 Test
71 refs. Same spike-first approach.

## ⚪ 3. Third pass - Adonai/YHWH/Elohim → Lord/LORD/Lord GOD consistency

**Scope, per operator's prior research:** in the Hebrew source, the
Tetragrammaton (YHWH) is conventionally rendered `LORD` (all caps) in most
English Bibles, `Adonai` as `Lord`, and the combined `Adonai YHWH` as
`Lord GOD`. This pass checks whether **docm/RWB's actual usage is internally
consistent** with whichever convention RWB has adopted - not necessarily the
standard convention, since RWB already has an established practice
(observed repeatedly during Test 70/71 editing) of substituting `God` for
WEBU's `the LORD` in many verses.

**Open question this pass must answer first, before "consistency" is even
checkable:** what *is* RWB's actual intended rule? Candidates observed so
far, none yet confirmed as "the rule":
- Blanket `the LORD` → `God` substitution regardless of underlying Hebrew
  name (YHWH vs. Adonai vs. combined) - the simplest reading of what's been
  seen, but never explicitly confirmed as the intended design.
- A name-aware substitution that's supposed to preserve the YHWH/Adonai
  distinction sometimes, but has drifted/been applied inconsistently.

**This pass is therefore two steps, not one:**
1. **Define** the actual intended rule (ask the operator directly, or infer
   from a large enough sample of docm usage + any existing RWB style-guide
   documentation, if one exists).
2. **Verify** docm/WEBU/`rwb.txt` consistency against that defined rule,
   using the same census/diff tooling pattern (a name-usage census is a
   natural extension of the existing pattern-census tool - censusing
   `LORD`, `Lord`, `Lord GOD`, `God` occurrences per verse across all three
   sources rather than quote-mark triplets).

### 🟡 Related, deferred task (operator-flagged, not blocking Phase 4)

Compare Adonai/YHWH/Elohim → Lord/LORD/Lord GOD usage across **NIV, NKJV,
and WEBU** for cross-version consistency, framed as a "does this align with
the Study Bible" check. Explicitly **not** blocking the docm/RWB-internal
consistency check above - current focus stays on WEBU per existing project
scope. NIV/NKJV are copyrighted, modern translations - obtaining/using their
text for automated comparison needs its own licensing check before any
tooling is built against them (unlike WEB/WEBU's public-domain status,
which is *why* they were chosen as the baseline - see
`project_i18n_architecture_vision` memory / the architecture assessment in
the Test 70/71 plan doc). Record findings here when this is picked up; don't
start the comparison itself yet.

## ⚪ 4. Other comparisons on this baseline

Placeholder - no specific comparisons identified yet beyond items 1-3.
Add them here as they're identified, each as its own `⚪` line with scope,
rather than trying to enumerate them speculatively now.

## ⚪ 5. Additional considerations (running list - review and prioritize as added)

Space for anything that surfaces during 1-4 that isn't itself a comparison
pass but affects scope/priority/risk. Add dated entries here rather than
scattering notes elsewhere, so this plan doc stays the single place to check
"what have we learned since this was written."

- (none yet)

## ⚪ 6. Align with the future Strong's numbers implementation (Phase 5)

Phase 5 (in the Test 70/71 plan doc) will need to attach Strong's numbers
per verse, sourced from WEBU's `\w word|strong="H1234"\w*` tags, into a
future `rwb.usfm` export. Whatever tooling/process Phase 4 builds (the
docm-vs-rwb diff mode, the name-usage census) should be designed so Phase 5
doesn't have to redesign it:

- Keep everything **verse-keyed** (already the established pattern across
  R3/R6/R7 - `Map<"Book C:V", text>`) - Phase 5's word-level Strong's
  alignment will need to key off the same verse references without a
  translation layer.
- Don't discard provenance (SHA-256 input hashes, commit-hash snapshots)
  when building Phase 4's new tooling - Phase 5 will want the same
  "which exact docm/WEBU state was this built from" traceability that R3/R6/R7
  already provide.
- The Phase 4 "third pass" (LORD/Adonai) is itself relevant groundwork for
  Phase 5: Strong's-number alignment will need to know where `rwb.txt`'s
  wording (e.g. `God` for `the LORD`) diverges from WEBU's word-for-word
  text, since a naive positional word-aligner (flagged as a known risk in
  Phase 5's existing notes) would misattribute the Strong's number for
  `LORD` onto RWB's `God` otherwise.

## Suggested sequencing

1 → 2 (mechanical, same new tool, lower risk, spike-first) → 3 (needs a
definition step before it's even checkable) → 4/5 (open-ended, filled in
as work proceeds) → 6 is a standing design constraint across all of the
above, not a separate sequential step - check new tooling against it as
it's built, not after.

**Status: ⚪ all items not started.** This document is the plan only;
no execution has begun.
