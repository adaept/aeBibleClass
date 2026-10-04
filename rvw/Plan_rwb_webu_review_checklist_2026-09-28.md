# R15 plan: a reviewable rwb/webu (and later rwb/webbe) checklist file - 2026-09-28

Follow-on to the R14 categorizer work (`aeRWB/tools/web-diff/README.md`'s
"docm-vs-WEBU/WEBBE full change categorization" section;
`project_docm_categorization_and_nt_variant_policy` memory). That work
found the multi-hunk `other` bucket has hit diminishing returns for
classifier-rule batch-mining - 1,116 of 1,211 distinct leftover hunk
shapes occur exactly once (checked against `aeRWB` commit `7645e9e1`,
`aeBibleClass` commit `6c95c71`: WEBU `other` 1,454/12,201 = 11.9%, WEBBE
`other` 1,952/14,243 = 13.7%). What remains is genuine editorial
divergence needing individual human judgment, not more pattern-finding.
**Updated 2026-09-29 (see the six addenda below the Status section;
Addendum 6 also has a session-close checkpoint)**: six rounds of
follow-up work (`aeRWB` commits `f197665` through `b0d844b`;
`aeBibleClass` commit `f37ecc9`) brought this to **WEBU 1,226/12,200 =
10.0%, WEBBE 1,684/14,242 = 11.8%** (from the original 50.7%/61.7%) -
four rounds were classifier fixes (sourcing a real KJV text, Addendum 3,
was the single largest reduction of the whole effort); Addendum 5 was
different in kind - a genuine "WEBU is better" defect hunt (not a
classifier gap) that found and fixed 6 real docm defects (3 dehumanizing
pronoun downgrades, 3 outright typos); Addendum 6 closed out the
archaic-verb-form thread (lie/lay/laid). The plan itself (the actual
review-checklist tool, §§1-10) remains **not built** - see Addendum 6's
checkpoint for next-session priorities.

**This is a plan document only - nothing here is built yet.** Per the
operator's 10-point brief (2026-09-28), captured verbatim as the ten
numbered requirements below, then worked through with a concrete design,
a status table, and an explicit pros/cons/risks/suggestions section.

## Status

✅ Built 2026-09-29, then REVISED same day into a two-file split (pending +
permanent ledger) after real use - see Addendum 7 for the full session
account, not repeated here. All four §7 open questions confirmed as their
recommended defaults before building: inline `rwb[ ]`/`rwb[x]` checkbox
(later split into pending-file + ledger-file, still inline-checkbox
*within* each file), permanent acceptance semantics, git-tracked output,
WEBU-only first pass. Built as `aeRWB/tools/web-diff/
generate-review-checklist.mjs` (`npm run docm.review-checklist -- webu`),
with shared parsing/idempotency helpers in `lib.mjs`
(`parseChecklistFile`/`splitChecklistEntries`/`formatChecklistFile`,
covered by `lib.test.mjs`) and the "connected to the classifier" feedback
wired into `docm-webu-webbe-categorize.mjs` (an `operator-accepted` tally,
sourced from the ledger file, layered on top of - not replacing - the raw
`other` bucket).

**Session-close state (2026-10-01)**: WEBU `other` 862 pending + 275
accepted (2026-09-30 session start) → **722 pending + 400 accepted**
(12,146 changed verses, 5.9% pending) as of 2026-10-01's close - see
Addendum 9. WEBBE untouched (1,595/14,191 = 11.2%, no ledger started,
WEBU-first scope unchanged). All work committed and pushed in both
`aeRWB` and `aeBibleClass` - see Addendum 9 and the
`project_rwb_review_checklist_tool` memory for the full account, and
`sync/session_manifest.txt` for the cross-session handoff.

**✅ WEBU R15 fully closed, 2026-10-03/04 - see Addendum 13.** WEBU
`other` 21 pending (all deferred judgment calls, Addendum 12) → **0
pending, 1,092 accepted**. WEBBE still the full, untouched backlog
(1,560/14,177 = 11.0%) and is the next body of work.

## Addendum, 2026-09-28: a real classifier bug found via this plan's own worked example

**The operator's first review of this plan caught a real bug by noticing
the original §2 example looked wrong** - both illustrative verses used
(Genesis 2:5, Genesis 4:7) were, on inspection, already fully explained by
the classifier (not actually in the `other` bucket at all), which meant
they were misleading as "still needs review" examples. That observation
turned out to point at a genuine, confirmed classifier gap, not just a bad
example pick - worth recording here as part of this plan's own
development, not just fixed silently upstream.

**Investigation.** Searched the WEBU `other` bucket for any verse still
carrying a leftover bare `"not"` fragment (the signature of a relocated
contraction, e.g. `won't it be` -> `will it not be`) - found **24**. Root
causes, both in `aeRWB/tools/web-diff/docm-webu-webbe-categorize.mjs`:

1. **A structural limitation in the `reordering` fallback**: it required
   the *entire* verse's leftover del/ins word multiset to match exactly
   before crediting the relocation - so a relocated `not` sharing a verse
   with any other unrelated change (a real synonym swap, a second typo)
   blocked recognition entirely, for the whole verse. 23 of the 24 found
   verses hit this.
2. **A leading-punctuation gap** in `splitPunct`/`DIVINE_TOKEN_RE`/
   `EVIL_RE`/`DISASTER_RE` - parentheses weren't in the stripped-punctuation
   class (only quotes/comma/etc. were), so e.g. `"(isn't"` in
   Genesis 19:20 wasn't recognized as a contraction at all. 1 of the 24
   found verses hit this one specifically.

**Fix** (commit `f197665`, `aeRWB`): the reordering fallback now cancels
matching words one at a time between the leftover multisets instead of
requiring an exact whole-verse match; whatever remains uncancelled (if
anything) still correctly leaves the verse `other`, but is now tagged
`reordering` alongside `other` as a contributing reason, and - critically
- the verse's `unexplainedHunks` now surface only the genuinely
unexplained remainder instead of being buried under relocation noise.
Parentheses added to the punctuation classes.

**Result: only 1 of the 24 verses (Genesis 19:20) became fully explained.**
The other 23 correctly remain in `other` - they were never false
positives of "other", they have real additional content the reordering
noise was obscuring. WEBU `other` moved 1,454 -> 1,453, WEBBE 1,952 ->
1,951 (essentially flat at one decimal place). **This is expected, not a
disappointing result** - the fix's value is diagnostic accuracy (the
`other` bucket's remaining entries now show their TRUE unexplained
content), which is exactly what this checklist tool needs to be trustworthy,
not bucket-size reduction for its own sake.

**A genuine, previously-unknown docm defect surfaced as a side effect**:
Micah 2:7 currently reads *"Don my words not do good to him who walks
blamelessly?"* - "Don" is not a word; almost certainly should read "Do"
(WEBU: *"Don't my words do good..."*, a rhetorical question). Not fixed as
part of this classifier work - flagged here for the operator to confirm
and fix in Word via the same one-at-a-time workflow used throughout the
R14 effort, separate from this plan.

**Why this matters for the R15 design above**: it's a concrete, real
demonstration that the review-checklist's `webu`/`rwb` side-by-side format
(§2) is doing real work even before any tooling exists - a human glancing
at two stacked lines caught something a purely mechanical classifier run
had missed. Reinforces that the checklist's job is to make genuinely
leftover content *legible*, not to chase the `other` percentage to zero by
any means.

## Addendum 2, 2026-09-29: the same worked example caught a second, bigger bug

**The operator's very next review of this plan - after the example was
corrected to `1 Kings 22:18` in the addendum above - immediately caught
another real bug in that same example**: `"evil?"` -> `"always bad?"`
wasn't recognized as RWB's documented theological softening pattern
(`project_rwb_editorial_philosophy`'s "avoids literal renderings like
'God does evil'"), because `isTheologicalHunk` only matched WEBU-side
"evil" being replaced by one of exactly three whitelisted words
(`disaster`/`calamity`/`trouble`) - "bad" wasn't on the list.

**Investigation** (per the operator's explicit ask to check the process,
not just patch the one instance): searched every hunk in the whole
docm-vs-WEBU diff (not just the current `other` bucket) for any hunk that
drops the bare word "evil" from WEBU's side, and grouped by what replaces
it. Found **49 distinct replacement words/phrases** - not 3:
`disaster` (58x), `harm` (10x), `bad` (10x), `calamity` (6x), `trouble`
(5x), then a long tail of context-specific rewordings (`wicked`, `wrong`,
`terrible`, `misfortune`, `destructive`, `tragedy`, `mischief`, even
Hebrews 10:22's "evil conscience" -> "guilty conscience"). **Operator
confirmed this is real and intentional, not overreach**: RWB avoids the
word "evil" wherever it isn't describing a person's own moral
culpability, deliberately aligning with the NIV over the general KJV
reading on this specific point ("God cannot do evil by definition") - a
genuinely broader application of the editorial-philosophy principle than
the 2-3 examples it was first documented with.

**Fix** (`aeRWB` commit `2c8f2c7`): dropped the replacement-word
whitelist entirely - `isTheologicalHunk` now fires on ANY hunk that drops
bare "evil" from WEBU's side, regardless of what (if anything) replaces
it. The removal of the word is the signal, not any particular
replacement. A follow-up fan-out scan (same methodology, checked against
every other WEBU word in the remaining `other` bucket, not just "evil")
found one more real instance of the same pattern - the plural **"evils"**
(-> disasters/calamities/troubles, 3 distinct replacements) had the exact
same gap for the same reason; fixed in commit `7796fc1`. Nothing else in
that broader scan showed the same signature - every other high-fan-out
WEBU word (`for`, `are`, `from`, `the`, `not`, `you`, `let`, `him`, `who`,
`like`, `then`, `that`, `them`, `their`) was ordinary function-word noise,
not a hidden systematic pattern.

**Result: WEBU `other` 1,453 -> 1,385 (-68 total across both fixes),
WEBBE 1,951 -> 1,886 (-65).** Unlike Addendum 1's reordering fix, this one
*did* move the bucket size substantially - the difference is that "evil"
avoidance is a genuinely single, well-defined signal (one word being
removed) rather than a structural detection limitation entangled with
many different individual pieces of real content.

**Process lesson, worth keeping for any future classifier rule**: a rule
that requires matching BOTH sides of a hunk against a small, hand-picked
word list is a warning sign when the underlying editorial principle is a
*removal*, not a *substitution into a closed vocabulary* - "is this word
being removed" is a far more robust signal than "is it being replaced by
one of these three synonyms I happened to sample." Worth checking any
future two-sided narrow-whitelist rule against this same question before
trusting its coverage.

## Addendum 3, 2026-09-29: a real KJV text source, and the largest single reduction yet

**The operator's next review of the (now Addendum-2-corrected) `1 Kings
22:18` example moved on to Genesis 3:15 (`"hostility"` -> `"enmity"`) and
asked something new: is this a real KJV word match, or another guess?**
Checking that question honestly required admitting a real gap - this
project had **no local KJV text**, so any "matches KJV" claim up to this
point (the on-in/will-shall/even-emphasis idiom categories) was grounded
in general register recognition, not verse-by-verse ground truth.

**The operator asked directly whether sourcing a real KJV text made
sense** - it did. Found and verified `https://eBible.org/Scriptures/
eng-kjv2006_usfm.zip` (Pure Cambridge Edition, 1769 text) - same site,
same USFM format, same UTF-8 encoding as the already-integrated WEBU/WEBBE
drops, downloaded by the operator into `aeBibleClass/eng-kjv2006_usfm/`
(gitignored, matching the other two). `aeRWB` commit `3ec0bf9` adds
`export-kjv.mjs`/`kjv-loader.mjs` (mirroring `export-webbe.mjs` exactly,
no new parsing logic) producing `kjv.txt` - 31,102 verses, matching the
docm's own known-good baseline count exactly.

**Rejected an earlier alternative** (`openbible.com/textfiles/kjv.txt`,
a plain-text dump from the same site `web.txt` already comes from) -
same simple format, but unconfirmed encoding and `[bracket]`-notation for
supplied words that would have needed new cleanup code. The eBible.org
USFM route needed zero new parsing logic and guaranteed UTF-8.

**Verification methodology** (not a blind "does this word appear
anywhere in the KJV" check - that was tried first and produced too much
noise): for every remaining single-word-substitution hunk in the `other`
bucket, checked whether the docm's word is present in the REAL KJV text
at that EXACT verse reference and WEBU's word is absent from it -
per-verse ground truth, not a generic table. An unfiltered first pass
found 280 "confirmed" hits, but many were common short words (`"the"`,
`"was"`, `"who"`) that trivially appear in nearly any KJV verse
regardless of real causation. Restricting to distinctive words (length
>= 4 letters, excluding a ~50-word stopword list) narrowed this to
**198 high-confidence hits** - several patterns recurring enough to be
certain: `"Baptizer"` -> `"Baptist"` (15x), `"give"` -> `"render"` (10x),
`"chest(s)"` -> `"breast(s)"` (7x), Ezekiel's temple-vision `"nave"`/
`"temple"` vocabulary (7x), `"tunic(s)"` -> `"coat(s)"` (4x), the rest
genuine single-occurrence matches (`"sulfur"` -> `"brimstone"`, `"hades"`/
`"tartarus"` -> `"hell"`, `"hall"` -> `"porch"`, and many more).

**Built as a real classifier category, not a one-time list**: `aeRWB`
commit `956a080` adds `kjv-word-choice`, checking every remaining
single-word substitution against `kjv.txt` live at classify time (a
lazily-loaded, memoized reference bible), not a static word-pair table -
this means any FUTURE docm edit that happens to match real KJV wording
at its verse gets picked up automatically, the same way `nt-variant-
addition`'s ref-keyed check does for the 4 known textual-variant verses.

**Result: WEBU `other` 1,385 -> 1,239 (-146, now 10.2%), WEBBE 1,886 ->
1,697 (-189, now 11.9%)** - by a wide margin the single largest reduction
across this entire categorization effort (compare Addendum 1's -68 and
Addendum 2's -68). Confirms the operator's original observation
(Genesis 3:15) was correct, and that it was one instance of a much larger,
now-verifiable pattern.

**A genuine textual-tradition finding surfaced as a side effect, distinct
from ordinary vocabulary preference**: 2 Samuel 21:8 - WEBU reads
`"Merab"` (the modern scholarly emendation, since Michal is stated
childless at 2 Samuel 6:23), the KJV reads `"Michal"` (the Masoretic
Text's literal, textually disputed reading), and the docm already reads
`"Michal"` - matching the KJV/traditional side, the Old Testament
counterpart to the New Testament NU/TR variant policy documented
elsewhere in the categorizer. Not flagged as an error - flagged for the
operator's awareness/explicit confirmation, since it's a real textual
decision riding along inside what otherwise looks like an ordinary
word-choice hunk, not something this check's stopword/length filter
would ever catch or exclude on its own.

## Addendum 4, 2026-09-29: co-occurrence blocking, generalized further

**The operator's review of the Genesis 2:19 example asked exactly the
right diagnostic question**: "we've already dealt with 'the LORD God' in
divine-names, so why is this verse still here?" Checking `classifyVerse`
directly confirmed the divine-names hunk WAS already correctly explained
(`reasons: ["divine-names", "kjv-word-choice"]` after this addendum's fix)
- the verse stayed in `other` only because of a second hunk,
`"the man" -> "Adam"`, which `kjv-word-choice` (Addendum 3) didn't catch:
it only checked 1-word-to-1-word substitutions, and this is 2 words
(`"the"`, `"man"`) to 1 (`"Adam"`).

**Fix**: extended `isKjvWordChoiceHunk` to also match `"the/a/an NOUN" ->
"ProperName"`, dropping the contentless article and allowing a shorter
minimum length for the noun (generic nouns standing in for a name are
often short - "man", "boy", "son"). Verified against real `kjv.txt` text
the same way as the 1-word case, not a guess.

**Answering the operator's follow-up ("are there others with the same
problem")** - searched the remaining `other` bucket for the same
"already-explained hunk blocked by one leftover hunk" shape more broadly
(not just article+noun) and found two more clean, strongly recurring KJV
idioms, each verified against every single occurrence before coding:

- **`"the sky"` -> `"heaven"`** (6x: Daniel 7:13, Matthew 24:30,
  Acts 1:11 x2, Acts 4:24, Revelation 20:11) - the KJV never uses "the
  sky" anywhere in the corpus.
- **`"according to"` -> `"after"`** (4x: Acts 24:14, Romans 8:4,
  2 Corinthians 5:16 x2) - the KJV's well-known "after the flesh"/"after
  the Spirit" phrasing.

**A broader raw scan (167 candidate multi-word-del hunks) also surfaced a
lot of noise** - common short words (`"then"` -> `"and"`, `"who"` ->
`"that"`, `"could"` -> `"should"`) that pass a bag-of-words KJV-presence
check almost by coincidence, the same failure mode the stopword filter
was built to catch for single-word matches. Not coded into any rule -
these need individual review, not a blanket pattern, same judgment call
as Addendum 3's stopword list.

**Result: WEBU `other` 1,239 -> 1,232 (-7), WEBBE 1,697 -> 1,690 (-7)**
(`aeRWB` commit `caccd7a`). Smaller than Addendum 3's KJV-source win, but
notable for a different reason: this round wasn't found by scanning the
data directly - it came entirely from the operator asking *why* a
specific already-partially-explained verse was still showing up, which
is exactly the review-checklist's own intended workflow (§6) working
correctly even before the tool itself is built.

## Addendum 5, 2026-09-29: a genuine "WEBU is better" defect hunt, not a classifier fix

**Different in kind from Addenda 1-4**: reviewing this plan's Exodus 2:9
example, the operator noticed docm's `"nursed it"` is objectively worse
than WEBU's `"nursed him"` - Moses's sex is already established, so "it"
reads as dehumanizing, not a stylistic choice. This isn't a classifier
gap (nothing to explain away) - it's a real content defect in the docm,
found the same way as the original 12-verse pure-deletion cluster back
at the start of this whole effort.

**Targeted heuristic searches, not a blanket scan** (each checked against
the FULL remaining `other` bucket, not just the one example verse):

1. **Personal pronoun -> "it"/"its" downgrades** (he/him/his referring to
   a person, replaced by "it"): found 8 hits across 5 verses. 3 were
   genuine defects, all a known-sex child called "it" - **Exodus 2:9**,
   **2 Samuel 12:15**, **1 Kings 3:21** (the last one internally
   inconsistent - the same sentence later calls the same referent
   "son"). The other 3 (Ephesians 5:25-27, "her" -> "it" for "the
   assembly"/church) were judged a DIFFERENT, defensible category -
   de-personifying an institution, not misgendering an individual - and
   deliberately excluded, not fixed.
2. **Wrong-sex pronoun swaps** (he/him/his <-> she/her): **zero found**,
   both before and after the fixes below.
3. **Numeral/digit changes**: **zero found**.
4. **Negation drops** (bare "not"/"never" removed with nothing
   compensating anywhere in the verse - real meaning-reversal risk):
   34 hits, individually reviewed. 31 were safe (restructured with
   "neither"/"nor", or matching real KJV phrasing like "spared to take",
   "let not me", "without sin unto salvation" - verified against
   `kjv.txt`, not guessed). **3 were genuine typos, not edits**:
   Luke 16:28's `"wil"` (missing a letter), 2 Corinthians 12:18's
   `"Didt"` (garbled), Hebrews 11:5's `"wouldnot"` (missing a space) -
   confirmed by checking all three broken forms appear ZERO times in
   either `kjv.txt` or `engwebu.txt`, ruling out an archaic-spelling
   explanation.
5. **Count-noun mismatches** (son/sons, man/men, etc.): 4 hits, all
   checked against `kjv.txt` and confirmed SAFE - e.g. Psalm 140:1's
   singular "the evil man"/"the violent man" is the KJV's own exact
   wording (WEBU's plural "men" is the outlier, not docm).

**All 6 confirmed defects fixed** (`aeBibleClass` commit `f37ecc9`) via
the same one-at-a-time Word-edit workflow used throughout. Verified: 4 of
6 fully resolved by the categorizer; the other 2 (2 Samuel 12:15,
Luke 16:28) still show `other` for a separate, unrelated, pre-existing,
low-priority reason (a dropped leading "Then"; an em-dash/comma
punctuation shape) - not defects, not addressed here.

**Follow-up round found zero new defects** - re-ran all five heuristics
against the post-fix state: pronoun downgrades, gender swaps, and
numeral changes all at 0; the 4 count-noun cases and all remaining
negation-drops re-confirmed safe. This specific defect class appears
exhausted for now - further "WEBU is better" hunting would need a
different heuristic angle (offered to the operator, not yet pursued:
dropped proper names replaced by vague pronouns, location-word swaps).

Result: WEBU `other` 1,232 -> 1,228, WEBBE 1,690 -> 1,686.

## Addendum 6, 2026-09-29: lie/lay/laid, and a session-close checkpoint

**The operator's review of the Ruth 3:4 example** ("lay" vs "lie") asked
directly for the KJV-preferred archaic verb form, same instinct as every
prior addendum. Confirmed against real `kjv.txt`: Ruth 3:4's "lay thee
down" and Ruth 3:7's "laid her down" are both exact KJV wording -
`isKjvWordChoiceHunk` had missed them only because `lie`/`lay`/`laid` are
3-4 letters, under the general length>=4 filter built for common
FUNCTION words, not distinctive irregular VERB forms.

**Fix, two parts** (`aeRWB` commits `b0d844b`): (1) a small curated
`KJV_SHORT_WORD_ALLOWLIST` (`lie`/`lay`/`laid`/`lain`) exempted from the
length filter - checked first whether a broader curated set of ~30
archaic/modern irregular-verb pairs (drank/drunk, sang/sung, ran/run,
ate/eat, etc.) found more misses; it found none, confirming those longer
forms already clear the length filter on their own. (2) Ruth 3:4 needed
a SECOND fix: the KJV's own text uses BOTH "lie" and "lay" in that verse
for two different clauses, so the check's usual "old word must be absent
from KJV" safety rule incorrectly disqualified a real match - relaxed
that rule specifically for the small pre-vetted allowlist, not generally.

**Result: WEBU `other` 1,228 -> 1,226, WEBBE 1,686 -> 1,684.**

### Session-close checkpoint, 2026-09-29

**This specific investigative thread (lie/lay and the broader "similar
examples" archaic-verb-form hunt) is closed** - confirmed exhaustive, no
further action pending on it. **The R15 plan as a whole remains NOT
closed** - the review-checklist tool described in the requirements below
still has not been built; everything in Addenda 1-6 happened while
reviewing this PLAN's own worked examples, which is a strong signal the
plan's core idea (side-by-side `rwb`/`webu` review surfaces real findings
even manually) works, but doesn't substitute for building it.

**Cumulative result across this whole document's addenda**: WEBU `other`
50.7% (original, `rvw/Code_review 2026-09-25.md` item 4) -> **10.0%**
(1,226/12,200), WEBBE 61.7% -> **11.8%** (1,684/14,242). Real docm content
defects found and fixed across every addendum combined: ~36 verses (see
`project_docm_categorization_and_nt_variant_policy` memory for the
full account). One item flagged but not yet resolved: 2 Samuel 21:8's
"Michal"/"Merab" textual-tradition question (Addendum 3) - awaiting the
operator's explicit confirmation, not blocking anything.

**For the next session, in priority order:**
1. Decide whether to actually build the R15 review-checklist tool now
   (§§1-10 below), or continue the manual "review a worked example,
   find + fix what it surfaces" pattern that's been working well without
   it - both are legitimate; this is the operator's call, not a
   default.
2. If continuing manually: the remaining `other` bucket (WEBU 1,226,
   WEBBE 1,684) is dominated by single-hunk substitutions - `npm run
   docm.categorize` then filter `category === "other"` for the next
   batch to review, same technique used throughout this document.
3. Explicit confirmation still wanted on 2 Samuel 21:8 (Merab/Michal).
4. Further "WEBU is better" defect-hunt angles not yet tried (offered,
   not pursued): dropped proper names replaced by vague pronouns,
   location-word swaps, other semantic-risk shapes beyond the five
   already checked (Addendum 5).

## Addendum 7, 2026-09-29: the tool was actually built, revised, and used - session close

**Answers Addendum 6's item 1**: the operator chose to build the R15 tool
(§§1-10 below) rather than continue purely manually. Built as
`aeRWB/tools/web-diff/generate-review-checklist.mjs` (`npm run
docm.review-checklist -- webu`), with shared parsing/idempotency helpers in
`lib.mjs` and a classifier-feedback connection in
`docm-webu-webbe-categorize.mjs`. All four of §7's open questions were
confirmed at their recommended defaults before building - see that section
for the options; §3's design is what actually shipped, WITH one revision
below.

**Design revision, same day, from real use**: the original single-file
design (§2/§3 as written) kept `rwb[x]`-approved verses sitting in the same
file forever, as an in-file audit trail. The first real approval
round-trip (13 verses checked, file closed and reopened) showed this is
"confusing and distracting to work from" in practice - not a bug, the code
worked exactly as designed and confirmed, but the confirmed semantics
weren't what was actually useful day to day. **Fixed by splitting into two
files**: `rwb-webu-review.txt` is now PENDING-ONLY (a checked verse
graduates out entirely on the next regeneration), and a new
`rwb-webu-accepted.txt` is the permanent, append-only ledger
`docm-webu-webbe-categorize.mjs`'s `operator-accepted` tally now reads.
Full rationale and mechanics in `generate-review-checklist.mjs`'s own
header comment and the `project_rwb_review_checklist_tool` memory - not
re-litigated here. This supersedes §2/§3's single-file design for the
checkbox-persistence question specifically; everything else in §§1-10
(file shape, scope, git-tracking decision, idempotency's core
byte-identical-text rule) still holds.

**A second, unrelated design-confirmation gap found the same way**: when
first asked "how many of the remaining `other`-bucket verses containing
'the LORD' need review" (a real question, not hypothetical - the operator
noticed the pattern), the first answer was wrong - it checked only whether
`reasons` included the `divine-names` tag, missing that some hunks are
explained via `formatting` instead (a pure case-only `LORD`→`Lord` change).
Corrected: the right check is whether the LORD-touching hunk itself appears
in `unexplainedHunks`. Re-checked properly, ALL 123 such verses (at the
time) already had that hunk fully explained - the divine-name swap was
never the reason any of them were still listed; a different, unrelated
hunk in each verse was the real remaining item. Same process lesson as
Addendum 4: verify the specific claim against `classifyVerse`'s actual
output, don't reason from the category tag alone.

**Classifier fix from the same investigation**: `isDivineNamesHunk` missed
a divine name directly joined to the next word by an em dash with no space
(`"LORD—instead"`, `"LORD—a"`) - a first fix attempt (adding "—" to the
punctuation-stripping regex) only covered a TRAILING em dash, not one with
a real word glued on past it; the working fix splits each word on "—"
before testing each piece. Caught only by re-verifying every individual
case the fix claimed to resolve, not trusting the aggregate before/after
count - 2 of the first 3 target verses were still unresolved after the
first attempt.

**Real docm defects found and fixed this session** (all individually
verified against WEBU/KJV before fixing, same one-at-a-time Word-edit
workflow as every prior session): 12 missing-space defects (punctuation
glued to the next word/sentence) + 1 stray-semicolon typo (Micah 4:3) +
~30 more dropped/added/duplicated/misordered-word defects across Genesis,
Exodus, Leviticus, Numbers, Deuteronomy, Joshua, Judges, 1-2 Samuel,
1 Kings, Nehemiah, Isaiah, John. One flagged and caught BEFORE committing:
2 Samuel 19:36 was over-trimmed to a broken sentence ("Your servant just go
over" - missing "will") in one export, then fixed before being confirmed.
One flagged via a live-docm screenshot: Genesis 9:27 - the pending-list
entry showed "May Canaan" not matching WEBU's "Let Canaan", operator
confirmed the LIVE docm already said "Let Canaan" (screenshot with
formatting marks on, ruling out any hidden-character theory) - the export
had simply gone stale; re-exporting fixed it, not a tool bug.

**New work-list item, not built**: WEB confirmed as a modernization of the
1901 ASV (previously undocumented in this repo) - proposed as a future R17
(source ASV text same as R16's `kjv.txt`, add an `asv-word-choice`
category). A second, much smaller candidate ("stuff" avoidance, matching
the existing no-contractions-style rationale) was investigated and found
to only affect 1 remaining verse - not built, left as manual approval.
Both documented in `aeRWB/tools/web-diff/README.md`'s R15 section, not
duplicated here.

**Session-close numbers**: WEBU `other` 1,226 (session start) → **1,084
pending + 88 accepted** = 1,172 total still tracked (12,163 changed
verses, 8.9% pending). WEBBE untouched this session (still 1,630/14,207 =
11.5%, no ledger started - WEBU-first scope, unchanged from §1).

## Addendum 8, 2026-09-30: verse-by-verse review continued in batches, two new standing-accept patterns, session close

Picked up Addendum 7's next-session item 1 (continue the verse-by-verse
review using the two-file workflow). Worked through the pending list in
~10-30-verse batches, presenting each batch's findings (defect vs.
accept-pattern vs. genuinely uncertain) before touching the file, one
finding at a time with a recommendation, waiting for confirmation before
editing - the established review-fix process for this project. KJV/ASV
systematic cross-checking stayed
deferred per Addendum 7's own scope decision (the real scholarly grounding
is Strong's-based, not KJV/ASV - see `project_scholarly_grounding_plan`
memory), but **targeted KJV lookups (biblegateway.com) turned out to be the
fast, decisive way to resolve individual "uncertain" holds** - checking the
actual KJV wording confirmed or refuted almost every archaic-vs-modern
word-choice question and every ambiguous-pronoun-antecedent question raised
this session. Not a reversal of the deferral - a narrower, one-off
verification tool, not a systematic pass.

**Two new standing-accept patterns confirmed** (see
`project_rwb_editorial_philosophy` memory for the full writeup):
1. Docm keeping a more literal/archaic KJV-style rendering where WEBU
   modernized the wording (e.g. "maiden" vs "girl", "cried, and said" vs
   "cried out and said") - a same-meaning register/word-choice swap, not a
   defect.
2. Trivial connector-word/punctuation swaps ("and"/"then", comma/semicolon)
   and pronoun-for-clear-antecedent-proper-name swaps, when no meaning
   changes - extended from pattern 1 mid-session, confirmed by the
   operator.

An ambiguous-pronoun hold (1 Chronicles 4:17, "she bore Miriam" vs webu's
interpretive "Mered's wife bore Miriam") was resolved by confirming KJV/ASV
both use the same literal "she" - a genuine textual crux in the Hebrew, not
a docm error. Documented as a reusable precedent: when a hold is an
ambiguous pronoun/reference, check whether KJV/ASV use the SAME ambiguity,
not just any different wording.

**Real docm defects found and fixed this session** (each individually
verified against WEBU, several also checked against KJV before deciding):
missing/dropped words, subject-verb and number-agreement errors (e.g. "eat
he who" → "eat whoever", singular/plural antecedent mismatches), a
meaning-reversal defect (**2 Chronicles 11:16**: "made way for her" →
"seized her" - opposite actions), a theological-clarity gap (**2 Chronicles
11:15**: "male goats...calves" → "male goat and calf idols," since the
plain wording lost the point that these were idols, not literal
livestock), a quantifier/verb softening (**2 Chronicles 11:23**: "found
wives" → "sought many wives," matching both webu and KJV - **fix still
outstanding, not yet applied as of this addendum**), several title-before-
name capitalization defects ("king Jehoash"/"king Josiah"/"king David" →
capitalized), and a handful of stray-space/preposition/tense fixes.

**One process hiccup, caught and corrected**: a requested "of"→"by" fix for
1 Chronicles 24:27 landed on the adjacent, near-identical verse 24:26
instead (both start "The sons of Merari..."), introducing a new defect
there. Caught on the next diff review (comparing the FULL diff, not just
the requested verses) before committing as intentional; both verses
corrected in a follow-up export. Worth remembering: when two adjacent
verses share near-identical leading text, a manual find/replace fix is at
real risk of landing on the wrong one - diff the full changed-file output,
not just the targeted verse, before treating a fix as confirmed.

**Tooling fix, same session**: `generate-review-checklist.mjs`'s
provenance line only recorded a date, not a time - ambiguous across
multiple regenerations in one day. Fixed to include time (`aeRWB` commit
`0e51d50`).

**Session-close numbers**: WEBU `other` 1,084 (session start) → **862
pending + 275 accepted** (12,146 changed verses, 7.1% pending). WEBBE
untouched this session (still 1,595/14,191 = 11.2%, no ledger started -
WEBU-first scope, unchanged). One fix from this session's own batches
(2 Chronicles 11:23) is not yet applied - see next-session tasks in
`sync/session_manifest.txt`.

## Addendum 9, 2026-10-01: review continued (2 Chronicles 11:23 -> Isaiah 56:10), fourth standing-accept pattern, new Test 90, ImportAllVBAFiles finding

Continued Addendum 8's next-session item 1/2 (apply the one outstanding
fix, then keep working the pending list in batches). Same workflow as
Addendum 8: present each batch with a recommendation (standing-pattern
accept / likely defect / genuinely uncertain), wait for confirmation,
fix in Word + re-export + re-run for defects, mark `rwb[x]` + re-run for
accepts. 9 batches worked this session, covering 2 Chronicles 11:23
through Isaiah 56:10.

**Fourth standing-accept pattern confirmed** (see
`project_rwb_editorial_philosophy` memory for the full writeup): docm
inserts inline speaker labels ("Lover", "Friends", "Beloved") into Song of
Solomon's verse text and groups multiple webu-numbered verses' dialogue
under one docm verse number. Confirmed intentional by the operator -
"shows the speakers clearly, uses a specific style" - not a verse-split
bug. Scoped to this one book; any docm-vs-webu diff there with an inline
speaker label is accepted by default going forward.

**Real docm defects found and fixed this session** (each verified against
WEBU, most also cross-checked against KJV): dropped words (Proverbs 22:5
"far"; Psalm 105:35/113:7/122:3 dropped subjects or a dangling relative
clause, Ecclesiastes 2:8 a garbled trailing phrase), number-agreement
errors (Ezra 8:31/8:33, Psalm 72:15, Isaiah 3:11 - a "them...his...him"
mix within one verse, two separate touches needed), a homophone typo
(Psalm 84:9 "you're" -> "your"), a tense defect (Psalm 108:10 "has led"
-> "will lead"), meaning-altering word-choice defects (Proverbs 28:17
"life blood" -> "blood guilt"; Ecclesiastes 7:27/7:29 "scheme(s)" ->
clearer wording; Isaiah 53:3 restored "sorrows" where docm had
repeated "suffering" twice, losing the KJV's word variety), a broken-
grammar fix (Esther 9:1 "was turned out" -> "turned out"), a stray-space
typo recurring twice (Isaiah 36:1, 39:3 - "King Hezekiah , ..." - this
specific pattern is what prompted the new Test 90 below), a missing
proper-name hyphen (Isaiah 39:1 "Merodach Baladan" -> "Merodach-Baladan"),
and a title-capitalization defect (Proverbs 31:1 "king Lemuel" -> "King
Lemuel", same family as prior king-name fixes). One quantifier/number
fix (2 Chronicles 11:23, carried over from Addendum 8 as "outstanding")
was applied first thing this session.

**New test added, same session, operator-initiated** (unrelated to the
R15 review itself, but prompted directly by the Isaiah 36:1/39:3 stray-
space pattern found above): Test 90, `CountSpaceBeforePunctuation`, added
to `aeBibleClass.cls` per the standard 8-location checklist in
`md/Adding_To_Bible_Test_Class.md`. Checks for a space immediately before
`, . : ; ! ?` or a closing `)` anywhere in the document (opening marks
and apostrophes deliberately excluded - apostrophes already covered by
the contractions tests 56-65). Uses the established `m_lastHint`
first-violation-hint convention so a FAIL is immediately actionable.
`RUN_THE_TESTS(90)` confirmed PASS (0 violations) after import.

**ImportAllVBAFiles investigation, same session**: hit the known
`Skipped=4` anomaly (`feedback_importallvbafiles_error17` memory) while
importing Test 90. Operator's own troubleshooting found a likely root
cause not previously documented: the project must be **compiled BEFORE**
running `ImportAllVBAFiles`, not just after (the existing guidance only
covered compiling afterward, to avoid a separate silent-hang issue). One
paired observation (uncompiled -> fails every retry; compiled first ->
clean run) - not yet proven fully deterministic, but strong enough to
become the new default recovery step. Documented in
`feedback_importallvbafiles_error17` memory with the full sequence.

**Session-close numbers**: WEBU `other` 862 pending + 275 accepted
(2026-09-30 session start) → **722 pending + 400 accepted** (12,146
changed verses, 5.9% pending). 125 verses accepted into the ledger this
session; 24 real defects fixed (across 2 Chronicles, Ezra, Esther, Psalms,
Proverbs, Ecclesiastes, and Isaiah) via 7 commits. WEBBE untouched this
session (still 1,595/14,191 = 11.2%, no ledger started, WEBU-first scope
unchanged). No outstanding fixes carried over this time - next session
can start straight from the pending list wherever it resumes after
Isaiah 56:10.

## The ten requirements, as given

1. Update the memory/README with current status, but do not consider the
   categorization work closed yet. **✅ Done** as a prerequisite to this
   plan - see `aeRWB/tools/web-diff/README.md` commit `5fd4f4d` and the
   `project_docm_categorization_and_nt_variant_policy` memory update.
2. Prepare a plan that will list the remaining "leftover-hunk shapes" to
   create one file with both comparisons per verse.
3. The file will be the shape of `rwb.txt` itself and also UTF-8.
4. It will add a new tab-separated starting field of `rwb` and `webu` for
   each verse record.
5. The script generating the file will also allow the same to be run
   against WEBBE so that comparison can be made for that separate
   publication too.
6. This file is intended to be reviewed in GitHub Desktop so the list of
   changes can be readily reviewed.
7. A mechanism is desired such that a checkmark can be added to the `rwb`
   record to signal its acceptance, connected to the classifier such that
   it is idempotent and continues to work across a number of editing and
   approval cycles.
8. The list is expected to reduce significantly; any remainder will be
   corrected such that a new export from the docm run through the
   classifier will allow them to also be checked off.
9. This process should conclude the review of the rwb/webu (and later
   rwb/webbe) work.
10. Pros/cons/benefits/risks and suggestions for consideration to be
    reviewed in this planning doc.

## 1. Scope and definition of done

**In scope for the first pass: WEBU only** (point 9's own phrasing, "and
later rwb/webbe", plus the general-divergence risk noted in §8 below -
reviewing the same content twice in parallel wastes effort). WEBBE's
checklist is the same tool, same file shape, run a second time once WEBU's
list is worked down - not a separate design.

**"Done" is objectively defined**: the checklist file's unchecked-entry
count reaches zero. At that point every docm-vs-WEBU divergence has either
been (a) fixed in the docm, (b) explained by a classifier rule, or (c)
explicitly reviewed and accepted by the operator as deliberate RWB
editorial content. This is a stronger, more auditable completion criterion
than "the `other` percentage looks low enough."

## 2. File shape

Matches `rwb.txt`'s own conventions (UTF-8, tab-separated, `Book C:V` ref
key) per requirement 3, with one addition per requirement 4 - a leading
`rwb`/`webu` field, so **every verse becomes two adjacent lines** instead
of one:

```
RWB-REVIEW
docm-vs-WEBU, generated <date> from aeBibleClass <docm commit> / aeRWB <engwebu.txt commit> - see aeRWB/tools/web-diff/README.md R15
rwb[ ]	Ruth 3:4	It shall be, when he lies down, that you shall note the place where he is lying. Then you shall go in, uncover his feet, and lay down. Then he will tell you what to do."
webu	Ruth 3:4	It shall be, when he lies down, that you shall note the place where he is lying. Then you shall go in, uncover his feet, and lie down. Then he will tell you what to do."
```

- Two header lines, matching `rwb.txt`'s own `translation`/`source` header
  shape (so a reader immediately recognizes the family of file this is).
- One `rwb`/`webu` pair per unresolved verse, in canonical book order
  (same order `diffBibles` already walks) - not grouped by category, since
  the whole point is these are the *un*explained ones.
- The `rwb` field carries the acceptance mark (requirement 7) - proposed
  as `rwb[ ]` (unchecked) / `rwb[x]` (checked), a plain-text checkbox the
  operator edits directly in a text editor or even GitHub Desktop's own
  diff-adjacent editor. `webu` never carries a mark - it's read-only
  reference context sitting next to the record being judged.
- **"rwb" here means the docm's current text** (the same content that
  becomes `rwb.txt` after a sync), not `rwb.txt` itself - `rwb.txt` can
  lag behind an unsynced docm fix by a cycle, and reviewing stale text
  would defeat the purpose.
- **Explicitly NOT interchangeable with `rwb.txt`** despite the shape
  resemblance - four columns after splitting on the `rwb[ ]`/`webu`
  marker vs. `rwb.txt`'s plain two, and every other line is WEBU reference
  text, not RWB content. No existing tool (`docm.rwb.diff`,
  `apply-docm-rwb-*-sync.mjs`, etc.) should ever be pointed at this file.

## 3. Idempotency and the checkmark/classifier connection

Requirement 7 asks for the checkmark to be "connected to the classifier"
and to survive "a number of editing and approval cycles." This is the
part of the design with the most real risk (see §5) and the part most
worth the operator's attention before building anything.

**Proposed mechanism, mirroring a pattern already proven in this
codebase** (`divine-names-census.mjs`'s `RATIFIED_EXCEPTIONS`/
`KNOWN_INSCRIPTION_REFS`/`KNOWN_TITLE_REFS` - ref-keyed manual overrides
that sit *alongside* the mechanical classifier, not inside it):

1. **Source of truth for acceptance state is the checklist file itself**,
   git-tracked (see §4) - not a separate ledger. On each regeneration, the
   script reads the *previous* committed version of the file first.
2. For every verse currently in the categorizer's `other` bucket:
   - If it existed in the previous file **with `rwb[x]`** *and* both the
     `rwb` and `webu` text are byte-identical to what was there before ->
     carry the checkmark forward (`rwb[x]`), still listed.
   - If it existed before but the text changed (either side) -> **reset to
     `rwb[ ]`** - an approval is tied to the exact text pair it was given
     for, never silently carried onto different content. This is the
     single most important correctness rule in the whole design.
   - If it's new (never seen before) -> `rwb[ ]`.
3. For every verse that was in the previous file but is **no longer** in
   the current `other` bucket at all (whether because the docm was fixed,
   or because a classifier rule now explains it) -> **drop it from the
   file entirely.** Its disappearance from the diff *is* the resolution
   signal (requirement 8) - no separate "resolved" list needed.
4. **The "connected to the classifier" part**: once a verse carries
   `rwb[x]`, feed that ref set back into `docm.categorize` itself as a new
   ref-keyed override (same shape as `nt-variant-addition`'s
   `KNOWN_VARIANT_ADDITION_REFS`) - e.g. an `operator-accepted` category -
   so the categorizer's own `other` percentage in its console/summary
   output reflects human sign-off too, not just mechanical pattern
   matches. Read from the checklist file at classify time (or a small
   generated sidecar derived from it - see the open question in §7).

## 4. Script and tracking

- New script, next in the existing `R`-numbering: `aeRWB/tools/web-diff/
  generate-review-checklist.mjs` (R15), `npm run docm.review-checklist --
  webu` / `-- webbe` (source parameterized per requirement 5, reusing
  `classifyVerse`/`splitHunks`/`diffBibles` exactly like `docm.categorize`
  does - no new diffing logic, just new output shaping).
- Output: `aeRWB/rwb-webu-review.txt` (and later `rwb-webbe-review.txt`),
  at repo root alongside `rwb.txt`/`engwebu.txt`/`webbe.txt` - **git-tracked,
  not gitignored**, unlike `categorize/`. This is a deliberate departure
  from this repo's usual "regenerable output stays gitignored" convention
  (see `.gitignore`'s existing `diff/`/`census/`/`webbe-diff/`/
  `categorize/` entries) - required because requirement 6's GitHub Desktop
  review workflow only works on a tracked file (untracked files show as a
  single opaque "new file" add, not a line-by-line diff), and because the
  checkmark state itself needs to persist and be reviewable via git
  history across sessions.

## 5. Pros / Cons / Benefits / Risks (requirement 10)

**Benefits**
- Turns an otherwise-unreviewable 1,454-verse list into a concrete,
  git-diffable, incrementally-workable artifact - GitHub Desktop's native
  diff view needs no new tooling to consume this.
- Durable audit trail: git history on this one file *is* the review log -
  who accepted what, when, visible directly in commit history.
- Architecturally consistent - reuses the ref-keyed-override pattern
  already proven for `nt-variant-addition` and in
  `divine-names-census.mjs`, rather than inventing a new mechanism.
- Generalizes to WEBBE with the same script, no new design.
- Gives the categorization effort a genuine, objective "done" (count -> 0)
  instead of an open-ended percentage to eyeball.

**Risks / open design questions**
- **Scale.** ~1,454 verses x 2 lines is a ~2,900-line file. A single
  GitHub Desktop diff of that size across many small edits per session
  may be slow or unwieldy to review in one sitting. (Splitting per-book is
  a possible future refinement - not in the requirements as given, noted
  as a suggestion below rather than assumed.)
- **Round-trip parsing fragility - the biggest risk in this design.** The
  generator must re-parse its own *previous* output faithfully. If the
  operator hand-edits the file in ways beyond toggling `rwb[ ]`/`rwb[x]`
  (adds a note, reorders lines, fixes a typo in the `webu` reference text
  by mistake), the next regeneration needs a well-defined, tested
  contract for what happens - silently guessing wrong here could lose
  real review work. Needs explicit test coverage before this is trusted
  with real approval cycles, the same rigor `lib.test.mjs` already applies
  to the parsing primitives this would build on.
- **Checkbox semantics are ambiguous as specified.** Does `rwb[x]` mean
  "permanently correct as deliberate RWB content" (never revisit) or "I
  looked, fine for now" (still open to reconsideration)? The mechanism as
  designed treats it as the former (a checked verse silently drops out of
  future review once resolved-or-accepted) - worth the operator
  confirming that's the intended meaning before building it.
- **File-format collision risk.** Structurally resembling `rwb.txt`
  invites an accidental "wire it into an existing sync tool" mistake
  later, precisely because it looks similar. Mitigated by naming
  (`rwb-webu-review.txt`, not anything containing bare `rwb.txt`) and by
  the README documenting the distinction explicitly (§2 above) - still
  worth flagging as a real risk class, not just a naming nicety.
- **Departure from this repo's gitignore convention.** Every other
  regenerable categorization artifact (`categorize/`, `diff/`, `census/`)
  is gitignored; this one deliberately isn't, for good reason (§4) - but
  it means every regeneration is a commit-worthy event, and the operator
  should expect more commit traffic on this one file than the rest of the
  gitignored-output family.
- **WEBBE overlap.** Per requirement 9's own "(and later rwb/webbe)"
  phrasing, running both checklists in parallel would mean reviewing
  substantially the same underlying content twice (WEBU and WEBBE only
  differ in spelling, already handled separately by the
  `spelling-variant` category) - sequencing WEBU-first, WEBBE-second
  avoids that duplication.

**Suggestions for consideration (not requested, offered as options)**
- Sequence strictly WEBU-first, WEBBE-second (matches requirement 9's own
  ordering) rather than generating both checklists at once.
- A free-text notes column (a 4th tab-separated field on the `rwb` line)
  for the operator's own bookkeeping - e.g. "deliberate paraphrase" vs.
  "fixed in Word, pending re-export" - genuinely optional, easy to add
  later without a redesign if wanted.
- Have the generator **warn**, not just silently reset, when a
  previously-`[x]`-checked verse's text changed and its mark was dropped -
  makes the "why did this come back?" case legible without digging into
  git history.
- Per-book splitting of the output file as a scale mitigation, only if the
  single-file version proves unwieldy in practice - deliberately not
  designed in up front, since it adds real complexity (N files instead of
  1, cross-file idempotency) for a problem that might not materialize.

## 6. Workflow once built

1. `npm run docm.review-checklist -- webu` generates/updates
   `rwb-webu-review.txt` against the current docm export.
2. Operator reviews in GitHub Desktop - each verse pair is one small,
   readable diff hunk if anything changed since last run.
3. Operator either (a) edits the docm in Word for a real defect, re-runs
   `ExportDocmVersesToRWBFormat`, then regenerates the checklist (the
   fixed verse drops off automatically - §3 step 3), or (b) marks `rwb[x]`
   directly in the file for a verse judged to be deliberate RWB content,
   commits that change themselves.
4. Repeat until the file's unchecked-entry count reaches zero.
5. Only then, generate and begin the WEBBE checklist the same way.

## 7. Open questions for the operator before implementation begins

- Confirm the `rwb[ ]`/`rwb[x]` inline-checkbox design (§2/§3) over the
  alternative of a separate small sidecar "accepted-refs" ledger file
  (lower round-trip risk, but the checkmark wouldn't live directly next to
  what's being judged, and a second file to keep in sync is its own
  complexity).
- Confirm checkbox semantics: permanent acceptance vs. "revisit later" (§5).
- Confirm this file should be git-tracked (departure from this repo's
  usual convention) rather than gitignored with approvals stored
  elsewhere.
- Confirm scope (WEBU-only first pass, per §1) before WEBBE work starts.

## Addendum 10, 2026-10-02: review continued (Isaiah 57:8 -> Amos 4:2), 12 real defects, a stale-ledger-entry gap found and documented

Continued Addendum 9's next-session item 1 (keep working the pending list
in batches). Same workflow as Addenda 8/9: present each batch with a
recommendation (standing-pattern accept / likely defect / genuinely
uncertain), wait for confirmation, fix in Word + re-export + re-run for
defects, mark `rwb[x]` + re-run for accepts. 9 batches worked this
session, covering Isaiah 57:8 through Amos 4:2 (244 verses).

**Several cases where docm looked wrong but WEBU had actually diverged
from KJV**, confirmed by direct KJV cross-reference each time rather than
assumed: Ezekiel 11:19's "within you" mid-verse pronoun shift and Ezekiel
31:10's "you...he" shift are both genuine KJV quirks, not docm errors;
Daniel 11:17's "to corrupt her" and Daniel 11:38's "in his place" both
match KJV word-for-word where WEBU had substituted different wording;
Hosea 7:6's "their baker sleeps" is the literal KJV reading (WEBU's "their
anger smolders" is the one that reflects a different scholarly
interpretation); Hosea 11:9 and 14:4 also had docm matching KJV's
pronouns/wording exactly against a WEBU variant. Worth remembering for
future batches: a divergence from WEBU is not evidence of a docm defect
by itself - check KJV before assuming either direction.

**Real docm defects found and fixed this session** (each verified against
WEBU, most also cross-checked against KJV): two plain typos (Jeremiah 9:5
"commiting" -> "committing"; a stray mid-sentence period in Jeremiah
24:9), a wrong-word typo (Lamentations 2:9 "here" -> "where"), a meaning-
distorting number/word swap (Ezekiel 42:3 "the third floor" -> "three
stories", matching KJV/WEBU), a dangling participle with a missing verb
(Ezekiel 45:9 "...execute justice and righteousness; dispossessing my
people" -> "...righteousness. Cease dispossessing my people"), a
theologically significant singular/plural slip (Daniel 3:14 "my god" ->
"my gods", since Nebuchadnezzar is polytheistic), an awkward stacked-
preposition phrase (Daniel 6:20 "near to the den to Daniel" -> "near to
the den where Daniel was"), a wrong body part (Daniel 10:5 "thighs" ->
"waist", matching WEBU's modernization of KJV's "loins"), a dropped
clause (Hosea 2:23 "they will say, 'My God!'" -> "...'You are my God!'"),
a meaning-altering word swap (Hosea 4:10 "abandoned giving to God" ->
"abandoned listening to God", since the KJV sense is "stopped heeding/
obeying," not "stopped donating"), a number-agreement error (Hosea 11:3
"I took them by his arms" -> "...by their arms", same defect family as
the earlier Isaiah 3:11 fix), and a garbled idiom (Hosea 14:2 "we offer
our lips like bulls" -> "we offer the praise of our lips...like bulls for
sacrifice", clarifying KJV's "render the calves of our lips" metaphor).

**A real, previously-undocumented tool gap found and documented** (not
fixed - no code change, just discovered live and written up): a ledger
entry (`rwb-webu-accepted.txt`) for a verse that fully resolves - stops
being an `other`-bucket diff at all, rather than just changing text - is
never revisited by the regeneration loop, since that loop only walks refs
still present in the current `other` bucket. The stale pre-fix text is
left sitting in the ledger forever unless hand-corrected; harmless (the
verse itself is fine, `docm.categorize` no longer flags it as a diff) but
inaccurate as an audit trail. Found live via Ezekiel 23:7 (fixed
"whoever" -> "whomever" in Word, but the ledger still held the old
"whoever" wording after regeneration) and hand-corrected in the ledger
file directly. Documented in `aeRWB/tools/web-diff/README.md`'s R15
section as a known gap, with the symptom to watch for
(`docm.categorize`'s `operator-accepted` count running below the ledger's
line count) and the manual-fix procedure. No automatic cleanup built -
not requested, and the gap is rare and easy to spot/fix by hand when it
occurs.

**Session-close numbers**: WEBU `other` 722 pending + 400 accepted
(2026-10-01 session close) → **478 pending + 635 accepted** (12,137
changed verses, 3.9% pending). 235 verses accepted into the ledger this
session; 12 real defects fixed (across Jeremiah, Lamentations, Ezekiel,
Daniel, and Hosea). WEBBE untouched this session (1,570/14,182 = 11.1%,
down slightly from the defect fixes reducing the shared WEBBE diff too;
no ledger started, WEBU-first scope unchanged). No outstanding fixes
carried over - next session can start straight from the pending list
wherever it resumes after Amos 4:2.

## Addendum 11, 2026-10-02 (same day, continued): review continued (Amos 4:10 -> Zechariah 3:6), 3 more real defects, a second stale-ledger-entry instance confirmed

Continued straight on from Addendum 10 in the same session, same workflow.
2 more batches worked, covering Amos 4:10 through Zechariah 3:6 (62
verses).

**3 more real docm defects found and fixed**: a sentence-initial
capitalization typo (Amos 6:2 "...Philistines. are they better" -> "...
Are they better"), a subject/object inversion (Micah 7:10 "my enemy will
see me" -> "My eyes will see her", matching KJV's actual "mine eyes shall
behold her" - docm had the roles reversed), and an internal pronoun
inconsistency (Habakkuk 1:9 "Their hordes...He gathers" -> "...They
gather", fixing a plural-to-singular slip within one verse - distinct
from the sustained singular "he/him" used consistently in docm's 1:10-1:12,
which reads as a deliberate stylistic choice and was left as-is).

**Stale-ledger-entry gap (documented in Addendum 10) confirmed to recur
naturally, not just a one-off**: `docm.categorize`'s `operator-accepted`
count ran 1 below the ledger's line count again this session. Investigated
with a throwaway verification script (not committed - confirmed the
one stale entry was the SAME Ezekiel 23:7 fixed last session, now sitting
harmlessly in the ledger because it's a perfect match with WEBU and so
never gets visited by the regeneration loop, exactly as documented). No
new action needed - this is the same already-understood, harmless gap,
not a new bug. Confirms the gap is a structural property of the tool
(any verse that reaches a perfect WEBU match this way will do this), not
a rare fluke worth deeper investigation.

**Session-close numbers (combining Addenda 10 and 11, one continuous
session)**: WEBU `other` 722 pending + 400 accepted (2026-10-01 session
close) → **416 pending + 696 accepted** (12,137 changed verses, 3.4%
pending). 296 verses accepted into the ledger this session; 15 real
defects fixed total (12 in Addendum 10's batches, 3 here). WEBBE untouched
by review (1,569/14,182 = 11.1%, no ledger started). No outstanding fixes
carried over - next session starts from the pending list after Zechariah
3:6.

## Addendum 12, 2026-10-03: review completed to the end of Revelation - WEBU backlog effectively closed, 8 more real defects, 21 judgment calls deferred to the operator, a self-correction lesson recorded

Picked up where Addendum 11 left off (after Zechariah 3:6) and continued
the same two-file workflow in long batches to the literal end of the
Bible - Zechariah 3:9 through Revelation 22:20, 395 verses, the entire
remaining WEBU `other` backlog.

**8 more real docm defects found and fixed**: Matthew 1:6 (dropped "the
king"), Matthew 7:14 (dropped "few", leaving a broken sentence), Matthew
20:2 ("se nt" - stray mid-word space, matching the "be lieve" pattern
found later at John 10:37), John 16:3 ("cause" missing its "be-" prefix),
Acts 4:28 ("council" where "counsel" was meant - a homophone mix-up
changing an assembly into advice/will), Romans 4:18 ("Besides hope" where
KJV/WEBU both have "Against hope" - a different theological point, not a
synonym), John 10:37 ("be lieve" - the same stray-space pattern as Matthew
20:2), plus two early-session items the operator fixed directly after a
brief "uncertain" flag (Zechariah 13:1 "spring"->"fountain" matching KJV,
and a stray extra space before an em dash at Mark 2:10).

**A self-correction worth recording as a standing caution, not just a
one-off**: Luke 24:46 was initially flagged as a defect (docm's one-clause
"Thus it is written: The Christ will suffer and rise..." looked like it
was missing KJV's "and thus it behoved/was necessary" clause). The
operator asked what "behoved" actually meant (Greek *dei*, divine
necessity) - a reasonable question that led to overstating WEB's
"necessary" wording as the authoritative reading, when it's actually one
translator's choice among several (ESV/NASB/NIV/NLT/CSB all use
"would/should suffer" instead, no "necessary" at all). The flag was
revised to call out docm's *structure* instead - until the operator
pointed out that NIV's own rendering ("Thus it is written: The Messiah
will suffer and rise...") uses the exact same one-clause structure as
docm. Both framings were wrong, for the same underlying reason: treating
KJV/WEB's phrasing as the normative baseline and docm's divergence from
it as presumptively a defect, without checking whether other major
translations independently land on docm's side. **Lesson recorded**:
before flagging a structural or wording divergence as a defect, check it
against more than one major translation (not just KJV/WEB) - a divergence
shared with NIV, ESV, NASB, NLT, or CSB is evidence of a legitimate
translation choice, not evidence of an error. Luke 24:46 itself was
retracted and accepted once this was clear.

**21 genuine judgment calls deferred to the operator**, rather than
resolved unilaterally - each is a real translation/theological question,
not a simple wording slip, and deserves the operator's own read:

- Zechariah 8:23 - docm's "will take hold... they will take hold"
  repetition may mirror KJV's own correlative "shall take hold... even
  shall take hold" idiom, or may be an accidental duplication - never
  resolved either way.
- 1 Corinthians 1:20 - docm's "lawyer" vs KJV "disputer"/WEBU "debater";
  the Greek (*syzētētēs*) means "one who argues/debates," which "lawyer"
  (a legal profession) doesn't capture well.
- 2 Corinthians 4:14 - docm's "raise us also **with** Jesus" vs KJV "**by**
  Jesus"/WEBU "**through** Jesus" - companionship vs agency framing of the
  resurrection, both defensible elsewhere in Paul.
- 2 Corinthians 11:2 - docm's "**married**" vs KJV "**espoused**"/WEBU
  "**promised in marriage**" - completed marriage vs betrothal, a real
  difference given the church-as-bride-awaiting-the-wedding theme.
- 2 Corinthians 12:16 - docm puts "being crafty, I caught you with
  deception" in quotation marks (framing it as a quoted accusation from
  opponents); KJV has no quotes (Paul's own ironic admission).
- Ephesians 2:1 - docm's "As for you, you were dead..." drops the
  anticipatory "you were **made alive**" clause that both KJV ("you hath he
  quickened, who were dead...") and WEBU include - a known hard sentence
  to render (the Greek runs on as one long clause from v1 to v5).
- Ephesians 2:5 - a stray space before an em dash ("Christ **—**by grace"),
  same cosmetic pattern as the already-fixed Mark 2:10.
- Colossians 2:8 and 2:20 (same issue, both deferred together) - docm's
  "**elements** of the world" vs WEBU's "**elemental spirits** of the
  world"; the Greek *stoicheia* is genuinely disputed in Pauline
  scholarship (basic/physical principles vs. spiritual/angelic cosmic
  powers), and KJV takes a third position ("rudiments") entirely.
- 1 Timothy 2:9 - docm adds "**just**" ("not **just** with braided hair...")
  not present in KJV or WEBU, softening an absolute prohibition into a
  "not primarily" framing - a real interpretive shift on a verse with
  genuine denominational disagreement already.
- 1 Timothy 2:14 - docm's "became **a sinner**" (general moral state) vs
  KJV "was **in the transgression**"/WEBU "**fallen into disobedience**"
  (a specific transgressive act) - different theological weight.
- 1 Timothy 2:15 - docm's "**sanctification**" vs KJV/WEBU "**holiness**" -
  related but distinct terms (process vs state).
- Hebrews 5:7 - docm's "**reverent submission**" vs KJV "he **feared**"/WEBU
  "**godly fear**" - different emphasis (obedience vs awe).
- Hebrews 7:22 - docm's "**collateral**" vs KJV "**surety**"/WEBU
  "**guarantee**" - the Greek (*engyos*) means a person who personally
  guarantees, not a pledged financial asset; "collateral" changes the
  category of what's being described.
- 1 Peter 3:3 - the same "**just**" addition as 1 Timothy 2:9 ("Let your
  beauty be **not just** the outward adorning...") - same recurring pattern,
  not a one-off.
- Jude 1:4 - a genuine, well-known textual/grammatical dispute: docm's
  "denying Jesus Christ, our only sovereign and Lord" sidesteps the
  question (present in KJV and WEBU, each resolved differently) of
  whether the verse asserts Jesus Christ directly as "God" or distinguishes
  "God" (the Father) and "Jesus Christ" (the Son) as two separate denied
  figures.
- Revelation 3:2 - docm's "**keep** the things that remain" vs KJV/WEBU
  "**strengthen**"; the Greek (*stērison*) specifically means
  strengthen/reinforce, and the context (Sardis's works are dying) favors
  the stronger reading - leans real defect, not just a judgment call.
- Revelation 3:9 - docm's "I **give** some of the synagogue of Satan..."
  vs KJV/WEBU "I **make**..." - notably, docm's own text uses "make" later
  in the *same verse* for the parallel clause, so this also reads as an
  internal inconsistency, not just a KJV divergence.
- Revelation 16:16 - docm's "**Megiddo**" vs KJV/WEBU "**Harmagedon**"
  (Armageddon) - these are related but different terms; this is the sole
  biblical source of the word "Armageddon," and substituting "Megiddo"
  loses it - leans real, significant defect.
- Revelation 16:21 - a present/past tense mismatch ("this plague **is**
  exceedingly severe" vs the past-tense narrative around it, and vs
  KJV/WEBU's "**was**") - minor, likely just a grammar slip.
- Revelation 21:9 - docm's "I will show you **the wife, the Lamb's
  bride**" vs KJV/WEBU "**the bride, the Lamb's wife**" - the terms appear
  transposed; KJV's order makes sense as an appositive (bride, then
  clarified as the Lamb's wife), docm's reversed order reads oddly - leans
  real defect (word-order transposition).

**Session-close numbers**: WEBU `other` 416 pending + 696 accepted
(2026-10-02 session close) → **21 pending + 1,080 accepted** (12,131
changed verses, 0.2% pending) - effectively closed; the 21 remaining are
all deferred judgment calls above, not unreviewed backlog. WEBBE untouched
by review (1,560/14,177 = 11.0%, no ledger started - full backlog,
unlike WEBU). 395 verses reviewed and accepted into the ledger this
session (384 batch-accepted + the Luke 24:46 retraction); 8 real defects
fixed. Next session: operator reviews the 21 deferred items above at
their own pace (no urgency, no deadline); once resolved, WEBU R15 is
fully closed and WEBBE's ledger/review pass (not yet started) becomes the
next body of work under this same workflow.

## Addendum 13, 2026-10-03/04: all 21 deferred judgment calls resolved - WEBU R15 fully closed, 0 pending

Walked through Addendum 12's full 21-item deferred list with the operator,
one at a time, same present-finding/wait-for-decision workflow as every
other verse this tool has ever reviewed.

**The 4 items flagged as "leans real defect" were all confirmed and fixed**:
- Revelation 3:2: "keep" -> "strengthen" (now byte-identical to WEBU).
- Revelation 3:9: "give" -> "make" in the first clause, resolving the
  verse's own internal inconsistency (now byte-identical to WEBU).
- Revelation 16:16: "Megiddo" -> "Armageddon" (docm's own spelling choice;
  WEBU's is "Harmagedon" - same word, different transliteration, not
  re-matched byte-for-byte but the core defect - losing the word
  "Armageddon" entirely - is resolved).
- Revelation 21:9: transposition fixed to "the bride, the wife of the
  Lamb" (a deliberate of-genitive variant, not WEBU's exact "the Lamb's
  wife" - accepted as a stylistic choice, not re-matched byte-for-byte).

**A genuine NEW typo found and fixed mid-review, not present in the
original 21-item list**: fixing 2 Corinthians 11:2 ("married" ->
"promised in marriage") left a duplicated "you" ("promised you in
marriage **you** to one husband") on the first attempt - a find-replace
artifact, not a translation question. Caught before accepting, fixed on
re-export, now byte-identical to WEBU. Worth remembering: every Word edit
in this workflow still needs its own round-trip verification, even ones
that look like simple word swaps.

**Of the 17 pure judgment calls, two were resolved differently than
either KJV/WEBU's own choice** after the operator supplied outside
cross-checks (biblegateway.com parallel views), both following the
precedent already established for exactly this situation (see the
project memory's "Exception confirmed, 2026-09-30" entry, extended here
to disputed-wording defect flags, not just ambiguous-pronoun holds):
- **2 Corinthians 4:14** ("with Jesus" vs KJV "by"/WEBU "through"): the
  operator pointed at the parallel-translations view directly - "with" is
  independently attested among major translations, confirming it's a
  legitimate reading rather than a preposition slip. Accepted as-is.
- **Jude 1:4** (one-person vs two-figures dispute): checking the parallel
  view confirmed the one-person reading (Jesus Christ directly called
  both "Master/Sovereign" and "Lord") is the majority modern-translation
  position (ESV/NIV/NASB/NLT/CSB) - and that docm's existing text already
  matches it. The real finding: WEBU's own "God" insertion isn't in the
  Greek text at all (not even in ASV, WEB's own ancestor) - it's WEBU's
  own added theological gloss, making WEBU the outlier here, not docm.
  Accepted docm as-is, no change made.

**The remaining 15 were resolved by direct discussion, no outside
cross-check needed**: Zechariah 8:23 (kept docm's KJV-echoing emphatic
repetition), 1 Corinthians 1:20 ("lawyer" -> "debater", matching WEBU;
the separate "world"/"age" wording difference was left alone as a
KJV-matching variant), 2 Corinthians 11:2 (see typo note above), 2
Corinthians 12:16 (quotation marks removed, now reads as Paul's own
ironic statement, KJV-style), Ephesians 2:1 (added the missing "made
alive" clause: "And you he made alive, who were dead..."), Ephesians 2:5
(stray space before the em dash fixed, now byte-identical to WEBU),
Colossians 2:8 and 2:20 (both "elements" -> "elemental spirits", now
byte-identical to WEBU), 1 Timothy 2:9 (dropped the added "just", now
byte-identical to WEBU), 1 Timothy 2:14 ("became a sinner" -> "fell into
sin" - a deliberate middle-ground phrasing between docm's original state-framing
and WEBU/KJV's act-framing, prompted by the operator noticing NIV itself
reads "became a sinner" verbatim), 1 Timothy 2:15 ("sanctification" ->
"holiness", now byte-identical to WEBU), Hebrews 5:7 (kept "reverent
submission" as-is, a deliberate emphasis choice), Hebrews 7:22
("collateral" -> "guarantor" - not WEBU's "guarantee," but a sharper fix:
the Greek *engyos* names a person who personally guarantees, not a
pledged asset, and "guarantor" captures that better than either the old
"collateral" or WEBU's own "guarantee"), 1 Peter 3:3 (dropped "just",
same pattern as 1 Timothy 2:9), Revelation 16:21 ("is" -> "was", a plain
tense-agreement fix, now byte-identical to WEBU).

**Final numbers**: WEBU `other` 21 pending + 1,080 accepted (2026-10-03
session close) → **0 pending, 1,092 accepted** (12,131 changed verses,
0.0% pending). **WEBU's R15 ledger/review pass is now fully closed.**
WEBBE remains the full, untouched backlog (1,560/14,177 = 11.0%, no
`rwb-webbe-review.txt`/`rwb-webbe-accepted.txt` exist yet) and is the next
body of work under this same workflow - see this plan's own Status
section and `generate-review-checklist.mjs`'s existing (built, never run)
`-- webbe` parameterization.
