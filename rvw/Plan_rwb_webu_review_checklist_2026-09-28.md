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

✅ Built 2026-09-29 (the review-checklist tool itself, §§1-10) - operator
chose to build now rather than continue purely manually (see Addendum 6's
session-close checkpoint). All four §7 open questions confirmed as their
recommended defaults: inline `rwb[ ]`/`rwb[x]` checkbox, permanent
acceptance semantics, git-tracked output, WEBU-only first pass. Built as
`aeRWB/tools/web-diff/generate-review-checklist.mjs` (`npm run
docm.review-checklist -- webu`), with the shared parsing/idempotency
helpers in `lib.mjs` (`parseChecklistFile`/`buildChecklistEntries`/
`formatChecklistFile`, covered by `lib.test.mjs`) and the "connected to the
classifier" feedback wired into `docm-webu-webbe-categorize.mjs` (an
`operator-accepted` tally, sourced from the checklist file, layered on top
of - not replacing - the raw `other` bucket). First real run against the
live corpus: 1,226 verses (0 accepted), matching the R14 baseline exactly.
Full design in `aeRWB/tools/web-diff/README.md`'s new "Review-checklist
file (R15)" section. Not yet committed in `aeRWB` - awaiting operator
review before push (per the "no auto-push in aeRWB" rule).

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
