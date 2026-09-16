# Phase 4 Pass 3 sub-plan - divine-name presentation (LORD/GOD/Lord/God) - 2026-09-15

Follow-on to `Plan_rwb_phase4_content_sync_2026-09-15.md` Pass 3. Every claim
below was checked directly against the actual corpus files in this repo
before being written down - commands are reproducible. Where something is
*not* directly verifiable from what's available locally, that's stated
explicitly rather than assumed.

## 1. Primary sources - stated explicitly

**Old Testament source language:** Hebrew (Masoretic Text tradition).
`engwebu_usfm/00-FRTengwebu.usfm` does not name a specific critical edition
(e.g. Leningrad Codex/BHS) - WEB's OT lineage runs through the 1901 ASV,
which itself follows the Masoretic tradition standard for that era. **Not
independently verified which specific digital Masoretic edition backs
WEBU's tagging** - if that precision is ever needed, it would have to be
requested from eBible.org/Haiola directly; not claimed here without
evidence.

**New Testament source language:** Greek. Per `00-FRTengwebu.usfm` itself:
*"The World English Bible is an update of the American Standard Version
(ASV)... The New Testament was updated in places to conform to the
Byzantine Majority Text reconstruction of the original Greek manuscripts."*
This is explicit, quotable, primary documentation from WEBU's own front
matter - cite this, not a general assumption.

**The interlinear bridge to those primary languages, as tagged in this
corpus:** every `\w word|strong="H1234"\w*` (Hebrew) / `strong="G1234"`
(Greek) tag in `engwebu_usfm/*.usfm`. This is WEBU's own claim about which
Hebrew/Greek lexeme each English word translates - **see §3 for why this
claim is only partially reliable**, verified directly, not assumed.

**Key Hebrew lexemes for this pass** (standard Strong's numbers,
confirmed present in this corpus - see §2):
- `H3068` יהוה (YHWH, the Tetragrammaton) - traditionally rendered `LORD`
- `H3069` - the special Masoretic pointing used when Adonai directly
  precedes YHWH (to avoid saying "Adonai Adonai" aloud) - traditionally
  rendered `GOD` in the compound `Lord GOD`
- `H136` אֲדֹנָי (Adonai, "Lord/my Lord") - standalone, not preceding YHWH
- `H430` אֱלֹהִים (Elohim, "God") - **confirmed absent from this corpus's
  tagging entirely, see §3.2 - cannot be used for this pass**
- `H3050` יָהּ (Yah, poetic short form, e.g. "Hallelu-**Yah**")

**Key Greek lexemes:** `G2962` κύριος (kurios, "lord/Lord") - used for both
generic human lordship *and* as the standard NT/Septuagint rendering of
YHWH; `G2316` θεός (theos, "God"). **See §3.3 - kurios's dual use is a real
disambiguation problem, not a tagging bug.**

## 2. WEBU's own documented convention (quoted directly, not paraphrased)

Found via footnotes at the first occurrence in each OT book (Genesis,
Exodus, Leviticus, Numbers, Deuteronomy each carry their own copy):

> "When rendered in ALL CAPITAL LETTERS, 'LORD' or 'GOD' is the translation
> of God's Proper Name (Hebrew 'יהוה', usually pronounced Yahweh)."

And at Ezekiel 2:4 specifically (`Lord` + `GOD` compound):

> "The word translated 'Lord' is 'Adonai.'"

And at Genesis 1:1 specifically:

> "The Hebrew word rendered 'God' is 'אֱלֹהִים' (Elohim)."

**Also directly relevant, from `00-FRTengwebu.usfm`:** *"This is the
'updated' version of the World English Bible which uses 'LORD' or 'GOD' in
place of 'Yahweh' or 'Yah' in the Old Testament"* (from `copr.htm`). **This
means the pre-update 2013 `web.txt` baseline in this repo literally spells
out `Yahweh`** - confirmed: `web.txt` has 5,792 occurrences of the literal
string "Yahweh"; `rwb.txt` and the docm both have **zero**. RWB's divine-name
convention isn't a marginal stylistic choice - it's already a comprehensive,
completed replacement of thousands of instances of the old WEB's literal
transliteration, done at some point before this session's work began.

## 3. Gotchas - verified, not hypothetical

### 3.1 WEBU's own capitalization convention is the most reliable signal - not the Strong's tags

This is the load-bearing finding of this research pass, and it changes the
whole design of Pass 3.

**Verified:** in `engwebu_usfm/02-GENengwebu.usfm`, Genesis 1:1's "God" is
tagged `strong="H8064"` - which is not Elohim, it's the Strong's number for
**"heavens"** (a nearby word in the same verse). The book's own footnote at
that exact word correctly says *"The Hebrew word rendered 'God' is
'Elohim'"* - so the human-written footnote is right, but the automated
inline tag is wrong.

**This is not a one-off.** Checked every "God" occurrence in Genesis: none
are tagged `H430` (Elohim). Checked across the entire corpus
(`engwebu_usfm/*.usfm`): **`H430` never appears as a Strong's tag anywhere
in this WEBU build - zero occurrences, for a lexeme that appears roughly
2,600 times in the actual Hebrew Bible.** This is a specific, verifiable
gap in this corpus's automated word-alignment pipeline, not a general
tagging failure (`H3068`/YHWH tags 6,569+ times correctly on "LORD" and
appears reliable there).

**Conclusion: word-level `\w` Strong's tags cannot be trusted as the primary
signal for identifying Elohim/God occurrences in this pass**, and should be
treated with suspicion generally for high-frequency short words. This is
the same caution already flagged in the Test 70/71 plan's Phase 5 notes
("naive positional word-alignment will misattribute Strong's numbers") -
this research provides the first concrete, reproducible proof of it, not
just the theoretical risk.

**What to use instead:** WEBU's own **surface-form capitalization
convention** (`LORD`/`GOD` all-caps vs. `Lord`/`God` title-case) is a
deliberate *human translation choice*, explicitly documented in its own
footnotes (§2) - far more trustworthy than an automated inline tag. Pass 3's
primary method should be a capitalization-pattern census (reusing the exact
architecture already built for Tests 70/71's quote-mark census - see §5),
**not** Strong's-tag extraction. Strong's tags are usable only as a
secondary corroboration signal, and only for `H3068`/`G2962`/`G2316`, which
verified as reliable - never for `H430`, which isn't taggable in this
corpus at all.

### 3.2 `H430` (Elohim) is untaggable in this corpus - a hard limitation, not a design choice

Restated plainly because it matters: there is no reliable, tag-based way in
this specific WEBU build to distinguish "this 'God' is Elohim" from "this
'God' is something else." Any future tooling (this pass or Phase 5) must
not assume otherwise.

### 3.3 NT `kurios` (G2962) is genuinely ambiguous, not a tagging error

Unlike Elohim, this is a real linguistic ambiguity, not a corpus defect.
`kurios` is the NT's standard rendering of both **generic human lordship**
("sir," "master," "owner") *and* the Septuagint/NT convention of using
`kurios` **for YHWH** when quoting or alluding to OT texts. Confirmed in the
data: `LORD` (all-caps, definitely-YHWH usage) appears once in the NT with a
`G2962` tag (an OT quotation carrying the capitalization through), while
`Lord` (title-case) carries `G2962` **637 times** for ordinary/ambiguous
usage. **Resolving which is which requires exegetical context (is this an OT
quotation, is it referring to Jesus, is it a human address), not a
mechanical rule** - flag this as inherently manual-review territory in the
NT, unlike the OT where capitalization does the disambiguating work
directly.

### 3.4 ALL-CAPS text can mean "this is an inscription," not "this is the divine name"

Verified real false positives in the current docm:

- **Exodus 28:36**: `'HOLY TO GOD.'` - text engraved on a priestly gold
  plate, rendered in caps as a *typographic convention for quoting an
  inscription*, unrelated to the LORD/GOD divine-name convention.
- **Acts 17:23**: `'TO AN UNKNOWN GOD.'` - Paul quoting a pagan altar
  inscription at Athens - same typographic convention, same false-positive
  risk.

**A naive regex for all-caps `GOD`/`LORD` will flag both of these as
divine-name-convention hits. They are not.** Any census tooling for this
pass must be able to exclude (or at least separately bucket) inscription-
style quotations - likely via a manual exception list initially (there
won't be many), not an attempted general rule.

### 3.5 A stray, likely-missed conversion

**Romans 9:28**: `...because the LORD will make a short work upon the
earth."` - the **only** `LORD` (all-caps) left anywhere in the current
docm. This is Paul quoting Isaiah 10:22-23, carrying the OT's capitalized
convention into an NT quotation. Given RWB's conversion is otherwise 100%
complete (`web.txt` had 5,792 "Yahweh"s, `rwb.txt`/docm have zero), this
single leftover reads as an overlooked spot in an otherwise-thorough pass,
not a deliberate exception - **flag for operator review as part of Pass 3's
worklist**, don't silently "fix" it without confirming there isn't a
reason it was left alone (e.g. deliberately preserving OT-quotation
capitalization in the NT - possible, but unconfirmed).

### 3.6 Compound collapse: `"LORD God"` → `"God"` loses information

Verified: WEBU's Genesis 2:4 (`the LORD God made the earth...`) tags
**both** "LORD" and "God" with the same `H3068` - meaning the tagging
pipeline treats the two-word English phrase as one Hebrew concept (יהוה
אלהים, "YHWH Elohim"). RWB's rendering of the same verse is simply `God`
(one word) - confirmed via docm and `rwb.txt`. This is a **real, deliberate
simplification** (dropping "LORD" entirely from the compound, not
relabeling it), consistent with RWB's stated philosophy of avoiding literal/
complex renderings - but it means the distinct-name information (that two
different underlying words compose this phrase) is not recoverable from
RWB's text alone. Worth naming explicitly as a considered trade-off, not an
oversight, when the style rule is ratified (§6).

## 4. Baseline data (verified counts, not estimates)

| Surface form | `web.txt` (2013) | docm (current) | WEBU (current) |
|---|---:|---:|---:|
| `Yahweh` (literal) | 5,792 | 0 | 0 (see §2 - replaced by design) |
| `LORD` (all-caps) | - | **1** (Romans 9:28, §3.5) | 6,570 |
| `GOD` (all-caps) | - | **2** (both inscriptions, §3.4 - not real hits) | 314 |
| `Lord GOD` (compound) | - | **0** | present (subset of the 314/6,570 above) |
| `Lord` (title-case) | - | 1,990 | 1,020 |
| `God` (title-case) | - | 7,075 | (not cleanly countable - `H430` untaggable, §3.2) |

RWB has already executed a near-total (not 100% - see §3.5) replacement of
WEBU's `LORD`/`GOD`/`Lord GOD` convention with `God`, and uses `Lord` more
broadly than WEBU does (1,990 vs. 1,020) - the gap is presumably ordinary
human-lordship address ("my lord the king," etc.) plus whatever Adonai
usage RWB does preserve as `Lord`. **This needs verse-level sampling to
characterize precisely - not yet done, next step (§5).**

## 5. Methodology for the actual audit (not yet built/run)

1. **Build a new census tool** (`aeRWB/tools/web-diff`, same architecture
   as `census.mjs`/`docm-rwb-diff.mjs`) that, per verse, extracts WEBU's
   surface-form divine-name profile: counts of `LORD`, `GOD`, `Lord GOD`
   (as an adjacent pair), `Lord`, and (with the `H430` caveat noted) `God`.
   This is capitalization-pattern matching, not Strong's-tag extraction
   (§3.1) - reuses existing `parseBible`/pattern-matching infrastructure,
   no new parsing concepts needed.
2. **Cross-reference against docm**, verse by verse, the same way
   `docm-rwb-diff.mjs` already does for quote marks - but comparing
   *categories* (does this verse have a YHWH-per-WEBU occurrence, and if
   so what did RWB render it as) rather than exact text equality (RWB's
   wording legitimately differs, per the established policy throughout
   this project).
3. **Exclude known inscription false positives** (§3.4) via an explicit,
   reviewed exception list - start with the two found here, expect to find
   more while running the tool for real (Scripture has several other
   quoted-inscription passages).
4. **Produce a reviewable worklist**, not a bulk auto-classification -
   same "make every deviation apparent for outside review" discipline as
   R3/R4 and everything else in this project. Given `H430`'s untaggability,
   expect a meaningful "can't automatically classify, needs a human/AI
   read" bucket, not just clean pass/fail.
5. **Only after that data exists**, revisit the presentation-style
   determination in §6 with real, complete counts rather than the partial
   sample here.

## 6. Presentation-style determination - proposed, not unilaterally finalized

Based on the verified baseline (§4) and RWB's stated editorial philosophy
(avoids literal WEB/KJV-tradition constructs, avoids colloquialism, aligns
with KJV tradition otherwise - see `project_rwb_editorial_philosophy`
memory), the **already-overwhelmingly-in-effect** rule appears to be:

- **YHWH (`H3068`/`H3069`+`H3068` compound) → `God`** (not `LORD`, not `LORD
  God`, not `Lord GOD`) - the traditional all-caps convention is dropped
  entirely, and the "LORD God" compound is collapsed to a single word
  rather than kept as two.
- **Adonai (`H136`, standalone) and ordinary human lordship → `Lord`** -
  not yet verse-level confirmed which of these two categories accounts for
  more of RWB's 1,990 `Lord` occurrences; needs §5's audit before this half
  of the rule can be stated with the same confidence as the YHWH half.

**This is a proposal grounded in real data, not a final ruling** - given
its scale (thousands of verses already affected) and its theological
weight (this is exactly the kind of choice `project_rwb_editorial_philosophy`
says needs deliberate judgment, not silent inference), **this needs the
operator's explicit ratification** before being treated as the definitive
Phase 4 Pass 3 rule, the same way Option A needed sign-off for the more
structurally-invasive 2 Kings 19:13 fix. Once ratified, this becomes the
formal, documented style rule this section's title asks for - as clearly
defined as the docm's approved paragraph/character styles are today.

## 7. Connection to Phase 5 (Strong's numbers) - direct, not incidental

This research changes Phase 5's plan, not just Pass 3's:

- The Test 70/71 plan's existing Phase 5 note already recommended
  **verse-level** Strong's blocks over word-level alignment as "the first,
  safe milestone," flagging word-level misattribution as a *future risk*.
  **This pass found concrete, reproducible proof that the risk is already
  realized** in this exact corpus (`H430` untaggable; Genesis 1:1's "God"
  mistagged `H8064`) - Phase 5 should treat verse-level-only as a hard
  requirement for this WEBU build, not a cautious starting point to
  possibly relax later.
- Any Phase 5 tooling that wants to know "which words in this verse relate
  to the divine name" should reuse Pass 3's capitalization-based
  classification (§5) rather than trying to re-derive it from `\w` tags -
  Pass 3 is effectively doing Phase 5's divine-name groundwork now.

## Status

⚪ Research and this sub-plan done (2026-09-15). Tool-building (§5) and the
actual audit not yet started. Style-rule ratification (§6) pending operator
confirmation - needed before or alongside the audit, not strictly after.
