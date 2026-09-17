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

**Addendum, 2026-09-16 (operator reassessment request): this is a
tag-metadata defect, not evidence of a deliberate literal-interpretation
choice - re-verified, not just re-asserted.** Two completely different
things live under "WEBU" and must not be conflated: (1) the translation
text itself - the actual English words the WEBU/ASV-lineage translators
chose over 20+ years, which genuinely does lean literal/formal-equivalence
(see `project_rwb_editorial_philosophy` for RWB's deliberate departure from
this, e.g. "evil"->"disaster") - real, human, defensible scholarship,
untouched by anything below; and (2) the inline `strong="H1234"` word-level
tags in this specific `engwebu_usfm` build - a separate interlinear-
alignment metadata layer, typically produced by a different (often
automated) process layered on top of a finished translation.

Re-checked the raw USFM directly (not just re-citing the earlier finding):
Genesis 1:1 tags **"In," "beginning-adjacent words," and "God" all with
`strong="H8064"`** - the Strong's number for "heavens," which appears
nowhere in "In the beginning, God." Corpus-wide, `H8064` is applied **149
times within Genesis alone**, while the actual word "heavens" appears only
about 5 times in that book - a ~30x over-application impossible under any
real per-word semantic alignment. `H430` (Elohim) still confirmed **0**
occurrences corpus-wide, versus `H3068` (YHWH) at **13,217** - reliable and
correctly placed. This pattern (a handful of codes cycling positionally
across unrelated words, wildly out of proportion to real lexeme frequency)
is the signature of a broken/naive automated word-to-tag projection, not an
interpretive decision - a deliberate "more literal reading of God's nature"
would show up in the **words chosen**, not in mislabeling "heavens" onto
"In," "and," and "God" in the same sentence, which asserts nothing
theological and carries no interpretive content at all. **Conclusion
unchanged, evidence stronger**: this specific corpus's word-level Strong's
tags remain unusable for identifying Elohim/God occurrences, and this
finding says nothing about WEBU's translation quality or its genuinely
literal-leaning wording choices - those live entirely in row one, verified
separately and already trusted elsewhere in this plan (e.g. §2's `LORD`/
`GOD` capitalization convention, which WEBU's translators documented
explicitly in their own footnotes and which this plan already treats as
reliable).

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

## 4. Baseline data - corrected 2026-09-16, with the method-error that produced the first version

**First-pass numbers (2026-09-15, since superseded - kept for the record,
not deleted, per this project's progressive-history convention):**

| Surface form | docm (current) | WEBU (first pass) |
|---|---:|---:|
| `LORD` (all-caps) | 1 | 6,570 |
| `GOD` (all-caps) | 2 | 314 |
| `Lord` (title-case) | 1,990 | 1,020 |
| `God` (title-case) | 7,075 | not cleanly countable |

**What was wrong with it:** two separate errors, found by trying to
reconcile the numbers instead of presenting them side by side and stopping.
(1) The WEBU-side counts included **every book in `engwebu_usfm`**, which
ships WEBU's full ecumenical set - Tobit, Judith, Wisdom, Sirach, Baruch,
1-4 Maccabees, etc. - none of which exist in the 66-book Protestant canon
`docm`/`rwb.txt` track. This inflated every WEBU count with divine-name
references from books that have no docm counterpart at all - an
apples-to-oranges comparison. (2) The `God` count (7,075) used `grep -c`
(counts *matching lines*, i.e. verses with at least one hit) while every
other count used total occurrences - not the same metric, not comparable to
the others in the same table.

**Corrected method:** reused `aeRWB`'s own `loadEngwebu()` (already
filters to the 66 canonical books via `USFM_BOOK_NAMES` - the same
function `census.mjs`/`export-engwebu.mjs` already trust), and counted
total occurrences consistently on both sides.

| Surface form | WEBU (66 canonical books) | docm (current) |
|---|---:|---:|
| `LORD` alone (not in `LORD God`) | 6,528 | 2 (see §3.7/§3.5 - both false/stray, not real hits) |
| `LORD God` (YHWH+Elohim, apposition) | 42 | 0 |
| `Lord GOD` (Adonai+YHWH) | 288 | 0 |
| `Lord God` (title+title - **docm's actual form of the Adonai+YHWH compound**, missed entirely in the first pass) | 0 | **584** |
| `Lord` alone | 851 | 2,017 |
| `God` alone | 3,994 | 10,187 |
| lowercase `lord` (human address, e.g. "my lord the king") | 256 | 275 (comparable - RWB is not collapsing this into the divine-name convention) |

### Reconciling the corrected numbers - a 4th rule found by checking real verses, not just aggregates

Checked 5,507 individual WEBU verses containing plain `LORD` against their
docm counterparts (not just totals) - surfaced a pattern the first pass
missed entirely:

| # | WEBU pattern | → docm | Basis |
|---|---|---|---|
| 1 | `LORD` alone | `God` | confirmed, e.g. Genesis 6:3 |
| 2 | `LORD God` (apposition) | `God` (collapsed to one word) | confirmed 6/6 sampled, Genesis 2:4-2:15 |
| 3 | **`LORD your/our/my God`** (YHWH immediately before a possessive+Elohim) | **`Lord your/our God`** (not `God your God`) | **new finding** - 608 occurrences confirmed in WEBU (Exodus 6:7, 8:10, 8:26-28, Genesis 27:20, etc.) |
| 4 | `Lord GOD` (Adonai+YHWH) | `Lord God` | confirmed, Genesis 15:2/15:8 |
| 5 | `Lord` alone (Adonai/vocative address) | `Lord` (unchanged) | confirmed, Exodus 4:10/5:22 (the "O Lord" vocative survives alongside a separate "the LORD" → "God" in the *same* verse) |
| 6 | `God` alone (Elohim) | `God` (unchanged) | trivial |
| - | ~~Idiomatic exception, at least once confirmed~~ **RETRACTED, 2026-09-16** | ~~e.g. `the day of the LORD` → `the day of the Lord`, Malachi 4:5~~ | Task 2's tool checked all 14 `the day of the LORD` occurrences directly: 13 render as plain rule 1 (`God`); Malachi 4:5 is a one-off inconsistency, not a rule - see Task 2 below |

**Arithmetic reconciliation** (canonical-book counts, rule 3 netted out of
both `God` and `Lord`):

| | Predicted | Actual (docm) | Gap |
|---|---:|---:|---:|
| `God` | 6,528+42+288+3,994−608 = **10,244** | 10,187 | **57 (0.6%)** |
| `Lord` | 851+288+608 = **1,747** | 2,017 | **270 (13%)** |

The `God` column closes to well within noise once rule 3 is included (was
665/6% before finding it). `Lord`'s gap shrank from 878 (essentially
unexplained, ~85% of the observed total) to 270 (13%) - real progress, but
**not closed**, and not to be closed further by guessing at more aggregate
patterns. The remaining 270 needs the actual verse-level audit tool (§5),
which is exactly why §5 is a tool-building task and not something to
finish by hand-counting.

### 3.7 A second, unrelated bug found via this research: Psalms book-name mismatch

`docm-verses.txt` labels every Psalms verse **`"Psalms 1:1"`** (plural).
`web.txt`, `rwb.txt`, and `aeRWB/tools/web-diff/lib.mjs`'s own
`USFM_BOOK_NAMES` table (already tested, already used by every tool built
this session) all use **`"Psalm 1:1"`** (singular). This means **every
reference-keyed comparison between docm and web.txt/rwb.txt/WEBU silently
fails for all 2,461 Psalms verses** - `docm-rwb-diff.mjs` and
`apply-docm-rwb-sync.mjs` would report "not found" or simply never compare
Psalms at all, with no error raised. Never surfaced before now because
neither Test 70 nor Test 71's patterns happened to land in Psalms. This is
a bug in the VBA export routine (`ExportDocmVersesToRWBFormat` /
wherever its book-name list lives), unrelated to divine names except that
this is how it was found - **needs its own fix, tracked as Task 1 below,
before Pass 3's own verse-level tooling (which depends on ref-matching
working) is built.**

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

## 6. Presentation-style determination - final rule set, ready for ratification

Superseded twice since the 2026-09-15 two-part draft: once by the
four-rule model (§4, before real per-verse data existed), and now by the
**complete, verse-audited seven-rule model**, confirmed against all 9,016
WEBU divine-name occurrences via Task 3's census tool (99.94% clean match,
0 unexplained mismatches):

1. **YHWH alone → `God`** (or `Yah` - the poetic short form, H3050, which
   WEBU's own ALL-CAPS convention can't distinguish from full YHWH on the
   surface; both are legitimate RWB outcomes for the same WEBU `LORD`).
2. **YHWH in the `LORD God` apposition → `God`** (collapsed to one word).
3. **YHWH immediately before a possessive+Elohim (`LORD your/our/my God`)
   → `Lord your/our/my God`**, *not* `God` - avoids the nonsensical "God
   your God."
4. **Adonai in the `Lord GOD` compound → `Lord God`** (both words kept,
   not collapsed).
5. **Adonai alone (vocative/human address) → `Lord`** (unchanged).
6. **Elohim alone → `God`** (unchanged).
7. **`GOD`/`LORD of Armies` (recurring military-divine title, mostly
   Isaiah/Jeremiah/Amos) → `Lord God of Hosts`** (translating "Armies" →
   "Hosts" too) - confirmed across all 19 occurrences.

**Known, confirmed exceptions to the above (5 verses, all deliberate, not
gaps):**
- Exodus 28:36, Exodus 39:30, Acts 17:23, Zechariah 14:20 - quoted
  ALL-CAPS *inscriptions*, unrelated to the divine-name convention (§3.4).
- Revelation 19:16 - "KING OF KINGS AND LORD OF LORDS" is a compound title
  (Christ), not a substitutable divine name (confirmed against Revelation
  17:14's parallel, ordinary-title-case phrase).
- **Malachi 4:5 - operator-ratified, 2026-09-16**: kept as "the day of the
  Lord," the sole exception to rule 1's "day of the LORD → day of God"
  outcome (used at the other 13 occurrences). Reasoning: last verse of the
  OT, one of its most quoted/weighted verses, ~90% of surveyed
  translations use `LORD` (small caps) - judged to warrant RWB's
  KJV-tradition-alignment principle over its own internal default here.
- **Zechariah 14:20 - corrected 2026-09-17** (`Plan_scholarly_grounding_2026-09-16.md`
  Task 1 re-verification): the docm had rendered this inscription as "HOLY
  TO THE LORD" (unconverted) while the identical "HOLY TO THE LORD"
  inscription at Exodus 28:36/39:30 (same Hebrew word, `strong="H3068"`
  YHWH, in `engwebu_usfm`) was rendered "HOLY TO GOD." Operator-decided:
  applied rule 1 uniformly - Zechariah 14:20 now also reads "HOLY TO GOD,"
  matching the Exodus pattern. Fixed directly in `Blank Bible Copy.docm`;
  `rpt/docm-verses.txt` re-exported and diff-confirmed as the only content
  change. The bullet above (this section, 2026-09-15/16) is left as
  originally written per this project's progressive-history convention -
  this entry is the correction record, not a silent rewrite of it.

**Evolution for the record:** the 2026-09-15 two-part draft ("YHWH →
`God`," "Adonai/human lordship → `Lord`," undifferentiated by context) was
directionally correct but missed rule 3 entirely and had no real per-verse
evidence. The 2026-09-16 four-rule model (§4) added rule 3 and closed the
aggregate gaps to 0.6%/13%, but still rested on hand-sampling and included
one rule (the idiom exception) that turned out wrong when checked against
all 14 real occurrences. This seven-rule model is the first version
audited against every single occurrence in the corpus, not a sample -
0 unexplained verses remain.

**Ratification requested.** Given the scale (thousands of verses) and
theological weight of this choice (per
`project_rwb_editorial_philosophy`), **operator ratification is still
needed** - now *after* Task 3's real audit closed the gap (99.94% verified
clean, not an aggregate estimate), as this section originally called for.

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

✅ Research done (2026-09-15/16). ✅ Task 1 (Psalms/Song of Solomon
book-name fix) done and verified. ✅ Task 2 (audit tool) done, aeRWB
`ea6b79d`. ✅ Task 3 (review + fix) done, aeRWB `2b70402`,
aeBibleClass `3dedbfc` - **9,011/9,016 WEBU divine-name occurrences match
cleanly (99.94%), 0 unexplained mismatches, 5 confirmed exceptions.** The
seven-rule model (§6) is complete and audited against every occurrence,
not a sample. **Only remaining step: operator ratification of §6's final
rule set** (requested, not yet given).

✅ **Exceptions re-verified 2026-09-17** (`Plan_scholarly_grounding_2026-09-16.md`
Task 1) - all 5 confirmed exceptions individually re-read against
primary sources; one real inconsistency found and fixed (Zechariah
14:20's inscription now matches Exodus 28:36/39:30's, see §6's dated
bullet above), commit `<pending>`.

## Next-session tasks, in order

- **✅ Task 1 - Done 2026-09-16, aeBibleClass `663c36e`.** Fixed in
  `ExportDocmVersesToRWBFormat` (`basRWBTextExport.bas`), not the shared
  `aeBibleCitationClass.GetCanonicalBookTable()` (correct as-is for its own
  general-purpose citation use). **A systematic check of all 66 canonical
  names found a second mismatch beyond Psalms**: `"Song of Songs"`
  (aeBibleCitationClass) vs. `"Song of Solomon"` (`web.txt`/`rwb.txt`/
  WEBU) - both fixed with a narrow override in the export routine only.
  Verified: re-exported docm now emits `Psalm` (2,450 occurrences, 0
  `Psalms`) and `Song of Solomon` (0 `Song of Songs`); totals unchanged
  (31053/46/3/0); reference-keyed lookups against `rwb.txt` now succeed for
  both books (spot-checked Psalm 23:1, Song of Solomon 1:1).
- **✅ Task 2 - Done 2026-09-16, aeRWB `ea6b79d`.** Built
  `tools/web-diff/divine-names-census.mjs` (R9, `npm run pass3.census`),
  reusing `parseBible`/`loadEngwebu` per the plan. **First real run: 9,016
  WEBU verses have a divine-name occurrence; 8,971 match cleanly (99.5%),
  17 need a content fix, 28 need a human/AI read** - a dramatic
  improvement over the hand-sampled 13%/6% aggregate estimates above,
  because those estimates turned out to be measuring the wrong thing in
  two ways (both found only by running the real tool, not by refining the
  aggregate math further - exactly why §5 called for a tool, not more
  hand-counting):
  - **The `the day of the LORD` idiom (§4/§6) does not generalize.**
    Checked all 14 occurrences directly: 13 render as the ordinary rule-1
    `God` (Isaiah 13:9, Lamentations 2:22, Ezekiel 7:19, Joel 1:15/2:1/
    2:11/3:14, Amos 5:18/5:20, Obadiah 1:15, Zephaniah 1:7/1:8/1:14/1:18/
    2:2/2:3 - confirmed by direct inspection). Malachi 4:5 (the sole
    example the idiom was drawn from) is the outlier, not the rule -
    §4/§6's idiom row below is **retracted**, not merely unconfirmed.
  - **Rule 1 needed a second acceptable outcome: `Yah`**, the poetic short
    form of YHWH (H3050). WEBU's own ALL-CAPS convention folds "Yah" into
    the same `LORD` surface form as full YHWH (H3068) - undetectable from
    WEBU's text alone. 17 of the first 20 apparent "`God` missing" cases
    (mostly Psalms, e.g. "Praise the LORD!" -> "Praise Yah!") were this,
    not gaps.
  - The remaining **17 real mismatches**: 13 are a newly-discovered
    genuine defect (lowercase `lord` instead of `Lord` in the "Lord your
    God" phrase - Exodus 20:2, 10 Deuteronomy verses, Joshua 1:13, 1 Kings
    1:17, Jeremiah 40:2), 1 is Malachi 4:5's now-unexplained inconsistency
    (needs an operator decision), 1 is Acts 7:20 dropping "to God"
    entirely, 1 is the already-known Romans 9:28 leftover (§3.5), and 1
    (Revelation 19:16) looks like a classifier false-positive on the
    "LORD OF LORDS" title phrase rather than a real gap.
  - The **28 review cases**: 24 standalone `GOD` (mostly Isaiah/Jeremiah -
    likely the `H3069` Adonai+YHWH pointing without an immediately
    preceding "Lord," not yet individually confirmed) plus the 4 known
    inscription exceptions (§3.4).
  Full detail: `aeRWB/census/divine-names-worklist.md` (gitignored,
  regenerate with `npm run pass3.census`) and `tools/web-diff/README.md`'s
  R9 section.
- **✅ Task 3 - Done 2026-09-16, aeRWB `2b70402`, aeBibleClass
  `3dedbfc`.** Reviewed all 17 mismatches + 28 review cases individually:
  - **19 of the 24 standalone-`GOD` review cases confirmed as a genuine 7th
    rule**, not ambiguity: `GOD`/`LORD of Armies` -> `Lord God of Hosts`
    (translating "Armies"->"Hosts" too), verified across every occurrence.
    The remaining 5 plain standalone `GOD` cases confirmed as ordinary
    rule 6 (`God`, unchanged) - both folded into the tool, no longer held
    as "review."
  - **Revelation 19:16** confirmed a classifier false positive - "KING OF
    KINGS AND LORD OF LORDS" is a compound title (Christ), not a
    substitutable divine name (Revelation 17:14's parallel phrase uses
    normal title-case and needs no exception). Kept as a single named
    exception, not generalized from one example.
  - **Malachi 4:5 - operator-ratified exception**: kept as "the day of the
    Lord," not the ordinary "day of God" outcome the other 13 "day of the
    LORD" occurrences use. Reasoning: Malachi 4:5 is the last verse of the
    OT and one of its most quoted/weighted verses (the Elijah/messenger-
    of-the-covenant prophecy); a translation survey found ~90% use `LORD`
    (small caps), only one uses `God` - judged to warrant honoring RWB's
    KJV-tradition-alignment principle over its own internal idiom default
    here specifically. Recorded ref-keyed in the tool
    (`RATIFIED_EXCEPTIONS`) so it won't be re-flagged.
  - **The 15 real content defects were fixed in the docm**: 13 identical
    lowercase `lord`->`Lord` casing fixes (Exodus 20:2, 10 Deuteronomy
    verses, Joshua 1:13, 1 Kings 1:17, Jeremiah 40:2), Acts 7:20 (restored
    dropped "to God"), Romans 9:28 (leftover all-caps `LORD` -> `Lord`).
  - **Final verified result**: re-exported docm (still 31102/0/0), re-ran
    the census tool - **9,011/9,016 match (99.94%), 0 mismatch, 5 confirmed
    non-actionable exceptions** (4 known inscriptions + the Revelation
    19:16 title). The remaining §6 style-rule ratification is below.
