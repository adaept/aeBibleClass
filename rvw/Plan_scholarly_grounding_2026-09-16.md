# Scholarly grounding plan - next-session task list, 2026-09-16

Operator request at end-of-session: capture four follow-on items as tasks
for next time, and connect them to a longer-term goal (RWB as a sound
scholarly work, fit for peer review and i18n). This doc is that task list
plus a first-pass framing - none of the four items below have been
researched yet; where I've noted a candidate answer, it's flagged
explicitly as unverified recollection, not a finding, per this project's
own "check the primary source before asserting" discipline (repeatedly
earned the hard way this session - see `Plan_pass3_divine_names_2026-09-15.md`
§3.1's Strong's-tag reassessment and §4/§6's retracted idiom rule for two
concrete examples of claims that didn't survive contact with real data).

## Task 1 - Investigate the remaining divine-name outliers/exceptions

Operator's phrasing: "5/6 remaining outliers" - **ambiguous, needs
clarifying next session.** `Plan_pass3_divine_names_2026-09-15.md`'s Task 3
closed with exactly **5** confirmed non-actionable exceptions (Exodus
28:36, Exodus 39:30, Acts 17:23, Zechariah 14:20 - quoted inscriptions;
Revelation 19:16 - compound title). If the operator means a 6th item
exists that wasn't tracked, or means something else entirely (e.g.
re-checking the 19 "of Armies" occurrences individually rather than
trusting the 19/19 pattern-match, or the "Yah" alternative's 17
occurrences), that needs to be pinned down first.

**What to actually do, once scope is confirmed:** re-open
`aeRWB/census/divine-names-worklist.md` (or re-run `npm run pass3.census`
if it's gone stale) and manually read each of the flagged verse(s)
directly against the docm - the pattern established this session
(Malachi 4:5, the "day of the LORD" idiom, the "of Armies" idiom) is that
confirming a pattern generalizes requires checking **every** instance, not
trusting a plausible-looking majority. Don't assume the 5 known exceptions
are the end of the story just because the aggregate math closed to 99.94% -
that was exactly the kind of premature confidence that produced the wrong
idiom rule earlier in this same investigation.

## Task 2 - "Song of Solomon" vs "Song of Songs" vs "Song": confirm the right export name

**Operator's position, stated directly:** the docm's own convention (and
the Hebrew-tradition standard, שיר השירים) is "Song of Songs." The common
English search term is "Solomon" (already supported as a search-box
alias). The SBL abbreviation standard - which RWB otherwise follows - uses
"Song." Given RWB should follow the SBL standard rather than blindly
inherit WEB's book-naming convention, **the export should show "Song,"**
not "Song of Solomon." Operator explicitly frames this as "I consider
using the SBL standard more correct than blindly following WEB."

**Why this needs investigation before implementing, not just accepting -
my honest, current-knowledge assessment (unverified):**

1. **The principle is sound.** Task 1 of Phase 4 (aeBibleClass `663c36e`)
   made `ExportDocmVersesToRWBFormat` override this book's name specifically
   to match `web.txt`/`rwb.txt`/WEBU's inherited "Song of Solomon" - purely
   so reference-keyed diffing against those external files would work, not
   because "Song of Solomon" was judged more correct. If RWB has its own
   SBL-grounded citation standard (`aeSBL_Citation_Class.cls`/
   `basSBL_Citation_EBNF.bas` - see `project_sbl_citation_class.md`), that
   standard - not WEB's inherited title - should be the actual source of
   truth for what RWB calls this book. The operator's reasoning here is
   the same shape as the "God does evil" vs. softened-reading principle:
   RWB is not obligated to blindly inherit WEB/KJV-tradition choices.

2. **Concrete open question: is "Song" the SBL *full* name or only the SBL
   *citation abbreviation*?** In the SBL Handbook of Style (as I recall
   it, not yet re-verified against the actual Handbook or against this
   project's own `aeSBL_Citation_Class.cls`), the distinction is normally
   three-way: full name in running prose = "Song of Songs"; short-form
   used only inside parenthetical citations = "Song"; WEB/KJV-tradition
   title = "Song of Solomon." If the export's `Book C:V<TAB>text` format is
   meant to hold **full** book names (every other one of the 66 canonical
   books uses its full name - "Genesis," "1 Corinthians," not "Gen," "1
   Cor"), then the SBL-correct full name to use would be **"Song of
   Songs,"** not the abbreviated "Song" - unless this project's own
   `aeSBL_Citation_Class.cls` already treats "Song" as this book's
   canonical (non-abbreviated) identifier for a documented reason. **Check
   the actual class/EBNF source before implementing either way** - this is
   exactly the kind of claim ("SBL uses X") that needs a primary-source
   check, not recollection.

3. **Concrete consistency risk if "Song" (abbreviated) is used while all
   other 65 books stay full-name:** `aeRWB`'s `web.txt`/`rwb.txt`/
   `lib.mjs`'s `USFM_BOOK_NAMES` and `aeBibleCitationClass.GetCanonicalBookTable()`
   all use full, non-abbreviated names for every other book. Introducing
   an SBL abbreviation for exactly one book, without a matching change
   everywhere it's cross-referenced, risks silently reintroducing the
   **exact same class of ref-matching bug** Task 1 was built to fix (the
   Psalms/Song-of-Songs mismatch that broke every Psalms comparison
   silently). Any fix here needs to either (a) update every place this
   book's name is used for ref-matching consistently, or (b) build real
   alias resolution (matching what "the search box" apparently already
   does) so canonical storage and user-facing search can differ safely
   without breaking tooling.

**Recommendation for next session:** read `aeSBL_Citation_Class.cls`/
`basSBL_Citation_EBNF.bas` directly to see what this project's own SBL
implementation already says about this book's name, cross-reference
against the actual SBL Handbook of Style's book-abbreviation table (not
memory), then decide the correct canonical full name vs. citation-only
abbreviation before touching the export again.

## Task 3 - Source Strong's numbers, Hebrew source text, and Greek NT source text properly

This session found concrete, reproducible proof that `engwebu_usfm`'s
embedded Strong's tags are **not usable** for this project's purposes
(`H430`/Elohim never tagged anywhere in the corpus; `H8064`/"heavens"
over-applied ~30x its real frequency, including onto unrelated words like
"In" and "God" - see `Plan_pass3_divine_names_2026-09-15.md` §3.1). This
was worked around for Pass 3 via capitalization-pattern matching instead
of tags, but that workaround doesn't scale to Phase 5's actual goal
(attaching real Strong's numbers to the RWB text for study/lookup
features). A better-sourced Strong's dataset is needed.

**Three things to research, none decided or verified yet:**

1. **Strong's numbering itself** - where is the best-maintained, current,
   clearly-licensed digital Strong's dataset? Candidates to check
   (recalled from general knowledge, NOT verified against current license
   terms or maintenance status - this is exactly the research task):
   - **OpenScriptures Hebrew Bible (OSHB)** - open, git-tracked
     morphological/Strong's tagging for the Hebrew Bible, based on the
     Westminster Leningrad Codex.
   - **STEPBible's TAHOT/TAGNT** (Translators Amalgamated Hebrew OT /
     Greek NT) - openly licensed (stated CC BY-style terms, verify
     exactly), purpose-built for Bible translation tooling, Strong's-linked.
   - Whatever dataset actually backs `engwebu_usfm`'s tags currently (to
     understand *why* it's broken here specifically - a tooling bug in
     eBible.org/Haiola's pipeline, vs. a bad source dataset - matters for
     deciding whether to report the bug upstream too).
2. **Hebrew original source text** - the Westminster Leningrad Codex (WLC)
   is the usual free/public-scholarly-use base text behind OSHB and most
   open Hebrew Bible tooling; confirm current availability, exact license,
   and whether it's the right critical-text choice for RWB's purposes (vs.
   BHS, which is not freely licensed).
3. **Greek NT source text** - WEBU's own front matter states the NT
   "was updated in places to conform to the Byzantine Majority Text
   reconstruction" (quoted directly in `Plan_pass3_divine_names_2026-09-15.md`
   §1). For consistency with WEBU's own stated basis, the natural
   candidate is a Byzantine/Majority-Text edition (e.g. Robinson-Pierpont
   Byzantine Textform) rather than the modern eclectic critical text
   (Nestle-Aland/UBS, which is copyrighted by the German Bible Society and
   not freely licensed) - **verify this reasoning and the actual license
   terms before treating it as decided.**

**License restrictions are a hard gate, not a detail** - per this
project's stated public-domain-leaning philosophy
(`project_i18n_architecture_vision.md`: "public-domain licensing per
copy.church"), any source adopted here needs its license checked and
recorded explicitly before any tooling depends on it, the same discipline
already applied to WEBU itself in this plan's §1/§2.

## Task 4 - Synthesis: a plan to strengthen RWB toward peer-review-ready, i18n-ready scholarship

Once Tasks 1-3 have real findings (not before), write a proper plan
connecting them to the existing longer-term threads already on record:

- **Phase 5 (Strong's-number attachment)** - `Plan_pass3_divine_names_2026-09-15.md`
  §7 already flags that Phase 5 must treat verse-level Strong's blocks as
  a hard requirement given this corpus's proven word-level unreliability,
  and that Pass 3's capitalization-based classification is effectively
  doing Phase 5's divine-name groundwork already. Task 3's better-sourced
  Strong's data is what would let Phase 5 actually proceed with confidence
  instead of working around a known-broken tag layer indefinitely.
- **i18n/web/mobile architecture vision** (`project_i18n_architecture_vision.md`)
  - the long-term goal of a Windows+Linux Bible-dev pathway independent of
  SIL/Paratext, with public-domain licensing throughout. Correctly-sourced,
  clearly-licensed Hebrew/Greek/Strong's data is a direct prerequisite for
  that vision, not a separate concern.
- **Peer-review readiness** - the specific things a scholarly reviewer
  would ask first: what critical text(s) underlie the translation, what
  Strong's/morphological data backs any lookup features, and whether
  editorial departures from source translations (RWB's own philosophy -
  `project_rwb_editorial_philosophy.md`) are documented with stated
  reasoning rather than left implicit. This session's `Plan_pass3_divine_names_2026-09-15.md`
  is itself a working example of the standard to hold future documentation
  to (primary-source citations, retracted claims kept visible not deleted,
  reasoning recorded for every departure) - that documentation discipline
  is as much a part of "peer-review-ready" as the source data itself.

**Not to be conflated:** RWB's own editorial departures from WEB (avoiding
literal-but-harsh renderings, decontraction) are deliberate theological/
stylistic choices, already documented and out of scope for this task list -
Task 3 is about the *scholarly infrastructure* (source texts, numbering,
licensing) underneath the translation, not revisiting those editorial
choices themselves.

## Status

⚪ Not started - task list only, recorded per operator request at
end-of-session 2026-09-16. Task 1's scope needs a clarifying question
answered first ("5/6" - which exact items). Tasks 2-3 are research tasks
requiring primary-source verification before any implementation. Task 4
depends on 1-3 producing real findings.
