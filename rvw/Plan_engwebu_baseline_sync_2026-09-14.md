# Plan - engwebu_usfm baseline sync (Tests 70/71 and family) - 2026-09-14

**✅ Tests 70 and 71 are both closed and rebaselined (2026-09-15)** - see
their respective sections near the end of this file for full results.
**Next session starts here:** the small "Policy correction" note above
(Jeremiah 27:8 opening only - **27:22 was a false positive, corrected
2026-09-15, see that note**) - then Phase 4 (`rwb.txt` sync from the docm),
not yet started.

**Decision (operator, 2026-09-14, later same day):** `engwebu_usfm` (WEBU) is
now the **authoritative target** for quote-pattern instances (Tests 70/71),
superseding this plan's earlier "reference for sanity-checking, not a source
to copy from" framing (§1 item 3, §3). The docm and `rwb.txt` both need
editing to match WEBU's pattern - this is not a rebaseline-only fix. A new,
bigger goal was also set: **`rwb.txt` should become generated *from* the
docm** (with a clear diff/update record), replacing the current hand-tracked-
against-`web.txt` model. See the "2026-09-14 decision update" section below
for what this changes; §1-§6 below are left as originally written (the
investigation that led to this decision), not retroactively rewritten.

**Scope:** `C:\adaept\aeBibleClass\engwebu_usfm` was dropped in as a new
reference corpus. This plan covers (a) confirming what it is, (b) using it to
derive a real expected baseline for Test 70 (and investigating Test 71) instead
of the placeholder `0`, (c) updating `rwb.txt` (in the sibling `aeRWB` repo,
`C:\adaept\aeRWB`) with the punctuation it's missing, (d) a DRY mechanism so
future tests in the same family don't each need a bespoke one-off
investigation like this one, (e) a future step for Strong's numbers in the WIP
USFM export, and (f) the update process for the next time eBible.org ships a
new `engwebu_usfm` drop. **No baseline values or `rwb.txt` content are changed
by this plan itself** - it is the investigation and the proposed path; each
phase below is a separate, reviewable step.

## 1. Confirmed facts

1. **`engwebu_usfm` is genuine USFM with Strong's numbers.** Verified in
   `02-GENengwebu.usfm`: standard USFM markers (`\id`, `\c`, `\v`, `\p`, `\f`)
   plus inline word-tagging `\w word|strong="H1234"\w*` on nearly every word.
   87 files: 66 book files + front matter, glossary, `copr.htm`,
   `gentiumplus.css`, `keys.asc`, and the source `engwebu_usfm.zip`.
2. **Provenance confirmed via `copr.htm`:** "HTML generated with Haiola by
   eBible.org 12 Sep 2026 from source files dated 12 Sep 2026." World English
   Bible **Updated** (WEBU) - public domain (only the name "World English
   Bible" is trademarked; the text itself is free to copy/modify per
   `copr.htm`).
3. **RWB's relationship to this text**, per the operator and
   `C:\adaept\aeRWB\rvw\Code_review 2026-06-16.md` §6: the Study Bible
   (aeBibleClass's `.docx`/`.docm`) was originally based on **WEB**, not WEBU,
   with an early-2013-era `web.txt` snapshot (`https://openbible.com/textfiles/
   web.txt`, dated 2025-06-03 in the aeRWB repo) as the tracked baseline.
   `rwb.txt` is the **diverging RWB text**, hand-tracked against that `web.txt`
   baseline via the existing `tools/web-diff` harness (R3, see §4). WEBU
   (2026, this drop) is a **newer, related-but-different** public-domain
   edition (uses "LORD"/"GOD" in place of "Yahweh"/"Yah" - see `copr.htm`) -
   RWB is not a byte-for-byte descendant of WEBU, so WEBU is a **reference for
   sanity-checking**, not a source to blindly copy from.
4. **aeRWB (`C:\adaept\aeRWB`) is the live source-text repo.** `git log -1` =
   `524409d Fix I'm to I am` - this is the Test 55 fix from this same session
   (the "i'm" contraction correction), confirming `rwb.txt` is kept in sync
   with the aeBibleClass `.docm` by hand, verse-by-verse, not by an automated
   export yet (see §6, Phase 5).

## 2. The data: Test 70/71 pattern counts across all four texts

| Source | Test 70 (U+201C U+2018 U+201C) | Test 71 (U+201D U+2019 U+201D) |
|---|---:|---:|
| `engwebu_usfm` (WEBU, 2026-09-12) | **75** | **63** |
| Word `.docm` (current live RWB, via `RUN_THE_TESTS(70)`/`(71)`) | **73** | **15** |
| `aeRWB/rwb.txt` (hand-tracked RWB text file) | **0** | **35** |
| `aeRWB/web.txt` (2013-era WEB baseline) | **0** | **35** |

(`rwb.txt` and `web.txt` are **byte-identical** on both patterns - confirmed
via diff - meaning `rwb.txt` never inherited this punctuation from anywhere;
it's simply absent from the 2013 WEB lineage entirely.)

## 3. Reading the data

- **Test 70 is the clean case.** Docm (73) sits close to WEBU (75) - a 2-count
  gap is well within normal editorial drift (RWB has its own wording changes).
  This confirms the placeholder `0` was never realistic: nested open-quote
  triplets (a quote opening inside a quote inside a quote - e.g. Ezekiel's
  reported-speech chains, which account for 73 of WEBU's 75 hits) are a
  **normal feature of this text**, not a formatting bug. **Recommendation:
  rebaseline Test 70's `Expected1BasedArray` entry to 73** (the docm's current,
  already-verified count - see the Test 70 FAIL output already captured this
  session), matching the pattern used for every other stale-baseline fix this
  session (Tests 27-29, 34/35, 64).
- **Test 71 is the anomaly - do not blindly rebaseline.** The docm (15) is
  *lower* than both its own upstream text-family relatives: `rwb.txt`/
  `web.txt` (35) and WEBU (63). Every other comparison in this plan runs
  WEBU >= docm >= older-WEB; Test 71 inverts that. Two live hypotheses,
  neither confirmed:
  1. **Genuine editorial removal** - RWB's editing intentionally simplified
     some nested closing-quote sequences (e.g. added a space, or restructured
     the closing punctuation) somewhere between the 2013 WEB baseline and
     today's docm.
  2. **Detection loss, not content loss** - Word's AutoCorrect/AutoFormat (or
     a prior character-style/spacing fix, e.g. the smart-quote work referenced
     in Tests 66-69) may have inserted an invisible character (NNBSP, a
     character-style boundary, etc.) between the three quote glyphs in some
     instances, so `CountContraction`'s raw 3-character adjacency match
     silently stops matching even though the nested quotation is still there
     visually.
  **Action before rebaselining Test 71:** spot-check a handful of the 35
  `rwb.txt`/`web.txt` hit locations against the same verse in the current
  docm (verse references are recoverable via `aeRWB`'s `rwb.txt` line
  format - see §4) to determine which hypothesis holds. This is a short,
  bounded check, not a rebaseline-blind step.

## 4. DRY: extend the existing `tools/web-diff` harness (aeRWB), don't hand-roll a new one

`aeRWB/tools/web-diff` (R3, scaffolded 2026-06-16) already does almost exactly
what's needed: a **verse-keyed**, **deterministic**, **read-only-on-sources**
diff between `web.txt` and `rwb.txt` (`lib.mjs`: `parseBible` ->
`Map<"Book C:V", text>`; `diff.mjs` writes `diff/web-rwb-register.{jsonl,md}`
+ a summary with input SHA-256 provenance). Extending this, rather than
writing a one-off script in `aeBibleClass`, is the DRY move the operator asked
for (item 10) - it keeps the source-text tooling in the source-text repo, and
every future "is this baseline real" question (not just 70/71) reuses it.

**The same underlying question applies to the whole `CountContraction`
family in `aeBibleClass.cls`, not just 70/71** - worth naming explicitly so
future sessions don't re-derive this:

| Tests | Pattern family | Is a source-corpus baseline meaningful? |
|---|---|---|
| 52-65 | Literal contractions (`i'm`, `it's`, etc.) via `ContractionArrayU` | Mostly **should be 0** (editorial rule: uncontracted forms) - not this plan's concern, these are correctness checks, not content-frequency checks. |
| 66-69 | Space/NNBSP + smart-quote spacing bugs | **Should be 0** - pure formatting-error detection, unrelated to source content. |
| **70-71** | Nested-quote triplets (adjacent double+single+double) | **Not inherently 0** - a real typographic feature of narrative text with nested reported speech. This is the family this plan addresses. |

If a future test turns out to belong in the "not inherently 0, needs a real
corpus baseline" bucket like 70/71, it should reuse the mechanism built in
Phase 2 below rather than trigger a fresh investigation.

## 5. Phased plan

### Phase 0 - gitignore `engwebu_usfm` (DONE this session)

Added `/engwebu_usfm` to `aeBibleClass/.gitignore` with a comment pointing
back to this plan. Rationale: it's a large (~9 MB), third-party, periodically-
refreshed source drop - not something this repo should carry a bulk copy of
under version control (mirrors how `/Bibles`, `/Bible`, `/txt` are already
ignored for the same reason).

### Phase 1 - Add WEBU as a third parsed source to `web-diff`

In `aeRWB/tools/web-diff`, add a USFM-to-verse-map parser (new pure function
in `lib.mjs`, same shape as `parseBible`) that:

- Walks the 66 `NN-BBBengwebu.usfm` book files in `engwebu_usfm`.
- Strips `\w text|strong="H1234"\w*` down to plain `text` (regex extraction,
  no Strong's data kept at this stage - that's Phase 5).
- Strips footnotes (`\f + ... \f*`), cross-refs, and the `\+wh ... \+wh*`
  Hebrew/Greek inline markup.
- Emits the same `Map<"Book C:V", text>` shape `parseBible` already produces,
  so every existing diff/report function in `lib.mjs` works unchanged against
  it (this is the DRY payoff - one parser addition, zero changes to the
  diff/report logic).
- Book-name mapping note: WEBU's `\toc2`/`\h` gives the display name (e.g.
  "Genesis") - reuse `rwb.txt`'s existing book-name spellings (it already
  handles "1 Samuel", "Song of Solomon", etc. per `lib.mjs`'s header comment)
  so keys line up across all three sources without a separate mapping table.

**✅ Done (2026-09-15):** commit the derived `engwebu.txt` alongside
`web.txt`/`rwb.txt` in `aeRWB` only - not duplicated into `aeBibleClass`,
which stays the raw-USFM-drop side (gitignored per Phase 0). Same
reproducibility rationale as originally recommended, plus the operator's
concrete reason for deciding now: a committed `engwebu.txt` shows a normal,
reviewable diff in GitHub Desktop whenever a new WEBU drop lands - a
gitignored/regenerate-on-demand file would make that change invisible.
Implemented in `aeRWB` as `tools/web-diff/export-engwebu.mjs` (R7, `npm run
web.engwebu`); pointer + tracking story in that repo's root `README.md` and
`tools/web-diff/README.md`.

- ✅ **aeRWB commit:** `18440da` (2026-09-15) - "Add WEBU pattern census (R6)
  and committed engwebu.txt export (R7)", reviewed by operator in GitHub
  Desktop, pushed on Claude's explicit go-ahead.

### Phase 2 - Add a reusable "pattern census" mode

A new function in `lib.mjs` (e.g. `censusPattern(verseMap, pattern)`) that
counts occurrences of an arbitrary substring/regex per verse and returns
per-book and total tallies - i.e. the general form of what §2's table did by
hand for Tests 70/71. A thin CLI wrapper (`diff.mjs --census <pattern>` or
similar) runs it across all three parsed sources (`web.txt`, `rwb.txt`,
`engwebu.txt`) in one pass. This is the concrete DRY deliverable: the next
"is this aeBibleClass baseline real" question becomes one CLI run against
already-parsed sources, not a fresh round of ad hoc `grep`.

### Phase 3 - Resolve Test 70/71 baselines

- Test 70: rebaseline `Expected1BasedArray` (position 70) in
  `aeBibleClass.cls` to **73**, matching the already-verified docm count (§3).
- Test 71: run the Phase 3 spot-check (§3) first; rebaseline only once the
  15-vs-35-vs-63 gap is explained, not before.

### Phase 4 - Sync `rwb.txt` content

Per operator item 6/8: the docm already has correct nested-quote punctuation
for "many" of the Test 70 instances that `rwb.txt` is missing entirely (0 of
75/73). Correction order, most-trusted source first:

1. **Word docm is the primary source** for any verse it already has the
   correct punctuation in - it's the actively-edited master, and (per §1
   item 3) RWB has intentional wording differences from both WEB and WEBU, so
   copying from the docm preserves RWB's own edits.
2. **`engwebu_usfm` is the fallback** only for verses where the docm's
   nested-quote punctuation is itself still using the plain/straight form (not
   yet touched) - i.e. filling a genuine gap, not overwriting an RWB-specific
   editorial choice.
3. Mechanically: extend the Phase 1/2 tooling with a **verse-level diff-and-
   patch worklist** (reuse R3's existing `status: changed/added/removed`
   register shape from `web-rwb-register.jsonl`) that lists exactly which
   `rwb.txt` verses need the Test-70-pattern punctuation added, sourced from
   the docm where available. This should stay a **reviewed, verse-by-verse**
   correction (matching R3/R4's existing "make every deviation apparent for
   outside review" design goal in `tools/web-diff/README.md`), not a bulk
   find-replace - `rwb.txt` is the public-domain-destined deliverable, so
   silent bulk edits are exactly what R3/R4 exist to prevent.

### Phase 5 (future) - Strong's numbers into a `rwb.usfm` export

Ties into `basUSFM_Export.bas`'s existing page-range USFM exporter (WIP -
`ExportUSFM_PageRange`, `ConvertParagraphToUSFM`, etc., `src/basUSFM_Export.bas`).
Once `rwb.txt` is synced (Phase 4) and the export task matures past
page-range chunks to a full-document `rwb.usfm`, Strong's numbers can be
interleaved by matching each exported verse against the corresponding
`engwebu_usfm` verse's `\w word|strong="H1234"\w*` tokens.

**Flag now, don't solve now:** word-level alignment between two
independently-worded translations (RWB's edited wording vs. WEBU's) is not a
trivial 1:1 mapping - RWB verses will have different word counts/order than
WEBU in many places (that's the entire point of R3/R4's diff register). A
naive positional word-alignment will misattribute Strong's numbers. Recommend
starting with **verse-level** Strong's blocks (a `\v` line's WEBU Strong's set
attached as a whole, not word-by-word) as the first, safe milestone, with
true word-level alignment as an explicit stretch goal requiring its own
design pass (likely wants the same verse-keyed diff infrastructure from
Phase 1, extended with token-level alignment - e.g. a Levenshtein/LCS word
aligner over each verse pair).

### Phase 6 - Process for the next `engwebu_usfm` drop

When eBible.org/WorldEnglish.Bible ships a newer WEBU (`copr.htm`'s "source
files dated" line changes):

1. Download the new Haiola USFM zip from `https://eBible.org/engwebu/` (or
   `https://WorldEnglish.Bible`), extract to `C:\adaept\aeBibleClass\
   engwebu_usfm` (gitignored per Phase 0 - always safe to overwrite/re-extract
   in place).
2. Re-run `npm run web.engwebu` (in `aeRWB`) to regenerate the committed
   `engwebu.txt` - diff it against the previous committed version first, so
   the change itself is reviewable, same determinism guarantee R3 already
   provides for `web.txt`/`rwb.txt` today.
3. Re-run the Phase 2 pattern census for every test in the "not inherently 0"
   bucket (§4 table - currently just 70/71) and compare against the current
   `aeBibleClass.cls` baselines. A shift signals either a WEBU text update
   worth reviewing for `rwb.txt` (Phase 4, repeated) or just noise (RWB has
   already diverged past the point where a WEBU-side change matters).
4. No `aeBibleClass` baseline should be changed automatically from this - the
   census is a **signal to investigate**, matching this plan's own Test 70 vs.
   Test 71 lesson (§3): a corpus count is context, not an automatic answer.

## 6. Summary of what this plan changes right now vs. later

| Action | When |
|---|---|
| `.gitignore` entry for `/engwebu_usfm` | **Done this session** |
| This plan document | **Done this session** |
| `engwebu.txt` committed alongside `web.txt`/`rwb.txt` (Phase 1 decision) | ✅ Done (aeBibleClass `89959d3`, aeRWB `18440da`, both 2026-09-15) |
| Test 70 rebaseline to 73 | Next: quick, low-risk, data already confirmed |
| Test 71 spot-check + rebaseline | Next: bounded investigation, ~30 min |
| `web-diff` WEBU parser + pattern census (Phases 1-2) | Follow-up session, `aeRWB` repo |
| `rwb.txt` verse-by-verse sync (Phase 4) | Follow-up, after Phases 1-2 tooling exists |
| Strong's numbers in `rwb.usfm` (Phase 5) | Future, after the USFM exporter matures past page-range WIP |
| Recurring-drop process (Phase 6) | Documented now, executed whenever eBible.org next updates WEBU |

**Superseded by the 2026-09-14 decision update below** - Test 70/71 are no
longer simple rebaselines; see that section for the revised path.

## 2026-09-14 decision update - WEBU is authoritative for quote patterns

**What changed:** §3's framing ("WEBU is a reference for sanity-checking,
not a source to blindly copy from") is superseded for quote-pattern
instances specifically. The operator has decided `engwebu_usfm` is the
target both the docm and `rwb.txt` should match for Tests 70/71. This
means the gap is not a baseline-drift question anymore - it's an
**editorial task**: the docm is missing quote-pattern instances that WEBU
has, and those need to be added (not just measured).

### Revised targets

| Test | Docm (before edits) | WEBU (target) | Gap to close |
|---|---:|---:|---:|
| 70 (open-triple) | 73 | **75** | 2 instances |
| 71 (close-triple) | 15 | **63** | 48 instances |

Per the operator: "the closing quotes should follow the pattern of
engwebu_usfm" - i.e. Test 71's much larger gap (48, vs. Test 70's 2) is
expected and is the harder, main piece of this task, not a sign something
is wrong with the count.

### Revised Phase 3 - was "rebaseline to current docm count," now "edit the docm to match WEBU, then baseline"

The `Expected1BasedArray` values for Tests 70/71 should **not** be set to
73/15 (the docm's pre-edit count) - they should be set to **75/63** (WEBU's
count) once the docm has actually been edited to add the missing instances.
Setting the baseline first (as originally planned in §5 Phase 3) would just
make Tests 70/71 FAIL again the moment the WEBU parser (Phase 1) exists and
the editorial work begins - better to treat 75/63 as the target from the
outset now that the decision is made, and land the baseline change together
with the completed edits, not before them.

**Mechanical approach:** build the Phase 1 (WEBU verse-map parser) and
Phase 2 (pattern census) tooling first - not to re-litigate whether WEBU is
the target (decided), but because it's the only way to get a **verse-level
worklist** of exactly which 2 + 48 = 50 verses need editing. Hand-finding 50
verse locations by re-running `grep` per book (as done for the 3-verse spot
check earlier today) does not scale to this - the census tooling from §5
Phase 2 is now a prerequisite for the editorial work, not a nice-to-have.

### Revised Phase 4 - was "docm primary, WEBU fallback," now "WEBU primary for both docm and rwb.txt"

§5 Phase 4's correction order (docm trusted first, WEBU only fills genuine
gaps) is superseded for these two patterns: WEBU is now the pattern both
`rwb.txt` **and the docm** should be edited to match. This does not undo
RWB's own independent wording choices elsewhere (still governed by R3/R4's
existing diff-and-categorize discipline) - it's scoped specifically to
where these two quote-triplet patterns should appear, which is now an
"adopt WEBU's punctuation" decision, not a "preserve RWB's existing choice"
one.

### New goal - generate `rwb.txt` from the docm, not hand-track it against `web.txt`

This is a bigger structural change than the quote-pattern fix itself: **the
operator wants `rwb.txt` to become a generated artifact of the docm**, with
a clear, reviewable record of what changed on each generation - replacing
the current model where `rwb.txt` is hand-edited verse-by-verse to track
drift against the 2013 `web.txt` baseline (§1 item 3/4).

**Why this matters for the quote-pattern task specifically:** once the docm
is edited to match WEBU's quote patterns, that edit needs to reach `rwb.txt`
somehow. Manually re-typing 50 verses into `rwb.txt` by hand (the current
sync model) is exactly the error-prone, unscalable process this new goal is
meant to replace. Sequencing: the quote-pattern editorial work (revised
Phase 3/4 above) is the **first real test case** for a docm-to-`rwb.txt`
generator, not a separate, blocking prerequisite - but building at least a
minimal version of the generator before doing the 50-verse edit makes the
edit itself land in `rwb.txt` for free, instead of needing a second manual
pass.

**Shape of the generator (not fully designed yet, flagging the open
questions):**

- Needs to walk the docm's `VerseText` paragraphs (same bounded `For Each
  ActiveDocument.Paragraphs` pattern already proven throughout this
  session - `GetMarkerTotals`, `CountStyleParagraphsNotAligned`, etc., not
  a `Range.Find` scan) and emit one line per verse in `rwb.txt`'s existing
  format (`Book Chapter:Verse<TAB>text`, per `aeRWB/tools/web-diff/lib.mjs`'s
  header comment) - this is a **simpler, plain-text sibling** of the WIP
  `basUSFM_Export.bas` exporter, not the same thing (no USFM markers, no
  Strong's numbers - that's still Phase 5, later).
- "Clear record of updates and differences" (operator's words) - reuses
  the existing `aeRWB/tools/web-diff` harness's R3 diff-register concept
  (`web-rwb-register.jsonl`/`.md`), but pointed at **old `rwb.txt` vs.
  newly-generated `rwb.txt`** instead of `web.txt` vs. `rwb.txt` - same
  verse-keyed, word-level-diff, deterministic machinery, new pair of
  inputs.
- Open question for whoever builds this: does the generator run from
  VBA (writing `rwb.txt` directly to the `aeRWB` working copy, matching
  where `ImportThisDocumentFile`-style file I/O already lives in this
  codebase) or does it export an intermediate file that the existing
  Node.js `web-diff` tooling then consumes? Recommend VBA-side generation
  (the docm is the source of truth and already has all the paragraph-walk
  infrastructure) writing directly into `aeRWB/rwb.txt`'s format, with the
  Node-side harness doing only the diffing/reporting, matching each side's
  existing strengths - but this is a design decision for whoever picks up
  this phase, not settled here.

### Deferred - WEB Updates changelog review

Once the quote-pattern editorial work above is complete, the operator wants
a follow-up pass through `https://worldenglish.bible/webupdates.php` (the
WEB/WEBU update changelog) - explicitly **after**, not concurrent with, the
editing work here. Not scoped further in this plan; revisit when reached.

### WEBU parser - DONE, verified (aeRWB repo, file edits only - operator commits/pushes)

`aeRWB/tools/web-diff/lib.mjs` gained `USFM_BOOK_NAMES` (66-book mapping,
verified against `rwb.txt`'s exact spellings - "Psalm" singular, "Song of
Solomon"), `parseUsfmBook`, `mergeUsfmBooks`, `censusPattern`, plus
`census.mjs` (CLI) and 11 new unit tests (21/21 passing). Verified against
the real corpus - exactly matches every number already confirmed by hand
earlier this session (Test 70 pattern: 75 in `engwebu_usfm`, 0 in
`web.txt`/`rwb.txt`; Test 71 pattern: 63 in `engwebu_usfm`, 35 in
`web.txt`/`rwb.txt`). `npm run web.census -- "<pattern>"` writes a
per-verse worklist to `census/` (gitignored, regenerable, matching `diff/`'s
existing treatment) - this is the "method to see changes before approval"
for the WEBU side (operator point 1). **Not committed/pushed** - ready for
review in GitHub Desktop.

### Text equality is not automatic - must be explicitly defined (operator, 2026-09-14)

Before the docm/`rwb.txt`/WEBU three-way comparison can mean anything,
"identical" has to be a defined, deliberate comparison, not naive string
`===`. Risks specific to this project that would silently produce wrong
"changed"/"identical" classifications otherwise:

- **Encoding.** `web.txt`/`rwb.txt` are UTF-8 **with a BOM** (confirmed,
  `aeRWB/rvw/Code_review 2026-06-16.md` §4). VBA's `FileSystemObject.
  CreateTextFile(path, True, True)` "Unicode" flag writes **UTF-16LE**, not
  UTF-8 - using it for a docm dump would produce a file that "looks like
  text" but fails byte-for-byte comparison against `rwb.txt` for every
  single line. Any VBA-side writer must use `ADODB.Stream` with
  `Charset = "utf-8"` instead (see the new `ExportDocmVersesToRWBFormat`
  below).
- **Embedded control characters vs. the one-verse-per-line format.** A
  manual line break (`Chr(11)`) inside a verse's Word `Range.Text`, or the
  trailing paragraph mark (`Chr(13)`) that `Range.Text` includes at a
  paragraph's end, would corrupt the `Book C:V<TAB>text` line format if
  written verbatim (a stray embedded newline mid-line breaks every
  downstream line-based parser, including `parseBible`'s own "no-tab"
  anomaly detector). These must be explicitly collapsed to a space, not
  passed through - and this collapsing is itself a content decision worth
  documenting, since it means the dump is not a 100%-raw character capture.
- **Quote characters - the one thing that must NOT be normalized.** Unlike
  the above, curly vs. straight quotes, and the exact nesting-quote
  codepoints (U+201C/U+2018/U+201D/U+2019) that this whole investigation is
  about, must be captured **verbatim** - any cleanup routine reused from
  elsewhere in this codebase (e.g. `CleanTextForUTF8` in
  `basUSFM_Export.bas`, used for the unrelated USFM-export path) needs to be
  checked line-by-line for what it strips before being reused here, since a
  routine that's safe for USFM export could silently destroy the exact
  signal this comparison needs. (Checked 2026-09-14: `CleanTextForUTF8`
  only strips soft hyphens/zero-width chars/control characters below
  `Chr(32)` other than tab/CR/LF - it does not touch quote characters, so
  it's safe to reuse as a base layer, but this check needs repeating for
  any *other* cleanup function considered for this pipeline.)

**Working rule going forward:** any function in this pipeline that
transforms text must state, in a comment, exactly which characters it
touches and why - "cleaned" is not a sufficient description on its own.

### Revised summary table

| Action | When |
|---|---|
| `.gitignore` entry for `/engwebu_usfm` | **Done 2026-09-14** |
| This plan document + decision update | **Done 2026-09-14** |
| `web-diff` WEBU parser + pattern census (Phases 1-2) | ✅ Done - aeRWB `18440da`/`9afdc62` |
| `engwebu.txt` generator (committed derived text, not `rwb.txt` yet) | ✅ Done - aeRWB `18440da` (see Phase 1 decision above; `rwb.txt` sync itself is still Phase 4, not started) |
| Docm edits: 2 instances (Test 70) + 48 instances (Test 71) to match WEBU | ✅ Done - aeBibleClass `0b60a35` (Test 70), `7cfc01d`/`97581a4`/`f4eae9e`/`e789957` (Test 71 batches) |
| Rebaseline Tests 70/71 to 75/63 | ✅ Done - aeBibleClass `0b60a35` (70→75), `e789957` (71→63) |
| Strong's numbers in `rwb.usfm` (Phase 5) | Future, unchanged - after the USFM exporter matures |
| Recurring-drop process (Phase 6) | Documented, unchanged |
| WEB Updates changelog review (`webupdates.php`) | Deferred - after the quote-pattern editorial work is complete |

## ✅ Minimum edit test (do this first, 2026-09-14) - the model for every other edit

Before touching all 50 worklist verses, do the smallest possible one first -
the 2-verse Test 70 worklist - as an end-to-end proof of the whole
pipeline (docm edit -> VBA export -> Node census -> worklist) before
committing to the much larger 48-verse Test 71 pass. **Use this exact
process for every subsequent verse/pattern edit, both patterns.**

### Task

- [x] **Jeremiah 19:7** - docm now reads `"'"I will make the counsel...`
      (open-double + open-single + open-double), matching WEBU exactly.
- [x] **Jeremiah 27:8** - docm now reads `"'"It will happen...` (3 marks:
      open-double + open-single + open-double). **Editorial call made:** the
      operator chose the **minimum match** (3 marks, satisfies the Test 70
      pattern check), not the full 4-mark WEBU fidelity match (`"'"'`) this
      section flagged as the recommended-but-optional alternative. Deliberate
      choice, not an oversight - noted here per this section's own "make it
      deliberately" instruction. "says God" (not WEBU's "says the LORD")
      correctly left alone, per RWB's own Yahweh/LORD->God convention.
- [x] Document saved (VBA code also exported to `src/`).

**🟡 Policy correction (2026-09-15, supersedes the "minimum match" call above):**
the operator wants WEBU treated as the **exact punctuation baseline** going
forward - full character-for-character fidelity, not "whatever's minimally
sufficient to pass the automated pattern test" - so that i18n/web/mobile
tooling built later never has to special-case a "cosmetic" divergence
between RWB and its WEBU source. Two concrete, not-yet-applied follow-ups
this reveals, both **outside this 2-verse pilot's original scope** (queued
for the Test 71 batch pass rather than reopening this pilot):
- ⚪ **Jeremiah 27:8 opening** - docm has 3 marks (`"'"`); WEBU has 4
  (`"'"'`, an extra `'` right before "It"). Needs the missing `'` inserted
  to match WEBU exactly. **Verified 2026-09-15 by tracing the full v2-22
  nesting** (4 levels: L1 `"` v2, L2 `'` v4, L3 `"` v4, L4 `'` v5, all
  closing together at the end of v11, already fixed in the Test 71 batch) -
  v8's 4-mark cluster is WEBU *restating* all 4 already-open levels at a
  paragraph break, not opening 4 new ones needing separate closes. This is
  a real, confirmed gap and WEBU's own punctuation here is internally
  balanced, not a bug.
- ~~⚪ Jeremiah 27:22 closing~~ **❌ Not a real gap - corrected 2026-09-15.**
  Originally recorded here as docm `...to this place.'"` (2 marks) vs. WEBU
  `...to this place.'"'` (3 marks) - that WEBU transcription was **wrong**,
  a transcription error made without checking the raw USFM source. Re-
  verified directly against `engwebu_usfm/25-JERengwebu.usfm`: WEBU actually
  ends `...to this place.'"` - **2 marks, identical to docm.** No edit
  needed; the two closes required at that point (the L4-chain's second
  `'` opened mid-v22, and L1''' opened back at v16) are already both
  present. Left struck through rather than deleted, per this doc's
  progressive-history convention - the lesson (verify claims against the
  raw source before acting on them, not just against an earlier turn's own
  output) is worth keeping visible.

**🔴 Known blind spot (2026-09-15, the real gap above is not currently
tool-detected):** Test 70/71's census is a **presence check per verse**
(`docm has the 3-char pattern somewhere` / `WEBU has it somewhere`), not a
**nesting-structure check**. Jeremiah 27:8 has the pattern in *both* sources
(3 marks in docm, 4 in WEBU), so the Editorial worklist correctly excludes
it - "worklist empty" only ever meant "pattern present everywhere WEBU has
it," never "docm's nesting depth matches WEBU's." Nothing today would catch
this automatically; it's tracked only as the ⚪ item above until closed.
**What would close it** (considered and rejected/accepted below -
see the "4-mark test?" discussion in the architecture assessment's prep
list): not a new fixed-pattern `aeBibleClass` test (doesn't generalize past
this one depth or this one language's marks, doesn't localize to a verse,
wrong layer - content-fidelity vs. document-hygiene); instead, a
per-language-parameterized nesting-structure checker in `aeRWB` - **not** a
full text-equality diff (RWB's intentional wording divergence from WEB/WEBU
is out of scope for this, by design - see the prep-checklist item below).

### ✅ Verification process (the model for every other edit) - all four steps passed 2026-09-15

1. ✅ `RUN_THE_TESTS(70)` → **75** (was 73).
2. ✅ `ExportDocmVersesToRWBFormat` re-run → same `31053`/`46`/`3` totals as
   before - nothing broke elsewhere.
3. ✅ `npm run web.census -- "$(printf '“‘“')"` in `aeRWB` (file edits only -
   operator reviewed/pushed) → `docm-verses.txt` hits = **75** (was 73),
   editorial worklist **empty** (0 WEBU verses unmatched).
4. ✅ `Expected1BasedArray` position 70 rebaselined 73 → 75 in
   `aeBibleClass.cls` (already present in the exported code; confirmed
   correct only after steps 1-3 above passed).

**Provenance (freshness marker for this snapshot):** `rpt/docm-verses.txt` as
committed in `aeBibleClass` `0b60a35` (2026-09-15). The census counts above
(75/75, empty worklist) are only guaranteed current as of that commit - if
`git log -- rpt/docm-verses.txt` shows anything newer, re-run steps 2-3
before trusting these numbers.

### ✅ Applying this model to the remaining 48 Test 71 verses - Done 2026-09-15

Same four-step process, scaled up: edit all 48 verses from
`census/201D-2019-201D-worklist.md` - **per the 2026-09-15 policy correction
above, always take full WEBU character-for-character fidelity, not the
minimum needed to pass the pattern test** (superseding this section's
original "decide fidelity-vs-minimum per verse" framing) - then run the same
verification sequence once at the end (`RUN_THE_TESTS(71)` = 63, re-export,
re-census confirms 63 hits and an empty worklist, then rebaseline position 71
to 63) rather than one verse at a time - the two-verse Test 70 pass is the
proof the pipeline works; the 48-verse Test 71 pass is the same process at
scale, not a new process.

**Result:** all 48 verses done in four batches (2-verse spike + 10 + 10 + 10
+ 16), each independently verified. Several turned out to need a **mid-verse
insert** rather than an end-of-verse append (the WEBU close lands before
trailing narration, not at the verse's literal end) - e.g. 1 Kings 12:24,
2 Chronicles 11:4/34:28, Isaiah 38:8, Mark 7:11 - confirming the plan's
original caution that WEBU's actual nesting depth/position varies
verse-to-verse and must be checked individually, not assumed.

- ✅ `RUN_THE_TESTS(71)` → **63** (exact match with WEBU; was 15).
- ✅ `ExportDocmVersesToRWBFormat` re-run after every batch → same
  `31053`/`46`/`3` totals throughout - nothing else broke.
- ✅ Final `npm run web.census -- "”’”"` in `aeRWB` → 63 hits,
  editorial worklist **empty**.
- ✅ `Expected1BasedArray` position 71 rebaselined **0 → 63** in
  `aeBibleClass.cls` (it had never been baselined past the original
  placeholder `0`, unlike Test 70's prior real value of 73).

**Provenance:** `aeBibleClass` `e789957` (2026-09-15) - the final batch
commit; `rpt/docm-verses.txt` and the rebaseline both landed there. If
`git log -- rpt/docm-verses.txt` shows anything newer, re-verify before
trusting this section.

**Still open, not part of this closed batch** (per the policy-correction
note above - queue for a future small pass, not blocking): Jeremiah 27:8's
opening mark, found during the Test 70 pilot but outside its scope, and not
census-detectable by either Test 70 or 71's exact fixed pattern (see the
"Known blind spot" note above). Jeremiah 27:22, originally also flagged
here, turned out to be a false positive on re-verification (2026-09-15) -
see the policy-correction note above.

## 2026-09-15 architecture assessment - i18n/web/mobile/docx/client-server pathway

**Status: 🔵 forward-looking / non-blocking.** Recorded for historical
reference per operator request. Nothing here gates the Test 70/71 WIP above -
this is deliberately a "spike" (see the note on that below), not a commitment.

### Why this came up now

The operator framed Test 70/71's WEBU-fidelity work as more than an English/
Word exercise: it's a first concrete step toward a longer-term goal of
keeping a **realistic Windows+Linux Bible-translation development pathway**
that doesn't have to route through the SIL/Wycliffe/Tyndale/Paratext
ecosystem for basic text tooling, while still being able to interoperate
with it (a stated future integration target is Paratext's mobile-app-
generation track). Context given for this:

- Word/`.docm` is explicitly named a **sunset track** - a pragmatic current
  editing surface, not the intended long-term home. The forward architecture
  spans **i18n, web, mobile, docx, and client-server**, not just this repo.
- A documented cultural friction point in Bible-translation tech: much of
  that scholarship gravitates to Linux/open-source tooling (Paratext,
  Haiola, Crossway) and is wary of Windows/Microsoft "lock-in" (traced by
  the operator to the Ballmer era); some prior open efforts in this space
  (BLINK, SILAS) are now abandonware - a live cautionary tale about
  single-maintainer Bible-tech tooling going stale.
- Licensing philosophy: [copy.church/explain/importance](https://copy.church/explain/importance/)
  argues (fetched and summarized 2026-09-15) that restrictively-copyrighted
  Bible translations create real access barriers - "block anyone who would
  benefit... but isn't able or prepared to pay," turn ordinary sharing into
  unintentional law-breaking, stunt community improvement/adaptation, and
  remain vulnerable to being "retracted at any time" if a rights-holder's
  position changes - and argues nearly-free digital distribution makes that
  restriction avoidable. The page doesn't mention AI/tech reuse directly,
  but the same argument extends naturally there. This is the direction the
  operator is leaning for RWB - consistent with anchoring on WEB/WEBU (both
  public-domain) as the punctuation/text baseline rather than a
  restrictively-licensed modern translation.

### Assessment

**Pros**
- A public-domain-first baseline (WEB/WEBU, and RWB's own transparent diff-
  register model - R3/R4) is inherently friendlier to reuse than a
  copyrighted translation would be: no licensing negotiation needed for
  offline apps, web tooling, or AI-assisted translation work - directly in
  the spirit of the copy.church argument above.
- The `aeRWB` tooling built this session (R3 diff register, R6 pattern
  census, R7 `engwebu.txt` export) is **already format-agnostic in design**:
  verse-keyed maps, provenance-hashed inputs, deterministic regeneration.
  None of that logic is Word/VBA-specific - it's a legitimate nucleus for a
  future web/mobile/server layer, not throwaway scaffolding.
- Getting WEBU-fidelity punctuation exactly right **once**, at the source,
  means every future consumer (web render, mobile render, docx export, USFM
  export) inherits correct Unicode quote-nesting instead of every platform
  independently rediscovering the same bugs - a real, generalizable payoff
  from work that looks like a narrow English/Word fix today.
- An automated **quote-nesting depth-balance checker** (walk a book's text,
  track open/close depth, flag anywhere it goes negative or fails to return
  to baseline) is cheap to build, is exactly the kind of thing code catches
  reliably that a human (or an unverified AI claim) skimming dense nested
  speech will not, and is language-agnostic - it generalizes directly to
  i18n QA, not just English WEBU-matching. **Reinforced 2026-09-15**: the
  original Jeremiah 27:8 finding was real (found by `grep`, invisible on a
  normal read-through), but the *paired* 27:22 claim recorded alongside it
  turned out to be an AI transcription error that went uncaught for several
  turns - only caught when directly re-tracing the raw source instead of
  trusting an earlier turn's own output. A mechanical balance-checker
  wouldn't have made that mistake in the first place.

**Cons / Risks**
- `.docm`/VBA is a dead end for a multi-client (web/mobile/server)
  architecture. Every hour invested deepening Word-specific automation
  (ribbon UI, VBA classes) doesn't port - risk of the "generate `rwb.txt`
  from the docm" model becoming too load-bearing before any real export path
  exists, which would make leaving Word *harder*, not easier, later.
- Building outside the Paratext/USFM/SIL ecosystem risks **reinventing
  infrastructure** that ecosystem has spent decades hardening: USFM edge
  cases, checks, terminology tools, and - critically for i18n - right-to-left
  and complex-script rendering, and per-language quotation conventions that
  differ from English/WEBU's. This plan's own English-only assumptions
  (hardcoded `USFM_BOOK_NAMES`, literal `“`/`‘` pattern constants) would all
  need to be revisited for a second language.
- The Windows-vs-Linux reputational friction is a real adoption/credibility
  cost *if* this is ever pitched to the wider Bible-tech community for
  collaboration or funding - independent of whether the underlying tooling
  (Node.js in `aeRWB`) is actually cross-platform already (it is).
- BLINK/SILAS abandonware is a direct precedent for the sustainability risk
  of a single-maintainer Bible-tech tooling effort - worth naming plainly
  since it applies here too.
- The stated Paratext mobile-app-generation integration is currently just an
  intention - no design work yet, so it's a real, undiscovered dependency,
  not something already de-risked by today's work.

**Benefits (if the longer path is pursued)**
- A genuinely open, cross-platform, permissively-licensed Bible text +
  tooling stack that doesn't require going through SIL/Wycliffe/Tyndale/
  Paratext for basic text-correctness work, while still able to interoperate
  with Paratext where useful (the stated mobile-app-generation plan).
- RWB positioned as AI/LLM-friendly reference data by design (public-domain-
  first), which the operator identifies as an increasingly relevant use case
  even though copy.church's own page doesn't make that specific argument.
- The existing R3/R4/R6/R7 tooling in `aeRWB` is a legitimate starting
  nucleus for a future verse-keyed API/service layer, not a rewrite-from-
  scratch situation.

**Feasibility verdict:** realistic as a **multi-year, non-blocking side
architecture goal** - not realistic as a near-term replacement for Paratext/
USFM tooling, and not something to design head-on right now. The current
WIP (Test 70/71, R3/R4/R6/R7) is genuinely reusable prep regardless of which
specific web/mobile/server stack eventually gets chosen, which is exactly
why it's a "no wasted steps" investment matching the operator's stated risk
tolerance (don't advance one step now to reverse three later). The main risk
to actively manage is scope creep into Word/VBA-specific solutions that
don't generalize - already mitigated today by keeping Word-specific code in
`aeBibleClass` (test harness) and the portable/reusable logic in `aeRWB`'s
Node tooling; keep that boundary deliberate as this continues.

### Architecture prep/planning tasks (non-blocking - track with the emoji legend)

- ⚪ Define a canonical, format-agnostic verse-text interchange schema (e.g.
  `{ book, chapter, verse, text, provenance }` JSON) that `aeRWB`'s
  `web.txt`/`rwb.txt`/`engwebu.txt` loaders could also emit - a stepping
  stone toward web/mobile/server consumption without committing to a
  specific database or API yet.
- ⚪ **Build a per-language-parameterized quote-nesting-structure checker**
  (2026-09-15, corrects/replaces two earlier framings of this item - see
  below) as a new `aeRWB` tool: walks a verse/book/corpus and validates
  quotation-mark *nesting structure* (balance, correct open/close
  alternation, correct depth at known reference points like Jeremiah 27:8),
  **not full verse-text equality** - RWB's intentionally divergent wording
  vs. WEB/WEBU is a separate, already-handled concern (R3/R4's diff
  register, reviewed at commit time), explicitly **not** something this
  checker should flag. Scope is punctuation-nesting fidelity only.
  - **Must be parameterized by a per-language quote-mark-set table**
    (ordered `[open, close]` pairs per nesting level, direction-aware), not
    hardcoded to `U+201C`/`U+2018`/`U+201D`/`U+2019` the way Test 70/71 (and
    any fixed-pattern test) necessarily is - see the language survey below
    for why a hardcoded English/WEBU mark set would be silently meaningless
    for other languages, not just incomplete.
  - Closes the immediate Jeremiah 27:8 blind spot (docm's 3-mark vs. WEBU's
    4-mark original currently looks "done" to the Editorial worklist because
    both sides merely *contain* the pattern) as one concrete, near-term use
    of the same general mechanism. (27:22 was also flagged here originally
    but turned out to be a false positive - see the policy-correction note
    above; not an example of this blind spot after all.)
  - **Considered and rejected as a `aeBibleClass` fix:** a new fixed 4-mark
    `Test 87`/`88` in `aeBibleClass.cls` was considered - rejected because
    (1) a fixed pattern only catches this one specific depth in this one
    language's mark set, not arbitrary nesting depth or a different
    language's marks, (2) it returns a document-wide count like Test 70/71
    already do, not a verse reference, and (3) nesting-structure validation
    against a configurable per-language rule set is a different concern
    than Tests 1-86's self-contained, English-hardcoded document-hygiene
    checks - the right layer is `aeRWB`'s verse-keyed tooling, not a new VBA
    test. (If a fixed English-only regression guard is ever still wanted
    alongside this, it must be **appended** as the next unused test number,
    e.g. `87`/`88` - inserting between 69 and 70 would force renumbering
    every test through 86 and break every historical "Test 70"/"Test 71"
    reference in this doc and elsewhere; `aeBibleClass.cls`'s test dispatch
    has no execution-order dependency between cases, so there's no
    technical reason to insert mid-sequence either.)

**2026-09-15 language survey - why this can't be a hardcoded English
pattern** (answers "in what cases does this punctuation not apply, and how
is it resolved elsewhere," asked while scoping the item above):

- **Different glyph system entirely:** Japanese/Chinese/Korean traditionally
  use corner brackets (「...」 primary, 『...』 nested), not curly quotes at
  all. A translation in these languages would show **zero** hits on Test
  70/71's exact codepoints forever, regardless of whether its own nesting is
  correct - a "PASS" there would be meaningless, not reassuring.
- **Different glyphs, same underlying problem:** French/Russian/Greek and
  much of continental Europe use guillemets « » (`U+00AB`/`U+00BB`) as the
  primary mark, often with curly or angle quotes nested inside; German/
  Polish use low-high style „..." (`U+201E` opening). The adjacent-marks-at-
  a-nesting-boundary problem still exists, just with different codepoints
  this test doesn't reference.
- **Reversed polarity:** British English conventionally nests the *opposite*
  way from American/WEBU - single quotes outer, double inner. A pattern
  hardcoded to double-single-double checks for the wrong order if pointed at
  UK-convention text.
- **No stacked-mark nesting at all:** some literary and SIL-developed
  minority-language orthographies use a **quotation dash** (em-dash at the
  start of a reported-speech line) instead of paired quote marks - zero
  quotation-mark codepoints is *correct* there, by design.
- **RTL scripts** (Arabic, Hebrew): direction mirrors visually; convention
  may use guillemets (borrowed via French influence) or other direction-
  aware marks, not a drop-in case for a Latin-oriented codepoint sequence.
- **How Paratext/the USFM ecosystem actually resolves this:** each
  translation project configures its own quotation-mark table (ordered
  open/close pairs per nesting level, specific to that language), and
  validation checks nesting *structure* against that configured table, not
  a hardcoded mark set. This is the design the checker above should follow.
- ⚪ Research Paratext's project/USFM interchange and its mobile-app-
  generation track's expected input format, to scope the minimal adapter
  surface needed to feed RWB text into that pipeline without redesigning
  RWB's own workflow around it.
- ⚪ Explicitly decide and document RWB's target license posture (e.g.
  CC0/public-domain-style, per the copy.church argument, vs. some other
  permissive license) and record the rationale - currently only implicit
  (inherited from WEB's own public-domain status).
- ⚪ Prototype one non-English round-trip through the existing verse-keyed
  pipeline (even a toy/test language) to surface the Latin-script/English-
  only assumptions baked into today's tooling (book-name table, literal
  quote-character constants) before scaling to real i18n.
- ⚪ Explicitly evaluate whether the docm/Word ribbon layer stays the primary
  authoring surface long-term, or becomes just one of several editing
  front-ends over a shared backend text store - named directly because the
  operator has called Word a "sunset track."
- ⚪ Survey Haiola/Paratext-adjacent open tooling for reuse-as-a-library/
  service opportunities, to actively counter the "reinventing SIL's tooling"
  risk named above rather than discovering it the hard way.

### Note on methodology - the "spike" pattern

The operator has flagged the **minimum-edit-test / two-verse pilot before
the 48-verse batch** approach used for Test 70 above as something to keep
using generally, not just this once: prove a pipeline end-to-end on the
smallest possible real case before committing to the full-scale version.
Continue defaulting to this pattern for future batches of similar size/risk
(e.g. Test 71's 48 verses, and any future architecture work from the list
above that has a natural minimal-first-case shape).
