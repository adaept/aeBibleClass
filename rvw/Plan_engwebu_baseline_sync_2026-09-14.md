# Plan - engwebu_usfm baseline sync (Tests 70/71 and family) - 2026-09-14

**Next session starts here:** ["Minimum edit test (do this first, 2026-09-14)"](#minimum-edit-test-do-this-first-2026-09-14---the-model-for-every-other-edit)
near the end of this file - a 2-verse checklist (Jeremiah 19:7, 27:8) with
the exact before/after text and a 4-step verification process. This is the
model to repeat for the remaining 48 Test 71 verses once it's confirmed
working.

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

**Open decision (flag for operator, not decided here):** should the derived
plain-text `engwebu.txt` be committed alongside `web.txt`/`rwb.txt` (small,
public domain, gives the harness's SHA-256 provenance pinning and
`npm test` determinism for free - same treatment as `web.txt` today), or
regenerated on demand from a gitignored local `engwebu_usfm` drop (per Phase 0)
each time the harness runs? Recommend the former (commit the derived text,
not the raw USFM/Strong's files) for reproducibility, but this is the
operator's call.

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
2. Re-run the Phase 1 USFM parser to regenerate the derived `engwebu.txt` in
   `aeRWB` (if Phase 1's "commit the derived text" option was chosen - diff it
   against the previous committed version first, so the change itself is
   reviewable, same determinism guarantee R3 already provides for `web.txt`/
   `rwb.txt` today).
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

### WEBU parser - DONE, verified (aeRWB repo, file edits only per [[feedback-aerwb-no-autopush]])

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
| `web-diff` WEBU parser + pattern census (Phases 1-2) | **Next - now a prerequisite**, not just DRY infrastructure, since it's the only practical way to find the 50 verses needing edits |
| Minimal docm -> `rwb.txt` generator + diff register | **Next**, sequenced alongside/before the 50-verse edit so the edit lands in `rwb.txt` automatically |
| Docm edits: 2 instances (Test 70) + 48 instances (Test 71) to match WEBU | **Next - start with the minimum edit test below**, using the Phase 1/2 worklist |
| Rebaseline Tests 70/71 to 75/63 | **After** the docm edits land, not before |
| Strong's numbers in `rwb.usfm` (Phase 5) | Future, unchanged - after the USFM exporter matures |
| Recurring-drop process (Phase 6) | Documented, unchanged |
| WEB Updates changelog review (`webupdates.php`) | Deferred - after the quote-pattern editorial work is complete |

## Minimum edit test (do this first, 2026-09-14) - the model for every other edit

Before touching all 50 worklist verses, do the smallest possible one first -
the 2-verse Test 70 worklist - as an end-to-end proof of the whole
pipeline (docm edit -> VBA export -> Node census -> worklist) before
committing to the much larger 48-verse Test 71 pass. **Use this exact
process for every subsequent verse/pattern edit, both patterns.**

### Task

- [ ] **Jeremiah 19:7** - docm currently has `"I will make the counsel...`
      (one opening double-quote). WEBU has `"'"I will make the counsel...`
      (open-double + open-single + open-double). Insert `'"` immediately
      after the existing `"`, before "I will make" - two more opening
      marks needed to match WEBU.
- [ ] **Jeremiah 27:8** - docm currently has `"'It will happen...`
      (open-double + open-single - 2 marks). WEBU has
      `"'"'It will happen...` (open-double + open-single + open-double +
      open-single - 4 marks, one nesting level deeper than 19:7).
      **Decision needed while editing:** the Test 70 pattern only checks
      the *leading three* characters (`"'"`), which WEBU's actual 4-mark
      sequence already starts with - so inserting just **one** more `"`
      (giving `"'"It will happen...`) is enough to pass the test, but
      doesn't fully match WEBU's actual nesting depth. Inserting the full
      4-mark sequence (`"'"'`) matches WEBU exactly, per "closing quotes
      should follow the pattern of engwebu_usfm." Recommend the full
      4-mark match for fidelity, but this is a real editorial call, not a
      mechanical one - make it deliberately, not by default.
  - **Leave alone:** WEBU says "says the LORD" at 27:8 where the docm has
    "says God" - that's RWB's own independent Yahweh/LORD->God editorial
    convention, unrelated to this quote-pattern fix. Don't copy WEBU's
    wording here, only its quote punctuation.
- [ ] Save the document (not just export the VBA code - see item 9 of
      `Code_review 2026-09-14.md` for why this distinction matters).

### Verification process (the model for every other edit)

1. In Word's Immediate window: `RUN_THE_TESTS(70)` - expect **75** (up
   from 73). If it's not 75, stop and re-check the edit before going
   further - don't proceed to step 2 on a wrong count.
2. Re-export the docm dump: `ExportDocmVersesToRWBFormat` (full run, no
   `maxVerses` limit - the two edited verses could be anywhere in the
   document). Confirm the Immediate window still reports the same
   `31053`/`46`/`3` totals as before (or whatever the current baseline is)
   - a *different* skip/duplicate count would mean something broke
   elsewhere, not just the two intended verses changing.
3. In `aeRWB` (file edits only, do not commit/push -
   [[feedback-aerwb-no-autopush]]): re-run
   `npm run web.census -- "$(printf '“‘“')"` (or the equivalent for your
   shell). Confirm:
   - `docm-verses.txt` hits = **75** (was 73).
   - The editorial worklist in `census/201C-2018-201C-worklist.md` is now
     **empty**.
4. Only once all three checks pass, rebaseline `Expected1BasedArray`
   position 70 in `aeBibleClass.cls` from 73 to 75 (matches the pattern
   already used for every other rebaseline this session - edit the code,
   don't just accept the live PASS, per item 9's "verify against git, not
   a single live PASS" lesson).

### Applying this model to the remaining 48 Test 71 verses

Same four-step process, scaled up: edit all 48 verses from
`census/201D-2019-201D-worklist.md` (deciding fidelity-vs-minimum per verse
the way 27:8 required above, since WEBU's actual nesting depth may vary
verse-to-verse), then run the same verification sequence once at the end
(`RUN_THE_TESTS(71)` = 63, re-export, re-census confirms 63 hits and an
empty worklist, then rebaseline position 71 to 63) rather than one verse at
a time - the two-verse Test 70 pass is the proof the pipeline works; the
48-verse Test 71 pass is the same process at scale, not a new process.
