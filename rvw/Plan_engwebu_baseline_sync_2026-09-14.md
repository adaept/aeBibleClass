# Plan - engwebu_usfm baseline sync (Tests 70/71 and family) - 2026-09-14

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
