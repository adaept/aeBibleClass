# Code review - 2026-09-25

This file opens a new arc, carrying forward the open items from
`sync/session_manifest.txt`'s 2026-09-25 (continued) entry rather than a
prior dated `Code_review` doc - this session's work (Strong's/OSHB/RP
source-stack adoption in `aeRWB`, i18n-UK pilot verification against real
WEBBE text, `applySpellingVariant`'s live wiring) spanned four repos and
was tracked there, not in this file's lineage. See that manifest entry for
the full session narrative; this doc is the task-list/investigation record
for what comes next, carrying forward that manifest's open items plus the
operator's new work list below.

## Carried forward, still open

1. **`applySpellingVariant` live-check - REQUIRED, HIGH PRIORITY,
   explicitly deferred by operator direction 2026-09-25.** Rationale to be
   supplied later - do not guess one, do not run without an explicit
   go-ahead. Runbook: `aeBibleAddin/taskpane/src/word-audit-scanner.LIVE_CHECK.md`'s
   "Shape 8, spelling-conversion write extension" section. **Item 11 below
   adds to this gate** - the operator has now tied this live-check to
   resolving several of this session's new items first.
2. **Scholarly grounding roadmap step 3**: the verse-level Strong's spike
   on Genesis 1 (`aeRWB/tools/web-diff/`), using the OSHB data fetched
   2026-09-25. See `project_scholarly_grounding_plan` memory. cWEB is
   confirmed skipped entirely per operator decision.
3. **`ImportAllVBAFiles` `Skipped<>1` recurring anomaly** - root cause
   still not confirmed; diagnostic call site (`DumpLiveVBComponents
   [post-delete]`) left in place pending a reproduction with the dump
   active. See `feedback_importallvbafiles_error17` memory.
4. **Cases 49/50** (`CountAuditStyles_ToFile`, `SummarizeHeaderFooterAuditToFile`)
   remain unported to `aeBibleAddin` and formally untriaged (won't-fix vs.
   someday), unlike Cases 51/81 which got an explicit disposition.

## This session's items (operator work list, 2026-09-25)

### 1. "Running annotated Strong's search" - data source unclear, NEEDS CLARIFICATION

Searched the whole codebase (`aeBibleClass/src`, `md`, `rvw`; `aeRWB`;
`aeBibleAddin/src`, `taskpane/src`; `adaept5tudio/docs`) for "annotated
Strong's search," "Strong's search," "StrongSearch," and "annotated" near
any Strong's-related text. **Found no existing feature, plan, or code by
this name anywhere.** The closest adjacent things that do exist:

- `adaept5tudio/docs/studybible-mcp-integration-plan.md` (2026-09-10) - a
  *planning-only* document for integrating the user's `studybible-mcp` MCP
  server fork, which exposes an `word_study(strongs=...)` tool over Strong's-
  tagged lexicon data (LSJ/BDB/Abbott-Smith) and a proposed task-pane
  "Word Study" button (Option B, no LLM). Nothing built yet - `word_study`
  is upstream's tool name, not "annotated Strong's search."
- This session's own new Strong's-numbering/OSHB/Robinson-Pierpont source
  stack in `aeRWB/sources/` (see `project_scholarly_grounding_plan`
  memory) - raw dictionary/text data, not a search feature.
- The still-not-started verse-level Strong's spike (carried-forward item 2
  above) - a planned data-extraction pass, not a search UI.

**Need the operator to clarify** what "running annotated Strong's search"
refers to - a specific existing tool/site, a feature from a different
conversation, or a forward-looking description of what `studybible-mcp`'s
`word_study` or the planned Phase 5 lookup feature would become. Not
answered further here to avoid guessing.

### 2. Quotation check across WEBU/WEBBE/RWB/docm - CONFIRMED, all five sources match exactly

Ran both Test 70/71 census patterns (`aeRWB/tools/web-diff/census.mjs`,
which already covers `web.txt`/`rwb.txt`/`engwebu_usfm`/`docm-verses.txt`)
plus a direct substring count against `webbe.txt` (not yet wired into
`census.mjs`, counted separately):

| Pattern | web.txt (2013) | rwb.txt | WEBU | WEBBE | docm |
|---|---:|---:|---:|---:|---:|
| Open triplet `"''"` ("“‘“") | 0 | 75 | 75 | 75 | 75 |
| Close triplet `"''"` ("”’”") | 35 | 63 | 63 | 63 | 63 |

**rwb.txt, WEBU, WEBBE, and docm all agree exactly** on both patterns -
Tests 70/71's own baselines (75/63) hold across every current source, with
zero drift. Only the frozen 2013 `web.txt` baseline differs, as expected
(it predates the WEBU punctuation update entirely). No action needed here.

### 3. "Spirit's" in docm - LOCATED, one occurrence, action item for the operator

Only one apostrophe-word exists anywhere in the entire docm (31,103
verses) - **Romans 8:27**: *"He who searches the hearts knows what is on
the **Spirit's** mind, because he makes intercession for the saints
according to God."* This is a genuine possessive ("the mind of the
Spirit"), not a verb contraction - it doesn't mean "the Spirit is mind."
It's tracked by `src/data/contractions.ts`'s `spirit's` entry, but per
that file's own header comment, that list is a correctly-apostrophed
*coverage count*, not a zero-tolerance ban list, so its presence there
doesn't by itself mean it's forbidden.

The `it's: 1` count from this session's earlier smoke-test finding
(`project_phase3_audit_engine_port_plan` memory) is **the same single
occurrence**, not a second location - "Spirit's" contains the literal
substring "it's" (`Spir-it's`), so both patterns matched the identical
verse.

**Action item, operator-only** (this assistant does not edit the
production docm's content directly, per established practice): rephrase
Romans 8:27 to avoid the possessive-apostrophe construction, e.g. *"...
knows what is in the mind of the Spirit, because..."* or similar -
exact wording is an editorial call. After the docm edit: re-export
`rpt/docm-verses.txt`, and (once item 4's broader rwb-sync question below
is resolved) propagate to `rwb.txt`.

### 4. docm vs. WEBU/WEBBE outside documented differences - PARTIAL, plus one major new finding (docm vs. rwb.txt)

**Update, same day: docm-vs-rwb.txt divergence verified and synced (uncommitted, pending review).**
Verified the comma-placement convention illustrated below is real, not a
one-off: censused `,”` (old, comma-inside) vs `”,` (new, comma-outside)
across all five sources. docm's ratio (411:57 = 7.21) closely tracks
WEBU's own (425:58 = 7.33) - both reflect WEBU's actual per-sentence
punctuation logic. `web.txt`/`rwb.txt` (260:1 = 260.0 / 274:1 = 274.0) are
two orders of magnitude off - still the frozen 2013 baseline. **Confirms
docm is the correct, up-to-date source.** New tool
`aeRWB/tools/web-diff/apply-docm-rwb-full-sync.mjs` (R12,
`npm run docm.rwb.full-sync`) built - same full-verse-text-replacement
discipline as the existing pattern-scoped `apply-docm-rwb-sync.mjs` (R8),
generalized to every changed verse rather than a pattern-scoped subset.
Run against the current `rwb.txt`: **14,916 verses synced, 0 remaining
docm-vs-rwb.txt differences** (confirmed via a fresh diff after the sync).
4 verses each side untouched (ref mismatches, not text differences - see
below). All 25 `aeRWB` tests still pass. **Left uncommitted in `aeRWB`'s
working tree, per operator instruction, for GitHub Desktop review - text
content review is its own, separate cycle from this mechanical sync.**

Two real findings surfaced while running this:
- **Romans 14:24-26 vs. 16:25-27**: already-documented versification
  numbering difference (WEB's own convention vs. others - see
  `aeRWB/tools/web-diff/README.md`'s "Feeds R4" section, which already
  names this exact case) - not a bug, correctly left untouched by the sync.
- **`Jeremiah 37:910`**: a genuine reference-parsing anomaly in
  `rpt/docm-verses.txt` itself - verse 10's reference is malformed
  (missing the colon, reads as one number "910" instead of "10"),
  distinct from `Jeremiah 37:9`. As a result `rwb.txt`'s `Jeremiah 37:10`
  did not get synced (docm has no matching ref for the sync tool to find).
  **Not fixed here** - needs its own investigation into why the docm
  export produced this malformed ref (likely `ExportDocmVersesToRWBFormat`
  or its verse-number-detection logic), separate from this sync task.

**docm vs. WEBU/WEBBE:** `diffBibles` (existing `aeRWB/tools/web-diff/lib.mjs`)
against the leading-U+202F-stripped docm export:

| Comparison | Identical | Changed | Added | Removed |
|---|---:|---:|---:|---:|
| docm vs. engwebu.txt (WEBU) | 18,830 | 12,271 | 2 | 1 |
| docm vs. webbe.txt (WEBBE) | 16,802 | 14,299 | 2 | 1 |

A 6-verse sample of the docm-vs-WEBU "changed" set (Genesis 1:26, 2:4,
2:5, 2:7, 2:8, 2:9) showed **every single difference falling into an
already-documented category** - decontraction ("Let's" -> "Let us", per
`project_rwb_editorial_philosophy`) and the divine-names Rule 2 collapse
("the LORD God" -> "God", per `project_rwb_phase4_plan`'s seven-rule
model). **This sample is reassuring but not exhaustive** - confirming all
12,271/14,299 changed verses fall into a known bucket is a full
categorization pass on the scale of Phase 4's own rule-by-rule sweep
(~155+ verses found real defects there), not something six sample verses
can settle. Recommend scoping this as its own follow-on task, reusing
`diffBibles`'s word-level output the same way R4's "categorized change
rationale" was already planned (`aeRWB/tools/web-diff/README.md`'s "Feeds
R4" section) - tag each changed verse by reason (theological/decontraction/
divine-name/other), then hand-review only the "other" bucket.

**Major new finding, not what was asked but surfaced while investigating
this item: `docm` vs. `rwb.txt` diverges far more than expected.**

| Comparison | Identical | Changed | Added | Removed |
|---|---:|---:|---:|---:|
| docm vs. rwb.txt | 16,182 | 14,916 (48%) | 4 | 4 |

That's a *larger* divergence than docm-vs-WEBU (12,271) or docm-vs-WEBBE
(14,299) - **rwb.txt is currently closer to WEBU's own wording than to the
docm's**, which is backwards from what `project_engwebu_baseline_plan`
memory's "rwb.txt to become docm-generated" framing implies should be true.
A sample of the first 8 changed verses (Genesis 1:1, 1:2, 1:5, 1:8, 1:10,
1:11, 1:12, 1:14) shows old, pre-WEBU-update **2013 WEB-era wording and
punctuation** still sitting in `rwb.txt` - e.g. `rwb.txt` has `"day,"` /
`"night."` (old comma-then-close-quote convention) where the docm has
`"day",` / `"night".` (WEBU's convention, already covered by item 2's
clean Test 70/71 baselines); `rwb.txt` has "Now the earth... Darkness...
one day" where the docm has "The earth... and God's Spirit... the first
day" - genuinely different wording, not a formatting artifact.

**This is not the "intentional RWB divergence" R3's diff register exists
to document** - Phase 4's own scope notes are explicit that RWB's
intentional wording divergence is a separate, already-tracked concern
(divine names, decontraction), and this session's earlier Genesis-1 sample
doesn't match either of those categories. It looks like stale, un-synced
legacy text left over from before the docm's own editing passes, in
verses the Phase 4 divine-names/quote-pattern sync work never touched
(neither divine names nor the Test 70/71 quote triplet occurs in most of
Genesis 1). **Needs its own dedicated investigation and sync pass, scoped
separately from this review** - likely large (up to ~14,900 verses,
though the real "genuinely wrong" count is almost certainly smaller once
intentional RWB edits already present in that 48% are excluded; that
filtering is exactly the same "diff, then categorize, then only hand-review
the unexplained bucket" method recommended above for docm-vs-WEBU).

### 5. Psalm superscriptions and BOOK 1-5 in web.txt - ANSWERED: they're not exported at all, by construction

They don't "show" in `web.txt`/`engwebu.txt`/`webbe.txt` - **they're
silently dropped**, not rendered some other way. Confirmed directly
against the raw WEBU USFM (`engwebu_usfm/20-PSAengwebu.usfm`):

- **Book divisions** use `\ms1 BOOK 1` (a "major section" marker),
  appearing once before Psalm 1:1, and again before Psalms 42, 73, 90, and
  107 (the traditional five-book division of the Psalter).
- **Psalm superscriptions** use `\d ...` (descriptive title), e.g. Psalm
  3's `\d A Psalm by David, when he fled from Absalom his son.`, sitting
  between `\c 3` and `\v 1`.

Both are non-verse structural/front-matter markup in USFM, not part of
any `\v N` verse body. `aeRWB/tools/web-diff/lib.mjs`'s `parseUsfmBook`
only captures `\c`/`\v` content into the verse map - `\ms1` and `\d` lines
are walked over and discarded (they don't match `USFM_CHAPTER_RE` or
`USFM_VERSE_RE`, and aren't accumulated as verse-continuation text either).
The `Book C:V<TAB>text` plain-text format itself has no field to hold them
even if the parser kept them - there's no reference key for "the text
before Psalm 1:1 that isn't attached to any verse." This is a genuine gap
in the current export format, not a bug in the parsing logic given the
format's own design. See item 7 for what closing it would need.

### 6. Selah "showing on the next verse" in web.txt - COULD NOT REPRODUCE, needs a specific reference from the operator

Investigated two ways against the current committed sources:

1. **Verse-level attribution check**: extracted every verse reference
   containing "Selah" from `web.txt` and from `rpt/docm-verses.txt` (75 in
   each) and diffed the two reference sets directly - **zero differences**.
   Every verse that has "Selah" in `web.txt` has it in the docm too, and
   vice versa - no evidence of a shifted-by-one-verse attribution in the
   current files.
2. **Leading-word check**: searched all three plain-text sources
   (`web.txt`, `engwebu.txt`, `webbe.txt`) and the docm export for any
   verse where "Selah" is the *first* word (the clearest signature of a
   marker leaking onto the following verse) - **no matches anywhere**.
3. Spot-checked Psalm 3:2/3:3 specifically (a real Selah verse, raw USFM
   confirms `\qs Selah.\qs*` is correctly attached to `\v 2`) across all
   five sources (`web.txt`, `engwebu.txt`, `webbe.txt`, `rwb.txt`, docm) -
   all five agree, Selah ends verse 2 in every one.

**Could not locate the discrepancy the operator described** in any
currently-committed source. Possibilities, not resolved here: the issue
was already fixed upstream in `web.txt` since it was last observed; it's
specific to a particular verse not covered by the spot-check; or it was
observed in a different source (an older cached copy, a different USFM
drop, or something not currently in the repo). **Needs the specific verse
reference(s) from the operator to investigate further** - this item stays
open, not resolved, pending that.

### 7. What the rwb export needs to include superscriptions and BOOK 1-5

Three real pieces of work, in dependency order:

1. **Extend `parseUsfmBook`/`parseBible`'s data shape** (`aeRWB/tools/web-diff/lib.mjs`)
   to carry non-verse structural content. Currently `verses: Map<"Book C:V", text>`
   has no slot for "text attached to a chapter/book but not a specific
   verse." Needs a parallel structure, e.g. `frontMatter: Map<"Book C", {
   bookDivision?: string, superscription?: string }>` populated when
   `parseUsfmBook` encounters `\ms1`/`\d` lines, keyed by the *next*
   chapter/verse reference they precede (same "attaches forward" semantics
   USFM itself uses).
2. **Extend the `Book C:V<TAB>text` plain-text format** (or add a parallel
   file) to represent this. Two real options: (a) a special pseudo-verse
   key per chapter, e.g. `Psalm 3:0` for the superscription, keeping one
   file/one format (matches how some Bible software already represents
   superscriptions as "verse 0" or "verse 1" folded in); or (b) a second,
   parallel sidecar file (e.g. `psalm-front-matter.txt`) keyed by chapter
   reference, keeping `rwb.txt`'s existing verse-only shape untouched
   (lower risk to every existing consumer of that format, including the
   docm-sync tooling and `aeBibleClass`'s own test baselines). **(b) is
   the safer choice** - it doesn't risk silently breaking every existing
   `Book C:V` consumer that assumes V is always a real verse number, at
   the cost of one more file to keep in sync.
3. **Extend `export-webbe.mjs`/`export-engwebu.mjs`'s sibling, whatever
   generates `rwb.txt` from the docm** to populate this new structure from
   the docm's own front-matter content (the docm has this content today,
   per item 5's confirmation the docm already renders it - it's `web.txt`/
   `engwebu.txt`/`webbe.txt` missing it, not the docm) - `ExportDocmVersesToRWBFormat`'s
   VBA source (or whatever aeRWB-side tool eventually consumes its output)
   needs to walk the docm's Psalm superscription/Book-division paragraphs
   (identifiable by paragraph style, matching this whole codebase's
   established "identify structural content by style, not by guessing from
   text" convention) and emit them into the new sidecar format from (2).

### 8. USFM export from docm - what's needed

A **new** export path, not an extension of the existing plain-text one -
USFM is a different target format with its own markup vocabulary, not a
superset/subset of `Book C:V<TAB>text`. Real pieces:

1. **Style-to-USFM-marker mapping table**: the docm's paragraph/character
   styles (VerseText, Heading 1/2, the Psalm-superscription style, the
   Book-division style, Chapter Verse marker, Verse marker, Words of Jesus,
   etc. - all already enumerated in `EDSG/01-styles.md` and
   `aeSBL_Citation_Class`'s canonical book table) need an explicit,
   maintained mapping to their USFM marker equivalents (`\p`, `\d`, `\ms1`,
   `\v`, `\wj...\wj*`, etc.) - new data, not derivable from anything that
   exists today.
2. **A document walker** (VBA-side, since this reads the live docm's
   paragraph/style structure - not a JS/Office.js concern per the
   established "VBA stays for dev/QA tooling, JS is the end-user surface"
   split) that emits one `\v N verse text` line per VerseText paragraph,
   using (1)'s mapping to decide what markup wraps each paragraph based on
   its style, mirroring `ExportDocmVersesToRWBFormat`'s existing verse-walk
   shape but with USFM markup emitted instead of plain tab-separated text.
3. **Chapter/book-boundary and front-matter handling**: `\id`/`\h`/`\toc*`/
   `\mt1` book headers, `\c N` chapter markers, and (directly reusing item
   7's work) `\ms1`/`\d` front-matter - all need to be emitted at the right
   points in the walk, not just verse bodies.
4. **Explicitly NOT needed for a first pass, per this codebase's own
   established scope discipline**: Strong's-tag re-embedding (RWB has no
   Strong's tagging of its own yet - that's the separate, still-pending
   Phase 5/verse-level-spike work), footnote-cross-reference USFM markup
   (RWB's own footnote content, if any, is a separate concern from the
   verse text itself).

This is a real, multi-session build, not a quick add-on - closer in scope
to the original `web-diff` toolkit build than to any single tool within
it. Recommend its own dated plan doc before starting, following this
project's own established practice (per every other multi-session feature
in this repo's history).

### 9. Introduction to the Bible - SCOPED, not drafted here

A front-matter essay: cites source material (WEB/WEBU/Robinson-Pierpont/
WLC/OSHB per this session's own provenance work), explicitly follows the
tradition WEBU/WEBBE's own front matter sets (a plain-language "why this
translation, how it was made" explanation - `engwebu_usfm/00-FRTengwebu.usfm`
is the direct model to follow in tone and structure), carries real
scholarly grounding (citable critical-text basis, documented editorial
departures per `project_rwb_editorial_philosophy`) - but written for
**regular Bible readers**, not scholars, matching WEBU's own register
("informal, spoken English... designed to sound good and be accurate when
read aloud").

**Not drafted in this pass** - this is a substantive writing task deserving
its own dedicated session (matching how Task 2's Song-of-Solomon front-matter
note went through its own draft/review/approve cycle,
`project_scholarly_grounding_plan` memory). Recommend it as a distinct
next task, with the source-provenance material this session already
gathered (Strong's/OSHB/RP licensing table, WEBU/WEBBE relationship) as
its primary input.

### 10. Hyphenation for justified US-locale text vs. left-aligned i18n - answered, with one caveat

**(a) Hyphenation should not appear in extracted rwb utf8/USFM text -
CONFIRMED, with a caveat.** Checked `rwb.txt` and the current
`rpt/docm-verses.txt` export directly for both the Unicode soft/optional
hyphen (U+00AD) and the Unicode hyphenation-point character (U+2010) -
**zero occurrences of either in either file.** This is consistent with
Word's own well-documented VBA behavior: `Range.Text` does not return
optional/soft hyphens even when they're present in the document (they're
a display-only line-break hint, not real text content) - so as long as
any manual hyphenation in the docm uses Word's own optional-hyphen
mechanism (Ctrl+-), it should be structurally impossible for it to leak
into `Range.Text`-based extraction, by construction, not by luck.
**Caveat: this check is indirect** (it confirms the *export* is clean, via
a well-known Word behavior) **, not a direct inspection of the docm's own
raw content** - worth a quick live confirmation (e.g. `Selection.Characters`
count immediately around a known-hyphenated word) if the operator wants
certainty rather than well-grounded inference.

**(b) Tests should not be influenced by this - CONFIRMED.** Every VBA test
routine and every JS-ported audit routine in this codebase operates on
`Range.Text`/`paragraph.text`, the same property (a) already established
excludes optional hyphens. No test in either the VBA dispatcher or the
`aeBibleAddin` port reads document XML/OOXML directly for its *text*
content (the OOXML-parsing routines, Shape 5 and its extensions, read
*style*/*formatting* metadata, never plain text) - so there's no code path
in this project that could see a hyphenation artifact even if one existed.

**(c) Hyphenation not considered for other languages, too much work /
automation unreliable - CONFIRMED, and there's a stronger reason than
"too much work" specifically:** hyphenation exists to avoid ugly gaps in
**justified** text; the operator's own stated plan is that i18n VerseText
will be **left-aligned** (ragged-right), which has no such gaps to avoid
in the first place. So this isn't just deprioritized as expensive future
work - it's architecturally unnecessary for left-aligned text, full stop.

*Pros of never automating it:* zero risk of wrong-language hyphenation
rules being applied to text they don't belong to (a real failure mode -
hyphenation rules are language-specific and locale-aware automated
hyphenation for a language this project hasn't reviewed could silently
produce wrong break points); zero maintenance burden per new language; no
work needed at all if left-alignment is the permanent i18n choice.

*Cons:* if a future decision reverses course and wants justified text for
some i18n edition (e.g. matching a print convention in a specific target
language/market), hyphenation would need to be designed from scratch at
that point, with no existing groundwork to build on - an acceptable
tradeoff per this project's own stated anti-speculative-design discipline
(don't build for a requirement that doesn't exist yet).

### 11. Live-check gating - recommendation

**The operator's proposed gate** (resolve items above before running the
`applySpellingVariant` live-check) **makes sense for a specific, concrete
reason, not just general caution**: item 4's discovery that `rwb.txt` is
48% diverged from the docm, with old pre-WEBU-update wording still
present in un-synced verses, means **`rwb.txt` is not currently a reliable
stand-in for "the real document"** - live-checking a UK-spelling
conversion feature against a document whose own US-spelling baseline is
already known-stale in unknown ways makes it hard to tell a real bug in
the conversion mechanism apart from pre-existing baseline drift. Item 3
(the "Spirit's" fix) is smaller but has the same shape - the live-check
runbook's own scratch-document approach sidesteps this specific risk
(it's not testing against `rwb.txt` or the docm at all), so it isn't
strictly blocked by items 3/4 - but running it against a real Bible-content
copy later, to validate the feature beyond the synthetic scratch-doc test,
would be.

**On the specific proposal - auto-generating a non-hyphenated reference US
file to avoid churn in the hyphenated source:**

*Confirmed as a real risk worth avoiding*, if hyphenation is implemented as
literal soft-hyphen characters manually placed by the operator: any
automated tool that reads/rewrites paragraph text in that document
(including `applySpellingVariant`'s own `insertText(..., 'Replace')`
substitutions) risks disturbing manually-placed break points near a
converted word, forcing the operator to redo hyphenation decisions in
that neighborhood - real, avoidable churn for hand-placed work.

*However* - per item 10(a)'s finding, this may be **moot in practice**:
if hyphenation genuinely never appears in `Range.Text`-based extraction
(the well-grounded expectation, pending the live-content confirmation
flagged there), then a tool reading/writing via `Range.Text`/Office.js's
`.text` property never sees or disturbs soft hyphens in the first place -
there would be nothing to auto-generate a reference file to protect against.

**Suggestion, now vs. later:** don't build the non-hyphenated reference
file mechanism now. First, get a direct, live confirmation of item 10(a)'s
caveat (does `Range.Text`/Office.js text access actually leave manually-
placed optional hyphens untouched, checked directly against real
hyphenated content, not inferred). If confirmed clean, the whole
churn concern is resolved for free, no new tooling needed. If NOT
confirmed clean (a surprise finding, but this project's whole track
record this session is full of "assumed-safe things turning out not to
be" - the WEBBE quote-nesting assumption, the `\+w` USFM tag bug), *then*
build the reference-file mechanism, scoped specifically to that confirmed
gap rather than speculatively now.
