# Bug investigation + fix - Jeremiah 37:910 stray CVM marker - 2026-10-04

## Symptom

`rpt/docm-verses.txt` (produced by `basRWBTextExport.ExportDocmVersesToRWBFormat`)
contained a malformed reference line:

```
Jeremiah 37:910	 For though you had struck the whole army of the Chaldeans...
```

verse 10's text, filed under a verse number that doesn't exist in the
chapter (Jeremiah 37 has 21 verses). Jeremiah 37:10 itself was absent -
this reference was unreachable by any downstream per-verse tool (aeRWB
sync, WEBU/WEBBE review, etc.) under its correct address. First flagged
as a carried-over backlog item in `Plan_rwb_webu_review_checklist_2026-09-28.md`
§6c and `Plan_rwb_webbe_review_2026-10-04.md` line 104-105, not
root-caused until this session.

## Root cause - a genuine docm content defect, not a script bug

Inspected the live paragraph via Word COM (`word/document.xml`, confirmed
against the running `WINWORD.EXE` instance, document `Blank Bible
Copy.docm`). Jeremiah 37:10's `VerseText` paragraph had **three** runs
where it should have two:

```xml
<w:r><w:rPr><w:rStyle w:val="ChapterVersemarker"/></w:rPr><w:t>37</w:t></w:r>
<w:r><w:rPr><w:rStyle w:val="ChapterVersemarker"/></w:rPr><w:t>9</w:t></w:r>   <!-- stray -->
<w:r><w:rPr><w:rStyle w:val="Versemarker"/></w:rPr><w:t>10 </w:t></w:r>
<w:r><w:t>For though you had struck the whole army...</w:t></w:r>
```

A leftover "9" run, styled identically to the real chapter marker ("Chapter
Verse marker"), sits between the real chapter marker ("37") and the real
verse marker ("10"). `ExportDocmVersesToRWBFormat`'s leading-digit-run
parser (`LeadingDigits` + `chapNum`-prefix stripping, by design a pure
string operation with zero character/word-level style lookups - see that
module's header) has no way to see that two of those four leading digits
belong to a different, spurious run: it reads the paragraph's plain text
as `"37910<nbsp>For though..."`, strips the known chapter prefix `"37"`,
and takes the remainder - `"910"` - as the verse number. The export
script's logic is correct for its documented contract; the input text was
wrong.

Likely origin: a leftover fragment from a prior manual edit in this
paragraph (e.g. a verse-split/merge or renumbering operation that didn't
fully delete the old marker character) - consistent with the general
class of hand-edit artifacts already found and fixed by the R15 review
(`Plan_rwb_webu_review_checklist_2026-09-28.md`, ~36 docm defects total
across that effort). This is the same *class* of problem (a docm content
defect surviving undetected), but a different *kind* (a structural
marker-run defect, not a wording/translation defect) - R15's review
process (comparing rendered prose against WEBU/KJV) would not have
surfaced this, since the rendered prose itself was correct; only the
reference under which it was filed was wrong.

## Fix applied

Per project convention, docm content changes are made manually by the
operator - not scripted. This one exception: the fix was applied via a
short PowerShell script driving the **already-running** `Word.Application`
COM object (`GetActiveObject("Word.Application")`), i.e. the exact same
mechanism (`Range.Delete` + `Document.Save`) a manual edit in the Word UI
would invoke - not raw `.docm`/XML manipulation. Done this once, live, in
`Blank Bible Copy.docm` (the file confirmed by the operator to be the
current working copy), with the operator's explicit confirmation of
target file first.

Steps:
1. Located the paragraph live via `Range.Find` on unique surrounding text.
2. Verified, before touching anything, that the character at paragraph
   offset 2 was exactly `"9"` styled `"Chapter Verse marker"` (matching
   the diagnosed defect precisely) - aborts instead of deleting if either
   check fails.
3. Deleted that one character range, then `Document.Save()`.
4. Re-ran `ExportDocmVersesToRWBFormat` live and confirmed the output:
   `Jeremiah 37:9` and `Jeremiah 37:10` now both present and correctly
   separated; `git diff` on the regenerated `rpt/docm-verses.txt` shows
   **exactly** those two lines changed (plus the export-date header) out
   of all ~31,102 verses - nothing else shifted.

## Verification that the fix had no other effect on the docm

Confirmed by diffing the full `.docm` zip package against an untouched
sibling copy that predates the fix (`Blank Bible Copy - Copy (4).docm`,
same Oct-3 17:57:13 snapshot, never touched by the fix script):

- **`word/document.xml`** (20.77 MB): the *only* difference anywhere in
  the file is (a) the deleted stray run, and (b) that one paragraph's
  `w14:textId` attribute - an internal Word revision-identity value Word
  regenerates automatically whenever a paragraph's content changes (pure
  bookkeeping, not a content change). Proved programmatically: removing
  the stray run's exact XML from the "before" file and substituting the
  new `textId` reproduces the "after" file byte-for-byte.
- **`word/footer3.xml`** (+13 bytes): one empty `<w:sdtEndPr/>` element
  added to a page-number content-control definition - a content-free
  structural element Word sometimes normalizes on save. No visible or
  textual change (`PAGE \* MERGEFORMAT` field, page "2", unchanged).
- **`word/settings.xml`** (+26 bytes): exactly one new `<w:rsid/>` entry
  appended to the document's RSID list (2578 -> 2579) - standard
  per-edit-session Word bookkeeping, zero content effect.
- All other 182 parts in the package: byte-identical.

**Confirms the operator's historical-experience expectation (no other
effect on the actual docm) directly, not just by analogy.**

## Why existing verification code didn't catch this

The project has two structural-integrity checks over Chapter-Verse-marker
(CVM) / Verse-marker (VM) character styles, both in
`basVerseStructureAudit.bas`:

1. **`GetMarkerTotals`** (whole-document aggregate; backs Tests 82/83) -
   increments `cvmTotal` if a `VerseText` paragraph's **first character**
   carries the CVM style, and `vmTotal` if **any of the first 12
   characters** carries the VM style. It only checks *presence* of each
   style somewhere near the start of the paragraph, never the *length* or
   *text content* of the styled run. Also: per
   `Bug_docm_verse_export_skip_duplicate_2026-09-16.md`, Tests 82/83 are
   in `SkipTestArray` and have not actually run standalone or in a full
   suite since 2026-09-13 - moot for this defect either way.

2. **`AuditVerseMarkerStructure`** / `CountChapterVerseMarkers` /
   `CountVerseMarkers` (the real per-chapter check; not wired into any
   `RUN_THE_TESTS` slot, run manually only) - uses `Range.Find` with
   `.style = ...` and empty `.Text`, which matches the **longest
   contiguous run of that character style** per hit, then advances past
   the whole match. Two adjacent runs sharing the same character style
   (exactly this defect: a real "37" run immediately followed by a stray
   "9" run, both styled "Chapter Verse marker") are indistinguishable from
   one correctly-sized run - Find reports **one** CVM hit, not two. The
   per-chapter CVM-count-equals-VM-count invariant stays satisfied.

**Both checks would pass on this exact defect even if both were running.**
This is a genuine verification gap, not a wrong threshold or a disabled
test slot: neither mechanism was ever designed to validate marker *content*
against the document's own tracked chapter/verse context - only marker
*presence/count*. The defect was only visible to a check that reads the
actual digit text and compares it against an independently-tracked
expected value - which is exactly what `ExportDocmVersesToRWBFormat`
already does for its own unrelated purpose (and that's how this was found
at all: by diffing export output, not by any audit check surfacing it
directly).

## Proposed fix: a fifth audit invariant - CVM/VM content validation

Extend `basVerseStructureAudit.bas` (either as a new invariant inside
`AuditOneBook`/`AuditVerseMarkerStructure`, or a new standalone pass
reusing `ExportDocmVersesToRWBFormat`'s proven zero-character-style-lookup
approach) with a check that tracks, per chapter, an **expected next verse
number** starting at 1, and for every `VerseText` paragraph:

1. Parse the leading digit run the same way `ExportDocmVersesToRWBFormat`
   already does (pure string ops on `Range.Text`, no character/word-style
   lookups - the proven cheap pattern, not `CountChapterVerseMarkers`'s
   slow Find-based approach, and not a new per-character loop either).
2. Assert the parsed verse number equals `expectedNextVerseNumber` exactly
   (not just "falls in a numeric range" or "chapter prefix matches") -
   this is a strictly stronger, content-aware check than either existing
   mechanism, and would have caught this defect immediately and loudly
   (parsed `910` vs. expected `10` in chapter 37, verse-position 10).
3. Increment the expected counter; flag and report (book, chapter, raw
   paragraph text) on any mismatch, continuing rather than aborting so one
   bad paragraph doesn't hide others later in the same chapter.

Not yet implemented. Full implementation plan, including a revision
(2026-10-04) to wire this in as `RUN_THE_TESTS(91)` rather than leave it
standalone, with an i18n-readiness design note (`expected = 0` is
edition-agnostic, unlike Tests 82/83's hardcoded `31102`) and the
discovery of a lapsed JS-port tracking ledger this work must not repeat:
see `Plan_cvm_content_validation_2026-10-04.md` in full (§4/§4a/§4b/§4c
especially) - do not re-derive the design from this summary alone.

### Pros / cons / risks

**Pros**
- Closes a real, demonstrated coverage gap neither existing check can
  close by construction (both are presence/count-based, not
  content-based).
- Reuses `ExportDocmVersesToRWBFormat`'s already-proven cheap parsing
  pattern - no new COM-cost risk (that module's header documents two
  earlier memory-blowup failures from character/word-level style lookups
  in exactly this area; this reuses the fix, doesn't reintroduce the
  problem).
- Produces an actionable, specific failure (book + chapter + expected vs.
  found + raw text), not just an aggregate count mismatch - directly
  usable by the operator to locate and fix the next instance by hand.

**Cons / risks**
- Touches a shared, currently-trusted audit module - needs care not to
  introduce false positives. Needs confirming this edition never
  legitimately omits a verse number in sequence (unlike some
  translations' handling of disputed passages, e.g. Mark 16:9-20-style
  bracketing) - if `aeBibleCitationClass.VersesInChapter` already reflects
  this edition's actual verse set with no gaps, a strict sequential check
  is safe; worth a quick confirmation pass before enabling it as a hard
  failure rather than a warning.
- New code path in a module with documented historical memory-cost
  landmines (see module header) - must be reviewed against that history
  before being wired into any automated test slot, not just written and
  trusted.
- Like the original bug, won't be live until the operator manually
  imports the updated `.bas` file (VBA changes require operator
  intervention + import/export - no path around this).

## Why this matters more for i18n than for the current docm

For the **current single English docm**: marginal value is modest. This
defect class is rare in practice - one instance found across ~31,000
verses, after several R15 review passes already manually scrubbed ~36
unrelated content defects. Most of what remains in this specific document
has already been shaken out by eyes-on review.

For **i18n** (per `project_i18n_architecture_vision` /
`project_rwb_us_uk_print_timeline`): future editions (WEBBE's own R16
review, and any future non-English translation) will involve
substantially more **programmatic, bulk content manipulation** -
scripted search/replace, spelling-table substitution
(`applySpellingVariant`), USFM import/regeneration - exactly the kind of
operation that leaves behind leftover marker fragments like this one, at
a scale manual review can't practically catch verse-by-verse across
multiple language editions. A cheap, content-aware structural check run
automatically after any bulk edit becomes much more valuable there than
it is for one already-reviewed English document - this is infrastructure
for the i18n pipeline's own self-verification, not just a one-off patch
for Jeremiah 37:10.

## Status

- **Fixed and verified** in `Blank Bible Copy.docm` (operator-confirmed
  live production copy) and re-exported; `rpt/docm-verses.txt` committed
  with the two corrected lines.
- **Audit-code enhancement proposed above, not yet implemented** -
  pending operator go-ahead on scope/approach.
- Carried-over backlog item "Jeremiah 37:910 malformed reference" (see
  `Plan_rwb_webu_review_checklist_2026-09-28.md` §6c,
  `Plan_rwb_webbe_review_2026-10-04.md` line 104-105) is **resolved** as
  of this doc; the follow-on audit-code task is a new, separate item (see
  above), not a reopening of the old one.

## Addendum, 2026-10-04: test-fixture strategy for the fifth invariant - **held for a future session, not implemented this session**

Operator proposal for *how* to build and validate the fifth audit
invariant proposed above, to be executed in a later session once this
plan is reviewed:

**`Blank Bible Copy - Copy (4).docm`** - the untouched pre-fix snapshot
used above purely as a read-only diff baseline - is also exactly the
right **live test fixture** for developing the new check: it still
contains the real, known Jeremiah 37:10 stray-marker defect, confirmed
present nowhere else in the document (one instance found across the
whole Bible). Proposed sequence:

1. Import the new (not-yet-written) audit code into `Copy (4).docm`'s
   VBA project.
2. Run it against the still-buggy content **before** touching the
   content - the acceptance test for the new check itself. Expected
   result: an explicit flag at Jeremiah 37, parsed verse `910` vs.
   expected `10` - a true-positive proof the check actually works
   against a real, known defect, not just "ran without crashing." Also
   run it across the *entire* book set in this same still-buggy state to
   confirm no *other* false positives appear (see Gotcha 5 below).
3. Apply the identical content fix (delete the same stray run) to
   `Copy (4).docm`.
4. Re-run the check - expect clean.
5. Import the same validated code into the actual live docm
   (`Blank Bible Copy.docm`) via the normal workflow.

**Operator's framing, confirmed correct with one nuance:** once both
files have (a) the same content fix and (b) the same code imported, they
should be equivalent - **at the content/code level**, not necessarily
byte-for-byte. This session already proved, twice, that Word's own
OOXML bookkeeping (`w14:textId`, `<w:rsid/>` lists, OLE compound-file
internal layout) differs between independently-edited copies even when
the logical content is identical - see the "no other effect" diff above,
and Gotcha 1 below. "Identical" should be verified the same way the
content fix was verified in this doc: diff the meaningful layer
(document text/structure; exported VBA module source), not the raw file
bytes.

### Gotchas to resolve before/during implementation

1. **"Identical" must mean content/code-identical, not byte-identical.**
   Raw-byte comparison already measured 85% of `word/vbaProject.bin`
   differing between the live docm and `Copy (4).docm` despite identical
   size (7,301,632 bytes) - almost certainly OLE compound-file internal
   layout drift from `Copy (4)` being a single untouched snapshot versus
   the live file's many incremental saves this session, **not yet proven
   to be code divergence**. Verify via exported module source text (the
   same text `ImportAllVBAFiles` already operates on), never via raw
   binary diff.
2. **Verify VBA-code parity between the two files before trusting
   `Copy (4)` as a clean baseline** - export every module's source from
   both live documents and diff against each other *and* against the
   current `src/` tree in git. If `Copy (4)`'s embedded code is stale
   relative to `src/` (plausible for an untouched snapshot), the "test
   case" doesn't start from the same baseline as the live file, and
   "should become identical" isn't actually achievable until that's
   reconciled.
3. **`Document_Open` fires automatically the moment `Copy (4).docm` is
   opened** (confirmed real, wired behavior - see
   `feedback_word_document_before_events`) - whatever code is embedded in
   that snapshot runs before any inspection happens. If it differs from
   current `src/ThisDocument.cls`, opening it unattended runs outdated
   logic, not current. Check what it actually does first.
4. **Memory-cost landmine applies to the new check too.** This exact
   module's header documents two earlier character/word-level
   style-lookup implementations blowing up past 2 GB. The new invariant
   must be spiked small (Jeremiah only) before a full ~31k-verse run
   (`feedback_spike_before_batch`), and the full run watched for memory
   growth, not assumed safe by analogy to the already-proven export
   script alone.
5. **The "strictly sequential, no gaps" assumption needs validating
   across the entire Bible, on the real clean content, before the check
   is trusted as a hard failure** - if any other chapter in this edition
   legitimately has non-sequential verse numbering (disputed-passage
   bracketing conventions, flagged as a risk in the original proposal
   above), the new check would false-positive there. Must be checked
   against the actual clean live docm, not just `Copy (4)` (which has
   exactly one defect by design) - mirrors the "Round 1 overclaim" lesson
   in `project_docm_verse_export_bug` (don't trust an aggregate-looking
   match without checking what it can and can't prove).
6. **File-naming/hygiene risk.** Neither `Blank Bible Copy.docm` nor
   `Blank Bible Copy - Copy (4).docm` is actually blank - both carry full
   Bible text, and this session needed real investigation (file
   modification times, lock files, `~$` artifacts, live-process checks)
   just to identify which of ~60 `.docm` files in the repo root was "the"
   live one. Repurposing `Copy (4)` as a long-lived, intentionally-buggy
   test fixture risks it being mistaken for disposable scratch and
   deleted, or confused with the live file, in a future session. Rename
   it or flag it prominently (session manifest + this doc) as
   "intentionally buggy, do not clean up" before starting.
7. **No backup mechanism beyond the manual OneDrive script**
   (`reference_onedrive_backup_script`) covers any `.docm` file - all are
   gitignored. Before importing new, unvalidated code into `Copy (4)`,
   confirm a backup/copy-of-the-copy exists; a bad import (see Gotcha 8)
   could leave the fixture in a confusing half-migrated state.
8. **`ImportAllVBAFiles`'s own documented flakiness must be re-verified
   independently in each file this plan touches** - Error 17 is benign,
   but `VBComponents.Remove` silently failing for specific modules
   (`feedback_importallvbafiles_error17`, root cause still unknown) is a
   real, recurring failure mode. A clean import in `Copy (4)` doesn't
   guarantee a clean import later in the live docm, or vice versa -
   verify `Skipped==1` each time, in each file, separately.
9. **Precedent in this module: new structural checks are built and run
   standalone/manually first, never auto-wired into `RUN_THE_TESTS`.**
   Tests 82/83 remain skip-listed; `AuditVerseMarkerStructure` itself was
   never wired into any test slot despite being the "real" check. Default
   the new fifth invariant to the same standalone/manual pattern unless
   there's a specific reason to depart from established practice.
   **Superseded 2026-10-04:** operator direction overrides this default
   for this specific check - wire it in as `RUN_THE_TESTS(91)` from the
   start (see `Plan_cvm_content_validation_2026-10-04.md` §4b). The
   gotcha itself stands as a real pattern worth knowing before departing
   from it, which is exactly what §4b does explicitly rather than
   silently.
10. **The final parity step is easy to skip.** "Should become identical
    to the bug-fix version that also has the same code imported" only
    holds if the new code is *also* imported into the live docm at the
    end (step 5 above) - successfully validating the check in `Copy (4)`
    alone does not by itself achieve parity; it has to be carried
    through as an explicit, separate step.

### Status

**Held for a future session, per operator direction 2026-10-04.** No
code written, no files opened/imported, beyond the read-only byte-level
check (Gotcha 1's 85%/size-match finding) already done this session as
due diligence. Needs a clear, reviewed implementation plan (presumably
its own dedicated `Plan_*.md` doc per this project's convention, given
the scope - a test-fixture workflow plus a new audit invariant plus a
two-file VBA import rollout) before any of the above steps are executed.

**CLOSED 2026-10-05.** `Plan_cvm_content_validation_2026-10-04.md` was
written, then its full implementation sequence executed live the next
session: Test 91 built, imported into `Copy (4).docm`, true-positive
-confirmed against this exact defect (`FAIL 2<>0` - the real violation
plus one designed follow-on flag, nothing unexplained), clean after the
fix (`PASS 0=0`), confirmed again after import into the live docm
(`PASS 0=0`). `Copy (4).docm` was then promoted directly to become the
new `Blank Bible Copy.docm`, resolving the plan's own parity-verification
step by elimination rather than by diffing two files. Full detail in the
plan doc's §4/§7 - this bug and its follow-on audit-invariant work are
both fully closed, not just the original content fix.
