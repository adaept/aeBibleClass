# Plan: CVM/VM content-validation audit invariant + test-fixture rollout - 2026-10-04

**This is a plan document only - nothing here is built or executed yet.**
Per the operator's explicit framing (2026-10-04): capture the plan now,
implement later. Follow-on to
`Bug_jeremiah_37_10_stray_cvm_marker_2026-10-04.md`, which root-caused
and fixed one real docm content defect (Jeremiah 37:10's stray marker
character) and found a genuine gap in this project's existing
marker-integrity checks. Read that doc in full before starting
implementation - this plan does not re-derive its findings, only acts on
them.

## 0. Problem statement (summary - see the bug doc for full detail)

`basVerseStructureAudit.bas` has two existing CVM (Chapter Verse marker)
/ VM (Verse marker) integrity checks, `GetMarkerTotals` (whole-document
aggregate, backs skip-listed Tests 82/83) and `AuditVerseMarkerStructure`
/ `CountChapterVerseMarkers`/`CountVerseMarkers` (per-chapter, never
wired into `RUN_THE_TESTS`, run manually only). **Both are
presence/count-based, not content-based** - neither can detect a stray
character sharing the correct style sitting adjacent to a real marker,
because `GetMarkerTotals` only checks style presence near the start of a
paragraph, and `Range.Find` with a style filter matches the *longest
contiguous run* of that style, silently absorbing an adjacent
same-styled stray character into one hit. This exact defect (Jeremiah
37:10, verse parsed as `910` instead of `10`) would have passed both
checks even running clean. Confirmed by code-reading both mechanisms, not
assumed.

## 1. Goal

Add a fifth audit invariant that validates CVM/VM run **content**, not
just presence/count: per chapter, track an expected-next-verse-number
counter starting at 1, and for every `VerseText` paragraph assert the
parsed verse number equals that counter exactly, incrementing after each
check. A stray digit (missing or extra) breaks the sequence immediately
and loudly - this is a strictly stronger check than anything currently
in place, and would have caught Jeremiah 37:10 on the spot.

## 2. Design constraints (non-negotiable, from prior incidents in this exact module)

- **Reuse `ExportDocmVersesToRWBFormat`'s proven parsing pattern**: one
  `Range.Text` read per `VerseText` paragraph, pure string ops
  (`LeadingDigits`, chapter-prefix stripping) - zero character- or
  word-level style lookups in the hot path. This module's own header
  documents two earlier implementations blowing past 2 GB of memory from
  exactly that kind of per-character/word COM call in a ~35k-paragraph
  walk. Do not reintroduce that pattern, even by accident, inside a new
  invariant.
- **Spike before batch** (confirmed-good project pattern,
  `feedback_spike_before_batch`): prove the new check on the smallest
  real known case (Jeremiah 37 alone) before running it across the full
  ~31k-verse document, and watch memory during the full run rather than
  assuming safety by analogy.
- **Wired into `RUN_THE_TESTS` as a new numbered test (revised
  2026-10-04 - see §4a/§4b below), not left standalone.** The original
  draft of this plan deferred this, following this module's own
  precedent (`AuditVerseMarkerStructure` itself was never wired in;
  Tests 82/83 remain skip-listed over a year later). Operator direction
  overrides that precedent for this check specifically: build it as a
  real, numbered `RUN_THE_TESTS` slot from the start, with the
  i18n-readiness design constraint in §4a, rather than leaving it a
  manually-run diagnostic that has to be remembered and re-run by hand
  (exactly the fate that befell `AuditVerseMarkerStructure` itself).
- **VBA changes are never live until manually imported** via the VBE,
  regardless of who writes the `.bas` text
  (`feedback_docm_manual_vba_import_convention`,
  `feedback_importallvbafiles_error17`). Every step below that touches
  code ends with "operator imports," not "done."

## 3. Pre-flight checklist - resolve before writing any code

These come directly from the 10 gotchas in the bug doc's addendum,
reordered here as blocking pre-flight checks rather than a flat list:

1. **Verify `Blank Bible Copy - Copy (4).docm`'s embedded VBA code is not
   stale relative to the current `src/` tree.** Export every module's
   source from both the live docm and `Copy (4)` (live Word COM, the same
   technique already used to fix the content defect) and diff against
   each other and against `src/` in git. A raw-byte check this session
   found `word/vbaProject.bin` 85% different by byte count between the
   two files despite identical size (7,301,632 bytes) - likely just OLE
   compound-file layout drift from differing save histories, but **not
   proven**. Do not proceed to step 4 below until this is resolved one
   way or the other.
2. **Confirm a backup of `Copy (4).docm` exists** before importing
   unvalidated code into it (no `.docm` is git-tracked; the only backup
   mechanism is the manual OneDrive script,
   `reference_onedrive_backup_script`).
3. **Check what `Copy (4).docm`'s embedded `Document_Open` actually does**
   before opening it - confirmed real/wired behavior
   (`feedback_word_document_before_events`), not a no-op. If it differs
   from current `src/ThisDocument.cls`, opening it unattended runs
   outdated logic.
4. **Rename or prominently flag `Copy (4).docm`** as "intentionally
   buggy, do not clean up" before starting - neither it nor
   `Blank Bible Copy.docm` is actually blank, and this session needed a
   real investigation (mtimes, lock files, live-process checks) just to
   identify which of ~60 `.docm` files in the repo root was "the" live
   one. Don't recreate that confusion for a fixture meant to persist
   across sessions.

## 4. Implementation sequence

1. Write the new invariant (design per §1-2) in `basVerseStructureAudit.bas`
   - present the diff to the operator one piece at a time per
     `feedback_code_review_process`, not as one large unreviewed change.
2. **Wire it in as `RUN_THE_TESTS(91)`** - the next free slot (`MaxTests`
   is `90` as of this session: Test 88 `AuditColorConstants`, Test 89
   `CountBuiltInHyperlinkStyleRuns`, Test 90
   `CountSpaceBeforePunctuation`). Follow the standard 8-step checklist in
   `md/Adding_To_Bible_Test_Class.md` exactly - this is a normal new test,
   not a special case:
   1. `MaxTests` 90 -> 91.
   2. `Expected1BasedArray` - append expected value **`0`** (see §4a - this
      is deliberately edition-agnostic, not a hardcoded English-specific
      count).
   3. `Expected1BasedArray`'s comment line - append `91`.
   4. `GetPassFail` - `Case 91: ResultArray(TestNum) = CountSequentialVerseNumberViolations()`
      (name TBD at write time; thin wrapper delegating to the new
      `basVerseStructureAudit` routine, matching the Test 88 pattern -
      `feedback_class_encapsulation`).
   5. `RunBibleClassTests` - `RunTest (91)` after `RunTest (90)`.
   6. `RunTest` - `Case 91`, matching label/output convention.
   7. `OutputTestReport` - `Case 91`, same label.
   8. `GetTestDescription` - `Case 91`, describing the rule in the same
      style as Cases 82/83/87 above it.
3. Operator imports into `Copy (4).docm` (not the live docm) via the
   normal `ImportAllVBAFiles` workflow - verify `Skipped==1`
   independently in this file (`feedback_importallvbafiles_error17`'s
   flakiness is not guaranteed to reproduce or not-reproduce the same way
   twice).
4. Run `RUN_THE_TESTS(91)` against `Copy (4).docm`'s still-buggy content.
   Acceptance criteria:
   - **True positive**: `FAIL`, with the result/hint identifying Jeremiah
     37, parsed verse `910` vs. expected `10` (not just a bare nonzero
     count - per §4a, the whole point of this invariant over Tests 82/83
     is an actionable, specific failure).
   - **No false positives elsewhere** - if any appear, STOP and determine
     whether they're (a) a real additional defect (good - same category
     as this whole investigation), or (b) evidence the "strictly
     sequential, no gaps" assumption is wrong for some legitimate reason
     (e.g. a disputed-passage bracketing convention) - do not assume (a)
     without checking, mirroring the "Round 1 overclaim" lesson in
     `project_docm_verse_export_bug` (an aggregate-looking result was
     trusted before checking what it could and couldn't actually prove).
   - Watch memory usage through the full run; abort and re-scope if it
     trends toward the documented failure pattern that sidelined Tests
     82/83.
5. Once clean except for the one known true positive: apply the same
   content fix to `Copy (4).docm` (delete the stray run, same as the live
   docm fix) and re-run `RUN_THE_TESTS(91)` - expect `PASS 0=0`.
6. Operator imports the same validated code into the live docm
   (`Blank Bible Copy.docm`) via the normal workflow - verify `Skipped==1`
   independently again (does not inherit step 3's result). Run
   `RUN_THE_TESTS(91)` there too - expect `PASS 0=0` on the already-fixed
   content.
7. Verify content/code parity between the two files **at the right
   layer** - document.xml text/structure diff (as already done for the
   original content fix) plus exported-module-source diff (per §3 item
   1's method) - not raw file/byte comparison. Document the result
   either way.
8. **Update the JS-port tracking ledger** - see §4c below. Not optional;
   this is the step that's been silently skipped for the last two VBA
   test additions (Tests 89 and 90).

### 4a. Why `expected = 0` makes this invariant i18n-ready by construction

Tests 82/83's `Expected1BasedArray` entries hardcode `31102` - a count
baked to *this specific English edition's* total verse count. Any future
non-English edition (a different docm, a different total verse count if
versification ever differs) would need its own hardcoded expected value
before Tests 82/83-style checks could run against it unmodified - a real
per-edition maintenance cost.

The new invariant's expected value is **`0`** - "zero sequential-number
violations found," the same structural contract (`feedback_class_encapsulation`-
style separation of a stateless rule from edition-specific data)
regardless of which language or edition's docm is `ActiveDocument` when
`RUN_THE_TESTS(91)` runs. The per-chapter expected-verse-*count* data it
reads (`aeBibleCitationClass.VersesInChapter`) is already the project's
existing single canonical source, used unchanged by every other check in
this module - this plan does not introduce a second one. **This is the
concrete, minimal sense in which the new test is "i18n-ready":** no
per-edition constant to update, no code branch keyed to a specific
language, and it runs correctly against any future docm that follows the
same CVM/VM/VerseText structural convention, the same way Tests 84-90
already do. It does **not** by itself solve cross-language versification
differences (if a future edition's canon genuinely numbers verses
differently, `VersesInChapter`'s own table would need an edition-specific
variant first - a pre-existing, separate concern this plan does not
create or solve).

### 4b. Precedent departure, noted explicitly

§2 originally deferred `RUN_THE_TESTS` wiring, citing
`AuditVerseMarkerStructure`'s own history (built, never wired, run
manually only when someone remembers to). Operator direction (2026-10-04)
is to wire this one in regardless - the i18n-readiness framing in §4a is
*why* this check specifically is worth the departure: it is cheap
(reuses the proven O(n) pattern, §2), edition-agnostic (§4a), and exactly
the kind of check that's useless if it only runs when someone remembers
to run it by hand - which is demonstrably what happened to
`AuditVerseMarkerStructure` itself (unwired since its creation,
per `project_docm_verse_export_bug`) and is the direct reason this whole
defect went undetected as long as it did.

### 4c. JS-port parity ledger - update required, gap found while writing this plan

`adaept5tudio/docs/aeBibleClass-word-addin-conversion-plan.md` §15.11
("Phase E deferred/skipped-Case tracker") is this project's existing
mechanism for tracking which VBA `RUN_THE_TESTS` Cases have/haven't been
ported to the JS add-in - its own text says "update it whenever a new
Case gets deferred/skipped, or when one of these gets resolved." It has
**not** been updated for the last two VBA test additions: Test 89
(`CountBuiltInHyperlinkStyleRuns`, added `aeBibleClass` commit `680779d`,
2026-09-21) and Test 90 (`CountSpaceBeforePunctuation`, commit
`1a9922e`, 2026-10-01) appear nowhere in that document, despite the
ledger itself being edited as late as 2026-09-22 (after Test 89 already
existed). **Confirmed by direct grep, not assumed** - see
`project_jeremiah_37_10_stray_cvm_marker` memory and
§6/the new "JS-port parity gap" analysis this session added to
`adaept5tudio/docs/aeBibleClass-word-addin-conversion-plan.md` for the
full finding and a proposed fix (a repeatable parity check, not just a
reminder to update the table by hand again). Test 91 (this plan) must
not repeat that lapse - step 8 above is the forcing function; see that
doc's new section for the actual mechanism proposed to stop relying on
memory for this going forward.

## 5. Pros / cons / risks (carried from the bug doc, confirmed)

**Pros:** closes a demonstrated coverage gap neither existing check can
close by construction; reuses an already-proven cheap parsing pattern
(no new COM-cost risk); produces an actionable, specific failure (book +
chapter + expected vs. found + raw text) instead of an aggregate count
match.

**Cons/risks:** touches a shared, currently-trusted audit module;
depends on the "no legitimate verse-number gaps" assumption holding
across all 66 books, not yet verified; two separate VBA imports (test
fixture, then live docm), each independently exposed to
`ImportAllVBAFiles`'s known-but-not-root-caused flakiness; a
not-yet-resolved question (§3 item 1) about whether the test fixture's
starting VBA state is even a valid baseline; wiring into `RUN_THE_TESTS`
(§4/§4b, a departure from this module's own "standalone, manual" history)
means a bug in the new check's logic now runs automatically on every
full-suite pass instead of only when manually invoked - raises the bar
on getting the acceptance criteria in step 4 right before this ships.

## 6. Explicitly out of scope for this plan

- Root-causing `GetMarkerTotals`'s original memory issue that got Tests
  82/83 skip-listed (`project_docm_verse_export_bug`, follow-up item 1 -
  separate, long-standing, not blocking this work). Test 91 is a new,
  independent slot with its own proven-cheap parsing pattern (§2) - it
  does not depend on, and should not be blocked by, that unresolved
  issue.
- Any further content-defect sweep beyond what this specific check
  surfaces (that is R15/R16 review territory, separate backlog).
- Building the general-purpose VBA-Case/JS-Case parity-check script
  proposed in §4c/the `adaept5tudio` conversion-plan update - naming the
  gap and proposing the mechanism is in scope here; building that script
  is its own follow-on task, not bundled into this plan.
- Reconciling cross-language versification differences in
  `aeBibleCitationClass.VersesInChapter` (§4a) - out of scope until a
  concrete future edition actually needs it.

## 7. Status

Plan only, written 2026-10-04, revised same day (§2/§4/§6: the new
invariant is now planned as `RUN_THE_TESTS(91)`, not a standalone-only
routine, with the JS-port ledger update folded in as step 8/§4c rather
than left implicit). Awaiting operator review/go-ahead before any
implementation step begins.
