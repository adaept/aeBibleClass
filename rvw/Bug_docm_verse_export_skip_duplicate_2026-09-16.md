# Bug investigation - docm verse-export skip/duplicate discrepancy - 2026-09-16

## Symptom

`ExportDocmVersesToRWBFormat` reports `wrote 31053 verses ... skipped=46
duplicates=3 unknownBookHeadings=0` on every run this session. `31053 + 46
+ 3 = 31102`, matching `aeBibleClass.cls` Tests 82/83's expected baseline
(`Expected1BasedArray` positions 82/83, both `31102`).

**Initial (wrong) claim, corrected here:** that this arithmetic match proves
the document is structurally sound and the 46/3 are export-side parsing
weaknesses on otherwise-valid content. **This was overclaimed** - Tests
82/83 cannot actually prove per-chapter numeric correctness (see below), so
that conclusion wasn't earned. Operator caught this by asking whether
Tests 82/83 verify numeric B/C/V correctness (not just marker presence) and
whether a canonical per-chapter source table is checked anywhere.

## What Tests 82/83 actually verify (and don't)

`aeBibleClass.CountVerseMarker`/`CountChapterVerseMarker` (test slots 82/83)
call `basVerseStructureAudit.GetMarkerTotals`, which walks every `VerseText`
paragraph once and counts:
- `cvmTotal` - paragraphs whose **first character** carries the "Chapter
  Verse marker" character style.
- `vmTotal` - paragraphs where any of the **first 12 characters** carries
  the "Verse marker" character style.

**This is a whole-document aggregate total, not a per-chapter check.** Per
`GetMarkerTotals`'s own header comment, this is a deliberate speed
shortcut: the real per-chapter check (`CountVerseMarkers`/
`CountChapterVerseMarkers`, called per-chapter from `AuditOneBook`) is
"correct but slow (300-2700 s) because Word's Find degenerates on
character-style runs in a large document" - the aggregate version "a single
pass through 35k paragraphs completes in seconds."

**Consequence:** a duplicate verse in one chapter and a missing verse in
another chapter would leave both aggregate totals at exactly 31,102,
completely undetected. Tests 82/83 passing does **not** prove:
- every chapter has the canonically-correct verse count,
- there are no duplicate chapter:verse references,
- there are no gaps/omissions in the numbering.

## The check that WOULD catch this - unused this session

`basVerseStructureAudit.AuditVerseMarkerStructure` (public, standalone,
**not called from any `RUN_THE_TESTS` slot**) does the real check:

1. Walks every canonical book (via `aeBibleCitationClass.GetCanonicalBookTable`).
2. Walks every chapter (Heading 2) within each book.
3. **Per chapter**, compares the actual `VM`/`CVM` count against
   `aeBibleCitationClass.VersesInChapter(bookName, chIdx)` - a genuine
   canonical per-chapter verse-count source table, distinct from the
   book-level chapter-count table used elsewhere.
4. Also checks CVM count == VM count per chapter (the "one Chapter Verse
   marker + one Verse marker per verse" structural rule) - catches a verse
   missing its leading CVM run even if the VM-only count happens to match.
5. Writes `rpt/VerseStructureAudit.txt` plus an Immediate-window summary,
   listing every book/chapter with a mismatch by name.

**This has apparently never been run this session** (no `rpt/
VerseStructureAudit.txt` referenced, and it's absent from every
`RUN_THE_TESTS` log seen). Given its documented 300-2700 second runtime,
this is very plausibly why it isn't part of routine testing - the
operator's instinct ("tests that take a long time... in the skip category")
was directionally correct: not that a *scheduled* test gets silently
skipped, but that the *rigorous* version of this check is a separate,
slow, manually-invoked tool that automatic testing substitutes a fast
aggregate for.

## Pre-run code review (operator request, 2026-09-16) - before spending the 5-45 minute runtime

**✅ Real bug found and fixed, aeBibleClass `<pending commit>`:** `GetMaxVerse`
(the function `VersesInChapter`/`AuditOneBook` ultimately depends on) had an
off-by-one bounds check: `Chapter > UBound(maps(BookID)) + 1` let
`Chapter = UBound+1` pass validation, then crash indexing the array.
**Verified this is a genuine off-by-one, not a 0-based-array blind spot**:
`ToOneBasedLongArray` explicitly converts every book's literal `Array(...)`
(0-based by VBA default in this module - no `Option Base 1`) into a true
1-based array via `ReDim temp(1 To Count)`, and `AssertOneBased` checks this
on every call - confirmed for Genesis, `UBound(maps(1))` is genuinely `50`,
so `Chapter=51` should never have passed. Fixed: dropped the `+ 1`.

**Correction to an initial overclaim in this same review:** first assessed
this as a crash risk for `AuditVerseMarkerStructure` specifically. On
tracing the actual call path, `AuditOneBook` calls `VersesInChapter`, not
`GetMaxVerse` directly - and `VersesInChapter` has its own separate guard
(`Chapter > maxCh`, using the canonical `GetMaxChapter` table) that
intercepts an out-of-range chapter *before* reaching `GetMaxVerse`'s buggy
line, returning `0` cleanly instead of crashing. **This bug was real but
was not actually a crash risk for the run about to happen** - worst case
it would have shown as a normal "MISMATCH" line. Still correct to fix
(defense in depth; other callers of `GetMaxVerse` may not have the same
guard - `ValidateSBLReference` calls it more directly and wasn't fully
traced here, out of scope for this pass). Confirmed via
`basTEST_aeBibleCitationClass.bas`'s own `Test_GetMaxVerse`: it only tests
grossly-invalid inputs (chapter `999`), never the exact `UBound+1`
boundary - exactly why this escaped detection until now.

**✅ Minor, fixed alongside:** `VersesInChapter`'s error handler said
`"...of Class aeSBL_Citation_Class"` (wrong class name - this code lives in
`aeBibleCitationClass.cls`) and used `MsgBox` instead of `Debug.Print` (this
project's convention for error handlers). Fixed. **Not fixed, flagged for a
future separate pass:** the identical wrong-class-name pattern appears in
11 other functions throughout `aeBibleCitationClass.cls` - clearly a
leftover from an earlier rename, out of scope for this bug's fix.

**Confirmed, not a bug:** the operator confirmed all 5 single-chapter books
(Obadiah, Philemon, 2 John, 3 John, Jude) do have a genuine `Heading 2`
styled "CHAPTER 1" in the docm - the single-chapter-book false-positive
concern raised earlier in this review is ruled out.

## Selah / PsalmSuperscription / Psalms BOOK - confirmed not the cause

- `Selah` is a **character style** applied to a word *inside* an ordinary
  `VerseText` paragraph - doesn't remove the paragraph from either count.
- `PsalmSuperscription` (e.g. "A Psalm by David") and the Psalms "BOOK"
  division headers are **separate paragraph styles**, not `VerseText`.
  Both `GetMarkerTotals` and `ExportDocmVersesToRWBFormat`'s main loop gate
  on `StyleName = "VerseText"` identically, so these paragraphs are
  excluded from both counts equally - they don't explain the 46/3
  discrepancy and need no special handling in the export's skip logic.

## What we actually know vs. don't know

**Known:** the export's 46 skips + 3 duplicates come from its own naive
digit-parsing of `VerseText` paragraph plain text (chapter-number-prefix
string matching) - documented in the export routine itself, unrelated to
character-style presence.

**Not known (this is the open question):** whether those 46+3 paragraphs
are (a) structurally fine per the canonical per-chapter table and only
failing the export's naive parser, (b) genuine docm numbering defects
(duplicates, gaps, or misnumbered verses) that Tests 82/83's aggregate
happens not to expose, or (c) some mix of both.

## Path to resolution (not yet executed)

1. **Run `AuditVerseMarkerStructure`** (accept the 5-45 minute runtime) to
   get the authoritative per-chapter report against the canonical
   `VersesInChapter` table - this is the one piece of code in the project
   that can actually answer "is the docm's B/C/V numbering correct."
2. **Add targeted diagnostic logging to `ExportDocmVersesToRWBFormat`**
   (Immediate-window `Debug.Print`, per this project's error-handling
   convention) that prints the raw paragraph text for each of the 46
   skips and 3 duplicates, so the specific verses can be identified.
3. **Cross-reference the two reports** - do the audit's flagged
   chapters/books line up with the export's skipped/duplicate paragraphs?
   If yes, the docm has real numbering defects to fix. If the audit comes
   back clean (0 issues, 31,102/31,102 per-chapter) while the export still
   skips 46/3, the defect is confirmed to be in the export's parser only.
4. **Only after that**, decide whether any of Phase 4's prior work (Pass 1/
   2 sync, Pass 3 divine-name census) needs re-running against a corrected
   `docm-verses.txt` - unlikely to change Test 70/71 results (neither
   pattern's verses are known to be among the 46/3), but Pass 3's still-
   open 13% `Lord`-count gap is a candidate worth re-checking once the real
   picture is known.

## Process bug found alongside this investigation (operator, 2026-09-16): the release process has no way to catch this

**Confirmed real, not hypothetical:** `README.md` documents `RUN_THE_TESTS`
as *"Run all tests"* - the only testing workflow this repo's public docs
describe. `AuditVerseMarkerStructure` is not one of the 86 numbered
`RUN_THE_TESTS` slots and is never mentioned in `README.md` at all. Anyone
following the documented testing workflow - including a future release
process - would have no way to know this check exists, let alone that it
needs to be run separately. This is exactly how a canonical-versification
regression could ship undetected.

**Placement note (per `feedback_public_vs_internal_docs`, not yet
decided):** this repo's `README.md` is external-user-facing only -
maintainer/release-process runbooks belong in adaept5tudio's private docs,
not here. So "add a prominent note" needs a location decision before it's
written:
- Source-code comments (near `RUN_THE_TESTS`'s dispatch in
  `basTest_aeBibleClass.bas`, and in `AuditVerseMarkerStructure`'s own
  header) - appropriate regardless of public/private, maintainer-facing by
  nature either way.
- A release-process checklist, if a dedicated one is wanted - operator to
  decide whether that belongs in adaept5tudio's private docs (per the
  standing rule) or this repo's `rvw/` (internal working notes, arguably
  not "docs" in the public sense the standing rule was written to keep
  clean).

### Task - not yet done

- ⚪ Add a prominent, hard-to-miss note (debug/code-comment, documentation,
  and any other relevant surface) that `AuditVerseMarkerStructure` exists,
  is not part of `RUN_THE_TESTS`, and must be run explicitly before any
  release - specific location(s) pending the placement decision above.

## Status

🟡 Pre-run code review done, two real bugs found and fixed (off-by-one in
`GetMaxVerse`; wrong class name/`MsgBox` in `VersesInChapter`'s error
handler). A third, process-level bug found and logged (above) - the
release process has no documented way to know `AuditVerseMarkerStructure`
needs to run at all - not yet fixed, placement pending. `AuditVerseMarkerStructure`
about to be run for the first time. No diagnostic logging added to the
export yet; no docm changes made.
