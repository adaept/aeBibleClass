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

**Correction (2026-09-16, after actually running both `AuditVerseMarkerStructure`
and `RUN_THE_TESTS(82)`):** the framing above was wrong in an important way.
It's not "a fast aggregate substitute runs instead of the rigorous check" -
**nothing runs at all.** `RUN_THE_TESTS(82)` returns `SKIP` with `Result =
-1` (the untouched `InitializeGlobalResultArrayToMinusOne` initialization
value) - `GetMarkerTotals` never executes, standalone or in a full suite.
Confirmed via `MakeSkipTestArray`: `SkipTestArray = Array(82, 83)`, checked
uniformly by `IsSkipTest` regardless of run mode. Full history in
`rvw/Code_review 2026-09-13.md` item 7: a 2026-09-13 full-suite run hung at
Test 82 with memory climbing past 2GB (a COM-object accumulation problem
specific to running it after 8 other heavy tests in one call stack, not a
correctness bug in the counting logic - standalone it worked, `162.81s`,
genuine `PASS 31102=31102`, at a time before whatever edit created the 3
John defect below). The fix added 82/83 to `SkipTestArray`, which - as that
doc explicitly notes - **also disabled them standalone**, not just in full
runs. The underlying memory cause was never root-caused; explicitly
deprioritized once the suite was unblocked.

**The operator's original instinct was exactly right, more literally than
first credited:** Tests 82/83 are not just "slow so a fast substitute runs
instead" - they are **in a literal skip list** and have provided **zero
verification signal since 2026-09-13**. `AuditVerseMarkerStructure`, run
today, is the first real verse-structure check since that date, and it
found a genuine defect on its first run (see below).

## Pre-run code review (operator request, 2026-09-16) - before spending the 5-45 minute runtime

**✅ Real bug found and fixed, aeBibleClass `48307be`:** `GetMaxVerse`
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

## ✅ 3 John fixed and verified (2026-09-16)

The operator merged the two split paragraphs back together in the docm.
Two intermediate attempts surfaced real sub-issues along the way (a
partial deletion left `CVM Count 14 <> VM Count 15` - an orphaned `Verse
marker`-styled run with no paired `Chapter Verse marker`; found via
`ReportDigitAtCursor_Diagnostics`, `src/basTEST_aeBibleTools.bas` line 845,
which reports a character's style plus the one immediately before it -
exactly suited to hunting a style-transition boundary). Final
`AuditVerseMarkerStructure` re-run: **`31102 / 31102`, 0 structural
issues** - the docm's B/C/V numbering is now canonically correct across
the entire Bible, confirmed by the one tool in this project actually
capable of proving that.

## `AuditVerseMarkerStructure` result (2026-09-16) - first real run, first real finding

Ran successfully in 205.42s (well inside the documented 300-2700s range,
no memory issue - this function is not `GetMarkerTotals`, a different
implementation). Result: **`31103 / 31102` verses found, 1 structural
issue**:

```
3 John 1: expected verses=14  found=15
```

**Root cause, verified directly against WEBU:** WEBU has exactly 14 verses
in 3 John; its verse 14 is one continuous sentence: *"...but I hope to see
you soon. Then we will speak face to face. Peace be to you. The friends
greet you. Greet the friends by name."* The docm has this **split into two
separate verses/paragraphs** - `3 John 1:14` ("...face to face.") and `3
John 1:15` ("Peace be to you...by name.") - each independently numbered
and each carrying its own genuine "Verse marker" styling (confirmed by the
audit's own count, not just visible digit text). This is a real content/
formatting defect, not an intentional versification choice - not present
in WEBU, and `aeBibleCitationClass`'s canonical table correctly expects 14.

**Not among the export's 46 skips/3 duplicates:** both `3 John 1:14` and
`1:15` parse as clean, distinct, valid references in `docm-verses.txt` - no
skip or duplicate flagged there. So this defect is real but was invisible
to the export's own anomaly detection too - it only surfaced via the
canonical per-chapter cross-reference.

**What this resolves:** the audit found **exactly one** issue across the
entire Bible - every other book/chapter's marker count matches canonical
exactly. This makes it far more likely the export's 46 skips/3 duplicates
are genuinely export-parser-only weaknesses (per the export's own
documented naive digit-parsing), not a large population of hidden
numbering defects - 3 John is confirmed real, but appears to be an
isolated case, not the tip of a larger iceberg. Diagnostic logging on the
export (originally planned as step 2 below) would still pin down the exact
46+3 causes precisely, but the canonical-correctness question this whole
investigation started from is now largely answered: **the docm's B/C/V
numbering is correct except for this one confirmed defect.**

## New finding: post-3-John-fix export undercounts by exactly 1 paragraph

Re-ran the export after the 3 John fix and after `AuditVerseMarkerStructure`
confirmed `31102/31102, 0 issues`. Result: `wrote 31052 verses ...
skipped=46 duplicates=3` - **31052+46+3 = 31101, one short of the
now-confirmed-correct canonical total 31102** (previously masked: before the
fix it was `31053+46+3=31102`, which looked consistent only because the
then-actual defective total was `31103`, one high from the 3 John split -
two independent one-off errors that happened to cancel in the arithmetic
check, not evidence the export was sound).

Re-read `ExportDocmVersesToRWBFormat`'s full body
(`src/basRWBTextExport.bas`) to confirm the counters' relationship by
construction: every `VerseText`-styled paragraph increments `visitedCount`
by exactly 1, and then exactly one of `lineCount`, `skipCount`, or
`dupCount` (mutually exclusive branches, no other path). So
`visitedCount == lineCount + skipCount + dupCount` always, by the code's own
structure - the gap is not a tallying bug between the four counters. It
means the export's paragraph walk itself is only visiting **31101**
`VerseText`-styled paragraphs, one fewer than the 31102 marker-carrying
verses the canonical audit confirmed exist.

**Working hypothesis, not yet confirmed:** `AuditOneBook`'s per-chapter
`CountVerseMarkers`/`CountChapterVerseMarkers` (the slow, `Find`-based
functions backing `AuditVerseMarkerStructure`) search a chapter-bounded
`Range` for the "Verse marker"/"Chapter Verse marker" **character** styles
directly - with no paragraph-style filter at all. The export, by contrast,
only ever visits paragraphs whose **paragraph** style is literally
`"VerseText"`. If exactly one paragraph somewhere in the docm carries
genuine Verse-marker/Chapter-Verse-marker character styling but has the
**wrong paragraph style** (not `"VerseText"` - e.g. accidentally left as
`"BodyText"` or similar after an edit), the audit would count it correctly
while the export would skip it entirely - silently, not even as one of the
46 skips, since the paragraph never enters the `VerseText` branch at all.
This would exactly explain a clean 1-paragraph gap that produces no export
warning.

**Candidate tool to test this without touching the export's hot path:**
`basVerseStructureAudit.AuditCharStyleUsage(StyleName, bWriteFile,
bAnomaliesOnly)` already exists and, per its own documented `bAnomaliesOnly`
behavior, suppresses "runs where paraStyle=VerseText AND position=START" -
i.e. it's designed to surface exactly a marker-styled run in an unexpected
paragraph-style/position context. Not yet run for this purpose - the plan
is to call `AuditCharStyleUsage("Verse marker", True, True)` (and the same
for `"Chapter Verse marker"`) and inspect the anomaly list for the one
paragraph that would explain this gap, before considering any change to the
export itself. This deliberately avoids adding character/word-level style
lookups to `ExportDocmVersesToRWBFormat`'s hot path, which the module's own
header explicitly warns against (two earlier versions of this kind of
lookup caused multi-GB memory blowups, the same class of issue as
`GetMarkerTotals`/Tests 82/83).

**✅ Diagnostic logging added to the export itself** (separate from the
above, still useful for the 46/3 question): `ExportDocmVersesToRWBFormat`
now `Debug.Print`s the raw paragraph text (truncated to 60 chars) and
context for every skip and every duplicate, at the exact point each is
detected, per this project's error-handling convention (Immediate window,
not `MsgBox`). Not yet re-run/reviewed against a live export - next export
run will show the 46+3 explicitly instead of only their counts.

**✅ Hypothesis confirmed - real defect found: Psalm 4:2 mis-styled as
"Psalms BOOK".** First tried `AuditCharStyleUsage("Verse marker", True,
True)` to test this - aborted partway through (Ctrl+Break) after
discovering its anomaly filter isn't discriminating for this style: "Verse
marker" character-styling legitimately sits *after* the "Chapter Verse
marker" prefix within a verse paragraph (matches how `GetMarkerTotals`
counts them - CVM checked at char 1, VM checked anywhere in the first 12
chars), so its position is normally `MID`, not `START` - the position-based
anomaly test flagged effectively 100% of correct verses (19000+ "anomalies"
observed with zero narrowing), making it useless for isolating one bad
paragraph, and the unbounded whole-document `Find` loop was also visibly
degrading in speed per the same documented Word/`Find`-on-character-styles
issue (per-1000-run time climbing: 485s at 19000, 622s at 20000, 733s at
21000, 852s at 22000).

Wrote a new, targeted, much cheaper routine instead:
`basVerseStructureAudit.FindMarkerStyleOutsideVerseText(ByRef hitCount As
Long, Optional bWriteFile As Boolean = True)` - complementary to
`GetMarkerTotals`: paragraph-level iteration (proven fast) restricted to
the small minority of paragraphs NOT styled `VerseText` (2,725 of ~35k),
checking each with the exact same first-char/first-12-chars technique
`GetMarkerTotals` already uses on the `VerseText` set. Ran in 32.82s.
**Result: exactly 1 hit** - `Psalms BOOK`-styled paragraph, first-char-style
`Chapter Verse marker`, excerpt "42 You sons of men, how long shall my
glory be turned into dishonor?..." - this is **Psalm 4:2**, carrying
genuine marker character-styling but parented under the wrong paragraph
style (should be `VerseText`). This is the confirmed root cause of the
31101-vs-31102 export gap: `AuditVerseMarkerStructure`'s character-style-
only search counts it (hence the confirmed-correct canonical total of
31102), while `ExportDocmVersesToRWBFormat` (and `GetMarkerTotals`/Tests
82/83, were they not hard-skipped) only ever visit `VerseText`-styled
paragraphs, so this verse is invisible to them - not even logged among the
46 skips, since it never enters the `VerseText` branch at all.

**Fix required:** operator to correct Psalm 4:2's paragraph style to
`VerseText` in the docm directly (same category of fix as the 3 John
defect - a document data correction, not a code change).

**✅ Fixed and verified, 2026-09-16.** Re-ran
`FindMarkerStyleOutsideVerseText`: non-`VerseText` paragraphs scanned
dropped `2725 -> 2724` (Psalm 4:2 now correctly counted as `VerseText`),
**hits = 0**. Test 87 would now pass.

**✅ Added as a new permanent test, `Test 87`
(`CountMarkerStyleOutsideVerseText`, expected baseline `0`)** - `MaxTests`
bumped 86 -> 87, wired into all four parallel dispatch switches
(`GetTestDescription`, `GetPassFail`'s `ResultArray` case, both Immediate/
buffer loggers) plus `Expected1BasedArray` (position 87 = 0) and the
`RunTest(87)` call added after `RunTest(86)` in `RunBibleClassTests`, with
a `DoEvents` in between matching the existing 74-86 pattern. Backed by a
new wrapper, `CountMarkerStyleOutsideVerseText()`, which calls
`FindMarkerStyleOutsideVerseText` with `bWriteFile:=True` (writes
`rpt\MarkerStyleOutsideVerseText.txt` on every run, matching
`CountHeaderFooterStyleViolations`'s convention of always leaving an audit
file behind). Not added to `SkipTestArray` - unlike `GetMarkerTotals`
(Tests 82/83), this routine only touches ~2,725 non-`VerseText` paragraphs
(not all ~35k), so its risk of the same full-suite COM-accumulation memory
blowup is expected to be much lower, but this is **not yet proven** - a
full-suite `RUN_THE_TESTS` run (not yet performed since adding Test 87)
should be watched for memory growth the same way the 2026-09-13 incident
was diagnosed, before treating Test 87 as fully safe long-term. Test 87
does **not** replace `AuditVerseMarkerStructure`: it only catches "marker
styling in the wrong paragraph style," not duplicate/missing verse numbers
within a correctly-styled `VerseText` paragraph - the release-process
reminder (`RunBibleClassTests`, after `RunTest(87)`) was updated to say so
explicitly.

## 46 skips / 3 duplicates - root-caused via the new diagnostic logging, mostly fixed

With the per-skip/duplicate `Debug.Print` logging live, analyzed the raw
text of each entry (cross-referenced against WEB wording to identify the
real verse - not yet independently verified against a live source for
every case, treat identifications below as high-confidence, not certain).
Confirmed by the operator to be two distinct root causes:

1. **Individual C:V marker digit typos** (majority of the 46) - a digit
   missing from the front or back of the chapter or verse number in the
   marker text itself (e.g. Psalm 94:11's marker read `4:11`, missing the
   leading `9`; Revelation 19:13's read `1:13`, missing the trailing `9`;
   a few were substitutions/insertions rather than deletions - e.g. Psalm
   25:17 read `26:17`, a `5`->`6` typo, and Psalm 119:92 had an extra
   duplicated leading `1`). Genuine content typos, fixed by the operator
   directly in the docm, one by one.
2. **1 Peter 4's Heading 2 read "MYCHAPTER 4" instead of "CHAPTER 4"** -
   explains the 15-entry cluster logged as `book="1 Peter" chapNum=2` with
   digit runs `41`..`419` (exactly matching all 19 verses of the real
   1 Peter 4): the chapter heading wasn't recognized/parsed correctly, so
   `chapNum` never advanced for that chapter's verses. Fixed by the
   operator correcting the heading text.

**Re-run after these fixes:** `wrote 31095 ... skipped=4 duplicates=3`
(`31095+4+3=31102`, arithmetic still closes against the canonical total).
**✅ All 7 fixed and verified, 2026-09-16.** Operator confirmed all 4 skip
hypotheses correct (Psalm 12:3, Psalm 25:17 - was showing `6`, Psalm
119:92 - extra leading `1`, Revelation 12:2). Re-export:
`skipped=0 duplicates=3` (`31099+0+3=31102`). Of the 3 "duplicates," the
operator confirmed the hypothesis exactly: **2 were missing-number typos**
(Matthew 10:35 and Acts 11:16, matching the predicted misidentification -
not true duplicates, no content was at risk), and **1 was a genuine
duplicate paragraph** (John 16:6, the lowest-confidence prediction, now
confirmed correct too). Final re-export:

```
ExportDocmVersesToRWBFormat: wrote 31102 verses to ...\rpt\docm-verses.txt
  skipped=0 duplicates=0 unknownBookHeadings=0
```

**31102 written, 0 skipped, 0 duplicates - exact match to the canonical
total with zero anomalies of any kind.** This closes the investigation
that started from the original `31053/46/3` arithmetic-mismatch report.

## What we actually know vs. don't know

**Known:** the export's 46 skips + 3 duplicates come from its own naive
digit-parsing of `VerseText` paragraph plain text (chapter-number-prefix
string matching) - documented in the export routine itself, unrelated to
character-style presence. **Known (new):** the docm has exactly one
confirmed canonical-numbering defect (3 John, above), and it's not among
the 46/3.

**Not known:** the exact identity of the 46+3 export-skipped/duplicate
paragraphs (still needs diagnostic logging, step 2 below, to confirm they
really are all parser-only and not a second, different-shaped defect the
per-chapter audit's own methodology happens not to catch).

## Path to resolution

1. ~~Run `AuditVerseMarkerStructure`~~ **✅ Done.**
2. ~~Decide how to fix 3 John~~ **✅ Done and verified - `31102/31102`, 0 issues.**
3. ~~Add targeted diagnostic logging to `ExportDocmVersesToRWBFormat`~~
   **✅ Done** - prints the raw paragraph text for each skip/duplicate at
   detection time. Not yet re-run against a live export. **Superseded in
   priority by the new finding above**: the export now undercounts total
   `VerseText` paragraphs visited by exactly 1 relative to the canonical
   audit (31101 vs 31102), a different question than the 46/3 shape - the
   diagnostic logging answers the 46/3 question but not the 1-paragraph
   gap, which needs `AuditCharStyleUsage` (see above) instead.
4. **Separately, decide whether to root-cause the `GetMarkerTotals` memory
   issue** (`rvw/Code_review 2026-09-13.md` item 7, deprioritized at the
   time) so Tests 82/83 could be safely restored to `SkipTestArray`-free
   operation - currently they provide zero signal at all, standalone or in
   a full suite, and have since 2026-09-13. Independent of `AuditVerseMarkerStructure`
   continuing to exist as the authoritative slow check either way.
5. **Only after 3**, decide whether any of Phase 4's prior work (Pass 1/
   2 sync, Pass 3 divine-name census) needs re-running against a corrected
   `docm-verses.txt` (the 3 John fix changed that book's text - Pass 1/2's
   scope was Tests 70/71's patterns only, neither known to touch 3 John,
   but not re-verified) - Pass 3's still-open 13% `Lord`-count gap remains
   a candidate worth re-checking regardless.

## Process bug found alongside this investigation (operator, 2026-09-16): the release process has no way to catch this

**Confirmed real, not hypothetical - and more severe than first stated:**
`README.md` documents `RUN_THE_TESTS` as *"Run all tests"* - the only
testing workflow this repo's public docs describe. `AuditVerseMarkerStructure`
is not one of the 86 numbered `RUN_THE_TESTS` slots and is never mentioned
in `README.md` at all. **Worse than "a different, better check exists
outside RUN_THE_TESTS":** the two tests that were supposed to cover this
inside `RUN_THE_TESTS` (82/83) are themselves hard-skipped
(`SkipTestArray`, see above) and have provided zero signal since
2026-09-13. So there was no automated verse-structure verification of any
kind - fast or slow - until `AuditVerseMarkerStructure` was run by hand
today, and it found a real defect (3 John) on its first run. Anyone
following the documented testing workflow - including a future release
process - would have had no way to know any of this.

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

### Task - done 2026-09-16

**Placement decided (operator):** the actual release-checklist entry goes
in adaept5tudio's private docs, per the standing rule - plus an in-repo
pointer and a runtime reminder, so the note is visible from every angle
someone might encounter it.

- ✅ **Release-process guard added**, `adaept5tudio` `4446c74` -
  `adaept5tudio/docs/aeBibleClass-word-addin-conversion-plan.md` section
  10.5, new step "6a" (non-disruptive - avoids renumbering steps 7-8, which
  have historical progress notes tied to their numbers) - full explanation
  of why `AuditVerseMarkerStructure` isn't equivalent to `RUN_THE_TESTS`
  82/83, with a pointer back to this doc.
- ✅ **In-repo pointer**: this bug doc itself, cross-referenced from the
  release-process doc above (mutual pointer, findable from either side).
- ✅ **Runtime reminder, aeBibleClass `869f46c`**: `aeBibleClass.cls`
  `RunBibleClassTests` now prints `"REMINDER: AuditVerseMarkerStructure is
  NOT included above - run it separately before any release..."` once,
  after `RunTest(86)`, only on a full-suite run (not per individual test -
  found and corrected a placement mistake first: initially added inside
  `RunTest` itself, which runs once *per test* and would have spammed the
  reminder 86 times per full run).

## Status

✅ **Closed, 2026-09-16.** `ExportDocmVersesToRWBFormat` now reports
`wrote 31102 verses ... skipped=0 duplicates=0 unknownBookHeadings=0` -
exact match to the canonical total with zero anomalies. Full resolution
chain: `AuditVerseMarkerStructure` found and the operator fixed a real 3
John split-verse defect (confirmed `31102/31102, 0 structural issues`);
the export's own 1-paragraph undercount was root-caused to Psalm 4:2 being
mis-styled `Psalms BOOK` instead of `VerseText` (fixed, and now guarded by
a permanent regression test, `Test 87`); the remaining 46 skips + 3
duplicates were entirely individual C:V marker digit typos (missing/
extra/substituted digits) plus one mis-typed Heading 2 ("MYCHAPTER 4")
that broke chapter tracking for all of 1 Peter 4 - all identified via the
new diagnostic logging, confirmed and fixed by the operator one at a time.
Release-process guard added (adaept5tudio doc + in-repo pointer + runtime
reminder), all ✅. Two real code bugs found and fixed along the way
(`GetMaxVerse` off-by-one; `VersesInChapter` error-handler class name/
`MsgBox`).

**Follow-up items, not blocking, tracked separately:**
1. ~~Run a full-suite `RUN_THE_TESTS` at least once to confirm Test 87
   doesn't reproduce the Tests-82/83-style full-suite memory blowup~~ **✅
   Done, 2026-09-16, across two full-suite runs.** Test 87 passed (`0=0`)
   both times - it did **not** hang or fail like Tests 82/83 did. Runtime
   varied across three total full-suite runs this session: `540.64s`
   (86-test, pre-Test-87), `1683.93s` (87-test, first run after adding
   Test 87 and the Test 16 rebaseline), `1336.95s` (87-test, after also
   fixing Tests 77/78). The middle run's `~1143s` jump initially looked
   concerning, but the third run landing in between weakens the case that
   Test 87 itself causes a stable slowdown - more consistent with ordinary
   run-to-run variance (Word's per-session performance was separately
   observed degrading elsewhere in this investigation, e.g.
   `AuditCharStyleUsage`'s accelerating per-batch slowdown). **Closed as
   non-issue** - no further action needed unless it recurs.
2. Decide whether to finally root-cause the `GetMarkerTotals` memory issue
   so Tests 82/83 themselves could be restored - not done, not decided.
3. `docm-verses.txt` is now clean (31102/0/0) - Phase 4's Pass 1/2 sync and
   Pass 3's divine-name census could be re-run against this corrected
   export if still relevant (see `rvw/Plan_rwb_phase4_content_sync_2026-09-15.md`
   and `rvw/Plan_pass3_divine_names_2026-09-15.md`).

**Unrelated side-finding while full-suite-verifying this bug's fix, ✅
closed:** Tests 77/78 (`CountApprovedStylesWithUnhideWhenUsedOn`/
`CountApprovedStylesWithWrongPriority`) failed on `Default Paragraph Font`
- an exact recurrence of a 2026-09-14 incident (`rvw/Code_review
2026-09-14.md`), most likely re-triggered by this session's heavy amount of
direct in-Word editing (any run being touched can make Word silently
re-surface an approved style - exactly what Test 77 exists to catch). Root
cause of the *recurrence*: the `UnhideWhenUsed` half of the 2026-09-14 fix
was only ever a manual Immediate-window line, never committed as reusable
code - so it had to be reconstructed from scratch instead of rerun. Fixed
properly this time, aeBibleClass `fcb8172`: folded `s.UnhideWhenUsed =
False` into `PromoteApprovedStyles`'s existing per-style loop (alongside
the `Priority` fix it already applied), so one call now fixes both
properties whenever this recurs. Verified: `PromoteApprovedStyles` run +
document saved + full-suite re-run - Tests 77/78 both `PASS` at `0`, full
suite otherwise clean (only Tests 82/83 skip, as already tracked above).
