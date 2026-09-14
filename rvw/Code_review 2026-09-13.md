# Code review - 2026-09-13 carry-forward

This file opens a fresh review arc on 2026-09-13. The previous
carry-forward arc [`rvw/Code_review 2026-06-01.md`](Code_review%202026-06-01.md)
is now **closed for new work** and remains the authoritative dated
history for items 1-16 listed there. This session's scope was narrow
(Test 72 hang investigation, `ThisDocument.cls` license-header loss on
import/export), so **items 1-16 below are carried forward by reference,
not reverified** - several (notably item 1, the aeRibbon release track)
show evidence of further progress since 2026-06-01 (`aeRibbon/releases/
1.0.0+bc71416/BUILD_RECORD.txt` and `RELEASE_TRACK_CONTEXT.md` record a
G8 run against `Radiant-Word-Bible.docx` on 2026-09-12) that this
session did not review in detail. Re-verify status before acting on any
carried-forward item below.

Status tag legend (continued):

- **OPEN** - actively pending, all known prerequisites met.
- **PARTIAL** - partially complete; specific remaining work listed.
- **DEFERRED** - not started, waiting on a specific trigger.
- **FUTURE** - speculative; revisit only when conditions warrant.
- **DONE** - completed and verified this session.
- **CARRIED (UNVERIFIED)** - copied forward from the 2026-06-01 arc
  without reverification; status may have changed since.

## Open carry-forward (priority order)

### 1. Test 72 is testing the wrong thing - VerseText/PsalmSuperscription alignment scope gap (HIGH) - OPEN 2026-09-13

**Confirmed.** Test 72 (`HasLeftAlignedParagraph`, `src/aeBibleClass.cls`)
is out of date against a formatting decision already made in the
document: `PsalmSuperscription` paragraphs (e.g. "A Psalm by David, when
he fled from Absalom his son.", found at printed page 417) are now
**intentionally** left-aligned. Test 72's design predates that decision
- it was written as a coarse sanity check ("does *any* left-aligned
paragraph exist in the body range at all"), so once `PsalmSuperscription`
correctly started using left alignment, the test now PASSes on that
exact evidence while the number stored in `Expected1BasedArray` (`0`)
still reflects the old "no left alignment anywhere in the body" world -
i.e. the FAIL is a stale baseline, not a code defect. The fix applied
earlier this session (see item 3) made the test able to *run and report*
correctly; it did not, and was not asked to, fix what the test *checks*.

**Root design problem:** "at least one left-aligned paragraph exists"
is the wrong shape of assertion once more than one style is allowed to
be left-aligned by design. The rule that actually matters editorially is
narrower and per-style:

- **`VerseText` must always be justified.** No `VerseText` paragraph
  should ever be left- or right-aligned - this is the real integrity
  rule Test 72 was reaching for. Correct assertion: count of `VerseText`
  paragraphs with `Alignment <> wdAlignParagraphJustify` should be **0**.
- **`PsalmSuperscription` is now allowed/expected to be left-aligned.**
  This needs its own gate, not exclusion from Test 72's scope only by
  accident.
- **`Psalms BOOK`** (confirmed exact approved style name, see
  `GetApprovedStyles()` in `src/basTEST_aeBibleConfig.bas:58`) has a
  related but separately-scoped alignment expectation per the operator's
  note below.

**New tests needed (operator-specified 2026-09-13):**

i. Count of `PsalmSuperscription` paragraphs that are **not**
   left-aligned - expect **0** (baseline: all `PsalmSuperscription`
   paragraphs are left-aligned by design).
ii. Count of `Psalms BOOK`-styled paragraphs that are **not**
    left-aligned - expect **0**.
iii. Replacement for Test 72 itself: count of `VerseText` paragraphs
    that are **not** justified - expect **0**.

**Broader gap (operator's note):** this project has no per-approved-style
"expected alignment / expected formatting" registry - each style's
formatting contract is currently implicit (embedded in ad hoc tests like
Test 72, or in `AuditOneStyle` calls in `basTEST_aeBibleConfig.bas` which
check font/size/spacing but not alignment as a QA gate). Per-style
expected-alignment (and likely other per-style QA facts) should become a
first-class table, the same way `GetApprovedStyles()` /
`GetApprovedStylesByType()` are the SSOT for style membership - this
generalizes beyond alignment (e.g. keep-with-next, spacing) and beyond
this document, since the same registry is what a future non-English
translation would need to QA against to keep formatting correct.

**Implementation shape recommendation:** follow the bounded per-paragraph
`For Each ActiveDocument.Paragraphs` walk already proven in
`GetHeaderFooterStyleTotals`, `GetMarkerTotals`, and
`AuditOrphanBodyTextParagraphs` (see item 2, 2026-06-01 arc) rather than
a `Range.Find`-based scan - avoids both the `O(N^2)` trap flagged for
`AuditCharStyleUsage` and the page-navigation fragility that caused
items 3-4 below.

**Not implemented this session** - documented per operator request for
review before code changes. Suggested next step: confirm the three new
test specs above, then implement as Tests 85-87 (or renumber per Test
72's disposition - retire vs. repurpose to the `VerseText` check is an
open naming decision).

### 2. Test 72 magic numbers -> constants (MEDIUM) - OPEN 2026-09-13

The page-range boundary for Test 72 is duplicated as literal numbers in
four places, already caught drifting once this session (the
`GetTestDescription` text still says pages "18-931" while the live call
had already moved to `(19, 925)` by the time the hang was fixed - a
maintainer had edited the call site but not the other three copies):

- `src/aeBibleClass.cls:1042` - `GetTestDescription`, Case 72, prose text
- `src/aeBibleClass.cls:1271` - `GetPassFail`, Case 72, the actual call
  (`HasLeftAlignedParagraph(19, 925)`)
- `src/aeBibleClass.cls:1484` - `Debug.Print` report-row label (literal
  string duplicate of the call)
- `src/aeBibleClass.cls:1687` - `BufAppend` report-row label (literal
  string duplicate of the call)

**Fix direction:** add two class-level constants near the existing
`Private Const` block (e.g. `BIBLE_BODY_START_PAGE = 19`,
`BIBLE_BODY_END_PAGE = 925`), reference them from the call site, and
build the two report-label strings and the description text
dynamically from the same constants instead of repeating literals. Not
implemented this session - flagged for the same reason as item 1 (code
change, not requested yet this turn).

### 3. GoToAdjustedPage hang + navigation bug (HIGH) - DONE 2026-09-13

`RUN_THE_TESTS(72)` hung indefinitely. Root-caused to two stacked bugs
in `GoToAdjustedPage` (`src/aeBibleClass.cls`), both fixed:

- **Bug A (why it hung forever):** the loop had no termination guard
  beyond "adjusted page number equals target" - if the target page was
  never reached, `Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext`
  simply stops advancing at the last page without raising an error, so
  the `Do...Loop` spun forever.
- **Bug B (why it never reached the target even on the correct 900+
  page document):** `Which:=wdGoToNext` requires a prior *absolute*
  page-type `GoTo` to establish a navigation reference point. The
  function only did `sel.HomeKey wdStory` first (not a page-type GoTo),
  so `wdGoToNext` never advanced at all - the adjusted page number was
  stuck at `1` regardless of actual document length. Diagnosed live:
  first fix (a stall-guard around the `Next`-based loop) converted the
  hang into a fast, correct FAIL reporting "stalled at 1" in 0.18 s -
  which is what exposed Bug B.

**Fix:** rewrote `GoToAdjustedPage` to navigate by **absolute** physical
page number in a `For physPage = 1 To totalPages` loop
(`ActiveDocument.ComputeStatistics(wdStatisticPages)` for the bound),
using `Which:=wdGoToAbsolute, Count:=physPage` each iteration. This is
provably terminating (no guard needed - the loop is bounded by the
document's real page count) and sidesteps the `Next`-reference-point
quirk entirely. Verified live: Test 72 now completes in 1.38 s and
correctly finds/report page 417 (see item 1).

Also added a missing diagnostic hint: `HasLeftAlignedParagraph` now sets
`m_lastHint` to `"Page <n>: <first 60 chars>"` on a match, matching the
`m_lastHint` convention used by other Count functions (e.g.
`FormatEmptyRunHint`) - previously Test 72's FAIL rows always printed
"(no hint provided by test function)", which is what made item 1's root
cause (the Psalm superscription hit) invisible until this was added.

### 4. `ThisDocument.cls` license header silently dropped on every import/export (MEDIUM) - DONE 2026-09-13

`ImportAllVBAFiles` (`src/basImportWordGitFiles.bas`) explicitly skipped
`ThisDocument.cls` on every run:

```vb
Else
    colSkipped.Add strFile & " (ThisDocument)"
    intSkipped = intSkipped + 1
End If
```

`ThisDocument` is a built-in Word document module - `VBComponents.Import`
errors if you try to import over an existing component name, which is
presumably why it was skipped rather than handled. The effect: any edit
made to `src/ThisDocument.cls` (including the dual-license header added
under [[project_ribbon_dotm_docx_model]]-adjacent licensing work, item 16
of the 2026-06-01 arc) never reached the live VBA project, so every
export re-produced the stale, header-less version - discovered when the
header vanished from `src/ThisDocument.cls` twice in the same session
even after being manually restored in the repo file.

**Fix:** added `ImportThisDocumentFile` (`src/basImportWordGitFiles.bas`)
- reads the `.cls` text, strips the VBE-managed
`VERSION`/`BEGIN`/`END`/`Attribute` header lines (component metadata,
not valid `CodeModule` text), and replaces `ThisDocument`'s live
`CodeModule` in place via `DeleteLines` + `AddFromString` with everything
from `Option Explicit` onward. Verified live over two import/export
cycles - header (and a subsequent whitespace trim) now survives the
round-trip cleanly (`git diff` empty against HEAD both times).

**Caveat surfaced during rollout:** `basImportWordGitFiles.bas` is
itself one of the files `ImportAllVBAFiles` would need to already be
running the new version of to pick up its own fix - and
`DeleteAllModulesExceptImporter` explicitly protects it from
self-deletion/self-reimport by name. The operator had to manually
reload `basImportWordGitFiles.bas` once before the fix took effect. No
code action needed - just a one-time manual step, now done.

### 5. Tests 42, 51, 72 unskipped and rebaselined - CONFIRMED intentional (MEDIUM) - RESOLVED 2026-09-13

Working-tree changes present at the start of this session (from an
operator export, not from Claude) touched `Expected1BasedArray` and
`MakeSkipTestArray` in `src/aeBibleClass.cls`:

- **`SkipTestArray`** changed from `Array(42, 51, 72)` to `Array()`
  (`src/aeBibleClass.cls:352`, old value commented out in place) - all
  three tests unskipped together, not just 72.
- **`Expected1BasedArray`** (`src/aeBibleClass.cls:250`) values changed
  for Test 16 (`33822` -> `33827`), Test 37 (`19` -> `99`), Test 49
  (`15` -> `42`), and Test 50 (`147` -> `145`).

**Confirmed by operator 2026-09-13:** this was deliberate - Tests 42
(`CountBoldFootnotesWordLevel`) and 51 (`CountAndCreateDefinitionForH2`)
were unskipped and their results already verified directly by the
operator (independent of this session's Test 72 work). This is
progress toward closing the "tests that only run as `SKIP!!!!`" gap.
Expected to need re-verification again after further document edits -
that is normal baseline churn, not drift to chase down. No further
action needed on this item; safe to include in the next commit.

## Carried forward from 2026-06-01 (not reverified this session)

The following items are copied forward by reference only. See
[`Code_review 2026-06-01.md`](Code_review%202026-06-01.md) for full
detail on each - status tags below are as of 2026-06-01 and have not
been re-checked:

1. Run aeRibbon Gates G1-G8 and ship v1.0.0 (HIGH) - was "PREP DONE,
   READY FOR BUILD"; **note:** `aeRibbon/releases/1.0.0+bc71416/
   BUILD_RECORD.txt` and a new `RELEASE_TRACK_CONTEXT.md` show a G8 run
   recorded 2026-09-12, i.e. this has very likely progressed since -
   re-read those files before treating this as still at prep stage.
2. `AuditCharStyleUsage` quadratic-time fix (HIGH) - CARRIED (UNVERIFIED).
3. Header/Footer + Section SHAPE LOCKDOWN (MEDIUM) - CARRIED (UNVERIFIED).
4. Revisit failed tests and verify status/code/performance (MEDIUM) -
   CARRIED (UNVERIFIED); item 5 above (this file) is a direct instance
   of this standing task and should probably be folded into it next time
   both are touched.
5. Date-rule sweep follow-ups (MEDIUM) - CARRIED (UNVERIFIED).
6. File-write code audit against FSO rule (MEDIUM) - CARRIED (UNVERIFIED).
7. +1 CVM anomaly at ParaStart=3087864 (editorial) - CARRIED (UNVERIFIED).
8. EDSG `10-list-paragraph-bug.md` Step 0 snippet correction (LOW) -
   CARRIED (UNVERIFIED).
9. Test 38 kind-distribution + structural-phrasing follow-ups (LOW) -
   CARRIED (UNVERIFIED).
10. Normal style audit (LOW, DEFERRED) - CARRIED (UNVERIFIED).
11. Finding 5 (ribbon nav) - Word limitation, no action available.
12. SoftHyphenSweep_FootnotesOnly sister routine (DEFERRED).
13. SHA_ReplaceHard i18n consideration (FUTURE).
14. Architecture rule - class encapsulation + module/class safety
    boundary (RULE) - standing rule, see [[feedback_class_encapsulation]].
15. aeRWB source-text repo relationship (RULE/NOTE) - CARRIED
    (UNVERIFIED); operator's fuller linkage details may have since
    arrived.
16. Dual-license headers (LICENSING) - was "DONE; CLASS + PRODUCTION
    .bas COVERAGE COMPLETE" as of 2026-06-17; **directly relevant to
    item 4 above** - the coverage was complete in the repo files, but
    this session found the *live VBA project's* `ThisDocument` never
    actually received it due to the import skip, now fixed.

## Pointer back to the closed arc

Full dated history through 2026-06-01 is in
[`rvw/Code_review 2026-06-01.md`](Code_review%202026-06-01.md), which
itself points back through the 2026-05-28, 2026-05-16, and earlier arcs.
