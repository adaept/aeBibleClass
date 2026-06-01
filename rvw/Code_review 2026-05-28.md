# Code review - 2026-05-28 carry-forward

This file opens a fresh review arc on 2026-05-28. The previous
arc [`rvw/Code_review 2026-05-16.md`](Code_review%202026-05-16.md)
is now **closed for new work**; that file remains the
authoritative dated history for everything between 2026-05-16
and 2026-05-28, including:

- **Ribbon v1.0.0 prep finished 2026-05-17** - `BUILD.md` step-3
  count fixed (4 `.bas` + 2 `.cls`); trim re-run clean (133
  kept, 70 removed); small-refresh release classification.
- **Hide-sweep wired into `WordEditingConfig` 2026-05-17.**
  `HideUnapprovedBuiltInStyles` now runs between
  `PromoteApprovedStyles` and `DumpPrioritiesSorted`. Test 45
  PASS at 0 on the live document.
- **TestReport.txt per-slot descriptions** populated for all 74
  slots across nine batches; baseline numbers removed from
  description text per operator-rebaseline lesson.
- **Test 75 / 76 / 77 / 78 added** - approved-cohort discipline
  gated end-to-end: `BaseStyle=""` + no LinkToListTemplate
  (75), `AutomaticallyUpdate=False` (76), `UnhideWhenUsed=
  False` (77), Priority equality to array position (78).
- **Three persistently-missing placeholders retired
  2026-05-17** - `AppendixTitle`, `AppendixBody`,
  `BodyTextContinuation` removed from `GetApprovedStyles`;
  taxonomy reached first fully-clean run in project history.
- **Test 79 added 2026-05-20** - `CountNumericOrdinals`
  relocated from `Module1.bas` into `aeBibleClass.cls`; date-
  class numeric-ordinal metric. Coexists with the broader
  module-level `Test_NoSuperscriptOrdinals` (2026-05-19).
- **2026-05-28 session** - Introduction SpaceBefore spec
  realigned 0 -> 12 to match operator's style update;
  `Default Paragraph Font` promoted to approved (array position
  34) so character-style inheritance into `VerseText`
  paragraphs stays functional. Taxonomy at 57 PASS / 0 FAIL,
  Test 78 PASS.

Status tag legend (continued):

- **OPEN** - actively pending, all known prerequisites met.
- **PARTIAL** - partially complete; specific remaining work listed.
- **DEFERRED** - not started, waiting on a specific trigger.
- **FUTURE** - speculative; revisit only when conditions warrant.
- **RECOVERED** - surfaced from a prior arc where it was dropped
  off the carry-forward chain.

## Open carry-forward (priority order)

### 1. Run aeRibbon Gates G1-G8 and ship v1.0.0 (HIGH) - PREP DONE 2026-05-17, READY FOR BUILD

Still the next active release-track item. Build-side prep
(trim, BUILD.md correction, small-refresh classification)
remains valid; nothing in the 2026-05-28 session touched ribbon
code. G8 spot-check on the new book aliases (`JSH`, `JDG`, etc.)
still queued.

**Operator action** (Word GUI work; not driveable from here):

1. Build `aeRibbon/template/aeRibbon.dotm` per
   `aeRibbon/BUILD.md` steps 1-8.
2. Editor/Developer produces the production Bible `.docx` per
   `BUILD.md` "Producing the production Bible `.docx`".
3. Run Gates G1-G8 from `aeRibbon/QA_CHECKLIST.md`. Record
   results in
   `aeRibbon/releases/1.0.0+bc71416/BUILD_RECORD.txt`.
4. Append a row to `aeRibbon/RELEASES.md` and
   `git tag v1.0.0+bc71416`.

Full prior analysis and gate definitions: see § 1 in
[`Code_review 2026-05-16.md`](Code_review%202026-05-16.md).

### 2. Revisit failed tests and verify status / code / performance (MEDIUM) - OPEN

Carry-forward from 2026-05-16 § 12. Trigger remains the same:
next time a slot FAILs, walk the function before rebaselining.
Slot-by-slot revisits will surface candidates for the same kind
of split applied to Test 22 (one bundled signal -> two or three
disjoint ones).

Known candidates from prior arcs:

- **Test 30 source-comment vs expected mismatch.**
  `CountHeaderStyleUsage` function comment says "Expected = 0"
  but the `values` array baseline is 70. Decide whether the
  rule was relaxed (update comment) or the expected drifted
  (rebaseline or restore).
- **Count-baseline tests 24, 27, 29, 30, 32-35, 37, 47, 49,
  50, 51** - drift could be either editorial or accidental;
  walk the function before bumping the expected.
- **Slow slots** - apply the Test 22 perf lens (cheap length
  guard, iterate the small collection, avoid Range.Text when
  Len(Range.Text) will do).

### 3. Date-rule sweep follow-ups (MEDIUM) - OPEN

Carry-forward from 2026-05-19 / 2026-05-20.

- **Pair 06 in `Date_Example.txt` still `pending`** - operator
  to decide whether the 300-600 AD range stands or needs
  amendment; rerun `ApplyDateRule_2026_05_19` once decided.
- **Apply the date rule to the 20 example passages in the live
  document**; rerun `Test_NoSuperscriptOrdinals` and Test 79
  and confirm the counts move toward the target deltas.
- **Book-number ordinal policy still DEFERRED** (`1st Samuel`
  vs `1 Samuel`). Once decided, either extend
  `Test_NoSuperscriptOrdinals` to enforce zero, or add a
  matching pre-pass that strips the suffixes. Test 79
  (digit-prefixed only) is already in place; the module-level
  sweep is the broader one to gate.

Target end state:

- Test 79 (`CountNumericOrdinals`): 0 (date-class only).
- `Test_NoSuperscriptOrdinals`: 0 only once the book-number
  ordinal policy is decided and applied.

### 4. EDSG `10-list-paragraph-bug.md` Step 0 snippet correction (LOW) - OPEN 2026-05-28

Discovered during the 2026-05-28 audit analysis. The Step 0
diagnostic snippet uses `Not (s.LinkToListTemplate Is Nothing)`
as a read-side check. `Style.LinkToListTemplate` is **write-
only** in Word's object model (compile error "argument not
optional" if invoked as a getter). The working read-side
detection is the `ListTemplates -> ListLevels -> LinkedStyle`
graph traversal already used in Test 75
(`CountApprovedStylesWithListParagraphRisk`).

**Action:** update the EDSG snippet to mirror the Test 75
implementation. Conceptual framing of Step 0 stands; only the
code line needs replacement.

**Originated:** noted in 2026-05-16 arc closing entry; logged
here for follow-through on the next EDSG pass.

### 5. Finding 5 (ribbon nav) - umbrella OPEN (DEFERRED, WORD LIMITATION) - RECOVERED

Word-side limitation; no action available. Remains in the
register for awareness.

### 6. SoftHyphenSweep_FootnotesOnly sister routine (DEFERRED)

Surfaced during the 2026-05-08 SHA build; waiting on a
footnote-specific trigger before implementation.

### 7. SHA_ReplaceHard i18n consideration (FUTURE)

Speculative; revisit when a non-English target translation
materialises.

### 8. Normal style audit (LOW, DEFERRED)

Carry-forward from 2026-05-16 § 17 sub-item 1. `Normal` is
intentionally unaudited as the "pin-everything-else-above"
anchor. A bucket-1 entry would need a `DumpStyleProperties
"Normal", True` capture and a new `AuditOneStyle` line. Worth
doing opportunistically next time `Normal` is touched, not
chasing as its own work item.

**Note:** the 2026-05-28 `DumpAllApprovedStyles` output showed
`Normal.QuickStyle` flipped True -> False vs the previously
committed dump; audit does not check QuickStyle, so no spec
change required, but it is a hint that `Normal` is drifting
quietly. Bump in priority if a second drift surfaces.

### 9. Architecture rule - class encapsulation + module/class as casual-coder safety boundary (RULE, 2026-05-15)

Codified as a feedback memory and documented in the 2026-05-15
arc. Standing rule, not an action item - listed here so it
remains visible during slot-by-slot review work.

Full rule and worked examples: see § 9 in
[`Code_review 2026-05-15.md`](Code_review%202026-05-15.md).

### 10. AuditCharStyleUsage quadratic-time fix (HIGH) - OPEN 2026-05-31

`basVerseStructureAudit.AuditCharStyleUsage` walks character-
style runs via document-scope `Range.Find`, then calls
`oRng.Paragraphs(1)` after every match to resolve the enclosing
paragraph. Word resolves that lookup by walking the Paragraphs
collection from doc start, so cost per match is `O(position-in-
doc)` and the total scan is `O(N²)` where N = match count.

**Observed on 2026-05-31 live doc:**

- CVM run (31,103 matches): completed in 406 s.
- VM run (same shape): heartbeat 1k-runs/5s early, climbed to
  >60 s per 1k-runs batch deeper in; aborted before completion.

The performance is unacceptable now and will degrade further as
the document grows.

**Fix direction:** rewrite to walk `ActiveDocument.Paragraphs`
once via `For Each`, and inside each paragraph do a paragraph-
scoped `Range.Find` (or `Characters(1)` / `Characters(Last)`
checks for the common START/END cases). Same bounded-per-
paragraph shape already proven in `EnsureVerseMarkerCounts`
(slots 82+83) and `AuditOrphanBodyTextParagraphs`. Eliminates
the `oRng.Paragraphs(1)` lookup entirely.

**Out of scope for this session.** Diagnostic was usable in
ANOMALIES-ONLY mode for the CVM run; the +1 anomaly was found
via the CVM scan before the curve became prohibitive.

### 11. File-write code audit against FSO rule (MEDIUM) - OPEN 2026-05-31

Per [[feedback_fso_file_writes]]: `rpt/` writers in re-callable
routines must use `FSO.CreateTextFile`, not `Open ... For
Output As`, to avoid Err 55/70 from leaked handles and OS share
locks. Several writers were converted in earlier sessions; a
full sweep has not been done.

**Action:** grep all `.bas` and `.cls` files for `Open ` and
`Output As` / `Append As` patterns. For every writer that
targets `rpt/` and may be invoked more than once per session,
convert to the FSO pattern used in
`CountAuditStyles_ToFile` / `CountAuditCharacterStyles_ToFile` /
`WriteCharStyleUsageFile`. Leave alone any genuine one-shot
diagnostic writers that can't be re-entered.

## 2026-05-31 - Tests 81-83 added (character-style audits) + +1 CVM anomaly found

Three new test slots wire character-style coverage into the
audit pipeline alongside the existing paragraph-style coverage
(Test 49 baseline 51, etc.). All three required a `MaxTests`
bump from 80 to 83 and matching extensions of `ResultArray`,
`m_HintArray`, `GetPassFailArray`, `TestTimingArray`, and the
1-based `values` baseline array.

**Test 81 - `CountAuditCharacterStyles_ToFile` (PASS at 0).**
Presence audit. For every approved character style returned by
the new `basTEST_aeBibleConfig.GetApprovedStylesByType(wd
StyleTypeCharacter)` helper, runs a single bounded `Find.
Execute` against `ActiveDocument.Content` and records
Present/ABSENT to `rpt\Character Style Usage.txt`. Returns the
count of absent (expected 0). Runtime ~1 s. Drift surfaces as
"first absent: <style>" hint.

**Test 82 - `CountVerseMarker` (PASS at 31,102).**
**Test 83 - `CountChapterVerseMarker` (PASS at 31,102).**
Both back into `basVerseStructureAudit.GetMarkerTotals` (new
public helper) which walks `ActiveDocument.Paragraphs` once.
For every `VerseText` paragraph it checks
`Characters(1).Style.NameLocal == "Chapter Verse marker"` and
scans `Characters(1..12)` for `"Verse marker"`. Module-level
cache (keyed by `ActiveDocument.FullName`) makes slot 83 a
sub-second read after slot 82 populates it, even across
separate `aeBibleClass` instances (OneTest mode reinstantiates
the class per test).

**Algorithm-iteration history** (preserved for the next time
someone reaches for document-scope `Range.Find` on a 31k-
matches problem):

1. Single document-scope `Range.Find` loop with `.Text=""` +
   `.Style=s` + `.Collapse wdCollapseEnd`: ran to 2.7 GB
   memory leak on the linked-styles superset.
2. Narrowed allowlist to `wdStyleTypeCharacter`: dropped to
   600 MB still climbing.
3. Per-paragraph `Range.Find` on `p.Range.Duplicate`: same
   shape leak (31k transient Range RCWs).
4. Document-scope `Range.Find` with `.Format=True` and
   explicit `oRng.Start = oRng.End` / `oRng.End = endPos` re-
   bound (pattern lifted from
   `basVerseStructureAudit.CountVerseMarkers`): completed -
   one style at 205 s, the other at 2730 s (45 min). Counts
   31,103/31,103.
5. Per-chapter scoping of the same proven Find pattern via
   `GetMarkerTotals` walking H2 boundaries: 326 s + 170 s. Both
   31,103 still.
6. Paragraph-walk with `Characters(i).Style.NameLocal` check
   on VerseText paragraphs only (current implementation):
   222 s first run, 0.02 s cached. Counts 31,102/31,102 -
   matches the canonical baseline.

The semantic shift between (4-5) and (6) is intentional and
documented in `GetMarkerTotals`'s header: the slot-82/83 metric
is "VerseText paragraphs satisfying the design rule" not "total
CVM/VM runs anywhere in the document." The latter is what
surfaced the +1 drift below.

**+1 drift found at `ParaStart=3087864` (open editorial item).**

Find-based total counts (31,103) differed from paragraph-rule
counts (31,102) by exactly one. Running the new ANOMALIES-ONLY
mode of `AuditCharStyleUsage` (`bAnomaliesOnly:=True`, third
parameter) on `"Chapter Verse marker"` surfaced the locus:

```
Run #22396 | ParaStart=3087864 | Style=VerseText | first-char-style=Chapter Verse marker | Phase2: KEEP-AS-VerseText
  Chapter Verse marker at END of paragraph (offset 146 of 147)
  Excerpt: "215 neither shall he stand who handles the bow; and he who is swift of foot will"
```

A single stray CVM-styled character at the tail of one
VerseText paragraph (offset 146 of 147 - last character before
the paragraph mark). The paragraph itself is structurally
correct (CVM at char 1, VM nearby). The pair VM partner was not
confirmed; the VM ANOMALIES-ONLY scan hit the `O(N²)`
slowdown described in § 10 and was aborted. Operator can
navigate via `GoToPos 3087864` and inspect the end-of-
paragraph character manually.

**`AuditCharStyleUsage` improvements applied 2026-05-31:**

- New optional 3rd parameter `bAnomaliesOnly As Boolean = False`.
  When True, suppresses per-run dumps for the expected
  case (`paraStyle=VerseText` AND `posLabel=START`), and only
  emits detail blocks for anomalies. Auto-bumps the safety cap
  from 5,000 to 100,000 to cover full-doc scans without
  exhausting the report buffer.
- Summary now reports
  `Anomalies (paraStyle<>VerseText OR position<>START): N`.
- Report file now embeds start time, finish time, and duration
  at the tail (was start time only in header).
- Progress heartbeat every 1,000 runs via `Debug.Print` +
  `DoEvents` - made the `O(N²)` curve in § 10 visible (and
  Ctrl+Break responsive).

**Wiring touched in `aeBibleClass.cls`:**

- `MaxTests` 80 -> 83.
- Three new module-level state vars: `m_verseScanDone`,
  `m_verseMarkerCount`, `m_chapterVerseMarkerCount` (redundant
  inner cache on top of `basVerseStructureAudit`'s module-
  level cache; harmless).
- `GetTestDescription` Case 81 / 82 / 83 added.
- `GetPassFail` Case 81 / 82 / 83 dispatch + `m_HintArray`
  copy.
- `Debug.Print` and `BufAppend` Case 81 / 82 / 83 rows.
- `Expected1BasedArray` `values` array extended with
  `0, 31102, 31102`.
- New private routines: `CountAuditCharacterStyles_ToFile`,
  `CountVerseMarker`, `CountChapterVerseMarker`,
  `EnsureVerseMarkerCounts`.

**Wiring touched in `basTEST_aeBibleConfig.bas`:**

- New `Public Function GetApprovedStylesByType(wantType As
  WdStyleType) As Variant`. Walks `GetApprovedStyles()` and
  asks Word for each style's `.Type`. SSOT stays the flat name
  list; this helper materializes the para/char taxonomy at
  runtime so neither slot 81 nor any future consumer needs a
  hand-maintained second list.

**Wiring touched in `basVerseStructureAudit.bas`:**

- New `Public Sub GetMarkerTotals(vmTotal, cvmTotal)` plus
  module-level cache vars `m_cachedDocName`, `m_cachedVMTotal`,
  `m_cachedCVMTotal`, `m_cacheValid`. The cache invalidates
  when `ActiveDocument.FullName` changes.
- `AuditCharStyleUsage` extended with `bAnomaliesOnly`,
  progress heartbeat, and timing tail (per § 10 follow-up).

**Verification (live document, 2026-05-31):**

- Test 81 PASS at 0 in ~1 s (presence audit clean).
- Test 82 PASS at 31,102 in 222 s (first run, populates cache).
- Test 83 PASS at 31,102 in 0.02 s (cache hit).
- `AuditUnconvertedVerseParagraphs True` returns 0 - confirms
  no CVM-led paragraphs sit outside `BodyText` / `VerseText`.
- `AuditCharStyleUsage "Chapter Verse marker", True, True`
  surfaced the single anomaly above (406 s, completes despite
  the quadratic curve).

**Follow-ups (open).**

- § 10 above: rewrite `AuditCharStyleUsage` to walk paragraphs
  instead of doc-scope Find.
- Operator to inspect `ParaStart=3087864` and decide whether to
  clean the stray CVM character or accept and rebaseline a
  future "total VM/CVM runs in doc" diagnostic to 31,103.
- VM anomaly partner unconfirmed; will likely surface at the
  same `ParaStart` once § 10 fix lands and the VM scan can
  complete.

## 2026-05-30 - Test 11 / 33 / 38 hint pass + Test 80 added (Bare empty para split)

Three hint-or-diagnostic gaps closed across the existing slot
taxonomy, plus a Test 22-style split applied to Test 38 once the
underlying structural reality came to light.

**Test 11 - hint added to `CountFindNumberDashNumber`.**
Mirrors the Test 79 (`CountNumericOrdinals`) convention. On the
first `[0-9]+-[0-9]+` match, sets `m_lastHint = "page <N> :
<match-text>"` via `rng.Information(wdActiveEndPageNumber)`.
Previously the FAIL row printed `(no hint provided by test
function)`.

**Test 33 - hint bug fixed in `CountLinefeed`.**
`CountLinefeed` is shared by Test 32 (no arg, finds `^l`) and
Test 33 (`" "` arg, finds `" {1}^l"`). The pre-fix hint captured
the first match of whatever pattern was active, so for Test 33
it always returned a *passing* (space-preceded) instance.

Fix: gated the primary-loop hint capture to the Test 32 path
(`IsMissing(space)`), then added a second pass in the Test 33
branch that iterates `^l` matches, inspects the character at
`scanRng.Start - 1`, and records the first instance whose
predecessor is **not** a space. Hint shape:
`page <N> : prev="<char>" ctx=<±20-char window>`. On a PASS run
no violation exists and `m_lastHint` stays empty - correct
signal.

**Test 38 - structural reality surfaced; description corrected;
Test 80 added as the `Bare`-only split.**

Originated as "Test 38 has no hint" and grew into a multi-step
investigation:

1. **First hint added** captured the first occurrence with page,
   paragraph index, and preceding-paragraph snippet. Operator
   reported the hint pointed at paragraph #1 (an existing
   baseline empty), not the +1 drift - the dataset was too dense
   for a single-hit hint to be diagnostic.

2. **Switched to per-occurrence dump** at
   `rpt\EmptyParagraphs.txt` (TSV, overwrite each run, same
   convention as `StyleTaxonomyAudit.txt`). Hint became
   `<N> empty paragraphs - see rpt\EmptyParagraphs.txt`.
   Operator-side workflow: re-run, `git diff` the file, the
   drift row stands out.

3. **CRLF / TSV hygiene** - first dump produced `\r\r\n` line
   endings on rows whose snippet ended with a pilcrow `\r`
   (VBA's `Print #` appends its own `\r\n` after data). Some
   snippets also contained embedded tabs that corrupted the
   TSV. Fixed by trailing-semicolon `Print #` + explicit
   `vbCrLf`, plus snippet sanitisation
   (`Replace vbCr/vbLf/vbTab -> " "` then `RTrim$`). Operator
   verification: a 1-row diff exactly matched a single
   document edit (tab added to a previously-empty paragraph).

4. **Classification column added** to surface what kinds of
   "empty" paragraphs exist. First attempt used:

   - `ParagraphFormat.PageBreakBefore` for **PBB**, and
   - `para.Range.End = Sections(1).Range.End` for **section-end**.

   Result: 143 `Bare` + 2 `PBB`, no section labels at all -
   wrong. The positional equality is unreliable (off-by-one
   when the section break char isn't included in the
   paragraph's range).

5. **Second attempt** compared section indices of this
   paragraph and the next - also wrong. Diagnostic columns
   added (`this_sec`, `next_sec`, `nx_ch`, `pbb`) to the
   report revealed the actual structure: **`this_sec ==
   next_sec` for every row** (the section transition happens
   across multiple Word paragraphs, not at the empty-paragraph
   boundary), and **`nx_ch = 12`** (the page/section break
   char) on essentially every row. The empty paragraph is
   *followed by* the break char, not at the section end in
   the index sense.

6. **Correct predicate** (applied 2026-05-30):

   - `PBB` if `ParagraphFormat.PageBreakBefore = True`.
   - Else if char at `para.Range.End` is `Chr(12)`:
     cross-reference a pre-built `secStartPos() ->
     secStartType()` map; if a section starts at
     `para.Range.End + 1`, classify by `SectionStart` enum
     (`SBNP` / `SBC` / `SBEP` / `SBOP` / `SBNC`); otherwise
     plain page break `PB`.
   - Else if char = `Chr(14)`: `CB`.
   - Else: `Bare`.

   Labels match Word's Layout > Breaks menu. Distribution on
   the live document after the fix: **142 SBNP, 2 PBB, 1 SBC,
   0 Bare**. Every flagged "empty" paragraph is a structural
   carrier; the test was measuring section/page-break density,
   not stray pilcrows.

7. **File-IO hardening surfaced en route.** A mid-run failure
   left a VBA file handle open, causing Err 55 ("File already
   open") on the next invocation. PROC_ERR updated to close
   the file handle defensively. Then a follow-up run hit Err
   70 ("Permission denied") at the `Open For Output` call -
   resolved by switching to
   `Scripting.FileSystemObject.CreateTextFile` (late binding,
   same pattern as `basStyleInspector.bas:179`). FSO opens
   the file with more permissive share semantics and writes
   `\r\n` natively, removing the `Print #` semicolon dance
   for line termination. The Err 70 itself ultimately
   required a Word restart to clear the orphan lock;
   investigation didn't pinpoint the holder (suspected WSL
   read-side cache).

**Test 38 description corrected.** Was: "Rule: Bare empty
paragraphs (pilcrow only) at the accepted baseline". Now:
"Count of empty paragraphs (Range.Text = Chr(13) only); in
practice almost all are structural carriers for Section/Page
breaks (SBNP/SBC/PBB) classified in rpt\EmptyParagraphs.txt -
CountEmptyParagraphs should match the expected baseline.
Paired with Test 80 for the truly-bare subset." Acknowledges
the structural reality without changing the baseline.

**Test 80 added - `CountBareEmptyParagraphs`.** Mirrors the
Test 22 -> 22 / 38 / 74 split pattern from the 2026-05-15 arc:
Test 38 keeps the structural-inclusive baseline (216), Test 80
isolates the truly-bare residual. Predicate: `PBB = False` AND
char at `Range.End` is neither `Chr(12)` nor `Chr(14)` AND
`Range.Text = Chr(13)`. Expected 0. On a Bare hit, sets
`m_lastHint = "page <N> : paragraph #<idx>"`. Standalone walk -
no ordering dependency on Test 38.

Wiring touched all four switches in one pass:

- `MaxTests` 79 -> 80.
- `Expected1BasedArray` values array extended with `0`.
- `GetTestDescription` Case 80 - "Rule: No truly-bare empty
  paragraphs (pilcrow only, not hosting a Section/Page/Column
  break - subset of Test 38 filtered to kind=Bare) -
  CountBareEmptyParagraphs should return 0".
- `GetPassFail` Case 80 dispatch + `m_HintArray` copy.
- `Debug.Print` and `BufAppend` Case 80 rows.

**Verification (live document):**

- `RUN_THE_TESTS 11` - PASS with hint surfacing the first
  digit-dash-digit token.
- `RUN_THE_TESTS 33` - PASS path keeps `m_lastHint` empty
  (no non-space-preceded linefeed exists); a FAIL would now
  point at the offender.
- `RUN_THE_TESTS 38` - PASS at 216; `rpt\EmptyParagraphs.txt`
  distribution 142 SBNP / 2 PBB / 1 SBC / 0 Bare.
- `RUN_THE_TESTS 80` - PASS at 0 (consistent with the 0 Bare
  count from the report).

**Coverage closed:** the FAIL desensitisation that prompted the
session - "Test 38 says +1 but the hint points at an existing
baseline entry" - is now structurally addressed. Future drifts
land in one of two places: Test 38 baseline (a new structural
carrier appeared) or Test 80 hit (a stray bare paragraph appeared,
hint provided), making the editorial decision actionable.

**Follow-ups (open).**

- Decide whether `kind`-distribution drift (e.g., SBNP vs SBC
  ratio changes between runs) should itself be a tested
  invariant. Currently the report is freeform; only the total
  count is gated by Test 38.
- The structural-reality phrasing in Test 38's description may
  be worth promoting into `EDSG/01-styles.md` or
  `EDSG/04-qa-workflow.md` if operators reference test
  descriptions in QA narrative.

## Pointer back to the closed arc

Full dated history of the work that produced this carry-forward
state is in [`rvw/Code_review 2026-05-16.md`](Code_review%202026-05-16.md).
That file (and the arcs it points back to) covers:

- The hide-sweep wiring into `WordEditingConfig`.
- The Test 75 / 76 / 77 / 78 approved-cohort discipline arc.
- The retirement of the three persistently-missing placeholders
  (AppendixTitle, AppendixBody, BodyTextContinuation).
- Test 79 add + Module1 retirement step.
- The date-formatting rule and `Test_NoSuperscriptOrdinals`
  add.
- The 2026-05-28 Introduction SpaceBefore + Default Paragraph
  Font promotion session.
