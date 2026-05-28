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
