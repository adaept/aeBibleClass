# Code review - 2026-05-15 carry-forward

This file opens a fresh review arc on 2026-05-15. The previous arc
[`rvw/Code_review 2026-05-14.md`](Code_review%202026-05-14.md) is now
**closed for new work**; that file remains the authoritative dated
history for everything between 2026-05-14 and 2026-05-15, including:

- **BookHyperlink custom style** replaces use of Word's built-in
  `Hyperlink` character style for the doc's one-form hyperlink
  rule. Pins all four font properties explicitly (Carlito 9 +
  palette DarkBlue + single underline). `DefineBookHyperlinkStyle`,
  `LockBookHyperlinks`, `AuditBookHyperlinkStyling` are the trio
  of routines. Earlier `LockHyperlinksToPalette` /
  `AuditHyperlinkStyling` retired; `LockHyperlinksAlwaysBlue` alias
  deleted.
- **Built-in Hyperlink no longer touched.** The earlier approach
  modified `Styles("Hyperlink").Font` directly; this proved fragile
  (Word can reset it on theme operations) and inadequate (can't
  enforce font/size against paragraph inheritance). Built-in style
  now belongs to the upcoming hide-sweep bucket alongside Heading
  4-9 and Office365 collaboration styles.
- **`AuditNonPaletteStyleColors` Font.Name and Font.Size checks
  added** earlier in the session, which is what surfaced the
  inheritance bug that drove the BookHyperlink refactor.
- **`py/normalize_vba.py` extended** with `Word.Field` casing
  rules (one direct property pattern, one `As Word.Field`
  declaration pattern). Pre-existing lowercase occurrences in
  3 files normalised on one pass.

Status tag legend (continued from 2026-05-14 arc):

- **OPEN** — actively pending, all known prerequisites met.
- **PARTIAL** — partially complete; specific remaining work listed.
- **DEFERRED** — not started, waiting on a specific trigger.
- **FUTURE** — speculative; revisit only when conditions warrant.
- **RECOVERED** — surfaced from a prior arc where it was dropped
  off the carry-forward chain.

## Open carry-forward (priority order)

### 1. Run aeRibbon Gates G1-G8 and ship v1.0.0 (HIGH) - OPEN

The production export gateway is in place; nothing has been built
or gated yet. **Next active release-track item** and gates the
hand-off to the author for comments-only review.

**Why high:** every other ribbon-side improvement (signing,
auto-docx-from-docm, ribbon UX iteration) sits behind a first
successful gated build. Also the highest-leverage validation of
the trim script: any false drop will surface in G6 (compile) or
G8 (navigation smoke).

**Action:**

1. Build `aeRibbon/template/aeRibbon.dotm` per `aeRibbon/BUILD.md`.
   - Inject `aeRibbon/template/customUI14.xml` via
     `wsl python3 py/inject_ribbon.py`.
   - Import the 5 files from `aeRibbon/src/` into the template VBA
     project.
   - Set `RIBBON_VERSION` constant + custom property
     `aeRibbonVersion` to match `aeRibbon/VERSION`
     (`1.0.0+bc71416`).
   - Debug -> Compile VBAProject: must be zero errors.
2. Editor/Developer produces the production Bible `.docx` per
   `BUILD.md` "Producing the production Bible `.docx`" (manual
   File -> Save As `.docx` from the dev `.docm` - Option 1).
3. Run Gates G1-G8 from `aeRibbon/QA_CHECKLIST.md`. Record results
   in `aeRibbon/releases/1.0.0+bc71416/BUILD_RECORD.txt`.
4. Append a row to `aeRibbon/RELEASES.md`.
   `git tag v1.0.0+bc71416`.

**Expected blockers / what to watch for:**

- Trim script's call-graph is conservative-overinclusive; false
  drops surface as missing-Sub compile errors in G6. Fix by adding
  an explicit root rather than reverting to manual cherry-pick.
- VBA lifecycle hooks (`Class_Initialize`, `Class_Terminate`,
  `AutoExec`) are now always preserved if defined.
- G8 must show no macro-security warning on docx open - confirms
  the dotm/docx split is architecturally clean.

Originated 2026-05-12 with the gateway commits `bc71416` + `70bcff3`.
Carried unchanged through 2026-05-14.

### 2. Item 13 remaining work — built-in hide-sweep + test wiring (MEDIUM) - PARTIAL

Pass 1 of item 13 closed 2026-05-14 (`AuditNonPaletteStyleColors`
added, custom-style anomaly count brought to 0 by deleting the
orphan `Error` style). The 2026-05-15 BookHyperlink refactor
*strengthens* the case for hiding built-in `Hyperlink` and
`FollowedHyperlink` since neither is now in use - they must be
hidden so authors can't pick them from the Style gallery and
accidentally reintroduce the inheritance bug.

Remaining work, priority order:

**2.1 Hide-sweep for Word built-in noise (MEDIUM).** 122+ built-in
styles are skipped by the audit because they're not under
editorial control. After the 2026-05-15 BookHyperlink work, also
explicitly hide:

- `Hyperlink` (built-in, now superseded by `BookHyperlink`)
- `FollowedHyperlink` (built-in, no longer used)

New routine `HideUnapprovedBuiltInStyles` in `basStyleInspector`.
The approved-styles list is the SSOT defined in
`basTEST_aeBibleConfig.GetApprovedStyles()` (extracted 2026-05-15
from the prior in-line `Array(...)` inside `PromoteApprovedStyles`).
For every `BuiltIn = True` style whose name is NOT returned by
`GetApprovedStyles`, set:

```vba
.Priority = 99
.QuickStyle = False
.UnhideWhenUsed = False
```

The three-property pattern (not just Priority) matters because
`UnhideWhenUsed = True` re-surfaces a style in the gallery the
moment any run touches it - including paste operations.

Built-in styles that ARE in `GetApprovedStyles` (and therefore
left visible by the sweep): `Normal`, `Title`, `Heading 1`,
`Heading 2`, `Footnote Reference`, `Footnote Text`. Everything
else under editorial control is a custom (BuiltIn=False) style
and is untouched by this routine - custom-style discipline is
covered by `AuditNonPaletteStyleColors`.

**2.2 Wire `AuditNonPaletteStyleColors` into RUN_THE_TESTS
(MEDIUM) - DONE 2026-05-15.** Permanent custom-style
colour-discipline test. Return value 0 is the assertion.
Assigned to slot 44 (reused from `TestSlotAvailable` placeholder;
`values(44)` was already 0). Three sites in `aeBibleClass.cls`
updated: `GetPassFail` Case 44, `RunTest` `Debug.Print` label,
`BufAppend` label. Function called with default args
(`IncludeBuiltIn = False`) - custom + linked styles only, which
is the typical custom-style discipline signal.

**2.3 `CountUnapprovedVisibleStyles` test (MEDIUM) - DONE
2026-05-15.** Walks styles, counts those that are neither in
the `GetApprovedStyles` SSOT nor properly hidden (Priority=99 AND
QuickStyle=False AND UnhideWhenUsed=False; applied to all styles
regardless of BuiltIn). Combined with the hide-sweep this gives
the strong rule: "only approved styles visible to the editor /
translator." Assigned to slot 45 (reused from the obsolete
`CountFindNotEmphasisBlack` stub, which was removed). Function
lives inside `aeBibleClass.cls` as a `Private Function` per the
"test functions live in the class" rule recorded 2026-05-15.
Skips `wdStyleTypeTable` / `wdStyleTypeList`. Returns -1 on
internal error (will FAIL against expected 0). Prints violating
NameLocal list to Immediate on violations > 0.

**2.4 `AuditBookHyperlinkStyling` wired into RUN_THE_TESTS
(MEDIUM) - DONE 2026-05-15.** Permanent BookHyperlink discipline
test alongside `CountActiveHyperlinks` (slot 17). Assigned to
slot 46 (reused from the obsolete `CountFindNotEmphasisRed` -1
stub, removed). Per the class-encapsulation rule (§ 9 below /
`EDSG/12-module-vs-class.md`), the function body was migrated
from `basStyleInspector.bas` into `aeBibleClass.cls` as a
`Public Function`; a thin two-line delegate stub remains in
`basStyleInspector.bas` so `?AuditBookHyperlinkStyling` from the
Immediate window keeps working. This pair is the canonical
template for the still-pending slot-44 retro fix.

**2.5 `AuditThemeColorUsage` (LOWER PRIORITY).** Walks every
style, reports any with
`Font.TextColor.ObjectThemeColor <> wdThemeColorNone`. Catches
Office theme leaks. Informational; useful after the hide-sweep
to confirm coverage.

**2.6 `AuditDeliberateColourCompliance` (LOWER PRIORITY).** For
each named-deliberate-colour style (Hyperlink-family,
Footnote Reference, Verse marker, Chapter Verse marker, Words of
Jesus, EmphasisRed, BookHyperlink), verify `Font.Color` matches
the expected palette entry. Requires a style → palette-name
registry in `basBiblePalette`.

**Item 13 closes** when 2.1-2.4 are done and clean. 2.5 and 2.6
are graduate items.

**Status 2026-05-15:** 2.1, 2.2, 2.3, 2.4 all DONE this session.
Pending retro fix (slot 44 `AuditNonPaletteStyleColors` body
migration into the class) ALSO completed 2026-05-15 using the 2.4
body-in-class + delegate-stub template. `StyleTypeName` promoted
from `Private` to `Public` in `basStyleInspector` so the class
can call it (stateless lookup → fine to remain in the module per
the rule); `ColorDisplay` (used only by the migrated function)
moved into the class as `Private Function`. Item 13 fully
closes when operator-verification of slot 44 + slot 46 passes.

**Operator-verification snapshot 2026-05-15 (closes Item 13):**

- `HideUnapprovedBuiltInStyles` (first run): 24 newly hidden,
  92 already hidden, 251 skipped (locked). Newly hidden set
  includes `Hyperlink`, `FollowedHyperlink`, `Smart Hyperlink`,
  `SmartLink` (the gallery-leak vectors that motivated the
  BookHyperlink reframe), plus `Heading 3`-`Heading 9`,
  `TOC Heading`, `Caption`, `No Spacing`, `List Paragraph`,
  `Bibliography`, `Book Title`, `Subtitle`, `Strong`,
  `Balloon Text`, `Default Paragraph Font`, `Hashtag`,
  `Mention`, `Unresolved Mention`, `HTML Preformatted`,
  `HTML Typewriter`.
- `RUN_THE_TESTS 17` (`CountActiveHyperlinks`): PASS at 0.
- `RUN_THE_TESTS 44` (`AuditNonPaletteStyleColors`): PASS at 0.
- `RUN_THE_TESTS 45` (`CountUnapprovedVisibleStyles`): PASS at 0.
- `RUN_THE_TESTS 46` (`AuditBookHyperlinkStyling`): PASS at 0,
  15 BookHyperlink-styled runs checked.

**Built-ins diagnostic (`?AuditNonPaletteStyleColors(True)`):** the
include-built-ins one-off classified 176 styles total. Tier 1: 147,
Tier 2: 9, Theme: 16, Anomaly: 4. The 16 theme-coloured built-ins
are `Block Text`, `Caption`, `Heading 4`-`Heading 9`,
`Intense Emphasis`, `Intense Quote`, `Intense Reference`, `Quote`,
`Subtitle`, `Subtle Emphasis`, `Subtle Reference`, `TOC Heading` -
all carry Word's default Office theme colours. The 4 anomalies are
Microsoft-defined UI accent RGBs that ship with Word built-ins:
`Hashtag` (`#2B579A`), `Mention` (`#2B579A`),
`Placeholder Text` (`#666666`), `Unresolved Mention` (`#605E5C`).
All 20 of these built-ins are hidden by `HideUnapprovedBuiltInStyles`
(they appear in the 2.1 newly-hidden list above), so they do not
appear in the editor's Style gallery and do not affect editorial
discipline. The default `AuditNonPaletteStyleColors` (custom +
linked only) correctly reports 0 anomalies. Including built-ins is
informational only - useful as a one-off audit when investigating
a paste-from-elsewhere theme leak, not a reason to extend the
return-value contract.

Carried from 2026-05-12 item 13 (reframed 2026-05-14, advanced
again 2026-05-15 with the BookHyperlink addition).

### 3. Prescriptive-spec pass for `LineSpacingRule = Exactly` (MEDIUM) - CLOSED 2026-05-16

Two paragraph styles in bucket 1 still hold `LineSpacingRule =
Exactly` against the QA-checklist preference of `Single`:

- `CustomParaAfterH1` — `Exactly 10`
- `Footnote Text` — `Exactly 8`

**Status 2026-05-16:** `DumpStyleProperties` was run on 11 styles
(see `rpt/Styles/` dated 2026-05-16) confirming the outcome:

- `CustomParaAfterH1` — fixed to Single (`LineSpacingRule = 0`,
  `LineSpacing = 12`); taxonomy line in
  `basTEST_aeBibleConfig.bas` updated to match.
- `Footnote Text` — retained at `Exactly 8`. **Known exception:**
  promoting to Single caused space irregularity on the line
  carrying the Footnote Reference number. Flagged as potentially
  requiring adjustment for i18n due to font differences. Taxonomy
  entry annotated in-place.

These were intentionally retained when the `BaseStyle = ""` half
of the prior prescriptive-pass round was completed; the
`LineSpacingRule` change is a separate prescriptive decision per
style, not a batch.

The previously-listed `Heading 2`, `Psalms BOOK`, `Brief` cases
were resolved in the 2026-05-06 QA-alignment pass.

**Recommendation:** treat as two single-property decisions; for
each, a doc-side edit then one-line taxonomy update.

**Status check needed:** carried forward from 2026-05-14 arc
RECOVERED tag. Confirm relevance before scheduling: are these two
styles still on the prescriptive-pass radar, or has the
QA-checklist preference been revisited?

Original analysis: `rvw/Code_review 2026-04-25.md` § "Spec
promotion: descriptive vs prescriptive (decision)";
`rvw/Code_review 2026-05-06.md` § 2 (status updates);
`rvw/Code_review 2026-05-07.md` § 2.

### 4. Taxonomy audit — full-coverage final-state goal (LOW-MEDIUM, ASPIRATIONAL) - RECOVERED

State at last accounting (`Code_review 2026-05-07.md`):
**25 fully specified + 4 existence-verified + 3 not-yet-created
+ 5 tab-stops verified = 37 distinct style entries across 44
checks.**

The four existence-verified entries are character styles or
hard-to-spec paragraph styles awaiting a decision:

- `BookIntro` — paragraph; NOT FOUND in document. Decide:
  create the style (then promote with full spec) or remove
  from `approved`.
- `TheHeaders`, `TheFooters` — paragraph; structural. Promotion
  to bucket 1 needs a decision on what properties are
  meaningful for headers / footers (font / size mostly;
  alignment varies). Keep `wdColorAutomatic` per the two-tier
  colour discipline.
- `Footnote Reference` — character; bucket-1 promotion was the
  endpoint of the AuditOneStyle char-style extension. Verify
  this is now bucket 1 after the 2026-05-12 prerequisite
  closure.

Three not-yet-created (expected FAIL until each `Define*`
routine is run): `BodyTextContinuation`, `AppendixTitle`,
`AppendixBody`. Decide for each: create + promote, or remove
from `approved`.

**Recommendation:** a 10-minute state check via
`RUN_TAXONOMY_STYLES` will quantify how much was incidentally
closed by intervening arcs. Update the count then decide
whether the umbrella item warrants explicit attention or is
well-served by being a callout in `EDSG/01-styles.md`.

Note 2026-05-15: `BookHyperlink` was added to `approved` and to
the `AuditOneStyle` taxonomy as a fully-specified bucket-1
character style. That advances the count by one fully-specified
entry. Update the state-check number after the next run.

Original analysis: `rvw/Code_review 2026-04-25.md` § "Important —
taxonomy audit final-state goal" callout;
`rvw/Code_review 2026-05-07.md` § 3.

### 5. EDSG documentation refresh — VerseText-aware (LOW) - RECOVERED

Now that `VerseText` is the live verse-paragraph style (since
2026-05-01), EDSG needs opportunistic refresh on:

- **`EDSG/01-styles.md`** — `VerseText` at priority 31 in the
  priority snapshot; reframe `BodyText` as the residual
  non-verse paragraph style (front matter, chapter intros,
  chapter-end content). Per-round progress callouts have been
  kept current; the broader page narrative still references the
  pre-VerseText state.
- **`EDSG/06-i18n.md`** — note `VerseText` as the primary
  translation target.
- **`EDSG/02-editing-process.md`** — Stage 1 step references
  could mention `AuthorListItem*` as the canonical example for
  the `BaseStyle = ""` rule.
- **`EDSG/04-qa-workflow.md`** — "Current state" section dated
  2026-04-26 still mentions the priorities 38-41 reserved gap
  and 43-styles count; superseded by 2026-04-29 `SpeakerLabel`
  insertion and again by 2026-05-01 `VerseText` insertion.

Note 2026-05-15: `EDSG/01-styles.md` has been touched again this
session (Companion rule rewrite for BookHyperlink). The
VerseText-narrative refresh remains an outstanding item from
2026-05-07 and is still relevant.

**Recommendation:** opportunistic update next time these pages
are visited for substantive edits. Not blocking anything.

Original analysis: `rvw/Code_review 2026-05-07.md` § 5.

### 6. Finding 5 (ribbon nav) — umbrella OPEN (DEFERRED, WORD LIMITATION) - RECOVERED

Fix (A) resolved the primary caret-not-visible symptom. The
residual is the **idle-commit focus leak**: Word's customUI
`editBox` auto-commits on idle (~1 s) and returns focus to the
document body, so any subsequent Tab is a document Tab.

**Status:** **WORD LIMITATION, NO VBA-SIDE FIX AVAILABLE.**
KeyTips are the supported Office UX path for cross-control jumps
and bypass Tab entirely.

**Forward options (deferred):**

- **Auto-fire Go on valid `(book, chapter, verse)` triple** —
  code-side option to remove the final Tab → Go step. Tradeoff:
  nav fires before user expects it; would need a `bAutoFire`
  toggle.
- **VSTO/WPF ribbon rewrite** — only path to true ribbon-owned
  focus management. Major rewrite; deferred indefinitely.

Original analysis: `rvw/Code_review 2026-04-25.md` § "Finding 5 —
terminology correction"; `rvw/Code_review 2026-05-07.md` § 4.

### 7. SoftHyphenSweep_FootnotesOnly sister routine (DEFERRED)

No triggering need yet; build the sister routine only if a
`Footnote Text` font-change pass leaves stray soft hyphens in
footnote bodies.

Originated `rvw/Code_review 2026-05-08.md` § 3b.

### 8. SHA_ReplaceHard i18n consideration (FUTURE)

Revisit only if a non-English edition adopts soft hyphens as
semantic compound-break markers (German, Dutch, some
Slavic-language typesetting conventions occasionally do).

Originated `rvw/Code_review 2026-05-08.md` § 3c.

### 9. Architecture rule — class encapsulation + module/class as casual-coder safety boundary (RULE, 2026-05-15)

Established 2026-05-15 while wiring item 13 / 2.3 and 2.4.

**Rule.** Class-related code stays inside the class. If a class
needs behaviour from elsewhere, it calls into another class
(`aeAssertClass`, `aeLoggerClass`, `aeBibleCitationClass`,
`aeLongProcessClass`, `aeRibbonClass`, `aeUpdateCharStyleClass`) -
not into a module. Stateless specs / lookup tables / config SSOTs
may remain in modules.

**Why - architectural side.** The project already runs a class
spine. Putting test code and coherent stateful workflow inside
classes keeps the slot dispatch and its dependencies in one
readable file; module-side test bodies require hunting and
diverge from the SSOT principle the class itself embodies.

**Why - social side (the larger benefit).** Most VBA contributors
edit modules and avoid classes - classes signal "an invariant is
being maintained here." By concentrating stateful behaviour
inside classes, the file boundary itself becomes a permission
gate: a future i18n / editorial contributor who opens only `.bas`
files cannot break a class invariant they never saw. They edit
the palette table, the abbreviation list, the approved-styles
SSOT, the ribbon XML - and never need to know what a class *is*.
If they find themselves opening a `.cls` file, that's the signal
to stop and ask. The boundary doubles as a blast-radius limiter.

**Why `basBiblePalette` stays a module (not a class).** It is a
stateless lookup table. Converting it would dissolve the safety
signal (every palette-row edit would require opening a class
file) while buying no behavioural gain. Classes for stateful
actors and test code; modules for spec / data / lookup. The
discipline only works if the module/class split *means
something*, and the right meaning is "stateless data vs stateful
actor," not "everything important is a class."

**How applied this session.**

- 2.3: `CountUnapprovedVisibleStyles` placed as `Private Function`
  inside `aeBibleClass.cls` (not in `basStyleInspector`).
- 2.4: `AuditBookHyperlinkStyling` body moved into
  `aeBibleClass.cls`; thin delegate stub left in
  `basStyleInspector.bas` for the existing
  `?AuditBookHyperlinkStyling` Immediate-window usage.
- Retro fix completed 2026-05-15 (later same session):
  `AuditNonPaletteStyleColors` (slot 44) body migrated into
  `aeBibleClass.cls` using the 2.4 template. `StyleTypeName`
  promoted to `Public` in `basStyleInspector` (stateless lookup,
  stays in module per the rule). `ColorDisplay` moved into the
  class as `Private Function`.

**EDSG anchor.** [`EDSG/12-module-vs-class.md`](../EDSG/12-module-vs-class.md)
gives the contributor-facing version of this rule and should be
read by anyone planning an i18n / editorial code change before
opening a `.cls` file.

### 10. TestReport.txt — per-slot one-line descriptions (MEDIUM, 2026-05-16)

Approved 2026-05-16 as a follow-up to the unified First-hit
refactor. Goal: make each PASS/FAIL row in `rpt/TestReport.txt`
self-explanatory without diving into the class. Function names
like `Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed` and
`HasLeftAlignedParagraph(18, 931)` are cryptic six months out;
the description shifts the report from "test surface" to
"living test plan."

**Shape (chosen 2026-05-16):**

- Storage: new `Public Function GetTestDescription(num As Long)
  As String` inside `aeBibleClass.cls`, returning the description
  for each slot via `Select Case`. Single SSOT; no risk of
  duplicating across the three label sites (`GetPassFail`,
  `RunTest`, `OutputTestReport`).
- Emission: add a second indented line per test, only emitted
  when the description is non-empty. Parallel to the existing
  `>> First hit:` shape so the report has a consistent
  "subordinate lines under a PASS/FAIL header" pattern.
  Example:

  ```
  PASS              Copy ()       Test = 9      7             7             CountPeriodSpaceLeftParenthesis
                                                                          Detects ".  (" sequences left over from legacy footnote artifact.
  ```
- Width: do not widen the existing column block; keep the
  description as a wrapped subordinate line.
- Gradual rollout: scaffold `GetTestDescription` with empty
  strings; fill in one description per slot as each test gets
  touched, or in a single dedicated review pass.

**Pros:** report becomes self-documenting; writing 73 one-liners
forces an audit of each slot (would have surfaced the
`[obso]` stubs we removed this past session); pairs naturally
with a future "publish test plan" EDSG appendix.

**Cons / honest cost:** 73 sentences to write; risk of staleness
unless the description is kept near the slot dispatch (the SSOT
choice mitigates this); cannot be auto-derived from code in VBA
(no reflection).

**Plumbing scope:**

1. Add `Public Function GetTestDescription(num As Long) As String`
   skeleton in `aeBibleClass.cls` with `Select Case 1 To 73`
   returning `""` for each.
2. Extend `RunTest` post-Select block: if
   `GetTestDescription(num)` is non-empty, `Debug.Print , , ,
   "  " & GetTestDescription(num)`.
3. Extend `OutputTestReport` post-Select block similarly using
   `BufAppend`.
4. Fill in descriptions as a separate authoring pass.

Originated: 2026-05-16 follow-up to the unified First-hit
emission (this file § "operator-verification snapshot" trail).

### 11. Slot 5 retired + Slot 6 upgraded - empty-paragraph discipline tightened (2026-05-16)

Two related changes this session triggered by the slot-by-slot
description authoring in § 10.

**Slot 5 retired.** On review of slot 5
(`CountWhiteSpaceAndCarriageReturn`) it was confirmed to be a strict
subset of slot 3 (`CountSpaceFollowedByCarriageReturn`): both search
for `" ^13"`, slot 5 adds a `Font.Color = wdColorWhite` filter.
Slot 3 catches every occurrence slot 5 could catch; slot 3 also
populates `m_lastHint` while slot 5 did not. Slot 5 cannot signal
anything slot 3 misses. Function `CountWhiteSpaceAndCarriageReturn`
removed; slot 5 reassigned to `TestSlotAvailable` placeholder.

**Slot 6 upgraded.** `CountQuadrupleParagraphMarks` searched
`^13^13^13^13` (three consecutive empty paragraphs between content)
as the violation threshold for "author using CRLF as vertical
spacing." Two issues surfaced on review:

- The threshold was too loose - two consecutive empty paragraphs
  is already a violation in editorial discipline.
- The accepted exception (one empty paragraph followed by a page /
  column / section break, used as legitimate vertical spacing)
  cannot be expressed as a text-only `Find` pattern. A break can
  live in three different shapes that pure `^13`-counting cannot
  see:
  1. `Chr(12)` (page break) or `Chr(14)` (column break) inserted
     inline in a paragraph.
  2. `Paragraph.PageBreakBefore = True` on the following paragraph
     - no visible character at all.
  3. A section break, stored as the *type* of the terminating
     paragraph mark.

  Tightening to `^13^13^13` via Find would flag legitimate
  "empty + page break" runs as violations - the exact kind of
  false-positive noise just retired in slot 5.

**Implementation chosen.** Path (b) from the 2026-05-16 review
discussion: paragraph-walk replacement. New function
`CountConsecutiveEmptyParagraphsNotPrecedingBreak` walks
`Document.Paragraphs`, detects runs of length >= 2 where
`Len(p.Range.Text) = 1`, and at each run boundary checks three
exception predicates via a helper `HasBreakAtRunBoundary`:

1. Next content paragraph has `PageBreakBefore = True`.
2. Last empty paragraph in the run contains `Chr(12)` or
   `Chr(14)` inline.
3. Last empty paragraph terminates a section
   (`wdActiveEndSectionNumber` differs from next paragraph's).

If none of the three exceptions holds, the run counts. A run that
trails the document end is always counted (no following content
can justify trailing whitespace).

**`m_lastHint` shape.** First-violation hint records
`"Page <n>, line <m>: <k> consecutive empty paragraphs"` via
helper `FormatEmptyRunHint`, using
`Range.Information(wdActiveEndPageNumber)` and
`wdFirstCharacterLineNumber`. Document-end runs get a
`" (document end)"` suffix so the FAIL line says where to look
without re-running with diagnostics. Parallel to the run-of-the
-mill First-hit shape used elsewhere.

**Class-encapsulation alignment (§ 9).** Function lives inside
`aeBibleClass.cls` as `Private Function` along with its two
private helpers (`HasBreakAtRunBoundary`, `FormatEmptyRunHint`).
No module-side delegate stub needed - this is purely test
dispatch, not an Immediate-window helper. Slot 6 is the cleanest
example so far of "stateful slot logic stays in the class" since
all three exception predicates are paragraph-walk side effects.

**Status:** code changes complete 2026-05-16. Awaiting operator
verification: `RUN_THE_TESTS 6` should still return 0 on the
production document. If new violations surface (likely - the
threshold is now tighter), they are real editorial signal, not
test bug; address by removing the offending empty paragraphs or
documenting the run as legitimate spacing.

**Slot availability after this change.** Slot 5 is the sole free
slot. Slots 42, 51, 72 are skipped (heavy / conditional), not
available. Slots 44, 45, 46 are populated with custom-style
audits (Item 13 work). Net pool: 1 slot.

Originated: 2026-05-16 slot-by-slot description authoring of § 10
surfaced the slot 3 / slot 5 redundancy; slot 6 caveat surfaced in
the same review pass.

**Operator-verification follow-up 2026-05-16.**

- *Adjusted page-number fix.* First operator run hinted "Page 689"
  but the footer at that location read 682-683. Cause:
  `FormatEmptyRunHint` was using
  `Range.Information(wdActiveEndPageNumber)`, which returns the
  sequential internal index ignoring Roman-numeral front matter
  and numbering restarts. Changed to
  `wdActiveEndAdjustedPageNumber`, matching the convention already
  established in `HasLeftAlignedParagraph` (slot 72) which solved
  the same problem for the same reason. Re-run reported
  "Page 681" - matches what the editor sees in the footer.
- *Styled-paragraph mechanics learned.* The actual offending stack
  on page 681 was `[empty][empty][empty][BodyTextTopLineCPBB
  without a tab][content]`. A paragraph styled
  `BodyTextTopLineCPBB` (centered, page-break-before) with no tab
  is structurally indistinguishable from an empty paragraph by the
  `Len(p.Range.Text) = 1` test, so it was being counted as part of
  the empty run. The break-exception predicate
  (`nextContent.PageBreakBefore = True`) attaches to the
  *paragraph carrying* the page-break-before, not to the empties
  preceding it - so this case was correctly identified as a
  4-empty run terminating at content, no exception triggered.
  Editorial fix (adding the tab character to the styled paragraph
  to make it structurally non-empty) restored PASS. The rule held;
  the document was wrong. Good signal that the path-(b)
  implementation discriminates the empty-vs-content boundary at
  the right level.
- *Slot 6 description authored 2026-05-16.* Description filled in
  via § 10 `GetTestDescription` case branch; emission confirmed in
  the post-fix PASS report ("Detects runs of length >= 2 with
  three exceptions ...").

## Pointer back to the closed arc

Full dated history of the work that produced this carry-forward
state is in [`rvw/Code_review 2026-05-14.md`](Code_review%202026-05-14.md).
That file (and the arcs it points back to) covers:

- The BookHyperlink design, implementation, and verification.
- The 2026-05-15 `AuditHyperlinkStyling` extension that surfaced
  the size-11 bug driving the BookHyperlink refactor.
- The `Word.Field` casing normaliser additions.

Older arcs of historical relevance:

- [`Code_review 2026-05-12.md`](Code_review%202026-05-12.md) —
  Palette infrastructure (`basBiblePalette`), no-clickable-
  hyperlinks rule v1 (built-in Hyperlink style locked to palette),
  item 11 cleanup family, item 13 Pass 1 (two-tier colour
  discipline). 9 items closed this arc.
- [`Code_review 2026-05-11.md`](Code_review%202026-05-11.md) —
  AuditOneStyle char-style extension, ribbon alias bug fix,
  aeRibbon production export gateway design.
- [`Code_review 2026-05-08.md`](Code_review%202026-05-08.md) —
  BodyTextIndent + BookIntro removals, soft-hyphen sweep
  end-to-end build.
- [`Code_review 2026-05-07.md`](Code_review%202026-05-07.md) —
  AuditVerseMarkerStructure CVM extension; **source of items
  3, 4, 5, 6 above (RECOVERED tag)** which fell off the
  carry-forward chain between 2026-05-07 and 2026-05-08.
- [`Code_review 2026-05-06.md`](Code_review%202026-05-06.md) —
  VerseText migration (`BodyText` → `VerseText` paragraph-style
  switch).
- [`Code_review 2026-04-30.md`](Code_review%202026-04-30.md) —
  Reference rename (Solomon → Song of Songs), AuditOneStyle
  char-style audit kicker, body-content number-prefix decision.
