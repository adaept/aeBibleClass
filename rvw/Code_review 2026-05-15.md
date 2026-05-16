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
One pending retro fix carries over (slot 44 `AuditNonPaletteStyleColors`
body migration into the class), but Item 13's gating sub-items are
complete; operator-verification of slot 46 closes it.

Carried from 2026-05-12 item 13 (reframed 2026-05-14, advanced
again 2026-05-15 with the BookHyperlink addition).

### 3. Prescriptive-spec pass for `LineSpacingRule = Exactly` (MEDIUM) - RECOVERED

Two paragraph styles in bucket 1 still hold `LineSpacingRule =
Exactly` against the QA-checklist preference of `Single`:

- `CustomParaAfterH1` — `Exactly 10`
- `Footnote Text` — `Exactly 8`

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
- Retro fix pending: `AuditNonPaletteStyleColors` (slot 44) body
  still lives in `basStyleInspector`. Migrate using the same
  body-in-class + stub-in-module template established by 2.4.

**EDSG anchor.** [`EDSG/12-module-vs-class.md`](../EDSG/12-module-vs-class.md)
gives the contributor-facing version of this rule and should be
read by anyone planning an i18n / editorial code change before
opening a `.cls` file.

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
