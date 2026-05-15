# Code review - 2026-05-14 carry-forward

This file opens a fresh review arc on 2026-05-14. The previous arc
[`rvw/Code_review 2026-05-12.md`](Code_review%202026-05-12.md) is now
**closed for new work**; that file remains the authoritative dated
history for everything between 2026-05-12 and 2026-05-14, including:

- **Palette infrastructure** — `basBiblePalette.bas` added, 12 named
  colours, late-bound, theme-extensible API. `Module1.HexToRGB`,
  `basTEST_aeBibleTools.GetColorNameFromHex`,
  `basTEST_aeBibleTools.ListAndCountFontColors` rewired to delegate.
  `wdColorAutomatic` and `wdUndefined` handled as distinct sentinels;
  no phantom-colour buckets in the histogram.
- **Footnote Reference colour correction.** Live style confirmed Blue
  (`#0000FF`, 296 references). `Module1.EnsureFootnoteReferenceStyleColor`
  corrected from Purple to Blue; palette `Usage` field corrected.
- **No-clickable-hyperlinks rule** codified. `LockHyperlinksToPalette`
  rewritten: unlinks active Hyperlinks (step 0), locks Hyperlink +
  FollowedHyperlink styles to palette `DarkBlue` (`#000080`), forces
  colour + underline on every Hyperlink-styled run across all stories.
  Test 17 redefined `CountActiveHyperlinks`; expected 0 across all
  StoryRanges; current state PASS at 0/0. 14 visible-as-link runs
  remain (12 concordance URL stubs + 2 newly-unlinked), all
  rule-compliant.
- **Item 11 cleanup family.** 7 `#C00000` Jesus quotations migrated
  to DarkRed; 5 proper-noun false-positive Hyperlink-styled runs
  stripped manually; footnote 218 duplicate FR-styled paragraph mark
  restyled; footnote-story URL deleted.
- **Item 13 Pass 1.** Two-tier colour discipline established:
  `wdColorAutomatic` (default-text) and palette-registered colours
  are the only allowed Font.Color values on custom styles. New
  `AuditNonPaletteStyleColors` 5-bucket classifier (Tier 1 / Tier 2 /
  Theme / Anomaly / Skipped). Orphan custom `Error` style deleted.
  Custom-style anomaly count now 0.
- **EDSG additions.** Section "State-aware styles: print-locking"
  (Hyperlink lock pattern), section "Companion rule: no clickable
  hyperlinks anywhere" (rule + audits), section "Colour discipline:
  two tiers, no third" (full tier rule + audit map).
- **Test 17 history this arc.** `CountRedFootnoteReferences` (dead)
  → `CountFootnoteHyperlinks` (footnote-only) → `CountActiveHyperlinks`
  (all stories). Each handoff documented in the closed file.
- **Items closed this arc** (numbering from the closed file): 2, 3,
  4, 5 (WONTFIX), 6, 7 (WONTFIX), 10, 11, 12. Plus item 13 Pass 1.

Items below are the **open** carry-forward set, ordered roughly by
unlock-to-effort ratio - work that removes blockers for multiple
downstream items, or that closes a category of risk, at the top.

Each item is marked with a status tag:

- **OPEN** — actively pending, all known prerequisites met.
- **PARTIAL** — partially complete; specific remaining work listed.
- **DEFERRED** — not started, waiting on a specific trigger.
- **FUTURE** — speculative; revisit only when conditions warrant.
- **RECOVERED** — surfaced from a prior arc where it was dropped
  off the carry-forward chain. Examination needed: still relevant,
  still scoped correctly?

## Open carry-forward (priority order)

### 1. Run aeRibbon Gates G1-G8 and ship v1.0.0 (HIGH) - OPEN

The production export gateway is in place; nothing has been built
or gated yet. This is the **next active release-track item** and
gates the hand-off to the author for comments-only review.

**Why high:** every other ribbon-side improvement (signing,
auto-docx-from-docm, ribbon UX iteration) sits behind a first
successful gated build. Also the highest-leverage validation of the
trim script: any false drop will surface in G6 (compile) or G8
(navigation smoke).

**Action:**

1. Build `aeRibbon/template/aeRibbon.dotm` per `aeRibbon/BUILD.md`.
   - Inject `aeRibbon/template/customUI14.xml` via
     `wsl python3 py/inject_ribbon.py`.
   - Import the 5 files from `aeRibbon/src/` into the template VBA
     project.
   - Set `RIBBON_VERSION` constant + custom property
     `aeRibbonVersion` to match `aeRibbon/VERSION` (`1.0.0+bc71416`).
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

### 2. Item 13 remaining work — built-in hide-sweep + test wiring (MEDIUM) - PARTIAL

Pass 1 of item 13 closed 2026-05-14 (`AuditNonPaletteStyleColors`
added, custom-style anomaly count brought to 0 by deleting the
orphan `Error` style). Remaining work falls into four pieces, in
priority order:

**2.1 Hide-sweep for Word built-in noise (MEDIUM).** 122 built-in
styles are currently skipped by the audit because they're not under
editorial control. Most carry theme colours or hand-set RGB
defaults (Heading 4-9, Caption, Quote, Hashtag, Mention,
Placeholder Text, Unresolved Mention, etc.). They aren't applied
to any production text but they show up in the Styles pane and
gallery, polluting the editor's options.

New routine `HideUnapprovedBuiltInStyles` in `basStyleInspector`.
For every `BuiltIn = True` style NOT in the `approved` array, set:

```vba
.Priority = 99
.QuickStyle = False
.UnhideWhenUsed = False
```

The three-property pattern (not just Priority) matters because
`UnhideWhenUsed = True` re-surfaces a style in the gallery the
moment any run touches it - including paste operations.

**2.2 Wire `AuditNonPaletteStyleColors` into RUN_THE_TESTS
(MEDIUM).** Permanent custom-style colour-discipline test. Return
value 0 is the assertion. Slot number to be assigned by operator.

**2.3 `CountUnapprovedVisibleStyles` test (MEDIUM).** Walks styles,
counts those that are neither in `approved` nor properly hidden
(BuiltIn + Priority=99 + QuickStyle=False + UnhideWhenUsed=False).
Combined with the hide-sweep this gives the strong rule: "only
approved styles visible to the editor / translator." Slot TBD.

**2.4 `AuditThemeColorUsage` (LOWER PRIORITY).** Walks every style,
reports any with `Font.TextColor.ObjectThemeColor <> wdThemeColorNone`.
After the hide-sweep, theme-coloured built-ins are hidden but the
property values remain theme-encoded. Informational audit; useful
for surfacing any new theme-colour leaks from imports/paste.

**2.5 `AuditDeliberateColourCompliance` (LOWER PRIORITY).** For
each named-deliberate-colour style (Hyperlink, Footnote Reference,
Verse marker, Chapter Verse marker, Words of Jesus, EmphasisRed),
verify `Font.Color` matches the expected palette entry. Catches
drift (e.g., Hyperlink getting reset to its default by a template
operation). Requires a style → palette-name registry. Per operator
decision 2026-05-14, registry lives in `basBiblePalette`.

**Item 13 closes** when 2.1, 2.2, 2.3 are done and clean. 2.4 and
2.5 are graduate items - useful but not blocking item 13.

Carried from 2026-05-12 item 13 (reframed 2026-05-14).

### 3. Prescriptive-spec pass for `LineSpacingRule = Exactly` (MEDIUM) - RECOVERED

Two paragraph styles in bucket 1 still hold `LineSpacingRule =
Exactly` against the QA-checklist preference of `Single`:

- `CustomParaAfterH1` - `Exactly 10`
- `Footnote Text` - `Exactly 8`

These were intentionally retained when the `BaseStyle = ""` half
of the prior prescriptive-pass round was completed; the
`LineSpacingRule` change is a separate prescriptive decision per
style, not a batch.

The previously-listed `Heading 2`, `Psalms BOOK`, `Brief` cases
were resolved in the 2026-05-06 QA-alignment pass.

**Recommendation:** treat as two single-property decisions; for
each, a doc-side edit then one-line taxonomy update.

**Why RECOVERED:** carried forward from `Code_review 2026-05-07.md`
item 2 but dropped from `Code_review 2026-05-08.md` onward.
Confirm relevance before scheduling: are these two styles still
on the prescriptive-pass radar, or has the QA-checklist preference
been revisited?

Original analysis: `rvw/Code_review 2026-04-25.md` § "Spec
promotion: descriptive vs prescriptive (decision)";
`rvw/Code_review 2026-05-06.md` § 2 (status updates);
`rvw/Code_review 2026-05-07.md` § 2.

### 4. Taxonomy audit — full-coverage final-state goal (LOW-MEDIUM, ASPIRATIONAL) - RECOVERED

State at last accounting (`Code_review 2026-05-07.md`): **25 fully
specified + 4 existence-verified + 3 not-yet-created + 5 tab-stops
verified = 37 distinct style entries across 44 checks.**

The four existence-verified entries are character styles or
hard-to-spec paragraph styles awaiting a decision:

- `BookIntro` — paragraph; NOT FOUND in document. Decide: create
  the style (then promote with full spec) or remove from
  `approved`.
- `TheHeaders`, `TheFooters` — paragraph; structural. Promotion
  to bucket 1 needs a decision on what properties are even
  meaningful for headers / footers (font / size mostly; alignment
  varies). (Note 2026-05-14: these two were the subject of the
  rejected "convert to literals" sub-task in item 13. They keep
  `wdColorAutomatic` per the two-tier discipline; their bucket-1
  promotion needs decisions on the non-colour properties.)
- `Footnote Reference` — character; bucket-1 promotion was the
  endpoint of the AuditOneStyle char-style extension. Verify if
  this is now bucket 1 after the 2026-05-12 prerequisite closure.

Three not-yet-created (expected FAIL until each `Define*` routine
is run): `BodyTextContinuation`, `AppendixTitle`, `AppendixBody`.
Decide for each: create + promote, or remove from `approved`.

**Why RECOVERED:** carried forward from `Code_review 2026-05-07.md`
item 3 but dropped from `Code_review 2026-05-08.md` onward.
Bucket-2 → bucket-1 promotions did happen between then and now
(8 character styles promoted, `Selah` / `EmphasisBlack` etc.
upgraded), but the umbrella goal — every approved style mapped
with real specs — has not had an explicit progress check since.

Recommendation: a 10-minute state check via `RUN_TAXONOMY_STYLES`
will quantify how much was incidentally closed by the intervening
arcs. Update the count then decide whether the umbrella item
warrants explicit attention or is well-served by being a callout
in `EDSG/01-styles.md`.

Original analysis: `rvw/Code_review 2026-04-25.md` § "Important —
taxonomy audit final-state goal" callout;
`rvw/Code_review 2026-05-07.md` § 3.

### 5. EDSG documentation refresh — VerseText-aware (LOW) - RECOVERED

Now that `VerseText` is the live verse-paragraph style (since
2026-05-01), EDSG needs opportunistic refresh on:

- **`EDSG/01-styles.md`** — `VerseText` at priority 31 in the
  priority snapshot; reframe `BodyText` as the residual non-verse
  paragraph style (front matter, chapter intros, chapter-end
  content). Per-round progress callouts have been kept current; the
  broader page narrative still references the pre-VerseText state.
- **`EDSG/06-i18n.md`** — note `VerseText` as the primary translation
  target.
- **`EDSG/02-editing-process.md`** — Stage 1 step references could
  mention `AuthorListItem*` as the canonical example for the
  `BaseStyle = ""` rule.
- **`EDSG/04-qa-workflow.md`** — "Current state" section dated
  2026-04-26 still mentions the priorities 38-41 reserved gap and
  43-styles count; superseded by 2026-04-29 `SpeakerLabel`
  insertion and again by 2026-05-01 `VerseText` insertion.

**Why RECOVERED:** carried forward from `Code_review 2026-05-07.md`
item 5 but dropped from `Code_review 2026-05-08.md` onward.
`EDSG/01-styles.md` has been touched since with three additions
(state-aware styles, no-clickable-hyperlinks, two-tier colour
discipline) but the VerseText-narrative integration was not
addressed.

**Recommendation:** opportunistic update next time these pages are
visited for substantive edits. Not blocking anything.

Original analysis: `rvw/Code_review 2026-05-07.md` § 5.

### 6. Finding 5 (ribbon nav) — umbrella OPEN (DEFERRED, WORD LIMITATION) - RECOVERED

Fix (A) resolved the primary caret-not-visible symptom. The
residual is the **idle-commit focus leak**: Word's customUI
`editBox` auto-commits on idle (~1 s) and returns focus to the
document body, so any subsequent Tab is a document Tab.

**Status:** **WORD LIMITATION, NO VBA-SIDE FIX AVAILABLE.** KeyTips
are the supported Office UX path for cross-control jumps and
bypass Tab entirely.

**Forward options (deferred):**

- **Auto-fire Go on valid `(book, chapter, verse)` triple** —
  code-side option to remove the final Tab → Go step. Tradeoff: nav
  fires before user expects it; would need a `bAutoFire` toggle.
- **VSTO/WPF ribbon rewrite** — only path to true ribbon-owned
  focus management. Major rewrite; deferred indefinitely.

**Why RECOVERED:** carried forward from `Code_review 2026-05-07.md`
item 4 but dropped from `Code_review 2026-05-08.md` onward. The
ribbon work-track since then has focused on the production-export
gateway (which is item 1 above), leaving Finding 5 unaddressed in
the carry-forward chain. The deferred-indefinitely framing still
holds.

Original analysis: `rvw/Code_review 2026-04-25.md` § "Finding 5 —
terminology correction"; `rvw/Code_review 2026-05-07.md` § 4.

### 7. SoftHyphenSweep_FootnotesOnly sister routine (DEFERRED)

No triggering need yet; build the sister routine only if a
`Footnote Text` font-change pass leaves stray soft hyphens in
footnote bodies.

Carried from 2026-05-12 item 8; originated `rvw/Code_review
2026-05-08.md` § 3b.

### 8. SHA_ReplaceHard i18n consideration (FUTURE)

Revisit only if a non-English edition adopts soft hyphens as
semantic compound-break markers (German, Dutch, some Slavic-
language typesetting conventions occasionally do).

Carried from 2026-05-12 item 9; originated `rvw/Code_review
2026-05-08.md` § 3c.

## Pointer back to the closed arc

Full dated history of the work that produced this carry-forward
state is in [`rvw/Code_review 2026-05-12.md`](Code_review%202026-05-12.md).
That file (and the arcs it points back to) covers:

- The complete palette infrastructure design and rollout.
- The no-clickable-hyperlinks rule design, audit, and verification.
- Item 11's three-stage classification + cleanup
  (`#7F9698`/wdUndefined diagnosis, `#C00000` migration, 5-proper-
  noun strip, footnote 218 paragraph-mark restyle).
- Test 17's three-step evolution
  (CountRedFootnoteReferences → CountFootnoteHyperlinks →
  CountActiveHyperlinks).
- Item 13 Pass 1's reframe from "convert to literals" to
  two-tier colour discipline + audits.

Older arcs of historical relevance:

- [`Code_review 2026-05-11.md`](Code_review%202026-05-11.md) —
  AuditOneStyle char-style extension, ribbon alias bug fix,
  aeRibbon production export gateway design.
- [`Code_review 2026-05-08.md`](Code_review%202026-05-08.md) —
  BodyTextIndent + BookIntro removals, soft-hyphen sweep
  end-to-end build.
- [`Code_review 2026-05-07.md`](Code_review%202026-05-07.md) —
  AuditVerseMarkerStructure CVM extension, taxonomy parameteriz-
  ation. **Items 2, 3, 4, 5 of that arc are surfaced as
  items 3, 4, 6, 5 (respectively) in this review under the
  RECOVERED tag**, having fallen off the carry-forward chain
  between 2026-05-07 and 2026-05-08.
- [`Code_review 2026-05-06.md`](Code_review%202026-05-06.md) —
  VerseText migration (`BodyText` → `VerseText` paragraph-style
  switch); 12 closures.
- [`Code_review 2026-04-30.md`](Code_review%202026-04-30.md) —
  Reference rename (Solomon → Song of Songs), AuditOneStyle
  char-style audit kicker, body-content number-prefix decision.
