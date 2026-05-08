# Code review - 2026-05-07 carry-forward

This file opens a fresh review arc on 2026-05-07. The previous arc
[`rvw/Code_review 2026-05-06.md`](Code_review%202026-05-06.md) is now
**closed for new work**; that file remains the authoritative dated
history for everything between 2026-05-06 and 2026-05-07, including
the `AuditVerseMarkerStructure` CVM extension, four taxonomy
QA-alignment rounds (BaseStyle = "" enforcement plus bucket-1
promotions spanning Psalm, front-matter, TOC, and approved-list
styles), the `ContentsRef` tab-stop coverage, the `Footnote Reference`
base-style rebase to `Default Paragraph Font`, and the documented rule
that every approved character style must be based on `Default
Paragraph Font`.

Items below are the **open** carry-forward set. Each entry links back
to the section in the prior review where the full rationale lives.

## Closed in the prior arc (recorded here for completeness)

- **`AuditVerseMarkerStructure` CVM coverage extension.** Added
  per-chapter `CountChapterVerseMarkers` invariant; caught 2 real
  defects (Genesis 8 missing CVM, Job 37 stray CVM); user repaired;
  final SUMMARY 31102 / 31102 verses, 0 structural issues. See
  `rvw/Code_review 2026-05-06.md` § 7.
- **Taxonomy resync after QA-alignment partial pass.** Five paragraph
  styles adjusted in the document toward QA preferences; three
  taxonomy lines updated to match (`Heading 2`, `Brief`,
  `Psalms BOOK`). Result: 24 PASS / 4 FAIL across 28 checks. § 9.
- **BaseStyle = "" prescriptive pass + Psalm bucket-1 promotions.**
  `AuditOneStyle` extended with optional `sExpBaseStyle` parameter
  (sentinel `"<skip>"`); six paragraph styles now enforce
  `BaseStyle = ""` (`CustomParaAfterH1`, `Brief`, `Footnote Text`,
  `Psalms BOOK`, `PsalmAcrostic`, `PsalmSuperscription`).
  `PsalmAcrostic` and `PsalmSuperscription` promoted from bucket 2 to
  bucket 1. Result: 26 PASS / 4 FAIL across 30 checks. § 10.
- **Front-matter & TOC bucket-1 promotion.** 12 new bucket-1 entries
  for front-matter / TOC / index styles (priorities 4-17); `Title`
  promoted from bucket 2; `ContentsRef` gained `BaseStyle = ""`.
  `ContentsRef` tab stop added (4 -> 5 styles in tab-stop block).
  Result: 39 PASS / 4 FAIL across 43 checks. §§ 11-12.
- **Character-style basing rule documented.** Every approved
  character style must be based on `Default Paragraph Font` (not
  `Normal`, not another paragraph style). Rationale: overlay
  behavior - the run must adopt the host paragraph's font. § 13.
- **`Footnote Reference` base-style fix (doc-side).** Was incorrectly
  based on `"Normal text"`; user rebased to `"Default Paragraph
  Font"`; re-dump confirmed. Audit coverage still gated on item 1.
  § 14.
- **Round-4 BaseStyle = "" + `CenterSubText` promotion.** Eight more
  paragraph styles enforce `BaseStyle = ""` (`AuthorListItem`,
  `AuthorListItemBody`, `AuthorListItemTab`, `AuthorBookRefHeader`,
  `AuthorBookRef`, `Heading 1`, `DatAuthRef`, `Heading 2`).
  `CenterSubText` promoted to bucket 1 with full descriptive spec.
  Final-state at arc close: `RUN_TAXONOMY_STYLES` -> **40 PASS / 4
  FAIL across 44 checks** (pending re-import verification on test
  `.docm`; production `.docm` import still pending). § 15.

## Open carry-forward items

### 1. `AuditOneStyle` - extend for character-style properties

Currently `AuditOneStyle` checks only paragraph-style properties
(font name/size, alignment, indent, line-spacing, space before/after,
bold, italic, BaseStyle). The three approved character styles
(`Chapter Verse marker`, `Verse marker`, `Footnote Reference`) cannot
be fully audited - they remain in bucket 2 (existence-verified only)
because the audit cannot meaningfully assert character-style-relevant
properties.

**Required for:** `Footnote Reference`, `Chapter Verse marker`,
`Verse marker` to graduate from bucket 2 to bucket 1.

**Scope:** add 3-4 character-style-relevant property checks
(`bExpBold`, `bExpItalic`, `lExpColor`) and a character-flavored
`sExpBaseStyle` (expecting `"Default Paragraph Font"` per item 13's
rule from the prior arc). Or split into a sibling `AuditOneCharStyle`
with character-style-relevant checks only - same pattern as
`AuditStyleTabs`.

**Additional dependent decisions:**

- **BaseStyle locale literal.** English Word reports `"Default
  Paragraph Font"`; localized builds may differ. Spec the literal
  for English first; a locale-tolerant form is a future concern.
- **Color check.** Three character-style colors are descriptive
  baselines as of 2026-05-06: `Chapter Verse marker` ~ orange
  (`42495`), `Verse marker` ~ medium green (`7915600`),
  `Footnote Reference` ~ blue (`16711680`). User confirmation of
  intent on each is a separate prerequisite (see item 8 below).

Original analysis: `rvw/Code_review 2026-04-25.md` § "Footnote
Reference - deferred to bucket 2 (Q2 decision)" plus
`rvw/Code_review 2026-05-06.md` §§ 13-15.

### 2. Prescriptive-spec pass for `LineSpacingRule = Exactly` (residual)

Two paragraph styles in bucket 1 still hold `LineSpacingRule =
Exactly` against QA-checklist preference of `Single`:

- `CustomParaAfterH1` - `Exactly 10`
- `Footnote Text` - `Exactly 8`

These were intentionally retained when the BaseStyle = "" half of the
prior round was completed; the `LineSpacingRule` change is a separate
prescriptive decision per style, not a batch.

The previously-listed `Heading 2`, `Psalms BOOK`, `Brief` cases were
resolved in the prior arc's QA-alignment pass.

**Recommendation:** treat as two single-property decisions; doc-side
edit then one-line taxonomy update each.

Original analysis: `rvw/Code_review 2026-04-25.md` § "Spec promotion:
descriptive vs prescriptive (decision)" plus
`rvw/Code_review 2026-05-06.md` § 2 (status updates).

### 3. Taxonomy audit - full-coverage final-state goal

State at this arc open: **25 fully specified + 4 existence-verified +
3 not-yet-created + 5 tab-stops verified = 37 distinct style entries
across 44 checks.** All four existence-verified entries are character
styles or hard-to-spec paragraph styles awaiting a decision:

- `BookIntro` - paragraph; NOT FOUND in document. Either create the
  style (then promote with full spec) or remove from `approved`.
- `TheHeaders`, `TheFooters` - paragraph; structural. Promotion to
  bucket 1 needs a decision on what properties are even meaningful
  for headers / footers (font / size mostly; alignment varies).
- `Footnote Reference` - character; promotion blocked on item 1.

Three not-yet-created (expected FAIL until each `Define*` routine is
run): `BodyTextContinuation`, `AppendixTitle`, `AppendixBody`. These
are the four NOT-FOUND FAILs in the steady-state `RUN_TAXONOMY_STYLES`
output.

Original analysis: `rvw/Code_review 2026-04-25.md` § "Important -
taxonomy audit final-state goal" callout, plus
`EDSG/01-styles.md` callout (kept current via per-round bullets).

### 4. Finding 5 (ribbon nav) - umbrella OPEN

Fix (A) resolved the primary caret-not-visible symptom. The residual
is the **idle-commit focus leak**: Word's customUI `editBox`
auto-commits on idle (~1 s) and returns focus to the document body,
so any subsequent Tab is a document Tab.

**Status:** **WORD LIMITATION, NO VBA-SIDE FIX AVAILABLE.** KeyTips
are the supported Office UX path for cross-control jumps and bypass
Tab entirely.

**Forward options (deferred):**

- **Auto-fire Go on valid `(book, chapter, verse)` triple** -
  code-side option to remove the final Tab -> Go step. Tradeoff: nav
  fires before user expects it; would need a `bAutoFire` toggle.
- **VSTO/WPF ribbon rewrite** - only path to true ribbon-owned focus
  management. Major rewrite; deferred indefinitely.

Original analysis: `rvw/Code_review 2026-04-25.md` § "Finding 5 -
terminology correction".

### 5. EDSG documentation refresh - VerseText-aware

Now that `VerseText` is the live verse-paragraph style, EDSG needs
opportunistic refresh on:

- **`EDSG/01-styles.md`** - `VerseText` at priority 31 in the
  priority snapshot; reframe `BodyText` as the residual non-verse
  paragraph style (front matter, chapter intros, chapter-end
  content). The 2026-05-06 progress callout has been kept current
  with per-round bullets - the broader page narrative still
  references the pre-VerseText state.
- **`EDSG/06-i18n.md`** - note `VerseText` as the primary translation
  target.
- **`EDSG/02-editing-process.md`** - Stage 1 step references could
  mention `AuthorListItem*` as the canonical example for the
  `BaseStyle = ""` rule.
- **`EDSG/04-qa-workflow.md`** - "Current state" section dated
  2026-04-26 still mentions the priorities 38-41 reserved gap and
  43-styles count; superseded by 2026-04-29 `SpeakerLabel` insertion
  and again by 2026-05-01 `VerseText` insertion.

**Recommendation:** opportunistic update next time these pages are
visited for substantive edits.

### 6. Body-content number prefixes - keep manual (deferred-decision record)

Decision standing from `rvw/Code_review 2026-04-30.md`: keep manual
text prefixes (`"1. "`, `"2. "`, ...) on `AuthorBookRef` paragraphs;
do **not** convert to `{ DOCVARIABLE }` fields. Carried forward only
as a deferred-decision record so the trigger-to-revisit isn't lost.

**Trigger to revisit:** active i18n rollout where a target locale
requires non-Arabic numerals in body content, substantially different
prefix punctuation that can't be handled by a one-pass reformatter,
or per-paragraph content substitution that today's manual prefixes
can't accommodate.

Full reasoning: `rvw/Code_review 2026-04-30.md` § "8. Body-content
number prefixes".

### 7. Session manifest

`sync/session_manifest.txt` now contains three 2026-05-06 session
blocks (taxonomy resync + BaseStyle = "" prescriptive pass +
front-matter promotion). The 2026-05-01 -> 2026-05-06 src/ edits
prior to those (VerseText migration, orphan-audit, char-style-audit
parameterization, AuditVerseMarkerStructure CVM extension) are
**not** yet itemized as manifest entries - they are referenced in the
prior review file but absent from the import checklist.

**Action:** consolidate into a single fresh-manifest entry covering
2026-05-01 -> 2026-05-07 edits, or split per src/ file modified.
Either way, the manifest must be the authoritative import list when
re-syncing both `.docm` files.

### 8. `Footnote Reference` color - intent confirmation

Color sanity-check from item 14 in the prior arc remains open.
Current dump shows `Font.Color = 16711680` which decodes as
`&H00FF0000`. Word stores color as BGR, so this value is
`wdColorBlue` (B=255, G=0, R=0). If footnote reference numbers are
expected to render in red (the common convention), this is a latent
bug. User to confirm intended color.

If the intent is blue: leave as-is and capture as the descriptive
baseline when item 1 lands.

If the intent is red: change in the document to `wdColorRed = 255`
(`&H000000FF`, BGR 0/0/255), re-dump, then capture as the
descriptive baseline.

The same confirmation is wanted for the other two character-style
colors (descriptive baselines noted in `rvw/Code_review 2026-05-06.md`
§ 15): `Chapter Verse marker` ~ orange (`42495`), `Verse marker` ~
medium green (`7915600`).

## Pointer back to the closed arc

Full dated history of the work that produced this carry-forward state
is in [`rvw/Code_review 2026-05-06.md`](Code_review%202026-05-06.md).
That file includes:

- The complete `AuditVerseMarkerStructure` CVM extension record,
  including the two real defects it caught and verified.
- Four taxonomy QA-alignment rounds with full per-style spec deltas.
- The character-style basing rule and its rationale.
- The `Footnote Reference` `Normal text` -> `Default Paragraph Font`
  rebase record.

Anything in this 2026-05-07 file should reference back to that arc
for the *why*; this file holds only the **what is still open**.

## BodyTextIndent removed; bucket-1 promotion candidates

`BodyTextIndent` was deleted from the active document after a usage
audit (see `ListBodyTextIndentUsage` in `basAuditDocument.bas`)
returned zero paragraphs carrying the style. Taxonomy updated to
match:

- `basTEST_aeBibleConfig.bas` `PromoteApprovedStyles` array - removed
  `"BodyTextIndent"`.
- `basTEST_aeBibleConfig.bas` `RUN_TAXONOMY_STYLES` - removed the
  `AuditOneStyle "BodyTextIndent" ...` line; bucket-1 header comment
  updated from "16 fully specified" to "15 fully specified" and the
  style name dropped from the inline bucket list.
- `rpt/StyleTaxonomyAudit.txt` will regenerate on the next
  `RUN_TAXONOMY_STYLES`; the stale `PASS BodyTextIndent` line
  disappears at that point.

Left untouched (tooling, no taxonomy effect):

- `basFixDocxRoutines.DefineBodyTextIndentStyle` - still callable if
  a future document needs the style re-created. Mark for deletion in
  a follow-up if the decision is permanent.
- `basVerseStructureAudit.bas` block-comment reference at line 824
  (notes `BodyTextIndent` as a Phase-2 conversion candidate). The
  comment is historical and stays as-is per the
  "review docs are progressive history" rule applied to dated audit
  notes.

### Styles dumped this session (specs captured, bucket-1 promotion candidates)

Properties for the following styles have been written via
`DumpStyleProperties` (output under `rpt/Styles/`). They currently
live in `PromoteApprovedStyles` (priority list) but not in
`RUN_TAXONOMY_STYLES` bucket 1 (fully-specified). With specs now
captured they are eligible for bucket-1 promotion in a follow-up
pass:

- Footnote Text *(already in bucket 1; re-dumped for confirmation)*
- Psalms BOOK *(already in bucket 1; re-dumped for confirmation)*
- PsalmSuperscription *(already in bucket 1; re-dumped for confirmation)*
- Selah
- PsalmAcrostic *(already in bucket 1; re-dumped for confirmation)*
- SpeakerLabel
- EmphasisBlack
- EmphasisRed
- Words of Jesus
- AuthorSectionHead
- AuthorQuote
- AuthorBookSections

Promotion is held until each `AuditOneStyle ...` line can be written
against a known-good descriptive spec and verified PASS.

### Post-edit RUN_TAXONOMY_STYLES result

```
RUN_TAXONOMY_STYLES: 39 PASS  4 FAIL  -> rpt\StyleTaxonomyAudit.txt
```

Breakdown (43 total checks):

- Bucket 1 (fully specified): 31 PASS / 0 FAIL
- Bucket 2 (existence verified): 3 PASS / 1 FAIL
  - FAIL `BookIntro` - NOT FOUND in document (carried over).
- Bucket 3 (not yet created): 0 PASS / 3 FAIL (expected -
  `BodyTextContinuation`, `AppendixTitle`, `AppendixBody`).
- Tab stops: 5 PASS / 0 FAIL.

Delta from prior baseline `38 PASS / 4 FAIL across 42 checks`
(commit `d513542`): -1 PASS from removing `BodyTextIndent`, then
+2 PASS / +1 check observed from intervening promotions not yet
reflected in the bucket-1 header comment. Header comment in
`basTEST_aeBibleConfig.bas` realigned to current reality
(31 bucket-1 / 4 existence / 3 not-yet-created / 5 tab stops =
43 checks).

## BookIntro removed (obsolete)

`BookIntro` is obsolete - the `Introduction` style is used for that
role. Taxonomy updated:

- `basTEST_aeBibleConfig.bas` `PromoteApprovedStyles` array - removed
  `"BookIntro"`.
- `basTEST_aeBibleConfig.bas` `RUN_TAXONOMY_STYLES` bucket 2 -
  removed the `AuditOneStyle "BookIntro", ...` line; header comment
  realigned (37 styles / 5 tab-stop specs / 42 checks; bucket 2 now
  3 entries: `TheHeaders`, `TheFooters`, `Footnote Reference`).

Outstanding decision: the tooling routines `DefineBookIntroStyle`
and `ApplyBookIntroAfterDatAuthRef` in `basFixDocxRoutines.bas`
(approx 180 lines, 1100-1281) are now dead code. Deletion proposed
but pending explicit confirmation due to blast radius.

## DefineBodyTextIndentStyle removed

The orphaned `DefineBodyTextIndentStyle` Sub in
`basFixDocxRoutines.bas` (banner + body, 77 lines) was deleted.
`BodyTextIndent` removal is permanent; the recreate routine is no
longer needed.

## NEW TASK: Define colors used in the docx

Use of Word Themes / Theme Colors is **not allowed anywhere** in
this document. Every color reference must be expressed as an
explicit RGB / `wdColor*` constant captured as part of the
descriptive style baseline.

Scope:

- Enumerate every style and direct-formatting site that carries a
  non-default color (paragraph styles, character styles, run-level
  overrides, table/shading, ribbon-driven highlights).
- Confirm intended color for each site and capture as a long
  RGB / BGR literal in the style spec (no `wdColorAutomatic`, no
  theme-color tints, no `wdColorByPalette`).
- Lock in the three known character-style colors (open from
  `2026-05-06.md` § 15 / this file's "Footnote Reference color"
  section): `Footnote Reference`, `Chapter Verse marker`,
  `Verse marker`.
- Add a taxonomy check (extension of `AuditOneStyle` or a sibling
  routine) that fails any style whose color resolves through a
  theme rather than an explicit literal.

This subsumes the open "Footnote Reference color" confirmation
above; resolve there as a sub-step of this task.

## BookIntro tooling deleted

`DefineBookIntroStyle` and `ApplyBookIntroAfterDatAuthRef`
(approx 185 lines) removed from `basFixDocxRoutines.bas`. Both
referenced an obsolete style; `Introduction` is the active
replacement. No callers in the project remain.

## Bucket-1 / bucket-2 promotions (8 styles)

`DumpStyleProperties` outputs under `rpt/Styles/` were used to
write `AuditOneStyle` lines for the 8 dumped candidates. Split
by Word style type:

**Bucket 1 (fully specified, paragraph styles - 3 added):**

| Style | Font | Size | Align | LineRule | LineSp | SpB / SpA | Bold / Italic | BaseStyle |
|---|---|---|---|---|---|---|---|---|
| SpeakerLabel | Carlito | 9 | Left | Single | 12 | 3 / 2 | 0 / 0 | "" |
| AuthorSectionHead | Liberation Serif | 14 | Left | Single | 12 | 12 / 6 | 0 / -1 | "" |
| AuthorBookSections | Carlito | 11 | Left | Single | 12 | 0 / 0 | 0 / 0 | "Normal" |

`AuthorBookSections` carries a single tab stop (378 pt, Right, Dots)
and was added to the `AuditStyleTabs` block. It was also added to
`PromoteApprovedStyles` (was Priority=99, now near the
`AuthorBookRef` cluster).

**Bucket 2 (existence + font/size only, character styles - 5 added):**

`Selah`, `EmphasisBlack`, `EmphasisRed`, `Words of Jesus`,
`AuthorQuote` - all `Carlito` 9pt. Bold / Italic / Color checks
deferred until `AuditOneStyle` is extended for character styles
(same parking pattern as `Footnote Reference`). The color values
captured in the dumps (`EmphasisRed`, `Words of Jesus`,
`AuthorQuote`) feed the new "Define colors used in the docx" task.

Header comment in `basTEST_aeBibleConfig.bas` realigned: 45 styles
(34 bucket-1 / 8 bucket-2 / 3 not-yet-created) + 6 tab-stop specs =
**51 checks**.

Expected on next `RUN_TAXONOMY_STYLES`: 47 PASS / 4 FAIL across
51 checks (current 39 PASS + 3 paragraph-style PASS + 5
character-style PASS + 1 tab-stop PASS).
