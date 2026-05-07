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
