# 01 — Approved style taxonomy

The Study Bible uses a deliberate, ordered set of paragraph and
character styles. This page is the human-readable index. Code-level
authority lives in:

- **`approved` array** in `src/basTEST_aeBibleConfig.bas` — sequence
  and membership.
- **`RUN_TAXONOMY_STYLES`** — expected property values for QA.

> ## ⚑ Important — taxonomy audit final-state goal
>
> The current `RUN_TAXONOMY_STYLES` covers a curated 19-style subset
> (**9 fully specified + 7 existence-verified + 3 not-yet-created**,
> as of 2026-04-29). **This is a transitional state, not the
> destination.** The final-state resolution is for the audit to map
> **every approved style** with a real (non-sentinel) expected spec,
> so that any property drift on any approved style is caught
> immediately.
>
> Promoted-but-unaudited styles (e.g. `SpeakerLabel` as of
> 2026-04-29) are temporarily off-radar for property drift —
> `DumpAllApprovedStyles` only confirms existence and priority, not
> font / size / alignment / indents / line-spacing / spacing.
>
> Each move from bucket 2 (existence-verified, full spec pending)
> into bucket 1 (fully specified) is a measurable step toward full
> drift coverage. Progress so far:
>
> - **2026-04-29** — promoted 7 styles to bucket 1: `Heading 1`,
>   `Heading 2`, `CustomParaAfterH1`, `DatAuthRef`, `Brief`,
>   `Psalms BOOK`, `Footnote Text`. Bucket 1 is 2 -> 9.
> - **2026-05-06** — taxonomy resync after a partial QA-alignment
>   pass on five paragraph styles (`Heading 2`, `Brief`,
>   `Psalms BOOK`, `CustomParaAfterH1`, `Footnote Text`). Three
>   audit lines updated to match new descriptive specs
>   (`Heading 2`, `Brief`, `Psalms BOOK`). `RUN_TAXONOMY_STYLES`
>   now reports 24 PASS / 4 FAIL — all four FAILs are NOT-FOUND
>   placeholders. See `rvw/Code_review 2026-05-06.md` § 9.
> - **2026-05-06** — first prescriptive-property pass:
>   `AuditOneStyle` extended with optional `sExpBaseStyle`;
>   `BaseStyle = ""` invariant enforced on `CustomParaAfterH1`,
>   `Brief`, `Footnote Text`, `Psalms BOOK`. `PsalmAcrostic` and
>   `PsalmSuperscription` promoted from bucket 2 to bucket 1 with
>   full descriptive specs (also at `BaseStyle = ""`). Bucket 1 is
>   9 -> 11; total checks 28 -> 30. `RUN_TAXONOMY_STYLES`:
>   **26 PASS / 4 FAIL**. See `rvw/Code_review 2026-05-06.md` § 10.
> - **2026-05-06** — front-matter & TOC bucket-1 promotion:
>   12 new bucket-1 entries (`FrontPageTopLine`, `TitleEyebrow`,
>   `TitleVersion`, `FrontPageBodyText`, `BodyTextTopLineCPBB`,
>   `Acknowledgments`, `AuthorBodyText`, `Contents`,
>   `BibleIndexEyebrow`, `BibleIndex`, `Introduction`,
>   `TitleOnePage`); `Title` promoted from bucket 2; `ContentsRef`
>   gained `BaseStyle = ""`; `ContentsRef` tab stop added to the
>   tab-stops block. Bucket 1: 11 -> 24. Tab-stop coverage:
>   4 -> 5 styles. `RUN_TAXONOMY_STYLES`: **39 PASS / 4 FAIL across
>   43 checks**. See `rvw/Code_review 2026-05-06.md` §§ 11-12.
> - `Lamentation` removed from the audit (style deleted).
> - `Footnote Reference` added to the audit, parked in bucket 2
>   pending an extension to `AuditOneStyle` to check
>   character-style Bold / Italic / Color properties.
> - Specs encoded as **descriptive** (capture current document
>   state) rather than prescriptive — see `rvw/Code_review
>   2026-04-25.md` "Spec promotion: descriptive vs prescriptive"
>   for decision rationale.

Status: WIP. The page walk has validated priorities 1–36; positions
**37 and beyond are still pending re-validation** as the walk
continues. The array's order through priority 36 reflects book
occurrence (the canonical convention; see
[04-qa-workflow](04-qa-workflow.md)).

## Current approved order (snapshot)

Priority order from `WordEditingConfig` (which runs
`PromoteApprovedStyles` then `DumpPrioritiesSorted`). For a
page-keyed view, run `ListApprovedStylesByBookOrder` — the live
document is the authority, this snapshot ages.

Snapshot rebuilt 2026-05-17 from `GetApprovedStyles()` in
`basTEST_aeBibleConfig.bas`. **52 entries** in the array (48
present in the document + 4 tracking placeholders). VerseText is
the live verse-paragraph style since 2026-05-01.

### Approved styles (full list)

| Prio | Style | Notes |
|---:|---|---|
| 1 | TheHeaders | |
| 2 | BodyText | residual non-verse paragraph (front matter, chapter intros) |
| 3 | TheFooters | |
| 4 | FrontPageTopLine | |
| 5 | TitleEyebrow | |
| 6 | Title | |
| 7 | TitleVersion | |
| 8 | FrontPageBodyText | |
| 9 | BodyTextTopLineCPBB | |
| 10 | Acknowledgments | |
| 11 | AuthorBodyText | |
| 12 | Contents | |
| 13 | ContentsRef | |
| 14 | BibleIndexEyebrow | |
| 15 | BibleIndex | |
| 16 | BibleIndexList | |
| 17 | Introduction | |
| 18 | TitleOnePage | |
| 19 | AuthorListItem | canonical `BaseStyle = ""` example |
| 20 | AuthorListItemBody | |
| 21 | AuthorListItemTab | |
| 22 | AuthorBookRefHeader | |
| 23 | AuthorBookRef | |
| 24 | AuthorBookSections | |
| 25 | CenterSubText | |
| 26 | Heading 1 | |
| 27 | CustomParaAfterH1 | LineSpacing fixed to Single 2026-05-16 |
| 28 | Brief | |
| 29 | DatAuthRef | |
| 30 | Heading 2 | |
| 31 | Chapter Verse marker | |
| 32 | Verse marker | |
| 33 | **VerseText** | **primary verse paragraph style** |
| 34 | Footnote Reference | |
| 35 | BookHyperlink | one-form hyperlink rule (2026-05-15) |
| 36 | Footnote Text | `LineSpacingRule = Exactly 8` known exception |
| 37 | Psalms BOOK | |
| 38 | PsalmSuperscription | |
| 39 | Selah | |
| 40 | PsalmAcrostic | |
| 41 | SpeakerLabel | |
| 42 | BodyTextContinuation | not present in document — tracking placeholder |
| 43 | AppendixTitle | not present in document — tracking placeholder |
| 44 | AppendixBody | not present in document — tracking placeholder |
| 45 | EmphasisBlack | |
| 46 | EmphasisRed | |
| 47 | Words of Jesus | |
| 48 | AuthorSectionHead | |
| 49 | ParallelHeader | |
| 50 | ParallelText | |
| 51 | Normal | last approved entry — anchor for the hide-sweep |
| 52 | FargleBlargle | deliberate canary — confirms missing-style diagnostic |

### Missing from document

`BodyTextContinuation`, `AppendixTitle`, `AppendixBody`,
`FargleBlargle` — in the `approved` array but not present in the
current document; reported by `PromoteApprovedStyles` as a
diagnostic. Kept as tracking placeholders pending decisions on
each (create + populate, or remove from array).

### Removed from the array

- `Lamentations` — removed 2026-04-26 (book content standardized
  on `BodyText`). Orphan `style_Lamentations.txt` auto-cleaned.
- `BookIntro` — removed; pending decision on whether to define
  and promote, or close as not needed.
- `BodyTextIndent` — removed during the 2026-05-01 VerseText
  migration; no longer in the approved array.
- `AuthorQuote` — removed; front matter usage was never
  finalized.

### Reserved gaps

None as of 2026-05-17. The prior gap at 39–42 was filled by
`SpeakerLabel`, `BodyTextContinuation`, `AppendixTitle`,
`AppendixBody`. Future insertions follow at the next free
priority.

## Style categories

Loose grouping for orientation. Authoritative roles live in
`RUN_TAXONOMY_STYLES`.

### Front matter (priorities 1–23)

Title block, headers/footers, contents and index pages, author
introduction, list-item conventions, author-reference header.
`TheHeaders` / `TheFooters` apply via the EvenPages header/footer
slot — see [05-headers-footers](05-headers-footers.md).

Notable members: `FrontPageTopLine`, `TitleEyebrow`, `Title`,
`TitleVersion`, `FrontPageBodyText`, `BodyTextTopLineCPBB`,
`Acknowledgments`, `AuthorBodyText`, `Contents`, `ContentsRef`,
`BibleIndexEyebrow`, `BibleIndex`, `Introduction`, `TitleOnePage`,
`ListItem`, `ListItemBody`, `ListItemTab`,
`AuthorBookRefHeader`, `AuthorBookRef`, `CenterSubText`.

### Body text (priorities 26+)

`Heading 1`, `Heading 2`, `BodyText` (residual non-verse
paragraphs), `VerseText` (the verse paragraph style since
2026-05-01), `CustomParaAfterH1`, `DatAuthRef`, `Brief`.
Verse-level character styles: `Chapter Verse marker`,
`Verse marker`, `Words of Jesus`, `EmphasisBlack`, `EmphasisRed`,
`BookHyperlink`.

### Special book treatments

- `Psalms BOOK` — book-level Psalms heading style.
- `PsalmSuperscription` — the prefatory line attributed to
  authorship / context (e.g., "A psalm of David").
- `Selah` — the Hebrew musical / liturgical interjection.
- `PsalmAcrostic` — Hebrew-letter section markers in acrostic
  Psalms (notably Psalm 119).

### Footnotes

`Footnote Reference`, `Footnote Text`.

### Author commentary

`AuthorBodyText`, `AuthorBookRefHeader`, `AuthorBookRef`,
`AuthorBookSections`, `AuthorSectionHead`. Distinct font family
from the Bible body to signal commentary.

### Anchor

`Normal` (priority 51) — deliberately the last approved entry
before the `FargleBlargle` canary. Replaced operationally by
`BodyText` (for non-verse paragraphs) and `VerseText` (for
verses); retained in the array as a "pin-everything-else-above"
anchor.

## QA checklist for every approved style

These four properties should default as below for almost every
approved style. Documented exceptions only. See `rvw/Code_review
2026-04-21.md` § QA checklist for the original definition.

| # | Property | Expected | UI equivalent | Applies to |
|---|----------|----------|---------------|------------|
| 1 | `.BaseStyle` | `""` | Style based on: **(no style)** | All styles ([why](10-list-paragraph-bug.md)) |
| 2 | `.AutomaticallyUpdate` | `False` | "Automatically update" checkbox NOT selected | Paragraph only |
| 3 | `.QuickStyle` | `False` | Style does not appear in the Quick Styles gallery | All styles |
| 4 | `.ParagraphFormat.LineSpacingRule` | `0` (`wdLineSpaceSingle`) | Line spacing: Single | Paragraph only |

`AutomaticallyUpdate = True` is the silent killer of style discipline
— direct formatting on one paragraph silently rewrites the style for
all others using it. Always `False`.

## Colour discipline: two tiers, no third

Every style's `Font.Color` must fall into exactly one of two tiers.
There is no third option.

| Tier | Value | Intent | Examples |
|---|---|---|---|
| **1. Default text** | `wdColorAutomatic` (`-16777216`) | Render as the page's foreground colour. Black on white in default theme; flips with dark mode; lands black on printed PDF. | TheHeaders, TheFooters, Selah, EmphasisBlack, Body Text, most paragraph styles. |
| **2. Deliberate colour** | A palette-registered `RgbLong` value | Render in a specific named colour regardless of page. Lives in [`basBiblePalette`](../src/basBiblePalette.bas) `GetPalette()`. | Hyperlink → DarkBlue, Footnote Reference → Blue, Verse marker → Emerald, Chapter Verse marker → Orange, Words of Jesus → DarkRed. |

Anything else is an anomaly. Common cases:

- Hand-typed `RGB(0, 0, 0)` masquerading as "default text" — should be `wdColorAutomatic`. Locking explicit black makes the run invisible on a dark background and forfeits theme portability for zero print benefit.
- An off-palette specific colour like `#C00000` where the palette has `#800000` (`DarkRed`) — should be repointed to the palette entry, or, if the off-palette colour is deliberate, *added* to the palette as a new named entry.
- `Font.ObjectThemeColor <> wdThemeColorNone` — Office theme colour. **Banned outright**: too niche, too template-coupled, not portable to non-Office renderers.

### Why two tiers and not "explicit literals everywhere"

The earlier-proposed "convert all Automatic to explicit `wdColorBlack`" rule was rejected (2026-05-14) on these grounds:

- **Print target via PDF lands the same.** Automatic → black on paper, explicit `wdColorBlack` → black on paper. No print difference.
- **Theme portability survives.** Automatic flips correctly with dark mode; locked black does not. For a future online edition, all "default text" needs to invert — that's what Automatic delivers.
- **Translator workflow stays clean.** A translator inheriting this doc never has to know "normal text is RGB(0,0,0)." They work with style names and never touch colour metadata.
- **Single discipline, two states.** Style colour is either Automatic *(default-text intent)* or a palette-named value *(deliberate-colour intent)*. No third state to debate; anomalies are obvious.

### Audits

Two checks cover the rule:

- **`AuditNonPaletteStyleColors`** (`basStyleInspector`) — walks every style, classifies `Font.Color` into Automatic / Palette / Anomaly. Returns anomaly count; expected 0 in steady state. Drives manual decisions on each anomaly (repoint Automatic, repoint to palette, or add to palette).
- **`AuditThemeColorUsage`** *(planned)* — walks every style, reports any with `Font.ObjectThemeColor <> wdThemeColorNone`. Expected 0. Wired into `RUN_THE_TESTS` once stable.

A third audit, **`AuditDeliberateColourCompliance`** *(planned)*, will verify each named-deliberate-colour style carries the exact palette value expected (e.g., Hyperlink → DarkBlue, not some hand-typed near-blue). Needs a style → palette-name registry stored in `basBiblePalette`.

## State-aware styles: print-locking

Some Word character styles change appearance based on interactive
state. The canonical case is the **Hyperlink / FollowedHyperlink**
pair: an unvisited link renders one way, a clicked-once link shifts to
the FollowedHyperlink colour. That state shift is fine for digital
reading but breaks print-target consistency — the printed page reflects
the reader's current click history rather than a stable design.

**Print-locking pattern.** Force the state-aware style to match its
non-state counterpart, so the visited / followed state has no visible
effect:

1. Set `Styles("Hyperlink").Font` to the target palette colour +
   underline.
2. Set `Styles("FollowedHyperlink").Font` to the **same** colour +
   underline.
3. For every existing instance (here, every `Doc.Hyperlinks`), force
   `Range.Style = Hyperlink` so manual run-level formatting can't
   override the style.

Reference implementation: `LockHyperlinksToPalette` in
`basTEST_aeBibleTools` (sources colour from
[`basBiblePalette`](../src/basBiblePalette.bas) → `"DarkBlue"`,
`#000080`). Drift audit: `?AuditHyperlinkStyling` in
`basStyleInspector` — returns 0 when every hyperlink in the doc
matches the convention.

**Why DarkBlue, not Blue.** `Footnote Reference` style already
occupies pure Blue (`#0000FF`). Keeping `Hyperlink` on the same colour
collides in any colour audit and reads similarly in print. Moving
hyperlinks to `DarkBlue` (`#000080`, matches `wdColorDarkBlue`)
disambiguates the two roles without sacrificing the "this is a link"
underline signal. The pattern extends to any future state-aware
character style — pick a print-stable palette colour, then lock both
states to it.

### Companion rule: no clickable hyperlinks anywhere

**Hyperlinks in this document are visual references, not interactive
controls.** Every run that looks like a link must be display-only —
the underlying `Hyperlink` object must be absent — and styled with
the **`BookHyperlink`** custom character style (not Word's built-in
`Hyperlink`).

Rationale:

- The document's primary target is **print**. Reader holding the
  printed book cannot click anything; live link objects serve no
  purpose.
- Clickability is a **future-mode concern** for an eventual online
  edition. At that build time, online-edition-specific tooling can
  re-attach link objects to the styled text. The print master stays
  clean.
- Single-form discipline simplifies i18n and translator work — a
  translator never has to think about link mechanics, only display
  styling.

**One-form definition for this doc:** a hyperlink is a web URL
pointing to an online tool, rendered as text with the `BookHyperlink`
character style. `BookHyperlink` pins all four font properties
explicitly:

| Property | Value |
|---|---|
| `Font.Name` | `Carlito` |
| `Font.Size` | `9` |
| `Font.Color` | palette `DarkBlue` (`#000080`) |
| `Font.Underline` | `wdUnderlineSingle` |

Word's built-in `Hyperlink` and `FollowedHyperlink` styles are **not
used here**. They inherit font and size from the paragraph context,
which means a hyperlink inside (for example) an `AuthorListItemTab`
paragraph at 11pt renders at 11pt — breaking the print uniformity
the rule demands. The custom `BookHyperlink` style is fully
self-contained and renders identically regardless of surrounding
paragraph style.

Word's other clickable mechanisms — `REF` / `PAGEREF` / internal
bookmark Hyperlinks — are **not in use here**. Any that appear are
anomalies to remove.

### Why a custom style and not the built-in

Earlier work (2026-05-13/14) tried to lock the built-in
`Hyperlink` style to palette DarkBlue. The audit appeared to pass
but a test on 2026-05-15 demonstrated the gap: a hyperlink pasted
into an `AuthorListItemTab` paragraph carried size 11 (inherited
from the paragraph) while the rule demands size 9. The built-in
style cannot enforce font/size; the inheritance from paragraph
context wins. A custom character style with explicit Font.Name and
Font.Size resolves the structural issue.

### Enforcement

- **`DefineBookHyperlinkStyle`** (`basFixDocxRoutines`) — one-shot,
  creates the `BookHyperlink` character style with the four pinned
  properties. Idempotent (skips if already present).
- **`LockBookHyperlinks`** (`basTEST_aeBibleTools`) — three-step
  workflow:
  1. Walk every story; migrate any run styled built-in `Hyperlink`
     to `BookHyperlink` (catches paste-ins and Word's URL
     auto-format output).
  2. Walk every story's `Hyperlinks` collection; restyle each
     `hl.Range` to `BookHyperlink`, then `Hyperlink.Delete` to
     remove the click target.
  3. Walk every story; force-apply the four `BookHyperlink`
     properties on every `BookHyperlink`-styled run (idempotent
     override; strips paste-in direct formatting).
- **`AuditBookHyperlinkStyling`** (`basStyleInspector`) — verifies
  the four properties on every `BookHyperlink`-styled run. Per-
  property mismatch reporting. Expected 0 anomalies after the lock.
- **RUN_THE_TESTS slot 17 → `CountActiveHyperlinks`**
  (`src/aeBibleClass.cls`). Sums `story.Hyperlinks.Count` across
  all StoryRanges. Expected value 0; any non-zero result is a rule
  violation.

The no-hyperlinks-in-footnotes rule is **subsumed** by this broader
rule (zero hyperlinks anywhere implies zero in footnotes).

Pattern: when a style discipline implies a content rule, codify the
content rule as its own test rather than leaning on the style audit
to catch it indirectly. Two audits cover the two concerns:

- `AuditBookHyperlinkStyling` — *how* a hyperlink is dressed (all
  four font properties must match the pinned values).
- `CountActiveHyperlinks` (test 17) — *whether* an active link
  object exists at all (expected 0 anywhere).

### Per-installation recommendation: disable URL auto-format

Word's "AutoFormat As You Type" feature auto-converts typed URLs
into active `Hyperlink` objects styled with the built-in `Hyperlink`
character style. That's incompatible with the rule above — every
typed URL becomes a rule violation that `LockBookHyperlinks` then
has to migrate.

**Disable it once, per editor's installation:**

1. File → Options → Proofing → AutoCorrect Options...
2. AutoFormat As You Type tab.
3. Under "Replace as you type", uncheck **"Internet and network
   paths with hyperlinks"**.
4. (Optional) Also AutoFormat tab → same checkbox, off.

This is a per-installation Word setting, not a per-document setting,
so it cannot be enforced through the doc itself. Editors and
translators forking the project should make this change as part of
their initial setup. With it disabled, typed URLs remain plain text
until styled deliberately to `BookHyperlink`; pastes from the web
no longer auto-create link objects.

If a translator misses the step, `LockBookHyperlinks` will catch
and migrate any built-in `Hyperlink` runs on its next run. The
setting prevents the work, not the rule.

## How to add a new style

See [02-editing-process](02-editing-process.md) § Style design.
