# 01 — Approved style taxonomy

The Study Bible uses a deliberate, ordered set of paragraph and
character styles. This page is the human-readable index. Code-level
authority lives in:

- **`approved` array** in `src/basTEST_aeBibleConfig.bas` — sequence
  and membership.
- **`RUN_TAXONOMY_STYLES`** — expected property values for QA.

Status: WIP. The page walk has validated priorities 1–33; positions
**34 and beyond are still pending re-validation** as the walk
continues. The array's order through priority 33 reflects book
occurrence (the canonical convention; see
[04-qa-workflow](04-qa-workflow.md)).

## Current approved order (snapshot)

Priority order from `WordEditingConfig` (which runs
`PromoteApprovedStyles` then `DumpPrioritiesSorted`). For a
page-keyed view, run `ListApprovedStylesByBookOrder` — the live
document is the authority, this snapshot ages.

### Validated (priorities 1–33)

| Prio | Style |
|---:|---|
| 1 | TheHeaders |
| 2 | BodyText |
| 3 | TheFooters |
| 4 | FrontPageTopLine |
| 5 | TitleEyebrow |
| 6 | Title |
| 7 | TitleVersion |
| 8 | FrontPageBodyText |
| 9 | BodyTextTopLineCPBB |
| 10 | Acknowledgments |
| 11 | AuthorBodyText |
| 12 | Contents |
| 13 | ContentsRef |
| 14 | BibleIndexEyebrow |
| 15 | BibleIndex |
| 16 | Introduction |
| 18 | ListItem |
| 19 | ListItemBody |
| 20 | ListItemTab |
| 21 | AuthorBookRefHeader |
| 22 | AuthorBookRef |
| 23 | TitleOnePage |
| 24 | CenterSubText |
| 25 | Heading 1 |
| 26 | CustomParaAfterH1 |
| 27 | Brief |
| 28 | DatAuthRef |
| 29 | Heading 2 |
| 30 | Chapter Verse marker |
| 31 | Verse marker |
| 32 | Footnote Reference |
| 33 | Footnote Text |

### Pending re-validation (priorities 34+)

Order inherited from earlier passes; will be re-walked.

| Prio | Style |
|---:|---|
| 34 | Lamentations |
| 35 | Psalms BOOK |
| 36 | BodyTextIndent |
| 41 | EmphasisBlack |
| 42 | EmphasisRed |
| 43 | Words of Jesus |
| 44 | AuthorSectionHead |
| 45 | AuthorQuote |
| 46 | Normal |

### Gaps

Priorities 17, 37–40 are intentional reserved gaps for future
insertions without wholesale renumbering.

### Missing from document

The following are in the `approved` array but not present in the
current document; reported by `PromoteApprovedStyles` as a
diagnostic. Kept in the array as tracking placeholders:

- `BodyTextContinuation`
- `BookIntro`
- `AppendixTitle`
- `AppendixBody`
- `FargleBlargle` (deliberate canary — confirms the missing-style
  diagnostic is wired correctly)

## Style categories

Loose grouping for orientation. Authoritative roles live in
`RUN_TAXONOMY_STYLES`.

### Front matter (priorities 1–24)

Title block, headers/footers, contents and index pages, author
introduction, list-item conventions, author-reference header.
`TheHeaders` / `TheFooters` apply via the EvenPages header/footer
slot — see [05-headers-footers](05-headers-footers.md).

Notable members: `FrontPageTopLine`, `TitleEyebrow`, `Title`,
`TitleVersion`, `FrontPageBodyText`, `BodyTextTopLineCPBB`,
`Acknowledgments`, `AuthorBodyText`, `Contents`, `ContentsRef`,
`BibleIndexEyebrow`, `BibleIndex`, `Introduction`, `ListItem`,
`ListItemBody`, `ListItemTab`, `AuthorBookRefHeader`,
`AuthorBookRef`, `TitleOnePage`, `CenterSubText`.

### Body text (priorities 25+)

`Heading 1`, `Heading 2`, `BodyText`, `BodyTextIndent`,
`CustomParaAfterH1`, `DatAuthRef`, `Brief`. Verse-level styles:
`Chapter Verse marker`, `Verse marker`, `Words of Jesus`,
`EmphasisBlack`, `EmphasisRed`.

### Special book treatments

`Psalms BOOK`, `Lamentations` — book-level stylistic differences
(e.g., indentation patterns).

### Footnotes

`Footnote Reference`, `Footnote Text`.

### Author commentary

`AuthorBodyText`, `AuthorBookRefHeader`, `AuthorBookRef`,
`AuthorSectionHead`, `AuthorQuote`. Distinct font family from the
Bible body to signal commentary.

### Anchor

`Normal` (priority 46) — deliberately the last entry. Replaced
operationally by `BodyText`; retained in the array as a
"pin-everything-else-above" anchor.

## QA checklist for every approved style

These four properties should default as below for almost every
approved style. Documented exceptions only. See `rvw/Code_review
2026-04-21.md` § QA checklist for the original definition.

| # | Property | Expected | UI equivalent | Applies to |
|---|----------|----------|---------------|------------|
| 1 | `.BaseStyle` | `""` | Style based on: **(no style)** | All styles |
| 2 | `.AutomaticallyUpdate` | `False` | "Automatically update" checkbox NOT selected | Paragraph only |
| 3 | `.QuickStyle` | `False` | Style does not appear in the Quick Styles gallery | All styles |
| 4 | `.ParagraphFormat.LineSpacingRule` | `0` (`wdLineSpaceSingle`) | Line spacing: Single | Paragraph only |

`AutomaticallyUpdate = True` is the silent killer of style discipline
— direct formatting on one paragraph silently rewrites the style for
all others using it. Always `False`.

## How to add a new style

See [02-editing-process](02-editing-process.md) § Style design.
