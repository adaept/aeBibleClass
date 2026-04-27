# 01 — Approved style taxonomy

The Study Bible uses a deliberate, ordered set of paragraph and
character styles. This page is the human-readable index. Code-level
authority lives in:

- **`approved` array** in `src/basTEST_aeBibleConfig.bas` — sequence
  and membership.
- **`RUN_TAXONOMY_STYLES`** — expected property values for QA.

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

Latest run (2026-04-26): **43 approved styles succeeded**, ~4 sec.

### Validated (priorities 1–36)

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
| 17 | TitleOnePage |
| 18 | ListItem |
| 19 | ListItemBody |
| 20 | ListItemTab |
| 21 | AuthorBookRefHeader |
| 22 | AuthorBookRef |
| 23 | CenterSubText |
| 24 | Heading 1 |
| 25 | CustomParaAfterH1 |
| 26 | Brief |
| 27 | DatAuthRef |
| 28 | Heading 2 |
| 29 | Chapter Verse marker |
| 30 | Verse marker |
| 31 | Footnote Reference |
| 32 | Footnote Text |
| 33 | Psalms BOOK |
| 34 | PsalmSuperscription |
| 35 | Selah |
| 36 | PsalmAcrostic |

### Pending re-validation (priorities 37+)

Order inherited from earlier passes; will be re-walked.

| Prio | Style |
|---:|---|
| 37 | BodyTextIndent |
| 42 | EmphasisBlack |
| 43 | EmphasisRed |
| 44 | Words of Jesus |
| 45 | AuthorSectionHead |
| 46 | AuthorQuote |
| 47 | Normal |

### Reserved gaps

Priorities 38–41 are reserved for future insertions without
wholesale renumbering. (Earlier "gap at 17" was not a deliberate
reservation; it was a `TitleOnePage` duplicate in the array,
fixed 2026-04-26 — `TitleOnePage` now correctly holds 17.)

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

`Lamentations` was previously listed; **removed from the array**
on 2026-04-26 (book content standardized on `BodyText` for now).
The orphan `style_Lamentations.txt` was auto-cleaned by the
`DumpAllApprovedStyles` orphan prompt.

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

### Body text (priorities 24+)

`Heading 1`, `Heading 2`, `BodyText`, `BodyTextIndent`,
`CustomParaAfterH1`, `DatAuthRef`, `Brief`. Verse-level styles:
`Chapter Verse marker`, `Verse marker`, `Words of Jesus`,
`EmphasisBlack`, `EmphasisRed`.

### Special book treatments

- `Psalms BOOK` — book-level Psalms heading style.
- `PsalmSuperscription` — the prefatory line attributed to
  authorship / context (e.g., "A psalm of David").
- `Selah` — the Hebrew musical / liturgical interjection.
- `PsalmAcrostic` — Hebrew-letter section markers in acrostic
  Psalms (notably Psalm 119).
- `Lamentation` (singular) is audited via `AuditOneStyle` but
  not currently in the promoted approved array; revisit when its
  role is decided.

### Footnotes

`Footnote Reference`, `Footnote Text`.

### Author commentary

`AuthorBodyText`, `AuthorBookRefHeader`, `AuthorBookRef`,
`AuthorSectionHead`, `AuthorQuote`. Distinct font family from the
Bible body to signal commentary.

### Anchor

`Normal` (priority 47) — deliberately the last entry. Replaced
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
