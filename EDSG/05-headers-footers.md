# 05 — Headers and footers

Status: WIP. Captures what's known after the 2026-04-24 audit; will
expand as the page walk continues.

## Document structure

The Study Bible document has **148 sections**. Most sections are
"Linked to Previous" — they inherit header / footer content from an
earlier section. A handful of sections own unlinked headers / footers
that define the actual content.

## Header / footer slots

Word offers three slots per section, indexed `1..3`:

| Index | `WdHeaderFooterIndex` constant | Role |
|-------|-------------------------------|------|
| 1 | `wdHeaderFooterEvenPages` | Even-page header / footer |
| 2 | `wdHeaderFooterPrimary` | Primary (odd-page if Different Even/Odd; otherwise all) |
| 3 | `wdHeaderFooterFirstPage` | First-page header / footer (if Different First Page) |

The Study Bible uses **Different Even / Odd Pages**. The `TheHeaders`
and `TheFooters` styles live in `Header(1)` / `Footer(1)` (even-page
slot) — the `Header(2)` / `Footer(2)` primary slot in section 1 is
empty with the built-in `Header` / `Footer` style.

That detail caught the QA tooling out: an early manual check looked at
`Sections(1).Headers(2)` (primary), saw the empty built-in style, and
mistakenly reported the document had no `TheHeaders` content. Always
check all three slots.

## Diagnostic — `DumpHeaderFooterStyles`

```vba
DumpHeaderFooterStyles
```

Writes `rpt/Styles/header_footer_audit.txt` — one line per
`Section x Slot` combination. Sample (from the actual document):

```
Sec 001 Header(1)            paras=1  style=TheHeaders   text=[<tab>]
Sec 001 Header(2)            paras=1  style=Header       text=[]
Sec 001 Footer(1)            paras=1  style=TheFooters   text=[]
Sec 002 Header(1) linked     paras=1  style=TheHeaders   text=[<tab>]
Sec 002 Header(3)            paras=1  style=TheHeaders   text=[<tab>]
Sec 004 Footer(1)            paras=1  style=TheFooters   text=[1]
... (888 lines total for 148 sections)
```

`linked` indicates `LinkToPrevious = True` for that slot.

## Why `Find` doesn't work in headers

`ListApprovedStylesByBookOrder` uses `Range.Find` on the main body and
on footnotes / endnotes. For headers and footers, the same approach
returns no match — Find with `.Text = ""` plus `.Style = oStyle`
doesn't reliably hit content that consists of only a tab or only a
paragraph mark.

The fix in `FirstPageInParagraphs` is to **iterate paragraphs
directly** and compare `Para.Style.NameLocal` to the target. Headers
and footers are tiny (typically 1 paragraph), so iteration is cheap
and exact.

## Why `Information(wdActiveEndPageNumber)` is wrong for headers

For a paragraph inside a header story, Word returns a section-anchor-
related page number (in this document: 417 — coincidentally the
Psalms section transition), not the page where the header first
applies. The fix: if a paragraph match is found in a header / footer,
ignore the reported page and let the page-1 fallback kick in.

The fallback is exact for this document because `TheHeaders` /
`TheFooters` tile from page 1. If a future style uses a header that
legitimately starts on a later page (e.g., a "Maps appendix"
header beginning on page 500), the fallback over-reports as page 1.
Flagged YAGNI in `rvw/Code_review 2026-04-21.md`.

## Linked vs unlinked

When auditing, distinguish:

- **Unlinked anchor**: a section that owns the actual content. The
  `LinkToPrevious` flag is `False`. Editing here changes the content
  for downstream linked sections.
- **Linked**: inherits from the previous unlinked anchor. Editing
  here forks the content, breaking the link.

For pure consumption (the page-walk QA), linked sections still report
the inherited content. For changes, find the upstream anchor and
edit there.

## Open questions for future page walks

- Do all unlinked anchors agree (all use `TheHeaders` / `TheFooters`)?
  Spot-check the audit file as the page walk proceeds.
- Are there sections with `Different First Page` that need the
  `Header(3)` / `Footer(3)` slot to use the same styles?
- Are there book-specific header overrides (e.g., Psalms with a
  different running head)?
