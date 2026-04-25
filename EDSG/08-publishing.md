# 08 — Publishing

Status: skeleton. The EDSG is intended to be publishable in three
forms: web (markdown via GitHub), digital document (`.docx`), and
print (PDF / physical). The print form uses the **same Study Bible
styles and templates** as the Bible itself — dogfooding the system.

## Three output forms

### 1. Web — markdown via GitHub

Source of truth. Lives in `/EDSG/`, rendered automatically by GitHub.
No build step required; commits are immediately readable.

Audience: developers, editors, anyone with a browser.

### 2. Digital document — `.docx`

Imported into a Word document for offline reading or controlled
distribution. The import targets the Study Bible template so the
EDSG inherits the same styles, fonts, and layout conventions.

Open question: which import path?

- **Pandoc** — markdown → docx with style mapping. Most flexible,
  needs tuning so each markdown construct (heading, table, code
  block, callout) maps to an approved style.
- **Word's built-in markdown import** (where available) — less
  control, may not preserve style mappings.
- **Custom VBA import** — write a small routine that ingests
  `/EDSG/*.md` and emits a styled docx. Most work; tightest fit.

Recommendation pending — likely Pandoc plus a custom style mapping
file.

### 3. Print — PDF / physical book

Generated from the docx form via Word's PDF export. The print form is
intentional dogfooding:

- Every approved style in [01-styles](01-styles.md) gets exercised by
  the EDSG content.
- Page layout decisions (headers, footers, margins, indents) are
  validated on a real document that isn't the Bible.
- If the EDSG looks wrong in print, the Bible would too.

## Markdown ↔ Style mapping

Each markdown construct maps to one approved style. Draft:

| Markdown | Approved style |
|---|---|
| `# H1 heading` | `Heading 1` |
| `## H2 heading` | `Heading 2` |
| Body paragraph | `BodyText` |
| Indented quote | `BodyTextIndent` (when defined) |
| List item | `ListItem` |
| Code block | TBD — needs a `CodeBlock` style? |
| Inline code | TBD — character style |
| Table | TBD — Word table style |
| Link | Standard hyperlink character |

`CodeBlock` and the inline-code character style are not yet in the
approved list. Adding them is the first concrete EDSG-driven style
addition once the publishing path is selected.

## Repo connection

- Source markdown: `/EDSG/`.
- Generated docx: `/EDSG/build/EDSG.docx` (gitignored — derived).
- Generated PDF: `/EDSG/build/EDSG.pdf` (gitignored — derived).
- Build script: `/EDSG/build.cmd` or `/EDSG/build.sh` (TBD).

Build artifacts are not committed; the source markdown is the only
versioned form.

## Open questions

1. Print template — clone the Study Bible `.docm` directly, or a
   derivative with EDSG-specific section breaks?
2. Pagination — same headers / footers as the Bible, or EDSG-specific
   (e.g., "EDSG — Chapter 4")?
3. Cover / front matter for the EDSG — same `FrontPageTopLine` /
   `Title` / `TitleVersion` styles applied with EDSG content, or
   different layout?
4. Distribution — print-on-demand, internal PDF only, public release?

These shape `08-publishing.md` substantially. This page expands when
they're answered.
