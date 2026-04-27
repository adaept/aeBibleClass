# Editing and Design Style Guide (EDSG)

Operational manual for editors and developers working on the Study
Bible — especially anyone preparing a localization (i18n).

This guide answers four questions:

1. **What styles exist** and what each is for.
2. **How to apply, audit, and extend** them.
3. **Which routines** to run at each editing stage.
4. **How QA gates** a release.

The EDSG is itself produced using the same Study Bible styles and
templates — dogfooding the system.

## Source-of-truth map

Different artifacts answer different questions. Know which to consult.

| Question | Authority | See |
|---|---|---|
| What property values should a style have? | `RUN_TAXONOMY_STYLES` constants in code | [01-styles](01-styles.md), [03-inspection-tools](03-inspection-tools.md) |
| What are the actual indent measurements? | The Word UI ruler (WRIST principle) | [01-styles](01-styles.md), [02-editing-process](02-editing-process.md) |
| What order should the approved-style array be in? | `ListApprovedStylesByBookOrder` output | [04-qa-workflow](04-qa-workflow.md) |
| Why was decision X made? | `rvw/Code_review YYYY-MM-DD.md` | [09-history](09-history.md) |
| What is the current synthesized state? | This guide | — |

## Pages

| # | File | Purpose | Status |
|---|------|---------|--------|
| 1 | [01-styles.md](01-styles.md) | Approved style taxonomy and current array | WIP — validated up to priority 36 |
| 2 | [02-editing-process.md](02-editing-process.md) | Routines mapped to editing steps | Mature |
| 3 | [03-inspection-tools.md](03-inspection-tools.md) | `basStyleInspector` reference | Mature |
| 4 | [04-qa-workflow.md](04-qa-workflow.md) | Book-order canonical priority workflow | Mature |
| 5 | [05-headers-footers.md](05-headers-footers.md) | Section / header / footer conventions | WIP |
| 6 | [06-i18n.md](06-i18n.md) | Localization considerations | Skeleton |
| 7 | [07-super-test-runs.md](07-super-test-runs.md) | Architectural QA supervisor | Placeholder — pending implementation |
| 8 | [08-publishing.md](08-publishing.md) | Producing the docx / PDF | Skeleton |
| 9 | [09-history.md](09-history.md) | Pointers into `rvw/` | Mature |

## Operative principles

- **WRIST** — Word ruler is the practical source of truth for indent
  measurements. Read indents off the ruler in the UI; encode the
  values into `RUN_TAXONOMY_STYLES` in code.
- **Book order = priority order** — the `approved` array reflects
  reading order through the document, not alphabetical or historical.
- **Progressive history** — `rvw/` files are dated, append-only
  snapshots. Never retroactively rewrite earlier sections.
- **ASCII only in VBA source** — `.bas` and `.cls` files use
  hyphen-minus, not em-dash. Markdown can use either.
- **Late binding** — all COM objects via `As Object` plus
  `CreateObject`. No project references added.
- **Identifier casing preserved** — never normalize VBA identifier
  case; the git commit normalizer relies on stability.

## Repo connection

The EDSG is tracked in the same repo as the code. Each EDSG update
commits alongside the code or style change that motivated it — the
commit message and the EDSG diff form a paired audit trail.
