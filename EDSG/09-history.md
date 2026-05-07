# 09 — History — pointers into rvw/

The EDSG synthesizes the *current* state. The `rvw/Code_review
YYYY-MM-DD.md` files preserve the *historical* state: dated,
append-only, no retroactive rewrites. When a decision in the EDSG
needs justification, the rvw/ files are where the rationale lives.

## How rvw/ works

- One file per major editing arc, dated.
- Each file appends section by section (`§ Foo - YYYY-MM-DD`).
- Earlier sections are never rewritten; corrections are made by
  appending a new section that supersedes.
- Status fields (PENDING / IN PROGRESS / DONE / DEFERRED) are
  updated in place; the surrounding analysis is not.

## Active and recent files

| File | Span | Topics |
|---|---|---|
| [`rvw/Code_review 2026-05-07.md`](../rvw/Code_review%202026-05-07.md) | 2026-05-07 → | Active carry-forward arc; open items only |
| [`rvw/Code_review 2026-05-06.md`](../rvw/Code_review%202026-05-06.md) | 2026-05-06 → 2026-05-07 | AuditVerseMarkerStructure CVM extension; four taxonomy QA-alignment rounds (BaseStyle = "" + bucket-1 promotions); ContentsRef tab-stop; Footnote Reference base-style rebase; character-style basing rule documented |
| [`rvw/Code_review 2026-04-30.md`](../rvw/Code_review%202026-04-30.md) | 2026-04-30 → 2026-05-06 | VerseText migration closure; Solomon -> Song of Songs sweep; orphan repair; EmphasisRed cleanup; AuditCharStyleUsage parameterization |
| [`rvw/Code_review 2026-04-25.md`](../rvw/Code_review%202026-04-25.md) | 2026-04-25 → 2026-04-30 | List Paragraph migration (Phases 0-6); WEB versification; spec promotions; Finding 5 ribbon focus; tab-stop infrastructure |
| [`rvw/Code_review 2026-04-21.md`](../rvw/Code_review%202026-04-21.md) | 2026-04-21 → 2026-04-24 | Style inspector module; QA workflow; orphan cleanup; WRIST principle; SUPER_TEST_RUNS proposal |
| (earlier dated reviews) | | Older arcs — see `rvw/` directory listing |

The full list of dated reviews is the `rvw/` directory itself.

## Decision archaeology — common queries

When you need to know *why* something is the way it is, search rvw/
first.

| Question | Where to look |
|---|---|
| Why does `AutomaticallyUpdate = False` for every approved style? | `rvw/Code_review 2026-04-21.md` § QA checklist |
| Why is `BaseStyle = ""` instead of `Normal`? | `rvw/Code_review 2026-04-21.md` § Style inspector |
| Why are header/footer styles in `Headers(1)` not `Headers(2)`? | `rvw/Code_review 2026-04-21.md` § Paragraph-iteration fix; § DumpHeaderFooterStyles diagnostic |
| Why does `ListApprovedStylesByBookOrder` skip story types 6–11? | `rvw/Code_review 2026-04-21.md` § Skip header/footer types in StoryRanges walk |
| Why is page 1 used as fallback for header/footer hits? | `rvw/Code_review 2026-04-21.md` § Header/Footer page-1 fallback |
| Why does the `approved` array order match book order? | `rvw/Code_review 2026-04-21.md` § QA workflow goal and current state |
| What is WRIST and where did it come from? | `rvw/Code_review 2026-04-21.md` § Historical reference - "ruler as source of truth for indents" |
| Why is `SUPER_TEST_RUNS` deferred? | `rvw/Code_review 2026-04-21.md` § 6 |
| Why was AuthorRef renamed to AuthBookRef? | `rvw/Code_review 2026-04-21.md` § (rename note) |

## Significant commits

When a decision has a single commit-anchor, prefer linking the
commit. Notable:

| Commit | Subject | Why it matters |
|---|---|---|
| `200ead8` | Taxonomy test is working | Established `RUN_TAXONOMY_STYLES` as source of truth for style property values |
| `27136bb` | Word ruler is source of truth for indents (WRIST) | Recorded the WRIST principle |
| `7b8cef8` | FIXED — QA workflow goal statement | Captured "book order = canonical priority" |
| `8cbcabb` | FIXED — Create DumpStyleProperties | First entry of `basStyleInspector.bas` |
| `dac2c12` | FIXED — DumpAllApprovedStyles | Bulk audit added |
| `04695ed` | FIXED — Move style reports | Created `rpt/Styles/` |

Commit hashes are stable; subjects may not match exactly if the
history is later edited (rebase / amend) — verify with `git log`.

## When to write rvw/ vs EDSG

Both, paired:

- **rvw/** — the decision, the rationale, the alternatives
  considered, the cost/benefit. Append-only. Future-you reads this
  to understand "why."
- **EDSG** — the current state and the operational instructions.
  Updated in place. New readers consult this to understand "how."

A typical change touches both: the rvw/ file gets a new section
documenting the decision; the relevant EDSG page gets updated to
reflect the new state.
