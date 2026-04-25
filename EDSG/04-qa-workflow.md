# 04 — QA workflow

## The principle

**The book-order output IS the canonical priority sequence.**

`ListApprovedStylesByBookOrder` walks the document and reports each
approved style at the page where it first appears. That output —
sorted `(Page ascending, Priority ascending)` — defines the order
that the `approved` array in `src/basTEST_aeBibleConfig.bas` should
follow.

Not a hint. Not a guideline. The array is wrong if it disagrees with
the output.

## Why book order

Three reasons:

1. **Onboarding navigation** — a translator or editor reading the
   array can jump to the page where each style first matters and see
   it in context.
2. **Review structure** — a printed proof of the document can be
   walked top-to-bottom alongside the array, line by line.
3. **Mechanical synchronization** — array generation can become
   automatic from book order; deviations are bugs, not preferences.

## The cycle

Each pass is the same five steps:

1. **Walk** a chunk of pages (e.g., pages 12–25 next).
2. **Audit**: `ListApprovedStylesByBookOrder True`.
3. **Compare** the output against the current `approved` array. If
   they diverge:
   - Update the array — reorder existing entries, add any new styles
     encountered along the way.
   - For new styles, also follow Stage 1 of
     [02-editing-process](02-editing-process.md) — define them in
     code if they don't yet exist.
4. **Repromote**: `WordEditingConfig` resets every paragraph /
   character style to `Priority = 99`, then promotes the array in
   order.
5. **Re-verify**: `ListApprovedStylesByBookOrder` again. The two
   should now align. If not, return to step 3.

`DumpAllApprovedStyles` after the cycle catches orphan dump files left
by any rename.

## What "approved" means in code

Defined operationally, not declaratively:

```
approved style := paragraph or character style with Priority <> 99
```

`PromoteApprovedStyles` (in `src/basTEST_aeBibleConfig.bas`) is the
mechanism. It first sets every paragraph/character style to priority
99, then assigns 1, 2, 3, ... to the names in the `approved` array
(in array order). Names in the array that are not present in the
document are reported as missing — preserved as a tracking
mechanism (e.g., `BodyTextContinuation`, `BookIntro`,
`AppendixTitle`, `AppendixBody`) plus the deliberate `FargleBlargle`
canary.

After running `WordEditingConfig`:

- A style with `Priority < 99` is approved (and its number = its
  position in the array).
- Anything else is not approved (whether built-in like `Default
  Paragraph Font`, or legacy like `BodyTextTopLineCPBB` before it
  was added to the array).

## QA checklist (per-style)

Apply to every approved style. See [01-styles](01-styles.md) §
QA checklist for the table. Four checks:

1. `BaseStyle = ""` (no inheritance)
2. `AutomaticallyUpdate = False` (paragraph styles only)
3. `QuickStyle = False` (no gallery clutter)
4. `LineSpacingRule = wdLineSpaceSingle` (paragraph styles only)

`DumpStyleProperties` puts these in front of you per style;
`DumpAllApprovedStyles` does it in bulk. Documented exceptions are
recorded inline in [01-styles](01-styles.md).

## Current state — 2026-04-25

- Pages 1–11 walked; array aligned through approximately priority 16.
- Priorities 17+ inherited from earlier ad-hoc ordering; pending
  re-validation as the page walk proceeds.
- `Normal` (priority 14, page 6) — appears used in the main body but
  was previously assumed replaced by `BodyText`. Decision pending:
  keep as anchor or prune.
- `BodyTextIndent` (priority 18) — `[not used]`. Decision pending:
  keep for planned use, or prune.
- `AuthorQuote` (priority 36) — `[not used]`. Deferred to front
  matter work.

## Headless caveat

`DumpAllApprovedStyles` shows an interactive `MsgBox` on orphan
detection. Not safe to chain into a non-interactive batch (e.g., a
future `SUPER_TEST_RUNS`). If that becomes a need, add a
`bSkipPrompt As Boolean` argument. Flagged YAGNI for now.
