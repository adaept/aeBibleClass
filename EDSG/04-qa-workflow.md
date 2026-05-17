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
mechanism (e.g., `BodyTextContinuation`, `AppendixTitle`,
`AppendixBody`) plus the deliberate `FargleBlargle` canary.

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

> **⚑ Final-state goal:** `RUN_TAXONOMY_STYLES` should ultimately map
> *every* approved style with a real (non-sentinel) expected spec, so
> any property drift on any approved style is caught immediately. The
> current 19-entry curated subset is transitional. See
> [01-styles § ⚑ Important — taxonomy audit final-state goal](01-styles.md#-important--taxonomy-audit-final-state-goal)
> for the full callout and progress framing.

> **⚑ Word `List Paragraph` numbering-engine bug:** the QA-checklist
> rule `BaseStyle = ""` is not stylistic — it dodges a long-known
> Word bug that hangs the application on Modify Style for any style
> inheriting from `List Paragraph` in large documents. See
> [10-list-paragraph-bug](10-list-paragraph-bug.md) for symptom,
> cause, common bad advice, and the all-VBA migration recipe.

## Current state — 2026-05-17 (latest)

- **`GetApprovedStyles()` array**: 52 entries. Of those, 48 are
  present in the document; 4 are tracking placeholders
  (`BodyTextContinuation`, `AppendixTitle`, `AppendixBody`, and
  the deliberate `FargleBlargle` canary).
- **`VerseText` is the live verse paragraph style** since
  2026-05-01 (priority 33). `BodyText` is now the residual
  non-verse paragraph style (front matter, chapter intros,
  chapter-end content).
- **Recent additions since the 2026-04-26 snapshot**:
  - `VerseText` (33) — verse paragraph style; replaced
    `BodyText` for verse content on 2026-05-01.
  - `BookHyperlink` (35) — one-form hyperlink character style
    added 2026-05-15; replaces direct use of the built-in
    `Hyperlink` style.
  - `BibleIndexList` (16), `AuthorBookSections` (24),
    `ParallelHeader` (49), `ParallelText` (50) — promoted
    during the same period.
  - `SpeakerLabel` (41), `BodyTextContinuation` (42),
    `AppendixTitle` (43), `AppendixBody` (44) — filled the
    former 38–41 gap.
- **Removals since the 2026-04-26 snapshot**:
  - `BodyTextIndent` — removed during the VerseText migration.
  - `AuthorQuote` — removed; front matter usage never
    finalized.
  - `BookIntro` — removed; decision deferred on define-and-
    promote vs close.
- **`Normal`** — priority 51 (was 47), last approved entry
  before the `FargleBlargle` canary. Operational role replaced
  by `BodyText` and `VerseText`; kept as the
  "pin-everything-else-above" anchor.
- **Pending re-validation**: most styles have been touched
  during the empty-paragraph discipline, hide-sweep, and
  LineSpacingRule prescriptive passes. Per-style state lives in
  `RUN_TAXONOMY_STYLES` (`basTEST_aeBibleConfig.bas`); the
  taxonomy audit umbrella tracks coverage progress.

### Reserved gaps

None as of 2026-05-17. The prior 38–41 gap was filled. Future
insertions take the next free priority.

## Headless caveat

`DumpAllApprovedStyles` shows an interactive `MsgBox` on orphan
detection. Not safe to chain into a non-interactive batch (e.g., a
future `SUPER_TEST_RUNS`). If that becomes a need, add a
`bSkipPrompt As Boolean` argument. Flagged YAGNI for now.
