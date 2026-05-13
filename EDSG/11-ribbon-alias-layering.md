# Ribbon Book-Combo Alias Layering

How the ribbon book-combo resolves what the user types when they
press Tab.

## The two resolution layers

There are two independent name-resolution systems in the project.
For a long time they were not connected; the ribbon combo bypassed
the alias map entirely. This page documents the contract after the
2026-05-11 fix.

| Layer | Code | Authority | Used by |
|---|---|---|---|
| 1. Alias map | `aeBibleCitationClass.ResolveAlias` -> `aliasMap` | Canonical book name for any registered alias (e.g. "JB" -> "Job") | SBL citation parser, `basVerseStructureAudit`, ribbon combo (since 2026-05-11) |
| 2. H1 substring | `headingData` substring `Like` scan in `aeRibbonClass.OnBookChanged` | Best-effort partial typing | Ribbon combo for partials that are not formal aliases |

The ribbon combo now tries layer 1 first; on `Unknown book alias`
it falls back to layer 2.

## Why this matters - the first-substring-wins bug

The legacy single-layer scan walked books 1..66 and exited on the
first `Like "*pattern*"` match. That produced order-dependent
collisions whenever the typed text was a substring of an
earlier book's name:

| Typed | Legacy bound to | Intended |
|---|---|---|
| `NAH` | `Jonah` (32) - J-O-**N-A-H** | `Nahum` (34) |
| `JB` | (no match) | `Job` |
| `PRV` | (no match) | `Proverbs` |
| `SG` | (no match) | `Song of Songs` |
| `1 KGS` | (no match) | `1 Kings` |

Aliases that are not substrings of the canonical H1 simply did not
work, and aliases that *were* substrings of an earlier book bound
to the wrong book. Neither failure was documented anywhere.

## The fix

`aeRibbonClass.OnBookChanged` consults `ResolveAlias` first:

1. Try `aeBibleCitationClass.ResolveAlias(text)`.
2. If it succeeds, match the returned **canonical name** against
   `headingData` with `Like "*canonical*"`. `aliasMap` is now
   authoritative for any alias it recognises.
3. If `ResolveAlias` raises `Unknown book alias` (or layer 1 finds
   no `headingData` match), fall back to the legacy substring
   `Like "*pattern*"` scan against `headingData`, where `pattern`
   is the user's typed text after `NormalizeBookInput`. This
   preserves partial-typing UX (`Genesi` + Tab still finds Genesis).

The two layers are wired sequentially in `OnBookChanged`: layer 1
runs first, layer 2 runs only if layer 1 leaves `matchFound = False`.
Inputs shorter than two characters short-circuit before either
layer runs (single-letter prohibition matches the alias map's own
rule against single-letter keys).

## Search expansion - generic ribbon lookup

The ribbon lookup is intentionally generic now: anything in the
alias map resolves, and anything the alias map does not know still
gets a fair shot via the layer-2 `Like` scan. The **breadth of the
map** is therefore what determines how forgiving the ribbon feels
to a typist.

The authoritative inventory of accepted abbreviations is
[`md/BibleAbbreviationList.md`](../md/BibleAbbreviationList.md).
It is a unified, deduplicated, publisher-grade list that merges:

- Traditional English publishing (KJV-lineage, commentaries,
  devotional works).
- Standard church / academic abbreviations.
- Digital shortest-form systems (Logos-style, concordances,
  BibleStudyTools).

Every entry in that document is mirrored in
`aeBibleCitationClass.GetBookAliasMap` (UPPERCASE, ASCII, two or
more characters - the single-letter prohibition still holds).
Closed-up no-space forms (`1SA`, `2PE`, `1JO`) live alongside the
spaced forms (`1 SA`, `2 PE`, `1 JN`) so the ribbon resolves either
convention without a normalisation pass.

**End-to-end resolution for a typed token:**

1. `NormalizeBookInput(text)` uppercases the typed text and
   collapses whitespace. Result is `pattern`.
2. `ResolveAlias(text)` looks up `pattern` (after its own
   normalisation) in `aliasMap`. Hit -> returns the canonical book
   name (`"Nahum"`, `"1 John"`, ...). Miss -> raises
   `Unknown book alias`.
3. **Layer 1 - canonical Like scan.** If layer 1 returned a
   canonical name, walk `headingData(1..66, 0)` and accept the
   first H1 that matches `Like "*canonical*"`. Canonical names are
   designed to not be substrings of each other in book-order, so
   this is collision-free for the standard 66-book set.
4. **Layer 2 - typed-pattern Like scan (fallback).** If layer 1
   produced no match (either `ResolveAlias` raised, or the
   canonical name was not present in `headingData`), walk
   `headingData` again with `Like "*pattern*"` against the
   normalised typed text. This catches partial typing
   (`"Genesi"`, `"Reve"`, `"Eccles"`) and any abbreviation a future
   editor types before the map is extended.

**When to extend the map vs. rely on the fallback:** prefer
extending the map. The fallback is a safety net for partials, not
a documentation surface. If a real-world abbreviation needs to
work, add it to `md/BibleAbbreviationList.md` and mirror it in
`GetBookAliasMap`. The ribbon will pick it up on the next
`aliasMap` build (lazy, first-use, cached).

## Implications for editors and developers

- **Add a new abbreviation:** put it in `aliasMap` (in
  `aeBibleCitationClass.cls` `GetBookAliasMap`). The ribbon will
  pick it up automatically. No ribbon code change required.
- **Order of entries in `aliasMap` is irrelevant.** It is a
  `Scripting.Dictionary` keyed by hash. Group entries by book for
  human readability only.
- **No regeneration is needed.** The map is built lazily on
  first use and cached.
- **Partial typing still works** for substrings that aren't in
  the map, via the layer-2 fallback.

## Caveat - canonical-substring collisions

Layer 1 still uses `Like "*canonical*"` against `headingData`, so
if a canonical name is a substring of an earlier H1, the earlier
H1 wins. The current 66-book table has no such collisions
(e.g. `"John"` matches the standalone H1 "John" at index 43 before
the `"1 John"`/`"2 John"`/`"3 John"` H1s; `"1 John"` only matches
"1 John" itself, etc.). New book additions or H1 renames should
be reviewed against this constraint.

## Cross-references

- `src\aeRibbonClass.cls` `OnBookChanged` - the two-layer logic.
- `src\aeBibleCitationClass.cls` `ResolveAlias`, `GetBookAliasMap` -
  layer 1 implementation.
- [`md/BibleAbbreviationList.md`](../md/BibleAbbreviationList.md) -
  authoritative abbreviation inventory mirrored into the alias map.
- `rvw\Code_review 2026-05-11.md` - original bug report and
  two-layer fix decision.
- `rvw\Code_review 2026-05-12.md` (2026-05-13 entry) - alias-map
  expansion from `md/BibleAbbreviationList.md`.
