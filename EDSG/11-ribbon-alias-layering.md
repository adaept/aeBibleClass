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
   `headingData`. `aliasMap` is now authoritative for any alias
   it recognises.
3. If `ResolveAlias` raises `Unknown book alias`, fall back to the
   legacy substring `Like` scan against `headingData`. This
   preserves partial-typing UX (`Genesi` + Tab still finds Genesis).

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
- `rvw\Code_review 2026-05-11.md` - bug report and fix decision.
