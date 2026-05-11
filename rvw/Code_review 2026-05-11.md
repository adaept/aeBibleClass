# Code review - 2026-05-11 carry-forward

This file opens a fresh review arc on 2026-05-11. The previous arc
[`rvw/Code_review 2026-05-08.md`](Code_review%202026-05-08.md) is now
**closed for new work**; that file remains the authoritative dated
history for everything between 2026-05-08 and 2026-05-11, including:

- **Row character-count & pitch diagnostic** built end-to-end
  (Phase A survey, Phase C histogram, Phase B interactive
  navigator). Operating doc at
  [`md/Row_Pitch_Diagnostic.md`](../md/Row_Pitch_Diagnostic.md).
  Validated on page 522 with the closed-loop Survey -> Histogram
  -> manual-fix -> re-run cycle.
- **Character-style hygiene pass**: new `AuditCharStyleBases` and
  `ScanCharStyleApplications` diagnostics in `basStyleInspector.bas`;
  9 cruft styles deleted; deletable cruft 8 -> 2 remaining;
  decision rule and palette categorisation captured.
- **Taxonomy refresh**: `TitleOnePage` font baselined to
  `Liberation Serif`; three new bucket-1 paragraph styles added
  (`BibleIndexList`, `ParallelHeader`, `ParallelText`); `AuthorQuote`
  removed; `Title` / `BibleIndex` / `AuthorBookSections` drift
  re-baselined. Final state **52 PASS / 3 FAIL** (the 3 expected
  not-yet-created styles).
- **Normalizer entries** added for `NextChar` / `NextPara` /
  `NextLine` loop labels so the VBA IDE preserves their casing.

Items below are the **open** carry-forward set, ordered with the
most effective work at the top of the list. "Most effective" =
highest unlock-to-effort ratio - work that removes blockers for
multiple downstream items, or that closes a category of risk
rather than a single instance.

## Open carry-forward (priority order)

### 1. Extend AuditOneStyle for character-style properties (HIGH)

`AuditOneStyle` currently checks paragraph properties only.
Character styles park in bucket 2 with font / size existence only.
Extension needed to check **Bold / Italic / Color** for character
styles.

**Why this is top of the list:** a single change in one routine
unlocks **promotion of 7 character styles** to bucket 1, and is a
hard prerequisite for closing the colour-audit task (item 2 below).
Most leveraged work available.

Once landed, promote from bucket 2 to bucket 1:

- `Footnote Reference` (built-in, Applied)
- `Selah` (custom, Applied)
- `EmphasisBlack` (custom, Applied)
- `EmphasisRed` (custom, Applied)
- `Words of Jesus` (custom, Applied)
- `Chapter Verse marker` (custom, Applied) - **add per
  `Code_review 2026-05-08.md` 6h**
- `Verse marker` (custom, Applied) - **add per
  `Code_review 2026-05-08.md` 6h**

Originated in `rvw/Code_review 2026-05-07.md`; updated
promotion list per 2026-05-08 6h.

### 2. Define colors used in the docx (HIGH)

Use of Word Themes / Theme Colors is **not allowed anywhere** in
this document. Every color reference must be expressed as an
explicit RGB / `wdColor*` constant captured as part of the
descriptive style baseline.

**Why high:** doc-wide audit guarantee. Pairs naturally with
item 1 because the character-style audit extension is the place
the colour check will live.

Scope:

- Enumerate every style and direct-formatting site that carries a
  non-default color (paragraph styles, character styles, run-level
  overrides, table / shading, ribbon-driven highlights).
- Confirm intended color for each site and capture as a long
  RGB / BGR literal in the style spec (no `wdColorAutomatic`, no
  theme-color tints, no `wdColorByPalette`).
- Lock in the three known character-style colors carried from the
  prior arc:
  - `Footnote Reference` - dump shows `Font.Color = 16711680`
    (`&H00FF0000`, BGR = blue). User to confirm intent (red is the
    common convention).
  - `Chapter Verse marker` - orange (`42495`).
  - `Verse marker` - medium green (`7915600`).
- Add a taxonomy check (extension of `AuditOneStyle` or a sibling
  routine) that fails any style whose color resolves through a
  theme rather than an explicit literal.

Originated in `rvw/Code_review 2026-05-07.md` "NEW TASK: Define
colors used in the docx" section.

### 3. Re-base remaining character styles to Default Paragraph Font (MEDIUM)

`AuditCharStyleBases` returned 13 offenders on first run (after
exclusion of `Default Paragraph Font` itself). 9 of those were
deleted as palette cruft (per 6h); the remaining offenders need
re-basing rather than deletion.

**Action:**

1. Re-run `?AuditCharStyleBases` to get the current offender list
   (post-deletion, post-`Page Number -> Footer Char` resolution).
2. For each remaining offender, set
   `ActiveDocument.Styles("<name>").BaseStyle = "Default Paragraph Font"`.
3. Re-run `?AuditCharStyleBases`; expect **0**.

Special case: `Page Number -> Footer Char` is chained inheritance
through the `Footer` paragraph style's auto-generated linked
character style. Repoint directly to `Default Paragraph Font`.

Originated in `rvw/Code_review 2026-05-08.md` 6b / 6g.

### 4. Delete `Normal text` custom character style (MEDIUM)

`Normal text` is the last remaining custom-and-Unapplied character
style after the 9-style cleanup. Generic name, not in any
promotion list, not carried by any run. Strong delete candidate.

**Action:**

1. Final visual check that no run carries it (or trust
   `?ScanCharStyleApplications`, which already performed exactly
   this check on 2026-05-10).
2. `ActiveDocument.Styles("Normal text").Delete`.
3. Re-run `?ScanCharStyleApplications`; expect Custom Unapplied
   count = 0.

Originated in `rvw/Code_review 2026-05-08.md` 6h.

### 5. Apply Row Pitch Diagnostic to two un-hyphenated 10-page ranges (MEDIUM)

The diagnostic toolchain (Phase A / B / C) is production-ready and
validated on a single page. The next step is the **actual
use case** - identify two un-hyphenated 10-page ranges, run the
full survey -> histogram -> review cycle per
[`md/Row_Pitch_Diagnostic.md`](../md/Row_Pitch_Diagnostic.md).

**Why medium not high:** tooling is in place, this is just doing
the work. High-value for output quality but does not unblock
other items.

Expected outcome:

- Stable per-side median pitch (Left and Right within ~0.1 pt).
- A clear suspect tail (Pitch > median + 1.0 pt).
- Reduced suspect count after the manual-hyphen pass; medians
  tighten further.

Originated in `rvw/Code_review 2026-05-08.md` 4d.

### 6. Verify anomalous 2-column sections in production (LOW-MEDIUM)

`SoftHyphen_DiagnoseLayout` flagged **2 anomalies** (sections with
2-col geometry deviating from standard `196.9 / 14.4 / 196.9` by
more than 0.5 pt). `GetColumnBoundsForPage` reads the section's
own `PageSetup` so the worker should handle these correctly, but
production runs that actually touch those sections are the only
confirmation.

Known anomaly: section 135 starting page 913 -
`Col1.Width=186.1  SpaceAfter=36.0  Col2.Width=186.1`.
Calibration on page 913 returned 7 finds, all Active, 0
OutsideBody - bounds correctly resolved. Watch for the second
anomaly in `rpt\SoftHyphen_Layout.log` "-- Anomalies --" block
when a production sweep covers it.

**Why low-medium:** opportunistic - resolves naturally as
production sweeps progress through the document.

Originated in `rvw/Code_review 2026-05-08.md` 3a.

### 7. Optional --clear helper for RowCharCount survey driver (LOW)

`rpt\RowCharCount.csv` is opened with `For Append`. Re-running
the survey on the same page without first clearing the CSV
duplicates that page's rows in the histogram input. Currently the
user must clear manually between same-range re-runs.

Possible follow-up: add a `--clear` parameter (or a sibling
`RunRowCharCountSurvey_Across_Pages_From_Fresh`) that truncates
the CSV first. Or detect overlap with already-present pages and
refuse / prompt.

**Why low:** the manual workflow is documented in
[`md/Row_Pitch_Diagnostic.md`](../md/Row_Pitch_Diagnostic.md) and
works fine. Pure quality-of-life.

Originated in `rvw/Code_review 2026-05-08.md` 4g (append-mode
caveat).

### 8. SoftHyphenSweep_FootnotesOnly sister routine (DEFERRED)

Out of scope for the body sweep that landed 2026-05-07. If a
font-change pass on `Footnote Text` leaves stray soft hyphens in
footnote bodies, build `SoftHyphenSweep_FootnotesOnly` as a
sister routine rather than overloading the body sweep with a
flag. Headers / footers explicitly excluded - they will not
contain soft hyphens.

**Why deferred:** no triggering need yet.

Originated in `rvw/Code_review 2026-05-08.md` 3b.

### 9. SHA_ReplaceHard i18n consideration (FUTURE)

Currently the only action on a Stray soft-hyphen find is delete.
If a future non-English edition adopts soft hyphens as semantic
compound-break markers (German, Dutch, and some Slavic-language
typesetting conventions occasionally do), revisit `SHA_ReplaceHard`
to replace the soft hyphen with a hard hyphen-minus rather than
deleting. Until then the action set stays binary - delete or
preserve.

**Why future:** no planned i18n editions on the immediate horizon.

Originated in `rvw/Code_review 2026-05-08.md` 3c.

## Pointer back to the closed arc

Full dated history of the work that produced this carry-forward
state is in [`rvw/Code_review 2026-05-08.md`](Code_review%202026-05-08.md).
That file includes:

- The full Row Pitch Diagnostic design discussion, phase-by-phase
  implementation, page-522 closed-loop validation.
- The `AuditCharStyleBases` and `ScanCharStyleApplications`
  design, exclusion rationale, scan results, palette
  categorisation (categories A-E), and built-in vs custom
  distinction.
- The taxonomy edits (TitleOnePage / Title / BibleIndex /
  AuthorBookSections re-baselining; BibleIndexList /
  ParallelHeader / ParallelText additions; AuthorQuote removal;
  count header rebalancing 51 -> 55).
- The `RUN_TAXONOMY_STYLES` confirmation 52 PASS / 3 FAIL.

Anything in this 2026-05-11 file should reference back to that arc
for the *why*; this file holds only the **what is still open**.

## Status updates (append-only)

### 2026-05-11 - Item 1 CLOSED

`AuditOneStyle` extended with optional `vExpColor` parameter
(sentinel `-2` = skip), appended after `sExpBaseStyle` so existing
positional callers were unaffected. Color check block mirrors the
Bold / Italic pattern in `src\basTEST_aeBibleConfig.bas`.

All 9 character / paragraph styles from the promotion list have been
moved from bucket 2 to bucket 1 with descriptive specs captured via
`DumpStyleProperties`:

- `TheHeaders`, `TheFooters` (paragraph; Noto Sans 9pt; Color
  `-16777216` = wdColorAutomatic - flagged for item 2)
- `Footnote Reference` (Carlito 9pt Bold; Color `16711680` BGR blue,
  intent confirmed)
- `Selah` (Carlito 9pt; Color `-16777216` - flagged for item 2)
- `EmphasisBlack` (Carlito 9pt Bold; Color `-16777216` - flagged
  for item 2)
- `EmphasisRed` (Carlito 9pt Bold; Color `128` BGR dark-red)
- `Words of Jesus` (Carlito 9pt; Color `128`; BaseStyle "")
- `Chapter Verse marker` (Noto Sans 5pt Bold; Color `42495` orange;
  added per 2026-05-08 6h)
- `Verse marker` (Noto Sans 8pt Bold; Color `7915600` green; added
  per 2026-05-08 6h)

`TheFooters` tab stop added to `AuditStyleTabs` (1 stop at 7.2 pt,
Left, Spaces). Existence-verified bucket is now empty.

Header doc-block recounted: **49 style audits + 9 tab-stop audits =
58 checks total**; 46 fully specified / 0 existence-verified / 3
not-yet-created.

Verified post-7-style-promotion run: `RUN_TAXONOMY_STYLES` = 53 PASS
/ 3 FAIL. Expected after the two `Chapter Verse marker` / `Verse
marker` additions: 55 PASS / 3 FAIL (the 3 expected not-yet-created
`Define*` styles).

**Item 2 hand-off:** four styles carry `wdColorAutomatic` (`-16777216`)
as their descriptive baseline - `TheHeaders`, `TheFooters`, `Selah`,
`EmphasisBlack`. Item 2's colour-literal pass converts these to
explicit RGB / BGR literals. The audit harness is now in place to
enforce the post-conversion values.
