# Code review - 2026-05-08 carry-forward

This file opens a fresh review arc on 2026-05-08. The previous arc
[`rvw/Code_review 2026-05-07.md`](Code_review%202026-05-07.md) is now
**closed for new work**; that file remains the authoritative dated
history for everything between 2026-05-07 and 2026-05-08, including:

- `BodyTextIndent` removal (taxonomy + `DefineBodyTextIndentStyle` Sub).
- `BookIntro` removal (taxonomy + `DefineBookIntroStyle` and
  `ApplyBookIntroAfterDatAuthRef` tooling).
- 8 styles dumped via `DumpStyleProperties` and promoted: 3 paragraph
  styles to bucket 1 (`SpeakerLabel`, `AuthorSectionHead`,
  `AuthorBookSections`) and 5 character styles to bucket 2 (`Selah`,
  `EmphasisBlack`, `EmphasisRed`, `Words of Jesus`, `AuthorQuote`).
- Taxonomy header realignment to `45 styles + 6 tab-stop specs = 51
  checks` (`47 PASS / 4 FAIL` predicted).
- New "Define colors used in the docx" task registered.
- **Soft-hyphen sweep** built end-to-end: design discussion, JIS B5
  geometry capture from `JUDE - Sample.docm`, mirrored-margin handling,
  per-page section-aware bound resolution, calibration helper,
  diagnostic, two-pass worker (Active / Stray / OutsideBody), driver,
  Yes / No / Cancel prompt with selection scroll-into-view, dry-run
  gate, append-mode CSV + log under `rpt\`. Validated on pages
  21 (Recto), 22 (Verso), 913 (anomalous section 135), and 911
  (3 Stray finds confirmed). Confirmed working in production.

Items below are the **open** carry-forward set. Each entry links back
to the section in the prior review where the full rationale lives.

## Open carry-forward

### 1. Define colors used in the docx (active task)

Use of Word Themes / Theme Colors is **not allowed anywhere** in this
document. Every color reference must be expressed as an explicit RGB /
`wdColor*` constant captured as part of the descriptive style baseline.
Originated in `rvw/Code_review 2026-05-07.md` "NEW TASK: Define colors
used in the docx" section.

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
  routine) that fails any style whose color resolves through a theme
  rather than an explicit literal.

### 2. Character-style audit extension for `AuditOneStyle`

`AuditOneStyle` currently checks paragraph properties only. Character
styles park in bucket 2 with font / size existence only. Extension
needed to check **Bold / Italic / Color** for character styles.

Once landed, promote from bucket 2 to bucket 1:

- `Footnote Reference`
- `Selah`
- `EmphasisBlack`
- `EmphasisRed`
- `Words of Jesus`
- `AuthorQuote`

This unblocks completion of the dump-and-promote arc that started
2026-05-07.

### 3. Soft-hyphen sweep - production use and follow-ups

Production sweep is **operational**. Routines live in
`src\basWordRepairRunner.bas`:

- `SoftHyphen_DiagnoseLayout` - section-by-section PageSetup dump to
  `rpt\SoftHyphen_Layout.log`, anomaly detection.
- `SoftHyphen_CalibrateColumns(pageNum)` - per-page calibration with
  Recto / Verso classification, append output to
  `rpt\SoftHyphenCalibration.csv` and `.log`.
- `SoftHyphenSweep_ByColumnContext_SinglePage(pageNum, mode, ByRef ...)`
  - two-pass worker with Yes / No / Cancel prompt.
- `RunSoftHyphenSweep_Across_Pages_From(startPage, pageCount, dryRun)`
  - driver, all args required (no `Optional`).

#### 3a. Anomalous 2-column sections to verify in production

`SoftHyphen_DiagnoseLayout` flagged **2 anomalies** (sections with
2-col geometry deviating from standard `196.9 / 14.4 / 196.9` by more
than 0.5 pt). `GetColumnBoundsForPage` reads the section's own
`PageSetup` so the worker should handle these correctly, but a real
production sweep that touches those sections is the only confirmation.

Known anomaly: section 135 starting page 913 -
`Col1.Width=186.1  SpaceAfter=36.0  Col2.Width=186.1`. Calibration
on page 913 returned 7 finds, all Active, 0 OutsideBody - bounds were
correctly resolved against the section's own geometry, not the
neighbouring 196.9 sections. Watch for the second anomaly in
`rpt\SoftHyphen_Layout.log` "-- Anomalies --" block when a
production sweep covers it.

#### 3b. `Footnote Text` soft-hyphen sweep (possible sister routine)

Out of scope for the body sweep that just landed (Q3 / 2026-05-07).
If a font-change pass on `Footnote Text` leaves stray soft hyphens in
footnote bodies, build `SoftHyphenSweep_FootnotesOnly` as a sister
routine rather than overloading the body sweep with a flag. Headers /
footers explicitly excluded - they will not contain soft hyphens.

#### 3c. `SHA_ReplaceHard` future i18n consideration

Currently the only action on a Stray find is delete. If a future
non-English edition adopts soft hyphens as semantic compound-break
markers (German, Dutch, and some Slavic-language typesetting
conventions occasionally do), revisit `SHA_ReplaceHard` to replace
the soft hyphen with a hard hyphen-minus rather than deleting. Until
then the action set stays binary - delete or preserve.

## Pointer back to the closed arc

Full dated history of the work that produced this carry-forward state
is in [`rvw/Code_review 2026-05-07.md`](Code_review%202026-05-07.md).
That file includes:

- The complete `BodyTextIndent` and `BookIntro` removal records,
  including the tooling-routine deletion decisions.
- Four bucket-promotion records for the 8 dumped styles plus the
  taxonomy header realignment to 51 checks.
- The "Define colors used in the docx" task creation rationale.
- The full soft-hyphen design discussion (Q1-Q6 with resolutions),
  JIS B5 geometry derivation from `JUDE - Sample.docm`, the
  mirrored-margin correction, the diagnostic anomaly-detection
  addition, and the two-pass worker design.

Anything in this 2026-05-08 file should reference back to that arc
for the *why*; this file holds only the **what is still open**.
