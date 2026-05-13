# Code review - 2026-05-12 carry-forward

This file opens a fresh review arc on 2026-05-12. The previous arc
[`rvw/Code_review 2026-05-11.md`](Code_review%202026-05-11.md) is now
**closed for new work**; that file remains the authoritative dated
history for everything between 2026-05-11 and 2026-05-12, including:

- **`AuditOneStyle` extended for character-style colour** (item 1 of
  the 2026-05-11 arc CLOSED). Seven character styles and two paragraph
  styles promoted bucket-2 -> bucket-1 with descriptive specs;
  `RUN_TAXONOMY_STYLES` 53 PASS / 3 FAIL after the promotion run.
- **Ribbon book-combo alias bug FIXED.** `aeRibbonClass.OnBookChanged`
  now consults `ResolveAlias` first; the legacy substring scan stays as
  fallback. New EDSG page
  [`EDSG/11-ribbon-alias-layering.md`](../EDSG/11-ribbon-alias-layering.md)
  captures the two-layer contract.
- **aeRibbon production export gateway established.** New `aeRibbon/`
  directory, new `py/ribbon_export_trim.py`, new
  `md/aeProductionRibbonPlan.md`. Six plan decisions approved including
  the dotm/docx production split. No `src/` changes; dev `.docm` files
  do not need re-import.

Items below are the **open** carry-forward set, ordered by
unlock-to-effort ratio - work that removes blockers for multiple
downstream items, or that closes a category of risk, at the top.

## Open carry-forward (priority order)

### 1. Run aeRibbon Gates G1-G8 and ship v1.0.0 (HIGH)

The production export gateway is in place; nothing has been built or
gated yet. This is the **next active release-track item** and gates the
hand-off to the author for comments-only review.

**Why high:** every other ribbon-side improvement (signing,
auto-docx-from-docm, ribbon UX iteration) sits behind a first
successful gated build. Also the highest-leverage validation of the
trim script: any false drop will surface in G6 (compile) or G8
(navigation smoke).

**Action:**

1. Build `aeRibbon/template/aeRibbon.dotm` per `aeRibbon/BUILD.md`.
   - Inject `aeRibbon/template/customUI14.xml` via
     `wsl python3 py/inject_ribbon.py`.
   - Import the 5 files from `aeRibbon/src/` into the template VBA
     project.
   - Set `RIBBON_VERSION` constant + custom property
     `aeRibbonVersion` to match `aeRibbon/VERSION` (`1.0.0+bc71416`).
   - Debug -> Compile VBAProject: must be zero errors.
2. Editor/Developer produces the production Bible `.docx` per
   `BUILD.md` "Producing the production Bible `.docx`" (manual
   File -> Save As `.docx` from the dev `.docm` - Option 1).
3. Run Gates G1-G8 from `aeRibbon/QA_CHECKLIST.md`. Record results in
   `aeRibbon/releases/1.0.0+bc71416/BUILD_RECORD.txt`.
4. Append a row to `aeRibbon/RELEASES.md`. `git tag v1.0.0+bc71416`.

**Expected blockers / what to watch for:**

- The trim script's call-graph is **conservative-overinclusive** by
  design (token-level identifier match, case-insensitive). False
  positives are harmless; **false drops** would surface as missing-Sub
  compile errors in G6. If G6 fails, identify the dropped routine and
  add it as an explicit root (or fix its caller to keep it reachable)
  rather than reverting to manual cherry-pick.
- VBA lifecycle hooks (`Class_Initialize`, `Class_Terminate`,
  `AutoExec`) are now always preserved if defined; verify this still
  holds after any future edit to `py/ribbon_export_trim.py`.
- G8 must show **no macro-security warning on docx open** - this is
  the architectural claim of the dotm/docx split. If the warning
  appears, the docx is not actually code-free (likely a Save-As mode
  selection error).

Originated 2026-05-12 with the gateway commits `bc71416` + `70bcff3`.

### 2. Define colors used in the docx (HIGH)

Carried forward from `rvw/Code_review 2026-05-11.md` item 2. Now has
**concrete inputs** from the 2026-05-11 hand-off: four styles carry
`wdColorAutomatic` (`-16777216`) as their descriptive baseline and
require explicit-literal conversion - `TheHeaders`, `TheFooters`,
`Selah`, `EmphasisBlack`. Three other character-style colours already
captured: `Footnote Reference` (`16711680` BGR blue, confirmed),
`Chapter Verse marker` (`42495` orange), `Verse marker` (`7915600`
green).

Use of Word Themes / Theme Colors is **not allowed anywhere**. Every
color reference must be an explicit RGB / `wdColor*` constant captured
in the descriptive style baseline.

**Action:**

- Enumerate every style and direct-formatting site that carries a
  non-default color (paragraph styles, character styles, run-level
  overrides, table / shading, ribbon-driven highlights).
- Convert the four `wdColorAutomatic` baselines to explicit RGB / BGR
  literals after confirming intent.
- Add a taxonomy check (extension of `AuditOneStyle` or sibling
  routine) that fails any style whose color resolves through a theme
  rather than an explicit literal.

Originated `rvw/Code_review 2026-05-07.md`; promoted to active by the
2026-05-11 Item 1 hand-off.

### 3. Re-base remaining character styles to Default Paragraph Font (MEDIUM)

Carried forward from `rvw/Code_review 2026-05-11.md` item 3. No change
in status this arc.

**Action:**

1. Re-run `?AuditCharStyleBases` to get the current offender list.
2. For each offender, set
   `ActiveDocument.Styles("<name>").BaseStyle = "Default Paragraph Font"`.
3. Re-run; expect **0**.

Special case still pending: `Page Number -> Footer Char` chained
inheritance - repoint directly to `Default Paragraph Font`.

Originated `rvw/Code_review 2026-05-08.md` 6b / 6g.

### 4. Delete `Normal text` custom character style (MEDIUM)

Carried forward from `rvw/Code_review 2026-05-11.md` item 4. Last
remaining custom-and-Unapplied character style after the 9-style
cleanup. `?ScanCharStyleApplications` already confirmed no run carries
it.

**Action:** `ActiveDocument.Styles("Normal text").Delete`; re-run
`?ScanCharStyleApplications`; expect Custom Unapplied count = 0.

Originated `rvw/Code_review 2026-05-08.md` 6h.

### 5. Apply Row Pitch Diagnostic to two un-hyphenated 10-page ranges (MEDIUM)

Carried forward from `rvw/Code_review 2026-05-11.md` item 5. Tooling
ready; identify the ranges and run the
survey -> histogram -> review cycle per
[`md/Row_Pitch_Diagnostic.md`](../md/Row_Pitch_Diagnostic.md).

Expected outcome:

- Stable per-side median pitch (Left and Right within ~0.1 pt).
- Clear suspect tail (Pitch > median + 1.0 pt).
- Reduced suspect count after the manual-hyphen pass; medians tighten
  further.

Originated `rvw/Code_review 2026-05-08.md` 4d.

### 6. Verify anomalous 2-column sections in production (LOW-MEDIUM)

Carried forward from `rvw/Code_review 2026-05-11.md` item 6. Resolves
naturally as production sweeps progress through the document.

Known anomaly: section 135 starting page 913 -
`Col1.Width=186.1  SpaceAfter=36.0  Col2.Width=186.1` already validated
on page 913 (7 finds, all Active, 0 OutsideBody). Watch
`rpt\SoftHyphen_Layout.log` "-- Anomalies --" block for the second
anomaly during a production sweep.

Originated `rvw/Code_review 2026-05-08.md` 3a.

### 7. Optional --clear helper for RowCharCount survey driver (LOW)

Carried forward from `rvw/Code_review 2026-05-11.md` item 7. Pure QoL;
the manual workflow in
[`md/Row_Pitch_Diagnostic.md`](../md/Row_Pitch_Diagnostic.md) works
fine. Open as a possible follow-up only.

Originated `rvw/Code_review 2026-05-08.md` 4g.

### 8. SoftHyphenSweep_FootnotesOnly sister routine (DEFERRED)

Carried forward from `rvw/Code_review 2026-05-11.md` item 8. No
triggering need yet; build the sister routine only if a
`Footnote Text` font-change pass leaves stray soft hyphens in footnote
bodies.

Originated `rvw/Code_review 2026-05-08.md` 3b.

### 9. SHA_ReplaceHard i18n consideration (FUTURE)

Carried forward from `rvw/Code_review 2026-05-11.md` item 9. Revisit
only if a non-English edition adopts soft hyphens as semantic
compound-break markers.

Originated `rvw/Code_review 2026-05-08.md` 3c.

## Pointer back to the closed arc

Full dated history of the work that produced this carry-forward state
is in [`rvw/Code_review 2026-05-11.md`](Code_review%202026-05-11.md).
That file (and the arcs it points back to) covers:

- The `AuditOneStyle` colour-check extension and the bucket-2 -> bucket-1
  promotions.
- The ribbon book-combo alias bug root-cause analysis and two-layer
  contract.
- The aeRibbon production export gateway design, plan decisions, and
  commit sequence.

Anything in this 2026-05-12 file should reference back to those arcs
for the *why*; this file holds only the **what is still open**.

## Status updates (append-only)

_(none yet)_
