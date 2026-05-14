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

### 3. Re-base remaining character styles to Default Paragraph Font (MEDIUM) - CLOSED 2026-05-13

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

### 4. Delete `Normal text` custom character style (MEDIUM) - CLOSED 2026-05-13

Carried forward from `rvw/Code_review 2026-05-11.md` item 4. Last
remaining custom-and-Unapplied character style after the 9-style
cleanup. `?ScanCharStyleApplications` already confirmed no run carries
it.

**Action:** `ActiveDocument.Styles("Normal text").Delete`; re-run
`?ScanCharStyleApplications`; expect Custom Unapplied count = 0.

Originated `rvw/Code_review 2026-05-08.md` 6h.

Resolution: see 2026-05-13 entry below. `Normal text` did not
actually exist; the palette entry was Word's built-in `Normal`
(undeletable). `AuthorQuote` was the real deletable custom style
and has been deleted.

### 5. Apply Row Pitch Diagnostic to two un-hyphenated 10-page ranges (MEDIUM) - WONTFIX 2026-05-13

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

Resolution 2026-05-13: **WONTFIX.** In practice the
survey -> histogram -> review cycle takes longer per page than a
straight manual read-through with hyphen insertion. The diagnostic
remains available in
[`md/Row_Pitch_Diagnostic.md`](../md/Row_Pitch_Diagnostic.md) for
any future case where a wide unhyphenated region needs an
objective second opinion, but it is no longer on the active
worklist.

### 6. Verify anomalous 2-column sections in production (LOW-MEDIUM) - CLOSED 2026-05-13

Carried forward from `rvw/Code_review 2026-05-11.md` item 6. Resolves
naturally as production sweeps progress through the document.

Known anomaly: section 135 starting page 913 -
`Col1.Width=186.1  SpaceAfter=36.0  Col2.Width=186.1` already validated
on page 913 (7 finds, all Active, 0 OutsideBody). Watch
`rpt\SoftHyphen_Layout.log` "-- Anomalies --" block for the second
anomaly during a production sweep.

Originated `rvw/Code_review 2026-05-08.md` 3a.

### 7. Optional --clear helper for RowCharCount survey driver (LOW) - WONTFIX 2026-05-13

Carried forward from `rvw/Code_review 2026-05-11.md` item 7. Pure QoL;
the manual workflow in
[`md/Row_Pitch_Diagnostic.md`](../md/Row_Pitch_Diagnostic.md) works
fine. Open as a possible follow-up only.

Originated `rvw/Code_review 2026-05-08.md` 4g.

Resolution 2026-05-13: **WONTFIX.** Parent workflow (item 5, Row
Pitch Diagnostic) closed WONTFIX the same day - manual hyphen
insertion is faster than the survey -> histogram -> review cycle.
With the survey driver no longer on the active worklist, the
`--clear` QoL helper has no consumer. The driver itself remains
available for ad-hoc use; deleting the prior report manually before
re-running is the documented workaround if anyone reaches for it.

### 8. SoftHyphenSweep_FootnotesOnly sister routine (DEFERRED)

Carried forward from `rvw/Code_review 2026-05-11.md` item 8. No
triggering need yet; build the sister routine only if a
`Footnote Text` font-change pass leaves stray soft hyphens in footnote
bodies.

Originated `rvw/Code_review 2026-05-08.md` 3b.

### 10. Research: legacy red-color usages and Footnote Reference value conflict (RESEARCH)

Surfaced during item 2 (palette consolidation, 2026-05-13).

**Question 1 - why does `aeBibleClass.CountRedFootnoteReferences`
probe for `RGB(255,0,0)` runs?** The "Footnote Reference"
character style is set to Purple `#663399` by
`Module1.EnsureFootnoteReferenceStyleColor`. A counter that scans
for explicit bright-red footnote references implies that at some
point the production docx contained hand-coloured red footnote
markers - either an older colour scheme, a paste-from-another-doc
artifact, or a deliberate now-obsolete convention. The probe is
still wired but its current count is unknown. Action: run
`?CountRedFootnoteReferences` (or expose it publicly first if it
is still Private) against the production docx and capture the
result. If 0, the probe is dead code and can be removed; if >0,
the surviving runs need to be reviewed and either re-styled or
documented as intentional.

**Question 2 - the Footnote Reference colour conflict.**
`basTEST_aeBibleConfig.AuditOneStyle` audits the "Footnote
Reference" style at `16711680` (= `RGB(0,0,255)`, Blue).
`Module1.EnsureFootnoteReferenceStyleColor` sets the same style
to `#663399` (Purple). One of these is stale. Action: read the
live style colour
(`?ActiveDocument.Styles("Footnote Reference").Font.Color`) and
align both code paths to that value. The new palette entry
already flags both colors (`Blue` and `Purple`) with cross-refs
in the `Usage` field so the conflict is visible from
`DumpPalette` output.

**Question 3 - any other "force black" run-level overrides?**
The earlier discussion conflated `wdColorAutomatic` and explicit
`wdColorBlack`. `UpdateBlackToAutomatic` relaxes explicit black
back to Automatic; there is no inverse routine. But the
production docx may still carry occasional run-level
`wdColorBlack` (= 0) overrides from legacy paste operations.
Action: a one-off histogram run via
`ListAndCountFontColors` (already routed through the new
palette in Step B) should reveal any `Black #000000` bucket. If
the count is non-trivial, decide whether to relax those runs to
Automatic via `UpdateBlackToAutomatic` or leave them as
deliberate overrides.

No code changes for this item - it is a diagnostics + decisions
ticket. Open follow-ons (re-style red footnotes, reconcile
Footnote Reference colour, sweep explicit blacks) become their
own items once the three answers are in hand.

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

### 2026-05-12 - Item 1 build sequence (manual + automated split)

For reference during the v1.0.0 build. The dotm build cannot be fully
automated from a Claude session: Word's UI (Save-As, VBA editor Import,
Compile, Custom Properties) is interactive. The split below is the
minimum-handoff sequence.

**Manual steps (Editor/Developer in Word 365):**

1. **Create the empty `.dotm`.**
   - Word -> File -> New -> Blank document.
   - File -> Save As -> "Save as type" =
     **Word Macro-Enabled Template (`*.dotm`)**.
   - Save to
     `C:\adaept\aeBibleClass\aeRibbon\template\aeRibbon.dotm`.
   - **Close Word completely.** The file must not be open during XML
     injection.

2. **After XML injection** (automated step below): reopen
   `aeRibbon\template\aeRibbon.dotm`, press **Alt+F11**:
   - File -> Import File... and import in this order from
     `aeRibbon\src\`:
     1. `basUIStrings.bas`
     2. `basRibbonDeferred.bas`
     3. `aeBibleCitationClass.cls`
     4. `aeRibbonClass.cls`
     5. `basBibleRibbonSetup.bas`
   - In `basBibleRibbonSetup`, after `Option Explicit`, add:
     ```vb
     Public Const RIBBON_VERSION As String = "1.0.0+bc71416"
     ```
   - **Debug -> Compile VBAProject.** Must complete with **zero errors**
     (this is Gate G6's compile sub-check).

3. **Custom property + save.**
   - File -> Info -> Properties -> Advanced Properties -> Custom tab.
   - Name `aeRibbonVersion`, Type Text, Value `1.0.0+bc71416`. Add. OK.
   - Save the `.dotm`. Close Word.

**Automated step (between manual 1 and manual 2):**

```bash
wsl python3 py/inject_ribbon.py aeRibbon/template/aeRibbon.dotm
```

Notes:
- `inject_ribbon.py` takes a positional file path; it always reads the
  ribbon XML from `customUI14backupRWB.xml` at the repo root.
  `aeRibbon/template/customUI14.xml` is a tracked snapshot of that
  file - the two must stay in sync.
- The script requires the target `.dotm` to be closed in Word.

**Watch points during this build:**

- If VBA compile (manual step 2) fails on a missing-Sub error, the trim
  script dropped a routine that turned out to be reachable. **Do not**
  hand-patch `aeRibbon/src/`; instead fix the root in
  `py/ribbon_export_trim.py` (add the routine name to roots or fix the
  caller it was reached from), re-run the script, and re-import the
  affected file in the VBA editor.
- After the build is green, the production Bible `.docx` is produced
  per `aeRibbon/BUILD.md` "Producing the production Bible `.docx`"
  (manual File -> Save As `.docx` from the dev `.docm`) before Gate
  G8 can run.

Documented in `aeRibbon/BUILD.md` (canonical) and mirrored here for
review-arc context.

### 2026-05-12 - Item 1 first build attempt: compile GREEN

`aeRibbon.dotm` v1.0.0+bc71416 built. **Debug -> Compile VBAProject:
zero errors.** Three gaps were uncovered and closed during the build;
all three are now defended in the toolchain so future builds will not
re-encounter them.

**Gap 1 - `inject_ribbon.py` could only replace, not bootstrap.**

The existing script assumed `customUI/customUI14.xml` was already
present in the target zip. A freshly-saved empty `.dotm` has no
customUI part at all, so the script errored out with
"use RibbonX Editor to add one first" - which is exactly what the
project bans (`[[feedback_ribbon_injector]]`).

Fix: `py/inject_ribbon.py` extended with a bootstrap mode (auto-detected
when `customUI/customUI14.xml` is absent). Bootstrap adds the three
customUI parts, patches `_rels/.rels` to add the ribbon-extensibility
relationship, and (always, in both modes) patches `[Content_Types].xml`
to ensure `Default Extension="png" ContentType="image/png"` is declared.
The customUI image (`adaept.png`) is staged once at
`aeRibbon/template/images/adaept.png` so the bootstrap is self-contained.

This means **the v1.0.0 build no longer needs RibbonX Editor for the
initial template build** - the python pipeline is sufficient.

**Gap 2 - "unreadable content" on first open of the bootstrapped dotm.**

Symptom: Word reported unreadable content when opening the
post-bootstrap `aeRibbon.dotm`. Root cause: the empty `.dotm` Word
created had `[Content_Types].xml` without a `Default Extension="png"`
entry. When the bootstrap added `customUI/images/adaept.png`, the
package became invalid (every part needs a declared content type).

Fix folded into gap 1 above: `patch_content_types()` runs
unconditionally (idempotent) in both bootstrap and replace modes, so
even an inadvertently-stripped `[Content_Types].xml` gets repaired on
every injection.

**Gap 3 - `.cls` files imported as standard modules (not class modules).**

Symptom: VBA editor's Project Explorer showed all `.cls` files under
"Modules", not "Class Modules". The trimmed files had a valid
`VERSION 1.0 CLASS` / `BEGIN` / `END` / `Attribute VB_Name` header, but
the VBA editor's `.cls` header parser is strict about **CRLF**
line endings and silently demotes a file to standard-module mode if it
sees only LF.

Root cause: `py/ribbon_export_trim.py` used `read_text()` (which
normalises CRLF -> LF in memory) and `write_text(..., newline="")`
(which writes whatever is in memory). The dev `src/` is CRLF; the
exported `aeRibbon/src/` was LF-only.

Fix: `write_text` now uses `newline="\r\n"` for both trimmed files
and as-is copies. The exported `aeRibbon/src/` is now CRLF
unconditionally, regardless of in-memory state.

**Gap 4 - `basSBL_VerseCountsGenerator.bas` mis-classified.**

Symptom: Compile failed with
`Sub or Function not defined: GetVerseCounts`. `GetChapterVerseMap` in
`aeBibleCitationClass.cls` calls `GetVerseCounts()`, which lives in
`basSBL_VerseCountsGenerator.bas`. The plan §2.2 had excluded that file
as a "generator" by filename. The file is in fact **mixed-purpose**:
one runtime accessor (`GetVerseCounts` + helper `ToOneBasedLongArray`)
plus three dev-time routines (`GeneratePackedVerseStrings_FromDictionary`,
`VerifyPackedVerseMap`, `ExpectedChapterCounts`).

Fix: `basSBL_VerseCountsGenerator.bas` moved into `TRIM_FILES` in
`py/ribbon_export_trim.py`. The call-graph trim correctly keeps the
two runtime routines and drops the three dev-time routines. Result
recorded in `aeRibbon/RoutineLog.md`: KEPT `GetVerseCounts`,
`ToOneBasedLongArray`; REMOVED the three generators.

**Lesson:** filename-based exclusion of `src/` files is unreliable for
mixed-purpose modules. Safer default going forward: feed every
`.bas`/`.cls` to the trim script and let call-graph reachability decide
membership. The current `TRIM_FILES`/`ASIS_FILES` lists are now
empirically validated for v1.0.0 but should be revisited for v1.1.0
when widening the production surface.

**Documentation updates:**

- `aeRibbon/BUILD.md` - import procedure corrected (order does not
  matter; multi-select supported; verify Class Modules vs. Modules in
  Project Explorer); file count corrected to 3 `.bas` + 2 `.cls`;
  `basSBL_VerseCountsGenerator.bas` added to the parts table with its
  trim status.

**Where we are:** Gates G1-G5 satisfied (pre-build). Gate G6's compile
sub-check is **GREEN**. Remaining for v1.0.0 release:

- G6: `RIBBON_VERSION` constant set in `basBibleRibbonSetup`
  (manual step); custom property `aeRibbonVersion` set on the `.dotm`.
- G7: Open `aeRibbon-host.docx` (still to be authored - see
  `aeRibbon/docx/README_host_docx.md`) with template attached; verify
  tab renders without error.
- G8: Open production Bible `.docx` (to be produced from current dev
  `.docm` per `BUILD.md` "Producing the production Bible `.docx`") and
  run the navigation smoke checklist.

### 2026-05-12 - Item 1 G6 CLOSED (+ LogHeadingData src/ fix)

Two further gaps closed during the G6 finish; G6 is now done.

**Gap 5 - `LogHeadingData` Path-not-found at template load (src/ fix).**

Symptom: opening `aeRibbon.dotm` as a template raised
`Error 76 (Path not found) in procedure LogHeadingData of Class
aeRibbonClass` via a MsgBox at ribbon load. Latent bug, never triggered
in dev because `C:\adaept\aeBibleClass\` happens to have an `rpt\`
subfolder.

Call chain: ribbon XML `onLoad="RibbonOnLoad"` ->
`basBibleRibbonSetup.RibbonOnLoad` -> `aeRibbonClass.OnRibbonLoad` ->
`EnableButtonsRoutine` -> `CaptureHeading1s` (essential book scan) then
`LogHeadingData` (diagnostic CSV writer). `LogHeadingData` opened
`ActiveDocument.Path & "\rpt\HeadingLog.txt"` for output; any host
folder without an `rpt\` subfolder raised error 76. Would have hit G7
(`aeRibbon\docx\`), G8 (production Bible docx folder), and every
end-user's filesystem.

Fix (Option A, approved): one-line guard at the top of `LogHeadingData`
in `src/aeRibbonClass.cls`:
```vb
If Dir(ActiveDocument.Path & "\rpt", vbDirectory) = "" Then Exit Sub
```
Routine now exits silently when no `rpt\` exists beside the active
document; dev behaviour where `rpt\` exists is unchanged.

This is the first **src/** change of the production-export work.
`sync/session_manifest.txt` updated to flag `src/aeRibbonClass.cls` for
re-import into the dev `.docm` files. `py/ribbon_export_trim.py`
re-run; `aeRibbon/src/aeRibbonClass.cls` carries the fix.

**Gap 6 - Editing the template vs. a new doc from the template.**

Symptom: after double-clicking `aeRibbon.dotm` in Explorer, Word's title
bar reads "Document1" instead of `aeRibbon.dotm`. Double-clicking a
`.dotm` tells Word to create a *new transient document from the
template*, not to edit the template itself. VBE saves still route into
the template (because the VBA project lives in the template that owns
it), but Word's main File -> Save prompts to save Document1 (the wrong
file). This is a UX trap rather than a bug.

Fix (documentation):
- `aeRibbon/BUILD.md` build-step 1 now explicitly warns against
  double-clicking the `.dotm` for editing; right-click -> Open (or
  File -> Open in Word) is the correct path; title bar should read
  `aeRibbon.dotm`.
- G6 save instruction now says **Ctrl+S in VBE**, not File -> Save in
  Word, so the save targets the template regardless of which Word
  document is visible.
- Note added: on close, if Word prompts "Save changes to Document1?",
  click **Don't Save**.

**Gap 7 - "Advanced Properties" UI gone in current Word 365.**

Symptom: File -> Info -> Properties -> Advanced Properties (the
documented path to add custom document properties) does not exist in
current Word 365 builds.

Fix: `aeRibbon/BUILD.md` G6 step 2 replaced with the VBE Immediate
window route - version-independent, single command:
```vb
ThisDocument.CustomDocumentProperties.Add Name:="aeRibbonVersion", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="1.0.0+bc71416"
```
Verification line documented. Runtime error 5 on `?` query interpreted
as "property doesn't exist yet - re-run Add". Rebuild path
(`= "1.0.0+..."`) documented for in-place updates.

**G6 result:**

- `RIBBON_VERSION` constant landed in `basBibleRibbonSetup`.
- Custom property `aeRibbonVersion = 1.0.0+bc71416` confirmed via
  Immediate-window `?` query.
- `aeRibbon.dotm` re-opens as a template with **no MsgBox**; ribbon
  load is clean.
- Compile remains GREEN.

Next: G7 (open `aeRibbon-host.docx` with template loaded; verify tab
renders), then G8 (navigation smoke against the production Bible
`.docx`).

### 2026-05-12 - Item 1 G7 CLOSED

Opened `aeRibbon/docx/aeRibbon-host.docx` in a fresh Word session with
`aeRibbon.dotm` loaded from `%APPDATA%\Microsoft\Word\STARTUP\`.

Visible result:

- **Radiant Word Bible** tab appears in the ribbon.
- `Alt` shows the `Y2` tab keytip; `Alt, Y2` switches to the tab.
- Selectors and Prev/Next buttons render **enabled** (see note below).
  Only **Go** and **New Search** render greyed-out (`m_currentBookIndex = 0`).
- No error dialog at load.

Immediate window trace (confirmed by paste):
```
>> RibbonOnLoad at 23:49:45
RibbonController: Class_Initialize at 23:49:45
RibbonController: Ribbon ready at 23:49:45
RibbonController: EnableButtonsRoutine
CaptureHeading1s: Stored 0 Heading 1 entries (saved=True).
```

`AutoExec` did not print in this capture - expected: `AutoExec` fires
when the template loads at Word startup (from the STARTUP folder),
before VBE is opened to view the Immediate window. Per-docx open does
not re-fire `AutoExec`. The `LogHeadingData` line is also (correctly)
absent - the guard added in Gap 5 fires because no `rpt\` exists beside
`aeRibbon-host.docx`.

**Design clarification recorded in QA_CHECKLIST.md G7:**

The expectation that "all selectors render disabled" was wrong - it
matched the conceptual state-machine diagram in `md/Ribbon Design.md`
but not the implementation. Actual design (verified against
`aeRibbonClass.cls`):

| Control | Enabled state | Rationale |
|---|---|---|
| Prev/Next Book, Chapter, Verse | always True | "always-enable (#599)"; click handlers guard bounds |
| Book/Chapter/Verse selectors | always True | `m_ribbon.Invalidate` from `onChange` is deferred (fires after Tab routing); selectors must be enabled tab stops from initial render or Tab would skip past them |
| Go | `m_currentBookIndex <> 0` | greys until a book is selected |
| New Search | `m_currentBookIndex <> 0` | same |

`QA_CHECKLIST.md` G7 row updated to reflect the actual design with
inline rationale pointing at the relevant `GetEnabled` callbacks.

**G7 result:** PASS. Ready for G8.

### 2026-05-13 - Gap 8: dual "Radiant Word Bible" tabs after STARTUP staging

Symptom: after copying `aeRibbon.dotm` into
`%APPDATA%\Microsoft\Word\STARTUP\` (G6 step 3), reopening the canonical
`aeRibbon\template\aeRibbon.dotm` produced **two** ribbon tabs both
labelled "Radiant Word Bible".

Root cause: dual-load. Word treats the STARTUP folder as a global
template directory; everything in it loads in every Word session. When
the canonical `.dotm` was then opened directly for editing, Word loaded
**both** copies simultaneously. Each copy declares the same customUI
tab, so two tabs render. The tabs look identical but each is bound to
its own `aeRibbonClass` instance - clicking a control on one tab uses
that template's VBA, the other tab uses the other template's VBA. A
real testing footgun.

Resolution (no code change - workflow rule):

- The STARTUP-folder copy is a **deployment artefact**, not an editing
  target. Canonical source-of-truth is always
  `aeRibbon\template\aeRibbon.dotm`.
- **Before any further template edit:** close Word, delete the
  STARTUP-folder copy, then open the canonical.
- **After saving:** close Word, re-copy the freshly-saved canonical
  back into STARTUP only if needed for the next docx smoke test.
- **Never have both copies present at the same time.**

Doc updates:

- `aeRibbon/BUILD.md` G6 step 3 now carries the workflow rule
  explicitly: "STARTUP copy is a deployment artefact, not an editing
  target" + the delete-before-edit / re-copy-after-save cycle.

**Status:** dual-tab condition cleared by deleting the STARTUP copy.
G8 still pending; will run with **single** ribbon tab loaded from the
canonical (or from STARTUP after editing is complete).

### 2026-05-13 - .Color casing demotion (normalizer gap)

Symptom: `src/basTEST_aeBibleConfig.bas:544` rendered
`If oStyle.Font.color <> CLng(vExpColor) Then` (lowercase `color`) when
the canonical form is `Font.Color`. Cause: the VBE auto-case behaviour
demoted `Color` -> `color` when an identifier with that exact spelling
was typed lowercase elsewhere in the project; VBE then back-propagated
the lowercase form across every reference. `py/normalize_vba.py` had no
rule for `.Color` so the normalize-before-commit pass let the
corruption through.

Root-cause class: same family as the previously-fixed `Space()`
(issue #616) and other identifier-casing rules - any identifier the VBE
auto-corrects must have a normalizer rule, or VBE wins and the
canonical casing rots.

Fix: `py/normalize_vba.py` - one new rule, inserted immediately after
`.Font` (the typical access path is `Font.Color`):
```python
(r'(?i)\.Color\b',          '.Color',           '.Color property on Font/Style/object (Font.Color access)'),
```

Normalizer re-run against `src/`: **75 replacements across 9 files**:
`Module1.bas` (15), `aeBibleClass.cls` (9), `basAuthorStyles.bas` (4),
`basFixDocxRoutines.bas` (6), `basStyleInspector.bas` (2),
`basTEST_aeBibleConfig.bas` (2), `basTEST_aeBibleFonts.bas` (3),
`basTEST_aeBibleTools.bas` (29), `basWordRepairRunner.bas` (5).

Post-fix grep on `src/` for any lowercase `.color`: zero hits.

All 9 files listed `[IMPORT]` in `sync/session_manifest.txt` for
re-import into the dev `.docm` files this session. Production
`aeRibbon/src/` is unaffected - none of the trimmed production routines
reference `.color` (the regenerated `aeRibbon/src/` re-runs would
include the corrected casing automatically the next time
`py/ribbon_export_trim.py` runs).

### 2026-05-13 - RIBBON_VERSION declaration site moved to dev source

Improvement: `Public Const RIBBON_VERSION As String = ""` added at the
top of `src/basBibleRibbonSetup.bas` (line 5, module-header region
before any routine). The trim script preserves the module-header
region byte-for-byte, so the declaration carries through to
`aeRibbon/src/basBibleRibbonSetup.bas` automatically on every export.

Before this change: `aeRibbon/BUILD.md` G6 step 1 told the operator to
**add** the constant line during the build. That worked but had two
weaknesses:
- The constant had no home in the dev source. Anyone reading
  `src/basBibleRibbonSetup.bas` had no signal that production carries
  a version constant.
- Every build had to remember the exact line syntax.

After this change: the declaration is permanent in dev source as an
empty-string sentinel. The per-release build step is now reduced to
**set the value** (paste the version string from `aeRibbon/VERSION`).
The sentinel must **not** be committed back to `src/` with a populated
value - source stays empty, build sets the value per release.

Doc updates:
- `aeRibbon/BUILD.md` G6 step 1 rewritten: "set the value for this
  release" instead of "add the line". Inline note that the empty
  sentinel stays in `src/`; only the template's copy carries the
  populated value.
- `sync/session_manifest.txt` marks `src/basBibleRibbonSetup.bas`
  `[IMPORT]` so the dev `.docm` files pick up the declaration too
  (dev imports will see `RIBBON_VERSION = ""` - harmless, unused on
  the dev side until/unless someone wires a callsite).

Verification: `py/ribbon_export_trim.py` re-run; trimmed
`aeRibbon/src/basBibleRibbonSetup.bas` line 5 reads
`Public Const RIBBON_VERSION As String = ""`. The build operator
edits the template's copy of the module post-import, not the source.

### 2026-05-13 - BibleAbbreviationList.md created + GetBookAliasMap expanded

Improvement: new reference doc `md/BibleAbbreviationList.md` captures
the unified, deduplicated non-SBL abbreviation set drawn from
traditional English publishing (KJV-lineage), standard
church/academic abbreviations, and digital shortest-form systems
(Logos-style, concordances, BibleStudyTools). The doc is formatted
as proper Markdown (H1/H2/H3/H4 hierarchy, bulleted books with bold
names, no tables) so it renders cleanly in VS Code and any Markdown
viewer.

`GetBookAliasMap` in `src/aeBibleCitationClass.cls` was then
extended to include every form listed in the new reference. The
single-letter prohibition still holds (comment unchanged); all
additions are two-or-more characters. Closed-up no-space forms
(e.g. `1SA`, `2PE`, `1JO`) are added alongside the existing
spaced forms (`1 SA`, `2 PE`, `1 JN`) so parsers can resolve
either convention.

Additions by book (UPPERCASE map keys, ASCII only - per the
in-VBA-ASCII rule):

- **OT.** `NB` (Numbers), `JSH` (Joshua), `JDG`/`JDGS`/`JG`
  (Judges), `RTH` (Ruth), `1SA`/`2SA` (Samuel), `1KI`/`2KI`
  (Kings), `1CH`/`2CH` (Chronicles), `PSS`/`PSM` (Psalms),
  `ECCLES`/`QOH` (Ecclesiastes - `QOH` for Qoheleth),
  `SOS` (Song of Songs), `JR` (Jeremiah), `EZK` (Ezekiel),
  `JNH` (Jonah), `MC` (Micah), `ML` (Malachi).
- **NT.** `MRK`/`MR` (Mark), `JHN` (John), `RM` (Romans),
  `1CO`/`2CO` (Corinthians), `EPHES` (Ephesians), `PHP`
  (Philippians), `1 TH`/`1TH`/`2 TH`/`2TH` (Thessalonians),
  `1TI`/`2TI` (Timothy), `PHM` (Philemon), `JM` (James),
  `1PE`/`2PE` (Peter), `1 JO`/`1 JHN`/`1JO`/`2 JO`/`2JO`/`3 JO`/`3JO` (Johannine epistles).

No removals. All pre-existing keys (`GEN`, `MATT`, `1 SAM`, etc.)
remain to preserve dictionary lookups elsewhere in the class
(`ResolveAlias`, audit routines).

Edit scope: `src/aeBibleCitationClass.cls` only. The aeRibbon
production copy `aeRibbon/src/aeBibleCitationClass.cls` is a
trim-generated artifact and will be refreshed by
`py/ribbon_export_trim.py` on the next ribbon build; no manual
edit there.

Verification deferred to the next test-harness run
(`basTEST_aeBibleCitationClass`) - additions are pure
`aliasMap.Add` calls with unique keys, so the risk is duplicate-key
runtime errors if any addition collides with an existing entry.
Spot-checked: no collisions in the additions above against the
pre-existing key set.

### 2026-05-13 - Item 4 CLOSED (Normal text was Normal; AuthorQuote deleted)

Diagnosis for item 4 ("Delete `Normal text` custom character
style") showed `BuiltIn = False`, `Type = 2` (wdStyleTypeCharacter),
`BaseStyle = "Default Paragraph Font"` and no dependents
(`WhoReferencesNormalText` returned nothing). The Styles-pane
**Delete** entry remained greyed regardless, and the style was
set `Priority = 99` as the hide-not-delete fallback.

Correction on closer look: the palette entry under suspicion was
the built-in `Normal` style, **not** a custom `Normal text` style.
There is no `Normal text` style in this document - the name was
inherited from an earlier review note and never verified against
the live palette. Built-in `Normal` is non-deletable by design;
that is the correct end state, not a defect.

The actual remaining custom-and-Unapplied character style was
`AuthorQuote`. It has now been deleted via
`ActiveDocument.Styles("AuthorQuote").Delete`. Re-run of
`?ScanCharStyleApplications` is expected to report Custom
Unapplied = 0.

Net effect: Item 4 closes with a name correction
(`Normal text` -> `AuthorQuote`) and the cleanup goal met. No
remaining deletable character-style cruft.

### 2026-05-13 - Items 5 and 7 WONTFIX (Row Pitch Diagnostic shelved)

Item 5 (Apply Row Pitch Diagnostic to two un-hyphenated 10-page
ranges) closed **WONTFIX**: in practice the
survey -> histogram -> review cycle takes longer per page than a
straight manual read-through with hyphen insertion. The diagnostic
remains documented at
[`md/Row_Pitch_Diagnostic.md`](../md/Row_Pitch_Diagnostic.md) as
an objective second opinion for any future wide-unhyphenated case,
but is no longer on the active worklist.

Item 7 (Optional `--clear` helper for RowCharCount survey driver)
closed **WONTFIX** as a downstream consequence: the helper's only
consumer was the item-5 workflow. With the parent workflow shelved
there is no demand for the QoL flag. The survey driver itself
remains available for ad-hoc use; manual deletion of the prior
report file before re-running is the documented workaround.

### 2026-05-13 - Item 6 CLOSED (second 2-col anomaly validated)

Item 6 (Verify anomalous 2-column sections in production) carried
two anomalies surfaced by `SoftHyphen_DiagnoseLayout`:

- Section 123 (page 886): Col1=186.1/36.0 Col2=186.1
- Section 135 (page 913): Col1=186.1/36.0 Col2=186.1

Section 135 was already validated upstream
(7 finds, all Active, 0 OutsideBody). Today section 123 was
validated with a single-page dry-run sweep:

```
RunSoftHyphenSweep_Across_Pages_From 886, 1, True
SoftHyphenSweep p886 (Verso): 9 find(s) - 9 Active, 0 Stray
                              (0 Removed, 0 Skipped), 0 OutsideBody
```

Pass criterion (0 OutsideBody) met; bonus 0 Stray. The
classifier's column-X constants accommodate the narrower
186.1/36.0 variant without retuning. Both 2-col anomalies are now
confirmed harmless geometry variants, not classifier risks.

Item 6 closes.

### 2026-05-13 - Item 3 CLOSED (four character styles repointed)

`?AuditCharStyleBases` reported four offenders:

```
Endnote Reference  ->  (none)
Hyperlink          ->  (none)
Page Number        ->  Footer Char
Words of Jesus     ->  (none)
```

Three were built-in Word styles standing as standalone roots
(empty `BaseStyle`); one was the chained-inheritance special case
called out in the carry-forward (`Page Number -> Footer Char`).
`Footer Char` itself was not an offender and was left untouched.

Repoint block:

```
ActiveDocument.Styles("Page Number").BaseStyle       = "Default Paragraph Font"
ActiveDocument.Styles("Endnote Reference").BaseStyle = "Default Paragraph Font"
ActiveDocument.Styles("Hyperlink").BaseStyle         = "Default Paragraph Font"
ActiveDocument.Styles("Words of Jesus").BaseStyle    = "Default Paragraph Font"
```

Post-block `?AuditCharStyleBases` -> **0**. All 34 in-use character
styles (excluding the root `Default Paragraph Font`) are now
explicitly rooted in the default.

Caveat for the built-ins (`Hyperlink`, `Endnote Reference`,
`Page Number`): Word can snap `BaseStyle` back to the original on
theme switches, template reattach, or paste-from-HTML operations.
If a future audit run re-surfaces these three, that is Word
restoring built-in defaults, not a regression in our cleanup. The
custom `Words of Jesus` repoint is stable.

No visual changes observed - the repoints were taxonomy-only;
none of the rerouted hops carried meaningful intermediate font
properties.

Item 3 closes.

### 2026-05-13 - Item 2 Step A: basBiblePalette.bas added

Step A of item 2 (Define colors used in the docx): new
self-contained module `src/basBiblePalette.bas` introduces a
single source of truth for the named colors used and allowed in
the production document. Step A is purely additive - no existing
call sites are rewired. Step B (refactor `Module1.HexToRGB`,
`basTEST_aeBibleTools.GetColorNameFromHex`,
`basTEST_aeBibleTools.ListAndCountFontColors` to delegate to the
new module) is deferred until the palette is validated.

Public API:

- `GetPalette(theme)` - returns a `Scripting.Dictionary` keyed by
  Name -> nested `Scripting.Dictionary` of seven fields (Name, R,
  G, B, RgbLong, HexCode, Usage). Only `theme = "Default"` is
  populated; `"Dark"` and `"Colorblind"` raise "not implemented"
  so call sites can be wired now and themes added later without
  an API change. (Nested-dict layout chosen over a `Public Type`
  record because VBA forbids passing UDTs declared in .bas
  modules to late-bound functions - a Dictionary stays in a .bas
  module without that restriction.)
- `ColorFromName(name)` -> RgbLong (raises if unknown).
- `NameFromColor(rgbLong)` -> Name (returns "" if unknown -
  audit-friendly).
- `LongToHex(rgbLong)` -> "#RRGGBB" (byte-correct; fixes the
  BGR-order bug in `aeBibleClass.ColorToHex` which Hex-encodes
  the raw Long).
- `HexToLong(hex)` -> RgbLong (replaces `Module1.HexToRGB`).
- `LongToRgbString(rgbLong)` -> "(R,G,B)" (replaces private
  `basTEST_aeBibleTools.RGBToString`).
- `DumpPalette` - diagnostic dump to Immediate window.

Palette content: 12 named colors (Black, White, Red, DarkRed,
Green, DarkGreen, Emerald, Blue, Gold, Orange, Purple, Gray).
The earlier "15" estimate over-counted because the in-doc
semantic colors (FootnotePurple, ChapterVerseOrange,
VerseMarkerEmerald) deduplicate against the generic names
(Purple, Orange, Emerald). The `Usage` field on each record
documents all document roles a color plays so the count stays
honest while still surfacing semantic intent in `DumpPalette`
output.

Design decisions captured in the module header:

- `wdColorAutomatic` deliberately excluded - it is a sentinel
  ("inherit, will be black in default theme"), not a color. Theme
  work depends on body text staying `wdColorAutomatic` so page-
  background inversion does the right thing. Pulling it into the
  palette would tempt callers to swap it out and break that
  mechanism.
- Office `ObjectThemeColor` deliberately excluded - too niche,
  too template-coupled, not portable.

Late binding throughout. No project references added.

Verification (deferred to next VBE session):
```
DumpPalette
?ColorFromName("Purple")        ' expect 10040166
?LongToHex(10040166)            ' expect "#663399"
?NameFromColor(RGB(255,165,0))  ' expect "Orange"
```

Three research questions surfaced during Step A are captured as
new item 10 below (legacy red-footnote probe, Footnote Reference
colour conflict between audit and ensure routines, possible
residual `wdColorBlack` overrides).

### 2026-05-13 - Item 2 Step A verified + Emerald catch

`DumpPalette` plus three round-trip probes confirm Step A:

```
?ColorFromName("Purple")        -> 10040166
?LongToHex(10040166)            -> #663399
?NameFromColor(RGB(255,165,0))  -> "Orange"
```

All 12 entries render correctly with their R/G/B/Long/Hex
fields populated from the `RGB()` and byte-decompose helpers.

**Side catch: Emerald.** Earlier chat math gave `RGB(80,200,120)`
as `Long = 7849040`; the live `DumpPalette` shows `7915600` (the
correct value: `80 + 200*256 + 120*65536`). The palette is the
source of truth from this point on - any hand-computed colour
literal anywhere else in the codebase must be cross-checked
against `DumpPalette` rather than trusted on its own. Action
folded into research item 10 question 3 (residual overrides
audit will surface any Emerald-equivalent literals).

### 2026-05-13 - Item 2 Step B: legacy call sites rewired

Step B refactors three legacy call sites to delegate to
`basBiblePalette`. Each change is small, surgical, and behaviour-
preserving (or behaviour-fixing in one case noted below).

- `Module1.EnsureFootnoteReferenceStyleColor` no longer hardcodes
  `"#663399"`. It now reads `ColorFromName("Purple")`. The
  semantic intent ("apply the Footnote Reference colour") is
  preserved without coupling the routine to a specific hex
  literal - a future palette swap or audit re-point will not
  require editing this routine.
- `Module1.HexToRGB` is now a one-line shim that delegates to
  `basBiblePalette.HexToLong`. Kept under its original name so
  any external caller resolving it still works; new code should
  call `HexToLong` directly.
- `basTEST_aeBibleTools.GetColorNameFromHex` is now a 5-line
  shim that delegates to `NameFromColor` (with hex->Long
  conversion via `HexToLong`). Preserves the historical
  "Unknown Color" return string when the value is not in the
  palette.
- `basTEST_aeBibleTools.ListAndCountFontColors` rewritten to
  tally by raw `Font.Color` Long (rather than hex string),
  resolve names via `NameFromColor` at print time, and report
  `wdColorAutomatic` as a distinct row ("Automatic (inherit)")
  rather than crushing it into a `(0,0,0)` bucket. This last
  change is the small behaviour fix: previously the
  byte-decompose math on the `-16777216` sentinel silently
  reported `RGB(0,0,0) #000000 - Black`, conflating Automatic
  runs with explicit-black runs in the histogram. Research
  item 10 question 3 depends on these two being distinct, so
  the fix is load-bearing for the next step.

No new public API surface. No project references added. Late
binding preserved.

Verification (deferred to next VBE session):
- `EnsureFootnoteReferenceStyleColor` still prints the same
  `Count of Footnote Reference = N` line.
- `?HexToRGB("#663399")` returns the same `10040166` it always
  did (now via shim).
- `ListAndCountFontColors` output: should match the prior
  format for non-Automatic colours and add a distinct
  `wdColorAutomatic` row for body text.

### 2026-05-13 - Item 10 research probes + Footnote Reference correction

Three diagnostic probes run against the production docx; results
folded back into code and palette.

**Q1 - red-footnote probe:** `Red (#FF0000) footnote references:
0`. The `CountRedFootnoteReferences` function in `aeBibleClass`
is scanning for content that no longer exists. Confirmed dead
code; queued for removal in a follow-on item.

**Q2 - Footnote Reference live colour:**
```
?ActiveDocument.Styles("Footnote Reference").Font.Color  -> 16711680
?LongToHex(...)                                          -> #0000FF
?NameFromColor(...)                                      -> Blue
```

Live state is **Blue**, not Purple. This flips the conflict
documented in item 10:

- `basTEST_aeBibleConfig.AuditOneStyle` (audits Blue 16711680)
  is correct and matches the live doc.
- `Module1.EnsureFootnoteReferenceStyleColor` (was setting
  Purple `#663399`) would have corrupted 296 existing Blue
  references if run. Now corrected to
  `ColorFromName("Blue")`.
- The palette `Usage` field on Blue/Purple was likewise wrong
  in Step A; corrected this session - Blue now documents the
  Footnote Reference role (296 occurrences), Purple is
  reclassified as palette-only.

**Q3 - colour histogram:**

| Colour                | Count   | Palette name | Note |
|-----------------------|---------|--------------|------|
| `wdColorAutomatic`    | 872,359 | Automatic    | Body text - expected. |
| `#800000` DarkRed     | 47,874  | DarkRed      | Words of Jesus / EmphasisRed - expected. |
| `#7F9698`             | 32,001  | Unknown      | Gray-blue. Needs identification. |
| `#0000FF` Blue        | 296     | Blue         | Footnote References (matches Q2). |
| `#C00000`             | 153     | Unknown      | Darker red variant. Needs identification. |
| `#FFA500` Orange      | 1       | Orange       | One stray; investigate. |
| `#000000` Black       | 0       | -            | No explicit-black overrides. |
| `#50C878` Emerald     | 0       | -            | None at run level (applied via style). |
| `#663399` Purple      | 0       | -            | Not present in doc (confirms reclassification). |

Histogram caveat: shows **run-level explicit overrides**, not
rendered colours. Styled colours (Verse marker Emerald,
Chapter Verse marker Orange, Footnote Reference Blue when
applied via style chain) read `wdColorAutomatic` on the run
and roll into the 872K Automatic count. The colours that show
up here are direct overrides, not inherited.

**Resolution and follow-ons:**

Item 10 Q1 (red probe) and Q2 (Footnote Reference conflict)
resolved this session. The remaining open work is identification
of `#7F9698` and `#C00000` and the lone explicit Orange - these
become **item 11** below to keep item 10 a clean record of the
three original questions.

### 11. Identify unnamed colours in production docx (RESEARCH)

Surfaced 2026-05-13 from item 10 Q3 histogram. The production
docx carries three non-palette colours at run level that need
semantic identification before they can be added to
`basBiblePalette` (or relaxed to the appropriate palette entry):

- **`#7F9698` (32,001 occurrences)** - gray-blue. Largest
  unknown by far. Significant semantic content. Sample a few
  runs to determine the role (commentary author colour? Note
  attribution? Header text?).
- **`#C00000` (153 occurrences)** - darker red than the
  standard `#800000`. Likely a pre-standardisation Words of
  Jesus variant, or a section-heading red.
- **`#FFA500` Orange, 1 occurrence** - a single explicit
  Chapter-Verse-equivalent Orange override. Either a stray
  legacy artifact or a deliberate one-off.

No code changes for this item; sample-and-document only.
Action: navigate to a sample of each colour via `Find > More >
Format > Font > Colour`, capture the surrounding context and
character/paragraph style, decide whether to (a) add to palette
with a semantic name, (b) repoint to an existing palette entry,
or (c) leave as documented exception.

### 12. Remove dead CountRedFootnoteReferences probe (LOW)

Surfaced 2026-05-13 from item 10 Q1. The probe returned 0
against the production docx. Either delete the function
outright (and any callers in `aeBibleClass.cls`), or leave it
with a comment noting "historical-zero, retained for regression
catch." Low priority; not blocking anything.
