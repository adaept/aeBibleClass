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

### 2. Define colors used in the docx (HIGH) - CLOSED 2026-05-13

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

### 2026-05-13 - Item 2 CLOSED (infrastructure); item 13 spawned

Item 2 ("Define colors used in the docx") closes today on the
strength of:

- New `basBiblePalette.bas` module (Step A) - single source of
  truth, 12 named colours, late-bound, theme-extensible API.
- Three legacy call sites rewired (Step B) - `Module1.HexToRGB`
  shim, `EnsureFootnoteReferenceStyleColor` palette-driven (and
  corrected to Blue), `basTEST_aeBibleTools.GetColorNameFromHex`
  shim, `ListAndCountFontColors` palette-driven with distinct
  `wdColorAutomatic` row.
- Histogram baseline captured (item 10 Q3) - we now know what
  colours actually appear in the production docx at run level.

Two scoped sub-tasks from the original item 2 description were
not completed and are spawned as descendants rather than left as
silent debt:

- **Item 11** (RESEARCH): identify `#7F9698`, `#C00000`, lone
  Orange.
- **Item 13** (MEDIUM): convert the four
  `wdColorAutomatic`-baselined styles (`TheHeaders`,
  `TheFooters`, `Selah`, `EmphasisBlack`) to explicit literals,
  plus add a `wdThemeColorNone` taxonomy check.

Item 13 is sequenced after item 11 because the histogram
unknowns may belong to the very styles item 13 needs to convert
- doing 11 first surfaces the literal value that 13 should
write.

### 2026-05-13 - Correction: histogram scope, plus item 11 first probe

Earlier today's "Item 10 research probes" entry described the
histogram caveat as: *"shows run-level explicit overrides, not
rendered colours."* That was wrong. The histogram via
`ActiveDocument.Words` reads `Range.Font.Color`, and Word resolves
the style chain when you read `Font.Color` on a Range - so the
histogram counts the **resolved (effective rendered) colour**,
including style-inherited values. Find with `.Font.Color = X` is
the routine that matches only explicit run-level overrides.

The two scopes are distinct:

| Tool | Reads | Scope |
|---|---|---|
| `ListAndCountFontColors` (via `ActiveDocument.Words`) | resolved color | run override + style inheritance |
| `DescribeFirstRunOfColor` (via Find) | explicit override only | run override |
| `DescribeStylesCarryingColor` (walks styles) | style's Font.Color | style chain |

The mistake came from conflating Find's behavior with the
histogram's. The two helpers are now complementary by design and
the module-header notes in `basBiblePalette.bas` were corrected
to match.

First-pass results for item 11:

- **`#7F9698` (32,001 occurrences) - NOT FOUND by
  `DescribeFirstRunOfColor`.** None of the 32K runs carry an
  explicit override; the colour is being applied via style
  chain. `DescribeStylesCarryingColor` (helper added this
  session) is the right tool to identify which style. Pending
  next probe.
- **`#C00000` (153) - hand-coloured Jesus quotation in
  AuthorBodyText.** First match: page 959, paragraph and run
  both `AuthorBodyText`, text `"Blessed is he who takes no
  offense at me."` (Luke 7:22). This is a Words-of-Jesus
  quotation that pre-dates the `WordsOfJesus` character-style
  convention - 153 scattered instances of hand-coloured
  quotations inside commentary text. Note the colour is
  `#C00000` (192,0,0), not the standard `#800000` (128,0,0) the
  `WordsOfJesus` style carries. Cleanup target: repoint to
  `WordsOfJesus` style and remove the run-level override. The
  editorial decision is whether the unified colour should be
  `#800000` or `#C00000`.
- **`#FFA500` Orange (1) - redundant override on Genesis 1:1
  Chapter-Verse marker.** Page 27, paragraph `VerseText`, run
  style `Chapter Verse marker`, text `"1"`. The run's style
  already supplies Orange; this is a duplicate explicit
  override (likely an artifact of when the style was first
  applied to that run). Cleanup: remove run-level Font.Color
  and let the style provide it.

Next probe (pending re-import of `basBiblePalette.bas`):

```vba
DescribeStylesCarryingColor RGB(127,150,152)   ' identify the #7F9698 style
```

### 2026-05-13 - Item 11: #7F9698 was wdUndefined (histogram bug)

`DescribeStylesCarryingColor RGB(127,150,152)` also returned
**NOT FOUND**. No style carries the colour, and Find can't
locate any run with it. The histogram nevertheless reports
32,001 runs at this value.

Diagnosis: `9999999` decimal is the `wdUndefined` sentinel that
Word returns from `Range.Font.Color` when the range spans
**mixed colours**. `9999999 = 127 + 150*256 + 152*65536`, which
byte-decomposes to `RGB(127, 150, 152) #7F9698` - a coincidence,
not a real colour.

`ActiveDocument.Words` iterates word-by-word; many words straddle
a colour boundary (a word partly inside a styled run, partly
outside). Every such mixed word reads `wdUndefined` from
`Range.Font.Color` and got binned into the phantom #7F9698
bucket. So the 32,001 count is real (that many mixed-color words
exist), but it does not represent a colour in the document.

The histogram had the same class of bug for `wdColorAutomatic`
before Step B - silently lumped into `RGB(0,0,0)` Black via
byte-decompose math. Today the analogous fix is applied to
`wdUndefined`:

- `basTEST_aeBibleTools.ListAndCountFontColors` now reports
  `wdUndefined` as its own row: `"wdUndefined (9999999) -
  Mixed (range spans multiple colors)"`. No phantom-color bucket.

Item 11 outstanding work after this correction:

- ~~Identify `#7F9698` style.~~ Resolved as wdUndefined sentinel;
  not a real colour.
- **`#C00000` (153)**: hand-coloured Jesus quotations in
  AuthorBodyText runs. Cleanup target. Editorial: keep `#C00000`
  or migrate to `#800000` WordsOfJesus style.
- **`#FFA500` Orange (1)**: redundant override on Genesis 1:1
  Chapter Verse marker. Cleanup: remove run-level override.

Once those two cleanups are decided and applied, item 11 closes.
Re-running `ListAndCountFontColors` after this session's
histogram fix should produce a cleaner picture: the
`wdUndefined` row will appear (informational), and the bogus
`(127,150,152) #7F9698` row will be gone.

### 2026-05-13 - Histogram caveat + accurate per-color counter

Post-fix histogram showed Blue = 296, but the production docx
has 1000 footnotes - expected Blue is ~2000 (1000 reference
markers in the body + 1000 matching markers at the start of each
footnote in the Footnotes story). The 296 figure is a Word-level
granularity artifact:

- `ActiveDocument.Words` iterates Word-by-Word.
- A Word that contains a footnote reference marker glued to its
  anchor word (`Lord.<ref>`) spans two different colours and
  `Range.Font.Color` returns `wdUndefined` for the whole Word.
- Those Words bin into the `wdUndefined` row (32,001 count is
  consistent with this), not into Blue.
- Single-character coloured runs - footnote refs, verse number
  markers, chapter markers - are systematically undercounted in
  their colour row.
- Plus `ActiveDocument.Words` walks MainText only; the matching
  reference markers in the Footnotes story aren't counted at
  all.

This makes the histogram a fast-but-vague approximation.
Acceptable as a "what colours are present at all" sanity check,
not for accurate per-colour counts.

**Two additions this session to give accurate per-colour counts:**

- `basBiblePalette.CountRunsWithColor(rgbLong) As Long` -
  Find-based scan across all primary StoryRanges. Returns the
  exact run count for the given colour. Slower than the
  histogram but authoritative.
- `basBiblePalette.ReportRunsWithColor(rgbLong)` - same scan,
  prints per-story breakdown plus total. Useful for seeing
  WHERE the runs are.
- `ListAndCountFontColors` now prints a caveat header before
  its rows explaining the Word-level limit and pointing to the
  two new helpers.

Expected values for sanity-check probes:

| Colour | Expected | Source |
|---|---|---|
| Blue (Footnote Reference) | ~2,000 | 1000 footnotes x 2 (body marker + footnote-start marker) |
| Orange (Chapter Verse marker) | = N verses | one per Chapter/Verse start |
| Emerald (Verse marker)        | = N verses | one per verse marker |
| DarkRed (WordsOfJesus / EmphasisRed) | ~47K | matches histogram order of magnitude |
| `#C00000` (legacy quotes)     | ~153      | hand-coloured Jesus quotes |

Strays detection (runs of an expected colour appearing in
unexpected styles, or vice versa) is a separate need surfaced by
this analysis - tracked as a follow-on under item 11 once the
authoritative counts above are confirmed.

### 2026-05-13 - Item 11 authoritative counts

`CountRunsWithColor` / `ReportRunsWithColor` against the live
production docx:

| Colour | Actual | Expected | Notes |
|---|---:|---:|---|
| Blue (Footnote Reference) | 2,015 | 2,000 | 1000 body + 1000 footnote-start + 15 surplus |
| Orange (Chapter Verse marker) | 31,102 | 31,102 | clean match - canonical Protestant verse count |
| Emerald (Verse marker) | 31,102 | 31,102 | clean match |
| DarkRed (WordsOfJesus + EmphasisRed) | 2,262 | -- | contiguous runs; histogram's 47,874 was Word count |
| `#C00000` legacy quotes | 7 | ~153 (Words) | 7 contiguous runs, each spanning many words |

Two clean validations (Orange, Emerald) - the
`Chapter Verse marker` and `Verse marker` styles are applied
uniformly across all 31,102 verses with no missing or duplicate
markers.

Blue surplus breakdown via `ReportRunsWithColor`:

```
ReportRunsWithColor: (0,0,255) #0000FF
  MainText            1014
  Footnotes           1001
  TOTAL               2015
```

Editorial reading (from the operator at the doc):

- MainText surplus = 14: Word's built-in `Hyperlink` character
  style is Blue. These are legitimate hyperlinks, not strays.
- Footnotes surplus = 1: a single genuine stray in the Footnotes
  story. Needs identification.

#C00000 reduced from "~153 cleanup targets" to "7 hand-coloured
quotations" - same span of content, just counted at the right
granularity. Editorial decision (keep `#C00000` vs migrate to
`#800000` WordsOfJesus) still pending, but the scope is smaller
than first reported.

Next probe (proposed):
`ListRunsOfColorByStyle ColorFromName("Blue")` to group the
2,015 Blue runs by character-style name. The one Footnotes-story
stray will appear as the row with count 1 in some style that
isn't `Footnote Reference` or `Hyperlink`.

### 2026-05-14 - Item 11 ListRunsOfColorByStyle results

```
Blue (#0000FF) - by run style:
  Footnote Reference   2001   <-- one extra
  Hyperlink              12
  AuthorListItemTab       2   <-- new finding
  TOTAL                2015

#C00000 - by run style:
  AuthorBodyText          7
  TOTAL                   7

Orange (#FFA500) - by run style:
  Chapter Verse marker  31102
  TOTAL                31102

Emerald (#50C878) - by run style:
  Verse marker          31102
  TOTAL                31102
```

**Clean validations:** Orange and Emerald show single-style
profiles at the canonical 31,102 verse count - zero strays in
either family. Chapter Verse marker and Verse marker styling
is structurally perfect.

**#C00000:** Single-style profile - all 7 runs in
`AuthorBodyText`. Confirms these are 7 hand-coloured Jesus
quotations inside commentary text. Cleanup decision pending
(keep `#C00000` distinct vs. migrate to standard
`WordsOfJesus` style + `#800000`).

**Blue refinement:** the earlier "14 hyperlinks" was 12
hyperlinks + 2 `AuthorListItemTab` Blue runs - a new finding.
And the Footnotes-story surplus is itself styled
`Footnote Reference` (the style is correct, the *count* is
wrong: 2001 vs. expected 2000). So the stray is either:

- A duplicated FR marker on one footnote (two markers where
  one belongs), or
- An orphan FR-styled character somewhere in the Footnotes
  story with no associated footnote.

Cannot be located by colour scan alone - the style is normal,
only the count is anomalous. Next probe (proposed):
`AuditFootnoteReferenceMarkers` - walks every
`ActiveDocument.Footnotes(i)`, counts FR-styled characters
inside each `footnote.Range`, flags the footnote whose count
!= 1.

The 2 `AuthorListItemTab` Blue runs are solvable without a new
helper - Word's
`Find > Format > Style: AuthorListItemTab + Font: Blue` will
navigate to both directly.

### 2026-05-14 - Item 11 stray located: footnote 218

`AuditFootnoteReferenceMarkers` (added to `basStyleInspector`)
identified the Blue surplus:

```
total FR runs in Footnotes story=1001  (inside footnote.Range=1001, orphans=0)
  ANOMALY footnote(218): FR markers=2 (expected 1)  page=402
  text=[lit. "adversary". The devil, fallen angel Lucifer who tempte ...]
per-footnote check - 1 footnote(s) with FR count != 1.
```

**Footnote 218 on page 402 has 2 Footnote-Reference-styled
markers inside its body where 1 belongs.** Body text begins
"lit. 'adversary'. The devil, fallen angel Lucifer who tempte..."

The probe went through three iterations before producing this
clean result - worth recording the trajectory because the same
class of false-negative will likely surface again on other
audits:

1. **Iteration 1 - per-footnote `Find.Style = "Footnote
   Reference"` against `footnote.Range`** returned FR count = 0
   for all 1000 footnotes. Either `footnote.Range` excludes the
   auto-numbered marker or `Find.Style = "<string>"` does not
   match field-result characters. Wrong methodology.
2. **Iteration 2 - scan Footnotes story for FR runs, classify
   each by which `footnote.Range` contains its Start position**
   reported all 1001 FR runs as orphans. The body markers sit
   at `footnote.Range.Start - 1` (one char before the body
   proper), and a strict `>= Start` test rejected every
   legitimate marker.
3. **Iteration 3 - same scan with `MARKER_GAP = 5` backward
   tolerance and a per-footnote tally** classified all 1001
   correctly, and the per-footnote tally surfaced footnote 218
   with FR count = 2. This is the stray.

Used `oDoc.Styles(FR_STYLE)` (Style object) rather than the
string `"Footnote Reference"` on `Find.Style` - the object form
matches reliably, the string form does not in this context.
A flood guard (`MAX_PRINT = 20`) was also added so that if
classification ever breaks again, Immediate is not flooded with
1000+ lines.

Cleanup: open footnote 218 body, delete the duplicate auto-
numbered marker. Editorial (manual), not scripted - destructive
edit on production content. Re-running the audit afterward
should return 0.

Resolution: the "duplicate marker" turned out to be a
**paragraph mark inside the footnote body that carried the
`Footnote Reference` character style**. Deleting the paragraph
mark itself was not desirable (would merge two paragraphs). The
fix was to **strip the character style** from the paragraph
mark while leaving the structural ¶ in place: select the mark,
apply `Default Paragraph Font` via Ctrl+Shift+S (the Apply
Styles dialog - Default Paragraph Font does not appear in the
Styles pane under the default "Recommended" filter).

Post-fix audit:

```
total FR runs in Footnotes story=1000  (inside footnote.Range=1000, orphans=0)
per-footnote check - 0 footnote(s) with FR count != 1.
 0
```

Blue total in Footnotes story drops from 1001 to 1000 as
expected. The 2015 Blue runs across the doc are now
2000 (FR, 1000+1000) + 12 (Hyperlink) + 2 (AuthorListItemTab) =
2014 plus whatever the structural-paragraph-mark count is after
this fix. To reconcile precisely:

```vba
?CountRunsWithColor(ColorFromName("Blue"))   ' expect 2014
ListRunsOfColorByStyle ColorFromName("Blue") ' expect FR=2000, Hyp=12, ALI=2
```

(Pending verification next session.)

### 2026-05-14 - Hyperlinks moved to DarkBlue (palette-driven, locked)

Hyperlinks previously shared pure Blue (`#0000FF`) with Footnote
Reference - the cause of the 2014-vs-2000 audit ambiguity earlier
today. Moving Hyperlink + FollowedHyperlink to DarkBlue
(`#000080`, matches `wdColorDarkBlue`) disambiguates the two
roles in every colour audit and improves print legibility.

Changes:

- **`basBiblePalette.bas`**: new palette entry `DarkBlue` =
  `RGB(0,0,128)` = `#000080`. Usage field documents the
  Hyperlink + FollowedHyperlink role and the audit-separation
  rationale.
- **`basTEST_aeBibleTools.bas`**: existing
  `LockHyperlinksAlwaysBlue` rewritten as
  `LockHyperlinksToPalette` - same three-step body (Hyperlink
  style + FollowedHyperlink style + per-Hyperlink range), but
  the colour is now sourced from `ColorFromName("DarkBlue")`
  rather than hardcoded `wdColorBlue`. The old name remains as
  a one-line alias for one cycle to defend against forgotten
  external callers; the alias will be removed next session.
- **`basStyleInspector.bas`**: new `AuditHyperlinkStyling`
  function walks `ActiveDocument.Hyperlinks`, verifies each is
  styled `Hyperlink` and coloured palette-DarkBlue, reports any
  anomalies. Expected return is 0 after `LockHyperlinksToPalette`
  runs; non-zero on subsequent audits means a pasted hyperlink
  has escaped the convention.
- **`EDSG/01-styles.md`**: new section *"State-aware styles:
  print-locking"* documents the design pattern (Hyperlink /
  FollowedHyperlink as canonical case, lock both states to the
  same palette colour, audit via the matching `Audit*` function).

Operator sequence to apply:

```
LockHyperlinksToPalette         ' rewrites all 12 hyperlinks
?AuditHyperlinkStyling          ' expect 0
?CountRunsWithColor(ColorFromName("Blue"))      ' expect 2000 (drops by 12)
?CountRunsWithColor(ColorFromName("DarkBlue"))  ' expect 12
ListRunsOfColorByStyle ColorFromName("Blue")    ' expect single row: FR=2000
ListRunsOfColorByStyle ColorFromName("DarkBlue")' expect: Hyperlink=12 (and any
                                                ' AuthorListItemTab that picked
                                                ' up the new colour through
                                                ' style inheritance, TBD)
```

The 2 `AuthorListItemTab` Blue runs may or may not be related
to hyperlinks (they appeared in the Blue histogram before this
change). The post-change DarkBlue grouping will reveal: if
AuthorListItemTab stays in the Blue bucket, the 2 runs are
independent of hyperlinks and need their own investigation;
if they migrate to DarkBlue, they were inheriting from
Hyperlink and the migration is transparent.

### 2026-05-14 - Hyperlink lock applied; 2-vs-14 collection gap reframed

Post-`LockHyperlinksToPalette` measurements:

```
AuditHyperlinkStyling: 2 hyperlinks checked, 0 anomalies (initial version)
CountRunsWithColor(Blue)         = 2000   (drops from 2014, pure FR now)
CountRunsWithColor(DarkBlue)     = 14     (operator-confirmed correct)
ListRunsOfColorByStyle(Blue)     = Footnote Reference 2000  (clean)
ListRunsOfColorByStyle(DarkBlue) = Hyperlink 14             (clean)
```

The 2 `AuthorListItemTab` Blue runs migrated to DarkBlue
transparently (they were inheriting from the Hyperlink style;
the style colour change propagated automatically).

**2-vs-14 collection gap:**
`ActiveDocument.Hyperlinks.Count = 2` but 14 runs carry the
Hyperlink character style. Operator-confirmed: all 14 are in
MainText and are **concordance navigation links** - most likely
REF / PAGEREF / HYPERLINK field-result runs styled as
Hyperlink rather than first-class `Hyperlink` collection
objects. The collection only includes the 2 with an actual
`.Address` property.

**Implication for the lock and audit routines:** iterating
`Doc.Hyperlinks` covers only the 2 collection objects; the
other 12 are reached only through style-level inheritance
(steps 1 + 2 of LockHyperlinksToPalette which modify the
Style.Font.Color globally). That worked here, but the per-
instance step 3 missed 12, and the audit checked only 2 - the
"0 anomalies / 2 checked" report was misleading about coverage.

**Fix this session:**

- **`LockHyperlinksToPalette`** rewritten to walk every
  `StoryRange`, Find runs styled `Hyperlink`, and force their
  `Font.Color` / `Font.Underline`. Covers all 14 (and any
  future field-result runs). Reports the total locked in the
  completion MsgBox so coverage is visible.
- **`AuditHyperlinkStyling`** rewritten on the same model:
  Find-by-style across all StoryRanges, verify colour +
  underline. Reports anomaly count over the full styled
  population, not just collection-Hyperlinks. Anomaly rows
  now include the StoryType, page, current colour, current
  underline, and a text snippet.
- **`ReportHyperlinkStoryDistribution`** added as a one-shot
  diagnostic: per StoryRange, prints `Hyperlinks.Count` vs
  Hyperlink-styled-run count. The gap between the two columns
  shows how much of the style-discipline picture the
  collection misses. Useful baseline before/after any future
  hyperlink work.

Post-rewrite verification (pending re-import):

```vba
?AuditHyperlinkStyling             ' expect 0 anomalies / 14 styled runs checked
ReportHyperlinkStoryDistribution   ' expect MainText: Hyperlinks.Count=2, styled runs=14
```

The story-distribution probe confirms the operator's
"all 14 in MainText" reading and quantifies the field-result
vs collection-object split (expected 2 collection + 12
field-result = 14 styled total, all in MainText).

### 2026-05-14 - Hyperlink lock post-rewrite measurements + footnote-link finding

Re-running after the Find-by-style rewrite:

```
LockHyperlinksToPalette
  19 Hyperlink-styled runs across all stories; visited state neutralized.

?AuditHyperlinkStyling
  19 Hyperlink-styled run(s) checked, 0 anomaly/anomalies.

ReportHyperlinkStoryDistribution
  StoryType=1 (MainText)   Hyperlinks.Count=2  Hyperlink-styled runs=19
  StoryType=2 (Footnotes)  Hyperlinks.Count=1  Hyperlink-styled runs=0
  TOTAL across stories: collection=3  styled runs=19
```

**Why 19 styled runs vs 14 visible/clickable links (operator's
manual count).** `Find.Font.Color = X` matches **explicit
run-level overrides only**, not inherited colour from the style
chain. The earlier `CountRunsWithColor(DarkBlue) = 14` was
under-counting: 14 Hyperlink-styled runs carried an explicit
DarkBlue override (from the first lock pass), 5 inherited
DarkBlue from `Styles("Hyperlink").Font.Color` without an
explicit override. Visually identical, count divergent. The new
step-3 walker forced explicit DarkBlue + underline on all 19, so
all match Find criteria now. (Same class of explicit-vs-inherited
calibration we hit earlier with `wdUndefined` mixed-color Words
and footnote markers; recurring sharp edge worth a EDSG note.)

**Operator-confirmed editorial rule: no hyperlinks in footnotes.**
The Footnote-story collection Hyperlink is therefore a
**rule-violation finding**, not a coverage gap to patch out. Live
state:

```
?ActiveDocument.StoryRanges(wdFootnotesStory).Hyperlinks(1).Address
  http://archaeology.about.com/od/jterms/qt/jericho.htm
?...Hyperlinks(1).Range.Style.NameLocal -> "Footnote Text"
?...Hyperlinks(1).Range.Font.Color      -> -16777216 (wdColorAutomatic)
```

A real external URL was inserted into a footnote and never
restyled. It renders as plain Footnote Text - no DarkBlue, no
underline, no clickable affordance. Per the rule, the cleanup is
**review and delete the hyperlink object** (or, if the URL is
worth retaining as a citation, replace with plain text). NOT
restyle to Hyperlink.

The lock routine was deliberately **not** extended to walk
collection Hyperlinks across all stories: doing so would have
"fixed" this finding silently by restyling and recolouring,
masking future violations of the no-hyperlinks-in-footnotes
rule. Current behaviour (Find-by-style in step 3 only) preserves
the finding's visibility.

**Proposed formalisation as a RUN_THE_TESTS entry:** the rule
"no Hyperlink collection entries in `wdFootnotesStory`" is a
single-assert audit, trivially codified. Operator will propose
a test number; implementation will move from `basStyleInspector`
into `aeBibleClass`'s test surface when the slot is assigned.
Shape:

```vba
Public Function AuditFootnoteHyperlinks() As Long
    Dim story As Word.Range, n As Long
    For Each story In ActiveDocument.StoryRanges
        If story.StoryType = wdFootnotesStory Then
            n = story.Hyperlinks.Count
            Exit For
        End If
    Next story
    Debug.Print "AuditFootnoteHyperlinks: " & n & " collection Hyperlink(s) in Footnotes story (expected 0)."
    AuditFootnoteHyperlinks = n
End Function
```

**Open follow-on:** the 5 Hyperlink-styled runs in MainText that
operator did NOT count among the 14 visible/clickable links.
They carry the style without being collection-Hyperlinks (and
aren't the 12 concordance field-result navigation either, based
on the 2 + 12 = 14 expected split). Likely stale style
application from imports or hand-styled "looks like a link"
text. Listing them per-run for visual identification is a
follow-on probe if needed; deferred pending decision on whether
to chase.

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

### 13. Convert Automatic-baselined styles to explicit literals + taxonomy check (MEDIUM)

Spawned 2026-05-13 when item 2 was closed. The palette
infrastructure (Step A + Step B) is in place, but the original
item-2 scope also called for:

- Converting four character/paragraph styles whose
  `basTEST_aeBibleConfig.AuditOneStyle` baseline is
  `wdColorAutomatic` (`-16777216`) to explicit RGB / `wdColor*`
  literals after confirming editorial intent: `TheHeaders`,
  `TheFooters`, `Selah`, `EmphasisBlack`.
- A taxonomy check (extension of `AuditOneStyle` or a sibling
  routine) that fails any style whose colour resolves through a
  theme rather than an explicit literal. No Office theme colours
  are allowed anywhere.

Neither of those landed in Step A/B. They depend on per-style
editorial decisions (what should the explicit literal *be*?)
which the histogram doesn't answer. Sequence: identify the
unknown colours from item 11 first, since `TheHeaders` /
`TheFooters` may already be rendering as one of the unnamed
colours and the "convert to literal" answer falls out
automatically. Then run a small extension to `AuditOneStyle` to
flag any `ObjectThemeColor <> wdThemeColorNone`.
