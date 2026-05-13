# aeRibbon — Build Instructions

> Production build steps for the Radiant Word Bible navigation ribbon
> template (`aeRibbon.dotm`). See `md/aeProductionRibbonPlan.md` for
> background.

## Prerequisites

- Microsoft Word 365 (Windows).
- Macro-enabled file format (`.dotm`).
- VBA editor (Alt+F11) accessible.
- WSL + Python 3 (for the trim + ribbon-XML injection scripts).
- A fresh dev export: run `wsl python3 py/ribbon_export_trim.py` from the
  repo root. This must complete successfully and produce
  `aeRibbon/src/` + `aeRibbon/RoutineLog.md` with no warnings.

## Files going into the template

From `aeRibbon/src/`:

| File | Origin | Trim status |
|---|---|---|
| `basBibleRibbonSetup.bas`        | dev `src/` | trimmed via call-graph (see RoutineLog) |
| `basRibbonDeferred.bas`          | dev `src/` | as-is (all-public, all reachable) |
| `basSBL_VerseCountsGenerator.bas`| dev `src/` | trimmed (keeps `GetVerseCounts` + helper; drops generators) |
| `basUIStrings.bas`               | dev `src/` | as-is |
| `aeBibleCitationClass.cls`       | dev `src/` | trimmed |
| `aeRibbonClass.cls`              | dev `src/` | trimmed |

Files intentionally **not** included:
- `aeBibleClass.cls` — test-runner class; 0/85 routines reachable from
  any ribbon callback. Confirmed by call-graph trim.
- `ThisDocument.cls` — Word manages the template's own `ThisDocument`
  (decision §7.2 in `md/aeProductionRibbonPlan.md`).
- All other modules in `src/` — author/maintainer tooling, tests,
  generators (full list in §2.2 of the plan).

The production document model lives in `aeRibbonClass` +
`aeBibleCitationClass` (+ `GetVerseCounts` for verse counts data).

From `aeRibbon/template/`:

- `customUI14.xml` — copy of `customUI14backupRWB.xml`.

`ThisDocument.cls` is **not** in `aeRibbon/src/`. The template's own
`ThisDocument` is created by Word; if a `Document_Open` body is ever needed
in future versions, paste it into the template's `ThisDocument` manually
during build.

## Build steps

1. **Create the template.** In Word 365: File → New → Blank document →
   File → Save As → choose **Word Macro-Enabled Template (`*.dotm`)** →
   save as `aeRibbon/template/aeRibbon.dotm`. Close Word.

2. **Inject ribbon XML.** From repo root:
   ```bash
   wsl python3 py/inject_ribbon.py aeRibbon/template/aeRibbon.dotm
   ```
   The script always reads `customUI14backupRWB.xml` at the repo root.
   `aeRibbon/template/customUI14.xml` is a tracked snapshot of that file;
   keep them in sync (the trim/release pipeline copies one to the other).
   (Per `[[feedback_ribbon_injector]]` — never use RibbonX Editor for this
   project; it has a known load bug.)

3. **Import VBA modules.** Open `aeRibbon.dotm` in Word, then Alt+F11.
   - In the VBA editor: File → Import File... — multi-select supported.
     Import order does **not** matter (VBA resolves references at compile
     time, not import time).
   - The current `aeRibbon/src/` set is **3 `.bas` + 2 `.cls`** (5 files):
     `basBibleRibbonSetup.bas`, `basRibbonDeferred.bas`,
     `basSBL_VerseCountsGenerator.bas`, `aeBibleCitationClass.cls`,
     `aeRibbonClass.cls`.
   - After import, verify in Project Explorer that the two `.cls` files
     appear under **Class Modules** and the three `.bas` files appear
     under **Modules**. If a `.cls` lands under Modules, the file has
     LF-only line endings — re-run `py/ribbon_export_trim.py` (it now
     forces CRLF) and re-import.
   - Do **not** import `ThisDocument.cls` (it isn't present).

4. **Stamp the version.** Open `basBibleRibbonSetup` in the VBA editor and
   confirm (or add at the top) a constant matching `aeRibbon/VERSION`:
   ```vb
   Public Const RIBBON_VERSION As String = "1.0.0+bc71416"
   ```
   Also set the template's custom document property `aeRibbonVersion` to
   the same value (File → Info → Properties → Advanced Properties → Custom).

5. **Compile.** In the VBA editor: Debug → Compile VBAProject. Resolve
   any errors before proceeding. There must be zero compile errors.

6. **Save and close.** File → Save in Word.

7. **Smoke-check load.** Open a fresh Word session, then open
   `aeRibbon/docx/aeRibbon-host.docx` (or any blank `.docx`) with the
   template attached (File → Options → Add-ins → Manage: Templates → Go →
   Add → select `aeRibbon.dotm`). Confirm the **Radiant Word Bible** tab
   appears and `RibbonOnLoad` fires (visible in Immediate window).

8. **Release record.** Append a row to `aeRibbon/RELEASES.md` with:
   version, build date, dev SHA (from `git rev-parse --short HEAD`),
   QA gate results, and SHA-256 of the `.dotm`:
   ```bash
   sha256sum aeRibbon/template/aeRibbon.dotm
   ```
   Copy the built `.dotm` and the `RoutineLog.md` snapshot into
   `aeRibbon/releases/<version>/`.

## Producing the production Bible `.docx`

Per plan §7 decision 5, this is a **manual step** owned by the
Editor/Developer (Option 1).

1. Open the current dev `Peter-USE REFINED English Bible CONTENTS.docm`
   in Word 365.
2. File → Save As → choose **Word Document (`*.docx`)** as the format.
   Word will warn that VBA will be removed — that is the desired outcome:
   the production document must be code-free so the author can open it
   for comments-only review without macro-security prompts.
3. Save as `aeRibbon/docx/Radiant-Word-Bible.docx` (final filename TBD).
4. Verify by reopening: no macro-security banner appears; the Bible
   content is intact.
5. The Editor/Developer attaches `aeRibbon.dotm` once on their machine
   (File → Options → Add-ins → Templates) and runs Gate G8 against this
   `.docx`. The same template can be shipped to the author later for
   their own attach step.

This is expected to be re-run for every release until the
build/test loop stabilises. If/when the manual step becomes a release
pain point, revisit plan §7 decision 5 to consider automating via Word
COM (Option 2).

## Re-building after a dev-source change

The export gateway is **idempotent**:

1. `git pull` / sync dev changes into `src/`.
2. `wsl python3 py/ribbon_export_trim.py` — regenerates `aeRibbon/src/`
   and `aeRibbon/RoutineLog.md`.
3. `git diff aeRibbon/` shows the actual production-surface impact of
   the dev change.
4. Re-run steps 1–8 above for a fresh `.dotm`.

## Code signing

Deferred for v1.0.0 (per plan §7 decision 3). Users enable macros via the
Word Trust Center on first open. Candidate for a later MINOR release if a
cert becomes available.
