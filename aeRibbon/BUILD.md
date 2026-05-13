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

   **Important — re-opening the template for editing.** Once the
   `.dotm` exists, do **not** double-click it in Explorer to edit it.
   Double-clicking a `.dotm` tells Word to create a *new transient
   document from the template* (the title bar reads "Document1"); your
   edits will be saved by the VBA editor into the template anyway, but
   the workflow is confusing and the main Word "Save" prompts for
   Document1 (the wrong file). Instead:
   - **Right-click** `aeRibbon.dotm` in Explorer → **Open** (capital O,
     the option above "New"). This opens the template itself for editing
     — title bar reads `aeRibbon.dotm`.
   - Or, in Word: File → Open → navigate to `aeRibbon.dotm`. Same
     result.

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

## Gate procedures — finishing v1.0.0

These steps pick up after the build (steps 1–8 above) and walk through
the remaining gates in order: **G6 finish → G7 → G8**.

### Attaching the template — choose one method

The ribbon tab only appears when `aeRibbon.dotm` is **loaded by Word**.
Two options:

- **Global template via Startup folder (recommended for the
  Editor/Developer's test loop).** Copy `aeRibbon.dotm` to
  `%APPDATA%\Microsoft\Word\STARTUP\` and restart Word. The template
  loads automatically in every Word session; the **Radiant Word Bible**
  ribbon tab appears whenever any Word document is open. No per-document
  attachment needed.
- **Per-document attachment (use this later when shipping to the
  author).** Open the target `.docx`, then File → Options → Add-ins →
  Manage: **Templates** → Go → in the "Templates and Add-ins" dialog,
  click **Attach...** next to "Document template" → select
  `aeRibbon.dotm` → OK. The ribbon appears only when that specific
  document is open.

The Startup-folder method is recommended below; both produce the same
runtime behaviour.

### G6 finish — version constants

The compile sub-check is already green from the build steps. Two
artefacts still need to land for G6 to close.

1. **`RIBBON_VERSION` constant.** In the VBA editor, open
   `basBibleRibbonSetup`. Immediately after `Option Explicit`, add:
   ```vb
   Public Const RIBBON_VERSION As String = "1.0.0+bc71416"
   ```
   Re-run **Debug → Compile VBAProject** — must stay green.

2. **Custom document property `aeRibbonVersion`.**

   The "Advanced Properties" entry under File → Info → Properties has
   been removed in current Word 365 builds. Use the VBA Immediate
   window instead — it's version-independent and writes the same
   underlying property.

   - With `aeRibbon.dotm` open, press Alt+F11 → Ctrl+G (Immediate
     window).
   - Paste this single line and press Enter (silent success - the
     Immediate window shows no output on a successful Add):
     ```vb
     ThisDocument.CustomDocumentProperties.Add Name:="aeRibbonVersion", LinkToContent:=False, Type:=msoPropertyTypeString, Value:="1.0.0+bc71416"
     ```
   - Verify by pasting:
     ```vb
     ?ThisDocument.CustomDocumentProperties("aeRibbonVersion").Value
     ```
     The Immediate window should print `1.0.0+bc71416`. If it raises
     **runtime error 5** the property was never added - re-run the
     `Add` line above, then re-run the `?` query.
   - **Save from the VBA editor**, not from Word: in the VBE, File →
     Save (or Ctrl+S). This saves the **template** that owns the VBA
     project — i.e. `aeRibbon.dotm`. Word's main File → Save would
     prompt to save the visible document, which (if you opened the
     `.dotm` by double-click) is a transient "Document1" you don't want
     to keep.
   - If you opened the `.dotm` correctly (right-click → Open or File →
     Open), Word's title bar reads `aeRibbon.dotm` and main File →
     Save also works — both routes write to the same template.
   - Close VBE, then close Word. If Word prompts to save Document1
     (only when you opened via double-click), click **Don't Save** —
     Document1 is the transient new-doc-from-template and isn't part
     of the build.

   Re-running the `Add` line on the same template raises error 5
   ("property already exists") - that's expected on a rebuild. To
   replace an existing value:
   ```vb
   ThisDocument.CustomDocumentProperties("aeRibbonVersion").Value = "1.0.0+bc71416"
   ```

3. **Stage the template for Word to load.**
   - Copy `C:\adaept\aeBibleClass\aeRibbon\template\aeRibbon.dotm` to
     `%APPDATA%\Microsoft\Word\STARTUP\` (open Explorer, paste
     `%APPDATA%\Microsoft\Word\STARTUP\` into the address bar, then
     paste the file). Or use the per-document method described above.
   - **Critical workflow rule:** the STARTUP-folder copy is a
     **deployment artefact**, not an editing target. The canonical
     source-of-truth is always `aeRibbon\template\aeRibbon.dotm`. If
     both copies are present and you right-click → Open the canonical,
     Word loads both at once and you will see **two "Radiant Word
     Bible" tabs**. They look identical, but each is bound to a
     different VBA instance - a real footgun for testing.
   - **Before any further template edit:** delete the STARTUP-folder
     copy, then open the canonical. **After saving:** close Word,
     re-copy the freshly-saved canonical back into STARTUP if you want
     it auto-loaded for the next docx smoke test. Never have both
     copies present at the same time.

G6 closes when steps 1–3 are done and the next Word session loads the
template without error.

### Note: `LogHeadingData` Path-not-found at template load

If the template-load ever raises `Error 76 (Path not found) in procedure
LogHeadingData of Class aeRibbonClass`, the `.dotm` is carrying a
pre-2026-05-12 copy of `aeRibbonClass`. Cause: `LogHeadingData` writes a
diagnostic CSV at `ActiveDocument.Path & "\rpt\HeadingLog.txt"`; any host
folder without an `rpt\` subfolder hit a hard error 76. Fixed in
`src/aeRibbonClass.cls` 2026-05-12 with a one-line guard. Resolution: in
the VBA editor, **Remove `aeRibbonClass`** (don't export) and re-import
`aeRibbon/src/aeRibbonClass.cls`, Compile, Ctrl+S, close, re-open.

The same fix benefits any dev `.docm` hosted in a folder without `rpt\`.
Re-import `src/aeRibbonClass.cls` into the dev `.docm` files per
`sync/session_manifest.txt`.

### G7 — empty host docx smoke

Verifies the template loads and the ribbon tab renders. **No
navigation is tested here** — the host docx has no Bible content.

1. **Create `aeRibbon-host.docx`** (one-time) per
   `aeRibbon/docx/README_host_docx.md`:
   - Word → File → New → Blank document.
   - Paste this single paragraph as the only content:
     > Attach `aeRibbon.dotm` via File > Options > Add-ins (Manage:
     > Templates → Go → Add), then open a Radiant Word Bible `.docx` to
     > see the **Radiant Word Bible** ribbon tab.
   - File → Save As → Word Document (`*.docx`) →
     `C:\adaept\aeBibleClass\aeRibbon\docx\aeRibbon-host.docx`.
   - Close Word.

2. **Run the smoke check.**
   - Start a fresh Word session (so the Startup-folder template
     re-loads cleanly).
   - Open `aeRibbon-host.docx`.
   - **Expected:**
     - No error dialog.
     - **Radiant Word Bible** tab appears in the ribbon.
     - Pressing **Alt** shows the tab keytip `Y2`.
     - Pressing `Alt, Y2` switches to the tab; controls render
       (selectors visible but disabled — no Bible structure in the docx
       to navigate).
   - Open the VBA editor (Alt+F11) → Immediate window (Ctrl+G). Confirm
     load messages from `RibbonOnLoad` and `AutoExec` appear (these are
     `Debug.Print` statements in the source).

3. **Record the result** in
   `aeRibbon/releases/1.0.0+bc71416/BUILD_RECORD.txt`:
   - Tab appeared: yes/no
   - RibbonOnLoad printed: yes/no
   - AutoExec printed: yes/no
   - Any error dialogs: paste verbatim or "none"

G7 closes when all three "expected" items are met and there are no
error dialogs.

### G8 — production Bible docx navigation smoke

This is the **real navigation test** against an actual Bible `.docx`
produced for this release.

1. **Produce the production Bible `.docx`** (per the next section,
   "Producing the production Bible `.docx`"). This is a manual
   Save-As from the dev `.docm`. Drop the result at
   `C:\adaept\aeBibleClass\aeRibbon\docx\Radiant-Word-Bible.docx`.

2. **Open the docx in a fresh Word session.**
   - **Expected: no macro-security warning.** This is the architectural
     claim of the dotm/docx split — the docx must be code-free. If a
     macro warning appears, the docx is not actually code-free and G8
     fails immediately; fix the Save-As step (Word may have offered
     "macro-enabled" by mistake) before continuing.
   - The **Radiant Word Bible** tab appears (template loaded via
     Startup folder).
   - Book selector is enabled.

3. **Run the navigation checklist** from `aeRibbon/QA_CHECKLIST.md`
   §G8. Highlights:
   - **Mouse path:** click Book, type `Jn`, click Chapter, type `3`,
     click Verse, type `16`, click **Go**. Cursor must land at
     John 3:16.
   - **Keyboard path:** `Alt, Y2, B`, `Jn`, Tab, `3`, Tab, `16`, Tab,
     Enter. Cursor must land at John 3:16.
   - **Next Book:** `Alt, Y2, ]` — advances one book and navigates.
   - **Next Chapter:** `Alt, Y2, .` — advances within current book.
   - **Boundary at Revelation 22:21:** `Alt, Y2, >` shows last-verse
     message in status bar (no nav).
   - **Boundary at Genesis 1:1:** `Alt, Y2, [` shows first-book
     message.
   - **New Search:** `Alt, Y2, S` resets to NoSelection state;
     Chapter/Verse rows disable.
   - **About:** `Alt, Y2, A` opens the About dialog showing
     `RIBBON_VERSION`.

4. **Record results** in
   `aeRibbon/releases/1.0.0+bc71416/BUILD_RECORD.txt` — one line per
   QA_CHECKLIST item, plus the SHA-256 of `aeRibbon.dotm` (`wsl
   sha256sum aeRibbon/template/aeRibbon.dotm`).

5. **Append the release row** to `aeRibbon/RELEASES.md` and tag:
   ```
   git tag v1.0.0+bc71416
   ```

G8 closes when every navigation item passes and the BUILD_RECORD +
RELEASES row + tag are committed.

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
