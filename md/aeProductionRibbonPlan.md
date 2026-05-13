# aeProductionRibbon — Production Export Plan

**Status:** APPROVED (§7 items 1–5) — ready to execute §8 on user "go".
No files copied or generated yet.
**Date:** 2026-05-12
**Author:** Claude (Opus 4.7) — for review by peterennis
**Scope:** Establish `aeRibbon/src/` as the production export gateway and
`aeRibbon/template/` as the Word 365 `.dotm` delivery vehicle for the
Radiant Word Bible navigation ribbon.

---

## 1. Goals (from the request)

1. Populate `C:\adaept\aeBibleClass\aeRibbon\src\` with **only** the files
   needed to run the Bible ribbon.
2. Keep routines/classes that are necessary; remove the rest from the copies.
3. **Do not modify any file under `C:\adaept\aeBibleClass\src\`** — that
   folder remains the development source of truth. All trimming happens on
   the copies inside `aeRibbon/src/`.
4. Maintain a log of every routine kept and every routine removed.
5. Treat `aeRibbon/` as the **production export gateway** — the only path
   from dev source to a shippable Word template.
6. Produce a Word 365 `.dotm` template + a `.docx` end-user document that
   pairs with it. Template contains no Bible text.
7. Bible text stays in the existing `.docm` (e.g. `Peter-USE REFINED English
   Bible CONTENTS.docm`); the template attaches to it.
8. This file — captures the plan.
9. QA release process built on the existing `basTEST_*` harness.
10. Versioning + tracking scheme rooted in dev-side identifiers.

---

## 2. Production scope — files in / files out

### 2.1 Hard dependencies of the running ribbon

Verified by grepping callback names in `customUI14backupRWB.xml` against
`Public Sub` declarations, and chasing class references inside those
files:

| File (dev `src/`) | Role | Action for `aeRibbon/src/` |
|---|---|---|
| `customUI14backupRWB.xml` (at repo root) | Ribbon XML; declares all callbacks | **COPY** as `aeRibbon/template/customUI14.xml` |
| `basBibleRibbonSetup.bas` | All `onAction` / `getEnabled` / `getLabel` / keytip callbacks; singleton `Instance()`; `AutoExec` | **COPY, TRIM** |
| `aeRibbonClass.cls` | State machine, navigation, status-bar messages | **COPY, TRIM** |
| `aeBibleClass.cls` | Document model: chapter/verse scan, ScrollIntoView, focus restore | **COPY, TRIM** |
| `aeBibleCitationClass.cls` | Canonical book table, alias resolution, verse counts | **COPY, TRIM** |
| `basRibbonDeferred.bas` | `OnTime`-style deferred chapter nav + status bar | **COPY AS-IS** (tiny) |
| `basUIStrings.bas` | Centralised user-visible strings (i18n hook) | **COPY AS-IS** (tiny) |
| `ThisDocument.cls` | Document-open wiring (if it loads the ribbon) | **COPY** (verify first) |

### 2.2 Explicitly excluded (development-only)

These never enter `aeRibbon/src/`:

- All `basTEST_*.bas` (8 files, ~6,800 lines) — test harness stays in dev.
- All `X*` prefixed test slow-runners.
- `basAuditDocument.bas`, `basVerseStructureAudit.bas`, `basStyleInspector.bas`,
  `basAuthorStyles.bas`, `basFixDocxRoutines.bas`, `basWordRepairRunner.bas`,
  `basWordSettingsDiagnostic.bas` — author/maintainer tooling.
- `basChangeLog_*.bas` — historical change logs.
- `basImportWordGitFiles.bas`, `aeWordGitClass.cls` — dev import pipeline.
- `basUSFM_Export.bas`, `basSBL_VerseCountsGenerator.bas` — generators.
- `basBibleRibbon_OLD.bas`, `bas_TODO.bas`, `Module1.bas` — legacy / scratch.
- `aeAssertClass.cls`, `aeLoggerClass.cls`, `aeLongProcessClass.cls`,
  `aeUpdateCharStyleClass.cls`, `IaeLongProcessClass.cls` — only referenced
  from test/maintenance entry points. Confirmed: aeBibleClass/Citation/Ribbon
  classes contain `aeAssert.*` only inside `Sub Test_*` routines that will
  be removed during trimming.

### 2.3 What "TRIM" means

For each file marked TRIM:

1. Identify every `Public`/`Private` `Sub`/`Function`/`Property` inside it.
2. Build a call graph rooted at the **ribbon callback set** declared in
   `customUI14backupRWB.xml` (the 50+ `Get*` / `On*` entry points listed in
   `basBibleRibbonSetup.bas`) plus `AutoExec`, `RibbonOnLoad`,
   `Document_Open` (if present in `ThisDocument`).
3. Anything not transitively reachable from those roots is **removed** from
   the production copy.
4. Specifically: every `Sub Test_*`, `Sub RUN_*`, `Sub DEBUG_*`, `Sub PURGE_*`
   that is not wired into ribbon XML gets stripped.
5. Each removal is appended to the routine log (see §3).

No identifier casing changes, no signature changes, no semantic edits.
Trim = delete-only. (Per `[[feedback_casing]]` and the dev-source-is-truth
principle.)

---

## 3. Routine log — format

A single file: `aeRibbon/RoutineLog.md`. One row per routine considered.

```markdown
| File | Routine | Decision | Reachable from | Notes |
|---|---|---|---|---|
| aeRibbonClass.cls | GoToVerseByScan | KEPT | OnGoClick, OnNextVerseClick, ... | core nav |
| aeRibbonClass.cls | DEBUG_DumpState | REMOVED | (none) | dev-only |
| aeBibleCitationClass.cls | Test_ParseToken | REMOVED | (none) | test |
| aeBibleCitationClass.cls | GetCanonicalBookTable | KEPT | aeRibbonClass.OnBookChanged | book list source |
```

Generated by a one-shot Python pass (`py/ribbon_export_trim.py`, new) that:

1. Tokenises each `.bas`/`.cls` into routines (regex on `^(Public|Private)?\s*(Sub|Function|Property)\s+\w+`).
2. Walks identifier references to build the call graph.
3. Writes the trimmed file copies under `aeRibbon/src/` + the log row.
4. Is **idempotent** — re-running produces identical output for the same
   dev `src/`.

This is preferable to a hand edit: it can be re-run on every dev sync, and
the diff in `aeRibbon/src/` will only reflect dev-side semantic change.

---

## 4. The Word 365 template + companion docx

### 4.1 Layout

```
aeRibbon/
├── src/                      # trimmed VBA — production export gateway
├── template/
│   ├── customUI14.xml        # copy of customUI14backupRWB.xml
│   ├── aeRibbon.dotm         # built template (Word 365)
│   └── README_template.md    # build instructions
├── docx/
│   └── aeRibbon-host.docx    # placeholder host doc (no Bible text)
├── RoutineLog.md             # generated by export script
├── VERSION                   # semantic version, e.g. 1.0.0
└── BUILD.md                  # how to build & install
```

### 4.2 What the template is and isn't

- The `.dotm` is a **macro template** — it carries the VBA project and the
  customUI14 ribbon XML, but no Bible content.
- Attaching the template to the existing
  `Peter-USE REFINED English Bible CONTENTS.docm` (Tools → Templates and
  Add-ins → Document template) makes the **Radiant Word Bible** ribbon tab
  appear in Word whenever that document is open.
- The companion `.docx` is an empty host showing the ribbon works against a
  fresh document — useful for QA and demos. The ribbon will only navigate
  if the open document contains the canonical Heading-1 book / Heading-2
  chapter / verse-style structure; the empty docx is for **install/load
  verification**, not navigation tests.

### 4.3 Build steps (will be encoded in `aeRibbon/BUILD.md`)

1. Create a blank `.dotm` in Word 365.
2. Open the VBA editor; import every `.bas` and `.cls` from `aeRibbon/src/`.
3. Inject ribbon XML via the existing pipeline:
   `wsl python3 py/inject_ribbon.py --target aeRibbon/template/aeRibbon.dotm
   --xml aeRibbon/template/customUI14.xml`
   (per `[[feedback_ribbon_injector]]` — never RibbonX Editor for this
   project.)
4. Save. Compute `VERSION` + git SHA stamp (see §6).
5. Sign the template if a code-signing cert is available (optional for now).

---

## 5. QA release process

Built on the existing `basTEST_*` harness — those modules stay in dev `src/`
and are run against the dev `.docm`, **not** against `aeRibbon/`.

### 5.1 Gate sequence

| Gate | Where it runs | What it checks |
|---|---|---|
| G1 — Unit | dev `.docm`, VBA `RUN_*` entry points in basTEST_aeBibleClass / basTEST_aeBibleCitationClass | Citation parse, book/chapter/verse table integrity |
| G2 — Citation block | `basTEST_aeBibleCitationBlock` | Citation rendering |
| G3 — Config / styles | `basTEST_aeBibleConfig` | Style taxonomy intact |
| G4 — Tools | `basTEST_aeBibleTools` | Document tool surface |
| G5 — Export trim | `python3 py/ribbon_export_trim.py --check` | No unreachable routine survived; no reachable routine dropped |
| G6 — Template build | manual or scripted | `.dotm` builds, ribbon XML injects without error |
| G7 — Smoke (host docx) | Word 365, open `aeRibbon-host.docx` | Tab appears, controls render, AutoExec/RibbonOnLoad fire, `RibbonOnLoad` errors == 0 |
| G8 — Smoke (real Bible docm) | Word 365, open `Peter-USE REFINED English Bible CONTENTS.docm` with template attached | Type `Jn`, Tab, `3`, Tab, `16`, Tab, Enter → cursor lands at John 3:16; KeyTip path `Alt, Y2, B/C/V/G` works; Prev/Next at boundaries shows status-bar message |

G1–G5 are automatable / re-runnable; G6–G8 are human-in-the-loop and
checklisted in `aeRibbon/QA_CHECKLIST.md`.

### 5.2 Release artefacts

Each release produces, under `aeRibbon/releases/<version>/`:

- `aeRibbon-<version>.dotm`
- `RoutineLog.md` (snapshot)
- `BUILD_RECORD.txt` — git SHA, build date, gates passed, gate operator
- SHA-256 of the `.dotm`

---

## 6. Versioning and tracking

### 6.1 Version scheme

`MAJOR.MINOR.PATCH+<git-short-sha>` — semantic, with the dev-source SHA
appended for traceability. Example: `1.0.0+48fd425`.

- **MAJOR** — ribbon XML schema change or breaking callback signature change.
- **MINOR** — new ribbon control, new keytip, new public callback.
- **PATCH** — bug fix in `aeRibbonClass` / `aeBibleClass` /
  `aeBibleCitationClass` that does not change callback surface.

### 6.2 Where the version is stored

- `aeRibbon/VERSION` — plain text, single line.
- `basBibleRibbonSetup.bas` — `Public Const RIBBON_VERSION As String = "..."`
  injected by the export script, surfaced through `OnAdaeptAboutClick`
  (About dialog).
- `aeRibbon/template/aeRibbon.dotm` — custom document property
  `aeRibbonVersion`.

### 6.3 Tracking mapping (dev SHA → production build)

`aeRibbon/RELEASES.md` — append-only table:

```markdown
| Version | Date | Dev SHA | Built from src/ | Gates passed | Notes |
|---|---|---|---|---|---|
| 1.0.0+48fd425 | 2026-05-12 | 48fd425 | src/ @ 48fd425 | G1-G8 | initial export |
```

This satisfies the "comprehensive versioning and tracking" item — any
production `.dotm` in the field can be tied back to an exact dev-source SHA
via its About-dialog version string.

---

## 7. Decisions (resolved 2026-05-12)

1. **Trim approach — APPROVED: automated.**
   `py/ribbon_export_trim.py` (new) builds the call graph from the
   `customUI14backupRWB.xml` callback roots, writes trimmed copies into
   `aeRibbon/src/`, and appends to `RoutineLog.md`. Idempotent and
   re-runnable on every dev sync.

2. **`ThisDocument.cls` — APPROVED: build-step, not a copy-target.**
   `ThisDocument` is not freely importable. The production `.dotm`
   keeps its own `ThisDocument`. `BUILD.md` documents the (currently
   empty) `Document_Open` body to paste in if/when one is needed.
   No file at `aeRibbon/src/ThisDocument.cls`.

3. **Code signing — APPROVED: deferred for v1.0.0.**
   Users enable macros via Word Trust Center on first open. Signing is
   a candidate for a later MINOR release if a cert becomes available.

4. **Companion docx — APPROVED: `aeRibbon-host.docx`.**
   Empty Word document, one instruction paragraph
   ("Attach `aeRibbon.dotm` via File > Options > Add-ins, then open
   your Bible `.docm` to see the Radiant Word Bible tab.").
   No Bible text. Used for G7 install/load smoke only.

5. **AutoExec vs. Document_Open — APPROVED: keep `AutoExec`, no
   `Document_Open` promotion.**

   Rationale (full pros/cons in conversation 2026-05-12):
   - `RibbonOnLoad` is the load-bearing functional entry point — Word
     always fires it when the customUI14 XML loads. That stays
     unchanged.
   - `AutoExec`'s current job is a diagnostic pre-warm of the
     singleton via `Instance()`. The singleton is lazy-instantiated
     by every callback anyway, so `AutoExec` is harmless and gives a
     deterministic `Debug.Print` timestamp for load-order debugging.
   - `Document_Open` placed in the **template's** `ThisDocument` does
     **not** fire for host documents — only when the `.dotm` itself
     opens. For the attached-template deployment model (the expected
     production install path) it would silently no-op.
   - Per-document logic, if ever needed, belongs in the **host
     `.docm`'s** `ThisDocument` as a future MINOR change, not in the
     template.

---

## 8. Proposed execution order (after approval)

1. Write `py/ribbon_export_trim.py` + first run → populates `aeRibbon/src/`
   and `aeRibbon/RoutineLog.md`.
2. Copy `customUI14backupRWB.xml` → `aeRibbon/template/customUI14.xml`.
3. Write `aeRibbon/BUILD.md`, `aeRibbon/QA_CHECKLIST.md`,
   `aeRibbon/RELEASES.md` (empty table).
4. Set `aeRibbon/VERSION` → `1.0.0+<sha>`.
5. **User-side step** — build the `.dotm` in Word 365 per `BUILD.md`
   (Claude cannot run Word interactively).
6. Run G1–G8; record results; commit `aeRibbon/releases/1.0.0+<sha>/`.

---

## 9. Memory notes

This plan honours the standing project rules:
- `[[feedback_casing]]` — no identifier-casing changes.
- `[[feedback_late_binding]]` — no new project references added.
- `[[feedback_ascii_in_vba]]` — ASCII-only in any new VBA we touch.
- `[[feedback_ribbon_injector]]` — XML injected via `py/inject_ribbon.py`.
- `[[feedback_session_manifest]]` — every session that produces files under
  `aeRibbon/src/` writes `sync/session_manifest.txt`.
- `[[feedback_code_review_process]]` — this document is the proposal step;
  no source files are copied or generated until the user says go.
