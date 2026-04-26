# Code Review 2026-04-25

Continuation of `rvw/Code_review 2026-04-21.md`. The prior file remains
the authoritative progressive history through 2026-04-24; this file
carries forward only the live state and adds new work starting today.

---

## § Carry-forward from Code_review 2026-04-21 - state as of 2026-04-25

### Completed (2026-04-21 to 2026-04-24)

- **`basStyleInspector.bas`** module created with five entry points and
  several private helpers:
  - `DumpStyleProperties(name [, writeFile])` - single-style dump,
    paste-ready VBA properties, optional `rpt\Styles\style_<name>.txt`.
  - `DumpAllApprovedStyles` - batch dump in priority order, error-trapped
    per style, integrated orphan-file cleanup with single MsgBox prompt.
  - `ListApprovedStylesByBookOrder([writeFile])` - first-occurrence page
    per approved style across main body, footnotes, endnotes, and
    every section's headers / footers; sorted by `(Page, Priority)`
    ascending; `[not used]` block at end. Optional
    `rpt\Styles\styles_book_order.txt`.
  - `DumpHeaderFooterStyles` - read-only audit of every section x
    header/footer slot to `rpt\Styles\header_footer_audit.txt`.
  - `StartTimer / EndTimer` - session-scoped `Scripting.Dictionary`
    timing pair for "expected (last run) vs actual" feedback.
- **`rpt/Styles/`** subdirectory created; 31+ `style_*.txt` dumps moved
  in via `git mv`, code paths updated.
- **QA workflow established**: ListApprovedStylesByBookOrder output IS
  the canonical priority sequence for the `approved` array in
  `basTEST_aeBibleConfig.bas`. Pages 1-11 walked; array aligned through
  priority 16.
- **Bug fixes applied** (with rationale captured in 2026-04-21):
  - Character-style guard for paragraph-only properties (error 5900).
  - Sort: secondary key on priority for stable tied-page order.
  - StoryRanges walk skips header/footer types 6-11 (Sections walk
    handles them deterministically).
  - Header/footer paragraph iteration with page-1 fallback (Find blind
    to tab-only header content; Information() returns misleading
    section-anchor page).
  - Orphan dump cleanup after rename (`ContentsCPBB` -> `Contents`
    verified 2026-04-24).

### Pending - carried forward

| Item | From | Status |
|---|---|---|
| Walk pages 12+ to align approved array | QA workflow | WIP |
| Identify and add new styles encountered in pages 12+ | QA workflow | WIP |
| Fix or remove `Normal` (priority 14, page 6) | Fix #3 | PENDING decision |
| Decide on `BodyTextIndent` (`[not used]`) | Fix #3 | PENDING |
| Decide on `AuthorQuote` (`[not used]`) | Front matter TBD | DEFERRED |
| `TitleEyebrow` / `Title` formalization | Front matter | PENDING |
| `DefineFrontPageStyles` per book-order findings | Front matter | PENDING |
| Add author styles to `RUN_TAXONOMY_STYLES` | Style spec | PENDING |
| Allowed fonts / fallback fonts / CJK prep | i18n queue | PENDING |
| `SUPER_TEST_RUNS` global verification command | § 6 of 2026-04-21 | DEFERRED until taxonomy stable |
| DOCVARIABLE `UpdatePageNumbers` | Front matter | DEFERRED |

### Operative principles

- **WRIST** - Word ruler is the practical source of truth for indent
  measurements (read off the ruler in UI); `RUN_TAXONOMY_STYLES`
  constants are the code-side source of truth for property values.
- **Book order = priority order** - the approved array reflects
  reading order, not alphabetical or historical.
- **Progressive history** - rvw/ files are dated snapshots; never
  retroactively rewrite earlier sections.
- **ASCII only in VBA** - no em-dashes / en-dashes in `.bas` / `.cls`.
- **Late binding** - all COM objects via `As Object` + `CreateObject`.
- **Identifier casing preserved** - never normalize VBA identifier case.

---

## § EDSG (Editing and Design Style Guide) - plan - 2026-04-25

### Goal

A single web-readable, docx-importable, physically-publishable document
that explains the editing and design conventions of the Study Bible -
the operational manual for any developer or editor (especially for
i18n work) who needs to apply, audit, or extend the style taxonomy.

### Target audience

Primary: a developer / editor onboarding to a translation / localization
of the Study Bible. They need to know:

- What styles exist, what each is for, and where each first appears.
- How to apply, verify, and extend them.
- Which routines to run at each editing stage.
- How the QA process gates a release.

Secondary: future maintainers of the codebase who need an organized
synthesis instead of scrolling through dated `rvw/` files.

### Output channels

| Channel | Format | Notes |
|---|---|---|
| Web | Markdown rendered via GitHub | Source of truth, lives in `/EDSG/` |
| Document | Imported into a `.docx` | Same approved styles as the Study Bible itself |
| Print | PDF / physical book | Uses the Study Bible's template - the EDSG is itself styled like the Bible |

The print form is intentional dogfooding: if the EDSG can be produced
from the same templates and routines as the Bible, the templates and
routines work.

### Folder layout

```
/EDSG/
  README.md                 - landing page, links to all others
  01-styles.md              - approved style taxonomy, current array
  02-editing-process.md     - routines mapped to editing steps
  03-inspection-tools.md    - DumpStyleProperties, DumpAllApprovedStyles, etc.
  04-qa-workflow.md         - book-order workflow, approved-array sync
  05-headers-footers.md     - section / header / footer conventions
  06-i18n.md                - locale considerations, font fallback, RTL/CJK prep
  07-super-test-runs.md     - architectural supervisor (when implemented)
  08-publishing.md          - producing the docx/PDF using Study Bible styles
  09-history.md             - pointers into rvw/ for decision archaeology
```

`README.md` is the single starting page; every other file is reachable
in two clicks.

### Source-of-truth model recorded in EDSG

| Asset | Source of truth | Where in EDSG |
|---|---|---|
| Style property values | `RUN_TAXONOMY_STYLES` constants in code | 01-styles, 03-inspection-tools |
| Indent measurements | Word UI ruler (WRIST principle) | 01-styles, 02-editing-process |
| Approved-list order | `ListApprovedStylesByBookOrder` output vs the array | 04-qa-workflow |
| Decision history | rvw/ files (dated, append-only) | 09-history |
| Current synthesis | EDSG itself | n/a |

### Connection to GitHub

- `/EDSG/` lives in this repo (single source).
- Commits group EDSG updates with the code/style changes that motivated
  them - the commit message and the EDSG diff form a paired audit
  trail.
- GitHub renders markdown directly for browsing.
- Reference commit hashes for major decisions (e.g., "WRIST principle -
  see commit `27136bb`").

---

## § Editing process page - outline for `02-editing-process.md`

This page connects each routine to the editing step it supports. Draft
table of contents:

### Workflow stages

1. **Style design**
   - Decide name (book-order priority position, USFM marker if any)
   - Define properties in `DefineXxxStyle` routine in
     `basFixDocxRoutines.bas`
   - Add to `approved` array in `basTEST_aeBibleConfig.bas`
   - Add expected values to `RUN_TAXONOMY_STYLES`

2. **Apply style in document**
   - Manual: Word Styles pane / shortcut
   - Read indent values off the ruler (WRIST)

3. **Single-style audit**
   - `DumpStyleProperties "<Name>"` - paste-ready dump to Immediate
   - `DumpStyleProperties "<Name>", True` - also writes
     `rpt\Styles\style_<Name>.txt`
   - QA checklist (4 properties: BaseStyle, AutomaticallyUpdate,
     QuickStyle, LineSpacingRule) - see 2026-04-21 § for the table

4. **Bulk audit + cleanup**
   - `WordEditingConfig` to repromote priorities
   - `DumpAllApprovedStyles` - writes all current styles, prompts to
     delete orphans (e.g., after renames)

5. **Order verification**
   - `ListApprovedStylesByBookOrder` - generates canonical book order
   - Compare to current `approved` array; reorder array to match

6. **Header / footer changes**
   - `DumpHeaderFooterStyles` - audit every section x slot
   - Identify Linked-to-Previous chains and unlinked anchors

7. **Pre-commit gate**
   - `SUPER_TEST_RUNS` (when implemented per § 6 of 2026-04-21)

### Per-routine quick reference

For each routine: signature, what it produces, when to run it, sample
output snippet, related routines.

### Anti-patterns / gotchas

- Don't apply styles via direct formatting + `AutomaticallyUpdate=True`
  (silent rewrite of the style def for everyone).
- Don't reorder the `approved` array manually without re-running
  `ListApprovedStylesByBookOrder` to verify.
- Don't delete `rpt/Styles/` files by hand; let the orphan-cleanup
  prompt do it.

---

## § SUPER_TEST_RUNS as architectural QA supervisor

Deferred per § 6 of 2026-04-21 (status: implement after taxonomy is
stable). EDSG ties in at the final stage:

- `07-super-test-runs.md` documents the master report and how each
  suite contributes to the release-readiness signal.
- Each EDSG style/process page links forward to the SUPER_TEST_RUNS
  suite that validates it (e.g., 01-styles links to "Suite 1: Style
  taxonomy").
- `SuperTestReport.txt` is the canonical pre-publication health
  snapshot. A clean run is the release gate for any locale build.
- Suite sequence (carried forward from § 6): Style taxonomy -> Document
  diagnostics -> Font audit -> Header/footer audit -> Scripture parser.

The architectural framing in EDSG: every editing routine has a *human*
workflow (how to use it) and a *test* anchor (what catches its
mistakes). SUPER_TEST_RUNS is where those test anchors converge.

---

## § Pros / Cons / Benefits

### Pros

- **Single onboarding doc** for i18n editors instead of forensic
  reading of `rvw/` history.
- **Dogfooding** - producing the EDSG with the Bible's own styles
  exercises and validates the templates.
- **Synthesis surface** - EDSG can present the *current* state cleanly
  while rvw/ keeps the *historical* state intact.
- **i18n unblocking** - future translators need a stable doc anchored
  to current state, not a 2000-line dated review.
- **Forcing function** for naming and concept consistency - writing for
  a new reader exposes ad-hoc terminology.
- **Same-style requirement** for the printed EDSG forces issues out of
  the woodwork: missing styles, edge cases in page layout, whatever.

### Cons

- **Doc maintenance overhead** while styles still in flux (pages 12+
  not yet walked, several decisions pending).
- **Drift risk** - if EDSG falls out of sync with code, it becomes
  worse than no doc.
- **Time split** - parallel work means each session has to choose
  between progress on styles vs progress on docs.
- **Some sections deferred** - `07-super-test-runs.md` can't be filled
  in until SUPER_TEST_RUNS exists; readers will see "coming soon"
  placeholders.

### Benefits

- Captures decisions while fresh (vs archeology in 6 months).
- Establishes the single-source-of-truth map (which artifact is
  authoritative for which question).
- Creates the i18n entry point on day 1 of the localization effort.
- Gives the project a publishable artifact independent of the Bible
  itself - the methodology is shareable.

---

## § Cost - implementation now vs later

### Now (start parallel to style work)

- **Setup**: ~2-3 hours (folder, README scaffold, 9 file skeletons,
  initial content for mature areas: 01-styles, 03-inspection-tools,
  04-qa-workflow).
- **Ongoing**: ~20-30 min per session that touches styles, slotting an
  EDSG update next to the rvw/ entry.
- **Risk**: low - in-flux sections explicitly marked, mature sections
  stable.

### Later (after taxonomy fully stable)

- **Backfill**: ~10-15 hours of concentrated writing once style work
  freezes - reconstructing rationale, re-reading rvw/ progression,
  re-running routines to capture sample output.
- **Risk**: medium-high - decisions blur in memory; "we'll do docs
  later" has a known failure mode of never happening; i18n work
  blocks waiting for docs.

### Recommendation

**Start now, incrementally.** Concretely:

1. Create `/EDSG/README.md` and 9 file skeletons in this commit cycle.
2. Fill in the **mature** sections first - 01-styles, 03-inspection-tools,
   04-qa-workflow - they reflect work already done.
3. Mark in-flux sections clearly (`Status: WIP - taxonomy walk through
   page 11 only; pending pages 12+`).
4. Update EDSG entries alongside future style work, paired with
   matching rvw/ entries.
5. `07-super-test-runs.md` stays a placeholder until § 6 of 2026-04-21
   moves out of DEFERRED.

---

## § Open questions before writing actual `/EDSG/*.md`

(Not blockers for the plan; would help shape the content.)

1. First localization target - which language / script first? Affects
   what i18n.md prioritizes (RTL, CJK, accented Latin).
2. Print template for the EDSG - same Study Bible `.docm` template, or
   a derivative? Affects how 08-publishing.md is structured.
3. Visual conventions in the markdown - call-out blocks, code samples,
   embedded screenshots? GitHub-flavored MD admonitions or plain?
4. Audience tone - terse engineer-style or more narrative for editors
   who may not be developers?

### Status

EDSG plan: **DRAFT - awaiting user review**.
Skeleton creation: **PENDING approval**.
Initial mature-section content: **PENDING approval**.

---

## § EDSG scaffolding - 2026-04-25

### Done

`/EDSG/` directory created with 10 files:

| File | Status | Notes |
|---|---|---|
| `README.md` | Complete | Landing page, source-of-truth map, page index, operative principles |
| `01-styles.md` | Mature - WIP marker on pages 12+ | Approved style snapshot, categories, QA checklist |
| `02-editing-process.md` | Mature | 7 workflow stages, anti-patterns, per-routine quick-ref |
| `03-inspection-tools.md` | Mature | Full reference for `basStyleInspector` public + private API |
| `04-qa-workflow.md` | Mature | Book-order canonical-priority workflow, 5-step cycle |
| `05-headers-footers.md` | WIP | What's known after the audit; gotchas (Headers(1) vs Headers(2), Find blindness, Information() misleading) |
| `06-i18n.md` | Skeleton | Awaits first-locale decision; current font inventory |
| `07-super-test-runs.md` | Placeholder | Recaps the deferred design from `rvw/Code_review 2026-04-21.md` § 6 |
| `08-publishing.md` | Skeleton | Three output forms (web/docx/print); markdown-to-style mapping draft; open questions |
| `09-history.md` | Mature | Pointers into `rvw/`; decision-archaeology table; significant commits |

### Decisions during scaffolding

- ASCII hyphens used throughout EDSG markdown for consistency with
  recent `rvw/` content (memory permits em-dashes in markdown but
  uniformity wins).
- Cross-links use relative paths (`01-styles.md`) — GitHub renders
  them; same paths work for any docx/PDF importer that follows links.
- Decision archaeology table in `09-history.md` provides a question →
  rvw-section index, so future readers don't need to scan the full
  rvw/ history to find why a thing is the way it is.
- Inline file-status badges in `README.md` and the in-review table
  above mark every page as Mature / WIP / Skeleton / Placeholder so
  readers know what to trust.

### Pending follow-ups

- Answer the four open questions (first locale; print template;
  visual conventions; tone) before further fleshing out
  `06-i18n.md`, `08-publishing.md`.
- Choose markdown -> docx import path (likely Pandoc + style
  mapping); detail in `08-publishing.md` once chosen.
- Define a `CodeBlock` paragraph style and an inline-code character
  style — needed for `08-publishing.md` to map markdown code
  constructs.
- Replace `07-super-test-runs.md` placeholder content with
  operational documentation when SUPER_TEST_RUNS lands.
- Re-run `ListApprovedStylesByBookOrder` after pages 12+ walk; refresh
  the snapshot table in `01-styles.md`.

### Status

**SCAFFOLDED - 2026-04-25**. Mature pages (`02`, `03`, `04`, `09` and
`README.md`) are usable now; `01` is current through page 11; `05`,
`06`, `07`, `08` carry visible status markers.

---

## § Book ComboBox sizing - "2 Thessalonians too short" - 2026-04-25

### Symptom

Code picks `"2 Thessalonians"` as the longest book name to size the
ComboBox, but at runtime the combo clips the last few characters of
that very value (and others). User noted "the combo shows caps or
the dropdown is taking space."

### Two effects, both contributing

**1. Character count is a poor proxy for rendered width.**

ComboBoxes draw text in a proportional font. `Len("2 Thessalonians") =
15` ties with `"Song of Solomon"` and `"1 Thessalonians"`, but the
glyphs differ widely:

- `"Thessalonians"` is dominated by narrow letters (`i`, `l`, `s`,
  `t`, `n`).
- `"Song of Solomon"` has wide round letters (`o`, `g`, `S`, `m`)
  plus an extra space.
- All-caps versions add even more variance: `"OBADIAH"` (7 chars)
  rendered uppercase can rival a 15-char title-case string in pixel
  width.

The alias map in `aeBibleCitationClass.cls` stores keys uppercase
(`"1 THESSALONIANS"`, `"2 THESSALONIANS"`, etc.). If display anywhere
in the pipeline uses the uppercase form, rendered width exceeds the
length-based estimate.

**2. Dropdown arrow chrome eats text area.**

A Word ComboBox is two regions side by side: editable text area and
the drop-down arrow button. The arrow is fixed chrome (~17 px on a
standard form). If `Combo.Width` is set to the exact text width of
`"2 Thessalonians"`, the arrow overlays the last 1-2 characters.

### Recommended fix (analysis only - not yet implemented)

Both corrections needed; either alone is insufficient.

1. **Measure rendered width**, not `Len()`. Options:
   - Hidden `Label` control with `AutoSize = True`, looped over all
     book names in the casing actually displayed, take `MaxOf
     Label.Width`. Use the same font as the ComboBox so the
     measurement transfers exactly.
   - `TextWidth` via the Office drawing canvas.
2. **Add chrome padding**. After computing widest text width, add
   ~20 form units (about 24 pixels at standard scale) for the
   dropdown arrow plus a couple pixels of breathing room.

### Open follow-ups

- Locate the routine currently setting the combo width. Searched for
  `ComboBox`, `cboBook`, `longest`, `Width`, `MeasureString` in
  `src/`; no obvious owner found. May live in a UserForm code-behind
  not surfaced via grep on `.bas` / `.cls` (UserForm `.frm` /
  `.frx` files).
- Confirm whether the displayed casing is title-case or upper-case;
  shapes the candidate-string set used for measurement.

### Status

**ANALYSIS - 2026-04-25**. Awaiting confirmation of the owning
routine before applying the two-step fix.

### Update - it's RibbonX, not a UserForm - 2026-04-25

User identified the owning XML: a `<comboBox>` in `customUI14.xml`
with attribute `sizeString="2 Thessalonians"`. This is Ribbon XML,
not a Word UserForm.

Two changes to the earlier analysis:

- **The dropdown-arrow chrome is NOT an issue.** The Ribbon engine
  handles chrome itself when sizing from `sizeString`. Only effect
  remains is the proportional-font glyph variance.
- **`Width` cannot be set programmatically** for a ribbon comboBox -
  `sizeString` is the only knob. Replace the value, no code-side
  measurement loop needed.

Three combos found in `customUI14backupRWB.xml`, all using
`sizeString="2 Thessalonians"`:

- `cmbBook`
- `cmbChapter`
- `cmbVerse`

Earlier review (`Code_review - 2026-04-10a.md`) confirms all three
were sized identically on purpose for ribbon-row alignment. They
must change together.

### Fix applied

Replaced `sizeString="2 Thessalonians"` with
`sizeString="Song of Solomon"` on all three combos. `"Song of
Solomon"` measures wider in Segoe UI (round letters `o`, `S`, `g`,
`m` plus extra space) despite the same character count.

### Caveat - embedded XML in the .docm

The repo file is the tracking copy. The runtime ribbon reads the
customUI XML embedded inside the `.docm` package
(`/customUI/customUI14.xml` inside the zip). The repo edit alone
does not change the live ribbon - the same change must be pushed
into the embedded XML (e.g., via Office RibbonX Editor) before the
runtime sizing changes.

### Status

**APPLIED in repo - 2026-04-25**. Embedded `.docm` customUI update
**PENDING user action**.

### Embedded `.docm` updated via `py/inject_ribbon.py` - 2026-04-25

Ran:

```
wsl python3 py/inject_ribbon.py
```

Output: `REPLACED  customUI/customUI14.xml / Done. Blank Bible Copy.docm
updated.` Verified by re-opening the package and counting tokens:
3x `Song of Solomon`, 0x `2 Thessalonians`.

This is the working path for ribbon-XML updates in this project -
RibbonX Editor has a known load bug for this file, so the Python
injector is the sanctioned tool. Default target is
`Blank Bible Copy.docm`; pass an alternate path as the first argument
to update other `.docm` files (e.g., the `Peter-USE REFINED English
Bible CONTENTS.docm` working copy when ready).

### Status (final for this fix)

**APPLIED - 2026-04-25** in both `customUI14backupRWB.xml` (repo
tracking copy) and `Blank Bible Copy.docm` (runtime). Combo width
will now match the actual widest English book name. Locale-specific
overrides (other languages, all-caps display) flagged for future
i18n consideration.

---
