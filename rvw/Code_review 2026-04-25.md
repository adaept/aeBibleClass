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

### Round 3 - case mismatch revealed - 2026-04-25

`"Song of Solomon"` still didn't fit. Root cause: the displayed
text in the combo is taken from `headingData(idx, 0)` via
`GetBookText`, which is the literal `Heading 1` paragraph text
captured by `CaptureHeading1s`. The document's H1 paragraphs are
**uppercase** ("GENESIS", "EXODUS", etc.), so the actual rendered
text in the combo is uppercase, not title case. Title-case
`"Song of Solomon"` underestimated the rendered width.

#### Re-analysis - widest in uppercase

In Segoe UI uppercase the relative letter widths shift compared to
title case. Reasoning:

- `SONG OF SOLOMON` - 15 chars, 5 wide `O`s, an `M`, but title
  case favored it because of lowercase `o` runs that disappear in
  uppercase.
- `2 THESSALONIANS` - 15 chars, leading `2` (wider than `1` and
  wider than a space), wide opening `T H`, double `S S`, long
  medium-wide tail `A L O N I A N S`.
- `1 THESSALONIANS` - 15 chars, narrower opening because `1` is
  thinner than `2`.

User's on-screen test confirmed `"2 THESSALONIANS"` is the actual
widest in the ribbon's font - the leading digit `2` plus `T H`
plus double `S S` outweighs the round-letter advantage of
`SONG OF SOLOMON` once everything is uppercase.

#### Fix applied

`customUI14backupRWB.xml` updated: all 3 `sizeString` attributes
now `"2 THESSALONIANS"`. Verified count = 3, no other variants
remain. `wsl python3 py/inject_ribbon.py` ran successfully and the
embedded XML inside `Blank Bible Copy.docm` confirmed via re-read.

#### Lesson for the i18n queue

`sizeString` must match the **case actually displayed**, not just
the longest-by-character-count from a canonical book list. For
locales where the displayed text is mixed case, retest with the
mixed-case rendering. For all-caps locales, retest with all-caps
candidates.

This is now noted in `EDSG/06-i18n.md` queue and `08-publishing.md`
draft.

#### Status

**APPLIED - 2026-04-25** in both repo file and runtime `.docm`.
Awaiting user re-verification.

---

## § Stale heading cache in ribbon navigation - 2026-04-25

### Symptom

Reproducible sequence:

1. Search "GEN" in the ribbon - finds Heading 1 Genesis (correct).
2. Click Next book button - jumps to the section break before Exodus,
   not Exodus itself.
3. (Earlier in the session, an empty paragraph in `DatAuthRef` style
   was added between books - any insertion is enough.)
4. Search "ROM" then Go - lands on what looks like verse 1 of Romans
   (the chapter-reset is by design; see "Not a bug" below).

User hypothesis: navigation works on a clean document but breaks
after edits because heading data is stale. **Confirmed.**

### Root cause

`src/aeRibbonClass.cls` caches heading positions in three class-level
arrays:

- `headingData(1..66, 0..1)` - book name + char position of each
  Heading 1.
- `chapterData(1..66, 1..150)` - char position of each Heading 2.
- `m_currentBookPos` / `m_currentChapterPos` - derived.

`CaptureHeading1s` populates them. The body was gated by:

```vba
Static hasRun As Boolean
...
If hasRun Then GoTo PROC_EXIT
```

`Static` persists for the lifetime of the class instance, which lives
until Word closes. So `CaptureHeading1s` ran exactly once per session
and never refreshed. After any edit, every cached char position
downstream of the edit was N characters off, sending navigation into
section breaks or unrelated content.

The `If IsEmpty(headingData(1, 0)) Then CaptureHeading1s` guards at
lines 226 and 534 looked like rescan triggers but never fired - the
arrays were only "empty" if the class instance was destroyed, not
when the document was edited.

### Not a bug - verse 1 after Go

`GoToChapter` (line 622) sets `m_currentVerse = 1` deliberately
("Rule 2a"). Go-to-chapter resets verse to 1; the user types a
different verse explicitly to navigate elsewhere. Symptom #4 in the
report is intentional behavior, not part of the cache bug.

### Remediation - Option 2 (saved-flag invalidation)

Selected over Option 1 (always rescan, simplest but slowest) and
Option 3 (manual Refresh button, easily forgotten).

Use `ActiveDocument.Saved` as a freshness signal. The cache is valid
only when BOTH the previous scan AND the current state report
`Saved = True` - that combination guarantees no edits could have
happened in between. Any other combination forces a rescan.

#### Code changes (`src/aeRibbonClass.cls`)

1. New class-level state: `Private m_lastScanWasSaved As Boolean`.
2. Initialized to `False` in `Class_Initialize`.
3. `CaptureHeading1s` updated:
   - Now accepts `Optional ByVal bForce As Boolean = False` for an
     explicit override (e.g., a future Refresh button).
   - Static `hasRun` retained.
   - Cache-valid check replaces the unconditional gate:
     ```vba
     If hasRun And Not bForce And m_lastScanWasSaved And ActiveDocument.Saved Then
         Debug.Print "CaptureHeading1s: cache valid (no edits since last scan)."
         GoTo PROC_EXIT
     End If
     ```
   - `Erase headingData` / `Erase chapterData` before rescanning so a
     scan that finds fewer H1/H2 entries does not leave stale tail
     data.
   - Records the saved-state at end: `m_lastScanWasSaved =
     ActiveDocument.Saved`.

No call-site changes - the routine is still safely re-callable via
the existing `IsEmpty(headingData(1, 0)) Then CaptureHeading1s`
guards and via direct calls in nav paths.

#### Behavior trace

| State | `Saved` | `m_lastScanWasSaved` | Action |
|---|---|---|---|
| Doc opened (clean) | True | n/a (first call) | Scan; record `True`. |
| Read-only nav | True | True | Skip - cache valid. |
| User edits | False | True | Next nav rescans; record `False`. |
| More navs (still unsaved) | False | False | Each nav rescans (cost during editing - acceptable). |
| User saves | True | False | Next nav rescans; record `True`. |
| Read-only nav resumes | True | True | Skip again. |

### Limitation

During an editing burst, every navigation triggers a rescan because
`m_lastScanWasSaved = False` until the user saves. On a Bible-sized
document this is a couple of seconds per Next/Prev/Go. If that proves
painful, layer on a paragraph-count signature or hook
`Document.ContentControlOnEnter` / `Application.WindowSelectionChange`
for finer invalidation - flagged for later, not done now.

### Status

**APPLIED - 2026-04-25** in `src/aeRibbonClass.cls`. Awaiting
verification re-run with the original repro sequence.

### Verification re-run uncovered second bug - 2026-04-25

After the saved-flag fix, user reported:

> After some editing then nav forward with Next it starts to advance
> the cursor into the title of each next book. Prev will track back,
> with the cursor in the same offset, until it hits the region where
> it is good again and is at the start of each Heading 1.

Same symptom (stale positions, off by a constant offset downstream
of the edit), even with the saved-flag invalidation in place.

#### Second root cause - call sites bypass the staleness check

`CaptureHeading1s` got smarter about *when* to rescan, but it was
only being **called** in three places, two of which guarded the
call:

| Line | Context | Calls? |
|---|---|---|
| 196 | `EnableButtonsRoutine` (startup) | Unconditional |
| 228 | `GoToH1` (search book by name) | `If IsEmpty(headingData(1, 0)) Then CaptureHeading1s` - skipped when cache holds (stale) data |
| 549 | `OnBookChanged` (combo selection) | Same `IsEmpty` guard - same skip |
| `NextButton` | Next book | **Never called** |
| `PrevButton` | Prev book | **Never called** |
| `GoToChapter` / `FindChapterPos` | Chapter nav | **Never called** |

So `Next` / `Prev` / chapter navigation read `headingData` /
`chapterData` directly without ever consulting the staleness check.
After an edit, those paths happily used the wrong cached positions.

#### Second fix

Two targeted changes in `src/aeRibbonClass.cls`:

1. **Drop the `IsEmpty` guards** at lines 228 (`GoToH1`) and 549
   (`OnBookChanged`). Replace each with an unconditional
   `CaptureHeading1s` call. The routine's own
   `m_lastScanWasSaved And ActiveDocument.Saved` check decides
   whether to rescan; the cache-empty case still triggers a scan
   (`hasRun = False`).
2. **Add `CaptureHeading1s` at the top of `NextButton`,
   `PrevButton`, and `GoToChapter`**. One line each, with a comment
   noting it is a cheap no-op when the cache is valid.

Pattern at every nav entry point now reads:

```vba
On Error GoTo PROC_ERR
CaptureHeading1s     ' refresh if doc was edited; cheap no-op otherwise
' ... existing nav body reads headingData / chapterData ...
```

#### Why this is the right pattern

The "is the cache empty?" question becomes the routine's concern,
not the caller's. Every nav path becomes uniform: call
`CaptureHeading1s`, then read the data. No more "did I remember to
guard / unguard the call?" footguns.

#### Tradeoff

Adds one `Debug.Print` line per nav click in the cached-valid case
(`CaptureHeading1s: cache valid (no edits since last scan).`).
Cosmetic noise in the Immediate window; can be silenced after the
fix is trusted.

#### Status (combined)

**APPLIED - 2026-04-25** in `src/aeRibbonClass.cls`:

- Saved-flag invalidation in `CaptureHeading1s` (first fix).
- All nav paths now call `CaptureHeading1s` unconditionally
  (`NextButton`, `PrevButton`, `GoToH1`, `OnBookChanged`,
  `GoToChapter`).
- `Erase headingData` / `Erase chapterData` before each rescan to
  avoid stale tail entries.

Awaiting re-verification with the original "edit then Next" repro.

---

## § Book nav scroll-position consistency - 2026-04-25

### Symptom

After Prev/Next book navigation:

- **Prev** lands the new book's `Heading 1` near the **top** of the
  viewport (visually clear).
- **Next** lands the new book's `Heading 1` near the **bottom** of
  the viewport (visually confusing - the rest of the screen shows
  trailing content from the *previous* book).

User wants consistency: H1 at the top in both cases.

### Cause - Word's "scroll just enough" heuristic

Both `NextButton` and `PrevButton` end with the same call:

```vba
ActiveDocument.Range(pos, pos).Select
```

`.Select` moves the cursor and asks Word to make it visible. Word's
default scroll behavior is to scroll the **minimum** distance
needed to bring the cursor into view:

- **Prev**: target is *above* the current viewport. Word scrolls
  up; the cursor (start of new H1) lands near the top. Looks
  great.
- **Next**: target is *below* the current viewport. Word scrolls
  down; the cursor lands near the bottom of the new viewport with
  the rest filled by previous-book content.

The asymmetry is in Word, not the code. `.Select` doesn't take a
scroll-alignment argument.

### Fix - explicit `ScrollIntoView` after Select

Two-step pattern in both buttons:

```vba
Dim rTarget As Word.Range
Set rTarget = ActiveDocument.Range(pos, pos)
rTarget.Select                               ' moves cursor (still required)
ActiveWindow.ScrollIntoView rTarget, True    ' forces H1 to top of viewport
```

`ScrollIntoView` with `Start:=True` aligns the range to the top of
the visible region. For Prev, this is usually a no-op (H1 already
near top); for Next, it forces the desired upward scroll.

### Why this is not "Bug 19"

The existing comment warns that `ScrollIntoView` *alone* leaves the
cursor stale (Bug 19). The new pattern keeps `.Select` first
(cursor is moved), then uses `ScrollIntoView` only to override
Word's scroll heuristic. Cursor is correct AND the scroll position
is consistent.

### Tradeoff

- One extra `ScrollIntoView` call per Next/Prev click - fast.
- Prev becomes a no-op scroll when the H1 was already at the top -
  invisible to the user.
- Next now scrolls further than Word's default - the desired
  outcome.

### Status

**APPLIED - 2026-04-25** in `src/aeRibbonClass.cls` (`NextButton`
and `PrevButton`). Awaiting user re-verification.

### Round 2 - first scroll fix didn't take - 2026-04-25

User report after the first scroll fix:

> Next still appears near the bottom. Prev is lower than before and
> flashes the screen. It is distracting with fast nav.

Two problems with the first attempt:

1. **Zero-width range is degenerate for `ScrollIntoView`**.
   Calling `ActiveWindow.ScrollIntoView ActiveDocument.Range(pos,
   pos), True` passes a zero-length range. There is no meaningful
   "upper-left corner" for Word to align to the top. Word falls
   back to the same "ensure visible" heuristic that `.Select`
   already used - no improvement for Next, and a small
   not-actually-aligned nudge for Prev.
2. **Two scroll calls = two paints = visible flash**. `.Select`
   scrolled once (Word's default), then `ScrollIntoView` scrolled
   again (the nudge above). Each operation triggers a screen
   repaint. Bad on fast nav.

### Round 2 fix - non-zero range + ScreenUpdating off

Three changes per button:

1. Use a **non-zero range from H1 to end-of-document** as the
   `ScrollIntoView` target. That gives Word a real upper-left
   corner (the H1) to align to the top of the viewport.
2. Reverse the order: **scroll first, then `.Select`**. Once the
   H1 is at the top, the cursor position is already on screen so
   `.Select` does not trigger any further scroll.
3. Bracket the pair with `Application.ScreenUpdating = False /
   True` so only the final state paints.

```vba
Application.ScreenUpdating = False
Dim rView As Word.Range
Set rView = ActiveDocument.Range(pos, ActiveDocument.Content.End)
ActiveWindow.ScrollIntoView rView, True
ActiveDocument.Range(pos, pos).Select
Application.ScreenUpdating = True
```

The Round 1 (zero-width range, .Select-then-ScrollIntoView)
attempt is removed - replaced wholesale.

### Net behavior

- Next: H1 at top of viewport. One paint per click. No flicker.
- Prev: H1 at top of viewport (or unchanged when already there).
  One paint per click. No flicker.
- Both buttons render identically.

### Risk / fallback (not applied)

If `ScrollIntoView` with the H1-to-EOD range still misbehaves on
some Word versions, the proven fallback is `ActiveWindow.SmallScroll
Up:=99` to jump to the top of the current page first, then
`ScrollIntoView` again. Costs a second paint - skip unless this
fix proves unreliable.

### Status

**APPLIED - 2026-04-25 (round 2)** in `src/aeRibbonClass.cls`
(`NextButton`, `PrevButton`). Awaiting re-verification.

### Round 3 - cursor caret flash - 2026-04-26

User report after round 2:

> This works. No heading jumping, but the cursor is seen flashing
> from top or bottom according to prev/next direction. Is there a
> solution to that?

The H1 lands at the top in both directions (round 2 fix held).
What remained was a brief visible flash of the **caret** (text
insertion cursor) at the *old* position before snapping to the
new H1.

#### Why the caret slipped past `ScreenUpdating = False`

`Application.ScreenUpdating = False` suppresses Word's content
repaints, but **not** the OS-level caret. The caret is rendered
by Windows on top of the window via `CreateCaret` /
`SetCaretPos`, asynchronous to Word's paint cycle. So the caret
blinker draws the old position briefly between the
`ScrollIntoView` / `.Select` operations and the eventual repaint.

#### Round 3 fix - `LockWindowUpdate` (Win32)

Use `user32!LockWindowUpdate` to freeze the entire Word window's
paint pipeline - including the OS caret - around the scroll +
select pair. Word internally completes both operations; nothing
renders until unlock. One visible repaint, no caret afterimage.

Code changes in `src/aeRibbonClass.cls`:

1. New API import next to existing `LoadCursor` / `SetCursor`:
   ```vba
   Private Declare PtrSafe Function LockWindowUpdate Lib "user32" _
       (ByVal hWndLock As LongPtr) As Long
   ```
2. Both `NextButton` and `PrevButton` wrap the scroll + select
   block:
   ```vba
   LockWindowUpdate Application.hwnd
   Application.ScreenUpdating = False
   ' ScrollIntoView + Select
   Application.ScreenUpdating = True
   LockWindowUpdate 0
   ```
3. Each button's `PROC_ERR` handler restores both flags BEFORE
   showing the `MsgBox` so a crash mid-lock cannot leave Word's
   window frozen:
   ```vba
   PROC_ERR:
       Application.ScreenUpdating = True
       LockWindowUpdate 0
       MsgBox ...
   ```
   `LockWindowUpdate(0)` is a safe no-op when no lock is held.

#### Net behavior

- One visible paint per Next/Prev click.
- No caret flash at the old position.
- H1 still at top of viewport in both directions.
- Fast nav no longer distracting.

#### Safety footgun (mitigated)

If the routine errored between `LockWindowUpdate hwnd` and
`LockWindowUpdate 0`, Word's window would stay frozen. The
`PROC_ERR` unlock handles this. Both subs already had
`On Error GoTo PROC_ERR` plumbing; the unlock just slots into
the existing handler.

#### Status

**APPLIED - 2026-04-26 (round 3)** in `src/aeRibbonClass.cls`
(`NextButton`, `PrevButton`, plus the new `LockWindowUpdate`
declare). Awaiting user re-verification.

#### Round 3a - Application.Hwnd does not exist in Word - 2026-04-26

Compile error on first run:

```
Method or data not found:
    LockWindowUpdate Application.hwnd
```

Word's `Application` object has no `Hwnd` property. That property
is on Excel's and Outlook's Application objects but not Word's -
an inconsistency in the Office object model.

Fix: get Word's main-window handle via `user32!FindWindow`. Word's
main window class name is `"OpusApp"` (a legacy from when Word's
internal codename was "Opus").

Code changes in `src/aeRibbonClass.cls`:

1. New API import next to `LockWindowUpdate`:
   ```vba
   Private Declare PtrSafe Function FindWindow Lib "user32" _
       Alias "FindWindowA" _
       (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
   ```
2. Both `NextButton` and `PrevButton` lock-call sites changed:
   ```vba
   LockWindowUpdate FindWindow("OpusApp", vbNullString)
   ```
   `LockWindowUpdate 0` (unlock) and `PROC_ERR` unlock are
   unchanged - passing 0 is the documented unlock signal and does
   not need an hwnd.

`FindWindow` is fast (a single hash table lookup against the OS
window list); calling it inline once per Next/Prev click has no
perceptible cost. Caching the hwnd was considered but rejected -
it would not survive a Word window recreation (e.g., user closes
and reopens a doc), and the lookup is cheap.

#### Status

**APPLIED - 2026-04-26 (round 3a)** in `src/aeRibbonClass.cls`
(both nav subs + new `FindWindow` declare). Compile error
resolved. Awaiting user re-verification of the original caret-
flash fix.

#### Round 3b - LockWindowUpdate backed out, Selection.SetRange in - 2026-04-26

User report after round 3a:

> Cursor still jumps from top and bottom depending on prev/next.

The lock didn't help and may have *regressed* round 2: by freezing
Word's paint pipeline, `LockWindowUpdate` likely interfered with
`ScrollIntoView`'s commit, so `.Select`'s default scroll-to-cursor
heuristic won out (cursor at top for Prev, bottom for Next - the
exact pre-round-2 behavior).

Diagnostic via `?FindWindow("OpusApp", vbNullString)` returned
"sub or function not defined" - expected, because `Private Declare`
in a class is not callable from the Immediate window. The
diagnostic is moot; the symptom alone is enough to back out.

#### Backed out

In `src/aeRibbonClass.cls`:

- `LockWindowUpdate` and `FindWindow` `Private Declare` lines
  removed from the API import block at the top.
- Both nav subs no longer call
  `LockWindowUpdate FindWindow(...)` / `LockWindowUpdate 0`.
- Both `PROC_ERR` handlers no longer call `LockWindowUpdate 0`
  (`ScreenUpdating = True` retained).

#### New attempt - `Selection.SetRange` instead of `Range.Select`

`.Select` triggers Word's "scroll-to-cursor" heuristic;
`Selection.SetRange Start:=pos, End:=pos` sets the selection
WITHOUT firing that heuristic. If Word respects this in practice,
the OS caret has no reason to render at the old screen position
between the `ScrollIntoView` and the final paint.

Both `NextButton` and `PrevButton` updated:

```vba
Application.ScreenUpdating = False
Dim rView As Word.Range
Set rView = ActiveDocument.Range(pos, ActiveDocument.Content.End)
ActiveWindow.ScrollIntoView rView, True
Selection.SetRange Start:=pos, End:=pos
Application.ScreenUpdating = True
```

#### Net behavior expected

- H1 lands at top (round 2 outcome restored).
- No double-scroll wobble.
- If `SetRange` truly suppresses the scroll-to-cursor: no caret
  flash either.

#### If the caret still flashes

Document it as a known cosmetic OS-caret artifact and stop
chasing. Suppressing the OS caret cleanly requires owning the
editor child window's hwnd, which is more complexity than the
artifact warrants.

#### Status

**APPLIED - 2026-04-26 (round 3b)** in `src/aeRibbonClass.cls`.
LockWindowUpdate / FindWindow code fully removed; Range.Select
replaced with Selection.SetRange in both nav subs. Awaiting
verification.

#### Verification - heading correct, caret flash remains - 2026-04-26

User report after round 3b:

> The heading position is the same level on the screen and fast
> nav with prev/next still shows a cursor flash.

Outcome:

- **H1 placement**: correct. Both Next and Prev land the new
  book's H1 at the same screen level (top of viewport). No
  heading jump, no double-scroll wobble. The functional goal
  is met.
- **Caret flash**: remains on fast nav. Cursor briefly visible at
  old top/bottom position before snapping to new H1.

Neither `Selection.SetRange` nor `LockWindowUpdate` (round 3a/b)
suppressed the caret blink at the old pixel position. The OS
caret render is asynchronous to Word's paint cycle and to the
selection-update messages; the visible blink at the old pixel
position is a Windows / Word artifact below the level VBA can
cleanly reach without owning the editor's child window hwnd.

#### Decision: document as known cosmetic, stop chasing

Per the agreed fallback in round 3b: cleanly suppressing the
caret would require `FindWindowEx` to the editor child window
plus `HideCaret` / `ShowCaret` calls plus careful
exception-safe paired locking. Two more Win32 APIs, more
fragility, and a worse footgun than the artifact itself. Not
worth it.

The behavior is recorded as a known cosmetic limitation:
fast Next/Prev book nav may briefly flash the OS caret at the
previous screen position before it snaps to the new H1. The
heading lands correctly; the caret afterimage is harmless.

#### Final status (this thread)

`src/aeRibbonClass.cls` `NextButton` and `PrevButton`:

- Heading-cache staleness: **FIXED** (round 1, saved-flag
  invalidation; round 2, all nav paths call CaptureHeading1s).
- H1-at-top consistency: **FIXED** (round 2, ScrollIntoView with
  H1..EOD non-zero range + ScreenUpdating).
- Cursor caret flash: **WONTFIX** (cosmetic OS artifact; clean
  fix exceeds value).

---

## § Chapter / verse nav scroll-position consistency - 2026-04-26

Apply the book-nav scroll fix to chapter and verse navigation so
all three navigation tiers land their target heading at the same
viewport position.

### Symptom

When a chapter or verse is selected and displayed via Go or the
Prev/Next chapter/verse buttons, the chapter heading or verse
marker can land anywhere on the screen depending on Word's "ensure
cursor visible" minimum scroll. Visually inconsistent with the
top-of-viewport book-nav behavior just established.

### Two routines, same root cause as book nav round 1

| Routine | Pre-fix code | Issue |
|---|---|---|
| `GoToChapter` (line 639) | `ActiveDocument.Range(chPos, chPos).Select` | `.Select` only - no explicit `ScrollIntoView`; Word's default scroll heuristic decides the landing position |
| `GoToVerseByScan` (line 997) | `Range(r.Start, r.Start).Select` + `ScrollIntoView Range(r.Start, r.Start), True` | `.Select` + `ScrollIntoView` of a zero-width range - degenerate "no upper-left corner" case; Word falls back to the same heuristic |

### Fix - same three-step pattern as `NextButton` / `PrevButton`

For both routines, replace cursor placement + scroll with:

```vba
Application.ScreenUpdating = False
Dim rView As Word.Range
Set rView = ActiveDocument.Range(targetPos, ActiveDocument.Content.End)
ActiveWindow.ScrollIntoView rView, True
Selection.SetRange Start:=targetPos, End:=targetPos
Application.ScreenUpdating = True
```

Where `targetPos` is `chPos` for `GoToChapter` and `r.Start` for
`GoToVerseByScan`.

### Code changes

`src/aeRibbonClass.cls`:

1. **`GoToChapter`** - the single `.Select` line replaced with
   the three-step block. The `Application.StatusBar = SB_NAVIGATING`
   call is preserved (still useful during the scroll layout cost).
   The "Bug 22b" warning comment is replaced with a comment
   pointing to the new pattern's reasoning.
2. **`GoToVerseByScan`** - the `.Select` + `ScrollIntoView` pair
   inside the verse-found branch replaced with the three-step
   block. Local `rVsView` declared inside the `If Count = vsNum`
   branch.

No other call sites needed editing - all chapter / verse
navigation funnels through these two routines (Prev/Next chapter
buttons -> `GoToChapter`; Prev/Next verse buttons + Go after
verse text input -> `GoToVerse` -> `GoToVerseByScan`).

### Net behavior

- Book H1, chapter H2, and verse marker all land at the same
  level (top of viewport).
- One render per click; no double-scroll wobble.
- Caret flash on fast nav still present - same WONTFIX status as
  book nav (cosmetic OS-caret artifact).

### Status

**APPLIED - 2026-04-26** in `src/aeRibbonClass.cls`
(`GoToChapter`, `GoToVerseByScan`). Awaiting verification with
chapter / verse nav repros.

### Verified - 2026-04-26

User report:

> This works. The page jumps to where the eye is looking rather
> than chasing the cursor around the page.

All three navigation tiers - book Prev/Next, chapter Prev/Next,
verse Go - now anchor the target heading at the top of the
viewport. The cursor follows the eye instead of vice versa.

#### Final status (book + chapter + verse nav)

| Tier | Routines | H1/H2/verse at top | Caret flash |
|---|---|---|---|
| Book | `NextButton`, `PrevButton` | **FIXED** | WONTFIX (cosmetic) |
| Chapter | `GoToChapter` | **FIXED** | WONTFIX (carries over) |
| Verse | `GoToVerseByScan` (called by `GoToVerse`) | **FIXED** | WONTFIX (carries over) |

Heading-cache staleness: **FIXED** across all nav (round 1+2
saved-flag invalidation + every nav path calls `CaptureHeading1s`).

---
