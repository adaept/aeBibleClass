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

## § EDSG snapshot refresh - 2026-04-26

### Trigger

User shared current `WordEditingConfig` output showing the
approved array has advanced significantly since the EDSG was
scaffolded. Validated walk now reaches priority 33; new styles
added; one rename; one removal; `Normal` repositioned to anchor.

### Latest priority-sorted approved list

```
1   TheHeaders
2   BodyText
3   TheFooters
4   FrontPageTopLine
5   TitleEyebrow
6   Title
7   TitleVersion
8   FrontPageBodyText
9   BodyTextTopLineCPBB
10  Acknowledgments
11  AuthorBodyText
12  Contents
13  ContentsRef
14  BibleIndexEyebrow
15  BibleIndex
16  Introduction
18  ListItem
19  ListItemBody
20  ListItemTab
21  AuthorBookRefHeader
22  AuthorBookRef
23  TitleOnePage
24  CenterSubText
25  Heading 1
26  CustomParaAfterH1
27  Brief
28  DatAuthRef
29  Heading 2
30  Chapter Verse marker
31  Verse marker
32  Footnote Reference
33  Footnote Text
--- validated above this line ---
34  Lamentations
35  Psalms BOOK
36  BodyTextIndent
41  EmphasisBlack
42  EmphasisRed
43  Words of Jesus
44  AuthorSectionHead
45  AuthorQuote
46  Normal
```

Missing-from-document (warning): `BodyTextContinuation`,
`BookIntro`, `AppendixTitle`, `AppendixBody`, `FargleBlargle`.
Reserved gaps: 17, 37–40.

### Changes from prior snapshot

- **New approved styles**: `Introduction` (16), `ListItemTab` (20),
  `AuthorBookRefHeader` (21), `TitleOnePage` (23),
  `CenterSubText` (24).
- **Renamed**: `Lamentation` → `Lamentations` (English plural).
- **Removed from approved array**: `Book Title`.
- **Promoted into approved**: `BodyTextIndent` (now at 36; was
  `[not used]` placeholder previously).
- **Repositioned**: `Normal` moved to priority 46 — deliberately
  the last entry, anchor for "pin everything else above this."
  Operationally replaced by `BodyText`.

### EDSG file updates

`EDSG/README.md`:

- Status row for `01-styles.md` updated:
  `WIP — pages 1–11 walked` → `WIP — validated up to priority 33`.

`EDSG/01-styles.md`:

- WIP marker updated to "validated up to priority 33; positions 34
  and beyond pending re-validation."
- Snapshot table replaced. Page column dropped (current source for
  page-keyed view is `ListApprovedStylesByBookOrder`); table now
  split into Validated (1–33) and Pending re-validation (34+)
  blocks.
- New section "Reserved gaps" recording priorities 17, 37–40.
- New section "Missing from document" listing the array entries
  reported by `PromoteApprovedStyles`'s diagnostic, including the
  deliberate `FargleBlargle` canary.
- "Style categories" updated to reflect the new front-matter
  members, `Lamentations` rename, removal of `Book Title`, and
  `Normal` repositioned as anchor.
- "Front matter" range expanded to priorities 1–24 (was 1–15);
  "Body text" now starts at 25 (was 16).

`EDSG/04-qa-workflow.md`:

- "Current state — 2026-04-25" replaced with current state dated
  2026-04-26.
- Records the validated-through-33 line, the styles added /
  renamed / removed, the `Normal` repositioning, and the gaps.
- `BodyTextIndent` decision now resolved (in array at 36);
  `AuthorQuote` decision still pending.

`EDSG/03-inspection-tools.md`:

- Sample `ListApprovedStylesByBookOrder` output trimmed: the
  illustrative `[not used]` line previously cited `BodyTextIndent`
  at priority 18 (now stale on both axes). Replaced with a
  generic `<unused style>` placeholder so the sample doesn't age.

Other EDSG files (`02-editing-process.md`, `05-headers-footers.md`,
`06-i18n.md`, `07-super-test-runs.md`, `08-publishing.md`,
`09-history.md`) reviewed - no style-name references that needed
updating.

### Status

EDSG content **REFRESHED - 2026-04-26** to reflect the priority-33
validated state. Next refresh due when the page walk extends past
priority 33.

---

## § Publish EDSG to edsg.adaept.com - plan - 2026-04-26

### Goal

Host `/EDSG/` as a public website at `https://edsg.adaept.com`,
sourced from the GitHub repo, with i18n scaffolding, a docx export
that uses the Study Bible style template, and CI that exercises
the build weekly.

Constraints / preferences:

- US English is the primary content; i18n feedback wanted ASAP
  even at small scale.
- EDSG is intentionally smaller than the Study Bible itself - good
  test bed for the publishing pipeline before applying lessons to
  the Bible.
- Editor / translator audience first; developers second.
- Code signing extends to the EDSG.docx (single project, ongoing
  maintenance).

### Recommended stack (top of plan)

| Concern | Choice | Why |
|---|---|---|
| Hosting | GitHub Pages | Free; native to the repo; custom domain support |
| Static site generator | **Docusaurus 3** | First-class i18n, search, CJK-clean defaults, React-based theming |
| Custom domain | edsg.adaept.com via DNS CNAME | adaept.com already owned; no new registration |
| Docs license | CC BY-SA 4.0 | Preserves attribution; allows derivatives under same terms; widely understood |
| Docx export | Pandoc + reference template | Pandoc maps markdown to Word styles via reference doc |
| Code signing | Sectigo or DigiCert OV code-signing cert (multi-year) | Required hardware-token issuance since 2023; multi-year reduces renewal frequency |
| CI | GitHub Actions (cron weekly + on-push build) | Native, free for public repos |

Alternates worth knowing: MkDocs Material is simpler than
Docusaurus but i18n is plugin-based and less polished; VitePress
is fast but i18n is younger; raw GitHub Pages with Jekyll is the
zero-tooling baseline.

---

### Phase A - Public site live (do now)

Cheap, fast, unblocks everything downstream.

- [ ] **A1.** Create `/website/` (Docusaurus root) at the repo's
      project root. `EDSG/*.md` becomes the source content; the
      site builds into `/website/build/`.
- [ ] **A2.** Initialize Docusaurus 3 with English locale only.
      Sidebar generated from EDSG file numbering (01-, 02-, ...).
- [ ] **A3.** Pick and apply a docs theme. Default Docusaurus
      Classic is fine for v1; theme polish later.
- [ ] **A4.** Add `LICENSE-DOCS` at repo root (CC BY-SA 4.0) and
      a footer link from every EDSG page (Docusaurus theme config).
- [ ] **A5.** Configure GitHub Pages: deploy from GitHub Actions
      workflow that runs `npm run build` and publishes
      `/website/build/` to the `gh-pages` branch.
- [ ] **A6.** Add `CNAME` file containing `edsg.adaept.com` to
      the published artifact.
- [ ] **A7.** Configure DNS at adaept.com: CNAME record
      `edsg → adaept.github.io`.
- [ ] **A8.** Verify HTTPS via GitHub Pages auto-issue
      (Let's Encrypt). Wait for SSL provisioning (5-30 min).
- [ ] **A9.** Smoke-test: site loads, every EDSG page renders,
      navigation works, license footer visible.

Estimated effort: 3-5 hours for someone familiar with Docusaurus;
6-10 hours from scratch.

---

### Phase B - i18n scaffolding (do now, even if English-only initially)

The i18n shape is much cheaper to bake in from day 1 than
retrofit. Empty translation directories signal openness without
committing to translation work.

- [ ] **B1.** Enable Docusaurus i18n in `docusaurus.config.js`:
      `i18n: { defaultLocale: 'en', locales: ['en'] }`. Add
      placeholder for future locales.
- [ ] **B2.** Move EDSG content under `i18n/en/docusaurus-plugin-
      content-docs/current/` (Docusaurus convention) OR keep it
      at `EDSG/` with a docs plugin pointed there - decision
      pending the i18n process choice.
- [ ] **B3.** Add `i18n/<locale>/` skeleton directories for
      candidate first-translation locales (placeholder); see
      sub-list below.
- [ ] **B4.** Configure language switcher in the navbar (hidden
      until at least one non-English locale has content).
- [ ] **B5.** Translation workflow doc at
      `EDSG/06-i18n.md` — extend the existing skeleton with a
      concrete process: fork → translate → PR.

#### i18n process specifics

- [ ] **B-i18n-1.** Decide first-locale target. (Open question
      from EDSG plan; carry forward.) Likely candidates: French,
      Spanish (Latin script — lowest tooling friction); Japanese
      or Chinese (CJK — exercises font and layout assumptions).
- [ ] **B-i18n-2.** Establish translation source: human
      translator vs LLM-assisted vs hybrid. Document in
      `06-i18n.md`.
- [ ] **B-i18n-3.** Translation memory / glossary file at
      `i18n/glossary.yml` — Bible-specific terms with canonical
      translations per locale.
- [ ] **B-i18n-4.** PR review process: who validates translation
      accuracy. Likely: native-speaker reviewer per locale;
      flagged in CONTRIBUTING.md.
- [ ] **B-i18n-5.** Translation freshness CI check: weekly job
      reports per-locale staleness (English changes since last
      translation update).

#### CJK compatibility specifics

- [ ] **B-cjk-1.** Font stack in CSS: `font-family: -apple-system,
      "Noto Sans CJK SC", "Noto Sans CJK TC", "Noto Sans CJK JP",
      "Noto Sans CJK KR", system-ui, sans-serif;`. Loaded from
      Google Fonts or self-hosted (self-host preferred for
      privacy).
- [ ] **B-cjk-2.** CSS line-height bumped slightly for CJK locales
      (Japanese and Chinese benefit from ~1.7 vs 1.5 for Latin).
- [ ] **B-cjk-3.** Search: Docusaurus default uses Algolia
      DocSearch. Algolia handles CJK tokenization out of the box
      for Japanese; Chinese / Korean may need tokenization tuning.
      Local search plugin alternative: `@easyops-cn/docusaurus-
      search-local` — has CJK tokenizer support.
- [ ] **B-cjk-4.** Test page: a sample paragraph in Simplified
      Chinese, Japanese, and Korean to visually validate font
      fallback chain on first deploy.
- [ ] **B-cjk-5.** Right-to-left (RTL) is out of scope until
      Hebrew / Arabic locale is added; Docusaurus supports
      `direction: 'rtl'` per locale when the time comes.

#### Skeleton frame for languages

- [ ] **B-skel-1.** Create `i18n/<locale>/` for each candidate
      locale even before content exists. Each contains:
      - `code.json` (UI strings)
      - `docusaurus-plugin-content-docs/current/` (translated
        markdown)
      - `docusaurus-theme-classic/` (theme overrides if any)
- [ ] **B-skel-2.** Initial candidate locales (suggest 3 to
      cover scripts): `fr` (Latin), `zh-CN` (Simplified Chinese),
      `ja` (Japanese). Hebrew (`he`) added when RTL work begins.
- [ ] **B-skel-3.** Auto-generate stub markdown files in each
      locale that match the English structure but contain only
      a header and a "translation pending" marker. CI gate
      ensures every English page has at least a stub in every
      configured locale.

---

### Phase C - Docx export using EDSG / Bible styles (do soon)

The dogfooding goal: produce EDSG.docx using the same approved
styles as the Study Bible. Validates the templates against a
non-Bible document.

- [ ] **C1.** Decide reference template: clone the Study Bible
      `.docm` and strip Bible content, leaving only the styles?
      Or generate fresh from `RUN_TAXONOMY_STYLES`? Recommend
      the clone-and-strip — guarantees style fidelity to the
      Bible.
- [ ] **C2.** Save reference template at `EDSG/build/EDSG.docx`
      (gitignored output) with a source at
      `EDSG/build/reference.docx` (committed).
- [ ] **C3.** Pandoc command at `EDSG/build/build.cmd` (and
      `build.sh` for CI):
      ```
      pandoc README.md 01-styles.md ... -o build/EDSG.docx \
        --reference-doc=build/reference.docx \
        --toc --toc-depth=2
      ```
- [ ] **C4.** Markdown → style mapping table (Pandoc respects
      style names from the reference doc):
      - `# H1` → `Heading 1`
      - `## H2` → `Heading 2`
      - body → `BodyText`
      - inline code → custom character style (define if needed)
      - code block → custom paragraph style (define if needed)
      - tables → Word table style chosen from approved set
- [ ] **C5.** Add `CodeBlock` paragraph style + `InlineCode`
      character style to the approved array (currently absent;
      previously flagged in `08-publishing.md`).
- [ ] **C6.** CI step builds the docx and uploads as a workflow
      artifact (Actions free tier covers this).
- [ ] **C7.** Optional: PDF export via Pandoc + LaTeX (or via
      Word's PDF export driven by an Office Automation script on
      a Windows runner — heavier, defer).

---

### Phase D - Code signing the docx (do soon-ish)

The docx is the concrete dogfooded artifact; signing it means
editors trust its origin without "this file came from the
internet" warnings. Same cert can sign the Bible's `.docm`.

- [ ] **D1.** Decide cert vendor and term length. See "Cert
      vendor" notes below.
- [ ] **D2.** Acquire hardware token (HSM) from cert vendor -
      mandatory since June 2023 per CA/B Forum baseline
      requirements.
- [ ] **D3.** Configure Office to require macros to be signed
      (`File → Options → Trust Center → Macro Settings`).
- [ ] **D4.** Sign the EDSG.docx as part of the build pipeline
      (Windows runner with `signtool` or PowerShell
      `Set-AuthenticodeSignature`).
- [ ] **D5.** Sign the Bible `.docm` and the VBA project on
      release.
- [ ] **D6.** Document the signing process in
      `EDSG/08-publishing.md` and in a new
      `09-history.md` cross-reference.

#### Cert vendor analysis

User asked about a "no annual renewal" certificate. Honest
status as of 2026-04:

- **No major Microsoft-trusted CA currently issues code-signing
  certificates that never need renewal.** All commercial code
  signing certs (Sectigo, DigiCert, SSL.com, GlobalSign) require
  renewal at the end of their term.
- **Multi-year terms reduce renewal frequency**:
  - 1-year: cheapest per-year, most renewals
  - 2-year: ~15% per-year discount
  - 3-year: ~25% per-year discount, **maximum** allowed by CA/B
    Forum baseline requirements (no longer can be 5- or 10-year
    as some used to offer)
- **Current floor price**: ~$200-$400/year for OV (Organization
  Validation) code signing from a reseller; ~$600-$1000/year for
  EV (Extended Validation) code signing. EV signs without
  SmartScreen reputation warmup; OV requires building reputation
  via downloads.
- **What the user may have read about**: possibly self-signed
  certs (free, zero renewal, but **only trusted on the issuing
  machine** - not useful for distribution); or the (now-defunct)
  CAcert; or older "lifetime subscription" offers from CAs that
  no longer exist post-2017 baseline-requirements changes.
- **Recommendation**: Sectigo OV code signing, 3-year term, via
  a reseller like KSoftware or SignMyCode. ~$60-90/year
  effective rate. Hardware token included.
- **Alternative for VBA-only**: a Document Signing certificate
  (different EKU, often cheaper) signs the VBA project and
  validates inside Office, but does not sign external `.exe` /
  `.msi` artifacts. If signing scope is strictly "the docx and
  the docm and nothing else," this is a viable lower-cost
  path - confirm with the chosen CA that their Document Signing
  cert validates VBA projects.

---

### Phase E - GitHub bug reports + CI (do now for templates;
weekly cron later)

- [ ] **E1.** Issue templates at
      `.github/ISSUE_TEMPLATE/`:
      - `bug-report.yml`
      - `documentation-issue.yml` (EDSG-specific)
      - `translation-issue.yml` (i18n)
      - `feature-request.yml`
- [ ] **E2.** PR template at `.github/pull_request_template.md`
      with checklist (description / impacted files / testing
      done / docs updated).
- [ ] **E3.** GitHub Discussions enabled - separate space for
      open-ended questions vs Issues for bugs.
- [ ] **E4.** CONTRIBUTING.md at repo root - first-time
      contributor onboarding, links to relevant EDSG pages.
- [ ] **E5.** GitHub Actions workflows under `.github/workflows/`:
      - `site-build.yml` — on push to main, build and deploy
        Docusaurus
      - `docx-build.yml` — on push to main, build EDSG.docx,
        attach as artifact
      - `weekly.yml` — cron `0 0 * * 0` (Sundays 00:00 UTC),
        runs link-check, markdown lint, and (when ready)
        SUPER_TEST_RUNS-equivalent for the docs
- [ ] **E6.** Status badge on README pointing to the weekly
      job - "build status" surfaces failures immediately.

---

### Pros / Cons / Benefits

#### Pros (do now)

- **Public visibility now** invites the i18n feedback the user
  explicitly wants ASAP.
- **Smaller surface area than Bible** - mistakes here are cheap
  and recoverable; lessons learned apply to the Bible publishing
  pipeline later.
- **Forcing function** for license + signing decisions that have
  been in the queue.
- **GitHub Pages is free** - no hosting cost.
- **Docusaurus i18n is mature** - cheap to bake in early.

#### Cons (do now)

- **Bandwidth split** - while EDSG site is being built, fewer
  cycles for Bible style work and `SUPER_TEST_RUNS`.
- **Maintenance overhead** - a public site implies a SLA in the
  reader's mind even if not formally promised.
- **Cert cost** - even minimal Sectigo OV is ~$200/year up front.
- **Risk of premature publication** - parts of the EDSG are
  still WIP (priorities 34+). Status markers per page mitigate
  but don't eliminate.

#### Benefits

- **Real-world i18n testbed** - find tooling problems on a small
  doc before they bite the Bible.
- **Translator recruitment surface** - a public site with
  contribution paths attracts translators in a way that a
  GitHub repo alone does not.
- **Auditable artifact** - signed docx + signed VBA project +
  CC-licensed source = a defensible position for trademark / name
  attribution.
- **CI weekly cadence** establishes a heartbeat - regressions
  caught within 7 days even if no one's actively pushing.

---

### Cost - now vs later

#### Now

- **Setup time** (Phase A + B + E1-E5): ~15-25 hours of focused
  work, spread over 1-2 weeks at part-time pace.
- **Code signing cert** (Phase D): ~$60-$90/year amortized
  (3-year Sectigo OV via reseller).
- **DNS work** (Phase A7): minutes, no incremental cost.
- **Ongoing**: ~1 hour/week to keep CI green and merge
  translation contributions if/when they arrive.

#### Later

- **Setup time** (deferring 6-12 months): same total time but
  back-loaded. Risk: i18n shape decisions get made implicitly
  by ad-hoc retrofits instead of explicitly up front.
- **Translation cost** of waiting: any translator who could have
  contributed in the meantime is lost.
- **Visibility cost**: search engines take weeks to index a new
  site; later launch means later visibility.
- **Cert cost**: same; certs renew on the same calendar
  regardless of when this project starts.

#### Recommendation

**Phase A + B1-B5 + E1-E4 now** (week 1) - get the site live,
i18n scaffolded, contribution paths defined.

**Phase C + B-skel + B-cjk** within month 1 - dogfood the docx
export, populate language skeletons, validate CJK rendering.

**Phase D** before public announcement - signing has lead time
(token ordering, identity validation - 1-3 weeks for Sectigo OV).

**Phase E5-E6** after first week of normal commits - so the
weekly CI has something real to check.

**Phase B-i18n full process** when the first translator
volunteers - don't pre-build a process for a non-existent flow.

---

### Open questions to resolve before Phase A

1. Confirm `adaept.io` reference - was it a typo for
   `adaept.com`? Or a separate domain to also point? (Cheap to
   add a second CNAME if so.)
2. Reference template (Phase C1): clone-and-strip the Bible
   `.docm`, or generate fresh? Recommend clone-and-strip.
3. First-locale candidate (Phase B-i18n-1): pick one to
   un-stub.
4. Code signing cert vendor (Phase D1): Sectigo via reseller is
   the recommended starting point - confirm or override.
5. Docusaurus content path (Phase B2): `i18n/en/...` (full
   Docusaurus i18n layout) or `EDSG/` (current path, with docs
   plugin pointed there) - decision affects diff against today.

---

### Status

EDSG publication plan: **DRAFT - awaiting user review and
phase-by-phase approval**. No site, DNS, cert, or CI changes
applied. All decisions deliberately staged behind explicit
go-aheads.

---

## § Latest DumpAllApprovedStyles run + array observations - 2026-04-26

### Run output

```
DumpAllApprovedStyles: Done. 40 succeeded, 0 failed.
Orphan style dumps (in rpt\Styles but not in current approved list):
  C:\adaept\aeBibleClass\rpt\Styles\style_Lamentations.txt
  deleted: C:\adaept\aeBibleClass\rpt\Styles\style_Lamentations.txt
DumpAllApprovedStyles - actual 8.04 sec
```

40 styles dumped successfully (down from 41 - one removal
since the previous EDSG snapshot). Orphan-cleanup mechanism
worked end-to-end: `style_Lamentations.txt` was detected,
listed, prompted, deleted on confirmation. ~8 sec total runtime
including the orphan pass.

### Array observations (from `src/basTEST_aeBibleConfig.bas`)

While reconciling the count, two findings worth recording:

1. **`Lamentations` is no longer in the `approved` array.**
   Confirms why the file was orphaned. `Lamentation` (singular)
   appears only in the `AuditOneStyle` list at line 275, not in
   the promoted approved array. Whether this is intentional
   (Lamentations dropped pending decision) or a regression
   needs user confirmation.
2. **`TitleOnePage` appears twice** in the approved array (lines
   38 and 41). `PromoteApprovedStyles` iterates the array
   sequentially and assigns `s.Priority = i + 1`; on the second
   pass the priority is overwritten with the later position, so
   the earlier slot becomes a "dead" priority no style holds.

   This is almost certainly the actual cause of the **gap at
   priority 17** that the EDSG previously documented as a
   "reserved gap." It's not a deliberate reservation — it's a
   copy-paste artifact.

### EDSG file updates

`EDSG/01-styles.md`:

- "Pending re-validation" table updated: `Lamentations` removed;
  remaining priorities renumbered nominally.
- New "Known issues" subsection flags the `TitleOnePage`
  duplicate as the cause of priority 17's gap.
- "Missing from document" subsection appended with a note about
  the Lamentations removal and the orphan auto-cleanup that
  followed.
- "Special book treatments" updated: removed `Lamentations`
  from the listed members; added a note that `Lamentation`
  (singular) is in the audit list only.
- Validated count statement updated to 40.

`EDSG/04-qa-workflow.md`:

- "Current state" subsection updated with the latest-run
  results (40/0, ~8s, orphan cleaned).
- Lamentations removal noted in the change list.
- New "Known issue - duplicate in array" subsection documenting
  the `TitleOnePage` duplicate and its priority-17 side effect.
  Recommended fix flagged for the next array edit.
- "Reserved gaps" reframed to acknowledge that priority 17 is
  symptom-driven, not reserved.

### Recommended fix (not applied)

Remove the second `"TitleOnePage"` entry from line 41 of
`src/basTEST_aeBibleConfig.bas`. After re-running
`WordEditingConfig`, `TitleOnePage` will hold its earlier
priority and the gap at 17 will close (every priority below
its old slot will shift down by 1).

The fix is small but has cascading effects on every priority
≥ 17 in the document — flagged for explicit user approval before
applying so the EDSG snapshot can be refreshed in the same
commit cycle.

### Decision items for user

1. Confirm: was `Lamentations` removed intentionally? If
   intentional, what's the plan for that book's special
   formatting (currently no promoted style for it)? If
   accidental, restore.
2. Approve removing the `TitleOnePage` duplicate — yes (with
   priority shift acceptable) or no (live with the dead slot).

### Status

EDSG content **REFRESHED - 2026-04-26 (latest run)** to reflect
the 40-style state and the `TitleOnePage` duplicate finding.
Code edits to `basTEST_aeBibleConfig.bas` **PENDING user
approval** of the two decision items above.

---

## § Post-cleanup snapshot - 2026-04-26 (later that day)

User confirmed the two decision items and applied the fixes:

- **Lamentations removal**: intentional. Book content
  standardized on `BodyText` for now. No further action needed.
- **TitleOnePage duplicate**: fixed in
  `src/basTEST_aeBibleConfig.bas`. `TitleOnePage` now holds
  priority **17** (the previously "dead" slot is closed).
- **New styles added**: `PsalmSuperscription` (34), `Selah` (35),
  `PsalmAcrostic` (36) — Psalms-specific styles encountered
  while walking the Psalms book.

### Latest run output

```
DumpAllApprovedStyles: Done. 43 succeeded, 0 failed.
DumpAllApprovedStyles - actual 4.08 sec
```

Runtime down ~50% from the previous 8.04 sec - likely due to the
array no longer containing the duplicate (no double-promote +
overwrite).

### Latest priority order (from `WordEditingConfig`)

```
1   TheHeaders
2   BodyText
3   TheFooters
4   FrontPageTopLine
5   TitleEyebrow
6   Title
7   TitleVersion
8   FrontPageBodyText
9   BodyTextTopLineCPBB
10  Acknowledgments
11  AuthorBodyText
12  Contents
13  ContentsRef
14  BibleIndexEyebrow
15  BibleIndex
16  Introduction
17  TitleOnePage          <- previously stuck at gap; now holds slot
18  ListItem
19  ListItemBody
20  ListItemTab
21  AuthorBookRefHeader
22  AuthorBookRef
23  CenterSubText
24  Heading 1
25  CustomParaAfterH1
26  Brief
27  DatAuthRef
28  Heading 2
29  Chapter Verse marker
30  Verse marker
31  Footnote Reference
32  Footnote Text
33  Psalms BOOK
34  PsalmSuperscription   <- new
35  Selah                 <- new
36  PsalmAcrostic         <- new
--- validated above this line ---
37  BodyTextIndent
42  EmphasisBlack
43  EmphasisRed
44  Words of Jesus
45  AuthorSectionHead
46  AuthorQuote
47  Normal
```

Reserved gaps: 38-41 (now genuinely reserved, post-duplicate-fix).
Missing-from-document: `BodyTextContinuation`, `BookIntro`,
`AppendixTitle`, `AppendixBody`, `FargleBlargle` (canary).

### EDSG file updates (this refresh)

`EDSG/README.md`:

- Status row for `01-styles.md`: validated to priority **36**
  (was 33).

`EDSG/01-styles.md`:

- Snapshot rewritten with the 43-style list. Validated block
  expanded to 1-36; Pending block trimmed to 37+.
- `TitleOnePage` now correctly listed at priority 17.
- Three new entries: `PsalmSuperscription`, `Selah`,
  `PsalmAcrostic`.
- "Known issues" subsection (TitleOnePage duplicate) **removed**
  - bug fixed.
- "Reserved gaps" rewritten: 38-41 are now genuine reservations;
  noted that the priority-17 "gap" was a duplicate artifact, now
  resolved.
- "Special book treatments" expanded with descriptions for the
  three new Psalms styles (PsalmSuperscription = author /
  context line; Selah = Hebrew interjection; PsalmAcrostic =
  Hebrew-letter section markers, notably Psalm 119).

`EDSG/04-qa-workflow.md`:

- "Current state" subsection updated with the 43/0 run, the
  duplicate fix, the Lamentations removal confirmation, and the
  three new Psalms styles.
- "Known issue - duplicate in array" subsection **removed**
  - bug fixed.
- "Reserved gaps" reframed to genuine reservation status.

### Status

Snapshot **REFRESHED - 2026-04-26 (post-cleanup)** to reflect 43
approved styles, validated through priority 36, with the
TitleOnePage duplicate resolved and three Psalms-specific styles
added. Walk continues toward priority 37+ in the next QA cycle.

---

## § Plan - introduce `VerseText` style - 2026-04-26

### Goal

Replace `BodyText` with a new `VerseText` paragraph style on every
paragraph in the Bible body that begins with a "Chapter Verse
marker" character-style run (i.e., every verse paragraph). The
existing `BodyText` style remains for non-verse paragraphs (intros,
spacers, footnote contexts, etc.).

The new style is identical to `BodyText` except for the name.

### Stated benefits

- **First-occurrence semantics**: with the canonical
  book-order-by-first-occurrence rule, `VerseText` first appears
  on the first verse of the Bible (Genesis 1:1). Currently
  `BodyText` first appears on page 1 (front matter spacer),
  which is semantically misleading.
- **USFM mapping clarity**: `VerseText` maps cleanly to USFM
  `\v` (verse body). `BodyText` becomes the residual for non-verse
  paragraphs that map to `\p`, `\ip`, etc.

### Identification rule

A paragraph qualifies for conversion when both:

1. `paragraph.Style.NameLocal = "BodyText"`, AND
2. `paragraph.Range.Characters(1).Style.NameLocal = "Chapter Verse marker"`.

Per the existing `GoToVerseByScan` comment in `aeRibbonClass.cls`:

> Every verse in both study and print versions begins with a
> "Chapter Verse marker" run (chapter number) immediately followed
> by a "Verse marker" run (verse number).

So the rule catches exactly the verse-bearing paragraphs and
nothing else.

### Plan - phased

#### Phase 1 - define the new style (low risk)

- [ ] **1.1** Add `DefineVerseTextStyle` to
  `src/basFixDocxRoutines.bas` - clones every property of
  `BodyText` (BaseStyle = "", identical Font / ParagraphFormat
  fields, AutomaticallyUpdate = False, QuickStyle = False).
- [ ] **1.2** Add `"VerseText"` to the `approved` array in
  `src/basTEST_aeBibleConfig.bas`. Position: immediately after
  `"Verse marker"` so VerseText falls into the verse-marker
  cluster in book-order priority. Will renumber priorities 31+
  (current `Footnote Reference` etc.) by +1.
- [ ] **1.3** Add `VerseText` row to `RUN_TAXONOMY_STYLES` with
  the same expected property values as `BodyText` (separate
  row; not "based on" since BaseStyle = "").
- [ ] **1.4** Run `WordEditingConfig`; confirm
  `VerseText` appears in the priority list. Run
  `DumpStyleProperties "VerseText"`; confirm property dump
  matches `BodyText`.

#### Phase 2 - bulk conversion (one-time mutation, BACKUP FIRST)

- [ ] **2.1** Commit the clean repo + back up the working `.docm`
  to `~/Backups/` (or git-stash equivalent).
- [ ] **2.2** Add `ConvertBodyTextVersesToVerseText` to
  `src/basFixDocxRoutines.bas`:
  ```vba
  Public Sub ConvertBodyTextVersesToVerseText()
      Dim oPara As Object
      Dim nConverted As Long, nKept As Long, nScanned As Long
      For Each oPara In ActiveDocument.Paragraphs
          If oPara.Style.NameLocal = "BodyText" Then
              nScanned = nScanned + 1
              If oPara.Range.Characters(1).Style.NameLocal _
                              = "Chapter Verse marker" Then
                  oPara.Style = ActiveDocument.Styles("VerseText")
                  nConverted = nConverted + 1
              Else
                  nKept = nKept + 1
              End If
          End If
      Next oPara
      Debug.Print "ConvertBodyTextVersesToVerseText: scanned=" _
          & nScanned & " converted=" & nConverted _
          & " kept=" & nKept
  End Sub
  ```
- [ ] **2.3** Run the conversion. Note the reported counts.
  Expected: tens of thousands converted (one per verse;
  ~31,000 verses in the Bible). `kept` should be small
  (front-matter / non-verse paragraphs styled BodyText).
- [ ] **2.4** Verify by sampling: open random books / chapters
  in Word, click into verse paragraphs, confirm style shows
  `VerseText`. Click into non-verse paragraphs, confirm style
  shows `BodyText`.
- [ ] **2.5** Re-run `DumpAllApprovedStyles` - new file
  `style_VerseText.txt` should appear; orphan check passes.
- [ ] **2.6** Re-run `ListApprovedStylesByBookOrder` - confirm
  `VerseText` first-occurrence is now Genesis 1:1's page (page 1
  of Old Testament body, ~page 22 in the current document).
  `BodyText` first-occurrence shifts to whatever non-verse
  paragraph appears first (likely a front-matter spacer or
  intro).
- [ ] **2.7** Idempotency check: re-run
  `ConvertBodyTextVersesToVerseText`. Expected:
  `converted=0` (no remaining BodyText paragraphs match the
  criterion).

#### Phase 3 - audit and follow-up

- [ ] **3.1** Walk the converted document looking for surprises -
  any paragraph still styled BodyText that should be VerseText,
  or vice versa. Most likely edge case: footnote-text paragraphs
  that incidentally start with a marker run.
- [ ] **3.2** Update USFM export mapping (when implemented):
  `VerseText` → `\v <number> <text>` reconstruction;
  `BodyText` → `\p` or `\ip` per context.
- [ ] **3.3** Update EDSG:
  - `01-styles.md`: add `VerseText` to the snapshot,
    "Body text" category, and a brief rationale note.
  - `04-qa-workflow.md`: note the conversion as a milestone.
  - `06-i18n.md`: add `VerseText` as the primary translation
    target style.
- [ ] **3.4** Code search for hardcoded `"BodyText"` literals
  that may have assumed verse content:
  - `src/aeRibbonClass.cls` - any nav code targeting BodyText?
  - `src/basSBL_*` - parser code referencing BodyText?
  - Test routines - assertions about BodyText paragraph counts?

### Pros

- **Semantic clarity**: paragraph style now answers "is this a
  verse?" with a one-property check.
- **Book-order alignment**: the first-occurrence sort places
  `VerseText` near the start of the Bible body where it
  belongs, instead of `BodyText` showing on page 1 as
  front-matter spacer.
- **USFM round-trip readiness**: `\v` ↔ `VerseText` is a
  one-to-one mapping; `\p`/`\ip` ↔ `BodyText` covers the
  residual.
- **Better Find / Replace targeting**: "find all verse text" =
  Find with `.Style = VerseText`. Currently no clean way to do
  that without paragraph-walking.
- **Easier i18n review**: translators know which paragraphs
  carry scripture vs commentary vs UI text.
- **Audit aid**: any verse-style paragraph that doesn't start
  with a Chapter Verse marker run is a structural defect; the
  conversion routine surfaces them as `nKept > 0` for paragraphs
  Word labeled BodyText but without a marker. Inverse audit
  (post-conversion: find any VerseText paragraph that doesn't
  start with a Chapter Verse marker) catches the other direction.

### Cons

- **One-time mass mutation** of the document. If something is
  miscategorized, undo is per-paragraph. Hence the explicit
  backup step before Phase 2.
- **Priority renumbering** for everything below the new
  `VerseText` slot - cascades through the snapshot and any
  doc that cites a specific priority number. Mitigation: only
  the `01-styles.md` snapshot and the `04-qa-workflow.md`
  current state cite specific numbers; both are easy to refresh.
- **Style count +1**, marginal cost.
- **Existing code that searches for `"BodyText"`** may now miss
  verses. Phase 3.4 audit catches this; effort proportional to
  current count of literals (likely small).
- **No backwards compatibility**: documents from before this
  conversion that still use `BodyText` for verses will look the
  same visually but parse differently. Acceptable for this
  project (single primary document) but worth noting.

### Risk register

| Risk | Likelihood | Mitigation |
|---|---|---|
| Conversion misses some verses | Low | Phase 3.1 audit; idempotent re-run |
| Conversion catches non-verse paragraphs that incidentally start with Chapter Verse marker | Very low | The character style appears only at verse boundaries by design |
| `Characters(1).Style.NameLocal` returns paragraph style instead of character style on some edge case | Low | Test a sample first; if so, walk Range.StoryRanges or use the Find pattern instead |
| Some BodyText paragraph inside a footnote story qualifies and gets converted | Low | Conversion iterates `ActiveDocument.Paragraphs` which is the main story by default; add explicit story-scope guard if footnotes turn out to also use BodyText with markers |
| Hardcoded `"BodyText"` literal in code path breaks after the change | Low-medium | Phase 3.4 grep audit catches them |

### What this does NOT change

- `BodyText` style definition itself (kept; just used less).
- `Chapter Verse marker` and `Verse marker` character styles
  (unchanged).
- Existing approved-array order for priorities 1-30 - those
  styles keep their current positions; only items at and below
  the new `VerseText` slot shift down.
- Any visual rendering in Word - VerseText is a property-for-
  property clone of BodyText.

### USFM mapping detail (added benefit)

Post-conversion, the implicit Bible structure becomes more
machine-readable. Sketch of the mapping:

| Word style | USFM marker | Notes |
|---|---|---|
| `Heading 1` | `\toc1` / book title | book boundary |
| `Heading 2` | `\c <n>` | chapter boundary |
| `Chapter Verse marker` (char) | (consumed by `\v`) | chapter number per verse |
| `Verse marker` (char) | `\v <n> ` | verse start |
| `VerseText` | `\v <n> <text>` body | verse content paragraph |
| `BodyText` | `\p` or `\ip` | non-verse paragraph |
| `Brief` | `\is` | book intro |
| `DatAuthRef` | `\d` | descriptive / authorship line |
| `Footnote Reference` | `\f` open | footnote anchor |
| `Footnote Text` | (footnote story) | footnote body |
| `Words of Jesus` | `\wj ... \wj*` | red-letter |
| `Brief` | `\is` | section intro |
| `Psalms BOOK` | `\ms` | major section (Psalms book divisions I-V) |

This is sketch quality; full mapping work belongs in the SBL
parser side and is separate from this conversion plan.

### Status

Plan: **DRAFT - awaiting user approval**. No code or document
changes applied. Phases are gated: Phase 2 should not start
until Phase 1 is complete and verified; Phase 3 follow-ups can
overlap with normal work.

---

## § Plan - verse-marker structural audit (precondition for VerseText) - 2026-04-26

### Goal

Walk the entire Bible body and verify that every chapter contains
exactly the expected number of well-formed verse markers - using
the known canonical book/chapter/verse counts in
`aeBibleCitationClass.cls` (`ChaptersInBook`,
`VersesInChapter`) as ground truth. Run this audit BEFORE the
`VerseText` conversion so any structural defects are surfaced
and fixed first, eliminating the risk of mass-converting around
a malformed marker.

User-noted history: occasional past discovery of paragraphs where
"Chapter Verse marker" or "Verse marker" runs were messed up. No
systematic check existed; this plan adds one.

### Invariants to verify

For the main story (excluding footnote/header/footer stories):

1. **Books present**: every book in the canonical 66-book list
   has a `Heading 1` paragraph in document order.
2. **Chapter counts per book**: count of `Heading 2` paragraphs
   between consecutive `Heading 1`s matches
   `ChaptersInBook(book)`.
3. **Verse counts per chapter**: count of `Verse marker`
   character-style runs between consecutive `Heading 2`s
   matches `VersesInChapter(book, chapter)`.
4. **Marker pairing**: every `Verse marker` run is immediately
   preceded by a `Chapter Verse marker` run. (Per the existing
   `GoToVerseByScan` comment: "Every verse ... begins with a
   Chapter Verse marker run ... immediately followed by a
   Verse marker run.")
5. **Marker text**: each `Verse marker` run's text is a numeric
   string equal to the verse's expected sequence number
   (1, 2, 3, ..., N). Each `Chapter Verse marker`'s text is the
   chapter number for its containing chapter.
6. **No stray markers**: no `Verse marker` or `Chapter Verse
   marker` runs appear inside `Heading 1`, `Heading 2`,
   `Brief`, `DatAuthRef`, `Footnote Text`, or other non-verse
   contexts.

### Test design

Single Public entry point + several private helpers in a new
module. Read-only; produces a report file plus an Immediate
summary; flags every discrepancy with book + chapter + verse
context.

```vba
Public Sub AuditVerseMarkerStructure(Optional bWriteFile As Boolean = True)
    ' For each book in canonical order:
    '   For each chapter in that book:
    '     Find chapter range (Heading 2 .. next Heading 2 / Heading 1)
    '     Count Verse marker runs in range
    '     Compare to VersesInChapter
    '     For each Verse marker run:
    '       Check preceding run is Chapter Verse marker
    '       Check Verse marker text == expected number
    '       Check Chapter Verse marker text == chapter number
    ' Report: total expected vs actual verse counts per book and overall
End Sub
```

Output to `rpt/VerseStructureAudit.txt` (plus Immediate summary).
Format:

```
---- AuditVerseMarkerStructure: 2026-04-26 hh:nn:ss ----

Genesis              expected chapters=50  found=50   OK
  ch  1: expected verses=31   found=31   OK
  ch  2: expected verses=25   found=25   OK
  ...
Exodus               expected chapters=40  found=40   OK
  ...
Psalms               expected chapters=150 found=150  OK
  ch  119: expected verses=176 found=176 OK
  ...

ISSUES FOUND:
  Genesis 3:14   Verse marker text "14" but preceding run style is "BodyText" (expected "Chapter Verse marker")
  Exodus  20:17  Chapter Verse marker text "21" (expected "20")
  Psalms 119:177 unexpected extra Verse marker run after verse 176

SUMMARY: 31102 / 31102 verses found, 3 structural issues, 0 missing chapters, 0 missing books
```

### Implementation steps

#### Phase 1 - module + skeleton

- [ ] **1.1** New module `src/basVerseStructureAudit.bas` with
  `Option Explicit` / `Option Compare Text` /
  `Option Private Module` per project convention.
- [ ] **1.2** `Public Sub AuditVerseMarkerStructure` entry
  point with optional `bWriteFile` (default True). Uses
  `StartTimer` / `EndTimer` from `basStyleInspector` for
  expected/actual feedback (sub may take 30-60 sec on a full
  Bible; benefits from timing).
- [ ] **1.3** Private helper `IterateBooks` that walks
  `ActiveDocument.Paragraphs`, identifies each `Heading 1`,
  matches its text against the canonical book name list, and
  yields `(bookIndex, bookName, headingPos)` tuples.
- [ ] **1.4** Private helper `IterateChapters(bookRange)` that
  walks the book's range, identifies each `Heading 2`, and
  yields `(chapterNum, headingPos, endPos)` tuples for the
  chapter sub-range.
- [ ] **1.5** Private helper `CountVerseMarkers(chapterRange)`
  that uses `Range.Find` with `.Style = "Verse marker"` to
  count occurrences.

#### Phase 2 - invariant checks

- [ ] **2.1** Implement invariant 1 (books present) using book
  iteration vs canonical 66-name list.
- [ ] **2.2** Implement invariant 2 (chapter counts per book)
  using chapter iteration count vs `ChaptersInBook`.
- [ ] **2.3** Implement invariant 3 (verse counts per chapter)
  using verse-marker count vs `VersesInChapter`.
- [ ] **2.4** Implement invariant 4 (marker pairing) by walking
  each Verse marker run and inspecting the run immediately
  before it via Range.Characters or Range.MoveStart -1.
- [ ] **2.5** Implement invariant 5 (marker text correctness)
  by comparing `Verse marker` run text to the iteration index
  and `Chapter Verse marker` run text to the current chapter
  number.
- [ ] **2.6** Implement invariant 6 (no stray markers) by
  scanning Heading 1, Heading 2, Brief, DatAuthRef, etc.
  paragraph ranges for any Verse marker / Chapter Verse marker
  runs.

#### Phase 3 - report + integration

- [ ] **3.1** Format the per-book / per-chapter / issues sections
  as shown in the design sample. Color: ASCII only (per VBA
  feedback memory; report is read in many places, not just the
  VBE).
- [ ] **3.2** Write to `rpt/VerseStructureAudit.txt`. Also
  print summary line and any issues to Immediate window.
- [ ] **3.3** Integrate into the test suite. Position:
  **before** any taxonomy / style audit, since marker structure
  is a precondition. Concretely: when `SUPER_TEST_RUNS` lands,
  add `AuditVerseMarkerStructure` as Suite 0 (precondition).
  Until then, add as a manual run referenced by the
  conversion plan in the prior section.
- [ ] **3.4** Add to the EDSG `02-editing-process.md` as a
  pre-conversion verification step.

### Pros

- **Catches structural defects** before downstream work
  (VerseText conversion, USFM export, parser tests) silently
  consumes them.
- **Concrete pass/fail signal** based on canonical counts -
  no human review fatigue across 31,102 verses.
- **Reusable**: any future structural change benefits from the
  same audit as a precondition.
- **Documents invariants explicitly**: the test code is the
  authoritative statement of "what a well-formed verse looks
  like."
- **i18n robustness**: a translated Bible with shifted text
  lengths still has the same verse count - this audit confirms
  marker structure survived the translation work.
- **CI heartbeat**: weekly automated run flags any regression
  introduced by ongoing edits without waiting for the next
  conversion to surface them.

### Cons

- **Runtime** ~30-60 sec on a full Bible (one Find per chapter +
  per-run preceding-style check). Acceptable as a manual /
  weekly check; not for every commit.
- **Maintenance**: if marker conventions change (new style for
  Psalm acrostic verses, for example), the audit needs an
  update.
- **Edge-case complexity**: Psalms 119 (acrostic), Psalms with
  superscriptions, Selah mid-verse - all need to be modeled
  correctly so they don't generate false positives.
- **Discovers issues but doesn't fix them**: each flagged issue
  still needs manual inspection and correction.

### Benefits

- **Confidence baseline** for the VerseText conversion. If the
  audit passes, the conversion is safe; if it fails, fix the
  flagged issues first.
- **Prevents silent navigation errors**: a malformed Verse
  marker today causes `GoToVerseByScan` to misnavigate (the
  Nth Verse marker is the wrong verse). The audit catches
  these proactively.
- **Foundation for SUPER_TEST_RUNS**: a clean verse-structure
  audit is the natural Suite 0 of any global verification
  command.
- **Translator-friendly**: a translator working on a localized
  Bible can run the audit on their work-in-progress to confirm
  no markers were broken.

### Special cases to handle

| Case | Concern | Approach |
|---|---|---|
| Psalm 119 acrostic headings (Aleph, Beth, ...) | Non-verse paragraphs interspersed with verses | `IterateChapters` already handles via Heading 2 boundaries; verse counter only counts Verse marker runs, ignores acrostic paragraphs |
| `PsalmSuperscription` paragraphs | Non-verse paragraphs in some Psalms (e.g., Psalm 3 "A Psalm of David...") | Verse counter ignores them automatically |
| `Selah` | Mid-verse interjection | Should not contain Verse marker; if invariant 6 catches Verse markers inside Selah runs, flag |
| Psalms book division headings (`Psalms BOOK`) | Sit between Psalms book "chapters" | Skip when iterating chapters (not Heading 2) |
| Very short books (Obadiah, 1-3 John, Jude, Philemon) | One chapter only | Iteration handles same as multi-chapter; expected verse count from `VersesInChapter(book, 1)` |
| Books not in current document (BodyTextContinuation et al placeholders) | Already missing per `PromoteApprovedStyles` warning | Audit ignores - they're style placeholders, not book content |

### Decision items

1. Module name: `basVerseStructureAudit` proposed - confirm or
   override.
2. Output filename: `rpt/VerseStructureAudit.txt` proposed;
   alternative `rpt/Styles/VerseStructureAudit.txt` to keep all
   audits under `rpt/Styles/`. (Argument for the latter: the
   audit IS about style usage. Argument for the former: it's
   about document structure, not style definitions.)
3. Test-suite integration timing: add as standalone Public Sub
   immediately, then wire into `SUPER_TEST_RUNS` when that
   command lands. Or wait for `SUPER_TEST_RUNS` and add both
   together. Standalone-first is recommended (does not block on
   the deferred suite; the audit is independently useful for
   the VerseText conversion).

### Status

Plan: **DRAFT - awaiting user approval**. No code applied. The
user noted that "now vs later" is moot - this should run before
the VerseText conversion. Recommended sequence:

1. Approve this audit plan.
2. Implement and run the audit.
3. Resolve any flagged issues.
4. Then (re-)approve and run the VerseText conversion plan from
   the previous section.

### Decisions confirmed and module created - 2026-04-26

User approved:

- Module name: `basVerseStructureAudit` ✓
- Output filename: `rpt/VerseStructureAudit.txt` ✓
- Standalone-first integration (don't wait on `SUPER_TEST_RUNS`) ✓

`src/basVerseStructureAudit.bas` created. Implements core
invariants 1-3:

- `Public Sub AuditVerseMarkerStructure(Optional bWriteFile = True)` -
  entry point.
- `PopulateCanonical` - hardcodes the 66 canonical book names +
  chapter counts (mirrors `basTEST_aeBibleCitationClass`; same
  list, including project's "Solomon" rather than SBL's "Song").
- Walks Heading 1 paragraphs in the document, matches each to a
  canonical book ID by case-insensitive name compare, then
  delegates per-book analysis to `AuditOneBook`.
- `AuditOneBook` walks Heading 2 paragraphs in the book range,
  compares count to `ChaptersInBook` (passes/fails per book),
  then for each chapter calls `CountVerseMarkers` and compares
  to `aeBibleCitationClass.VersesInChapter(book, ch)`.
- `CountVerseMarkers` runs `Range.Find` with style filter in a
  loop, advancing past each match until end-of-chapter.
- Report format: per-book one-liner with per-chapter detail
  underneath; final SUMMARY with totals; ISSUES FOUND and
  MISSING BOOKS sections appended when applicable.
- Uses `StartTimer` / `EndTimer` from `basStyleInspector` for the
  expected/actual timing pattern.
- Output goes to `rpt/VerseStructureAudit.txt` (and Immediate).

Invariants 4-6 (marker pairing, marker text correctness, no
stray markers in non-verse paragraphs) deferred to a follow-up
once the core counts are clean. The verse-count mismatch alone
catches the "messed up markers" regression the audit was
designed to surface; the deeper invariants are belt-and-suspenders.

### Status

Audit module: **IMPLEMENTED - 2026-04-27** in
`src/basVerseStructureAudit.bas`. Awaiting first run to verify
clean results before VerseText conversion is approved.

Recommended invocation:

```vba
AuditVerseMarkerStructure              ' writes rpt/VerseStructureAudit.txt
```

---

## 2026-04-28 — VerseStructureAudit follow-up + Song-of-Songs canonicalization

Three concerns surfaced from the first run of `AuditVerseMarkerStructure` (output in `rpt/VerseStructureAudit.txt`) and a related decision to standardise on SBL `"Song of Songs"` instead of the project's prior `"Solomon"`. Findings and recommended fixes below — presented one at a time per the standard fix-process.

### Finding 1 — chapter/verse report accumulates across books (audit bug)

**Symptom.** Every book past Genesis prints a chapter-detail block that contains the previous books' chapter lines as well. E.g. the `Exodus` line at `rpt/VerseStructureAudit.txt:56` correctly reports `expected chapters=40 found=40 OK`, but the per-chapter detail underneath runs `ch 1..50` (Genesis's 50) followed by `ch 1..40` (Exodus's actual 40). Same pattern for every subsequent book; e.g. `Solomon` at line 6976 also shows `ISSUES` even though every chapter line is `OK`.

**Root cause.** `basVerseStructureAudit.AuditVerseMarkerStructure` declares `chapterReport`, `bookIssues`, `bookIssueDetail` *inside* the `For i = 1 To nH1` loop. In VBA, `Dim` inside a loop is hoisted to procedure scope and the variables are **not** re-initialised per iteration. `AuditOneBook` is `ByRef` on these variables and appends/increments without first clearing them, so the running totals leak from one book into the next. `bookIssues` accumulating is also why books with zero per-chapter mismatches still print `ISSUES` once any earlier book has a chapter-count mismatch.

**Recommended fix.** Reset the three accumulators at the top of each loop iteration in `src/basVerseStructureAudit.bas`:

```vba
For i = 1 To nH1
    chapterReport = vbNullString
    bookIssues = 0
    bookIssueDetail = vbNullString
    ' ... existing body ...
Next i
```

Local fix, low risk, no signature changes.

### Finding 2 — DRY violation: audit module owns its own canonical 66-book table

**Symptom.** `basVerseStructureAudit` defines `PopulateCanonical` (66-book name + chapter-count list) and `LookupBookID`, duplicating data already authoritative in `aeBibleCitationClass.GetCanonicalBookTable` (`aeBibleCitationClass.cls:886`) and the alias map (`:1361`). The two copies have already diverged (`names(22) = "Solomon"` in audit vs `"Song of Songs"` in the class).

**Root cause.** Initial implementation mirrored `basTEST_aeBibleCitationClass.bas:864-929`, which also keeps a private copy. `aeBibleCitationClass` already exposes everything the audit needs:

- `ChaptersInBook(bookName) As Long` — `:1956`
- `VersesInChapter(bookName, chapter) As Long` — `:1988` (already used)
- `GetCanonicalBookTable() As Object` — full `BookID → (id, name, chapters)` triples — `:886`
- `ResolveAlias(abbr, ByRef BookID) As String` — alias-tolerant lookup that also returns the canonical name — `:1896`

**Recommended refactor.** Replace `PopulateCanonical` and the local `canonNames` / `canonChapters` arrays with a small adapter that pulls from `GetCanonicalBookTable` once at the top of `AuditVerseMarkerStructure`:

```vba
Dim books As Object
Set books = aeBibleCitationClass.GetCanonicalBookTable
' books(k)(0)=BookID  books(k)(1)=Canonical name  books(k)(2)=ChaptersInBook
```

`LookupBookID` becomes a single `ResolveAlias` call wrapped in `On Error Resume Next` (so unknown H1 text still produces the existing `?? UNKNOWN H1` line rather than a hard error). This automatically picks up the new `"Song of Songs"` canonical name and any future SBL alias additions, eliminates the hand-maintained list, and aligns with the existing pattern already used at `:199` (`aeBibleCitationClass.VersesInChapter(...)`).

Note: `basTEST_aeBibleCitationClass.bas` keeps a deliberate independent copy as the **expected oracle** for the test; that is correct (test data must not derive from the system under test). The audit, by contrast, is a consumer of canonical truth, not its validator — so it should call into the class.

### Finding 3 — `ToSBLShortForm` lookup with `"Song of Songs"`

**What changed.** Switching project canonical from `"Solomon"` to `"Song of Songs"` means every Heading 1 / canonical reference passing through SBL routines now contains a multi-word book name with **two internal spaces**. The user reports a lookup error in `ToSBLShortForm` after the change.

**Static-analysis result — honest read.** `aeBibleCitationClass.ToSBLShortForm` (`:2845`) splits the input on the **last** space, which correctly yields `bookName = "Song of Songs"`, `numPart = "1:1"` for an input of `"Song of Songs 1:1"`. The alias map already has `"SONG OF SONGS" → 22` (`:1474`), so `ResolveAlias` should return `bID=22`, `abbr="Song"`, output `"Song 1:1"`. **From static analysis I cannot reproduce a lookup failure for `"Song of Songs N:V"` style input** — the last-space split is robust against multi-word book names with internal spaces.

**Three plausible failure modes worth checking before editing.** Honest recommendation: before patching `ToSBLShortForm`, capture the actual failing input and the `Err.Description` from the `MsgBox` it raises. Candidates:

1. **Input lacks a chapter:verse suffix.** If `canon = "Song of Songs"` (book name only, no `" 1:1"` tail), the last-space split produces `bookName = "Song of"`, `numPart = "Songs"` — `ResolveAlias("Song of")` is not in the alias map and raises `vbObjectError + 10 "Unknown book alias: Song of"`. The same bug applies to any multi-word book passed without a chapter suffix (`"1 Samuel"`, `"Song of Songs"`, etc.); it just never tripped before because the project canonical (`"Solomon"`) was a single token.
2. **Hidden whitespace.** A non-breaking space (`Chr(160)`) or accidental double space inside `"Song of Songs"` would not match `"SONG OF SONGS"` after `UCase$ Trim$` (Trim only strips outer whitespace). If the source data was retyped, this is plausible.
3. **Sister test in `Test_CanonicalNamesAndSBLTable`** (`basTEST_aeBibleCitationClass.bas:970-980`) extracts the abbreviation by `InStr(sblResult, " ")` — first space, not last. That logic is **already broken** for any book whose SBL abbreviation contains a space (`"1 Sam 1:1"` returns `"1"` rather than `"1 Sam"`). It happens to pass for `"Song"` only by coincidence (`"Song"` is one word). This is a pre-existing test-side bug, not a `ToSBLShortForm` bug, but flips into view as soon as the canonical name changes if the user is reading the failure message and assuming `ToSBLShortForm` itself is at fault.

**Recommended next step.** Hold the `ToSBLShortForm` patch until we see the failing input. If the failure is mode (1), the right fix is a defensive guard: when no space is found *after* a successful book-name lookup, fall back to whole-input alias resolution. If mode (2), the fix is in the data, not the code (or normalise internal whitespace in `ResolveAlias`). If mode (3), fix the test extractor to use `InStrRev`.

### Finding 4 — broader citation-code impact of the rename

A grep for `Solomon` / `Song of Songs` shows the rename touches more than just `ToSBLShortForm`:

| File | Line(s) | Status |
|---|---|---|
| `src/aeBibleCitationClass.cls` | 913 | Already `"Song of Songs"` ✓ |
| `src/aeBibleCitationClass.cls` | 1474, 1826 | Comments still reference `"Song of Solomon"`; alias map already accepts both. Cosmetic. |
| `src/basVerseStructureAudit.bas` | 301 | Still `"Solomon"` — fixed by Finding 2 refactor. |
| `src/basTEST_aeBibleCitationClass.bas` | 885 | Test oracle still `"Solomon"`. **Must change to `"Song of Songs"`** if the canonical is now Song of Songs, otherwise `Test_CanonicalNamesAndSBLTable` will read stale. |
| `src/basSBL_VerseCountsGenerator.bas` | 95 | Generator label `"Solomon"` — review whether this affects the generated `GetVerseCounts` map keys; if it's only a debug label it's fine, but worth confirming. |
| `src/XbasTESTaeBibleDOCVARIABLE.bas` | 527 | `VerifyBookNameFromDocVariable "Song", "Solomon"` — this expectation is **document-specific** (already flagged on `2026-03-16` review). If the source `.docm` Heading 1 is being changed too, update to `"Song of Songs"`; if not, leave as-is. |
| `md/Deterministic Structural Parser.md` | 83, 314 | Reference table + multi-word example — both should be updated. |
| `md/Editorial Design and Style Guide.md` and `EDSG/*.md` | n/a | No current references; no edits required. |
| `rvw/*` (older review docs) | many | **Do not retro-edit** — review docs are progressive history. |

### Suggested order of operations

1. Apply Finding 1 (accumulator reset) — purely local audit-module fix. **APPLIED — 2026-04-28**
2. Re-run `AuditVerseMarkerStructure` and confirm the per-book chapter detail lines match expected counts and `Solomon`/Song-of-Songs no longer shows phantom `ISSUES`. **CONFIRMED — 2026-04-28**
3. Apply Finding 2 (DRY refactor) — drops `PopulateCanonical` and `LookupBookID`; the audit module shrinks substantially and inherits the class's `"Song of Songs"` canonical name automatically. **APPLIED — 2026-04-28** (see diff summary below)
4. Capture the actual `ToSBLShortForm` failure (mode 1 / 2 / 3 from Finding 3) before patching the citation class.
5. Sweep `basTEST_aeBibleCitationClass.bas:885` and `md/Deterministic Structural Parser.md:83,314` to align the rename.

### Finding 2 — applied diff summary, `src/basVerseStructureAudit.bas`

- **`AuditVerseMarkerStructure`**: replaced local `canonNames(1 To 66)` / `canonChapters(1 To 66)` arrays plus the `PopulateCanonical` call with one line — `Set books = aeBibleCitationClass.GetCanonicalBookTable`. The `books` dictionary returns `Array(BookID, name, chapters)` per BookID; reads use `books(BookID)(1)` for name and `books(BookID)(2)` for chapter count.
- **`LookupBookID`**: rewritten as a thin wrapper around `aeBibleCitationClass.ResolveAlias` (with `On Error Resume Next` so unknown H1 text returns `0` and produces the existing `?? UNKNOWN H1` line). Drops the case-insensitive scan over local `canonNames`. Now accepts every alias the citation class recognises (`"Song of Songs"`, `"Song"`, `"Solomon"`, `"SG"`, etc.) without modification to the audit module.
- **`PopulateCanonical`**: deleted entirely (66 lines + helper sub).

Net: ~80 lines removed, single source of truth restored. The audit module now consumes canonical data from `aeBibleCitationClass` exactly as the rest of the codebase does.

Expected effect on next audit run: the `MISSING BOOKS: Solomon` and `Unknown H1 text: [SONG OF SONGS]` entries both clear (alias map already covers both); only the deferred document-content items (Romans 14, Hebrews 7) should remain.

### Finding 2 — follow-up bug surfaced by the DRY refactor

After Finding 2 was applied, the next audit run reported a new mismatch: `Nahum: chapter Count mismatch (expected 7, found 3)`. The document has 3 chapters in Nahum, which is correct (Nahum is a 3-chapter book; cross-referenced against `basSBL_VerseCountsGenerator.bas:107` which holds a 3-element array). The "expected 7" came from a stale value in the canonical book table at `aeBibleCitationClass.cls:925`:

```
books.Add 34, Array(34, "Nahum", 7)   ' WRONG — Nahum has 3 chapters
```

This bug was previously masked because the audit module kept its own correct copy in `PopulateCanonical` (`chapters(34) = 3`). The DRY refactor (Finding 2) routed the audit through the canonical table and surfaced the latent error. The verse-counts table was always correct — only the chapter-count metadata was wrong.

**Fix applied — 2026-04-28**: `aeBibleCitationClass.cls:925` updated to `books.Add 34, Array(34, "Nahum", 3)`. One-line edit.

Verse-total sanity check after the fix: 31,104 found vs 31,102 expected. The +2 delta corresponds exactly to the two deferred document-content items: Rom 14 (+3, contains 14:24-26 doxology) and Heb 7 (−1, missing one verse). No further structural mismatches expected.

### Correction — Romans doxology placement (WEB vs TR)

The earlier classification of Romans 14 as a "document content bug" was based on a misreading of the WEB translator note. Re-reading the note as published on eBible.org:

> "TR places Romans 14:24-26 at the end of Romans instead of at the end of chapter 14, and numbers these verses 16:25-27."

Subject of the sentence is **TR** (Textus Receptus — the Greek text underlying the KJV). The correct reading:

- **WEB**: doxology at end of **Romans 14**, numbered **14:24-26** → Rom 14 = **26** verses, Rom 16 = **24** verses.
- **TR / KJV tradition**: doxology at end of **Romans 16**, numbered **16:25-27** → Rom 14 = 23, Rom 16 = 27.

The verse-counts table at `basSBL_VerseCountsGenerator.bas:119` was originally seeded with the TR pattern (`...14, 23, 33, 27`), not the WEB pattern. Since the project source is the WEB Protestant Edition, the data must be corrected to the WEB placement.

**Fix applied — 2026-04-28**: `basSBL_VerseCountsGenerator.bas:119` Romans array updated:

| Chapter | Before (TR) | After (WEB) |
|---|---|---|
| Romans 14 | 23 | **26** |
| Romans 16 | 27 | **24** |

Net book total unchanged (compensating shift).

### Reclassification of the remaining audit issues

After this Romans correction the two outstanding items shift:

| Item | Status before correction | Status after correction |
|---|---|---|
| Rom 14 (26 found vs 23 expected) | document content bug | **CLEARS** — doc already matches WEB |
| Rom 16 (now 27 found vs 24 expected) | not flagged | **NEW ISSUE** — duplicate doxology in document at 16:25-27 |
| Heb 7 (27 found vs 28 expected) | document content bug | unchanged — still a missing verse marker |

User-confirmed reading of the source document: the doxology appears in **both** Rom 14 and Rom 16 with "slight differences" between the two copies. Likely an editorial merge from a TR-based source into a WEB-based source. Document-side fix is to delete the doxology from Romans 16 (verses 25-27) — WEB places it only at 14:24-26.

### Final clean state — 2026-04-28

Document-content fixes applied by user between audit runs:

- **Hebrews 7**: missing verse marker repaired in the `.docm`. Chapter now reports 28/28 OK.
- **Romans 16**: duplicate doxology (16:25-27) deleted from the `.docm`. Chapter now reports 24/24 OK. WEB keeps the doxology only at 14:24-26, which the document already had.

Final audit run:

```
SUMMARY: 31102 / 31102 verses found, 0 structural issue(s).
AuditVerseMarkerStructure - actual 66.52 sec
```

The found total **31,102** matches the KJV / WEB Protestant canon total documented in `basSBL_VerseCountsGenerator.bas:14-21`. All 66 books pass; all 1,189 chapters pass; all per-chapter verse counts match the WEB-aligned source data. The structural-audit baseline is clean.

### Summary of all 2026-04-28 work

| # | Finding | File(s) | Status |
|---|---|---|---|
| 1 | Per-book accumulator reset bug | `src/basVerseStructureAudit.bas` | APPLIED + CONFIRMED |
| 2 | DRY refactor — consume `aeBibleCitationClass.GetCanonicalBookTable` | `src/basVerseStructureAudit.bas` | APPLIED + CONFIRMED |
| 2.1 | Latent bug exposed: Nahum chapter count `7 → 3` | `src/aeBibleCitationClass.cls:925` | APPLIED + CONFIRMED |
| 3 | `ToSBLShortForm` "Song of Songs" lookup | `src/aeBibleCitationClass.cls` | DIAGNOSED — no defect reproducible from static analysis; awaiting failing input from user |
| 4 | Reference rename impact (Solomon → Song of Songs) | various test/md files | INVENTORIED — `basTEST_aeBibleCitationClass.bas:885` and `md/Deterministic Structural Parser.md:83,314` still pending |
| WEB-1 | OT Hebrew→English versification fixes (2 Sam 18/19, 2 Kgs 11/12, 2 Chr 13/14, 3 John) | `src/basSBL_VerseCountsGenerator.bas` | APPLIED + CONFIRMED |
| WEB-2 | Romans doxology placement corrected from TR pattern to WEB pattern (14=26, 16=24) | `src/basSBL_VerseCountsGenerator.bas:119` | APPLIED + CONFIRMED |
| DOC-1 | Hebrews 7 missing verse marker | `.docm` content | RESOLVED by user |
| DOC-2 | Duplicate doxology at Romans 16:25-27 | `.docm` content | RESOLVED by user |

Audit baseline: **clean — 31,102 / 31,102, 0 issues**.

---

## 2026-04-29 — Finding 3 closure + new Finding 5 (nav sync)

### Finding 3 — closed

User confirms ribbon navigation works for `"Song"` after the canonical rename: `aeBibleCitationClass.ResolveAlias("Song" | "Song of Songs" | "Solomon")` all return BookID 22, and `ToSBLShortForm` outputs `"Song"` correctly. No defect was reproducible from static analysis on the alleged "lookup error", and live-run confirms the alias path is sound. **CLOSED — 2026-04-29**.

### Finding 5 — first-click navigation sync (NEW)

**Symptom (user reproduction).**
1. Open the `.docm`. Pick book/chapter/verse in the ribbon. Click **Go**.
2. The status bar shows that nav happened, but **the cursor (caret) does not land in the document body**.
3. Pressing **Next Verse** "forces" the cursor into the document — it appears at the expected verse + 1.
4. Pressing **Next Chapter** *after that workaround* lands at the **wrong position** — somewhere off from the actual chapter H2.

User's own diagnosis: "*Some sync issue similar to the reason the cache was setup.*" That instinct is consistent with what the code shows.

**Code-side context.** Three relevant pieces in `src/aeRibbonClass.cls`:

- `OnGoClick` (`:1054`) → `GoToVerse vsNum` (`:1079`) — fires synchronously from the ribbon button click. Ribbon owns focus when this runs.
- `GoToVerse` (`:947`) → `GoToVerseByScan chPos, vsNum` (`:976`) — uses the documented three-step pattern (`ScrollIntoView` + `Selection.SetRange`) at `:1032-1033`. Caches `m_currentChapterPos = chPos` (`:965`) for subsequent `GoToVerse` calls in the same chapter.
- The codebase already has a deferred-navigation pattern: `basRibbonDeferred.UpdateStatusBarDeferred` (`:303`, `:343`, `:678`, `:981`) and `FocusBookDeferred` (Bug #597) — both use `Application.OnTime Now, ...` to let the ribbon click event clear before touching focus or selection. This was introduced specifically because operations done while ribbon focus is still alive get swallowed (Bug 21 — "*ScrollIntoView steals ribbon focus from OnTime context*").

**Hypothesis (honest, not yet verified).**

The first `OnGoClick` runs synchronously from a ribbon-owned event. `Selection.SetRange` writes the document's stored selection but Word does not render the caret until focus actually returns to the document body. The user then doesn't see the cursor, so they press **Next Verse**, which goes through a different code path (`OnNextVerseClick` → `GoToVerse m_currentVerse + 1`) that runs at a later message-pump tick when ribbon focus has cleared — and at that point the caret materialises. The subsequent **Next Chapter** computes a position relative to a `m_currentChapterPos` cache value that was populated in the first (un-rendered) call, before final layout had stabilised — so the H2 search range or the chapterData index may have been against partially-laid-out content, producing an off-by-some position.

This matches three of the existing comments in the source:
- `:200-201` — "*The first GoTo Book will warm the layout on demand (~12s, once per session)*" — first nav crosses an un-warmed layout boundary.
- `:284-296` — long block describing the three-step ScrollIntoView+SetRange dance and Bug 19 ("*ScrollIntoView does not move the document cursor*").
- `:600-608` — "*ScrollIntoView was here, Tab key presses ... were routed through ribbon*" — confirms that when ScrollIntoView fires from a ribbon-focus context, key routing breaks until focus returns to the document.

**Three candidate fixes (no edits applied).**

1. **Force document focus after OnGoClick.** Add `ActiveDocument.ActiveWindow.Activate` (or `Application.ActiveWindow.SetFocus` equivalent) at the end of `GoToVerse` after the ScrollIntoView+SetRange. *Risk:* may re-trigger Bug 21–style focus thrash on the ribbon; could cause keytip handling to lose state.
2. **Force layout completion before caching positions.** Insert `ActiveDocument.ComputeStatistics(wdStatisticPages)` (or a `DoEvents` loop) before `m_currentChapterPos = chPos` so the cached position reflects post-layout coordinates. *Risk:* adds latency to first nav; doesn't fix the caret-not-visible part.
3. **Defer the actual navigation via `Application.OnTime`** — mirrors the existing `FocusBookDeferred` / `UpdateStatusBarDeferred` pattern. `OnGoClick` would set `m_pendingVerse = vsNum` and schedule a `GoToVerseDeferred` routine; the deferred routine runs after the ribbon click event clears, when the document has focus and layout has settled. The `m_pendingVerse` plumbing already exists at `:48` and `:1127-1132` — partly stubbed out as a no-op (`'navigation trigger moved to OnGoClick (#600); m_pendingVerse is never set so this is a permanent no-op'`) but the structure is there to revive.

**Recommendation.** Option 3 is the most consistent with this codebase's existing patterns (all the other ribbon → document operations that touched focus or layout were eventually deferred for the same reason). It's also the lowest-risk: `Application.OnTime Now` was already proven safe by `UpdateStatusBarDeferred` and `FocusBookDeferred`. Option 1 alone won't solve the cached-position drift in step 4; option 2 alone won't solve the caret-not-visible in step 2. Option 3 plausibly addresses both, by running `GoToVerseByScan` at a tick when focus and layout are both settled.

**Before editing, I'd like to confirm:**

1. Repro is the **first nav after document open**, every time? Or only on some books/chapters?
2. Does it also happen if the user clicks somewhere in the document body **first** (giving the doc focus), and then uses ribbon Go?
3. After step 4 (off-by-some chapter position), is the offset always in the same direction, or does it vary?

The answer to (2) in particular discriminates between the focus hypothesis and the layout-cache hypothesis. If clicking in the document first makes the bug go away, it's pure focus (option 1 alone might be enough). If it persists even with prior document focus, it's layout-sync (option 3 is needed).

**Status:** DIAGNOSIS — awaiting reproduction details and direction on which option to pursue.

### Finding 5 — reproduction confirmed (2026-04-29)

User-confirmed reproduction details:

1. **When the cursor is already in the `.docx` body, ribbon Go places the caret correctly.** → confirms the failure is a **focus issue**, not a layout-cache issue. Selection mutations from ribbon-owned event handlers don't render a caret because Word only paints the I-beam when the document body owns focus.

2. **Entering a chapter number, then pressing Tab Tab quickly, sends the Tabs into the document body** (two tab characters typed). → initially mislabelled as a "focus race"; corrected diagnosis below — it is a **deterministic idle-commit focus leak**, not a race. See the 2026-04-29 terminology correction further down.

These two together rule out the layout/cache hypothesis as primary. It is purely a focus-handoff problem. The chapter-position drift on the *next* nav (the original step 4) is then explained as a consequence of the first nav running while focus was still on the ribbon — `Selection.SetRange` updates the stored selection but `ActiveWindow.ScrollIntoView` does not move the rendered viewport in the same way it would with document focus, leaving subsequent position computations against a partially-positioned view.

### Finding 5 — refined recommendation

Two-part fix, both small:

**A. Force document focus at the end of `GoToVerse`.** After the existing `ScrollIntoView` + `SetRange` (`aeRibbonClass.cls:1032-1033`), add a single line that transfers focus from the ribbon to the document body:

```vba
ActiveDocument.ActiveWindow.Activate
```

This is the canonical Word-API call for "give focus to this document window" and is what the codebase already uses elsewhere for window activation. It runs at the end of the navigation, so it doesn't disturb the ScrollIntoView/SetRange ordering.

**B. Defer the entire `GoToVerse` call from `OnGoClick` via `Application.OnTime`** — mirrors `FocusBookDeferred` (Bug #597) and `UpdateStatusBarDeferred`. The `m_pendingVerse` plumbing at `aeRibbonClass.cls:48` and `:1127-1132` is already in place but stubbed; reviving it means `OnGoClick` sets `m_pendingVerse = vsNum` and schedules `basRibbonDeferred.GoToVerseDeferred` (a new short routine that calls `ExecutePendingVerse`) via `Application.OnTime Now, ...`. By the time the deferred routine fires, the ribbon click event has cleared and focus naturally returns to the document body before navigation runs.

**Why both, not one or the other:** (A) alone fixes the caret-not-visible symptom but does not fix the Tab Tab → document race, because that race occurs *before* OnGoClick runs at all (it's between editBox commits and the next Tab). (B) alone fixes the caret-not-visible symptom *and* sidesteps the Tab race for Go, but does not protect against any future ribbon → document operation that hasn't been ported to the deferred pattern yet. Together they form belt-and-suspenders: (B) makes nav focus-safe by construction, and (A) is a defensive activation at the end so any path into `GoToVerse` (including direct calls from Prev/Next Verse buttons) ends with the document holding focus.

**The Tab behaviour in step (2)** is a separate Word/customUI focus-management item and is *not* fixed by either A or B. The corrected diagnosis (below) shows it is the platform's documented `editBox` idle-commit returning focus to the document, after which any Tab is a document Tab. There is no VBA-side fix; KeyTips are the supported path. Flagging as a separate item rather than mixing it into the same fix.

**Recommendation:** apply A first as a one-line, low-risk defensive fix and re-test. If symptom (1) clears with just A, leave B for later; if any residual focus issues remain, layer B on top. This way each change can be evaluated independently — consistent with the project's one-fix-at-a-time review pattern.

**Status:** AWAITING APPROVAL — propose to apply (A) only as the first step.

### Finding 5 — fix (A) applied 2026-04-29

`src/aeRibbonClass.cls` `GoToVerseByScan` — added `ActiveDocument.ActiveWindow.Activate` immediately after `Selection.SetRange`, before the `ScreenUpdating = True` that ends the navigation block. Three-line addition (one statement + two-line comment).

```vba
ActiveWindow.ScrollIntoView rVsView, True
Selection.SetRange Start:=r.Start, End:=r.Start
' Transfer focus from the ribbon to the document body so Word
' renders the caret. Without this the Selection is updated but
' the I-beam stays invisible until another action moves focus.
ActiveDocument.ActiveWindow.Activate
Application.ScreenUpdating = True
```

Why placed inside `GoToVerseByScan` and not in `OnGoClick`: every navigation path through the class — `OnGoClick`, Prev/Next Verse buttons, Prev/Next Chapter buttons, `ExecutePendingVerse` — funnels to `GoToVerse` and from there to `GoToVerseByScan`. Activating at the deepest common point ensures all entry points end with the document holding focus, without duplicating the call.

Test path:
1. Open the `.docm` (don't click in the document).
2. Pick book/chapter/verse in the ribbon. Click Go.
3. Caret should now appear at the target verse without needing a Next-Verse "force" press.
4. Press Next Chapter — should land at the correct H2 (no off-by-some).

If symptoms persist, layer fix (B) on top. The Tab Tab → document behaviour in step (2) of the original repro remains a separate item — corrected diagnosis below identifies it as Word's documented idle-commit, not a project defect.

**Status:** APPLIED — awaiting test result.

### Finding 5 — test result and closure note (2026-04-29)

User-confirmed test outcomes:

1. `ge` Tab `5` Tab Go → caret lands in the document at the correct position. **Fix (A) resolves the primary symptom.**
2. `ge` Tab `5`, wait 5 s, Tab → Tab character types into the document body. **Idle-commit focus leak** (terminology corrected from earlier "Tab race" — see 2026-04-29 correction below). Word's customUI `editBox` auto-commits the value after a short idle interval (~1 s) and returns focus to the document body; subsequent Tabs are then document Tabs. This is documented Word behaviour, not a project defect.
3. Next-button → Exodus, Tab `5` Tab `3` Tab Go → search works. Confirms the chained nav path is sound when no idle period elapses between Tabs (commits arrive before the idle-commit threshold expires).

**Constraints documented for the residual (separate item, not fixed by (A)):**

Word's customUI14 schema does **not** expose a public API to programmatically set focus on a specific ribbon control. `IRibbonUI` has `Invalidate`, `InvalidateControl`, `ActivateTab`, `ActivateTabMso` — and nothing for `SetFocus(controlId)`. Tab routing within the ribbon is platform-controlled; on `editBox` commit (whether explicit via Enter/Tab or implicit via idle) Word returns focus to the document by design. The behaviour is therefore a Word limitation, not a project defect.

Available paths if/when the Tab race is prioritised:
- **KeyTips (already wired):** `KT_BOOK / KT_CHAPTER / KT_VERSE / KT_GO` constants exist in `basUIStrings`. `Alt + <keytip>` is the canonical Office UX for cross-control jumps and bypasses Tab entirely. Documentation update only — no code change.
- **Auto-fire Go on valid chapter+verse:** condition the existing `OnChapterChanged` / `OnVerseChanged` handlers to invoke `OnGoClick` once both fields validate. Removes the need for the final Tab → Go step. Tradeoff: nav fires before user expects it.
- **VSTO / WPF ribbon rewrite:** would allow a true Tab focus chain. Major rewrite; deferred indefinitely.

**Status:** **OPEN** — fix (A) resolved primary symptom (caret-not-visible). Idle-commit focus leak (formerly "Tab race") tracked as a separate item; Finding 5 stays open as the umbrella ticket until the residual is addressed or explicitly closed as won't-fix.

### Finding 5 — terminology correction (2026-04-29)

User correctly pushed back on the term "Tab race". A race condition by definition resolves with timing — wait long enough and you get a consistent winner. The user's observation is the opposite: **after waiting 5 seconds, the Tab still goes to the document, deterministically**. That isn't a race; it's a steady-state focus-leak.

**Corrected diagnosis: idle-commit focus leak.**

Word's customUI `editBox` has two focus-loss triggers, both deterministic:

1. **Idle-commit:** after the user stops typing for ~1 s, Word treats the value as committed, fires `onChange`, and returns focus to the document body. Undocumented in the customUI XML reference but observable across Office versions.
2. **Explicit-commit:** Enter, Tab, or focus-loss commits immediately. Same result.

So when the user types `5` and pauses 5 s, focus has already been handed back to the document during the first ~1 s of idle. The next Tab is no longer "Tab inside the chapter editBox" — it's "Tab inside the document". A Tab in the document body inserts a tab character. Deterministic.

That also explains scenario (3) — rapid Tab chains arrive **before** the idle-commit threshold expires, so they're processed by the ribbon's internal tab routing. The moment the user pauses for ~1 s, that grace period ends and every subsequent Tab is a document Tab.

**Hidden bug or by-design Word behaviour?**

It is **Word behaving as designed**, not a defect in this project's code. Microsoft's official UX position is: ribbon-control navigation is **KeyTips**, not Tab. The Tab traversal in (3) is accidental — Microsoft never promised it. Disambiguation:

| Symptom | Hidden bug? |
|---|---|
| Caret doesn't appear after first Go | **Hidden bug — fixed by (A).** Code didn't return focus to the document. |
| Tab after chapter sometimes goes to next field, sometimes to document | **Not a code bug — Word's documented customUI behaviour.** Idle-commit hands focus to document. |
| Rapid Tab traversal works (3) | **Not officially supported.** Works by accident; could break across Office updates. |
| `5` doesn't commit until I press something | **Not a bug — by design.** customUI editBox waits for explicit commit or idle threshold. |

**Apology for the sloppy term.** "Tab race" in earlier notes implied a code-level timing defect we could resolve. The correct framing is platform-level deterministic focus-leak. Forward references to "Tab race" in this Finding should be read as "idle-commit focus leak".

**Recommendation unchanged.** No VBA-side fix is available. The supported path is **KeyTips** (already wired in `basUIStrings`). Auto-fire-on-valid-chapter+verse remains a viable code-side option to *avoid* the residual rather than fix it. VSTO/WPF rewrite remains the only path to true ribbon-owned focus management.

---

## 2026-04-29 — Style taxonomy snapshot refresh

### Run output

`WordEditingConfig` + `DumpAllApprovedStyles` + `RUN_TAXONOMY_STYLES`:

- `DumpAllApprovedStyles`: **44 succeeded, 0 failed, ~4.24 sec.**
- `RUN_TAXONOMY_STYLES`: **13 PASS  6 FAIL → rpt\StyleTaxonomyAudit.txt**.
- `PromoteApprovedStyles` reported five styles missing from the document — `BodyTextContinuation`, `BookIntro`, `AppendixTitle`, `AppendixBody`, and the deliberate canary `FargleBlargle` — all expected per the EDSG `01-styles.md` "Missing from document" list.

`SpeakerLabel` is now present at priority 37 (added by the recent `Add SpeakerLabel style` commit), pushing `BodyTextIndent` from 37 → 38 and the Emphasis / Author / Normal block from the 42–47 range to 43–48. The reserved-gap of four slots shifted from 38–41 to 39–42.

### Taxonomy routine header — corrected (2026-04-29)

`src/basTEST_aeBibleConfig.bas:222-241` had `Audits all 17 approved taxonomy styles ...`. The actual count of `AuditOneStyle` calls is **19** (2 fully-specified + 14 existence-verified + 3 not-yet-created), confirmed by the run's 13 PASS + 6 FAIL = 19 total.

Fix applied: PURPOSE block updated to state 19 styles and enumerate the three buckets:

```
PURPOSE:
  Audits the 19 approved-array taxonomy styles against their expected
  configuration and writes a structured report to rpt\StyleTaxonomyAudit.txt.
  Buckets:
    2 fully specified (all 7 properties verified) - BodyText, BodyTextIndent
    14 existence-verified (full spec pending)
    3 not yet created (expected FAIL until each Define* routine runs)
```

No code change to the audits themselves; header text only.

### EDSG `01-styles.md` — refreshed (2026-04-29)

Three small edits to capture the post-`SpeakerLabel` state:

1. "Latest run" stamp: `2026-04-26 / 43 approved styles` → `2026-04-29 / 44 approved styles`.
2. "Pending re-validation (priorities 37+)" table updated: `SpeakerLabel` inserted at 37; `BodyTextIndent` 37→38; `EmphasisBlack` 42→43; `EmphasisRed` 43→44; `Words of Jesus` 44→45; `AuthorSectionHead` 45→46; `AuthorQuote` 46→47; `Normal` 47→48.
3. "Reserved gaps" paragraph updated: `Priorities 38–41 are reserved` → `Priorities 39–42 are reserved`, with a note that the gap shifted +1 on 2026-04-29 when `SpeakerLabel` was added at priority 37.

Validated 1–36 table unchanged — `SpeakerLabel` correctly belongs in the pending-re-validation bucket per the EDSG's own convention until the page walk reaches it.

### Follow-up (not actioned)

`RUN_TAXONOMY_STYLES` does not currently include an `AuditOneStyle "SpeakerLabel", ...` entry, so the new style is **promoted but not actively audited**. By user decision (default), this is left as-is. Adding it would be a one-line addition to the existence-verified bucket: `AuditOneStyle "SpeakerLabel", "", 0, -1, -999, -1, -999, -999, -999`, bumping the total to 20.

**Status:** all three sub-items APPLIED — 2026-04-29.

> **Important — taxonomy audit final-state goal**
>
> The current 19-entry curated audit is a *transitional* state, not the
> destination. The final-state resolution is for `RUN_TAXONOMY_STYLES`
> to map **every approved style** with a real (non-sentinel) expected
> spec, so that any property drift on any approved style is caught
> immediately. Promoted-but-unaudited styles like `SpeakerLabel` today
> are temporarily off-radar for property drift — `DumpAllApprovedStyles`
> only confirms existence and priority, not font / size / alignment /
> indents / line-spacing / spacing-before-after.
>
> **This should have a prominent place in the EDSG** so the goal is
> visible to anyone reading the style taxonomy page, and so progress
> can be tracked: each move from bucket 2 (existence-verified, full
> spec pending) into bucket 1 (fully specified) is a measurable step
> toward full drift coverage.
>
> Action: add a top-level callout in `EDSG/01-styles.md` and link it
> from `EDSG/04-qa-workflow.md` so both QA-first and styles-first
> readers see the goal.

---

## 2026-04-29 — Spec promotion: descriptive vs prescriptive (decision)

User asked to promote eight styles from bucket 2 (existence-verified) to bucket 1 (fully specified) in `RUN_TAXONOMY_STYLES`, plus remove `Lamentation` and add `Footnote Reference`. Before encoding values, the question of *what* to encode arose.

### Two ways to set audit "expected" values

**Descriptive spec — capture what's there.** Read the live values from `rpt/Styles/style_*.txt` and encode them as expected. The audit then acts as a *drift detector*: alerts only when the style changes from its current state. Pro: passes immediately, captures today's state, no false negatives. Con: blesses any current-but-wrong values silently — if today's `Brief` style has `LineSpacingRule = Exactly` while the EDSG QA checklist says it should be Single, a descriptive audit codifies the wrong value as canonical.

**Prescriptive spec — capture what it should be.** Encode values that represent the design intent regardless of current state. The audit then acts as a *correction driver*: fails until the document agrees with the spec. Pro: turns the audit into a tracked to-do list closing one item at a time. Con: starts with FAILs as the normal state, "PASS count" is no longer a meaningful health metric until the corrections land.

### Concrete divergences observed in the eight dumped styles

The QA-checklist properties (`BaseStyle`, `AutomaticallyUpdate`, `QuickStyle`) aren't checked by `AuditOneStyle` — they're tracked separately. But several styles have `LineSpacingRule = 4 (Exactly)` against the QA-checklist preference of `0 (Single)` for paragraph styles, and `LineSpacingRule` *is* an `AuditOneStyle` arg. Notable cases: `Heading 2`, `CustomParaAfterH1`, `Brief`, `Psalms BOOK`, `Footnote Text`. Each of these would be a candidate for a prescriptive override (encode `0` instead of `4` and treat the resulting FAIL as a tracked correction item).

### Decision — path (a), purely descriptive (2026-04-29)

Path (a) chosen as the safer first move. Reasoning:

- The audit gets a known baseline immediately. PASS count is meaningful from day one as drift detection.
- Path (b)'s prescriptive overrides are a separate, deliberate exercise — one property at a time, each tracked as its own review item with rationale recorded. Folding them into the bucket-1 promotion would conflate two different decisions.
- Locking in today's values doesn't preclude prescriptive correction later — it just sets the *floor*. When ready to drive a correction (e.g. "all paragraph styles → `LineSpacingRule = Single` per QA checklist"), the audit's expected value is updated and the FAIL becomes the tracking signal.

Path (b) remains a viable next step. Tracked as a deferred work item: any style where the descriptive value is known to violate the EDSG QA checklist is a future prescriptive-override candidate.

### Footnote Reference — deferred to bucket 2 (Q2 decision)

`AuditOneStyle` only inspects font + paragraph-format properties. On a Character style (which Footnote Reference is), only Font.Name and Font.Size apply; Bold, Italic, Color are not audited. Promoting to bucket 1 today would be misleading — the "fully specified" label would mean only "every audit-able property" is set, which is just two properties out of the styled-character surface.

Decision: leave `Footnote Reference` in bucket 2 with sentinels for now, and **track as a follow-up** the work to extend `AuditOneStyle` to check Bold / Italic / Color so character-styles can be truly fully specified. Once that audit extension lands, `Footnote Reference` (and any other character style) can graduate into bucket 1 honestly.

### Source-code application — applied 2026-04-29

Following edits applied as a single batch under the descriptive-spec decision:

**`src/basTEST_aeBibleConfig.bas`:**
- PURPOSE comment block updated: bucket counts `2 + 14 + 3` → `9 + 7 + 3`; added explicit listing of bucket-1 members; cross-reference to this review's decision section.
- 7 `AuditOneStyle` calls promoted from bucket 2 to bucket 1 with descriptive specs read from `rpt/Styles/style_*.txt`:
  - `Heading 1`        — `"Noto Sans", 24, 1, 0,    0, 12, 144,  0`
  - `Heading 2`        — `"Noto Sans",  8, 1, 0,    4, 10,  12,  8`
  - `CustomParaAfterH1`— `"Noto Sans", 10, 1, 0,    4, 10,   0, 62`
  - `DatAuthRef`       — `"Noto Sans",  8, 1, 0,    0, 12,  11,  0`
  - `Brief`            — `"Noto Sans", 10, 1, 0,    4,  9.5, 0,  0`
  - `Psalms BOOK`      — `"Carlito",    9, 0, 14.4, 4, 10,  10,  0`
  - `Footnote Text`    — `"Carlito",    7, 3, 0,    4,  8,   0,  0`
- `Lamentation` line deleted (style was removed from approved array on 2026-04-26 per `EDSG/01-styles.md`).
- `Footnote Reference` added as a new bucket-2 entry with font + size descriptive (`"Carlito", 9`) and remaining args sentinel — parked here pending the deferred follow-up.

**`EDSG/01-styles.md`:**
- "Important — taxonomy audit final-state goal" callout updated: bucket counts `2 + 14 + 3` → `9 + 7 + 3`; "Progress so far" sub-list added with the 2026-04-29 promotions, the Lamentation removal, the Footnote Reference parking, and the descriptive-spec decision cross-reference.

**Total audit count unchanged at 19** (Lamentation out, Footnote Reference in cancels). Bucket distribution: bucket 1 grew from 2 to 9, bucket 2 shrank from 14 to 7, bucket 3 unchanged at 3.

### Deferred follow-up tasks (next refinement level)

1. **Extend `AuditOneStyle` to check character-style properties** (Bold, Italic, Color). Once landed, `Footnote Reference` (and any future character style) can graduate from bucket 2 to bucket 1 honestly. Captured in this review as the unblocker for the rest of the character-style audit coverage.
2. **Prescriptive-spec pass.** A separate, deliberate exercise: identify each style where the descriptive value violates the EDSG QA checklist (notable candidates today: `Heading 2 / CustomParaAfterH1 / Brief / Psalms BOOK / Footnote Text` all have `LineSpacingRule = 4 (Exactly)` against the checklist preference of `0 (Single)`; `BookIntro` is missing from the document and stays a bucket-2 placeholder). Each prescriptive override is a tracked review item with rationale recorded.

**Status:** APPLIED — 2026-04-29.

### Post-application audit run (2026-04-29)

`RUN_TAXONOMY_STYLES`: **14 PASS  5 FAIL → rpt\StyleTaxonomyAudit.txt** (was 13 PASS  6 FAIL prior to this batch — net −1 FAIL, +1 PASS, exactly as predicted).

All seven promoted styles PASS with their new descriptive specs. `Footnote Reference` (newly added to bucket 2) also PASS. `Lamentation` removed (was a pre-existing FAIL).

#### FAIL breakdown

| # | Style | Reason | Note |
|---|---|---|---|
| 1 | `BookIntro` | NOT FOUND in document | Pre-existing missing-tracker entry; style listed in approved array but not defined |
| 2 | `ListItem` | `Indent: expected 0 got -18` | **Pre-existing spec drift** — not introduced by today's changes |
| 3 | `BodyTextContinuation` | NOT FOUND | Bucket 3 — expected FAIL until Define routine runs |
| 4 | `AppendixTitle` | NOT FOUND | Bucket 3 — expected FAIL |
| 5 | `AppendixBody` | NOT FOUND | Bucket 3 — expected FAIL |

#### `ListItem` indent — leave-as-is (decision 2026-04-29)

Live document has `FirstLineIndent = -18` (hanging indent ~0.25 inch); audit's existing partial spec expects `0`. Two-way descriptive/prescriptive call left open: encoding `-18` would make this PASS by capturing current state; encoding `0` (current) keeps the FAIL as a work item if the live value is considered wrong. **Decision: leave as-is.** The FAIL is already serving as a tracked indicator, and a future task (separately defined) impacts this resolution path. Re-evaluate when that task lands.

This kind of partial-spec FAIL is exactly the case the deferred *prescriptive-spec pass* (recorded above) is designed to handle — each item gets a deliberate descriptive-vs-prescriptive call with rationale. `ListItem` is now an explicit member of that follow-up's input set.

**Status:** baseline established at 14/5; deferred items unchanged.

---

## 2026-04-29 — `List Paragraph` numbering-engine bug — analysis and migration recipe

### Honest caveats up front

This analysis was prepared without independent reproduction at the C++ source level. Inputs: user's direct experience plus publicly documented evidence (Microsoft Q&A threads, MVP blogs, Office Open XML spec). Assessment is "consistent with known evidence" rather than "verified at the engine level"; gaps are flagged inline.

### 1. Verifying the proposal

User's proposal: **"create your own style and do not use the `List Paragraph` / `List Item` inheritance."**

**Pros**

- **Avoids the cascade trigger.** The hang is consistent with Word's list engine eagerly recomputing all paragraphs that reference a `ListTemplate` when an inherited or linked style changes. Detaching from `List Paragraph` removes the inheritance edge that triggers the recompute on every style edit.
- **Aligns with the EDSG QA checklist.** `BaseStyle = ""` is already the project standard. `List Paragraph` inheritance violates it — so the refactor enforces an existing rule, not a new one.
- **Predictable across Word versions.** `List Paragraph` has had quiet behavioural changes across Word 2010 / 2013 / 2016 / 365. Standalone styles are version-stable.
- **VBA-only maintenance becomes safe.** With no list-template link, `s.Font.Size = 11` doesn't trigger a recompute storm.
- **Importable and reproducible.** Standalone styles round-trip through `.docx` cleanly; `List Paragraph`-derived styles often don't, because recipient documents inherit a different `List Paragraph` definition.

**Cons**

- **Loses Word's auto-numbering for those styles.** No `1. 2. 3.` or `a) b) c)` from the list engine. If numbering is required, render markers in the body or apply a separate dedicated VBA-driven list pattern.
- **Weakens Outline / Navigation pane integration.** Word's outline tooling reads from list-template structures; detached list-styled paragraphs don't show numbering in the navigation pane.
- **Breaks cross-references to numbered items.** Anything that says "see item 4(b)" depends on the list engine resolving the number.
- **Migration step itself carries hang risk.** The very work to escape the bug must be done carefully — VBA, batched, `ScreenUpdating = False`, never the Modify Style dialog.
- **A future Word feature may assume the list engine** (Tables of figures, captions, certain accessibility tools).

**Net read for this project.** A Bible-class document — long-form running text, verse markers, list usage limited to specific structural blocks (`ListItem`, `ListItemBody`, `ListItemTab`) — doesn't actually use any of the lost features. List items aren't cross-referenced by number. Outline view is anchored on Heading 1/2. The cons mostly don't apply.

**Recommendation: the proposal is sound for this project.** Approved 2026-04-29.

### 2. Real solution — five-step migration recipe

All-VBA path, never the Modify Style dialog:

- **Step 0 — Diagnostic.** `AuditListStyleRisk` enumerates approved paragraph styles flagging non-empty `BaseStyle` and any `LinkToListTemplate`. Expected to flag `ListItem`, `ListItemBody`, `ListItemTab`.
- **Step 1 — Define replacements in a blank `.docx`.** Critical isolation: a fresh empty Word document has no list-engine state to recompute against, so the cascade doesn't fire. Set `BaseStyle = ""` *before* any other property.
- **Step 2 — Transport to live document.** Either `Document.CopyStylesFromTemplate` (Organizer) or VBA direct-property-copy. Never the dialog.
- **Step 3 — Re-apply to existing paragraphs.** Batched, `ScreenUpdating = False`. With no list-template link on the new style, the list-engine recompute that causes the hang doesn't fire.
- **Step 4 — Decommission the old styles.** Remove from approved array; delete (or `Priority = 99`) after a clean audit pass.
- **Step 5 — Update `RUN_TAXONOMY_STYLES`.** Encode descriptive specs for the new style names.

Full code recipe lives in [`EDSG/10-list-paragraph-bug.md`](../EDSG/10-list-paragraph-bug.md). Created 2026-04-29.

### 3. EDSG documentation — applied 2026-04-29

- New page **`EDSG/10-list-paragraph-bug.md`** created. Contains: symptom, project policy, root cause read, Microsoft status, common bad advice (with rebuttals), the five-step migration recipe with code, what we lose / don't lose by detaching, cost-now vs cost-later, i18n implications, history.
- Cross-link added to **`EDSG/01-styles.md`** QA-checklist row `BaseStyle = ""` → "[why](10-list-paragraph-bug.md)" — the *why* is now one click from the rule.
- Cross-link added to **`EDSG/02-editing-process.md`** Stage 1 step 2 — explicitly forbids `List Paragraph` inheritance and `LinkToListTemplate` for any list-shaped style, with link to the recipe.
- Cross-link added to **`EDSG/04-qa-workflow.md`** below the existing taxonomy callout — second `⚑` callout flagging the bug and pointing at the new page.

Three independent reading paths (styles taxonomy, editing process, QA workflow) all surface the rule and the recipe. A reader entering at any of those pages sees the bug documented.

### 4. Cost-now vs cost-later — measurement framework

**Cost-now (one-off refactor):**

| Component | Estimate | Notes |
|---|---|---|
| Diagnostic (Step 0) | 15 min | Run, log, confirm 3 candidate styles |
| Define replacements in blank doc (Step 1) | 30-60 min | 3 styles × ~15-20 min; specs already exist in `style_*.txt` dumps |
| Transport to live doc (Step 2) | 15 min | One-shot VBA |
| Migrate paragraphs (Step 3) | 5-30 min runtime + verification | Depends on affected paragraph count |
| Decommission old (Step 4) | 15 min | Update `approved` array |
| Update taxonomy audit (Step 5) | 15 min | Three lines edited |
| Visual / audit verification | 30-60 min | Re-run `RUN_TAXONOMY_STYLES`, full doc visual scan |
| **Total** | **2-4 hours** | One sitting, one branch, one PR. Bounded. |

**Cost-later (deferred, recurring):**

| Component | Cost type | Probability per modification |
|---|---|---|
| Style modification hangs Word | Lost work + restart | Moderate-to-high in this 33,857-paragraph doc |
| Workarounds around the hang accumulate | Tech debt | Compounds linearly |
| New contributors hit the hang | Onboarding cost | High; unavoidable without docs |
| Refactor eventually mandatory anyway | Future task | Probability = 1 |
| i18n exposure | Multiplier on hang risk | Higher for longer translated docs |

**Single-number indicator.** Time the next Modify Style operation on a `List Paragraph`-derived style via the dialog. If it exceeds 10 seconds or shows "Not Responding", refactor is overdue. The canary has already chirped on this project.

**Math.** Cost-now is bounded at ~4 hours. Cost-later is unbounded and probability-1 of recurring. Net: do it now.

### 5. i18n implications

Two opposite forces, with a clear net direction.

**Forces pushing toward "refactor first, then i18n":**

- Translated documents are typically *longer* (English → German averages +30%, English → French +15-25%). Bigger documents hit the list-engine hang harder.
- Word's `NameLocal` aliasing for `List Paragraph` differs by Office UI language. A document built around inheritance can break on a Word installation with a different UI language. Standalone styles use literal names and round-trip cleanly.
- The list-rendering layer is the part most affected by RTL languages (Arabic, Hebrew). Custom RTL handling is cleaner in our own code than in Word's RTL list engine.

**Forces against (or neutral):**

- Word's list engine handles list-marker localisation natively (Roman numerals, alphabetic letters, locale-appropriate digits). Detaching means the project localises markers itself. For this document — verses are marked, lists are structural and finite — the cost is small.
- Number-format-as-content (where the marker becomes part of the body text) means markers need translation. Mostly neutral here.

**Net.** Refactoring first makes i18n cheaper and safer. Order: list-style refactor → then i18n rollout. If i18n ships first, every translated doc carries the same hang and the eventual refactor becomes a cross-locale change touching every translated artifact.

### Status of the actual refactor work — DEFERRED

User has approved the analysis and the EDSG documentation, but the refactor itself is **deferred** — to be run as a separate batch under the standard one-fix-at-a-time review process when scheduled. Estimated 2-4 hours focused work. Trigger: discretionary (do before the next round of large-doc style edits, before i18n rollout starts, or before the next "Word not responding" incident — whichever comes first).

**Status:** analysis recorded; EDSG documentation applied; refactor deferred.

### Correction — holding file extension `.docx` → `.docm` (2026-04-29)

Caught by user during review of the EDSG page: the migration recipe described the holding file as `.docx`, but a `.docx` cannot retain VBA — Word strips macros on save in macro-free formats. The holding file is meant to carry its own style-creation macro (so future contributors can clone the repo, open the file, run the macro, and reproduce a known state), so it must be `.docm` (macro-enabled document).

**File-extension reference table** (Office 2007+ format split — security boundary at the extension level so tooling can identify executable content without inspecting the zip):

| Extension | Kind | Macros |
|---|---|---|
| `.docx` | Document | Stripped on save |
| `.docm` | Document | Retained |
| `.dotx` | Template | Stripped on save |
| `.dotm` | Template | Retained |

**Design choice — option (A) approved:** holding file as `.docm` only, single artifact, accessed via Step 2(b) (`Documents.Open` + property copy). `.dotm` (macro-enabled template) reserved as a future option for shipping styles to documents that aren't macro-enabled.

**Fix applied to `EDSG/10-list-paragraph-bug.md`:**

- Step 1 heading: `.docx` → `.docm`.
- Step 1 holding-file save line: `tools/style_holding.docx` → `tools/style_holding.docm`.
- Step 2 (b) `Documents.Open` line: `style_holding.docx` → `style_holding.docm`.
- Added a `> File extension matters.` callout in Step 1 explaining the macro-free vs macro-enabled distinction so future readers don't repeat the mistake.
- Step 2 (a) Organizer line untouched — already correctly used `.dotm` (template).

No source-code change — the recipe is the only mention of the holding file in the project today.

### Correction — example style name `ListItem_v2` → `AuthorListItem` (2026-04-29)

User directive: rename the example replacement-style name in the migration recipe from the placeholder `ListItem_v2` to a project-meaningful `AuthorListItem`. The new name signals the actual intended use (author/list-item content) rather than a generic versioned placeholder.

Applied to `EDSG/10-list-paragraph-bug.md` — all occurrences of `ListItem_v2` replaced with `AuthorListItem` (single file, replace-all). Affects Step 1 (style definition), Step 2 (b) (transport), and any inline references in the recipe text.

No source-code change — `AuthorListItem` is not yet in the `approved` array; it becomes a real style only when the deferred refactor work is scheduled.

### List Paragraph migration — kickoff (2026-04-29)

User directive: **start the process** for creating `AuthorListItem` and the other replacement styles. The previously-deferred refactor moves from "deferred" to "in progress." Run as a multi-phase sequence under the standard one-fix-at-a-time review.

#### Multi-phase plan

| Phase | What | Who acts | What lands in repo |
|---|---|---|---|
| 0 | Diagnostic — run `AuditListStyleRisk`, collect output | User runs, I read | New diagnostic Sub in `src/basAuthorStyles.bas` |
| 1 | Define replacements in `tools/style_holding.docm` | I write `CreateAuthorStyles` Sub; user creates the `.docm` and runs it | Same `basAuthorStyles.bas` adds `CreateAuthorStyles`; new local `tools/style_holding.docm` (binary, not version-controlled) |
| 2 | Transport styles into live `.docm` | User runs transport routine | Same module adds `TransportAuthorStyles` |
| 3 | Migrate paragraphs from old → new style | User runs `MigrateParagraphs` per pair | Reusable Sub in `basAuthorStyles.bas` |
| 4 | Decommission old styles + update audit | I edit | `src/basTEST_aeBibleConfig.bas` `approved` array; `RUN_TAXONOMY_STYLES` audit list |
| 5 | Verify | User runs `RUN_TAXONOMY_STYLES` and confirms clean | (no further code edits unless audit reports unexpected) |

#### Decisions recorded — 2026-04-29

- **Naming — Option A (parallel rename), approved.** `ListItem → AuthorListItem`, `ListItemBody → AuthorListItemBody`, `ListItemTab → AuthorListItemTab`. Mechanical 1-for-1 mapping; preserves the existing three-tier structure. Consolidation, if warranted, is a separate later decision.
- **Module placement — option (i), approved.** New self-contained `src/basAuthorStyles.bas`. Easy to remove once the migration is verified and `RUN_TAXONOMY_STYLES` is clean. Doesn't pollute `basFixDocxRoutines.bas` or `basTEST_aeBibleConfig.bas`.
- **Phase 0 (diagnostic) — APPLIED 2026-04-29.** `src/basAuthorStyles.bas` created with `AuditListStyleRisk`. Output expected (prediction): three paragraph styles flagged — `ListItem`, `ListItemBody`, `ListItemTab` — confirming the migration scope. Awaiting user-side run.

#### Phase 1 placeholder

`CreateAuthorStyles` will be designed against the descriptive specs already captured in `rpt/Styles/style_*.txt` (specifically: `style_ListItem.txt`, `style_ListItemBody.txt`, `style_ListItemTab.txt`). That guarantees the new styles match the current visual rendering 1-for-1; any design corrections (the prescriptive-spec exercise) remain a separate follow-up. The `BaseStyle = ""` rule from the EDSG QA checklist applies — no inheritance from `List Paragraph` and no `LinkToListTemplate` on any of the three new styles.

#### Phase boundaries

- Each phase concludes with explicit user-side action (run a Sub, observe output, paste back to me).
- I propose; user approves; user runs; I read; we proceed.
- No phase merges into the next without intervening review.

**Status:** Phase 0 applied; awaiting `AuditListStyleRisk` run output before Phase 1.

### Phase 0 — first run output and corrections (2026-04-29)

#### `LinkToListTemplate` is a method, not a property — initial diagnostic failed compile

User's first run of `AuditListStyleRisk` raised "argument not optional" on the line:

```vba
hasLT = Not (s.LinkToListTemplate Is Nothing)
```

`Style.LinkToListTemplate` is a method (it takes a `ListTemplate` argument and is used to *establish* the link), not a read-only property. There is no public Word VBA equivalent for "is this style linked to a list template" that can be queried directly. **Fix: drop the `LinkToListTemplate` check** and rely on `BaseStyle` inheritance as the primary signal. The list-template-attachment-without-inheritance case is rare; a paragraph-level fallback (`Paragraphs.Range.ListFormat.ListTemplate Is Nothing` sampling) can be layered on later if Phase 0 output looks incomplete.

#### First successful run flagged zero styles — diagnostic was too narrow

After the fix, `AuditListStyleRisk` ran clean and reported `Flagged: 0 style(s).` Unexpected — predicted `ListItem`, `ListItemBody`, `ListItemTab` should have been flagged. Inspection of the existing dump files showed why:

| Style | `BaseStyle` (live) |
|---|---|
| `ListItem` | `"List Number"` |
| `ListItemBody` | `""` (already detached) |
| `ListItemTab` | `""` (already detached) |

Two corrections to the original assumptions:

1. **Migration scope is smaller than predicted.** Only `ListItem` actually has list-family inheritance to refactor. `ListItemBody` and `ListItemTab` are already structurally clean (`BaseStyle = ""`). If past hangs were on edits to `ListItem` (or on the list as a whole because `ListItem` is the root), fixing `ListItem` may be sufficient. If hangs occurred while editing `ListItemBody` / `ListItemTab` directly, the cause is something else (likely direct list-format attachment on paragraphs rather than style inheritance).
2. **Diagnostic was too narrow.** It checked only `"List Paragraph"` literally; `ListItem` inherits from `"List Number"`. Word has a whole family of list-related built-in paragraph styles — `List Paragraph`, `List`, `List Number`, `List Bullet`, `List Continue`, plus their numbered siblings — and any of these as a `BaseStyle` is the same engine-recompute risk vector.

#### Phase 0 v2 — broadened diagnostic (APPLIED 2026-04-29)

Rewrote `AuditListStyleRisk` with three changes:

(a) **Inheritance check broadened** to catch any list-family built-in: `LCase$(BaseStyle)` matches `"list paragraph"`, `"list"`, or wildcard-prefix matches `"list number*"` / `"list bullet*"` / `"list continue*"`. Catches `ListItem` (BaseStyle=`"List Number"`) and any other inheritance from list-family parents we haven't anticipated.

(b) **Two-output design** — section (A) flagged at-risk styles, section (B) full inventory of every paragraph style with non-empty `BaseStyle`. Section (B) gives one-time discovery visibility into anything else of interest (e.g., approved styles violating the `BaseStyle = ""` rule for non-list reasons).

(c) **Honest output naming** — "Inherits from list-family built-in" rather than just "List Paragraph", matching what's actually being checked.

Awaiting v2 run output. Section (A) should now show `ListItem` flagged (per the dump data); section (B) gives the full BaseStyle picture for the next round of decisions.

**Status:** Phase 0 v2 applied; awaiting v2 run.

### Phase 0 v2 — run output and scope decisions (2026-04-29)

#### Section (A) — flagged at-risk styles

Two styles flagged with list-family inheritance:

| Style | Priority | BaseStyle |
|---|---|---|
| `ListItem` | 18 | `List Number` |
| `AuthorBookRef` | 22 | `List Number` |

**Migration scope doubled vs original prediction.** `AuthorBookRef` was not anticipated — it inherits from `List Number` and carries the same Modify-Style hang risk as `ListItem`. The project has been operating with two list-engine-entangled approved styles, not one.

`ListItemBody` (19) and `ListItemTab` (20) — predicted candidates — drop out of the migration entirely. Their dump files confirm `BaseStyle = ""`; already structurally clean for the list-engine concern. Earlier hang reports may have surfaced via `ListItem` (the inheritance root) cascading through the list engine, even if the user perceived the edit was on a different style.

#### Section (B) — full inventory findings (108 paragraph styles with non-empty BaseStyle)

Most entries are Word built-ins inheriting from `Normal` (expected, not actionable). Among **approved** styles (priority < 99), six violate the EDSG QA-checklist rule `BaseStyle = ""` for non-list-family reasons (i.e., they inherit from `Normal`, not from a list-family parent):

| Style | Priority | BaseStyle |
|---|---|---|
| `CustomParaAfterH1` | 25 | `Normal` |
| `Brief` | 26 | `Normal` |
| `Footnote Text` | 32 | `Normal` |
| `Psalms BOOK` | 33 | `Normal` |
| `PsalmSuperscription` | 34 | `Normal` |
| `PsalmAcrostic` | 36 | `Normal` |

These do **not** trigger the list-engine hang (Normal has no list-engine entanglement), but they violate the broader QA-checklist rationale (predictability, version-stability, reproducibility). Lower urgency than the list-family migration.

#### Three decisions recorded — 2026-04-29

**(1) Migration scope — APPROVED.** Phase 1 targets two styles only: `ListItem` and `AuthorBookRef`. `ListItemBody` and `ListItemTab` removed from migration scope (already `BaseStyle = ""`).

**(2) `AuthorBookRef` naming approach — APPROVED, suffix-temp pattern.** Create replacement as `AuthorBookRefNew` in the holding `.docm`, transport, migrate paragraphs, delete the old `AuthorBookRef`, then VBA-rename `AuthorBookRefNew → AuthorBookRef`. Final state identical to today's name. Same-name overwrite (option 3 from the proposal) was rejected as it would trigger the very hang the migration is escaping.

**(3) Section (B) violations — APPROVED, not part of Phase 1.** The six `BaseStyle = "Normal"` approved styles are deferred to a separate pass, folded into the existing deferred prescriptive-spec exercise. Treating list-engine migration and BaseStyle-compliance cleanup as separate work items keeps each pass scoped and reviewable.

#### Updated migration target table

| Old style | Inheritance issue | Replacement (final name) | During-migration name |
|---|---|---|---|
| `ListItem` | BaseStyle = `List Number` | `AuthorListItem` | `AuthorListItem` (no temp; new name) |
| `AuthorBookRef` | BaseStyle = `List Number` | `AuthorBookRef` (unchanged) | `AuthorBookRefNew` (temp; renamed at end) |

#### What's next — Phase 1 design

`CreateAuthorStyles` Sub in `src/basAuthorStyles.bas` will define both replacement styles in a fresh blank `.docm` holding file. Specs will be **descriptive** (read from `rpt/Styles/style_ListItem.txt` and `rpt/Styles/style_AuthorBookRef.txt` if present, or fresh dumps if not). `BaseStyle = ""` set first; no `LinkToListTemplate` call. Phase 1 will be proposed as code-only edit before being applied.

**Status:** Phase 0 v2 run analysed; scope and naming decisions recorded; Phase 1 design pending.

---

## 2026-04-28 — Versification reconciliation: data follows WEB / English Protestant

### Decision

The project source text is based on the **World English Bible Protestant Edition** (WEB), per [eBible.org](https://ebible.org/web/):

> "This is the World English Bible Protestant Edition... It contains only the 66 books of the Old and New Testaments."

WEB versification matches the standard English Protestant tradition (KJV / ASV / NASB / ESV / NIV). All verse-count source data must therefore reflect English Protestant numbering, not Masoretic / Hebrew numbering. The project intent was already documented in `src/basSBL_VerseCountsGenerator.bas:14-21` ("66-book Protestant canon used by KJV, NIV ... 31,102 verses (KJV)"); some chapter pairs had silently drifted to Hebrew versification.

### Audit-found discrepancies — diagnosis

After applying Finding 1 (per-book accumulator reset), the audit reported nine chapter mismatches. Reclassified by root cause:

**(A) Source-data bugs (verse-count table is wrong; document is correct).** Hebrew/Masoretic numbering inadvertently used:

| Book / chapter | Current data | Should be (WEB) | Note |
|---|---|---|---|
| 2 Samuel 18 | 32 | 33 | Hebrew 19:1 ("O my son Absalom...") = English 18:33 |
| 2 Samuel 19 | 44 | 43 | Compensating split |
| 2 Kings 11 | 20 | 21 | Hebrew 12:1 = English 11:21 |
| 2 Kings 12 | 22 | 21 | Compensating split |
| 2 Chronicles 13 | 23 | 22 | Hebrew 14:1 = English 13:23 |
| 2 Chronicles 14 | 14 | 15 | Compensating split |
| 3 John | 15 | 14 | Some Greek editions split v.14 into 14+15; WEB does not |

**(B) Document content bugs (verse-count table is correct; document is wrong).** No data fix needed; flagged for separate document repair:

| Book / chapter | Expected (WEB) | Found in doc | Likely cause |
|---|---|---|---|
| Romans 14 | 23 | 26 | Document apparently includes a 14:24-26 doxology that WEB places at 16:25-27. |
| Hebrews 7 | 28 | 27 | Document is missing one verse marker. |

**(C) Pre-existing document issue, unrelated to versification.** Flagged earlier:

| Book / chapter | Note |
|---|---|
| Joshua 23 = 49 found vs 16 expected; book short one chapter | Likely a missing Heading 2 break causing chapter 23/24 to merge. |
| H1 "SONG OF SONGS" unrecognised | Audit module's local `PopulateCanonical` still has `"Solomon"`; resolved by Finding 2 DRY refactor. |

### Source citations for the data fix

- **WEB (World English Bible)** — eBible.org Protestant Edition: <https://ebible.org/web/>.
- **KJV reference totals**: 1,189 chapters / 31,102 verses across 66 books — already documented in `basSBL_VerseCountsGenerator.bas:14-21`.
- **Hebrew vs English split at 2 Sam 18/19, 2 Kgs 11/12, 2 Chr 13/14** — well-known one-verse Masoretic boundary shifts; cross-checked against KJV, NIV, ESV, and the WEB text on eBible.org.
- **3 John verse total (14)** — WEB and KJV both end at v.14; the v.15 split appears in some Greek critical editions (NA28) but is not used in WEB.

### Status

- Data fix for table (A) entries: **APPLIED — 2026-04-28** to `src/basSBL_VerseCountsGenerator.bas`.
- Document content fixes for (B) entries: **deferred** — these are .docm content edits, not code edits.
- Joshua 23 / chapter-break loss: **deferred** — also a document content fix.

### Applied diff summary — `src/basSBL_VerseCountsGenerator.bas`

Four `d.Add` lines edited; seven array values changed; net total verse delta = −1 (3 John).

| Line | Book | Chapter(s) | Before | After | Source |
|---|---|---|---|---|---|
| 83 | 2 Samuel | 18 | 32 | **33** | WEB / KJV — Hebrew 19:1 = English 18:33 |
| 83 | 2 Samuel | 19 | 44 | **43** | WEB / KJV — compensating split |
| 85 | 2 Kings  | 11 | 20 | **21** | WEB / KJV — Hebrew 12:1 = English 11:21 |
| 85 | 2 Kings  | 12 | 22 | **21** | WEB / KJV — compensating split |
| 87 | 2 Chronicles | 13 | 23 | **22** | WEB / KJV — Hebrew 14:1 = English 13:23 |
| 87 | 2 Chronicles | 14 | 14 | **15** | WEB / KJV — compensating split |
| 138 | 3 John | (single) | 15 | **14** | WEB ends at v.14; v.15 split is critical-edition-only |

Reference: <https://ebible.org/web/> (World English Bible, Protestant Edition).

Re-run `AuditVerseMarkerStructure` to confirm the (A) mismatches drop and only the (B) document-content items and the Joshua 23 break remain.

---
