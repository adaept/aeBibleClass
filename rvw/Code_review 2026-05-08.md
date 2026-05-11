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

### 4. Row character-count and pitch diagnostic (new task)

Follow-on to the soft-hyphen sweep. With stray soft hyphens handled,
the next visible body-text defect is rows that look loosely set -
rows with too few characters for the column width, which the
justifier stretches with wide inter-word spaces. These rows have no
soft hyphen at the end (otherwise they would already have been
hyphenated), so they need a separate diagnostic.

Plan file: `~/.claude/plans/greedy-napping-yeti.md` (approved
2026-05-08).

Operating doc: [`md/Row_Pitch_Diagnostic.md`](../md/Row_Pitch_Diagnostic.md)
- standard procedure, output file table, append-mode pitfalls,
worked validation from page 522.

#### 4a. Document structure assumption

This Bible is normalized - **one verse per paragraph**:

- Single-row verses are naturally left-aligned (paragraph ends
  mid-column, not justified). Ignore.
- Multi-row verses have a final row that contains the paragraph
  mark and is also left-aligned. Ignore.
- Earlier rows of multi-row verses ARE justified across the full
  column. These are the rows we measure.

Unified rule: **exclude any row that contains a paragraph mark from
the histogram and the suspects list.**

#### 4b. What "pitch" means

Pitch = `(rightX - leftX) / CharCount`, in points per character.
The justifier cannot stretch glyphs, so it stretches inter-word
spaces - a short justified line therefore has a higher average
pitch than a well-filled line. CharCount is a proxy; pitch is the
direct signal. Both are recorded; suspects are ranked by pitch.

#### 4c. Sampling strategy

Pick **two non-overlapping 10-page ranges** from sections that are
**not yet hyphenated**. Mixing already-hyphenated sections into the
sample skews the histogram peak rightward (hyphenated text packs
tighter), which would set the suspect threshold too lenient
elsewhere. Two ranges also confirm the baseline is stable rather
than range-specific.

#### 4d. Process steps

1. Identify two un-hyphenated 10-page ranges. Record the start pages.
2. Run survey range 1: `RunRowCharCountSurvey_Across_Pages_From firstStart, 10`.
3. Run survey range 2: `RunRowCharCountSurvey_Across_Pages_From secondStart, 10`. Both append to the same CSVs.
4. Build histogram: `BuildRowCharCountHistogram`. Inspect `rpt\RowCharCountHistogram.csv` in Excel; confirm a clean peak per side; sanity-check the pitch threshold.
5. Tune threshold if the suspect count looks unreasonable (too few = miss real defects; too many = noise).
6. Run interactive review: `ReviewRowCharCountSuspects`. Cycle through suspects Yes / No / Skip / Cancel; insert soft hyphens manually where Yes (same UX as the soft-hyphen sweep).
7. Re-run the soft-hyphen sweep over the same ranges to verify newly inserted soft hyphens classify as `Active`.
8. Optional second pass: re-run survey on the same ranges to confirm the suspect count drops.

#### 4e. Routines (in `src\basWordRepairRunner.bas`)

Phase A - Survey (LANDED 2026-05-08):

- `RunRowCharCountSurvey_Across_Pages_From(startPage, pageCount)` - driver, mirrors `RunSoftHyphenSweep_Across_Pages_From`. Read-only.
- `RowCharCountSurvey_SinglePage(pageNum, ByRef rowsCum, ByRef userCancelled)` - per-page worker. Walks paragraphs in `[pageStart, pageEnd]`; within each paragraph, walks character-by-character; groups characters into visual rows by Y position with `LINE_HEIGHT_TOLERANCE` (4.0 pt). Skips non-`wdMainTextStory` (headers / footers / footnotes excluded by construction). Emits to `rpt\RowCharCount.csv` and `rpt\RowCharCount.log`. `DoEvents` every 200 chars for cancel responsiveness.
- `FlushRowCharCountRow` (private) - writes one CSV record and updates per-page counters. Computes `Side` via existing `ClassifyColumnAt` from row's first-char X. Pitch = `(rightX - leftX) / max(charCount-1, 1)` pt/char (pen-advance form).

Phase C - Histogram (LANDED 2026-05-08):

- `BuildRowCharCountHistogram(Optional thresholdPt = 1.0)` - reads `rpt\RowCharCount.csv`; filters out paragraph-end rows, soft-hyphen-terminated rows, and non-body rows; buckets remaining rows by `CharCount` (1-char) and `Pitch` (0.1 pt) per side; computes per-side mode of CharCount and median of Pitch; writes histogram and suspects CSVs; appends a summary block to `rpt\RowCharCount.log`.
- `MedianOfSingles` (private) - insertion-sort median for the per-side pitch arrays.

Phase B - Interactive review (PENDING):

- `ReviewRowCharCountSuspects()` - interactive Yes / No / Skip / Cancel walker over `rpt\RowCharCountSuspects.csv`. Uses `RangeStart`/`RangeEnd` from the CSV to select each suspect row in Word, prompts for action, waits for the user to manually insert a soft hyphen on Yes, records the decision to `rpt\RowCharCountReview.log`. Mirrors `SoftHyphenSweep_ByColumnContext_SinglePage`'s prompt scaffolding.

#### 4f. Output files (under `rpt\`)

- `rpt\RowCharCount.csv` - per-row records. Header: `PageNum,PageSide,RowIndex,Side,Y,LeftX,RightX,CharCount,Pitch,LastCharCode,EndsWithSoftHyphen,IsParagraphEnd,RangeStart,RangeEnd,FirstChars`. Append-mode (multiple ranges accumulate).
- `rpt\RowCharCount.log` - per-page summary plus appended Phase C summary block (mode, median, threshold, suspect count).
- `rpt\RowCharCountHistogram.csv` - `Side,Metric,Bin,Frequency`. `Metric` is `CharCount` or `Pitch`. Overwritten on each Phase C run.
- `rpt\RowCharCountSuspects.csv` - rows with `Pitch > medianForSide + thresholdPt`. Full row passthrough plus `MedianPitchSide` and `PitchExcess` columns. Overwritten on each Phase C run.
- `rpt\RowCharCountReview.log` - reserved for Phase B; will record Yes / No / Skip per row.

Survey and histogram routines are read-only against the document.
Only the review phase will mutate, and only via user-confirmed
manual edits (same model as the soft-hyphen sweep).

#### 4g. Test notes

Phase A smoke test (2026-05-08): `RunRowCharCountSurvey_Across_Pages_From 522, 1` on the production doc returned `119 row(s) - body=119 outside=0 paraEnd=35 endShy=3` for a Verso page. 119 - 35 - 3 = 81 histogram-eligible rows, consistent with two columns of justified text on one page.

Phase C smoke test (2026-05-08): `BuildRowCharCountHistogram` on the same single-page survey returned `scanned=119 eligible=81 suspects=0 medianL=3.786 medianR=3.714` (default threshold 1.0 pt). Lowering to `BuildRowCharCountHistogram 0.5` returned `suspects=4`. Left and Right medians within 0.07 pt of each other - confirms `GetColumnBoundsForPage` mirror handling is symmetric.

Closed-loop validation (2026-05-08): user manually added 10 soft hyphens to the loose-row candidates surfaced on page 522, then cleared `rpt\RowCharCount.csv` and re-ran:

| Metric | Before | After | Delta |
|---|---|---|---|
| `endShy` | 3 | 13 | +10 (the manually inserted soft hyphens) |
| eligible rows | 81 | 71 | -10 (now excluded by the end-shy rule) |
| medianL | 3.786 pt | 3.714 pt | -0.072 (tightened; previously pulled up by loose rows) |
| medianR | 3.714 pt | 3.714 pt | unchanged |
| suspects @ 0.5 pt | 4 | 0 | none of the remaining rows are outliers vs the new baseline |

Two takeaways:

1. The **Left median collapsed onto the Right median exactly** (`3.714` both). The page's true well-set per-character pitch is `3.714 pt/char`. Loose rows visibly pull the Left median up; once excluded, the underlying baseline is uniform across columns - which is what we should expect on a JIS B5 mirrored two-column body.

2. **`suspects=0` is self-calibrating, not absolute.** The threshold tracks the per-side median, so each pass of "fix suspects, re-run" surfaces the *next* loosest cohort against a tighter baseline. The natural stopping condition is when no row exceeds the chosen threshold above a baseline that has stopped moving.

Operational implication: Phase B (interactive walker) is **optional, not required**. The manual workflow (`survey -> BuildRowCharCountHistogram -> open suspects CSV in Excel -> add soft hyphens in Word -> clear CSV -> re-run`) was demonstrated end-to-end on page 522. Phase B would speed up the per-suspect navigation step on multi-page runs but is not on the critical path.

Append-mode caveat (2026-05-08): `rpt\RowCharCount.csv` is opened with `For Append`. Re-running the survey on the same page without first clearing the CSV duplicates that page's rows in the histogram input. For the per-page edit cycle, clear the CSV between passes; for accumulating two 10-page samples, do not. Possible follow-up: add a `--clear` mode to the survey driver, or detect overlap with already-present pages and refuse.

### 5. Taxonomy correction: TitleOnePage font (DONE 2026-05-08)

`rpt\Styles\style_TitleOnePage.txt` was refreshed and shows
`Font.Name = "Liberation Serif"`. The taxonomy expectation in
`src\basTEST_aeBibleConfig.bas:321` had `"Times New Roman"` from
the original baseline. Updated the `AuditOneStyle` call to
`"Liberation Serif"`. All other 11 fields (size 36, centered
alignment, line spacing 12, space-before 144, space-after 8,
bold 0, italic 0, base style "") already matched the dump.

No taxonomy header recount needed - this was a value correction
within an existing bucket-1 paragraph-style entry, not an add or
remove.

### 6. Character-style base audit (NEW 2026-05-10)

New diagnostic function `AuditCharStyleBases` in
`src\basStyleInspector.bas`. Scans every character style with
`InUse = True` and prints, to the Immediate window, any style whose
`BaseStyle` is not `"Default Paragraph Font"`. Returns the offender
count. Per the project rule:

> Use "Default Paragraph Font" for ALL character styles. Never use
> "(no style)". - the difference between stable, predictable,
> audit-safe formatting vs. hidden inheritance, inconsistent
> rendering, and style-tree instability.

#### 6a. Function exclusion - "Default Paragraph Font" itself

`Default Paragraph Font` is the root character style. Querying its
`BaseStyle` returns empty because there is nothing above it to
inherit from, and it cannot be based on itself. Including it would
inflate the offender count by one and report a non-fixable false
positive. The loop skips it by name; the header docstring documents
the exclusion. The reported count therefore reflects only styles
that can actually be repointed to the default.

#### 6b. First-run finding (2026-05-10)

`?AuditCharStyleBases` returned **14** initially (before the
exclusion fix); **13** real fixes after excluding
`Default Paragraph Font`:

12 styles based on `(no style)` - in `Bold`, `Bold italic`,
`Italics`, `Endnote Reference`, `Figure`, `Figure Caption`,
`Footnote`, `Footnote Label`, `Footnote marker`, `Hyperlink`,
`Table Cell Head`, `Words of Jesus`. Several are Word built-ins
(e.g. `Hyperlink`, `Endnote Reference`); the `BaseStyle` property
remains settable on built-ins via VBA even when the UI displays an
"underlying properties" description.

1 chained-inheritance case: `Page Number  ->  Footer Char`.
`Footer Char` is the linked character style auto-generated by Word
for the `Footer` paragraph style. Per the rule this is exactly the
"hidden inheritance / style-tree instability" risk; repoint
`Page Number.BaseStyle` to `Default Paragraph Font` directly.

Action: re-base all 13 to `Default Paragraph Font`, then re-run
`?AuditCharStyleBases` and confirm it returns 0.

#### 6c. ScanCharStyleApplications (NEW 2026-05-10)

Companion function in `src\basStyleInspector.bas`. Distinguishes
character styles that are merely *present in the styles palette*
from those *actually carried by at least one run of text*.

Motivation: Word's `Style.InUse` flag is True for any custom
character style from the moment it is created, regardless of
whether any run carries it - so InUse alone cannot answer "is
this style live in the document?". On the first run of
`?AuditCharStyleBases` (2026-05-10), `Bold`, `Bold italic`, and
`Italics` appeared in the InUse list, but the user noted they are
not specifically applied anywhere. This function settles that
question.

Mechanism: for each in-use character style (excluding
`Default Paragraph Font` for symmetry with `AuditCharStyleBases`),
runs Find with `.Style = oStyle` and `.Text = ""` across each
primary StoryRange (main, footnotes, endnotes, headers, footers).
First hit short-circuits to `Applied`; no hit yields `Unapplied`.

Output to Immediate window:
```
<styleName>  ->  Applied
<styleName>  ->  Unapplied
ScanCharStyleApplications: checked N in-use character style(s) (excluding "Default Paragraph Font"); A Applied, U Unapplied.
```

Returns the count of `Unapplied` styles - the palette-cruft
candidates.

#### 6d. Decision rule for unapplied character styles

For each style returned `Unapplied`:

| Outcome | Decision | Action |
|---|---|---|
| 0 application sites | Palette cruft | Delete the style. Do not add to the approved list. |
| Few sites, accidental | Cleanup needed | Retag the offending runs to the intended style; then delete. |
| Few sites, intentional | Genuine but rare | Add to approved list (base = `Default Paragraph Font`); promote per the bucket-2 audit when ready. |
| Many sites | Genuine palette member | Add to approved list; re-base to `Default Paragraph Font`. |

The rule "promote only what is intentionally used" applies. Adding
unused styles to the approved list creates audit obligations
(`AuditOneStyle` calls, taxonomy header recount) for nothing that
produces visible output. Deletion is the cleanest way to remove
ambiguity.

Open question for `Bold` / `Bold italic` / `Italics`: pending
results of `?ScanCharStyleApplications` against the production
docx.

#### 6e. First scan and palette categorisation (2026-05-10)

`?ScanCharStyleApplications` returned **36 Unapplied / 9 Applied**
on first run. The 36 split into distinct categories explaining
why most of the unfamiliar names appeared:

**A. HTML legacy styles (Word built-ins)**
`HTML Acronym`, `HTML Cite`, `HTML Code`, `HTML Definition`,
`HTML Keyboard`, `HTML Sample`, `HTML Typewriter`, `HTML Variable`.
Activated the moment Word touches HTML content (paste from a web
page, import from `.htm`, or content brought from another doc that
itself once touched HTML). Once `InUse=True` flips, it sticks.

**B. Modern Office collaboration features (built-ins)**
`Hashtag`, `Mention`, `Unresolved Mention`, `Smart Hyperlink`,
`SmartLink`, `Placeholder Text`, `Comment Reference`. Activated by
touching comments, @-mentions, modern hyperlinks, or content
controls - even briefly during an editing session that was later
undone.

**C. Built-in semantic emphasis family**
`Strong`, `Emphasis`, `Subtle Emphasis`, `Intense Emphasis`,
`Subtle Reference`, `Intense Reference`, `Book Title`. Word's
built-in style-set / theme-driven emphasis. Switching theme or
style set flips these to `InUse=True` even without application.

**D. Built-in note / reference machinery**
`Endnote Reference`, `Footnote Reference` (Applied),
`Footnote`, `Footnote Label`, `Footnote marker`, `Line Number`,
`Page Number` (Applied), `Hyperlink` (Applied),
`FollowedHyperlink`. Some are actively used (Page Number on the
footer, Hyperlink, Footnote Reference). Others are part of the
same built-in family that activated alongside them.

**E. Custom styles never applied (the only deletable category)**
`AuthorQuote`, `Bold`, `Bold italic`, `Italics`, `Figure`,
`Figure Caption`, `Table Cell Head`, `Normal text`. These are
project-template leftovers - created during earlier design
iterations, never wired into a final ribbon / macro path.
Particularly suspect: `Bold` / `Italics` / `Bold italic` likely
predate the standardisation on `EmphasisBlack` / `EmphasisRed`.

#### 6f. ScanCharStyleApplications enhancement (2026-05-10)

Critical distinction added: `Style.BuiltIn` is True for categories
A-D (cannot be deleted - Word recreates them next session) and
False only for category E. Without surfacing this, the raw "36
Unapplied" number was misleading because most of those entries
are not actionable.

Function output enhanced:
- Per-style line now appends `[Builtin]` or `[Custom]`:
  ```
  Bold  ->  Unapplied  [Custom]
  HTML Sample  ->  Unapplied  [Builtin]
  Hyperlink  ->  Applied  [Builtin]
  ```
- Summary now splits both Applied and Unapplied counts by Builtin
  vs Custom and prints a final "Deletable cruft" line:
  ```
  Applied   : Builtin=N  Custom=N  (total N)
  Unapplied : Builtin=N  Custom=N  (total N)
  Deletable cruft (Unapplied & Custom): N
  ```
- **Return value changed from total Unapplied to Unapplied & Custom.**
  This is the action-list size: how many styles are candidates for
  `Style.Delete`. On this docx the new return value would be 8
  (category E above), down from the misleading 36.

This is an intentional contract change for a 2-day-old function;
no callers in production yet. If a future caller needs the raw
total, add a sister function rather than restoring the old return.

#### 6g. Action plan from 6e + 6f

1. Re-run `?ScanCharStyleApplications` and confirm the new return
   value is **8** (the category-E count).
2. For each of the 8 deletable styles, do one final visual
   confirmation that no run is using the style (Find with Style
   filter in Word UI, or trust the function which already
   performed exactly this check). Then delete via
   `ActiveDocument.Styles("<name>").Delete`.
3. For categories A-D Unapplied: leave alone. Word will recreate
   them on demand. Do not add to the approved style list.
4. Re-run `?AuditCharStyleBases` and confirm any remaining
   offenders are categories A-D (built-in) or `Page Number` (the
   chained-inheritance fix from 6b). Re-base the latter to
   `Default Paragraph Font`.

#### 6h. Second scan after cleanup (2026-05-10)

`?ScanCharStyleApplications` returned a much-cleaner state:

```
checked   : 36 in-use character style(s) (excluding "Default Paragraph Font")
Applied   : Builtin=3   Custom=6   (total 9)
Unapplied : Builtin=25  Custom=2   (total 27)
Deletable cruft (Unapplied & Custom): 2
```

**Delta from 6e:** in-use 45 -> 36 (9 styles deleted between
runs); deletable cruft 8 -> 2. The 9 removed were the category-E
candidates from 6e: `Bold`, `Bold italic`, `Italics`, `Figure`,
`Figure Caption`, `Footnote`, `Footnote Label`, `Footnote marker`,
`Table Cell Head`. Cleanup completed cleanly - no built-in
collateral damage.

**Real palette now visible (the 9 Applied):**

- Built-in (3): `Footnote Reference`, `Hyperlink`, `Page Number`.
- Custom (6): `Selah`, `EmphasisBlack`, `EmphasisRed`,
  `Words of Jesus`, `Chapter Verse marker`, `Verse marker`.

**Two custom Unapplied flagged for decision:**

- **`AuthorQuote`** - listed in section 2's bucket-2 promotion
  set alongside Selah / EmphasisBlack / EmphasisRed /
  Words of Jesus, but unlike those four it is **not actually
  applied to any run**. Either created for a planned feature that
  has not shipped (keep; base on `Default Paragraph Font`; defer
  promotion until first use), or residue from earlier drafts whose
  application sites were rewritten (decide whether to reintroduce
  or delete). Verify against project intent before acting.

- **`Normal text`** - not in any promotion list; generic name
  suggests template residue. Strong delete candidate; one final
  visual sweep, then `ActiveDocument.Styles("Normal text").Delete`.

**Two custom Applied missing from the section 2 promotion list:**

- `Chapter Verse marker` and `Verse marker` are both Applied
  Custom and clearly intentional (Bible verse numbering, with
  colour intent already locked in section 1: orange `42495` and
  medium green `7915600` respectively). Neither appears in
  section 2's bucket-2 list. **Add both to the bucket-2 promotion
  set when the `AuditOneStyle` character-style extension lands.**

**Built-in Unapplied (25):** unchanged - Word will recreate any
that get deleted on next theme switch / paste / comment use.
Ignore in audit reports; not actionable.

**Updated section 2 promotion list (proposed):**

- `Footnote Reference` (built-in, Applied)
- `Selah` (custom, Applied) - already listed
- `EmphasisBlack` (custom, Applied) - already listed
- `EmphasisRed` (custom, Applied) - already listed
- `Words of Jesus` (custom, Applied) - already listed
- `AuthorQuote` (custom, Unapplied) - **deferred** until first
  application or removed from set
- `Chapter Verse marker` (custom, Applied) - **add**
- `Verse marker` (custom, Applied) - **add**

This brings the bucket-2 audit in line with what is actually
carried by runs in the document.

#### 6i. Taxonomy update (DONE 2026-05-10)

`PromoteApprovedStyles` was updated to add three new paragraph
styles. Specs captured via `DumpStyleProperties`:

- **`BibleIndexList`** (Paragraph) - Liberation Serif 11pt,
  centered, line spacing 12, no space-before/after. 1 tab stop at
  172.8 pt (Right, Dots).
- **`ParallelHeader`** (Paragraph) - Carlito 11pt bold, left
  aligned, line spacing 12, space-after 8. No tab stops.
- **`ParallelText`** (Paragraph) - Carlito 7pt bold, left aligned,
  line spacing 12, space-after 8. 6 tab stops at 7.2 / 57.6 / 108
  / 158.4 / 208.8 / 259.2 pt (Left, Spaces) - parallel-passage
  column grid.

`AuthorQuote` removed from the bucket-2 existence-verified set per
the 6h decision (Unapplied + not in any near-term plan).

Taxonomy edits in `src\basTEST_aeBibleConfig.bas`:

1. Added three `AuditOneStyle` calls in the bucket-1
   fully-specified block, dated `2026-05-10`.
2. Added two `AuditStyleTabs` blocks: `BibleIndexList` (1 stop)
   and `ParallelText` (6 stops).
3. Removed the `AuditOneStyle "AuthorQuote"` call.
4. Updated the `RUN_TAXONOMY_STYLES` header comment:
   - 45 styles -> 47 (-1 AuthorQuote, +3 new)
   - 6 tab-stop specs -> 8 (+BibleIndexList, +ParallelText)
   - Total checks 51 -> 55
   - Fully-specified bucket 34 -> 37
   - Existence-verified bucket 8 -> 7

Predicted result: **52 PASS / 3 FAIL** (the 3 FAILs are the
not-yet-created `BodyTextContinuation`, `AppendixTitle`,
`AppendixBody`, unchanged from the prior arc). Confirm by running
`RUN_TAXONOMY_STYLES` and checking
`rpt\StyleTaxonomyAudit.txt`.

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
