# Row Character-Count and Pitch Diagnostic

## Purpose

Surface body-text rows in the two-column layout that the justifier
has stretched with **excessive inter-word spacing**. These rows have
no soft hyphen at the end (otherwise they would already break) and
are therefore invisible to `RunSoftHyphenSweep_Across_Pages_From`.
This diagnostic identifies them so a soft hyphen can be added at a
suitable break point.

## Concept

For each rendered row in a column:

- **CharCount** — characters in the row (proxy for "looseness").
- **Pitch** — `(rightX - leftX) / (charCount - 1)` in points per
  character. The justifier cannot stretch glyphs, so a short
  justified line stretches the *spaces* and the row's average pitch
  rises. Pitch is the direct signal; CharCount is a sanity check.

Rows that contain a paragraph mark are **excluded** (Bible is
normalized to one verse per paragraph; the last row of any verse is
naturally left-aligned). Rows that already end with a soft hyphen
are excluded too (already handled).

The peak of the pitch histogram per column = the page's natural
"well-set" pitch. Rows with `Pitch > median + threshold` are
flagged as **suspects**.

## Routines

All in `src\basWordRepairRunner.bas`. All are read-only against the
document; only the Phase B navigator changes the active selection
(no edits are ever performed by the macros).

| Routine | Phase | Purpose |
|---|---|---|
| `RunRowCharCountSurvey_Across_Pages_From startPage, pageCount` | A | Walk pages; emit one CSV row per visual line. |
| `RowCharCountSurvey_SinglePage` | A | Per-page worker (called by the driver). |
| `BuildRowCharCountHistogram [thresholdPt]` | C | Bucket eligible rows; compute mode + median; emit histogram and suspects CSVs. Default threshold = 1.0 pt. |
| `ReviewRowCharCountSuspects` | B | Navigator: each call jumps to the next suspect, selects its row, and scrolls into view. |
| `ReviewRowCharCountSuspects_Reset` | B | Clear navigator state (after re-running the histogram). |

## Output files

All under `<docDir>\rpt\`:

| File | Mode | Contents |
|---|---|---|
| `RowCharCount.csv` | append | Per-row records: `PageNum,PageSide,RowIndex,Side,Y,LeftX,RightX,CharCount,Pitch,LastCharCode,EndsWithSoftHyphen,IsParagraphEnd,RangeStart,RangeEnd,FirstChars` |
| `RowCharCount.log` | append | Per-page survey summary + Phase C summary block |
| `RowCharCountHistogram.csv` | overwrite | `Side,Metric,Bin,Frequency` for `Metric` in `{CharCount, Pitch}` |
| `RowCharCountSuspects.csv` | overwrite | Suspect rows + `MedianPitchSide,PitchExcess` columns |

`RowCharCount.csv` is **append-mode**. Two surveys on different
ranges accumulate. Two surveys on the *same* range duplicate that
range's rows in the histogram input — clear the file between
re-runs of the same range.

## Sampling rule

Pick the survey ranges from sections that are **not yet
hyphenated**. Mixing already-hyphenated content into the sample
packs the histogram peak rightward (tighter pitch) and pulls the
suspect threshold too lenient elsewhere. Two non-overlapping
10-page ranges give a baseline plus a cross-check that the median
is stable rather than range-specific.

## Standard procedure

1. **Pick two un-hyphenated 10-page ranges.** Note the start pages.

2. **Clear stale survey output** if `rpt\RowCharCount.csv` exists
   from a previous unrelated run. (Skip this on the first run.)

3. **Survey range 1:**
   ```
   RunRowCharCountSurvey_Across_Pages_From <startA>, 10
   ```

4. **Survey range 2:**
   ```
   RunRowCharCountSurvey_Across_Pages_From <startB>, 10
   ```
   Both surveys append to the same CSV.

5. **Build the histogram** (default threshold 1.0 pt over median):
   ```
   BuildRowCharCountHistogram
   ```
   Read the `medianL` and `medianR` values from the Immediate
   window. Open `rpt\RowCharCountHistogram.csv` in Excel; pivot or
   filter `Bin` vs `Frequency` per `Side` and `Metric`. Confirm a
   sharp peak in CharCount per side and a peak with a right tail in
   Pitch.

6. **Tune the threshold** if needed:
   ```
   BuildRowCharCountHistogram 0.5
   ```
   Lower threshold = more suspects = more candidates. Aim for a
   suspect count you can review without fatigue (a few per page is
   typical).

7. **Reset Phase B state** (only needed if you have run it before
   in this Word session):
   ```
   ReviewRowCharCountSuspects_Reset
   ```

8. **Walk the suspects.** Bind `ReviewRowCharCountSuspects` to a
   keyboard shortcut for fast cycling. The first call announces
   the loaded suspect count. Each subsequent call:
   - Selects the next suspect's row in the document.
   - Scrolls it into view.
   - Shows a status MsgBox naming page, side, CharCount, Pitch,
     and PitchExcess.
   - Dismiss the MsgBox; the row stays selected.
   - Add a soft hyphen with **Ctrl+Hyphen** at the appropriate
     break point in the selected row, or leave the row alone.
   - Re-invoke (the keyboard shortcut) to jump to the next suspect.

9. **Verify with the soft-hyphen sweep.** After the review pass:
   ```
   RunSoftHyphenSweep_Across_Pages_From <startA>, 10, True
   ```
   The newly inserted soft hyphens should classify as **Active**
   (line-breaking) in `rpt\SoftHyphenSweep.csv`. None should be
   classified as **Stray**.

10. **Re-survey to confirm.** Clear `rpt\RowCharCount.csv`, re-run
    the survey on the touched ranges, and re-run
    `BuildRowCharCountHistogram`. Expected pattern:
    - `endShy` count rises by the number of hyphens you added.
    - `eligible` count drops by the same amount.
    - `medianL` and `medianR` tighten (drop) toward the page's true
      well-set pitch.
    - Suspect count drops, often to 0 at the original threshold.
      Lower the threshold to find the next loosest cohort if a
      further pass is desired.

The diagnostic is **self-calibrating**: each fix-and-re-run cycle
exposes the next-loosest rows against a tighter baseline. Stop when
no row exceeds the chosen threshold above a baseline that has
stopped moving.

## Worked validation (page 522, Verso)

| Pass | endShy | eligible | medianL | medianR | suspects @ 0.5 |
|---|---|---|---|---|---|
| Initial | 3 | 81 | 3.786 | 3.714 | 4 |
| After +10 manual soft hyphens | 13 | 71 | 3.714 | 3.714 | 0 |

The Left median collapsed onto the Right median exactly. The
remaining eligible rows are uniform across columns; no further
suspects exist at threshold 0.5 pt over the new baseline.

## Append-mode pitfalls

- `rpt\RowCharCount.csv` accumulates. Re-surveying the same page
  duplicates rows in the histogram input. **Clear the CSV between
  passes that scan the same range.**
- `rpt\RowCharCountHistogram.csv` and `rpt\RowCharCountSuspects.csv`
  are **overwritten** on every `BuildRowCharCountHistogram` call,
  which is the desired behaviour — the suspects list always
  reflects the current state.
- `rpt\RowCharCount.log` is **append-mode**, preserving the run
  history of every survey + histogram pass.

## Notes

- Row grouping uses `LINE_HEIGHT_TOLERANCE = 4.0 pt` (same constant
  used by the soft-hyphen sweep). Y-jumps above this threshold
  start a new row.
- Pitch uses pen-advance form `(rightX - leftX) / (charCount - 1)`
  rather than `width / charCount` so a 1-character row reports
  pitch = 0 (and is excluded by being a paragraph-end row anyway).
- Headers, footers, and footnotes are excluded by the
  `wdMainTextStory` story-type guard.
- The survey runs ~1-2 minutes per 10 pages because each character
  position requires two `Information(...)` queries. Acceptable for
  a one-off diagnostic; revisit with a binary-search row-boundary
  detector if larger ranges become routine.
