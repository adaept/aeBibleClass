# 📘 Editorial Design and Style Guide

_Audit-safe layout standards, suffix diagnostics, and repair architecture for verse-based documents._

## 📌 Scope

This guide defines diagnostic rules and layout behavior used in automated repairs and pre-sweep evaluations of verse marker blocks across paginated documents. It captures typographic integrity, audit history, session tracking, and editorial safety.

## 🧩 Code Architecture Overview

The Editorial Design and Style Guide (EDSG) works in conjunction with VBA code embedded in a `.DOCM` file. The operation of the code is modular and audit-driven, with the following components:

- **`aeBibleClass`**  
  A class module responsible for running a battery of 40+ diagnostic tests on document content. These tests cover layout integrity, style consistency, suffix behavior, and marker structure.

- **`aewordgitClass`**  
  A class module used to export VBA code and audit logs to GitHub. It supports version tracking, task development, and collaborative refinement of repair logic.

- **`basWordRepairRunner`**  
  A standard module that executes layout diagnostics and automated repairs. It applies rules defined in EDSG to ensure safe, reversible, and editorially sound modifications.

## 🧠 EDSG Development Module Manifest (Items 1–10)

### **VBA Development Routines**

Purpose-built scaffolding to:

- Ensure all macro logic is auditable and version-aware.
- Maintain reversible transitions between Word and GitHub.
- Provide test harnesses that simulate layout, suffix, and style variants.
- Respect editorial boundaries—never modify punctuation or core content.

### 1. `basBibleRibbon`

- Applies macros from a custom Word Ribbon
- Extendable using Office RibbonX Editor.
- Designed for Office 365 Word only.

### 2. `basChangeLogaeBibleClass`

Change log engine for `aeBibleClass` operations:

- Timestamped, suffix-aware log entries.
- Captures transformations of sacred/editorial text.
- Supports layout boundary reconciliation and suffix audits.

### 3. `basChangeLogaewordgit`

Tracks automation-driven changes for `aewordgit` routines:

- Records macro-induced layout and style shifts.
- Links macro sessions with Git-exportable artifacts.
- Supports suffix normalization and audit-friendly exports.

### 4. `basImportWordGitFiles`

Git-to-Word rehydration tool:

- Applies macros from Git-exported modules back to Word safely.
- Verifies suffix resolution, style inheritance, and layout sanity.
- Supports preview loops and skip-logging before macro activation.
- Easily adaptable with direct access to the code.

### 5. `basTESTaeBibleClass`

Test initiator for BibleClass routines:

- Simulates Word inheritance quirks and suffix variants.
- Validates repairs against editorial safety constraints.
- Logs skipped cases with ASCII + style context.

### 6. `basTESTaeBibleFonts`

Project font related routines:

- 
- 
- 

### 7. `basTESTaeBibleTools`

Project development tools:

- 
- 
- 

### 8. `basTESTaewordgitClass`

Initiator for Git-bound Word exports:

- Runs dry-export tests for macro readability and fidelity.
- Audits namespace integrity and suffix tracking completeness.
- 

### 9. `basWordRepairRunner`

Document QA development tools:

- 
- 
- 

### 10. `basWordSettingsDiagnostic`

Word environment QA tools:

- 
- 
- 

---

## 🔄 EDSG Workflow — Word Automation Lifecycle

1. **Source Document in Word**
   - Contains layout anomalies, suffix variants, and legacy styles.

2. **aeBibleClass Macros**
   - Repairs layout and suffix inconsistencies.
   - Respects editorial boundaries and punctuation integrity.

3. **basChangeLogaeBibleClass**
   - Logs macro actions with timestamp, style context, and audit notes.
   - Records layout boundaries, suffix findings, and skipped cases.

4. **basTESTaeBibleClass**
   - Runs simulated tests using fake layouts or edge-case paragraphs.
   - Validates macro safety, logs skipped repairs with full ASCII trace.

5. **Macro Export to GitHub**
   - Macro modules, diagnostic reports, and style audits are versioned.

6. **basChangeLogaewordgit**
   - Captures export session context: suffix stats, macro results, and style map.
   - Tracks performance and module evolution across runs.

7. **basTESTaewordgitClass**
   - Dry-runs Git-bound macros for compatibility and namespace sanity.
   - Verifies macro readability and layout assumptions.

8. **GitHub Repository**
   - Stores modular code, suffix audit outputs, and diagnostic histories.
   - Supports rollback, team collaboration, and public reference.

9. **basImportWordGitFiles**
   - Safely rehydrates macro modules back into Word.
   - Allows preview/confirm flow, skips unsafe cases, maintains suffix state.

10. **Final Document Pass in Word**
    - Clean, audited, repaired document with full changelog support.
    - Suffix reports and audit logs retained for tracking or next-run forecasting.

---

## 🧭 Layout Roles and Styles

| Style         | Purpose                                         |
|---------------|--------------------------------------------------|
| `Heading 1`   | Title page or section header — never contains markers |
| `Heading 2`   | Chapter identifiers (e.g., `CHAPTER 5`) — paired with `"Chapter Verse marker"` = `5` |
| `Verse marker`| Contains numeric verse fragments like `1513` or `.1513` |
| `Chapter Verse marker`| Style applied to chapter digits (e.g., `5`) on `CHAPTER X` pages |

> ✅ A single page may contain **0–2 `Heading 2`** entries.

---

## 🛠 Page-Level Pre-Sweep Audit Features

| Feature ID | Metric | Description |
|------------|--------|-------------|
| 1 | `sessionID` | Timestamp or GUID tag per sweep run |
| 2 | `ascii12Count` | Count of layout wrappers using `Chr(12)` |
| 3 | `suffix160Count` | Count of suffixes using `Chr(160)` |
| 4 | `suffixHairSpaceCount` | Count of suffixes using `Hair Space` (`ChrW(8239)`) |
| 5 | `suffixOtherCount` | Count of suffixes using irregular spacing characters |
| 6 | `paragraphCount` | Total paragraphs on page — used for layout density |
| 7 | `chapterCount` | Count of `Heading 2` entries per page |
| 8 | `heading2Titles()` | Text content from each `Heading 2` block |
| 9 | `firstVerseMarker` | First verse digit found in layout order |
| 10 | `lastVerseMarker` | Last verse digit found on page |
| 11 | `excessiveSpacingFlag` | Boolean — flag triggered by overuse of NBSP, Hair Spaces, or alignment drift |
| 12 | `checksumValue` | Hash or digest of page text + styles for drift tracking |
| 13 | `qualityRating` | Label: `Clean`, `Minor Drift`, `Needs Rescan` |

---

## 🧪 Suffix Audit Standards

| ASCII | Name           | Treatment        |
|-------|----------------|------------------|
| 160   | Non-breaking space (NBSP) | Preferred |
| 8239  | Hair Space      | Acceptable if style/font audited |
| 32    | Regular space   | Risky — prone to layout breaks |
| Other | Thin/En space, punctuation | Logged as drift |

Suffixes should only follow complete digit blocks and pass style filter for `"Verse marker"` with appropriate font (`Calibri 9pt` preferred).

---

## 🧱 Repair Runner Safeguards

- Repairs only execute when:
  - Marker style + layout side are valid
  - No ambiguity in chapter context
  - No punctuation or editorial text at risk
- Audit runs before repair and logs:
  - All skipped markers
  - ASCII status of prefix and suffix characters
  - Font and style data on every marker

---

## 📊 Session Tracking

- Rescans only update sweep data for the relevant page.
- Multi-page runs include cumulative totals and per-page details.
- Historical snapshots allow delta comparison of audit metrics.
- Manual editorial adjustments (e.g. suffix replacement, style overrides) are tracked by checksum deviation.

---

## 🔧 Future Standards (Planned)

- Chapter-to-verse integrity checks (`e.g., CHAPTER 5 → markers 51–5x`)
- Repair threshold toggles based on wrapper density or suffix accuracy
- Auto-suggestions for normalization (e.g., replace Hair Spaces with NBSP in `"Verse marker"` context)
- Audit map rollup across entire document with quality heatmap

---

## 🧑‍🔬 Philosophy

Layout repair must never override editorial meaning. This guide exists to preserve structural integrity, typographic consistency, and auditable decision paths—so that every automation is reversible, explainable, and safe.

---
