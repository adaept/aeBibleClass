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

## 🔧 Suggested Refinement Targets for Audit Suite

### 📊 Paragraph Audit Drift

- **Trigger:** Test 16 mismatch (`16853` vs. `16471`)
- **Action:**  
  - Export paragraph context (style, section, shape type) when count differs from expected.
  - Add optional logging toggle for shape vs. inline content distinction.

---

### 📋 Header/Footer Audit Enhancement

- **Goal:** Improve granularity of tab-only paragraphs in headers/footers.
- **Action:**  
  - Split counts by shape type (textboxes vs. inline)
  - Include adjacent style info for anomaly detection

---

### 🎯 Style Coverage Analysis

- **Goal:** Detect unused or ghost styles inflating style pools.
- **Action:**  
  - Compare defined styles vs. applied styles.
  - Flag unused styles with section and frequency metadata.

---

### 🛡️ Redundancy Audit

- **Goal:** Catch unintentional function reuse across tests.
- **Action:**  
  - Script to flag identical function calls in adjacent test cases.
  - Use `Repeatable` or `IntentionalRepeat` flags to whitelist exceptions.

---

### 📈 Forecast & Runtime Metrics

- **Goal:** Tie audit density to performance profiling.
- **Action:**  
  - Log runtime per test
  - Add density metric (`Hits/Minute`) to forecast audit cost
  - Export to CSV with `SessionID` for historical tracking

---

### 🕵️‍♂️ Suffix Anomaly Deep-Dive

- **Goal:** Extend Hair Space and NBSP audit to adjacent punctuation
- **Action:**  
  - Track context before/after suffixes
  - Log anomalies with full ASCII and style signature

---

## 🧠 Word Style Management – Manual Restart Guide

This document outlines a modular strategy to safely reinitiate the Word Style Management thread after a prior session crash. Designed for stability, traceability, and iterative development.

---

### 🧩 Stepwise Restart (No History Required)

### 1. Clarify Scope

Define target focus:

- Layout drift
- Style inheritance
- Suffix audits
- Font cleanup

Set intent:

- Diagnostics only
- Include auto-repair logic

### 2. Select Module for Reactivation

Recommended starting points:

- `StyleUsageDistribution()` — audits style counts per page
- `TrackLayoutDrift()` — detects visual drift across multi-column regions
- `SuffixAuditTracker()` — logs suffix types per paragraph and page

### 3. Session Context Setup

Configure session parameters:

- Assign new Session ID (timestamp-based or manual)
- Logging preferences:
  - Verbosity level
  - ASCII layout mapping
  - Skipped case tracking
- Export options:
  - CSV
  - Embedded comment logs
  - GitHub-ready output

### 4. Minimal Test Run

Use a low-density test document or select 2–3 sample pages:

- Isolate visual drift or suffix anomalies
- Collect clean debug output
- Verify correctness before scaling to full document

---

## 🔧 Optional Enhancements

### Performance Modules

- Timing analysis
- Forecast tracking (to support session comparisons)

### Layout Anomaly Detection

- Split marker boundaries
- Visual misalignment across columns/pages

### Font Diagnostics

- Legacy font detection in headers, footers, and main body
- Font drift across sections and styles

---

## ✅ Action Summary

| Task                        | Status     |
|----------------------------|------------|
| Define scope               | 🔲 Pending |
| Select module              | 🔲 Pending |
| Configure session context  | 🔲 Pending |
| Launch test document       | 🔲 Pending |
| Apply optional enhancements| 🔲 Optional |

---

## 🧮 Version Control Recommendations

- Save module scripts per version with timestamp in GitHub
- Track suffix normalization metrics and export per run
- Use commit messages to annotate layout quirks and engine decisions

---

## 📄 Copilot Pages: Workflow Guide (July 2025)

## Overview

Copilot Pages is best used as a stable workspace for audit logs, macro iterations, and suffix tracking—not for fluid, spontaneous debugging. It offers editability, persistent context, and structured layout, but lacks the conversational nuance of chat.

---

## 🔍 Pages vs Chat Usage Matrix

| Task Type                   | Best Used In | Reasoning                                                  |
|----------------------------|--------------|------------------------------------------------------------|
| Macro iterations           | Pages        | Audit-friendly revisions with editable history             |
| Style audit logs           | Pages        | Add summaries, CSV-style breakdowns, suffix notes          |
| Punctuation-edge cases     | Chat         | Better for rapid-fire context and emotional nuance         |
| Font repair tracking       | Pages        | Consolidate detection notes and outcome logs               |
| Session instability analysis| Both        | Track failures in Pages, troubleshoot live in chat         |

---

## 🚀 How to Start a Page

- Hover over any Copilot response  
- Click **“Edit in a page”**  
- Or, create one manually via the “Pages” tab in the sidebar  
- Pages are auto-saved and support markdown-style formatting

---

## 💡 Tips for Clarity and Auditability

- Use headers like `# Layout Drift Audit — July 25`  
- Log skipped cases, suffix anomalies, repair summaries  
- Ask Copilot to refactor code, summarize output, or reformat content live within the Page

---

## 🧱 Why Pages Feel Clunky (Today)

- No conversational turn-taking or inline commentary  
- Less intuitive than chat for real-time diagnostics  
- Ideal for persistent records, not exploratory reasoning

---

## 📘 Official Help Resource

Visit: [Copilot Pages Help](http://aka.ms/copilot-pages-help)

---

## 🧑‍🔬 Follow-up Issues

- 🧠 Section 144: Header=77 (M) and Section 146: Header=73 (I)—are those initial glyphs from chapter metadata? Might be worth flagging for suffix tracking. (TestHeaderFooterStyleScan)

- 🔍 Section 147’s tab-tab (ASCII=9) pairing may mark an empty pair or control-only layout. Could use that as a soft indicator for skipped suffix density? (TestHeaderFooterStyleScan)
