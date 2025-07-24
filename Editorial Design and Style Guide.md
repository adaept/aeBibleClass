# 📘 Editorial Design and Style Guide

_Audit-safe layout standards, suffix diagnostics, and repair architecture for verse-based documents._

## 📌 Scope

This guide defines diagnostic rules and layout behavior used in automated repairs and pre-sweep evaluations of verse marker blocks across paginated documents. It captures typographic integrity, audit history, session tracking, and editorial safety.

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
