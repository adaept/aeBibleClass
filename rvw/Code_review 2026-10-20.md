# Code review - 2026-10-20

Carries forward the open items from `rvw/Code_review 2026-09-25.md` (the last dated
Code_review doc), the backlog lists in the old session manifest
(`sync/session_manifest_backup_2026-10-20.txt`, latest backlog at lines 3795-3825 and
3905-3925), and the 2026-10-19 task list now at the top of `sync/session_manifest.txt`.

State as of aeBibleClass `bd304ff` (2026-10-20). Full suite: 94/94 PASS, 1605 s.

## Legend

- Status: ✅ Done | 🟡 In progress | 🔴 Blocked | ⚪ Not started
- Priority: 🔴 High | 🟡 Med | 🟢 Low
- Done items record `commit <sha> (date)`. Open items record the commit they were last
  touched at, or `-` if none.

## Quick overview

| # | Item | Status | Priority | Commit |
|---|------|--------|----------|--------|
| 1 | `applySpellingVariant` live-check (gated, operator go-ahead required) | 🔴 Blocked | 🔴 High | built 4c7f7d8 area (2026-09-25), check not run |
| 2 | Test-ordering experiment: Find-heavy test before Tests 82/94 | ⚪ Not started | 🟡 Med | - |
| 3 | Defender CPU: lap-line MsMpEng or MpPerformanceRecording; exclusion is the operator's call | ⚪ Not started | 🟡 Med | - |
| 4 | Baseline tag re-pin (LAST, after JS matches newest VBA) | ⚪ Not started | 🟡 Med | - |
| 5 | W-2: fresh RadiantWordBible `InspectStyleUsage` run on the v1.1 export | 🔴 Blocked (v1.1 not built) | 🟡 Med | - |
| 6 | Genesis 1 verse-level Strong's spike (OSHB only) | ⚪ Not started | 🟡 Med | - |
| 7 | Hyphenation: live confirmation that `Range.Text` / Office.js leaves optional hyphens alone | ⚪ Not started | 🟡 Med | - |
| 8 | Canary ideas (call counters, Find-cost canary, paging-delta, "slower than previous" column) | ⚪ Not started | 🟢 Low | - |
| 9 | `ImportAllVBAFiles` `Skipped<>1`: why `VBComponents.Remove` fails | ⚪ Not started | 🟢 Low | - |
| 10 | USFM export from docm (needs its own plan doc first) | ⚪ Not started | 🟢 Low | - |
| 11 | "Introduction to the Bible" front-matter essay | ⚪ Not started | 🟢 Low | - |
| 12 | R16 WEBBE review (other 1,550/14,169; UK-print prerequisite) | ⚪ Not started | 🟢 Low | - |
| 13 | 2 Samuel 21:8 confirmation | ⚪ Not started | 🟢 Low | - |
| 14 | Cases 49/50 unported to aeBibleAddin, formally untriaged | ⚪ Not started | 🟢 Low | - |
| 15 | Timing-comparison discipline: fresh docm copy, note start canary | ✅ Standing practice | 🟢 Low | - |

## Open items, detail

### 1. `applySpellingVariant` live-check - 🔴 Blocked, 🔴 High

Required and high priority, but explicitly deferred by the operator on 2026-09-25;
rationale to follow. Do not run without an explicit go-ahead and do not guess a
rationale. The live wiring is built and committed (search/replace, format-preserving).
Runbook: `aeBibleAddin/taskpane/src/word-audit-scanner.LIVE_CHECK.md`, "Shape 8,
spelling-conversion write extension".

### 2. Test-ordering experiment - ⚪ Not started, 🟡 Med

Tests 82 (528 s) and 94 (684 s) are about 75% of the run and scale with the
style-read canary. Earlier notes say Find-based tests (79, 88) reset the slow state.
Put a Find-heavy test just before 82 and 94 and compare. Use a fresh docm copy and
record the start canary (item 15).

### 3. Defender - ⚪ Not started, 🟡 Med

Defender real-time scan was seen at about 12% CPU during some tests. Add MsMpEng CPU to
the lap line, or run `New-MpPerformanceRecording` / `Get-MpPerformanceReport`. Folder or
process exclusion (`C:\adaept`, `WINWORD.EXE`) is the operator's decision, not yet made.

### 4. Baseline tag re-pin - ⚪ Not started, 🟡 Med

From 2026-10-09. Done last, after JS matches the newest VBA.

### 5. W-2 RadiantWordBible run - 🔴 Blocked, 🟡 Med

Needs the v1.1 export. v1.1 is the next release, after the latest testing concludes, and
is not built yet. v1.0 shipped as code `73ffd51`, record `04cb54f` (2026-10-09).

### 6. Genesis 1 Strong's spike - ⚪ Not started, 🟡 Med

Scholarly grounding roadmap step 3: verse-level spike on Genesis 1 in
`aeRWB/tools/web-diff/`, OSHB data only (cWEB skipped by operator decision). Sources
fetched in roadmap step 2 (CLOSED 2026-09-25). Leads to Phase 5 in aeBibleAddin, then
STEPBible.

### 7. Hyphenation live confirmation - ⚪ Not started, 🟡 Med

From 2026-09-25 item 10(a). Export is clean (zero U+00AD / U+2010), but the check was
indirect. Confirm directly against real hyphenated docm content that `Range.Text` and
Office.js text access leave manually placed optional hyphens untouched. Only if that
fails, build the non-hyphenated reference-file mechanism. Hyphenation for other
languages is not planned (left-aligned i18n text makes it unnecessary).

### 8. Canary ideas - ⚪ Not started, 🟢 Low

Per-bucket call counters for slow tests, a Find-cost canary, a paging-delta flag, and a
"slower than previous runs" column from older TestLaps files.

### 9. `ImportAllVBAFiles` `Skipped<>1` - ⚪ Not started, 🟢 Low

Root cause confirmed: `VBComponents.Remove` silently fails for `basBiblePalette`,
`basTEST_aeBibleTools` and `Module1`. Why it fails is still unknown. Diagnostic call site
(`DumpLiveVBComponents [post-delete]`) is left in place.

### 10. USFM export from docm - ⚪ Not started, 🟢 Low

A new export path with a style-to-USFM marker map, a VBA document walker, and
book/chapter/front-matter handling (reuses the Psalm front-matter work). Strong's re-embedding
and footnote cross-references are out of scope for a first pass. Write a dated plan doc
before starting.

### 11. Introduction to the Bible - ⚪ Not started, 🟢 Low

Reader-level front-matter essay modelled on `engwebu_usfm/00-FRTengwebu.usfm`. Input is
the source-provenance material (Strong's / OSHB / RP licensing, WEBU / WEBBE
relationship). Needs its own draft/review/approve session.

### 12. R16 WEBBE review - ⚪ Not started, 🟢 Low

Plan only: `rvw/Plan_rwb_webbe_review_2026-10-04.md`. Not urgent (US prints before UK).
Three separate UK-print prerequisites: this review, spelling-substitution wiring, and VBA
hyphenation/alignment work.

### 13. 2 Samuel 21:8 confirmation - ⚪ Not started, 🟢 Low

Carried from every backlog list since 2026-09. The detail was not found in the
2026-09-25 doc; check the old manifest backup before starting.

### 14. Cases 49/50 - ⚪ Not started, 🟢 Low

`CountAuditStyles_ToFile` and `SummarizeHeaderFooterAuditToFile` are unported to
aeBibleAddin and have no formal disposition. Believed permanently out of scope but not
re-investigated. Decide won't-fix vs. someday.

### 15. Timing-comparison discipline - ✅ Standing practice

Open a fresh copy of the docm and note the start canary. Standalone numbers are not
comparable to in-suite numbers.

## Closed since 2026-09-25 (for traceability, not open)

| Item | Status | Commit |
|------|--------|--------|
| Spirit's (Romans 8:27) fixed in docm, propagated to rwb.txt | ✅ Done | aeRWB `ba87d5e` (2026-09-26) |
| Psalm superscriptions / BOOK 1-5 sidecar export (R13) | ✅ Done | aeRWB `5c52a8d` (2026-09-26) |
| docm vs. rwb.txt full sync (R12) | ✅ Done | closed 2026-09-26 |
| Jeremiah 37:10 stray CVM marker (`37:910`), Test 91 | ✅ Done | closed 2026-10-05 |
| JS-port parity script, Tests 90/91 | ✅ Done | closed 2026-10-05 |
| Divine-name outlier inconsistency (Zechariah 14:20) | ✅ Done | `e76ef8d` (2026-09-17) |
| Song of Solomon naming | ✅ Done | `fdd4a1a` (2026-09-22) |
| Test-suite speed-ups, 3292 s to 1605 s | ✅ Done | `2cb1fae` (2026-10-19) |
| "Running annotated Strong's search"; Selah attribution | ✅ Moot | closed 2026-09-26 |

## Provenance

Open-item statuses above were rebuilt from the 2026-09-25 doc, the 2026-10-05 and
2026-10-06 backlog lists in the manifest backup, and the 2026-10-19 task list. Items 13
and the exact commit for item 1's build were not re-verified against git this session.
