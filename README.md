# README.md

This file provides guidance when working with code in this repository.

## Project Overview

VBA automation framework for Microsoft Word 365 (Office 365) that audits, validates, and repairs an 800+ page Study Bible document. The project enforces layout consistency, style integrity, and typographic precision through 40+ diagnostic tests and a 14-stage Scripture reference parser.

**Platform:** Word 365 only. No Excel dependencies. All Immediate Window output must be ASCII-only.

## Key Commands

All macros run from the Word VBA Immediate Window or via the custom Ribbon.

**Run tests:**

```vba
RUN_THE_TESTS              ' Run all tests
RUN_THE_TESTS(8)           ' Run test #8 only
RUN_THE_TESTS("varDebug")  ' Run with debug output
```

**Export VBA to /src (for Git):** Use `aeWordGitClass` export routine from the live `.DOCM`.

**Scripture parser entry point:** `basSBL_Citation_EBNF` — call `ParseReference("Genesis 1:1")`.

## Architecture

### Source of Truth: the `.DOCM` file

All macro logic lives in the live `.DOCM` document. The `\src` directory is an export mirror for Git tracking only — code is pushed there via `aeWordGitClass`. Never edit `\src` files as a substitute for editing the `.DOCM`.

Deliverables are `.DOCX` files produced via Word's "Save As" — this separates output from tooling.

### Module Structure

| Module | Role |
| -------- | ------ |
| `aeBibleClass.cls` | Main test class — 40+ diagnostic tests (layout, style, fonts, markers) |
| `aeWordGitClass.cls` | Exports VBA to `\src`, Git integration, changelog management |
| `aeRibbonClass.cls` | Ribbon event handler class |
| `basSBL_Citation_EBNF.bas` | 14-stage deterministic Scripture reference parser |
| `basSBL_TestHarness.bas` | Parser test suite (alias coverage, tokenization, semantic flow) |
| `basSBL_TestFramework.bas` | Assertion library: `AssertTrue`, `AssertFalse`, `AssertEqual` |
| `basSBL_VerseCountsGenerator.bas` | Generates canonical verse-count tables |
| `basWordRepairRunner.bas` | Automated layout repairs (verse markers, styles, wrapping) |
| `basUSFM_Export.bas` | Exports Bible content to USFM format with audit logging |
| `basTEST_aeBibleTools.bas` | Document audits: color, font, style, empty paragraphs, sections |
| `basTEST_aeBibleFonts.bas` | Font diagnostic tests |
| `basTest_aeBibleClass.bas` | Entry point — `RUN_THE_TESTS()` orchestrator |
| `basAuditDocument.bas` | Document-level audit routines |
| `basFixDocxRoutines.bas` | FIX routines: headers, footers, and other document repairs |
| `basBibleRibbonSetup.bas` | Custom Word Ribbon implementation |
| `basImportWordGitFiles.bas` | Imports VBA source files back into the `.DOCM` |
| `basChangeLog_aeBibleClass.bas` | Issue tracker and changelog (numbered #NNN) |
| `basChangeLog_aeWordGitClass.bas` | Changelog for Git integration module |
| `basWordSettingsDiagnostic.bas` | Word application settings diagnostics |
| `bas_TODO.bas` | Future stages and open feature ideas |
| `Module1.bas` | Utility functions: font printing, book navigation, character analysis |
| `ThisDocument.cls` | Document-level event handlers |

Files prefixed with `X` (e.g., `XbasTESTaeBibleClass_SLOW.bas`) are long-running or deferred — not part of the normal test run.

### Scripture Parser (14 Stages)

Stages in `basSBL_Citation_EBNF.bas`:

1. Lexical Scan — tokenize input
2. Resolve Alias — map book names/abbreviations
3. Interpret Structure — parse chapter:verse
4. Validate Canonical — bounds checking
5. Rewrite Single-Chapter — handle Jude, Obadiah, Philemon, etc.
6. Compose Canonical — format output
7. Final Parser — verification
8-12. Range/List extensions — `Gen 1:1-2:3`, comma-separated lists
13-14. Contextual shorthand and compression

Stages 15+ (Span Normalization, Canonical Ordering, Verse Expansion) are planned in `bas_TODO.bas`.

### Issue Numbering Convention

Changes are tracked in `basChangeLog_aeBibleClass.bas` with sequential `#NNN` issue numbers. Task categories: `[doc] [test] [bug] [perf] [audit] [disc] [feat] [idea] [impr] [flow] [code] [wip] [clean] [obso] [regr] [refac] [opt]`.

Commit messages follow: `FIXED - #NNN - Description [category]`

### Audit Outputs (`\rpt` directory)

- `HeadingLog.txt` — Heading 1 paragraph index map
- `HeadingIndex.txt` — Heading structure
- `HeaderFooterAudit.txt` — Header/footer analysis
- `ExportedBible.usfm` — USFM export
- `RepairLog.txt` — Log of automated repairs
- `TestReport.txt` — Test run results
- `Style Usage Distribution.txt` — Style frequency report
- `USFM_Export_Log.txt` / `USFM_Validator_Log.txt` — USFM pipeline logs

### Documentation (`\md` directory)

- `Editorial Design and Style Guide.md` — EDSG rules, audit architecture, module manifest
- `Bias Guard.md` — Copilot suggestion filtering
- `Compact Strategy for Squashed Audit Commits.md` — Git workflow
- `Efficient Book-Chapter Navigation.md` — Navigation patterns
- `FIXED_AuditLog.md` — Resolved issue log

## Code Standards

- All modules use `Option Explicit` and `Option Compare Text`
- Error handling: `On Error GoTo PROC_ERR` with `PROC_EXIT` label pattern
- ASCII-only output to Immediate Window
- No content modification may bypass audit review — all repairs must be reversible and logged
- LGPL 3.0 licensed; all Class modules include copyright headers
