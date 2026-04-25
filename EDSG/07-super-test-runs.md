# 07 — SUPER_TEST_RUNS — architectural QA supervisor

Status: **placeholder**. The routine is designed (see § 6 of
`rvw/Code_review 2026-04-21.md`) but **deferred** until the style
taxonomy is stable. This page expands when implementation begins.

## Concept

A single entry point that runs every QA suite in sequence and produces
a master report. One command before any commit; one report to read
before any release. The architectural supervisor of the QA process.

## Suite sequence (planned)

| Order | Suite | Entry point | Output |
|---|---|---|---|
| 1 | Style taxonomy | `RUN_TAXONOMY_STYLES` | `rpt\StyleTaxonomyAudit.txt` |
| 2 | Document diagnostics | `RUN_THE_TESTS` | `rpt\TestReport.txt` |
| 3 | Font audit | (existing font test routines) | `rpt\FontAudit.txt` |
| 4 | Header / footer audit | `SummarizeHeaderFooterAuditToFile` | `rpt\HeaderFooterAudit.txt` |
| 5 | Scripture parser | `basSBL_TestHarness` entry point | (currently Immediate; needs file output before inclusion) |

Master report at `rpt\SuperTestReport.txt`: timestamp, one summary
line per suite (`PASS / FAIL / count`), then links to individual
report files for drill-down.

## Decisions already taken

- Name: `SUPER_TEST_RUNS` (caps, consistent with `RUN_THE_TESTS`).
- Location: new module `basVerificationSuite.bas` — orchestration
  separate from individual test logic; scales as suites are added.
- Each suite call wrapped in `On Error Resume Next` per-suite — one
  crashing suite must not silence the rest.
- A `Quick` mode flag will skip slow suites (marked `X` prefix
  convention already in use) for pre-commit checks.
- `SuperTestReport.txt` accumulates with timestamps (append, not
  overwrite) so trend analysis is possible.

## Connection to EDSG

When implemented, this page documents:

- What each suite validates.
- How to read the master report.
- Failure triage: which suite failed → which EDSG page explains the
  expected behavior.
- Release gating: a clean run is the pre-publication gate for any
  build (English or localized).

Each style and process page in this guide will link forward to the
SUPER_TEST_RUNS suite that validates it. Currently the back-links
read "validated by Suite N (pending)".

## Why deferred

`RUN_THE_TESTS` alone is several minutes; the full suite may be
10–20. Premature integration before the underlying suites stabilize
would mean rework. Implementation begins after:

- Style taxonomy fully walked (pages 12+ pending).
- Decisions resolved on `Normal`, `BodyTextIndent`, `AuthorQuote`.
- Font audit suite formalized and documented.
- Scripture parser harness updated to write to a file (currently
  Immediate-only; can't aggregate).

Once those land, this placeholder converts to operational
documentation.
