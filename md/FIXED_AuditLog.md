# Audits for Commit Log

## [#294] Cut a 0.1.1 release and tag it on GitHub

- **Fixed in:** [`b6f95f5`](https://github.com/adaept/aeBibleClass/commit/b6f95f5a1d2a8c2428f8a69146ef5566fb5c14b2)
- **Files:**
  - [`aeBibleClass.cls`](https://github.com/adaept/aeBibleClass/blob/b6f95f5/src/aeBibleClass.cls)
  - [`basChangeLogaeBibleClass.bas`](https://github.com/adaept/aeBibleClass/blob/b6f95f5/src/basChangeLogaeBibleClass.bas)
- **Summary:** Tagged release `v0.1.1` and updated changelog to reflect finalized fixes and documentation.
- **Audit Result:** ✅ `[OK] Release cut and changelog updated`

## [#300] Add md doc to outline a Compact Strategy for Squashed Audit Commits and reduce GitHub commit log spam

- **Fixed in:** [`7af3130`](https://github.com/adaept/aeBibleClass/commit/7af31303b2dae7204f23d6e0d5916307eaf38b05)
- **File:** `md/Compact Audit Strategy.md`
- **Line:** [L1](https://github.com/adaept/aeBibleClass/blob/7af31303/src/md/Compact%20Audit%20Strategy.md#L1)
- **Summary:** Introduced strategy document for squashed audit commits to reduce GitHub noise and improve changelog clarity.
- **Audit Result:** ✅ `[OK] Strategy doc added and changelog updated`

## [#279] Add routine to define H2 style and reapply it in the project

- **Fixed in:** [`7af3130`](https://github.com/adaept/aeBibleClass/commit/7af31303b2dae7204f23d6e0d5916307eaf38b05)
- **File:** `basTESTaeBibleTools.bas`
- **Line:** [L1](https://github.com/adaept/aeBibleClass/blob/7af31303/src/basTESTaeBibleTools.bas#L1)
- **Summary:** Added macro to define and reapply H2 style across project documents; includes header and [impr] tag.
- **Audit Result:** ✅ `[OK] Style macro added and changelog updated`

## [#306] Add audit log from squash #274

- **Fixed in:** [`7af3130`](https://github.com/adaept/aeBibleClass/commit/7af31303b2dae7204f23d6e0d5916307eaf38b05)
- **File:** `md/FIXED_AuditLog.md`
- **Line:** [L1](https://github.com/adaept/aeBibleClass/blob/7af31303/src/md/FIXED_AuditLog.md#L1)
- **Summary:** Inserted audit entry for previously squashed fix #274 to preserve traceability.
- **Audit Result:** ✅ `[OK] Squash audit entry added`

## [#307] Remove bGoTo16, not needed with use of run single test

- **Fixed in:** [`7af3130`](https://github.com/adaept/aeBibleClass/commit/7af31303b2dae7204f23d6e0d5916307eaf38b05)
- **File:** `basTESTaeBibleTools.bas`
- **Line:** [L42](https://github.com/adaept/aeBibleClass/blob/7af31303/src/basTESTaeBibleTools.bas#L42)
- **Summary:** Removed obsolete flag `bGoTo16`; logic now handled by single test runner.
- **Audit Result:** ✅ `[OK] Obsolete flag removed cleanly`

## [#308] Update all use of TestReportFlag to - If TestReportFlag And OneTest = 0

- **Fixed in:** [`7af3130`](https://github.com/adaept/aeBibleClass/commit/7af31303b2dae7204f23d6e0d5916307eaf38b05)
- **File:** `basTESTaeBibleTools.bas`
- **Line:** [L55](https://github.com/adaept/aeBibleClass/blob/7af31303/src/basTESTaeBibleTools.bas#L55)
- **Summary:** Refactored conditional logic to suppress report output during single test runs.
- **Audit Result:** ✅ `[OK] TestReportFlag logic updated`

## [#303] Fix single RUN_THE_TESTS(x) so it does not run AppendToFile and kill the full report

- **Fixed in:** [`7af3130`](https://github.com/adaept/aeBibleClass/commit/7af31303b2dae7204f23d6e0d5916307eaf38b05)
- **File:** `basTESTaeBibleTools.bas`
- **Line:** [L61](https://github.com/adaept/aeBibleClass/blob/7af31303/src/basTESTaeBibleTools.bas#L61)
- **Summary:** Corrected logic to prevent full report overwrite when running a single test.
- **Audit Result:** ✅ `[OK] Single test logic corrected`

## [#309] Add code to scan modules in .docm to flag early-bound object declarations

- **Fixed in:** [`7af3130`](https://github.com/adaept/aeBibleClass/commit/7af31303b2dae7204f23d6e0d5916307eaf38b05)
- **File:** `basTESTaeBibleTools.bas`
- **Line:** [L75](https://github.com/adaept/aeBibleClass/blob/7af31303/src/basTESTaeBibleTools.bas#L75)
- **Summary:** Added macro to scan `.docm` modules for early-bound declarations to support audit and refactor.
- **Audit Result:** ✅ `[OK] Early-bound scan macro added`

## [#301] AppendToFile should be "SKIPPED" [bug]

- **Fixed in:** [`47faa14`](https://github.com/adaept/aeBibleClass/commit/47faa142c479a485167755cef65eb87290399504)
- **File:** `basChangeLogaeBibleClass.bas`
- **Line:** [L185](https://github.com/adaept/aeBibleClass/blob/fcc07412eddc3c3498affa5c0955c1a3db0a9779/src/basChangeLogaeBibleClass.bas#L185)
- **Summary:** Corrected mislabeling of AppendToFile output; SKIPPED now outputs cleanly during session when required.
- **Audit Result:** ✅ `[OK] Logic fix verified via commit message and diagnostic trail`

## [#302] Move PrintCompactSectionLayoutInfo to basTESTaeBibleTools and update output path

- **Fixed in:** [`47faa14`](https://github.com/adaept/aeBibleClass/commit/47faa142c479a485167755cef65eb87290399504)
- **File:** `basTESTaeBibleTools.bas`
- **Line:** [L1](https://github.com/adaept/aeBibleClass/blob/fcc07412eddc3c3498affa5c0955c1a3db0a9779/src/basTESTaeBibleTools.bas#L1)
- **Summary:** Relocated macro for compact section layout reporting; updated output path to `rpt` folder and added header for audit context.
- **Audit Result:** ✅ `[OK] Macro relocation and output path audit confirmed`

## [#304] Add task type [wip] for pre-resolution changelog tagging

- **Fixed in:** [`47faa14`](https://github.com/adaept/aeBibleClass/commit/47faa142c479a485167755cef65eb87290399504)
- **File:** `basChangeLogaeBibleClass.bas`
- **Line:** [L187](https://github.com/adaept/aeBibleClass/blob/fcc07412eddc3c3498affa5c0955c1a3db0a9779/src/basChangeLogaeBibleClass.bas#L187)
- **Summary:** Introduced `[wip]` task type for early commit tagging—enables staging of partial fixes without audit disruption.
- **Audit Result:** ✅ `[OK] Task type logic operational and format-compatible`

## [#274] Fix output path so 'Style Usage Distribution.txt' goes to rpt folder, add code header

- **Fixed in:** [`47faa14`](https://github.com/adaept/aeBibleClass/commit/47faa142c479a485167755cef65eb87290399504)
- **File:** `basChangeLogaeBibleClass.bas`
- **Line:** [L12](https://github.com/adaept/aeBibleClass/blob/fcc07412eddc3c3498affa5c0955c1a3db0a9779/src/basChangeLogaeBibleClass.bas#L12)
- **Summary:** Updated output path logic and added documentation header for style usage report.
- **Audit Result:** ✅ `[OK] #274 found within module block`

## [#299] Final validator update and audit format clarification

- **Fixed in:** [`ff2aa10`](https://github.com/adaept/aeBibleClass/commit/ff2aa102a1aabcd00f330c6475693527ff79c200)
- **File:** `md/Editorial Design and Style Guide.md`
- **Line:** [L503](https://github.com/adaept/aeBibleClass/commit/ff2aa102a1aabcd00f330c6475693527ff79c200#diff-0ef90a4f6297d8bd3147bdae1da9222de0a0a1b9ea1f82a86f0178671d227029R503)
- **Summary:** Added formal specification of Markdown audit log format, clarifying UTF-8 compatibility and distinction from ASCII-only macro diagnostics.
- **Audit Result:** ✅ `[OK] Markdown audit format spec added and aligned with validator logic`

- **Fixed in:** [`53f1ade`](https://github.com/adaept/aeBibleClass/commit/53f1ade0e531a31021ea794ce1aa1f6f9fcfa96e)
- **File:** `md/FIXED_AuditLog.md`
- **Summary:** Regenerated audit entry for #299 with UTF-8 Markdown formatting. Clarified audit log format specification for future reference.
- **Audit Result:** ✅ `[OK] Markdown audit entry and format spec updated`

## [#299] Add initial README and Bias Guard md files

- **Fixed in:**
  - [`39a147b`](https://github.com/adaept/aeBibleClass/commit/39a147bd719b01da113dc4d367cf1dec1d319b96)
  - [`7e84cde`](https://github.com/adaept/aeBibleClass/commit/7e84cdee435d95084af20518ef3b76b5240633fb)
- **File:** `basChangeLogaeBibleClass.bas`
- **Line:** [L12](https://github.com/adaept/aeBibleClass/blob/main/src/basChangeLogaeBibleClass.bas#L12)
- **Summary:** Task #299 added to changelog block. Markdown documentation introduced:
  - `README.md`: Project scope, platform notes, and guiding principles
  - `Bias Guard.md`: Audit alignment rationale and integration notes
- **Audit Result:** ✅ `[OK] #299 found within module block`

## [#298] Use SSOT with Select Case statements for values such as num and verify with RUN_THE_TESTS

- **Fixed in:**
  - [`026b45f`](https://github.com/adaept/aeBibleClass/commit/026b45f0cc180ed0de5733240264b368bcc654eb)
  - [`1f712f0`](https://github.com/adaept/aeBibleClass/commit/1f712f01ff7bcdb504ba2e906e8e5244a834ad03)
- **File:** `basChangeLogaeBibleClass.bas`
- **Line:** [L12](https://github.com/adaept/aeBibleClass/blob/fcc07412eddc3c3498affa5c0955c1a3db0a9779/src/basChangeLogaeBibleClass.bas#L12)
- **Summary:** Refactored changelog entry and implemented SSOT logic with Select Case validation.
- **Audit Result:** ✅ `[OK] #298 found within module block`

## [#297] Create file to hold Audits for Commit Log

- **Fixed in:**
  - [`da0f97e`](https://github.com/adaept/aeBibleClass/commit/da0f97ee6a62defe528eb3fb6dc4fe27680fa830)
  - [`29aaa7f`](https://github.com/adaept/aeBibleClass/commit/29aaa7ffb0689e1038cb2d1c0014a980e4cc8af2)
- **File:** `basChangeLogaeBibleClass.bas`
- **Line:** [L14](https://github.com/adaept/aeBibleClass/blob/fcc07412eddc3c3498affa5c0955c1a3db0a9779/src/basChangeLogaeBibleClass.bas#L14)
- **Summary:** Added FIXED_AuditLog.md and updated changelog to include task #297.
- **Audit Result:** ✅ `[OK] #297 found within module block`

## [#296] ValidateTaskInChangelogModule

- **Fixed in:** [`a3f62b8`](https://github.com/adaept/aeBibleClass/commit/a3f62b85c8106efaf5bbfa5d07824474e23f1f82)
- **File:** `basChangeLogaeBibleClass.bas`
- **Line:** [L14](https://github.com/adaept/aeBibleClass/blob/fcc07412eddc3c3498affa5c0955c1a3db0a9779/src/basChangeLogaeBibleClass.bas#L14)
- **Summary:** Added validation macro to confirm task tags appear within changelog blocks.
- **Audit Result:** ✅ `[OK] #296 found within module block`
