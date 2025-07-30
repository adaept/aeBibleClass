# Audits for Commit Log

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
- **Audit Result:** `[OK] #298 found within module block`

## [#297] Create file to hold Audits for Commit Log

- **Fixed in:**
  - [`da0f97e`](https://github.com/adaept/aeBibleClass/commit/da0f97ee6a62defe528eb3fb6dc4fe27680fa830)
  - [`29aaa7f`](https://github.com/adaept/aeBibleClass/commit/29aaa7ffb0689e1038cb2d1c0014a980e4cc8af2)
- **File:** `basChangeLogaeBibleClass.bas`
- **Line:** [L14](https://github.com/adaept/aeBibleClass/blob/fcc07412eddc3c3498affa5c0955c1a3db0a9779/src/basChangeLogaeBibleClass.bas#L14)
- **Summary:** Added FIXED_AuditLog.md and updated changelog to include task #297.
- **Audit Result:** `[OK] #297 found within module block`

## [#296] ValidateTaskInChangelogModule

- **Fixed in:** [`a3f62b8`](https://github.com/adaept/aeBibleClass/commit/a3f62b85c8106efaf5bbfa5d07824474e23f1f82)
- **File:** `basChangeLogaeBibleClass.bas`
- **Line:** [L14](https://github.com/adaept/aeBibleClass/blob/fcc07412eddc3c3498affa5c0955c1a3db0a9779/src/basChangeLogaeBibleClass.bas#L14)
- **Summary:** Added validation macro to confirm task tags appear within changelog blocks.
- **Audit Result:** `[OK] #296 found within module block`
