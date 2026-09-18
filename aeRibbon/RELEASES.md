# aeRibbon Releases

Append-only ledger mapping each shipped `.dotm` back to its dev-source
SHA. See `md/aeProductionRibbonPlan.md` §6 for the versioning scheme.

| Version | Date | Dev SHA | Source tree | Gates passed | Notes |
|---|---|---|---|---|---|
| `1.0.0+9689917` | 2026-09-18 | `9689917` | `aeRibbon/src/` @ `9689917` | G7: **pass**. G8: **pass with known non-blocking bugs** (2 found -- see `aeRibbon/releases/1.0.0+9689917/BUILD_RECORD.txt`: Next-Chapter-after-Next-Book confirmed real, KeyTip focus-escape at boundaries won't-fix). G1–G6: not run (not a full v1.0.0 sign-off). | First release row. Verifies the About-dialog `RIBBON_VERSION` fix and the new `basImportWordRibbonGitFiles.bas` build-bootstrap automation (first live run, passed), in service of the `aeBibleAddin` docm/docx/JS-taskpane release-alignment procedure -- same pattern as `1.0.0+bc71416`'s own G8 subsection that fed `aeBibleAddin v0.1.0`. |
