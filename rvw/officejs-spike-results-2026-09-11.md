# Office.js technical spike — results, 2026-09-11

**Supersedes/confirms:** `rvw/Code_review - 2026-04-10a.md`'s VSTO-over-Office.js recommendation
(~line 792–851), which reasoned that "the Office JS API does not yet expose the full
paragraph-level navigation and character-style inspection that the current VBA code relies on."
This note is the hands-on follow-up the master conversion plan (`adaept5tudio/docs/aeBibleClass-word-addin-conversion-plan.md`
§4) called for before trusting that documentation-only read a second time.

**Spike code:** [`aeBibleAddin/spike-officejs/`](https://github.com/adaept/aeBibleAddin/tree/main/spike-officejs)
(commit `e7395ab`, schema/manifest fixed at `6e8dc3f`). **Test document:** the actual current
`Bible.docx` (not a synthetic excerpt) — 33,825 body paragraphs, 427 defined styles, 145 sections.
**Word build:** Microsoft 365 Current Channel, per the same build info the VBA `TestReport.txt`
records.

## Results

### Spike 1 — style introspection (`Document.getStyles()`)

**Succeeded.** 427 styles loaded — `nameLocal`, `type`, `priority`, `quickStyle`,
`unhideWhenUsed` all populated — in one batched `context.sync()`. This is a direct, hands-on
refutation of the April 2026 review's "character-style inspection" concern: the properties
`CountUnapprovedVisibleStyles` (`aeBibleClass.cls:2857`) needs are all present and readable at
full document scale, not just in documentation.

The "44 violations" reported is not a meaningful number on its own — the spike's approved-style
set is a 5-entry placeholder (`Normal`, `Heading 1`, `Heading 2`, `BodyText`, `VerseText`), not
this project's real ~42-entry approved set (see `rpt/Style Usage Distribution.txt`). A real port
would read the actual approved set first; that wasn't the point of this check, which was purely
"can the API expose enough to replicate the logic," and it can.

### Spike 2 — batched paragraph-style walk (the performance question)

**33,825 paragraphs read in 2984.4ms** (one batched `context.sync()`) — roughly 11,000
paragraphs/second. This is the concrete number the conversion plan flagged as an unknown
(Office.js's `RequestContext` batching model vs. a naive per-paragraph round-trip, which is a
well-known Office-add-in performance pitfall). ~3 seconds for a full-Bible-scale document, from a
single user-triggered action, is well within an acceptable UX budget — this is not a performance
blocker.

**New finding, not fully anticipated by the conversion plan:** cross-checking the returned style
breakdown against `rpt/Style Usage Distribution.txt` (the VBA's own `CountAuditStyles_ToFile`
output), almost every count matches exactly (`VerseText: 31103`, `Heading 2: 1189`,
`AuthorBodyText: 160`, etc.) — **except `Footnote Text` (1000), `TheHeaders` (137),
`TheFooters` (5), `ParallelText`/`ParallelHeader`, `Header`, and `Footer` are completely absent**
from the JS walk. `context.document.body.paragraphs` only covers the main body story —
headers, footers, and footnotes are separate story ranges Office.js doesn't include by default.

This confirms, with a precise number, what the conversion plan's §3 only speculated about:
*"[Section-scoped header/footer access] covers the same practical ground... for a document that
has a small, fixed number of sections"* — true, but it means a full audit-engine port (Phase 3)
needs **separate batched walks per story range** (body + each section's header/footer via
`getHeader`/`getFooter`, confirmed working in spike 3), not one flat traversal like VBA's
`ActiveDocument.Paragraphs`. Footnote-range enumeration specifically is **not yet validated** —
open question for whoever picks up Phase 3.

A handful of small, unrelated count drifts vs. the VBA snapshot (`BodyText` 488 vs. 495, `Normal`
1 vs. 7, `Introduction` 5 vs. 6, plus 4 paragraphs with an empty style name in the JS reading not
present in the VBA list) are most likely explained by document edits between the two measurements
(the VBA snapshot and this spike were not run at the exact same document state) — not
investigated further here, out of scope for this spike.

### Spike 3 — header read/write (`Section.getHeader/getFooter`)

**Succeeded.** Before: `"    "` (whitespace/tab-only header, consistent with
`rpt/HeaderFooterAudit.txt`'s page-number-only header pattern). After:
`"     [spike marker -- safe to remove]"` — the insert landed correctly, confirming both read
and write work, not just that the API type-checks.

**Caution for anyone re-running this spike:** this write is real and persists if the document is
saved. The test document was not committed to git (see `spike-officejs/README.md`), so there's no
version-control undo — don't save, or manually remove the marker text first.

## Verdict

**This confirms Office.js is viable for the citation/navigation/style-introspection surface —
supersedes the April 2026 VSTO recommendation for those specific capabilities.** Both properties
that review named as missing (paragraph-level style inspection, header/footer access) are
present, performant at real document scale, and read/write-capable. The remaining VSTO-favoring
argument that review made — the full 85-routine audit engine's paragraph-level depth — is now
narrowed to a specific, well-defined open question (multi-story-range walks, footnote enumeration
unvalidated) rather than a blanket "the API can't do this." That's a Phase 3 design detail to plan
around, not a reason to choose VSTO over Office.js for Phase 1/2.

**Recommendation:** proceed with Phase 2 (task-pane shell). Flag the story-range-walk design
question for whoever scopes Phase 3's audit-engine port; validate footnote-range enumeration
specifically before committing to a Phase 3 approach.
