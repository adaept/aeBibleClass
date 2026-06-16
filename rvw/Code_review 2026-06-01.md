# Code review - 2026-06-01 carry-forward

This file opens a fresh review arc on 2026-06-01. The previous arc
[`rvw/Code_review 2026-05-28.md`](Code_review%202026-05-28.md) is now
**closed for new work**; that file remains the authoritative dated
history for everything between 2026-05-28 and 2026-06-01, including:

- **Header/Footer style audit arc (2026-06-01).** Built
  `AuditHeaderFooterStyles` + `GetHeaderFooterStyleTotals` in
  `basVerseStructureAudit.bas`; rewrote them onto a `StoryRanges` /
  `NextStoryRange` walk so they enumerate orphaned first-page/even-page
  stories (same basis as the Style Usage Distribution). Live run found
  7 violations (2 the distribution masked under `Normal`); operator
  remediated the front-matter orphans; audit reconciled with the
  distribution at 0. **Test slot 84 activated - PASS at 0.**
- **Backup script (2026-06-01).** `synch-to-onedrive.bat` now snapshots
  the Claude project settings (memory + global `settings.json`) into
  `_claude_backup\` before rsync; `.gitignore` updated.
- **Tests 81-83 added (2026-05-31)** - character-style coverage
  (`CountAuditCharacterStyles_ToFile`, `CountVerseMarker`,
  `CountChapterVerseMarker`); surfaced a +1 CVM anomaly.
- **Test 80 + Test 11 / 33 / 38 work (2026-05-30)** - bare-empty-para
  split and hint passes.
- **2026-05-28 session** - Introduction SpaceBefore 0 -> 12; Default
  Paragraph Font promoted to approved.

Status tag legend (continued):

- **OPEN** - actively pending, all known prerequisites met.
- **PARTIAL** - partially complete; specific remaining work listed.
- **DEFERRED** - not started, waiting on a specific trigger.
- **FUTURE** - speculative; revisit only when conditions warrant.
- **RECOVERED** - surfaced from a prior arc where it was dropped off
  the carry-forward chain.

## Open carry-forward (priority order)

### 1. Run aeRibbon Gates G1-G8 and ship v1.0.0 (HIGH) - PREP DONE 2026-05-17, READY FOR BUILD

Still the next active release-track item. Build-side prep (trim,
BUILD.md correction, small-refresh classification) remains valid. G8
spot-check on the new book aliases (`JSH`, `JDG`, etc.) still queued.

**Operator action** (Word GUI; not driveable from here):

1. Build `aeRibbon/template/aeRibbon.dotm` per `aeRibbon/BUILD.md`
   steps 1-8.
2. Produce the production Bible `.docx` per BUILD.md.
3. Run Gates G1-G8 from `aeRibbon/QA_CHECKLIST.md`; record results in
   `aeRibbon/releases/1.0.0+bc71416/BUILD_RECORD.txt`.
4. Append a row to `aeRibbon/RELEASES.md` and `git tag v1.0.0+bc71416`.

Full prior analysis and gate definitions: § 1 in
[`Code_review 2026-05-28.md`](Code_review%202026-05-28.md) ->
[`Code_review 2026-05-16.md`](Code_review%202026-05-16.md).

### 2. AuditCharStyleUsage quadratic-time fix (HIGH) - OPEN 2026-05-31

`basVerseStructureAudit.AuditCharStyleUsage` walks character-style runs
via document-scope `Range.Find`, then calls `oRng.Paragraphs(1)` after
every match - an `O(position-in-doc)` lookup, so the scan is `O(N^2)`.
Observed: CVM run (31,103 matches) 406 s; VM run aborted (>60 s per 1k
deep in).

**Fix direction:** rewrite to walk `ActiveDocument.Paragraphs` once via
`For Each` with a paragraph-scoped `Range.Find` (or `Characters(1)` /
`Characters(Last)` for the common START/END cases) - the same bounded-
per-paragraph shape now proven in `GetHeaderFooterStyleTotals`,
`GetMarkerTotals`, and `AuditOrphanBodyTextParagraphs`. Eliminates the
`oRng.Paragraphs(1)` lookup. Full detail: § 10 in
[`Code_review 2026-05-28.md`](Code_review%202026-05-28.md).

### 3. Header/Footer + Section SHAPE LOCKDOWN (MEDIUM) - OPEN 2026-06-01

Standing structural-integrity task, distinct from the per-paragraph
style gate (which slot 84 now covers). The document is deliberately
**opinionated and single-shape**:

- **No unused / invalid sections.** First and last sections are
  intentional placeholder boundaries (J5 layout vs Word's default
  "Document 1"); everything between is the 66-book body.
- **No orphaned header/footer stories.** The 2026-06-01 remediation
  cleared the front-matter orphans; even-page stories were dropped and
  the slot-84 gate now holds the per-paragraph style at 0. A shape gate
  would additionally assert zero *orphaned* (compliant-but-unused)
  stories.
- **Sec 1 "Different First Page" is now intended shape.** Remediation
  left it ON (front matter has a distinct, compliant first page). A
  future shape test must assert this as expected, not flag it as drift.
- **`KeepWithNext` / `KeepTogether` locked down tightly** - pagination
  discipline is part of the shape contract; a future test should assert
  the intended keep-with-next map rather than let it drift.
- **i18n must NOT change the document shape.** Localization appends
  content within the existing shape; it does not add sections, headers,
  footers, or page-break structure.
- **Known distant-horizon exception:** bidirectional / RTL localization
  could need to bend the strict single-shape rule. Far out, explicitly
  out of scope now; flagged so shape-lockdown tests do not bake in
  assumptions that make RTL impossible later.

Candidate gates (once specified): zero orphaned HF stories; section
count equals the canonical expected (145 today, or a derived value);
placeholder boundary sections present and correct; Sec 1
Different-First-Page present; `KeepWithNext` matches the intended map.
Full origin analysis: the 2026-06-01 reconciliation entries in
[`Code_review 2026-05-28.md`](Code_review%202026-05-28.md).

### 4. Revisit failed tests and verify status / code / performance (MEDIUM) - OPEN

Carry-forward from 2026-05-16 § 12. Trigger: next time a slot FAILs,
walk the function before rebaselining (Test 22-style split candidates).
Known candidates: Test 30 source-comment vs expected mismatch
(`CountHeaderStyleUsage` says "Expected = 0" but baseline 70);
count-baseline tests 24, 27, 29, 30, 32-35, 37, 47, 49, 50, 51; slow
slots (apply the Test 22 perf lens). Full list: § 2 in
[`Code_review 2026-05-28.md`](Code_review%202026-05-28.md).

### 5. Date-rule sweep follow-ups (MEDIUM) - OPEN

Carry-forward from 2026-05-19 / 2026-05-20:

- **Pair 06 in `Date_Example.txt` still `pending`** - operator to
  decide the 300-600 AD range; rerun `ApplyDateRule_2026_05_19`.
- **Apply the date rule to the 20 example passages** in the live doc;
  rerun `Test_NoSuperscriptOrdinals` and Test 79.
- **Book-number ordinal policy still DEFERRED** (`1st Samuel` vs
  `1 Samuel`). Target: Test 79 = 0; `Test_NoSuperscriptOrdinals` = 0
  once the book-number policy is decided and applied.

Full detail: § 3 in
[`Code_review 2026-05-28.md`](Code_review%202026-05-28.md).

### 6. File-write code audit against FSO rule (MEDIUM) - OPEN 2026-05-31

Per [[feedback_fso_file_writes]]: `rpt/` writers in re-callable
routines must use `FSO.CreateTextFile`, not `Open ... For Output As`,
to avoid Err 55/70. A full sweep has not been done.

**Action:** grep all `.bas`/`.cls` for `Open ` + `Output As` /
`Append As`; for every re-callable writer targeting `rpt/`, convert to
the FSO pattern (`CountAuditStyles_ToFile`,
`WriteHeaderFooterStyleFile`, etc.). Leave genuine one-shot writers.
Full detail: § 11 in
[`Code_review 2026-05-28.md`](Code_review%202026-05-28.md).

### 7. +1 CVM anomaly at ParaStart=3087864 (editorial) - OPEN 2026-05-31

A single stray "Chapter Verse marker"-styled character at the tail of
one VerseText paragraph (offset 146 of 147) makes the Find-based CVM
total (31,103) differ from the paragraph-rule count (31,102) by one.
Paragraph is structurally correct.

**Operator:** `GoToPos 3087864`, inspect the end-of-paragraph
character; decide clean-up (remove the styling) or accept-and-rebaseline
a future "total VM/CVM runs in doc" diagnostic to 31,103. VM anomaly
partner unconfirmed; will likely surface at the same `ParaStart` once
the § 2 quadratic fix lands and the VM scan can complete.

### 8. EDSG `10-list-paragraph-bug.md` Step 0 snippet correction (LOW) - OPEN 2026-05-28

The Step 0 diagnostic snippet uses `Not (s.LinkToListTemplate Is
Nothing)` as a read-side check, but `Style.LinkToListTemplate` is
**write-only**. Update the snippet to mirror the Test 75
`ListTemplates -> ListLevels -> LinkedStyle` traversal. Conceptual
framing stands; only the code line needs replacement. Full detail: § 4
in [`Code_review 2026-05-28.md`](Code_review%202026-05-28.md).

### 9. Test 38 kind-distribution + structural-phrasing follow-ups (LOW) - OPEN 2026-05-30

- Decide whether `kind`-distribution drift in `rpt\EmptyParagraphs.txt`
  (e.g. SBNP vs SBC ratio) should itself be a tested invariant;
  currently only the total count is gated (Test 38 / Test 80).
- Consider promoting Test 38's structural-reality phrasing into
  `EDSG/01-styles.md` or `EDSG/04-qa-workflow.md` if operators
  reference test descriptions in QA narrative.

### 10. Normal style audit (LOW, DEFERRED)

`Normal` is intentionally unaudited as the "pin-everything-else-above"
anchor. **Watch:** the 2026-05-28 `DumpAllApprovedStyles` showed
`Normal.QuickStyle` flipped True -> False; audit does not check
QuickStyle (no spec change), but bump priority if a second drift
surfaces. Full detail: § 8 in
[`Code_review 2026-05-28.md`](Code_review%202026-05-28.md).

### 11. Finding 5 (ribbon nav) - umbrella OPEN (DEFERRED, WORD LIMITATION) - RECOVERED

Word-side limitation; no action available. Remains in the register for
awareness.

### 12. SoftHyphenSweep_FootnotesOnly sister routine (DEFERRED)

Surfaced during the 2026-05-08 SHA build; waiting on a footnote-specific
trigger before implementation.

### 13. SHA_ReplaceHard i18n consideration (FUTURE)

Speculative; revisit when a non-English target translation materialises.
Related to the i18n-shape-invariance note in § 3.

### 14. Architecture rule - class encapsulation + module/class safety boundary (RULE, 2026-05-15)

Standing rule (codified as [[feedback_class_encapsulation]]), not an
action item - listed so it stays visible during slot-by-slot work.
Full rule + worked examples: § 9 in
[`Code_review 2026-05-28.md`](Code_review%202026-05-28.md).

### 15. aeRWB source-text repo relationship (RULE/NOTE, 2026-06-16)

Recorded for cross-project awareness; the linkage details are still
**forthcoming from the operator** so the specifics below are the known
state, not a closed spec.

- **What aeRWB is.** `C:\adaept\aeRWB` is the upstream **source-text**
  repository for the Radiant Word Bible (RWB). It holds `web.txt` (the
  original World English Bible from openbible.com) and `rwb.txt` (the
  Radiant / refined English text). aeBibleClass is the consumer side:
  the Word Study-Bible `.docx`/`.docm` + the navigation ribbon
  (`LBL_TAB = "Radiant Word Bible"`) built *from* that text.
- **Shared identity.** aeRWB was renamed "Refined" -> "Radiant"
  (2026-06-16), so both repos now share the **Radiant Word Bible / RWB**
  naming and trademark (see `md\Ribbon Design.md` § trademark).
- **Supervisor onboarding (2026-06-16).** aeRWB is now a supervised
  project under the `adaept5tudio` hub - already option 2
  (`FOLDER_2=aeRWB`) in this repo's `synch-to-onedrive.bat`, and it now
  carries a committed `.claude\settings.json` (deny reading session
  `.jsonl`), a `.gitignore` (excludes `_claude_backup\`), and its own
  `rvw\Code_review 2026-06-16.md` review lineage.
- **i18n tie-in (source-text side of items 3 + 13).** aeRWB's text
  files are **UTF-8 with BOM and must stay UTF-8** for future i18n /
  translation work. This is the upstream half of the same i18n
  discipline the document tracks downstream: item 3 keeps localization
  *within the existing document shape*; the source text must stay
  lossless so a localized RWB edition can be produced without
  corrupting non-ASCII content.

**Next:** fold the operator's forthcoming details (how aeRWB text feeds
the production `.docx`, version/sync direction, translation pipeline)
into the next arc once supplied.

## Pointer back to the closed arc

Full dated history of the work that produced this carry-forward state is
in [`rvw/Code_review 2026-05-28.md`](Code_review%202026-05-28.md). That
file (and the arcs it points back to) covers the Header/Footer audit +
slot-84 arc, the StoryRanges rewrite and orphaned-story remediation, the
backup-script change, Tests 80-83, and the 2026-05-28 SpaceBefore /
Default Paragraph Font session.
