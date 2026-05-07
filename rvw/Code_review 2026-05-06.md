# Code review — 2026-05-06 carry-forward

This file opens a fresh review arc on 2026-05-06. The previous arc
[`rvw/Code_review 2026-04-30.md`](Code_review%202026-04-30.md) is now
**closed for new work**; that file remains the authoritative dated
history for everything between 2026-04-30 and 2026-05-06, including
the full VerseText migration record (Phases 1.1 through 2 closure),
the `Solomon` → `Song of Songs` rename sweep, the orphan-paragraph
audit + repair, the `EmphasisRed` paragraph-mark cleanup, and the
character-style audit parameterization.

Items below are the **open** carry-forward set. Each entry links back
to the section in the prior review where the full rationale lives.

## Closed in the prior arc (recorded here for completeness)

- **VerseText migration.** `BodyText` → `VerseText` paragraph-style
  migration complete on both test and production `.docm`. Final
  `ConvertParagraphContinuationVersesToVerseText` reports
  `scanned=0 converted=0 kept=0` (idempotent — no continuations
  remain). See `rvw/Code_review 2026-04-30.md` § "Phase 2 — CLOSED
  2026-05-06".
- **`Solomon` → `Song of Songs` rename sweep.** All five identified
  references plus one bonus comment in `aeBibleCitationClass.cls`
  updated. `Test_SongOfSongs_AllAliases` added (40 new assertions).
  `Run_All_SBL_Tests`: 222 / 0 PASS.
- **Orphan `BodyText` paragraph audit + repair.** 4 punctuation
  orphans repaired via `GoToPos` + Backspace. Final audit: 0 orphans;
  343 chapter intros + 175 chapter-end content excluded as legitimate
  non-verse content.
- **`EmphasisRed` paragraph-mark cleanup.** 227 paragraph marks
  carrying `EmphasisRed` character formatting reset via adapted
  `RUN_THE_TESTS(43)`. Test stable at 0.
- **Character-style usage audits.** `AuditCharStyleUsage` parameterized
  from the prior `AuditSelahUsage`. Selah / EmphasisBlack / EmphasisRed
  / Words of Jesus all surveyed: 3,429 total runs, all inside
  Phase-2-eligible verse paragraphs (0 policy flags).
- **Style taxonomy run state at arc close.** `RUN_TAXONOMY_STYLES`:
  **24 PASS / 4 FAIL across 28 checks**. Four FAILs are NOT-FOUND
  placeholders (`BookIntro`, `BodyTextContinuation`, `AppendixTitle`,
  `AppendixBody`). `VerseText` at priority 31 (45 styles promoted by
  `WordEditingConfig`). Test and production `.docm` in lockstep.

## Open carry-forward items

### 1. `AuditOneStyle` — extend for character-style properties

Currently `AuditOneStyle` only checks paragraph-style properties (font
name/size, alignment, indent, line spacing, space before/after).
Character styles like `Footnote Reference` are parked in bucket 2
(existence-verified) because the audit cannot meaningfully fully-specify
them — Bold, Italic, Color are not in the check list.

**Required for:** `Footnote Reference` to graduate from bucket 2 to
bucket 1.

**Scope:** add ~3 optional property arguments to `AuditOneStyle`
(`bExpBold`, `bExpItalic`, `lExpColor`) with sentinels (e.g. `-2` for
skip on Bold/Italic, `-1` for skip on Color since
`wdColorAutomatic = -16777216` is a real value). Or split into a
sibling `AuditOneCharStyle` with character-style-relevant checks only
— same pattern as `AuditStyleTabs` (Phase 6c).

Original analysis: `rvw/Code_review 2026-04-25.md` § "Footnote
Reference — deferred to bucket 2 (Q2 decision)" (2026-04-29).

### 2. Prescriptive-spec pass for known QA-checklist drift

The current taxonomy audit encodes **descriptive** specs (capture
today's values). Several known QA-checklist violations are tolerated
rather than driven to correction:

**`LineSpacingRule = Exactly` on paragraph styles** (QA checklist
preference is `Single`):
- `Heading 2` — `Exactly 10`
- `CustomParaAfterH1` — `Exactly 10`
- `Brief` — `Exactly 9.5`
- `Psalms BOOK` — `Exactly 10`
- `Footnote Text` — `Exactly 8`

**`BaseStyle = "Normal"` on approved styles** (QA checklist preference
is `""`):
- `CustomParaAfterH1`, `Brief`, `Footnote Text`, `Psalms BOOK`,
  `PsalmSuperscription`, `PsalmAcrostic`

Original analysis: `rvw/Code_review 2026-04-25.md` § "Spec promotion:
descriptive vs prescriptive (decision)" (2026-04-29) and § "Section (B)
full inventory findings" (2026-04-29).

**Recommendation:** treat as a series of one-property-at-a-time
decisions, each tracked as its own review item with rationale. Not a
single batch.

**Status update 2026-05-06:** partial QA-alignment pass applied to the
document (user-side) on five styles — `Footnote Text`, `Psalms BOOK`,
`Brief`, `CustomParaAfterH1`, `Heading 2`. Net effect on
`LineSpacingRule`: `Heading 2` and `Psalms BOOK` moved from `Exactly`
to `Single` (rule 0); `Brief` moved to `Multiple` (rule 5, 13.9pt);
`CustomParaAfterH1` and `Footnote Text` retained `Exactly` for now.
See item 9 below for the taxonomy resync that followed.

**Status update 2026-05-06 (BaseStyle half — DONE for all six):** the
six styles listed above (`CustomParaAfterH1`, `Brief`, `Footnote Text`,
`Psalms BOOK`, `PsalmSuperscription`, `PsalmAcrostic`) are now all at
`BaseStyle = ""` in the document, and the taxonomy enforces it via a
new optional `sExpBaseStyle` parameter on `AuditOneStyle` (sentinel
`"<skip>"`). `PsalmSuperscription` and `PsalmAcrostic` were also
promoted from bucket 2 to bucket 1 with full descriptive specs. See
item 10 below.

**Still open (LineSpacingRule = Exactly half):** `CustomParaAfterH1`
(`Exactly 10`) and `Footnote Text` (`Exactly 8`) remain at `Exactly`
pending a future prescriptive decision per QA preference.

### 3. Taxonomy audit — full-coverage final-state goal

Documented in `EDSG/01-styles.md` callout. Current state at arc close:
24 PASS / 4 FAIL across 28 checks. Four FAILs are NOT-FOUND placeholders
(`BookIntro`, `BodyTextContinuation`, `AppendixTitle`, `AppendixBody`).
Each move from bucket 2 → bucket 1 (when descriptive specs are encoded
for an existence-verified style) is a measurable step toward "every
approved style mapped with real specs". Remaining bucket-2 candidates:
`BookIntro`, `ListItemTab` legacy slot, `TheHeaders`, `TheFooters`,
`Title`, `Footnote Reference` (blocked on item 1 above),
`AuthorListItemTab` placeholder before tabs were audited.

Note: now that `VerseText` is the live verse-paragraph style across the
document, EDSG should reflect it as the primary translation target and
`BodyText` as the residual non-verse paragraph style. See item 5 below.

Original analysis: `rvw/Code_review 2026-04-25.md` § "Important —
taxonomy audit final-state goal" callout (2026-04-29).

### 4. Finding 5 (ribbon nav) — umbrella OPEN

Fix (A) resolved the primary caret-not-visible symptom. The residual is
the **idle-commit focus leak**: Word's customUI `editBox` auto-commits
on idle (~1 s) and returns focus to the document body, so any
subsequent Tab is a document Tab.

**Status:** **WORD LIMITATION, NO VBA-SIDE FIX AVAILABLE.** Documented
in the prior review. KeyTips are the supported Office UX path for
cross-control jumps and bypass Tab entirely.

**Forward options (deferred):**

- **Auto-fire Go on valid `(book, chapter, verse)` triple** — code-side
  option to remove the final Tab → Go step. Tradeoff: nav fires before
  user expects it; would need a `bAutoFire` toggle.
- **VSTO/WPF ribbon rewrite** — only path to true ribbon-owned focus
  management. Major rewrite; deferred indefinitely.

Original analysis: `rvw/Code_review 2026-04-25.md` § "Finding 5 —
terminology correction" (2026-04-29).

### 5. EDSG documentation refresh — VerseText-aware

Now that VerseText is the live verse-paragraph style, EDSG needs an
opportunistic refresh:

- **`EDSG/01-styles.md`** — add `VerseText` at priority 31 to the
  priority snapshot; reframe `BodyText` as the residual non-verse
  paragraph style (front matter, chapter intros, chapter-end content).
- **`EDSG/06-i18n.md`** — note `VerseText` as the primary translation
  target (the conversion's stated benefit: `VerseText` is a precise
  selector for verse paragraphs, which `BodyText` no longer is).
- **`EDSG/02-editing-process.md`** Stage 1 step references could
  mention `AuthorListItem*` as the canonical example for the
  `BaseStyle = ""` rule (currently uses generic phrasing).
- **`EDSG/04-qa-workflow.md`** "Current state" section dated 2026-04-26
  still mentions priorities 38-41 reserved gap and the 43-styles count
  — superseded by the 2026-04-29 SpeakerLabel insertion (now 39-42
  reserved, 44 styles) and again by the 2026-05-01 `VerseText`
  insertion (now 40-43 reserved, 45 styles).

**Recommendation:** opportunistic update next time these pages are
visited for substantive edits.

### 6. Body-content number prefixes — keep manual (decision recorded 2026-04-30)

Decision standing from the prior arc: keep manual text prefixes
(`"1. "`, `"2. "`, …) on `AuthorBookRef` paragraphs; do **not** convert
to `{ DOCVARIABLE }` fields. Carried forward only as a deferred-decision
record so the trigger-to-revisit isn't lost.

**Trigger to revisit:** active i18n rollout where a target locale
requires non-Arabic numerals in body content, substantially different
prefix punctuation that can't be handled by a one-pass reformatter, or
per-paragraph content substitution that today's manual prefixes can't
accommodate.

Full reasoning: `rvw/Code_review 2026-04-30.md` § "8. Body-content
number prefixes — keep manual, no docvariables (decision 2026-04-30)".

### 7. `AuditVerseMarkerStructure` — missing `Chapter Verse marker` coverage (BUG, 2026-05-06)

**Project rule (now stated explicitly in `basVerseStructureAudit.bas` header):** every verse paragraph (now `VerseText`) leads with a `Chapter Verse marker` character-style run (chapter number) **immediately followed by** a `Verse marker` character-style run (verse number). One CVM + one VM per verse — no exceptions.

**Bug:** user found a verse paragraph in the document where the leading `Chapter Verse marker` was missing (only the `Verse marker` was present). `AuditVerseMarkerStructure` reported the chapter as OK.

**Root cause:** the audit never inspects `Chapter Verse marker` at all. `CountVerseMarkers` (`src/basVerseStructureAudit.bas:228`) finds character-style runs whose style is `"Verse marker"` only:

```vba
.style = oDoc.Styles("Verse marker")
```

That count is compared against `aeBibleCitationClass.VersesInChapter(bookName, chIdx)` (lines 204-205). Chapter boundaries come from `Heading 2` paragraph starts (line 176). Nothing in the routine ever enumerates, counts, or asserts the presence of `Chapter Verse marker`.

The failure mode is exactly what the design allows:

- The verse paragraph's leading `Chapter Verse marker` run is missing, but the verse-number portion is still styled `Verse marker`.
- `CountVerseMarkers` still finds the VM run → chapter total matches `VersesInChapter` → reported `OK`.
- The chapter's `Heading 2` is also still present → chapter count matches → also `OK`.

The audit name oversells what it does. It is really a **verse-count-by-chapter** audit. The `Chapter Verse marker` half of "verse-marker structure" is not audited.

**Fix scope (one per verse, the stated rule):**

1. Add `CountChapterVerseMarkers(oDoc, startPos, endPos)` mirroring `CountVerseMarkers` but with `.style = oDoc.Styles("Chapter Verse marker")`.
2. In `AuditOneBook`, per chapter, also compute `foundCVMs = CountChapterVerseMarkers(...)` and assert `foundCVMs == foundVerses` (CVM count must equal VM count, since the rule is one CVM + one VM per verse).
3. Status flips to `MISMATCH` and increments `bookIssues` when CVM ≠ VM. Issue line shows both counts: `expected verses=N  found VM=N  found CVM=M  MISMATCH (CVM gap)`.
4. Optional stricter check (deferred — costlier, may not be needed once counts match): walk each verse paragraph and verify the **adjacency** rule (CVM run immediately followed by VM run as the paragraph's first two character-style runs). The count-equality check catches the symptom; the adjacency check would catch the structural pathology of a CVM existing somewhere else in the paragraph.

**Header update applied 2026-05-06:** `src/basVerseStructureAudit.bas` header now explicitly states the project verse-marker rule (one CVM + one VM per verse, no exceptions). After the code fix below, the header was updated again to promote CVM coverage from "known gap" to "invariant 4".

#### Fix — `CountChapterVerseMarkers` extension APPLIED 2026-05-06

`src/basVerseStructureAudit.bas`:

- New private function `CountChapterVerseMarkers(oDoc, startPos, endPos)` — parallel to `CountVerseMarkers`, identical body except `.style = oDoc.Styles("Chapter Verse marker")`. Same 20,000 safety cap; same range-walk pattern.
- `AuditOneBook` per-chapter loop now also computes `foundCVMs` and asserts `foundCVMs = foundVerses` (the "one CVM + one VM per verse" rule). On mismatch, sets `cvmStatus = "CVM-MISMATCH"`, increments `bookIssues`, and appends an issue line of the form:

  ```
  GENESIS 1: CVM count 1532 <> VM count 1533 (one CVM+VM per verse rule)
  ```

- `chapterReport` line extended to show both counts and both statuses:

  ```
  ch   1: expected verses= 31  found VM= 31  found CVM= 31  OK/OK
  ```

  The `status/cvmStatus` pair makes it obvious which dimension failed when something does fail (`OK/CVM-MISMATCH`, `MISMATCH/OK`, or `MISMATCH/CVM-MISMATCH`).

- Header invariants list updated: gap removed; CVM coverage promoted to invariant 4 ("per-chapter Chapter Verse marker Count equals Verse marker Count").

**Behavior change:** the audit now flags the failure mode the user discovered. A verse paragraph missing its leading CVM run will reduce `foundCVMs` by 1, the chapter's CVM/VM equality breaks, status becomes `OK/CVM-MISMATCH` (or `MISMATCH/CVM-MISMATCH` if VM is also off), and the issue line surfaces in the `ISSUES FOUND` section of the report.

**Adjacency check (deferred):** count-equality catches the symptom (CVM run is missing somewhere). It does **not** catch the structural pathology of a CVM run existing in a verse paragraph but not adjacent to (and immediately preceding) the VM run. If that pathology turns up, a paragraph-walk pass that inspects each verse paragraph's first two character-style runs would close it. Not implemented now — counts-equal is the cheaper invariant and matches the rule the user stated.

**Status:** code applied. Awaiting user-side action: re-import `basVerseStructureAudit.bas` into the test `.docm`, run `AuditVerseMarkerStructure`, paste the SUMMARY line and any new `CVM-MISMATCH` issue lines. Expected initial result: at least one `CVM-MISMATCH` line corresponding to the missing-CVM verse the user found. Once the data defect is repaired and re-audited, expected steady state is 0 issues across all 66 books.

#### Verified 2026-05-06 — fix caught two real defects

First run after re-import:

```
ISSUES FOUND:
  Genesis 8: CVM count 21 <> VM count 22 (one CVM+VM per verse rule)
  Job 37:    CVM count 25 <> VM count 24 (one CVM+VM per verse rule)

SUMMARY: 31102 / 31102 verses found, 2 structural issue(s).
AuditVerseMarkerStructure - actual 131.72 sec
```

Two real data defects, opposite signs:

- **Genesis 8** — CVM short by 1 (21 vs 22 VMs). One verse missing its leading `Chapter Verse marker` run. This is the failure mode the user originally found that was passing the old VM-only audit.
- **Job 37** — CVM long by 1 (25 vs 24 VMs). A `Chapter Verse marker` run exists somewhere in the chapter that isn't paired with a VM — likely a stray CVM left from an editing operation, or a run that should have been a VM but got the CVM character style instead.

User repaired both defects. Re-run:

```
SUMMARY: 31102 / 31102 verses found, 0 structural issue(s).
AuditVerseMarkerStructure - actual 131.66 sec
```

**Audit clean.** All 31,102 verses now satisfy the "one CVM + one VM per verse" rule across all 66 books.

#### Performance note

Run time ~131.7 sec on the production-scale `.docm` (consistent across both runs). The CVM extension roughly doubled the per-chapter find work (CVM scan + VM scan instead of VM only). No optimization needed — the audit is run on demand, not in a hot loop, and the 2-minute cost is acceptable for full structural verification. If runtime becomes a concern, the two scans could be unified by walking each verse paragraph once and inspecting its first two character-style runs (which would also enable the deferred adjacency check in the same pass).

**Item 7 closed.**

### 8. Session manifest

`sync/session_manifest.txt` from the prior arc covered src/ edits
through 2026-04-30. The 2026-05-01 → 2026-05-06 VerseText, orphan-audit,
and char-style-audit edits should be recorded in a fresh manifest entry
(or the existing one updated) as the import checklist for any fresh
`.docm` re-sync.

### 9. Taxonomy resync after QA-alignment partial pass — CLOSED 2026-05-06

After the user adjusted five paragraph styles in the document toward
QA guidelines (see item 2 status update), `RUN_TAXONOMY_STYLES` was
re-run and produced three new FAILs (`Heading 2`, `Brief`,
`Psalms BOOK`) because the audit's expected values still encoded the
pre-change descriptive specs.

**Inputs:** dumped style snapshots in `rpt/Styles/` —
`style_Heading_2.txt`, `style_Brief.txt`, `style_Psalms_BOOK.txt`,
`style_CustomParaAfterH1.txt`, `style_Footnote_Text.txt`. The latter
two already matched the existing taxonomy and needed no edit.

**Edits applied to `src/basTEST_aeBibleConfig.bas` (lines 282–298 region):**

- `Heading 2` — `LineRule 4 -> 0` (Single), `LineSpacing 10 -> 12pt`.
- `Brief` — `LineRule 4 -> 5` (Multiple), `LineSpacing 9.5 -> 13.9pt`,
  `SpaceAfter 0 -> 8pt`.
- `Psalms BOOK` — `Indent 14.4 -> 0`, `LineRule 4 -> 0` (Single),
  `LineSpacing 10 -> 12pt`, `SpaceAfter 0 -> 8pt`.

**Verification:** after re-import to the test `.docm`,
`RUN_TAXONOMY_STYLES` reports **24 PASS / 4 FAIL across 28 checks**.
The remaining four FAILs are the expected NOT-FOUND placeholders
(`BookIntro`, `BodyTextContinuation`, `AppendixTitle`,
`AppendixBody`). All five touched-by-QA styles now PASS.

**Note on spec character:** these are still **descriptive** specs
(captured to match document state). Per item 2, a future prescriptive
pass would drive the remaining `LineSpacingRule = Exactly` values on
`CustomParaAfterH1` and `Footnote Text` toward QA preferences, plus
the `BaseStyle = "Normal"` drift.

**Item 9 closed.**

### 10. BaseStyle = "" prescriptive pass + Psalm bucket-1 promotions — CLOSED 2026-05-06

First single-property prescriptive decision from item 2's list,
applied as three steps:

**Step 1 — `AuditOneStyle` extended.** New optional parameter
`sExpBaseStyle As String = "<skip>"` and a corresponding check block
modeled on the existing Bold/Italic pattern. All 17 prior callers
default to skip, so the extension is a strict superset (zero
behavior change at point of merge).

**Step 2 — invariant turned on for the four already-bucket-1
styles** that the QA list flagged: `CustomParaAfterH1`, `Brief`,
`Psalms BOOK`, `Footnote Text`. Each call site appended `, ""`. Doc
state already complied (verified via `rpt/Styles/` dumps), so
`RUN_TAXONOMY_STYLES` stayed at 24 PASS / 4 FAIL.

**Step 3 — `PsalmAcrostic` and `PsalmSuperscription` promoted from
bucket 2 to bucket 1** with full descriptive specs from
`rpt/Styles/`:

- `PsalmAcrostic` — Carlito 9pt, Center, Single 12pt, SpB/SpA = 3,
  `BaseStyle = ""`. (Dump also shows `SmallCaps = -1` and
  `QuickStyle = False`; neither dimension is checked by
  `AuditOneStyle` — same as for every other audited style.)
- `PsalmSuperscription` — Carlito 8pt, Left, Multiple 13.9pt,
  SpB/SpA = 2, Italic = -1, `BaseStyle = ""`. First dump for this
  style still showed `BaseStyle = "Normal"`; user re-applied the QA
  edit, re-dumped, and confirmed `""` before the audit line was
  added.

**Verification:** `RUN_TAXONOMY_STYLES` reports **26 PASS / 4 FAIL
across 30 checks**. Four FAILs remain the expected NOT-FOUND
placeholders (`BookIntro`, `BodyTextContinuation`, `AppendixTitle`,
`AppendixBody`). Bucket 1 grew 9 -> 11.

**Property-coverage gap noted:** `AuditOneStyle` still does not check
`SmallCaps`, `QuickStyle`, `Underline`, `Color`, `RightIndent`,
`LeftIndent`, `KeepWithNext`, `OutlineLevel`, `NextParagraphStyle`,
`AutomaticallyUpdate`, `WidowControl`, `KeepTogether`,
`PageBreakBefore`. Each is a candidate for a future
"one-property-at-a-time" extension when the corresponding QA
preference is decided.

**Item 10 closed.**

### 11. Front-matter & TOC styles bucket-1 promotion — CLOSED 2026-05-06

User dumped 14 front-matter / TOC / index styles to `rpt/Styles/`
after a QA-alignment pass: `FrontPageTopLine`, `TitleEyebrow`,
`Title`, `TitleVersion`, `FrontPageBodyText`, `BodyTextTopLineCPBB`,
`Acknowledgments`, `AuthorBodyText`, `Contents`, `ContentsRef`,
`BibleIndexEyebrow`, `BibleIndex`, `Introduction`, `TitleOnePage`.
All 14 dumps show `BaseStyle = ""`.

**Edits applied to `src/basTEST_aeBibleConfig.bas`:**

- 12 new bucket-1 entries added in a "Front matter & TOC styles
  (priorities 4-17)" group below `PsalmSuperscription`. Each carries
  full descriptive specs (Font / Size / Align / Indent / LineRule /
  LineSp / SpB / SpA / Bold / Italic) plus `BaseStyle = ""`.
- `Title` promoted from bucket 2 (existence-only placeholder) to
  bucket 1 with full spec (Times New Roman 36pt, Center, Single 12pt,
  SpB/SpA = 0, `BaseStyle = ""`). Bucket-2 placeholder removed.
- `ContentsRef` (already in bucket 1) had `, ""` appended to enforce
  `BaseStyle = ""`. All other properties already matched the dump.

**Spec character notes:**

- `BodyTextTopLineCPBB` dump shows `PageBreakBefore = -1` and
  `WidowControl = 0` — neither dimension is checked by `AuditOneStyle`,
  so these aren't enforced; descriptive only.
- `ContentsRef` dump shows 1 explicit tab stop (`Position=378
  Align=Right Leader=Dots`). Not yet covered by `AuditStyleTabs` -
  candidate for a future tab-stop audit addition (see item 12 below).
- `AuthorBodyText` has `FirstLineIndent = 23.75` (the only first-line
  indent in the front-matter group) and `Alignment = 3` (Justify);
  encoded as such.

**Verification expected:** `RUN_TAXONOMY_STYLES` -> **38 PASS / 4 FAIL
across 42 checks**. Bucket 1: 11 -> 24. Bucket 2: 5 -> 4 (Title
removed). Four FAILs remain the expected NOT-FOUND placeholders
(`BookIntro`, `BodyTextContinuation`, `AppendixTitle`,
`AppendixBody`).

**Verified 2026-05-06:** `RUN_TAXONOMY_STYLES` reports **38 PASS / 4
FAIL across 42 checks** after re-import. Item 11 closed.

### 12. ContentsRef tab-stop coverage — CLOSED 2026-05-06

`rpt/Styles/style_ContentsRef.txt` shows one explicit tab stop
(`Position=378 Align=Right Leader=Dots`). The `AuditStyleTabs` block
covered `AuthorListItem`, `AuthorListItemTab`, `AuthorBookRef`, and
`AuthorBookRefHeader` but not `ContentsRef`.

**Edit applied** to `src/basTEST_aeBibleConfig.bas` (Tab stops
verified section):

```vba
AuditStyleTabs "ContentsRef", _
    Array(378, wdAlignTabRight, wdTabLeaderDots)
```

**Verified 2026-05-06:** `RUN_TAXONOMY_STYLES` reports **39 PASS / 4
FAIL across 43 checks**. Tab-stop coverage: 4 -> 5 styles. Item 12
closed.

## Pointer back to the closed arc

Full dated history of the work that produced this carry-forward state
is in [`rvw/Code_review 2026-04-30.md`](Code_review%202026-04-30.md).
That file includes:

- The complete VerseText migration record (Phases 1.1 through 2
  closure) with every audit, repair, and verification.
- The `Solomon` → `Song of Songs` rename sweep and the
  `Test_SongOfSongs_AllAliases` test addition.
- The orphan-paragraph audit refinement (v1 too greedy → refined rule
  → 4 real orphans → 0 after manual repair).
- The `EmphasisRed` paragraph-mark cleanup.
- The `AuditCharStyleUsage` parameterization and four-style survey.

Anything in this 2026-05-06 file should reference back to that arc
for the *why*; this file holds only the **what is still open**.
