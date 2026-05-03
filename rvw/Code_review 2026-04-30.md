# Code review — 2026-04-30 carry-forward

This file opens a fresh review arc on 2026-04-30. The previous arc
[`rvw/Code_review 2026-04-25.md`](Code_review%202026-04-25.md) is now
**closed for new work**; that file remains the authoritative dated
history for everything between 2026-04-25 and 2026-04-30, including
the full List Paragraph migration record (Phases 0 through 6), the
WEB versification corrections, and the spec-promotion decisions.

Items below are the **open** carry-forward set. Each entry links back
to the section in the prior review where the full rationale lives.

## Closed (recorded here for completeness)

- **VerseStructureAudit baseline.** 31,102 / 31,102 verses, 0 structural
  issues. Document content matches the WEB Protestant Edition source
  exactly. Audit module DRY-refactored against `aeBibleCitationClass`.
- **List Paragraph migration.** End-to-end complete on test copy and
  production. Three styles renamed (`AuthorListItem`, `AuthorListItemBody`,
  `AuthorListItemTab`); `AuthorBookRef` rebased standalone. All
  `BaseStyle = ""`. `AuditListStyleRisk` reports 0 flagged.
- **Tab-stop infrastructure.** `DumpStyleProperties`, `CopyOneStyle`, and
  `AuditStyleTabs` now all handle `ParagraphFormat.TabStops`. The
  taxonomy audit has full coverage of every audit-able property on
  every approved paragraph style.
- **`AuthorBookRef` promoted to fully-spec + tab-stop audited (2026-04-30).**
  User added tab stops manually to both docs. Six-step promotion run:
  fresh `DumpStyleProperties`; one `AuditOneStyle "AuthorBookRef", "Carlito",
  11, 0, -18, 0, 12, 0, 11` line added in bucket 1; one `AuditStyleTabs
  "AuthorBookRef", Array(36, wdAlignTabLeft, wdTabLeaderSpaces),
  Array(378, wdAlignTabRight, wdTabLeaderDots)` line added in tab-stops
  bucket; PURPOSE block updated to 21 styles + 2 tab-stop specs = 23
  total checks. **Verified 2026-04-30: `RUN_TAXONOMY_STYLES` reports
  18 PASS / 5 FAIL across 23 checks**, both new entries (`AuthorBookRef`
  and `AuthorBookRef (TabStops)`) landed clean.
- **Style taxonomy run state.** `RUN_TAXONOMY_STYLES`: **18 PASS / 5 FAIL
  across 23 checks** (post-`AuthorBookRef` promotion, 2026-04-30). Five
  FAILs are pre-tracked items (see deferred list below).
- **Phase 7 — Bold / Italic audit coverage (2026-04-30).** `AuditOneStyle`
  extended with two optional args (`vExpBold`, `vExpItalic`, sentinel `-2`).
  All 12 fully-specified entries got descriptive Bold/Italic specs.
  `ContentsRef` and `AuthorBookRefHeader` added as new bucket-1 entries
  with full font + paragraph specs and (for `AuthorBookRefHeader`) a
  tab-stop entry at 381.6 pt Right Spaces. The `AuthorBookRefHeader`
  Bold drift was caught and corrected: dump showed `Bold = 0` (False)
  contradicting design intent; user fixed the live style to `Bold = -1`
  (Path A) before encoding. **Verified 2026-04-30: `RUN_TAXONOMY_STYLES`
  reports 21 PASS / 5 FAIL across 26 checks**, all three new entries
  PASS (`ContentsRef`, `AuthorBookRefHeader`, `AuthorBookRefHeader (TabStops)`).
  Bold drift on `AuthorBookRefHeader` is now caught immediately by the
  audit — any future flip back to `Bold=False` will FAIL the style row.

- **Bucket 1 grew 10 → 12.** Added `ContentsRef` and `AuthorBookRefHeader`.
  All twelve now audited against descriptive Bold / Italic in addition
  to the original eight property classes.

- **Author* trio promoted to bucket 1 (2026-05-01).** `AuthorListItem`,
  `AuthorListItemBody`, `AuthorListItemTab` moved from bucket 2
  (existence-verified) to bucket 1 (fully specified) with descriptive
  font + paragraph + Bold/Italic specs lifted from fresh dumps. Closes
  the long-running `AuthorListItem` FirstLineIndent drift (expected 0
  vs live -18): descriptive spec encodes -18, FAIL clears. New
  `AuditStyleTabs` entry added for `AuthorListItem` (1 stop at 36 pt
  Left Spaces — newly added by user since the migration completed).
  Audit count: 23 styles + 4 tab-stop specs = **27 total checks**;
  bucket distribution: 15 / 5 / 3. **Verified 2026-05-01:
  `RUN_TAXONOMY_STYLES` reports 23 PASS / 4 FAIL across 27 checks**;
  all five `AuthorListItem*` rows PASS (3 style + 2 tab-stop). The
  4 remaining FAILs are all NOT-FOUND placeholders (`BookIntro`,
  `BodyTextContinuation`, `AppendixTitle`, `AppendixBody`).
- **Ribbon focus fix (Finding 5 fix A).** Caret renders correctly on
  first nav from a ribbon-owned event handler.
- **`ToSBLShortForm` "Song of Songs" lookup error (Finding 3).** No
  defect reproducible from static analysis; live-run confirms alias
  path sound. Closed 2026-04-29.

## Open carry-forward items

### 1. Reference rename sweep — `Solomon` → `Song of Songs` (CLOSED 2026-05-01)

Originally listed as open. **Closed 2026-05-01** — Phase A rename sweep applied across all five identified locations plus one additional comment in `aeBibleCitationClass.cls` discovered during the verification sweep:

| File | Change |
|---|---|
| `src/basTEST_aeBibleCitationClass.bas:885` | `canonNames(22) = "Solomon"` → `"Song of Songs"` |
| `src/basSBL_VerseCountsGenerator.bas:95` | Context label `"Solomon"` → `"Song of Songs"` |
| `src/XbasTESTaeBibleDOCVARIABLE.bas:527` | `VerifyBookNameFromDocVariable "Song", "Solomon"` → `"Song", "Song of Songs"` |
| `md/Deterministic Structural Parser.md:83` | Table row `Solomon` → `Song of Songs` |
| `md/Deterministic Structural Parser.md:314` | Multi-word example `"Song of Solomon"` → `"Song of Songs"` |
| `src/aeBibleCitationClass.cls:1826` | Multi-word example comment `"Song of Solomon"` → `"Song of Songs"` |

**Intentionally kept** (non-canonical references with documented purpose):

- `aeBibleCitationClass.cls:1472` — comment "Song of Songs, aka Song of Solomon" documenting that `"Song of Solomon"` is a recognised alias (it's in the alias map at line 1478 already).
- `basChangeLog_aeBibleClass.bas:17, 18` — change-log entries (#617, #618). Dated history; no retroactive rewrites per `EDSG/09-history.md` policy.

#### Phase B — existing citation tests (closed 2026-05-01)

User ran `Run_All_SBL_Tests`. Result:

```
Tests Run:  44
Failures:   0
Result: PASS
```

All 44 citation-class tests pass after the rename. The two highest-impact tests (`Test_Stage1_AliasCoverage` and `Test_CanonicalNamesAndSBLTable`) now exercise the `"Song of Songs"` canonical name through the alias map and SBL short-form path, both clean.

Side note: `Test_Stage1_AliasCoverage` and most other `Test_Stage*` Subs depend on a module-level `aeAssert` that's initialised inside `Run_All_SBL_Tests`. They aren't designed to be invoked standalone — running them individually from the Immediate window raises `Object variable or with block variable not set` because `aeAssert` is `Nothing`. The aggregator is the supported entry point. Documented here for future debugging.

#### Phase C — `Test_SongOfSongs_AllAliases` added (closed 2026-05-01)

Focused alias-coverage Sub appended to `basTEST_aeBibleCitationClass.bas` and wired into `Run_All_SBL_Tests` immediately after `Test_Stage1_AliasCoverage`. Coverage:

- **All 13 alias forms** (canonical, three case variants of canonical, plus 10 documented short aliases including `Solomon` as a still-valid alias) → BookID 22, canonical name `"Song of Songs"`.
- **Negative case** — `"Song of Solomon"` (multi-word, not in alias map) raises. Documents the boundary explicitly so any future addition of `"SONG OF SOLOMON"` to the alias map flips this assertion deliberately.
- **`ChaptersInBook`** via canonical and two aliases (`"Song"`, `"Solomon"`) — all 8.
- **`VersesInChapter`** for all 8 chapters against the WEB-aligned data (17/17/11/16/16/13/13/14).
- **`ToSBLShortForm`** for the canonical input → SBL `"Song N:V"` output.

Uses the suite-level `aeAssert` (no local instantiation) — matches the `Test_Stage*` pattern in the codebase.

**Status:** rename sweep + tests complete. Confirmed 2026-05-01 via `Run_All_SBL_Tests`:

```
Tests Run: 222
Failures:  0
Result:    PASS
```

`Test_SongOfSongs_AllAliases` contributed 40 new PASS assertions (26 alias × 2 + 1 negative + 3 ChaptersInBook + 8 VersesInChapter + 2 ToSBLShortForm). The previous "44 / 0" the user saw in the Immediate window was the **`Run_Extra_Tests`** summary (a separate aeAssert instance running after `Run_All_SBL_Tests` completes); the Stage suite's actual summary scrolled off the Immediate buffer but is preserved in `rpt\SBL_Tests.UTF8.txt`.

#### Original open-state record — preserved for traceability

The five-row table below was the original list when this item was open. Each row is now resolved per the closure tables above. Kept verbatim for audit history.

| File | Line | Original context |
|---|---|---|
| `src/basTEST_aeBibleCitationClass.bas` | 885 | Test oracle: `canonNames(22) = "Solomon"`. Needed to be `"Song of Songs"` to match the citation class's canonical form. |
| `src/basSBL_VerseCountsGenerator.bas` | 95 | Context label string passed to `ToOneBasedLongArray`. Cosmetic — did not affect data. |
| `src/XbasTESTaeBibleDOCVARIABLE.bas` | 527 | `VerifyBookNameFromDocVariable "Song", "Solomon"` — document-specific assertion. Earlier review (2026-03-16) marked this as "correct for this document"; status changed once project canonical moved to `"Song of Songs"`. |
| `md/Deterministic Structural Parser.md` | 83 | Reference table row: `Solomon`. |
| `md/Deterministic Structural Parser.md` | 314 | Multi-word example: `"Song of Solomon"`. |

Original analysis: `rvw/Code_review 2026-04-25.md` § "Finding 4 — broader citation-code impact of the rename" (2026-04-28).

### 2. `VerseText` style introduction + bulk conversion (original goal — preconditions now met)

**Plan location:** `rvw/Code_review 2026-04-25.md` § "Plan - introduce `VerseText` style - 2026-04-26".

**Goal recap:** introduce a `VerseText` paragraph style identical in spec to `BodyText`, then bulk-convert every paragraph in the Bible body that begins with a `Chapter Verse marker` character-style run from `BodyText` to `VerseText`. `BodyText` retains for non-verse paragraphs (intros, spacers, footnote contexts).

**Why this is on the carry-forward.** This is the **original goal** of the entire 2026-04-25 → 2026-05-01 work arc. Every audit infrastructure piece built in that arc — `AuditVerseMarkerStructure`, the descriptive-spec promotion of bucket-1 styles, the migration of List Paragraph–entangled styles, the `Bold` / `Italic` audit extension — was scaffolding to make the VerseText conversion safe to run. The conversion itself was never executed.

**Precondition status (all green as of 2026-05-01):**

| Precondition | Status |
|---|---|
| `AuditVerseMarkerStructure` reports clean | ✓ 31,102 / 31,102 verses, 0 structural issues |
| Style taxonomy baseline established | ✓ 23 PASS / 4 FAIL across 27 checks, all 4 FAILs are NOT-FOUND placeholders |
| `BodyText` is fully-spec audited (descriptive) | ✓ bucket 1, locked at Carlito 9, Justify, LineSpacing 10 Exactly, Bold/Italic False |
| List Paragraph migration complete | ✓ no list-engine entanglement on any approved style |
| Reference rename (Solomon → Song of Songs) | ✓ closed 2026-05-01 |

**Two-phase execution (per the original plan, abbreviated):**

#### Phase 1 — define the style (low risk)

- **1.1** Add `DefineVerseTextStyle` to `src/basFixDocxRoutines.bas` — clones every property of `BodyText`. `BaseStyle = ""`, no list-engine attachment.
- **1.2** Add `"VerseText"` to the `approved` array in `src/basTEST_aeBibleConfig.bas`. Position: immediately after `"Verse marker"` (puts it in the verse-marker cluster). Renumbers downstream priorities by +1.
- **1.3** Add `VerseText` row to `RUN_TAXONOMY_STYLES`, fully-spec, same expected values as `BodyText` (separate row; not "based on" since `BaseStyle = ""`).
- **1.4** Run `WordEditingConfig` + `DumpStyleProperties "VerseText"` to confirm the new style matches `BodyText` property-for-property.

#### Phase 2 — bulk conversion (one-time mutation, BACKUP FIRST)

- **2.1** Commit clean repo + back up the working `.docm`.
- **2.2** Add `ConvertBodyTextVersesToVerseText` to `src/basFixDocxRoutines.bas` — iterates all paragraphs, converts `BodyText` → `VerseText` when first character's style is `"Chapter Verse marker"`.
- **2.3** Run conversion. Expected: ~31,000 conversions (one per verse); small `kept` count (front-matter / non-verse paragraphs styled BodyText).
- **2.4** Visual verification by sampling.
- **2.5** Re-run `RUN_TAXONOMY_STYLES` and `AuditVerseMarkerStructure` — confirm both still clean.
- **2.6** Update EDSG (`01-styles.md` snapshot, `06-i18n.md` notes on VerseText as primary translation target).

**Stated benefits (from the original plan):**

- **First-occurrence semantics** — `VerseText` first appears at Genesis 1:1, the natural canonical anchor. `BodyText` currently first appears on page 1 (front-matter spacer), which is semantically misleading.
- **USFM mapping clarity** — `VerseText` → `\v` (verse body); `BodyText` → `\p` / `\ip` / etc. for non-verse paragraphs.
- **Find / Replace by style** becomes meaningful — `VerseText` is the precise selector for "all verse paragraphs", which is currently impossible with `BodyText` (mixed with front-matter content).

### Step-by-step execution plan (2026-05-01, approved)

#### Pre-flight (Step 0) — preconditions confirmed green 2026-05-01

| Check | Required | Verified |
|---|---|---|
| `AuditVerseMarkerStructure` | 31,102 / 31,102, 0 issues | ✓ |
| `RUN_TAXONOMY_STYLES` baseline | 23 PASS / 4 FAIL across 27 checks | ✓ |
| `BodyText` is fully-spec audited | bucket 1, descriptive | ✓ |
| Both `.docm` files in sync | same audit baseline | ✓ |
| Working tree clean | `git status` empty | (user-managed) |

#### Phase 1 — Define `VerseText` style (low risk, reversible)

Order of operations:

| Step | Action | Edit target | Check after |
|---|---|---|---|
| 1.1 | Add `DefineVerseTextStyle` Sub | `src/basFixDocxRoutines.bas` | VBA compiles |
| 1.2 | Run `DefineVerseTextStyle` in test `.docm` | (user action, Immediate) | `DumpStyleProperties "VerseText", True` matches `style_BodyText.txt` byte-for-byte (excluding name + priority) |
| 1.3 | Add `"VerseText"` to `approved` array | `src/basTEST_aeBibleConfig.bas` | VBA compiles |
| 1.4 | Add `VerseText` row to `RUN_TAXONOMY_STYLES` bucket 1 | `src/basTEST_aeBibleConfig.bas` | VBA compiles; PURPOSE block updated |
| 1.5 | Re-import both modules into test `.docm` | (user action) | VBA editor shows updated revisions; no compile errors |
| 1.6 | Run `WordEditingConfig` | (user action) | 45 styles promoted; `VerseText` at priority 31 |
| 1.7 | Run `RUN_TAXONOMY_STYLES` | (user action) | **24 PASS / 4 FAIL across 28 checks** |
| 1.8 | Visual sanity in Word | (user action) | `VerseText` listed in Styles pane; no paragraph yet using it |
| 1.9 | Replicate Steps 1.5-1.8 on production `.docm` | (user action) | Both docs at the same post-Phase-1 baseline |

**Phase 1 verification gate:** all 1.x checks green. If anything is off, **stop**; diagnose; do **not** proceed to Phase 2.

#### Phase 2 — Bulk conversion (one-time mutation; backup first)

| Step | Action | Edit target | Check after |
|---|---|---|---|
| 2.1 | Backup the production `.docm` to versioned filename | (user action) | Backup file exists; opens in Word |
| 2.2 | Add `ConvertBodyTextVersesToVerseText` Sub | `src/basFixDocxRoutines.bas` | VBA compiles |
| 2.3 | Run on test `.docm` | (user action) | `converted ≈ 31,102`; `kept` is small |
| 2.4 | Visual spot-check on test `.docm` | (user action) | Verses now `VerseText`; non-verse `BodyText`; rendering unchanged |
| 2.5 | Re-run audits on test `.docm` | (user action) | `RUN_TAXONOMY_STYLES` 24/4; `AuditVerseMarkerStructure` 31102/31102 |
| 2.6 | Idempotency check on test `.docm` | (user action) | Second run reports `converted = 0` |
| 2.7 | Run on production `.docm` (Steps 2.3-2.6) | (user action) | Same outcomes |

**Identification rule (used by `ConvertBodyTextVersesToVerseText`):** a paragraph qualifies for conversion when both:

1. `paragraph.Style.NameLocal = "BodyText"`, AND
2. `paragraph.Range.Characters(1).Style.NameLocal = "Chapter Verse marker"`

**Phase 2 verification gate:** all 2.x checks green on test, then on production. If anything is off, **stop**; restore from backup if needed.

#### Phase 3 — Documentation closeout

| Step | Edit target | What |
|---|---|---|
| 3.1 | `EDSG/01-styles.md` | Add `VerseText` at priority 31 in the priority snapshot |
| 3.2 | `EDSG/06-i18n.md` | Add `VerseText` as the primary translation target |
| 3.3 | `rvw/Code_review 2026-04-30.md` | Mark item 2 as CLOSED with verified counts |
| 3.4 | `sync/session_manifest.txt` | Record VerseText work as a new session theme |

#### Failure recovery

- **Phase 1:** delete `VerseText` style via VBA; revert array + audit edits; `WordEditingConfig` reasserts prior state. No data loss.
- **Phase 2 mid-conversion:** the conversion is reversible by inverse Sub (`VerseText` → `BodyText` for paragraphs whose first character is `Chapter Verse marker`). For non-trivial issues, restore from Step 2.1 backup.

### Execution log

#### Phase 1.1 — `DefineVerseTextStyle` Sub APPLIED 2026-05-01

`src/basFixDocxRoutines.bas` — new Sub added between `DefineBodyTextStyle` and `DefineBodyTextIndentStyle`:

- `BaseStyle = ""` set first (per EDSG List Paragraph rule)
- `NextParagraphStyle = self` (verses follow verses)
- `AutomaticallyUpdate = False`, `QuickStyle = False` (QA-checklist)
- Font: Carlito 9, Bold/Italic False
- Paragraph: Justify, Exactly 10pt, FirstLineIndent 0, LeftIndent 0, SpaceBefore/After 0
- RERUN SAFE: exits without changes if `VerseText` already exists
- USFM mapping documented: `\v <number> <text>` (verse body content)
- Standard `PROC_EXIT` / `PROC_ERR` pattern matching project convention

**Status:** Phase 1.1 applied. Awaiting user-side action for Phase 1.2 (`DefineVerseTextStyle` execution + dump verification).

#### Phase 1.2 — `DefineVerseTextStyle` executed on test `.docm` (2026-05-01)

User ran `DefineVerseTextStyle` followed by `DumpStyleProperties "VerseText", True`. Output confirms byte-identical match to `style_BodyText.txt` except for name, priority (1, pending Phase 1.6 promotion to 31), and `NextParagraphStyle` (self — same self-pointing pattern as BodyText). All 13 paragraph properties match exactly.

#### Phase 1.3 — `"VerseText"` added to `approved` array (APPLIED 2026-05-01)

`src/basTEST_aeBibleConfig.bas` `approved` array:

```
"Heading 2", "Chapter Verse marker", "Verse marker", _
"VerseText", _
"Footnote Reference", "Footnote Text", "Psalms BOOK", _
```

`VerseText` placed immediately after `Verse marker` per the original 2026-04-26 plan. Position 31 in book-occurrence order. Downstream priorities (`Footnote Reference` and below) shift +1 when `WordEditingConfig` runs.

#### Phase 1.4 — `VerseText` row added to `RUN_TAXONOMY_STYLES` (APPLIED 2026-05-01)

`src/basTEST_aeBibleConfig.bas` bucket 1, immediately after `BodyText`:

```vba
AuditOneStyle "BodyText", "Carlito", 9, 3, 0, 4, 10, 0, 0, 0, 0
AuditOneStyle "VerseText", "Carlito", 9, 3, 0, 4, 10, 0, 0, 0, 0
AuditOneStyle "BodyTextIndent", "Carlito", 9, 3, 14.4, 4, 10, 0, 0, 0, 0
```

Identical descriptive spec to `BodyText`: Carlito 9, Justify, FirstLineIndent=0, Exactly 10pt, SpaceBefore/After=0, Bold=0, Italic=0. PURPOSE block updated: 24 styles + 4 tab-stops = **28 total checks**; bucket 1 grew 15 → 16.

**Status:** Phases 1.1-1.4 applied. Awaiting user-side actions for Phases 1.5-1.7 (re-import + WordEditingConfig + RUN_TAXONOMY_STYLES verification).

#### Phases 1.5-1.7 — verified clean on test `.docm` (2026-05-01)

`WordEditingConfig` post-import: 45 styles promoted (was 44). `VerseText` lands at priority **31** as planned. Cascade verified: `Footnote Reference 31→32`, `Footnote Text 32→33`, `Psalms BOOK 33→34`, `PsalmSuperscription 34→35`, `Selah 35→36`, `PsalmAcrostic 36→37`, `SpeakerLabel 37→38`, `BodyTextIndent 38→39`. Reserved gap shifted from 39-42 to **40-43**. `EmphasisBlack 43→44`, `EmphasisRed 44→45`, `Words of Jesus 45→46`, `AuthorSectionHead 46→47`, `AuthorQuote 47→48`, `Normal 48→49`. No new missing-style warnings.

`RUN_TAXONOMY_STYLES`: **24 PASS / 4 FAIL across 28 checks**. New `VerseText` row PASSes; `BodyText` and `BodyTextIndent` rows still PASS. Audit log shows `PASS BodyText`, `PASS VerseText`, `PASS BodyTextIndent` consecutively in bucket 1.

**Phase 1 verification gate: GREEN on test `.docm`.**

**Status:** Phases 1.5-1.7 verified. Awaiting Phase 1.8 (visual sanity in Word) and Phase 1.9 (replicate on production).

#### Phase 1.8 — visual sanity PASS on test `.docm` (2026-05-01)

User confirmed: `VerseText` listed in Styles pane; no paragraph yet using it (Phase 2 will do the conversion). Phase 1's promise — "define only, no reassignment" — held.

#### Pre-Phase-2 Selah investigation (2026-05-01)

User flagged a potential edge case: `Selah` is a style applied to text in some Psalms, and "it is still part of the verse." Worth understanding before Phase 2 locks in the conversion rule.

**Finding:** `Selah` is a **character style**, not a paragraph style. Dump confirms:

```
'--- Selah  (Type=Character, Priority=35) ---
.BaseStyle = "Default Paragraph Font"
.Font.SmallCaps = -1
```

That changes the analysis:

- **Typical case:** verse paragraph contains a `Selah` character run inside it. Paragraph style is `BodyText`; first character is `Chapter Verse marker`; paragraph qualifies for Phase 2 conversion. The `Selah` character run inside the paragraph is preserved by paragraph-style reassignment (character styles overlay paragraph styles).
- **Edge cases (unknown without survey):**
  - Selah-only paragraphs (line containing just "Selah" with character style on the word). First character would be `Selah`, not `Chapter Verse marker` → fails Phase 2 rule → stays as `BodyText`. Policy question: should this be `VerseText` too?
  - Paragraph continuations or other non-standard verse-paragraph shapes containing Selah runs.

**Phase 1.9 PAUSED** until Selah usage is surveyed and policy decision made (if needed). Phase 1 source edits remain in place; production import deferred until policy is settled.

#### `AuditSelahUsage` diagnostic — APPLIED 2026-05-01

`src/basVerseStructureAudit.bas` — new public Sub `AuditSelahUsage` plus private helper `WriteSelahUsageFile`. Read-only walk of the main story:

- Locates every `Selah` character-style run via `Range.Find` with `.style = oDoc.Styles("Selah")`.
- For each run reports: paragraph start offset, paragraph style, first-character style, Phase-2 conversion verdict (`CONVERT` / `KEEP-AS-<style>`), Selah position within the paragraph (`START` / `MID` / `END` based on offset), short text excerpt (80 chars).
- Flags `POLICY DECISION` candidates: any `BodyText` paragraph containing a Selah run that the Phase 2 rule would NOT convert.
- Summary: total runs / CONVERT count / KEEP-AS-other count / policy-flag count.

Output: Immediate window + `rpt\SelahUsageAudit.txt`. Default `bWriteFile = True` matches the project's audit-Sub convention.

**Status:** `AuditSelahUsage` Sub applied. Awaiting user-side action: re-import `basVerseStructureAudit.bas` into the test `.docm`, run `AuditSelahUsage`, paste summary section. Result decides whether Phase 2 conversion rule needs extension before Phase 1.9 production import.

#### `AuditSelahUsage` — verified 2026-05-01

User ran on test `.docm`. Summary:

```
Total Selah character runs: 76
  CONVERT (verse paragraph, Phase 2 will reassign to VerseText): 76
  KEEP-AS-other (paragraph not caught by Phase 2 rule): 0
  Policy decision flags (BodyText paragraph not converted): 0
```

**All 76 Selah runs live inside verse paragraphs.** Zero edge cases. Phase 2 conversion rule needs no adjustment — when the paragraph reassigns from `BodyText` to `VerseText`, the embedded `Selah` character run survives because character styles overlay paragraph styles.

Sample run trace confirms the structural pattern:

```
Run #76 | ParaStart=3143000 | Style=BodyText | first-char-style=Chapter Verse marker | Phase2: CONVERT
  Selah at END of paragraph (offset 167 of 173)
  Excerpt: "...for the salvation of your people, for the salvation of your ano..."
```

Selah runs typically sit at the END of the verse paragraph (visible in the offset metric — 167 of 173 chars in the example), matching the textual convention.

**Phase 1.9 UNBLOCKED.** Phase 2 design is settled.

#### Phase 1.9 — replicate Phase 1 on production `.docm` (next user action)

Re-import the same source modules (`basFixDocxRoutines.bas`, `basTEST_aeBibleConfig.bas`, `basVerseStructureAudit.bas`) into the production `.docm`'s VBA project, then:

1. `DefineVerseTextStyle` — create the style.
2. `DumpStyleProperties "VerseText", True` — verify byte-identical to BodyText.
3. `WordEditingConfig` — promote (45 styles, VerseText at priority 31).
4. `RUN_TAXONOMY_STYLES` — expect 24 PASS / 4 FAIL across 28 checks.
5. `AuditSelahUsage` — expect same 76/76/0/0 result (production should mirror the test copy's structure).

Once production is at the post-Phase-1 baseline, we proceed to **Phase 2.1** (backup the production `.docm` before bulk conversion).

**Status:** Phase 1.9 ready to execute on production.

#### Phase 2.0 — pre-Phase-2 audits (additional preconditions identified 2026-05-01)

User identified four issues during Phase 1.8 visual inspection that need resolution before Phase 2 conversion locks in:

1. **Orphan `BodyText` paragraphs** — verse continuations created by stray paragraph marks mid-verse. These would NOT be converted by Phase 2 (no `Chapter Verse marker` first-char) but semantically belong to the preceding verse. **Document data defect.**
2. **`EmphasisBlack` character runs** — applied to text inside verse paragraphs.
3. **`EmphasisRed` character runs** — same pattern.
4. **`Words of Jesus` character runs** — same pattern.

For #2-4, the analysis is identical to Selah's: character-style runs inside verse paragraphs are preserved by Phase 2's paragraph-style reassignment, since character styles overlay paragraph styles in Word. Provided all instances live inside convertible verse paragraphs, no rule change is needed. Need to *survey* to confirm.

For #1, the orphan paragraphs are structural defects in the document content. Phase 2's rule cannot help — the orphan doesn't begin with `Chapter Verse marker`, so it stays as `BodyText`. Repair must happen at the data layer (manually merge each orphan back into its preceding paragraph).

#### Plan — re-sequenced 2026-05-01 (Step 3 first per user direction)

| Step | Action | Edit target / actor | What |
|---|---|---|---|
| 3 | Add `AuditOrphanBodyTextParagraphs` Sub | `src/basVerseStructureAudit.bas` | Discovery diagnostic for orphan continuations |
| 4 | Run audit on test `.docm` | (user action) | Capture orphan list |
| 5 | Manually merge orphans | `.docm` content | Delete stray paragraph marks; repeat per book |
| 6 | Re-run `AuditOrphanBodyTextParagraphs` | (user action) | Confirm 0 orphans |
| 7 | Generalize `AuditSelahUsage` → `AuditCharStyleUsage(styleName)` | `src/basVerseStructureAudit.bas` | Parameterize the Selah Sub |
| 8 | Run for all 4 character styles | (user action) | Verify Selah / EmphasisBlack / EmphasisRed / Words of Jesus |
| 9 | Resume Phase 1.9 (production replication) | (user action) | Same as before — define style + audit row + WordEditingConfig + RUN_TAXONOMY_STYLES |
| 10 | Phase 2 (backup + bulk conversion) | (later) | Original Phase 2 plan unchanged |

**Reasoning for Step 3 first:** orphan repair changes paragraph structure (deletes paragraph marks, merging two paragraphs into one). Character-style audits run after orphan repair will reflect the cleaned structure — character runs that currently span an orphan boundary may end up entirely inside a single verse paragraph after repair. Doing the orphan audit first avoids re-running the character audits.

#### Step 3 — `AuditOrphanBodyTextParagraphs` APPLIED 2026-05-01

`src/basVerseStructureAudit.bas` — new public Sub `AuditOrphanBodyTextParagraphs` plus private helper `WriteOrphanFile`. Read-only walk of the main story:

- Tracks current book via last `Heading 1` paragraph; tracks `seenFirstH2InCurrentBook` flag (set on first H2 of each book; reset on next H1).
- For each `BodyText` paragraph **inside the verse-bearing region** (after first H2 of book, before next H1) whose first character is **not** `Chapter Verse marker`, reports it.
- Per-orphan fields: book name, ParaStart, first-character style name (or `(empty)` for empty paragraphs), size category (EMPTY / SHORT < 30 chars / LONG >= 30 chars), 80-char excerpt.
- Empty paragraphs are reported separately because they're typically accidental extra Enter keystrokes rather than real content orphans.

**Excluded from audit (legitimate non-verse BodyText):**
- Front matter (before first H1).
- Book introductions (between H1 and first H2 of each book).

Output: Immediate window + `rpt\OrphanBodyTextAudit.txt`.

**Status:** Step 3 code applied. Awaiting user-side action: re-import `basVerseStructureAudit.bas`, run `AuditOrphanBodyTextParagraphs` on test `.docm`, paste summary block.

#### Step 3 — initial run too greedy (2026-05-01)

User ran on test `.docm`. Result: **328 orphan candidates** (80 EMPTY + 130 SHORT + 118 LONG). Inspection of the trailing entries showed mostly false positives — book-end content (e.g., "Introducing the Divine Principle" appearing in REVELATION region after the last verse) and whitespace-only paragraphs.

**Root cause:** the v1 rule "BodyText after first H2 of book without CVM first-char" was too greedy. It correctly excluded *book intros* (before first H2) but did not exclude:

1. **Chapter intros** — BodyText paragraphs between an H2 (chapter heading) and the first actual verse paragraph of that chapter.
2. **Chapter-end content** — BodyText paragraphs after the last verse of a chapter but before the next H2 (or H1, or end of document).

User's structural insight: "Each page with Heading 1 has text that is not VerseText and is rightfully BodyText. It is always the next page that starts verse 1." That implied the gap between H1/H2 and the first verse paragraph can contain legitimate non-verse BodyText.

#### Step 3 — refined algorithm (APPLIED 2026-05-02)

Refined detection rule: **a BodyText paragraph is an orphan only if it sits BETWEEN two verse paragraphs in the same chapter.**

Single-pass algorithm with a 1000-element buffer:

| Event | Action |
|---|---|
| H1 | Discard pending buffer (post-last-verse of previous book); reset state; `seenFirstVerseInCurrentChapter = False` |
| H2 | Discard pending buffer (post-last-verse of previous chapter); `seenFirstVerseInCurrentChapter = False` |
| BodyText with CVM (verse) | If `seenFirstVerseInCurrentChapter`, **flush buffer as confirmed orphans**; set `seenFirstVerseInCurrentChapter = True`; clear buffer |
| BodyText without CVM, after first verse of chapter | Add to buffer (suspended judgement) |
| BodyText without CVM, before first verse of chapter | Increment `chapterIntroCount` (excluded) |
| End of document | Discard remaining buffer (post-last-verse-of-last-book) |

Side counters in the summary report show:
- `Chapter intros (before first verse of chapter)`: count of paragraphs excluded as legitimate non-verse content between H2 and first verse.
- `Chapter-end content (after last verse, before next H2/H1)`: count of paragraphs excluded as legitimate non-verse content trailing the last verse of a chapter.

This gives noise-vs-signal visibility: out of all non-CVM BodyText paragraphs in book regions, the report shows how many are excluded vs flagged.

**Implementation notes:**

- Buffer cap is 1000 paragraphs between two verses — a sane upper bound. Overflow is silently dropped (vanishingly unlikely scenario).
- Buffer columns: `(ParaStart, firstCharStyle, sizeCat, paraLen, excerpt)` — five strings per row.
- On flush, sizeCat (`EMPTY`/`SHORT`/`LONG`) is mapped to the appropriate counter and rendered as a label like `SHORT (15 chars)`.

**Status:** Step 3 refined Sub applied 2026-05-02. Awaiting user-side action: re-import `basVerseStructureAudit.bas`, re-run `AuditOrphanBodyTextParagraphs` on test `.docm`, paste new summary block. Decision tree:
- **0 orphans** → proceed to Step 7 (parameterize char-style audit).
- **Small count (<50)** → manually repair, re-audit, then Step 7.
- **Still hundreds** → structural investigation needed.

### 3. `AuditOneStyle` — extend for character-style properties

Currently `AuditOneStyle` only checks paragraph-style properties (font name/size, alignment, indent, line spacing, space before/after). Character styles like `Footnote Reference` are parked in bucket 2 (existence-verified) because the audit cannot meaningfully fully-specify them — Bold, Italic, Color are not in the check list.

**Required for:** `Footnote Reference` to graduate from bucket 2 to bucket 1.

**Scope:** add ~3 optional property arguments to `AuditOneStyle` (`bExpBold`, `bExpItalic`, `lExpColor`) with sentinels (e.g. `-2` for skip on Bold/Italic, `-1` for skip on Color since `wdColorAutomatic = -16777216` is a real value). Or split into a sibling `AuditOneCharStyle` with character-style-relevant checks only — same pattern as `AuditStyleTabs` (Phase 6c).

Original analysis: `rvw/Code_review 2026-04-25.md` § "Footnote Reference — deferred to bucket 2 (Q2 decision)" (2026-04-29).

### 4. Prescriptive-spec pass for known QA-checklist drift

The current taxonomy audit encodes **descriptive** specs (capture today's values). Several known QA-checklist violations are tolerated rather than driven to correction:

**`LineSpacingRule = Exactly` on paragraph styles** (QA checklist preference is `Single`):
- `Heading 2` — `Exactly 10`
- `CustomParaAfterH1` — `Exactly 10`
- `Brief` — `Exactly 9.5`
- `Psalms BOOK` — `Exactly 10`
- `Footnote Text` — `Exactly 8`

**`BaseStyle = "Normal"` on approved styles** (QA checklist preference is `""`):
- `CustomParaAfterH1`, `Brief`, `Footnote Text`, `Psalms BOOK`, `PsalmSuperscription`, `PsalmAcrostic`

**`AuthorListItem` FirstLineIndent drift:** ~~expected `0` (audit), live `-18` (hanging). Currently FAILs on every `RUN_TAXONOMY_STYLES`. 2026-04-29 leave-as-is decision: keep the FAIL as a tracked indicator. Revisit during the prescriptive pass.~~ **CLOSED 2026-05-01** by promoting `AuthorListItem` to bucket 1 with descriptive spec `FirstLineIndent = -18`. The drift was a documented design intent all along; descriptive promotion encodes the live state and clears the FAIL. See "Author* trio promoted to bucket 1" closure record below.

Original analysis: `rvw/Code_review 2026-04-25.md` § "Spec promotion: descriptive vs prescriptive (decision)" (2026-04-29) and § "Section (B) full inventory findings" (2026-04-29).

**Recommendation:** treat as a series of one-property-at-a-time decisions, each tracked as its own review item with rationale. Not a single batch.

### 5. Taxonomy audit — full-coverage final-state goal

Documented in `EDSG/01-styles.md` callout. Current state: 9 fully-specified + 8 existence-verified + 3 not-yet-created + 1 tab-stop = transitional toward "every approved style mapped with real specs".

Each move from bucket 2 → bucket 1 (when descriptive specs are encoded for an existence-verified style) is a measurable step. SpeakerLabel, Heading 1, Heading 2, etc. that are already promoted reduce the bucket-2 count further; the remaining bucket-2 styles (`BookIntro`, `ListItemTab` legacy slot, `TheHeaders`, `TheFooters`, `Title`, `Footnote Reference`, `AuthorListItemTab` placeholder before tabs were audited) are the candidates.

Original analysis: `rvw/Code_review 2026-04-25.md` § "Important — taxonomy audit final-state goal" callout (2026-04-29).

### 6. Finding 5 (ribbon nav) — umbrella OPEN

Fix (A) resolved the primary caret-not-visible symptom. The residual is the **idle-commit focus leak**: Word's customUI `editBox` auto-commits on idle (~1 s) and returns focus to the document body, so any subsequent Tab is a document Tab.

**Status:** **WORD LIMITATION, NO VBA-SIDE FIX AVAILABLE.** Documented in the prior review. KeyTips are the supported Office UX path for cross-control jumps and bypass Tab entirely.

**Forward options (deferred):**

- **Auto-fire Go on valid `(book, chapter, verse)` triple** — code-side option to remove the final Tab → Go step. Tradeoff: nav fires before user expects it; would need a `bAutoFire` toggle.
- **VSTO/WPF ribbon rewrite** — only path to true ribbon-owned focus management. Major rewrite; deferred indefinitely.

Original analysis: `rvw/Code_review 2026-04-25.md` § "Finding 5 — terminology correction" (2026-04-29).

### 7. Optional EDSG documentation refresh

Minor consistency items noticed during the migration work:

- **`EDSG/01-styles.md` "Missing from document" list** still lists `BookIntro` (not in document but kept as a tracking placeholder) — accurate.
- **`EDSG/02-editing-process.md`** Stage 1 step references could mention `AuthorListItem*` as the canonical example for the `BaseStyle = ""` rule (currently uses generic phrasing).
- **`EDSG/04-qa-workflow.md`** "Current state" section dated 2026-04-26 still mentions priorities 38-41 reserved gap and the 43-styles count — superseded by the 2026-04-29 SpeakerLabel insertion (now 39-42 reserved, 44 styles). Documentation lag, not a blocker.

**Recommendation:** opportunistic update next time these pages are visited for substantive edits.

### 8. Body-content number prefixes — keep manual, no docvariables (decision 2026-04-30)

User considered replacing manual text prefixes (`"1. "`, `"2. "`, …) on `AuthorBookRef` paragraphs with Word doc-variable fields for future i18n flexibility. **Decision: keep manual text prefixes. Revisit only if/when i18n is actively rolled out and a target locale needs non-Arabic numbering.**

#### Reasoning recorded

**Pros considered (theoretical i18n benefit):**
- Locale-aware numbering substitution (Arabic-Indic, Hebrew letters, RTL ordering).
- Programmatic renumber without text edits.
- Separation of presentation (number) from content (citation).

**Cons (practical, drove the decision):**
- Each prefix becomes a `{ DOCVARIABLE }` field — visual clutter when field codes toggled, fragile across some copy-paste targets.
- No native auto-renumber. Inserting between #5 and #6 still requires VBA renumber logic; not better than retyping literals.
- Discovery cost for future contributors and AI assistants seeing fields in the body.
- Edit overhead vs typing `"5. "`.
- No native locale routing in Word docvariables — they're document-wide single values; localization layer would have to be built anyway.
- Project's existing i18n strategy (`basUIStrings`, `check_i18n.py`) targets ribbon UI labels and status-bar messages, **not body content**. Body-content i18n is a different problem class.
- Number prefixes are mostly language-neutral — Western Arabic digits are the academic-citation standard across most plausible target locales. Tiny punctuation variants (`1.` vs `1)`) are easier handled by a one-pass VBA reformatter than by a docvariable layer.

**Better tool when i18n becomes active:** locale-aware VBA prefix generator that switches on `Document.LanguageID` or a project-level locale setting and writes prefix text on demand. Decouples from Word's field machinery.

**Where docvariables ARE the right tool (different use case):** single-string substitutions appearing once or in known template positions (e.g., a "Bibliography" heading, dates, version strings). Number prefixes don't fit that pattern.

#### Trigger to revisit

Active i18n rollout where a target locale requires:
- Non-Arabic numerals in body content, OR
- Substantially different prefix punctuation that can't be handled by a one-pass reformatter, OR
- Per-paragraph content substitution that today's manual prefixes can't accommodate.

Until then, manual text prefixes stand.

### 9. Session manifest

`sync/session_manifest.txt` written 2026-04-30 covering this arc's src/ edits. Use as the import checklist for any fresh `.docm` re-sync. Final-state verification commands listed at the end of the manifest.

## Pointer back to the closed arc

Full dated history of the work that produced this carry-forward state is in
[`rvw/Code_review 2026-04-25.md`](Code_review%202026-04-25.md). That file
includes:

- Phase 0 through Phase 6 of the List Paragraph migration with all
  decisions, pre-flight checks, and verifications recorded.
- The descriptive-vs-prescriptive decision framework.
- The Romans doxology TR-vs-WEB clarification.
- The Word `customUI` focus-handling analysis (Finding 5).
- The terminology correction from "Tab race" to "idle-commit focus leak".

Anything in this 2026-04-30 file should reference back to that arc
for the *why*; this file holds only the **what is still open**.
