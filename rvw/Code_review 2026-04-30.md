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
- **Ribbon focus fix (Finding 5 fix A).** Caret renders correctly on
  first nav from a ribbon-owned event handler.
- **`ToSBLShortForm` "Song of Songs" lookup error (Finding 3).** No
  defect reproducible from static analysis; live-run confirms alias
  path sound. Closed 2026-04-29.

## Open carry-forward items

### 1. Reference rename sweep — `Solomon` → `Song of Songs`

Confirmed open as of 2026-04-30. Four files still contain the older
`Solomon` form:

| File | Line | Context |
|---|---|---|
| `src/basTEST_aeBibleCitationClass.bas` | 885 | Test oracle: `canonNames(22) = "Solomon"`. Should be `"Song of Songs"` to match the citation class's canonical form. |
| `src/basSBL_VerseCountsGenerator.bas` | 95 | Context label string passed to `ToOneBasedLongArray`. Cosmetic — does not affect data. |
| `src/XbasTESTaeBibleDOCVARIABLE.bas` | 527 | `VerifyBookNameFromDocVariable "Song", "Solomon"` — document-specific assertion. Earlier review (2026-03-16) marked this as "correct for this document"; verify if the assertion is still right after the canonical rename. |
| `md/Deterministic Structural Parser.md` | 83 | Reference table row: `Solomon`. Update to `Song of Songs`. |
| `md/Deterministic Structural Parser.md` | 314 | Multi-word example: `"Song of Solomon"`. Decide whether to leave (it's an alias) or rename to `"Song of Songs"`. |

Original analysis: `rvw/Code_review 2026-04-25.md` § "Finding 4 — broader citation-code impact of the rename" (2026-04-28).

**Recommendation:** small batched edit pass when convenient. Risk: low (none of these affect runtime behaviour except potentially the test oracle, which the audit has not flagged). Each rename should preserve identifier casing per the project's normalization rule.

### 2. `AuditOneStyle` — extend for character-style properties

Currently `AuditOneStyle` only checks paragraph-style properties (font name/size, alignment, indent, line spacing, space before/after). Character styles like `Footnote Reference` are parked in bucket 2 (existence-verified) because the audit cannot meaningfully fully-specify them — Bold, Italic, Color are not in the check list.

**Required for:** `Footnote Reference` to graduate from bucket 2 to bucket 1.

**Scope:** add ~3 optional property arguments to `AuditOneStyle` (`bExpBold`, `bExpItalic`, `lExpColor`) with sentinels (e.g. `-2` for skip on Bold/Italic, `-1` for skip on Color since `wdColorAutomatic = -16777216` is a real value). Or split into a sibling `AuditOneCharStyle` with character-style-relevant checks only — same pattern as `AuditStyleTabs` (Phase 6c).

Original analysis: `rvw/Code_review 2026-04-25.md` § "Footnote Reference — deferred to bucket 2 (Q2 decision)" (2026-04-29).

### 3. Prescriptive-spec pass for known QA-checklist drift

The current taxonomy audit encodes **descriptive** specs (capture today's values). Several known QA-checklist violations are tolerated rather than driven to correction:

**`LineSpacingRule = Exactly` on paragraph styles** (QA checklist preference is `Single`):
- `Heading 2` — `Exactly 10`
- `CustomParaAfterH1` — `Exactly 10`
- `Brief` — `Exactly 9.5`
- `Psalms BOOK` — `Exactly 10`
- `Footnote Text` — `Exactly 8`

**`BaseStyle = "Normal"` on approved styles** (QA checklist preference is `""`):
- `CustomParaAfterH1`, `Brief`, `Footnote Text`, `Psalms BOOK`, `PsalmSuperscription`, `PsalmAcrostic`

**`AuthorListItem` FirstLineIndent drift:** expected `0` (audit), live `-18` (hanging). Currently FAILs on every `RUN_TAXONOMY_STYLES`. 2026-04-29 leave-as-is decision: keep the FAIL as a tracked indicator. Revisit during the prescriptive pass.

Original analysis: `rvw/Code_review 2026-04-25.md` § "Spec promotion: descriptive vs prescriptive (decision)" (2026-04-29) and § "Section (B) full inventory findings" (2026-04-29).

**Recommendation:** treat as a series of one-property-at-a-time decisions, each tracked as its own review item with rationale. Not a single batch.

### 4. Taxonomy audit — full-coverage final-state goal

Documented in `EDSG/01-styles.md` callout. Current state: 9 fully-specified + 8 existence-verified + 3 not-yet-created + 1 tab-stop = transitional toward "every approved style mapped with real specs".

Each move from bucket 2 → bucket 1 (when descriptive specs are encoded for an existence-verified style) is a measurable step. SpeakerLabel, Heading 1, Heading 2, etc. that are already promoted reduce the bucket-2 count further; the remaining bucket-2 styles (`BookIntro`, `ListItemTab` legacy slot, `TheHeaders`, `TheFooters`, `Title`, `Footnote Reference`, `AuthorListItemTab` placeholder before tabs were audited) are the candidates.

Original analysis: `rvw/Code_review 2026-04-25.md` § "Important — taxonomy audit final-state goal" callout (2026-04-29).

### 5. Finding 5 (ribbon nav) — umbrella OPEN

Fix (A) resolved the primary caret-not-visible symptom. The residual is the **idle-commit focus leak**: Word's customUI `editBox` auto-commits on idle (~1 s) and returns focus to the document body, so any subsequent Tab is a document Tab.

**Status:** **WORD LIMITATION, NO VBA-SIDE FIX AVAILABLE.** Documented in the prior review. KeyTips are the supported Office UX path for cross-control jumps and bypass Tab entirely.

**Forward options (deferred):**

- **Auto-fire Go on valid `(book, chapter, verse)` triple** — code-side option to remove the final Tab → Go step. Tradeoff: nav fires before user expects it; would need a `bAutoFire` toggle.
- **VSTO/WPF ribbon rewrite** — only path to true ribbon-owned focus management. Major rewrite; deferred indefinitely.

Original analysis: `rvw/Code_review 2026-04-25.md` § "Finding 5 — terminology correction" (2026-04-29).

### 6. Optional EDSG documentation refresh

Minor consistency items noticed during the migration work:

- **`EDSG/01-styles.md` "Missing from document" list** still lists `BookIntro` (not in document but kept as a tracking placeholder) — accurate.
- **`EDSG/02-editing-process.md`** Stage 1 step references could mention `AuthorListItem*` as the canonical example for the `BaseStyle = ""` rule (currently uses generic phrasing).
- **`EDSG/04-qa-workflow.md`** "Current state" section dated 2026-04-26 still mentions priorities 38-41 reserved gap and the 43-styles count — superseded by the 2026-04-29 SpeakerLabel insertion (now 39-42 reserved, 44 styles). Documentation lag, not a blocker.

**Recommendation:** opportunistic update next time these pages are visited for substantive edits.

### 7. Body-content number prefixes — keep manual, no docvariables (decision 2026-04-30)

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

### 8. Session manifest

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
