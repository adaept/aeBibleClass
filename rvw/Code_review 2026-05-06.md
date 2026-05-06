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

### 7. Session manifest

`sync/session_manifest.txt` from the prior arc covered src/ edits
through 2026-04-30. The 2026-05-01 → 2026-05-06 VerseText, orphan-audit,
and char-style-audit edits should be recorded in a fresh manifest entry
(or the existing one updated) as the import checklist for any fresh
`.docm` re-sync.

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
