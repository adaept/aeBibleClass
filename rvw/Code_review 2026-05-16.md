# Code review - 2026-05-16 carry-forward

This file opens a fresh review arc on 2026-05-16. The previous arc
[`rvw/Code_review 2026-05-15.md`](Code_review%202026-05-15.md) is
now **closed for new work**; that file remains the authoritative
dated history for everything between 2026-05-15 and 2026-05-16,
including:

- **Test 5 added** — `CountApprovedStylesInGallery`. Enforces the
  editorial policy that approved styles must not appear in the
  Styles ribbon gallery. Closes the QuickStyle audit gap for the
  approved cohort (Test 45 covers the non-approved cohort).
- **LineSpacingRule prescriptive pass closed.** `CustomParaAfterH1`
  fixed to Single (rule 0, LineSpacing 12); `Footnote Text`
  retained at Exactly 8 as a known exception (i18n-flagged).
  Taxonomy in `basTEST_aeBibleConfig.bas` updated.
- **Test 22 split into 22 / 38 / 74.** The bundled
  `CountEmptyParagraphsWithFormatting` was decomposed into three
  disjoint detectors after a step-back analysis surfaced that
  Test 38 already covered the bare-empty case as a subset:
  - Test 38 (`CountEmptyParagraphs`) - bare empty, expected
    rebaselined 153 -> 216.
  - Test 22 (`CountWhitespacePaddedEmptyParagraphs`) - spaces
    around the pilcrow only, expected 0.
  - Test 74 (`CountEmptyParagraphsWithInlineContent`, new) -
    integrity check for visually-empty paragraphs carrying
    InlineShapes / Fields / Bookmarks, expected 0.
  Combined runtime 7.6s vs prior single-test 390s.
- **Slot 5 retire + Slot 6 upgrade.** Empty-paragraph discipline
  tightened via path-(b) paragraph walk with three break-exception
  predicates. Operator-verification trail confirmed
  `wdActiveEndAdjustedPageNumber` is the right page index for
  hint emission.

Status tag legend (continued):

- **OPEN** - actively pending, all known prerequisites met.
- **PARTIAL** - partially complete; specific remaining work listed.
- **DEFERRED** - not started, waiting on a specific trigger.
- **FUTURE** - speculative; revisit only when conditions warrant.
- **RECOVERED** - surfaced from a prior arc where it was dropped
  off the carry-forward chain.

## Open carry-forward (priority order)

### 1. Run aeRibbon Gates G1-G8 and ship v1.0.0 (HIGH) - OPEN

The production export gateway is in place; nothing has been built
or gated yet. **Next active release-track item** and gates the
hand-off to the author for comments-only review.

Original analysis and gate definitions: see prior arcs via the
2026-05-15 carry-forward.

### 2. Item 13 remaining work - built-in hide-sweep + test wiring (MEDIUM) - PARTIAL

Pass 1 closed 2026-05-14 (`AuditNonPaletteStyleColors` added,
custom-style anomaly count brought to 0).

Remaining work:

- **2.1 Hide-sweep for Word built-in noise (MEDIUM).** 122+
  built-in styles bypass the audit because they're not under
  editorial control. After the 2026-05-15 BookHyperlink work, the
  built-in `Hyperlink` and `FollowedHyperlink` styles also belong
  in the sweep - neither is in use after the refactor and
  authors must not be able to pick them from the gallery and
  reintroduce the inheritance bug. Test 5
  (`CountApprovedStylesInGallery`) added 2026-05-16 protects the
  approved cohort; Test 45 (`CountUnapprovedVisibleStyles`)
  protects the non-approved cohort; the missing piece is the
  maintenance sweep that gets the built-ins into the
  "properly hidden" state Test 45 enforces.

Full prior analysis: see § 2 in
[`Code_review 2026-05-15.md`](Code_review%202026-05-15.md).

### 3. TestReport.txt - per-slot one-line descriptions (MEDIUM)

Carrying forward from 2026-05-15 § 10. Goal: each PASS/FAIL row
in `rpt/TestReport.txt` self-explanatory without diving into the
class.

**Progress at carry-forward (2026-05-17):** 14 of 74 slots
populated in `GetTestDescription` - slots 1, 2, 3, 4, 5, 6, 16,
17, 18, 19, 22, 28, 38, 74. Remaining ~60 slots fall through
`Case 7 To 74` to the empty-string default.

**Authoring strategy:** opportunistic per-touch (slot gets a
description when its surrounding code is edited) plus optional
dedicated authoring passes. Per-touch has been working well -
the 2026-05-16 session added 22, 28, 38, 74 organically while
splitting Test 22.

Full design rationale and emission shape: see § 10 in
[`Code_review 2026-05-15.md`](Code_review%202026-05-15.md).

### 4. Taxonomy audit - full-coverage final-state goal (LOW-MEDIUM, ASPIRATIONAL) - RECOVERED

State at last accounting: **25 fully specified + 4
existence-verified + 3 not-yet-created + 5 tab-stops verified =
37 distinct style entries across 44 checks** (+1 from BookHyperlink
add 2026-05-15).

Recommendation: a 10-minute state check via
`RUN_TAXONOMY_STYLES` will quantify how much was incidentally
closed by intervening arcs. Update the count then decide whether
the umbrella item warrants explicit attention or is well-served
by being a callout in `EDSG/01-styles.md`.

Full per-style decisions list: see § 4 in
[`Code_review 2026-05-15.md`](Code_review%202026-05-15.md).

### 5. EDSG documentation refresh - VerseText-aware (LOW) - CLOSED 2026-05-17

Investigation surfaced wider drift than the original carry-forward
summary suggested: `GetApprovedStyles()` had grown from ~44 to 52
entries, with VerseText at priority 33 (not 31 as the prior
summary said), plus `BibleIndexList`, `AuthorBookSections`,
`BookHyperlink`, `ParallelHeader`, `ParallelText`,
`SpeakerLabel`, `BodyTextContinuation`, `AppendixTitle`,
`AppendixBody` all added; `BodyTextIndent`, `AuthorQuote`,
`BookIntro` removed.

**Refreshed (one-shot rebuild from the SSOT in
`basTEST_aeBibleConfig.bas`):**

- `EDSG/01-styles.md` — priority table rebuilt as a single
  unified 52-row table with notes column; "Pending re-validation"
  / "Reserved gaps" framing retired (gaps were filled); category
  prose updated (`Body text`, `Author commentary`, `Anchor`,
  `Special book treatments`).
- `EDSG/04-qa-workflow.md` — "Current state" rewritten and dated
  2026-05-17; reflects 52-entry array, lists additions and
  removals since the 2026-04-26 snapshot.
- `EDSG/06-i18n.md` — `VerseText` added as the primary
  translation target; `BodyTextIndent` line removed.

Mechanical-dump script (option 2 from the investigation report)
was not built; reassess if the array starts churning again.

### 6. EDSG/02-editing-process.md - AuthorListItem* as canonical BaseStyle="" example (LOW) - OPEN

Carried independently from the 2026-05-17 EDSG refresh.
`AuthorListItem` (priority 19) was identified as the cleanest
worked example of the `BaseStyle = ""` rule. Opportunistic edit
to Stage 1 of the editing process narrative; not blocking.

### 7. Finding 5 (ribbon nav) - umbrella OPEN (DEFERRED, WORD LIMITATION) - RECOVERED

Word-side limitation; no action available. Remains in the
register for awareness.

### 8. SoftHyphenSweep_FootnotesOnly sister routine (DEFERRED)

Surfaced during the 2026-05-08 SHA build; waiting on a footnote-
specific trigger before implementation.

### 9. SHA_ReplaceHard i18n consideration (FUTURE)

Speculative; revisit when a non-English target translation
materialises.

### 10. Architecture rule - class encapsulation + module/class as casual-coder safety boundary (RULE, 2026-05-15)

Codified as a feedback memory and documented in the 2026-05-15
arc. Standing rule, not an action item - listed here so it
remains visible during slot-by-slot review work.

Full rule and worked examples: see § 9 in
[`Code_review 2026-05-15.md`](Code_review%202026-05-15.md).

## Pointer back to the closed arc

Full dated history of the work that produced this carry-forward
state is in [`rvw/Code_review 2026-05-15.md`](Code_review%202026-05-15.md).
That file (and the arcs it points back to) covers:

- The BookHyperlink design, implementation, and verification.
- The Test 5 / Test 22 split / Test 74 add sequence.
- The Slot 5 retire + Slot 6 upgrade arc with operator-verification
  trail.
- The LineSpacingRule prescriptive pass closure.
