# Code Review - 2026-04-20

## Carry-Forward from 2026-04-18

Continues from `rvw/Code_review 2026-04-18.md`.

---

## § 1 — Status of Previous Session (2026-04-18/19) Carry-Forward

### Completed items (closed this session)

| Item | Detail | Status |
|------|--------|--------|
| Bug #598 — C/V Prev/Next disabled after book selection | `m_currentChapter = 1` / `m_currentVerse = 1` in `OnBookChanged` | **CLOSED — 2026-04-18** |
| Bug #599 — First-load Tab goes to document | All six Prev/Next buttons always-enabled | **CLOSED — 2026-04-18** |
| GoButton — XML injection | `customUI14backupRWB.xml` inserted into `Blank Bible Copy.docm` via `inject_ribbon.py` | **CLOSED — 2026-04-18** |
| GoButton — VBA callbacks | `GetGoEnabled` / `OnGoClick` in `aeRibbonClass.cls`; wrappers in `basBibleRibbonSetup.bas` | **CLOSED — 2026-04-18** |
| Status bar — `"Navigating ..."` and SBL citation | Deferred write via `UpdateStatusBarDeferred`; persists through background layout | **CLOSED — 2026-04-18** |
| Status bar — invalid input / boundary messages | `SB_*` constants in `basUIStrings.bas`; `FormatMsg` helper | **CLOSED — 2026-04-19** |
| Pro #6 gap — range limit in messages | `m_chapterTextValid`, `m_chapterMax`, `m_verseTextValid`, `m_verseMax`; range in all boundary/error messages | **CLOSED — 2026-04-19** |
| `basRibbonStrings` → `basUIStrings` rename | `git mv`; `Attribute VB_Name` updated; all references updated | **CLOSED — 2026-04-19** |
| `SB_NAVIGATING` / `SB_WARM_CACHE` i18n extraction | Both inline strings added to `basUIStrings.bas` | **CLOSED — 2026-04-19** |
| `KT_GO` constant — keytip consistency gap | Added to `basUIStrings.bas`; `GetGoKeytip` in `basBibleRibbonSetup.bas`; XML `keytip="G"` → `getKeytip="GetGoKeytip"` | **CLOSED — 2026-04-20** |
| `KT_NEW_SEARCH` → `KT_SEARCH` rename | Renamed in `basUIStrings.bas`, `basBibleRibbonSetup.bas`, `normalize_vba.py` | **CLOSED — 2026-04-20** |
| Ribbon XML re-injected | `inject_ribbon.py` run after `getKeytip="GetGoKeytip"` change; `Blank Bible Copy.docm` updated | **CLOSED — 2026-04-20** |
| `Ribbon Design.md` — state machine diagram | Mermaid `stateDiagram-v2` + state table + transition rules added as new section | **CLOSED — 2026-04-20** |
| `Ribbon Design.md` — verse section phrasing | Parallel with chapter: "Press Go to navigate, or use ◀ ▶ to step through verses immediately" | **CLOSED — 2026-04-20** |
| `Ribbon Design.md` — dropdown overpromise | Softened: type-first is primary path; dropdown noted as not yet populated | **CLOSED — 2026-04-20** |
| `Ribbon Design.md` — typing vs. navigation rule | Added: "Typing values sets the navigation target; navigation only occurs when Go or a Prev/Next button is used." | **CLOSED — 2026-04-20** |
| Code cleanup — `aeRibbonClass.cls` | Step labels removed; dead methods archived at bottom | **CLOSED — 2026-04-19** |
| Code cleanup — `basBibleRibbonSetup.bas` | Module header; archived callbacks section; test helpers section | **CLOSED — 2026-04-19** |
| Code cleanup — `basRibbonDeferred.bas` | Active vs. archived deferred entry points separated | **CLOSED — 2026-04-19** |
| `aeBibleCitationClass.cls` hardening | GALATIANS typo fixed; Solomon documented; service contract comments; 66-book self-test added | **CLOSED — 2026-04-19** |
| `normalize_vba.py` — `Count` standalone | `\bCount\b` entry added | **CLOSED — 2026-04-19** |

### Open items (carry-forward)

| Item | Detail | Status |
|------|--------|--------|
| Bug #597 | New Search should focus `cmbBook` — Option A/B/C documented; awaiting decision | **OPEN** |
| Bug 16 | Keytip badges end-to-end test — re-test after `GetGoKeytip` injection | **PENDING** |
| Bug 22 / 23a | First-nav layout delay (~6–17s) | **KNOWN LIMITATION** |
| Bug 27 | Enter in Chapter does not navigate | **SUPERSEDED — absorbed by GoButton** |
| Step 7 | OLD_CODE cleanup — dead stubs in `aeRibbonClass.cls` | **PENDING** |
| WarmLayoutCache rewrite | Replace `Range.Select` with `ScrollIntoView`; re-enable deferred warm | **FUTURE** |
| Search tracking reset | Test `Selection.SetRange` from `OnTime` context | **FUTURE** |
| Import modules | `aeRibbonClass.cls`, `basBibleRibbonSetup.bas`, `basRibbonDeferred.bas`, `basUIStrings.bas` all modified — must be imported into VBA project | **PENDING** |

---

## § 2 — Architecture: Navigation State Machine

Formalized 2026-04-20. The ribbon controller is a four-state machine:

| State | Book row | Chapter row | Verse row | Go |
|---|---|---|---|---|
| `NoSelection` | active | — | — | — |
| `BookSelected` | active | active | — | — |
| `ChapterSelected` | active | active | active | active |
| `VerseSelected` | active | active | active | active |

**Transition rules (from `Ribbon Design.md` state machine section):**

- State advances on valid input in each field; invalid input stays in current state (error in status bar).
- Prev/Next buttons navigate immediately and stay in the same state.
- Go navigates and stays in the same state.
- New Search resets to `NoSelection` from any state.
- State never advances past `VerseSelected`.

**Key private state fields in `aeRibbonClass.cls`:**

| Field | Type | Purpose |
|-------|------|---------|
| `m_currentBookIndex` | `Long` | Current book (1–66); 0 = no selection |
| `m_currentChapter` | `Long` | Current chapter; 0 = no selection |
| `m_currentVerse` | `Long` | Current verse; 0 = no selection |
| `m_bookTextValid` | `Boolean` | False if last `OnBookChanged` input was unrecognised |
| `m_chapterTextValid` | `Boolean` | False if last `OnChapterChanged` input was out of range |
| `m_chapterMax` | `Long` | Max chapter stored at rejection time for error message |
| `m_verseTextValid` | `Boolean` | False if last `OnVerseChanged` input was out of range |
| `m_verseMax` | `Long` | Max verse stored at rejection time for error message |

---

## § 3 — Architecture: Navigation Rules (from § 28 / 2026-04-13)

| Rule | Description |
|------|-------------|
| 1 | Navigation requires all three fields (Book, Chapter, Verse) to be filled |
| 2 | Book is always required — no default |
| 2a | When Book is confirmed, Chapter and Verse are immediately set to 1 |
| 3 | Tab past Chapter accepts the displayed value |
| 4 | Tab past Verse accepts the displayed value |
| 5 | Navigation fires only from Go or Prev/Next buttons — never implicitly from `onChange` |
| 6 | Prev/Next buttons are always enabled; click handlers guard boundaries and write status bar messages |
| 7 | Prev/Next button presses update all three B/C/V fields |

---

## § 4 — i18n Architecture (basUIStrings.bas)

All user-facing strings live in `basUIStrings.bas` as `Public Const` values.

**KeyTip constants:**

| Constant | Value | Control |
|----------|-------|---------|
| `KT_BOOK` | `"B"` | Book comboBox |
| `KT_CHAPTER` | `"C"` | Chapter comboBox |
| `KT_VERSE` | `"V"` | Verse comboBox |
| `KT_PREV_BOOK` | `"["` | Previous Book button |
| `KT_NEXT_BOOK` | `"]"` | Next Book button |
| `KT_PREV_CHAPTER` | `","` | Previous Chapter button |
| `KT_NEXT_CHAPTER` | `"."` | Next Chapter button |
| `KT_PREV_VERSE` | `"<"` | Previous Verse button |
| `KT_NEXT_VERSE` | `">"` | Next Verse button |
| `KT_GO` | `"G"` | Go (navigate) button |
| `KT_SEARCH` | `"S"` | New Search button |
| `KT_ABOUT` | `"A"` | About (adaept) button |

Note: `KT_NEW_SEARCH` renamed to `KT_SEARCH` — aligns with keytip convention (first letter of action).

**Status bar constants (static — use directly):**

| Constant | Value |
|----------|-------|
| `SB_NAVIGATING` | `"Navigating ..."` |
| `SB_WARM_CACHE` | `"Bible: building navigation index..."` |
| `SB_INVALID_BOOK` | `"Invalid input for Book - enter a book name or abbreviation"` |
| `SB_ALREADY_FIRST_BOOK` | `"Already at first book"` |
| `SB_ALREADY_LAST_BOOK` | `"Already at last book"` |

**Status bar constants (dynamic — use with `FormatMsg`):**

| Constant | Template |
|----------|---------|
| `SB_INVALID_CHAPTER` | `"Invalid input for Chapter - out of range (1-{0})"` |
| `SB_INVALID_VERSE` | `"Invalid input for Verse - out of range (1-{0})"` |
| `SB_ALREADY_FIRST_CHAPTER` | `"Already at first chapter of {0} (1-{1})"` |
| `SB_ALREADY_LAST_CHAPTER` | `"Already at last chapter of {0} (1-{1})"` |
| `SB_ALREADY_FIRST_VERSE` | `"Already at first verse of {0} {1} (1-{2})"` |
| `SB_ALREADY_LAST_VERSE` | `"Already at last verse of {0} {1} (1-{2})"` |

`FormatMsg(template, ParamArray args)` — four-line `Replace`-loop helper; lives in `basUIStrings.bas`.

---

## § 5 — Known Limitations (carry-forward, no fix available)

| Item | Detail |
|------|--------|
| Bug 22 / 23a | First navigation to distant book: ~6–17s layout cost. One-time per session. |
| Bug 27 | Enter in Chapter does not navigate — `onChange` cannot distinguish Enter from keystroke. Superseded by GoButton. |
| Focus mode stale display | ComboBox shows user-typed text (e.g., `"rev"`) after returning from Focus mode. Win32 buffer vs. `getText` callback; no VBA fix available. |
| Status bar flash | Post-`onAction` Word refresh briefly overwrites SBL citation; deferred write shortens but does not eliminate the flash. |
| Status bar ephemeral | Word overwrites status bar on hover, selection, and other events; citation and error messages may disappear. |

---

## § 6 — Module Import Checklist

All of the following `src/` files have been modified since last import into the VBA project.
Must be imported (Remove old → Import new) before testing:

| File | Changes |
|------|---------|
| `src/aeRibbonClass.cls` | GoButton, status bar, state flags, boundary messages, archived methods |
| `src/basBibleRibbonSetup.bas` | `GetGoEnabled`, `OnGoClick`, `GetGoKeytip` wrappers; `KT_SEARCH` reference; reorganised |
| `src/basRibbonDeferred.bas` | `GoToVerseDeferred` stubbed; `UpdateStatusBarDeferred` added; active/archived sections |
| `src/basUIStrings.bas` | Full module: all `SB_*` and `KT_*` constants, `FormatMsg`, `KT_GO`, `KT_SEARCH` |

**Import procedure** (with `Blank Bible Copy.docm` open in Word):

1. Alt+F11 → open VBA editor
2. For each file: right-click module in Project Explorer → Remove (No to export) → File → Import File
3. Ctrl+S to save `.docm`
4. Close and reopen document (or run `RibbonOnLoad` manually) to reinitialise ribbon

---

## § 7 — Session Status Summary (2026-04-20)

| Item | Status |
|------|--------|
| Bug #597 — New Search focus to cmbBook | **OPEN** |
| Bug 16 — Keytip badges end-to-end test | **PENDING — re-test after GetGoKeytip injection** |
| Bug 22 / 23a — First-nav layout delay | **KNOWN LIMITATION** |
| Bug 27 — Enter in Chapter | **SUPERSEDED by GoButton** |
| Step 7 — OLD_CODE cleanup | **PENDING** |
| GoButton — full implementation | **DONE** |
| Status bar feedback — all paths | **DONE** |
| i18n — basUIStrings.bas complete | **DONE** |
| Ribbon Design.md — state machine diagram | **DONE** |
| Ribbon Design.md — doc/UX accuracy | **DONE** |
| Module imports into VBA project | **PENDING** |
| WarmLayoutCache rewrite | **FUTURE** |
| Search tracking reset | **FUTURE** |

---

## § 8 — Proposal: Control ID Constants (UI Contract Completion)

### Suggestion

Add control ID constants to `basUIStrings.bas` as a third surface alongside keytips and
status strings:

```vba
' -- Control IDs (structural — match id= attributes in customUI14.xml) -----------
Public Const CTRL_CMB_BOOK        As String = "cmbBook"
Public Const CTRL_CMB_CHAPTER     As String = "cmbChapter"
Public Const CTRL_CMB_VERSE       As String = "cmbVerse"
Public Const CTRL_BTN_PREV_BOOK   As String = "PrevBookButton"
Public Const CTRL_BTN_NEXT_BOOK   As String = "NextBookButton"
Public Const CTRL_BTN_PREV_CH     As String = "PrevChapterButton"
Public Const CTRL_BTN_NEXT_CH     As String = "NextChapterButton"
Public Const CTRL_BTN_PREV_VERSE  As String = "PrevVerseButton"
Public Const CTRL_BTN_NEXT_VERSE  As String = "NextVerseButton"
Public Const CTRL_BTN_GO          As String = "GoButton"
Public Const CTRL_BTN_NEW_SEARCH  As String = "NewSearchButton"
Public Const CTRL_BTN_ABOUT       As String = "adaeptButton"
```

Replace string literals in `aeRibbonClass.cls` `InvalidateControl` calls (6–8 sites)
with the corresponding constants.

### Pros

| # | Pro |
|---|-----|
| 1 | **Typo prevention** — `CTRL_CMB_BOOK` vs. `"cmbBokk"` — compiler catches the former; runtime silently ignores the latter |
| 2 | **Single rename point** — if a control ID changes in XML, one constant update propagates everywhere |
| 3 | **Completes the UI contract** — keytips, status strings, and control IDs are the three surfaces where VBA and XML share names; all three in one module makes the contract explicit and auditable |
| 4 | **Enables future logging/telemetry** — `InvalidateControl CTRL_CMB_BOOK` is loggable; a string literal requires parsing to interpret |
| 5 | **normalize_vba.py** — `CTRL_*` entries become authoritative and self-documenting |
| 6 | **VSTO port** — control IDs map to `FindControl(Id:=...)` calls; constants make the mapping mechanical |

### Cons

| # | Con |
|---|-----|
| 1 | **Control IDs are very stable** — rename-safety benefit is theoretical for this project |
| 2 | **No compile-time XML validation** — a mismatch between constant value and XML attribute is still a silent runtime failure; constants prevent typos only, not drift |
| 3 | **`basUIStrings` scope creep** — control IDs are structural, not localised; mixing with user-facing text blurs the module purpose |
| 4 | **No current logging or telemetry** — telemetry benefit is speculative; infrastructure for a non-existent consumer is premature |
| 5 | **Small call-site count** — `InvalidateControl` is called in ~6–8 places; literals are short, consistent, and grep-findable |

### Benefits summary

| Area | Impact |
|------|--------|
| Typo safety | Real but low-risk — existing literals are short, consistent, and tested |
| Refactoring | Low value now — control IDs are stable; meaningful only if ribbon is redesigned |
| Completeness | Genuine — the UI contract is currently two-thirds expressed |
| Telemetry / logging | Speculative — no framework exists to consume it |
| VSTO port | Valid long-term — but distant |

### Cost estimate

| Timing | Effort |
|--------|--------|
| **Now** | ~45 min — add 12 constants, replace 6–8 literals in `aeRibbonClass.cls`, add `CTRL_*` entries to `normalize_vba.py` |
| **Later** | Same cost — control IDs are stable; no drift accumulates |

### Decision

**DEFERRED — low urgency, same cost either way.**

The typo-prevention benefit is real but small given the short, stable ID strings already
in use. The stronger argument — completing the UI contract — is valid but cosmetic at
this stage.

**Condition for promoting to "do now":** if Bug #597 (New Search focus) is implemented
in a way that requires referencing control IDs in a new context, or if any
logging/debug infrastructure is added, pull this in at that point. It fits naturally in
any session already touching `basUIStrings.bas`.

**If done:** add a separate `-- Control IDs` section in `basUIStrings.bas` rather than
mixing with keytips — the distinction between user-facing text and structural XML
identifiers is worth preserving in the module layout.

---

## § 9 — Project Goals: Revised Analysis

### Context for revision

The § 8 proposal was evaluated in isolation. The following 10 project-goal points change
the weighting of several conclusions. Each point is analysed; a revised near/far-term
priority table follows.

---

### Point-by-point analysis

**1. No installed user base**

There is currently no migration burden. Architectural changes that would otherwise
require backward-compatible shims can be made cleanly now. This is a narrow window —
once users exist, changes to control IDs, constant names, or module structure carry
a deployment cost. The implication is: establish the right patterns *before* the
first release, not after.

**Impact on § 8:** Elevates urgency slightly. Control ID constants, i18n completeness,
and test coverage are cheaper to add before users exist than after a release freeze.

---

**2. i18n goal — ribbon and code require zero changes per locale**

The current state is partially complete:

| Surface | i18n status |
|---------|------------|
| Keytip strings (`KT_*`) | ✅ All in `basUIStrings.bas`; `getKeytip` callbacks in XML |
| Status bar strings (`SB_*`) | ✅ All in `basUIStrings.bas` |
| **Tab keytip (`Y2`)** | ❌ **Not in XML — auto-assigned by Word at runtime** |
| Ribbon **labels** (`label="Go"`, `label="About"`, tab label, group label) | ❌ **Hardcoded in XML** |
| `sizeString="2 Thessalonians"` (comboBox width) | ❌ Hardcoded — may be wrong for other languages |
| Control IDs | N/A — structural, not localised |

The ribbon labels are the most significant gap. `label="Go"`, `label="About"`,
`label="Radiant Word Bible"` (tab), and `label="Bible Navigation"` (group) are all
static XML attributes. True zero-change i18n requires `getLabel` callbacks for every
visible label, wired through `basUIStrings` constants — the same pattern already
applied to keytips.

`sizeString` sets the comboBox pixel width to match the longest expected entry.
`"2 Thessalonians"` is the longest English book name. A German or Finnish locale may
have a longer canonical name; a `getItemWidth` callback or a `LBL_COMBO_SIZE_STRING`
constant would make this locale-aware.

**Impact on § 8:** Control ID constants are a minor i18n concern (IDs are not
localised). The label gap is the material i18n work. The existing `getKeytip` pattern
is already the template — applying it to `getLabel` is a well-defined extension.

---

**3. Testing automation — verify i18n without manual processes**

The current development workflow is manual: run a test sub via Alt+F8, read
`Debug.Print` output in the Immediate window, compare visually. This is adequate for
unit-level logic tests but cannot verify:

- That every user-facing string in `aeRibbonClass.cls` comes from a `basUIStrings`
  constant (i18n completeness)
- That every `InvalidateControl` call uses a string that matches a real XML control ID
- That `normalize_vba.py` covers every identifier the VBA IDE could corrupt

A Python script run at commit time (as a pre-commit check or CI step) could scan
`src/aeRibbonClass.cls` for inline string literals that look like status bar text or
control IDs, and fail if any are found outside `basUIStrings.bas`. This is a 1–2 hour
tool that provides ongoing automated i18n regression coverage.

**Impact on § 8:** Automated string-completeness checking is the primary mechanism
for verifying the i18n goal. It does not require control ID constants to exist — it
can grep for `InvalidateControl "` (literal with quote) as the signal. Control ID
constants would convert that grep from "find literals" to "verify constants are used."

---

**4. aeAssertClass — use in testing as much as possible**

`aeAssertClass` provides `AssertEqual`, `AssertTrue`, `AssertFalse`, optional logger,
and a pass/fail summary. It is well-suited for:

| Test category | aeAssert suitable? |
|---------------|-------------------|
| Business logic: `ChaptersInBook`, `VersesInChapter`, `ToSBLShortForm` | ✅ Already used |
| `basUIStrings` constant value correctness | ✅ `AssertEqual KT_GO, "G"` etc. |
| `FormatMsg` output correctness | ✅ Direct |
| Ribbon state machine transitions | ❌ Requires UI interaction — not automatable in VBA |
| `InvalidateControl` call correctness | ❌ Requires ribbon instance + Word UI |
| Keytip badge rendering (Bug 16) | ❌ Visual — manual only |
| API connectivity (Point 10) | ❌ Network — manual or separate tooling |

A `basTEST_basUIStrings.bas` module using `aeAssertClass` could verify all constant
values and `FormatMsg` behaviour in a single Alt+F8 run. This is a direct, low-cost
contribution to the i18n verification goal.

---

**5. Testing gaps — areas for improvement**

| Gap | Risk | Suggested approach |
|-----|------|--------------------|
| No test for `basUIStrings` completeness (all strings extracted) | Silent i18n regression | Python pre-commit script scanning for inline literals |
| No test verifying XML control IDs match `InvalidateControl` literals | Silent ribbon breakage on rename | Control ID constants + `basTEST_basUIStrings` |
| No ribbon state machine test | Tab-order and enable-state regressions undetected until manual test | Partially mitigable with unit tests on state-flag logic |
| No regression test for all 66 books navigation | `GoToVerseByScan` correctness only tested ad-hoc | `basTEST_aeRibbonClass` with `aeAssertClass` — requires document open |
| `normalize_vba.py` coverage | Uncovered identifiers silently corrupt on export | Add a test that exports then normalizes and diffs — zero diff = pass |
| No performance baseline | Layout delay regression undetected | Timer-based test in `basTEST_aeRibbonClass` for first-nav cost |
| i18n label gap (see Point 2) | Labels remain in English in all locales | `getLabel` callbacks + `LBL_*` constants in `basUIStrings` |

---

**6. Free version = current working version (reader-focused)**

The free version is the navigation ribbon as it stands. The architectural implication is
that the free version modules should be **frozen at release** — no modifications to
satisfy subscription-only requirements. Subscription features must live in separate
modules and extend the ribbon without touching the free core.

This means:
- `aeRibbonClass.cls`, `basUIStrings.bas`, `basBibleRibbonSetup.bas`, and
  `basRibbonDeferred.bas` are the free-version core — stabilise and lock
- Subscription ribbon controls go in a separate ribbon group or tab, wired to a
  separate class (`aeStudyRibbonClass` or similar)
- Shared data classes (`aeBibleCitationClass`, `aeBibleClass`) are already
  standalone and reusable unchanged

**Impact on § 8:** Control ID constants in `basUIStrings.bas` are part of the free
core. If the subscription version adds new controls, their IDs go in a separate module
(`basStudyUIStrings`?) to keep the free-version contract closed.

---

**7. Subscription version — serious study features**

The subscription version extends, not replaces, the free version. The ribbon
architecture already supports extension via additional groups; the VBA pattern
(singleton class per feature area, shim wrappers in a setup module) is established.

Key architectural requirement: the subscription VBA modules must not create a
compile dependency on the free-version core. If a user has only the free `.docm`,
the free version must run without errors — subscription modules must be absent, not
present-but-disabled.

This argues for **separate `.docm` files** (free and subscription ship different
documents) sharing the same VBA source modules via import, rather than a single
document with feature flags.

---

**8. SBL Citation browsing — click-through on citation strings**

This builds directly on existing infrastructure:
- `aeSBL_Citation_Class` (EBNF parser, closed #521)
- `aeBibleCitationClass.ToSBLShortForm`
- Navigation via `aeRibbonClass.GoToVerse`

The feature: user selects or clicks a citation string (e.g., "Matt 5:14–16 par."); the
add-in parses it, resolves the book/chapter/verse, and navigates. For a list of
citations, a task pane or modeless form would allow click-through browsing.

This is subscription-tier, far term. The infrastructure is ready; the UI layer
(selection detection, task pane) is new work.

---

**9. Verse-of-the-Day — user-defined, About button area**

User-defined VotD implies:
- A settings store (Word document variables via `ActiveDocument.Variables` are the
  natural VBA choice — no external dependency, persists with the `.docm`)
- A UI for selection (modeless form or ribbon flyout)
- A display mechanism (status bar, task pane, or ribbon label)
- Navigation to the VotD passage on demand

This is a self-contained feature with no dependency on subscription infrastructure.
It could ship in the free version as a discovery/engagement feature, or be
subscription-only. Decision has UX implications for the "About" button scope.

---

**10. Bible version comparison — combo box + local + API**

This is the most architecturally complex planned feature:

| Sub-feature | Complexity | Notes |
|-------------|-----------|-------|
| Ribbon combo for installed local versions | Medium | File-system scan for `.docm` / structured text Bibles |
| Display parallel passage in a pane | Medium | Word task pane or second document window |
| API connectivity (Bible Gateway, API.Bible, etc.) | High | HTTP via `WinHTTP` / `XMLHTTP` (late binding — consistent with project pattern) |
| Auth/key management for paid APIs | High | Credential store — `ActiveDocument.Variables` or Windows Credential Manager |
| Public domain API fallback (bible-api.com, wldeh/bible-api) | Low | No auth; JSON response parsing |
| Offline graceful degradation | Medium | API call fails → show only local versions |

All API calls use late binding COM (`CreateObject("WinHTTP.WinHttpRequest.5.1")`) —
consistent with the existing no-added-references policy.

This feature is subscription-tier, far term. The combo box in the ribbon is a new
XML element with its own `getItemCount` / `getItemLabel` / `onChange` callbacks —
the established pattern from `cmbBook` / `cmbChapter` / `cmbVerse` applies directly.

---

### Revised § 8 assessment — Control ID constants

Given the project goals, the original "defer — same cost later" recommendation changes:

| Factor | Original weight | Revised weight |
|--------|----------------|----------------|
| Typo prevention | Low | Low (unchanged) |
| Rename safety | Low | Low (unchanged) |
| i18n contract completeness | Cosmetic | **Moderate** — forms part of the auditable i18n surface |
| Automated testing support | Not applicable | **Moderate** — i18n completeness script can check for literals vs. constants |
| Subscription version extension | Not considered | **Moderate** — establishes the pattern before new controls are added |
| Free-version freeze | Not considered | **Moderate** — locking the free core before release is cleaner with constants in place |

**Revised recommendation: do when setting up test infrastructure — not standalone.**

The constants themselves are a 45-minute task. Their value is fully realised only when
paired with a `basTEST_basUIStrings` module and/or a Python i18n-completeness script.
Do all three together as one "test infrastructure" session rather than adding the
constants in isolation.

---

### Near vs. Far term priority table

#### Near term — free version stabilisation (before first user)

| Priority | Item | Effort | Dependency |
|----------|------|--------|------------|
| 1 | Tab keytip — add `keytip="Y2"` to XML; correct `basUIStrings` comment; re-inject | 30 min | See § 10 |
| 2 | Import all modified modules into VBA project | 30 min | Blocks all testing |
| 3 | Bug #597 — New Search focus to `cmbBook` | 1–2 hr | None |
| 4 | Bug 16 — Keytip badges end-to-end test (incl. `GetGoKeytip`) | 30 min manual | Modules imported |
| 5 | `getLabel` callbacks for all ribbon labels (`LBL_*` constants) | 2–3 hr | Completes i18n goal |
| 6 | `basTEST_basUIStrings` — aeAssert tests for all constants and `FormatMsg` | 1 hr | None |
| 7 | Python i18n-completeness script — scan `src/` for inline literals | 1–2 hr | None |
| 8 | Control ID constants + replace `InvalidateControl` literals | 45 min | Do with item 6/7 |
| 9 | Step 7 — OLD_CODE cleanup | 30 min | None |

#### Far term — subscription version (after free version ships)

| Priority | Item | Notes |
|----------|------|-------|
| 1 | SBL Citation click-through browsing | Infrastructure ready; UI layer is new |
| 2 | Verse-of-the-Day (free or subscription decision needed) | `ActiveDocument.Variables` storage |
| 3 | Bible version comparison — local installed versions | Ribbon combo; file-system scan |
| 4 | Bible version comparison — public domain APIs | `WinHTTP` late binding; JSON parsing |
| 5 | Bible version comparison — paid API auth | Credential management |
| 6 | Subscription `.docm` build — separate from free | Import shared modules; add subscription modules |

---

### Summary

The project goals shift the analysis from "is this architecturally tidy?" to "does this
support a verified, shippable free version and a scalable path to subscription?"

The highest-leverage near-term investment is the combination of:
1. **Tab keytip fix** — lock `Y2` explicitly in the XML (see § 10)
2. `getLabel` callbacks (closes the i18n label gap — the only remaining i18n hole)
3. `basTEST_basUIStrings` + Python completeness script (provides automated i18n
   regression coverage — replaces the manual "check all strings" process)
4. Control ID constants (completes the UI contract before the free version is frozen)

These four items together are a half-day of work. They establish the automated
verification path that makes Point 2 (zero i18n changes needed) a tested guarantee
rather than a design intention.

---

## § 10 — Tab Keytip `Y2`: i18n Risk and Mitigation

### Current state

The `<tab>` element in `customUI14backupRWB.xml` has **no `keytip=` attribute**:

```xml
<tab id="RWB" label="Radiant Word Bible">
```

Word auto-assigns the keytip at runtime. In the English development environment it
assigned `Y2`. The `basUIStrings.bas` header comment states:

> *"Ribbon tab keytip (Alt, Y2) is defined in customUI XML, not in this module."*

This comment is **incorrect**. `Y2` is not in the XML — it is the value Word
auto-assigned during testing. The comment implies it is a deliberate choice written
into the XML; it should say "auto-assigned by Word at runtime — see § 10."

### The auto-assignment algorithm

When Word renders the ribbon it assigns keytip badges to tabs:

1. Built-in tabs receive locale-specific fixed keytips (`H`=Home in English,
   `S`=Start in German, `D`=Démarrer in French, etc.).
2. For custom tabs Word tries the first letter of the `label=` attribute. If taken
   it appends a digit (`R2`, `R3`...). If all combinations of that letter are taken
   it moves to the next available letter.
3. The assignment depends on the **installed Word locale** and on **which other
   add-ins are loaded**.

### Prior attempt: `keytip="RW"`

`RW` was tested as a fixed two-character keytip. It **conflicted with the Review
tab** in English Word and triggered the Word Count operation, which takes
approximately 20 seconds on the Bible document. `RW` is not a safe choice.

### The API constraint

Unlike `<button>` and `<comboBox>` which support `getKeytip` callbacks, the
`<tab>` element in the Office Fluent ribbon schema (2009 customUI namespace) supports
**only the static `keytip=` attribute**. There is no `getKeytip` for tabs.
Therefore:

- The tab keytip **cannot be made dynamic** via a VBA callback
- It cannot be extracted to `basUIStrings.bas` as a runtime constant
- It must be a hard-coded static value in the XML, or left unset (auto-assigned)

### Mitigation: explicitly lock `Y2` in the XML

`Y2` is already a two-character combination — the same property that made `RW`
appealing — but Word itself selected it as available and non-conflicting in the
English environment. The correct fix is to **declare it explicitly** rather than
rely on the auto-assignment producing the same result on every machine:

```xml
<tab id="RWB" label="Radiant Word Bible" keytip="Y2">
```

**Why this is the right choice:**

| Property | Assessment |
|----------|-----------|
| Non-conflicting in English Word | ✅ Confirmed — Word chose it |
| Two-character — cannot collide with single-letter built-in keytips | ✅ |
| Mnemonic value | ❌ None — known limitation; accepted |
| Deterministic across reinstalls and add-in changes | ✅ Once explicit in XML |
| Safe in non-English locales | ⚠️ Unverified — `Y` is not a common built-in initial in most European locales, making `Y2` a low-risk choice, but not guaranteed |

### Residual risk

If a non-English Word installation has assigned `Y` or `Y2` to a built-in tab or
another add-in, Word's conflict resolution is undefined. The user may see a
disambiguating prompt or the keytip may not function. This is an accepted limitation
of the Office Fluent ribbon API — no workaround exists at the XML or VBA level.

The risk is judged low: `Y` is not an initial letter in the standard tab names of
major European locales (German: H/S/E/S/S/Ü/A/N; French: A/I/M/R/P/A/D/A).

### Required changes

| File | Change |
|------|--------|
| `customUI14backupRWB.xml` | Add `keytip="Y2"` to `<tab id="RWB">` |
| `py/inject_ribbon.py` | Run after XML edit to update `Blank Bible Copy.docm` |
| `src/basUIStrings.bas` | Correct the header comment (remove misleading "defined in customUI XML" claim) |

Documentation (`Ribbon Design.md`, review files) already uses `Y2` throughout —
no documentation changes needed.

**Status: PROPOSED — pending approval to add `keytip="Y2"` to the XML.**
