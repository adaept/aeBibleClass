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
| Bug #597 — New Search focus to cmbBook | **CLOSED AS KNOWN LIMITATION — Option A attempted and reverted; see § 11 note** |
| Bug 16 — Keytip badges end-to-end test | **PENDING — re-test after GetGoKeytip injection** |
| Bug 22 / 23a — First-nav layout delay | **KNOWN LIMITATION** |
| Bug 27 — Enter in Chapter | **SUPERSEDED by GoButton** |
| Step 7 — OLD_CODE cleanup | **PENDING** |
| GoButton — full implementation | **DONE** |
| Status bar feedback — all paths | **DONE** |
| i18n — basUIStrings.bas complete | **DONE** |
| Ribbon Design.md — state machine diagram | **DONE** |
| Ribbon Design.md — doc/UX accuracy | **DONE** |
| Module imports into VBA project | **DONE — 2026-04-20 — ImportAllVBAFiles run manually** |
| Tab keytip `Y2` locked in XML; `basUIStrings` comment corrected; re-injected | **DONE — 2026-04-20** |
| Session manifest (`sync/session_manifest.txt`) — developer sync process | **DONE — 2026-04-20** |
| WarmLayoutCache rewrite | **FUTURE** |
| Search tracking reset | **FUTURE** |
| VSTO / VB.NET migration analysis | **CARRY FORWARD — see § 12** |
| `.dotm` template architecture — Bible content as `.docx` | **CARRY FORWARD — see § 13** |
| Split document architecture — page numbers, footnotes, cross-file navigation | **CARRY FORWARD — see § 14** |

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
| 7 | `.dotm` template architecture — Bible content as `.docx` | `ThisDocument.VBProject.Name`; `IsBibleDocument()` guard; see § 13 |
| 8 | Split document architecture — per-book or OT/NT | Page numbers, footnotes, cross-file navigation, merge script; see § 14 |
| 9 | VSTO / VB.NET migration — subscription version | Start with VSTO setup; `basUIStrings` → `.resx`; UI Automation for Bug #597; see § 12 |

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

**Status: DONE — 2026-04-20. `keytip="Y2"` added to XML; `basUIStrings.bas` comment corrected; re-injected.**

---

## § 11 — Developer Sync: File State Verification Between Claude and VBA Editor

### Problem

Two actors modify the same `src/` files:

1. **Claude** — edits files directly on disk (`src/*.bas`, `src/*.cls`)
2. **Developer** — imports modified files into the Word VBA editor; may make further
   changes there and export back to `src/` via `ImportAllVBAFiles` / export routines

The risk: Claude holds a stale mental model of a file's content. If the developer
exports a VBA-edited version back to `src/` and Claude then edits the same file
without reading it fresh, Claude's edit is based on the version it last wrote, not
the current disk state. The `Edit` tool requires a prior `Read`, which mitigates
this within a single session — but across sessions the risk is real.

The developer notes: *"This is most likely a benign event but serves as a reminder."*
A lightweight warning mechanism is sufficient; correctness guarantees are not required.

### Why simple MD5 is imperfect

A VBA editor round-trip (import → edit → export) reformats code:

- Line endings may normalise
- Whitespace around operators may change
- `normalize_vba.py` then runs and changes casing

This means the MD5 of a file Claude wrote will almost always differ from the MD5
after a VBA export of equivalent content. A naive MD5 comparison would produce
**false positives on every round-trip** — warning on every import/export cycle
even when no intentional change was made.

False positives are acceptable per the developer's framing ("most likely benign"),
but a high false-positive rate degrades the signal into noise.

### Proposed approach: git-object hash on committed state

Rather than comparing disk MD5s, compare against the **last committed git object
hash** for each file. The workflow becomes:

1. Claude edits `src/X.bas` and the session ends.
2. Developer runs `normalize_vba.py`, imports the file, and **commits** — this is
   the sync point. The committed hash is now the authoritative "known good" state.
3. Developer makes VBA changes, exports, normalizes, and commits again.
4. Claude starts the next session: `git status` immediately shows any files that
   have been modified since the last commit. If `src/X.bas` is listed as modified
   (`M`), it has changed since the last known commit — prompt before editing.

**`git status` is already the sync signal.** The committed hash is the shared
reference point both actors implicitly agree on.

### Complementary: session manifest file

After any session where Claude edits `src/` files, write a manifest listing the
files changed:

```
sync/session_manifest.txt   (committed alongside the src/ edits)
```

Format:
```
# Claude session 2026-04-20
# Files modified — import these into VBA before testing
src/customUI14backupRWB.xml  (XML only — no VBA import needed)
src/basUIStrings.bas
src/basBibleRibbonSetup.bas
```

This gives the developer a checklist of what to import without requiring them to
remember which files changed across a long session. It supplements `git diff --stat`
with intent (why each file changed is in the commit message; the manifest is the
import checklist).

### Lightweight checksum script (optional — if git is insufficient)

If a between-commit warning is needed (Claude edits a file, developer has not yet
committed, and the developer subsequently exports from VBA before committing):

```python
# py/sync_check.py
# Usage:
#   python3 py/sync_check.py record src/basUIStrings.bas src/aeRibbonClass.cls
#   python3 py/sync_check.py verify src/basUIStrings.bas src/aeRibbonClass.cls

import hashlib, json, sys
from pathlib import Path

MANIFEST = Path('sync/checksums.json')

def md5(path):
    return hashlib.md5(Path(path).read_bytes()).hexdigest()

if sys.argv[1] == 'record':
    data = json.loads(MANIFEST.read_text()) if MANIFEST.exists() else {}
    for f in sys.argv[2:]:
        data[f] = md5(f)
    MANIFEST.parent.mkdir(exist_ok=True)
    MANIFEST.write_text(json.dumps(data, indent=2))
    print(f'Recorded {len(sys.argv)-2} file(s).')

elif sys.argv[1] == 'verify':
    data = json.loads(MANIFEST.read_text()) if MANIFEST.exists() else {}
    for f in sys.argv[2:]:
        stored = data.get(f)
        current = md5(f)
        if stored is None:
            print(f'NEW (no prior record): {f}')
        elif stored != current:
            print(f'CHANGED since last record: {f}')
        else:
            print(f'OK: {f}')
```

`sync/checksums.json` is committed so Claude can read it at the start of a session
and warn if any file it previously modified has since changed on disk.

**Acknowledged limitation:** a VBA export of unchanged content will report
`CHANGED` due to reformatting. The developer treats this as expected and
non-blocking — it is a prompt to review, not an error.

### Recommendation

| Mechanism | Cost | Signal quality | Verdict |
|-----------|------|----------------|---------|
| `git status` before each edit session | Zero — already available | High for committed changes | **Use always** |
| Session manifest (`sync/session_manifest.txt`) | 5 min per session | High — explicit import checklist | **Implement now** |
| `py/sync_check.py` checksum script | 1 hr one-time | Moderate — false positives on VBA export | **Implement if manifest proves insufficient** |

**Immediate action:** Claude writes `sync/session_manifest.txt` at the end of each
session listing files it modified. Developer uses it as the import checklist.
The checksum script is held in reserve.

**Status: APPROVED — 2026-04-20.**

### Bug #597 — Option A attempted and reverted

**Compile error (first):** `Application.SendKeys` does not exist on Word's
`Application` object — it is an Excel-only method. Fixed to standalone `SendKeys`.

**Runtime bug (second — fatal):** After clicking New Search with the cursor in the
document (e.g., at a chapter marker after navigating to Revelation), `SendKeys "%Y2B"`
inserted `2B` into the document at the cursor position.

**Root cause:** VBA `SendKeys` `%` prefix sends the following character as an
**Alt+X chord** (Alt held while pressing X). Ribbon keytip navigation requires
a different sequence:

| Actual UI sequence | What `SendKeys "%Y2B"` sends |
|-------------------|------------------------------|
| Tap Alt alone → keytip mode activates | Alt+Y (chord) — not the same |
| Press Y (plain) | — |
| Press 2 (plain) → tab selected | 2 (plain, but focus is now ambiguous) |
| Press B (plain) → cmbBook focused | B (plain, inserted into document) |

The `%X` syntax in `SendKeys` cannot express "tap Alt alone" — it always produces
a chord. There is no VBA `SendKeys` syntax that replicates the keytip tap sequence.
Any characters after a failed Alt+X chord that does not activate the ribbon are
typed into whichever document area has focus.

**Decision: Option C — accepted as known limitation.**

After clicking New Search the user re-focuses `cmbBook` manually:
- Click the Book comboBox, or
- Press `Alt, Y2, B`

`FocusBookDeferred` is retained in `basRibbonDeferred.bas` as a commented
dead-end with the failure documented inline. `OnNewSearchClick` in
`aeRibbonClass.cls` retains a comment explaining the decision.

---

### Implementation — first manifest written

`sync/session_manifest.txt` created for the 2026-04-20 session. It records:

- All `src/` files modified by Claude, with import status (`[IMPORT]`, `[XML]`,
  `[SCRIPT]`, `[DOC]`, `[DONE]`)
- Which changes were made before vs. after the `ImportAllVBAFiles` run
- A final import checklist for the developer

**Ongoing convention:**

- Claude writes or updates `sync/session_manifest.txt` at the end of every session
  that modifies `src/` files
- The developer uses it as the import checklist after each session
- At the start of the next session Claude runs `git status` and notes any files
  that have changed since the last commit before editing them
- The checksum script (`py/sync_check.py`) remains deferred — implement if the
  manifest alone proves insufficient to catch between-commit drift

### Session manifest: expected process (reference)

**Developer workflow per session:**

```
1. Claude edits src/ files during the session
2. Session ends → Claude writes/updates sync/session_manifest.txt
3. Developer opens sync/session_manifest.txt
4. For each [IMPORT] line: Remove old module in VBA editor → Import updated file
5. For each [XML] line: run py/inject_ribbon.py (already done by Claude if possible)
6. For each [SCRIPT] / [DOC] line: no action needed
7. Run normalize_vba.py on any exported files
8. git add + git commit (this is the sync point both actors share)
```

**Claude workflow at session start:**

```
1. Run git status — review any modified files
2. If a file Claude last modified shows as changed: read it fresh before editing;
   note the change in the session log
3. Check sync/session_manifest.txt for any outstanding [IMPORT] items from prior
   sessions that may not have been completed
```

**Import status key used in manifest:**

| Tag | Meaning |
|-----|---------|
| `[IMPORT]` | Must be imported into VBA editor (Remove old → Import new) |
| `[XML]` | Ribbon XML — run `py/inject_ribbon.py`; no VBA import |
| `[SCRIPT]` | Python/tooling file; no VBA import needed |
| `[DOC]` | Documentation only; no import needed |
| `[DONE]` | Already imported this session; noted for record-keeping |

**When the manifest is NOT needed:**
If Claude only edits documentation (`md/`, `rvw/`), Python scripts (`py/`), or the
ribbon XML (and `inject_ribbon.py` is run immediately), no VBA import is required
and the manifest can note `[DOC]`/`[XML]` entries only.

---

## § 12 — VSTO / VB.NET Migration: Solutions to Known VBA Limitations

**Carry-forward item.** Arose from Bug #597 (SendKeys failure). The analysis covers
which current VBA limitations are solved by a VSTO port, which are not, and the
architectural implications for the free vs. subscription version split.

### Bug #597 — SendKeys / ribbon focus

The VBA `SendKeys` `%X` syntax always produces an Alt+X chord. Ribbon keytip mode
requires Alt tapped alone then plain keypresses — not expressible in VBA. VSTO
provides two clean solutions:

**Option 1 — P/Invoke `SendInput`**

`SendInput` (Win32 API) sends individual key-down and key-up events separately,
exactly replicating what a user does at the keyboard:

```vbnet
' 1. VK_MENU keydown  → enters keytip mode
' 2. VK_MENU keyup
' 3. 'Y' keydown/up   → narrows to Y prefix
' 4. '2' keydown/up   → selects RWB tab
' 5. 'B' keydown/up   → focuses cmbBook
```

Characters are sent after Alt is fully released — correct sequence, not a chord.
No leakage into the document.

**Option 2 — UI Automation**

Locates and focuses a ribbon control by automation ID, bypassing keytip sequences
entirely:

```vbnet
Imports System.Windows.Automation

Dim wordElement = AutomationElement.FromHandle(New IntPtr(Globals.ThisAddIn.Application.Hwnd))
Dim bookCombo = wordElement.FindFirst(
    TreeScope.Descendants,
    New PropertyCondition(AutomationElement.AutomationIdProperty, "cmbBook"))
bookCombo?.SetFocus()
```

No dependence on keytip sequences, locale, or keyboard state.

---

### Full limitation comparison

| Limitation | VBA status | VSTO / VB.NET |
|---|---|---|
| **Bug #597** — focus ribbon control | No solution | ✅ P/Invoke `SendInput` or UI Automation |
| **`Application.OnTime`** pattern | Fragile; fires after event cycle | ✅ `Async/Await` + `Task.Delay` — explicit, testable, cancellable |
| **Late binding** (`CreateObject`) | No IntelliSense; runtime errors only | ✅ Early binding via Office interop assemblies; compile-time checking |
| **API integration** (Bible Gateway etc.) | `WinHTTP` COM; no async; manual JSON | ✅ `HttpClient` + `Async/Await`; `System.Text.Json` or Newtonsoft |
| **Testing** | `aeAssertClass` + `Debug.Print`; no mocking | ✅ NUnit / MSTest; mocking with Moq; CI integration |
| **i18n strings** | `basUIStrings.bas` constants | ✅ `.resx` resource files; `ResourceManager`; design-time locale switching |
| **Casing normalization** | `normalize_vba.py` required after every export | ✅ Not needed — VB.NET compiler enforces casing; no export/import cycle |
| **`Option Private Module` / `OnTime` name resolution** | Requires `basRibbonDeferred` to omit `Option Private Module` | ✅ Not applicable — .NET namespaces and delegates |
| **Status bar flash** after `onAction` | Deferred write via `OnTime`; flash unavoidable | ⚠️ Async post-action write still possible; flash is Office behaviour, not VBA |
| **Focus mode stale display** | No fix — Win32 buffer vs. `getText` callback | ❌ Same Win32/COM ribbon architecture underneath |
| **`<tab>` keytip — no `getKeytip`** | Static `keytip="Y2"` in XML; no callback | ❌ Same Office XML schema constraint |
| **Layout delay** (Bug 22/23a) | Word's layout engine; no VBA event | ❌ Word internal — unchanged |

### What VSTO does NOT solve

Three limitations are Office-architectural, not VBA-architectural:

1. **`<tab>` has no `getKeytip` callback** — customUI XML schema; same in VSTO
2. **Focus mode stale display** — Win32 comboBox buffer vs. `getText`; same in VSTO
3. **Layout delay** — Word's internal pagination engine; independent of add-in layer

### Migration cost by module

| Current VBA | VSTO equivalent | Migration cost |
|---|---|---|
| `basUIStrings.bas` `Public Const` | `.resx` file; `My.Resources.xxx` | Low — direct name mapping |
| `FormatMsg(template, args)` | `String.Format` built-in | Trivial — delete `FormatMsg` |
| `aeRibbonClass.cls` singleton | `ThisAddIn`-scoped ribbon class | Low |
| `basBibleRibbonSetup.bas` shim wrappers | Eliminated — ribbon XML callbacks wire directly to class methods | Removed entirely |
| `basRibbonDeferred.bas` `OnTime` pattern | `Async/Await` in ribbon event handlers | Medium — rewire all deferred subs |
| `aeAssertClass` tests | NUnit test project | Medium — port test logic; gain CI |
| Late-binding COM (`CreateObject`) | Project references + `Imports` | Low — mechanical |
| `normalize_vba.py` | Not needed | Deleted |
| `py/inject_ribbon.py` | Still needed — `.docm` XML injection unchanged | Retained |

### Implications for free vs. subscription split

The free version (current VBA ribbon) can remain VBA if the known limitations are
accepted. The subscription version has strong reasons to be VSTO from the start:

- Bible API integration requires `HttpClient` + async — VBA `WinHTTP` is viable
  but significantly more complex
- UI Automation (Bug #597 fix) is needed for the New Search focus UX
- NUnit test infrastructure is needed to verify i18n and regression coverage at scale
- The subscription version ribbon (additional controls, Bible version combo) is
  easier to manage with VSTO Ribbon Designer and early binding

**Recommended path:** free version ships as VBA (current architecture); subscription
version is built as a VSTO add-in that ships alongside or replaces the `.docm`.
The `basUIStrings` → `.resx` migration is the cleanest handoff point — it was
designed for this from the start.

### Status

**CARRY FORWARD** — no implementation decision required now. Relevant when
subscription version planning begins. The Bug #597 analysis is the concrete trigger:
the first subscription-version development session should open with VSTO setup rather
than continuing to work around VBA limitations.

---

## § 13 — `.dotm` Template Architecture: Bible Content as `.docx`

**Carry-forward proposal.** Arose from the Bug #597 / `SendKeys` limitation and the
goal of supporting Bible document editors who should not need VBA knowledge.

### Overview

Word loads ribbon customizations and VBA macros from any `.dotm` file placed in the
Word STARTUP folder (`%AppData%\Microsoft\Word\STARTUP\`). That template is loaded
globally for every Word session — its ribbon and VBA are available to all open
documents, including plain `.docx` files.

This separates the add-in from the content entirely:

```
BibleAddIn.dotm   ← ribbon XML + all VBA  (developer-maintained)
Genesis.docx      ← Bible content only, no VBA  (editor-maintained)
Exodus.docx       ← same
...
```

### How callback resolution works

When Word sees `onAction="OnGoClick"` in ribbon XML it searches:
1. The active document's VBA project
2. All loaded template VBA projects (STARTUP folder)

`OnGoClick` lives in `BibleAddIn.dotm` — Word finds it there. The active `.docx`
needs no VBA at all. `Application.ActiveDocument` in the template code operates on
whichever document the user is working in — already how the current code is written.

### Required code change — `ThisDocument.VBProject.Name`

Every `Application.OnTime` call currently builds the project name from the active
document:

```vba
' Current — resolves to the content document's project (wrong in template context):
projName = Application.ActiveDocument.VBProject.Name

' Required — resolves to the template's own project:
projName = ThisDocument.VBProject.Name
```

`ThisDocument` in a `.dotm` module refers to the template itself. `Application.OnTime`
uses that name to find `basRibbonDeferred` subs — it must match the template project
name, not the content document's project name.

This is the **only structural VBA change** required. `Instance()`, `m_ribbon.Invalidate`,
`aeRibbonClass`, and all other code work unchanged.

### Context guard — ribbon inert on non-Bible documents

The ribbon appears for all open documents. A document variable identifies Bible
documents; every `GetEnabled` callback checks it:

```vba
Private Function IsBibleDocument() As Boolean
    On Error Resume Next
    IsBibleDocument = (Application.ActiveDocument.Variables("RWB_Document").Value = "1")
    On Error GoTo 0
End Function

Public Function GetGoEnabled(control As IRibbonControl) As Boolean
    GetGoEnabled = IsBibleDocument() And (m_currentBookIndex <> 0)
End Function
```

Each Bible `.docx` carries `RWB_Document = "1"` as its identity marker. All ribbon
controls are disabled (invisible in practice) when any non-Bible document is active.

### Architecture comparison

| Factor | Current `.docm` | Template `.dotm` + `.docx` |
|--------|----------------|---------------------------|
| Bible content format | `.docm` — macro-enabled required | `.docx` — plain Word document |
| Editor VBA knowledge | Must not break macros | None — content only |
| Merge / diff of content | Binary `.docm`; fragile | `.docx`; diffable via pandoc / docx2txt |
| Ribbon scope | Only when `.docm` is open | All Word sessions; context-guarded |
| Context guard needed | No | Yes — `IsBibleDocument()` |
| Code change required | None | `ActiveDocument.VBProject.Name` → `ThisDocument.VBProject.Name` in all `OnTime` calls |
| Distribution | Single `.docm` | Template installed once; `.docx` files freely shared |
| `basRibbonDeferred` `Option Private Module` omission | Still required | Same — unchanged |
| VSTO migration path | `.docm` → VSTO add-in | `.dotm` → VSTO add-in (same step, cleaner handoff) |

### Known limitations — unchanged from current

| Limitation | Status |
|---|---|
| `<tab>` keytip `Y2` — no `getKeytip` callback | Same constraint |
| Focus mode stale display | Same Win32 issue |
| Bug #597 — `SendKeys` / ribbon focus | Same VBA limitation; resolved in VSTO (§ 12) |
| Layout delay (Bug 22/23a) | Word-internal; unchanged |

### Implementation cost

| Task | Effort |
|------|--------|
| Create `BibleAddIn.dotm` from current `.docm` (copy VBA project, ribbon XML) | 30 min |
| Replace `ActiveDocument.VBProject.Name` with `ThisDocument.VBProject.Name` | 30 min — grep all `OnTime` calls in `aeRibbonClass.cls` and `basRibbonDeferred.bas` |
| Add `IsBibleDocument()` guard to all `GetEnabled` callbacks | 1 hr |
| Add `RWB_Document = "1"` document variable to each Bible `.docx` | Per-document; one-time setup sub |
| Test: open a non-Bible `.docx` — ribbon should be fully disabled | 15 min manual |
| Test: open a Bible `.docx` — full ribbon function | Existing manual test suite |

Total estimated effort: **~2–3 hours.**

### Relationship to VSTO migration (§ 12)

The `.dotm` approach is the cleanest stepping stone to VSTO:

```
Phase 1 (now):      .docm  →  BibleAddIn.dotm + Bible.docx
Phase 2 (subscription):  BibleAddIn.dotm  →  BibleAddIn VSTO add-in
                          Bible.docx unchanged
```

Phase 2 requires no change to Bible content files — editors are completely
insulated from the infrastructure migration.

### Status

**CARRY FORWARD** — recommended as the next architectural step before the subscription
version begins. Implement when the free version is stable and the first Bible
document editors are onboarded. The `ThisDocument.VBProject.Name` change is the
gate item; everything else follows from it.

---

## § 14 — Split Document Architecture: Code, Workflow, Page Numbers, Footnotes

**Carry-forward item.** Continues from § 13 (`.dotm` architecture). Covers three
developer questions: what goes in the `.dotm`, how the development workflow changes,
and the implications of splitting the Bible content into multiple `.docx` files.

---

### 1. Does all code go in the `.dotm`?

**Yes — the `.dotm` is the single development environment.** Everything currently in
the `.docm` VBA project moves to `BibleAddIn.dotm`:

- Ribbon classes: `aeRibbonClass`, `basBibleRibbonSetup`, `basRibbonDeferred`
- Data classes: `aeBibleCitationClass`, `aeBibleClass`, `aeSBL_Citation_Class`
- Utility modules: `basUSFM_Export`, `basAuditDocument`, `basAddHeaderFooter`
- Test modules: `basTEST_*`, `aeAssertClass`, `aeLoggerClass`
- All helper modules: `basUIStrings`, `normalize_vba.py` targets

Content `.docx` files contain **zero VBA**. The `.dotm` is where you code from,
always. Utility subs that operate on document structure already use
`Application.ActiveDocument` — they work from the template unchanged. The
`IsBibleDocument()` guard applies only to ribbon `GetEnabled` callbacks; Alt+F8
utility subs are trusted by the developer to run on the right document.

---

### 2. Development workflow change

**Current:**
```
Open BibleClass.docm
  → Bible text is there
  → Alt+F11 → VBA editor → all code is in one project
```

**With `.dotm` + `.docx`:**
```
Word starts → BibleAddIn.dotm loads automatically from STARTUP folder
  → VBA editor always shows BibleAddIn project in Project Explorer
  → To code: Alt+F11 → select BibleAddIn in Project Explorer → edit
  → To edit Bible text: File → Open → Genesis.docx (or whichever book)
  → Both visible simultaneously; switch between them normally
```

The `.dotm` is always present in the VBA editor's Project Explorer —
you never open it specifically to reach the code. To edit the template itself:

```
File → Open → BibleAddIn.dotm
  → Opens as a document window (content is empty — template only)
  → Alt+F11 → its project is active in the editor
```

The git workflow (`normalize_vba.py`, export/import, session manifest) targets
the `.dotm`'s modules — otherwise identical to the current process.

**Net change for the developer:** minor. Code is always accessible via Alt+F11
regardless of which document is open. The `.docm` single-file convenience is
replaced by the `.dotm` always-present convenience.

---

### 3. Split document: page numbers, footnotes, layout delay

#### Layout delay

Splitting by book or section largely eliminates Bug 22/23a. Each per-book `.docx`
is small enough that Word paginates it immediately — the 6–17 second first-load
cost disappears within a file. Cross-book navigation incurs a file-open delay
instead, which is brief and predictable.

#### Page numbers

Each `.docx` starts page numbering at 1 by default.

| Approach | Mechanism | Trade-off |
|----------|-----------|-----------|
| Restart per file | Default Word behaviour | Fine for screen; page refs don't match print |
| Manual start-at per file | Insert → Page Number → Format → Start at N | Fragile — recalculate all if any earlier file changes length |
| Remove during editing | Strip page number fields while editing | Cleanest for editing; regenerate at print/PDF stage |
| Master document | Word built-in master/sub-document | Notorious for corruption — not recommended |

**Recommendation:** Remove page number fields from editing copies. Page numbers are
only meaningful in final printed/PDF form. Generate the final document by merging
files at print time (Python `python-docx` or a Word macro), at which point page
numbering is set once and correctly.

#### Footnote numbers

| Convention | Mechanism | Notes |
|------------|-----------|-------|
| Restart per book | Each `.docx` footnotes start at 1 | Natural for Bible — independent per book |
| Restart per chapter | Word section break → footnote numbering option | Finer-grained; requires section breaks at chapter boundaries |
| Continuous across books | Manual start-at per file | Fragile; not appropriate for Bible use |

**Recommendation:** Restart per book (default when split by book file). Matches
how Study Bibles present footnotes. No coordination required between files.

#### Cross-file navigation — ribbon impact

`GoToVerseByScan` currently scans the entire open document. With split files:

- **Within-book navigation** (Prev/Next Chapter, Prev/Next Verse, Go to verse):
  unchanged — scans the active file
- **Cross-book navigation** (Prev/Next Book, or Go to a different book):
  requires opening a different file

The ribbon needs a book-to-file mapping in the `.dotm`:

```vba
Private Function BookFilePath(bookIndex As Long) As String
    ' Maps book index (1-66) to file path
    ' e.g., index 1 → "C:\Bible\01_Genesis.docx"
    ' Stored as a Const array, config file, or .dotm document variable
End Function

' In PrevButton / NextButton:
Dim targetPath As String
targetPath = BookFilePath(m_currentBookIndex - 1)
Application.Documents.Open targetPath
' then navigate to last chapter/verse of that book
```

Opening a new document replaces the layout-delay wait with a file-open pause.
For most navigation (within a book) the experience is faster. Cross-book jumps
are infrequent enough that the file-open pause is acceptable.

#### `headingData` scan

Currently built by scanning all H1 headings in the open document. With split
files each document contains only one book's H1 headings — the scan becomes
faster and `headingData` is rebuilt per file. The `DocumentOpen` event in the
`.dotm` triggers a rescan and ribbon invalidate when a new Bible `.docx` is opened.

---

### Upfront work required before splitting

| Task | Notes |
|------|-------|
| Define split boundaries | Per book (66 files) or per section (OT/NT = 2 files) |
| Decide footnote numbering convention | Restart per book recommended |
| Decide page number strategy | Remove during editing; regenerate at print time |
| Create file-naming convention | `01_Genesis.docx` — sortable, predictable |
| Build book-to-file path mapping | Array or config in `.dotm` |
| Add `RWB_Document` variable to each file | One-time setup macro |
| Handle `DocumentOpen` event in `.dotm` | Triggers `headingData` rescan + ribbon invalidate |
| Implement cross-book navigation | `PrevButton`/`NextButton` open adjacent file |
| Write merge script for final print/PDF | Python `python-docx` or Word macro |

---

### Recommended phasing

A pragmatic intermediate step before committing to 66 files:

**Phase A — Split into two files (OT + NT):**
- Cuts layout delay roughly in half immediately
- Cross-file navigation is one boundary only (Malachi → Matthew)
- Book-to-file mapping has two entries
- Merge script is trivial
- Minimal architectural change to the ribbon

**Phase B — Split per book (66 files):**
- Eliminates layout delay completely
- Full book-to-file mapping (66 entries)
- Full cross-file navigation architecture
- Enables parallel editing by multiple contributors per book

**Phase C — Merge script for print/PDF:**
- Concatenates all files with correct page and footnote numbering
- Runs independently of the editing workflow

---

### Three-phase migration picture

```
Current:   BibleClass.docm (all-in-one)
Phase 1:   BibleAddIn.dotm + Bible.docx (one content file)        ← § 13
Phase A:   BibleAddIn.dotm + OT.docx + NT.docx                    ← § 14, Phase A
Phase B:   BibleAddIn.dotm + 01_Genesis.docx ... 66_Revelation.docx ← § 14, Phase B
Phase 2:   BibleAddIn VSTO + 01_Genesis.docx ... (unchanged)      ← § 12
```

Each phase is independently deployable. Editors are insulated from every
infrastructure transition after Phase 1.

---

### Status

**CARRY FORWARD — for further review.** Decision points before implementation:

1. Split boundary: OT/NT first, or go directly to per-book?
2. Footnote convention confirmed: restart per book?
3. Page number strategy confirmed: strip during editing, regenerate at print?
4. File-naming convention agreed
5. `DocumentOpen` event handler design reviewed against current `OnRibbonLoad` flow

---

## § 15 — Split document architecture abandoned

**Decision (2026-04-20):** The split document approach (§ 14) is abandoned.

**Rationale:** The Study Bible is a completed work — all 66 books, footnotes numbered
1–1000, 146 sections, ~10 picture maps. The upfront cost to split the document plus the
back-end merge cost (merge script, page number regeneration, footnote renumbering) plus
the ongoing testing burden does not balance against a one-time user instruction:

> "GoTo Revelation takes ~20 seconds to load this 900+ page document for editing.
>  Your patience is appreciated."

The layout delay is a known, disclosed limitation of the single-document approach.
It is not a defect — it is a predictable consequence of document size. Splitting the
document trades that constraint for a different set of constraints (merge tooling,
cross-file navigation, multi-file editing workflow) that are harder to explain and harder
to maintain.

**Architecture remains:** single `.docm` (current) → `BibleAddIn.dotm + Bible.docx` (§ 13 — Phase 1, when ready).

---

## § 16 — i18n label gap: batch plan

### Scope

Three tasks confirmed for a single implementation session:

#### Task A — LBL_* constants + `getLabel` callbacks

Four visible ribbon labels are currently static strings in XML with no i18n path:

| Control | XML attribute | Proposed constant |
|---------|--------------|-------------------|
| `<tab id="RWB">` | `label="Radiant Word Bible"` | `LBL_TAB` |
| `<group id="NavGroup">` | `label="Bible Navigation"` | `LBL_GROUP` |
| `<button id="GoButton">` | `label="Go"` | `LBL_GO` |
| `<button id="adaeptButton">` | `label="About"` | `LBL_ABOUT` |

Work:
- Add `LBL_*` constants to `basUIStrings.bas`
- Replace static `label=` with `getLabel=` in XML for all four elements
- Add four `GetXxxLabel` callbacks to `basBibleRibbonSetup.bas`

`showLabel="false"` controls (Prev/Next buttons, comboBoxes, NewSearchButton) have no
visible label — no `getLabel` needed for i18n.

#### Task B — Control ID constants + fix `InvalidateControl` literals

All control ID strings in `InvalidateControl` calls are currently inline literals.
Define `CTRL_*` constants in `basUIStrings.bas` and replace every literal.

**Preexisting bug discovered:** Three `InvalidateControl` calls in `OnGoClick`
(lines 1004, 1010, 1015 of `aeRibbonClass.cls`) use wrong IDs:

| Wrong ID (current) | Correct XML id |
|--------------------|----------------|
| `"BookComboBox"` | `"cmbBook"` |
| `"ChapterComboBox"` | `"cmbChapter"` |
| `"VerseComboBox"` | `"cmbVerse"` |

`InvalidateControl` silently no-ops on an unknown ID. Effect: after an invalid Go
attempt the comboBox display does not refresh. Fix is natural when constants are
introduced — the constant carries the correct ID.

#### Task C — Python i18n-completeness scan

New script `py/check_i18n.py`: scans all `src/*.bas` and `src/*.cls` for string
literals that look like UI text and are not references to `basUIStrings` constants.
Produces a baseline report. Run after Tasks A and B to confirm no new inline
literals were introduced.

### Import plan (end of session)

| File | Reason |
|------|--------|
| `src/basUIStrings.bas` | LBL_* + CTRL_* constants (Tasks A + B combined) |
| `src/basBibleRibbonSetup.bas` | GetXxxLabel callbacks (Task A) |
| `src/aeRibbonClass.cls` | CTRL_* replacements + bug fix (Task B) |
| XML inject | getLabel= attributes (Task A) |

Python script `py/check_i18n.py` — no VBA import.

### Testing batch (after imports)

Single manual pass covering:
- All keytip badges (Alt+Y2 → B, C, V, [, ], ,, ., <, >, G, S, A)
- All four localised labels visible in ribbon UI
- Invalid Go attempt → verify comboBox now refreshes (bug fix confirmation)
- i18n scan script runs clean (no violations in current src/)

### py/check_i18n.py baseline results

Ribbon-active code is clean. 9 violations remain in `aeRibbonClass.cls`, all in archived
or legacy methods (About dialog, `GoToVerseSBL` stub, `GoToH1Direct` InputBox). These are
Step 7 targets — no action this session.

| Location | String | Category |
|----------|--------|----------|
| line 175 | `"Hello, adaept World!"` | About dialog — Step 7 |
| line 176 | `"adaeptMsg  = "`, `"About adaept"` | About dialog — Step 7 |
| line 212 | `"GoToVerseSBL - Parser not yet implemented."` | archived stub — Step 7 |
| line 228 | `"Enter a Book Name..."`, `"Go To Bible Book"` | `GoToH1Direct` InputBox — Step 7 |
| line 242 | `"Book not found! No Heading 1 matches: '"` | `GoToH1Direct` error — Step 7 |
| line 487 | `"...the truth shall make you free."` | About dialog quote — Step 7 |

### Status

**DONE — implemented 2026-04-20.**

Files changed: `src/basUIStrings.bas`, `src/basBibleRibbonSetup.bas`, `src/aeRibbonClass.cls`,
`src/basRibbonDeferred.bas`, `customUI14backupRWB.xml`, `py/check_i18n.py` (new).

Awaiting manual imports + testing batch (see session manifest).

---

## § 17 — Book comboBox text not refreshed after Go navigation

### Observed behaviour

User types "rev", presses Go. Navigation to Revelation is correct. Book comboBox
continues to display "rev". Switching to another ribbon tab and returning, or any
tab-away/return, does not fix the display.

Developer initially suspected a fundamental comboBox limitation.

### Root cause

Not a limitation — a missing `InvalidateControl` call.

`GetBookText` (line 517, `aeRibbonClass.cls`) is correctly implemented:

```vba
Public Function GetBookText(control As IRibbonControl) As String
    ' returns headingData(m_currentBookIndex, 0) — canonical book name
```

The `getText` callback is only called when the control is explicitly invalidated. After
a successful `GoToVerse` in `OnGoClick`, `CTRL_BOOK` was never invalidated. The
comboBox retained whatever the user had typed, indefinitely.

`InvalidateControl CTRL_BOOK` was also missing from `GoToH1Deferred`
(`basRibbonDeferred.bas`) — the Prev/Next Book path. That path correctly invalidated
`CTRL_NEXT_BOOK` and `CTRL_PREV_BOOK` (the nav buttons) but not the comboBox text.

### Fix

**`src/aeRibbonClass.cls` — `OnGoClick`:** one line added after `GoToVerse vsNum`:

```vba
    GoToVerse vsNum
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl CTRL_BOOK
```

**`src/basRibbonDeferred.bas` — `GoToH1Deferred`:** one line added:

```vba
    rc.GoToH1Direct
    rc.InvalidateControl CTRL_BOOK     ' ← added
    rc.InvalidateControl CTRL_NEXT_BOOK
    rc.InvalidateControl CTRL_PREV_BOOK
```

### Testing

After importing both files:
- Type a book abbreviation (e.g. "rev"), press Go → comboBox must show "Revelation"
- Use Prev Book / Next Book buttons → comboBox must update to new book name

### Status

**FIXED — 2026-04-20.** Requires import of `src/aeRibbonClass.cls` and
`src/basRibbonDeferred.bas`.

---

## § 18 — Book comboBox reverts to user-typed text after tab switch (known limitation)

### Observed behaviour

1. User types "gen" in Book comboBox.
2. Clicks Chapter (or Verse) — `OnBookChanged` fires, `m_ribbon.Invalidate` runs,
   `GetBookText` is called, comboBox updates to display "GENESIS". **Correct.**
3. User clicks another ribbon tab (e.g. Home), then returns to RWB.
4. Book comboBox shows "gen" — the original user-typed text.

### Investigation

`OnBookChanged` (line 571, `aeRibbonClass.cls`) calls `m_ribbon.Invalidate` (full
invalidation of all controls). This fires while the RWB tab is visible: `GetBookText`
is called and returns `headingData(m_currentBookIndex, 0)` — "GENESIS" is displayed.
The invalidation is consumed immediately.

When the user switches to another tab, no pending invalidation remains. On returning to
RWB, the ribbon re-renders the comboBox from its cached user-input state ("gen") rather
than issuing a fresh `getText` call — because no dirty flag is set.

### Root cause

Office Fluent ribbon `comboBox` controls maintain two distinct internal values:

| Value | Source | Lifetime |
|-------|--------|----------|
| User-input text | `onChange` / keyboard | Persists across tab switches |
| Callback text | `getText` on invalidation | Valid only while invalidation is pending |

On tab-switch, the control reverts to the user-input text. There is no `onShow` or
tab-activation callback in the customUI schema — no API hook to trigger a
re-invalidation when the tab becomes visible again.

### Distinction from § 17

§ 17 was a missing `InvalidateControl CTRL_BOOK` after `GoToVerse` — fixable because
the fix point (after Go) is reachable from code. **Fixed.**

This bug occurs between `OnBookChanged` and the next navigation commit. The `Invalidate`
in `OnBookChanged` already fires correctly; the limitation is the tab-switch re-render
behaviour, which has no addressable fix point in the standard ribbon API.

### Scope

Applies equally to Chapter and Verse comboBoxes. Less visible there because typed
numerals and canonical numerals are identical (typing "5" vs canonical "5"). Most
noticeable for Book because abbreviations differ from canonical names.

### Impact

**Cosmetic only.** Navigation state (`m_currentBookIndex`, `m_currentChapter`,
`m_currentVerse`) is always correct — `OnBookChanged` sets it before the display
updates. After the user presses Go, the canonical name is displayed (§ 17 fix).

### Status

**Known limitation — no VBA fix available.** The Office Fluent ribbon comboBox
preserves user-typed text across tab switches; re-invalidation requires a
tab-activation callback that the customUI schema does not provide.

If a future VSTO port is undertaken (§ 12), this limitation disappears: WinForms/WPF
controls can be updated programmatically at any time, and the `Ribbon.RibbonTab`
activated event is exposed via the managed Office object model.

---

## § 19 — Test 73: CountInvisibleCharacters added to Bible QA suite

### Motivation

Invisible Unicode characters (zero-width spaces, non-joiners, byte-order marks, word
joiners) are visually silent but can corrupt Word's Find/Replace results, style
normalization passes, and USFM export output. A systematic test ensures none are
introduced during editing.

The detection function already existed in `basTEST_aeBibleConfig.bas` as a standalone
diagnostic (`TestInvisible` / `CountInvisibleCharacters`). This session promoted it to
a numbered QA test so it runs automatically with every full test pass.

### Characters tested

| Code point | Name |
|------------|------|
| U+200B | ZERO WIDTH SPACE |
| U+200C | ZERO WIDTH NON-JOINER |
| U+200D | ZERO WIDTH JOINER |
| U+FEFF | ZERO WIDTH NO-BREAK SPACE (BOM) |
| U+2060 | WORD JOINER |

### Changes made

**`src/aeBibleClass.cls`** — six coordinated changes required by the test framework:

| Change point | Detail |
|-------------|--------|
| `MaxTests` constant | 72 → 73 (sizes `ResultArray` and `GetPassFailArray` arrays) |
| `Expected1BasedArray` values | `, 0` appended — expected = no invisible chars |
| `GetPassFail` Case 73 | `ResultArray(TestNum) = CountInvisibleCharacters()` |
| `RunBibleClassTests` sequence | `RunTest(73)` added after `RunTest(72)` |
| `RunTest` Case 73 | `Debug.Print` line with `"CountInvisibleCharacters"` label |
| `OutputTestReport` Case 73 | Same label written to `rpt/TestReport.txt` |
| New private function | `Private Function CountInvisibleCharacters() As Long` |

The class function returns `Long` (total count across all story ranges) rather than the
`String` report returned by the source function in `basTEST_aeBibleConfig`. This matches
the numeric comparison pattern used by every other test in the framework.

**Algorithm:** `UBound(Split(r.Text, targetChar))` equals the occurrence count because
splitting a string on a character that appears N times produces N+1 parts. Applied
across all story ranges (body, headers, footers, footnotes, text boxes).

**`src/basTEST_aeBibleConfig.bas`** — one change:

`CountInvisibleCharacters` visibility `Private` → `Public`. The class function is a
separate `Long`-returning variant; no naming conflict because the class resolves its
own method first. The public version in `basTEST_aeBibleConfig` remains available for
standalone use via `TestInvisible` or directly from the Immediate Window.

### Process documentation

`md/Adding_To_Bible_Test_Class.md` created. Contains:
- Architecture diagram of the full test dispatch chain
- 8-step checklist for adding any new test
- Decision guide: copy logic into class vs. call across modules
- Test 73 walkthrough as worked example
- Run instructions and expected pass/fail output formats

### Running

```
RUN_THE_TESTS(73)          ' standalone
RUN_THE_TESTS              ' full suite — test 73 included at position 73
```

Expected pass output:
```
PASS        Copy ()     Test = 73       0               0               CountInvisibleCharacters
```

If the test fails, `TestInvisible` in `basTEST_aeBibleConfig` provides a per-character
breakdown with Unicode labels and occurrence counts.

### Status

**IMPLEMENTED — 2026-04-20.** Import confirmed and test passing.

---

## § 20 — Bible Class Test Infrastructure: In-Depth Analysis

**2026-04-20 — requested after Test 73 import and verification.**

Three improvement areas analysed: (1) progress visibility, (2) iterative failure
location, (3) output formats (UTF-8 + Markdown).

---

### Area 1 — Progress Visibility and Stuck Detection

#### Current behaviour

The test loop in `RunBibleClassTests` calls `RunTest(n)` for each test with no
feedback before execution begins. The Immediate Window is silent during test
execution; the result line appears only after `GetPassFail` returns. For slow
tests (Test 42 — `CountBoldFootnotesWordLevel` — expected ~80 seconds), the
window is blank for over a minute with no way to distinguish "running" from
"crashed".

`bTimeAllTests = True` captures elapsed time per test but prints it after the
result line — "Routine Runtime: X.XX seconds" — so it confirms how long
something took, not that it is still running.

There is no `DoEvents` anywhere in the class. Word's message loop is starved
during long `Find.Execute` loops, making the application unresponsive to user
input and preventing the Immediate Window from refreshing.

`AppendToFile` opens and closes the report file on every single call. For a
full 73-test run with header/footer lines, this is approximately 80 separate
file I/O operations, each paying the open-seek-close cost on a path inside the
document's own folder.

#### Recommendations

**1a. Pre-test announcement in `RunTest`**

Add a `Debug.Print` line immediately before the `GetPassFail(num)` call in
`RunTest`. The function name is already in the `Select Case` below — duplicate
it into the pre-announce line:

```vba
Private Function RunTest(num As Integer, Optional SkipTest As Variant) As Boolean
    ...
    startTime = Timer

    Debug.Print ">> Starting Test " & num    ' <-- ADD THIS

    GetPassFail (num)
    ...
```

This is a one-line change. It does not require the Select Case label because
even the bare number is enough to know which test is running. The "Routine
Runtime" line that follows provides the elapsed duration. Together they bracket
each test without restructuring anything.

**1b. Yield to the message loop between tests**

`DoEvents` cannot be injected inside `Find.Execute` loops without risk (Word's
object model is not re-entrant during an active Find). However, calling
`DoEvents` once at the top of `RunTest`, after the pre-announce print but
before `GetPassFail`, yields control briefly so Word can repaint the Immediate
Window:

```vba
    Debug.Print ">> Starting Test " & num
    DoEvents                                 ' let Word repaint before blocking
    GetPassFail (num)
```

One `DoEvents` per test is safe — there is no active Find at that point.

**1c. Batch the report file writes**

Replace the 80-call `AppendToFile` pattern with a single write at the end of
the run. Accumulate all test report lines in a module-level or local
`Collection` or dynamic String buffer during the run, then write the complete
file once when all tests are done. The data is already available in
`ResultArray`, `GetPassFailArray`, and `oneBasedExpectedArray` at that point.

A minimal approach: collect lines into a `String` variable using `& vbCrLf`,
then write it in one `Open/Print/Close` block:

```vba
' At top of RunBibleClassTests: Dim reportBuf As String
' In each test: reportBuf = reportBuf & FormatReportLine(num) & vbCrLf
' At the end:   WriteBuf reportBuf, TestReportFileName
```

This reduces file I/O from ~80 operations to 1. The tradeoff is that a crash
mid-run produces no partial report. If incremental crash-safety is required,
write every 10 tests instead of every test.

---

### Area 2 — Iterative Failure Location

#### Current behaviour

Every Count* function returns a `Long` (total violation count). When a test
fails, the report shows the count and the expected value. The editor then must
manually search the entire document to find the first violation.

Example:
```
FAIL!!!!    Copy ()     Test = 3        8               0     CountSpaceFollowedByCarriageReturn
```

"8 occurrences" tells you how many are wrong but not where the first one is.
For a 900-page Bible, the first occurrence could be anywhere.

The Find-based count functions (`CountDoubleSpaces`, `CountSpaceFollowedByCarriageReturn`,
etc.) already traverse each match one at a time inside a `Do While .Execute`
loop. The first match is visited on the first loop iteration — capturing its
location at that point costs nothing beyond storing two integers.

#### Recommendation — first-hit hint array

Add a parallel class-level array `m_HintArray(1 To MaxTests) As String`.
Populate each slot during `GetPassFail` only when the count exceeds zero.
Print the hint immediately after a FAIL result in `RunTest`.

**Step 1 — Declare the hint array** (near `ResultArray` and `GetPassFailArray`):

```vba
Private m_HintArray(1 To MaxTests) As String
```

Reset it in `InitializeGlobalResultArrayToMinusOne`:

```vba
Dim i As Integer
For i = 1 To MaxTests : m_HintArray(i) = "" : Next i
```

**Step 2 — Capture first-hit location in Find-loop functions**

Pattern (shown for `CountDoubleSpaces`):

```vba
Private Function CountDoubleSpaces() As Integer
    ...
    Dim firstHit As String
    firstHit = ""
    Do While .Execute
        doubleSpaceCount = doubleSpaceCount + 1
        If doubleSpaceCount = 1 Then          ' first match only
            firstHit = "Para " & rng.Paragraphs(1).Range.Information(wdActiveEndAdjustedPageNumber) _
                      & " pg ~" & rng.Information(wdActiveEndAdjustedPageNumber)
        End If
        rng.Collapse wdCollapseEnd
    Loop
    CountDoubleSpaces = doubleSpaceCount
    ' Caller sets m_HintArray — see GetPassFail pattern below
End Function
```

Because the function is private and returns only a Long, the simplest
integration is to have each Count function set a shared module-level variable
(`m_lastHint As String`) and have the `GetPassFail` Case block copy it into
`m_HintArray(TestNum)` after the call:

```vba
Case 1
    ResultArray(TestNum) = CountDoubleSpaces()
    m_HintArray(TestNum) = m_lastHint      ' set by CountDoubleSpaces on first hit
```

**Step 3 — Print hint after FAIL in `RunTest`**

```vba
' After the existing Debug.Print result line:
If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then
    Debug.Print , , , "  >> First hit: " & m_HintArray(num)
End If
```

**Scope of change**

Paragraph-iterating functions (Tests 26–31, 36–38, 43–44, etc.) use `For Each
para In doc.Paragraphs` loops — same pattern: capture on first iteration.
Unicode search functions (Tests 66–71) and the `Split`-based
`CountInvisibleCharacters` (Test 73) can provide a story-type label as the
hint ("Main body" / "Header story" / "Footnote story") without page info.

Tests that call `CheckAllHeaders` or audit-style functions already return
structured data to files — their hint could be "see rpt/HeadingLog.txt".

This change does not alter the test count, the expected-value comparison, or
the pass/fail decision. It is additive hint metadata.

---

### Area 3 — Output Formats: UTF-8 and Markdown

#### 3a — UTF-8 output via aeLoggerClass

`aeLoggerClass` is already in the project. Its interface:

| Method | Effect |
|--------|--------|
| `Log_Init(path)` | Creates/overwrites file with UTF-8 BOM; writes session header (timestamp, user, machine) |
| `Log_Write(msg)` | Prepends `HH:nn:ss \|` timestamp; rewrites full buffer to file (crash-safe) |
| `Log_Close()` | Writes END marker |

A UTF-8 report alongside the existing ASCII `TestReport.txt` requires:
- A second report path constant, e.g. `TestReportUTF8FileName = "TestReportUTF8.txt"`
- A module-level `Private m_log As Object` instance
- `m_log.Log_Init` at the start of the test run (after `vbYes`)
- `m_log.Log_Write` once per test (same data as `AppendToFile`)
- `m_log.Log_Close` at the end

Because `aeLoggerClass` uses late binding (`As Object` / `CreateObject`), no
reference changes are needed. The existing `AppendToFile` calls remain
untouched — the logger is additive.

The UTF-8 output is particularly valuable for tests 52–71 (contraction and
Unicode sequence tests) where the function label contains non-ASCII characters
that `AppendToFile` writes as `?` in the ASCII stream.

#### 3b — Markdown report

A Markdown report makes test results readable in any viewer (VS Code, GitHub,
browser) without the columnar fixed-width formatting that the current text file
requires.

Proposed format (`rpt/TestReport.md`):

```markdown
# Bible QA Test Report
Generated: 2026-04-20 14:30:00
BibleClass VERSION: x.x  Word: 16.0.xxxxx

| Status | Test | Result | Expected | Function |
|--------|------|--------|----------|----------|
| PASS | 1 | 0 | 0 | CountDoubleSpaces |
| FAIL | 3 | 8 | 0 | CountSpaceFollowedByCarriageReturn |
...

**Total Runtime:** 4.23 seconds
```

Implementation approach: add a `Private Function FormatMdReportLine(num As
Integer) As String` that returns the `| ... |` row, and a `WriteMarkdownReport`
private sub that opens a single file, writes the header table, all test rows,
and the footer. Call it once at the end of `RunBibleClassTests` alongside
`RunTotalTimeTestSession`.

The Markdown report does not replace the existing `TestReport.txt` — it is an
additional output. Both can coexist.

---

### Summary Table

| Area | Change | Scope | Risk |
|------|--------|-------|------|
| Progress: pre-announce | `Debug.Print ">> Starting Test " & num` before `GetPassFail` in `RunTest` | 1 line | None |
| Progress: DoEvents | `DoEvents` after pre-announce, before `GetPassFail` | 1 line | None |
| Progress: batch file I/O | Accumulate report buffer, write once | Medium refactor | Partial report lost on crash |
| First-hit location | `m_HintArray`, `m_lastHint`, hint print in `RunTest` | Medium addition | None (purely additive) |
| UTF-8 output | `aeLoggerClass` instance alongside `AppendToFile` | Low — additive | None |
| Markdown output | `FormatMdReportLine` + `WriteMarkdownReport` | Low — additive | None |

Highest value, lowest risk changes: pre-announce (1a) and DoEvents (1b). These
two lines resolve the "is it stuck?" problem immediately.

### Status

**ANALYSED — 2026-04-20.** Implementation plan approved — see § 21.

---

## § 21 — Bible Class Test Infrastructure: Implementation Plan

**2026-04-21 — step-by-step, one approval per step.**

| Step | Description | Status |
|------|-------------|--------|
| 1 | Pre-test announcement — `Debug.Print ">> Starting Test " & num` | **DONE — 2026-04-21** |
| 2 | DoEvents between tests | **DONE — 2026-04-21** |
| 3 | Batch AppendToFile writes | **DONE — 2026-04-21** |
| 4 | First-hit hint infrastructure (arrays + print hook) | Pending |
| 5 | First-hit capture in Count functions (failing tests first) | Pending |
| 6 | UTF-8 output via aeLoggerClass | Pending |
| 7 | Markdown report | Pending |

### Step 1 — Pre-test announcement

**File:** `src/aeBibleClass.cls`

One line added in `RunTest` (line 1004), immediately before `GetPassFail(num)`:

```vba
Debug.Print ">> Starting Test " & num
GetPassFail (num)
```

**Test:** `RUN_THE_TESTS(42)` — Immediate Window must show `>> Starting Test 42`
before the result line.  
**Pass criteria:** Two lines in order: announce, then result.

**Status: IMPLEMENTED — 2026-04-21. Import `src/aeBibleClass.cls` and run
`RUN_THE_TESTS(42)` to verify.**

### Step 2 — DoEvents between tests

**File:** `src/aeBibleClass.cls`

One line added in `RunTest`, between the announce print and `GetPassFail(num)`:

```vba
Debug.Print ">> Starting Test " & num
DoEvents
GetPassFail (num)
```

**Status: IMPLEMENTED — 2026-04-21.**

### Step 3 — Batch AppendToFile writes

**File:** `src/aeBibleClass.cls`

Changes:
- `Private m_ReportBuf As String` declared at class level (reset to `""` in `InitializeGlobalResultArrayToMinusOne`)
- `Private Sub BufAppend(text As String)` — appends `text & vbCrLf` to `m_ReportBuf`
- `Private Sub FlushReportBuf()` — opens `TestReportFileName` for Append once, writes full buffer, closes, clears buffer
- All 78 `AppendToFile TestReportFileName, expr` calls replaced with `BufAppend expr`
- `FlushReportBuf` called once at end of `RunBibleClassTests` (after last `BufAppend`, before `RunBibleClassTests = True`)
- `AppendToFile("TotalTimeReport.txt", ...)` unchanged — separate file, not batched

**Test:** Full `RUN_THE_TESTS` — `rpt/TestReport.txt` content must be identical to
the pre-change baseline. Run time equal or faster.  
**Pass criteria:** File content unchanged; no regression in pass/fail counts.

**Status: IMPLEMENTED — 2026-04-21. Import `src/aeBibleClass.cls` and run full
`RUN_THE_TESTS` to verify.**

**File:** `src/aeBibleClass.cls`

One line added in `RunTest`, between the announce print and `GetPassFail(num)`:

```vba
Debug.Print ">> Starting Test " & num
DoEvents
GetPassFail (num)
```

`DoEvents` is called once per test at a safe point — no active Find is running.
Yields to Word's message loop so the Immediate Window repaints and the ribbon
remains clickable during the blocking Find execution.

**Test:** `RUN_THE_TESTS(42)` — Word must not enter "not responding" state during
the ~80-second run. Announce line must appear in Immediate Window before the
block begins.  
**Pass criteria:** UI responsive throughout; announce visible immediately.

**Status: IMPLEMENTED — 2026-04-21. Import `src/aeBibleClass.cls` and run
`RUN_THE_TESTS(42)` to verify.**

---

## § 22 — Bug: Test 36 — Stop in CountFooterParagraphsWithFooterStyle

**2026-04-21**

### Symptom

`RUN_THE_TESTS` halts mid-run with a VBA `Stop` statement inside
`CountFooterParagraphsWithFooterStyle`. Execution breaks at the first footer
paragraph found that uses the built-in Word "Footer" style instead of the
project style "TheFooters".

### Root cause

The function was written as a diagnostic probe, not a counting function. On the
first match it selects the paragraph, prints a message, and executes `Stop` to
drop the developer into the editor at that location. This was useful during
initial investigation but makes the test non-runnable in a full suite.

```vba
' Original — halts on first match, never returns a count
If para.style = "Footer" Then
    Count = Count + 1
    para.Range.Select
    Debug.Print "Found paragraph with Footer style. Stopping at this location."
    Stop       ' breaks full-suite run
End If
```

### Rule being enforced

All footer paragraphs must use the project paragraph style `"TheFooters"`.
The built-in Word style `"Footer"` is not used in this document. Any paragraph
still carrying `"Footer"` is a gap in style normalization.

### Fix routine

`ReapplyTheFootersToAllFooters` in `basTEST_aeBibleTools.bas`:

- Iterates every section, every footer, every paragraph
- Applies `p.style = "TheFooters"` unconditionally
- Logs each updated paragraph (previous style, ASCII value, hex) to the
  Immediate Window for audit
- Does not touch footer content or page numbering

Note: `FixTheFooters` in `basAddHeaderFooter.bas` is a different tool — it
rebuilds footer content and consecutive page numbering from the cursor position.
It is not the correct fix for a style-only normalization issue.

### Changes made — `src/aeBibleClass.cls`

**`CountFooterParagraphsWithFooterStyle`** — `Stop`, `Select`, and diagnostic
`Debug.Print` removed; function now counts all violations and returns cleanly:

```vba
If para.style = "Footer" Then
    Count = Count + 1
End If
```

**`RunTest` Case 36 and `OutputTestReport` Case 36** — fix routine appended to
the function label so it appears in both the Immediate Window and `TestReport.txt`:

```
CountFooterParagraphsWithFooterStyle  FIX: ReapplyTheFootersToAllFooters
```

### Expected value

Expected = `0`. The test will FAIL until `ReapplyTheFootersToAllFooters` is
run and all footer paragraphs carry `"TheFooters"`. This is the correct
enforcement posture — the test is red until the document is clean.

### Workflow when test fails

```
RUN_THE_TESTS(36)             ' confirms failure, shows count
ReapplyTheFootersToAllFooters ' fixes all sections; logs each change to Immediate Window
RUN_THE_TESTS(36)             ' must now PASS with result = 0
```

### Status

**IMPLEMENTED — 2026-04-21. Import `src/aeBibleClass.cls` and run
`RUN_THE_TESTS(36)` to verify count is returned without stopping.**

**Verified — 2026-04-21. Test 36 PASS (result = 0, runtime = 0.07 s).**

---

## § 23 — Bug: Test 72 — Word Not Responding / Force Close

**2026-04-21**

### Symptom

`RUN_THE_TESTS` (full suite) causes Word to become unresponsive when Test 72
(`HasLeftAlignedParagraph(18, 931)`) executes. The application requires force
close and the final report is not generated.

### Two independent failure modes

#### Failure mode 1 — `GoToAdjustedPage` infinite loop

`GoToAdjustedPage` positions the selection by calling `sel.GoTo wdGoToNext`
in a loop until `wdActiveEndAdjustedPageNumber` equals `targetPage`:

```vba
sel.HomeKey wdStory
Do
    If sel.Information(wdActiveEndAdjustedPageNumber) = targetPage Then Exit Do
    sel.GoTo What:=wdGoToPage, Which:=wdGoToNext
Loop
```

There is no exit condition for the case where `wdGoToNext` stops advancing.
At the last page of the document Word silently keeps the selection on the final
page — `wdGoToNext` does not raise an error and does not move the cursor.
`wdActiveEndAdjustedPageNumber` never changes, the `If` condition is never
met, and the loop runs forever. This is the force-close scenario.

Secondary fragility: if the document has Roman-numeral front matter (e.g.,
pages i–xvii), the adjusted page number 18 appears twice — once in the front
matter (Roman numeral xviii = numeric 18) and once in the Bible body. The loop
exits at the first hit, which may be the wrong page. Not currently triggered
for startPage=18 if front matter is fewer than 18 pages, but is a structural
risk.

#### Failure mode 2 — format-only Find across 900+ pages (no DoEvents)

After positioning to page 18, `HasLeftAlignedParagraph` runs a format-only
`sel.Find.Execute` with no text pattern, `Format = True`, and
`ParagraphFormat.Alignment = wdAlignParagraphLeft`:

```vba
With sel.Find
    .ClearFormatting
    .Text = ""
    .Forward = True
    .Wrap = wdFindStop
    .Format = True
    .ParagraphFormat.Alignment = wdAlignParagraphLeft
End With
found = sel.Find.Execute
```

The expected value for Test 72 is `0` — the document should contain no
left-aligned paragraphs in pages 18–931; all body paragraphs should be
justified. When the document is clean the Find scans all remaining pages
without a match before stopping at `wdFindStop`. A format-only scan checks
every paragraph's formatting properties one by one with no text shortcut.
On a 16,471-paragraph document with no `DoEvents` this starves Word's message
loop → "not responding."

Note: if the document is NOT clean (a left-aligned paragraph exists early),
the Find terminates on the first hit and returns in under a second. The slow
path is the passing case — the test is hardest to run when the document is
correct.

### Immediate fix — add Test 72 to SkipTestArray

`SkipTestArray = Array(42, 51, 72)` — Test 72 is skipped in the full suite
(printed as `SKIP!!!!`) and runs only when called explicitly as
`RUN_THE_TESTS(72)`. This unblocks the full suite immediately.

### Full fix (deferred — see game plan)

Two changes required in `src/aeBibleClass.cls`:

**1. `GoToAdjustedPage` — stall detection:**

```vba
Dim prevPage As Long
prevPage = -1
Do
    curPage = sel.Information(wdActiveEndAdjustedPageNumber)
    If curPage = targetPage Then Exit Do
    If curPage = prevPage Then Exit Do  ' no longer advancing — target not found
    prevPage = curPage
    sel.GoTo What:=wdGoToPage, Which:=wdGoToNext
Loop
```

**2. `HasLeftAlignedParagraph` — replace format-only Find with paragraph
iteration** with `DoEvents` every 500 paragraphs. This also enables a
first-hit hint (page, text excerpt) for Step 5 of the infrastructure plan
(§ 21).

### Status

**IMMEDIATE FIX IMPLEMENTED — 2026-04-21.**
`SkipTestArray = Array(42, 51, 72)` in `src/aeBibleClass.cls`.
Import and run full `RUN_THE_TESTS` — Test 72 must print `SKIP!!!!` and the
suite must complete without force close.
Full rewrite deferred pending game plan review of complete test results.

**Verified — 2026-04-21. Test 72 SKIP!!!!. Full suite completed (483 s).**

---

## § 24 — Bug: FlushReportBuf — Error 52 Bad File Name or Number

**2026-04-21**

### Symptom

After the full suite completes and all results are printed to the Immediate
Window, the final `rpt/TestReport.txt` is not written. The error handler in
`RunBibleClassTests` fires:

```
ERROR in aeBibleClass.RunBibleClassTests | Erl: 0 | Err: 52 | Bad file name or number
```

The Immediate Window output is complete and correct — only the file write
fails.

### Root cause

`TestReportFileName` is set in `Expected1BasedArray` as a **full path**:

```vba
TestReportFileName = doc.Path & "\rpt\TestReport.txt"
' e.g. C:\adaept\...\rpt\TestReport.txt
```

`FlushReportBuf` (introduced in Step 3) then prepended `folderPath & "\"` on
top of it unconditionally:

```vba
folderPath = ActiveDocument.Path & "\rpt"
finalPath = folderPath & "\" & TestReportFileName
' result: C:\adaept\...\rpt\C:\adaept\...\rpt\TestReport.txt  ← invalid
```

The original `AppendToFile` handled both a bare filename and a full path
correctly via an `InStr` check. That guard was not carried forward into
`FlushReportBuf`.

### Fix — `src/aeBibleClass.cls` — `FlushReportBuf`

Same `InStr` path check as `AppendToFile`:

```vba
If InStr(TestReportFileName, "\") = 0 Then
    finalPath = folderPath & "\" & TestReportFileName  ' bare filename
Else
    finalPath = TestReportFileName                     ' already a full path
End If
```

### Status

**IMPLEMENTED — 2026-04-21. Import `src/aeBibleClass.cls` and run full
`RUN_THE_TESTS` — `rpt/TestReport.txt` must be written on completion.**

**Verified — 2026-04-21. Report written successfully (230 s run).**

---

## § 25 — Test Report: Completeness and Limitations

**2026-04-21 — first complete report after Steps 1–3 and bug fixes §§ 22–24.**

### Report as written

73 tests total: 3 SKIP (42, 51, 72), 70 executed.

| Status | Count | Tests |
|--------|-------|-------|
| PASS | 50 | — |
| FAIL | 17 | 3, 16, 26, 27, 28, 29, 32, 33, 34, 35, 47, 49, 50, 55, 64, 70, 71 |
| SKIP | 3 | 42, 51, 72 |

### Limitations and gaps

**1. No error status.** Test 13 shows `PASS` in the report but threw an
Overflow error during execution (see § 26). The report has no column or marker
for "function errored" — a failed function that returns 0 is indistinguishable
from a clean zero-violation result. Any test with expected = 0 can silently
false-pass on error.

**2. Timing absent from report.** Per-test runtime appears in the Immediate
Window (`Routine Runtime: X.XX seconds`) but is not written to `TestReport.txt`.
The report records only the total runtime. Slow outliers (Test 73: 50 s) are
not visible in the file.

**3. Unicode labels corrupted.** Tests 52–71 (contraction and Unicode sequence
tests) show `?` in place of Unicode characters in the function label column.
`AppendToFile` writes ASCII; multi-byte code points are replaced with `?`.
Example: `CountContraction("wouldn?t")` instead of `CountContraction("wouldn't")`.
Step 6 (UTF-8 via aeLoggerClass) will address this.

**4. SKIP rows show result = -1.** Skipped tests display `result = -1` (the
unrun sentinel) alongside `expected = 0`. The -1 is correct internally but is
confusing in the report — it implies the test ran and returned -1 rather than
was intentionally skipped.

**5. Expected-value block is the entire first 5 lines.** The expected-value
dump at the top of the report (Tests 1–73) duplicates the per-test rows and
serves no readability purpose in the report file. It was useful during
development but adds noise now.

**6. No document identification.** The report does not record which document
was tested (filename, full path, or word count). Two runs on different documents
produce identical headers.

---

## § 26 — Bug: Test 13 — Overflow, False PASS, Silent Error Masking

**2026-04-21**

### Symptom

Test 13 (`CountEmptyParasWithNoThemeColor`) prints an Overflow error to the
Immediate Window during the run but shows `PASS` in both the Immediate Window
and `TestReport.txt`. The error is not captured in the report. The result of 0
matches the expected value of 0, but the function never executed.

### Root cause — dead variable triggers overflow

`CountEmptyParasWithNoThemeColor` declares:

```vba
Dim emptyParaCount As Integer
Dim totalParaCount As Integer    ' ← dead variable; never used in logic
...
totalParaCount = ActiveDocument.Paragraphs.Count   ' 33,854 → Overflow
```

`Integer` in VBA is 16-bit: maximum value 32,767. The document has 33,854
paragraphs. Assigning `Paragraphs.Count` to an `Integer` raises Error 6
(Overflow) immediately — before any paragraph is iterated.

`totalParaCount` is never used in logic. All `Debug.Print` lines that
referenced it are commented out. It is a dead variable left over from
diagnostic development. Removing it eliminates the overflow entirely.

### Root cause — PROC_ERR masks errors as false PASSes

Every Count function in the class uses the same error handler pattern:

```vba
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountXxx | ..."
    Resume PROC_EXIT
```

`Resume PROC_EXIT` exits the function with its return value in whatever state
it was at the point of error. For an `Integer` or `Long` return type, the
uninitialised value is 0. If the expected result for that test is also 0, the
comparison passes and `GetPassFailArray(n) = "PASS    "` is written to the
report. The error message goes only to the Immediate Window.

**This means: any test where expected = 0 cannot be trusted if the function
could error for any reason.**

Tests with expected = 0 in the current suite: 1, 2, 4, 5, 6, 7, 8, 11, 12,
13, 14, 17, 19, 20, 21, 23, 26 (partial), 30, 36, 39, 40, 41, 43, 44, 45,
46, 48, 52–54, 56–63, 66–70, 73 — the majority of the suite.

### Systematic risk — Integer overflow in 32 Count functions

Every Count function except `CountTotalParagraphs` is declared `As Integer`.
Internal count variables are also `As Integer` in many functions. VBA Integer
is 16-bit (max 32,767). Three risk categories:

| Risk | Description | Functions affected |
|------|-------------|-------------------|
| **High** | Variable assigned `Paragraphs.Count` (33,854) directly | Test 13 — `totalParaCount` |
| **Medium** | Return value could exceed 32,767 if violations are numerous | All 32 `As Integer` functions |
| **Low** | Internal counters over header/footer ranges only | Tests 28, 29, 30, 31 — counts sections × paragraphs per header, bounded by 147 × ~3 |

Header/footer functions (Tests 28–31) iterate section headers, not body
paragraphs — their `totalCount` variables accumulate at most a few hundred and
are safe. Body-paragraph functions are safe as long as violations stay below
32,767, which is true for a clean document. But the type is wrong regardless.

### Recommended fixes

**Fix A — Test 13 immediate:** remove dead `totalParaCount` variable and its
assignment. Change `emptyParaCount As Integer` → `As Long`. Change function
return type `As Integer` → `As Long`.

**Fix B — Systematic return types:** change all 32 Count function declarations
from `As Integer` to `As Long` and all internal `Dim Count As Integer`,
`Dim totalCount As Integer` to `As Long`. This is a class-wide find-and-replace
with verification.

**Fix C — Error sentinel:** in the PROC_ERR handler of every Count function,
set the return value to `-1` before `Resume PROC_EXIT`:

```vba
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountXxx | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    CountXxx = -1    ' ← sentinel: ensures any error shows FAIL, not PASS
    Resume PROC_EXIT
```

Since expected values are never -1, this guarantees every errored test shows
`FAIL!!!!` in the report rather than silently matching an expected value of 0.
The error message in the Immediate Window identifies the cause; the FAIL in
the report makes it visible without scrolling.

### Priority

Fix A (Test 13) is trivial and should be done first. Fixes B and C are
class-wide and should be done together as one import. They constitute part of
the game plan for the infrastructure improvement pass.

### Status

**ANALYSED — 2026-04-21. No code changes yet. Fixes A, B, C deferred to
game plan.**
