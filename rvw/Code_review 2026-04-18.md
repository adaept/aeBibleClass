# Code Review - 2026-04-18

## Carry-Forward from 2026-04-13

Continues from `rvw/Code_review - 2026-04-13.md`.

---

## § 1 — Status of Previous Session (2026-04-13) Carry-Forward

### Completed items (closed)

| Item | Detail |
|------|--------|
| Bug 12 | Tab trap at last Book/Chapter/Verse | **CLOSED** |
| Bug 13 | Tab after Chapter steals focus to document | **CLOSED** |
| Bug 14 | Alt+R triggers Review / Word Count | **CLOSED — keytip="RW" removed** |
| Bug 15 | RWB tab unreachable from keyboard | **CLOSED — Y2 confirmed** |
| Bug 17 | Book selection scrolls document | **CLOSED** |
| Bug 18 | GoToChapter uses ScrollIntoView (not .Select) | **CLOSED** |
| Bug 19 | Next/Prev Book navigates from stale cursor | **CLOSED — m_currentBookIndex used** |
| Bug 20 | Tab from Chapter (inline ScrollIntoView) | **CLOSED** |
| Bug 21 | Deferred GoToChapter steals ribbon focus | **CLOSED — ExecutePendingChapter is no-op stub** |
| Bug 22b | Snap-back to previous verse | **CLOSED — Range.Select in button handlers** |
| Bug 23b | Tab after multi-digit chapter → document | **CLOSED — no self-invalidation** |
| Bug 23c | PrevVerseButton blocks Tab path | **CLOSED — always-enable at boundary** |
| Bug 24 | First-load Tab to document after book selection | **CLOSED — superseded by Fix 7** |
| Bug 25a | First-load verse Tab still goes to document | **CLOSED — Fix 7 (GetChapterEnabled/GetVerseEnabled always True)** |
| Bug 25b | GoToVerse wrong verse in Psalm 119 | **CLOSED — GoToVerseByScan unified path** |
| Bug 26 | Tab after chapter entry goes to document | **CLOSED — Fix 9 + § 28 architecture** |
| Bug 28 | Invalid Chapter/Verse leaves stale display | **CLOSED — deferred ResetChapterDisplay/ResetVerseDisplay** |
| Bug 29 | First-load Tab regression from Rule 2a Step 1 | **CLOSED — display/state separation** |
| Fix 8 | Pre-built chapterData array — O(1) chapter lookup | **CLOSED** |
| Step 5 | GoToVerse timing test (Psalm 119:176) | **CLOSED — GoToVerseByScan confirmed correct** |

### Open items (carry-forward)

| Item | Detail | Status |
|------|--------|--------|
| Bug 16 | Keytip badges end-to-end test | **PENDING — re-test after XML re-import** |
| Bug 22 | First nav to distant book slow (~10s) | **KNOWN LIMITATION — one-time session cost** |
| Bug 23a | Layout delay Psalms (~6s first nav) | **KNOWN LIMITATION — same class as Bug 22** |
| Bug 27 | Enter in Chapter does not navigate | **KNOWN LIMITATION — onChange cannot distinguish Enter from keystroke** |
| Step 7 | OLD_CODE cleanup — dead stubs | **PENDING** |
| WarmLayoutCache rewrite | Replace Range.Select with ScrollIntoView; re-enable deferred warm | **FUTURE** |
| Search tracking reset | Test Selection.SetRange from OnTime context | **FUTURE** |
| Layout pre-warm | Deferred ScrollIntoView warm at open | **FUTURE** |

---

## § 2 — Architecture Context (Default-Fill + Action-Gate, from § 28)

The navigation model adopted at the end of the 2026-04-13 session:

| Rule | Description |
|------|-------------|
| 1 | Navigation requires all three fields (Book, Chapter, Verse) to be filled |
| 2 | Book is always required — no default |
| 2a | When Book is confirmed, Chapter and Verse are immediately set to 1 |
| 3 | Tab past Chapter accepts the displayed value (1 if default, or user-entered) |
| 4 | Tab past Verse accepts the displayed value (1 if default, or user-entered) |
| 5 | Navigation fires only after B/C/V are all filled (verse confirmation trigger) |
| 6 | Prev/Next buttons are always enabled; click handlers guard boundaries |
| 7 | Prev/Next button presses update all three B/C/V fields |

**Key invariants in current code:**
- `GetChapterEnabled` and `GetVerseEnabled` always return `True` (Fix 7) — comboBoxes are always Tab stops
- `GetPrevChapterEnabled` and `GetNextChapterEnabled` return `(m_currentChapter > 0)`
- `GetPrevVerseEnabled` and `GetNextVerseEnabled` return `(m_currentChapter > 0)`
- `OnChapterChanged` does NOT call `InvalidateControl` (Bug 26 fix)
- `ExecutePendingChapter` invalidates verse row controls via `Application.OnTime` after Tab routing completes

---

## § 3 — New Bugs Reported: #597, #598, #599

Test results received 2026-04-18. Three bugs from the current session:

```
' #599 - First load Gen tab tab tab 119 tab sets focus in docm, second use tab will go through all controls [bug]
' #598 - Gen tab fills C/V with 1/1 but does not enable C/V Prev/Next buttons [bug]
' #597 - New Search should set the focus in cmbBook and not the docm [bug]
```

---

## § 4 — Bug #598: C/V Prev/Next Buttons Disabled After Book Selection

### Symptom

Typing `Gen` in `cmbBook` and pressing Tab:
- `cmbChapter` and `cmbVerse` both display `1` (correct via `GetChapterText` / `GetVerseText`)
- `PrevChapterButton`, `NextChapterButton`, `PrevVerseButton`, `NextVerseButton` remain **disabled**

### Root cause

`OnBookChanged` set `m_currentChapter = 0` and `m_currentVerse = 0` to implement the
display/state separation from the Bug 29 fix. `GetChapterText` returned `"1"` via its
middle branch (`ElseIf m_currentBookIndex > 0 Then GetChapterText = "1"`), decoupling
the visual from the state. However the Prev/Next enabled callbacks use the state:

```vba
Public Function GetPrevChapterEnabled(control As IRibbonControl) As Boolean
    GetPrevChapterEnabled = (m_currentChapter > 0)   ' False when m_currentChapter = 0
End Function
```

`m_currentChapter = 0` → all four Prev/Next Chapter and Verse buttons disabled.
The display says "1" but the state says "nothing confirmed" — buttons stay off.

### Fix applied (2026-04-18)

In `OnBookChanged`, changed:

```vba
' Before:
m_currentChapter = 0   ' display "1" via GetChapterText; keep 0 so chapter buttons stay disabled
m_currentVerse = 0     ' display "1" via GetVerseText; keep 0 so verse buttons stay disabled

' After:
m_currentChapter = 1   ' default chapter 1 — enables Prev/Next Chapter buttons (#598)
m_currentVerse = 1     ' default verse 1 — enables Prev/Next Verse buttons (#598)
```

**Effect on existing GetChapterText / GetVerseText:**

With `m_currentChapter = 1`, the **first** branch fires:
```vba
If m_currentChapter > 0 Then GetChapterText = CStr(m_currentChapter)   ' returns "1"
```
The display result is identical ("1"). The middle branch (`ElseIf m_currentBookIndex > 0`)
becomes dead code in the normal flow but is harmless.

**Effect on click guards:**

All boundary guards already handle chapter = 1 and verse = 1 correctly:
- `OnPrevChapterClick`: `If m_currentChapter > 1 Then GoToChapter ...` — no-op at 1
- `OnNextChapterClick`: `If m_currentChapter < ChaptersInBook(...)` — navigates to 2
- `OnPrevVerseClick`: `If m_currentVerse > 1 Then GoToVerse ...` — no-op at 1
- `OnNextVerseClick`: `If m_currentVerse < VersesInChapter(...)` — navigates to 2

**Effect on GoToVerse guard:**

```vba
If m_currentChapter = 0 Then GoTo PROC_EXIT   ' assert: confirm chapter first
```

With chapter = 1 this guard passes — correct, chapter is now confirmed at 1.

### Potential interaction with Bug 29

Bug 29 (2026-04-13) was triggered because setting `m_currentChapter = 1` in `OnBookChanged`
caused the deferred `m_ribbon.Invalidate` to enable `PrevChapterButton` / `NextChapterButton`.
Tab from `cmbChapter` then hit `NextChapterButton` instead of `cmbVerse` on first load.

The display/state separation fix (Bug 29 resolution) explicitly kept `m_currentChapter = 0` to
prevent this. Reverting to `m_currentChapter = 1` reintroduces the same structural condition.

**Why Bug #598 supersedes Bug 29:** Bug #599 ("second use goes through all controls") implies
the expected Tab path includes all Prev/Next buttons as active Tab stops. If Bug #599 is fixed
by making Prev/Next buttons always-enabled (same pattern as Fix 7 for the comboBoxes), then the
Tab path is consistent regardless of `m_currentChapter` value. Bug 29 is absorbed by the #599 fix.

**Status: pending test.** If Bug 29 re-emerges (Tab from cmbChapter hits NextChapterButton on
first load), Bug #599 must be fixed before or concurrently with #598.

### Status

| Item | Status |
|------|--------|
| Bug #598 — C/V Prev/Next disabled after book selection | **CONFIRMED — Buttons now Enabled — 2026-04-18** |

---

## § 5 — Bug #599: First-Load Tab Goes to Document; Second Use Works

### Test results (2026-04-18)

**`rev [Tab] [Tab]` (after #598 fix):**
- Chapter = 1, Verse = 1, all Prev/Next enabled ✓
- Continued Tab reaches New Search ✓

**`gen [Tab] [Tab] [Tab] 119 [Tab]` — first load:**
- Last Tab goes to document. Chapter stays at 1.
- `119` was entered and reset (bad data — Genesis has 50 chapters).
- Exact intermediate Tab path after `cmbVerse` not yet confirmed.

**`gen [Tab] [Tab] [Tab] 119 [Tab]` — second use (no New Search between runs):**
- Tab 1 → `NextBookButton` (enabled — prior interaction updated cache)
- Tab 2 → `PrevChapterButton`
- Tab 3 → `cmbChapter` — user types `119` → bad data → display resets to `1`
- Last Tab → `NextChapterButton` ("ends up at Next") ✓

### What the tests confirm

The Tab count to reach `cmbChapter` differs between loads:

| Load | `NextBookButton` cache | Tabs to `cmbChapter` | Last Tab after `119` |
|------|------------------------|----------------------|----------------------|
| First | DISABLED (`m_currentBookIndex = 0` at initial render) | 1 | Goes to document — intermediate path not confirmed |
| Second use | ENABLED (prior run updated cache) | 3 | `NextChapterButton` |

The root cause is `NextBookButton` (and `PrevBookButton`) being disabled in the initial render
cache, reducing the Tab count from 3 to 1. With fewer Tab stops between `cmbBook` and `cmbChapter`,
the sequence `[Tab][Tab][Tab] 119 [Tab]` overshoots `cmbChapter` on first load and exits the
ribbon to the document before `119` can be entered there.

### Root cause

The ribbon builds its Tab-routing cache at **initial render** before `OnRibbonLoad`. At that
moment `m_currentBookIndex = 0`, so:

| Control | Initial cache |
|---------|---------------|
| `cmbBook` | ENABLED |
| `NextBookButton` / `PrevBookButton` | **DISABLED** |
| `PrevChapterButton` / `NextChapterButton` | DISABLED |
| `cmbChapter` | ENABLED (always — Fix 7) |
| `PrevVerseButton` / `NextVerseButton` | DISABLED |
| `cmbVerse` | ENABLED (always — Fix 7) |
| `NewSearchButton` | DISABLED |
| `adaeptButton` | ENABLED |

`OnRibbonLoad` calls `m_ribbon.Invalidate` synchronously, but `m_currentBookIndex` is still 0
at that point — cache unchanged. When the user types `gen` + Tab, `OnBookChanged` fires and
calls `m_ribbon.Invalidate` from within `onChange` — **deferred**, fires after the current
event cycle including Tab routing. Tab routing for the first Tab reads the stale cache where
`NextBookButton` is DISABLED, giving only 1 hop to `cmbChapter` instead of 3.

On second use (no New Search): the previous `gen` interaction left `m_currentBookIndex = 1`
and its deferred Invalidate updated the cache — `NextBookButton` is now ENABLED. The next
`gen` + Tab fires `OnBookChanged` again but Tab routing reads the **already-updated** cache
from the prior session, giving 3 hops to `cmbChapter`.

### Proposed fix

Extend "always-enable at boundary" (Fix 7 pattern) to all **six** Prev/Next buttons, including
`GetPrevBkEnabled` and `GetNextBkEnabled`:

```vba
' GetPrevBkEnabled / GetNextBkEnabled — before:
GetPrevBkEnabled = (m_currentBookIndex > 0)
GetNextBkEnabled = (m_currentBookIndex > 0)

' After:
GetPrevBkEnabled = True   ' OnPrevButtonClick guards: If m_currentBookIndex <= 1 GoTo PROC_EXIT
GetNextBkEnabled = True   ' OnNextButtonClick guards: If m_currentBookIndex <= 0 or >= 66 GoTo PROC_EXIT

' GetPrevChapterEnabled / GetNextChapterEnabled — before:
GetPrevChapterEnabled = (m_currentChapter > 0)
GetNextChapterEnabled = (m_currentChapter > 0 And m_currentBookIndex > 0)

' After:
GetPrevChapterEnabled = True   ' OnPrevChapterClick guards: If m_currentChapter > 1
GetNextChapterEnabled = True   ' OnNextChapterClick guards: If m_currentChapter = 0 or m_currentBookIndex = 0

' GetPrevVerseEnabled / GetNextVerseEnabled — before:
GetPrevVerseEnabled = (m_currentChapter > 0)
GetNextVerseEnabled = (m_currentChapter > 0)

' After:
GetPrevVerseEnabled = True   ' OnPrevVerseClick guards: If m_currentVerse > 1
GetNextVerseEnabled = True   ' OnNextVerseClick guards: If m_currentVerse = 0 or m_currentChapter = 0 or m_currentBookIndex = 0
```

**Resulting Tab order from initial render (consistent on all loads):**
```
PrevBookButton → cmbBook → NextBookButton → PrevChapterButton → cmbChapter → NextChapterButton
→ PrevVerseButton → cmbVerse → NextVerseButton → NewSearchButton → adaeptButton
```

`NewSearchButton` is not always-enabled; it stays disabled until a book is selected. Tab skips
it on first load, which is correct — New Search has nothing to reset.

With this fix the sequence `gen [Tab][Tab][Tab] 119 [Tab]` follows the same path on first load
as on second use: Tab 1 → `NextBookButton`, Tab 2 → `PrevChapterButton`, Tab 3 → `cmbChapter`.

**Status: fix applied — 2026-04-18. Pending test: verify Tab path is consistent on first load and second use.**

### Status

| Item | Status |
|------|--------|
| Bug #599 — first-load Tab goes to document | **FIX APPLIED — 2026-04-18 — pending test** |

---

## § 6 — Bug #597: New Search Should Focus cmbBook

### Symptom

After clicking the New Search button, focus returns to the document body instead of `cmbBook`.

### Root cause

Office ribbon `onAction` callbacks always return focus to the document after the button
activates — this is by design in the Win32/Office Fluent ribbon architecture. There is
no `IRibbonUI.FocusControl()` method. The current `OnNewSearchClick` resets state and
calls `m_ribbon.Invalidate` (synchronous, correct), but cannot redirect where focus goes.

### Options

| Option | Approach | Risk |
|--------|----------|------|
| A | Deferred `Application.SendKeys` with keytip sequence (Alt → tab keytip → "B") | Fragile — requires knowing auto-assigned tab keytip; locale-sensitive |
| B | Deferred `Application.SendKeys "{F6}"` to cycle focus to ribbon | Partial — lands somewhere in ribbon, not specifically cmbBook |
| C | Accept as known limitation — document the Tab/keytip workflow | No code change; user navigates manually |

The auto-assigned keytip for the "Radiant Word Bible" tab is not hard-coded in `customUI14.xml`
(no `keytip=` attribute on `<tab>`). The keytip Word assigns at runtime depends on which
single letters are still free after all built-in tabs. This value cannot be determined from
VBA code.

**Decision pending:** If Option A is chosen, the user must supply the auto-assigned tab
keytip character observed when pressing Alt with the Bible tab active.

### Status

| Item | Status |
|------|--------|
| Bug #597 — New Search should focus cmbBook | **OPEN — awaiting decision on approach; keytip value needed for Option A** |

---

## § 7 — Session Status Summary (2026-04-18)

| Item | Status |
|------|--------|
| Bug #597 — New Search focus to cmbBook | **OPEN** |
| Bug #598 — C/V Prev/Next disabled after book selection | **CONFIRMED — Buttons now Enabled — 2026-04-18** |
| Bug #599 — First-load Tab goes to document | **FIX APPLIED — 2026-04-18 — pending test** |
| Bug 16 — Keytip badges end-to-end test | **PENDING** |
| Bug 22 / 23a — First-nav layout delay | **KNOWN LIMITATION** |
| Bug 27 — Enter in Chapter does not navigate | **KNOWN LIMITATION** |
| Step 7 — OLD_CODE cleanup | **PENDING** |
| WarmLayoutCache rewrite | **FUTURE** |
| Search tracking reset | **FUTURE** |
