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

```txt
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
| Bug #599 — First-load Tab goes to document | **CONFIRMED — Tab order consistent on all loads — 2026-04-18** |
| Bug 16 — Keytip badges end-to-end test | **PENDING** |
| Bug 22 / 23a — First-nav layout delay | **KNOWN LIMITATION** |
| Bug 27 — Enter in Chapter does not navigate | **SUPERSEDED — see § 8 GoButton proposal** |
| Step 7 — OLD_CODE cleanup | **PENDING** |
| Nav error feedback — status bar | **PROPOSED — see § 8** |
| GoButton — explicit navigation trigger | **PROPOSED — see § 8** |
| WarmLayoutCache rewrite | **FUTURE** |
| Search tracking reset | **FUTURE** |

---

## § 8 — Proposal: GoButton (Explicit Navigation Trigger)

### Context

The always-enable Tab fix (§ 5, #599) confirmed that Tab order through the ribbon is
now consistent on first load and all subsequent uses. However, navigation has no
explicit trigger visible to the user. The current model fires navigation implicitly
when the user Tabs past `cmbVerse` — an invisible action-gate that is not obvious
from the UI alone.

Two related items motivate this proposal:

1. **Bug 27** (Enter in Chapter does not navigate) — `onChange` cannot distinguish
   the Enter key from a normal keystroke; no workaround exists within the ribbon API.
2. **Nav error feedback** — status bar messages were proposed for invalid comboBox
   entry. Those messages are more actionable if there is a clear recovery step ("press
   Enter to navigate after correcting the value").

### Proposal

Add a large ribbon button — **GoButton** — positioned between `NextVerseButton` and
`NewSearchButton`. The button represents the Enter key and serves as the explicit
navigation trigger.

**Proposed ribbon XML placement:**

```txt
PrevBookButton → cmbBook → NextBookButton
→ PrevChapterButton → cmbChapter → NextChapterButton
→ PrevVerseButton → cmbVerse → NextVerseButton
→ [GoButton]          ← NEW
→ NewSearchButton → adaeptButton
```

**Proposed label / keytip:**

- Label: `Go` (or `↵` if font renders reliably in ribbon)
- Keytip: `G`
- Size: `large` (ribbon XML `size="large"`)

### Pros

| # | Pro |
|---|-----|
| 1 | **Explicit action** — user knows exactly when navigation fires; no hidden Tab-trigger |
| 2 | **Absorbs Bug 27** — Enter-in-Chapter is no longer needed; GoButton is the intended trigger |
| 3 | **Status bar feedback gains a recovery path** — "Invalid entry — correct and press Go (G)" is actionable |
| 4 | **Simplifies navigation trigger logic** — `OnGoClick` becomes the single fire point; deferred verse execution path in `OnVerseChanged` can be removed or reduced |
| 5 | **Matches keyboard mental model** — Tab through B/C/V, then Tab to Go, press Space (or click) |
| 6 | **Clarifies NewSearchButton role** — New Search resets only; Go navigates; two distinct actions, not combined |
| 7 | **Large size signals importance** — visually anchors the "do it" step in the control group |
| 8 | **May resolve Bug #597** — `OnGoClick` fires via `onAction`; its deferred Invalidate is synchronous post-action; a `FocusBookDeferred` sub in `basRibbonDeferred` could be scheduled from `OnNewSearchClick` now that the action-completion model is explicit |

### Cons

| # | Con |
|---|-----|
| 1 | **Ribbon real estate** — a large button takes more horizontal space than the current layout |
| 2 | **Paradigm shift** — users who learned the Tab-auto-navigate flow must adapt; Tab past Verse would no longer fire navigation |
| 3 | **Double-trigger risk** — if the implicit Tab-past-Verse trigger is kept alongside GoButton, both paths must be kept consistent; easier to remove the implicit path entirely |
| 4 | **XML re-import required** — `customUI14.xml` must be edited and the document re-imported to add the control; this is a deployment step, not just a VBA change |
| 5 | **GetGoEnabled callback needed** — button should be disabled until a book is selected (same guard as `NewSearchButton`), or always-enabled with a click guard (same pattern as §5) |

### Benefit analysis

| Area | Impact |
|------|--------|
| Bug 27 | Closed — Enter-in-Chapter becomes irrelevant |
| Bug #597 | Likely resolved — `OnNewSearchClick` can schedule `FocusBookDeferred` from `basRibbonDeferred`; GoButton makes the action boundary clear, removing ambiguity about when focus should return |
| Nav error UX | Elevated — status bar message + keytip `G` gives a complete "see error → fix → press G" loop |
| Code complexity | Reduced — single navigation entry point replaces deferred verse-trigger path |
| Tab order | Unchanged from §5 fix; GoButton inserts as a natural endpoint before New Search |
| Documentation | Simplified — "fill B/C/V, press Go" is a one-sentence instruction |

### Decisions — resolved 2026-04-18

| Question | Decision |
|----------|----------|
| Remove implicit Tab-past-Verse trigger? | **Yes** — clean single entry point |
| GoButton always-enabled or state-gated? | **Disabled until book selected** — consistent with `GetNewSearchEnabled` |
| Icon | **`mso EndOfDocument` or `mso PageNextWord`** — verify visually; both candidates for Enter/Go semantics |
| Keytip | **`G`** |

**Status: APPROVED — implementation plan in § 9.**

---

## § 9 — GoButton: Implementation Plan

### Scope

Three files change: `customUI14.xml` (ribbon XML), `aeRibbonClass.cls` (callbacks + trigger removal),
`basRibbonDeferred.bas` (GoToVerseDeferred stub).

---

### Step 1 — customUI14.xml: insert GoButton

Insert between `NextVerseButton` and `NewSearchButton`. Exact element:

```xml
<button id="GoButton"
        label="Go"
        size="large"
        keytip="G"
        imageMso="EndOfDocument"
        getEnabled="GetGoEnabled"
        onAction="OnGoClick" />
```

`imageMso="PageNextWord"` is the alternative candidate. Verify both visually in the ribbon
before committing. Re-import `customUI14.xml` into the `.docm` after edit.

**Resulting Tab order:**
```
PrevBookButton → cmbBook → NextBookButton
→ PrevChapterButton → cmbChapter → NextChapterButton
→ PrevVerseButton → cmbVerse → NextVerseButton
→ GoButton → NewSearchButton → adaeptButton
```

---

### Step 2 — aeRibbonClass.cls: add GetGoEnabled and OnGoClick

**GetGoEnabled** — same guard pattern as `GetNewSearchEnabled` (line 939):

```vba
' -- Go (navigate) -------------------------------------------------------------

Public Function GetGoEnabled(control As IRibbonControl) As Boolean
    GetGoEnabled = (m_currentBookIndex <> 0)   ' disabled until book selected
End Function

Public Sub OnGoClick(control As IRibbonControl)
    On Error GoTo PROC_ERR
    If m_currentBookIndex = 0 Then GoTo PROC_EXIT        ' guard: no book
    If m_currentChapter = 0 Then GoTo PROC_EXIT          ' guard: no chapter
    Dim vsNum As Long
    vsNum = m_currentVerse
    If vsNum < 1 Then vsNum = 1                          ' default verse 1
    GoToVerse vsNum
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OnGoClick of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub
```

---

### Step 3 — aeRibbonClass.cls: remove implicit navigation trigger from OnVerseChanged

Current `OnVerseChanged` (line 771): validates verse, sets `m_pendingVerse`, then schedules
`GoToVerseDeferred` via `Application.OnTime`.

**After change**, `OnVerseChanged` validates and stores state only — does not schedule navigation:

```vba
Public Sub OnVerseChanged(control As IRibbonControl, text As String)
    ' onChange fires on Enter and on each keystroke.
    ' Validates and stores m_currentVerse. Navigation fires from OnGoClick only.
    On Error GoTo PROC_ERR
    Dim projNameVs As String
    projNameVs = Application.ActiveDocument.VBProject.Name
    If Not IsNumeric(Trim(text)) Then
        Application.OnTime Now, projNameVs & ".basRibbonDeferred.ResetVerseDisplayDeferred"
        GoTo PROC_EXIT
    End If
    If m_currentChapter = 0 Or m_currentBookIndex = 0 Then GoTo PROC_EXIT
    Dim vsNum As Long
    vsNum = CLng(Trim(text))
    Dim bookName As String
    bookName = CStr(headingData(m_currentBookIndex, 0))
    If vsNum < 1 Or vsNum > aeBibleCitationClass.VersesInChapter(bookName, m_currentChapter) Then
        Application.OnTime Now, projNameVs & ".basRibbonDeferred.ResetVerseDisplayDeferred"
        GoTo PROC_EXIT
    End If
    m_currentVerse = vsNum   ' store confirmed value; GoButton fires navigation
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OnVerseChanged of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub
```

**Key changes from current:**

- `m_pendingVerse = vsNum` → `m_currentVerse = vsNum` (no pending state needed)
- Remove `Application.OnTime ... GoToVerseDeferred` call
- `ExecutePendingVerse` becomes dead code — to be removed or stubbed in Step 5

---

### Step 4 — aeRibbonClass.cls: review OnChapterChanged for parallel implicit trigger

`OnChapterChanged` (line ~629) schedules `GoToChapterDeferred` → `ExecutePendingChapter`.
With GoButton as the sole navigation trigger, `OnChapterChanged` should also validate and
store only — no deferred navigation.

`ExecutePendingChapter` currently invalidates verse-row controls. That Invalidate call
must move to `OnGoClick` (fire after navigation) or remain in a renamed deferred sub that
does NOT call navigation.

**Scope note:** this step must be confirmed against current `OnChapterChanged` and
`ExecutePendingChapter` code before applying.

---

### Step 5 — basRibbonDeferred.bas: stub GoToVerseDeferred

`GoToVerseDeferred` (line 47) calls `ExecutePendingVerse`. With `ExecutePendingVerse`
removed, this becomes a dead stub:

```vba
Public Sub GoToVerseDeferred()
    ' Dead stub — navigation trigger moved to OnGoClick (GoButton, #600).
    ' Instance().ExecutePendingVerse
End Sub
```

---

### Step 4 finding — OnChapterChanged

`OnChapterChanged` does NOT schedule `GoToChapterDeferred` — there was no implicit nav
trigger in the chapter path. `ExecutePendingChapter` (called from `GoToChapterDeferred`)
only invalidates verse controls; it performs no navigation. Both remain unchanged.

### Status

| Step | Item | Status |
|------|------|--------|
| 1 | XML — insert GoButton (`customUI14backupRWB.xml`) | **DONE — 2026-04-18** |
| 2 | VBA — GetGoEnabled / OnGoClick (`aeRibbonClass.cls`) | **DONE — 2026-04-18** |
| 3 | VBA — OnVerseChanged: remove nav trigger, `m_currentVerse` direct | **DONE — 2026-04-18** |
| 4 | VBA — OnChapterChanged: no implicit trigger confirmed — no change | **DONE — 2026-04-18** |
| 5 | VBA — GoToVerseDeferred: stubbed (`basRibbonDeferred.bas`) | **DONE — 2026-04-18** |

### Ribbon XML injection

#### Automated — py/inject_ribbon.py (preferred)

The script replaces `customUI/customUI14.xml` inside the `.docm` zip in place.
Run from a WSL terminal with the `.docm` closed in Word:

```bash
cd /mnt/c/adaept/aeBibleClass
python3 py/inject_ribbon.py 'Blank Bible Copy.docm'
```

Expected output:

```txt
REPLACED  customUI/customUI14.xml
Done.  Blank Bible Copy.docm updated.
```

**Applied 2026-04-18** — GoButton entry injected successfully.

#### Manual procedure (fallback — no WSL/Python available)

A `.docm` file is a zip archive. Any zip tool (7-Zip, Windows Explorer zip,
WinRAR) can open and edit it directly.

1. **Close `Blank Bible Copy.docm` in Word** — the file must not be open.
2. **Make a backup copy** — e.g., `Blank Bible Copy BACKUP.docm` — before editing.
3. **Open the `.docm` as a zip archive:**
   - In 7-Zip: right-click the file → *7-Zip → Open archive*
   - In Windows Explorer: rename `.docm` → `.zip`, then open the folder
4. **Navigate to `customUI/`** inside the archive.
5. **Replace `customUI14.xml`:**
   - Delete the existing `customUI14.xml` from the archive.
   - Drag `customUI14backupRWB.xml` from the project root into the `customUI/` folder.
   - Rename it to `customUI14.xml` inside the archive.
6. **Save and close the archive.**
   - If you renamed to `.zip` in step 3, rename back to `.docm`.
7. **Open `Blank Bible Copy.docm` in Word** — the updated ribbon loads automatically.

> Note: do not edit `customUI/_rels/customUI14.xml.rels` or `[Content_Types].xml`
> unless adding the `customUI` part for the first time. Replacing the existing
> `customUI14.xml` content requires no relationship or content-type changes.

---

## § 10 — Bugs: "Macro not found" / GoButton stays disabled after book selection

### Root cause

Both symptoms have the same cause. `inject_ribbon.py` updates **only** `customUI/customUI14.xml`
inside the `.docm` zip — the ribbon XML. It does not touch `vbaProject.bin`, which contains
the compiled VBA code.

The callbacks `GetGoEnabled` and `OnGoClick` were added to `src/aeRibbonClass.cls`, and
`GoToVerseDeferred` was stubbed in `src/basRibbonDeferred.bas`, by editing the `src/` files
directly. Those changes are not yet inside the running VBA project. Word therefore:

- Cannot resolve `GetGoEnabled` → button defaults to **disabled** (Bug: GoButton not enabled)
- Cannot resolve `OnGoClick` → **"Macro not found"** error on click

### Fix — import updated modules into the VBA project

There is no automated VBA import script in this project. The `src/` files are the
git-tracked source; `vbaProject.bin` is updated by importing through the VBA editor.

**Files to import:**

- `src/aeRibbonClass.cls` — contains `GetGoEnabled`, `OnGoClick`, and all other changes
- `src/basRibbonDeferred.bas` — contains stubbed `GoToVerseDeferred`

**Procedure (with `Blank Bible Copy.docm` open in Word):**

1. Press **Alt+F11** to open the VBA editor.
2. In the **Project Explorer** (Ctrl+R if not visible), expand the project for `Blank Bible Copy`.
3. **Remove the old module** — right-click `aeRibbonClass` → *Remove aeRibbonClass* → when
   prompted to export first, click **No** (the `src/` file is already up to date).
4. **Import the updated file** — *File → Import File* →
   navigate to `C:\adaept\aeBibleClass\src\aeRibbonClass.cls` → *Open*.
5. Repeat steps 3–4 for **`basRibbonDeferred`** using `src\basRibbonDeferred.bas`.
6. Press **Ctrl+S** (or save from Word) to save the `.docm` with the updated VBA project.
7. **Reload the ribbon** — close and reopen the document, or run `RibbonOnLoad` manually
   from the VBA editor (F5 with cursor in the sub) to reinitialise the ribbon instance.

### Workflow note

The current project workflow is **VBA-editor-first**:

```txt
Edit in VBA editor → Export to src/ → normalize_vba.py → git commit
```

When `src/` files are edited directly (as in this session), the reverse step is needed:

```txt
Edit src/ → Import into VBA editor → Save .docm → (re-export to confirm round-trip)
```

`py/normalize_vba.py` has been updated with normalizer entries for `GetGoEnabled` and
`OnGoClick` so that a subsequent export round-trip preserves their casing correctly.

### Status

| Item | Status |
|------|--------|
| "Macro not found" on GoButton click | **CLOSED — missing wrappers added to basBibleRibbonSetup.bas** |
| GoButton disabled after book selection | **CLOSED — same fix** |
| normalize_vba.py — GetGoEnabled / OnGoClick entries | **DONE — 2026-04-18** |
| Navigation confirmed working — all Prev/Next controls | **CONFIRMED — 2026-04-18** |

---

## § 11 — Status Bar: "Navigating ..." and SBL Citation Feedback

### Context

First-time navigation to a distant book (e.g., Revelation) carries a ~17-second layout
cost (Bug 22 / 23a — known limitation). With no visual feedback during this period the
ribbon and document appear frozen. A status bar message provides the minimum signal that
the operation is in progress.

### Fix applied — 2026-04-18

`Application.StatusBar = "Navigating ..."` set immediately before the expensive call in
both navigation paths, cleared with `Application.StatusBar = False` after completion:

| Method | Placement |
|--------|-----------|
| `GoToVerse` (`aeRibbonClass.cls`) | Before `GoToVerseByScan chPos, vsNum` |
| `GoToChapter` (`aeRibbonClass.cls`) | Before `ActiveDocument.Range(chPos, chPos).Select` |

Both paths restore the Word default status bar on completion. The comment references § 11
for the SBL citation decision.

**Files changed:** `src/aeRibbonClass.cls` — import required.

### Proposal — SBL short form citation in status bar after each successful navigation

After navigation completes, display the current reference in SBL short form instead of
restoring the Word default status bar. Example: `Gen 1:1`, `Ps 119:176`, `Rev 22:21`.

#### Pros

| # | Pro |
|---|-----|
| 1 | **Confirms navigation** — user sees the exact reference reached, not just that something happened |
| 2 | **Persists between actions** — unlike "Navigating ..." which clears, the citation remains until the next navigation or Word overwrites it |
| 3 | **Useful with Prev/Next buttons** — rapid chapter/verse stepping shows current position after each click |
| 4 | **Consistent with Bible software convention** — most Bible applications display current reference in a status area |
| 5 | **Recovers context after long wait** — after a 17-second load, the citation confirms the destination was reached |
| 6 | **SBL infrastructure already present** — `aeSBL_Citation_Class` and book name data in `headingData` are available |
| 7 | **Documentation value** — "status bar shows current reference" is a one-line user instruction |

#### Cons

| # | Con |
|---|-----|
| 1 | **Status bar is shared and ephemeral** — Word overwrites it on hover, selection, and many other events; citation may disappear unexpectedly |
| 2 | **SBL abbreviation table needed** — `headingData` stores full book names; generating `Gen` from `Genesis` requires an abbreviation lookup; this is a new data dependency |
| 3 | **Coupling** — navigation code gains a formatting dependency (citation string builder); violates single-responsibility if done inline |
| 4 | **Stale on error** — if navigation exits early (guard fails silently), the previous citation remains displayed and may mislead |
| 5 | **Not accessible** — screen readers do not announce status bar changes; users relying on accessibility get no benefit |
| 6 | **User documented to ignore it** — the status bar is noted as typically ignored in use (§ 8); upfront documentation is the mitigation |

#### Implementation options

| Option | Approach | Notes |
|--------|----------|-------|
| A | Full book name — `Genesis 1:1` | No new data; looks verbose for long names |
| B | SBL short abbreviation — `Gen 1:1` | Requires abbreviation lookup table; cleanest for Bible use |
| C | Ribbon comboBox values already show B/C/V | Status bar adds redundant info unless citation includes book abbreviation |

**Recommendation:** Option B, implemented as a private helper `SBLStatusText` in
`aeRibbonClass.cls` that reads from a `Const` array of 66 abbreviations (parallel to
`headingData` book order). Called at the end of `GoToVerse` and `GoToChapter` in place of
`Application.StatusBar = False`.

**Status: PROPOSED — awaiting approval.**

### Reusing GetBookAliasMap for SBL display — Pros/Cons

`aeBibleCitationClass.GetBookAliasMap` builds a `Scripting.Dictionary` mapping alias
strings (uppercase) to book index numbers. Each book group is added in this order:

```txt
Full name → SBL short form → shorter alternates
e.g.: "GENESIS", "GEN", "GE", "GN"  →  all map to index 1
```

The map contains the SBL abbreviations but does **not mark which entry is canonical**.
The SBL form is not always at a fixed insertion position:

| Book | Insertion order | SBL canonical | Position |
|------|----------------|---------------|----------|
| Genesis | GENESIS, **GEN**, GE, GN | GEN | 2nd |
| Exodus | EXODUS, **EXOD**, EXO, EX | EXOD | 2nd |
| Judges | JUDGES, JUDGE, **JUDG**, JGS | JUDG | 3rd |
| Ruth | **RUTH**, RUT, RU | Ruth (no abbrev) | 1st |
| Psalms | PSALMS, PSALM, PSA, **PS** | Ps | 4th |

A simple "take the Nth entry" rule is not reliable across all 66 books.

#### Pros

| # | Pro |
|---|-----|
| 1 | **No new file or module** — abbreviations live alongside alias definitions in `aeBibleCitationClass` |
| 2 | **Single source of truth** — any correction to a book alias propagates to both input parsing and display |
| 3 | **`aliasMap` is already lazily initialised and cached** — no overhead after first navigation |
| 4 | **SBL form is already in the map** — the canonical abbreviation is one of the recognised aliases; no new string data is needed, only a mechanism to identify which alias is canonical |
| 5 | **Public API already exists** — `GetBookAliasMap` is accessible from `aeRibbonClass` without new references |

#### Cons

| # | Con |
|---|-----|
| 1 | **Map direction is inverted** — `aliasMap` is alias→index; display needs index→canonical. Requires a reverse lookup structure |
| 2 | **Canonical form not marked** — SBL position varies by book (2nd, 3rd, 4th, or full name). A positional rule fails for Judges, Ruth, Psalms and others |
| 3 | **Runtime reverse-build cost** — iterating 264+ keys to build an index→abbrev map; acceptable as a one-time initialisation but adds complexity |
| 4 | **Coupling** — `aeRibbonClass` display logic depends on `aeBibleCitationClass` internals; a display concern mixes with a data concern |
| 5 | **Implicit contract** — the canonical SBL form must be added at a defined position in each book group; this convention is not enforced and can silently break if entries are reordered |

#### Proposed resolution — superseded by ToSBLShortForm discovery

`GetSBLAbbrev` + `sblMap` are **not needed**. `aeBibleCitationClass` already has:

```vba
Public Function ToSBLShortForm(ByVal canon As String) As String
```

`ToSBLShortForm` (line 2835) contains the complete 66-book `Select Case` abbreviation
table and already strips the chapter number for single-chapter books:

```vba
' SBL shorthand for single-chapter books omits the chapter number.
' Canonical form is "Jude 1:6"; SBL output is "Jude 6".
If GetMaxChapter(bID) = 1 Then
    cpPos = InStr(numPart, ":")
    If cpPos > 0 Then numPart = Mid$(numPart, cpPos + 1)
End If
```

Single-chapter books handled: Obadiah (31), Philemon (57), 2 John (63), 3 John (64),
Jude (65). `GetSingleChapterBookSet` (line 1240) documents the full set.

The status bar call in `GoToVerse` and `GoToChapter` becomes:

```vba
Application.StatusBar = aeBibleCitationClass.ToSBLShortForm( _
    CStr(headingData(m_currentBookIndex, 0)) & " " & m_currentChapter & ":" & m_currentVerse)
```

`headingData(m_currentBookIndex, 0)` supplies the full book name (e.g., `"Genesis"`),
which `ResolveAlias` inside `ToSBLShortForm` recognises. No new method, no `sblMap`,
no second table of any kind.

**Pending actions before implementation:**

- Revert `Private sblMap As Object` line added to `aeBibleCitationClass.cls`
- Add `Application.StatusBar = aeBibleCitationClass.ToSBLShortForm(...)` to `GoToVerse` and `GoToChapter`
- Add `ToSBLShortForm` wrapper to `basBibleRibbonSetup.bas` — not needed (called directly, not via ribbon XML)
- Add `ToSBLShortForm` to `normalize_vba.py` normaliser

**Status: DONE — 2026-04-18. Import `aeRibbonClass.cls` to activate.**

### Bug: Prev/Next clicks show Word's "Page X of Y" status — overwriting SBL citation

#### Cause

`onAction` callbacks always trigger Word to refresh its own status bar (page, section,
word count) **after** the callback returns. The `ToSBLShortForm` call inside `GoToVerse`
and `GoToChapter` fires before that refresh, so Word overwrites the citation.

#### Fix — deferred status bar write via Application.OnTime

Same pattern as `GoToVerseDeferred` / `ExecutePendingChapter`. Schedule the status bar
write via `Application.OnTime Now` so it fires after Word's own refresh:

1. `onAction` fires → nav runs → **schedules `UpdateStatusBarDeferred`**
2. `onAction` returns → Word writes `"Page X of Y, Words: XXXX"`
3. `OnTime` fires → `UpdateStatusBarDeferred` → SBL citation overwrites

#### Known limitation — flash still occurs

The Word status bar refresh (step 2) and the deferred write (step 3) happen in rapid
succession — typically tens of milliseconds — but the flash is **not eliminated**, only
shortened. There is no Word VBA API to suppress the post-`onAction` status bar refresh
for individual controls. Hiding the status bar entirely (`Application.DisplayStatusBar`)
would be more disruptive than the flash itself.

Accepted as expected behaviour. For the slow first-load path the flash is negligible;
for fast repeated Prev/Next clicks it may be briefly visible.

#### Implementation

| File | Change |
|------|--------|
| `aeRibbonClass.cls` — `GoToVerse` | Replace direct `StatusBar` write with `OnTime` schedule |
| `aeRibbonClass.cls` — `GoToChapter` | Same |
| `aeRibbonClass.cls` | Add `Public Sub UpdateStatusBar` |
| `basRibbonDeferred.bas` | Add `Public Sub UpdateStatusBarDeferred` |
| `normalize_vba.py` | Add `UpdateStatusBar` / `UpdateStatusBarDeferred` entries |

**Status: DONE — 2026-04-18.**

### Bug: Prev/Next Book does not update the status bar

#### Root cause

`PrevButton` and `NextButton` in `aeRibbonClass.cls` correctly update `m_currentBookIndex`,
`m_currentChapter = 1`, `m_currentVerse = 1` and call `m_ribbon.Invalidate`, but did not
schedule `UpdateStatusBarDeferred`. The deferred write was only wired into `GoToVerse` and
`GoToChapter` — not the book navigation path.

Note: `GoToH1Deferred` is the old InputBox-based book lookup (not used by Prev/Next Book).
Prev/Next Book call `PrevButton` / `NextButton` directly via `OnPrevButtonClick` /
`OnNextButtonClick`.

#### Fix applied — 2026-04-18

Added `Application.OnTime Now, ... & ".basRibbonDeferred.UpdateStatusBarDeferred"` at the
end of both `PrevButton` and `NextButton`, after `m_ribbon.Invalidate`. State is already
correct at that point — `m_currentChapter = 1` and `m_currentVerse = 1` (Rule 2a), so
`UpdateStatusBar` will display e.g., `"Exod 1:1"` after pressing Next from Genesis.

Import `aeRibbonClass.cls` to activate.

| Item | Status |
|------|--------|
| Prev/Next Book status bar update | **DONE — 2026-04-18** |
| Prev/Next Chapter status bar update | **DONE — 2026-04-18** |
| Prev/Next Verse status bar update | **DONE — 2026-04-18** |
| GoButton (OnGoClick) status bar update | **DONE — 2026-04-18** |

---

### Bug: "Navigating ..." disappears when REVELATION appears; spinning continues ~10s

#### Observed behaviour

| Time | Event |
|------|-------|
| 0s | GoButton pressed; `"Navigating ..."` set in status bar |
| ~7s | Document scrolls to show REVELATION heading; status bar message **gone** |
| ~7–17s | Word busy cursor continues spinning; status bar empty |

#### Explanation

`GoToVerseByScan` scans paragraph by paragraph for the Nth Verse marker style run.
When the target is found it calls `Range.Select`, which causes two things:

1. Word scrolls the document view to the selected position — **REVELATION becomes visible**.
2. `Range.Select` **returns control to VBA immediately**, before Word's layout engine
   has finished paginating the document from that position.

Our code had `Application.StatusBar = False` immediately after `GoToVerseByScan`
returned. Because `Range.Select` returns to VBA before layout completes, the status bar
was cleared at the wrong moment — exactly when REVELATION appeared, with ~10 seconds
of Word background rendering still pending.

**VBA has no event for "Word layout engine finished."** There is no
`Application.OnLayoutComplete` or equivalent callback. The ~10-second spinning cursor
after the view scrolls is entirely inside Word's rendering subsystem, outside VBA's
call stack.

#### Fix applied — 2026-04-18

Removed `Application.StatusBar = False` from `GoToVerse` and `GoToChapter`.
`"Navigating ..."` now persists until the next navigation overwrites it with a new
message. A comment in the code records the reason.

When the SBL citation feedback (§ 11 proposal) is implemented, the citation string
will replace `"Navigating ..."` on completion — providing a positive confirmation
that the layout pass is done, because the citation is only written after
`GoToVerseByScan` returns and all state has been updated.

| Item | Status |
|------|--------|
| "Navigating ..." clears prematurely | **FIXED — 2026-04-18 — message persists through background layout** |
| GoToChapter — same fix applied | **FIXED — 2026-04-18** |

---

### Bug: Book comboBox displays "rev" after returning from Focus mode

#### Observed behaviour

1. User types `rev` in the Book comboBox and presses Go — document navigates to REVELATION.
2. `OnGoClick` calls `Instance().InvalidateControl "BookComboBox"` (via `m_ribbon.Invalidate`); `GetBookText` callback returns `"REVELATION"`.
3. The comboBox display updates to `"REVELATION"` — correct.
4. User clicks **Focus** in the status bar (or View → Focus); Focus mode activates.
5. User presses **Escape** (or closes Focus mode) — ribbon re-renders.
6. The Book comboBox now shows **`"rev"`** — the user's original typed text, not `"REVELATION"`.

#### Explanation — Win32 edit buffer vs getText callback

An Office ribbon `<comboBox>` control has two separate text values:

| Layer | What it holds | Who controls it |
|-------|--------------|-----------------|
| **Win32 edit buffer** | The raw string last typed or programmatically set by the host | Win32 control state; persists until the control is destroyed and recreated |
| **getText-derived display** | The string returned by the `getText` callback during `Invalidate` | VBA callback; applied only during an Invalidate pass |

When the user types `"rev"` and presses Tab or Go, the Win32 edit buffer stores `"rev"`.
`Invalidate` fires the `getText` callback which returns `"REVELATION"`, and the **display**
updates — but the Win32 buffer still contains `"rev"`.

**Focus mode destroys and recreates the ribbon** (it is a full application-level layout
switch, not a simple redraw). When the ribbon is rebuilt, the comboBox Win32 control is
created fresh. Word/Office initialises it from the **native control state**, which is the
Win32 edit buffer value — `"rev"` — not the `getText` callback. The `getText` callback is
not fired at control creation; it fires only during a subsequent `Invalidate` pass.

#### Why there is no VBA fix

- **No event for Focus mode entry/exit.** There is no `Application.OnFocusModeEnter`,
  `DocumentBeforeFocusMode`, or equivalent callback. VBA cannot detect the transition.
- **Invalidate cannot be scheduled.** Without an event trigger, `Application.OnTime`
  cannot be scheduled at the right moment. An `OnTime`-based polling workaround would
  fire during normal operation and cause unnecessary ribbon flicker.
- **`getText` callback is not called at control creation.** The callback fires during
  `Invalidate` only. The fresh comboBox is initialised from internal Office/Win32 state,
  bypassing the callback entirely.

#### Status: Known limitation — no fix available within current architecture

The behaviour is a consequence of how Office ribbon controls interact with the Win32
subsystem during host-level layout switches. It affects all Office applications that use
comboBox ribbon controls with user-typed input.

**Mitigation (documented, not implemented):** Users who notice the stale display can
press Tab or click the comboBox and type a new value; the next `Invalidate` pass
(triggered by any ribbon interaction) will restore the callback-derived display.

| Item | Status |
|------|--------|
| Book comboBox shows stale user text after Focus mode | **KNOWN LIMITATION — no VBA fix available** |

---

### Bug: Invalid comboBox input — status bar and display remain stale

#### Observed behaviour

1. A successful navigation has been made; status bar shows e.g., `"Gen 3:5"`.
2. User types `"sfdgs55jsfdr"` (or any unrecognised string) in the Book comboBox and clicks Go.
3. Nothing navigates — the guard inside `OnGoClick` / `OnBookChanged` rejects the input.
4. Status bar still shows `"Gen 3:5"` — the previous successful citation.
5. ComboBox still displays the nonsense string.
6. Prev/Next Book/Chapter/Verse navigate based on the last **valid** internal state
   (`m_currentBookIndex`, `m_currentChapter`, `m_currentVerse`) while the comboBox
   display continues to show the invalid string.

The same class of error applies to out-of-bounds numeric input:

| Input location | Example invalid value | Expected message |
|----------------|-----------------------|-----------------|
| Book comboBox | `"sfdgs55jsfdr"` (unrecognised alias) | `"Invalid input for Book — enter a book name or abbreviation"` |
| Chapter comboBox | `"99"` when book has 28 chapters | `"Invalid input for Chapter — out of range (1–28)"` |
| Verse comboBox | `"0"` or `"999"` | `"Invalid input for Verse — out of range (1–N)"` |

#### Root cause

`OnBookChanged`, `OnChapterChanged`, and `OnVerseChanged` validate input and silently
return on failure. No feedback is written to the status bar and the comboBox display is
not corrected. The status bar retains whatever was last written by a successful
navigation.

#### Proposed fix — status bar error feedback

Write a descriptive error string to `Application.StatusBar` on any validation failure
inside `OnGoClick` and the `OnXxxChanged` handlers.

Three error categories:

| Category | Trigger | Proposed status bar text |
|----------|---------|--------------------------|
| **Book unrecognised** | `ResolveAlias` returns 0 or fails | `"Invalid input for Book — enter a book name or abbreviation"` |
| **Chapter out of range** | Parsed integer < 1 or > `GetMaxChapter(bookID)` | `"Invalid input for Chapter — out of range (1–N)"` |
| **Verse out of range** | Parsed integer < 1 or > `GetMaxVerse(bookID, chapter)` | `"Invalid input for Verse — out of range (1–N)"` |

Non-numeric chapter/verse input (e.g., letters in the chapter field) is treated as
out-of-range after `Val()` or `CInt()` parsing returns 0.

The comboBox display issue (nonsense text persisting) is a separate concern: calling
`m_ribbon.InvalidateControl "BookComboBox"` after a failed parse will trigger `GetBookText`
which returns the last valid book name, correcting the display. Same applies to
`ChapterComboBox` and `VerseComboBox`.

#### Pros

| # | Pro |
|---|-----|
| 1 | **Immediate user feedback** — user knows why nothing happened without having to guess |
| 2 | **Status bar is already the feedback channel** — consistent with `"Navigating ..."` and SBL citation patterns established this session |
| 3 | **ComboBox display corrected** — `InvalidateControl` after rejection resets the display to the last valid value; user sees the active state, not stale nonsense |
| 4 | **Prev/Next remain correct** — navigation continues from the last valid state; the error message clarifies why the comboBox shows something different |
| 5 | **Low coupling** — error writes happen at the guard site (`OnGoClick`, `OnBookChanged`, `OnChapterChanged`, `OnVerseChanged`); no new methods or modules required |
| 6 | **Range limit in message** — `"out of range (1–28)"` teaches the user the valid range without a dialog box |
| 7 | **Deferred pattern not needed** — error messages do not compete with Word's post-`onAction` status bar refresh; they persist until the next navigation or `OnTime` write |

#### Cons

| # | Con |
|---|-----|
| 1 | **Status bar is ephemeral** — Word will overwrite the error message on hover, selection change, or next Prev/Next click; user may miss it |
| 2 | **Range data needed at call site** — displaying `"out of range (1–28)"` requires `GetMaxChapter` / `GetMaxVerse` to be called even when navigation is not proceeding; minor overhead |
| 3 | **`InvalidateControl` adds a ribbon round-trip per rejection** — visible as a brief flicker if the user types rapidly; acceptable given the use case (deliberate Go click) |
| 4 | **Error message is overwritten by next Prev/Next** — if the user immediately clicks Prev/Next after an invalid Book entry, the error disappears and is replaced by the SBL citation of the nav result; sequence may confuse |
| 5 | **Chapter/Verse validation in `OnXxxChanged` fires on every keystroke** — attaching error feedback there would produce per-keystroke error messages; error feedback should be deferred to `OnGoClick` only |

#### Benefits

- Eliminates the silent-failure UX: the application appears unresponsive when invalid input is entered. A status bar message makes the rejection explicit.
- `InvalidateControl` after rejection restores comboBox display consistency with internal state, removing the misleading Prev/Next behaviour (navigates correctly but display shows nonsense).
- Establishes a consistent error feedback pattern for all three comboBox controls, reducing future ambiguity if new validation rules are added.

#### Prev/Next boundary — gap identified

The six Prev/Next buttons were changed to always return `True` from `GetXxxEnabled`
(§ 5 — always-enable at boundary pattern). The click handlers (`PrevButton`,
`NextButton`, `OnPrevChapterClick`, `OnNextChapterClick`, `OnPrevVerseClick`,
`OnNextVerseClick`) guard boundaries silently: they detect the limit and return without
navigating, but write **nothing** to the status bar.

This means:

- Clicking PrevBook at Genesis produces no feedback — the button appears to do nothing.
- Clicking NextVerse at the last verse of the chapter produces no feedback — same.
- The same silence applies to all six Prev/Next boundary cases.

Boundary status bar messages follow the same pattern as invalid input messages:

| Button | Boundary condition | Proposed status bar text |
|--------|--------------------|--------------------------|
| PrevBook | Already at first book (index 1) | `"Already at first book"` |
| NextBook | Already at last book (index 66) | `"Already at last book"` |
| PrevChapter | Already at chapter 1 | `"Already at first chapter of [Book]"` |
| NextChapter | Already at last chapter of current book | `"Already at last chapter of [Book]"` |
| PrevVerse | Already at verse 1 | `"Already at first verse of [Book] [C]"` |
| NextVerse | Already at last verse of current chapter | `"Already at last verse of [Book] [C]"` |

These are written directly (not deferred) — boundary guards return early before
any navigation, so there is no post-`onAction` Word status bar refresh to race against.

#### Scope

| File | Change |
|------|--------|
| `aeRibbonClass.cls` — `OnGoClick` | After book alias fails: write error to status bar; call `InvalidateControl "BookComboBox"` |
| `aeRibbonClass.cls` — `OnGoClick` | After chapter out-of-range: write error; call `InvalidateControl "ChapterComboBox"` |
| `aeRibbonClass.cls` — `OnGoClick` | After verse out-of-range: write error; call `InvalidateControl "VerseComboBox"` |
| `aeRibbonClass.cls` — `OnBookChanged` | On alias failure: write error to status bar; call `InvalidateControl "BookComboBox"` |
| `aeRibbonClass.cls` — `OnChapterChanged` | Validation only — no error message here (see Con 5); error reported at Go time |
| `aeRibbonClass.cls` — `OnVerseChanged` | Same — validation only; no per-keystroke error message |
| `aeRibbonClass.cls` — `PrevButton` | At boundary: write `"Already at first book"` to status bar |
| `aeRibbonClass.cls` — `NextButton` | At boundary: write `"Already at last book"` to status bar |
| `aeRibbonClass.cls` — `OnPrevChapterClick` | At boundary: write `"Already at first chapter of [Book]"` |
| `aeRibbonClass.cls` — `OnNextChapterClick` | At boundary: write `"Already at last chapter of [Book]"` |
| `aeRibbonClass.cls` — `OnPrevVerseClick` | At boundary: write `"Already at first verse of [Book] [C]"` |
| `aeRibbonClass.cls` — `OnNextVerseClick` | At boundary: write `"Already at last verse of [Book] [C]"` |

**Status: DONE — 2026-04-19.**

#### Implementation notes

- `m_bookTextValid As Boolean` added to private state (initialized `True` in `Class_Initialize`).
- `OnBookChanged`: sets `m_bookTextValid = False` and writes status bar error on no-match;
  sets `m_bookTextValid = True` on successful alias resolution. No `InvalidateControl`
  here — calling it mid-typing would discard characters the user is still entering.
- `OnGoClick`: checks `Not m_bookTextValid` after the `m_currentBookIndex = 0` guard;
  writes the same error message and calls `InvalidateControl "BookComboBox"` to restore
  display to the last resolved book name before returning.
- `PrevButton` / `NextButton`: boundary split into `<= 0` (silent, no book set) and
  `= 1` / `>= 66` (error message written).
- Chapter/Verse boundary messages include the full book name from `headingData` and,
  for verse boundaries, the current chapter number — e.g., `"Already at last verse of Genesis 3"`.

---

### Status bar strings — i18n extraction

#### Context

The status bar messages added in the invalid-input and boundary fix are inline string
literals in `aeRibbonClass.cls`. `basUIStrings.bas` already exists as the declared
home for all user-facing ribbon strings. Its header states:

> *i18n: to localise the ribbon, edit only this module.*
> *VSTO port: replace this module with a .resx resource file. Constant names map directly to resource keys.*

The status bar messages are user-facing text that should live there — but they raise
a structural question: some messages include runtime data (book name, chapter number)
that cannot be expressed as a compile-time `Const`.

#### Is a separate module needed?

No. The case for a new module (e.g., `basStatusStrings`) rests on scope separation — 
"ribbon XML keytips" vs "status bar messages" — but `basUIStrings` already declares
itself as the home for **all** ribbon UI strings. Adding a second module for status bar
messages would split the i18n surface across two files, requiring translators and VSTO
porters to edit two locations. One module is sufficient.

If the module name feels too narrow after the additions, a rename to `basRibbonResources`
or `basUIStrings` is the appropriate response — not a new module.

#### Static vs dynamic messages

Messages fall into two groups:

| Group | Example | VBA representation |
|-------|---------|-------------------|
| **Static** | `"Invalid input for Book — enter a book name or abbreviation"` | `Public Const` — no runtime data |
| **Dynamic** | `"Already at last verse of Genesis 3"` | Contains book name + chapter; cannot be a `Const` |

Three implementation options for dynamic messages:

**Option A — Const prefix, caller concatenates**

```vba
' basUIStrings:
Public Const SB_ALREADY_FIRST_CHAPTER As String = "Already at first chapter of "

' aeRibbonClass:
Application.StatusBar = SB_ALREADY_FIRST_CHAPTER & bookName
```

- Translator sees only the prefix; the full sentence structure is split across two files.
- Simple — no helper function needed.

**Option B — Format-string Const, shared FormatMsg helper**

```vba
' basUIStrings:
Public Const SB_ALREADY_FIRST_CHAPTER As String = "Already at first chapter of {0}"

' basUIStrings (or basUtility):
Public Function FormatMsg(ByVal template As String, ParamArray args() As Variant) As String
    Dim result As String
    result = template
    Dim i As Long
    For i = 0 To UBound(args)
        result = Replace(result, "{" & i & "}", CStr(args(i)))
    Next i
    FormatMsg = result
End Function

' aeRibbonClass:
Application.StatusBar = FormatMsg(SB_ALREADY_FIRST_CHAPTER, bookName)
```

- Translator sees the complete sentence template including the placeholder position.
- Most i18n-correct: placeholder position can move between languages (e.g., German reverses noun order).
- Adds one utility function.

**Option C — Public Function per message**

```vba
' basUIStrings:
Public Function SB_AlreadyFirstChapter(ByVal bookName As String) As String
    SB_AlreadyFirstChapter = "Already at first chapter of " & bookName
End Function
```

- Translator must read function bodies rather than constant values — harder to extract.
- Mixing functions and constants in a "strings" module is unconventional and complicates VSTO port (no direct .resx mapping for functions).

#### Pros and Cons by option

| | Option A (prefix Const) | Option B (format Const + FormatMsg) | Option C (function) |
|-|------------------------|-------------------------------------|---------------------|
| Translator experience | Sees only the prefix | Sees full template | Must read function body |
| Localisation correctness | Word order fixed | Word order flexible | Word order flexible |
| VSTO port (.resx) | Direct key mapping | Direct key mapping | No direct mapping |
| Code at call site | `Const & runtime` | `FormatMsg(Const, arg)` | `Function(arg)` |
| New infrastructure | None | One helper function | None |
| VBA convention | Standard | Non-standard but clear | Unusual in a "strings" module |

#### Recommendation

**Option B** for all messages that embed runtime data. Rationale:

- The existing `basUIStrings` header explicitly targets VSTO port and localisation; Option B is the only option that keeps all translatable text in constant strings while remaining word-order-safe.
- `FormatMsg` is a four-line utility that lives in `basUIStrings` itself (it is cohesive with string resources).
- Call sites in `aeRibbonClass.cls` remain readable: `FormatMsg(SB_ALREADY_FIRST_CHAPTER, bookName)`.

Static messages (no runtime data) use `Public Const` directly — same as keytips today.

#### Proposed constant names

| Constant / Function | Value / Template |
|--------------------|-----------------|
| `SB_INVALID_BOOK` | `"Invalid input for Book — enter a book name or abbreviation"` |
| `SB_ALREADY_FIRST_BOOK` | `"Already at first book"` |
| `SB_ALREADY_LAST_BOOK` | `"Already at last book"` |
| `SB_ALREADY_FIRST_CHAPTER` | `"Already at first chapter of {0}"` |
| `SB_ALREADY_LAST_CHAPTER` | `"Already at last chapter of {0}"` |
| `SB_ALREADY_FIRST_VERSE` | `"Already at first verse of {0} {1}"` |
| `SB_ALREADY_LAST_VERSE` | `"Already at last verse of {0} {1}"` |

`{0}` = book name, `{1}` = chapter number where applicable.

**Status: DONE — 2026-04-19.**

| File | Change |
|------|--------|
| `basUIStrings.bas` | Added 7 `SB_*` constants and `FormatMsg` helper |
| `aeRibbonClass.cls` | All 8 inline status bar strings replaced with `SB_*` constants and `FormatMsg` calls |

---

### Gap: Pro #6 — range limit not included in messages

#### Problem

Pro #6 of the approved design stated:
> *"Range limit in message — `'out of range (1–28)'` teaches the user the valid range without a dialog box"*

Two categories of range information were never implemented:

**1. Out-of-range typed chapter/verse at Go time**

`OnGoClick` does not detect that the chapter or verse field contained an out-of-range
number. Reason: `m_currentChapter` and `m_currentVerse` are always valid at Go time —
`OnChapterChanged` and `OnVerseChanged` reset them to the previous valid value via
`ResetChapterDisplayDeferred` / `ResetVerseDisplayDeferred` when out-of-range input is
detected. `OnGoClick` never saw invalid state and navigated silently.

To close this, the same `m_bookTextValid` pattern is needed for Chapter and Verse:
- `m_chapterTextValid As Boolean` + `m_chapterMax As Long`
- `m_verseTextValid As Boolean` + `m_verseMax As Long`

Set `False` (and store max) in `OnChapterChanged` / `OnVerseChanged` on rejection;
set `True` on acceptance. `OnGoClick` reads the flags and shows:

| Constant | Template |
|----------|---------|
| `SB_INVALID_CHAPTER` | `"Invalid input for Chapter - out of range (1-{0})"` |
| `SB_INVALID_VERSE` | `"Invalid input for Verse - out of range (1-{0})"` |

Reset both flags to `True` in `OnBookChanged` (new book resets chapter/verse to valid
defaults 1/1) and in `OnNewSearchClick`.

**2. Range missing from boundary messages**

All four chapter/verse boundary constants omit the valid range:

| Current | Required |
|---------|---------|
| `"Already at first chapter of {0}"` | `"Already at first chapter of {0} (1-{1})"` |
| `"Already at last chapter of {0}"` | `"Already at last chapter of {0} (1-{1})"` |
| `"Already at first verse of {0} {1}"` | `"Already at first verse of {0} {1} (1-{2})"` |
| `"Already at last verse of {0} {1}"` | `"Already at last verse of {0} {1} (1-{2})"` |

`{1}` = max chapters (chapter messages); `{2}` = max verses (verse messages).
`OnNextChapterClick` and `OnNextVerseClick` already have the max value in scope
(used in the boundary condition); `OnPrevChapterClick` and `OnPrevVerseClick` need
one additional lookup.

#### Full scope — 2026-04-19 fix

| File | Change |
|------|--------|
| `basUIStrings.bas` | Add `SB_INVALID_CHAPTER`, `SB_INVALID_VERSE`; update 4 boundary constants with range placeholder |
| `aeRibbonClass.cls` — private state | Add `m_chapterTextValid`, `m_chapterMax`, `m_verseTextValid`, `m_verseMax` |
| `aeRibbonClass.cls` — `Class_Initialize` | Initialise new fields |
| `aeRibbonClass.cls` — `OnChapterChanged` | Set flag + max on rejection; set `True` on acceptance |
| `aeRibbonClass.cls` — `OnVerseChanged` | Same |
| `aeRibbonClass.cls` — `OnBookChanged` | Reset chapter/verse flags to `True` on successful match |
| `aeRibbonClass.cls` — `OnNewSearchClick` | Reset all four new fields |
| `aeRibbonClass.cls` — `OnGoClick` | Add chapter and verse flag checks with range messages |
| `aeRibbonClass.cls` — `OnPrevChapterClick` | Add max lookup; update `FormatMsg` call |
| `aeRibbonClass.cls` — `OnNextChapterClick` | Store max in variable; update `FormatMsg` call |
| `aeRibbonClass.cls` — `OnPrevVerseClick` | Add max lookup; update `FormatMsg` call |
| `aeRibbonClass.cls` — `OnNextVerseClick` | Store max in variable; update `FormatMsg` call |

**Status: DONE — 2026-04-19.**
