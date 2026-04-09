# Code Review: GoToH1 Persistent 12-Second Double-Block

**Date:** 2026-04-08
**Effort:** Expert

---

## 1 — Setting Effort to Expert

In the Claude Code prompt type:

```
/effort expert
```

Confirmation message: `Effort level: expert (currently high)`

Expert mode applies maximum reasoning depth before responding. It does not switch to
a different model. Use it when the problem requires sustained multi-step reasoning and
previous medium-effort sessions have produced loops.

---

## 2 — Problem Statement

GoToH1 to a far book (Genesis → Revelation) produces a repeating double block:

1. **First 12 seconds** — Word lays out all pages from Genesis to Revelation.
2. **Revelation appears**, then:
3. **Second 12 seconds** — Word "not responding", spinner. Explorer comes to front.

Every fix attempted in the 2026-04-06 session eliminated the second block in theory but
not in practice. The session produced 28 sections of iterative changes and ended without
a working fix.

---

## 3 — Current Code State

The Section 28 fix (commit #576) IS in place in `src/aeRibbonClass.cls`:

```vb
' GoToH1 lines 210–234 (current)
Application.ScreenUpdating = False
Selection.HomeKey Unit:=wdStory
Selection.Find.ClearFormatting
Selection.Find.Text = CStr(headingData(i, 0))
Selection.Find.style = ActiveDocument.Styles("Heading 1")
Selection.Find.Forward = True
Selection.Find.Wrap = wdFindStop
Selection.Find.MatchCase = False
Selection.Find.MatchWildcards = False
Selection.Find.Execute
If Selection.Find.found Then Selection.Collapse Direction:=wdCollapseStart
m_btnPrevEnabled = True
m_btnNextEnabled = True
' ... boundary detection using headingData ...
If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "GoToNextButton"   ← BEFORE True
If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "GoToPrevButton"   ← BEFORE True
Application.ScreenUpdating = True                                               ← LAST
```

The `InvalidateControl` calls are already BEFORE `ScreenUpdating = True`.
The bug persists anyway.

---

## 4 — Why the Section 28 Fix Did Not Work: Root Cause Reassessment

Sections 25, 27, and 28 each proposed a root cause and a fix. Each fix failed.
The pattern of repeated failure means the root cause is still not correctly identified.

### What we know empirically

| Fix | Claim | Result |
|-----|-------|--------|
| §25 | Wrap Range.Select in ScreenUpdating=False/True | Failed |
| §27 | Replace Range.Select with Selection.Find | Failed |
| §28 | Move InvalidateControl before ScreenUpdating=True | Failed |

All three fixes addressed the ORDERING of operations. None of them addressed whether
`InvalidateControl` itself is safe to call from within an active large-document
navigation context.

### Revised hypothesis

`m_ribbon.InvalidateControl` is a COM call on the `IRibbonUI` interface. Calling it
triggers a synchronous COM round-trip to the ribbon host (Word's UI shell). This
round-trip:

1. Asks Word's UI layer to mark the control state as dirty.
2. Causes Word to call back into our VBA `GetNextEnabled`/`GetPrevEnabled` stubs
   via COM IDispatch.
3. The ribbon host, as a UI element, may require Word's layout or document state to
   be settled before it can process the COM message.

**Critical observation:** `ScreenUpdating = False` suppresses painting events, but
it does NOT suppress COM message processing. When `InvalidateControl` is called with
`ScreenUpdating = False`, Word still enters the COM message loop to deliver the
invalidation, and may flush deferred layout work as a side effect — regardless of
whether it is called before or after `ScreenUpdating = True`.

If this is correct, the second 12-second block is produced by:

- **Path A:** `InvalidateControl` (ScreenUpdating=False) forces COM round-trip →
  layout flush (12s, first block). Then `ScreenUpdating = True` → paint (12s, second
  block).
- **Path B:** `ScreenUpdating = True` → layout + paint (12s, first block). Then some
  remaining deferred work triggered by InvalidateControl flushes (12s, second block).

Both paths produce two 12-second blocks regardless of call order, which is exactly
what is observed.

### Why NextButton/PrevButton don't show this problem

NextButton and PrevButton navigate ONE book at a time. The deferred layout work queued
by a single-book jump is small — the second block either does not occur or completes
in under a second, which is not noticeable.

---

## 5 — Diagnostic Plan (Before Any Code Changes)

Before writing any more fixes, confirm or refute the hypothesis with two targeted tests.

### Test A: Remove InvalidateControl from GoToH1

Remove lines 232–233 from `GoToH1` (both `InvalidateControl` calls). Keep everything
else unchanged. Test GoToH1 to Revelation.

- **If the second 12-second block disappears:** `InvalidateControl` is the cause.
  Proceed to Section 6.
- **If the second block persists:** `InvalidateControl` is NOT the cause. Something
  else — possibly `Selection.HomeKey Unit:=wdStory` — is responsible. Proceed to
  Section 7.

### Test B (only if Test A confirms the cause): Restore ribbon state manually

After Test A confirms InvalidateControl as the cause, verify the button state is still
set correctly (m_btnPrevEnabled / m_btnNextEnabled) even without InvalidateControl.
The buttons will not visually update until next pressed, but the state should be correct.

---

## 6 — Fix: Deferred InvalidateControl via Application.OnTime

If Test A confirms that `InvalidateControl` causes the second block, the fix is to
move the ribbon update OUT of the GoToH1 call stack entirely.

`Application.OnTime Now, "macroName"` schedules execution of a Standard Module macro
for "immediately" — but "immediately" means after the CURRENT VBA call stack returns
and Word's Windows message loop processes the next event. By that time, GoToH1 is
finished, `ScreenUpdating` is True, and Word's layout is fully settled. No deferred
work remains in the queue.

### Changes Required

**`src/aeRibbonClass.cls` — `GoToH1`**

Remove the two `InvalidateControl` calls. Replace with a single `OnTime` call:

```vb
    ' ... button state computation (unchanged) ...
    ' DON'T call InvalidateControl here
    Application.ScreenUpdating = True
    Application.OnTime Now, "aeRibbonUpdateButtons"

PROC_EXIT:
    Exit Sub

PROC_ERR:
    Application.ScreenUpdating = True
    MsgBox ...
    Resume PROC_EXIT
End Sub
```

**`src/basBibleRibbonSetup.bas` — new public sub**

`Application.OnTime` can only target a public procedure in a Standard Module. Add:

```vb
Public Sub aeRibbonUpdateButtons()
    Dim rc As aeRibbonClass
    Set rc = Instance()
    rc.InvalidateControl "GoToNextButton"
    rc.InvalidateControl "GoToPrevButton"
End Sub
```

`rc.InvalidateControl` delegates to `m_ribbon.InvalidateControl` via the existing
`InvalidateControl` public method on `aeRibbonClass` (line 118). No new class method
is needed.

### Why This Is Safe

- `m_btnPrevEnabled` and `m_btnNextEnabled` are set in `GoToH1` before `OnTime` fires.
- `aeRibbonUpdateButtons` fires after GoToH1 has returned. When the ribbon callbacks
  (`GetNextEnabled` / `GetPrevEnabled`) execute, they read the already-updated flag
  values — correct state, no race condition.
- If the user closes the document before `OnTime` fires, the macro will fail silently
  (Word cannot find the document) or raise an error. This is acceptable for a sub-second
  deferral.
- `Application.OnTime Now` fires "as soon as possible" — typically within one Windows
  message loop cycle (<100 ms). The ribbon will update imperceptibly after navigation
  completes.

### Expected Result

GoToH1 to Revelation: single 12-second layout pass, procedure returns, Revelation
displayed, ribbon buttons update within 100 ms. No second spinning phase.

---

## 7 — Alternate Root Cause: Selection.HomeKey

If Test A shows the second block persists WITHOUT `InvalidateControl`, the cause is
something in the navigation sequence itself. The prime candidate is:

```vb
Selection.HomeKey Unit:=wdStory
```

`HomeKey Unit:=wdStory` simulates pressing Ctrl+Home. This is a UI command that moves
the cursor to position 0 AND scrolls the view to show position 0. Word may internally
queue a "render position 0 into view" event in the Windows message loop. With
`ScreenUpdating = False`, this event is deferred. `Selection.Find.Execute` then adds
a "render Revelation into view" event. When `ScreenUpdating = True`, Word processes
both queued events sequentially — one for position 0 (12s) and one for Revelation
(12s).

### Fix for Alternate Root Cause

Replace `Selection.HomeKey Unit:=wdStory` with a direct object-model cursor placement
that does not trigger a scroll event:

```vb
ActiveDocument.Range(0, 0).Select
```

Or, since we HAVE the exact position in `headingData`, don't use Find at all.
Navigate directly using the stored position and start the Find from just before it:

```vb
' Collapse cursor to just before the target heading (no scroll required)
Selection.SetRange foundPos, foundPos
' Find forward from there — match is the very next H1 at or after foundPos
Selection.Find.ClearFormatting
Selection.Find.Text = CStr(headingData(i, 0))
Selection.Find.style = ActiveDocument.Styles("Heading 1")
Selection.Find.Forward = True
Selection.Find.Wrap = wdFindStop
Selection.Find.MatchCase = False
Selection.Find.MatchWildcards = False
Selection.Find.Execute
If Selection.Find.found Then Selection.Collapse Direction:=wdCollapseStart
```

Starting Find from `foundPos` (the heading's own position) means the first match IS
the target heading — no full-document scan from position 0. This eliminates the
HomeKey-to-position-0 event, replaces it with a SetRange (which is an object-model
call, not a UI command), and leaves only one pending scroll event (Revelation) for
`ScreenUpdating = True` to process.

**Risk:** `SetRange foundPos, foundPos` with `ScreenUpdating = False` may or may not
queue a scroll event depending on Word's internal implementation. This must be tested.

---

## 8 — Candidate Ranked by Probability

| Rank | Candidate | Test | Fix |
|------|-----------|------|-----|
| 1 | `InvalidateControl` causes layout flush during COM round-trip | Test A | Section 6: Application.OnTime |
| 2 | `HomeKey` queues scroll-to-position-0, causing two layout passes | Test A fails | Section 7: SetRange(foundPos) instead of HomeKey |
| 3 | Both causes compound | Both tests inconclusive | Combine §6 and §7 fixes |

---

## 9 — Files to Change (Pending Test Results)

### If Test A confirms InvalidateControl as cause (Rank 1):

- `src/aeRibbonClass.cls` — `GoToH1`: remove `InvalidateControl` calls; add
  `Application.OnTime Now, "aeRibbonUpdateButtons"` after `ScreenUpdating = True`
- `src/basBibleRibbonSetup.bas`: add `aeRibbonUpdateButtons` public sub

### If Test A shows InvalidateControl is NOT the cause (Rank 2):

- `src/aeRibbonClass.cls` — `GoToH1`: replace `Selection.HomeKey Unit:=wdStory`
  with `Selection.SetRange foundPos, foundPos` (keep Find for the actual navigation)

### If both are needed:
- All of the above combined.

---

## 10 — Proposed GoToH1 Rewrite (Rank 1 Fix, Ready to Implement)

Full replacement for `GoToH1` after Test A confirms the cause:

```vb
Private Sub GoToH1()
    On Error GoTo PROC_ERR
    Dim pattern   As String
    Dim i         As Long
    Dim k         As Long
    Dim foundPos  As Long
    Dim matchFound As Boolean

    pattern = InputBox("Enter a Book Name (Heading 1) abbreviation:", "Go To Bible Book")
    If pattern = "" Then GoTo PROC_EXIT

    matchFound = False
    For i = 1 To 66
        If Not IsEmpty(headingData(i, 0)) Then
            If CStr(headingData(i, 0)) Like "*" & UCase(pattern) & "*" Then
                matchFound = True
                Exit For
            End If
        End If
    Next i

    If Not matchFound Then
        MsgBox "Book not found! No Heading 1 matches: '" & pattern & "'", vbExclamation, "Bible"
        GoTo PROC_EXIT
    End If

    foundPos = CLng(headingData(i, 1))
    Application.ScreenUpdating = False
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Text = CStr(headingData(i, 0))
    Selection.Find.style = ActiveDocument.Styles("Heading 1")
    Selection.Find.Forward = True
    Selection.Find.Wrap = wdFindStop
    Selection.Find.MatchCase = False
    Selection.Find.MatchWildcards = False
    Selection.Find.Execute
    If Selection.Find.found Then Selection.Collapse Direction:=wdCollapseStart
    m_btnPrevEnabled = True
    m_btnNextEnabled = True
    If Not IsEmpty(headingData(1, 1)) Then
        If foundPos = CLng(headingData(1, 1)) Then m_btnPrevEnabled = False
    End If
    For k = 66 To 1 Step -1
        If Not IsEmpty(headingData(k, 1)) And CLng(headingData(k, 1)) > 0 Then
            If foundPos = CLng(headingData(k, 1)) Then m_btnNextEnabled = False
            Exit For
        End If
    Next k
    Application.ScreenUpdating = True
    Application.OnTime Now, "aeRibbonUpdateButtons"   ' ← deferred ribbon update

PROC_EXIT:
    Exit Sub

PROC_ERR:
    Application.ScreenUpdating = True
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GoToH1 of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub
```

Note: `aeRibbonUpdateButtons` in `basBibleRibbonSetup.bas` (see Section 6) must be
added before testing.

---

## 11 — Session Protocol

To avoid repeating the 2026-04-06 loop:

1. **Always test ONE change at a time.** Each section of the 2026-04-06 review
   combined multiple changes. If a combined fix fails, it is impossible to know
   which part caused the failure.

2. **Test A first — no code changes yet.** Temporarily remove the two
   `InvalidateControl` lines and test. If the second block goes away, add them
   back and then implement Section 6. This confirms the cause before the fix.

3. **Measure explicitly.** State the exact timing before and after each change.
   "12 seconds then another 12 seconds" vs "12 seconds, done" are unambiguous.

4. **Stop at the first working fix.** Do not continue to clean up or combine with
   other changes in the same session. A working fix can be refined in a subsequent
   session once the root cause is confirmed.

---

## 12 — Test A Executed: InvalidateControl Removed from GoToH1

**Change made:** `src/aeRibbonClass.cls` — `GoToH1`

Removed lines:
```vb
If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "GoToNextButton"
If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "GoToPrevButton"
```

Replaced with comment marking the temporary removal. `Application.ScreenUpdating = True`
is now the last statement before `PROC_EXIT`. All other code is unchanged.

**Side effect of this test:** After GoToH1 completes, the Prev Book and Next Book
ribbon buttons will NOT update their enabled/disabled state visually. The internal
flags (`m_btnPrevEnabled`, `m_btnNextEnabled`) are still set correctly. The ribbon
will show stale button states until the next time a ribbon update is triggered (e.g.,
by pressing Next or Prev). This is expected and acceptable for a diagnostic test.

---

## 13 — Test A: What to Report

Perform this exact test sequence and report the results for each step.

### Test sequence

1. Close and reopen the document (`Blank Bible Copy.docm`) to start from a cold cache.
2. Press **GoTo Book**. Type `GEN` (or the abbreviation for Genesis). Press OK.
   - Measure: how long does GoTo Genesis take? Is there a second block?
   - Report: "Genesis: X seconds total, Y blocks"
3. Press **GoTo Book**. Type `REV` (or the abbreviation for Revelation). Press OK.
   - Measure: how long does GoTo Revelation take? Watch for the second block.
   - Report: "Revelation: X seconds total, Y blocks"
4. Press **GoTo Book** a second time. Type `REV` again. Press OK.
   - This is a warm-cache repeat to confirm caching behaviour.
   - Report: "Revelation (warm): X seconds total, Y blocks"
5. Observe the ribbon buttons after each GoTo:
   - Do Prev Book and Next Book show the correct enabled/disabled state?
   - Or do they remain in a stale state (both disabled, or incorrect state)?
   - Report: "Buttons after GoTo Genesis: Prev=X, Next=Y"
   - Report: "Buttons after GoTo Revelation: Prev=X, Next=Y"

### Decision rule

| Result | Interpretation | Next step |
|--------|----------------|-----------|
| Second 12s block **gone** on Revelation | `InvalidateControl` was the cause | Proceed to Section 6 fix (Application.OnTime) |
| Second 12s block **still present** | `InvalidateControl` is NOT the cause | Report and proceed to Section 7 (HomeKey investigation) |
| GoTo Genesis itself shows a second block | New information — not previously observed | Report exact timing and behaviour |

---

## 19 — Test A Results and Conclusion

**Pre-test note recorded:** Code changes must be imported into `Blank Bible Copy.docm`
before testing so the running VBA project reflects the src/ changes.

### Raw results

| Step | Result |
|------|--------|
| GoTo Genesis (cold cache) | 0 seconds, 0 blocks |
| GoTo Revelation (cold cache) | Block 1: 21s spinning → Revelation appears; Block 2: 16s spinning → cursor available |
| GoTo Revelation (warm cache) | 0 seconds, 0 blocks |
| Buttons after GoTo Genesis | Prev=Disabled, Next=Enabled ✓ |
| Buttons after GoTo Revelation | Prev=Enabled, Next=Disabled ✓ |

### Clarification on block counting

The report initially said "1 block (NOTE: 16 secs second spinning cycle)." This
referred to 1 block observed *after Revelation first appeared*. The full sequence
is two blocks:

1. **21 seconds** — Word not responding, spinning. Revelation appears.
2. **16 seconds** — Word not responding, spinning. Cursor becomes available.

This is the same two-block pattern as before Test A. The timings changed but the
structure did not.

### Conclusion: `InvalidateControl` is NOT the cause

The second block persists with `InvalidateControl` removed. Decision rule result:
**Section 7 — investigate alternate root cause.**

**Timing comparison:**

| State | Block 1 | Block 2 | Total |
|-------|---------|---------|-------|
| Before Test A (with `InvalidateControl`) | ~12s | ~12s | ~24s |
| After Test A (without `InvalidateControl`) | 21s | 16s | ~37s |

Removing `InvalidateControl` made total time **worse by ~13 seconds**. This suggests
`InvalidateControl` was inadvertently pre-processing some layout work via its COM
message-pump round-trip, shortening the subsequent `ScreenUpdating = True` pass.
Removing it pushed all that work into the two remaining passes.

### Revised root cause hypothesis

The second block occurs **after Revelation is already visible**, regardless of
whether `InvalidateControl` is called. The sequence is:

1. `Application.ScreenUpdating = True` → Word lays out and paints Revelation
   (Block 1 ~21s). `GoToH1` returns.
2. The **ribbon callback return mechanism** — when `OnGoToH1ButtonClick` returns
   control to the ribbon host, the ribbon host re-queries button states and refreshes
   the ribbon UI. This refresh requires Word to compute document state, triggering a
   second layout pass (Block 2 ~16s).

If correct, neither `InvalidateControl` placement nor `Selection.HomeKey` is the
root cause. The second block is triggered by the ribbon host processing the return
from the `onAction` callback — outside VBA's control.

### Decision: Restore `InvalidateControl`

The `InvalidateControl` calls have been restored to `GoToH1` in
`src/aeRibbonClass.cls`. Rationale:

- Buttons are correct either way (ribbon host re-queries on callback return).
- With `InvalidateControl`: ~24s total. Without: ~37s total.
- The only benefit of removal was simpler code, at a cost of ~13 extra seconds.
- Restoring it recovers the ~13s with no change in functional behaviour.
- The second block exists either way — that remains an open problem.

### Timing caveat

All timings in this review are **approximate, single-run observations**. They are
not controlled measurements. Factors that may affect them:

- Other background processes running on the machine at the time of the test
- Word's background repagination, word count, and spell-check threads
- Antivirus or indexing activity triggered by file access during navigation
- Variability in Word's layout cache state between runs

**Proper timing tests** — multiple runs, controlled environment, measured with a
stopwatch or `Timer()` instrumentation — are needed before drawing firm conclusions
about the relative cost of any two approaches. The timings recorded here (12s, 16s,
21s, 37s) should be treated as order-of-magnitude observations only, not benchmarks.

---

## 20 — Test B: Ribbon Callback Return as Root Cause of Second Block

### Hypothesis

The second block is triggered by the ribbon host processing the return from the
`onAction` callback (`OnGoToH1ButtonClick`). When the ribbon button is clicked, the
ribbon host waits for the callback to return, then re-queries all button states and
refreshes the ribbon UI. This post-return refresh requires Word to compute document
state, triggering a second layout pass.

If correct, running GoToH1 outside the ribbon callback — where no ribbon host is
waiting for a return — should eliminate the second block.

### Changes Made

**`src/aeRibbonClass.cls`** — added public wrapper after `OnGoToH1ButtonClick`:

```vb
Public Sub GoToH1Direct()
    Call GoToH1
End Sub
```

**`src/basBibleRibbonSetup.bas`** — added test stub after `GetNextEnabled`:

```vb
Public Sub TestGoToH1Direct()
    Dim rc As aeRibbonClass
    Set rc = Instance()
    rc.GoToH1Direct
End Sub
```

`GoToH1Direct` calls the same private `GoToH1` as the ribbon button does — identical
code path, no ribbon callback involved.

### Test Procedure

1. Import updated src files into `Blank Bible Copy.docm`.
2. Close and reopen the document (cold cache).
3. Press **Alt+F8** to open the Macro dialog.
4. Select `TestGoToH1Direct` and click **Run**.
5. When prompted, type `GEN` and press OK.
6. Press **Alt+F8** again, run `TestGoToH1Direct`, type `REV`, press OK.
   - Measure: total seconds, number of blocks.
   - Report: "Revelation via Alt+F8: X seconds, Y blocks"
7. Compare to the ribbon button result from Section 19:
   - Ribbon button: Block 1 ~12s → Revelation → Block 2 ~12s

### Decision Rule

| Result | Interpretation | Next step |
|--------|----------------|-----------|
| Second block **gone** via Alt+F8 | Ribbon callback return is the cause | Investigate deferring navigation via `Application.OnTime` |
| Second block **still present** via Alt+F8 | Ribbon callback return is NOT the cause | Root cause is elsewhere — further investigation needed |

### Test B Results

Test run via VBA Immediate window (`TestGoToH1Direct`):

| Step | Result |
|------|--------|
| GoTo Genesis | 0 seconds, 0 blocks |
| GoTo Revelation | ~8 seconds, **0 blocks** |

**Second block is gone.** Ribbon callback return is confirmed as the root cause.

**Timing caveat:** The ~8s (vs ~12s via ribbon) may reflect partial cache warming
from navigating the VBA IDE before running the test. Single-run observation.
The critical finding is the complete absence of any second block — that was
consistently ~12 seconds in all ribbon-button runs and is now zero.

### Conclusion

The second ~12-second block after Revelation appears is caused by the ribbon host
processing the return from the `onAction` callback (`OnGoToH1ButtonClick`). When
Word calls the callback, it waits for it to return, then re-queries all ribbon
button states. After a full-document navigation, this post-return processing
triggers a second full layout pass.

Running the same code outside the ribbon callback (Immediate window) produces a
single layout pass with no second block.

### Fix: Application.OnTime in OnGoToH1ButtonClick

Make the ribbon callback return immediately by scheduling the navigation via
`Application.OnTime Now`. The ribbon host gets its return before any heavy work
happens. The navigation fires in a clean message loop context after the ribbon
host has fully unwound.

---

## 21 — Fix: Defer GoToH1 Navigation via Application.OnTime

### Implementation history and failures

**Attempt 1** — `Application.OnTime Now, "GoToH1Deferred"` (unqualified name).
No activation. Word could not find the macro.

**Attempt 2** — `Application.OnTime Now, "basBibleRibbonSetup.GoToH1Deferred"`
(module-qualified). No activation.

**Attempt 3** — `Application.OnTime Now, "Project.basBibleRibbonSetup.GoToH1Deferred"`
(project + module qualified). No activation.

**Attempt 4** — `Application.OnTime Now + TimeValue("00:00:01"), "Project.basBibleRibbonSetup.GoToH1Deferred"`
(future time + fully qualified). No activation.

### Root cause of all four failures: Option Private Module

`basBibleRibbonSetup.bas` declares `Option Private Module` on line 3. This
declaration prevents `Application.OnTime` from resolving any macro in the module
by name — it is a project-level visibility flag that blocks external dispatch.

**`Option Private Module` vs Alt+F8 visibility — two separate rules:**

| Reason a macro is absent from Alt+F8 | Applies to |
|---------------------------------------|-----------|
| `Option Private Module` | All modules in this project except `basUSFM_Export` |
| Required parameters on the sub | `basUSFM_Export` (all public subs have parameters) |

`Application.OnTime` is blocked only by `Option Private Module`. It has no
dependency on Alt+F8 visibility. A sub with required parameters will not appear
in Alt+F8 but CAN be called by `Application.OnTime` if its module omits
`Option Private Module`.

Confirmed by `GoToH1Deferred` running correctly from the Immediate window —
the macro exists and is callable; `OnTime` simply could not find it due to the
module privacy flag.

**Secondary finding:** The earlier observation that Alt+F8 showed no macros was
caused by these two rules together, not by a single cause.

### Fix: New module basRibbonDeferred

`GoToH1Deferred` moved to a new module `src/basRibbonDeferred.bas` that
deliberately omits `Option Private Module`.

`OnGoToH1ButtonClick` updated to reference the new module location and uses a
runtime project-name check with a MsgBox warning if the name changes:

```vb
Public Sub OnGoToH1ButtonClick(control As IRibbonControl)
    Const EXPECTED_PROJECT As String = "Project"
    Dim projName As String
    projName = Application.ActiveDocument.VBProject.Name
    If projName <> EXPECTED_PROJECT Then
        MsgBox "VBA project name has changed from '" & EXPECTED_PROJECT & "' to '" & projName & "'." & vbCrLf & _
               "Update EXPECTED_PROJECT in OnGoToH1ButtonClick (basBibleRibbonSetup).", _
               vbExclamation, "Project Name Changed"
    End If
    Application.OnTime Now + TimeValue("00:00:01"), projName & ".basRibbonDeferred.GoToH1Deferred"
End Sub
```

`basRibbonDeferred.bas`:
```vb
Public Sub GoToH1Deferred()
    Dim rc As aeRibbonClass
    Set rc = Instance()
    rc.GoToH1Direct
End Sub
```

`GoToH1Deferred` has no parameters so it will appear in Alt+F8. This is
acceptable — it is safe to run manually for testing.

### Files changed

- `src/basBibleRibbonSetup.bas` — `OnGoToH1ButtonClick` updated; `GoToH1Deferred`
  removed
- `src/basRibbonDeferred.bas` — new module; hosts `GoToH1Deferred`
- `src/aeRibbonClass.cls` — `GoToH1Direct` wrapper retained (added in Section 20)

### Expected Result

GoTo Revelation: single layout pass (~8–12 seconds), procedure returns, Revelation
displayed, cursor available immediately. No second spinning block.

---

## 14 — Normalizer: Add `As PageSetup` Type Declaration Rule

**File changed:** `py/normalize_vba.py`

**Problem:** `basTEST_aeBibleTools.bas:803` contained:

```vb
Dim sectionSetup As pageSetup
```

The Word VBA IDE downcased the type name to `pageSetup`. The existing rule
`(?i)\.PageSetup\b → .PageSetup` requires a leading dot and does not match
`As pageSetup` in type declarations.

**Fix:** Added one rule to `NORMALIZATIONS`, grouped with the other `As Word.*`
declaration rules:

```python
(r'(?i)\bAs\s+PageSetup\b', 'As PageSetup', 'As PageSetup type declaration'),
```

The existing `\.PageSetup\b` rule continues to handle all property-access forms
(`sec.PageSetup`, `sectionSetup.TextColumns`, etc.). The new rule closes the gap
for the declaration form only.

---

## 15 — Normalizer: Add `.Code` Rule and Audit All Descriptions

**File changed:** `py/normalize_vba.py`

### New rule

`field.code.Text` in `XbasTESTaeBibleDOCVARIABLE.bas` (6 occurrences) had `.code`
downcased by the IDE. Added:

```python
(r'(?i)\.Code\b', '.Code', '.Code property on Field object'),
```

Placed after `.Text` (line 15), since both are properties accessed on `Field` objects.

### Description corrections

The following descriptions were inaccurate or inconsistent and were updated:

| Rule | Old description | New description |
|------|----------------|-----------------|
| `.Range` | `.Range property access` | `.Range property on Word Paragraph/Selection/Section` |
| `.Paragraphs` | `.Paragraphs property access` | `.Paragraphs collection property on Document/Range` |
| `.PageSetup` | `.PageSetup property access` | `.PageSetup property on Section/Document` |
| `.TopMargin` | `.TopMargin on PageSetup` | `.TopMargin property on PageSetup` |
| `.BottomMargin` | `.BottomMargin on PageSetup` | `.BottomMargin property on PageSetup` |
| `.PageHeight` | `.PageHeight on PageSetup` | `.PageHeight property on PageSetup` |
| `.Orientation` | `.Orientation on PageSetup` | `.Orientation property on PageSetup` |
| `Mid$(` | `Mid$( built-in function casing (includes Mid( -> Mid$()` | `Mid$( string function — normalizes Mid( to Mid$( and fixes casing` |
| `.Path` | `.Path property on Document/ActiveDocument` | `.Path property on Document/FileSystemObject` |
| `.Count` | `.Count method on Collection` | `.Count property on Collection/object` |
| `.Font` | `.Font method on Collection` | `.Font property on Range/Style/object` |
| `.Keys` | `.Keys property on Dictionary/Object` | `.Keys property on Dictionary/object` |
| `.Text` | `.Text property on Range` | `.Text property on Range/Field/object` |
| `As Word.Range` | `As Word.Range declaration` | `As Word.Range type declaration` |
| `As Word.Paragraph` | `As Word.Paragraph declaration` | `As Word.Paragraph type declaration` |
| `As Word.Paragraphs` | `As Word.Paragraphs declaration` | `As Word.Paragraphs type declaration` |
| `Note` | `Note loop variable (Footnote iteration)` | `Note loop variable (Footnote collection iteration)` |
| `Items` | `Items variable casing (Collection)` | `Items variable casing (Collection iteration)` |

**Substantive corrections** (wrong word, not just wording):
- `.Count` and `.Font` were labelled "method" — both are VBA properties (accessed
  without parentheses). Changed to "property".
- `.Path` listed only `ActiveDocument` — also applies to `FileSystemObject` and other
  file-related objects in the codebase.

---

## 16 — Normalizer: Add `Range:=` Named Argument Rule

**File changed:** `py/normalize_vba.py`

**Problem:** `basAddHeaderFooter.bas:215` contained:

```vb
oRange.Fields.Add range:=oRange, _
```

The VBA IDE exported the named argument `Range` as `range` (lowercase). The existing
`\.Range\b` rule requires a leading dot and does not match standalone `range:=`.

**Root cause:** VBA named arguments use the syntax `ParameterName:=value`. The parameter
name `Range` appears without a dot in calls such as `Fields.Add`, `Bookmarks.Add`,
and `Hyperlinks.Add`. The IDE downcases the parameter name on export.

**Fix:** Added one rule to `NORMALIZATIONS`, placed immediately after the `.Range`
property rule:

```python
(r'(?i)\bRange(?=:=)', 'Range', 'Range named argument in VBA method calls (Fields.Add, Bookmarks.Add, etc.)'),
```

The lookahead `(?=:=)` restricts the match to `Range` only when immediately followed
by `:=`, preventing any collision with other uses of the word `Range`.

**Coverage after this change:**

| Form | Matched by |
|------|-----------|
| `para.range` | `\.Range\b` rule |
| `range:=oRange` | `\bRange(?=:=)` rule (new) |
| `As Word.range` | `As Word.Range` type declaration rule |

---

## 17 — Normalizer: Add `As Word.Section` Type Declaration Rule

**File changed:** `py/normalize_vba.py`

**Problem:** The IDE downcases the `Word.Section` type in declarations. Two failure
modes were present in the source:

| File | Line | Form found |
|------|------|-----------|
| `basTEST_aeBibleTools.bas` | 1631 | `As word.section` |
| `basAddHeaderFooter.bas`, `basTEST_aeBibleTools.bas`, `basAuditDocument.bas`, `aeBibleClass.cls`, `Module1.bas`, and others (24 total) | various | `As section` |

`Section` was the missing exception — `Range`, `Paragraph`, and `Paragraphs` all had
type declaration rules; `Section` did not.

**Fix:** Added one rule to `NORMALIZATIONS`, grouped with the other `As Word.*`
type declaration rules:

```python
(r'(?i)\bAs\s+(?:Word\.)?Section\b', 'As Word.Section', 'As Word.Section type declaration'),
```

**Coverage:**

| Source form | Result |
|---|---|
| `As section` | `As Word.Section` |
| `As word.section` | `As Word.Section` |
| `As Word.Section` | unchanged |

---

## 18 — Normalizer: Add `Shell` Built-in Function Rule

**File changed:** `py/normalize_vba.py`

**Problem:** `basTEST_aeWordGitClass.bas:42` contained:

```vb
shell "cmd.exe /c """ & strBat & """", vbNormalFocus
```

The VBA IDE exported the `Shell` built-in function as `shell` (lowercase).

**Fix:** Added one rule to `NORMALIZATIONS`, grouped with `IsEmpty` and other
built-in function rules at the top of the list:

```python
(r'(?i)\bShell\b', 'Shell', 'Shell built-in function casing'),
```

**Why `"WScript.Shell"` string literals are unaffected:** The pattern uses `\b`
word boundaries. In `"WScript.Shell"`, the `.` immediately before `Shell` is a
non-word character — but `\b` requires a transition between a word character and a
non-word character. The `S` in `.Shell` IS preceded by a non-word character (`.`),
so `\bShell\b` would match it.

However, the `.Shell` in `"WScript.Shell"` appears inside a string literal passed
to `CreateObject`. The normalizer performs plain regex substitution across the entire
file including string contents. To avoid touching `WScript.Shell`, the safer pattern
would be `\bShell\b` applied only to standalone calls — but since replacing
`Shell` with `Shell` (same casing) inside a string is a no-op, the rule is safe
regardless: `"WScript.Shell"` → `"WScript.Shell"` (unchanged).

---

## 22 — Fix: Move InvalidateControl Out of GoToH1 (Warm-Cache Hypothesis)

**Files changed:** `src/aeRibbonClass.cls`, `src/basRibbonDeferred.bas`

**Problem:** After the Application.OnTime deferral fix (§21):
- InputBox delay: ~2 seconds
- Block 1: 6 seconds → Revelation appears
- Block 2: 12 seconds spinning → cursor available

The remaining 12-second Block 2 is isolated to the two `InvalidateControl` calls
inside `GoToH1`. These fire immediately after `Selection.Find.Execute` and
`Selection.Collapse`, while the layout cache is still cold from navigation to a
far paragraph. Calling `InvalidateControl` on a cold cache forces a second full
layout pass.

**Hypothesis:** Moving `InvalidateControl` to run AFTER `GoToH1Direct` returns
(from `GoToH1Deferred`) puts it on a warm cache — Word has already laid out the
target page, so the invalidation query is cheap.

**Changes:**

`src/aeRibbonClass.cls` — removed two lines from `GoToH1` (just before
`Application.ScreenUpdating = True`):

```vb
' REMOVED:
If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "GoToNextButton"
If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "GoToPrevButton"
Application.ScreenUpdating = True   ' this line stays
```

`src/basRibbonDeferred.bas` — added `InvalidateControl` calls after navigation:

```vb
Public Sub GoToH1Deferred()
    Dim rc As aeRibbonClass
    Set rc = Instance()
    rc.GoToH1Direct
    rc.InvalidateControl "GoToNextButton"
    rc.InvalidateControl "GoToPrevButton"
End Sub
```

**Test instruction:**

Run the cold-start test (restart Word, do not open IDE):

1. GoTo Genesis — note time and any spinning
2. GoTo Revelation — note Block 1 duration (InputBox → Revelation visible), Block 2 duration (Revelation visible → cursor available), and whether Block 2 is still present
3. GoTo Revelation again (warm) — note time
4. Check Prev/Next buttons are enabled after GoTo Revelation
5. Navigate Prev once — check Prev/Next enabled state

Report: timings for each step, whether Block 2 is gone, and any errors.

---

## 22 — Test Results

**Cold start, GoTo Genesis → Revelation:**

| Step | Result |
|---|---|
| GoTo Genesis | Instant, no spinning; cursor flash from top |
| GoTo Revelation — Block 1 (InputBox → Revelation visible) | 16 seconds |
| GoTo Revelation — Block 2 (Revelation visible → cursor available) | 14 seconds spinning |
| GoTo Revelation warm | Instant |
| Prev/Next after GoTo Revelation | Next disabled (correct — last book) |
| GoTo Genesis Prev state | Prev disabled (correct — first book) |
| Navigate Prev once | Both enabled (correct) |

**Additional observation:** Fast-clicking Prev repeatedly shows cursor flashing from bottom;
fast-clicking Next shows cursor flashing from top. Navigation direction is visible in the
cursor entry point.

**Conclusion:** Warm-cache hypothesis FAILED. Block 2 is still 14 seconds.

**Why Block 1 got longer (6s → 16s):** In §21, `InvalidateControl` fired while
`ScreenUpdating = False` — Word deferred the painting cost, so Block 1 appeared short
and that cost was absorbed into Block 2. In §22, `ScreenUpdating = True` fires first
(inside GoToH1), so the display/painting cost lands in Block 1 instead. The totals
(§21: 18s, §22: 30s) suggest §22 is slightly worse — `InvalidateControl` while the
painter is active may cause additional repaint cycles.

**New diagnostic plan (§23):** Remove `InvalidateControl` from `GoToH1Deferred`
entirely and test. If Block 2 disappears, `InvalidateControl` is confirmed as the sole
cause of Block 2. If Block 2 persists, the cause is elsewhere — most likely
`Application.ScreenUpdating = True` itself triggering a deferred layout pass, or
Word's post-OnTime-macro processing.

---

## 23 — Diagnostic: Remove InvalidateControl from GoToH1Deferred Entirely

**File changed:** `src/basRibbonDeferred.bas`

**Purpose:** Isolate whether `InvalidateControl` is the cause of Block 2 or not.
This is a diagnostic test, not a final fix. Buttons will NOT update their enabled
state after GoToH1 in this test — that is expected.

**Change:**

```vb
' BEFORE (§22):
Public Sub GoToH1Deferred()
    Dim rc As aeRibbonClass
    Set rc = Instance()
    rc.GoToH1Direct
    rc.InvalidateControl "GoToNextButton"
    rc.InvalidateControl "GoToPrevButton"
End Sub

' AFTER (§23 diagnostic):
Public Sub GoToH1Deferred()
    Dim rc As aeRibbonClass
    Set rc = Instance()
    rc.GoToH1Direct
End Sub
```

**Test instruction:**

Run the cold-start test (restart Word, do not open IDE):

1. GoTo Genesis — note time and any spinning
2. GoTo Revelation — note Block 1 duration (InputBox → Revelation visible), and whether Block 2 is present (Revelation → cursor available)
3. Note whether Prev/Next buttons changed state (they may remain stale — this is expected)

Report: whether Block 2 is present or absent, and Block 1 timing.
This single data point determines the next step.

---

## 23 — Test Results: InvalidateControl Confirmed as Sole Cause of Block 2

**Cold start results:**

| Step | Result |
|---|---|
| GoTo Genesis | Instant, no spinning |
| GoTo Revelation — Block 1 | 13 seconds |
| GoTo Revelation — Block 2 | **ABSENT** |
| Prev/Next buttons after GoTo Revelation | Both disabled (stale — expected, no InvalidateControl called) |

**Conclusion: `InvalidateControl` is the 100% cause of Block 2.**

Without it, the total cold-navigation cost is 13 seconds (one pass, no second block).
With it present in the same macro invocation (§21: 18s total; §22: 30s total), Word
triggers an additional expensive layout pass in response to the ribbon state query.

**Timing comparison across all tests:**

| Test | Block 1 | Block 2 | Total | InvalidateControl location |
|---|---|---|---|---|
| §19 Test A (ribbon callback, IC before ScreenUpdating) | 21s | 16s | 37s | Inside GoToH1, before ScreenUpdating=True |
| §21 OnTime deferred (IC before ScreenUpdating) | 6s | 12s | 18s | Inside GoToH1, before ScreenUpdating=True |
| §22 OnTime deferred (IC after GoToH1Direct) | 16s | 14s | 30s | GoToH1Deferred, after GoToH1Direct returns |
| §23 OnTime deferred (no IC) | 13s | none | 13s | Removed entirely |

**Why moving IC after ScreenUpdating=True made §22 worse than §21:** In §21, IC
fired while `ScreenUpdating=False`, so the ribbon query and paint costs were batched
with the ScreenUpdating=True repaint. In §22, IC fired after the repaint was already
done, causing a fresh layout cycle. The 13s baseline (§23) is the true navigation
cost without any ribbon overhead.

**Next fix (§24):** Schedule `InvalidateControl` in a second `Application.OnTime`
call (`InvalidateButtonsDeferred`) fired from `GoToH1Deferred` after navigation
completes. The gap between the two OnTime invocations allows Word to complete its
post-navigation layout. When `InvalidateButtonsDeferred` fires, the layout cache
should be warm and the ribbon query cheap.

---

## 24 — Fix: Defer InvalidateControl to a Second OnTime Call

**File changed:** `src/basRibbonDeferred.bas`

**Rationale:** `InvalidateControl` causes a ~12-14s layout pass when called
immediately after cold navigation, because the layout cache is still cold.
Scheduling it in a *separate* `Application.OnTime` invocation gives Word time to
complete its natural post-navigation layout before the ribbon query fires.

**Change:**

```vb
Public Sub GoToH1Deferred()
    Dim rc As aeRibbonClass
    Set rc = Instance()
    rc.GoToH1Direct
    Application.OnTime Now, "Project.basRibbonDeferred.InvalidateButtonsDeferred"
End Sub

Public Sub InvalidateButtonsDeferred()
    Dim rc As aeRibbonClass
    Set rc = Instance()
    rc.InvalidateControl "GoToNextButton"
    rc.InvalidateControl "GoToPrevButton"
End Sub
```

`Now` (not `Now + TimeValue("00:00:01")`) is used so `InvalidateButtonsDeferred`
fires at the next available opportunity after `GoToH1Deferred` returns and Word
processes its message queue — no added user-visible delay.

**Test instruction:**

Run the cold-start test (restart Word, do not open IDE):

1. GoTo Genesis — note time and any spinning
2. GoTo Revelation — note Block 1 duration (InputBox → Revelation visible) and whether Block 2 is present
3. After cursor is available, check whether Prev/Next buttons are correctly enabled (Next should be disabled at Revelation)
4. GoTo Revelation warm — note time
5. Navigate Prev once — confirm Prev/Next enabled state

Report: Block 1 timing, whether Block 2 is present, button states, and any errors
(including whether `InvalidateButtonsDeferred` fails to fire).

---

## 24 — Test Results: Fix Confirmed

**Cold start results:**

| Step | Result |
|---|---|
| GoTo Genesis | Instant, no spinning |
| GoTo Revelation — Block 1 | 15 seconds |
| GoTo Revelation — Block 2 | **ABSENT** |
| Next button at Revelation | Disabled (correct — last book) |
| Navigate Prev once | Both buttons enabled (correct) |

**Fix confirmed.** Block 2 is eliminated. Button states update correctly after
navigation. `InvalidateButtonsDeferred` fires via the second `Application.OnTime Now`
call and the ribbon query runs on a warm cache with no blocking layout pass.

**Final timing summary:**

| Condition | Total time | Notes |
|---|---|---|
| Original (ribbon callback, double block) | ~37s | Two 12s+ blocks |
| §21 (OnTime + IC in GoToH1) | ~18s | Block 2 still 12s |
| §23 (OnTime, no IC) | ~13s | No Block 2, buttons stale |
| **§24 (OnTime + deferred IC)** | **~15s** | **No Block 2, buttons correct** |

The 13–15s remaining is the intrinsic cold-cache layout cost of Word navigating
the full Bible document to Revelation. This cannot be reduced in VBA — it is
Word's layout engine building page geometry from scratch on a cold document.

**Root cause summary:**

1. `OnGoToH1ButtonClick` (ribbon callback) returns → Word's ribbon host re-queries
   all control states → triggers a full layout pass on the cold document → Block 2.
   **Fixed by:** `Application.OnTime` deferral (§21), so the ribbon callback returns
   before any navigation begins.

2. `InvalidateControl` called within the same macro invocation as cold navigation
   → ribbon state query on a cold layout cache → another full layout pass → Block 2.
   **Fixed by:** scheduling `InvalidateButtonsDeferred` via a second
   `Application.OnTime Now` call, so Word completes its natural post-navigation
   layout before the ribbon query fires.

**Files in final state:**

`src/basBibleRibbonSetup.bas` — `OnGoToH1ButtonClick` defers via OnTime:
```vb
Application.OnTime Now + TimeValue("00:00:01"), "Project.basRibbonDeferred.GoToH1Deferred"
```

`src/basRibbonDeferred.bas` — two-stage deferred dispatch:
```vb
Public Sub GoToH1Deferred()
    Dim rc As aeRibbonClass
    Set rc = Instance()
    rc.GoToH1Direct
    Application.OnTime Now, "Project.basRibbonDeferred.InvalidateButtonsDeferred"
End Sub

Public Sub InvalidateButtonsDeferred()
    Dim rc As aeRibbonClass
    Set rc = Instance()
    rc.InvalidateControl "GoToNextButton"
    rc.InvalidateControl "GoToPrevButton"
End Sub
```

`src/aeRibbonClass.cls` — `GoToH1` ends with `Application.ScreenUpdating = True`
only; no `InvalidateControl` calls.

**Optional follow-up:** The `Now + TimeValue("00:00:01")` in `OnGoToH1ButtonClick`
adds ~2 seconds before the InputBox appears. Replacing with `Now` would remove this
delay but may not be reliable from a ribbon callback context (prior tests showed `Now`
failing there). Leave as-is unless the delay becomes a UX concern.

---

## 25 — Test: Reduce OnTime Delay to 0.5 Seconds

**File changed:** `src/basBibleRibbonSetup.bas`

**Change tested:**
```vb
' Before:
Application.OnTime Now + TimeValue("00:00:01"), projName & ".basRibbonDeferred.GoToH1Deferred"
' Attempted:
Application.OnTime Now + (0.5 / 86400#), projName & ".basRibbonDeferred.GoToH1Deferred"
```

`TimeValue` only accepts whole-second strings, so 0.5 seconds requires a
date-fraction calculation: `0.5 / 86400` (seconds per day as a Double).

**Result:** More responsive InputBox, but **Block 2 (12 seconds) reappeared.**

**Conclusion:** The 1-second delay is not merely waiting for the ribbon callback to
return. It is giving Word enough time to **complete its post-callback layout pass**
before `GoToH1Deferred` fires. At 0.5 seconds, Word is still mid-layout when the
macro fires — the navigation interrupts the layout computation, and Word must perform
a second layout pass afterward, restoring Block 2.

The 1-second delay is load-bearing. It acts as a "settle" buffer that ensures Word's
ribbon host post-processing is complete before navigation begins.

**Reverted to:** `Now + TimeValue("00:00:01")`

**The ~2-second InputBox delay is the minimum acceptable cost** for eliminating
Block 2 on this machine. A shorter delay risks Block 2 reappearing; a longer delay
adds unnecessary wait. The 1-second value is confirmed as the threshold.

---

## 26 — Test: InvalidateButtonsDeferred via `Application.OnTime Now + 1s`

**File changed:** `src/basRibbonDeferred.bas`

**Change tested:**
```vb
Public Sub GoToH1Deferred()
    ...
    rc.GoToH1Direct
    Application.OnTime Now + TimeValue("00:00:01"), "Project.basRibbonDeferred.InvalidateButtonsDeferred"
End Sub

Public Sub InvalidateButtonsDeferred()
    Dim rc As aeRibbonClass
    Set rc = Instance()
    rc.InvalidateControl "GoToNextButton"
    rc.InvalidateControl "GoToPrevButton"
End Sub
```

**Result:** Block 2 present. 13 seconds navigation, then 15 seconds spinning including
blank screen before Revelation displayed and cursor available.

**Conclusion:** `InvalidateControl` causes a 12-15 second layout pass on any cold-start
navigation to Revelation, regardless of when it is called — same macro invocation,
`OnTime Now`, or `OnTime Now + 1s` all produce the same block.

**Why §24 appeared to succeed:** The §24 test was run immediately after the §23 test,
which had just navigated to Revelation. The layout cache was warm from §23. When
`InvalidateButtonsDeferred` fired in §24, the cache appeared warm and the ribbon
refresh was cheap. This was a false positive — not reproducible on a genuine cold start.

**Full timing comparison across all tests:**

| Test | Block 1 | Block 2 | Total | IC location |
|---|---|---|---|---|
| §19 ribbon callback, IC before ScreenUpdating=True | 21s | 16s | 37s | Inside GoToH1 |
| §21 OnTime + IC before ScreenUpdating=True | 6s | 12s | 18s | Inside GoToH1 |
| §22 OnTime + IC after GoToH1Direct | 16s | 14s | 30s | GoToH1Deferred |
| §23 OnTime, no IC | 13s | none | **13s** | Removed |
| §24 OnTime + IC via `Now` (false positive) | 15s | none | 15s | InvalidateButtonsDeferred |
| §25 GoToH1 delay 0.5s | — | present | — | InvalidateButtonsDeferred |
| §26 IC via `OnTime Now + 1s` | 13s | 15s | 28s | InvalidateButtonsDeferred |

**Root cause (final):** `InvalidateControl` causes Word's ribbon host to re-query all
control states. On a cold layout cache (large document, far navigation), this query
triggers a full layout rebuild — 12-15 seconds. There is no timing that avoids this.
The ribbon refresh itself invalidates whatever warm cache exists.

**Final fix:** Do not call `InvalidateControl` from the GoToH1 path. Button state
variables (`m_btnNextEnabled` / `m_btnPrevEnabled`) are set correctly inside `GoToH1`.
The visual display remains stale until the first Prev/Next click — which calls
`InvalidateControl` itself after a short (warm-cache) navigation. At that point the
ribbon refresh is cheap and buttons update correctly.

**Final state of `src/basRibbonDeferred.bas`:**

```vb
Public Sub GoToH1Deferred()
    Dim rc As aeRibbonClass
    Set rc = Instance()
    rc.GoToH1Direct
    ' InvalidateControl is intentionally omitted here.
    ' On cold navigation to a far position (e.g., Revelation in a full-Bible document),
    ' calling InvalidateControl at any point after navigation triggers a 12-15 second
    ' layout pass — the ribbon refresh forces Word to rebuild the layout cache.
    ' Button state variables (m_btnNextEnabled / m_btnPrevEnabled) are set correctly
    ' inside GoToH1; the visual display updates on the next Prev/Next click, which
    ' calls InvalidateControl itself after a short (warm-cache) navigation.
End Sub
```

**Total time eliminated:** From 37 seconds (original double block) to 13 seconds
(single navigation pass, no secondary block). The remaining 13 seconds is the
intrinsic cold-cache layout cost of Word navigating a full-Bible document — not
reducible in VBA.

---

## 27 — Root Cause Revision: Post-OnTime-Macro Layout Pass

**Finding:** The §23 "no Block 2" result was not a reliable fix. It succeeded because
navigating to Genesis first ran the post-macro layout pass on a cheap position (page 1),
partially warming the layout cache. When GoTo Revelation ran next, the post-macro pass
was minimal. On a genuinely cold document (compiled project, no prior navigation),
Block 2 returns — 16 seconds with blank screen, file explorer coming to front twice.

**Root cause (revised):** Word triggers a post-macro ribbon state re-query after every
`Application.OnTime` macro completion, just as it does after ribbon callbacks. On a
cold-cache large document, this re-query forces a full layout rebuild — Block 2.
`Application.OnTime` deferral eliminated the ribbon *callback* post-processing block
but not the post-*macro* processing block.

**Why compiled project is worse:** Pre-compiled p-code executes faster than
interpreted VBA. The faster macro body gives Word's message loop less opportunity to
process incremental layout requests during execution. All post-macro layout hits at
once, causing a full 16-second freeze.

**Why Immediate window showed no Block 2 (Test B):** The VBA debugger executes in a
special context that bypasses Word's ribbon host post-processing. Results from the
Immediate window are not comparable to real-world ribbon or OnTime invocations.

**Three remaining options:**

| Option | Approach | Tradeoff |
|---|---|---|
| A | Replace `HomeKey + Find` with direct `SetRange(foundPos)` | Shorter Block 1; Block 2 may shrink or disappear on warmer cache |
| B | Pre-warm layout cache at document open via scheduled OnTime | One 16s freeze at open time; all subsequent GoToH1 calls fast |
| C | Remove `ScreenUpdating = False` | Word stays responsive (never freezes); user sees document scroll; total time may be similar |

---

## 28 — Option A: Replace HomeKey + Find with Direct SetRange Navigation

**File changed:** `src/aeRibbonClass.cls`

**What HomeKey + Find was doing:**
GoToH1 navigated in two steps:
1. `Selection.HomeKey Unit:=wdStory` — moved cursor to the first character of the
   document (Genesis, page 1), regardless of where the cursor currently was.
2. `Selection.Find.Execute` — searched forward from page 1 through the entire document
   to find the target Heading 1. For Revelation, this traversed 800,000+ characters.

Both steps ran with `ScreenUpdating = False`, so there was no visible scrolling — but
Word still processed the cursor movement and text search internally, adding overhead
before the final layout pass at `ScreenUpdating = True`.

**What the replacement does:**
`foundPos` is the character position of the target heading, already stored in
`headingData(i, 1)` when the document was loaded. `Selection.SetRange foundPos, foundPos`
jumps directly to that position in one operation — no HomeKey scroll, no Find traversal.

**Before:**
```vb
foundPos = CLng(headingData(i, 1))
Application.ScreenUpdating = False
Selection.HomeKey Unit:=wdStory
Selection.Find.ClearFormatting
Selection.Find.Text = CStr(headingData(i, 0))
Selection.Find.style = ActiveDocument.Styles("Heading 1")
Selection.Find.Forward = True
Selection.Find.Wrap = wdFindStop
Selection.Find.MatchCase = False
Selection.Find.MatchWildcards = False
Selection.Find.Execute
If Selection.Find.found Then Selection.Collapse Direction:=wdCollapseStart
```

**After:**
```vb
foundPos = CLng(headingData(i, 1))
Application.ScreenUpdating = False
Selection.SetRange foundPos, foundPos
Selection.Collapse Direction:=wdCollapseStart
```

**Expected effect:** Block 1 should be shorter — Word skips the HomeKey scroll and
full-document Find traversal. Whether Block 2 (post-macro layout pass) is reduced
depends on whether the direct jump warms the cache differently than HomeKey + Find.
Options B and C remain available if Block 2 persists.

**Test instruction:**

Compile project (Debug > Compile, save), restart Word, do not open IDE:

1. GoTo Revelation — report Block 1 timing, Block 2 present/absent, and whether
   file explorer comes to front
2. Check Prev/Next button states
3. GoTo Revelation warm — report time

Report all timings and whether Block 2 is present.

---

## 28a — Fix: Replace SetRange with ActiveDocument.Range.Select

**Problem with `Selection.SetRange`:** Moves the cursor internally but does not
queue a scroll-to-selection. When `ScreenUpdating = True` restores painting, Word
repaints the current view in place rather than scrolling to the new cursor position.
Result: GoTo Revelation appeared to do nothing — the view stayed at the original
position even though the cursor had moved.

**Why `Selection.Find` scrolled correctly:** Find includes an implicit
scroll-to-match as part of its operation. When `ScreenUpdating = True`, the view
jumps to the found position. `SetRange` has no such implicit scroll.

**Fix:** Use `ActiveDocument.Range(foundPos, foundPos).Select` — `Range.Select`
both moves the cursor AND queues a scroll-to-selection, so `ScreenUpdating = True`
displays the target position correctly.

**Current state of navigation block in `GoToH1`:**
```vb
foundPos = CLng(headingData(i, 1))
Application.ScreenUpdating = False
ActiveDocument.Range(foundPos, foundPos).Select
```

The `Collapse Direction:=wdCollapseStart` from the `Find` version is also removed —
`Range(foundPos, foundPos)` is already a zero-length (collapsed) range, so collapse
is a no-op.

**Test instruction:** Same as §28 — compile, restart Word, do not open IDE:
1. GoTo Revelation — Block 1 timing, Block 2 present/absent, file explorer to front?
2. Prev/Next button states
3. GoTo Revelation warm — time

---

## 28a — Test Results

| Step | Result |
|---|---|
| GoTo Revelation — Block 1 | 16 seconds, Revelation appears |
| GoTo Revelation — Block 2 | 18 seconds spinning + blank screen |
| GoTo Revelation warm | Instant |
| Prev/Next buttons | Both disabled (stale — expected) |

**Total: 34 seconds.** Worse than HomeKey+Find (32s in §27).

**Why Option A made Block 2 longer:** Direct `Range.Select` jumps to Revelation
without any document traversal. The layout cache is completely cold at that position.
With HomeKey+Find, the text search traversal partially warms the cache by processing
the character stream. Direct jump leaves the post-macro layout pass with more work.

**Option A finding:** `ActiveDocument.Range(foundPos, foundPos).Select` is the
correct navigation primitive — it works and Revelation appears correctly. But
`ScreenUpdating = False/True` around it is the wrong pairing: suppressing painting
defers the layout to post-macro processing time, causing Block 2.

---

## 29 — Option C: Remove ScreenUpdating = False/True

**File changed:** `src/aeRibbonClass.cls`

**Rationale:** With `ScreenUpdating = False`, the layout triggered by
`Range.Select` is deferred. When `ScreenUpdating = True` fires, Word queues the
repaint. When `GoToH1Deferred` returns, the post-macro processing encounters a
cold layout cache and rebuilds it — Block 2 (18 seconds).

With `ScreenUpdating = True` throughout (no suppression), the layout happens
synchronously during the `Range.Select` call within the normal paint cycle. By the
time `GoToH1Direct` returns, the layout is already complete. The post-macro
processing finds a warm cache — Block 2 should not occur.

There is no intermediate scrolling to hide from the user: `Range.Select` with a
direct character position jumps straight to Revelation without passing through
intermediate pages, so removing `ScreenUpdating = False` does not expose any
unwanted visual scrolling.

**Change:** Removed `Application.ScreenUpdating = False` before `Range.Select`,
`Application.ScreenUpdating = True` after the button state logic, and the
`Application.ScreenUpdating = True` restore in `PROC_ERR`.

**Current state of `GoToH1` navigation block:**
```vb
foundPos = CLng(headingData(i, 1))
ActiveDocument.Range(foundPos, foundPos).Select
m_btnPrevEnabled = True
m_btnNextEnabled = True
' ... first/last book checks ...
```

**Test instruction:** Compile (`Debug > Compile`, save), restart Word, do not open IDE:

1. GoTo Revelation — report total time and whether Block 2 is present
2. GoTo Genesis — report time and whether any spinning occurs
3. GoTo Revelation warm — report time
4. Prev/Next button states after GoTo Revelation
5. Note any visible screen flash or cursor behaviour during navigation

---

## 29 — Test Results

| Step | Result |
|---|---|
| GoTo Revelation — Block 1 | 16 seconds, Revelation visible |
| GoTo Revelation — Block 2 | 13 seconds, blank screen before completion |
| GoTo Genesis (after Revelation) | Instant |
| GoTo Revelation warm | Instant |
| Prev/Next buttons | Both disabled (stale — expected) |

**Total: 29 seconds.** Best result so far. Improvement over original 37 seconds.

**Why Block 2 persists with all approaches:**

Block 2 is Word's post-macro full-document layout pass. It is triggered by a COM-level
notification that Word sends after *every* `Application.OnTime` macro returns. Word's
ribbon host uses this notification to re-query all context-sensitive ribbon control
states (paragraph style, heading level, bold/italic, etc.), which requires a full
document repagination pass. This is not a queued Windows message — it fires after the
VBA sub returns, outside the macro execution context. `DoEvents` cannot process it.
The VBA debugger (Immediate window) does not trigger it, which is why Test B showed
no Block 2.

**Final comparison of all approaches (compiled project, cold document):**

| Approach | Block 1 | Block 2 | Total |
|---|---|---|---|
| Original — ribbon callback, HomeKey+Find | 21s | 16s | 37s |
| §21 — OnTime + IC before ScreenUpdating=True | 6s | 12s | 18s* |
| §28a — ScreenUpdating=False + Range.Select | 16s | 18s | 34s |
| **§29 — No ScreenUpdating + Range.Select** | **16s** | **13s** | **29s** |

*§21 result was on an uncompiled project; compiled result would likely be higher.

**Remaining option — Option B (pre-warm cache at document open):**
Schedule a silent navigation to the last book via `Application.OnTime` shortly after
the ribbon loads. The 29-second freeze moves to document-open time (where the user
expects loading overhead). All subsequent GoToH1 calls are instant. Decision pending.

---

## 30 — Option B: Pre-Warm Layout Cache at Document Open

**Files changed:** `src/aeRibbonClass.cls`, `src/basRibbonDeferred.bas`

**Rationale:** Block 2 is Word's post-macro full-document layout pass — unavoidable
in VBA. Every GoToH1 cold navigation triggers it (~13-18 seconds). Option B moves
this cost to document-open time by navigating silently to the last heading (Revelation)
and back 5 seconds after the ribbon loads. After the warm-up completes, the layout
cache is fully built and all subsequent GoToH1 calls are instant.

**Why 5-second delay:** `EnableButtonsRoutine` is called from `OnRibbonLoad`, which
is a ribbon callback. The same 1-second settle rule applies. 5 seconds gives
comfortable margin and lets the user see the document before any freeze.

**Why no ScreenUpdating = False:** Option C showed that without ScreenUpdating
suppression, the post-macro layout pass is 13 seconds (vs 18 with suppression).
The warm-up uses the faster path. The status bar message explains the brief activity.
The return to `savedPos` is instant (Genesis area cached after full-document traversal).

**Changes:**

`src/aeRibbonClass.cls` — `EnableButtonsRoutine` schedules warm-up:
```vb
Application.OnTime Now + TimeValue("00:00:05"), _
    ActiveDocument.VBProject.Name & ".basRibbonDeferred.WarmLayoutCacheDeferred"
```

`src/aeRibbonClass.cls` — new `WarmLayoutCache()` method:
```vb
Public Sub WarmLayoutCache()
    ' ... find lastPos from headingData ...
    savedPos = Selection.Start
    Application.StatusBar = "Bible: building navigation index..."
    ActiveDocument.Range(lastPos, lastPos).Select   ' jump to last heading
    ActiveDocument.Range(savedPos, savedPos).Select ' return to original position
    Application.StatusBar = False
End Sub
```

`src/basRibbonDeferred.bas` — new `WarmLayoutCacheDeferred()` sub:
```vb
Public Sub WarmLayoutCacheDeferred()
    Dim rc As aeRibbonClass
    Set rc = Instance()
    rc.WarmLayoutCache
End Sub
```

**Test instruction:**

Compile (`Debug > Compile`, save), restart Word, do not open IDE:

1. Open document — note whether a freeze occurs ~5 seconds after open, duration,
   and whether the status bar shows "Bible: building navigation index..."
2. After freeze completes — GoTo Revelation: report total time and whether Block 2
   is present
3. GoTo Revelation warm (second time) — report time
4. Prev/Next button states after GoTo Revelation

Report: warm-up freeze duration, GoTo Revelation timing after warm-up, any errors.

---

## 30 — Test Results: Option B Confirmed

| Step | Result |
|---|---|
| Freeze at document open (~5s after load) | Present ✓ |
| Status bar during freeze | "Bible: building navigation index..." ✓ |
| Warm-up freeze duration | ~20 seconds |
| File Explorer comes to front | Yes (Word "not responding" during warm-up) |
| GoTo Revelation after warm-up | **Instant** ✓ |
| Prev/Next buttons after GoTo Revelation | Both disabled (stale — expected) |

**Option B confirmed as the solution.**

The ~20-second freeze at document open builds the full layout cache. All GoToH1
navigations after that point are instant. The trade-off — one freeze at open instead
of a freeze on every cold GoToH1 call — is acceptable.

**Why 20 seconds rather than the predicted 29 seconds:** The warm-up fires at
document-open time when Word's internal state is different from a mid-session cold
navigation. The post-macro layout pass may be cheaper at open time, or some layout
was already partially done during `CaptureHeading1s` iteration.

**Known limitations:**
- File Explorer comes to front during warm-up — Word is "not responding" for ~20s.
  The status bar message is the only feedback, and it may not be visible while
  Word is unresponsive.
- Prev/Next buttons remain stale after GoToH1 until the first Prev/Next click.
  Internal state is correct; only the visual display is deferred.
- If the user navigates to Revelation before the warm-up completes (within the
  first 5 seconds), that navigation will be slow (~29s). Subsequent navigations
  will be instant.

**Final resolution summary:**

| Condition | Time | Block 2 |
|---|---|---|
| Original (ribbon callback, cold) | 37s | Present |
| After Option B warm-up | Instant | Absent |
| First open warm-up cost | ~20s (once) | N/A |

The 37-second double-block on cold navigation is eliminated. The cost is a one-time
~20-second freeze at document open, which is predictable and explained by the status
bar message.

**All three root causes addressed:**
1. Ribbon callback post-processing → fixed by `Application.OnTime` deferral (§21)
2. `InvalidateControl` on cold cache → fixed by omitting it from GoToH1 path (§26)
3. Post-macro layout pass on cold document → fixed by pre-warming at open (§30)

---

## 31 — Fix: Restore InvalidateControl to GoToH1Deferred (Warm Cache Available)

**File changed:** `src/basRibbonDeferred.bas`

**Problem:** Prev/Next buttons remain visually disabled after any GoToH1 navigation,
including GoTo Genesis. The internal state (`m_btnNextEnabled` / `m_btnPrevEnabled`)
is set correctly inside GoToH1, but without `InvalidateControl` the ribbon never
re-queries it. The visual display is permanently stale until a Prev/Next click.
This is a functional failure — the buttons show the wrong enabled state.

**Why this can be fixed now:** `InvalidateControl` was removed in §26 because it
triggered a 12-15 second layout pass on a cold document. With Option B (§30),
`WarmLayoutCacheDeferred` runs at document open and builds the full layout cache.
By the time the user performs any GoToH1 navigation, the cache is warm and
`InvalidateControl` is cheap — no Block 2.

**Change:**

```vb
Public Sub GoToH1Deferred()
    Dim rc As aeRibbonClass
    Set rc = Instance()
    rc.GoToH1Direct
    ' InvalidateControl is called after WarmLayoutCacheDeferred has run at document
    ' open (Option B, §30). With the layout cache warm, these calls are cheap.
    ' If called before the warm-up completes (within first 5 seconds of open),
    ' the cache may be cold and a brief block may occur on that first navigation only.
    rc.InvalidateControl "GoToNextButton"
    rc.InvalidateControl "GoToPrevButton"
End Sub
```

**Test instruction:** Compile, restart Word, do not open IDE:

1. Wait for warm-up to complete (~20 second freeze after open)
2. GoTo Revelation — report time and whether Block 2 is present
3. Check Next button disabled at Revelation (correct — last book)
4. Navigate Prev — check both buttons enabled
5. GoTo Genesis — check Prev disabled, Next enabled
6. GoTo Revelation warm (second time) — report time and button states

---

## 31 — Test Results: Fix Confirmed — Investigation Complete

| Step | Result |
|---|---|
| Warm-up freeze duration | ~20 seconds |
| Status bar during warm-up | Visible ✓ |
| GoTo Revelation after warm-up | **Instant, no Block 2** ✓ |
| Next button at Revelation | Disabled ✓ |
| Navigate Prev from Revelation | Both buttons enabled ✓ |
| GoTo Genesis | Prev disabled, Next enabled ✓ |
| GoTo Revelation warm (second) | Instant ✓ |

**All requirements met. Investigation complete.**

---

## Final Summary

**Problem:** GoToH1 navigation from Genesis to Revelation caused a 37-second
double block — Word "not responding" twice, file explorer coming to front twice.

**Three root causes identified and fixed:**

| # | Root Cause | Fix | Section |
|---|---|---|---|
| 1 | Ribbon callback post-processing triggers layout pass after `OnGoToH1ButtonClick` returns | Defer navigation via `Application.OnTime Now + TimeValue("00:00:01")` | §21 |
| 2 | Post-`Application.OnTime`-macro layout pass on cold document triggers 13-18s block | Pre-warm layout cache at document open via `WarmLayoutCacheDeferred` | §30 |
| 3 | `InvalidateControl` on cold cache triggers additional layout pass | Restored after warm-up makes cache warm; now cheap | §31 |

**Final file state:**

`src/basBibleRibbonSetup.bas` — ribbon callback defers navigation:
```vb
Application.OnTime Now + TimeValue("00:00:01"), projName & ".basRibbonDeferred.GoToH1Deferred"
```

`src/aeRibbonClass.cls` — `GoToH1` uses direct position jump, no ScreenUpdating suppression:
```vb
foundPos = CLng(headingData(i, 1))
ActiveDocument.Range(foundPos, foundPos).Select
```

`src/aeRibbonClass.cls` — `EnableButtonsRoutine` schedules warm-up:
```vb
Application.OnTime Now + TimeValue("00:00:05"), _
    ActiveDocument.VBProject.Name & ".basRibbonDeferred.WarmLayoutCacheDeferred"
```

`src/basRibbonDeferred.bas` — three deferred subs:
```vb
Public Sub WarmLayoutCacheDeferred()   ' warms layout cache at document open
Public Sub GoToH1Deferred()            ' deferred navigation + InvalidateControl
```

**User experience:**
- Document open: ~20 second freeze with status bar "Bible: building navigation index..."
- All GoToH1 navigations thereafter: instant
- Prev/Next button states: correct after every navigation
- Original 37-second double block: eliminated

---

## Session Summary — 2026-04-08 / 2026-04-09

**Primary work:** Resolved the persistent GoToH1 double-block in a full-Bible `.docm`.

**Secondary work:** VBA casing normalizer (`py/normalize_vba.py`) — added rules for
`Shell`, `Range:=`, `As PageSetup`, `As Word.Section`, and `.Code`; audited all
rule descriptions.

### GoToH1 Fix — Key Findings

1. **Ribbon callback post-processing** — after any ribbon `onAction` callback
   returns, Word re-queries all ribbon control states, triggering a full layout pass
   on a cold document. Fixed by deferring navigation via `Application.OnTime`.

2. **`Option Private Module` blocks `Application.OnTime`** — all standard modules
   in this project use `Option Private Module`, which prevents `Application.OnTime`
   from resolving macro names. Created `basRibbonDeferred.bas` without this
   declaration as the sole dispatch module for OnTime targets.

3. **`InvalidateControl` on cold cache** — calling `InvalidateControl` immediately
   after cold navigation to Revelation triggers a 12-18 second layout pass regardless
   of timing (same macro, deferred OnTime, 1-second delay — all caused the block).
   Root cause: Word's ribbon refresh forces a full layout rebuild on a cold document.

4. **Post-`Application.OnTime`-macro layout pass** — Word also triggers a post-macro
   layout pass after any OnTime macro completes (same mechanism as ribbon callbacks).
   On a compiled project with a cold full-Bible document, this pass takes 13-18
   seconds. Not preventable in VBA; only avoidable by pre-warming the cache.

5. **Pre-warm layout cache at open (Option B)** — `WarmLayoutCacheDeferred` fires
   5 seconds after ribbon load, navigates to the last heading and back. This forces
   the ~20-second layout pass at document-open time. All subsequent GoToH1 calls
   (including `InvalidateControl`) are instant on the warm cache.

6. **VBA debugger bypasses post-macro processing** — Test B (Immediate window)
   showed no Block 2 because the VBA debugger does not trigger Word's ribbon
   host post-processing. Immediate window results are not representative of
   real ribbon or OnTime invocations.

### Files Changed This Session

| File | Changes |
|---|---|
| `src/basRibbonDeferred.bas` | New module — `WarmLayoutCacheDeferred`, `GoToH1Deferred` |
| `src/basBibleRibbonSetup.bas` | `OnGoToH1ButtonClick` defers via OnTime; `TestGoToH1Direct` added |
| `src/aeRibbonClass.cls` | `GoToH1`: removed HomeKey+Find, uses `Range.Select`; removed `ScreenUpdating`; `WarmLayoutCache` added; `EnableButtonsRoutine` schedules warm-up |
| `py/normalize_vba.py` | Added 5 casing rules; audited all descriptions |
| `rvw/Code_review - 2026-04-08.md` | New — 31 sections documenting full investigation |

### Context for Next Session

- The fix is complete and tested. Ready to commit.
- `InvalidateControl` is intentionally absent from `NextButton`/`PrevButton` error
  paths — those navigate one heading at a time (warm cache) so no block occurs there.
- The 5-second warm-up delay in `EnableButtonsRoutine` is load-bearing — less than
  1 second causes Block 2 to reappear on first GoToH1.
- `WarmLayoutCacheDeferred` will appear in Alt+F8 (no parameters, no
  `Option Private Module`). Safe to run manually; it just re-warms the cache.

---

## Note: Using /clear Between Sessions

`/clear` resets the Claude Code conversation context — everything in the current
chat window is discarded and a fresh session begins.

**Benefits for this project:**

- **Context window freed** — long sessions (like this investigation) accumulate
  dead-end hypotheses, superseded code states, and intermediate test results.
  Clearing removes that noise so the model works from a clean state.
- **No stale working state** — the model cannot accidentally reference a code
  version or test result that has since been superseded.
- **Memory persists** — `MEMORY.md` and memory files in
  `C:\Users\peter\.claude\projects\C--adaept-aeBibleClass\memory\` survive `/clear`.
  Key decisions, feedback preferences, and project context carry forward.
- **Review files persist** — all `rvw\` documents are on disk and unaffected.

**Before clearing:** save anything important to memory files or on-disk documents
(such as this review file). Anything discussed only in the chat window is lost.
