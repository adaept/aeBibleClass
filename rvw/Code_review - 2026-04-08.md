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

### Changes

**`src/basBibleRibbonSetup.bas`** — `OnGoToH1ButtonClick` now schedules navigation
instead of executing it. `GoToH1Deferred` is the scheduled target:

```vb
Public Sub OnGoToH1ButtonClick(control As IRibbonControl)
    Application.OnTime Now, "GoToH1Deferred"
End Sub

Public Sub GoToH1Deferred()
    Dim rc As aeRibbonClass
    Set rc = Instance()
    rc.GoToH1Direct
End Sub
```

The ribbon callback returns in microseconds. `GoToH1Deferred` fires on the next
Windows message loop cycle — after the ribbon host has completed its post-callback
processing — and runs `GoToH1` in a clean context with no ribbon host waiting.

**`src/aeRibbonClass.cls`** — `GoToH1Direct` wrapper (added in Section 20) is now
the permanent entry point for scheduled navigation. No change needed to the class.

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
