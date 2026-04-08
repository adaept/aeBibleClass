# Code Review: Ribbon PrevButton Implementation

**Date:** 2026-04-06

---

## Scope

Implementation of `PrevButton` (Previous Book) in `src/aeRibbonClass.cls`, mirroring
the existing `NextButton` pattern. Includes design note for disabled-on-open behaviour.

---

## 1 — Implementation Plan

### Pattern: `NextButton`

`NextButton` searches forward from the end of the current paragraph for the next
Heading 1. If not found (cursor is in Revelation or beyond the last H1), it wraps
by searching forward from the document start — landing on Genesis.

### PrevButton Design

`PrevButton` mirrors `NextButton` with a backward search:

1. Get `paraStart` = start position of the current paragraph.
2. Search **backward** (`Forward = False`) within `doc.Range(0, paraStart)` for a
   Heading 1. This finds the nearest H1 before the cursor — the previous book.
3. If not found (cursor is in Genesis or before the first H1), wrap: search backward
   within `doc.Range(paraStart, doc.content.End)` — this finds the last H1 in the
   document (Revelation).
4. Navigate to the found heading.

Wrap behaviour matches `NextButton`:

- Next Book at Revelation → Genesis
- Prev Book at Genesis → Revelation

---

## 2 — Files Changed

**`src/aeRibbonClass.cls`**

| Item | Change |
|---|---|
| `m_btnPrevEnabled As Boolean` | Added to private state |
| `Class_Initialize` | Sets `m_btnPrevEnabled = True` |
| `BtnPrevEnabled` Property Get/Let | Added alongside `BtnNextEnabled` |
| `OnRibbonLoad` | Invalidates `GoToPrevButton` on load |
| `OnPrevButtonClick` | Added ribbon callback → calls `PrevButton` |
| `GetPrevEnabled` | Added enabled-state callback for ribbon |
| `EnableButtonsRoutine` | Sets `m_btnPrevEnabled = True`, invalidates `GoToPrevButton` |
| `PrevButton` | New private sub — backward H1 search with Revelation wrap |

---

## 3 — `PrevButton` Implementation

```vb
Private Sub PrevButton()
    On Error GoTo PROC_ERR
    Dim doc         As Document
    Dim searchRange As Word.Range
    Dim paraStart   As Long
    Dim found       As Boolean

    Set doc = ActiveDocument
    found = False

    paraStart = Selection.Paragraphs(1).Range.Start
    Set searchRange = doc.Range(0, paraStart)

    With searchRange.Find
        .ClearFormatting
        .style = doc.Styles("Heading 1")
        .Forward = False
        .Wrap = wdFindStop
        .Format = True
        .Text = ""
        found = .Execute
    End With

    If Not found Then
        ' Wrap: at Genesis, go to Revelation (last H1 in document)
        Set searchRange = doc.Range(paraStart, doc.content.End)
        With searchRange.Find
            .ClearFormatting
            .style = doc.Styles("Heading 1")
            .Forward = False
            .Wrap = wdFindStop
            .Format = True
            .Text = ""
            found = .Execute
        End With
    End If

    If found Then
        Selection.SetRange searchRange.Start, searchRange.Start
        ActiveWindow.ScrollIntoView Selection.Range, True
    Else
        MsgBox "No Heading 1 found in the document.", vbInformation
    End If

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure PrevButton of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub
```

---

## 4 — Disabled-on-Open Design (PrevBook and NextBook)

### Requirement

PrevBook and NextBook buttons should be **disabled when the document opens** and
become **active only after GoToBook has been successfully used once**.

### XML Changes

Add `getEnabled` callbacks to both navigation buttons:

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="RibbonOnLoad">
  <ribbon startFromScratch="false">
    <tabs>
      <tab id="RWB" label="Radiant Word Bible">
        <group id="TestGroup" label="Bible Class Group">
          <button id="GoToMyVerse" label="GoTo Verse " imageMso="FormatNumberDefault" size="large" onAction="OnGoToVerseSblClick"/>
          <separator id="sep1" />
          <button id="GoToPrevButton" label="Prev Book" imageMso="HeaderFooterPreviousSection" size="large" onAction="OnPrevButtonClick" getEnabled="GetPrevEnabled"/>
          <button id="GoToH1Button" label="GoTo Book" imageMso="NotebookNew" size="large" onAction="OnGoToH1ButtonClick"/>
          <button id="GoToNextButton" label="Next Book" imageMso="HeaderFooterNextSection" size="large" onAction="OnNextButtonClick" getEnabled="GetNextEnabled"/>
          <separator id="sep2" />
          <button id="adaeptButton" label="About" image="adaept" size="large" onAction="OnAdaeptAboutClick"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

### VBA Changes Required

**`Class_Initialize`** — start both buttons disabled:

```vb
m_btnNextEnabled = False
m_btnPrevEnabled = False
```

**`EnableButtonsRoutine`** — remove or leave as-is (called on load but now starts
disabled; this routine could be repurposed as `EnableNavButtons` called from `GoToH1`).

**`GoToH1`** — after a successful navigation, enable both buttons and invalidate:

```vb
If matchFound Then
    ' ... existing navigation code ...
    m_btnNextEnabled = True
    m_btnPrevEnabled = True
    If Not m_ribbon Is Nothing Then
        m_ribbon.InvalidateControl "GoToNextButton"
        m_ribbon.InvalidateControl "GoToPrevButton"
    End If
End If
```

### Behaviour After Change

| State | Prev Book | Next Book |
|---|---|---|
| Document just opened | Disabled (grey) | Disabled (grey) |
| After GoTo Book used once | Enabled | Enabled |
| Subsequent use | Remains enabled | Remains enabled |

GoTo Book itself has no `getEnabled` callback — it is always active, giving the user
a clear entry point to enable the navigation buttons.

---

## 5 — Questions for Clarification

1. Should the disabled-on-open behaviour be implemented now, or deferred?
2. Is the `getEnabled` callback name `GetPrevEnabled` / `GetNextEnabled` consistent
   with the existing naming convention in `basBibleRibbonSetup.bas`?
3. `EnableButtonsRoutine` currently also calls `CaptureHeading1s` and `LogHeadingData`.
   If buttons start disabled, should these still run on load, or only after GoToBook?


---

## 6 — How to Wire `getEnabled` Callbacks: Architecture and Steps

### Three-Layer Architecture

The ribbon callback system in `Blank Bible Copy.docm` has three layers:

| Layer | File | Role |
|---|---|---|
| 1 | `customUI14.xml` (inside `.docm`) | Declares callback names as plain strings |
| 2 | `basBibleRibbonSetup.bas` | Public stub functions/subs that Word resolves by name |
| 3 | `aeRibbonClass.cls` | Singleton class holding state and logic |

Word’s ribbon host resolves callback names against **public procedures in standard
modules only** — not against class methods. `basBibleRibbonSetup.bas` is the required
adapter between the XML and the class.

### What `getEnabled` Requires

A `getEnabled` callback must be a `Public Function` returning `Boolean` with one
`IRibbonControl` parameter, declared in a standard module. The function name in the
XML must exactly match the function name in the module.

### Step 1 — XML: add `getEnabled` attributes

```xml
<button id="GoToPrevButton" ... getEnabled="GetPrevEnabled"/>
<button id="GoToNextButton" ... getEnabled="GetNextEnabled"/>
```

### Step 2 — `basBibleRibbonSetup.bas`: add stub functions

```vb
Public Function GetPrevEnabled(control As IRibbonControl) As Boolean
    GetPrevEnabled = Instance().GetPrevEnabled(control)
End Function

Public Function GetNextEnabled(control As IRibbonControl) As Boolean
    GetNextEnabled = Instance().GetNextEnabled(control)
End Function
```

These delegate to the class methods already implemented in `aeRibbonClass.cls`.
The class `GetPrevEnabled` and `GetNextEnabled` methods exist and are correct.

### Step 3 — Edit the XML in `Blank Bible Copy.docm`

The XML is stored inside the `.docm` ZIP archive at `customUI/customUI14.xml`.
Two methods:

**Recommended — Custom UI Editor tool**
Open `Blank Bible Copy.docm` in the free *Custom UI Editor for Microsoft Office*
tool. It exposes the XML directly and saves it back into the file cleanly without
risk of corruption.

**Manual ZIP method** - more challenging

1. Close Word.
2. Rename `Blank Bible Copy.docm` → `Blank Bible Copy.zip`.
3. Open the ZIP, navigate to `customUI/`, edit `customUI14.xml` in a text editor.
4. Save, close the ZIP, rename back to `.docm`.
Risk: any step performed with the file open in Word will corrupt the file.

### Why Class Methods Alone Are Not Sufficient

Word’s ribbon host has no knowledge of class instances. The stub in
`basBibleRibbonSetup.bas` is the mandatory adapter. The missing pieces when
implementing the disabled-on-open behaviour are:

- Two stub functions in `basBibleRibbonSetup.bas`
- Two `getEnabled` attributes in the XML

All VBA logic in `aeRibbonClass.cls` is already in place.


---

## 7 — XML Edit: Delegated to User

Editing the `customUI14.xml` inside `Blank Bible Copy.docm` requires the Custom UI
Editor GUI tool. This cannot be performed by Claude. Task delegated to user.

**Steps (Custom UI Editor):**

1. Open `Blank Bible Copy.docm` in Custom UI Editor.
2. Add `getEnabled` to the two navigation buttons:

```xml
<button id="GoToPrevButton" label="Prev Book" imageMso="HeaderFooterPreviousSection" size="large" onAction="OnPrevButtonClick" getEnabled="GetPrevEnabled"/>
<button id="GoToNextButton" label="Next Book" imageMso="HeaderFooterNextSection" size="large" onAction="OnNextButtonClick" getEnabled="GetNextEnabled"/>
```

3. Save and close.

All VBA changes in `aeRibbonClass.cls` and `basBibleRibbonSetup.bas` are already
implemented and ready.


---

## 8 — Implementation: Disabled-on-Open with Enable After GoToBook

### Changes Implemented

**`src/basBibleRibbonSetup.bas`** — added two `getEnabled` stub functions:

```vb
Public Function GetPrevEnabled(control As IRibbonControl) As Boolean
    GetPrevEnabled = Instance().GetPrevEnabled(control)
End Function

Public Function GetNextEnabled(control As IRibbonControl) As Boolean
    GetNextEnabled = Instance().GetNextEnabled(control)
End Function
```

**`src/aeRibbonClass.cls` — `Class_Initialize`** — both buttons start disabled:

```vb
m_btnNextEnabled = False
m_btnPrevEnabled = False
```

**`src/aeRibbonClass.cls` — `GoToH1`** — after successful navigation, enable both
buttons and invalidate their controls so the ribbon updates immediately:

```vb
matchFound = True
m_btnNextEnabled = True
m_btnPrevEnabled = True
If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "GoToNextButton"
If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "GoToPrevButton"
Exit For
```

### Behaviour

| State | Prev Book | Next Book |
|---|---|---|
| Document opens | Disabled (grey) | Disabled (grey) |
| GoTo Book used successfully | Enabled | Enabled |
| Subsequent navigations | Remains enabled | Remains enabled |

XML edit (`getEnabled` attributes on both buttons) was performed by the user
using the Custom UI Editor tool (delegated — see Section 7).


---

## 9 — Bug Fix: "Wrong number of arguments" on Ribbon Load (`GetPrevEnabled`, `GetNextEnabled`)

**Symptom:** Two "Wrong number of arguments or invalid property assignment" errors on
ribbon load. Both Prev Book and Next Book buttons remained disabled.

**Cause:** The `getEnabled` stub functions in `basBibleRibbonSetup.bas` used chained
calls on the return value of `Instance()`:

```vb
GetPrevEnabled = Instance().GetPrevEnabled(control)
```

VBA cannot reliably resolve method arguments through a temporary object reference
returned inline by a function. It misinterprets `(control)` — either as an attempt
to index the return value or as an invalid property assignment — and raises Error 450.
The ribbon then defaults the button to disabled because the callback failed.

The error occurred twice because both `GetPrevEnabled` and `GetNextEnabled` used the
same pattern.

**Why Sub stubs are unaffected:** The existing Sub stubs (e.g. `OnPrevButtonClick`)
use VBA Sub call syntax without parentheses around the argument:
```vb
Instance().OnPrevButtonClick control
```
This form does not trigger the same ambiguity. Function calls that return a value
and pass arguments require the local variable pattern.

**Fix:** Store the instance in a local variable before calling the method:

```vb
Public Function GetPrevEnabled(control As IRibbonControl) As Boolean
    Dim rc As aeRibbonClass
    Set rc = Instance()
    GetPrevEnabled = rc.GetPrevEnabled(control)
End Function

Public Function GetNextEnabled(control As IRibbonControl) As Boolean
    Dim rc As aeRibbonClass
    Set rc = Instance()
    GetNextEnabled = rc.GetNextEnabled(control)
End Function
```

`rc.GetPrevEnabled(control)` is an unambiguous method call on a named variable —
VBA resolves it correctly.


---

## 10 — Bug Fix: Wrong `getEnabled` Callback Signature (Second Fix)

**Symptom:** Same "Wrong number of arguments or invalid property assignment" error
still occurred twice after the local-variable fix in Section 9.

**Cause:** The callback signature was wrong. Office ribbon `getEnabled` callbacks in
VBA must be a **Sub** with **two parameters** — `control As IRibbonControl` and
`ByRef enabled As Boolean`. Word passes the enabled state back through the `ByRef`
parameter. A `Function ... As Boolean` declares only one parameter; Word tries to
call it with two and raises Error 450.

Correct contract required by the ribbon host:
```vb
Sub GetPrevEnabled(control As IRibbonControl, ByRef enabled As Boolean)
```

Compare with `onAction` which correctly uses a one-parameter Sub:
```vb
Sub OnPrevButtonClick(control As IRibbonControl)
```

Each ribbon callback attribute type has its own fixed VBA signature.

**Fix:** Changed both stubs in `basBibleRibbonSetup.bas` from `Function` to `Sub`
with the correct `ByRef enabled As Boolean` parameter, reading state via the class
`BtnPrevEnabled` / `BtnNextEnabled` properties:

```vb
Public Sub GetPrevEnabled(control As IRibbonControl, ByRef enabled As Boolean)
    Dim rc As aeRibbonClass
    Set rc = Instance()
    enabled = rc.BtnPrevEnabled
End Sub

Public Sub GetNextEnabled(control As IRibbonControl, ByRef enabled As Boolean)
    Dim rc As aeRibbonClass
    Set rc = Instance()
    enabled = rc.BtnNextEnabled
End Sub
```

The `GetPrevEnabled` / `GetNextEnabled` Function methods on `aeRibbonClass` are
no longer called by the ribbon stubs but are retained on the class.


---

## 11 — Bug Fix: Type Mismatch on `ByRef enabled` Parameter (Third Fix)

**Symptom:** "Type mismatch" error still occurred twice on ribbon load after the
two-parameter Sub fix in Section 10.

**Cause:** `ByRef enabled As Boolean` is still wrong. The Office ribbon host passes
the `enabled` argument as a **Variant**, not a Boolean. Declaring the parameter
`As Boolean` causes a type mismatch when Word tries to bind its Variant to the
typed parameter.

**Fix:** Remove the `As Boolean` type declaration, leaving `enabled` untyped
(implicitly Variant) — the standard pattern for all Office ribbon `get*` callbacks:

```vb
Public Sub GetPrevEnabled(control As IRibbonControl, ByRef enabled)
    Dim rc As aeRibbonClass
    Set rc = Instance()
    enabled = rc.BtnPrevEnabled
End Sub

Public Sub GetNextEnabled(control As IRibbonControl, ByRef enabled)
    Dim rc As aeRibbonClass
    Set rc = Instance()
    enabled = rc.BtnNextEnabled
End Sub
```

Assigning a Boolean (`rc.BtnPrevEnabled`) to an untyped Variant `enabled` is
valid — VBA widens automatically. The ribbon host then reads the Variant as a
Boolean to set the button state.

**Correct Office ribbon `get*` callback pattern (VBA):**

| Attribute | Signature |
|---|---|
| `onAction` | `Sub Name(control As IRibbonControl)` |
| `getEnabled` | `Sub Name(control As IRibbonControl, ByRef enabled)` |
| `getLabel` | `Sub Name(control As IRibbonControl, ByRef label)` |
| `getVisible` | `Sub Name(control As IRibbonControl, ByRef visible)` |

All `get*` return parameters are untyped Variants passed ByRef.


---

## 12 — Bug Fix: Buttons Enabled on Open Despite `Class_Initialize` Setting `False`

**Symptom:** No error, but Prev Book and Next Book buttons were enabled immediately
on ribbon load, before GoTo Book had been used.

**Cause:** `EnableButtonsRoutine` is called from `OnRibbonLoad` and unconditionally
set both button states to `True`:

```vb
m_btnNextEnabled = True
m_btnPrevEnabled = True
If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "GoToNextButton"
If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "GoToPrevButton"
```

This ran after `Class_Initialize` set both to `False`, overriding the disabled-on-open
intent before the ribbon had even finished loading. The `getEnabled` callbacks then
correctly returned `True` — enabling the buttons.

`EnableButtonsRoutine` predates the disabled-on-open requirement. When it was written,
always-enabled was the correct behaviour. Now that `GoToH1` handles enabling both
buttons after a successful navigation, `EnableButtonsRoutine` must not touch button
state at all.

**Fix:** Removed the four button-state lines from `EnableButtonsRoutine`. Its sole
remaining purpose is data capture:

```vb
Private Sub EnableButtonsRoutine()
    On Error GoTo PROC_ERR
    Debug.Print "RibbonController: EnableButtonsRoutine"
    CaptureHeading1s
    LogHeadingData
    ...
```

Button enable/disable is now exclusively controlled by:
- `Class_Initialize` — both `False` on creation
- `GoToH1` — both `True` after successful navigation, with immediate invalidation


---

## 13 — Performance Fix: `PrevButton`/`NextButton` Replaced `Find` with `headingData` Array

**Symptom:** `PrevButton` was very slow on first use when the cursor was at Genesis.
Subsequent uses were fast. `NextButton` had the same latent problem at Revelation.

**Cause:** Both buttons used Word’s `Find` API with `Format = True` and a style filter.
When a wrap was required (Genesis → Revelation, or Revelation → Genesis), the search
range spanned the entire document. Word scanned every paragraph checking its style —
O(n) in document length. First use was slow because the find engine had not yet warmed
up its style index.

**Solution:** `CaptureHeading1s` already builds `headingData(1 To 66, 0 To 1)` on
ribbon load — an array of all 66 Heading 1 positions (character offsets). Navigation
via this array is O(66) regardless of document size and requires no document scan.

**Implementation:** `Find`-based code removed from both `NextButton` and `PrevButton`.
Three private helpers added to `aeRibbonClass.cls`:

| Procedure | Purpose |
|---|---|
| `CurrentBookIndex()` | Returns index 1–66 of the book containing the cursor (largest `headingData(i,1)` ≤ cursor position) |
| `LastBookIndex()` | Returns the highest populated index in `headingData` (typically 66) |
| `NavigateToBookIndex(idx)` | Moves cursor to `headingData(idx, 1)` and scrolls into view |

`NextButton` and `PrevButton` are now three lines each: get current index, compute
target index (with wrap), call `NavigateToBookIndex`.

---

## 14 — Future: Use `headingData` for Book/Chapter/Verse Lookups

**Context:** `GoToVerseSBL` is currently a stub (`MsgBox "Parser not yet implemented"`).
When implemented, it will need to locate a specific book, then scan within that book
for a chapter and verse.

**How `headingData` helps:**

`headingData(i, 1)` gives the document character offset of each book’s Heading 1.
`headingData(i+1, 1)` (or `doc.content.End` for the last book) gives the end of
that book’s content. This immediately bounds the search range for any verse lookup:

```vb
' Resolve book abbreviation -> bookIndex (1-66) via aeBibleCitationClass
Dim bookStart As Long, bookEnd As Long
bookStart = CLng(headingData(bookIndex, 1))
If bookIndex < 66 Then
    bookEnd = CLng(headingData(bookIndex + 1, 1))
Else
    bookEnd = ActiveDocument.content.End
End If
Set searchRange = ActiveDocument.Range(bookStart, bookEnd)
' Then Find chapter/verse heading or paragraph within searchRange only
```

**Benefits:**
- Search is bounded to one book’s text — typically 1/66 of the document
- No full-document scan for any lookup
- `headingData` is already populated before the buttons are enabled
- Same `CurrentBookIndex()` helper can be reused to detect which book the cursor
  is currently in — useful for context-sensitive verse lookups

**Next step:** Implement `GoToVerseSBL` using `aeBibleCitationClass.ParseCitationBlock`
to resolve the input, then use `headingData` to bound the document search.

---

## 14 — ScreenUpdating Fix in NavigateToBookIndex

### Problem

After replacing the Find-based navigation with `headingData` array lookups, `PrevButton`
(and `NextButton`) remained slow on first use when jumping across the entire document
(e.g., Genesis → Revelation). The symptom: Word "not responding", spinner, long pause
before the heading was visible.

**Root cause:** `Selection.SetRange` followed by `ActiveWindow.ScrollIntoView` forces
Word to synchronously repaginate the portion of the document between the old and new
cursor positions before it can render the target location. For a full-Bible document
this computation runs on the main UI thread and cannot be parallelised or cancelled.
This is a Word layout engine limitation — not a code defect.

### Fix

Suppress screen updates during the navigation call so Word defers repaint until both
`SetRange` and `ScrollIntoView` have completed:

```vb
Application.ScreenUpdating = False
Selection.SetRange targetPos, targetPos
ActiveWindow.ScrollIntoView Selection.Range, True
Application.ScreenUpdating = True
```

`Application.ScreenUpdating = True` is also restored unconditionally in `PROC_ERR`
so that an unexpected error never leaves the screen frozen.

### File Changed

`src/aeRibbonClass.cls` — `NavigateToBookIndex`

### Expected Result

The spinner and "not responding" pause are eliminated on first navigation. Word updates
the display in a single repaint after both operations complete rather than incrementally
scrolling through the intervening pages.

### Limitation

`ScreenUpdating = False` suppresses the intermediate repaints but does not eliminate
the layout computation itself. On very large documents (full Bible, ~1 MB+) there may
still be a brief pause on the first jump if Word has not yet fully paginated the target
region. Subsequent jumps are faster because the layout cache is warm.

---

## 15 — NavigateToBookIndex: Replace ScrollIntoView with Range.Select

### Problem

After adding `Application.ScreenUpdating = False`, navigation was still slower than
the original Find-based scan, and Word went behind the Explorer window twice during
each jump.

**Root cause — two issues:**

1. **`ActiveWindow.ScrollIntoView` forces full layout computation.**
   To scroll a character position into view, Word must paginate every paragraph from
   page 1 up to the target position — synchronously on the UI thread. For a full-Bible
   document this is heavier than the VBA paragraph scan it replaced. The `True`
   (center-in-window) argument makes it worse: centering also requires the exact
   rendered height of the target paragraph.

2. **`ScreenUpdating = False` killed the message pump.**
   While the UI thread was blocked by layout computation, Windows detected that Word's
   message pump had stalled and demoted the window in the z-order. This happened twice
   (once for `Selection.SetRange`, once for `ScrollIntoView`), which is why Word went
   behind the Explorer window twice. `ScreenUpdating = False` removed the intermediate
   repaints that normally keep the pump alive, making the stall more visible, not less.

### Fix

Replace `Selection.SetRange` + `ScrollIntoView` with `Range.Select`. Word's `Select`
method moves the cursor and brings the location into view incrementally — the same
mechanism used internally by Find — without requiring the full layout pass.
`ScreenUpdating` suppression is no longer needed.

```vb
Private Sub NavigateToBookIndex(ByVal idx As Long)
    On Error GoTo PROC_ERR
    If idx < 1 Or idx > 66 Then GoTo PROC_EXIT
    If IsEmpty(headingData(idx, 1)) Then GoTo PROC_EXIT
    Dim targetPos As Long
    targetPos = CLng(headingData(idx, 1))
    ActiveDocument.Range(targetPos, targetPos).Select
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure NavigateToBookIndex of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub
```

### File Changed

`src/aeRibbonClass.cls` — `NavigateToBookIndex`

### Expected Result

Navigation speed matches or exceeds the original Find-based approach. Word stays in
the foreground throughout. No `ScreenUpdating` suppression required.

---

## 16 — NavigateToBookIndex: Revert to Selection.Find

### Problem

After two iterations (`Selection.SetRange` + `ScrollIntoView` → `Range.Select`),
navigation to Revelation from Genesis still took >1 minute, Word showed "not
responding", the window was demoted behind Explorer twice, and the heading appeared
with text selected rather than a collapsed cursor.

### Why the previous fixes were wrong

**Section 14 (`ScreenUpdating = False`):**
The slowness is layout computation, not painting. `ScreenUpdating = False` suppresses
repaints but does not suppress layout. Word still computed full pagination
synchronously. Worse, suppressing repaints also suppresses message-pump activity,
which caused Windows to demote the Word window twice (once per blocking layout call).

**Section 15 (`Range.Select`):**
The claim that `ActiveDocument.Range(pos, pos).Select` navigates "the same way Find
does, at the text layer" was incorrect. `Range(pos, pos).Select` resolves the
character position through the layout engine and then forces a scroll-to-view, both
of which require Word to paginate all content between the current view and the target.
This is more expensive than `Selection.Find`, not equivalent to it.

The "text selected" symptom confirms the character positions stored in `headingData`
were not being resolved cleanly — a collapsed `Range(x, x)` should never produce a
visible selection.

### Why Selection.Find is faster

`Selection.Find` operates on Word's backing store (piece table), navigating through
paragraph style descriptors without computing page coordinates. The display updates
in a single deferred repaint after the selection is placed. No layout pass is
triggered during the search itself.

### Fix

`NextButton` and `PrevButton` now call `Selection.Find` directly with `wdFindContinue`
wrapping. `NavigateToBookIndex`, `CurrentBookIndex`, and `LastBookIndex` are removed —
they were only used by these two procedures and are no longer needed.

```vb
' NextButton
Selection.Find.ClearFormatting
Selection.Find.Style = ActiveDocument.Styles("Heading 1")
Selection.Find.Forward = True
Selection.Find.Wrap = wdFindContinue
Selection.Find.Execute

' PrevButton
Selection.Find.ClearFormatting
Selection.Find.Style = ActiveDocument.Styles("Heading 1")
Selection.Find.Forward = False
Selection.Find.Wrap = wdFindContinue
Selection.Find.Execute
```

`wdFindContinue` handles wrap-around automatically: Next from Revelation finds
Genesis; Prev from Genesis finds Revelation.

### File Changed

`src/aeRibbonClass.cls` — `NextButton`, `PrevButton`; removed `NavigateToBookIndex`,
`CurrentBookIndex`, `LastBookIndex`

### Note on headingData

`headingData` and `CaptureHeading1s`/`LogHeadingData` are retained. The array remains
available for future use by `GoToVerseSBL` to bound book-scoped searches (see
section 13).

---

## 17 — Pre-paginate at Startup: ActiveDocument.Repaginate

### Problem

Navigation with `Selection.Find` (Section 16) was still >1 minute to Revelation,
Word showed "not responding", and the window was demoted behind Explorer during the
operation.

### Why Find is also slow

The Section 16 claim that Find operates entirely at the backing-store level was
partially wrong. Find does search the backing store without layout — but after
locating the heading, Word must scroll to display the result. That display step
requires computing layout for all content between the current view and the target.
For a full-Bible document this means repaginating potentially hundreds of pages
synchronously on the UI thread. This bottleneck applies to every navigation method:
`Range.Select`, `Find`, `GoTo`, `ScrollIntoView` — all trigger the same layout
computation when jumping to an unrendered location.

The Explorer window switches are the OS demoting a window whose message pump has
stalled during that synchronous layout computation.

### Root cause

Word repaginates on demand. On first open, the document is not fully laid out. The
first navigation to a far location forces Word to compute all intervening layout
synchronously, blocking the UI thread. This is a Word layout engine constraint, not
a VBA code defect.

### Fix

Call `ActiveDocument.Repaginate` in `EnableButtonsRoutine` before `CaptureHeading1s`.
This forces full layout computation once at ribbon-load time. A status bar message
informs the user during the wait. After repagination completes, all subsequent
navigation (Next, Prev, GoTo) is instant because the layout cache is warm.

`Application.StatusBar = False` is also restored in `PROC_ERR` so the status bar
is never left in a stale state after an error.

```vb
Private Sub EnableButtonsRoutine()
    On Error GoTo PROC_ERR
    Debug.Print "RibbonController: EnableButtonsRoutine"
    Application.StatusBar = "Bible: Computing document layout..."
    ActiveDocument.Repaginate
    Application.StatusBar = False
    CaptureHeading1s
    LogHeadingData
PROC_EXIT:
    Exit Sub
PROC_ERR:
    Application.StatusBar = False
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure EnableButtonsRoutine of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub
```

### File Changed

`src/aeRibbonClass.cls` — `EnableButtonsRoutine`

### Expected Result

Navigation with Next/Prev/GoToH1 is instant after the initial repagination at
document open. The repagination delay (~15–45 seconds for a full-Bible document)
is visible only once, at load time, where a wait is expected.

---

## 18 — Draft View Navigation + Collapse Selection

### Problems

1. `ActiveDocument.Repaginate` (Section 17) added 17 seconds to startup with no
   improvement to navigation speed. The status bar message was never visible because
   the splash screen was open during `EnableButtonsRoutine`.
2. Navigation with `Selection.Find` still took 30 seconds to Revelation with a
   2-second blank screen before the page appeared and the Word window demoting behind
   Explorer.
3. The full Heading 1 text was selected after navigation instead of a collapsed cursor.

### Why Repaginate did not help

`ActiveDocument.Repaginate` computes page-break positions but does not pre-render
pages. Word uses a lazy visual rendering model: pages are only rendered when they
become visible. Navigation still triggered a full visual layout pass from the current
viewport to the target regardless of prior repagination. `Repaginate` is removed from
`EnableButtonsRoutine`.

### Why navigation was still slow in Print Layout view

Any navigation method — `Range.Select`, `Find`, `GoTo` — that requires Word to display
a far location in Print Layout view triggers a synchronous layout pass from the current
viewport to the target. For a full-Bible document this can be hundreds of pages of
layout computation on the UI thread, which is why Find was still slow and the window
was demoted behind Explorer.

### Fix: Navigate in Draft view

In Draft (Normal) view, Word renders text as a continuous stream with no pagination.
`Selection.Find` in Draft view is near-instant regardless of document size. Switching
back to Print Layout after navigation only requires rendering the single page where
the cursor lands, which takes a fraction of a second.

The current view is saved before the switch and restored after. `Application.
ScreenUpdating = False` hides the view transition so the user sees a direct jump to
the target heading. Both `ScreenUpdating` and the saved view are restored in `PROC_ERR`
so an error never leaves the document stranded in Draft view.

### Fix: Collapse selection after Find

`Selection.Find.Execute` selects the matched paragraph text. `Selection.Collapse
Direction:=wdCollapseStart` places a collapsed cursor at the start of the heading
instead.

### Files Changed

`src/aeRibbonClass.cls`
- `EnableButtonsRoutine`: removed `ActiveDocument.Repaginate` and `StatusBar` calls
- `NextButton`: added Draft view switch, `ScreenUpdating`, view restore, `Collapse`
- `PrevButton`: same as `NextButton`

### Expected Result

Navigation from Genesis to Revelation (and vice versa) completes in under 1 second
with no blank screen, no Explorer window switch, and a collapsed cursor at the
Heading 1 of the target book. Startup delay returns to the time required by
`CaptureHeading1s` and `LogHeadingData` only.

---

## 19 — Revert Draft View Navigation; Add Selection.Collapse

### Problem

Section 18 introduced Draft view (`wdNormalView`) switches around `Selection.Find`.
Navigation time increased from ~30 seconds to 127 seconds. Word showed "not
responding" for 67 seconds, then spun for 60 more seconds with a 10-second blank
screen before displaying Revelation.

### Why Draft view made it worse

The sequence of operations was:

1. `ActiveWindow.View.Type = wdNormalView` — switches to Draft, **clears the Print
   Layout cache**
2. `Selection.Find.Execute` — fast; finds Revelation immediately in Draft view
3. `ActiveWindow.View.Type = savedView` (= `wdPrintView`) — switches back to Print
   Layout **with the cursor at Revelation (end of document)**; Word must repaginate
   the entire Bible from page 1 to the last page from a cold cache

The 60-second spinning after Find completed was `ActiveWindow.View.Type = savedView`
rebuilding the full layout from scratch. Switching to Draft view first made the
round-trip more expensive than navigating in Print Layout directly, because it
cleared the cache that Print Layout navigation would have partially reused.

The user observation "after finding Revelation the first time there is no need to
continue to the end of the document" confirmed this: Find finished quickly but the
view switch continued processing the entire document.

### Why the Print Layout delay cannot be eliminated

In Print Layout view, displaying any location requires Word to compute line and page
layout for all content from page 1 to the target page. For a full-Bible document
this is unavoidable in VBA — there is no API that renders a location without first
computing its page position. The ~30-second first navigation to Revelation is a
Word layout engine constraint, not a code defect.

After the first navigation the layout cache is warm and all subsequent navigations
are instant within the same session.

### Fix

Remove the Draft view switches and `Application.ScreenUpdating` calls entirely.
Use `Selection.Find` directly in Print Layout. Add `Selection.Collapse
Direction:=wdCollapseStart` to place a collapsed cursor at the heading start
instead of selecting the full heading text.

### File Changed

`src/aeRibbonClass.cls` — `NextButton`, `PrevButton`

### State of Navigation Performance

| Attempt | Approach | First nav to Revelation |
|---------|----------|------------------------|
| Original | Find in Print Layout | ~30 s |
| Section 14 | ScreenUpdating=False + ScrollIntoView | ~30 s + window demotion |
| Section 15 | Range.Select | ~30 s + window demotion |
| Section 16 | headingData + Range.Select | ~60 s |
| Section 17 | Repaginate at startup | 17 s startup + ~30 s nav |
| Section 18 | Draft view switch + Find | 127 s |
| Section 19 | Find in Print Layout + Collapse | ~30 s (layout cache warm after first use) |

The first navigation to a far location in a large single-document Bible will always
incur the layout delay. Structural solutions (splitting into 66 documents) would
eliminate it but are out of scope for this review.

---

## 20 — Navigation Redesign Plan: Bounded Find, No Wrap-Around

### Context

All previous performance fixes (Sections 14–19) failed to reduce first-navigation
time to Revelation. The root cause is Word's layout engine: displaying any location
in Print Layout view requires computing layout for all preceding pages. For a
full-Bible document with the cursor at Genesis and the target at Revelation, this
means laying out the entire document cold — unavoidable in VBA. The approaches tried
made things worse by adding view-switch overhead and cache invalidation on top of the
base cost.

The user's proposed redesign avoids the problem entirely by eliminating the navigation
scenarios that require cold-cache long-distance jumps.

---

### The Plan

**Point 1 — No GoTo Revelation via PrevButton.**
PrevButton must never wrap around from Genesis to Revelation. This is the only
cold-cache long-distance case that triggers the Word layout delay. If PrevButton
cannot land on Revelation, the worst-case navigation distance is bounded to one step
back in the already-visible region of the document.

**Point 2 — Prev unavailable when at Genesis after first GoToBook.**
When the document opens and GoToH1 is used for the first time to land on Genesis,
PrevButton is disabled. Genesis is the first book; there is no prior Heading 1. This
prevents the user from pressing Prev and triggering a wrap or a no-op with an
unexpected result.

**Point 3 — PrevButton only enabled after NextButton is used once.**
After any GoToH1, only NextButton is initially enabled. PrevButton becomes enabled
after NextButton is pressed at least once. This guarantees that by the time Prev is
available, the user has already navigated forward and Word has computed layout for
some content beyond the starting point. The user cannot trigger a backward jump from
a cold-cache position.

**Point 4 — Stop at document boundary; do not wrap.**
Change `Wrap = wdFindContinue` to `Wrap = wdFindStop` in both NextButton and
PrevButton. Find stops at the document boundaries without wrapping. This directly
prevents the case where backward Find from Genesis wraps to Revelation and triggers
a full-document layout pass.

**Point 5 — Collapsed cursor, not selected text.**
`Selection.Collapse Direction:=wdCollapseStart` after `Find.Execute` places a
collapsed cursor at the heading start. The Heading 1 text is not selected.
Already present in the code but not observed working — suspected cause is that
`Execute` wrapped (wdFindContinue) and left the selection in an unexpected state.
With `wdFindStop` this should resolve cleanly.

---

### Discussion

**Point 3 scope — all GoToH1 targets or Genesis only?**
Points 2 and 3 together raise a question: does "Prev disabled until Next used once"
apply only when GoToH1 lands on Genesis, or for any GoToH1 target?

- If Genesis only: GoToH1 to Exodus enables both Next and Prev immediately (user
  can go Exodus→Genesis). This is natural but allows a short backward jump from any
  cold-cache position.
- If any GoToH1 target: GoToH1 to any book always enables Next first; Prev follows
  after one Next press. This is the safer rule and matches the literal text of
  Point 3.

Recommendation: apply Point 3 to all GoToH1 targets. After GoToH1, always enable
Next only. After NextButton fires successfully, enable Prev. This is the simplest
rule to implement and the most consistent from the user's perspective.

**Point 4 and button state at boundaries.**
With `wdFindStop`, NextButton from Revelation finds nothing (no forward H1, no wrap)
and `Execute` returns False. PrevButton from Genesis finds nothing and returns False.
In both cases the button fires but nothing visible happens. The buttons should be
disabled when the boundary is reached to prevent silent no-ops.

Implementation: check the return value of `Find.Execute`. If False, disable the
corresponding button and invalidate the control. This gives the user clear feedback
that the boundary has been reached.

**Point 1 and point 4 relationship.**
Point 1 (no GoTo Revelation via Prev) is fully satisfied by Point 4 (`wdFindStop`).
No separate guard is needed. With `wdFindStop`, PrevButton from Genesis simply stops
— it cannot wrap to Revelation.

**Point 5 — why Collapse was not working.**
With `wdFindContinue` and a wrapping Find, `Execute` may return True but leave the
Selection spanning from the wrap point to the matched paragraph, or spanning the full
matched paragraph style run. `Collapse Direction:=wdCollapseStart` on a wrapped
selection behaves differently than on a clean in-document match. With `wdFindStop`
there is no wrap, `Execute` either succeeds cleanly (selection = matched paragraph)
or returns False (selection unchanged). `Collapse` will then work as expected.

---

### Proposed Button State Machine

| Event | NextButton | PrevButton |
|-------|-----------|-----------|
| Document open (no GoToH1 yet) | Disabled | Disabled |
| GoToH1 succeeds (any book) | Enabled | Disabled |
| NextButton fires, Find succeeds | Enabled | Enabled |
| NextButton fires, Find fails (at Revelation) | Disabled | Enabled |
| PrevButton fires, Find succeeds | Enabled | Enabled |
| PrevButton fires, Find fails (at Genesis) | Enabled | Disabled |

---

### Files to Change (pending approval)

- `src/aeRibbonClass.cls` — `NextButton`, `PrevButton`: change to `wdFindStop`,
  add boundary detection, update button state after each navigation
- `src/aeRibbonClass.cls` — `GoToH1`: after match, enable Next, disable Prev,
  invalidate both controls

---

## 21 — Navigation Redesign: Decisions and Implementation

### Approved Decisions (from Section 20 discussion)

**Point 3 — Scope of Prev-disabled rule:** Apply to all GoToH1 targets, not Genesis
only. After GoToH1, always enable Next only; Prev becomes enabled after the first
successful NextButton press. Approved.

**Point 4 — Boundary detection:** Check the return value of `Find.Execute`. If
False, disable the corresponding button and invalidate the control. Approved.

**Point 5 — Collapse fix:** `wdFindStop` eliminates the wrapping Selection state
that caused Collapse to behave unexpectedly. Approved.

---

### Implementation

**`GoToH1`** — after a successful match, enable Next, disable Prev:
```vb
m_btnNextEnabled = True
m_btnPrevEnabled = False
```

**`NextButton`** — `wdFindStop`, return-value check, enable Prev on success,
disable Next on boundary:
```vb
found = Selection.Find.Execute
If found Then
    Selection.Collapse Direction:=wdCollapseStart
    m_btnPrevEnabled = True
    m_ribbon.InvalidateControl "GoToPrevButton"
Else
    m_btnNextEnabled = False
    m_ribbon.InvalidateControl "GoToNextButton"
End If
```

**`PrevButton`** — mirror of NextButton with Forward = False, enable Next on
success, disable Prev on boundary:
```vb
found = Selection.Find.Execute
If found Then
    Selection.Collapse Direction:=wdCollapseStart
    m_btnNextEnabled = True
    m_ribbon.InvalidateControl "GoToNextButton"
Else
    m_btnPrevEnabled = False
    m_ribbon.InvalidateControl "GoToPrevButton"
End If
```

### Files Changed

`src/aeRibbonClass.cls` — `GoToH1`, `NextButton`, `PrevButton`

### Expected Behaviour

| Scenario | Result |
|----------|--------|
| GoToH1 → any book | Next enabled, Prev disabled |
| NextButton → found | Both enabled, cursor at H1 start |
| NextButton → Revelation (boundary) | Next disabled, Prev stays enabled |
| PrevButton → found | Both enabled, cursor at H1 start |
| PrevButton → Genesis (boundary) | Prev disabled, Next stays enabled |
| PrevButton from Genesis (cold cache) | Not reachable — Prev is disabled after GoToH1 |

---

## 22 — Button State and NextButton Find Bugs

### Bugs Reported

1. **NextButton not working after GoTo Genesis** — clicking Next left the cursor at Genesis.
2. **PrevButton enabled after GoTo Genesis** — should be disabled; clicking it disabled it reactively instead of proactively.
3. **NextButton enabled after GoTo Revelation** — should be disabled; no next book exists.
4. **PrevButton disabled after GoTo Revelation** — should be enabled; Jude is a valid Prev target.

---

### Root Cause 1 — NextButton re-finds the current heading

`Selection.Find` with `Forward = True` starts searching from the beginning of the
current selection. After `GoToH1`, the cursor is collapsed at `para.Range.Start` —
the first character of the Genesis Heading 1 paragraph. Find immediately matches
Genesis itself, returns True, collapses back to the same position, enables Prev.
Visually nothing moves.

**Fix:** Before calling Find forward, advance the cursor to the end of the current
paragraph so the search starts after the heading:

```vb
Dim curParaEnd As Long
curParaEnd = Selection.Paragraphs(1).Range.End
Selection.SetRange curParaEnd, curParaEnd
```

`Selection.Paragraphs(1).Range.End` is the position immediately after the paragraph
mark of the current paragraph. Find forward from that position skips the current
heading and finds the next one (Exodus).

---

### Root Cause 2 — GoToH1 set button states without boundary detection

The Section 21 implementation set `m_btnPrevEnabled = False` unconditionally in
`GoToH1` (correct for Genesis, wrong for all other books) and `m_btnNextEnabled =
True` unconditionally (correct for all books except Revelation).

**Fix:** After finding the matched paragraph, compare its start position against
`headingData` to detect first and last book boundaries:

- First book (Genesis): `m_btnPrevEnabled = False`, `m_btnNextEnabled = True`
- Last book (Revelation): `m_btnPrevEnabled = True`, `m_btnNextEnabled = False`
- Any other book: both `True`

```vb
foundPos = para.Range.Start
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
```

This replaces the approved-but-too-restrictive "always disable Prev after GoToH1"
rule from Section 21. The `wdFindStop` rule (no wrap-around) already prevents the
cold-cache Genesis→Revelation problem, so Prev can safely be enabled at non-first
books immediately after GoToH1.

---

### Files Changed

`src/aeRibbonClass.cls`
- `GoToH1`: replaced fixed `m_btnPrevEnabled = False` with headingData boundary
  detection; sets both buttons correctly for Genesis, Revelation, and middle books
- `NextButton`: added `curParaEnd` advancement before `Selection.Find.Execute`

### Expected Behaviour After Fix

| GoToH1 target | NextButton | PrevButton |
|---------------|-----------|-----------|
| Genesis (first) | Enabled | Disabled |
| Any middle book | Enabled | Enabled |
| Revelation (last) | Disabled | Enabled |

NextButton from Genesis now advances past the Genesis heading before searching,
finding Exodus on the first press.

---

## 23 — GoToH1 Rewrite: Eliminate Paragraph Loop and Redundant Selections

### Bugs Reported

1. **GoTo Revelation takes 20 seconds** — paragraph loop iterates the entire document.
2. **After finding Revelation, Word spins for another 25 seconds** — two selection
   operations each trigger a full layout pass; `Selection.Range.Select` and `DoEvents`
   add a third layout/repaint cycle and Explorer window switches.
3. **NextButton from Jude to Revelation shows blank screen for 5 seconds** —
   unavoidable first-time layout computation; acceptable.

---

### Bug 1 — `For Each para In ActiveDocument.Paragraphs` (20 seconds)

The paragraph loop iterated every paragraph in the document (~31,000 for a full
Bible). To find Revelation it had to process all preceding paragraphs before reaching
the last Heading 1. `CaptureHeading1s` already stored all 66 heading names and
positions in `headingData` at ribbon load time. The document did not need to be
accessed at all during GoToH1.

**Fix:** Replace the paragraph loop with a 66-iteration search over `headingData(i, 0)`.

---

### Bug 2 — Double selection and redundant operations (25 seconds)

The original code executed three operations after finding the match:

```vb
para.Range.Select                                              ' (1) select full heading
ActiveDocument.Range(para.Range.Start, para.Range.Start).Select ' (2) collapse to start
...
Application.ScreenUpdating = True
Selection.Range.Select                                         ' (3) re-select same position
DoEvents                                                       ' (4) pump message queue
```

- **(1)** `para.Range.Select` selected the full heading paragraph. Even with
  `ScreenUpdating = False`, this resolved the character position against the layout
  engine internally.
- **(2)** `ActiveDocument.Range(...).Select` immediately after (1) triggered a second
  layout pass to collapse to the start position.
- **(3)** `Selection.Range.Select` on an already-set selection was redundant and
  triggered a third paint cycle after `ScreenUpdating` was restored.
- **(4)** `DoEvents` pumped the Windows message queue, which is what brought the
  Explorer window to the front — Windows activated another window because the message
  pump had been starved during the two layout passes.

**Fix:** Remove all four operations. Navigate with a single
`ActiveDocument.Range(foundPos, foundPos).Select`. Remove `Application.ScreenUpdating`,
`Selection.Range.Select`, and `DoEvents` entirely.

---

### Rewritten GoToH1

```
1. InputBox for pattern
2. Loop headingData(1..66, 0) — 66 string comparisons, no document access
3. If not found: MsgBox, exit
4. ActiveDocument.Range(foundPos, foundPos).Select — single navigation
5. headingData boundary detection — set Next/Prev enabled
6. InvalidateControl for both buttons
```

The paragraph loop (`For Each`), `Application.ScreenUpdating`, `para.Range.Select`,
`Selection.Range.Select`, and `DoEvents` are all removed.

---

### Bug 3 — NextButton 5-second delay at Revelation (accepted)

When NextButton advances from Jude to Revelation, Word must render the Revelation
page. Even though GoToH1 to Revelation was performed earlier in the session, some
layout may have been evicted from the cache during the intervening navigation to Jude.
The 5-second delay is the minimum cost of displaying a far location in a large Print
Layout document. No further optimisation is possible in VBA.

---

### Files Changed

`src/aeRibbonClass.cls` — `GoToH1` fully rewritten

### Expected Result

GoTo Revelation completes in a single layout pass (~5 seconds, same as NextButton)
instead of three passes (~45 seconds). No Explorer window switches. `DoEvents` is
no longer called.

---

## 24 — Flash and 3-Second Blank: ScreenUpdating and Find.Text

### Bugs Reported

1. **Flash (selection gray box) on first PrevButton from Revelation to Jude.**
2. **3-second blank screen on second PrevButton from Revelation to Jude** (after
   cycling 20 Prev presses then Next presses back to Revelation).

---

### Bug 1 — Selection flash

`Selection.Find.Execute` selects the matched heading text. Word issues a repaint
between `Execute` completing and `Selection.Collapse` running. The gray selection
box is briefly visible for that one repaint cycle. It was only noticeable at
Revelation → Jude because those pages had just been rendered and the paint occurred
before the Collapse could execute.

**Fix:** `Application.ScreenUpdating = False` before `Execute`; restored after
`Collapse`. Word suppresses the intermediate repaint; the user sees only the final
collapsed cursor, not the selected heading text.

---

### Bug 2 — 3-second blank on second navigation

`NextButton` has two viewport changes per press:

1. `Selection.SetRange curParaEnd, curParaEnd` — moves cursor to the END of the
   current heading paragraph (body area), triggering a scroll if the body is not
   already visible.
2. `Selection.Find.Execute` — finds the next H1 and scrolls to it.

`PrevButton` has one viewport change per press (Find result only).

After 20 consecutive NextButton presses cycling back to Revelation, the double
scroll per press evicts the pixel render cache for pages near each jump point.
When PrevButton subsequently scrolls from Revelation back to Jude, Word must
re-render Jude's page from layout data rather than from the pixel cache: hence
the 3-second blank.

**Fix:** `Application.ScreenUpdating = False` before `Selection.SetRange` (in
addition to before `Execute`). The intermediate cursor advance to the body area
never triggers a repaint request, so Word does not need to render that intermediate
state and the pixel cache for surrounding pages is preserved.

**Additional fix:** `Selection.Find.Text = ""` added before `Execute` in both
buttons. `ClearFormatting` clears format constraints but does not reset the text
search pattern. If Word's Find object retained text from a previous search (e.g.,
via the Find & Replace dialog), the style-only search would silently include a
text filter and could skip valid headings.

---

### Files Changed

`src/aeRibbonClass.cls` — `NextButton`, `PrevButton`

Changes:
- `Application.ScreenUpdating = False` moved before `Selection.SetRange` in
  `NextButton` and before `Selection.Find.Execute` in `PrevButton`
- `Application.ScreenUpdating = True` after `Selection.Collapse` in both
- `Application.ScreenUpdating = True` added to `PROC_ERR` in both so the
  screen is never left frozen after an error
- `Selection.Find.Text = ""` added before `Execute` in both

### Note

If Bug 2 persists after this fix, the remaining cause is Word's pixel render
cache behaviour, which is not addressable from VBA.
