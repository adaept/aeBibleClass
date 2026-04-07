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
