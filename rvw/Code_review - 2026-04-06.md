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
