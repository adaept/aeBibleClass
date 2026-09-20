VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Copyright (c) 2025-2026 Peter F. Ennis
'SPDX-License-Identifier: LGPL-3.0-or-later OR LicenseRef-adaept-Commercial
'DUAL-LICENSED. You may use this file under EITHER of:
'  (1) the GNU Lesser General Public License, version 3.0 or (at your option)
'      any later version  -  https://www.gnu.org/licenses/lgpl-3.0.txt ; OR
'  (2) a commercial / proprietary license available from adaept (Peter Ennis),
'      permitting use in closed-source / proprietary works WITHOUT the LGPL
'      copyleft obligations. Contact the copyright holder for commercial terms.
'As the sole copyright holder, adaept may license this file under either option.
'This library is distributed WITHOUT ANY WARRANTY; without even the implied
'warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.

' Document_Open: only startup logic is the ribbon-conflict guard below.
' The commented-out lines were used only for document-open timing tests.
Private Sub Document_Open()
    'ActiveWindow.View.Type = wdNormalView
    'ActiveDocument.ActiveWindow.Caption = "Loaded at: " & Now()
    CheckForConflictingRibbonAddIn
End Sub

' Document_BeforeSave (added 2026-09-19): auto-enforces EDSG/01-styles.md's
' "no clickable hyperlinks anywhere" rule on every save, so hyperlink creep
' can never survive into a saved file. Real-world trigger: an editing
' session that fixes license/map references (which often carry pasted
' links) can leave an active Hyperlink object behind even after applying
' the BookHyperlink character style by hand - the style and the click
' target are two independent things in Word's object model, and only
' removing the latter satisfies RUN_THE_TESTS slot 17
' (CountActiveHyperlinks, expected 0). "Right-click > Remove Hyperlink"
' fixes one instance; this hook makes the fix unconditional going forward
' instead of relying on remembering to run LockBookHyperlinks by hand.
'
' Cheap check first (sum story.Hyperlinks.Count across every StoryRange,
' the same measure CountActiveHyperlinks in aeBibleClass.cls uses) - the
' heavier three-step LockBookHyperlinks (full-document Find scans) only
' runs when there's actually something to fix. silent:=True suppresses its
' completion MsgBox (Debug.Print only) so saving never pops a modal - see
' basTEST_aeBibleTools.LockBookHyperlinks's own comment for that parameter.
'
' This document's docm/VBA pathway is NOT the enforcement point for
' translated editions, which are edited as .docx via aeBibleAddin's JS
' taskpane (VBA's own text/encoding handling isn't reliable for non-Latin/
' UTF-8 scripts) - see that taskpane's "Lock Hyperlinks" button and its
' auto-run-on-open call for the equivalent JS-side enforcement, and
' EDSG/01-styles.md's "Companion rule" section for why the two pathways
' need separate mechanisms (Office.js has no before-save event to hook).
Private Sub Document_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim story As Word.Range
    Dim activeCount As Long

    For Each story In ActiveDocument.StoryRanges
        activeCount = activeCount + story.Hyperlinks.Count
    Next story

    If activeCount > 0 Then
        Debug.Print "Document_BeforeSave: " & activeCount & _
            " active hyperlink(s) found before save -- auto-running LockBookHyperlinks."
        LockBookHyperlinks silent:=True
    End If
End Sub

' This document carries its own embedded ribbon (customUI/customUI14.xml,
' tab id "RWB" - see py/inject_ribbon.py) for VBA dev work. aeRibbon.dotm
' is a disposable build artefact used only for aeRibbon/aeBibleAddin
' ribbon-build testing (see aeRibbon/BUILD.md's G7/G8 gates) and is never
' needed for ordinary VBA dev work on this SSOT docm. If it's left
' installed in Word's global Startup folder, its identical "RWB" tab
' loads alongside this document's own copy and Word renders the tab
' twice - confirmed 2026-09-19 (operator report: two "Radiant Word Bible"
' tabs on open).
Private Sub CheckForConflictingRibbonAddIn()
    On Error GoTo PROC_ERR
    Const ADDIN_NAME As String = "aeRibbon.dotm"
    Dim addInPath As String
    Dim ai As AddIn

    addInPath = Application.StartupPath & "\" & ADDIN_NAME
    If Dir(addInPath) = "" Then Exit Sub   ' not installed - nothing to do

    If MsgBox("'" & ADDIN_NAME & "' is installed as a global Word add-in in:" & vbCrLf & _
              Application.StartupPath & vbCrLf & vbCrLf & _
              "This document already carries its own copy of the same ribbon tab " & _
              "for VBA dev work, so both will show side by side (""Radiant Word Bible"" " & _
              "twice) as long as the add-in stays installed." & vbCrLf & vbCrLf & _
              ADDIN_NAME & " is only needed for aeRibbon/aeBibleAddin ribbon-build " & _
              "testing and is regenerated by that build procedure - it is safe to delete." & _
              vbCrLf & vbCrLf & "Delete it now?", _
              vbYesNo + vbExclamation, "Duplicate ribbon tab detected") = vbYes Then

        ' Unload before deleting - Word holds the file open as a loaded
        ' global template while Installed=True; Kill fails with Err 70
        ' (Permission denied) against a file still open in this process.
        For Each ai In Application.AddIns
            If StrComp(ai.Name, ADDIN_NAME, vbTextCompare) = 0 Then
                ai.Installed = False
                Exit For
            End If
        Next ai

        Kill addInPath
        Debug.Print "CheckForConflictingRibbonAddIn: deleted " & addInPath
    Else
        Debug.Print "CheckForConflictingRibbonAddIn: user declined deletion; " & _
                    "duplicate tab will persist this session."
    End If

    Exit Sub
PROC_ERR:
    Debug.Print "ERROR in ThisDocument.CheckForConflictingRibbonAddIn | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description & _
        " - '" & ADDIN_NAME & "' may still be loaded; delete it manually from " & _
        Application.StartupPath & " (close Word first if it's locked) if this recurs."
End Sub


