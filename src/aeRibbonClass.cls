VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aeRibbonClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'Copyright (c) 2025-2026 Peter F. Ennis
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation;
'version 3.0, or (at your option) any later version.
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'Lesser General Public License for more details.
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, visit
'https://www.gnu.org/licenses/lgpl-3.0.txt

' ================================================================
' Author:   Peter F. Ennis
' Date:     March 10, 2025
' Comment:  Bible Ribbon Class for Word 365 VBA
' History:  See comment details, basChangeLog_aeBibleClass, commit messages on GitHub
' GitHub:   https://github.com/adaept/aeBibleClass
' ================================================================

' -- Private State -------------------------------------------------------------
Private m_ribbon         As IRibbonUI
Private m_ribbonReady    As Boolean
Private m_btnNextEnabled As Boolean
Private m_btnPrevEnabled As Boolean
Private m_bookmarkIndex  As Long
Private m_savedPos       As Long
Private m_bookAbbr       As String
Private headingData(1 To 66, 0 To 1) As Variant

' Windows API - cursor control
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" ( _
    ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

' -- Lifecycle -----------------------------------------------------------------

Private Sub Class_Initialize()
    Debug.Print "RibbonController: Class_Initialize at " & Format(Now, "hh:nn:ss")
    m_btnNextEnabled = False
    m_btnPrevEnabled = False
    m_ribbonReady = False
    m_bookmarkIndex = 0
    m_savedPos = 0
    m_bookAbbr = vbNullString
End Sub

Private Sub Class_Terminate()
    Debug.Print "RibbonController: Class_Terminate at " & Format(Now, "hh:nn:ss")
    Set m_ribbon = Nothing
    m_ribbonReady = False
End Sub

' -- Properties ----------------------------------------------------------------

Public Property Get RibbonReady() As Boolean
    RibbonReady = m_ribbonReady
End Property

Public Property Get BtnNextEnabled() As Boolean
    BtnNextEnabled = m_btnNextEnabled
End Property

Public Property Let BtnNextEnabled(val As Boolean)
    m_btnNextEnabled = val
End Property

Public Property Get BtnPrevEnabled() As Boolean
    BtnPrevEnabled = m_btnPrevEnabled
End Property

Public Property Let BtnPrevEnabled(val As Boolean)
    m_btnPrevEnabled = val
End Property

' -- Ribbon Load ---------------------------------------------------------------

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    On Error GoTo PROC_ERR
    Set m_ribbon = ribbon
    m_ribbonReady = True
    Debug.Print "RibbonController: Ribbon ready at " & Format(Now, "hh:nn:ss")
    m_ribbon.InvalidateControl "GoToNextButton"
    m_ribbon.InvalidateControl "GoToPrevButton"
    Call EnableButtonsRoutine
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OnRibbonLoad of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Invalidation --------------------------------------------------------------

Public Sub InvalidateAll()
    On Error GoTo PROC_ERR
    If Not m_ribbon Is Nothing Then m_ribbon.Invalidate
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure InvalidateAll of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

Public Sub InvalidateControl(controlId As String)
    On Error GoTo PROC_ERR
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl controlId
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure InvalidateControl of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Ribbon Callbacks ----------------------------------------------------------

Public Sub OnGoToVerseSblClick(control As IRibbonControl)
    Call GoToVerseSBL
End Sub

Public Sub OnPrevButtonClick(control As IRibbonControl)
    Call PrevButton
End Sub

Public Sub OnGoToH1ButtonClick(control As IRibbonControl)
    Call GoToH1
End Sub

Public Sub OnNextButtonClick(control As IRibbonControl)
    Call NextButton
End Sub

Public Sub OnAdaeptAboutClick(control As IRibbonControl)
    MsgBox "Hello, adaept World!" & vbCrLf & _
           "adaeptMsg  = " & AdaeptMsg, vbInformation, "About adaept"
End Sub

' -- Enabled State Callback ----------------------------------------------------

Public Function GetNextEnabled(control As IRibbonControl) As Boolean
    GetNextEnabled = m_btnNextEnabled
End Function

Public Function GetPrevEnabled(control As IRibbonControl) As Boolean
    GetPrevEnabled = m_btnPrevEnabled
End Function

' -- Initialization ------------------------------------------------------------

Private Sub EnableButtonsRoutine()
    On Error GoTo PROC_ERR
    Debug.Print "RibbonController: EnableButtonsRoutine"
    CaptureHeading1s
    LogHeadingData
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure EnableButtonsRoutine of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- GoTo Verse (stub - parser to be implemented) ------------------------------

Private Sub GoToVerseSBL()
    MsgBox "GoToVerseSBL - Parser not yet implemented."
End Sub

' -- GoTo Book (H1) ------------------------------------------------------------

Private Sub GoToH1()
    On Error GoTo PROC_ERR
    Dim pattern  As String
    Dim i        As Long
    Dim k        As Long
    Dim foundPos As Long
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
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "GoToNextButton"
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "GoToPrevButton"
    Application.ScreenUpdating = True

PROC_EXIT:
    Exit Sub

PROC_ERR:
    Application.ScreenUpdating = True
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GoToH1 of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Next Book (H1) ------------------------------------------------------------

Private Sub NextButton()
    On Error GoTo PROC_ERR
    Dim found    As Boolean
    Dim foundPos As Long
    Dim k        As Long
    Dim curParaEnd As Long
    curParaEnd = Selection.Paragraphs(1).Range.End
    Application.ScreenUpdating = False
    Selection.SetRange curParaEnd, curParaEnd
    Selection.Find.ClearFormatting
    Selection.Find.Text = ""
    Selection.Find.style = ActiveDocument.Styles("Heading 1")
    Selection.Find.Forward = True
    Selection.Find.Wrap = wdFindStop
    found = Selection.Find.Execute
    If found Then
        Selection.Collapse Direction:=wdCollapseStart
        foundPos = Selection.Range.Start
        m_btnPrevEnabled = True
        m_btnNextEnabled = True
        For k = 66 To 1 Step -1
            If Not IsEmpty(headingData(k, 1)) And CLng(headingData(k, 1)) > 0 Then
                If foundPos = CLng(headingData(k, 1)) Then m_btnNextEnabled = False
                Exit For
            End If
        Next k
    Else
        m_btnNextEnabled = False
    End If
    Application.ScreenUpdating = True
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "GoToNextButton"
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "GoToPrevButton"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    Application.ScreenUpdating = True
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure NextButton of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Prev Book (H1) ------------------------------------------------------------

Private Sub PrevButton()
    On Error GoTo PROC_ERR
    Dim found    As Boolean
    Dim foundPos As Long
    Application.ScreenUpdating = False
    Selection.Find.ClearFormatting
    Selection.Find.Text = ""
    Selection.Find.style = ActiveDocument.Styles("Heading 1")
    Selection.Find.Forward = False
    Selection.Find.Wrap = wdFindStop
    found = Selection.Find.Execute
    If found Then
        Selection.Collapse Direction:=wdCollapseStart
        foundPos = Selection.Range.Start
        m_btnNextEnabled = True
        m_btnPrevEnabled = True
        If Not IsEmpty(headingData(1, 1)) Then
            If foundPos = CLng(headingData(1, 1)) Then m_btnPrevEnabled = False
        End If
    Else
        m_btnPrevEnabled = False
    End If
    Application.ScreenUpdating = True
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "GoToNextButton"
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "GoToPrevButton"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    Application.ScreenUpdating = True
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure PrevButton of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Heading Data --------------------------------------------------------------

'===========================================================================================
' Purpose: Capture Heading 1 text and paragraph index into a static array (1 to 66, 0 to 1)
'   - Only runs once per session
'   - Heading text excludes vbCrLf
'   - Optimized for speed and reproducibility
'   - Audit-traceable: logs execution status to Immediate Window
'===========================================================================================
Public Sub CaptureHeading1s()
    On Error GoTo PROC_ERR
    Static hasRun As Boolean
    Dim para As Word.Paragraph
    Dim i        As Long
    Dim paraText As String

    If hasRun Then
        Debug.Print "CaptureHeading1s: Already executed this session."
        GoTo PROC_EXIT
    End If

    i = 1
    For Each para In ActiveDocument.Paragraphs
        If para.style = "Heading 1" Then
            If i > 66 Then Exit For
            paraText = Trim(Replace(para.Range.Text, vbCr, ""))
            headingData(i, 0) = paraText
            headingData(i, 1) = para.Range.Start
            i = i + 1
        End If
    Next para

    hasRun = True
    Debug.Print "CaptureHeading1s: Stored " & i - 1 & " Heading 1 entries."

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure CaptureHeading1s of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

'=============================================================================================
' Maintain a persistent 66-row CSV file + header, with session-aware ParaIndex updates
' sessionID,Heading1,BookName,ParaIndex
'   - Only updates ParaIndex and timestamp if changed
'   - Entire file is rewritten each time
'   - Prior values are preserved for unchanged rows
'   - Git-friendly: diffs show timestamped changes only
'=============================================================================================
Public Sub LogHeadingData()
    On Error GoTo PROC_ERR
    Const ROWS        As Long = 66
    Dim fPath         As String
    Dim sessionStamp  As String
    Dim i             As Long
    Dim fNum          As Integer
    Dim csvLines(1 To ROWS) As String
    Static lastParaStarts(1 To ROWS) As Long
    Static lastBookNames(1 To ROWS)  As String
    Static lastTimestamps(1 To ROWS) As String
    Dim changed As Boolean

    fPath = ActiveDocument.Path & "\rpt\HeadingLog.txt"
    sessionStamp = Format(Now, "yyyy-mm-dd hh:nn:ss")

    For i = 1 To ROWS
        If Not IsEmpty(headingData(i, 0)) Then
            If headingData(i, 1) <> lastParaStarts(i) Then
                lastParaStarts(i) = headingData(i, 1)
                lastBookNames(i) = headingData(i, 0)
                lastTimestamps(i) = sessionStamp
                changed = True
            End If
        End If
        csvLines(i) = lastTimestamps(i) & ",H1[" & i & "]," & lastBookNames(i) & "," & lastParaStarts(i)
    Next i

    fNum = FreeFile
    Open fPath For Output As #fNum
    Print #fNum, "sessionID,Heading1,BookName,ParaIndex"
    For i = 1 To ROWS
        Print #fNum, csvLines(i)
    Next i
    Close #fNum

    If changed Then
        Debug.Print "LogHeadingData: File updated at " & sessionStamp
    Else
        Debug.Print "LogHeadingData: No changes - file preserved"
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LogHeadingData of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Cursor Helpers ------------------------------------------------------------

Private Function SaveCursor() As Long
    SaveCursor = Selection.Start
End Function

Private Sub RestoreCursor(ByVal pos As Long)
    Selection.SetRange pos, pos
    Selection.Collapse Direction:=wdCollapseStart
End Sub

' -- About Message -------------------------------------------------------------

Private Function AdaeptMsg() As String
    AdaeptMsg = """...the truth shall make you free.""" & " John 8:32 (KJV)"
End Function

