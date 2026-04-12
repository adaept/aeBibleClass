Attribute VB_Name = "basBibleRibbonSetup"
Option Explicit
Option Private Module

' -- Singleton Instance --------------------------------------------------------
Private s_instance As aeRibbonClass

Public Function Instance() As aeRibbonClass
    On Error GoTo PROC_ERR
    If s_instance Is Nothing Then
        Set s_instance = New aeRibbonClass
    End If
    Set Instance = s_instance
PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Instance of Module basBibleRibbonSetup"
    Resume PROC_EXIT
End Function

' -- Bootstrap -----------------------------------------------------------------

Public Sub AutoExec()
    On Error GoTo PROC_ERR
    Debug.Print "basBibleRibbonSetup: AutoExec at " & Format(Now, "hh:nn:ss")
    Dim rc As aeRibbonClass
    Set rc = Instance()
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure AutoExec of Module basBibleRibbonSetup"
    Resume PROC_EXIT
End Sub

Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    On Error GoTo PROC_ERR
    Debug.Print ">> RibbonOnLoad at " & Format(Now, "hh:nn:ss")
    Instance().OnRibbonLoad ribbon
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RibbonOnLoad of Module basBibleRibbonSetup"
    Resume PROC_EXIT
End Sub

' -- Callback Stubs ------------------------------------------------------------

' OnGoToVerseSblClick removed from ribbon XML (comboBox design replaces large GoTo Verse button)
' Implementation preserved in aeRibbonClass.cls.GoToVerseSBL for reference.
' Public Sub OnGoToVerseSblClick(control As IRibbonControl)
'     Instance().OnGoToVerseSblClick control
' End Sub

Public Sub OnPrevButtonClick(control As IRibbonControl)
    'Debug.Print ">> OnPrevButtonClick at " & Format(Now, "hh:nn:ss")
    Instance().OnPrevButtonClick control
End Sub

' OnGoToH1ButtonClick removed from ribbon XML (GoTo Book is now the Book comboBox).
' GoToH1 implementation preserved in aeRibbonClass.cls for reference.
' Public Sub OnGoToH1ButtonClick(control As IRibbonControl)
'     Const EXPECTED_PROJECT As String = "Project"
'     Dim projName As String
'     projName = Application.ActiveDocument.VBProject.Name
'     Application.OnTime Now + TimeValue("00:00:01"), projName & ".basRibbonDeferred.GoToH1Deferred"
' End Sub

Public Sub OnNextButtonClick(control As IRibbonControl)
    'Debug.Print ">> OnNextButtonClick at " & Format(Now, "hh:nn:ss")
    Instance().OnNextButtonClick control
End Sub

Public Sub OnAdaeptAboutClick(control As IRibbonControl)
    'Debug.Print ">> OnAdaeptAboutClick at " & Format(Now, "hh:nn:ss")
    Instance().OnAdaeptAboutClick control
End Sub

Public Sub GetPrevEnabled(control As IRibbonControl, ByRef enabled)
    Dim rc As aeRibbonClass
    Set rc = Instance()
    enabled = rc.BtnPrevEnabled
End Sub

' Test stub — runs GoToH1 outside the ribbon callback (Alt+F8 -> TestGoToH1Direct)
' If the second block disappears vs ribbon button, ribbon callback return is the cause.
Public Sub TestGoToH1Direct()
    Dim rc As aeRibbonClass
    Set rc = Instance()
    rc.GoToH1Direct
End Sub

Public Sub GetNextEnabled(control As IRibbonControl, ByRef enabled)
    Dim rc As aeRibbonClass
    Set rc = Instance()
    enabled = rc.BtnNextEnabled
End Sub

' -- Book comboBox callbacks ---------------------------------------------------

Public Sub GetBookEnabled(control As IRibbonControl, ByRef enabled)
    enabled = Instance().GetBookEnabled(control)
End Sub

Public Sub GetBookCount(control As IRibbonControl, ByRef count)
    count = Instance().GetBookCount(control)
End Sub

Public Sub GetBookItemLabel(control As IRibbonControl, index As Long, ByRef label)
    label = Instance().GetBookItemLabel(control, index)
End Sub

Public Sub GetBookItemID(control As IRibbonControl, index As Long, ByRef id)
    id = Instance().GetBookItemID(control, index)
End Sub

Public Sub GetBookText(control As IRibbonControl, ByRef text)
    text = Instance().GetBookText(control)
End Sub

Public Sub OnBookChanged(control As IRibbonControl, text As String)
    Instance().OnBookChanged control, text
End Sub

' -- Chapter row callbacks -----------------------------------------------------

Public Sub GetPrevChapterEnabled(control As IRibbonControl, ByRef enabled)
    enabled = Instance().GetPrevChapterEnabled(control)
End Sub

Public Sub GetNextChapterEnabled(control As IRibbonControl, ByRef enabled)
    enabled = Instance().GetNextChapterEnabled(control)
End Sub

Public Sub GetChapterEnabled(control As IRibbonControl, ByRef enabled)
    enabled = Instance().GetChapterEnabled(control)
End Sub

Public Sub GetChapterText(control As IRibbonControl, ByRef text)
    text = Instance().GetChapterText(control)
End Sub

Public Sub OnChapterChanged(control As IRibbonControl, text As String)
    Instance().OnChapterChanged control, text
End Sub

Public Sub OnChapterAction(control As IRibbonControl, text As String)
    Instance().OnChapterAction control, text
End Sub

Public Sub OnPrevChapterClick(control As IRibbonControl)
    Instance().OnPrevChapterClick control
End Sub

Public Sub OnNextChapterClick(control As IRibbonControl)
    Instance().OnNextChapterClick control
End Sub

' -- Verse row callbacks -------------------------------------------------------

Public Sub GetPrevVerseEnabled(control As IRibbonControl, ByRef enabled)
    enabled = Instance().GetPrevVerseEnabled(control)
End Sub

Public Sub GetNextVerseEnabled(control As IRibbonControl, ByRef enabled)
    enabled = Instance().GetNextVerseEnabled(control)
End Sub

Public Sub GetVerseEnabled(control As IRibbonControl, ByRef enabled)
    enabled = Instance().GetVerseEnabled(control)
End Sub

Public Sub GetVerseText(control As IRibbonControl, ByRef text)
    text = Instance().GetVerseText(control)
End Sub

Public Sub OnVerseChanged(control As IRibbonControl, text As String)
    Instance().OnVerseChanged control, text
End Sub

Public Sub OnVerseAction(control As IRibbonControl, text As String)
    Instance().OnVerseAction control, text
End Sub

Public Sub OnPrevVerseClick(control As IRibbonControl)
    Instance().OnPrevVerseClick control
End Sub

Public Sub OnNextVerseClick(control As IRibbonControl)
    Instance().OnNextVerseClick control
End Sub

' -- New Search ----------------------------------------------------------------

Public Sub GetNewSearchEnabled(control As IRibbonControl, ByRef enabled)
    enabled = Instance().GetNewSearchEnabled(control)
End Sub

Public Sub OnNewSearchClick(control As IRibbonControl)
    Instance().OnNewSearchClick control
End Sub

