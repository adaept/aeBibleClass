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

Public Sub OnGoToVerseSblClick(control As IRibbonControl)
    'Debug.Print ">> OnGoToVerseSblClick at " & Format$(Date + Timer / 86400#, "yyyy-mm-dd hh:nn:ss.000")
    Instance().OnGoToVerseSblClick control
End Sub

Public Sub OnPrevButtonClick(control As IRibbonControl)
    'Debug.Print ">> OnPrevButtonClick at " & Format(Now, "hh:nn:ss")
    Instance().OnPrevButtonClick control
End Sub

Public Sub OnGoToH1ButtonClick(control As IRibbonControl)
    'Debug.Print ">> OnGoToH1ButtonClick at " & Format(Now, "hh:nn:ss")
    Instance().OnGoToH1ButtonClick control
End Sub

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

Public Sub GetNextEnabled(control As IRibbonControl, ByRef enabled)
    Dim rc As aeRibbonClass
    Set rc = Instance()
    enabled = rc.BtnNextEnabled
End Sub

