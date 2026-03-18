Attribute VB_Name = "basBibleRibbonSetup"
Option Explicit
Option Private Module

' -- Singleton Instance --------------------------------------------------------
Private s_instance As aeRibbonClass

Public Function Instance() As aeRibbonClass
    If s_instance Is Nothing Then
        Set s_instance = New aeRibbonClass
    End If
    Set Instance = s_instance
End Function

' -- Bootstrap -----------------------------------------------------------------

Public Sub AutoExec()
    Debug.Print "basBibleRibbonSetup: AutoExec at " & Format(Now, "hh:nn:ss")
    Dim rc As aeRibbonClass
    Set rc = Instance()
End Sub

Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Debug.Print ">> RibbonOnLoad at " & Format(Now, "hh:nn:ss")
    Instance().OnRibbonLoad ribbon
End Sub

' -- Callback Stubs ------------------------------------------------------------

Public Sub OnGoToVerseSblClick(control As IRibbonControl)
    'Debug.Print ">> OnGoToVerseSblClick at " & Format$(Date + Timer / 86400#, "yyyy-mm-dd hh:nn:ss.000")
    Instance().OnGoToVerseSblClick control
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

