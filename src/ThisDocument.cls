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
Option Compare Text

Private Sub Document_Open()
    Call AutoExec
    Debug.Print "After Document_Open AutoExec at: " & Format(Now, "hh:nn:ss")
    Application.OnTime Now + TimeValue("00:00:02"), "DelayedRibbonRefresh"
End Sub

Sub DelayedRibbonRefresh()
    If ribbonIsReady Then
        ribbonUI.InvalidateControl "GoToNextButton"
        Debug.Print "Ribbon controls refreshed at: " & Format(Now, "hh:nn:ss")
    Else
        ' Retry again in 2 seconds if not ready
        Application.OnTime Now + TimeValue("00:00:02"), "DelayedRibbonRefresh"
    End If
End Sub

