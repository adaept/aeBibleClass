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

'Private Sub Document_Close()
'    ' Show Ribbon
'    Application.CommandBars.ExecuteMso "MinimizeRibbon"
'End Sub

'Private Sub Document_Open()
'
'    Call AutoExec
'    Debug.Print "After Document_Open AutoExec at: " & Format(Now, "hh:nn:ss")
'    Application.OnTime Now + TimeValue("00:00:02"), "DelayedRibbonRefresh"
'
'    If Not UserConfirmed("Confirm?", "Hide Use Interface?") Then Exit Sub
'
'    ' Hide Ribbon
'    Application.CommandBars.ExecuteMso "MinimizeRibbon"
'
'    ' Hide Status Bar
'    Application.DisplayStatusBar = False
'
'    ' Hide/Show Scroll Bars
'    With ActiveWindow
'        .DisplayHorizontalScrollBar = False
'        .DisplayVerticalScrollBar = True
'    End With
'
'    ' Toggle Ruler
'    ActiveWindow.DisplayRulers = Not ActiveWindow.DisplayRulers
'
'End Sub

'Sub DelayedRibbonRefresh()
'    If ribbonIsReady Then
'        ribbonUI.InvalidateControl "GoToNextButton"
'        Debug.Print "Ribbon controls refreshed at: " & Format(Now, "hh:nn:ss")
'    Else
'        ' Retry again in 2 seconds if not ready
'        Application.OnTime Now + TimeValue("00:00:02"), "DelayedRibbonRefresh"
'    End If
'End Sub

