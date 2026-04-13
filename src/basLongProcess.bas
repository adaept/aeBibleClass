Attribute VB_Name = "basLongProcess"
'==============================================================================
' basLongProcess  -  Long-Running Process Entry Points
' ----------------------------------------------------------------------------
' Thin public skeleton. All batch loop logic, progress persistence, and logging
' live in aeLongProcessClass. Concrete task logic lives in implementations of
' IaeLongProcessClass (e.g. aeUpdateCharStyleClass).
'
' Entry points - call from Immediate Window:
'   TestUpdateCharStyle       start/resume UpdateCharacterStyle task
'   StopTask                  stop the active task
'   TestResetUpdateCharStyle  reset progress for UpdateCharacterStyle task
'
' SetWordHighPriority is an opt-in utility - call manually before a long task.
'==============================================================================
Option Explicit
Option Compare Text
Option Private Module

Private s_runner As aeLongProcessClass

' -----------------------------------------------------------------------------
' Test stubs - single-word Immediate Window entry points
' -----------------------------------------------------------------------------
Public Sub TestUpdateCharStyle()
    Dim t As New aeUpdateCharStyleClass
    StartOrResume t
End Sub

Public Sub TestResetUpdateCharStyle()
    Dim t As New aeUpdateCharStyleClass
    ResetTask t
End Sub

' -----------------------------------------------------------------------------
' StartOrResume - create runner if needed and run the task
' -----------------------------------------------------------------------------
Public Sub StartOrResume(task As IaeLongProcessClass)
    On Error GoTo PROC_ERR
    If s_runner Is Nothing Then Set s_runner = New aeLongProcessClass
    s_runner.Run task
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure StartOrResume of Module basLongProcess"
    Resume PROC_EXIT
End Sub

' -----------------------------------------------------------------------------
' StopTask - signal the active runner to stop at next item boundary
' -----------------------------------------------------------------------------
Public Sub StopTask()
    If Not s_runner Is Nothing Then s_runner.StopTask
End Sub

' -----------------------------------------------------------------------------
' ResetTask - clear saved progress so next StartOrResume begins from item 1
' -----------------------------------------------------------------------------
Public Sub ResetTask(task As IaeLongProcessClass)
    On Error GoTo PROC_ERR
    If s_runner Is Nothing Then Set s_runner = New aeLongProcessClass
    s_runner.ResetTask task
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ResetTask of Module basLongProcess"
    Resume PROC_EXIT
End Sub

' -----------------------------------------------------------------------------
' SetWordHighPriority - opt-in WMI call to raise WINWORD.EXE priority
' Call manually before starting a long task if needed.
' -----------------------------------------------------------------------------
Public Sub SetWordHighPriority()
    On Error GoTo PROC_ERR
    Dim objWMIService As Object
    Dim colProcesses  As Object
    Dim objProcess    As Object
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'WINWORD.EXE'")
    For Each objProcess In colProcesses
        objProcess.SetPriority 128
    Next objProcess
PROC_EXIT:
    Set objWMIService = Nothing
    Set colProcesses = Nothing
    Set objProcess = Nothing
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure SetWordHighPriority of Module basLongProcess"
    Resume PROC_EXIT
End Sub

' -----------------------------------------------------------------------------
' UpdateCharacterStyle - legacy task stub (to be moved to aeUpdateCharStyleClass
' in Step 5 of the long-process framework implementation plan)
' -----------------------------------------------------------------------------
Public Sub UpdateCharacterStyle(Optional ByVal pageNumber As Integer = 0)
    On Error GoTo PROC_ERR
    Dim doc       As Document
    Dim para      As Word.Paragraph
    Dim rng       As Word.Range
    Dim StyleName As String
    Dim updateCount As Integer
    Dim startTime As Double
    Dim endTime   As Double
    Dim runTime   As Double
    Dim minutes   As Integer
    Dim seconds   As Integer

    SetWordHighPriority
    startTime = Timer

    If pageNumber = 0 Then
        Debug.Print "Page number required"
        GoTo PROC_EXIT
    End If

    Set doc = ActiveDocument
    StyleName = "Chapter Verse marker"
    updateCount = 0

    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=pageNumber
    Debug.Print "Starting at Page " & pageNumber

    For Each para In doc.Paragraphs
        If para.Range.Information(wdActiveEndPageNumber) >= pageNumber Then
            For Each rng In para.Range.Characters
                If rng.style = StyleName Then
                    rng.style = StyleName
                    updateCount = updateCount + 1
                    If updateCount >= 5000 Then
                        Debug.Print "Done 5000"
                        endTime = Timer
                        runTime = endTime - startTime
                        minutes = Int(runTime / 60)
                        seconds = Int(runTime Mod 60)
                        Debug.Print "Routine Runtime: " & Format(minutes, "00") & ":" & Format(seconds, "00") & " minutes and seconds"
                        GoTo PROC_EXIT
                    End If
                    DoEvents
                End If
            Next rng
        End If
    Next para
    Debug.Print "Done!"

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure UpdateCharacterStyle of Module basLongProcess"
    Resume PROC_EXIT
End Sub
