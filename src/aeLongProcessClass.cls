VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aeLongProcessClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

' =============================================================================
' aeLongProcessClass - Long-Running Process Runner
' -----------------------------------------------------------------------------
' Executes any IaeLongProcessClass task in batches with DoEvents between Items
' to keep Word responsive. Progress is persisted to a rpt/ sidecar file so a
' task can be stopped and resumed across sessions.
'
' Usage (Immediate Window):
'   Dim t As New aeUpdateCharStyleClass   ' or any IaeLongProcessClass
'   Dim r As New aeLongProcessClass
'   r.Run t
'
' To stop a running task call StopTask on the same instance, or use the
' basLongProcess.StopTask entry point which delegates to the active runner.
' =============================================================================

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private m_continueProcessing As Boolean
Private m_lastProcessedItem  As Long
Private m_progressPct        As Double

Public BatchSize As Long   ' Items per batch  (default 50)
Public PauseMs   As Long   ' pause between batches in ms (default 1000)

Private Sub Class_Initialize()
    BatchSize = 50
    PauseMs = 1000
End Sub

' -----------------------------------------------------------------------------
' Run - execute the task in batches
' -----------------------------------------------------------------------------
Public Sub Run(task As IaeLongProcessClass)
    On Error GoTo PROC_ERR
    m_continueProcessing = True
    LoadProgress task.TaskName

    Dim log As New aeLoggerClass
    Dim logStarted As Boolean
    logStarted = False
    If ActiveDocument.Path <> "" Then
        log.Log_Init ActiveDocument.Path & "\rpt\LongProcess_" & task.TaskName & ".txt"
        logStarted = True
    End If

    Dim totalItems As Long
    totalItems = task.ItemCount

    If logStarted Then log.Log_Write "Run started: " & task.TaskName & "  Items=" & totalItems & "  BatchSize=" & BatchSize

    If totalItems = 0 Then
        Debug.Print "aeLongProcessClass.Run: no Items to process"
        If logStarted Then log.Log_Write "No Items to process - exiting"
        GoTo PROC_EXIT
    End If

    If m_lastProcessedItem = 0 Then m_lastProcessedItem = 1
    If logStarted Then log.Log_Write "Resuming from item " & m_lastProcessedItem

    Dim startIndex As Long
    Dim endIndex   As Long
    Dim i          As Long

    For startIndex = m_lastProcessedItem To totalItems Step BatchSize
        endIndex = startIndex + BatchSize - 1
        If endIndex > totalItems Then endIndex = totalItems

        ' Note: ScreenUpdating and Options.Pagination are intentionally NOT modified
        ' here. Setting Options.Pagination = False raises error 4605 if a Find/Replace
        ' or other dialog is open when DoEvents yields to the user. Individual task
        ' classes may manage their own Word application state if required.

        For i = startIndex To endIndex
            If Not m_continueProcessing Then
                m_lastProcessedItem = i
                m_progressPct = (m_lastProcessedItem / totalItems) * 100
                SaveProgress task.TaskName
                If logStarted Then log.Log_Write "Stopped at item " & i & " (" & Format(m_progressPct, "0.00") & "%)"
                GoTo PROC_EXIT
            End If

            If Not task.ExecuteItem(i) Then
                m_lastProcessedItem = i
                m_progressPct = (m_lastProcessedItem / totalItems) * 100
                SaveProgress task.TaskName
                If logStarted Then log.Log_Write "Task requested stop at item " & i
                GoTo PROC_EXIT
            End If

            DoEvents
        Next i

        m_progressPct = (endIndex / totalItems) * 100
        m_lastProcessedItem = endIndex + 1
        SaveProgress task.TaskName

        Debug.Print task.TaskName & " progress: " & Format(m_progressPct, "0.00") & "%"
        If logStarted Then log.Log_Write "Batch complete: " & endIndex & "/" & totalItems & " (" & Format(m_progressPct, "0.00") & "%)"

        PauseWithDoEvents PauseMs
    Next startIndex

    m_lastProcessedItem = 0
    m_progressPct = 100
    SaveProgress task.TaskName
    Debug.Print task.TaskName & ": complete - " & totalItems & " Items processed"
    If logStarted Then log.Log_Write "Complete: " & totalItems & " Items processed"

PROC_EXIT:
    If logStarted Then log.Log_Close
    Exit Sub
PROC_ERR:
    If logStarted Then log.Log_Close
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Run of Class aeLongProcessClass"
    Resume PROC_EXIT
End Sub

' -----------------------------------------------------------------------------
' StopTask - signals the batch loop to stop at the next item boundary
' -----------------------------------------------------------------------------
Public Sub StopTask()
    m_continueProcessing = False
End Sub

' -----------------------------------------------------------------------------
' ResetTask - clears saved progress so the next Run starts from item 1
' -----------------------------------------------------------------------------
Public Sub ResetTask(task As IaeLongProcessClass)
    On Error GoTo PROC_ERR
    m_lastProcessedItem = 0
    m_progressPct = 0
    SaveProgress task.TaskName
    Debug.Print task.TaskName & ": progress reset"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ResetTask of Class aeLongProcessClass"
    Resume PROC_EXIT
End Sub

' -----------------------------------------------------------------------------
' SaveProgress - write progress to rpt/ sidecar file (key=value format)
' -----------------------------------------------------------------------------
Private Sub SaveProgress(ByVal TaskName As String)
    On Error GoTo PROC_ERR
    If ActiveDocument.Path = "" Then GoTo PROC_EXIT
    Dim fPath As String
    fPath = ActiveDocument.Path & "\rpt\LongRunningProgress_" & TaskName & ".txt"
    Dim fNum As Integer
    fNum = FreeFile
    Open fPath For Output As #fNum
    Print #fNum, "LastProcessedItem=" & m_lastProcessedItem
    Print #fNum, "ProgressPercentage=" & m_progressPct
    Close #fNum
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure SaveProgress of Class aeLongProcessClass"
    Resume PROC_EXIT
End Sub

' -----------------------------------------------------------------------------
' LoadProgress - read progress from rpt/ sidecar file if it exists
' -----------------------------------------------------------------------------
Private Sub LoadProgress(ByVal TaskName As String)
    On Error GoTo PROC_ERR
    m_lastProcessedItem = 0
    m_progressPct = 0
    If ActiveDocument.Path = "" Then GoTo PROC_EXIT
    Dim fPath As String
    fPath = ActiveDocument.Path & "\rpt\LongRunningProgress_" & TaskName & ".txt"
    If Dir$(fPath) = "" Then GoTo PROC_EXIT
    Dim fNum As Integer
    Dim oneLine As String
    Dim eqPos As Long
    Dim key As String
    Dim val As String
    fNum = FreeFile
    Open fPath For Input As #fNum
    Do While Not EOF(fNum)
        Line Input #fNum, oneLine
        eqPos = InStr(oneLine, "=")
        If eqPos > 0 Then
            key = Left$(oneLine, eqPos - 1)
            val = Mid$(oneLine, eqPos + 1)
            If key = "LastProcessedItem" Then m_lastProcessedItem = CLng(val)
            If key = "ProgressPercentage" Then m_progressPct = CDbl(val)
        End If
    Loop
    Close #fNum
    Debug.Print "LoadProgress: " & TaskName & "  LastProcessedItem=" & m_lastProcessedItem
PROC_EXIT:
    Exit Sub
PROC_ERR:
    If fNum > 0 Then Close #fNum
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LoadProgress of Class aeLongProcessClass"
    Resume PROC_EXIT
End Sub

' -----------------------------------------------------------------------------
' PauseWithDoEvents - sleep loop that keeps Word responsive
' -----------------------------------------------------------------------------
Private Sub PauseWithDoEvents(ByVal milliseconds As Long)
    Dim startTime As Single
    startTime = Timer
    Do While Timer < startTime + (milliseconds / 1000)
        DoEvents
        Sleep 10
    Loop
End Sub
