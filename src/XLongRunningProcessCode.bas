Attribute VB_Name = "XLongRunningProcessCode"
'==============================================================================
' XLongRunningProcessCode — Long-Running Process Utilities (DEFERRED)
' ----------------------------------------------------------------------------
' X-prefix convention: excluded from the normal test run. Contains routines
' that take significant time to complete (WMI priority setting, per-character
' style updates, progress tracking via CustomDocumentProperties).
' Run manually from the Immediate Window only.
'==============================================================================
Option Explicit
Option Compare Text
Option Private Module

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim lastProcessedParagraph As Long
Dim continueProcessing As Boolean
Dim progressPercentage As Double
       
Sub PauseWithDoEvents(milliseconds As Long)
' Combining `Sleep` with `DoEvents` can help keep the application responsive
    Dim startTime As Single
    startTime = Timer
    Do While Timer < startTime + (milliseconds / 1000)
        DoEvents
        Sleep 10 ' Short sleep to keep the application responsive
    Loop
End Sub
    
Sub StartOrResumeUpdate()
    continueProcessing = True
    LoadProgress
    LongProcessSkeletonWithConsoleProgress
End Sub

Sub StopUpdate()
    continueProcessing = False
    SaveProgress
End Sub

Sub ResetProgress()
    lastProcessedParagraph = 0
    progressPercentage = 0
    SaveProgress
End Sub

Sub SaveProgress()
    On Error GoTo PROC_ERR
    Dim props As Object
    Set props = ActiveDocument.CustomDocumentProperties
    If Not CustomPropertyExists(props, "LastProcessedParagraph") Then
        props.Add "LastProcessedParagraph", False, 3, lastProcessedParagraph
    Else
        props("LastProcessedParagraph").value = lastProcessedParagraph
    End If
    If Not CustomPropertyExists(props, "ProgressPercentage") Then
        props.Add "ProgressPercentage", False, 3, progressPercentage
    Else
        props("ProgressPercentage").value = progressPercentage
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    Debug.Print "SaveProgress ERROR " & Err.Number & ": " & Err.Description
    Resume PROC_EXIT
End Sub

Private Function CustomPropertyExists(props As Object, ByVal propName As String) As Boolean
    Dim p As Object
    On Error Resume Next
    Set p = props(propName)
    CustomPropertyExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Sub LoadProgress()
    On Error Resume Next
    lastProcessedParagraph = ActiveDocument.CustomDocumentProperties("LastProcessedParagraph").value
    progressPercentage = ActiveDocument.CustomDocumentProperties("ProgressPercentage").value
    On Error GoTo 0
End Sub

Sub SetWordHighPriority()
    Dim objWMIService As Object
    Dim colProcesses As Object
    Dim objProcess As Object
    Dim strComputer As String
    Dim processName As String

    On Error GoTo PROC_ERR

    ' Set the computer and process name
    strComputer = "."
    processName = "WINWORD.EXE"

    ' Get the WMI service
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

    ' Get the processes with the specified name
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & processName & "'")

    ' Loop through the processes and set the priority to high
    For Each objProcess In colProcesses
        objProcess.SetPriority 128 ' 128 is the value for high priority
    Next objProcess

PROC_EXIT:
    ' Clean up
    Set objWMIService = Nothing
    Set colProcesses = Nothing
    Set objProcess = Nothing
    Exit Sub

PROC_ERR:
    Debug.Print "SetWordHighPriority ERROR " & Err.Number & ": " & Err.Description
    Resume PROC_EXIT
End Sub

Sub UpdateCharacterStyle(Optional ByVal pageNumber As Integer = 0)
    Dim doc As Document
    Dim para As Word.Paragraph
    Dim rng As Word.Range
    Dim styleName As String
    Dim updateCount As Integer
    Dim startTime As Double
    Dim endTime As Double
    Dim runTime As Double
    Dim minutes As Integer
    Dim seconds As Integer

    ' Set Word to high priority
    SetWordHighPriority

    ' Record the start timer for each test
    startTime = Timer

    If IsMissing(pageNumber) Then
        Debug.Print "Page number required"
        Exit Sub
    End If
    
    ' Set the document and style name
    Set doc = ActiveDocument
    styleName = "Chapter Verse marker" ' Replace with your character style name
    updateCount = 0

    ' Move the selection to the specified page number
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, count:=pageNumber
    Debug.Print "Starting at Page " & pageNumber

    ' Loop through each paragraph in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph is on or after the specified page
        If para.Range.Information(wdActiveEndPageNumber) >= pageNumber Then
            ' Process the paragraph (example: update character style)
            ' Loop through each range in the paragraph
            For Each rng In para.Range.Characters
                ' Check if the range has the specified character style
                If rng.style = styleName Then
                    ' Apply the style from the style gallery
                    rng.style = styleName
                    updateCount = updateCount + 1
                    ' Stop after the some number of updates
                    If updateCount >= 5000 Then
                        Debug.Print "Done 5000"
                        endTime = Timer
                        runTime = endTime - startTime
                        ' Convert elapsed time to minutes and seconds
                        minutes = Int(runTime / 60)
                        seconds = Int(runTime Mod 60)
                        Debug.Print "Routine Runtime: " & Format(minutes, "00") & ":" & Format(seconds, "00") & " minutes and seconds"
                        Exit Sub
                    End If
                    DoEvents ' Keep the application responsive
                End If
            Next rng
        End If
    Next para
    Debug.Print "Done!"
End Sub

Sub LongProcessSkeletonWithConsoleProgress()
    Dim doc As Document
    Set doc = ActiveDocument
        
    Dim totalParagraphs As Long
    totalParagraphs = doc.Paragraphs.count
        
    Dim batchSize As Long
    batchSize = 50 ' Number of paragraphs to update in each phase
        
    Dim startIndex As Long
    Dim endIndex As Long
    Dim i As Long
        
            
    ' Update the rest of the document in phases
    If lastProcessedParagraph = 0 Then lastProcessedParagraph = 1 ' Start from the beginning if not previously set
            
    For startIndex = lastProcessedParagraph To totalParagraphs Step batchSize
        endIndex = startIndex + batchSize - 1
        If endIndex > totalParagraphs Then endIndex = totalParagraphs
                
        Application.ScreenUpdating = False
        Options.Pagination = False
                
        For i = startIndex To endIndex
            If Not continueProcessing Then
                lastProcessedParagraph = i
                progressPercentage = (lastProcessedParagraph / totalParagraphs) * 100
                SaveProgress
                Exit Sub
            End If
            
            
            ' CODE GOES HERE
            
                    
            DoEvents ' Allow Word to process other events
        Next i
                
        Options.Pagination = True
        Application.ScreenUpdating = True
                
        ' Calculate and output progress to console
        progressPercentage = (endIndex / totalParagraphs) * 100
        Debug.Print "Progress: " & Format(progressPercentage, "0.00") & "%"
                
        ' Save progress
        lastProcessedParagraph = endIndex + 1
        SaveProgress
                
        ' Pause between phases to allow Word to catch up
        PauseWithDoEvents (60000) ' 1000 milliseconds = 1 second
    Next startIndex
            
    Debug.Print "Style update complete!"
End Sub
' This updated script saves the progress to custom document properties, ensuring that progress is remembered even after a computer restart
' When you start or resume the update, it loads the progress from these properties.

