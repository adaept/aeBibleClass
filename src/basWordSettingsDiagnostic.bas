Attribute VB_Name = "basWordSettingsDiagnostic"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

' === WordSettingsDiagnostic.bas ===
' === Main entry point ===
Sub RunWordSettingsAudit(Optional saveToFile As Boolean = False)
    Dim currentSettings As Object
    Dim targetSettings As Object
    Dim discrepancies As Object
    Dim outputText As String

    ' Create all dictionaries using late binding
    Set currentSettings = GetCurrentWordSettings()
    Set targetSettings = LoadTargetBaseline()
    Set discrepancies = CompareSettings(currentSettings, targetSettings)

    outputText = FormatDiagnostics(currentSettings, targetSettings, discrepancies)

    ' Output to Immediate Window
    Debug.Print outputText

    ' Optional: Export to file
    If saveToFile Then
        SaveReportToFile outputText, "WordSettingsAudit.txt"
        MsgBox "Audit saved to WordSettingsAudit.txt", vbInformation
    End If
End Sub

' === Gather current settings into a Dictionary ===
Function GetCurrentWordSettings() As Object
    Dim settings As Object
    Set settings = CreateObject("Scripting.Dictionary")
    With Options
        settings.Add "EnableLivePreview", .EnableLivePreview
        'settings.Add "ShowPasteOptions", .ShowPasteOptions     - only available in Excel
        'settings.Add "AllowClickAndType", .AllowClickAndType   - not defined in the Word VBA Options object
        settings.Add "EnableSound", .EnableSound
        'settings.Add "CheckGrammarWithSpelling", .CheckGrammarWithSpelling - not defined in Word 365 VBA
        ' Editor-based grammar settings are not exposed via VBA
        settings.Add "GrammarCheckStatus", "Not accessible via VBA (Editor-based)"
        'settings.Add "ShowAllFormattingMarks", .ShowAllFormattingMarks     - not defined in Word VBA
        'settings.Add "ShowTextBoundaries", .ShowTextBoundaries             - not a member of the Options object in Word VBA
        settings.Add "ShowTextBoundaries", GetShowTextBoundaries()         '- workaround
        settings.Add "OptimizeForWord97", ActiveDocument.OptimizeForWord97
        settings.Add "SaveInterval", .SaveInterval
        settings.Add "BackgroundSave", .BackgroundSave
    End With
    Set GetCurrentWordSettings = settings
End Function

Function GetShowTextBoundaries() As Variant
    On Error Resume Next
    Dim result As Variant

    ' Only check if view is Print Layout or Web Layout
    Select Case ActiveWindow.View.Type
        Case wdPrintView, wdWebView
            result = ActiveWindow.View.ShowTextBoundaries
        Case Else
            result = "Unsupported view mode: " & ActiveWindow.View.Type
    End Select

    GetShowTextBoundaries = result
End Function

' === Define or Load a baseline (can be replaced with a loader from external file) ===
Function LoadTargetBaseline() As Object
    Dim baseline As Object
    Set baseline = CreateObject("Scripting.Dictionary")
    baseline.Add "EnableLivePreview", True
    'baseline.Add "ShowPasteOptions", True
    'baseline.Add "AllowClickAndType", True
    baseline.Add "EnableSound", False
    'baseline.Add "CheckGrammarWithSpelling", True
    'baseline.Add "ShowAllFormattingMarks", False
    baseline.Add "ShowTextBoundaries", False
    baseline.Add "OptimizeForWord97", False
    baseline.Add "SaveInterval", 10
    baseline.Add "BackgroundSave", True
    Set LoadTargetBaseline = baseline
End Function

' === Compare two sets of settings ===
Function CompareSettings(current As Object, target As Object) As Object
    Dim key As Variant
    Dim discrepancies As Object
    Set discrepancies = CreateObject("Scripting.Dictionary")

    ' Loop through keys in target dictionary
    For Each key In target.Keys
        If current.Exists(key) Then
            'Debug.Print ">" & key
            If current(key) <> target(key) Then
                discrepancies.Add key, "Current: " & current(key) & " | Expected: " & target(key)
            End If
        Else
            discrepancies.Add key, "Missing in current settings"
        End If
    Next key

    Set CompareSettings = discrepancies
End Function

' === Format diagnostics for output ===
Function FormatDiagnostics(current As Object, target As Object, issues As Object) As String
    Dim result As String
    result = "== Word 365 Diagnostic Audit ==" & vbCrLf
    result = result & "Date: " & Format(Now, "yyyy-mm-dd hh:nn") & vbCrLf & vbCrLf

    result = result & "== Discrepancies ==" & vbCrLf
    If issues.count = 0 Then
        result = result & "None. All settings match baseline." & vbCrLf
    Else
        Dim key As Variant
        For Each key In issues.Keys
            result = result & key & ": " & issues(key) & vbCrLf
        Next key
    End If
    result = result & vbCrLf & "== Full Current Settings ==" & vbCrLf
    For Each key In current.Keys
        result = result & key & ": " & current(key) & vbCrLf
    Next key

    FormatDiagnostics = result
End Function

' === Save report to file ===
Sub SaveReportToFile(reportText As String, fileName As String)
    Dim filePath As String
    filePath = ThisDocument.Path & "\" & fileName
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, reportText
    Close #fileNum
End Sub
