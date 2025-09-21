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
        settings.Add "EnableSound", .EnableSound
        settings.Add "SaveInterval", .SaveInterval
        settings.Add "BackgroundSave", .BackgroundSave
    End With

    ' View-dependent workaround
    settings.Add "ShowTextBoundaries", GetShowTextBoundaries()

    ' Document-level setting
    settings.Add "OptimizeForWord97", ActiveDocument.OptimizeForWord97

    ' Editor settings - not exposed via VBA
    settings.Add "GrammarCheckStatus", "Not accessible via VBA (Editor-based)"

    ' Startup options - UI only
    settings.Add "ShowStartScreenOnLaunch", "Manual check: File > Options > General > Startup Options"
    settings.Add "MiniToolbarOnSelection", "Manual check: File > Options > General > User Interface Options"
    settings.Add "OpenUneditableFilesInReadingView", "Manual check: File > Options > General > Startup Options"
    settings.Add "ShowTellMeBox", "Manual check: File > Options > General > User Interface Options"
    settings.Add "EditorGrammarStyle", "Manual check: Home tab > Editor > Customize Suggestions"
    settings.Add "AutocorrectOptions", "Manual check: File > Options > Proofing > AutoCorrect Options"
    settings.Add "StyleSuggestions", "Manual check: Home tab > Editor > Customize Suggestions"
    settings.Add "FormatConsistencyChecker", "Manual check: File > Options > Advanced > Keep track of formatting"

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
    Dim key As Variant
    Const manualFlag As String = "[ ]" & " Manual check: "

    result = "== Word 365 Diagnostic Audit ==" & vbCrLf
    result = result & "Date: " & Format(Now, "yyyy-mm-dd hh:nn") & vbCrLf & vbCrLf

    result = result & "== Discrepancies ==" & vbCrLf
    If issues.count = 0 Then
        result = result & "None. All settings match baseline." & vbCrLf
    Else
        For Each key In issues.Keys
            result = result & key & ": " & issues(key) & vbCrLf
        Next key
    End If

    result = result & vbCrLf & "== Full Current Settings ==" & vbCrLf
    For Each key In current.Keys
        Select Case True
            Case InStr(current(key), "Manual check:") > 0
                result = result & "? " & key & ": " & current(key) & vbCrLf
            Case InStr(current(key), "Not accessible") > 0 Or InStr(current(key), "Unsupported") > 0
                result = result & "? " & key & ": " & current(key) & vbCrLf
            Case Else
                result = result & "[ ] " & key & ": " & FormatBoolean(current(key)) & vbCrLf
        End Select
    Next key

    result = result & vbCrLf & "== Manual UI Verifications ==" & vbCrLf
    For Each key In current.Keys
        If InStr(current(key), "File > Options") > 0 Or InStr(current(key), "Editor") > 0 Then
            result = result & manualFlag & key & " — " & current(key) & vbCrLf
        End If
    Next key

    FormatDiagnostics = result
End Function

Function FormatBoolean(value As Variant) As String
    If VarType(value) = vbBoolean Then
        FormatBoolean = IIf(value, "On", "Off")
    Else
        FormatBoolean = value
    End If
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

Sub ShowAllStyles()
    Dim s As style
    For Each s In ActiveDocument.Styles
        If s.Type = wdStyleTypeParagraph Or s.Type = wdStyleTypeCharacter Then
            Debug.Print "STYLE: " & s.NameLocal & _
                        " | InUse: " & s.InUse & _
                        " | QuickStyle: " & s.QuickStyle
        End If
    Next s
End Sub

Sub ShowMyStyles()
    Dim s As style
    Dim msg As String
    Dim styleCount As Integer

    msg = "Styles actively applied in body, headers, and footers:" & vbCrLf & vbCrLf
    styleCount = 0

    For Each s In ActiveDocument.Styles
        If s.Type = wdStyleTypeParagraph Or s.Type = wdStyleTypeCharacter Then
            If StyleIsAppliedAnywhere(s.NameLocal) Then
                styleCount = styleCount + 1
                msg = msg & s.NameLocal & vbTab & _
                      "QuickStyle=" & s.QuickStyle & vbCrLf

                Debug.Print "STYLE: " & s.NameLocal & _
                            " | QuickStyle=" & s.QuickStyle
            End If
        End If
    Next s

    If styleCount = 0 Then
        msg = "No styles matched usage in body or header/footer ranges."
        Debug.Print "INFO: No styles matched extended usage criteria."
    End If

    MsgBox msg, vbInformation, "Extended Style Audit"
End Sub

Function StyleIsAppliedAnywhere(sName As String) As Boolean
    Dim p As paragraph
    Dim sec As section

    On Error Resume Next

    ' Body paragraphs
    For Each p In ActiveDocument.paragraphs
        If p.style = sName Then
            StyleIsAppliedAnywhere = True
            Exit Function
        End If
    Next p

    ' Headers and footers
    For Each sec In ActiveDocument.Sections
        Dim hdrFtr As HeaderFooter
        For Each hdrFtr In sec.Headers
            For Each p In hdrFtr.range.paragraphs
                If p.style = sName Then
                    StyleIsAppliedAnywhere = True
                    Exit Function
                End If
            Next p
        Next hdrFtr
        For Each hdrFtr In sec.Footers
            For Each p In hdrFtr.range.paragraphs
                If p.style = sName Then
                    StyleIsAppliedAnywhere = True
                    Exit Function
                End If
            Next p
        Next hdrFtr
    Next sec

    On Error GoTo 0
End Function

Function StyleIsApplied(sName As String) As Boolean
    Dim p As paragraph
    On Error Resume Next
    For Each p In ActiveDocument.paragraphs
        If p.style = sName Then
            StyleIsApplied = True
            Exit Function
        End If
    Next p
    On Error GoTo 0
End Function

Sub HideUnusedStyles()
    Dim s As style
    For Each s In ActiveDocument.Styles
        If s.Type = wdStyleTypeParagraph Or s.Type = wdStyleTypeCharacter Then
            If Not s.InUse Then
                On Error Resume Next
                s.QuickStyle = False ' Hide from Ribbon gallery only
                On Error GoTo 0
            End If
        End If
    Next s
    MsgBox "Quick Style Gallery cleaned. Pane visibility cannot be modified via VBA.", vbInformation
End Sub




