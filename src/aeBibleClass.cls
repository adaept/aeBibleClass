VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aeBibleClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'Copyright (c) 2025 Peter F. Ennis
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation;
'version 3.0.
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'Lesser General Public License for more details.
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, visit
'http://www.gnu.org/licenses/lgpl-3.0.txt

'=================================================================================
' Author:   Peter F. Ennis
' Date:     February 16, 2025
' Comment:  Create class for QA of word Bible (REV)
' History:  See comment details, basChangeLogBibleClass, commit messages on GitHub
' GitHub:   https://github.com/adaept/BibleClass
'=================================================================================

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)

Private BibleClassVERSION As String
Private BibleClassVERSION_DATE As String

Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100
'
' Define the word constant in this class module
Private Const wdNotThemeColor As Long = -1
' Default use is to not time all tests individually
Private Const bTimeAllTests As Boolean = True
' Used to store Expected results
Private oneBasedExpectedArray As Variant
Private SkippedTestNum As Integer
Private Const MaxTests = 68
Private ResultArray(1 To MaxTests) As Long
Private GetPassFailArray(1 To MaxTests) As String
Private TestTimingArray(0 To MaxTests) As Variant
Private TestSessionID As String
Private result As Integer
Private TestReportFileName As String
Private testFileNum As Integer
Private Const TestReportFlag As Boolean = True
Private OneTest As Integer  ' Stores the value to run a single test
Private SkipTestArray As Variant       ' Used for heavy tests to skip
Private Const AposCP As Long = &H2019  ' Apostrophe Code Point
Private Apostrophe As String
Private ContractionArrayValues As Variant
Private contra As String               ' Contraction
Private Const CARTS As Integer = 52    ' Contraction Array RunTest Start
Private UnicodeString As String        ' A string of up to 3 code points
Private UnicodeLabelU As String        ' The UnicodeString label in U+ form
'==================================================================

Private Sub Class_Initialize()
    On Error GoTo 0
    'MsgBox Application.VBE.ActiveVBProject.Name
    'MsgBox Application.Caption
    If Application.VBE.ActiveVBProject.name = "BibleClass" Then
        Application.Caption = Application.VBE.ActiveVBProject.name & " " & BibleClassVERSION
    End If
    Debug.Print "Class_Initialize"
    BibleClassVERSION = "0.1.2"
    BibleClassVERSION_DATE = "January 31, 2026"
    ' Reset OneTest for each run
    OneTest = 0
    SkippedTestNum = 0
    Apostrophe = ChrW$(AposCP)
End Sub

Private Sub Class_Terminate()
    On Error GoTo 0
    Debug.Print
    Debug.Print "Class_Terminate"
    Debug.Print , Application.VBE.ActiveVBProject.name
    Debug.Print , "BibleClass VERSION: " & BibleClassVERSION
    Debug.Print , "BibleClass VERSION_DATE: " & BibleClassVERSION_DATE
End Sub

Public Property Get TheBibleClassTests(Optional ByVal varDebug As Variant) As Integer

    On Error GoTo PROC_ERR
    
    If Not CheckShowHideStatus Then
        MsgBox "Show/Hide is NOT set!", vbCritical, "HALT!"
        Stop
    End If

    'MsgBox "TheBibleClassTests"
    If IsMissing(varDebug) Then
        Debug.Print "Get TheBibleClassTests"
        Debug.Print , "varDebug IS missing so no parameter is passed to RunBibleClassTests"
        Debug.Print , "DEBUGGING IS OFF"
        TheBibleClassTests = RunBibleClassTests
    ElseIf varDebug = "varDebug" Then
        Debug.Print "Get TheBibleClassTests"
        Debug.Print , "varDebug = " & """varDebug""" & " and IS NOT missing so blnDebug is set to True"
        Debug.Print , "DEBUGGING TURNED ON"
        TheBibleClassTests = RunBibleClassTests(varDebug)
    ElseIf VarType(varDebug) = vbInteger Then
        Debug.Print "Get TheBibleClassTests"
        OneTest = varDebug
        Debug.Print ">Running Test OneTest = varDebug"  ', "varDebug = " & varDebug
        TheBibleClassTests = RunBibleClassTests(OneTest)
    Else
        Debug.Print "Unexpected Type Parameter for varDebug !!!"
    End If
    
PROC_EXIT:
    Exit Property

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TheBibleClassTests of Class BibleClass"
    Resume PROC_EXIT

End Property

Function CheckShowHideStatus() As Boolean
    CheckShowHideStatus = ActiveWindow.View.ShowAll
End Function

Private Sub InitializeGlobalResultArrayToMinusOne()

    On Error GoTo PROC_ERR

    Dim i As Integer, j As Integer, k As Integer
    For i = 1 To MaxTests
        ResultArray(i) = -1
    Next i
    ' Session ID to be stored in TestTimingArray(0)
    For j = 0 To MaxTests
        TestTimingArray(j) = -1
    Next j
    ' Initialize GetPassFailArray
    For k = 1 To MaxTests
        GetPassFailArray(k) = "x1x1x1x1"  ' Possible values are "PASS    ", "FAIL!!!!", "SKIP!!!!"
    Next k
    ' Optional: Print to Immediate window to confirm
    'For i = 1 To MaxTests
    '    Debug.Print "ResultArray(" & i & ") = " & ResultArray(i),
    'Next i

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure InitializeGlobalResultArrayToMinusOne of Class BibleClass"
    Resume PROC_EXIT
End Sub

Private Function ConvertToOneBasedArray(values As Variant) As Variant
    Dim strArray() As String
    Dim i As Integer

    On Error GoTo PROC_ERR

    ' Resize the array to start from index 1
    ReDim strArray(1 To UBound(values) - LBound(values) + 1)
    
    ' Store the values in the array starting from index 1
    For i = 1 To UBound(values) - LBound(values) + 1
        strArray(i) = values(i - 1 + LBound(values))
    Next i
    
    ' Return the 1-based array
    ConvertToOneBasedArray = strArray

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ConvertToOneBasedArray of Class BibleClass"
    Resume PROC_EXIT
End Function

Private Sub Expected1BasedArray()
    Dim doc As Document
    Dim values As Variant
    Dim i As Integer
    Dim outputString As String

    On Error GoTo PROC_ERR

    Set doc = ActiveDocument
    If Len(doc.Path) = 0 Then
        MsgBox "ActiveDocument has not been saved. Cannot set TestReportFileName.", vbCritical
        Stop
        Exit Sub
    End If

    If OneTest = 0 Then
        ' Do not make new TestReportFileName for single tests
        Set doc = ActiveDocument
        TestReportFileName = doc.Path & "\rpt\TestReport.txt"
        ' Open the debug file for writing - this will create a new file each time
        testFileNum = FreeFile
        Open TestReportFileName For Output As testFileNum
        Close testFileNum
    End If
    
    ' Define the Expected RunTest result values to store in the array
    '      RunTest 1  2  3  4  5  6  7  8  9  10 11 12 13 14 15 16     17 18  19 20 21 22   23 24    25 26 27   28  29  30 31 32   33   34   35   36 37  38   39 40 41 42 43 44 45 46 47 48 49  50   51    52 53 54 55 56 57 58 59 60 61 62 63 64 65 66 67 68)
    values = Array(0, 0, 0, 0, 0, 0, 0, 0, 7, 1, 0, 0, 0, 0, 0, 16471, 0, 66, 0, 0, 0, 153, 0, 1000, 1, 0, 147, 77, 70, 0, 1, 982, 982, 117, 117, 0, 19, 153, 0, 0, 0, 0, 0, 0, 0, 0, 3, 0, 16, 147, 1189, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0)

    ' Convert to a 1-based array
    oneBasedExpectedArray = ConvertToOneBasedArray(values)

    ' Output the values in the 1-based array
    'For i = 1 To UBound(oneBasedExpectedArray)
    '    Debug.Print "Index " & i & ": " & oneBasedExpectedArray(i)
    'Next i

    ' Access the array elements to write a string output of 15 expected results on one line
    For i = LBound(oneBasedExpectedArray) To 15
        outputString = outputString & " Test(" & i & ")" & oneBasedExpectedArray(i) & " "
        ' Trim the trailing space
        outputString = Trim(outputString)
    Next i
    Debug.Print outputString
    If TestReportFlag And OneTest = 0 Then AppendToFile TestReportFileName, outputString
    
    ' Next 15 results
    outputString = ""
    For i = 16 To 30
        outputString = outputString & " Test(" & i & ")" & oneBasedExpectedArray(i) & " "
        ' Trim the trailing space
        outputString = Trim(outputString)
    Next i
    Debug.Print outputString
    If TestReportFlag And OneTest = 0 Then AppendToFile TestReportFileName, outputString
    
    ' Next 15 results
    outputString = ""
    For i = 31 To 45
        outputString = outputString & " Test(" & i & ")" & oneBasedExpectedArray(i) & " "
        ' Trim the trailing space
        outputString = Trim(outputString)
    Next i
    Debug.Print outputString
    If TestReportFlag And OneTest = 0 Then AppendToFile TestReportFileName, outputString

    ' Next 15 results
    outputString = ""
    For i = 46 To 60
        outputString = outputString & " Test(" & i & ")" & oneBasedExpectedArray(i) & " "
        ' Trim the trailing space
        outputString = Trim(outputString)
    Next i
    Debug.Print outputString
    If TestReportFlag And OneTest = 0 Then AppendToFile TestReportFileName, outputString

    ' Next results
    outputString = ""
    For i = 61 To UBound(oneBasedExpectedArray)
        outputString = outputString & " Test(" & i & ")" & oneBasedExpectedArray(i) & " "
        ' Trim the trailing space
        outputString = Trim(outputString)
    Next i
    Debug.Print outputString
    If TestReportFlag And OneTest = 0 Then AppendToFile TestReportFileName, outputString

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Expected1BasedArray of Class BibleClass"
    Resume PROC_EXIT
End Sub

'==================================================================
' Manually update with values of tests that are to be skipped
'==================================================================
Private Sub MakeSkipTestArray()
    SkipTestArray = Array(42, 51)
End Sub

Private Function IsSkipTest(ByVal n As Long) As Boolean
    Dim i As Long
    For i = LBound(SkipTestArray) To UBound(SkipTestArray)
        If SkipTestArray(i) = n Then
            IsSkipTest = True
            Exit Function
        End If
    Next i
End Function

Private Function HexToUnicodeLabel(cp As Long) As String
    HexToUnicodeLabel = "U+" & Right$("0000" & Hex$(cp), 4)
End Function

Private Function MakeUnicodeSeq(cp1 As Long, Optional cp2 As Long = 0, Optional cp3 As Long = 0) As String
    Dim s As String
    s = ChrW(cp1)
    UnicodeLabelU = HexToUnicodeLabel(cp1) & ","
    'Debug.Print "1: " & UnicodeLabelU
    If cp2 <> 0 Then
        s = s & ChrW(cp2)
        UnicodeLabelU = UnicodeLabelU & HexToUnicodeLabel(cp2) & ","
        'Debug.Print "2: " & UnicodeLabelU
    End If
    If cp3 <> 0 Then
        s = s & ChrW(cp3)
        UnicodeLabelU = UnicodeLabelU & HexToUnicodeLabel(cp3)
        'Debug.Print "3"; " & UnicodeLabelU"
    End If
    MakeUnicodeSeq = s
End Function

Function ProcessUnicode(s As String) As String
    ProcessUnicode = s   ' return the string exactly as received
End Function

Private Function ContractionArrayU(idx As Integer) As String
    ' Replace `'` with unicode value in the string, Apostrophe, =ChrW$(AposCP), &H2019, RIGHT SINGLE QUOTATION MARK
    
    'Debug.Print ">ContractionArrayValues(idx) = " & ContractionArrayValues(idx)
    ContractionArrayU = Replace(ContractionArrayValues(idx), "'", Apostrophe, 1, 1)
    'Debug.Print ">ContractionArrayU) = " & ContractionArrayU
End Function

Private Sub CreateContractionArray()
    ContractionArrayValues = ConvertToOneBasedArray(Array( _
        "wouldn't", "isn't", "let's", "i'm", "hasn't", "didn't", "don't", "haven't", "weren't", "aren't", "can't", "shouldn't", "it's", _
        "spirit's", " '" _
        ))
    ' Now: LBound(values) = 1, UBound(values) = 13, values(1) = "wouldn't", values(15) = " '"
End Sub

'=====================================================================================
' AppendToFile - Centralized diagnostic output writer for Word macro sessions
'
' Description:
'   Appends a line of text to a target file located in the 'rpt' subfolder of the
'   active document's directory. If only a filename is provided, it auto-prepends
'   the correct path. Ensures folder exists and handles full file paths if needed.
'
' Parameters:
'   filePath (String) - Either:
'       (a) Just a filename (e.g., "Log.txt"): saved in rpt folder
'       (b) Full path override: used as-is
'   text (String) - The line of text to be appended (no trailing newline needed)
'
' Behavior:
'   - Auto-creates 'rpt' folder if missing
'   - Supports relative or full-path input
'   - Opens file in append mode; adds a newline after each Print
'
' Safety:
'   - Preserves existing content
'   - Validates file path integrity
'   - Can be called repeatedly without conflict
'
' Author: Peter
' Version: 1.0
'=====================================================================================
Sub AppendToFile(filePath As String, text As String)
    Dim fileNum As Integer
    Dim folderPath As String
    Dim finalPath As String

    folderPath = ActiveDocument.Path & "\rpt"
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

    If InStr(filePath, "\") = 0 Then
        finalPath = folderPath & "\" & filePath
    Else
        finalPath = filePath
    End If

    fileNum = FreeFile
    Open finalPath For Append As fileNum
    Print #fileNum, text
    Close fileNum
End Sub

'============================================================
' Routine: LogMessage
' Purpose: Returns a timestamped message string for logging.
' Inputs:
'   - msg: The message content to be timestamped.
' Returns:
'   - String formatted as "[YYYY-MM-DD HH:NN:SS] message"
' Notes:
'   - Used for diagnostics or trace logging.
'============================================================
Private Function LogMessage(msg As String) As String
    LogMessage = "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "] " & msg
End Function

'============================================================
' Routine: GenerateSessionID
' Purpose: Creates a unique session identifier based on current time.
' Returns:
'   - String formatted as "YYYYMMDD-HHNNSS"
' Notes:
'   - Used to tag test sessions for time-based traceability.
'============================================================
Private Function GenerateSessionID() As String
    ' Format: YYYYMMDD-HHMMSS
    GenerateSessionID = Format(Now, "yyyymmdd-hhnnss")
End Function

'============================================================
' Routine: RunTotalTimeTestSession
' Purpose: Runs a timing test and logs outcome with session ID.
' Inputs:
'   - testOutcome: String summary of timing result.
' Actions:
'   - Generates a unique SessionID.
'   - Ensures a \rpt folder exists in document path.
'   - Creates or appends to TotalTimeReport.txt.
'   - Logs header if new file.
'   - Appends session outcome to file using AppendToFile.
'   - Outputs session ID to immediate window.
' Notes:
'   - External dependency: AppendToFile routine must exist.
'============================================================
Private Sub RunTotalTimeTestSession(testOutcome As String)
    Dim sessionID As String
    sessionID = GenerateSessionID()

    Dim folderPath As String
    folderPath = ActiveDocument.Path & "\rpt"
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

    Dim reportName As String
    reportName = "TotalTimeReport.txt"
    
    Dim FileExists As Boolean
    FileExists = (Dir(folderPath & "\" & reportName) <> "")
    'Debug.Print "FileExists = " & FileExists, "reportName = " & reportName
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open folderPath & "\" & reportName For Append As #fileNum
    If Not FileExists Then
        Print #fileNum, "SessionID, TotalTime"
    End If
    Close #fileNum

    Call AppendToFile("TotalTimeReport.txt", sessionID & ", " & testOutcome)
    
    'MsgBox "Test complete. Session ID: " & sessionID, vbInformation
    Debug.Print "RunTotalTimeTestSession complete. Session ID: " & sessionID
End Sub

Private Sub DebugAndReportHeader()
    Debug.Print "1234567890123412345678901234123456789012341234567890123412345678901234"
    '            123456789    1234567    12345678    123456    12345678    12345678     NOTE: Comma separation of debug is 14 characters
    Debug.Print "Pass/Fail", "Copy ()", "Test Num", "Result", "Expected", "Function"
End Sub

Private Function RunBibleClassTests(Optional ByVal varDebug As Variant) As Boolean
    Dim startTime As Double
    Dim endTime As Double
    Dim runTime As Double
    
    ' Record the start time
    startTime = Timer

    'MsgBox "RunBibleClassTests"
    On Error GoTo PROC_ERR

    Dim objComponent As Object
    Set objComponent = ThisDocument.VBProject.VBComponents

    Debug.Print "RunBibleClassTests"
    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so blnDebug of RunBibleClassTests is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    ElseIf varDebug = "varDebug" Then
        Debug.Print , "varDebug = " & """varDebug""" & " and IS NOT missing so blnDebug is set to True"
        Debug.Print , "NOW DEBUGGING..."
    ElseIf VarType(varDebug) = vbInteger And OneTest <> 0 Then
        Debug.Print ">>Running Test " '& OneTest, "varDebug = " & varDebug
    Else
    End If
    
    Dim response As VbMsgBoxResult
    
    ' Display the message box with "Yes" and "No" buttons
    response = MsgBox("Run Bible Class Tests", vbYesNo + vbQuestion, "RunBibleClassTests")
    Select Case response
        Case vbYes
            'MsgBox "You chose Yes.", vbInformation, "Response"
            On Error Resume Next
            Debug.Print ">>>RunBibleClassTests = YES"
            On Error GoTo PROC_ERR
                        
            ' Set up arrays for actual and expected results
            Call InitializeGlobalResultArrayToMinusOne
            Call MakeSkipTestArray
            Call Expected1BasedArray
            ' Set up array of contractions
            Call CreateContractionArray
            
            If OneTest <> 0 Then
                DebugAndReportHeader
                Call RunTest(OneTest)
                Exit Function
            End If
            
            DebugAndReportHeader
            If TestReportFlag And OneTest = 0 Then AppendToFile TestReportFileName, "1234567890123412345678901234123456789012341234567890123412345678901234"
            If TestReportFlag And OneTest = 0 Then AppendToFile TestReportFileName, "Pass/Fail" & space(5) & "Copy ()" & space(7) & "Test Num" & space(6) & "Result" & space(8) & "Expected" & space(6) & "Function"
            
            RunTest (1)
            RunTest (2)
            RunTest (3)
            RunTest (4)
            RunTest (5)
            RunTest (6)
            RunTest (7)
            RunTest (8)
            RunTest (9)
            RunTest (10)
            RunTest (11)
            RunTest (12)
            RunTest (13)
            RunTest (14)
            RunTest (15)
            RunTest (16)
            RunTest (17)
            RunTest (18)
            RunTest (19)
            RunTest (20)
            RunTest (21)
            RunTest (22)
            RunTest (23)
            RunTest (24)
            RunTest (25)
            RunTest (26)
            RunTest (27)
            RunTest (28)
            RunTest (29)
            RunTest (30)
            RunTest (31)
            RunTest (32)
            RunTest (33)
            RunTest (34)
            RunTest (35)
            RunTest (36)
            RunTest (37)
            RunTest (38)
            RunTest (39)
            RunTest (40)
            RunTest (41)
            If IsSkipTest(42) Then
                Call RunTest(42, "SKIP")
            Else
                RunTest (42)
            End If
            RunTest (43)
            RunTest (44)
            RunTest (45)
            RunTest (46)
            RunTest (47)
            RunTest (48)
            RunTest (49)
            RunTest (50)
            If IsSkipTest(51) Then
                Call RunTest(51, "SKIP")
            Else
                RunTest (51)
            End If
            RunTest (52)
            RunTest (53)
            RunTest (54)
            RunTest (55)
            RunTest (56)
            RunTest (57)
            RunTest (58)
            RunTest (59)
            RunTest (60)
            RunTest (61)
            RunTest (62)
            RunTest (63)
            RunTest (64)
            RunTest (65)
            RunTest (66)
            RunTest (67)
            RunTest (68)
        Case vbNo
            'MsgBox "You chose No.", vbInformation, "Response"
            Debug.Print ">>>RunBibleClassTests = NO"
            Exit Function
    End Select

    ' Record the end time
    endTime = Timer
    
    ' Calculate the runtime
    runTime = endTime - startTime
    
    ' Output the runtime in seconds and hundredths of a second
    'MsgBox "Runtime: " & Format(runTime, "0.00") & " seconds"
    Dim TheTotalRuntime As String
    TheTotalRuntime = Format(runTime, "0.00")
    Debug.Print "The Total Runtime: " & TheTotalRuntime & " seconds"
    If TestReportFlag And OneTest = 0 Then Call RunTotalTimeTestSession(TheTotalRuntime)
    
    If TestReportFlag And OneTest = 0 Then AppendToFile TestReportFileName, "The Total Runtime: " & TheTotalRuntime & " seconds"
    
    Debug.Print LogMessage("Completed")
    If TestReportFlag And OneTest = 0 Then AppendToFile TestReportFileName, LogMessage("Completed")

    Debug.Print "BibleClass VERSION: " & BibleClassVERSION
    If TestReportFlag And OneTest = 0 Then AppendToFile TestReportFileName, "BibleClass VERSION: " & BibleClassVERSION
    
    Debug.Print "BibleClass VERSION_DATE: " & BibleClassVERSION_DATE
    If TestReportFlag And OneTest = 0 Then AppendToFile TestReportFileName, "BibleClass VERSION_DATE: " & BibleClassVERSION_DATE

    Debug.Print LogWordBuildInfo
    If TestReportFlag And OneTest = 0 Then AppendToFile TestReportFileName, LogWordBuildInfo
    
    RunBibleClassTests = True

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RunBibleClassTests of Class BibleClass"
    Resume PROC_EXIT
End Function

Private Function LogWordBuildInfo() As String
    Dim info As String
    info = "== Word Build Info ==" & vbCrLf
    info = info & "Application Name: " & Application.name & vbCrLf
    info = info & "Platform Version: " & Application.Version & vbCrLf
    info = info & "Platform Build: " & Application.Build & vbCrLf
    info = info & "Full Version String: " & Application.name & " " & Application.Version & " (" & Application.Build & ")" & vbCrLf
    info = info & "Marketing Version: 2506 (Build 18925.20158)" & vbCrLf
    info = info & "Source: Microsoft 365 Current Channel, July 2025 update"
    'Debug.Print info
    'MsgBox info, vbInformation, "Word Version Details"
    LogWordBuildInfo = info
End Function

Private Function U(codepoint As Long) As String
    U = ChrW(codepoint)
End Function

Private Function GetPassFail(TestNum As Integer) As Boolean
    On Error GoTo PROC_ERR
    
    ' Set explicit skip if the test is skipped
    If IsSkipTest(TestNum) Then
        GetPassFailArray(TestNum) = "SKIP!!!!"
        Exit Function
    End If
    
    Select Case TestNum
    Case 1
        ResultArray(TestNum) = CountDoubleSpaces
    Case 2
        ResultArray(TestNum) = CountDoubleSpacesInShapes
    Case 3
        ResultArray(TestNum) = CountSpaceFollowedByCarriageReturn
    Case 4
        ResultArray(TestNum) = CountDoubleTabs
    Case 5
        ResultArray(TestNum) = CountWhiteSpaceAndCarriageReturn
    Case 6
        ResultArray(TestNum) = CountQuadrupleParagraphMarks
    Case 7
        ResultArray(TestNum) = CountStyleWithSpaceAndNumber
    Case 8
        ResultArray(TestNum) = CountStyleWithNumberAndSpace
    Case 9
        ResultArray(TestNum) = CountPeriodSpaceLeftParenthesis
    Case 10
        ResultArray(TestNum) = CountNonBreakingSpaces
    Case 11
        ResultArray(TestNum) = CountFindNumberDashNumber
    Case 12
        ResultArray(TestNum) = CountNumberDashNumberInFootnotes
    Case 13
        ResultArray(TestNum) = CountEmptyParasWithNoThemeColor
    Case 14
        ResultArray(TestNum) = CountWhiteParagraphMarks
    Case 15
        ResultArray(TestNum) = CountSectionsWithDifferentFirstPage
    Case 16
        ResultArray(TestNum) = CountTotalParagraphs
    Case 17
        ResultArray(TestNum) = CountRedFootnoteReferences
    Case 18
        ResultArray(TestNum) = CountHeading1
    Case 19
        ResultArray(TestNum) = CountEmptyParasAfterH2
    Case 20
        ResultArray(TestNum) = CountFootnotesFollowedByDigit
    Case 21
        ResultArray(TestNum) = CountNotSpacesAfterFootnoteReferences
    Case 22
        ResultArray(TestNum) = CountEmptyParagraphsWithFormatting
    Case 23
        ResultArray(TestNum) = CountDeleteEmptyParagraphsBeforeHeading2
    Case 24
        ResultArray(TestNum) = CountFootnoteReferences
    Case 25
        ResultArray(TestNum) = CountFootnoteReferenceColors
    Case 26
        ResultArray(TestNum) = CheckAllHeaders("Empty")
    Case 27
        ResultArray(TestNum) = CheckAllHeaders("NotEmpty")
    Case 28
        ResultArray(TestNum) = CountTabFollowedByParagraphMarkInHeaders
    Case 29
        ResultArray(TestNum) = CountParagraphsWithoutTabInHeaders
    Case 30
        ResultArray(TestNum) = CountHeaderStyleUsage
    Case 31
        ResultArray(TestNum) = CountParagraphMarksPerHeaderSection
    Case 32
        ResultArray(TestNum) = CountLinefeed
    Case 33
        ResultArray(TestNum) = CountLinefeed(" ")
    Case 34
        ResultArray(TestNum) = CountManualLineBreaksAndWithSpace
    Case 35
        ResultArray(TestNum) = CountManualLineBreaksAndWithSpace(" ")
    Case 36
        ResultArray(TestNum) = CountFooterParagraphsWithFooterStyle
    Case 37
        ResultArray(TestNum) = CountDocTabOnlyParagraphs
    Case 38
        ResultArray(TestNum) = CountEmptyParagraphs
    Case 39
        ResultArray(TestNum) = CountParagraphMarks_ArialBlack
    Case 40
        ResultArray(TestNum) = CountParagraphMarks_ArialBlackDarkRed
    Case 41
        ResultArray(TestNum) = Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed
    Case 42
        ResultArray(TestNum) = CountBoldFootnotesWordLevel
    Case 43
        ResultArray(TestNum) = CountDarkRedStyledParagraphMarks
    Case 44
        ResultArray(TestNum) = CountParagraphMarks_CalibriDarkRed
    Case 45
        ResultArray(TestNum) = CountFindNotEmphasisBlack
    Case 46
        ResultArray(TestNum) = CountFindNotEmphasisRed
    Case 47
        ResultArray(TestNum) = CountTabOnlyParagraphs("Footer")
    Case 48
        ResultArray(TestNum) = AuditLiberationSansNarrowStyleDetails
    Case 49
        ResultArray(TestNum) = CountAuditStyles_ToFile
    Case 50
        ResultArray(TestNum) = SummarizeHeaderFooterAuditToFile
    Case 51
        ResultArray(TestNum) = CountAndCreateDefinitionForH2
    Case 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66
        contra = ContractionArrayU(TestNum - CARTS + 1)
        ResultArray(TestNum) = CountContraction(contra)
    Case 67
        UnicodeString = ProcessUnicode(MakeUnicodeSeq(&H202F, &H2019))  ' NNBSP U+202F followed by RSQM U+2019
        ResultArray(TestNum) = CountContraction(UnicodeString)
    Case 68
        UnicodeString = ProcessUnicode(MakeUnicodeSeq(&H202F, &H201D))  ' NNBSP U+202F followed by RDQM U+201D
        ResultArray(TestNum) = CountContraction(UnicodeString)
    Case Else
        Debug.Print "TestNum = " & TestNum & " >The test number is outside the accepted range"
    End Select
    
    If ResultArray(TestNum) = oneBasedExpectedArray(TestNum) Then
        GetPassFailArray(TestNum) = "PASS    "
    Else
        GetPassFailArray(TestNum) = "FAIL!!!!"
    End If

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GetPassFail of Class BibleClass"
    Resume PROC_EXIT
    
End Function

'=========================================================================
' Routine : Fixed14Str
' Purpose : Pads the input string to a fixed width of 14 characters.
' Input   : MyResult - String to be padded.
' Output  : String of exactly 14 characters, right-padded with spaces.
' Notes   : Useful for column-aligned output or diagnostics.
'=========================================================================
Private Function Fixed14Str(MyResult As String) As String
    Const fixedWidth = 14
    Fixed14Str = Left(MyResult & String(fixedWidth, " "), fixedWidth)
    'Debug.Print "The fixed width string is:" & vbCrLf & "[" & Fixed14Str & "]"
    'Debug.Print "[12345678901234]"
End Function

Private Function RunTest(num As Integer, Optional SkipTest As Variant) As Boolean
    Dim startTime As Double
    Dim endTime As Double
    Dim runTime As Double
    
    ' Record the start timer for each test
    startTime = Timer
    
    On Error GoTo PROC_ERR

    GetPassFail (num)

    Select Case num
    Case 1
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountDoubleSpaces"
    Case 2
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountDoubleSpacesInShapes"
    Case 3
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountSpaceFollowedByCarriageReturn"
    Case 4
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountDoubleTabs"
    Case 5
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountWhiteSpaceAndCarriageReturn"
    Case 6
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountQuadrupleParagraphMarks"
    Case 7
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountStyleWithSpaceAndNumber"
    Case 8
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountStyleWithNumberAndSpace"
    Case 9
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountPeriodSpaceLeftParenthesis"
    Case 10
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountNonBreakingSpaces"
    Case 11
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountFindNumberDashNumber"
    Case 12
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountNumberDashNumberInFootnotes"
    Case 13
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountEmptyParasWithNoThemeColor"
    Case 14
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountWhiteParagraphMarks"
    Case 15
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountSectionsWithDifferentFirstPage"
    Case 16
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountTotalParagraphs"
    Case 17
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountRedFootnoteReferences"
    Case 18
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountHeading1"
    Case 19
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountEmptyParasAfterH2"
    Case 20
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountFootnotesFollowedByDigit"
    Case 21
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountNotSpacesAfterFootnoteReferences"
    Case 22
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountEmptyParagraphsWithFormatting"
    Case 23
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountDeleteEmptyParagraphsBeforeHeading2"
    Case 24
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountFootnoteReferences"
    Case 25
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountFootnoteReferenceColors"
    Case 26
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CheckAllHeaders(" & "Empty" & ")"
    Case 27
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CheckAllHeaders(" & "NotEmpty" & ")"
    Case 28
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountTabFollowedByParagraphMarkInHeaders"
    Case 29
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountParagraphsWithoutTabInHeaders"
    Case 30
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountHeaderStyleUsage"
    Case 31
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountParagraphMarksPerHeaderSection"
    Case 32
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountLinefeed"
    Case 33
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountLinefeed(" & """ """ & ")"
    Case 34
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountManualLineBreaksAndWithSpace"
    Case 35
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountManualLineBreaksAndWithSpace(" & """ """ & ")"
    Case 36
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountFooterParagraphsWithFooterStyle"
    Case 37
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountDocTabOnlyParagraphs"
    Case 38
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountEmptyParagraphs"
    Case 39
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountParagraphMarks_ArialBlack"
    Case 40
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountParagraphMarks_ArialBlackDarkRed"
    Case 41
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed"
    Case 42
        'Debug.Print ">> Expected Routine Runtime ~ 80 seconds"
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountBoldFootnotesWordLevel"
    Case 43
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountDarkRedStyledParagraphMarks"
    Case 44
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountParagraphMarks_CalibriDarkRed"
    Case 45
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountFindNotEmphasisBlack"
    Case 46
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountFindNotEmphasisRed"
    Case 47
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountTabOnlyParagraphs(" & """Footer""" & ")"
    Case 48
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "AuditLiberationSansNarrowStyleDetails"
    Case 49
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountAuditStyles_ToFile"
    Case 50
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "SummarizeHeaderFooterAuditToFile"
    Case 51
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountAndCreateDefinitionForH2"
    Case 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "U+2019 CountContraction(" & """" & ContractionArrayU(num - CARTS + 1) & """" & ")"
    Case 67, 68
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), UnicodeLabelU & " CountUnicodeSeq(" & """" & UnicodeString & """" & "])"    ' Using [] to delineate unicode string in test output
    Case Else
        Debug.Print "num = " & num & " >>The test number is outside the accepted range"
    End Select
    'Debug.Print "!!! OneTest = " & OneTest, "num = " & num

    If TestReportFlag And OneTest = 0 And IsMissing(SkipTest) Then
        'Debug.Print "A>", "num = " & num
        Call OutputTestReport(num)
    ElseIf TestReportFlag And OneTest = 0 And Not IsMissing(SkipTest) Then
        'Debug.Print "B>", "num = " & num, "SkipTest = " & SkipTest
        Call OutputTestReport(num, SkipTest)
        'Stop
    End If

    endTime = Timer
    runTime = endTime - startTime
    
    ' Output the runtime in seconds and hundredths of a second
    'MsgBox "Routine Runtime: " & Format(runTime, "0.00") & " seconds"
    If bTimeAllTests Then Debug.Print "Routine Runtime: " & Format(runTime, "0.00") & " seconds"

PROC_EXIT:
    Exit Function
    
PROC_ERR:
    Dim answer As Integer
    answer = MsgBox("Err = " & Err.Number & " " & Err.Description & vbCrLf & "Do you want to continue?", vbQuestion + vbYesNo, "Continue or Stop")
    If answer = vbYes Then
        'MsgBox "You chose to continue.", vbInformation, "Continue"
        Resume
    Else
        'MsgBox "You chose to stop.", vbInformation, "Stop"
        'Exit Sub
        'Stop
    End If
    Debug.Print "!!! Error in Test num = " & num, "Function RunTest !!!"
    If num = 1 Then
        Debug.Print ">(1)"
        Debug.Print GetPassFail(1)
        Debug.Print "num = " & num
        Debug.Print "ResultArray(1) = " & ResultArray(1)
        Debug.Print "oneBasedExpectedArray(1) = " & oneBasedExpectedArray
        Debug.Print GetPassFail(1), "Copy ()", "Test = " & num, ResultArray(1), oneBasedExpectedArray(1), "CountDoubleSpaces"
    End If
    Stop
    End
End Function

Private Function OutputTestReport(num As Integer, Optional SkipTest As Variant) As Boolean
    On Error GoTo PROC_ERR
    
    Select Case num
    Case 1
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountDoubleSpaces"
    Case 2
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountDoubleSpacesInShapes"
    Case 3
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountSpaceFollowedByCarriageReturn"
    Case 4
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountDoubleTabs"
    Case 5
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountWhiteSpaceAndCarriageReturn"
    Case 6
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountQuadrupleParagraphMarks"
    Case 7
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountStyleWithSpaceAndNumber"
    Case 8
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountStyleWithNumberAndSpace"
    Case 9
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountPeriodSpaceLeftParenthesis"
    Case 10
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountNonBreakingSpaces"
    Case 11
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountFindNumberDashNumber"
    Case 12
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountNumberDashNumberInFootnotes"
    Case 13
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountEmptyParasWithNoThemeColor"
    Case 14
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountWhiteParagraphMarks"
    Case 15
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountSectionsWithDifferentFirstPage"
    Case 16
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountTotalParagraphs"
    Case 17
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountRedFootnoteReferences"
    Case 18
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountHeading1"
    Case 19
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountEmptyParasAfterH2"
    Case 20
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountFootnotesFollowedByDigit"
    Case 21
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountNotSpacesAfterFootnoteReferences"
    Case 22
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountEmptyParagraphsWithFormatting"
    Case 23
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountDeleteEmptyParagraphsBeforeHeading2"
    Case 24
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountFootnoteReferences"
    Case 25
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountFootnoteReferenceColors"
    Case 26
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CheckAllHeaders(" & "Empty" & ")"
    Case 27
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CheckAllHeaders(" & "NotEmpty" & ")"
    Case 28
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountTabFollowedByParagraphMarkInHeaders"
    Case 29
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountParagraphsWithoutTabInHeaders"
    Case 30
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountHeaderStyleUsage"
    Case 31
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountParagraphMarksPerHeaderSection"
    Case 32
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountLinefeed"
    Case 33
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountLinefeed(" & """ """ & ")"
    Case 34
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountManualLineBreaksAndWithSpace"
    Case 35
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountManualLineBreaksAndWithSpace(" & """ """ & ")"
    Case 36
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountFooterParagraphsWithFooterStyle"
    Case 37
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountDocTabOnlyParagraphs"
    Case 38
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountEmptyParagraphs"
    Case 39
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountParagraphMarks_ArialBlack"
    Case 40
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountParagraphMarks_ArialBlackDarkRed"
    Case 41
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed"
    Case 42
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountBoldFootnotesWordLevel"
    Case 43
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountDarkRedStyledParagraphMarks"
    Case 44
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountParagraphMarks_CalibriDarkRed"
    Case 45
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountFindNotEmphasisBlack"
    Case 46
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountFindNotEmphasisRed"
    Case 47
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountTabOnlyParagraphs(" & """Footer""" & ")"
    Case 48
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "AuditLiberationSansNarrowStyleDetails"
    Case 49
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountAuditStyles_ToFile"
    Case 50
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "SummarizeHeaderFooterAuditToFile"
    Case 51
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountAndCreateDefinitionForH2"
    Case 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "U+2019 CountContraction(" & """" & ContractionArrayU(num - CARTS + 1) & """" & ")"
    Case 67, 68
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & UnicodeLabelU & " CountUnicodeSeq([" & """" & UnicodeString & """" & "])"    ' Using [] to delineate unicode string in test output
    Case 999
        AppendToFile TestReportFileName, GetPassFailArray(num) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(50))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "SKIPPED"
    Case Else
        Debug.Print "num = " & num & " >>>The test number is outside the accepted range"
    End Select

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutPutTestReport of Class BibleClass"
    Resume PROC_EXIT

End Function

'==============================================================================
' Function: CountContraction
'
' Purpose:
'   Counts occurrences of a specific English contraction in the active document
'   using Word's native Find engine. The search is case-insensitive and optimized
'   for performance on large documents.
'
' Caveats:
'   - Matches only literal text using Word VBA's Find object.
'   - Will NOT find contraction substrings embedded within a larger word
'     (e.g., "it's" inside "Spirit's"), even though Ctrl+H may appear to do so.
'     This is a known difference between the Word UI search and the VBA Find API.
'   - Designed for standalone contractions (whole words), not sub-word matches.
'
' Additional:
'   - Can also accept up to 3 code points as a string when not calling
'     the contraction array.
'
' Parameters:
'   contraction  The contraction to search for, using Unicode U+2019 (ChrW(&H2019))
'                Tests 52~66
'   or           A Unicode string of up to 3 code points
'                Tests 67~
'
' Returns:
'   The number of matches found.
'==============================================================================
Private Function CountContraction(ByVal contraction As String) As Long
    Dim rng As range
    Dim count As Long

    Set rng = ActiveDocument.content
    count = 0

    With rng.Find
        .ClearFormatting
        .text = contraction
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .Wrap = wdFindStop
    End With

    Do While rng.Find.Execute
        count = count + 1
        rng.Collapse wdCollapseEnd
    Loop

    CountContraction = count
End Function

Public Sub TestContractionsExamples()
    Dim Apostrophe As String
    Apostrophe = ChrW(&H2019)

    'Debug.Print "Count of wouldn't  = " & CountContraction("wouldn" & Apostrophe & "t")
    'Debug.Print "Count of isn't     = " & CountContraction("isn" & Apostrophe & "t")
    'Debug.Print "Count of let's     = " & CountContraction("let" & Apostrophe & "s")
    'Debug.Print "Count of i'm       = " & CountContraction("i" & Apostrophe & "m")
    'Debug.Print "Count of hasn't    = " & CountContraction("hasn" & Apostrophe & "t")
    'Debug.Print "Count of didn't    = " & CountContraction("didn" & Apostrophe & "t")
    'Debug.Print "Count of don't     = " & CountContraction("don" & Apostrophe & "t")
    'Debug.Print "Count of haven't   = " & CountContraction("haven" & Apostrophe & "t")
    'Debug.Print "Count of weren't   = " & CountContraction("weren" & Apostrophe & "t")
    'Debug.Print "Count of aren't    = " & CountContraction("aren" & Apostrophe & "t")
    'Debug.Print "Count of can't     = " & CountContraction("can" & Apostrophe & "t")
    'Debug.Print "Count of shouldn't = " & CountContraction("shouldn" & Apostrophe & "t")
    'Debug.Print "Count of it's      = " & CountContraction("it" & Apostrophe & "s")
End Sub

' ======================================================================
' Macro Name   : CountAndCreateDefinitionForH2
' Purpose      : Enforces layout and paragraph rules for Heading 2 style.
'                - Enables KeepWithNext on Heading 2 paragraph style.
'                - Applies WidowControl to each Heading 2 paragraph.
'                - Explicitly disables KeepTogether to prevent override.
'                - Provides a total count of H2
' Audit Notes  : Logs all repair actions to Immediate Window.
'                Does NOT alter any content, punctuation, or quotes.
'                Applies paragraph-level enforcement only where style = "Heading 2"
' Safety Level : Editorial-safe. No deletions or format coercion.
' Last Updated : 20250805
' Author       : Peter
' ======================================================================
Private Function CountAndCreateDefinitionForH2()
    ' Update Heading 2 Keep With Next
    Dim count As Long
    count = 0

    Application.ScreenUpdating = False

    Dim s As style
    Set s = ActiveDocument.Styles("Heading 2")
    
    ' Apply KeepWithNext to paragraph formatting
    s.ParagraphFormat.KeepWithNext = True
    'Debug.Print "Heading 2 style updated: KeepWithNext = True"

    ' Enforce Heading 2 Paragraph Widow Orphan
    ' Disable KeepLines Together For Heading 2
    Dim para As paragraph
    For Each para In ActiveDocument.paragraphs
        If para.style = ActiveDocument.Styles("Heading 2") Then
            count = count + 1
            With para
                .WidowControl = True    ' enforces both widow and orphan control for that paragraph
                '.OrphanControl = True - Not needed,
                para.KeepTogether = False
            End With
            
            If count Mod 50 = 0 Then
                DoEvents  ' Keep UI responsive every 50 updates
            End If
        End If
    Next para
    
    Application.ScreenUpdating = True
    'Debug.Print "[repair] Widow/Orphan enforced at paragraph level for Heading 2"
    'Debug.Print "[repair] KeepLinesTogether disabled for Heading 2"
    'Debug.Print "count = " & count
    CountAndCreateDefinitionForH2 = count
End Function

'=========================================================================
' Function: SummarizeHeaderFooterAuditToFile
'
' Purpose:
'   Write a clean audit summary of all header/footer results
'   to the shared diagnostics file: rpt\HeaderFooterAudit.txt.
'
'   This file is overwritten only on failure and remains stable
'   across passing runs to avoid Git churn or session ID noise.
'
'   The audit contains linewise status of each header/footer
'   region, tracking anomalies such as style drift, legacy font
'   usage, or unexpected layout artifacts.
'
' Behavior:
'   - Overwrites the target file only if failure is detected
'   - Sorts output by section and header/footer type
'   - Skips export entirely if diagnostics are clean
'
' Output Format:
'   [FAIL] Section 3 Footer - Legacy font detected: TimesNewRoman
'   [PASS] Section 3 Header - Styles consistent
'
'   File: rpt\HeaderFooterAudit.txt
'
' References:
'   See commit changelog #270 for rationale
'
' Visibility:
'   Exported in diagnostics mode; safe for GitHub.io surfacing
'=========================================================================
Private Function SummarizeHeaderFooterAuditToFile()
    Dim sec As section, hf As HeaderFooter, p As paragraph
    Dim paraText As String, paraStyle As String
    Dim asciiVal As Long
    Dim headerOk As Boolean, footerOk As Boolean
    Dim dictAsciiCount As Object, dictAsciiDesc As Object
    Dim dictAllStyles As Object
    Dim dictBadHeaderStyles As Object, dictBadFooterStyles As Object
    Dim totalSections As Long
    Dim fs As Object, ts As Object
    Dim rptPath As String

    ' Create dictionary objects
    Set dictAsciiCount = CreateObject("Scripting.Dictionary")
    Set dictAsciiDesc = CreateObject("Scripting.Dictionary")
    Set dictAllStyles = CreateObject("Scripting.Dictionary")
    Set dictBadHeaderStyles = CreateObject("Scripting.Dictionary")
    Set dictBadFooterStyles = CreateObject("Scripting.Dictionary")

    headerOk = True: footerOk = True
    totalSections = ActiveDocument.Sections.count

    ' Prepare folder and file
    rptPath = ActiveDocument.Path & "\rpt"
    If Dir(rptPath, vbDirectory) = "" Then MkDir rptPath
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ts = fs.CreateTextFile(rptPath & "\HeaderFooterAudit.txt", True)

    ts.WriteLine "=== Header/Footer ASCII Audit Summary ==="
    ts.WriteLine "Total Sections: " & totalSections
    'Debug.Print "Total Sections: " & totalSections

    For Each sec In ActiveDocument.Sections
        For Each hf In sec.Headers
            If hf.Exists Then
                For Each p In hf.range.paragraphs
                    paraText = p.range.text
                    asciiVal = AscW(Left(paraText, 1))
                    paraStyle = p.style.NameLocal
                    If paraStyle <> "TheHeaders" Then
                        headerOk = False
                        dictBadHeaderStyles("SECTION " & sec.index & " = " & paraStyle) = ""
                    End If

                    If Not dictAsciiCount.Exists(asciiVal) Then dictAsciiCount(asciiVal) = 0
                    dictAsciiCount(asciiVal) = dictAsciiCount(asciiVal) + 1

                    If Not dictAsciiDesc.Exists(asciiVal) Then
                        If asciiVal >= 32 And asciiVal <= 126 Then
                            dictAsciiDesc(asciiVal) = ChrW(asciiVal)
                        ElseIf asciiVal = 9 Then
                            dictAsciiDesc(asciiVal) = "[Tab]"
                        ElseIf asciiVal = 13 Then
                            dictAsciiDesc(asciiVal) = "[CR]"
                        Else
                            dictAsciiDesc(asciiVal) = "[Non-printable]"
                        End If
                    End If

                    If Not dictAllStyles.Exists(paraStyle) Then dictAllStyles(paraStyle) = "Header"
                Next p
            End If
        Next hf

        For Each hf In sec.Footers
            If hf.Exists Then
                For Each p In hf.range.paragraphs
                    paraText = p.range.text
                    asciiVal = AscW(Left(paraText, 1))
                    paraStyle = p.style.NameLocal
                    If paraStyle <> "TheFooters" Then
                        footerOk = False
                        dictBadFooterStyles("SECTION " & sec.index & " = " & paraStyle) = ""
                    End If

                    If Not dictAsciiCount.Exists(asciiVal) Then dictAsciiCount(asciiVal) = 0
                    dictAsciiCount(asciiVal) = dictAsciiCount(asciiVal) + 1

                    If Not dictAsciiDesc.Exists(asciiVal) Then
                        If asciiVal >= 32 And asciiVal <= 126 Then
                            dictAsciiDesc(asciiVal) = ChrW(asciiVal)
                        ElseIf asciiVal = 9 Then
                            dictAsciiDesc(asciiVal) = "[Tab]"
                        ElseIf asciiVal = 13 Then
                            dictAsciiDesc(asciiVal) = "[CR]"
                        Else
                            dictAsciiDesc(asciiVal) = "[Non-printable]"
                        End If
                    End If

                    If Not dictAllStyles.Exists(paraStyle) Then dictAllStyles(paraStyle) = "Footer"
                Next p
            End If
        Next hf
    Next sec

    ' Output to file based on validation
    If headerOk And footerOk Then
        ts.WriteLine "--- ASCII Value Counts and Descriptions ---"
        Dim key As Variant
        For Each key In dictAsciiCount.Keys
            ts.WriteLine "ASCII=" & key & _
                " | Count=" & dictAsciiCount(key) & _
                " | Desc=" & dictAsciiDesc(key)
        Next key
        
        Dim styleKey As Variant
        Dim allGoodStyles As Boolean
        Dim tempDict As Object
        Set tempDict = CreateObject("Scripting.Dictionary")

        ' Build tempDict from dictAllStyles keys
        For Each styleKey In dictAllStyles.Keys
            tempDict(styleKey) = 1
        Next styleKey

        allGoodStyles = (tempDict.count = 2) _
                        And tempDict.Exists("TheHeaders") _
                        And tempDict.Exists("TheFooters")

        If allGoodStyles Then
            ts.WriteLine "--- Validated Styles ---"
            ts.WriteLine "Expected Header Style: TheHeaders"
            ts.WriteLine "Expected Footer Style: TheFooters"
        End If
        ts.WriteLine "Style Check: PASS"
        'Debug.Print "Style Check: PASS"
        SummarizeHeaderFooterAuditToFile = totalSections
    Else
        ts.WriteLine "Style Check: FAIL"
        'Debug.Print "Style Check: FAIL"
        SummarizeHeaderFooterAuditToFile = -1
        ts.WriteLine "--- Encountered Styles ---"
        For Each key In dictAllStyles.Keys
            ts.WriteLine "STYLE=" & key & " | LOCATION=" & dictAllStyles(key)
        Next key
        ts.WriteLine "--- Style Mismatches by Section ---"
        For Each key In dictBadHeaderStyles.Keys
            ts.WriteLine "[HEADER] " & key
        Next key
        For Each key In dictBadFooterStyles.Keys
            ts.WriteLine "[FOOTER] " & key
        Next key
    End If

    ts.WriteLine "=== Summary Complete ==="
    ts.Close
End Function

'==============================================================
' CountAuditStyles_ToFile
' -------------------------------------------------------------
' Purpose : Audits paragraph style usage across all story ranges
'           in the active Word document and outputs a frequency
'           distribution report to an external ASCII text file.
'
' Returns : Long - the number of unique styles detected.
'
' Behavior:
' - Iterates all story ranges (main text, headers, footers, etc.).
' - Counts paragraph occurrences per style using a Scripting.Dictionary.
' - Outputs results as "Style: Count" pairs, plus total paragraph count.
' - Writes to "\rpt\Style Usage Distribution.txt".
' - Skips UI prompts; uses silent logging for audit pipelines.
'
' Notes   :
' - Respects editorial boundaries by auditing only - no content changes.
' - Compatible with audit-friendly workflows and batch diagnostics.
' - Output is ASCII-only to maximize portability and diff clarity.
' - Assumes the "\rpt\" folder exists; add directory creation if needed.
'
' Author  : Peter
' Last Modified : 20250731
'==============================================================
Private Function CountAuditStyles_ToFile() As Long
    Dim rng As range
    Dim para As paragraph
    Dim paraStyle As String
    Dim paraCount As Long
    Dim keyCount As Long
    Dim styleDict As Object
    Set styleDict = CreateObject("Scripting.Dictionary")

    ' Build the distribution
    For Each rng In ActiveDocument.StoryRanges
        Do
            For Each para In rng.paragraphs
                paraStyle = para.style.NameLocal
                paraCount = paraCount + 1
                If Not styleDict.Exists(paraStyle) Then
                    styleDict.Add paraStyle, 1
                Else
                    styleDict(paraStyle) = styleDict(paraStyle) + 1
                End If
            Next para
            Set rng = rng.NextStoryRange
        Loop Until rng Is Nothing
    Next rng

    Dim fso As Object, outFile As Object, outputPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    outputPath = fso.GetParentFolderName(ActiveDocument.FullName) & "\rpt\Style Usage Distribution.txt"
    Set outFile = fso.CreateTextFile(outputPath, True)
    'Debug.Print "outputPath = " & outputPath

    outFile.WriteLine "=== Style Usage Distribution ==="
    Dim key As Variant
    keyCount = 0
    For Each key In styleDict.Keys
        keyCount = keyCount + 1
        'Debug.Print keyCount, key, styleDict(key)
        outFile.WriteLine key & ": " & styleDict(key)
    Next key
    outFile.WriteLine "Total paragraphs (all styles): " & paraCount
    outFile.Close

    'Debug.Print "Style usage written to:" & vbCrLf & outputPath
    'MsgBox "Style usage written to:" & vbCrLf & outputPath, vbInformation, "Audit Complete"
    'Debug.Print "keyCount = " & keyCount
    CountAuditStyles_ToFile = keyCount
End Function

Private Function AuditLiberationSansNarrowStyleDetails() As Integer
    Dim s As style
    Dim logBuffer As String
    Dim fName As String, fSize As Variant, pStyleType As String
    Dim styleCount As Long

    logBuffer = "=== Liberation Sans Narrow Style Definitions ===" & vbCrLf

    For Each s In ActiveDocument.Styles
        On Error Resume Next
        fName = s.font.name
        fSize = s.font.Size
        If fName = "Liberation Sans Narrow" Then
            styleCount = styleCount + 1
            Select Case s.Type
                Case wdStyleTypeParagraph: pStyleType = "Paragraph"
                Case wdStyleTypeCharacter: pStyleType = "Character"
                Case wdStyleTypeTable: pStyleType = "Table"
                Case wdStyleTypeList: pStyleType = "List"
                Case Else: pStyleType = "Other"
            End Select
            logBuffer = logBuffer & _
                "- Style: “" & s.NameLocal & "” | Type: " & pStyleType & _
                " | Font Size: " & fSize & vbCrLf
        End If
        On Error GoTo 0
    Next s

    'logBuffer = logBuffer & vbCrLf & "Total styles using Liberation Sans Narrow: " & styleCount
    'Debug.Print logBuffer
    'MsgBox "Style details printed. See Immediate Window.", vbInformation
    AuditLiberationSansNarrowStyleDetails = styleCount
End Function

Private Function CountTabOnlyParagraphs(areaType As String) As Long
    Dim sec As section
    Dim hdrFtr As HeaderFooter
    Dim para As paragraph
    Dim count As Long
    Dim paraText As String

    Dim targetType As WdHeaderFooterIndex
    Select Case LCase(areaType)
        Case "header": targetType = wdHeaderFooterPrimary
        Case "footer": targetType = wdHeaderFooterPrimary
        Case Else
            MsgBox "Invalid parameter. Use 'Header' or 'Footer'.", vbCritical
            Exit Function
    End Select

    For Each sec In ActiveDocument.Sections
        If LCase(areaType) = "header" Then
            Set hdrFtr = sec.Headers(targetType)
        Else
            Set hdrFtr = sec.Footers(targetType)
        End If

        If hdrFtr.Exists Then
            For Each para In hdrFtr.range.paragraphs

                ' Move the cursor to this paragraph in the footer
                para.range.Select

                paraText = Replace(para.range.text, vbCr, "")
                paraText = Replace(paraText, vbLf, "")
                paraText = Trim(paraText)

                If paraText = vbTab Then
                    count = count + 1
                    'Debug.Print "count = " & count
                    'Stop   ' Execution halts with the cursor at the footer location
                End If
            Next para
        End If
    Next sec

    CountTabOnlyParagraphs = count
End Function

Private Function CountFindNotEmphasisBlack()
    CountFindNotEmphasisBlack = FindNotEmphasisBlackRed("Automatic")
End Function

Private Function CountFindNotEmphasisRed()
    CountFindNotEmphasisRed = FindNotEmphasisBlackRed("Dark Red")
End Function

Private Function FindNotEmphasisBlackRed(fontColor As String) As Integer
' fontColor = Automatic - Arial Black, 8pt, "Normal" Style
' fontColor = Dark Red - Arial Black, 8pt, "Words of Jesus" Style
' This is fast and will skip over text that is of the style EmphasisBlack or EmphasisRed
' When count is 0 there is no more text that needs setting for style EmphasisBlack or EmphasisRed

    Dim rng As range
    Set rng = ActiveDocument.content
    Dim totalCount As Integer
    Dim wordColor As Long
    Dim wordStyle As style
    Dim myStyle As style

    If fontColor = "Automatic" Then
        wordColor = wdColorAutomatic
        Set wordStyle = ActiveDocument.Styles("Normal")
        Set myStyle = ActiveDocument.Styles("EmphasisBlack")
    ElseIf fontColor = "Dark Red" Then
        wordColor = wdColorDarkRed
        Set wordStyle = ActiveDocument.Styles("Words of Jesus")
        Set myStyle = ActiveDocument.Styles("EmphasisRed")
    Else
        MsgBox "Incorrect parameter. Automatic or Dark Red expected!", vbCritical, "FindNotEmphasisBlackRed"
        Stop
    End If

    'Debug.Print "wordStyle = " & wordStyle, "wordColor = " & wordColor, "wdColorDarkRed = " & wdColorDarkRed, _
                    "myStyle = " & myStyle
    totalCount = 0
    With rng.Find
        .ClearFormatting
        .style = ActiveDocument.Styles(wordStyle)
        .font.name = "Arial Black"
        .font.Size = 8
        .font.color = wordColor
        .text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True

        'Debug.Print ".style = " & rng.Find.style
        Do While .Execute
            ' Check if character style is NOT EmphasisBlack or EmphasisRed
            If Not rng.Characters(1).style = myStyle Then
                rng.Select
                totalCount = totalCount + 1
                MsgBox "Found matching text (not " & myStyle & ").", vbInformation
                Exit Function
            End If
            rng.Collapse wdCollapseEnd
        Loop

        'If totalCount = 0 Then MsgBox "totalCount = 0, No matching text found (excluding EmphasisBlack).", vbExclamation
    End With
    'Debug.Print "totalCount = " & totalCount
End Function

Private Function CountParagraphMarks_CalibriDarkRed() As Integer
    Dim para As paragraph
    Dim paraRange As range
    Dim totalCount As Long
    Dim foundFirst As Boolean
    Dim fontColor As WdColor

    totalCount = 0
    foundFirst = False

    For Each para In ActiveDocument.paragraphs
        Set paraRange = para.range
        paraRange.Start = paraRange.End - 1  ' Isolate the paragraph mark

        With paraRange.font
            fontColor = .color
            If .name = "Calibri" And _
               .Size = 9 And _
               (fontColor = wdColorDarkRed) Then

                totalCount = totalCount + 1
                If Not foundFirst Then
                    paraRange.Select
                    foundFirst = True
                    'Stop
                End If
            End If
        End With
    Next para

    'MsgBox "Total matching paragraph marks: " & totalCount, vbInformation
    'Debug.Print "totalCount = " & totalCount
    CountParagraphMarks_CalibriDarkRed = totalCount
End Function

Private Function CountDarkRedStyledParagraphMarks()
    Dim rng As range
    Dim totalMatches As Long
    Dim found As Boolean

    Set rng = ActiveDocument.content
    totalMatches = 0
    found = False

    With rng.Find
        .ClearFormatting
        .text = "^13"
        .font.name = "Calibri"
        .font.Size = 9
        .font.color = wdColorDarkRed
        .Forward = True
        .Wrap = wdFindStop
        .Format = True

        Do While .Execute
            totalMatches = totalMatches + 1

            If Not found Then
                rng.Select ' Move cursor to the first matching styled paragraph mark
                found = True
            End If

            rng.Collapse wdCollapseEnd
        Loop
    End With

    'MsgBox "Total styled paragraph marks found: " & totalMatches & vbCrLf & _
           IIf(found, "Cursor moved to first match.", "No matching paragraph marks found."), _
           vbInformation, "Search Complete"
    'Debug.Print "Total styled paragraph marks found: " & totalMatches & vbCrLf & _
           IIf(found, "Cursor moved to first match.", "No matching paragraph marks found.")
    CountDarkRedStyledParagraphMarks = totalMatches
End Function

Private Function CountBoldFootnotesWordLevel()
    Dim ft As footnote
    Dim aword As range
    Dim boldCount As Long
    Dim hasBold As Boolean

    For Each ft In ActiveDocument.Footnotes
        hasBold = False

        If ft.range.style = "Footnote Text" Then
            For Each aword In ft.range.words
                If aword.font.Bold = True Then
                    boldCount = boldCount + 1
                    hasBold = True
                    Exit For
                End If
            Next aword
        End If
    Next ft

    If boldCount > 0 Then
        MsgBox boldCount & " footnotes contain bold words in 'Footnote Text' style.", vbInformation
        Stop
    Else
        'MsgBox "No bold words found in 'Footnote Text' footnotes.", vbInformation
        CountBoldFootnotesWordLevel = 0
    End If
End Function

Private Function Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed()
    Dim rng As range
    Dim totalCount As Long
    Dim foundFirst As Boolean
    Dim docRange As range

    Set docRange = ActiveDocument.content
    totalCount = 0
    foundFirst = False

    Set rng = docRange.Duplicate

    With rng.Find
        .ClearFormatting
        .style = ActiveDocument.Styles("Words of Jesus")
        .font.name = "Arial Black"
        .font.Size = 8
        .font.color = wdColorDarkRed
        .text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True

        Do While .Execute
            If rng.Characters(1).style <> "EmphasisRed" Then
                totalCount = totalCount + 1

                If Not foundFirst Then
                    rng.Select
                    foundFirst = True
                End If
            End If
            rng.Collapse wdCollapseEnd
        Loop
    End With

    If foundFirst Then
        MsgBox "First match selected. Total matches (excluding EmphasisRed): " & totalCount, vbInformation
        Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed = totalCount
    Else
        'MsgBox "No matching text found (excluding EmphasisRed or wrong color).", vbExclamation
        'Debug.Print "NotEmphasisRed totalCount = " & totalCount
        Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed = totalCount
    End If
End Function

Private Function CountParagraphMarks_ArialBlackDarkRed() As Integer
    Dim para As paragraph
    Dim paraRange As range
    Dim totalCount As Long
    Dim foundFirst As Boolean
    Dim fontColor As WdColor

    totalCount = 0
    foundFirst = False

    For Each para In ActiveDocument.paragraphs
        Set paraRange = para.range
        paraRange.Start = paraRange.End - 1  ' Isolate the paragraph mark

        With paraRange.font
            fontColor = .color
            If .name = "Arial Black" And _
               .Size = 8 And _
               (fontColor = wdColorDarkRed) Then

                totalCount = totalCount + 1
                If Not foundFirst Then
                    paraRange.Select
                    foundFirst = True
                    Stop
                End If
            End If
        End With
    Next para

    'MsgBox "Total matching paragraph marks: " & totalCount, vbInformation
    'Debug.Print "totalCount = " & totalCount
    CountParagraphMarks_ArialBlackDarkRed = totalCount
End Function

Private Function CountParagraphMarks_ArialBlack() As Integer
    Dim para As paragraph
    Dim paraRange As range
    Dim totalCount As Long
    Dim foundFirst As Boolean
    Dim fontColor As WdColor

    totalCount = 0
    foundFirst = False

    For Each para In ActiveDocument.paragraphs
        Set paraRange = para.range
        paraRange.Start = paraRange.End - 1  ' Isolate the paragraph mark

        With paraRange.font
            fontColor = .color
            If .name = "Arial Black" And _
               .Size = 8 And _
               (fontColor = wdColorAutomatic Or fontColor = wdColorBlack) Then

                totalCount = totalCount + 1
                If Not foundFirst Then
                    paraRange.Select
                    foundFirst = True
                    'Stop
                End If
            End If
        End With
    Next para

    'MsgBox "Total matching paragraph marks: " & totalCount, vbInformation
    'Debug.Print "totalCount = " & totalCount
    CountParagraphMarks_ArialBlack = totalCount
End Function

Private Function CountEmptyParagraphs() As Long
    Dim para As paragraph
    Dim count As Long
    count = 0
    For Each para In ActiveDocument.paragraphs
        If Len(para.range.text) = 1 And para.range.text = vbCr Then
            count = count + 1
        End If
    Next para
    CountEmptyParagraphs = count
End Function

Private Function CountDocTabOnlyParagraphs() As Integer
' The final paragraph in a Word document always includes an extra paragraph mark that you can't delete.
' The extra count is usually the final paragraph, which Word always includes and may contain leftover tab formatting.
' This code can filter that out.
    Dim para As paragraph
    Dim countTabOnly As Long
    Dim paraText As String

    countTabOnly = 0

    For Each para In ActiveDocument.paragraphs
        paraText = para.range.text

        ' Skip final paragraph if it's the only one and contains just tab+CR
        If para.range.End = ActiveDocument.content.End And paraText = vbTab & vbCr Then
            ' Comment out to skip this one if desired
            countTabOnly = countTabOnly + 1
        Else
            ' Count if exactly tab + paragraph mark
            If paraText = vbTab & vbCr Then
                countTabOnly = countTabOnly + 1
            End If
        End If
    Next para

    'MsgBox "Paragraphs with only a tab and paragraph mark: " & countTabOnly, vbInformation
    'Debug.Print "Paragraphs with only a tab and paragraph mark: " & countTabOnly
    CountDocTabOnlyParagraphs = countTabOnly
End Function

Private Function CountFooterParagraphsWithFooterStyle() As Integer
    Dim sec As section
    Dim hdrFtr As HeaderFooter
    Dim para As paragraph
    Dim count As Long

    count = 0

    ' Loop through all sections in the document
    For Each sec In ActiveDocument.Sections
        ' Only check footers
        For Each hdrFtr In sec.Footers
            If hdrFtr.Exists Then
                For Each para In hdrFtr.range.paragraphs
                    If para.style = "Footer" Then
                        count = count + 1
                        para.range.Select
                        Debug.Print "Found paragraph with Footer style. Stopping at this location."
                        Stop
                    End If
                Next para
            End If
        Next hdrFtr
    Next sec

    'MsgBox "Number of paragraphs in footers with the style 'Footer': " & count, vbInformation
    'Debug.Print "Number of paragraphs in footers with the style 'Footer': " & count
    CountFooterParagraphsWithFooterStyle = count
End Function

Private Function CountManualLineBreaksAndWithSpace(Optional space As Variant) As Integer
    Dim fn As footnote
    Dim totalBreaks As Long
    Dim breaksWithSpaceBefore As Long
    Dim i As Long
    Dim content As String
    Dim currentChar As String
    Dim prevChar As String

    totalBreaks = 0
    breaksWithSpaceBefore = 0

    For Each fn In ActiveDocument.Footnotes
        content = fn.range.text
        For i = 1 To Len(content)
            currentChar = mid(content, i, 1)
            If Asc(currentChar) = 11 Then ' Chr(11) = manual line break
                totalBreaks = totalBreaks + 1
                If i > 1 Then
                    prevChar = mid(content, i - 1, 1)
                    If prevChar = " " Then
                        breaksWithSpaceBefore = breaksWithSpaceBefore + 1
                    ElseIf prevChar <> " " Then
                        Stop
                    End If
                End If
            End If
        Next i
    Next fn

    'MsgBox "Total manual line breaks: " & totalBreaks & vbCrLf & _
    '       " of which are preceded by a space: " & breaksWithSpaceBefore, vbInformation
    'Debug.Print "Total manual line breaks: " & totalBreaks & vbCrLf & _
    '       " of which are preceded by a space: " & breaksWithSpaceBefore
    If CStr(space) = " " Then
        'Debug.Print "Space"
        CountManualLineBreaksAndWithSpace = breaksWithSpaceBefore
        Exit Function
    End If
    If IsMissing(space) Then
        'Debug.Print "No Argument"
        CountManualLineBreaksAndWithSpace = totalBreaks
    Else
        MsgBox "Invalid Argument", vbCritical
        Stop
    End If
End Function

Private Function CountLinefeed(Optional space As Variant) As Integer
' Search in doc
    Dim doc As Document
    Dim rng As range
    Dim findText As String
    Dim count As Integer
    
    If IsMissing(space) Then
        'Debug.Print "IsMissing space"
        findText = "^l"
    Else
        'Debug.Print "Space Is Not Missing"
        findText = " {1}^l"     ' space followed by linefeed
    End If
    
    ' Set the document
    Set doc = ActiveDocument
       
    ' Search in the main document content
    Set rng = doc.content
    
    Application.ScreenUpdating = False
    With rng.Find
        .text = findText
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
        
        ' Loop through all instances and count them
        Do While .Execute
            count = count + 1
            ' Keep Word responsive every 50
            If count Mod 500 = 0 Then DoEvents
        Loop
    End With
    Application.ScreenUpdating = True
    
    ' Display the count
    'MsgBox "Number of matches found: " & count
    'Debug.Print "Number of matches found: " & count
    CountLinefeed = count
End Function

Private Function CountParagraphMarksPerHeaderSection() As Integer
    Dim doc As Document
    Dim sectionIndex As Integer
    Dim headerRange As range
    Dim paraCount As Integer
    
    ' Set the document
    Set doc = ActiveDocument
    paraCount = 1
    
    ' Loop through all sections
    For sectionIndex = 1 To doc.Sections.count
        ' Get the range of the header in the current section
        Set headerRange = doc.Sections(sectionIndex).Headers(wdHeaderFooterPrimary).range
        
        ' Count the number of paragraphs in the header
        paraCount = headerRange.paragraphs.count
        
        ' Print the result to the console
        'Debug.Print "Section " & sectionIndex & " header has " & paraCount & " paragraph(s)."
        
        ' Check if the paragraph count is more than one
        If paraCount > 1 Then
            MsgBox "Section " & sectionIndex & " header has more than one paragraph mark. Stopping at this section."
            Exit Function
        End If
    Next sectionIndex
    
    ' Display a message indicating the count is complete
    'MsgBox "Count of paragraph marks per header section is " & paraCount
    'Debug.Print "Count of paragraph marks per header section is " & paraCount
    CountParagraphMarksPerHeaderSection = paraCount
End Function

Private Function CountHeaderStyleUsage() As Integer
    Dim doc As Document
    Dim rng As range
    Dim styleName As String
    Dim count As Integer
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Define the style name to search for
    styleName = "Header"
    
    ' Initialize the range to search
    Set rng = doc.content
    
    ' Initialize the count
    count = 0
    
    ' Search for the style "Header"
    With rng.Find
        .ClearFormatting
        .style = styleName
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchWildcards = False
        
        ' Loop through all instances of the style "Header"
        Do While .Execute
            count = count + 1
        Loop
    End With
    
    ' Display the count
    'MsgBox "The style 'Header' is used " & count & " times in the document."
    'Debug.Print "The style 'Header' is used " & count & " times in the document."
    CountHeaderStyleUsage = count
End Function

Private Function CountParagraphsWithoutTabInHeaders() As Integer
    Dim doc As Document
    Dim sectionIndex As Integer
    Dim headerRange As range
    Dim para As paragraph
    Dim count As Integer
    Dim totalCount As Integer
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Initialize the total count
    totalCount = 0
    
    ' Loop through all sections
    For sectionIndex = 1 To doc.Sections.count
        ' Get the range of the header in the current section
        Set headerRange = doc.Sections(sectionIndex).Headers(wdHeaderFooterPrimary).range
        
        ' Initialize the count for the current header
        count = 0
        
        ' Loop through each paragraph in the header range
        For Each para In headerRange.paragraphs
            ' Check if the paragraph does not contain a tab
            If InStr(para.range.text, vbTab) = 0 Then
                count = count + 1
            End If
        Next para
        
        ' Update the total count
        totalCount = totalCount + count
        
        ' Display the count for the current header
        'MsgBox "Section " & sectionIndex & " header has " & count & " paragraphs without a tab."
        'Debug.Print "Section " & sectionIndex & " header has " & count & " paragraphs without a tab."
    Next sectionIndex
    
    ' Display the total count
    'MsgBox "Total paragraphs without a tab in all headers: " & totalCount
    'Debug.Print "Total paragraphs without a tab in all headers: " & totalCount
    CountParagraphsWithoutTabInHeaders = totalCount
End Function

Private Function CountTabFollowedByParagraphMarkInHeaders() As Integer
    Dim doc As Document
    Dim sectionIndex As Integer
    Dim headerRange As range
    Dim para As paragraph
    Dim count As Integer
    Dim totalCount As Integer
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Initialize the total count
    totalCount = 0
    
    ' Loop through all sections
    For sectionIndex = 1 To doc.Sections.count
        ' Get the range of the header in the current section
        Set headerRange = doc.Sections(sectionIndex).Headers(wdHeaderFooterPrimary).range
        
        ' Initialize the count for the current header
        count = 0
        
        ' Loop through each paragraph in the header range
        For Each para In headerRange.paragraphs
            ' Check if the paragraph contains a tab followed by a paragraph mark
            If InStr(para.range.text, vbTab & vbCr) > 0 Then
                count = count + 1
            End If
        Next para
        
        ' Update the total count
        totalCount = totalCount + count
        
        ' Display the count for the current header
        'MsgBox "Section " & sectionIndex & " header has " & count & " paragraphs with a tab followed by a paragraph mark."
        'Debug.Print "Section " & sectionIndex & " header has " & count & " paragraphs with a tab followed by a paragraph mark."
    Next sectionIndex
    
    ' Display the total count
    'MsgBox "Total paragraphs with a tab followed by a paragraph mark in all headers: " & totalCount
    'Debug.Print "Total paragraphs with a tab followed by a paragraph mark in all headers: " & totalCount
    CountTabFollowedByParagraphMarkInHeaders = totalCount
End Function

Private Function CheckAllHeaders(EmptyOrNotEmpty As String) As Integer
' Identifies headers with only paragraph marks as empty. Checks if the header contains any visible text or characters beyond just paragraph marks.
'1. Sets the active document.
'2. Initializes counters for empty headers, not empty headers, and total headers.
'3. Loops through all sections in the document.
'4. Defines the range of the header in each section.
'5. Checks if the header is empty by iterating through each paragraph in the header range and ensuring it contains more than just a paragraph mark.
'6. Updates the counters based on whether the header is empty or not.
'7. Displays a message box with the total number of headers, the number of empty headers, and the number of not empty headers.    Dim doc As Document
    Dim sectionIndex As Integer
    Dim headerRange As range
    Dim para As paragraph
    Dim isEmpty As Boolean
    Dim emptyCount As Integer
    Dim notEmptyCount As Integer
    Dim totalHeaders As Integer
    
    ' Set the document
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Initialize counters
    emptyCount = 0
    notEmptyCount = 0
    totalHeaders = 0
    
    ' Loop through all sections
    For sectionIndex = 1 To doc.Sections.count
        ' Get the range of the header in the current section
        Set headerRange = doc.Sections(sectionIndex).Headers(wdHeaderFooterPrimary).range
        
        ' Check if the header is empty
        isEmpty = True
        For Each para In headerRange.paragraphs
            If Len(Trim(para.range.text)) > 1 Then ' Check if paragraph contains more than just a paragraph mark
                isEmpty = False
                Exit For
            End If
        Next para
        
        ' Update counters
        totalHeaders = totalHeaders + 1
        If isEmpty Then
            emptyCount = emptyCount + 1
        Else
            notEmptyCount = notEmptyCount + 1
        End If
    Next sectionIndex
    
    'Debug.Print "Total headers: " & totalHeaders & vbCrLf & _
    '       "Empty headers: " & emptyCount & vbCrLf & _
    '       "Not empty headers: " & notEmptyCount
    
    If EmptyOrNotEmpty = "Empty" Then
        CheckAllHeaders = emptyCount
        Exit Function
    ElseIf EmptyOrNotEmpty = "NotEmpty" Then
        CheckAllHeaders = notEmptyCount
        Exit Function
    Else
        MsgBox "Parameter not valid", vbCritical, "CheckAllHeaders"
        Stop
    End If
End Function

Private Function CountFootnoteReferenceColors() As Integer
    Dim doc As Document
    Set doc = ActiveDocument
   
    Dim footnoteColors As Collection
    Set footnoteColors = New Collection
    
    Dim footnote As footnote
    Dim colorHex As String
    
    For Each footnote In doc.Footnotes
        colorHex = ColorToHex(footnote.Reference.font.color)
        On Error Resume Next
        footnoteColors.Add colorHex, colorHex
        On Error GoTo 0
    Next footnote
       
    Dim colorList As String
    colorList = "Colors of Footnote References in Hex: "
    
    Dim color As Variant
    Dim colorCount As Integer
    colorCount = 0
    For Each color In footnoteColors
        colorCount = colorCount + 1
        colorList = colorList & color & vbCrLf
    Next color
    
    'MsgBox colorList
    'Debug.Print colorList
    CountFootnoteReferenceColors = colorCount
End Function

Private Function ColorToHex(color As Long) As String
    ColorToHex = Right("000000" & Hex(color), 6)
End Function

Private Function CountFootnoteReferences() As Integer
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim footnoteCount As Integer
    footnoteCount = doc.Footnotes.count
       
    ' Display the results
    'MsgBox "Number of Footnote References: " & footnoteCount
    CountFootnoteReferences = footnoteCount
End Function

Private Function CountDeleteEmptyParagraphsBeforeHeading2() As Integer
' Avoids indexing (.Item(i)) — slow in Word.
' Does not call .Style unless needed.
' Only trims the text once per loop.
' Dramatically faster!
    Dim para As paragraph
    Dim prevText As String
    Dim prevIsEmpty As Boolean
    Dim count As Long
    Dim paraCount As Long

    count = 0
    paraCount = 0
    prevIsEmpty = False

    Application.ScreenUpdating = False
    ' Loop through paragraphs only once
    For Each para In ActiveDocument.range.paragraphs
        paraCount = paraCount + 1
        If para.style = "Heading 2" Then
            If prevIsEmpty Then
                count = count + 1
            End If
        End If

        ' Determine if current paragraph is empty (ignoring whitespace)
        prevText = Trim(para.range.text)
        prevIsEmpty = (prevText = "" Or prevText = Chr(13)) ' Chr(13) is vbCr
        ' Keep Word responsive every 500 paragraphs
        If paraCount Mod 500 = 0 Then DoEvents
    Next para
    Application.ScreenUpdating = True

    'MsgBox "Fast count of 'Heading 2' paragraphs preceded by an empty paragraph: " & count, vbInformation
    'Debug.Print "Fast count of 'Heading 2' paragraphs preceded by an empty paragraph: " & count
    CountDeleteEmptyParagraphsBeforeHeading2 = count
End Function

Private Function CountEmptyParagraphsWithFormatting() As Long
' To maximize speed and accuracy, stay in the Paragraphs collection (for true Word structure),
' but avoid anything expensive — no .Range.Text, .Style, etc. unless needed.
    Dim para As paragraph
    Dim rng As range
    Dim count As Long
    Dim paraCount As Long

    count = 0
    Set rng = ActiveDocument.content
    rng.Collapse Direction:=wdCollapseStart

    Application.ScreenUpdating = False
    ' Loop over all paragraphs using the built-in collection
    For Each para In ActiveDocument.paragraphs
        paraCount = paraCount + 1
        ' Work directly with Range.Text but trim just once
        If LenB(Trim$(Replace(para.range.text, vbCr, ""))) = 0 Then
            count = count + 1
        End If
        ' Keep Word responsive every 500 paragraphs
        If paraCount Mod 500 = 0 Then DoEvents
    Next para
    Application.ScreenUpdating = True

    'MsgBox "Accurate empty paragraph count: " & count, vbInformation
    'Debug.Print "Accurate empty paragraph count: " & count
    
    CountEmptyParagraphsWithFormatting = count
End Function

Private Function CountNotSpacesAfterFootnoteReferences() As Integer
' Also shows Footnote References and Following Characters (ASCII Values)
    Dim doc As Document
    Dim footnote As footnote
    Dim rng As range
    Dim spaceCount As Long
    Dim notSpaceCount As Long
    Dim debugInfo As String
    Dim asciiValue As Integer
    
    ' Initialize variables
    Set doc = ActiveDocument
    spaceCount = 0
    notSpaceCount = 0
    debugInfo = "Footnote References and Following Characters (ASCII Values):" & vbCrLf
    
    ' Loop through each footnote in the document
    For Each footnote In doc.Footnotes
        Set rng = footnote.Reference
        rng.Collapse wdCollapseEnd ' Collapse the range to the end of the footnote reference
        rng.MoveEnd wdCharacter, 1 ' Move the range to include the next character
        
        ' Get the ASCII value of the character
        asciiValue = Asc(rng.text)
        
        ' Exclude ASCII value 13 (carriage return)
        If asciiValue <> 13 Then
            ' Add debug information
            If asciiValue <> 32 Then    ' space (ASCII value 32)
                debugInfo = debugInfo & "Fn " & footnote.index & ": '" & rng.text & "' (ASC " & asciiValue & "), "
                'Debug.Print debugInfo
                notSpaceCount = notSpaceCount + 1
                'Stop
            End If
            
            If asciiValue = 32 Then
                spaceCount = spaceCount + 1
            End If
        End If
    Next footnote
    
    ' Display the count of spaces after footnote references
    'MsgBox "Number of spaces after footnote references: " & spaceCount
    ' Display debug information
    'MsgBox debugInfo
    
    ' Print the result to the console
    'Debug.Print "Number of spaces after footnote references: " & spaceCount
    'Debug.Print debugInfo
    CountNotSpacesAfterFootnoteReferences = notSpaceCount
End Function

Private Function CountFootnotesFollowedByDigit() As Integer
    Dim doc As Document
    Dim footnote As footnote
    Dim rng As range
    Dim digitCount As Long
    Dim debugInfo As String
    Dim asciiValue As Integer
    
    ' Initialize variables
    Set doc = ActiveDocument
    digitCount = 0
    debugInfo = "Footnote References and Following Characters (ASCII Values):" & vbCrLf
    
    ' Loop through each footnote in the document
    For Each footnote In doc.Footnotes
        Set rng = footnote.Reference
        rng.Collapse wdCollapseEnd ' Collapse the range to the end of the footnote reference
        rng.MoveEnd wdCharacter, 1 ' Move the range to include the next character
        
        ' Get the ASCII value of the character
        asciiValue = Asc(rng.text)
        
        ' Check if the character after the footnote reference is a digit (ASCII values 48-57)
        If asciiValue >= 48 And asciiValue <= 57 Then
            digitCount = digitCount + 1
        End If
        
        ' Add debug information
        debugInfo = debugInfo & "Footnote " & footnote.index & ": '" & rng.text & "' (ASCII: " & asciiValue & ")" & vbCrLf
    Next footnote
    
    ' Display the count of footnote references followed by a digit
    'MsgBox "Number of footnote references followed by a digit: " & digitCount
    
    ' Display debug information
    'MsgBox debugInfo
    'Debug.Print debugInfo
 
    ' Print the result to the console
    'Debug.Print "Number of footnote references followed by a digit: " & digitCount
    CountFootnotesFollowedByDigit = digitCount
End Function

Private Function CountEmptyParasAfterH2() As Integer
    Dim para As paragraph
    Dim emptyParaCount As Long
    Dim NextPara As paragraph
    
    ' Initialize the count
    emptyParaCount = 0
    
    ' Loop through each paragraph in the document
    For Each para In ActiveDocument.paragraphs
        ' Check if the paragraph style is Heading 2
        If para.style = ActiveDocument.Styles("Heading 2") Then
            ' Check the next paragraph
            Set NextPara = para.range.Next(Unit:=wdParagraph, count:=1).paragraphs(1)
            If Not NextPara Is Nothing Then
                ' Check if the next paragraph is empty
                If Len(Trim(NextPara.range.text)) = 1 Then ' Adjusted to check for a single paragraph mark
                    emptyParaCount = emptyParaCount + 1
                End If
            End If
        End If
    Next para
    
    ' Print the count to the console
    'Debug.Print "Number of empty paragraphs immediately after Heading 2: " & emptyParaCount
    CountEmptyParasAfterH2 = emptyParaCount
End Function

Private Function CountHeading1() As Integer
    Dim para As paragraph
    Dim heading1Count As Long
    
    ' Initialize the count
    heading1Count = 0
    
    ' Loop through each paragraph in the document
    For Each para In ActiveDocument.paragraphs
        ' Check if the paragraph style is Heading 1
        If para.style Like "*Heading 1*" Then   'para.style = ActiveDocument.Styles("Heading 1") Then
            heading1Count = heading1Count + 1
            'Debug.Print heading1Count & "[" & Replace(para.range.text, vbCr, "") & "]",
        End If
    Next para

    ' Print the count to the console
    'Debug.Print "Number of Heading 1 paragraphs: " & heading1Count
    CountHeading1 = heading1Count
    'Stop
End Function

Private Function CountRedFootnoteReferences() As Integer
    Dim footnote As footnote
    Dim redCount As Long
    Dim redColor As Long
    
    ' Define the red color using RGB values
    redColor = RGB(255, 0, 0)
    
    ' Initialize the count
    redCount = 0
    
    ' Loop through each footnote in the document
    For Each footnote In ActiveDocument.Footnotes
        If footnote.Reference.font.color = redColor Then
            redCount = redCount + 1
        End If
    Next footnote
    
    ' Print the count to the console
    'Debug.Print "Number of footnote references with red font: " & redCount
    CountRedFootnoteReferences = redCount
End Function

Private Function CountTotalParagraphs() As Long
    CountTotalParagraphs = ActiveDocument.paragraphs.count
End Function

Private Function CountSectionsWithDifferentFirstPage() As Integer
    Dim doc As Document
    Dim sec As section
    Dim count As Integer
    
    Application.ScreenUpdating = False
    
    Set doc = ActiveDocument
    count = 0
    
    ' Loop through each section in the document
    For Each sec In doc.Sections
        ' Check if the section does not have "Different First Page" selected
        If sec.pageSetup.DifferentFirstPageHeaderFooter Then
            count = count + 1
        End If
    Next sec
    
    Application.ScreenUpdating = True
    
    'MsgBox "Number of sections with 'Different First Page' selected: " & count
    CountSectionsWithDifferentFirstPage = count
End Function

Private Function CountWhiteParagraphMarks() As Integer
    Dim doc As Document
    Dim rng As range
    Dim count As Integer
    Dim firstFound As Boolean
    
    Set doc = ActiveDocument
    count = 0
    firstFound = False
    
    ' Initialize the range to search the entire document
    Set rng = doc.content
    
    ' Use the Find method to locate white paragraph marks
    With rng.Find
        .ClearFormatting
        .font.color = RGB(255, 255, 255) ' RGB value for white
        .text = "^p"
        .Forward = True
        .Wrap = wdFindStop
        
        ' Loop through all occurrences
        Do While .Execute
            ' Check if the found paragraph mark has the white color
            If rng.font.color = RGB(255, 255, 255) Then
                count = count + 1
                ' Go to the first white paragraph mark found
                If Not firstFound Then
                    rng.Select
                    firstFound = True
                End If
            End If
        Loop
    End With
    
    ' Display the count of white paragraph marks
    'MsgBox "Number of white paragraph marks: " & count
    CountWhiteParagraphMarks = count
End Function

Private Function CountEmptyParasWithNoThemeColor()
' The value of wdNotThemeColor is -1. This constant indicates that no theme color is applied.
    Dim para As paragraph
    Dim rng As range
    Dim emptyParaCount As Integer
    Dim themeColorUsed As Boolean
    Dim totalParaCount As Integer
    Dim colorCounts As Object
    Dim colorKey As String
    
    Set colorCounts = CreateObject("Scripting.Dictionary")
    emptyParaCount = 0
    totalParaCount = ActiveDocument.paragraphs.count
    
    For Each para In ActiveDocument.paragraphs
        Set rng = para.range
        themeColorUsed = False
        
        ' Check if the paragraph is empty
        If Len(rng.text) = 1 Then ' Only the paragraph mark
            ' Check if the theme color is used
            If rng.font.TextColor.ObjectThemeColor <> wdNotThemeColor Then
                themeColorUsed = True
                colorKey = CStr(rng.font.TextColor.ObjectThemeColor)
                
                If Not colorCounts.Exists(colorKey) Then
                    colorCounts.Add colorKey, 0
                End If
                
                colorCounts(colorKey) = colorCounts(colorKey) + 1
                emptyParaCount = emptyParaCount + 1
            End If
        End If
    Next para
    
    'Debug.Print "Total number of paragraphs: " & totalParaCount
    'Debug.Print "Number of empty paragraphs using theme color: " & emptyParaCount

    ' Print out the color count for each empty paragraph
    'Dim key As Variant
    'For Each key In colorCounts.Keys
    '    Debug.Print "Theme color " & key & " used in " & colorCounts(key) & " empty paragraphs."
    'Next key
    CountEmptyParasWithNoThemeColor = emptyParaCount
End Function

Private Function CountNumberDashNumberInFootnotes() As Integer
    Dim fn As footnote
    Dim regex As Object
    Dim matches As Object
    Dim matchCount As Long

    matchCount = 0

    ' Create RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .pattern = "\b\d{1,4}-\d{1,4}\b" ' Match numbers separated by a dash (e.g., 12-34)
        .Global = True
        .IgnoreCase = True
    End With

    ' Loop through all footnotes
    For Each fn In ActiveDocument.Footnotes
        Set matches = regex.Execute(fn.range.text)
        matchCount = matchCount + matches.count
    Next fn

    'MsgBox "Number of number-dash-number patterns in footnotes: " & matchCount, vbInformation
    'Debug.Print "Number of number-dash-number patterns in footnotes: " & matchCount
    CountNumberDashNumberInFootnotes = matchCount
End Function

Private Function CountFindNumberDashNumber() As Integer
    Dim rng As range
    Dim searchPattern As String
    Dim userInterrupt As VbMsgBoxResult
    Dim count As Integer
    
    searchPattern = "[0-9]{1,}-[0-9]{1,}" ' Pattern to find number-dash-number
    count = 0

    Set rng = ActiveDocument.content
    With rng.Find
        .text = searchPattern
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop

        Do While .Execute
            ' Move the selection to the found text
            rng.Select
            count = count + 1
            'userInterrupt = MsgBox("Found: " & rng.text, vbOKCancel)

            ' Check for user interruption
            'If userInterrupt = vbCancel Then
            '    MsgBox "Search interrupted. Total found: " & count
            '    Exit Function
            'End If

            ' Collapse the range to the end of the found text
            rng.Collapse Direction:=wdCollapseEnd
            rng.End = ActiveDocument.content.End
        
            If count Mod 50 = 0 Then
                DoEvents  ' Keep UI responsive every 50 updates
            End If
        Loop
    End With

    CountFindNumberDashNumber = count
End Function

Private Function CountNonBreakingSpaces() As Integer
    Dim rng As range
    Dim count As Long

    Set rng = ActiveDocument.content
    rng.Find.ClearFormatting

    With rng.Find
        .text = "^s"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
    End With

    count = 0
    Do While rng.Find.Execute
        count = count + 1
        rng.Collapse wdCollapseEnd
    
        If count Mod 50 = 0 Then
            DoEvents  ' Keep UI responsive every 50 updates
        End If
    Loop

    CountNonBreakingSpaces = count
End Function

Private Function CountPeriodSpaceLeftParenthesis() As Integer
    Dim rng As range
    Dim count As Long
    
    Set rng = ActiveDocument.content
    rng.Find.ClearFormatting
    
    With rng.Find
        .text = ". ("
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
    End With
    
    count = 0
    Do While rng.Find.Execute
        count = count + 1
        rng.Collapse wdCollapseEnd
    
        If count Mod 50 = 0 Then
            DoEvents  ' Keep UI responsive every 50 updates
        End If
    Loop
    
    CountPeriodSpaceLeftParenthesis = count
End Function

Private Function CountStyleWithNumberAndSpace() As Integer
    Dim rng As range
    Dim count As Long
    Dim styleName As String

    styleName = "Verse marker"

    Set rng = ActiveDocument.content
    rng.Find.ClearFormatting
    rng.Find.style = styleName

    With rng.Find
        .text = "[0-9] "
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchWildcards = True
    End With

    count = 0
    Do While rng.Find.Execute
        count = count + 1
        rng.Collapse wdCollapseEnd
        
        If count Mod 50 = 0 Then
            DoEvents  ' Keep UI responsive every 50 updates
        End If
    Loop

    CountStyleWithNumberAndSpace = count
End Function

Private Function CountStyleWithSpaceAndNumber() As Integer
    Dim rng As range
    Dim count As Long
    Dim styleName As String
    Dim firstFound As Boolean
    
    styleName = "Chapter Verse marker"

    Set rng = ActiveDocument.content
    rng.Find.ClearFormatting
    rng.Find.style = styleName
    firstFound = False

    With rng.Find
        .text = " [0-9]"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchWildcards = True
    End With

    count = 0
    Do While rng.Find.Execute
        count = count + 1
            If Not firstFound Then
                rng.Select
                firstFound = True
                'Stop
                Exit Do ' Exit the loop after finding the first match
            End If
        rng.Collapse wdCollapseEnd
    
        If count Mod 50 = 0 Then
            DoEvents  ' Keep UI responsive every 50 updates
        End If
    Loop

    CountStyleWithSpaceAndNumber = count
End Function

Private Function CountQuadrupleParagraphMarks() As Integer
    Dim rng As range
    Dim count As Long
    
    Set rng = ActiveDocument.content
    rng.Find.ClearFormatting
    
    With rng.Find
        .text = "^13^13^13^13"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
    End With
    
    count = 0
    Do While rng.Find.Execute
        count = count + 1
        rng.Collapse wdCollapseEnd
    Loop
    
    'MsgBox "CountQuadrupleParagraphMarks = " & count
    'Dim answer As Integer
    'answer = MsgBox("CountQuadrupleParagraphMarks Do you want to continue?", vbQuestion + vbYesNo, "Continue or Stop")
    'If answer = vbYes Then
    '    Resume
    'Else
    '    Stop
    'End If
    CountQuadrupleParagraphMarks = count
End Function

Private Function CountWhiteSpaceAndCarriageReturn() As Integer
    Dim rng As range
    Dim count As Long

    Set rng = ActiveDocument.content
    rng.Find.ClearFormatting
    rng.Find.font.color = wdColorWhite

    ' Set the Find properties
    With rng.Find
        .text = " ^13"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchWildcards = False
    End With

    count = 0
    Do While rng.Find.Execute
        count = count + 1
        rng.Collapse wdCollapseEnd
    Loop

    CountWhiteSpaceAndCarriageReturn = count
End Function

Private Function CountDoubleTabs() As Integer
    Dim doc As Document
    Dim rng As range
    Dim count As Long

    count = 0
    Set doc = ActiveDocument
    Set rng = doc.content

    Do While rng.Find.Execute(findText:="^t^t", Forward:=True, Wrap:=wdFindStop)
        count = count + 1
        rng.Collapse Direction:=wdCollapseEnd ' Move the range to the end of the found text
    Loop

    CountDoubleTabs = count
End Function

Private Function CountSpaceFollowedByCarriageReturn() As Integer
    Dim doc As Document
    Dim rng As range
    Dim count As Long

    count = 0
    Set doc = ActiveDocument
    Set rng = doc.content

    Do While rng.Find.Execute(findText:=" ^13", Forward:=True, Wrap:=wdFindStop)
        count = count + 1
        rng.Collapse Direction:=wdCollapseEnd ' Move the range to the end of the found text
    Loop

    CountSpaceFollowedByCarriageReturn = count
End Function

Private Function CountDoubleSpaces() As Integer
    Dim doc As Document
    Dim rng As range
    Dim doubleSpaceCount As Long
    
    doubleSpaceCount = 0
    Set doc = ActiveDocument
    Set rng = doc.content
    
    With rng.Find
        .text = "  " ' Double space
        .Format = False
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        Do While .Execute
            doubleSpaceCount = doubleSpaceCount + 1
            rng.Collapse wdCollapseEnd
        Loop
    End With
    
    CountDoubleSpaces = doubleSpaceCount
End Function

' Function to count occurrences of a substring in a string
Function CountOccurrences(ByVal text As String, ByVal substring As String) As Long
    Dim pos, count As Long
    pos = 1
    count = 0
    Do While InStr(pos, text, substring) > 0
        pos = InStr(pos, text, substring) + Len(substring)
        count = count + 1
    Loop
    
    CountOccurrences = count
End Function

' count double spaces inside shapes, including those within groups, with error handling
Private Function CountDoubleSpacesInShapes() As Integer
    Dim doc As Document
    Dim shp As shape
    Dim groupShp As shape
    Dim doubleSpaceCount As Long

    doubleSpaceCount = 0
    Set doc = ActiveDocument
    
    On Error Resume Next  ' Ignore errors and continue
    For Each shp In doc.Shapes
        Call ProcessShape(shp, doubleSpaceCount)
    Next shp
    On Error GoTo 0  ' Turn off error handling
    
    ' Display the count of double spaces in shapes
    CountDoubleSpacesInShapes = doubleSpaceCount
End Function

Sub ProcessShape(ByVal shp As shape, ByRef doubleSpaceCount As Long)
    Dim textRange As range
    Dim shapeText As String
    
    ' Check if the shape has a text frame
    If shp.Type = msoGroup Then
        ' Process each shape within the group
        Dim groupShp As shape
        For Each groupShp In shp.GroupItems
            Call ProcessShape(groupShp, doubleSpaceCount)
        Next groupShp
    Else
        If Not shp.TextFrame Is Nothing Then
            If shp.TextFrame.HasText Then
                ' Get the text range within the shape
                Set textRange = shp.TextFrame.textRange
                ' Get the text content of the shape
                shapeText = textRange.text
                ' Count double spaces in the shape's text
                doubleSpaceCount = doubleSpaceCount + CountOccurrences(shapeText, "  ")
            End If
        ElseIf Not shp.TextFrame2 Is Nothing Then
            If shp.TextFrame2.HasText Then
                ' Get the text range within the shape
                Set textRange = shp.TextFrame2.textRange
                ' Get the text content of the shape
                shapeText = textRange.text
                ' Count double spaces in the shape's text
                doubleSpaceCount = doubleSpaceCount + CountOccurrences(shapeText, "  ")
            End If
        End If
    End If
End Sub

