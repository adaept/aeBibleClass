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
Private Const bTimeAllTests As Boolean = False
Private Const bGoTo16 As Boolean = False
' Used to store Expected results
Private oneBasedExpectedArray As Variant
Private ResultArray(1 To 45) As Integer
Private result As Integer
Private TestReportFileName As String
Private testFileNum As Integer
Private Const TestReportFlag As Boolean = True
Private OneTest As Integer  ' Stores the value to run a single test
'==================================================================

Private Sub Class_Initialize()
    On Error GoTo 0
    'MsgBox Application.VBE.ActiveVBProject.Name
    'MsgBox Application.Caption
    If Application.VBE.ActiveVBProject.name = "BibleClass" Then
        Application.Caption = Application.VBE.ActiveVBProject.name & " " & BibleClassVERSION
    End If
    Debug.Print "Class_Initialize"
    BibleClassVERSION = "0.1.0"
    BibleClassVERSION_DATE = "July 19, 2025"
    ' Reset OneTest for each run
    OneTest = 0
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
        Debug.Print ">Running Test " & OneTest
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
    Dim i As Integer
    For i = 1 To 45
        ResultArray(i) = -1
    Next i
    ' Optional: Print to Immediate window to confirm
    'For i = 1 To 45
    '    Debug.Print "ResultArray(" & i & ") = " & ResultArray(i),
    'Next i
End Sub

Private Function ConvertToOneBasedArray(values As Variant) As Variant
    Dim strArray() As String
    Dim i As Integer

    InitializeGlobalResultArrayToMinusOne

    ' Resize the array to start from index 1
    ReDim strArray(1 To UBound(values) - LBound(values) + 1)
    
    ' Store the values in the array starting from index 1
    For i = 1 To UBound(values) - LBound(values) + 1
        strArray(i) = values(i - 1 + LBound(values))
    Next i
    
    ' Return the 1-based array
    ConvertToOneBasedArray = strArray
End Function

Private Sub Expected1BasedArray()
    Dim doc As Document
    Dim values As Variant
    Dim i As Integer
    Dim outputString As String

    Set doc = ActiveDocument
    TestReportFileName = doc.Path & "\TestReport.txt"
    
    ' Delete the old debug file if it exists
    If Dir(TestReportFileName) <> "" Then
        Kill TestReportFileName
    End If
    
    ' Open the debug file for writing
    testFileNum = FreeFile
    Open TestReportFileName For Output As testFileNum
    Close testFileNum
    
    ' Define the Expected RunTest result values to store in the array
    '      RunTest 1  2  3  4  5  6  7  8  9  10 11 12 13 14 15 16    17  18  19 20 21 22   23 24    25 26 27   28  29  30 31 32   33   34   35   36 37  38   39 40 41 42 43)
    values = Array(0, 0, 0, 0, 0, 0, 0, 0, 7, 1, 0, 0, 0, 0, 0, 16474, 0, 66, 0, 0, 0, 153, 0, 1000, 1, 0, 147, 77, 70, 0, 1, 648, 648, 117, 117, 0, 19, 153, 0, 0, 0, 0, 0)

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
    If TestReportFlag Then AppendToFile TestReportFileName, outputString
    
    ' Next 15 results
    outputString = ""
    For i = 16 To 30
        outputString = outputString & " Test(" & i & ")" & oneBasedExpectedArray(i) & " "
        ' Trim the trailing space
        outputString = Trim(outputString)
    Next i
    Debug.Print outputString
    If TestReportFlag Then AppendToFile TestReportFileName, outputString
    
    ' Next results
    outputString = ""
    For i = 31 To UBound(oneBasedExpectedArray)
        outputString = outputString & " Test(" & i & ")" & oneBasedExpectedArray(i) & " "
        ' Trim the trailing space
        outputString = Trim(outputString)
    Next i
    Debug.Print outputString
    If TestReportFlag Then AppendToFile TestReportFileName, outputString

End Sub

Sub AppendToFile(filePath As String, text As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Append As fileNum
    Print #fileNum, text
    Close fileNum
End Sub

Private Function LogMessage(msg As String) As String
    LogMessage = "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "] " & msg
End Function

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
        Debug.Print ">>Running Test " & OneTest
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
                        
            Call Expected1BasedArray
            Debug.Print "1234567890123412345678901234123456789012341234567890123412345678901234"
            '            123456789    1234567    12345678    123456    12345678    12345678     NOTE: Comma separation of debug is 14 characters
            Debug.Print "Pass/Fail", "Copy ()", "Test Num", "Result", "Expected", "Function"
            If TestReportFlag Then AppendToFile TestReportFileName, "1234567890123412345678901234123456789012341234567890123412345678901234"
            If TestReportFlag Then AppendToFile TestReportFileName, "Pass/Fail" & space(5) & "Copy ()" & space(7) & "Test Num" & space(6) & "Result" & space(8) & "Expected" & space(6) & "Function"
            
            If OneTest <> 0 Then
                RunTest (OneTest)
                Exit Function
            End If
            
            If bGoTo16 Then GoTo Test16
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
Test16:     RunTest (16)
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
            RunTest (42)
            RunTest (43)
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
    Debug.Print "Total Total Runtime: " & Format(runTime, "0.00") & " seconds"
    If TestReportFlag Then AppendToFile TestReportFileName, "Total Total Runtime: " & Format(runTime, "0.00") & " seconds"
    
    Debug.Print LogMessage("Completed")
    If TestReportFlag Then AppendToFile TestReportFileName, LogMessage("Completed")

    Debug.Print "BibleClass VERSION: " & BibleClassVERSION
    If TestReportFlag Then AppendToFile TestReportFileName, "BibleClass VERSION: " & BibleClassVERSION
    Debug.Print "BibleClass VERSION_DATE: " & BibleClassVERSION_DATE
    If TestReportFlag Then AppendToFile TestReportFileName, "BibleClass VERSION_DATE: " & BibleClassVERSION_DATE

    RunBibleClassTests = True

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RunBibleClassTests of Class BibleClass"
    Resume PROC_EXIT

End Function

Private Function GetPassFail(TestNum As Integer) As String
    Dim PassFail As String

    On Error GoTo 0
    ' Set explicit fail unless the test passes
    GetPassFail = "FAIL!!!!"
    Select Case TestNum
    Case 1
        ResultArray(1) = CountDoubleSpaces
        If ResultArray(1) = oneBasedExpectedArray(1) Then GetPassFail = "PASS    "
    Case 2
        ResultArray(2) = CountDoubleSpacesInShapes
        If ResultArray(2) = oneBasedExpectedArray(2) Then GetPassFail = "PASS    "
    Case 3
        ResultArray(3) = CountSpaceFollowedByCarriageReturn
        If ResultArray(3) = oneBasedExpectedArray(3) Then GetPassFail = "PASS    "
    Case 4
        ResultArray(4) = CountDoubleTabs
        If ResultArray(4) = oneBasedExpectedArray(4) Then GetPassFail = "PASS    "
    Case 5
        ResultArray(5) = CountWhiteSpaceAndCarriageReturn
        If ResultArray(5) = oneBasedExpectedArray(5) Then GetPassFail = "PASS    "
    Case 6
        ResultArray(6) = CountQuadrupleParagraphMarks
        If ResultArray(6) = oneBasedExpectedArray(6) Then GetPassFail = "PASS    "
    Case 7
        ResultArray(7) = CountStyleWithSpaceAndNumber
        If ResultArray(7) = oneBasedExpectedArray(7) Then GetPassFail = "PASS    "
    Case 8
        ResultArray(8) = CountStyleWithNumberAndSpace
        If ResultArray(8) = oneBasedExpectedArray(8) Then GetPassFail = "PASS    "
    Case 9
        ResultArray(9) = CountPeriodSpaceLeftParenthesis
        If ResultArray(9) = oneBasedExpectedArray(9) Then GetPassFail = "PASS    "
    Case 10
        ResultArray(10) = CountNonBreakingSpaces
        If ResultArray(10) = oneBasedExpectedArray(10) Then GetPassFail = "PASS    "
    Case 11
        ResultArray(11) = CountFindNumberDashNumber
        If ResultArray(11) = oneBasedExpectedArray(11) Then GetPassFail = "PASS    "
    Case 12
        ResultArray(12) = CountNumberDashNumberInFootnotes
        If ResultArray(12) = oneBasedExpectedArray(12) Then GetPassFail = "PASS    "
    Case 13
        ResultArray(13) = CountEmptyParasWithNoThemeColor
        If ResultArray(13) = oneBasedExpectedArray(13) Then GetPassFail = "PASS    "
    Case 14
        ResultArray(14) = CountWhiteParagraphMarks
        If ResultArray(14) = oneBasedExpectedArray(14) Then GetPassFail = "PASS    "
    Case 15
        ResultArray(15) = CountSectionsWithDifferentFirstPage
        If ResultArray(15) = oneBasedExpectedArray(15) Then GetPassFail = "PASS    "
    Case 16
        ResultArray(16) = CountTotalParagraphs
        If ResultArray(16) = oneBasedExpectedArray(16) Then GetPassFail = "PASS    "
    Case 17
        ResultArray(17) = CountRedFootnoteReferences
        If ResultArray(17) = oneBasedExpectedArray(17) Then GetPassFail = "PASS    "
    Case 18
        ResultArray(18) = CountHeading1
        If ResultArray(18) = oneBasedExpectedArray(18) Then GetPassFail = "PASS    "
    Case 19
        ResultArray(19) = CountEmptyParasAfterH2
        If ResultArray(19) = oneBasedExpectedArray(19) Then GetPassFail = "PASS    "
    Case 20
        ResultArray(20) = CountFootnotesFollowedByDigit
        If ResultArray(20) = oneBasedExpectedArray(20) Then GetPassFail = "PASS    "
    Case 21
        ResultArray(21) = CountNotSpacesAfterFootnoteReferences
        If ResultArray(21) = oneBasedExpectedArray(21) Then GetPassFail = "PASS    "
    Case 22
        ResultArray(22) = CountEmptyParagraphsWithFormatting
        If ResultArray(22) = oneBasedExpectedArray(22) Then GetPassFail = "PASS    "
    Case 23
        ResultArray(23) = CountDeleteEmptyParagraphsBeforeHeading2
        If ResultArray(23) = oneBasedExpectedArray(23) Then GetPassFail = "PASS    "
    Case 24
        ResultArray(24) = CountFootnoteReferences
        If ResultArray(24) = oneBasedExpectedArray(24) Then GetPassFail = "PASS    "
    Case 25
        ResultArray(25) = CountFootnoteReferenceColors
        If ResultArray(25) = oneBasedExpectedArray(25) Then GetPassFail = "PASS    "
    Case 26
        ResultArray(26) = CheckAllHeaders("Empty")
        If ResultArray(26) = oneBasedExpectedArray(26) Then GetPassFail = "PASS    "
    Case 27
        ResultArray(27) = CheckAllHeaders("NotEmpty")
        If ResultArray(27) = oneBasedExpectedArray(27) Then GetPassFail = "PASS    "
    Case 28
        ResultArray(28) = CountTabFollowedByParagraphMarkInHeaders
        If ResultArray(28) = oneBasedExpectedArray(28) Then GetPassFail = "PASS    "
    Case 29
        ResultArray(29) = CountParagraphsWithoutTabInHeaders
        If ResultArray(29) = oneBasedExpectedArray(29) Then GetPassFail = "PASS    "
    Case 30
        ResultArray(30) = CountHeaderStyleUsage
        If ResultArray(30) = oneBasedExpectedArray(30) Then GetPassFail = "PASS    "
    Case 31
        ResultArray(31) = CountParagraphMarksPerHeaderSection
        If ResultArray(31) = oneBasedExpectedArray(31) Then GetPassFail = "PASS    "
    Case 32
        ResultArray(32) = CountLinefeed
        If ResultArray(32) = oneBasedExpectedArray(32) Then GetPassFail = "PASS    "
    Case 33
        ResultArray(33) = CountLinefeed(" ")
        If ResultArray(33) = oneBasedExpectedArray(33) Then GetPassFail = "PASS    "
    Case 34
        ResultArray(34) = CountManualLineBreaksAndWithSpace
        If ResultArray(34) = oneBasedExpectedArray(34) Then GetPassFail = "PASS    "
    Case 35
        ResultArray(35) = CountManualLineBreaksAndWithSpace(" ")
        If ResultArray(35) = oneBasedExpectedArray(35) Then GetPassFail = "PASS    "
    Case 36
        ResultArray(36) = CountFooterParagraphsWithFooterStyle
        If ResultArray(36) = oneBasedExpectedArray(36) Then GetPassFail = "PASS    "
    Case 37
        ResultArray(37) = CountDocTabOnlyParagraphs
        If ResultArray(37) = oneBasedExpectedArray(37) Then GetPassFail = "PASS    "
    Case 38
        ResultArray(38) = CountEmptyParagraphs
        If ResultArray(38) = oneBasedExpectedArray(38) Then GetPassFail = "PASS    "
    Case 39
        ResultArray(39) = CountParagraphMarks_ArialBlack
        If ResultArray(39) = oneBasedExpectedArray(39) Then GetPassFail = "PASS    "
    Case 40
        ResultArray(40) = CountParagraphMarks_ArialBlackDarkRed
        If ResultArray(40) = oneBasedExpectedArray(40) Then GetPassFail = "PASS    "
    Case 41
        ResultArray(41) = Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed
        If ResultArray(41) = oneBasedExpectedArray(41) Then GetPassFail = "PASS    "
    Case 42
        ResultArray(42) = CountBoldFootnotesWordLevel
        If ResultArray(42) = oneBasedExpectedArray(42) Then GetPassFail = "PASS    "
    Case 43
        ResultArray(43) = CountDarkRedStyledParagraphMarks
        If ResultArray(43) = oneBasedExpectedArray(43) Then GetPassFail = "PASS    "
    Case Else
        Debug.Print "TestNum = " & TestNum & " >The test number is outside the accepted range"
    End Select
End Function

Private Function Fixed14Str(MyResult As String) As String
    Const fixedWidth = 14
    Fixed14Str = Left(MyResult & String(fixedWidth, " "), fixedWidth)
    'Debug.Print "The fixed width string is:" & vbCrLf & "[" & Fixed14Str & "]"
    'Debug.Print "[12345678901234]"
End Function

Private Function RunTest(num As Integer) As Boolean
    Dim startTime As Double
    Dim endTime As Double
    Dim runTime As Double
    
    ' Record the start timer for each test
    startTime = Timer
    
    On Error GoTo PROC_ERR
    Select Case num
    Case 1
        Debug.Print GetPassFail(1), "Copy ()", "Test = " & num, ResultArray(1), oneBasedExpectedArray(1), "CountDoubleSpaces"
        If TestReportFlag Then OutputTestReport (1)
    Case 2
        Debug.Print GetPassFail(2), "Copy ()", "Test = " & num, ResultArray(2), oneBasedExpectedArray(2), "CountDoubleSpacesInShapes"
        If TestReportFlag Then OutputTestReport (2)
    Case 3
        Debug.Print GetPassFail(3), "Copy ()", "Test = " & num, ResultArray(3), oneBasedExpectedArray(3), "CountSpaceFollowedByCarriageReturn"
        If TestReportFlag Then OutputTestReport (3)
    Case 4
        Debug.Print GetPassFail(4), "Copy ()", "Test = " & num, ResultArray(4), oneBasedExpectedArray(4), "CountDoubleTabs"
        If TestReportFlag Then OutputTestReport (4)
    Case 5
        Debug.Print GetPassFail(5), "Copy ()", "Test = " & num, ResultArray(5), oneBasedExpectedArray(5), "CountWhiteSpaceAndCarriageReturn"
        If TestReportFlag Then OutputTestReport (5)
    Case 6
        Debug.Print GetPassFail(6), "Copy ()", "Test = " & num, ResultArray(6), oneBasedExpectedArray(6), "CountQuadrupleParagraphMarks"
        If TestReportFlag Then OutputTestReport (6)
    Case 7
        Debug.Print GetPassFail(7), "Copy ()", "Test = " & num, ResultArray(7), oneBasedExpectedArray(7), "CountStyleWithSpaceAndNumber"
        If TestReportFlag Then OutputTestReport (7)
    Case 8
        Debug.Print GetPassFail(8), "Copy ()", "Test = " & num, ResultArray(8), oneBasedExpectedArray(8), "CountStyleWithNumberAndSpace"
        If TestReportFlag Then OutputTestReport (8)
    Case 9
        Debug.Print GetPassFail(9), "Copy ()", "Test = " & num, ResultArray(9), oneBasedExpectedArray(9), "CountPeriodSpaceLeftParenthesis"
        If TestReportFlag Then OutputTestReport (9)
    Case 10
        Debug.Print GetPassFail(10), "Copy ()", "Test = " & num, ResultArray(10), oneBasedExpectedArray(10), "CountNonBreakingSpaces"
        If TestReportFlag Then OutputTestReport (10)
    Case 11
        Debug.Print GetPassFail(11), "Copy ()", "Test = " & num, ResultArray(11), oneBasedExpectedArray(11), "CountFindNumberDashNumber"
        If TestReportFlag Then OutputTestReport (11)
    Case 12
        Debug.Print GetPassFail(12), "Copy ()", "Test = " & num, ResultArray(12), oneBasedExpectedArray(12), "CountNumberDashNumberInFootnotes"
        If TestReportFlag Then OutputTestReport (12)
    Case 13
        Debug.Print GetPassFail(13), "Copy ()", "Test = " & num, ResultArray(13), oneBasedExpectedArray(13), "CountEmptyParasWithNoThemeColor"
        If TestReportFlag Then OutputTestReport (13)
    Case 14
        Debug.Print GetPassFail(14), "Copy ()", "Test = " & num, ResultArray(14), oneBasedExpectedArray(14), "CountWhiteParagraphMarks"
        If TestReportFlag Then OutputTestReport (14)
    Case 15
        Debug.Print GetPassFail(15), "Copy ()", "Test = " & num, ResultArray(15), oneBasedExpectedArray(15), "CountSectionsWithDifferentFirstPage"
        If TestReportFlag Then OutputTestReport (15)
    Case 16
        Debug.Print GetPassFail(16), "Copy ()", "Test = " & num, ResultArray(16), oneBasedExpectedArray(16), "CountTotalParagraphs"
        If TestReportFlag Then OutputTestReport (16)
    Case 17
        Debug.Print GetPassFail(17), "Copy ()", "Test = " & num, ResultArray(17), oneBasedExpectedArray(17), "CountRedFootnoteReferences"
        If TestReportFlag Then OutputTestReport (17)
    Case 18
        Debug.Print GetPassFail(18), "Copy ()", "Test = " & num, ResultArray(18), oneBasedExpectedArray(18), "CountHeading1"
        If TestReportFlag Then OutputTestReport (18)
    Case 19
        Debug.Print GetPassFail(19), "Copy ()", "Test = " & num, ResultArray(19), oneBasedExpectedArray(19), "CountEmptyParasAfterH2"
        If TestReportFlag Then OutputTestReport (19)
    Case 20
        Debug.Print GetPassFail(20), "Copy ()", "Test = " & num, ResultArray(20), oneBasedExpectedArray(20), "CountFootnotesFollowedByDigit"
        If TestReportFlag Then OutputTestReport (20)
    Case 21
        Debug.Print GetPassFail(21), "Copy ()", "Test = " & num, ResultArray(21), oneBasedExpectedArray(21), "CountNotSpacesAfterFootnoteReferences"
        If TestReportFlag Then OutputTestReport (21)
    Case 22
        Debug.Print GetPassFail(22), "Copy ()", "Test = " & num, ResultArray(22), oneBasedExpectedArray(22), "CountEmptyParagraphsWithFormatting"
        If TestReportFlag Then OutputTestReport (22)
    Case 23
        Debug.Print GetPassFail(23), "Copy ()", "Test = " & num, ResultArray(23), oneBasedExpectedArray(23), "CountDeleteEmptyParagraphsBeforeHeading2"
        If TestReportFlag Then OutputTestReport (23)
    Case 24
        Debug.Print GetPassFail(24), "Copy ()", "Test = " & num, ResultArray(24), oneBasedExpectedArray(24), "CountFootnoteReferences"
        If TestReportFlag Then OutputTestReport (24)
    Case 25
        Debug.Print GetPassFail(25), "Copy ()", "Test = " & num, ResultArray(25), oneBasedExpectedArray(25), "CountFootnoteReferenceColors"
        If TestReportFlag Then OutputTestReport (25)
    Case 26
        Debug.Print GetPassFail(26), "Copy ()", "Test = " & num, ResultArray(26), oneBasedExpectedArray(26), "CheckAllHeaders(" & "Empty" & ")"
        If TestReportFlag Then OutputTestReport (26)
    Case 27
        Debug.Print GetPassFail(27), "Copy ()", "Test = " & num, ResultArray(27), oneBasedExpectedArray(27), "CheckAllHeaders(" & "NotEmpty" & ")"
        If TestReportFlag Then OutputTestReport (27)
    Case 28
        Debug.Print GetPassFail(28), "Copy ()", "Test = " & num, ResultArray(28), oneBasedExpectedArray(28), "CountTabFollowedByParagraphMarkInHeaders"
        If TestReportFlag Then OutputTestReport (28)
    Case 29
        Debug.Print GetPassFail(29), "Copy ()", "Test = " & num, ResultArray(29), oneBasedExpectedArray(29), "CountParagraphsWithoutTabInHeaders"
        If TestReportFlag Then OutputTestReport (29)
    Case 30
        Debug.Print GetPassFail(30), "Copy ()", "Test = " & num, ResultArray(30), oneBasedExpectedArray(30), "CountHeaderStyleUsage"
        If TestReportFlag Then OutputTestReport (30)
    Case 31
        Debug.Print GetPassFail(31), "Copy ()", "Test = " & num, ResultArray(31), oneBasedExpectedArray(31), "CountParagraphMarksPerHeaderSection"
        If TestReportFlag Then OutputTestReport (31)
    Case 32
        Debug.Print GetPassFail(32), "Copy ()", "Test = " & num, ResultArray(32), oneBasedExpectedArray(32), "CountLinefeed"
        If TestReportFlag Then OutputTestReport (32)
    Case 33
        Debug.Print GetPassFail(33), "Copy ()", "Test = " & num, ResultArray(33), oneBasedExpectedArray(33), "CountLinefeed(" & """ """ & ")"
        If TestReportFlag Then OutputTestReport (33)
    Case 34
        Debug.Print GetPassFail(34), "Copy ()", "Test = " & num, ResultArray(34), oneBasedExpectedArray(34), "CountManualLineBreaksAndWithSpace"
        If TestReportFlag Then OutputTestReport (34)
    Case 35
        Debug.Print GetPassFail(35), "Copy ()", "Test = " & num, ResultArray(35), oneBasedExpectedArray(35), "CountManualLineBreaksAndWithSpace(" & """ """ & ")"
        If TestReportFlag Then OutputTestReport (35)
    Case 36
        Debug.Print GetPassFail(36), "Copy ()", "Test = " & num, ResultArray(36), oneBasedExpectedArray(36), "CountFooterParagraphsWithFooterStyle"
        If TestReportFlag Then OutputTestReport (36)
    Case 37
        Debug.Print GetPassFail(37), "Copy ()", "Test = " & num, ResultArray(37), oneBasedExpectedArray(37), "CountDocTabOnlyParagraphs"
        If TestReportFlag Then OutputTestReport (37)
    Case 38
        Debug.Print GetPassFail(38), "Copy ()", "Test = " & num, ResultArray(38), oneBasedExpectedArray(38), "CountEmptyParagraphs"
        If TestReportFlag Then OutputTestReport (38)
    Case 39
        Debug.Print GetPassFail(39), "Copy ()", "Test = " & num, ResultArray(39), oneBasedExpectedArray(39), "CountParagraphMarks_ArialBlack"
        If TestReportFlag Then OutputTestReport (39)
    Case 40
        Debug.Print GetPassFail(40), "Copy ()", "Test = " & num, ResultArray(40), oneBasedExpectedArray(40), "CountParagraphMarks_ArialBlackDarkRed"
        If TestReportFlag Then OutputTestReport (40)
    Case 41
        Debug.Print GetPassFail(41), "Copy ()", "Test = " & num, ResultArray(41), oneBasedExpectedArray(41), "Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed"
        If TestReportFlag Then OutputTestReport (41)
    Case 42
        Debug.Print GetPassFail(42), "Copy ()", "Test = " & num, ResultArray(42), oneBasedExpectedArray(42), "CountBoldFootnotesWordLevel"
        If TestReportFlag Then OutputTestReport (42)
    Case 43
        Debug.Print GetPassFail(43), "Copy ()", "Test = " & num, ResultArray(43), oneBasedExpectedArray(43), "CountDarkRedStyledParagraphMarks"
        If TestReportFlag Then OutputTestReport (43)
   Case Else
        Debug.Print "num = " & num & " >>The test number is outside the accepted range"
    End Select

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
    End If
    'Debug.Print "!!! Error in Test num = " & num, "Function RunTest !!!"
    End
End Function

Private Function OutputTestReport(num As Integer) As Boolean
    On Error GoTo 0
    Select Case num
    Case 1
        AppendToFile TestReportFileName, GetPassFail(1) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(1))) & Fixed14Str(CStr(oneBasedExpectedArray(1))) & "CountDoubleSpaces"
    Case 2
        AppendToFile TestReportFileName, GetPassFail(2) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(2))) & Fixed14Str(CStr(oneBasedExpectedArray(2))) & "CountDoubleSpacesInShapes"
    Case 3
        AppendToFile TestReportFileName, GetPassFail(3) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(3))) & Fixed14Str(CStr(oneBasedExpectedArray(3))) & "CountSpaceFollowedByCarriageReturn"
    Case 4
        AppendToFile TestReportFileName, GetPassFail(4) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(4))) & Fixed14Str(CStr(oneBasedExpectedArray(4))) & "CountDoubleTabs"
    Case 5
        AppendToFile TestReportFileName, GetPassFail(5) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(5))) & Fixed14Str(CStr(oneBasedExpectedArray(5))) & "CountWhiteSpaceAndCarriageReturn"
    Case 6
        AppendToFile TestReportFileName, GetPassFail(6) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(6))) & Fixed14Str(CStr(oneBasedExpectedArray(6))) & "CountQuadrupleParagraphMarks"
    Case 7
        AppendToFile TestReportFileName, GetPassFail(7) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(7))) & Fixed14Str(CStr(oneBasedExpectedArray(7))) & "CountStyleWithSpaceAndNumber"
    Case 8
        AppendToFile TestReportFileName, GetPassFail(8) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(8))) & Fixed14Str(CStr(oneBasedExpectedArray(8))) & "CountStyleWithNumberAndSpace"
    Case 9
        AppendToFile TestReportFileName, GetPassFail(9) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(9))) & Fixed14Str(CStr(oneBasedExpectedArray(9))) & "CountPeriodSpaceLeftParenthesis"
    Case 10
        AppendToFile TestReportFileName, GetPassFail(11) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(10))) & Fixed14Str(CStr(oneBasedExpectedArray(10))) & "CountNonBreakingSpaces"
    Case 11
        AppendToFile TestReportFileName, GetPassFail(11) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(11))) & Fixed14Str(CStr(oneBasedExpectedArray(11))) & "CountFindNumberDashNumber"
    Case 12
        AppendToFile TestReportFileName, GetPassFail(12) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(12))) & Fixed14Str(CStr(oneBasedExpectedArray(12))) & "CountNumberDashNumberInFootnotes"
    Case 13
        AppendToFile TestReportFileName, GetPassFail(13) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(13))) & Fixed14Str(CStr(oneBasedExpectedArray(13))) & "CountEmptyParasWithNoThemeColor"
    Case 14
        AppendToFile TestReportFileName, GetPassFail(14) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(14))) & Fixed14Str(CStr(oneBasedExpectedArray(14))) & "CountWhiteParagraphMarks"
    Case 15
        AppendToFile TestReportFileName, GetPassFail(15) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(15))) & Fixed14Str(CStr(oneBasedExpectedArray(15))) & "CountSectionsWithDifferentFirstPage"
    Case 16
        AppendToFile TestReportFileName, GetPassFail(16) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(16))) & Fixed14Str(CStr(oneBasedExpectedArray(16))) & "CountTotalParagraphs"
    Case 17
        AppendToFile TestReportFileName, GetPassFail(17) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(17))) & Fixed14Str(CStr(oneBasedExpectedArray(17))) & "CountRedFootnoteReferences"
    Case 18
        AppendToFile TestReportFileName, GetPassFail(18) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(18))) & Fixed14Str(CStr(oneBasedExpectedArray(18))) & "CountHeading1"
    Case 19
        AppendToFile TestReportFileName, GetPassFail(19) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(19))) & Fixed14Str(CStr(oneBasedExpectedArray(19))) & "CountEmptyParasAfterH2"
    Case 20
        AppendToFile TestReportFileName, GetPassFail(20) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(20))) & Fixed14Str(CStr(oneBasedExpectedArray(20))) & "CountFootnotesFollowedByDigit"
    Case 21
        AppendToFile TestReportFileName, GetPassFail(21) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(21))) & Fixed14Str(CStr(oneBasedExpectedArray(21))) & "CountNotSpacesAfterFootnoteReferences"
    Case 22
        AppendToFile TestReportFileName, GetPassFail(22) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(22))) & Fixed14Str(CStr(oneBasedExpectedArray(22))) & "CountEmptyParagraphsWithFormatting"
    Case 23
        AppendToFile TestReportFileName, GetPassFail(23) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(23))) & Fixed14Str(CStr(oneBasedExpectedArray(23))) & "CountDeleteEmptyParagraphsBeforeHeading2"
    Case 24
        AppendToFile TestReportFileName, GetPassFail(24) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(24))) & Fixed14Str(CStr(oneBasedExpectedArray(24))) & "CountFootnoteReferences"
    Case 25
        AppendToFile TestReportFileName, GetPassFail(25) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(25))) & Fixed14Str(CStr(oneBasedExpectedArray(25))) & "CountFootnoteReferenceColors"
    Case 26
        AppendToFile TestReportFileName, GetPassFail(26) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(26))) & Fixed14Str(CStr(oneBasedExpectedArray(26))) & "CheckAllHeaders(" & "Empty" & ")"
    Case 27
        AppendToFile TestReportFileName, GetPassFail(27) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(27))) & Fixed14Str(CStr(oneBasedExpectedArray(27))) & "CheckAllHeaders(" & "NotEmpty" & ")"
    Case 28
        AppendToFile TestReportFileName, GetPassFail(28) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(28))) & Fixed14Str(CStr(oneBasedExpectedArray(28))) & "CountTabFollowedByParagraphMarkInHeaders"
    Case 29
        AppendToFile TestReportFileName, GetPassFail(29) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(29))) & Fixed14Str(CStr(oneBasedExpectedArray(29))) & "CountParagraphsWithoutTabInHeaders"
    Case 30
        AppendToFile TestReportFileName, GetPassFail(30) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(30))) & Fixed14Str(CStr(oneBasedExpectedArray(30))) & "CountHeaderStyleUsage"
    Case 31
        AppendToFile TestReportFileName, GetPassFail(31) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(31))) & Fixed14Str(CStr(oneBasedExpectedArray(31))) & "CountParagraphMarksPerHeaderSection"
    Case 32
        AppendToFile TestReportFileName, GetPassFail(32) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(32))) & Fixed14Str(CStr(oneBasedExpectedArray(32))) & "CountLinefeed"
    Case 33
        AppendToFile TestReportFileName, GetPassFail(33) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(33))) & Fixed14Str(CStr(oneBasedExpectedArray(33))) & "CountLinefeed(" & """ """ & ")"
    Case 34
        AppendToFile TestReportFileName, GetPassFail(34) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(34))) & Fixed14Str(CStr(oneBasedExpectedArray(34))) & "CountManualLineBreaksAndWithSpace"
    Case 35
        AppendToFile TestReportFileName, GetPassFail(35) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(35))) & Fixed14Str(CStr(oneBasedExpectedArray(35))) & "CountManualLineBreaksAndWithSpace(" & """ """ & ")"
    Case 36
        AppendToFile TestReportFileName, GetPassFail(36) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(36))) & Fixed14Str(CStr(oneBasedExpectedArray(36))) & "CountFooterParagraphsWithFooterStyle"
    Case 37
        AppendToFile TestReportFileName, GetPassFail(37) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(37))) & Fixed14Str(CStr(oneBasedExpectedArray(37))) & "CountDocTabOnlyParagraphs"
    Case 38
        AppendToFile TestReportFileName, GetPassFail(38) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(38))) & Fixed14Str(CStr(oneBasedExpectedArray(38))) & "CountEmptyParagraphs"
    Case 39
        AppendToFile TestReportFileName, GetPassFail(39) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(39))) & Fixed14Str(CStr(oneBasedExpectedArray(39))) & "CountParagraphMarks_ArialBlack"
    Case 40
        AppendToFile TestReportFileName, GetPassFail(40) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(40))) & Fixed14Str(CStr(oneBasedExpectedArray(40))) & "CountParagraphMarks_ArialBlackDarkRed"
    Case 41
        AppendToFile TestReportFileName, GetPassFail(41) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(41))) & Fixed14Str(CStr(oneBasedExpectedArray(41))) & "Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed"
    Case 42
        AppendToFile TestReportFileName, GetPassFail(42) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(42))) & Fixed14Str(CStr(oneBasedExpectedArray(42))) & "CountBoldFootnotesWordLevel"
    Case 43
        AppendToFile TestReportFileName, GetPassFail(43) & space(10) & "Copy ()" & space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(43))) & Fixed14Str(CStr(oneBasedExpectedArray(43))) & "CountDarkRedStyledParagraphMarks"
   Case Else
        Debug.Print "num = " & num & " >>>The test number is outside the accepted range"
    End Select
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
    Dim word As range
    Dim boldCount As Long
    Dim hasBold As Boolean

    For Each ft In ActiveDocument.Footnotes
        hasBold = False

        If ft.range.style = "Footnote Text" Then
            For Each word In ft.range.Words
                If word.font.Bold = True Then
                    boldCount = boldCount + 1
                    hasBold = True
                    Exit For
                End If
            Next word
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
                    Stop
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
    Dim hdrftr As HeaderFooter
    Dim para As paragraph
    Dim count As Long

    count = 0

    ' Loop through all sections in the document
    For Each sec In ActiveDocument.Sections
        ' Only check footers
        For Each hdrftr In sec.Footers
            If hdrftr.Exists Then
                For Each para In hdrftr.range.paragraphs
                    If para.style = "Footer" Then
                        count = count + 1
                        para.range.Select
                        Debug.Print "Found paragraph with Footer style. Stopping at this location."
                        Stop
                    End If
                Next para
            End If
        Next hdrftr
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
            currentChar = Mid(content, i, 1)
            If Asc(currentChar) = 11 Then ' Chr(11) = manual line break
                totalBreaks = totalBreaks + 1
                If i > 1 Then
                    prevChar = Mid(content, i - 1, 1)
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
                debugInfo = debugInfo & "Fn " & footnote.Index & ": '" & rng.text & "' (ASC " & asciiValue & "), "
                notSpaceCount = notSpaceCount + 1
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
        debugInfo = debugInfo & "Footnote " & footnote.Index & ": '" & rng.text & "' (ASCII: " & asciiValue & ")" & vbCrLf
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
            DoEvents
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

