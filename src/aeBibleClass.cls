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

'Copyright (c) 2025-2026 Peter F. Ennis
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation;
'version 3.0, or (at your option) any later version.
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'Lesser General Public License for more details.
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, visit
'https://www.gnu.org/licenses/lgpl-3.0.txt

'=================================================================================
' Author:   Peter F. Ennis
' Date:     March 16, 2025
' Comment:  Create class for QA of Radiant Word Bible (RWB)
' History:  See comment details, basChangeLog_aeBibleClass, commit messages on GitHub
' GitHub:   https://github.com/adaept/BibleClass
'
' Classification: QA / audit / editorial support module.
'   Not part of runtime ribbon navigation flow.
'=================================================================================

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)

Private BibleClassVERSION As String
Private BibleClassVERSION_DATE As String

Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100
'
Private Const MyVerseTextFont As String = "Carlito"
'
' Define the word constant in this class module
Private Const wdNotThemeColor As Long = -1
' Default use is to not time all tests individually
Private Const bTimeAllTests As Boolean = True
' Used to store Expected results
Private oneBasedExpectedArray As Variant
Private SkippedTestNum As Long
Private Const MaxTests = 73
Private ResultArray(1 To MaxTests) As Long
Private GetPassFailArray(1 To MaxTests) As String
Private TestTimingArray(0 To MaxTests) As Variant
Private TestSessionID As String
Private Result As Long
Private TestReportFileName As String
Private testFileNum As Long
Private Const TestReportFlag As Boolean = True
Private OneTest As Long  ' Stores the value to run a single test
Private SkipTestArray As Variant       ' Used for heavy tests to skip
Private Const AposCP As Long = &H2019  ' Apostrophe Code Point
Private Apostrophe As String
Private ContractionArrayValues As Variant
Private contra As String               ' Contraction
Private Const CARTS As Long = 52    ' Contraction Array RunTest Start
Private UnicodeString As String        ' A string of up to 3 code points
Private UnicodeLabelU As String        ' The UnicodeString label in U+ form
Private OneVersePerPara As Boolean     ' Main branch has this set True and v59 or other branch set False
Private m_ReportBuf As String          ' Accumulates TestReport lines; flushed once at end of run
Private m_lastFuncError As Boolean     ' Set True in any Count function PROC_ERR; checked in GetPassFail
Private m_lastHint As String           ' Set by Count functions on first violation; cleared in GetPassFail
Private m_HintArray(1 To MaxTests) As String  ' Per-test first-hit hint; printed in RunTest on FAIL
'==================================================================

Private Sub Class_Initialize()
    On Error GoTo PROC_ERR
    'MsgBox Application.VBE.ActiveVBProject.Name
    'MsgBox Application.Caption
    If Application.VBE.ActiveVBProject.Name = "BibleClass" Then
        Application.Caption = Application.VBE.ActiveVBProject.Name & " " & BibleClassVERSION
    End If
    Debug.Print "Class_Initialize"
    BibleClassVERSION = "0.1.2"
    BibleClassVERSION_DATE = "January 31, 2026"
    ' Reset OneTest for each run
    OneTest = 0
    SkippedTestNum = 0
    Apostrophe = ChrW$(AposCP)
    If FileNameStartsWithV59 Then
        OneVersePerPara = False
    Else
        OneVersePerPara = True
    End If
    Debug.Print "OneVersePerPara = " & OneVersePerPara
PROC_EXIT:
    Exit Sub
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.Class_Initialize | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Sub

Private Sub Class_Terminate()
    On Error GoTo PROC_ERR
    Debug.Print
    Debug.Print "Class_Terminate"
    Debug.Print , Application.VBE.ActiveVBProject.Name
    Debug.Print , "BibleClass VERSION: " & BibleClassVERSION
    Debug.Print , "BibleClass VERSION_DATE: " & BibleClassVERSION_DATE
PROC_EXIT:
    Exit Sub
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.Class_Terminate | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Sub

Public Property Get TheBibleClassTests(Optional ByVal varDebug As Variant) As Long
    On Error GoTo PROC_ERR
    
    If Not CheckShowHideStatus Then
        ' Make sure all hidden characters are visible
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
        TheBibleClassTests = -1
    End If
    
PROC_EXIT:
    Exit Property
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TheBibleClassTests of Class BibleClass"
    m_lastFuncError = True
    Resume PROC_EXIT
End Property

Private Function CheckShowHideStatus() As Boolean
    CheckShowHideStatus = ActiveWindow.View.ShowAll
End Function

'===============================================================
' Returns True if the active document filename starts with "v59"
'===============================================================
Private Function FileNameStartsWithV59() As Boolean
    On Error GoTo PROC_ERR

    Dim fileName As String

    ' Get filename only (no path)
    fileName = ActiveDocument.Name
    FileNameStartsWithV59 = (LCase$(Left$(fileName, 3)) = "v59")

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.FileNameStartsWithV59 | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Sub InitializeGlobalResultArrayToMinusOne()
    On Error GoTo PROC_ERR

    Dim i As Long, j As Long, k As Long
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
    m_ReportBuf = ""
    m_lastHint = ""
    Dim h As Long
    For h = 1 To MaxTests: m_HintArray(h) = "": Next h
    ' Optional: Print to Immediate window to confirm
    'For i = 1 To MaxTests
    '    Debug.Print "ResultArray(" & i & ") = " & ResultArray(i),
    'Next i

PROC_EXIT:
    Exit Sub
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.InitializeGlobalResultArrayToMinusOne | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Sub

Private Function ConvertToOneBasedArray(values As Variant) As Variant
    Dim strArray() As String
    Dim i As Long

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
    Debug.Print "ERROR in aeBibleClass.ConvertToOneBasedArray | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Sub Expected1BasedArray()
    Dim doc As Document
    Dim values As Variant
    Dim i As Long
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
    
    ' Define the Expected RunTest Result values to store in the array
    '      RunTest 1  2  3  4  5  6  7  8  9  10 11 12 13 14 15 16     17 18  19 20 21 22   23 24    25 26 27   28  29  30 31 32   33   34   35   36 37  38   39 40 41 42 43 44 45 46 47 48 49  50   51    52 53 54 55 56 57 58 59 60 61 62 63 64 65 66 67 68 69 70 71 72 73)
    values = Array(0, 0, 0, 0, 0, 0, 0, 0, 7, 1, 0, 0, 0, 0, 0, 33783, 0, 66, 0, 0, 0, 224, 0, 1000, 1, 0, 148, 77, 70, 0, 1, 982, 982, 117, 117, 0, 19, 153, 0, 0, 0, 0, 0, 0, 0, 0, 3, 0, 15, 147, 1189, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0)

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
    If TestReportFlag And OneTest = 0 Then BufAppend outputString
    
    ' Next 15 results
    outputString = ""
    For i = 16 To 30
        outputString = outputString & " Test(" & i & ")" & oneBasedExpectedArray(i) & " "
        ' Trim the trailing space
        outputString = Trim(outputString)
    Next i
    Debug.Print outputString
    If TestReportFlag And OneTest = 0 Then BufAppend outputString
    
    ' Next 15 results
    outputString = ""
    For i = 31 To 45
        outputString = outputString & " Test(" & i & ")" & oneBasedExpectedArray(i) & " "
        ' Trim the trailing space
        outputString = Trim(outputString)
    Next i
    Debug.Print outputString
    If TestReportFlag And OneTest = 0 Then BufAppend outputString

    ' Next 15 results
    outputString = ""
    For i = 46 To 60
        outputString = outputString & " Test(" & i & ")" & oneBasedExpectedArray(i) & " "
        ' Trim the trailing space
        outputString = Trim(outputString)
    Next i
    Debug.Print outputString
    If TestReportFlag And OneTest = 0 Then BufAppend outputString

    ' Next results
    outputString = ""
    For i = 61 To UBound(oneBasedExpectedArray)
        outputString = outputString & " Test(" & i & ")" & oneBasedExpectedArray(i) & " "
        ' Trim the trailing space
        outputString = Trim(outputString)
    Next i
    Debug.Print outputString
    If TestReportFlag And OneTest = 0 Then BufAppend outputString

PROC_EXIT:
    Exit Sub
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.Expected1BasedArray | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Sub

'==================================================================
' Manually update with values of tests that are to be skipped
'==================================================================
Private Sub MakeSkipTestArray()
    On Error GoTo PROC_ERR

    SkipTestArray = Array(42, 51, 72)

PROC_EXIT:
    Exit Sub
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.MakeSkipTestArray | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Sub

Private Function IsSkipTest(ByVal n As Long) As Boolean
    On Error GoTo PROC_ERR

    Dim i As Long
    For i = LBound(SkipTestArray) To UBound(SkipTestArray)
        If SkipTestArray(i) = n Then
            IsSkipTest = True
            Exit Function
        End If
    Next i

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.IsSkipTest | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function HexToUnicodeLabel(cp As Long) As String
    On Error GoTo PROC_ERR

    HexToUnicodeLabel = "U+" & Right$("0000" & Hex$(cp), 4)

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.HexToUnicodeLabel | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function MakeUnicodeSeq(cp1 As Long, Optional cp2 As Long = 0, Optional cp3 As Long = 0) As String
    On Error GoTo PROC_ERR

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

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.MakeUnicodeSeq | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function ProcessUnicode(s As String) As String
    ProcessUnicode = s   ' return the string exactly as received
End Function

Private Function ContractionArrayU(idx As Long) As String
    On Error GoTo PROC_ERR

    ' Replace `'` with unicode value in the string, Apostrophe, =ChrW$(AposCP), &H2019, RIGHT SINGLE QUOTATION MARK

    'Debug.Print ">ContractionArrayValues(idx) = " & ContractionArrayValues(idx)
    ContractionArrayU = Replace(ContractionArrayValues(idx), "'", Apostrophe, 1, 1)
    'Debug.Print ">ContractionArrayU) = " & ContractionArrayU

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.ContractionArrayU | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Sub CreateContractionArray()
    On Error GoTo PROC_ERR

    ContractionArrayValues = ConvertToOneBasedArray(Array( _
        "wouldn't", "isn't", "let's", "i'm", "hasn't", "didn't", "don't", "haven't", "weren't", "aren't", "can't", "shouldn't", "it's", _
        "spirit's" _
        ))
    ' Now: LBound(values) = 1, UBound(values) = 14, values(1) = "wouldn't", values(14) = "spirit's"

PROC_EXIT:
    Exit Sub
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CreateContractionArray | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Sub

'=====================================================================================
' BufAppend / FlushReportBuf - Batched TestReport writer
'
' BufAppend accumulates lines in m_ReportBuf during the test run.
' FlushReportBuf opens TestReportFileName once at the end and writes the full
' buffer in a single Append operation, replacing the per-line open/close pattern.
'=====================================================================================
Private Sub BufAppend(text As String)
    m_ReportBuf = m_ReportBuf & text & vbCrLf
End Sub

Private Sub FlushReportBuf()
    If m_ReportBuf = "" Then Exit Sub
    Dim fileNum As Long
    Dim folderPath As String
    Dim finalPath As String
    folderPath = ActiveDocument.Path & "\rpt"
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
    If InStr(TestReportFileName, "\") = 0 Then
        finalPath = folderPath & "\" & TestReportFileName
    Else
        finalPath = TestReportFileName  ' already a full path
    End If
    fileNum = FreeFile
    Open finalPath For Append As fileNum
    Print #fileNum, m_ReportBuf;
    Close fileNum
    m_ReportBuf = ""
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
Private Sub AppendToFile(filePath As String, text As String)
    Dim fileNum As Long
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
    On Error GoTo PROC_ERR

    LogMessage = "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "] " & msg

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.LogMessage | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
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
    On Error GoTo PROC_ERR

    ' Format: YYYYMMDD-HHMMSS
    GenerateSessionID = Format(Now, "yyyymmdd-hhnnss")

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.GenerateSessionID | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

'============================================================
' Routine: RunTotalTimeTestSession
' Purpose: Runs a timing test and logs outcome with session ID.
' Inputs:
'   - testOutcome: String summary of timing Result.
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
    On Error GoTo PROC_ERR

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
    
    Dim fileNum As Long
    fileNum = FreeFile
    
    Open folderPath & "\" & reportName For Append As #fileNum
    If Not FileExists Then
        Print #fileNum, "SessionID, TotalTime"
    End If
    Close #fileNum

    Call AppendToFile("TotalTimeReport.txt", sessionID & ", " & testOutcome)

    'MsgBox "Test complete. Session ID: " & sessionID, vbInformation
    Debug.Print "RunTotalTimeTestSession complete. Session ID: " & sessionID

PROC_EXIT:
    Exit Sub
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.RunTotalTimeTestSession | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Sub

Private Sub DebugAndReportHeader()
    On Error GoTo PROC_ERR

    Debug.Print "1234567890123412345678901234123456789012341234567890123412345678901234"
    '            123456789    1234567    12345678    123456    12345678    12345678     Note: Comma separation of debug is 14 characters
    Debug.Print "Pass/Fail", "Copy ()", "Test Num", "Result", "Expected", "Function"

PROC_EXIT:
    Exit Sub
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.DebugAndReportHeader | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
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
            If TestReportFlag And OneTest = 0 Then BufAppend "1234567890123412345678901234123456789012341234567890123412345678901234"
            If TestReportFlag And OneTest = 0 Then BufAppend "Pass/Fail" & Space(5) & "Copy ()" & Space(7) & "Test Num" & Space(6) & "Result" & Space(8) & "Expected" & Space(6) & "Function"
            
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
            RunTest (69)
            RunTest (70)
            RunTest (71)
            RunTest (72)
            RunTest (73)
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
    
    If TestReportFlag And OneTest = 0 Then BufAppend "The Total Runtime: " & TheTotalRuntime & " seconds"
    
    Debug.Print LogMessage("Completed")
    If TestReportFlag And OneTest = 0 Then BufAppend LogMessage("Completed")

    Debug.Print "BibleClass VERSION: " & BibleClassVERSION
    If TestReportFlag And OneTest = 0 Then BufAppend "BibleClass VERSION: " & BibleClassVERSION
    
    Debug.Print "BibleClass VERSION_DATE: " & BibleClassVERSION_DATE
    If TestReportFlag And OneTest = 0 Then BufAppend "BibleClass VERSION_DATE: " & BibleClassVERSION_DATE

    Debug.Print LogWordBuildInfo
    If TestReportFlag And OneTest = 0 Then BufAppend LogWordBuildInfo
    If TestReportFlag And OneTest = 0 Then FlushReportBuf

    RunBibleClassTests = True

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.RunBibleClassTests | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function LogWordBuildInfo() As String
    On Error GoTo PROC_ERR

    Dim info As String
    info = "== Word Build Info ==" & vbCrLf
    info = info & "Application Name: " & Application.Name & vbCrLf
    info = info & "Platform Version: " & Application.Version & vbCrLf
    info = info & "Platform Build: " & Application.Build & vbCrLf
    info = info & "Full Version String: " & Application.Name & " " & Application.Version & " (" & Application.Build & ")" & vbCrLf
    info = info & "Marketing Version: 2506 (Build 18925.20158)" & vbCrLf
    info = info & "Source: Microsoft 365 Current Channel, July 2025 update"
    'Debug.Print info
    'MsgBox info, vbInformation, "Word Version Details"
    LogWordBuildInfo = info

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.LogWordBuildInfo | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function U(codepoint As Long) As String
    On Error GoTo PROC_ERR

    U = ChrW(codepoint)

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.U | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function GetPassFail(TestNum As Long) As Boolean
    On Error GoTo PROC_ERR
    
    ' Set explicit skip if the test is skipped
    If IsSkipTest(TestNum) Then
        GetPassFailArray(TestNum) = "SKIP!!!!"
        Exit Function
    End If

    m_lastFuncError = False     ' reset before calling any Count function
    m_lastHint = ""             ' reset hint before calling any Count function

    Select Case TestNum
    Case 1
        ResultArray(TestNum) = CountDoubleSpaces
        m_HintArray(TestNum) = m_lastHint
    Case 2
        ResultArray(TestNum) = CountDoubleSpacesInShapes
        m_HintArray(TestNum) = m_lastHint
    Case 3
        ResultArray(TestNum) = CountSpaceFollowedByCarriageReturn
        m_HintArray(TestNum) = m_lastHint
    Case 4
        ResultArray(TestNum) = CountDoubleTabs
        m_HintArray(TestNum) = m_lastHint
    Case 5
        ResultArray(TestNum) = CountWhiteSpaceAndCarriageReturn
        m_HintArray(TestNum) = m_lastHint
    Case 6
        ResultArray(TestNum) = CountQuadrupleParagraphMarks
        m_HintArray(TestNum) = m_lastHint
    Case 7
        ResultArray(TestNum) = CountStyleWithSpaceAndNumber
        m_HintArray(TestNum) = m_lastHint
    Case 8
        ResultArray(TestNum) = CountStyleWithNumberAndSpace
        m_HintArray(TestNum) = m_lastHint
    Case 9
        ResultArray(TestNum) = CountPeriodSpaceLeftParenthesis
        m_HintArray(TestNum) = m_lastHint
    Case 10
        ResultArray(TestNum) = CountNonBreakingSpaces
        m_HintArray(TestNum) = m_lastHint
    Case 11
        ResultArray(TestNum) = CountFindNumberDashNumber
        m_HintArray(TestNum) = m_lastHint
    Case 12
        ResultArray(TestNum) = CountNumberDashNumberInFootnotes
        m_HintArray(TestNum) = m_lastHint
    Case 13
        ResultArray(TestNum) = CountEmptyParasWithNoThemeColor
        m_HintArray(TestNum) = m_lastHint
    Case 14
        ResultArray(TestNum) = CountWhiteParagraphMarks
        m_HintArray(TestNum) = m_lastHint
    Case 15
        ResultArray(TestNum) = CountSectionsWithDifferentFirstPage
        m_HintArray(TestNum) = m_lastHint
    Case 16
        ResultArray(TestNum) = CountTotalParagraphs
        m_HintArray(TestNum) = m_lastHint
    Case 17
        ResultArray(TestNum) = CountActiveHyperlinks
        m_HintArray(TestNum) = m_lastHint
    Case 18
        ResultArray(TestNum) = CountHeading1
        m_HintArray(TestNum) = m_lastHint
    Case 19
        ResultArray(TestNum) = CountEmptyParasAfterH2
        m_HintArray(TestNum) = m_lastHint
    Case 20
        ResultArray(TestNum) = CountFootnotesFollowedByDigit
        m_HintArray(TestNum) = m_lastHint
    Case 21
        ResultArray(TestNum) = CountNotSpacesAfterFootnoteReferences
        m_HintArray(TestNum) = m_lastHint
    Case 22
        ResultArray(TestNum) = CountEmptyParagraphsWithFormatting
        m_HintArray(TestNum) = m_lastHint
    Case 23
        ResultArray(TestNum) = CountDeleteEmptyParagraphsBeforeHeading2
        m_HintArray(TestNum) = m_lastHint
    Case 24
        ResultArray(TestNum) = CountFootnoteReferences
        m_HintArray(TestNum) = m_lastHint
    Case 25
        ResultArray(TestNum) = CountFootnoteReferenceColors
        m_HintArray(TestNum) = m_lastHint
    Case 26
        ResultArray(TestNum) = CheckAllHeaders("Empty")
        m_HintArray(TestNum) = m_lastHint
    Case 27
        ResultArray(TestNum) = CheckAllHeaders("NotEmpty")
        m_HintArray(TestNum) = m_lastHint
    Case 28
        ResultArray(TestNum) = CountTabFollowedByParagraphMarkInHeaders
        m_HintArray(TestNum) = m_lastHint
    Case 29
        ResultArray(TestNum) = CountParagraphsWithoutTabInHeaders
        m_HintArray(TestNum) = m_lastHint
    Case 30
        ResultArray(TestNum) = CountHeaderStyleUsage
        m_HintArray(TestNum) = m_lastHint
    Case 31
        ResultArray(TestNum) = CountParagraphMarksPerHeaderSection
        m_HintArray(TestNum) = m_lastHint
    Case 32
        ResultArray(TestNum) = CountLinefeed
        m_HintArray(TestNum) = m_lastHint
    Case 33
        ResultArray(TestNum) = CountLinefeed(" ")
        m_HintArray(TestNum) = m_lastHint
    Case 34
        ResultArray(TestNum) = CountManualLineBreaksAndWithSpace
        m_HintArray(TestNum) = m_lastHint
    Case 35
        ResultArray(TestNum) = CountManualLineBreaksAndWithSpace(" ")
        m_HintArray(TestNum) = m_lastHint
    Case 36
        ResultArray(TestNum) = CountFooterParagraphsWithFooterStyle
        m_HintArray(TestNum) = m_lastHint
    Case 37
        ResultArray(TestNum) = CountDocTabOnlyParagraphs
        m_HintArray(TestNum) = m_lastHint
    Case 38
        ResultArray(TestNum) = CountEmptyParagraphs
        m_HintArray(TestNum) = m_lastHint
    Case 39
        ResultArray(TestNum) = CountParagraphMarks_ArialBlack
        m_HintArray(TestNum) = m_lastHint
    Case 40
        ResultArray(TestNum) = CountParagraphMarks_ArialBlackDarkRed
        m_HintArray(TestNum) = m_lastHint
    Case 41
        ResultArray(TestNum) = Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed
        m_HintArray(TestNum) = m_lastHint
    Case 42
        ResultArray(TestNum) = CountBoldFootnotesWordLevel
        m_HintArray(TestNum) = m_lastHint
    Case 43
        ResultArray(TestNum) = CountParagraphMarksWithDarkRedFormatting
        m_HintArray(TestNum) = m_lastHint
    Case 44
        ResultArray(TestNum) = AuditNonPaletteStyleColors
        m_HintArray(TestNum) = m_lastHint
    Case 45
        ResultArray(TestNum) = CountUnapprovedVisibleStyles
        m_HintArray(TestNum) = m_lastHint
    Case 46
        ResultArray(TestNum) = AuditBookHyperlinkStyling
        m_HintArray(TestNum) = m_lastHint
    Case 47
        ResultArray(TestNum) = CountTabOnlyParagraphs("Footer")
        m_HintArray(TestNum) = m_lastHint
    Case 48
        ResultArray(TestNum) = AuditLiberationSansNarrowStyleDetails
        m_HintArray(TestNum) = m_lastHint
    Case 49
        ResultArray(TestNum) = CountAuditStyles_ToFile
        m_HintArray(TestNum) = m_lastHint
    Case 50
        ResultArray(TestNum) = SummarizeHeaderFooterAuditToFile
        m_HintArray(TestNum) = m_lastHint
    Case 51
        ResultArray(TestNum) = CountAndCreateDefinitionForH2
        m_HintArray(TestNum) = m_lastHint
    Case 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65
        contra = ContractionArrayU(TestNum - CARTS + 1)
        ResultArray(TestNum) = CountContraction(contra)
        m_HintArray(TestNum) = m_lastHint
    Case 66
        UnicodeString = ProcessUnicode(MakeUnicodeSeq(&H20, &H2019))            ' Space U+0020 followed by RSQM U+2019
        ResultArray(TestNum) = CountContraction(UnicodeString)
        m_HintArray(TestNum) = m_lastHint
    Case 67
        UnicodeString = ProcessUnicode(MakeUnicodeSeq(&H20, &H201D))            ' Space U+0020 followed by RDQM U+201D
        ResultArray(TestNum) = CountContraction(UnicodeString)
        m_HintArray(TestNum) = m_lastHint
    Case 68
        UnicodeString = ProcessUnicode(MakeUnicodeSeq(&H202F, &H2019))          ' NNBSP U+202F followed by RSQM U+2019
        ResultArray(TestNum) = CountContraction(UnicodeString)
        m_HintArray(TestNum) = m_lastHint
    Case 69
        UnicodeString = ProcessUnicode(MakeUnicodeSeq(&H202F, &H201D))          ' NNBSP U+202F followed by RDQM U+201D
        ResultArray(TestNum) = CountContraction(UnicodeString)
        m_HintArray(TestNum) = m_lastHint
    Case 70
        UnicodeString = ProcessUnicode(MakeUnicodeSeq(&H201C, &H2018, &H201C))  ' Double Opening Quote, Single Opening Quote, Double Opening Quote
        ResultArray(TestNum) = CountContraction(UnicodeString)
        m_HintArray(TestNum) = m_lastHint
    Case 71
        UnicodeString = ProcessUnicode(MakeUnicodeSeq(&H201D, &H2019, &H201D))  ' Double Closing Quote, Single Closing Quote, Double Closing Quote
        ResultArray(TestNum) = CountContraction(UnicodeString)
        m_HintArray(TestNum) = m_lastHint
    Case 72
        If OneVersePerPara Then
            ResultArray(TestNum) = HasLeftAlignedParagraph(18, 931)
        m_HintArray(TestNum) = m_lastHint
        End If
    Case 73
        ResultArray(TestNum) = CountInvisibleCharacters()
        m_HintArray(TestNum) = m_lastHint
    Case Else
        Debug.Print "TestNum = " & TestNum & " >The test number is outside the accepted range"
    End Select

    ' If any Count function errored, override Result with sentinel so test shows FAIL
    If m_lastFuncError Then
        ResultArray(TestNum) = -1
        m_HintArray(TestNum) = m_lastHint
        Debug.Print "!! Test " & TestNum & " Result overridden to -1 (function error)"
    End If

    If ResultArray(TestNum) = oneBasedExpectedArray(TestNum) Then
        GetPassFailArray(TestNum) = "PASS    "
    Else
        GetPassFailArray(TestNum) = "FAIL!!!!"
    End If

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.GetPassFail | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
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
    On Error GoTo PROC_ERR

    Const fixedWidth = 14
    Fixed14Str = Left(MyResult & String(fixedWidth, " "), fixedWidth)
    'Debug.Print "The fixed width string is:" & vbCrLf & "[" & Fixed14Str & "]"
    'Debug.Print "[12345678901234]"

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.Fixed14Str | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function RunTest(num As Long, Optional SkipTest As Variant) As Boolean
    Dim startTime As Double
    Dim endTime As Double
    Dim runTime As Double
    
    ' Record the start timer for each test
    startTime = Timer

    On Error GoTo PROC_ERR

    Debug.Print ">> Starting Test " & num
    DoEvents
    GetPassFail (num)

    Select Case num
    Case 1
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountDoubleSpaces"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 2
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountDoubleSpacesInShapes"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 3
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountSpaceFollowedByCarriageReturn"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 4
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountDoubleTabs"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 5
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountWhiteSpaceAndCarriageReturn"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 6
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountQuadrupleParagraphMarks"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 7
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountStyleWithSpaceAndNumber"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 8
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountStyleWithNumberAndSpace"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 9
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountPeriodSpaceLeftParenthesis"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 10
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountNonBreakingSpaces"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 11
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountFindNumberDashNumber"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 12
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountNumberDashNumberInFootnotes"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 13
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountEmptyParasWithNoThemeColor"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 14
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountWhiteParagraphMarks"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 15
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountSectionsWithDifferentFirstPage"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 16
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountTotalParagraphs"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 17
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountActiveHyperlinks"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 18
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountHeading1"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 19
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountEmptyParasAfterH2"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 20
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountFootnotesFollowedByDigit"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 21
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountNotSpacesAfterFootnoteReferences"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 22
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountEmptyParagraphsWithFormatting"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 23
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountDeleteEmptyParagraphsBeforeHeading2"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 24
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountFootnoteReferences"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 25
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountFootnoteReferenceColors"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 26
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CheckAllHeaders(" & "Empty" & ")"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 27
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CheckAllHeaders(" & "NotEmpty" & ")"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 28
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountTabFollowedByParagraphMarkInHeaders"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 29
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountParagraphsWithoutTabInHeaders"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 30
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountHeaderStyleUsage"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 31
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountParagraphMarksPerHeaderSection"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 32
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountLinefeed"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 33
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountLinefeed(" & """ """ & ")"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 34
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountManualLineBreaksAndWithSpace"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 35
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountManualLineBreaksAndWithSpace(" & """ """ & ")"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 36
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountFooterParagraphsWithFooterStyle  FIX: ReapplyTheFootersToAllFooters"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 37
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountDocTabOnlyParagraphs"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 38
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountEmptyParagraphs"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 39
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountParagraphMarks_ArialBlack"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 40
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountParagraphMarks_ArialBlackDarkRed"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 41
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 42
        'Debug.Print ">> Expected Routine Runtime ~ 80 seconds"
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountBoldFootnotesWordLevel"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 43
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountParagraphMarksWithDarkRedFormatting"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 44
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "AuditNonPaletteStyleColors"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 45
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountUnapprovedVisibleStyles"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 46
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "AuditBookHyperlinkStyling"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 47
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountTabOnlyParagraphs(" & """Footer""" & ")"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 48
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "AuditLiberationSansNarrowStyleDetails"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 49
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountAuditStyles_ToFile"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 50
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "SummarizeHeaderFooterAuditToFile"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 51
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountAndCreateDefinitionForH2"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "U+2019 CountContraction(" & """" & ContractionArrayU(num - CARTS + 1) & """" & ")"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 66, 67, 68, 69, 70, 71
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), UnicodeLabelU & " CountContraction(" & """" & UnicodeString & """" & "])"    ' Using [] to delineate unicode string in test output
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
    Case 72
        If OneVersePerPara Then
            Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "HasLeftAlignedParagraph(18, 931)"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
        End If
    Case 73
        Debug.Print GetPassFailArray(num), "Copy ()", "Test = " & num, ResultArray(num), oneBasedExpectedArray(num), "CountInvisibleCharacters"
        If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then Debug.Print , , , "  >> First hit: " & m_HintArray(num)
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
    Dim answer As Long
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
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function OutputTestReport(num As Long, Optional SkipTest As Variant) As Boolean
    On Error GoTo PROC_ERR
    
    Select Case num
    Case 1
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountDoubleSpaces"
    Case 2
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountDoubleSpacesInShapes"
    Case 3
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountSpaceFollowedByCarriageReturn"
    Case 4
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountDoubleTabs"
    Case 5
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountWhiteSpaceAndCarriageReturn"
    Case 6
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountQuadrupleParagraphMarks"
    Case 7
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountStyleWithSpaceAndNumber"
    Case 8
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountStyleWithNumberAndSpace"
    Case 9
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountPeriodSpaceLeftParenthesis"
    Case 10
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountNonBreakingSpaces"
    Case 11
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountFindNumberDashNumber"
    Case 12
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountNumberDashNumberInFootnotes"
    Case 13
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountEmptyParasWithNoThemeColor"
    Case 14
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountWhiteParagraphMarks"
    Case 15
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountSectionsWithDifferentFirstPage"
    Case 16
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountTotalParagraphs"
    Case 17
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountActiveHyperlinks"
    Case 18
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountHeading1"
    Case 19
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountEmptyParasAfterH2"
    Case 20
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountFootnotesFollowedByDigit"
    Case 21
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountNotSpacesAfterFootnoteReferences"
    Case 22
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountEmptyParagraphsWithFormatting"
    Case 23
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountDeleteEmptyParagraphsBeforeHeading2"
    Case 24
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountFootnoteReferences"
    Case 25
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountFootnoteReferenceColors"
    Case 26
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CheckAllHeaders(" & "Empty" & ")"
    Case 27
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CheckAllHeaders(" & "NotEmpty" & ")"
    Case 28
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountTabFollowedByParagraphMarkInHeaders"
    Case 29
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountParagraphsWithoutTabInHeaders"
    Case 30
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountHeaderStyleUsage"
    Case 31
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountParagraphMarksPerHeaderSection"
    Case 32
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountLinefeed"
    Case 33
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountLinefeed(" & """ """ & ")"
    Case 34
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountManualLineBreaksAndWithSpace"
    Case 35
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountManualLineBreaksAndWithSpace(" & """ """ & ")"
    Case 36
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountFooterParagraphsWithFooterStyle  FIX: ReapplyTheFootersToAllFooters"
    Case 37
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountDocTabOnlyParagraphs"
    Case 38
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountEmptyParagraphs"
    Case 39
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountParagraphMarks_ArialBlack"
    Case 40
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountParagraphMarks_ArialBlackDarkRed"
    Case 41
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed"
    Case 42
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountBoldFootnotesWordLevel"
    Case 43
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountParagraphMarksWithDarkRedFormatting"
    Case 44
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "AuditNonPaletteStyleColors"
    Case 45
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountUnapprovedVisibleStyles"
    Case 46
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "AuditBookHyperlinkStyling"
    Case 47
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountTabOnlyParagraphs(" & """Footer""" & ")"
    Case 48
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "AuditLiberationSansNarrowStyleDetails"
    Case 49
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountAuditStyles_ToFile"
    Case 50
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "SummarizeHeaderFooterAuditToFile"
    Case 51
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountAndCreateDefinitionForH2"
    Case 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "U+2019 CountContraction(" & """" & ContractionArrayU(num - CARTS + 1) & """" & ")"
    Case 66, 67, 68, 69, 70, 71
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & UnicodeLabelU & " CountContraction([" & """" & UnicodeString & """" & "])"    ' Using [] to delineate unicode string in test output
    Case 72
        If OneVersePerPara Then
            BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "HasLeftAlignedParagraph(18, 931)"
        End If
    Case 73
        BufAppend GetPassFailArray(num) & Space(10) & "Copy ()" & Space(7) & Fixed14Str("Test = " & num) & Fixed14Str(CStr(ResultArray(num))) & Fixed14Str(CStr(oneBasedExpectedArray(num))) & "CountInvisibleCharacters"
    Case Else
        Debug.Print "num = " & num & " >>>The test number is outside the accepted range"
    End Select

    If GetPassFailArray(num) = "FAIL!!!!" And m_HintArray(num) <> "" Then _
        BufAppend "  >> First hit: " & m_HintArray(num)

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.OutputTestReport | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

'====================================================================
' Function: CountInvisibleCharacters
'
' Purpose:
'   Counts invisible Unicode characters across all document stories.
'   These characters are visually silent but can corrupt search results,
'   style normalization, and USFM export if present.
'
' Characters checked:
'   U+200B  ZERO WIDTH SPACE
'   U+200C  ZERO WIDTH NON-JOINER
'   U+200D  ZERO WIDTH JOINER
'   U+FEFF  ZERO WIDTH NO-BREAK SPACE (byte-order mark)
'   U+2060  WORD JOINER
'
' Returns:
'   Long — total Count across all story ranges.
'   0 = clean (expected pass); > 0 = violations found (fail).
'
' Test: RUN_THE_TESTS(73)
'====================================================================
Private Function CountInvisibleCharacters() As Long
    On Error GoTo PROC_ERR

    Dim r As Word.Range
    Dim targets As Variant
    Dim counts() As Long
    Dim i As Long
    Dim total As Long

    targets = Array(ChrW(&H200B), ChrW(&H200C), ChrW(&H200D), ChrW(&HFEFF), ChrW(&H2060))
    ReDim counts(UBound(targets))

    For Each r In ActiveDocument.StoryRanges
        For i = 0 To UBound(targets)
            counts(i) = counts(i) + UBound(Split(r.Text, targets(i)))
        Next i
    Next r

    For i = 0 To UBound(counts)
        total = total + counts(i)
    Next i

    CountInvisibleCharacters = total

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountInvisibleCharacters | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

'====================================================================
' Function: HasLeftAlignedParagraph
'
' Purpose:
'   Determines whether at least one LEFT-ALIGNED paragraph exists
'   between two PRINTED page numbers (inclusive) in the main document
'   body, using Word's native Find engine for speed and accuracy.
'
' Parameters:
'   startPage (Long) - Printed page number where the search begins.
'   endPage   (Long) - Printed page number where the search ends.
'
' Returns:
'   Long
'     1  > A left-aligned paragraph exists within the page range.
'     0  > No left-aligned paragraph exists within the page range.
'
' Behavior:
'   - Uses Selection.Find (same engine as Ctrl+H).
'   - Searches forward only and DOES NOT wrap.
'   - Page comparison is based on ADJUSTED page numbering, so it
'     correctly handles Roman numerals, restarts, and mixed formats.
'
' Performance:
'   Optimized for very large documents.
'   Typically executes in < 1 second (Ctrl+H equivalent speed).
'
' Limitations / Notes:
'   - Searches only the Main Text Story (no headers, footers, notes).
'   - Detects actual paragraph alignment, not style definitions.
'   - Stops immediately after the first match is found.
'
'====================================================================
Private Function HasLeftAlignedParagraph( _
    ByVal startPage As Long, _
    ByVal endPage As Long) As Long

    On Error GoTo PROC_ERR

    Dim sel As Selection
    Dim foundPage As Long
    Dim found As Boolean

    Set sel = Application.Selection

    ' 1. Go to START page by PRINTED page number
    GoToAdjustedPage startPage

    ' 2. Setup paragraph-format-only Find (Ctrl+H equivalent)
    With sel.Find
        .ClearFormatting
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
    End With

    ' 3. Execute Find
    found = sel.Find.Execute

    If Not found Then
        HasLeftAlignedParagraph = 0
        Exit Function
    End If

    ' 4. Read PRINTED page number
    foundPage = sel.Information(wdActiveEndAdjustedPageNumber)
    'Debug.Print "foundPage = " & foundPage

    ' 5. Compare against end page (printed)
    If foundPage <= endPage Then
        HasLeftAlignedParagraph = 1
    Else
        HasLeftAlignedParagraph = 0
    End If

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.HasLeftAlignedParagraph | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Sub GoToAdjustedPage(ByVal targetPage As Long)
    On Error GoTo PROC_ERR

    Dim sel As Selection
    Set sel = Application.Selection

    sel.HomeKey wdStory
    Do
        If sel.Information(wdActiveEndAdjustedPageNumber) = targetPage Then Exit Do
        sel.GoTo What:=wdGoToPage, Which:=wdGoToNext
    Loop

PROC_EXIT:
    Exit Sub
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.GoToAdjustedPage | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Sub

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
'   - 20260215 Updated to use CountInStory (doc, footnotes, etc.)
'
' Additional:
'   - Can also accept up to 3 code points as a string when not calling
'     the contraction array.
'
' Parameters:
'   contraction  The contraction to search for, using Unicode U+2019 (ChrW(&H2019))
'                Tests 52~65
'   or           A Unicode string of up to 3 code points
'                Tests 66~
'
' Returns:
'   The number of matches found.
'==============================================================================
Private Function CountContraction(ByVal contraction As String) As Long
    On Error GoTo PROC_ERR

    Dim s As Word.Range
    Dim Count As Long

    Count = 0

    For Each s In ActiveDocument.StoryRanges
        Count = Count + CountInStory(s, contraction)
    Next

    CountContraction = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountContraction | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountInStory(rng As Word.Range, contraction As String) As Long
    On Error GoTo PROC_ERR

    Dim Count As Long
    Count = 0

    With rng.Find
        .ClearFormatting
        .Text = contraction
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Wrap = wdFindStop
    End With

    Do While rng.Find.Execute
        Count = Count + 1
        If m_lastHint = "" Then m_lastHint = Left(rng.Paragraphs(1).Range.Text, 60)
        rng.Collapse wdCollapseEnd
    Loop

    CountInStory = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountInStory | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

' ======================================================================
' Macro Name   : CountAndCreateDefinitionForH2
' Purpose      : Enforces layout and paragraph rules for Heading 2 style.
'                - Enables KeepWithNext on Heading 2 paragraph style.
'                - Applies WidowControl to each Heading 2 paragraph.
'                - Explicitly disables KeepTogether to prevent override.
'                - Provides a total Count of H2
' Audit Notes  : Logs all repair actions to Immediate Window.
'                Does NOT alter any content, punctuation, or quotes.
'                Applies paragraph-level enforcement only where style = "Heading 2"
' Safety Level : Editorial-safe. No deletions or format coercion.
' Last Updated : 20250805
' Author       : Peter
' ======================================================================
Private Function CountAndCreateDefinitionForH2()
    On Error GoTo PROC_ERR

    ' Update Heading 2 Keep With Next
    Dim Count As Long
    Count = 0

    Application.ScreenUpdating = False

    Dim s As Word.Style
    Set s = ActiveDocument.Styles("Heading 2")
    
    ' Apply KeepWithNext to paragraph formatting
    s.ParagraphFormat.KeepWithNext = True
    'Debug.Print "Heading 2 style updated: KeepWithNext = True"

    ' Enforce Heading 2 Paragraph Widow Orphan
    ' Disable KeepLines Together For Heading 2
    Dim para As Word.Paragraph
    For Each para In ActiveDocument.Paragraphs
        If para.style = ActiveDocument.Styles("Heading 2") Then
            Count = Count + 1
            With para
                .WidowControl = True    ' enforces both widow and orphan control for that paragraph
                '.OrphanControl = True - Not needed,
                para.KeepTogether = False
            End With
            
            If Count Mod 50 = 0 Then
                DoEvents  ' Keep UI responsive every 50 updates
            End If
        End If
    Next para
    
    Application.ScreenUpdating = True
    'Debug.Print "[repair] Widow/Orphan enforced at paragraph level for Heading 2"
    'Debug.Print "[repair] KeepLinesTogether disabled for Heading 2"
    'Debug.Print "Count = " & Count
    CountAndCreateDefinitionForH2 = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountAndCreateDefinitionForH2 | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
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
'   region, tracking anomalies such As Word.Style drift, legacy font
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
    On Error GoTo PROC_ERR

    Dim sec As Word.Section, hf As HeaderFooter, p As Word.Paragraph
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
    totalSections = ActiveDocument.Sections.Count

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
                For Each p In hf.Range.Paragraphs
                    paraText = p.Range.Text
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
                For Each p In hf.Range.Paragraphs
                    paraText = p.Range.Text
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

        allGoodStyles = (tempDict.Count = 2) _
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
    Set ts = Nothing
    Set fs = Nothing
    Set tempDict = Nothing
    Set dictAsciiCount = Nothing
    Set dictAsciiDesc = Nothing
    Set dictAllStyles = Nothing
    Set dictBadHeaderStyles = Nothing
    Set dictBadFooterStyles = Nothing

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.SummarizeHeaderFooterAuditToFile | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
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
' - Outputs results as "Style: Count" pairs, plus total paragraph Count.
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
    On Error GoTo PROC_ERR

    Dim rng As Word.Range
    Dim para As Word.Paragraph
    Dim paraStyle As String
    Dim paraCount As Long
    Dim keyCount As Long
    Dim styleDict As Object
    Set styleDict = CreateObject("Scripting.Dictionary")

    ' Build the distribution
    For Each rng In ActiveDocument.StoryRanges
        Do
            For Each para In rng.Paragraphs
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
    Set outFile = Nothing
    Set fso = Nothing

    'Debug.Print "Style usage written to:" & vbCrLf & outputPath
    'MsgBox "Style usage written to:" & vbCrLf & outputPath, vbInformation, "Audit Complete"
    'Debug.Print "keyCount = " & keyCount
    If m_lastHint = "" Then m_lastHint = keyCount & " styles - see rpt\Style Usage Distribution.txt"
    CountAuditStyles_ToFile = keyCount

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountAuditStyles_ToFile | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function AuditLiberationSansNarrowStyleDetails() As Long
    On Error GoTo PROC_ERR

    Dim s As Word.Style
    Dim logBuffer As String
    Dim fName As String, fSize As Variant, pStyleType As String
    Dim styleCount As Long

    logBuffer = "=== Liberation Sans Narrow Style Definitions ===" & vbCrLf

    For Each s In ActiveDocument.Styles
        On Error Resume Next
        fName = s.Font.Name
        fSize = s.Font.Size
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
                "- Style: �" & s.NameLocal & "� | Type: " & pStyleType & _
                " | Font Size: " & fSize & vbCrLf
        End If
        On Error GoTo 0
    Next s

    'logBuffer = logBuffer & vbCrLf & "Total styles using Liberation Sans Narrow: " & styleCount
    'Debug.Print logBuffer
    'MsgBox "Style details printed. See Immediate Window.", vbInformation
    AuditLiberationSansNarrowStyleDetails = styleCount

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.AuditLiberationSansNarrowStyleDetails | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountTabOnlyParagraphs(areaType As String) As Long
    On Error GoTo PROC_ERR

    Dim sec As Word.Section
    Dim hdrFtr As HeaderFooter
    Dim para As Word.Paragraph
    Dim Count As Long
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
            For Each para In hdrFtr.Range.Paragraphs

                ' Move the cursor to this paragraph in the footer
                para.Range.Select

                paraText = Replace(para.Range.Text, vbCr, "")
                paraText = Replace(paraText, vbLf, "")
                paraText = Trim(paraText)

                If paraText = vbTab Then
                    Count = Count + 1
                    If m_lastHint = "" Then _
                        m_lastHint = areaType & " Section " & sec.index
                End If
            Next para
        End If
    Next sec

    CountTabOnlyParagraphs = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountTabOnlyParagraphs | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

'=========================================================================
' AuditNonPaletteStyleColors  (Public; wired at slot 44)
' Two-tier colour discipline (see EDSG/01-styles.md):
'   Tier 1 - default-text intent : Font.Color = wdColorAutomatic
'   Tier 2 - deliberate colour   : Font.Color = a palette-registered value
'
' Five-bucket classification:
'   1. Tier 1   - Font.Color = wdColorAutomatic
'   2. Tier 2   - Font.Color in palette (NameFromColor returns non-empty)
'   3. Theme    - Font.ObjectThemeColor <> wdThemeColorNone. Office theme
'                 colour reference. Reported separately - banned by the
'                 rule but not the same as a hand-typed off-palette
'                 anomaly.
'   4. Anomaly  - none of the above. The actual editorial concern: a
'                 hand-typed off-palette RGB value. Return-value
'                 contribution.
'   5. Skipped  - Table or List style (Font.Color not meaningful),
'                 Font.Color read errored, or built-in style when
'                 IncludeBuiltIn is False.
'
' Return value = Anomaly Count only (the editorial-discipline assertion).
'
' Optional IncludeBuiltIn (default False): skip Word built-in styles.
' Pass True for a one-off diagnostic that covers everything.
'
' Public so the Immediate-window stub in basStyleInspector can delegate
' in. Canonical body lives here per the class-encapsulation rule
' (EDSG/12-module-vs-class.md).
'=========================================================================
Public Function AuditNonPaletteStyleColors( _
        Optional ByVal IncludeBuiltIn As Boolean = False) As Long
    Const WD_THEME_NONE As Long = -1   ' wdThemeColorNone
    Dim oDoc        As Word.Document
    Dim oStyle      As Word.Style
    Dim c           As Long
    Dim themeIdx    As Long
    Dim paletteName As String
    Dim styleType   As Long
    Dim okColor     As Boolean
    Dim okTheme     As Boolean
    Dim isBuiltIn   As Boolean
    Dim total       As Long
    Dim autoN       As Long
    Dim paletteN    As Long
    Dim themeN      As Long
    Dim anomalyN    As Long
    Dim skippedN    As Long
    Dim biSkippedN  As Long

    Set oDoc = ActiveDocument
    Debug.Print "AuditNonPaletteStyleColors: classifying Font.Color across all styles" & _
                IIf(IncludeBuiltIn, " (including built-ins)", " (custom + linked only)") & "..."
    Debug.Print "Tier 1=Automatic OK; Tier 2=Palette OK; Theme=banned-but-built-in; Anomaly=hand-typed off-palette."

    For Each oStyle In oDoc.Styles
        styleType = oStyle.Type
        If styleType = wdStyleTypeTable Or styleType = wdStyleTypeList Then
            skippedN = skippedN + 1
        Else
            On Error Resume Next
            isBuiltIn = False
            isBuiltIn = oStyle.BuiltIn
            On Error GoTo 0

            If isBuiltIn And Not IncludeBuiltIn Then
                biSkippedN = biSkippedN + 1
            Else
                c = 0
                okColor = False
                themeIdx = WD_THEME_NONE
                okTheme = False
                On Error Resume Next
                c = oStyle.Font.Color
                okColor = (Err.Number = 0)
                Err.Clear
                ' Theme colour lives on Font.TextColor.ObjectThemeColor in
                ' Word's modern object model. The direct Font.ObjectThemeColor
                ' is not a member.
                themeIdx = oStyle.Font.TextColor.ObjectThemeColor
                okTheme = (Err.Number = 0)
                On Error GoTo 0

                If Not okColor Then
                    skippedN = skippedN + 1
                Else
                    total = total + 1
                    If okTheme And themeIdx <> WD_THEME_NONE Then
                        themeN = themeN + 1
                        Debug.Print "  THEME    style=[" & oStyle.NameLocal & "]" & _
                                    "  type=[" & StyleTypeName(styleType) & "]" & _
                                    "  builtin=" & isBuiltIn & _
                                    "  ObjectThemeColor=" & themeIdx & _
                                    "  Font.Color=" & c & _
                                    "  InUse=" & oStyle.InUse
                    ElseIf c = wdColorAutomatic Then
                        autoN = autoN + 1
                    Else
                        paletteName = NameFromColor(c)
                        If Len(paletteName) > 0 Then
                            paletteN = paletteN + 1
                        Else
                            anomalyN = anomalyN + 1
                            Debug.Print "  ANOMALY  style=[" & oStyle.NameLocal & "]" & _
                                        "  type=[" & StyleTypeName(styleType) & "]" & _
                                        "  builtin=" & isBuiltIn & _
                                        "  Font.Color=" & c & _
                                        ColorDisplay(c) & _
                                        "  InUse=" & oStyle.InUse
                        End If
                    End If
                End If
            End If
        End If
    Next oStyle

    Debug.Print "---"
    Debug.Print "AuditNonPaletteStyleColors summary: " & total & " styles classified."
    Debug.Print "  Tier 1 - Automatic (default text):    " & autoN
    Debug.Print "  Tier 2 - Palette (deliberate colour): " & paletteN
    Debug.Print "  Theme  (Office theme colour, banned): " & themeN
    Debug.Print "  Anomaly (hand-typed off-palette):     " & anomalyN & " <- return value"
    Debug.Print "  Skipped (Table/List/error):           " & skippedN
    Debug.Print "  BuiltIn skipped:                      " & biSkippedN
    AuditNonPaletteStyleColors = anomalyN
End Function

' Private helper for AuditNonPaletteStyleColors. Format Font.Color
' for readable output. When value is outside the [0, 0xFFFFFF] RGB
' range it is a sentinel or theme reference, not a real colour -
' label it as such rather than byte-decomposing the low bits.
Private Function ColorDisplay(ByVal c As Long) As String
    If c >= 0 And c <= &HFFFFFF Then
        ColorDisplay = " " & LongToHex(c)
    Else
        ColorDisplay = " (sentinel/theme-encoded)"
    End If
End Function

'=========================================================================
' AuditBookHyperlinkStyling  (Public; wired at slot 46)
' Verify every BookHyperlink-styled run in the document carries the
' expected character properties. Walks all primary StoryRanges and
' Finds runs by character style "BookHyperlink".
'
' For each match, verify:
'   - Font.Name      = "Carlito"
'   - Font.Size      = 9
'   - Font.Color     = ColorFromName("DarkBlue")  (= 8388608 = #000080)
'   - Font.Underline = wdUnderlineSingle
'
' Anomalies are reported per-property: each mismatched property shows
' in the output. Expected 0 after LockBookHyperlinks runs.
'
' BookHyperlink replaced the built-in Hyperlink style as the doc's
' one-form hyperlink target. The built-in inherits font/size from
' paragraph context and so cannot be enforced uniformly; the custom
' style pins all four properties explicitly. See EDSG/01-styles.md
' "Companion rule: no clickable hyperlinks anywhere" for the rule.
'
' Public so the Immediate-window stub in basStyleInspector can
' delegate in: ?AuditBookHyperlinkStyling resolves via the stub.
' Canonical body lives here per the class-encapsulation rule
' (EDSG/12-module-vs-class.md).
'=========================================================================
Public Function AuditBookHyperlinkStyling() As Long
    Const EXPECTED_STYLE As String = "BookHyperlink"
    Const EXPECTED_FONT  As String = "Carlito"
    Const EXPECTED_SIZE  As Single = 9
    Const MAX_SNIP       As Long = 40
    Dim oDoc       As Word.Document
    Dim story      As Word.Range
    Dim probe      As Word.Range
    Dim runColor   As Long
    Dim runUL      As Long
    Dim runFont    As String
    Dim runSize    As Single
    Dim expected   As Long
    Dim pageNum    As Long
    Dim anomalies  As Long
    Dim total      As Long
    Dim snip       As String
    Dim storyName  As String
    Dim mismatch   As String

    Set oDoc = ActiveDocument
    expected = ColorFromName("DarkBlue")

    Debug.Print "AuditBookHyperlinkStyling: scanning all stories for """ & _
                EXPECTED_STYLE & """-styled runs; expecting font=" & EXPECTED_FONT & _
                " " & EXPECTED_SIZE & "pt + color=" & expected & " " & LongToHex(expected) & _
                " + underline single..."

    For Each story In oDoc.StoryRanges
        Set probe = story.Duplicate
        With probe.Find
            .ClearFormatting
            .Text = ""
            .style = oDoc.Styles(EXPECTED_STYLE)
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchWildcards = False
        End With
        Do While probe.Find.Execute
            total = total + 1
            runColor = probe.Font.Color
            runUL = probe.Font.Underline
            runFont = probe.Font.Name
            runSize = probe.Font.Size

            mismatch = ""
            If runFont <> EXPECTED_FONT Then mismatch = mismatch & " font=[" & runFont & "]"
            If runSize <> EXPECTED_SIZE Then mismatch = mismatch & " size=" & runSize
            If runColor <> expected Then mismatch = mismatch & " color=" & runColor & " " & LongToHex(runColor)
            If runUL <> wdUnderlineSingle Then mismatch = mismatch & " underline=" & runUL

            If Len(mismatch) > 0 Then
                anomalies = anomalies + 1
                pageNum = -1
                On Error Resume Next
                pageNum = probe.Information(wdActiveEndPageNumber)
                On Error GoTo 0
                snip = probe.Text
                If Len(snip) > MAX_SNIP Then snip = Left$(snip, MAX_SNIP) & " ..."
                snip = Replace(Replace(snip, vbCr, " | "), vbLf, " | ")
                storyName = "StoryType=" & story.StoryType
                Debug.Print "  ANOMALY " & storyName & _
                            "  page=" & pageNum & _
                            "  mismatch:" & mismatch & _
                            "  text=[" & snip & "]"
            End If
            probe.Collapse wdCollapseEnd
        Loop
    Next story

    Debug.Print "AuditBookHyperlinkStyling: " & total & " " & EXPECTED_STYLE & _
                "-styled run(s) checked, " & anomalies & " anomaly/anomalies."
    AuditBookHyperlinkStyling = anomalies
End Function

'=========================================================================
' CountUnapprovedVisibleStyles  (Private; wired at slot 45)
' Walks ActiveDocument.Styles (paragraph + character + linked; skips
' wdStyleTypeTable and wdStyleTypeList) and counts styles that are
' visible-and-unapproved.
'
' Compliant if either:
'   - NameLocal is in the approved-styles SSOT
'     (basTEST_aeBibleConfig.GetApprovedStyles), or
'   - "properly hidden": Priority = 99 AND QuickStyle = False AND
'     UnhideWhenUsed = False. Applies regardless of BuiltIn.
'
' Returns Long Count of violations (expected 0 after a clean
' HideUnapprovedBuiltInStyles run). On violations > 0, prints the
' offending NameLocal list to Immediate; silent when 0.
'=========================================================================
Private Function CountUnapprovedVisibleStyles() As Long
    On Error GoTo PROC_ERR
    Dim approved As Variant
    Dim approvedDict As Object
    Dim i As Long
    approved = GetApprovedStyles()
    Set approvedDict = CreateObject("Scripting.Dictionary")
    approvedDict.CompareMode = 1 ' TextCompare
    For i = LBound(approved) To UBound(approved)
        If Not approvedDict.Exists(CStr(approved(i))) Then
            approvedDict.Add CStr(approved(i)), True
        End If
    Next i

    Dim s As Word.Style
    Dim styleType As Long
    Dim violations As Long
    Dim hidden As Boolean
    Dim violationList As String
    violations = 0

    For Each s In ActiveDocument.Styles
        styleType = s.Type
        If styleType <> wdStyleTypeTable And styleType <> wdStyleTypeList Then
            If Not approvedDict.Exists(s.NameLocal) Then
                hidden = (s.Priority = 99) And (s.QuickStyle = False) And (s.UnhideWhenUsed = False)
                If Not hidden Then
                    violations = violations + 1
                    violationList = violationList & "  - " & s.NameLocal & vbCrLf
                End If
            End If
        End If
    Next s

    If violations > 0 Then
        Debug.Print "CountUnapprovedVisibleStyles: " & violations & " violation(s) -"
        Debug.Print violationList;
    End If

    CountUnapprovedVisibleStyles = violations

PROC_EXIT:
    Set approvedDict = Nothing
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountUnapprovedVisibleStyles | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    CountUnapprovedVisibleStyles = -1
    Resume PROC_EXIT
End Function

Private Function TestSlotAvailable() As Long
    On Error GoTo PROC_ERR

    TestSlotAvailable = -1

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.TestSlotAvailable | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountParagraphMarksWithDarkRedFormatting()
    On Error GoTo PROC_ERR

    Dim p As Word.Paragraph
    Dim Count As Long
    Count = 0
    For Each p In ActiveDocument.Paragraphs
        With p.Range.Characters.Last.Font
            If .Color = wdColorDarkRed Then
                Count = Count + 1
            End If
        End With
    Next p
    'Debug.Print "Count = " & Count
    CountParagraphMarksWithDarkRedFormatting = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in CountParagraphMarksWithDarkRedFormatting | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountBoldFootnotesWordLevel()
    On Error GoTo PROC_ERR

    Dim ft As footnote
    Dim aword As Word.Range
    Dim boldCount As Long
    Dim hasBold As Boolean

    For Each ft In ActiveDocument.Footnotes
        hasBold = False

        If ft.Range.style = "Footnote Text" Then
            For Each aword In ft.Range.words
                If aword.Font.Bold = True Then
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

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountBoldFootnotesWordLevel | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed()
    On Error GoTo PROC_ERR

    Dim rng As Word.Range
    Dim totalCount As Long
    Dim foundFirst As Boolean
    Dim docRange As Word.Range

    Set docRange = ActiveDocument.Content
    totalCount = 0
    foundFirst = False

    Set rng = docRange.Duplicate

    With rng.Find
        .ClearFormatting
        .style = ActiveDocument.Styles("Words of Jesus")
        .Font.Name = "Arial Black"
        .Font.Size = 8
        .Font.Color = wdColorDarkRed
        .Text = ""
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

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountParagraphMarks_ArialBlackDarkRed() As Long
    On Error GoTo PROC_ERR

    Dim para As Word.Paragraph
    Dim paraRange As Word.Range
    Dim totalCount As Long
    Dim foundFirst As Boolean
    Dim fontColor As WdColor

    totalCount = 0
    foundFirst = False

    For Each para In ActiveDocument.Paragraphs
        Set paraRange = para.Range
        paraRange.Start = paraRange.End - 1  ' Isolate the paragraph mark

        With paraRange.Font
            fontColor = .Color
            If .Name = "Arial Black" And _
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

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountParagraphMarks_ArialBlackDarkRed | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountParagraphMarks_ArialBlack() As Long
    On Error GoTo PROC_ERR

    Dim para As Word.Paragraph
    Dim paraRange As Word.Range
    Dim totalCount As Long
    Dim foundFirst As Boolean
    Dim fontColor As WdColor

    totalCount = 0
    foundFirst = False

    For Each para In ActiveDocument.Paragraphs
        Set paraRange = para.Range
        paraRange.Start = paraRange.End - 1  ' Isolate the paragraph mark

        With paraRange.Font
            fontColor = .Color
            If .Name = "Arial Black" And _
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

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountParagraphMarks_ArialBlack | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountEmptyParagraphs() As Long
    On Error GoTo PROC_ERR

    Dim para As Word.Paragraph
    Dim Count As Long
    Count = 0
    For Each para In ActiveDocument.Paragraphs
        If Len(para.Range.Text) = 1 And para.Range.Text = vbCr Then
            Count = Count + 1
        End If
    Next para
    CountEmptyParagraphs = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountEmptyParagraphs | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountDocTabOnlyParagraphs() As Long
    On Error GoTo PROC_ERR

' The final paragraph in a Word document always includes an extra paragraph mark that you can't delete.
' The extra Count is usually the final paragraph, which Word always includes and may contain leftover tab formatting.
' This code can filter that out.
    Dim para As Word.Paragraph
    Dim countTabOnly As Long
    Dim paraText As String

    countTabOnly = 0

    For Each para In ActiveDocument.Paragraphs
        paraText = para.Range.Text

        ' Skip final paragraph if it's the only one and contains just tab+CR
        If para.Range.End = ActiveDocument.Content.End And paraText = vbTab & vbCr Then
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

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountDocTabOnlyParagraphs | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountFooterParagraphsWithFooterStyle() As Long
    On Error GoTo PROC_ERR

    Dim sec As Word.Section
    Dim hdrFtr As HeaderFooter
    Dim para As Word.Paragraph
    Dim Count As Long

    Count = 0

    ' Loop through all sections in the document
    For Each sec In ActiveDocument.Sections
        ' Only check footers
        For Each hdrFtr In sec.Footers
            If hdrFtr.Exists Then
                For Each para In hdrFtr.Range.Paragraphs
                    If para.style = "Footer" Then
                        Count = Count + 1
                    End If
                Next para
            End If
        Next hdrFtr
    Next sec

    'MsgBox "Number of paragraphs in footers with the style 'Footer': " & Count, vbInformation
    'Debug.Print "Number of paragraphs in footers with the style 'Footer': " & Count
    CountFooterParagraphsWithFooterStyle = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountFooterParagraphsWithFooterStyle | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountManualLineBreaksAndWithSpace(Optional space As Variant) As Long
    On Error GoTo PROC_ERR

    Dim fn As footnote
    Dim totalBreaks As Long
    Dim breaksWithSpaceBefore As Long
    Dim i As Long
    Dim fnIdx As Long
    Dim content As String
    Dim currentChar As String
    Dim prevChar As String

    totalBreaks = 0
    breaksWithSpaceBefore = 0
    fnIdx = 0

    For Each fn In ActiveDocument.Footnotes
        fnIdx = fnIdx + 1
        content = fn.Range.Text
        For i = 1 To Len(content)
            currentChar = Mid$(content, i, 1)
            If Asc(currentChar) = 11 Then ' Chr(11) = manual line break
                totalBreaks = totalBreaks + 1
                If m_lastHint = "" Then _
                    m_lastHint = "Footnote " & fnIdx & ": " & Left(content, 50)
                If i > 1 Then
                    prevChar = Mid$(content, i - 1, 1)
                    If prevChar = " " Then
                        breaksWithSpaceBefore = breaksWithSpaceBefore + 1
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

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountManualLineBreaksAndWithSpace | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountLinefeed(Optional space As Variant) As Long
    On Error GoTo PROC_ERR

' Search in doc
    Dim doc As Document
    Dim rng As Word.Range
    Dim findText As String
    Dim Count As Long

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
    Set rng = doc.Content
    
    Application.ScreenUpdating = False
    With rng.Find
        .Text = findText
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
        
        ' Loop through all instances and Count them
        Do While .Execute
            Count = Count + 1
            If m_lastHint = "" Then m_lastHint = "First match: " & Left(rng.Text, 60)
            ' Keep Word responsive every 50
            If Count Mod 500 = 0 Then DoEvents
        Loop
    End With
    Application.ScreenUpdating = True
    
    ' Display the Count
    'MsgBox "Number of matches found: " & Count
    'Debug.Print "Number of matches found: " & Count
    CountLinefeed = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountLinefeed | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountParagraphMarksPerHeaderSection() As Long
    On Error GoTo PROC_ERR

    Dim doc As Document
    Dim sectionIndex As Long
    Dim headerRange As Word.Range
    Dim paraCount As Long

    ' Set the document
    Set doc = ActiveDocument
    paraCount = 1
    
    ' Loop through all sections
    For sectionIndex = 1 To doc.Sections.Count
        ' Get the range of the header in the current section
        Set headerRange = doc.Sections(sectionIndex).Headers(wdHeaderFooterPrimary).Range
        
        ' Count the number of paragraphs in the header
        paraCount = headerRange.Paragraphs.Count
        
        ' Print the Result to the console
        'Debug.Print "Section " & sectionIndex & " header has " & paraCount & " paragraph(s)."
        
        ' Check if the paragraph Count is more than one
        If paraCount > 1 Then
            MsgBox "Section " & sectionIndex & " header has more than one paragraph mark. Stopping at this section."
            Exit Function
        End If
    Next sectionIndex
    
    ' Display a message indicating the Count is complete
    'MsgBox "Count of paragraph marks per header section is " & paraCount
    'Debug.Print "Count of paragraph marks per header section is " & paraCount
    CountParagraphMarksPerHeaderSection = paraCount

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountParagraphMarksPerHeaderSection | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountHeaderStyleUsage() As Long
    On Error GoTo PROC_ERR

    ' Counts header paragraphs whose style is NOT "TheHeaders".
    ' Expected = 0: every paragraph in every primary header must use "TheHeaders".
    ' Searches header stories (not doc.Content, which excludes headers/footers).

    Dim doc As Document
    Dim sec As Word.Section
    Dim hdr As HeaderFooter
    Dim para As Word.Paragraph
    Dim Count As Long

    Set doc = ActiveDocument
    Count = 0

    For Each sec In doc.Sections
        Set hdr = sec.Headers(wdHeaderFooterPrimary)
        If hdr.Exists Then
            For Each para In hdr.Range.Paragraphs
                If para.style.NameLocal <> "TheHeaders" Then
                    Count = Count + 1
                    If m_lastHint = "" Then _
                        m_lastHint = "Section " & sec.index & ": style='" & para.style.NameLocal & "'"
                End If
            Next para
        End If
    Next sec

    CountHeaderStyleUsage = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountHeaderStyleUsage | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountParagraphsWithoutTabInHeaders() As Long
    On Error GoTo PROC_ERR

    Dim doc As Document
    Dim sectionIndex As Long
    Dim headerRange As Word.Range
    Dim para As Word.Paragraph
    Dim Count As Long
    Dim totalCount As Long

    ' Set the document
    Set doc = ActiveDocument
    
    ' Initialize the total Count
    totalCount = 0
    
    ' Loop through all sections
    For sectionIndex = 1 To doc.Sections.Count
        ' Get the range of the header in the current section
        Set headerRange = doc.Sections(sectionIndex).Headers(wdHeaderFooterPrimary).Range
        
        ' Initialize the Count for the current header
        Count = 0
        
        ' Loop through each paragraph in the header range
        For Each para In headerRange.Paragraphs
            ' Check if the paragraph does not contain a tab
            If InStr(para.Range.Text, vbTab) = 0 Then
                Count = Count + 1
                If m_lastHint = "" Then _
                    m_lastHint = "Section " & sectionIndex & ": " & Left(para.Range.Text, 50)
            End If
        Next para

        ' Update the total Count
        totalCount = totalCount + Count
        
        ' Display the Count for the current header
        'MsgBox "Section " & sectionIndex & " header has " & Count & " paragraphs without a tab."
        'Debug.Print "Section " & sectionIndex & " header has " & Count & " paragraphs without a tab."
    Next sectionIndex
    
    ' Display the total Count
    'MsgBox "Total paragraphs without a tab in all headers: " & totalCount
    'Debug.Print "Total paragraphs without a tab in all headers: " & totalCount
    CountParagraphsWithoutTabInHeaders = totalCount

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountParagraphsWithoutTabInHeaders | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountTabFollowedByParagraphMarkInHeaders() As Long
    On Error GoTo PROC_ERR

    Dim doc As Document
    Dim sectionIndex As Long
    Dim headerRange As Word.Range
    Dim para As Word.Paragraph
    Dim Count As Long
    Dim totalCount As Long

    ' Set the document
    Set doc = ActiveDocument
    
    ' Initialize the total Count
    totalCount = 0
    
    ' Loop through all sections
    For sectionIndex = 1 To doc.Sections.Count
        ' Get the range of the header in the current section
        Set headerRange = doc.Sections(sectionIndex).Headers(wdHeaderFooterPrimary).Range
        
        ' Initialize the Count for the current header
        Count = 0
        
        ' Loop through each paragraph in the header range
        For Each para In headerRange.Paragraphs
            ' Check if the paragraph contains a tab followed by a paragraph mark
            If InStr(para.Range.Text, vbTab & vbCr) > 0 Then
                Count = Count + 1
                If m_lastHint = "" Then _
                    m_lastHint = "Section " & sectionIndex & ": " & Left(para.Range.Text, 50)
            End If
        Next para

        ' Update the total Count
        totalCount = totalCount + Count
        
        ' Display the Count for the current header
        'MsgBox "Section " & sectionIndex & " header has " & Count & " paragraphs with a tab followed by a paragraph mark."
        'Debug.Print "Section " & sectionIndex & " header has " & Count & " paragraphs with a tab followed by a paragraph mark."
    Next sectionIndex
    
    ' Display the total Count
    'MsgBox "Total paragraphs with a tab followed by a paragraph mark in all headers: " & totalCount
    'Debug.Print "Total paragraphs with a tab followed by a paragraph mark in all headers: " & totalCount
    CountTabFollowedByParagraphMarkInHeaders = totalCount

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountTabFollowedByParagraphMarkInHeaders | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CheckAllHeaders(EmptyOrNotEmpty As String) As Long
    On Error GoTo PROC_ERR

' Identifies headers with only paragraph marks as empty. Checks if the header contains any visible text or characters beyond just paragraph marks.
'1. Sets the active document.
'2. Initializes counters for empty headers, not empty headers, and total headers.
'3. Loops through all sections in the document.
'4. Defines the range of the header in each section.
'5. Checks if the header is empty by iterating through each paragraph in the header range and ensuring it contains more than just a paragraph mark.
'6. Updates the counters based on whether the header is empty or not.
'7. Displays a message box with the total number of headers, the number of empty headers, and the number of not empty headers.    Dim doc As Document
    Dim sectionIndex As Long
    Dim headerRange As Word.Range
    Dim para As Word.Paragraph
    Dim IsEmpty As Boolean
    Dim emptyCount As Long
    Dim notEmptyCount As Long
    Dim totalHeaders As Long

    ' Set the document
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Initialize counters
    emptyCount = 0
    notEmptyCount = 0
    totalHeaders = 0
    
    ' Loop through all sections
    For sectionIndex = 1 To doc.Sections.Count
        ' Get the range of the header in the current section
        Set headerRange = doc.Sections(sectionIndex).Headers(wdHeaderFooterPrimary).Range
        
        ' Check if the header is empty
        IsEmpty = True
        For Each para In headerRange.Paragraphs
            If Len(Trim(para.Range.Text)) > 1 Then ' Check if paragraph contains more than just a paragraph mark
                IsEmpty = False
                Exit For
            End If
        Next para
        
        ' Update counters
        totalHeaders = totalHeaders + 1
        If IsEmpty Then
            emptyCount = emptyCount + 1
            If m_lastHint = "" And EmptyOrNotEmpty = "Empty" Then _
                m_lastHint = "First empty header: Section " & sectionIndex
        Else
            notEmptyCount = notEmptyCount + 1
            If m_lastHint = "" And EmptyOrNotEmpty = "NotEmpty" Then _
                m_lastHint = "First non-empty header: Section " & sectionIndex
        End If
    Next sectionIndex
    
    'Debug.Print "Total headers: " & totalHeaders & vbCrLf & _
    '       "Empty headers: " & emptyCount & vbCrLf & _
    '       "Not empty headers: " & notEmptyCount
    
    If EmptyOrNotEmpty = "Empty" Then
        CheckAllHeaders = emptyCount
        GoTo PROC_EXIT
    ElseIf EmptyOrNotEmpty = "NotEmpty" Then
        CheckAllHeaders = notEmptyCount
        GoTo PROC_EXIT
    Else
        MsgBox "Parameter not valid", vbCritical, "CheckAllHeaders"
        Stop
    End If

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CheckAllHeaders | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountFootnoteReferenceColors() As Long
    On Error GoTo PROC_ERR

    Dim doc As Document
    Set doc = ActiveDocument

    Dim footnoteColors As Collection
    Set footnoteColors = New Collection
    
    Dim footnote As footnote
    Dim colorHex As String
    
    For Each footnote In doc.Footnotes
        colorHex = ColorToHex(footnote.Reference.Font.Color)
        On Error Resume Next
        footnoteColors.Add colorHex, colorHex
        On Error GoTo 0
    Next footnote
       
    Dim colorList As String
    colorList = "Colors of Footnote References in Hex: "
    
    Dim color As Variant
    Dim colorCount As Long
    colorCount = 0
    For Each color In footnoteColors
        colorCount = colorCount + 1
        colorList = colorList & color & vbCrLf
    Next color
    
    'MsgBox colorList
    'Debug.Print colorList
    CountFootnoteReferenceColors = colorCount

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountFootnoteReferenceColors | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function ColorToHex(color As Long) As String
    On Error GoTo PROC_ERR

    ColorToHex = Right("000000" & Hex(color), 6)

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.ColorToHex | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountFootnoteReferences() As Long
    On Error GoTo PROC_ERR

    Dim doc As Document
    Set doc = ActiveDocument

    Dim footnoteCount As Long
    footnoteCount = doc.Footnotes.Count

    ' Display the results
    'MsgBox "Number of Footnote References: " & footnoteCount
    CountFootnoteReferences = footnoteCount

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountFootnoteReferences | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountDeleteEmptyParagraphsBeforeHeading2() As Long
    On Error GoTo PROC_ERR

' Avoids indexing (.Item(i))� slow in Word.
' Does not call .Style unless needed.
' Only trims the text once per loop.
' Dramatically faster!
    Dim para As Word.Paragraph
    Dim prevText As String
    Dim prevIsEmpty As Boolean
    Dim Count As Long
    Dim paraCount As Long

    Count = 0
    paraCount = 0
    prevIsEmpty = False

    Application.ScreenUpdating = False
    ' Loop through paragraphs only once
    For Each para In ActiveDocument.Range.Paragraphs
        paraCount = paraCount + 1
        If para.style = "Heading 2" Then
            If prevIsEmpty Then
                Count = Count + 1
            End If
        End If

        ' Determine if current paragraph is empty (ignoring whitespace)
        prevText = Trim(para.Range.Text)
        prevIsEmpty = (prevText = "" Or prevText = Chr(13)) ' Chr(13) is vbCr
        ' Keep Word responsive every 500 paragraphs
        If paraCount Mod 500 = 0 Then DoEvents
    Next para
    Application.ScreenUpdating = True

    'MsgBox "Fast Count of 'Heading 2' paragraphs preceded by an empty paragraph: " & Count, vbInformation
    'Debug.Print "Fast Count of 'Heading 2' paragraphs preceded by an empty paragraph: " & Count
    CountDeleteEmptyParagraphsBeforeHeading2 = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountDeleteEmptyParagraphsBeforeHeading2 | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountEmptyParagraphsWithFormatting() As Long
    On Error GoTo PROC_ERR

' To maximize speed and accuracy, stay in the Paragraphs collection (for true Word structure),
' but avoid anything expensive - no .Range.Text, .Style, etc. unless needed.
    Dim para As Word.Paragraph
    Dim rng As Word.Range
    Dim Count As Long
    Dim paraCount As Long

    Count = 0
    Set rng = ActiveDocument.Content
    rng.Collapse Direction:=wdCollapseStart

    Application.ScreenUpdating = False
    ' Loop over all paragraphs using the built-in collection
    For Each para In ActiveDocument.Paragraphs
        paraCount = paraCount + 1
        ' Work directly with Range.Text but trim just once
        If LenB(Trim$(Replace(para.Range.Text, vbCr, ""))) = 0 Then
            Count = Count + 1
        End If
        ' Keep Word responsive every 500 paragraphs
        If paraCount Mod 500 = 0 Then DoEvents
    Next para
    Application.ScreenUpdating = True

    'MsgBox "Accurate empty paragraph Count: " & Count, vbInformation
    'Debug.Print "Accurate empty paragraph Count: " & Count

    CountEmptyParagraphsWithFormatting = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountEmptyParagraphsWithFormatting | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountNotSpacesAfterFootnoteReferences() As Long
    On Error GoTo PROC_ERR

' Also shows Footnote References and Following Characters (ASCII Values)
    Dim doc As Document
    Dim footnote As footnote
    Dim rng As Word.Range
    Dim spaceCount As Long
    Dim notSpaceCount As Long
    Dim debugInfo As String
    Dim asciiValue As Long

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
        asciiValue = Asc(rng.Text)
        
        ' Exclude ASCII value 13 (carriage return)
        If asciiValue <> 13 Then
            ' Add debug information
            If asciiValue <> 32 Then    ' space (ASCII value 32)
                debugInfo = debugInfo & "Fn " & footnote.index & ": '" & rng.Text & "' (ASC " & asciiValue & "), "
                'Debug.Print debugInfo
                notSpaceCount = notSpaceCount + 1
                'Stop
            End If
            
            If asciiValue = 32 Then
                spaceCount = spaceCount + 1
            End If
        End If
    Next footnote
    
    ' Display the Count of spaces after footnote references
    'MsgBox "Number of spaces after footnote references: " & spaceCount
    ' Display debug information
    'MsgBox debugInfo
    
    ' Print the Result to the console
    'Debug.Print "Number of spaces after footnote references: " & spaceCount
    'Debug.Print debugInfo
    CountNotSpacesAfterFootnoteReferences = notSpaceCount

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountNotSpacesAfterFootnoteReferences | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountFootnotesFollowedByDigit() As Long
    On Error GoTo PROC_ERR

    Dim doc As Document
    Dim footnote As footnote
    Dim rng As Word.Range
    Dim digitCount As Long
    Dim debugInfo As String
    Dim asciiValue As Long

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
        asciiValue = Asc(rng.Text)
        
        ' Check if the character after the footnote reference is a digit (ASCII values 48-57)
        If asciiValue >= 48 And asciiValue <= 57 Then
            digitCount = digitCount + 1
        End If
        
        ' Add debug information
        debugInfo = debugInfo & "Footnote " & footnote.index & ": '" & rng.Text & "' (ASCII: " & asciiValue & ")" & vbCrLf
    Next footnote
    
    ' Display the Count of footnote references followed by a digit
    'MsgBox "Number of footnote references followed by a digit: " & digitCount
    
    ' Display debug information
    'MsgBox debugInfo
    'Debug.Print debugInfo
 
    ' Print the Result to the console
    'Debug.Print "Number of footnote references followed by a digit: " & digitCount
    CountFootnotesFollowedByDigit = digitCount

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountFootnotesFollowedByDigit | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountEmptyParasAfterH2() As Long
    On Error GoTo PROC_ERR

    Dim para As Word.Paragraph
    Dim emptyParaCount As Long
    Dim NextPara As Word.Paragraph

    ' Initialize the Count
    emptyParaCount = 0
    
    ' Loop through each paragraph in the document
    For Each para In ActiveDocument.Paragraphs
        ' Check if the paragraph style is Heading 2
        If para.style = ActiveDocument.Styles("Heading 2") Then
            ' Check the next paragraph
            Set NextPara = para.Range.Next(Unit:=wdParagraph, Count:=1).Paragraphs(1)
            If Not NextPara Is Nothing Then
                ' Check if the next paragraph is empty
                If Len(Trim(NextPara.Range.Text)) = 1 Then ' Adjusted to check for a single paragraph mark
                    emptyParaCount = emptyParaCount + 1
                End If
            End If
        End If
    Next para
    
    ' Print the Count to the console
    'Debug.Print "Number of empty paragraphs immediately after Heading 2: " & emptyParaCount
    CountEmptyParasAfterH2 = emptyParaCount

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountEmptyParasAfterH2 | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountHeading1() As Long
    On Error GoTo PROC_ERR

    Dim para As Word.Paragraph
    Dim heading1Count As Long

    ' Initialize the Count
    heading1Count = 0
    
    ' Loop through each paragraph in the document
    For Each para In ActiveDocument.Paragraphs
        ' Check if the paragraph style is Heading 1
        If para.style Like "*Heading 1*" Then   'para.style = ActiveDocument.Styles("Heading 1") Then
            heading1Count = heading1Count + 1
            'Debug.Print heading1Count & "[" & Replace(para.Range.Text, vbCr, "") & "]",
        End If
    Next para

    ' Print the Count to the console
    'Debug.Print "Number of Heading 1 paragraphs: " & heading1Count
    CountHeading1 = heading1Count
    'Stop

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountHeading1 | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

' CountActiveHyperlinks
' ---------------------
' Editorial rule: no clickable hyperlinks anywhere in the doc. Returns
' the sum of collection-Hyperlink objects across all StoryRanges
' (MainText, Footnotes, headers/footers, etc.). Expected value is 0;
' any non-zero Result is a rule violation - operator runs
' LockHyperlinksToPalette (which calls UnlinkActiveHyperlinks) to bring
' it back to 0.
'
' Replaces CountFootnoteHyperlinks (footnote-story only, 2026-05-14)
' which itself replaced CountRedFootnoteReferences (dead, 2026-05-14).
' The no-hyperlinks-in-footnotes rule is subsumed: zero across all
' stories implies zero in footnotes.
'
' Print is the primary target; clickability is a future-mode concern
' addressed at online-edition build time, not in this doc. See
' EDSG/01-styles.md "No clickable hyperlinks anywhere" for the
' editorial rule.
Private Function CountActiveHyperlinks() As Long
    On Error GoTo PROC_ERR

    Dim story As Word.Range
    Dim n     As Long
    n = 0
    For Each story In ActiveDocument.StoryRanges
        n = n + story.Hyperlinks.Count
    Next story
    CountActiveHyperlinks = n

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountActiveHyperlinks | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountTotalParagraphs() As Long
    On Error GoTo PROC_ERR

    CountTotalParagraphs = ActiveDocument.Paragraphs.Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountTotalParagraphs | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountSectionsWithDifferentFirstPage() As Long
    On Error GoTo PROC_ERR

    Dim doc As Document
    Dim sec As Word.Section
    Dim Count As Long

    Application.ScreenUpdating = False

    Set doc = ActiveDocument
    Count = 0
    
    ' Loop through each section in the document
    For Each sec In doc.Sections
        ' Check if the section does not have "Different First Page" selected
        If sec.PageSetup.DifferentFirstPageHeaderFooter Then
            Count = Count + 1
        End If
    Next sec
    
    Application.ScreenUpdating = True
    
    'MsgBox "Number of sections with 'Different First Page' selected: " & Count
    CountSectionsWithDifferentFirstPage = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountSectionsWithDifferentFirstPage | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountWhiteParagraphMarks() As Long
    On Error GoTo PROC_ERR

    Dim doc As Document
    Dim rng As Word.Range
    Dim Count As Long
    Dim firstFound As Boolean

    Set doc = ActiveDocument
    Count = 0
    firstFound = False
    
    ' Initialize the range to search the entire document
    Set rng = doc.Content
    
    ' Use the Find method to locate white paragraph marks
    With rng.Find
        .ClearFormatting
        .Font.Color = RGB(255, 255, 255) ' RGB value for white
        .Text = "^p"
        .Forward = True
        .Wrap = wdFindStop
        
        ' Loop through all occurrences
        Do While .Execute
            ' Check if the found paragraph mark has the white color
            If rng.Font.Color = RGB(255, 255, 255) Then
                Count = Count + 1
                ' Go to the first white paragraph mark found
                If Not firstFound Then
                    rng.Select
                    firstFound = True
                End If
            End If
        Loop
    End With
    
    ' Display the Count of white paragraph marks
    'MsgBox "Number of white paragraph marks: " & Count
    CountWhiteParagraphMarks = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountWhiteParagraphMarks | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountEmptyParasWithNoThemeColor()
    On Error GoTo PROC_ERR

' The value of wdNotThemeColor is -1. This constant indicates that no theme color is applied.
    Dim para As Word.Paragraph
    Dim rng As Word.Range
    Dim emptyParaCount As Long
    Dim themeColorUsed As Boolean
    Dim colorCounts As Object
    Dim colorKey As String

    Set colorCounts = CreateObject("Scripting.Dictionary")
    emptyParaCount = 0

    For Each para In ActiveDocument.Paragraphs
        Set rng = para.Range
        themeColorUsed = False
        
        ' Check if the paragraph is empty
        If Len(rng.Text) = 1 Then ' Only the paragraph mark
            ' Check if the theme color is used
            If rng.Font.TextColor.ObjectThemeColor <> wdNotThemeColor Then
                themeColorUsed = True
                colorKey = CStr(rng.Font.TextColor.ObjectThemeColor)
                
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

    ' Print out the color Count for each empty paragraph
    'Dim key As Variant
    'For Each key In colorCounts.Keys
    '    Debug.Print "Theme color " & key & " used in " & colorCounts(key) & " empty paragraphs."
    'Next key
    CountEmptyParasWithNoThemeColor = emptyParaCount

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountEmptyParasWithNoThemeColor | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountNumberDashNumberInFootnotes() As Long
    On Error GoTo PROC_ERR

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
        Set matches = regex.Execute(fn.Range.Text)
        matchCount = matchCount + matches.Count
    Next fn

    'MsgBox "Number of number-dash-number patterns in footnotes: " & matchCount, vbInformation
    'Debug.Print "Number of number-dash-number patterns in footnotes: " & matchCount
    CountNumberDashNumberInFootnotes = matchCount

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountNumberDashNumberInFootnotes | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountFindNumberDashNumber() As Long
    On Error GoTo PROC_ERR

    Dim rng As Word.Range
    Dim searchPattern As String
    Dim userInterrupt As VbMsgBoxResult
    Dim Count As Long

    searchPattern = "[0-9]{1,}-[0-9]{1,}" ' Pattern to find number-dash-number
    Count = 0

    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = searchPattern
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop

        Do While .Execute
            ' Move the selection to the found text
            rng.Select
            Count = Count + 1
            'userInterrupt = MsgBox("Found: " & rng.Text, vbOKCancel)

            ' Check for user interruption
            'If userInterrupt = vbCancel Then
            '    MsgBox "Search interrupted. Total found: " & Count
            '    Exit Function
            'End If

            ' Collapse the range to the end of the found text
            rng.Collapse Direction:=wdCollapseEnd
            rng.End = ActiveDocument.Content.End
        
            If Count Mod 50 = 0 Then
                DoEvents  ' Keep UI responsive every 50 updates
            End If
        Loop
    End With

    CountFindNumberDashNumber = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountFindNumberDashNumber | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountNonBreakingSpaces() As Long
    On Error GoTo PROC_ERR

    Dim rng As Word.Range
    Dim Count As Long

    Set rng = ActiveDocument.Content
    rng.Find.ClearFormatting

    With rng.Find
        .Text = "^s"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
    End With

    Count = 0
    Do While rng.Find.Execute
        Count = Count + 1
        rng.Collapse wdCollapseEnd
    
        If Count Mod 50 = 0 Then
            DoEvents  ' Keep UI responsive every 50 updates
        End If
    Loop

    CountNonBreakingSpaces = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountNonBreakingSpaces | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountPeriodSpaceLeftParenthesis() As Long
    On Error GoTo PROC_ERR

    Dim rng As Word.Range
    Dim Count As Long

    Set rng = ActiveDocument.Content
    rng.Find.ClearFormatting
    
    With rng.Find
        .Text = ". ("
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
    End With
    
    Count = 0
    Do While rng.Find.Execute
        Count = Count + 1
        rng.Collapse wdCollapseEnd
    
        If Count Mod 50 = 0 Then
            DoEvents  ' Keep UI responsive every 50 updates
        End If
    Loop
    
    CountPeriodSpaceLeftParenthesis = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountPeriodSpaceLeftParenthesis | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountStyleWithNumberAndSpace() As Long
    On Error GoTo PROC_ERR

    Dim rng As Word.Range
    Dim Count As Long
    Dim StyleName As String

    StyleName = "Verse marker"

    Set rng = ActiveDocument.Content
    rng.Find.ClearFormatting
    rng.Find.style = StyleName

    With rng.Find
        .Text = "[0-9] "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchWildcards = True
    End With

    Count = 0
    Do While rng.Find.Execute
        Count = Count + 1
        rng.Collapse wdCollapseEnd
        
        If Count Mod 50 = 0 Then
            DoEvents  ' Keep UI responsive every 50 updates
        End If
    Loop

    CountStyleWithNumberAndSpace = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountStyleWithNumberAndSpace | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountStyleWithSpaceAndNumber() As Long
    On Error GoTo PROC_ERR

    Dim rng As Word.Range
    Dim Count As Long
    Dim StyleName As String
    Dim firstFound As Boolean

    StyleName = "Chapter Verse marker"

    Set rng = ActiveDocument.Content
    rng.Find.ClearFormatting
    rng.Find.style = StyleName
    firstFound = False

    With rng.Find
        .Text = " [0-9]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchWildcards = True
    End With

    Count = 0
    Do While rng.Find.Execute
        Count = Count + 1
            If Not firstFound Then
                rng.Select
                firstFound = True
                'Stop
                Exit Do ' Exit the loop after finding the first match
            End If
        rng.Collapse wdCollapseEnd
    
        If Count Mod 50 = 0 Then
            DoEvents  ' Keep UI responsive every 50 updates
        End If
    Loop

    CountStyleWithSpaceAndNumber = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountStyleWithSpaceAndNumber | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountQuadrupleParagraphMarks() As Long
    On Error GoTo PROC_ERR

    Dim rng As Word.Range
    Dim Count As Long

    Set rng = ActiveDocument.Content
    rng.Find.ClearFormatting
    
    With rng.Find
        .Text = "^13^13^13^13"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
    End With
    
    Count = 0
    Do While rng.Find.Execute
        Count = Count + 1
        rng.Collapse wdCollapseEnd
    Loop
    
    'MsgBox "CountQuadrupleParagraphMarks = " & Count
    'Dim answer As Long
    'answer = MsgBox("CountQuadrupleParagraphMarks Do you want to continue?", vbQuestion + vbYesNo, "Continue or Stop")
    'If answer = vbYes Then
    '    Resume
    'Else
    '    Stop
    'End If
    CountQuadrupleParagraphMarks = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountQuadrupleParagraphMarks | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountWhiteSpaceAndCarriageReturn() As Long
    On Error GoTo PROC_ERR

    Dim rng As Word.Range
    Dim Count As Long

    Set rng = ActiveDocument.Content
    rng.Find.ClearFormatting
    rng.Find.Font.Color = wdColorWhite

    ' Set the Find properties
    With rng.Find
        .Text = " ^13"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchWildcards = False
    End With

    Count = 0
    Do While rng.Find.Execute
        Count = Count + 1
        rng.Collapse wdCollapseEnd
    Loop

    CountWhiteSpaceAndCarriageReturn = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountWhiteSpaceAndCarriageReturn | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountDoubleTabs() As Long
    On Error GoTo PROC_ERR

    Dim doc As Document
    Dim rng As Word.Range
    Dim Count As Long

    Count = 0
    Set doc = ActiveDocument
    Set rng = doc.Content

    Do While rng.Find.Execute(findText:="^t^t", Forward:=True, Wrap:=wdFindStop)
        Count = Count + 1
        rng.Collapse Direction:=wdCollapseEnd ' Move the range to the end of the found text
    Loop

    CountDoubleTabs = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountDoubleTabs | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountSpaceFollowedByCarriageReturn() As Long
    On Error GoTo PROC_ERR

    Dim doc As Document
    Dim rng As Word.Range
    Dim Count As Long

    Count = 0
    Set doc = ActiveDocument
    Set rng = doc.Content

    Do While rng.Find.Execute(findText:=" ^13", Forward:=True, Wrap:=wdFindStop)
        Count = Count + 1
        If m_lastHint = "" Then m_lastHint = Left(rng.Paragraphs(1).Range.Text, 60)
        rng.Collapse Direction:=wdCollapseEnd ' Move the range to the end of the found text
    Loop

    CountSpaceFollowedByCarriageReturn = Count

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountSpaceFollowedByCarriageReturn | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

Private Function CountDoubleSpaces() As Long
    On Error GoTo PROC_ERR

    Dim doc As Document
    Dim rng As Word.Range
    Dim doubleSpaceCount As Long

    doubleSpaceCount = 0
    Set doc = ActiveDocument
    Set rng = doc.Content
    
    With rng.Find
        .Text = "  " ' Double space
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

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in aeBibleClass.CountDoubleSpaces | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    m_lastFuncError = True
    Resume PROC_EXIT
End Function

' Function to Count occurrences of a substring in a string
Private Function CountOccurrences(ByVal text As String, ByVal substring As String) As Long
    Dim pos, Count As Long
    pos = 1
    Count = 0
    Do While InStr(pos, text, substring) > 0
        pos = InStr(pos, text, substring) + Len(substring)
        Count = Count + 1
    Loop
    
    CountOccurrences = Count
End Function

' Count double spaces inside shapes, including those within groups, with error handling
Private Function CountDoubleSpacesInShapes() As Long
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
    
    ' Display the Count of double spaces in shapes
    CountDoubleSpacesInShapes = doubleSpaceCount
End Function

Private Sub ProcessShape(ByVal shp As shape, ByRef doubleSpaceCount As Long)
    Dim textRange As Word.Range
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
                shapeText = textRange.Text
                ' Count double spaces in the shape's text
                doubleSpaceCount = doubleSpaceCount + CountOccurrences(shapeText, "  ")
            End If
        ElseIf Not shp.TextFrame2 Is Nothing Then
            If shp.TextFrame2.HasText Then
                ' Get the text range within the shape
                Set textRange = shp.TextFrame2.textRange
                ' Get the text content of the shape
                shapeText = textRange.Text
                ' Count double spaces in the shape's text
                doubleSpaceCount = doubleSpaceCount + CountOccurrences(shapeText, "  ")
            End If
        End If
    End If
End Sub

