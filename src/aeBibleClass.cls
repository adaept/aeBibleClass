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

Private Const BibleClassVERSION As String = "0.0.6"
Private Const BibleClassVERSION_DATE As String = "April 2, 2025"

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
Private Result As Integer
Private Expected As Integer
'=====================================================

Private Sub Class_Initialize()
    On Error GoTo 0
    'MsgBox Application.VBE.ActiveVBProject.Name
    'MsgBox Application.Caption
    If Application.VBE.ActiveVBProject.name = "BibleClass" Then
        Application.Caption = Application.VBE.ActiveVBProject.name & " " & BibleClassVERSION
    End If
    Debug.Print "Class_Initialize"
 End Sub

Private Sub Class_Terminate()
    On Error GoTo 0
    Debug.Print
    Debug.Print "Class_Terminate"
    Debug.Print , Application.VBE.ActiveVBProject.name
    Debug.Print , "BibleClass VERSION: " & BibleClassVERSION
    Debug.Print , "BibleClass VERSION_DATE: " & BibleClassVERSION_DATE
End Sub

Public Property Get TheBibleClassTests(Optional ByVal varDebug As Variant) As Boolean

    On Error GoTo PROC_ERR

    'MsgBox "TheBibleClassTests"
    If IsMissing(varDebug) Then
        Debug.Print "Get TheBibleClassTests"
        Debug.Print , "varDebug IS missing so no parameter is passed to RunBibleClassTests"
        Debug.Print , "DEBUGGING IS OFF"
        TheBibleClassTests = RunBibleClassTests
    Else
        Debug.Print "Get TheBibleClassTests"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to RunBibleClassTests"
        Debug.Print , "DEBUGGING TURNED ON"
        TheBibleClassTests = RunBibleClassTests(varDebug)
    End If

PROC_EXIT:
    Exit Property

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TheBibleClassTests of Class BibleClass"
    Resume PROC_EXIT

End Property

Function ConvertToOneBasedArray(values As Variant) As Variant
    Dim strArray() As String
    Dim i As Integer
    
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
    Dim values As Variant
    Dim i As Integer
    Dim outputString As String
    
    ' Define the Expected RunTest result values to store in the array
    '      RunTest 1  2  3  4  5  6  7  8  9  10 11 12 13 14 15  16  17 18  19 20 21  22  23 24    25)
    values = Array(0, 0, 0, 0, 0, 0, 0, 0, 7, 1, 2, 0, 0, 0, -1, -1, 0, 66, 0, 0, -1, -1, 0, 1000, 1)
    
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

    ' Next results
    outputString = ""
    For i = 16 To UBound(oneBasedExpectedArray)
        outputString = outputString & " Test(" & i & ")" & oneBasedExpectedArray(i) & " "
        ' Trim the trailing space
        outputString = Trim(outputString)
    Next i
    Debug.Print outputString

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
    Else
        Debug.Print , "varDebug IS NOT missing so blnDebug of RunBibleClassTests is set to True"
        Debug.Print , "NOW DEBUGGING..."
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
            Debug.Print "Pass/Fail", "Copy ()", "Test Num", "Result", "Expected", "Function"
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

    RunBibleClassTests = True

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RunBibleClassTests of Class BibleClass"
    Resume PROC_EXIT

End Function

Private Function GetPassFail(TestNum As Integer) As String
    Dim PassFail As String

    ' Set explicit fail unless the test passes
    GetPassFail = "FAIL!!!"
    Select Case TestNum
    Case 1
        Result = CountDoubleSpaces
        Expected = oneBasedExpectedArray(1)
        If Result = Expected Then GetPassFail = "PASS"
    Case 2
        Result = CountDoubleSpacesInShapes
        Expected = oneBasedExpectedArray(2)
        If Result = Expected Then GetPassFail = "PASS"
    Case 3
        Result = CountSpaceFollowedByCarriageReturn
        Expected = oneBasedExpectedArray(3)
        If Result = Expected Then GetPassFail = "PASS"
    Case 4
        Result = CountDoubleTabs
        Expected = oneBasedExpectedArray(4)
        If Result = Expected Then GetPassFail = "PASS"
    Case 5
        Result = CountWhiteSpaceAndCarriageReturn
        Expected = oneBasedExpectedArray(5)
        If Result = Expected Then GetPassFail = "PASS"
    Case 6
        Result = CountQuadrupleParagraphMarks
        Expected = oneBasedExpectedArray(6)
        If Result = Expected Then GetPassFail = "PASS"
    Case 7
        Result = CountStyleWithSpaceAndNumber
        Expected = oneBasedExpectedArray(7)
        If Result = Expected Then GetPassFail = "PASS"
    Case 8
        Result = CountStyleWithNumberAndSpace
        Expected = oneBasedExpectedArray(8)
        If Result = Expected Then GetPassFail = "PASS"
    Case 9
        Result = CountPeriodSpaceLeftParenthesis
        Expected = oneBasedExpectedArray(9)
        If Result = Expected Then GetPassFail = "PASS"
    Case 10
        Result = CountNonBreakingSpaces
        Expected = oneBasedExpectedArray(10)
        If Result = Expected Then GetPassFail = "PASS"
    Case 11
        Result = CountFindNumberDashNumber
        Expected = oneBasedExpectedArray(11)
        If Result = Expected Then GetPassFail = "PASS"
    Case 12
        Result = CountNumberDashNumberInFootnotes
        Expected = oneBasedExpectedArray(12)
        If Result = Expected Then GetPassFail = "PASS"
    Case 13
        Result = CountEmptyParasWithNoThemeColor
        Expected = oneBasedExpectedArray(13)
        If Result = Expected Then GetPassFail = "PASS"
    Case 14
        Result = CountWhiteParagraphMarks
        Expected = oneBasedExpectedArray(14)
        If Result = Expected Then GetPassFail = "PASS"
    Case 15
        Result = CountSectionsWithDifferentFirstPage
        Expected = oneBasedExpectedArray(15)
        If Result = Expected Then GetPassFail = "PASS"
    Case 16
        Result = CountTotalParagraphs
        Expected = oneBasedExpectedArray(16)
        If Result = Expected Then GetPassFail = "PASS"
    Case 17
        Result = CountRedFootnoteReferences
        Expected = oneBasedExpectedArray(17)
        If Result = Expected Then GetPassFail = "PASS"
    Case 18
        Result = CountHeading1
        Expected = oneBasedExpectedArray(18)
        If Result = Expected Then GetPassFail = "PASS"
    Case 19
        Result = CountEmptyParasAfterH2
        Expected = oneBasedExpectedArray(19)
        If Result = Expected Then GetPassFail = "PASS"
    Case 20
        Result = CountFootnotesFollowedByDigit
        Expected = oneBasedExpectedArray(20)
        If Result = Expected Then GetPassFail = "PASS"
    Case 21
        Result = CountSpacesAfterFootnotes
        Expected = oneBasedExpectedArray(21)
        If Result = Expected Then GetPassFail = "PASS"
    Case 22
        Result = CountEmptyParagraphs
        Expected = oneBasedExpectedArray(22)
        If Result = Expected Then GetPassFail = "PASS"
    Case 23
        Result = CountDeleteEmptyParagraphsBeforeHeading2
        Expected = oneBasedExpectedArray(23)
        If Result = Expected Then GetPassFail = "PASS"
    Case 24
        Result = CountFootnoteReferences
        Expected = oneBasedExpectedArray(24)
        If Result = Expected Then GetPassFail = "PASS"
    Case 25
        Result = CountFootnoteReferenceColors
        Expected = oneBasedExpectedArray(25)
        If Result = Expected Then GetPassFail = "PASS"
   Case Else
        Debug.Print "The test number is outside the accepted range"
    End Select
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
        Debug.Print GetPassFail(1), "Copy ()", "Test = " & num, Result, Expected, "CountDoubleSpaces"
    Case 2
        Debug.Print GetPassFail(2), "Copy ()", "Test = " & num, Result, Expected, "CountDoubleSpacesInShapes"
    Case 3
        Debug.Print GetPassFail(3), "Copy ()", "Test = " & num, Result, Expected, "CountSpaceFollowedByCarriageReturn"
    Case 4
        Debug.Print GetPassFail(4), "Copy ()", "Test = " & num, Result, Expected, "CountDoubleTabs"
    Case 5
        Debug.Print GetPassFail(5), "Copy ()", "Test = " & num, Result, Expected, "CountWhiteSpaceAndCarriageReturn"
    Case 6
        Debug.Print GetPassFail(6), "Copy ()", "Test = " & num, Result, Expected, "CountQuadrupleParagraphMarks"
    Case 7
        Debug.Print GetPassFail(7), "Copy ()", "Test = " & num, Result, Expected, "CountStyleWithSpaceAndNumber"
    Case 8
        Debug.Print GetPassFail(8), "Copy ()", "Test = " & num, Result, Expected, "CountStyleWithNumberAndSpace"
    Case 9
        Debug.Print GetPassFail(9), "Copy ()", "Test = " & num, Result, Expected, "CountPeriodSpaceLeftParenthesis"
    Case 10
        Debug.Print GetPassFail(10), "Copy ()", "Test = " & num, Result, Expected, "CountNonBreakingSpaces"
    Case 11
        Debug.Print GetPassFail(11), "Copy ()", "Test = " & num, Result, Expected, "CountFindNumberDashNumber"
    Case 12
        Debug.Print GetPassFail(12), "Copy ()", "Test = " & num, Result, Expected, "CountNumberDashNumberInFootnotes"
    Case 13
        Debug.Print GetPassFail(13), "Copy ()", "Test = " & num, Result, Expected, "CountEmptyParasWithNoThemeColor"
    Case 14
        Debug.Print GetPassFail(14), "Copy ()", "Test = " & num, Result, Expected, "CountWhiteParagraphMarks"
    Case 15
        Debug.Print GetPassFail(15), "Copy ()", "Test = " & num, Result, Expected, "CountSectionsWithDifferentFirstPage"
    Case 16
        Debug.Print GetPassFail(16), "Copy ()", "Test = " & num, Result, Expected, "CountTotalParagraphs"
    Case 17
        Debug.Print GetPassFail(17), "Copy ()", "Test = " & num, Result, Expected, "CountRedFootnoteReferences"
    Case 18
        Debug.Print GetPassFail(18), "Copy ()", "Test = " & num, Result, Expected, "CountHeading1"
    Case 19
        Debug.Print GetPassFail(19), "Copy ()", "Test = " & num, Result, Expected, "CountEmptyParasAfterH2"
    Case 20
        Debug.Print GetPassFail(20), "Copy ()", "Test = " & num, Result, Expected, "CountFootnotesFollowedByDigit"
    Case 21
        Debug.Print GetPassFail(21), "Copy ()", "Test = " & num, Result, Expected, "CountSpacesAfterFootnotes"
    Case 22
        Debug.Print GetPassFail(22), "Copy ()", "Test = " & num, Result, Expected, "CountEmptyParagraphs"
    Case 23
        Debug.Print GetPassFail(23), "Copy ()", "Test = " & num, Result, Expected, "CountDeleteEmptyParagraphsBeforeHeading2"
    Case 24
        Debug.Print GetPassFail(24), "Copy ()", "Test = " & num, Result, Expected, "CountFootnoteReferences"
    Case 25
        Debug.Print GetPassFail(25), "Copy ()", "Test = " & num, Result, Expected, "CountFootnoteReferenceColors"
   Case Else
        Debug.Print "The test number is outside the accepted range"
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
    Dim doc As Document
    Dim para As paragraph
    Dim deletionCount As Long
    Dim emptyCount As Long
    Dim userResponse As VbMsgBoxResult
    
    ' Initialize variables
    Set doc = ActiveDocument
    deletionCount = 0
    emptyCount = 0
    
    ' Loop through each paragraph in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph style is Heading 2
        If para.style = "Heading 2" Then
            ' Check if the previous paragraph exists and is not a section break
            If Not para.Previous Is Nothing Then
                ' Ensure the previous paragraph is not a section break
                If para.Previous.Range.Information(wdActiveEndAdjustedPageNumber) = para.Range.Information(wdActiveEndAdjustedPageNumber) Then
                    ' Check if the previous paragraph is empty
                    If Len(Trim(para.Previous.Range.text)) = 1 Then ' Empty paragraph has a length of 1 (paragraph mark)
                        emptyCount = emptyCount + 1
                        ' Prompt the user before deletion
                        'userResponse = MsgBox("Empty paragraph found before Heading 2. Do you want to delete it?", vbYesNoCancel + vbQuestion, "Delete Empty Paragraph")
                        'If userResponse = vbYes Then
                        '    para.Previous.Range.Delete ' Delete the empty paragraph
                        '    deletionCount = deletionCount + 1
                        'ElseIf userResponse = vbCancel Then
                        '    Exit For
                        'End If
                    End If
                End If
            End If
            
            ' Stop after 3 deletions
            'If deletionCount >= 3 Then
            '    Exit For
            'End If
        End If
    Next para
    
    ' Display the number of empty paragraphs found and deleted
    'MsgBox "Number of empty paragraphs found: " & emptyCount & vbCrLf & "Number of empty paragraphs deleted: " & deletionCount
    CountDeleteEmptyParagraphsBeforeHeading2 = deletionCount
End Function

Private Function CountEmptyParagraphs() As Long
    Dim para As paragraph
    Dim count As Long
    count = 0
    For Each para In ActiveDocument.Paragraphs
        If Len(para.Range.text) = 1 And para.Range.text = vbCr Then
            count = count + 1
        End If
    Next para
    CountEmptyParagraphs = count
End Function

Private Function CountSpacesAfterFootnotes() As Integer
' Also shows Footnote References and Following Characters (ASCII Values)
    Dim doc As Document
    Dim footnote As footnote
    Dim rng As Range
    Dim spaceCount As Long
    Dim debugInfo As String
    Dim asciiValue As Integer
    
    ' Initialize variables
    Set doc = ActiveDocument
    spaceCount = 0
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
    CountSpacesAfterFootnotes = spaceCount
End Function

Private Function CountFootnotesFollowedByDigit() As Integer
    Dim doc As Document
    Dim footnote As footnote
    Dim rng As Range
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
    Dim nextPara As paragraph
    
    ' Initialize the count
    emptyParaCount = 0
    
    ' Loop through each paragraph in the document
    For Each para In ActiveDocument.Paragraphs
        ' Check if the paragraph style is Heading 2
        If para.style = ActiveDocument.Styles("Heading 2") Then
            ' Check the next paragraph
            Set nextPara = para.Range.Next(Unit:=wdParagraph, count:=1).Paragraphs(1)
            If Not nextPara Is Nothing Then
                ' Check if the next paragraph is empty
                If Len(Trim(nextPara.Range.text)) = 1 Then ' Adjusted to check for a single paragraph mark
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
    For Each para In ActiveDocument.Paragraphs
        ' Check if the paragraph style is Heading 1
        If para.style = ActiveDocument.Styles("Heading 1") Then
            heading1Count = heading1Count + 1
        End If
    Next para
    
    ' Print the count to the console
    'Debug.Print "Number of Heading 1 paragraphs: " & heading1Count
    CountHeading1 = heading1Count
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
    CountTotalParagraphs = ActiveDocument.Paragraphs.count
End Function

Private Function CountSectionsWithDifferentFirstPage() As Integer
    Dim doc As Document
    Dim sec As Section
    Dim count As Integer
    
    Application.ScreenUpdating = False
    
    Set doc = ActiveDocument
    count = 0
    
    ' Loop through each section in the document
    For Each sec In doc.Sections
        ' Check if the section does not have "Different First Page" selected
        If sec.PageSetup.DifferentFirstPageHeaderFooter Then
            count = count + 1
        End If
    Next sec
    
    Application.ScreenUpdating = True
    
    'MsgBox "Number of sections with 'Different First Page' selected: " & count
    CountSectionsWithDifferentFirstPage = count
End Function

Private Function CountWhiteParagraphMarks() As Integer
    Dim doc As Document
    Dim rng As Range
    Dim count As Integer
    Dim firstFound As Boolean
    
    Set doc = ActiveDocument
    count = 0
    firstFound = False
    
    ' Initialize the range to search the entire document
    Set rng = doc.Content
    
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
    Dim rng As Range
    Dim emptyParaCount As Integer
    Dim themeColorUsed As Boolean
    Dim totalParaCount As Integer
    Dim colorCounts As Object
    Dim colorKey As String
    
    Set colorCounts = CreateObject("Scripting.Dictionary")
    emptyParaCount = 0
    totalParaCount = ActiveDocument.Paragraphs.count
    
    For Each para In ActiveDocument.Paragraphs
        Set rng = para.Range
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
    Dim footnote As footnote
    Dim rng As Range
    Dim matchCount As Integer
    Dim i As Integer
    Dim char As String
    
    matchCount = 0
    
    For Each footnote In ActiveDocument.Footnotes
        Set rng = footnote.Range
        For i = 1 To Len(rng.text)
            char = Mid(rng.text, i, 1)
            
            ' Check for a digit followed by a dash followed by another digit
            If IsNumeric(char) Then
                If i < Len(rng.text) And Mid(rng.text, i + 1, 1) = "-" Then
                    If i + 1 < Len(rng.text) And IsNumeric(Mid(rng.text, i + 2, 1)) Then
                        matchCount = matchCount + 1
                        i = i + 2 ' Skip the next two characters as they are part of the match
                    End If
                End If
            End If
        Next i
        DoEvents
    Next footnote

    'Debug.Print "Number of 'number-dash-number' matches in footnotes: " & matchCount
    CountNumberDashNumberInFootnotes = matchCount

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in function CountNumberDashNumberInFootnotes of Class BibleClass"
    Debug.Print "!!! Error in Test 11 !!!"
    Debug.Print "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in function CountNumberDashNumberInFootnotes of Class BibleClass"
    Resume PROC_EXIT
End Function

Private Function CountFindNumberDashNumber() As Integer
    Dim rng As Range
    Dim searchPattern As String
    Dim userInterrupt As VbMsgBoxResult
    Dim count As Integer
    
    searchPattern = "[0-9]{1,}-[0-9]{1,}" ' Pattern to find number-dash-number
    count = 0

    Set rng = ActiveDocument.Content
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
            rng.End = ActiveDocument.Content.End
        Loop
    End With

    CountFindNumberDashNumber = count
End Function

Private Function CountNonBreakingSpaces() As Integer
    Dim rng As Range
    Dim count As Long

    Set rng = ActiveDocument.Content
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
    Dim rng As Range
    Dim count As Long
    
    Set rng = ActiveDocument.Content
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
    Dim rng As Range
    Dim count As Long
    Dim styleName As String

    styleName = "Verse marker"

    Set rng = ActiveDocument.Content
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
    Dim rng As Range
    Dim count As Long
    Dim styleName As String
    Dim firstFound As Boolean
    
    styleName = "cvmarker"

    Set rng = ActiveDocument.Content
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
    Dim rng As Range
    Dim count As Long
    
    Set rng = ActiveDocument.Content
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
    Dim rng As Range
    Dim count As Long

    Set rng = ActiveDocument.Content
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
    Dim rng As Range
    Dim count As Long

    count = 0
    Set doc = ActiveDocument
    Set rng = doc.Content

    Do While rng.Find.Execute(FindText:="^t^t", Forward:=True, Wrap:=wdFindStop)
        count = count + 1
        rng.Collapse Direction:=wdCollapseEnd ' Move the range to the end of the found text
    Loop

    CountDoubleTabs = count
End Function

Private Function CountSpaceFollowedByCarriageReturn() As Integer
    Dim doc As Document
    Dim rng As Range
    Dim count As Long

    count = 0
    Set doc = ActiveDocument
    Set rng = doc.Content

    Do While rng.Find.Execute(FindText:=" ^13", Forward:=True, Wrap:=wdFindStop)
        count = count + 1
        rng.Collapse Direction:=wdCollapseEnd ' Move the range to the end of the found text
    Loop

    CountSpaceFollowedByCarriageReturn = count
End Function

Private Function CountDoubleSpaces() As Integer
    Dim doc As Document
    Dim rng As Range
    Dim doubleSpaceCount As Long
    
    doubleSpaceCount = 0
    Set doc = ActiveDocument
    Set rng = doc.Content
    
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
    Dim shp As Shape
    Dim groupShp As Shape
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

Sub ProcessShape(ByVal shp As Shape, ByRef doubleSpaceCount As Long)
    Dim textRange As Range
    Dim shapeText As String
    
    ' Check if the shape has a text frame
    If shp.Type = msoGroup Then
        ' Process each shape within the group
        Dim groupShp As Shape
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

