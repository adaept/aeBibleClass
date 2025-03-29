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

Private Const BibleClassVERSION As String = "0.0.5"
Private Const BibleClassVERSION_DATE As String = "March 23, 2025"

Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100
'
' Define the word constant in this class module
Private Const wdNotThemeColor As Long = -1
' Default use is to not time all tests individually
Private Const bTimeAllTests As Boolean = True
Private Const bGoTo16 As Boolean = True
' Used to store Expected results
Private oneBasedArray As Variant

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
    '      RunTest 1  2  3  4  5  6  7  8  9  10 11 12 13 14 15  16  17 18 19 20
    values = Array(0, 0, 0, 0, 0, 0, 0, 7, 1, 2, 0, 0, 0, 0, -1, -1, 0, 66)
    
    ' Convert to a 1-based array
    oneBasedArray = ConvertToOneBasedArray(values)
    
    ' Output the values in the 1-based array
    'For i = 1 To UBound(oneBasedArray)
    '    Debug.Print "Index " & i & ": " & oneBasedArray(i)
    'Next i

    ' Access the array elements to write a string output of 15 expected results on one line
    For i = LBound(oneBasedArray) To 15
        outputString = outputString & " Test(" & i & ")" & oneBasedArray(i) & " "
        ' Trim the trailing space
        outputString = Trim(outputString)
    Next i
    Debug.Print outputString

    ' Next results
    outputString = ""
    For i = 16 To UBound(oneBasedArray)
        outputString = outputString & " Test(" & i & ")" & oneBasedArray(i) & " "
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

Private Function RunTest(num As Integer) As Boolean
    Dim startTime As Double
    Dim endTime As Double
    Dim runTime As Double
    
    ' Record the start timer for each test
    startTime = Timer
    
    On Error GoTo PROC_ERR
    Select Case num
    Case 1
        Debug.Print "???", "Copy ()", "Test = " & num, CountDoubleSpaces, 0, "CountDoubleSpaces"
    Case 2
        Debug.Print "???", "Copy ()", "Test = " & num, CountDoubleSpacesInShapes, 0, "CountDoubleSpacesInShapes"
    Case 3
        Debug.Print "???", "Copy ()", "Test = " & num, CountSpaceFollowedByCarriageReturn, 0, "CountSpaceFollowedByCarriageReturn"
    Case 4
        Debug.Print "???", "Copy ()", "Test = " & num, CountDoubleTabs, 0, "CountDoubleTabs"
    Case 5
        Debug.Print "???", "Copy ()", "Test = " & num, CountWhiteSpaceAndCarriageReturn, 0, "CountWhiteSpaceAndCarriageReturn"
    Case 6
        Debug.Print "???", "Copy ()", "Test = " & num, CountQuadrupleParagraphMarks, 0, "CountQuadrupleParagraphMarks"
    Case 7
        Debug.Print "???", "Copy ()", "Test = " & num, CountStyleWithSpaceAndNumber, 0, "CountStyleWithSpaceAndNumber"
    Case 8
        Debug.Print "???", "Copy ()", "Test = " & num, CountStyleWithNumberAndSpace, 0, "CountStyleWithNumberAndSpace"
    Case 9
        Debug.Print "???", "Copy ()", "Test = " & num, CountPeriodSpaceLeftParenthesis, 7, "CountPeriodSpaceLeftParenthesis"
    Case 10
        Debug.Print "???", "Copy ()", "Test = " & num, CountNonBreakingSpaces, 1, "CountNonBreakingSpaces"
    Case 11
        Debug.Print "???", "Copy ()", "Test = " & num, CountFindNumberDashNumber, 2, "CountFindNumberDashNumber"
    Case 12
        Debug.Print "???", "Copy ()", "Test = " & num, CountNumberDashNumberInFootnotes, 0, "CountNumberDashNumberInFootnotes"
    Case 13
        Debug.Print "???", "Copy ()", "Test = " & num, CountEmptyParasWithNoThemeColor, 0, "CountEmptyParasWithNoThemeColor"
    Case 14
        Debug.Print "???", "Copy ()", "Test = " & num, CountWhiteParagraphMarks, 0, "CountWhiteParagraphMarks"
    Case 15
        Debug.Print "???", "Copy ()", "Test = " & num, CountSectionsWithDifferentFirstPage, -1, "CountSectionsWithDifferentFirstPage"
    Case 16
        Debug.Print "???", "Copy ()", "Test = " & num, CountTotalParagraphs, -1, "CountTotalParagraphs"
    Case 17
        Debug.Print "???", "Copy ()", "Test = " & num, CountRedFootnoteReferences, 0, "CountRedFootnoteReferences"
    Case 18
        Debug.Print "???", "Copy ()", "Test = " & num, CountHeading1, 66, "CountHeading1"
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
        If footnote.Reference.font.Color = redColor Then
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
        .font.Color = RGB(255, 255, 255) ' RGB value for white
        .text = "^p"
        .Forward = True
        .Wrap = wdFindStop
        
        ' Loop through all occurrences
        Do While .Execute
            ' Check if the found paragraph mark has the white color
            If rng.font.Color = RGB(255, 255, 255) Then
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
    rng.Find.font.Color = wdColorWhite

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

