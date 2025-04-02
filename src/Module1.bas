Attribute VB_Name = "Module1"
Sub ViewCodeDetails()
    Dim selectedText As String
    Dim msg As String
    Dim i As Integer

    ' Get the selected text
    selectedText = Selection.text

    ' Initialize the message string
    msg = "Code details for the selected text:" & vbCrLf & vbCrLf

    ' Loop through each character in the selected text
    For i = 1 To Len(selectedText)
        msg = msg & "Character " & i & ": " & Mid(selectedText, i, 1) & " (ASCII: " & Asc(Mid(selectedText, i, 1)) & ")" & vbCrLf
    Next i

    ' Display the code details in a message box
    MsgBox msg
End Sub

Sub PrintFontProperties()
    Dim sel As Selection
    Set sel = Selection
    With sel.font
        Debug.Print "Name: " & .name
        Debug.Print "Size: " & .Size
        Debug.Print "Bold: " & .Bold
        Debug.Print "Italic: " & .Italic
        Debug.Print "Underline: " & .Underline
        Debug.Print "Color: " & .Color
        Debug.Print "StrikeThrough: " & .StrikeThrough
        Debug.Print "DoubleStrikeThrough: " & .DoubleStrikeThrough
        Debug.Print "Subscript: " & .Subscript
        Debug.Print "Superscript: " & .Superscript
        Debug.Print "Shadow: " & .Shadow
        Debug.Print "Outline: " & .Outline
        Debug.Print "Emboss: " & .Emboss
        Debug.Print "Engrave: " & .Engrave
        Debug.Print "AllCaps: " & .AllCaps
        Debug.Print "Hidden: " & .Hidden
        Debug.Print "SmallCaps: " & .SmallCaps
        Debug.Print "Kerning: " & .Kerning
        Debug.Print "Spacing: " & .Spacing
        Debug.Print "Scaling: " & .Scaling
        Debug.Print "Position: " & .Position
        Debug.Print "Ligatures: " & .Ligatures
        Debug.Print "NumberForm: " & .NumberForm
        Debug.Print "NumberSpacing: " & .NumberSpacing
        Debug.Print "StylisticSet: " & .StylisticSet
        Debug.Print "ContextualAlternates: " & .ContextualAlternates
    End With
End Sub

Sub PrintBibleBook()
    Dim heading1Name As String
    Dim para As paragraph
    Dim startProcessing As Boolean
    Dim heading1Found As Boolean
    Dim heading2Found As Boolean
    
    ' Prompt user to enter the name of Heading 1
    heading1Name = InputBox("Enter the name of Heading 1:")
    heading1Name = UCase(heading1Name)
    
    startProcessing = False
    heading1Found = False
    heading2Found = False
    
    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        If para.style = "Heading 1" Then
            If InStr(para.Range.text, heading1Name) > 0 Then
                Debug.Print "Heading 1: " & para.Range.text
                startProcessing = True
                heading1Found = True
            Else
                startProcessing = False
                heading1Found = False
            End If
        End If
        
        If startProcessing Then
            If para.style = "Heading 2" Then
                Debug.Print "Heading 2: " & para.Range.text
                heading2Found = True
            ElseIf heading2Found Then
                Debug.Print para.Range.text
            End If
        End If
        
        If heading1Found And para.style = "Heading 1" And InStr(para.Range.text, heading1Name) = 0 Then
            Exit For
        End If
    Next para
End Sub

Function IsParagraphEmpty(paragraph As Range) As Boolean
    ' Check if the paragraph is empty
    If Len(paragraph.text) = 1 And paragraph.text = vbCr Then
        IsParagraphEmpty = True
    Else
        IsParagraphEmpty = False
    End If
End Function

Sub GoToParagraphIndex()
    Dim para As paragraph
    Dim paraIndex As Integer
    Dim targetIndex As Integer
    
    ' Prompt user to enter the index of the paragraph
    targetIndex = InputBox("Enter the index of the paragraph you want to go to:")
    
    ' Validate the entered index
    If targetIndex > 0 And targetIndex <= ActiveDocument.Paragraphs.count Then
        paraIndex = 1
        For Each para In ActiveDocument.Paragraphs
            If paraIndex = targetIndex Then
                para.Range.Select
                Exit Sub
            End If
            paraIndex = paraIndex + 1
        Next para
    Else
        MsgBox "Invalid index entered. Please enter a valid index between 1 and " & ActiveDocument.Paragraphs.count & "."
    End If
End Sub

Sub CountParagraphs()
    Dim columnBreakParagraphs As Long
    Dim textWrappingBreakParagraphs As Long
    Dim nextPageSectionBreakParagraphs As Long
    Dim continuousSectionBreakParagraphs As Long
    Dim evenPageSectionBreakParagraphs As Long
    Dim oddPageSectionBreakParagraphs As Long
    Dim debugFile As String
    
    ' Set the debug file path to the current document directory
    debugFile = ActiveDocument.Path & "\DebugTestFile.txt"
    
    ' Delete the old debug file if it exists
    If Dir(debugFile) <> "" Then
        Kill debugFile
    End If
    
    ' Count paragraphs
    columnBreakParagraphs = CountColumnBreakParagraphs()
    textWrappingBreakParagraphs = CountTextWrappingBreakParagraphs()
    nextPageSectionBreakParagraphs = CountNextPageSectionBreakParagraphs()
    continuousSectionBreakParagraphs = CountContinuousSectionBreakParagraphs()
    evenPageSectionBreakParagraphs = CountEvenPageSectionBreakParagraphs()
    oddPageSectionBreakParagraphs = CountOddPageSectionBreakParagraphs()
    
    ' Print the counts to the console (Immediate Window)
    Debug.Print "Total Paragraphs: " & totalParagraphs
    Debug.Print "Empty Paragraphs: " & emptyParagraphs
    Debug.Print "Paragraphs with Column Break: " & columnBreakParagraphs
    Debug.Print "Paragraphs with Text Wrapping Break: " & textWrappingBreakParagraphs
    Debug.Print "Paragraphs with Section Break (Next Page): " & nextPageSectionBreakParagraphs
    Debug.Print "Paragraphs with Section Break (Continuous): " & continuousSectionBreakParagraphs
    Debug.Print "Paragraphs with Section Break (Even Page): " & evenPageSectionBreakParagraphs
    Debug.Print "Paragraphs with Section Break (Odd Page): " & oddPageSectionBreakParagraphs
    
    ' Append the final results to the debug file
    AppendToFile debugFile, "Total Paragraphs: " & totalParagraphs
    AppendToFile debugFile, "Empty Paragraphs: " & emptyParagraphs
    AppendToFile debugFile, "Paragraphs with Column Break: " & columnBreakParagraphs
    AppendToFile debugFile, "Paragraphs with Text Wrapping Break: " & textWrappingBreakParagraphs
    AppendToFile debugFile, "Paragraphs with Section Break (Next Page): " & nextPageSectionBreakParagraphs
    AppendToFile debugFile, "Paragraphs with Section Break (Continuous): " & continuousSectionBreakParagraphs
    AppendToFile debugFile, "Paragraphs with Section Break (Even Page): " & evenPageSectionBreakParagraphs
    AppendToFile debugFile, "Paragraphs with Section Break (Odd Page): " & oddPageSectionBreakParagraphs
End Sub

Function CountColumnBreakParagraphs() As Long
    Dim rng As Range
    Dim count As Long
    
    ' Initialize count
    count = 0
    
    ' Set the range to the entire document
    Set rng = ActiveDocument.Content
    
    ' Use the Find method to search for column breaks (^b)
    With rng.Find
        .ClearFormatting
        .text = "^b"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        ' Loop through all occurrences of column breaks
        Do While .Execute
            count = count + 1
        Loop
    End With
    
    CountColumnBreakParagraphs = count
End Function

Function CountColumnBreakParagraphsSlowDebug() As Long
    Dim para As paragraph
    Dim count As Long
    Dim rng As Range
    Dim debugFile As String
    Dim fileNum As Integer
    Dim paraIndex As Long
    
    ' Initialize count and paragraph index
    count = 0
    paraIndex = 0
    
    ' Set the debug file path to the current document directory
    debugFile = ActiveDocument.Path & "\DebugColumnBreaks.txt"
    
    ' Delete the old debug file if it exists
    If Dir(debugFile) <> "" Then
        Kill debugFile
    End If
    
    ' Open the debug file for writing
    fileNum = FreeFile
    Open debugFile For Output As fileNum
    
    ' Loop through each paragraph in the document
    For Each para In ActiveDocument.Paragraphs
        paraIndex = paraIndex + 1
        Set rng = para.Range
        With rng.Find
            .ClearFormatting
            .text = "^b"
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            If .Execute Then
                count = count + 1
                Print #fileNum, "Paragraph " & paraIndex & " contains a column break (^b)"
            End If
        End With
    Next para
    
    ' Close the debug file
    Close fileNum
    
    CountColumnBreakParagraphsSlowDebug = count
End Function

Function CountColumnBreakParagraphsOptimized() As Long
    Dim rng As Range
    Dim count As Long
    Dim debugFile As String
    Dim fileNum As Integer
    Dim paraText As String
    
    ' Initialize count
    count = 0
    
    ' Set the debug file path to the current document directory
    debugFile = ActiveDocument.Path & "\DebugColumnBreaksOptimized.txt"
    
    ' Delete the old debug file if it exists
    If Dir(debugFile) <> "" Then
        Kill debugFile
    End If
    
    ' Open the debug file for writing
    fileNum = FreeFile
    Open debugFile For Output As fileNum
    
    ' Set the range to the entire document
    Set rng = ActiveDocument.Content
    
    ' Use the Find method to search for column breaks (^b)
    With rng.Find
        .ClearFormatting
        .text = "^b"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        ' Loop through all occurrences of column breaks
        Do While .Execute
            count = count + 1
            paraText = Left(rng.text, Len(rng.text) - 1) ' Remove paragraph mark
            Print #fileNum, "Paragraph text: " & paraText & " contains a column break (^b)"
        Loop
    End With
    
    ' Close the debug file
    Close fileNum
    
    CountColumnBreakParagraphsOptimized = count
End Function

Sub CompareDebugFiles()
    Dim debugFile1 As String
    Dim debugFile2 As String
    Dim fileNum1 As Integer
    Dim fileNum2 As Integer
    Dim line1 As String
    Dim line2 As String
    Dim paraTexts1 As Collection
    Dim paraTexts2 As Collection
    Dim item As Variant
    Dim differences As String
    
    ' Set the debug file paths
    debugFile1 = ActiveDocument.Path & "\DebugColumnBreaks.txt"
    debugFile2 = ActiveDocument.Path & "\DebugColumnBreaksOptimized.txt"
    
    ' Initialize collections
    Set paraTexts1 = New Collection
    Set paraTexts2 = New Collection
    
    ' Read the first debug file and store paragraph texts
    fileNum1 = FreeFile
    Open debugFile1 For Input As fileNum1
    Do While Not EOF(fileNum1)
        Line Input #fileNum1, line1
        paraTexts1.Add Mid(line1, 16) ' Extract paragraph text
    Loop
    Close fileNum1
    
    ' Read the second debug file and store paragraph texts
    fileNum2 = FreeFile
    Open debugFile2 For Input As fileNum2
    Do While Not EOF(fileNum2)
        Line Input #fileNum2, line2
        paraTexts2.Add Mid(line2, 16) ' Extract paragraph text
    Loop
    Close fileNum2
    
    ' Compare the collections and find differences
    differences = "Differences found:" & vbCrLf
    For Each item In paraTexts1
        If Not IsInCollection(paraTexts2, item) Then
            differences = differences & "Paragraph text: " & item & " not found in optimized results" & vbCrLf
        End If
    Next item
    For Each item In paraTexts2
        If Not IsInCollection(paraTexts1, item) Then
            differences = differences & "Paragraph text: " & item & " not found in original results" & vbCrLf
        End If
    Next item
    
    ' Display the differences
    MsgBox differences, vbInformation, "Comparison Results"
End Sub

Function IsInCollection(col As Collection, value As Variant) As Boolean
    Dim item As Variant
    IsInCollection = False
    For Each item In col
        If item = value Then
            IsInCollection = True
            Exit Function
        End If
    Next item
End Function

Function CountTextWrappingBreakParagraphs() As Long
    Dim para As paragraph
    Dim count As Long
    count = 0
    For Each para In ActiveDocument.Paragraphs
        With para.Range.Find
            .ClearFormatting
            .text = "^m"
            If .Execute Then
                count = count + 1
            End If
        End With
    Next para
    CountTextWrappingBreakParagraphs = count
End Function

Function CountNextPageSectionBreakParagraphs() As Long
    Dim para As paragraph
    Dim count As Long
    count = 0
    For Each para In ActiveDocument.Paragraphs
        If para.Range.Sections.count > 0 Then
            If para.Range.Sections(1).PageSetup.SectionStart = wdSectionNewPage Then
                count = count + 1
            End If
        End If
    Next para
    CountNextPageSectionBreakParagraphs = count
End Function

Function CountContinuousSectionBreakParagraphs() As Long
    Dim para As paragraph
    Dim count As Long
    count = 0
    For Each para In ActiveDocument.Paragraphs
        If para.Range.Sections.count > 0 Then
            If para.Range.Sections(1).PageSetup.SectionStart = wdSectionContinuous Then
                count = count + 1
            End If
        End If
    Next para
    CountContinuousSectionBreakParagraphs = count
End Function

Function CountEvenPageSectionBreakParagraphs() As Long
    Dim para As paragraph
    Dim count As Long
    count = 0
    For Each para In ActiveDocument.Paragraphs
        If para.Range.Sections.count > 0 Then
            If para.Range.Sections(1).PageSetup.SectionStart = wdSectionEvenPage Then
                count = count + 1
            End If
        End If
    Next para
    CountEvenPageSectionBreakParagraphs = count
End Function

Function CountOddPageSectionBreakParagraphs() As Long
    Dim para As paragraph
    Dim count As Long
    count = 0
    For Each para In ActiveDocument.Paragraphs
        If para.Range.Sections.count > 0 Then
            If para.Range.Sections(1).PageSetup.SectionStart = wdSectionOddPage Then
                count = count + 1
            End If
        End If
    Next para
    CountOddPageSectionBreakParagraphs = count
End Function

Sub AppendToFile(filePath As String, text As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Append As fileNum
    Print #fileNum, text
    Close fileNum
End Sub

Sub SearchParagraphs()
    Dim doc As Document
    Dim para As paragraph
    Dim count As Integer
    Dim firstOccurrenceIndex As Integer
    Dim foundFirst As Boolean
    
    Set doc = ActiveDocument
    count = 0
    firstOccurrenceIndex = -1
    foundFirst = False
    
    ' Loop through all paragraphs in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph contains only a page break or continuous page break
        If para.Range.text = Chr(12) Or para.Range.text = Chr(14) Then
            count = count + 1
            If Not foundFirst Then
                firstOccurrenceIndex = para.Range.Start
                foundFirst = True
            End If
        End If
    Next para
    
    ' Print the count and first occurrence index to the console
    Debug.Print "Count of paragraphs with only a page break and continuous page break: " & count
    Debug.Print "Index of the first occurrence: " & firstOccurrenceIndex
    
    ' Go to the first result in the document
    If firstOccurrenceIndex <> -1 Then
        doc.Range(firstOccurrenceIndex, firstOccurrenceIndex).Select
    End If
End Sub

Sub CountEmptyParagraphsWithAutomaticFont()
    Dim doc As Document
    Dim para As paragraph
    Dim count As Integer
    
    Set doc = ActiveDocument
    count = 0
    
    ' Loop through all paragraphs in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph is empty and has the font set to automatic
        If Len(para.Range.text) = 1 And para.Range.font.Color = wdColorAutomatic Then
            count = count + 1
        End If
    Next para
    
    ' Print the count to the console
    Debug.Print "Count of empty paragraphs with font set to automatic: " & count
End Sub

Sub GoToParagraphByCount(paragraphNumber As Integer)
    Dim doc As Document
    Dim para As paragraph
    Dim count As Integer
    
    Set doc = ActiveDocument
    count = 0
    
    ' Loop through all paragraphs in the document
    For Each para In doc.Paragraphs
        count = count + 1
        ' Check if the current paragraph is the one we want to go to
        If count = paragraphNumber Then
            ' Select the paragraph
            para.Range.Select
            Exit Sub
        End If
    Next para
    
    ' If the paragraph number is out of range, print a message to the console
    Debug.Print "Paragraph number " & paragraphNumber & " is out of range."
End Sub
 
Sub DetectFontColors()
    Dim para As paragraph
    Dim rng As Range
    Dim colorUsed As Boolean
    Dim themeColorUsed As Boolean
    Dim paraCount As Integer
    
    paraCount = 0
    
    For Each para In ActiveDocument.Paragraphs
        paraCount = paraCount + 1
        Set rng = para.Range
        colorUsed = False
        themeColorUsed = False
        
        If rng.font.Color <> wdColorAutomatic Then
            colorUsed = True
        End If
        
        If rng.font.TextColor.ObjectThemeColor <> wdThemeColorNone Then
            themeColorUsed = True
        End If
        
        If colorUsed Or themeColorUsed Then
            Debug.Print "Paragraph number with font or theme color: " & paraCount
            Exit Sub
        End If
    Next para
End Sub

Sub UpdateEmptyParasToNoThemeColor()
' Set the Font.Color property to wdColorAutomatic, which effectively removes any theme color
    Dim para As paragraph
    Dim rng As Range
    Dim totalParaCount As Integer
    Dim updatedParaCount As Integer
    
    totalParaCount = ActiveDocument.Paragraphs.count
    updatedParaCount = 0
    
    For Each para In ActiveDocument.Paragraphs
        Set rng = para.Range
        
        ' Check if the paragraph is empty
        If Len(rng.text) = 1 Then ' Only the paragraph mark
            ' Check if the theme color is not wdColorAutomatic
            If rng.font.Color <> wdColorAutomatic Then
                ' Update the font color to wdColorAutomatic
                rng.font.Color = wdColorAutomatic
                updatedParaCount = updatedParaCount + 1
            End If
        End If
    Next para
    
    Debug.Print "Total number of paragraphs: " & totalParaCount
    Debug.Print "Number of empty paragraphs updated to no theme color: " & updatedParaCount
End Sub

Sub UpdateBlackToAutomatic()
    Dim doc As Document
    Dim rng As Range
    Dim storyRange As Range
    
    Set doc = ActiveDocument
    
    ' Turn off screen updating for better performance
    Application.ScreenUpdating = False
    
    ' Loop through each story in the document
    For Each storyRange In doc.StoryRanges
        Set rng = storyRange
        Do
            ' Loop through each character in the range
            With rng.Find
                .ClearFormatting
                .font.Color = wdColorBlack
                .Replacement.ClearFormatting
                .Replacement.font.Color = wdColorAutomatic
                .text = ""
                .Replacement.text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = True
            End With
            rng.Find.Execute Replace:=wdReplaceAll
            
            ' Move to the next linked story range
            Set rng = rng.NextStoryRange
        Loop Until rng Is Nothing
    Next storyRange
    
    ' Turn on screen updating
    Application.ScreenUpdating = True
    
    ' Display a message indicating completion
    MsgBox "All black font colors have been updated to automatic."
End Sub

Sub ChangeFontColorRGB(oldR As Long, oldG As Long, oldB As Long, newR As Long, newG As Long, newB As Long)
    Dim rng As Range
    Dim oldColor As Long
    Dim newColor As Long
    Dim r As Long, g As Long, b As Long
    
    ' Define the old and new colors using RGB values
    oldColor = RGB(oldR, oldG, oldB)
    newColor = RGB(newR, newG, newB)
    
    ' Loop through each word in the document
    For Each rng In ActiveDocument.Words
        ' Extract the RGB values of the current font color
        r = (rng.font.Color And &HFF)
        g = (rng.font.Color \ &H100 And &HFF)
        b = (rng.font.Color \ &H10000 And &HFF)
        
        ' Compare the RGB values directly
        If r = oldR And g = oldG And b = oldB Then
            rng.font.Color = newColor
        End If
    Next rng
End Sub

Sub ChangeSpecificColor()
    'Call ChangeFontColorRGB(255, 0, 1, 255, 0, 0)
    Call ChangeFontColorRGB(37, 37, 37, 0, 0, 0)
End Sub

Sub TestGetColorNameFromHex()
    Dim hexColor As String
    Dim colorName As String
    
    hexColor = "#FF0000" ' Example hex color
    colorName = GetColorNameFromHex(hexColor)
    
    'MsgBox "The color name for " & hexColor & " is " & colorName
    Debug.Print "The color name for " & hexColor & " is " & colorName
End Sub
