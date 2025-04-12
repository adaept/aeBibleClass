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
        Debug.Print "Color: " & .color
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
            
            If Len(Trim(para.Range.text)) = 1 Then   ' Skip empty paragraph
                GoTo EmptyPara
            End If

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
EmptyPara:
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

Function CountEmptyParagraphs() As Long
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

Sub CountEmptyParagraphsWithAutomaticFont()
    Dim doc As Document
    Dim para As paragraph
    Dim count As Integer
    
    Set doc = ActiveDocument
    count = 0
    
    ' Loop through all paragraphs in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph is empty and has the font set to automatic
        If Len(para.Range.text) = 1 And para.Range.font.color = wdColorAutomatic Then
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
        
        If rng.font.color <> wdColorAutomatic Then
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
                .font.color = wdColorBlack
                .Replacement.ClearFormatting
                .Replacement.font.color = wdColorAutomatic
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
        r = (rng.font.color And &HFF)
        g = (rng.font.color \ &H100 And &HFF)
        b = (rng.font.color \ &H10000 And &HFF)
        
        ' Compare the RGB values directly
        If r = oldR And g = oldG And b = oldB Then
            rng.font.color = newColor
        End If
    Next rng
End Sub

Sub ChangeSpecificColor()
    'Call ChangeFontColorRGB(255, 0, 1, 255, 0, 0)
    Call ChangeFontColorRGB(37, 37, 37, 0, 0, 0)
End Sub

Sub EnsureFootnoteReferenceStyleColor()
    Dim doc As Document
    Dim para As paragraph
    Dim rng As Range
    Dim hexColor As String
    Dim rgbColor As Long
    Dim count As Integer
    
    ' Set the desired hex color (e.g., purple: #663399)
    hexColor = "#663399"
    
    ' Convert hex color to RGB
    rgbColor = HexToRGB(hexColor)
    
    ' Initialize variables
    Set doc = ActiveDocument
    
    count = 0
    ' Loop through each paragraph in the document
    For Each para In doc.Paragraphs
        ' Check if the paragraph style is Footnote Reference
        If para.style = "Footnote Reference" Then
            count = count + 1
            Set rng = para.Range
            ' Check if the style color is correctly set to the desired color
            If rng.font.color <> rgbColor Then
                ' Set the style color to the desired color
                rng.font.color = rgbColor
            End If
        End If
    Next para
    
    ' Display a message indicating the process is complete
    'MsgBox "Footnote Reference styles checked and updated to the desired color where necessary."
    Debug.Print "Count of Footnote Reference = " & count
End Sub

Function HexToRGB(hexColor As String) As Long
    Dim r As Long, g As Long, b As Long
    
    ' Remove the "#" character if present
    hexColor = Replace(hexColor, "#", "")
    
    ' Convert hex to RGB components
    r = CLng("&H" & Mid(hexColor, 1, 2))
    g = CLng("&H" & Mid(hexColor, 3, 2))
    b = CLng("&H" & Mid(hexColor, 5, 2))
    
    ' Combine RGB components into a single Long value
    HexToRGB = RGB(r, g, b)
End Function

Sub ReapplyFootnoteReferenceStyle()
    Dim footnote As footnote
    Dim doc As Document
    
    Set doc = ActiveDocument
    
    ' Loop through all footnotes in the document
    For Each footnote In doc.Footnotes
        ' Apply the Footnote Reference style to each footnote reference
        footnote.Reference.style = doc.Styles("Footnote Reference")
    Next footnote
End Sub

Function FirstPageFooterNotEmpty() As Boolean
    Dim doc As Document
    Dim footerRange As Range
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Get the range of the footer on the first page
    Set footerRange = doc.Sections(1).Footers(wdHeaderFooterPrimary).Range
    
    ' Check if the footer is not empty
    If Len(Trim(footerRange.text)) > 0 Then
        'MsgBox "The footer on the first page is not empty."
        FirstPageFooterNotEmpty = True
    Else
        'MsgBox "The footer on the first page is empty."
        FirstPageFooterNotEmpty = False
    End If
End Function

Function IsEmptyParagraph(p As paragraph) As Boolean
' Function to check if a paragraph is truly empty
    IsEmptyParagraph = (Len(p.Range.text) = 1 And p.Range.text = vbCr)
End Function

Sub CountTotallyEmptyParagraphs()
    Dim doc As Document
    Dim para As paragraph
    Dim sec As Section
    Dim hdr As HeaderFooter
    Dim ftr As HeaderFooter
    Dim footnote As footnote
    Dim shp As Shape
    Dim emptyParaCount As Long
    Dim emptyParaCountHeaders As Long
    Dim emptyParaCountFooters As Long
    Dim emptyParaCountFootnotes As Long
    Dim emptyParaCountTextBoxes As Long
    Dim grandTotal As Long
    Dim pageNum As Long
    
    Set doc = ActiveDocument
        
    ' Count empty paragraphs in the main document
    emptyParaCount = 0
    For Each para In doc.Paragraphs
        If IsEmptyParagraph(para) Then
            emptyParaCount = emptyParaCount + 1
        End If
    Next para
    
    ' Count empty paragraphs in headers
    emptyParaCountHeaders = 0
    For Each sec In doc.Sections
        For Each hdr In sec.Headers
            For Each para In hdr.Range.Paragraphs
                If IsEmptyParagraph(para) Then
                    emptyParaCountHeaders = emptyParaCountHeaders + 1
                End If
            Next para
        Next hdr
    Next sec
    
    ' Count empty paragraphs in footers and print page number to console
    emptyParaCountFooters = 0
    For Each sec In doc.Sections
        For Each ftr In sec.Footers
            For Each para In ftr.Range.Paragraphs
                ' Work around as IsEmptyParagraph does not work on first page with space as footer
                If Not FirstPageFooterNotEmpty Then
                    emptyParaCountFooters = emptyParaCountFooters + 1
                    ' Print page number to console
                    pageNum = para.Range.Information(wdActiveEndPageNumber)
                    Debug.Print "Empty paragraph found in footer on page: " & pageNum
                    ' Stop at the location of the first empty paragraph found in a footer
                    para.Range.Select
                    'MsgBox "Found an empty paragraph in a footer. Stopping at this location."
                    Debug.Print "Found an empty paragraph in a footer. Stopping at this location."
                    Exit Sub
                End If
            Next para
        Next ftr
    Next sec
    
    ' Count empty paragraphs in footnotes
    emptyParaCountFootnotes = 0
    For Each footnote In doc.Footnotes
        For Each para In footnote.Range.Paragraphs
            If IsEmptyParagraph(para) Then
                emptyParaCountFootnotes = emptyParaCountFootnotes + 1
            End If
        Next para
    Next footnote
    
    ' Count empty paragraphs in text boxes and stop at the first one found
    emptyParaCountTextBoxes = 0
    For Each shp In doc.Shapes
        If shp.Type = msoTextBox Then
            For Each para In shp.TextFrame.textRange.Paragraphs
                If IsEmptyParagraph(para) Then
                    emptyParaCountTextBoxes = emptyParaCountTextBoxes + 1
                    ' Stop at the location of the first empty paragraph found in a text box
                    para.Range.Select
                    'MsgBox "Found an empty paragraph in a text box. Stopping at this location."
                    Debug.Print "Found an empty paragraph in a text box. Stopping at this location."
                    Exit Sub
                End If
            Next para
        End If
    Next shp
    
    ' Calculate grand total
    grandTotal = emptyParaCount + emptyParaCountHeaders + emptyParaCountFooters + emptyParaCountFootnotes + emptyParaCountTextBoxes
    
    ' Display counts
    'MsgBox "Empty Paragraphs in Main Document: " & emptyParaCount & vbCrLf & _
    '       "Empty Paragraphs in Headers: " & emptyParaCountHeaders & vbCrLf & _
    '       "Empty Paragraphs in Footers: " & emptyParaCountFooters & vbCrLf & _
    '       "Empty Paragraphs in Footnotes: " & emptyParaCountFootnotes & vbCrLf & _
    '       "Empty Paragraphs in Text Boxes: " & emptyParaCountTextBoxes & vbCrLf & _
    '       "Grand Total: " & grandTotal
    Debug.Print "Empty Paragraphs in Main Document: " & emptyParaCount & vbCrLf & _
           "Empty Paragraphs in Headers: " & emptyParaCountHeaders & vbCrLf & _
           "Empty Paragraphs in Footers: " & emptyParaCountFooters & vbCrLf & _
           "Empty Paragraphs in Footnotes: " & emptyParaCountFootnotes & vbCrLf & _
           "Empty Paragraphs in Text Boxes: " & emptyParaCountTextBoxes & vbCrLf & _
           "Grand Total: " & grandTotal
End Sub

