Attribute VB_Name = "Module1"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Public Const wdThemeColorNone As Long = -1
Public Const wdPaperB5Jis As Integer = 11
Private Sections1Col As Integer
Private Sections2Col As Integer
Private SectionsOddPageBreaks As Integer
Private SectionsEvenPageBreaks As Integer
Private SectionsContinuousBreaks As Integer
Private SectionsNewPageBreaks As Integer

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
    For Each para In ActiveDocument.paragraphs
        If para.style = "Heading 1" Then
            If InStr(para.range.text, heading1Name) > 0 Then
                Debug.Print "Heading 1: " & para.range.text
                startProcessing = True
                heading1Found = True
            Else
                startProcessing = False
                heading1Found = False
            End If
        End If
        
        If startProcessing Then
            
            If Len(Trim(para.range.text)) = 1 Then   ' Skip empty paragraph
                GoTo EmptyPara
            End If

            If para.style = "Heading 2" Then
                Debug.Print "Heading 2: " & para.range.text
                heading2Found = True
            ElseIf heading2Found Then
                Debug.Print para.range.text
            End If
        End If
        
        If heading1Found And para.style = "Heading 1" And InStr(para.range.text, heading1Name) = 0 Then
            Exit For
        End If
EmptyPara:
    Next para
End Sub

Function IsParagraphEmpty(paragraph As range) As Boolean
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
    If targetIndex > 0 And targetIndex <= ActiveDocument.paragraphs.count Then
        paraIndex = 1
        For Each para In ActiveDocument.paragraphs
            If paraIndex = targetIndex Then
                para.range.Select
                Exit Sub
            End If
            paraIndex = paraIndex + 1
        Next para
    Else
        MsgBox "Invalid index entered. Please enter a valid index between 1 and " & ActiveDocument.paragraphs.count & "."
    End If
End Sub

Function CountTextWrappingBreakParagraphs() As Long
    Dim para As paragraph
    Dim count As Long
    count = 0
    For Each para In ActiveDocument.paragraphs
        With para.range.Find
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
    For Each para In ActiveDocument.paragraphs
        If para.range.Sections.count > 0 Then
            If para.range.Sections(1).pageSetup.sectionStart = wdSectionNewPage Then
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
    For Each para In ActiveDocument.paragraphs
        If para.range.Sections.count > 0 Then
            If para.range.Sections(1).pageSetup.sectionStart = wdSectionContinuous Then
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
    For Each para In ActiveDocument.paragraphs
        If para.range.Sections.count > 0 Then
            If para.range.Sections(1).pageSetup.sectionStart = wdSectionEvenPage Then
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
    For Each para In ActiveDocument.paragraphs
        If para.range.Sections.count > 0 Then
            If para.range.Sections(1).pageSetup.sectionStart = wdSectionOddPage Then
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
    For Each para In doc.paragraphs
        ' Check if the paragraph contains only a page break or continuous page break
        If para.range.text = Chr(12) Or para.range.text = Chr(14) Then
            count = count + 1
            If Not foundFirst Then
                firstOccurrenceIndex = para.range.Start
                foundFirst = True
            End If
        End If
    Next para
    
    ' Print the count and first occurrence index to the console
    Debug.Print "Count of paragraphs with only a page break and continuous page break: " & count
    Debug.Print "Index of the first occurrence: " & firstOccurrenceIndex
    
    ' Go to the first result in the document
    If firstOccurrenceIndex <> -1 Then
        doc.range(firstOccurrenceIndex, firstOccurrenceIndex).Select
    End If
End Sub

Sub CountEmptyParagraphsWithAutomaticFont()
    Dim doc As Document
    Dim para As paragraph
    Dim count As Integer
    
    Set doc = ActiveDocument
    count = 0
    
    ' Loop through all paragraphs in the document
    For Each para In doc.paragraphs
        ' Check if the paragraph is empty and has the font set to automatic
        If Len(para.range.text) = 1 And para.range.font.color = wdColorAutomatic Then
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
    For Each para In doc.paragraphs
        count = count + 1
        ' Check if the current paragraph is the one we want to go to
        If count = paragraphNumber Then
            ' Select the paragraph
            para.range.Select
            Exit Sub
        End If
    Next para
    
    ' If the paragraph number is out of range, print a message to the console
    Debug.Print "Paragraph number " & paragraphNumber & " is out of range."
End Sub
 
Sub DetectFontColors()
    Dim para As paragraph
    Dim rng As range
    Dim colorUsed As Boolean
    Dim themeColorUsed As Boolean
    Dim paraCount As Integer
    
    paraCount = 0
    
    For Each para In ActiveDocument.paragraphs
        paraCount = paraCount + 1
        Set rng = para.range
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
    Dim rng As range
    Dim storyRange As range
    
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
    Dim rng As range
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
    Dim rng As range
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
    For Each para In doc.paragraphs
        ' Check if the paragraph style is Footnote Reference
        If para.style = "Footnote Reference" Then
            count = count + 1
            Set rng = para.range
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
    Dim footerRange As range
    
    ' Set the document
    Set doc = ActiveDocument
    
    ' Get the range of the footer on the first page
    Set footerRange = doc.Sections(1).Footers(wdHeaderFooterPrimary).range
    
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
    IsEmptyParagraph = (Len(p.range.text) = 1 And p.range.text = vbCr)
End Function

Sub CountTotallyEmptyParagraphs()
    Dim doc As Document
    Dim para As paragraph
    Dim sec As section
    Dim hdr As HeaderFooter
    Dim ftr As HeaderFooter
    Dim footnote As footnote
    Dim shp As shape
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
    For Each para In doc.paragraphs
        If IsEmptyParagraph(para) Then
            emptyParaCount = emptyParaCount + 1
        End If
    Next para
    
    ' Count empty paragraphs in headers
    emptyParaCountHeaders = 0
    For Each sec In doc.Sections
        For Each hdr In sec.Headers
            For Each para In hdr.range.paragraphs
                If IsEmptyParagraph(para) Then
                    emptyParaCountHeaders = emptyParaCountHeaders + 1
                    para.range.Select
                    Debug.Print "Found an empty paragraph in a Header. Stopping at this location."
                    Exit Sub
                End If
            Next para
        Next hdr
    Next sec
    
    ' Count empty paragraphs in footers and print page number to console
    emptyParaCountFooters = 0
    For Each sec In doc.Sections
        For Each ftr In sec.Footers
            For Each para In ftr.range.paragraphs
                ' Work around as IsEmptyParagraph does not work on first page with space as footer
                If Not FirstPageFooterNotEmpty Then
                    emptyParaCountFooters = emptyParaCountFooters + 1
                    ' Print page number to console
                    pageNum = para.range.Information(wdActiveEndPageNumber)
                    Debug.Print "Empty paragraph found in footer on page: " & pageNum
                    ' Stop at the location of the first empty paragraph found in a footer
                    para.range.Select
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
        For Each para In footnote.range.paragraphs
            If IsEmptyParagraph(para) Then
                emptyParaCountFootnotes = emptyParaCountFootnotes + 1
            End If
        Next para
    Next footnote
    
    ' Count empty paragraphs in text boxes and stop at the first one found
    emptyParaCountTextBoxes = 0
    For Each shp In doc.Shapes
        If shp.Type = msoTextBox Then
            For Each para In shp.TextFrame.textRange.paragraphs
                If IsEmptyParagraph(para) Then
                    emptyParaCountTextBoxes = emptyParaCountTextBoxes + 1
                    ' Stop at the location of the first empty paragraph found in a text box
                    para.range.Select
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

Sub CountTypesTrulyEmptyParagraph()
    Dim para As paragraph
    Dim paraText As String
    Dim paraRange As range
    Dim sectionBreakFound As Boolean
    Dim nextChar As String
    
    Dim trulyEmptyCount As Long
    Dim sectionFormattedEmptyCount As Long
    Dim firstFound As Boolean

    trulyEmptyCount = 0
    sectionFormattedEmptyCount = 0
    firstFound = False

    For Each para In ActiveDocument.paragraphs
        Set paraRange = para.range
        paraText = Trim(paraRange.text)

        ' Remove final paragraph mark if present
        If Len(paraText) > 0 Then
            If Right(paraText, 1) = Chr(13) Or Right(paraText, 1) = Chr(11) Then
                paraText = Left(paraText, Len(paraText) - 1)
            End If
        End If

        ' Only process if the paragraph is effectively empty
        If paraText = "" Then
            sectionBreakFound = False

            ' Check for section break characters in the paragraph
            If InStr(paraRange.text, Chr(12)) > 0 Then
                sectionBreakFound = True
            End If

            ' Check if the paragraph spans multiple sections
            If paraRange.Sections.count > 1 Then
                sectionBreakFound = True
            End If

            ' Check the next character only if not at end of doc
            If paraRange.End < ActiveDocument.content.End Then
                nextChar = paraRange.Next(Unit:=wdCharacter, count:=1).text
                If nextChar = Chr(12) Then
                    sectionBreakFound = True
                End If
            End If

            ' Count accordingly
            If sectionBreakFound Then
                sectionFormattedEmptyCount = sectionFormattedEmptyCount + 1
            Else
                trulyEmptyCount = trulyEmptyCount + 1

                If Not firstFound Then
                    paraRange.Select
                    firstFound = True
                End If
            End If
        End If
    Next para

    ' Final message
    'MsgBox "Empty paragraph counts:" & vbCrLf & _
    '       "- Truly empty (no section formatting): " & trulyEmptyCount & vbCrLf & _
    '       "- Empty with section formatting: " & sectionFormattedEmptyCount, vbInformation, "Paragraph Summary"
    Debug.Print "Empty paragraph counts:" & vbCrLf & _
           "- Truly empty (no section formatting): " & trulyEmptyCount & vbCrLf & _
           "- Empty with section formatting: " & sectionFormattedEmptyCount

    If Not firstFound Then
        MsgBox "No truly empty paragraph found to select.", vbExclamation
    End If
End Sub

Sub FindSpecificFontOutsideMainBody()
    Dim targetFont As String
    targetFont = "Gentium" ' <-- change this to the font you want to find

    Dim storyRange As range
    Dim para As paragraph
    Dim fontName As String

    Application.ScreenUpdating = False
    Application.StatusBar = "Searching for font: " & targetFont

    For Each storyRange In ActiveDocument.StoryRanges
        If storyRange.StoryType <> wdMainTextStory Then
            Do
                For Each para In storyRange.paragraphs
                    fontName = para.range.font.name
                    If StrComp(fontName, targetFont, vbTextCompare) = 0 Then
                        para.range.Select
                        Application.StatusBar = False
                        Application.ScreenUpdating = True
                        MsgBox "Found font '" & targetFont & "' in non-main content.", vbInformation
                        Exit Sub
                    End If
                Next para
                Set storyRange = storyRange.NextStoryRange
            Loop While Not storyRange Is Nothing
        End If
    Next storyRange

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Font '" & targetFont & "' not found outside the main text.", vbInformation
End Sub

Sub CreateTemplateWithoutText()
    Dim doc As Document
    Dim templateDoc As Document
    Dim templatePath As String
    
    ' Get the current active document
    Set doc = ActiveDocument
    
    ' Create a new blank document based on styles and configurations from the original document
    Set templateDoc = Documents.Add
    
    ' Copy the entire content from the original document (styles and configurations)
    doc.range.Copy
    
    ' Paste the copied content into the new document (keeping the formatting)
    templateDoc.range.PasteAndFormat wdFormatOriginalFormatting
    
    ' Remove all the text from the new document (leave the styles and formatting)
    templateDoc.content.Delete
    
    ' Specify the path to save the template as a macro-enabled template
    templatePath = "C:\adaept\aeBibleClass\BibleTemplate.dotm" ' Modify this path as needed
    
    ' Save as a Word macro-enabled template (.dotm) using the numeric value 13
    ' 13 = wdFormatTemplateMacroEnabled (macro-enabled template)
    templateDoc.SaveAs2 fileName:=templatePath, FileFormat:=13 ' 13 = wdFormatTemplateMacroEnabled
    
    ' Close the new template document
    templateDoc.Close
    
    ' Notify the user
    MsgBox "Template without text has been saved successfully as: " & templatePath
End Sub

Function GetVerticalAlignmentName(valign As WdVerticalAlignment) As String
    Select Case valign
        Case wdAlignVerticalTop: GetVerticalAlignmentName = "Top"
        Case wdAlignVerticalCenter: GetVerticalAlignmentName = "Center"
        Case wdAlignVerticalJustify: GetVerticalAlignmentName = "Justify"
        Case wdAlignVerticalBottom: GetVerticalAlignmentName = "Bottom"
        Case Else: GetVerticalAlignmentName = "Unknown"
    End Select
End Function

Sub PrintCompactSectionLayoutInfo()
    Dim sec As section
    Dim i As Long
    Dim nOneCol As Long, nTwoCol As Long
    Dim nEvenPageBreak As Long, nOddPageBreak As Long
    Dim nContinuousBreak As Long, nNewPageBreak As Long
    Dim outputFile As String
    Dim outputText As String
    outputFile = "C:\adaept\aeBibleClass\DocumentLayoutReport.txt"  ' Change to your desired path

    ' Open the text file to write
    Open outputFile For Output As #1
    
    ' Write Header to the file
    outputText = "=== Layout Report ===" & vbCrLf
    outputText = outputText & "Doc: " & ActiveDocument.name & vbCrLf
    outputText = outputText & "Total Sections: " & ActiveDocument.Sections.count & vbCrLf & vbCrLf
    Print #1, outputText
    
    For i = 1 To ActiveDocument.Sections.count
        Set sec = ActiveDocument.Sections(i)
        
        outputText = "Section " & i & ": " & vbCrLf
        outputText = outputText & "Page: " & IIf(sec.pageSetup.orientation = wdOrientPortrait, "Portrait", "Landscape") & ", " & _
                    "Size: " & GetPaperSizeName(sec.pageSetup.paperSize) & ", " & _
                    "Columns: " & sec.pageSetup.TextColumns.count & vbCrLf
        If sec.pageSetup.TextColumns.count > 1 Then nTwoCol = nTwoCol + 1 Else nOneCol = nOneCol + 1
        
        ' Margins
        outputText = outputText & "Margins (inches): " & _
                    "Top: " & PointsToInches(sec.pageSetup.topMargin) & ", " & _
                    "Bottom: " & PointsToInches(sec.pageSetup.bottomMargin) & ", " & _
                    "Left: " & PointsToInches(sec.pageSetup.leftMargin) & ", " & _
                    "Right: " & PointsToInches(sec.pageSetup.rightMargin) & ", " & _
                    "Gutter: " & PointsToInches(sec.pageSetup.gutter) & vbCrLf
        
        ' Line Numbering
        If sec.pageSetup.LineNumbering.Active Then
            outputText = outputText & "Line Numbers: " & sec.pageSetup.LineNumbering.StartingNumber & ", " & _
                        "Increment: " & sec.pageSetup.LineNumbering.CountBy & vbCrLf
        End If
        
        ' Header/Footer settings
        outputText = outputText & "Header Distance: " & PointsToInches(sec.pageSetup.HeaderDistance) & ", " & _
                    "Footer Distance: " & PointsToInches(sec.pageSetup.FooterDistance) & vbCrLf
        
        ' Borders (if any)
        outputText = outputText & "Borders: " & _
                    "Top: " & GetBorderStyle(sec.Borders(wdBorderTop)) & ", " & _
                    "Bottom: " & GetBorderStyle(sec.Borders(wdBorderBottom)) & ", " & _
                    "Left: " & GetBorderStyle(sec.Borders(wdBorderLeft)) & ", " & _
                    "Right: " & GetBorderStyle(sec.Borders(wdBorderRight)) & vbCrLf
        
        ' Section Break Type
        Select Case sec.pageSetup.sectionStart
            Case wdSectionNewPage
                outputText = outputText & "Section Break: New Page" & vbCrLf
                nNewPageBreak = nNewPageBreak + 1
            Case wdSectionOddPage
                outputText = outputText & "Section Break: Odd Page" & vbCrLf
                nOddPageBreak = nOddPageBreak + 1
            Case wdSectionEvenPage
                outputText = outputText & "Section Break: Even Page" & vbCrLf
                nEvenPageBreak = nEvenPageBreak + 1
            Case wdSectionContinuous
                outputText = outputText & "Section Break: Continuous" & vbCrLf
                nContinuousBreak = nContinuousBreak + 1
            Case Else
                outputText = outputText & "Section Break: None" & vbCrLf
        End Select
        
        ' Write section data to file
        Print #1, outputText
        outputText = "" ' Reset outputText for the next section
    Next i
    
    ' Summary of Sections
    outputText = "Summary: " & vbCrLf
    outputText = outputText & "Sections with 1 Column: " & nOneCol & vbCrLf
    Sections1Col = nOneCol
    Debug.Print "Sections1Col = " & nOneCol
    outputText = outputText & "Sections with 2 Columns: " & nTwoCol & vbCrLf
    Sections2Col = nTwoCol
    Debug.Print "Sections2Col = " & nTwoCol
    outputText = outputText & "Sections with Odd Page Breaks: " & nOddPageBreak & vbCrLf
    SectionsOddPageBreaks = nOddPageBreak
    Debug.Print "SectionsOddPageBreaks = " & nOddPageBreak
    outputText = outputText & "Sections with Even Page Breaks: " & nEvenPageBreak & vbCrLf
    SectionsEvenPageBreaks = nEvenPageBreak
    Debug.Print "SectionsEvenPageBreaks = " & nEvenPageBreak
    outputText = outputText & "Sections with Continuous Breaks: " & nContinuousBreak & vbCrLf
    SectionsContinuousBreaks = nContinuousBreak
    Debug.Print "SectionsContinuousBreaks = " & nContinuousBreak
    outputText = outputText & "Sections with New Page Breaks: " & nNewPageBreak & vbCrLf
    SectionsNewPageBreaks = nNewPageBreak
    Debug.Print "SectionsNewPageBreaks = " & nNewPageBreak
    Print #1, outputText
    
    ' Close the file
    Close #1

    'MsgBox "Layout report saved to: " & outputFile, vbInformation
End Sub

Function PointsToInches(Points As Single) As String
    PointsToInches = Format(Points / 72, "0.00")
End Function

Function GetPaperSizeName(paperSizeValue As WdPaperSize) As String
    Select Case paperSizeValue
        Case wdPaperA4: GetPaperSizeName = "A4"
        Case wdPaperLetter: GetPaperSizeName = "Letter"
        Case wdPaperLegal: GetPaperSizeName = "Legal"
        Case 11: GetPaperSizeName = "B5 (JIS)" 'wdPaperB5Jis
        Case Else: GetPaperSizeName = "Other (" & paperSizeValue & ")"
    End Select
End Function

Function GetBorderStyle(border As border) As String
    If border.LineStyle = wdLineStyleNone Then
        GetBorderStyle = "None"
    Else
        GetBorderStyle = border.LineStyle & ", Color: " & border.color
    End If
End Function

Sub CountTabParagraphsFull()
    Dim doc As Document
    Dim sec As section
    Dim hdr As HeaderFooter
    Dim ftr As HeaderFooter
    Dim para As paragraph
    Dim rng As range
    Dim bodyCount As Long
    Dim headerCount As Long
    Dim footerCount As Long
    Dim grandTotal As Long
    
    Set doc = ActiveDocument
    bodyCount = 0
    headerCount = 0
    footerCount = 0

    ' Count in main document body
    For Each para In doc.paragraphs
        Set rng = para.range
        rng.End = rng.End - 1 ' Exclude the final paragraph mark
        If rng.text = vbTab Then
            bodyCount = bodyCount + 1
        End If
    Next para

    ' Count in headers and footers across all sections
    For Each sec In doc.Sections
        For Each hdr In sec.Headers
            For Each para In hdr.range.paragraphs
                Set rng = para.range
                rng.End = rng.End - 1
                If rng.text = vbTab Then
                    headerCount = headerCount + 1
                End If
            Next para
        Next hdr
        
        For Each ftr In sec.Footers
            For Each para In ftr.range.paragraphs
                Set rng = para.range
                rng.End = rng.End - 1
                If rng.text = vbTab Then
                    footerCount = footerCount + 1
                End If
            Next para
        Next ftr
    Next sec

    ' Calculate grand total
    grandTotal = bodyCount + headerCount + footerCount

    ' Display results
    MsgBox "Tab-Only Paragraphs Count:" & vbCrLf & _
           "Document Body: " & bodyCount & vbCrLf & _
           "Headers: " & headerCount & vbCrLf & _
           "Footers: " & footerCount & vbCrLf & _
           "Grand Total: " & grandTotal, _
           vbInformation, "Tab Paragraph Count"
End Sub

