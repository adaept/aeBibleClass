Attribute VB_Name = "Module1"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Public Const wdThemeColorNone As Long = -1
Public Const wdPaperB5Jis As Integer = 11

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
        msg = msg & "Character " & i & ": " & mid(selectedText, i, 1) & " (ASCII: " & Asc(mid(selectedText, i, 1)) & ")" & vbCrLf
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
    For Each rng In ActiveDocument.words
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
    r = CLng("&H" & mid(hexColor, 1, 2))
    g = CLng("&H" & mid(hexColor, 3, 2))
    b = CLng("&H" & mid(hexColor, 5, 2))
    
    ' Combine RGB components into a single Long value
    HexToRGB = RGB(r, g, b)
End Function

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

Sub CompareHeading1sWithShowHideToggle()
    Dim showList As Collection, hideList As Collection
    Dim i As Long, j As Long, k As Long
    Dim jShow(1 To 66) As Integer, kHide(1 To 66) As Integer
    Dim showTrue As String, showFalse As String
    Dim maxPage As Long
    Dim headingText As String
    Dim pageRange As range
    Dim para As paragraph
    Dim originalShowAll As Boolean

    Set showList = New Collection
    Set hideList = New Collection
    maxPage = ActiveDocument.range.Information(wdNumberOfPagesInDocument)

    ' --- Preserve original Show/Hide state ---
    originalShowAll = ActiveWindow.View.ShowAll

    Debug.Print "=== Comparison of Heading 1s with Show/Hide ON vs OFF ==="

    ' --- Pass 1: Show/Hide ON ---
    ActiveWindow.View.ShowAll = True
    j = 0
    For i = 1 To maxPage
        Set pageRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=i)
        Set pageRange = pageRange.GoTo(What:=wdGoToBookmark, name:="\page")

        headingText = ""
        For Each para In pageRange.paragraphs
            If para.style = "Heading 1" Then
                headingText = Replace(para.range.text, vbCr, "")
                showList.Add headingText
                j = j + 1
                jShow(j) = j
                showTrue = showTrue & " " & jShow(j) & ">" & i & "," & headingText
                'Debug.Print " " & jShow(j) & ">" & i & "," & headingText,
                Exit For
            End If
        Next para
    Next i
    Debug.Print showTrue
    Debug.Print "length = " & Len(showTrue)

    ' --- Pass 2: Show/Hide OFF ---
    ActiveWindow.View.ShowAll = False
    k = 0
    For i = 1 To maxPage
        Set pageRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=i)
        Set pageRange = pageRange.GoTo(What:=wdGoToBookmark, name:="\page")

        headingText = ""
        For Each para In pageRange.paragraphs
            If para.style = "Heading 1" Then
                headingText = Replace(para.range.text, vbCr, "")
                hideList.Add headingText
                k = k + 1
                kHide(k) = k
                showFalse = showFalse & " " & kHide(k) & ">" & i & "," & headingText
                'Debug.Print " " & kHide(k) & ">" & i & "," & headingText,
                Exit For
            End If
        Next para
    Next i
    Debug.Print showFalse
    Debug.Print "length = " & Len(showFalse)

    ' --- Restore original Show/Hide state ---
    ActiveWindow.View.ShowAll = originalShowAll

    Debug.Print "! Comparison complete."
End Sub

Sub CountAndDiagnoseFootnoteFormatting()
    Dim doc As Document
    Dim i As Long
    Dim ref As range
    Dim fn As footnote
    Dim errCount As Long
    Dim totalChecked As Long
    Dim posReported As Boolean

    Set doc = ActiveDocument
    errCount = 0
    totalChecked = 0
    posReported = False

    Debug.Print "Checking Footnote References..."

    ' Check main doc footnote references
    For i = 1 To doc.Footnotes.count
        Set ref = doc.Footnotes(i).Reference
        totalChecked = totalChecked + 1

        If Not IsFootnoteRefFormattedCorrectly(ref) Then
            errCount = errCount + 1
            If Not posReported Then
                Debug.Print "First incorrect formatting at character position: " & ref.Start
                Debug.Print "Mismatch details:"
                Debug.Print " - Font Name: " & ref.font.name
                Debug.Print " - Font Size: " & ref.font.Size
                Debug.Print " - Font Color: " & ref.font.color
                Debug.Print " - Superscript: " & ref.font.Superscript
                posReported = True
            End If
        End If

        If i Mod 100 = 0 Then Debug.Print "Checked: " & i & " footnotes..."
    Next i

    ' Check footnote numbers inside footnote text
    For Each fn In doc.Footnotes
        Set ref = fn.range.paragraphs(1).range.words(1)
        totalChecked = totalChecked + 1

        If Not IsFootnoteRefFormattedCorrectly(ref) Then
            errCount = errCount + 1
            If Not posReported Then
                Debug.Print "First incorrect footnote text number formatting at: " & ref.Start
                Debug.Print "Mismatch details:"
                Debug.Print " - Font Name: " & ref.font.name
                Debug.Print " - Font Size: " & ref.font.Size
                Debug.Print " - Font Color: " & ref.font.color
                Debug.Print " - Superscript: " & ref.font.Superscript
                posReported = True
            End If
        End If
    Next fn

    Debug.Print "Total checked: " & totalChecked
    Debug.Print "Total incorrect: " & errCount
End Sub

Function IsFootnoteRefFormattedCorrectly(rng As range) As Boolean
    With rng.font
        IsFootnoteRefFormattedCorrectly = (.name = "Segoe UI" Or .name = "Segoe UI Bold") _
            And .Size = 8 _
            And .color = wdColorBlue _
            And .Superscript = True
    End With
End Function

Sub TestPageRangeEnd()
    Selection.GoTo What:=wdGoToPage, name:="70"
    Selection.MoveRight Unit:=wdCharacter, count:=1
    Dim pageEnd As Long
    pageEnd = Selection.Bookmarks("\Page").range.End
    Selection.GoTo What:=wdGoToPage, name:="71"
    Selection.MoveRight Unit:=wdCharacter, count:=1
    Dim page71Start As Long
    page71Start = Selection.Start
    Debug.Print "Page 70 ends at: " & pageEnd
    Debug.Print "Page 71 starts at: " & page71Start
End Sub

Sub AuditFontUsage_ParagraphsAndHeadersFooters()
    Dim para As paragraph
    Dim fontMap As Object
    Dim fName As String
    Dim keyVar As Variant
    Dim logBuffer As String
    Dim sec As section
    Dim hf As HeaderFooter
    Dim hfTypes As Variant
    Dim hfKind As Variant

    Set fontMap = CreateObject("Scripting.Dictionary")

    ' Scan body paragraphs
    For Each para In ActiveDocument.paragraphs
        fName = para.range.Characters(1).font.name
        If Not fontMap.Exists(fName) Then
            fontMap.Add fName, 1
        Else
            fontMap(fName) = fontMap(fName) + 1
        End If
    Next para

    ' Define header/footer types
    hfTypes = Array(wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages)

    ' Scan header/footer paragraphs
    For Each sec In ActiveDocument.Sections
        For Each hfKind In hfTypes
            Set hf = sec.Headers(hfKind)
            If hf.Exists Then
                For Each para In hf.range.paragraphs
                    fName = para.range.Characters(1).font.name
                    If Not fontMap.Exists(fName) Then
                        fontMap.Add fName, 1
                    Else
                        fontMap(fName) = fontMap(fName) + 1
                    End If
                Next para
            End If

            Set hf = sec.Footers(hfKind)
            If hf.Exists Then
                For Each para In hf.range.paragraphs
                    fName = para.range.Characters(1).font.name
                    If Not fontMap.Exists(fName) Then
                        fontMap.Add fName, 1
                    Else
                        fontMap(fName) = fontMap(fName) + 1
                    End If
                Next para
            End If
        Next hfKind
    Next sec

    ' Output results
    logBuffer = "=== Font Usage Across Body, Headers, and Footers ===" & vbCrLf
    For Each keyVar In fontMap.Keys
        logBuffer = logBuffer & "- " & keyVar & ": " & fontMap(keyVar) & " paragraph(s)" & vbCrLf
    Next

    Debug.Print logBuffer
    MsgBox "Full font audit complete. See Immediate Window.", vbInformation
End Sub

Function CenturyRangeToYears(StartCentury As Long, EndCentury As Long) As String
    Dim sStart As Long, sEnd As Long
    Dim eStart As Long, eEnd As Long
    Dim Era As String

    ' AD (positive centuries)
    If StartCentury > 0 And EndCentury > 0 Then
        Era = "AD"

        ' Start of first century in range
        sStart = (StartCentury - 1) * 100 + 1
        ' End of last century in range
        eEnd = EndCentury * 100

        CenturyRangeToYears = CStr(sStart) & " to " & CStr(eEnd) & " Years " & Era
        Exit Function
    End If

    ' BC (negative centuries)
    If StartCentury < 0 And EndCentury < 0 Then
        Era = "BC"

        ' Convert to positive for math
        StartCentury = Abs(StartCentury)
        EndCentury = Abs(EndCentury)

        ' BC counts backward
        sStart = EndCentury * 100
        eEnd = (StartCentury - 1) * 100 + 1

        CenturyRangeToYears = CStr(sStart) & " to " & CStr(eEnd) & " Years " & Era
        Exit Function
    End If

    ' Mixed BC/AD ranges are historically invalid
    CenturyRangeToYears = ""
End Function

Public Sub CountLeftAligned(pageNumStart As Long, pageNumEnd As Long, ByRef someCount As Long)
    ' Passes someCount by reference, initiaize to -1 when calling and assign actual result at end of routine
    Dim pgRange As range
    Dim pageStart As Long, pageEnd As Long
    
    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNumStart))
    pageStart = pgRange.Start
    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNumEnd))
    pageEnd = pgRange.Start - 1
    Debug.Print "pageStart = " & pageStart
    Debug.Print "pageEnd   = " & pageEnd
    Debug.Print "someCount = " & someCount

    someCount = 40
    Debug.Print "someCount = " & someCount

    Dim i As Long
    i = pageStart
    
    

End Sub

'===============================================================================
' Procedure : CountNumericOrdinals
' Author    : Peter Ennis
'
' Purpose   :
'   Counts numeric ordinal abbreviations in the active Word document
'   (e.g., 1st, 2nd, 3rd, 4th, etc.), where the ordinal suffix is formatted
'   as superscript.
'
'   Only the suffix characters are counted:
'       st, nd, rd, th
'
'   The numeric portion itself is not counted.
'
' Design    :
'   This routine is optimized for very large documents (e.g., Bible or
'   scholarly texts) containing thousands of numeric references such as
'   chapter and verse numbers.
'
'   To ensure acceptable performance and prevent "Word Not Responding"
'   behavior, the code:
'     - Searches ONLY for superscripted ordinal suffixes (rare text)
'     - Avoids character-by-character scanning
'     - Avoids wildcard searches combined with formatting
'     - Uses Word's Find engine in its most efficient configuration
'
' Method    :
'   1. Iterates through the literal suffixes: "st", "nd", "rd", "th"
'   2. Uses Find with formatting enabled to locate superscripted matches
'   3. Verifies the immediately preceding character is a digit (0-9)
'   4. Tallies counts per suffix and maintains a grand total
'
' Review Mode:
'   Controlled by the Boolean variable "showReview" within the procedure.
'
'     True  - Each found ordinal is selected and displayed with a Yes/No
'             confirmation dialog for manual review.
'     False - Silent, high-performance counting only (recommended for
'             large documents).
'
' Output    :
'   Displays a summary message box reporting:
'     - Count of each ordinal suffix (st / nd / rd / th)
'     - Total number of numeric ordinals found
'
' Notes     :
'   - Superscript formatting is required for a match.
'   - Non-superscript ordinals are intentionally ignored.
'   - Roman numerals and spelled-out ordinals (e.g., "first") are not matched.
'   - The document is not modified.
'
'===============================================================================
Sub CountNumericOrdinals()

    Dim cntST As Long, cntND As Long, cntRD As Long, cntTH As Long
    Dim TOTAL As Long
    Dim showReview As Boolean

    ' ===== TOGGLE REVIEW MODE =====
    showReview = False
    ' ==============================

    Application.ScreenUpdating = False
    Application.StatusBar = "Counting numeric ordinals..."

    Dim suffixes As Variant
    suffixes = Array("st", "nd", "rd", "th")

    Dim doc As Document
    Set doc = ActiveDocument

    Dim i As Long
    For i = LBound(suffixes) To UBound(suffixes)

        Application.StatusBar = "Scanning superscript '" & suffixes(i) & "'..."
        DoEvents

        Dim rng As range
        Set rng = doc.content

        With rng.Find
            .ClearFormatting
            .text = suffixes(i)
            .MatchWildcards = False
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .font.Superscript = True   ' REQUIRED
        End With

        Do While rng.Find.Execute

            ' Check preceding character ONLY
            If rng.Start > doc.content.Start Then
                If doc.range(rng.Start - 1, rng.Start).text Like "[0-9]" Then

                    Select Case suffixes(i)
                        Case "st": cntST = cntST + 1
                        Case "nd": cntND = cntND + 1
                        Case "rd": cntRD = cntRD + 1
                        Case "th": cntTH = cntTH + 1
                    End Select

                    TOTAL = TOTAL + 1

                    If showReview Then
                        rng.Select
                        If MsgBox("Found ordinal: " & _
                                  doc.range(rng.Start - 1, rng.End).text & vbCrLf & _
                                  "Continue?", _
                                  vbYesNo + vbQuestion, _
                                  "Review Ordinal") = vbNo Then GoTo Done
                    End If
                End If
            End If

            rng.Collapse wdCollapseEnd
            DoEvents   ' keeps Word responsive
        Loop
    Next i

Done:
    Application.ScreenUpdating = True
    Application.StatusBar = False

    Debug.Print _
        "Numeric Ordinal Suffix Counts:" & vbCrLf & _
        "st: " & cntST & vbCrLf & _
        "nd: " & cntND & vbCrLf & _
        "rd: " & cntRD & vbCrLf & _
        "th: " & cntTH & vbCrLf & _
        "TOTAL: " & TOTAL & vbCrLf & _
        "Ordinal Count Results"

End Sub


