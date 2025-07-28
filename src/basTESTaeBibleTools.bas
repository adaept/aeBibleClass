Attribute VB_Name = "basTESTaeBibleTools"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Sub ListCustomXMLParts()
    Dim xmlPart As customXMLPart
    Dim i As Integer
    i = 1
    For Each xmlPart In ThisDocument.CustomXMLParts
        Debug.Print "Custom XML Part " & i & ": " & xmlPart.XML
        i = i + 1
    Next xmlPart
End Sub

Sub ListCustomXMLSchemas()
    Dim xmlPart As customXMLPart
    For Each xmlPart In ActiveDocument.CustomXMLParts
        Debug.Print xmlPart.NamespaceURI
    Next xmlPart
End Sub

Sub AddCustomUIXML()
    Dim xmlPart As customXMLPart
    Dim xmlContent As String
    ' Define XML structure
    xmlContent = "<?xml version='1.0' encoding='UTF-8'?>" & _
                 "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>" & _
                 "<ribbon><tabs><tab id='CustomTab' label='My Tab'></tab></tabs></ribbon>" & _
                 "</customUI>"
    ' Add XML part to document
    Set xmlPart = ActiveDocument.CustomXMLParts.Add(xmlContent)

    MsgBox "CustomUI XML added successfully!"
End Sub

Sub RemoveDuplicateCustomXMLParts()
    Dim xmlPart As customXMLPart
    Dim xmlParts As CustomXMLParts
    Dim essentialParts As Collection
    Dim duplicateParts As Collection
    Dim partName As String
    Dim i As Integer, j As Integer
    
    Set xmlParts = ActiveDocument.CustomXMLParts
    Set essentialParts = New Collection
    Set duplicateParts = New Collection
    
    ' Identify essential and duplicate parts
    For i = 1 To xmlParts.count
        partName = xmlParts(i).NamespaceURI
        If Not IsPartInCollection(essentialParts, partName) Then
            essentialParts.Add xmlParts(i), partName
        Else
            duplicateParts.Add xmlParts(i), partName
        End If
    Next i
    
    ' Remove duplicate parts
    For j = 1 To duplicateParts.count
        duplicateParts(j).Delete
    Next j
    
    ' Print names of essential and duplicate parts
    Debug.Print "Essential CustomXML Parts:"
    For i = 1 To essentialParts.count
        Debug.Print essentialParts(i).NamespaceURI
    Next i
    
    Debug.Print "Duplicate CustomXML Parts:"
    For j = 1 To duplicateParts.count
        Debug.Print duplicateParts(j).NamespaceURI
    Next j
End Sub

Function IsPartInCollection(col As Collection, partName As String) As Boolean
    Dim i As Integer
    IsPartInCollection = False
    For i = 1 To col.count
        If col(i).NamespaceURI = partName Then
            IsPartInCollection = True
            Exit Function
        End If
    Next i
End Function

Sub DeleteCustomUIXML()
    Dim xmlPart As customXMLPart
    Dim xmlParts As CustomXMLParts
    Dim i As Integer
    
    Set xmlParts = ActiveDocument.CustomXMLParts
    
    ' Loop through all CustomXMLParts to find and delete the customUI parts
    For i = xmlParts.count To 1 Step -1
        Set xmlPart = xmlParts(i)
        If xmlPart.NamespaceURI = "http://schemas.microsoft.com/office/2006/01/customui" Or _
                xmlPart.NamespaceURI = "http://schemas.microsoft.com/office/2009/07/customui" Then
            xmlPart.Delete
        End If
    Next i
    
    MsgBox "CustomUI XML parts deleted successfully!"
End Sub

Function GetColorNameFromHex(hexColor As String) As String
    Dim colorName As String
    
    ' Convert hex to uppercase for consistency
    hexColor = UCase(hexColor)
    
    ' Determine the color name based on the hex value
    Select Case hexColor
        Case "#FF0000"
            colorName = "Red"
        Case "#00FF00"
            colorName = "Green"
        Case "#006400"
            colorName = "Dark Green"
        Case "#50C878"
            colorName = "Emerald"
        Case "#0000FF"
            colorName = "Blue"
        Case "#FFD700"
            colorName = "Gold"
        Case "#FFA500"
            colorName = "Orange"
        Case "#663399"
            colorName = "Purple"
        Case "#FFFFFF"
            colorName = "White"
        Case "#000000"
            colorName = "Black"
        Case "#800000"
            colorName = "Dark Red"
        Case "#808080"
            colorName = "Gray"
        Case Else
            colorName = "Unknown Color"
    End Select
    
    ' Return the color name
    GetColorNameFromHex = colorName
End Function

Sub ListAndCountFontColors()
    Dim rng As range
    Dim colorDict As Object
    Dim colorKey As Variant
    Dim colorCount As Long
    Dim r As Long, g As Long, b As Long
    
    ' Create a dictionary to store color counts
    Set colorDict = CreateObject("Scripting.Dictionary")
    
    ' Loop through each word in the document
    For Each rng In ActiveDocument.Words
        ' Get the RGB values of the font color
        r = (rng.font.color And &HFF)
        g = (rng.font.color \ &H100 And &HFF)
        b = (rng.font.color \ &H10000 And &HFF)
        
        ' Create a key for the color in hex format
        colorKey = Right("0" & Hex(r), 2) & Right("0" & Hex(g), 2) & Right("0" & Hex(b), 2)
        
        ' Count the color occurrences
        If colorDict.Exists(colorKey) Then
            colorDict(colorKey) = colorDict(colorKey) + 1
        Else
            colorDict.Add colorKey, 1
        End If
    Next rng
    
    ' Print the results to the console
    For Each colorKey In colorDict.Keys
        colorCount = colorDict(colorKey)
        r = CLng("&H" & Left(colorKey, 2))
        g = CLng("&H" & Mid(colorKey, 3, 2))
        b = CLng("&H" & Right(colorKey, 2))
        
        Debug.Print "Color: RGB(" & r & ", " & g & ", " & b & ") - Hex: #" & colorKey & " - Count: " & colorCount & " - " & GetColorNameFromHex("#" & colorKey)
    Next colorKey
End Sub

Sub GetVerticalPositionOfCursorParagraph()
' Get the position of the para where the cursor is
    Dim doc As Document
    Dim rng As range
    Dim paraPos As Single
    
    Set doc = ActiveDocument
    Set rng = Selection.paragraphs(1).range
    
    ' Get the vertical position of the paragraph relative to the page
    paraPos = rng.Information(wdVerticalPositionRelativeToPage)
    
    ' Display the vertical position
    MsgBox "Vertical Position of the paragraph with the cursor: " & paraPos & " points"
End Sub

Sub FindFirstSectionWithDifferentFirstPage()
    Dim sec As section
    Dim i As Long

    For i = 1 To ActiveDocument.Sections.count
        Set sec = ActiveDocument.Sections(i)

        ' Check if Different First Page is enabled
        If sec.pageSetup.DifferentFirstPageHeaderFooter = True Then
            ' Select the header of the first page in this section
            sec.Headers(wdHeaderFooterFirstPage).range.Select

            MsgBox "Found in Section " & i & ": 'Different First Page' is enabled.", vbInformation
            Exit Sub
        End If
    Next i

    MsgBox "No sections with 'Different First Page' found.", vbInformation
End Sub

Sub FindFirstPageWithEmptyHeader()
    Dim sec As section
    Dim hdr As HeaderFooter
    Dim hdrText As String
    Dim i As Long
    Dim hdrType As Variant  ' Must be Variant for Array() to work

    For i = 1 To ActiveDocument.Sections.count
        Set sec = ActiveDocument.Sections(i)

        For Each hdrType In Array(wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages)
            Set hdr = sec.Headers(hdrType)

            If hdr.Exists And Not hdr.LinkToPrevious Then
                hdrText = Trim(hdr.range.text)

                If Right(hdrText, 1) = Chr(13) Then
                    hdrText = Left(hdrText, Len(hdrText) - 1)
                End If

                If hdrText = "" Then
                    hdr.range.Select
                    MsgBox "Found empty header in Section " & i & " (" & HeaderTypeName(hdrType) & ").", vbInformation
                    Exit Sub
                End If
            End If
        Next hdrType
    Next i

    MsgBox "No empty headers found.", vbInformation
End Sub

Function HeaderTypeName(hdrType As Variant) As String
    Select Case hdrType
        Case wdHeaderFooterPrimary: HeaderTypeName = "Primary"
        Case wdHeaderFooterFirstPage: HeaderTypeName = "First Page"
        Case wdHeaderFooterEvenPages: HeaderTypeName = "Even Pages"
        Case Else: HeaderTypeName = "Unknown"
    End Select
End Function

Sub OptimizedListFontsInDocument()
    Dim fontList As New Collection
    Dim doc As Document
    Dim para As paragraph
    Dim rng As range
    Dim fontName As String
    Dim i As Integer
    
    Set doc = ActiveDocument

    ' Loop through each paragraph in the document
    For Each para In doc.paragraphs
        Set rng = para.range
        fontName = rng.font.name
        On Error Resume Next
        ' Add unique fonts to the collection
        fontList.Add fontName, fontName
        On Error GoTo 0
    Next para
    
    ' Display the fonts in a message box
    Dim fontOutput As String
    fontOutput = "Fonts used in the document:" & vbCrLf
    For i = 1 To fontList.count
        fontOutput = fontOutput & "- " & fontList(i) & vbCrLf
    Next i
    'MsgBox fontOutput, vbInformation, "Fonts in Document"
    Debug.Print fontOutput
End Sub

Sub FindGentiumFromParagraph()
    Dim startParaNum As Long
    Dim para As paragraph
    Dim rng As range
    Dim charRange As range
    Dim i As Long, p As Long
    Dim totalParas As Long

    ' Ask user where to start
    startParaNum = val(InputBox("Enter paragraph number to start from:", "Start From Paragraph", 1))
    If startParaNum < 1 Then Exit Sub

    totalParas = ActiveDocument.paragraphs.count
    If startParaNum > totalParas Then
        MsgBox "There are only " & totalParas & " paragraphs in the document.", vbExclamation
        Exit Sub
    End If

    p = 0
    For Each para In ActiveDocument.paragraphs
        p = p + 1
        If p < startParaNum Then GoTo NextPara

        Set rng = para.range
        rng.End = rng.End - 1 ' Exclude paragraph mark

        For i = 1 To rng.Characters.count Step 10 ' Check every 10 chars
            Set charRange = rng.Characters(i)
            If charRange.font.name = "Gentium" Then
                charRange.Select
                MsgBox "Found Gentium font at paragraph " & p, vbInformation
                Application.StatusBar = False
                Exit Sub
            End If
        Next i

        If p Mod 100 = 0 Then
            Application.StatusBar = "Scanning paragraph " & p & " of " & totalParas & "..."
            DoEvents
        End If

NextPara:
    Next para

    Application.StatusBar = False
    MsgBox "Gentium font not found starting from paragraph " & startParaNum & ".", vbExclamation
End Sub

Sub GoToParagraph()
    Dim paraNum As Integer
    paraNum = (InputBox("Enter paragraph number:", "Goto Paragraph Number", 1))
    ActiveDocument.paragraphs(paraNum).range.Select
End Sub

Sub ListNonMainFonts_ByParagraph()
    Dim fontDict As Object
    Set fontDict = CreateObject("Scripting.Dictionary")

    Dim storyRange As range
    Dim para As paragraph
    Dim fontName As String
    Dim fontCount As Long
    Dim scannedParas As Long

    Application.ScreenUpdating = False
    Application.StatusBar = "Scanning fonts outside main text..."

    For Each storyRange In ActiveDocument.StoryRanges
        If storyRange.StoryType <> wdMainTextStory Then
            Do
                For Each para In storyRange.paragraphs
                    scannedParas = scannedParas + 1
                    fontName = para.range.font.name
                    If Len(fontName) > 0 Then
                        If Not fontDict.Exists(fontName) Then
                            fontDict.Add fontName, 1
                            fontCount = fontCount + 1
                        End If
                    End If

                    If scannedParas Mod 20 = 0 Then
                        Application.StatusBar = "Scanned " & scannedParas & " paragraphs... Fonts found: " & fontCount
                        DoEvents
                    End If
                Next para
                Set storyRange = storyRange.NextStoryRange
            Loop While Not storyRange Is Nothing
        End If
    Next storyRange

    Application.StatusBar = False
    Application.ScreenUpdating = True

    If fontDict.count = 0 Then
        MsgBox "No fonts found outside main text.", vbInformation
    Else
        Dim output As String, key As Variant
        output = "Fonts outside main document text:" & vbCrLf & vbCrLf
        For Each key In fontDict.Keys
            output = output & "- " & key & vbCrLf
        Next key
        'MsgBox output, vbInformation, "Non-Main Fonts"
        Debug.Print output
    End If
End Sub

Sub TestComp()
    CompareDocuments "C:\adaept\aeBibleClass\Peter-USE REFINED English Bible CONTENTS.docx", "C:\Users\peter\OneDrive\Documents\Peter-USE REFINED English Bible CONTENTS - Copy (49).docx"
End Sub

Sub CompareDocuments(original As String, modified As String)
' e.g. original = "C:\Path\To\Original.docx"
' e.g. "C:\Path\To\Modified.docx"
' - Original Document – The initial version of the document before changes were made.
' - Modified Document – The updated version that includes changes.
' - Comparison Document – The newly generated document that highlights differences between the original and modified versions.
' - The **comparison document** is a completely **new document** that shows changes such as insertions, deletions, and formatting modifications.
' - The **original** and **modified** documents remain **unchanged**—Word does **not** alter them.
' wdGranularityWordLevel
' - CompareFormatting (True) – Marks differences in formatting (e.g., font changes, bold/italic modifications).
' - CompareCaseChanges (True) – Highlights changes in letter case (e.g., "word" vs. "Word").
' - CompareWhitespace (True) – Tracks differences in spaces, paragraph breaks, and other whitespace variations.
' - CompareTables (True) – Compares changes within tables, including cell modifications.
' These options allow for a detailed comparison of documents, ensuring that even subtle changes are detected.
'
    Dim docOriginal As Document
    Dim docModified As Document
    Dim docComparison As Document
    Dim lastSlashPos As Integer
    Dim filePath As String
    
    lastSlashPos = InStrRev(original, "\") ' Find last occurrence of "\"
    If lastSlashPos > 0 Then
        filePath = Left(original, lastSlashPos) ' Get everything before the last "\"
    Else
        filePath = "" ' No path found, return empty string
    End If
    
    ' Open the original and modified documents
    Set docOriginal = Documents.Open(original)
    Set docModified = Documents.Open(modified)
    
    ' Create a comparison document
    Set docComparison = Application.CompareDocuments(docOriginal, docModified, wdCompareDestinationNew, _
        wdGranularityWordLevel, False, True, False, False)
    
    ' Save comparison result
    docComparison.SaveAs filePath & "\Comparison.docx"
    
    MsgBox "Comparison complete! See the document for tracked changes."
End Sub

Sub CountSearchHits()
    Dim searchTerm As String
    Dim count As Long
    Dim rng As range

    searchTerm = InputBox("Enter the text to search for:")
    If Len(searchTerm) = 0 Then Exit Sub

    count = 0
    Set rng = ActiveDocument.content
    With rng.Find
        .text = searchTerm
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False

        Do While .Execute
            count = count + 1
            rng.Collapse wdCollapseEnd
        Loop
    End With

    MsgBox "Found " & count & " instance(s) of '" & searchTerm & "'.", vbInformation
End Sub

Sub PrintHeading1sByLogicalPage()
    Dim i As Long
    Dim maxPage As Long
    Dim pageRange As range
    Dim para As paragraph
    Dim headingText As String
    Dim foundHeading As Boolean

    maxPage = ActiveDocument.range.Information(wdNumberOfPagesInDocument)

    Debug.Print "=== Heading 1s by Logical Page (GoTo ^H) ==="

    For i = 1 To maxPage
        Set pageRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=i)
        Set pageRange = pageRange.GoTo(What:=wdGoToBookmark, name:="\page") ' Get full page range

        foundHeading = False

        For Each para In pageRange.paragraphs
            If para.style = "Heading 1" Then
                headingText = Replace(para.range.text, vbCr, "")
                Debug.Print "Logical Page " & i & ": " & headingText
                foundHeading = True
                Exit For ' Only report first Heading 1 per page
            End If
        Next para

        If Not foundHeading Then
            ' Optional: Debug.Print "Logical Page " & i & ": No Heading 1"
        End If
    Next i
End Sub

Sub FixAndDiagnoseFootnoteReferences()
    Dim doc As Document
    Dim fn As footnote
    Dim fnRef As range
    Dim totalChecked As Long
    Dim totalIncorrect As Long
    Dim firstIncorrectPos As Long
    Dim firstFound As Boolean
    Dim mismatchDetail As String

    Set doc = ActiveDocument
    totalChecked = 0
    totalIncorrect = 0
    firstFound = False

    Debug.Print "FixAndDiagnoseFootnoteReferences"
    Debug.Print "Checking only footnote REFERENCE marks in main text..."

    For Each fn In doc.Footnotes
        Set fnRef = fn.Reference
        totalChecked = totalChecked + 1

        If totalChecked Mod 100 = 0 Then
            Debug.Print "Checked: " & totalChecked & " references..."
        End If

        If Not IsCorrectFootnoteFormat(fnRef, mismatchDetail) Then
            totalIncorrect = totalIncorrect + 1

            ' Attempt fix
            fnRef.font.Reset
            fnRef.style = doc.Styles("Footnote Reference")
            With fnRef.font
                .name = "Segoe UI"
                .Size = 8
                .Bold = True
                .color = wdColorBlue
                .Superscript = True
            End With

            If Not firstFound Then
                firstIncorrectPos = fnRef.Start
                firstFound = True
                Debug.Print "First incorrect formatting at character position: " & firstIncorrectPos
                Debug.Print "Mismatch details: " & vbCrLf & mismatchDetail
            End If
        End If
    Next fn

    Debug.Print "Total checked: " & totalChecked
    Debug.Print "Total incorrect: " & totalIncorrect

    If firstFound Then
        ' Move to the first incorrect reference (in main text)
        Selection.HomeKey Unit:=wdStory
        Selection.MoveRight Unit:=wdCharacter, count:=firstIncorrectPos
        Selection.Select
    Else
        MsgBox "All footnote reference formatting is correct."
    End If
End Sub

Function IsCorrectFootnoteFormat(rng As range, ByRef mismatch As String) As Boolean
    mismatch = ""
    IsCorrectFootnoteFormat = True
    With rng.font
        If rng.style <> "Footnote Reference" Then
            mismatch = mismatch & " - Style: " & rng.style & vbCrLf
            IsCorrectFootnoteFormat = False
        End If
        If .name <> "Segoe UI" Then
            mismatch = mismatch & " - Font Name: " & .name & vbCrLf
            IsCorrectFootnoteFormat = False
        End If
        If .Size <> 8 Then
            mismatch = mismatch & " - Size: " & .Size & vbCrLf
            IsCorrectFootnoteFormat = False
        End If
        If .Bold <> True Then
            mismatch = mismatch & " - Bold: " & .Bold & vbCrLf
            IsCorrectFootnoteFormat = False
        End If
        If .color <> wdColorBlue Then
            mismatch = mismatch & " - Color: " & .color & vbCrLf
            IsCorrectFootnoteFormat = False
        End If
        If .Superscript <> True Then
            mismatch = mismatch & " - Superscript: " & .Superscript & vbCrLf
            IsCorrectFootnoteFormat = False
        End If
    End With
End Function

Sub FixFootnoteNumberStyleInText()
    Dim fn As footnote
    Dim paraRange As range
    Dim firstRun As range

    For Each fn In ActiveDocument.Footnotes
        Set paraRange = fn.range.paragraphs(1).range
        Set firstRun = paraRange.Words(1) ' Usually the footnote number

        ' Apply Footnote Reference style
        firstRun.style = ActiveDocument.Styles("Footnote Reference")
    Next fn

    MsgBox "Footnote Reference style reapplied to footnote numbers in footnote text.", vbInformation
End Sub

Sub ReportPageLayoutMetrics(pageNum As Long)
    Dim pgRange As range
    Dim sectionSetup As pageSetup
    Dim numCols As Integer, isEven As Boolean
    Dim gutter As Single, pageWidth As Single
    Dim leftMargin As Single, rightMargin As Single
    Dim Spacing As Single, columnWidth As Single
    Dim sectionLeft As Single, colStart As Single
    Dim logBuffer As String, i As Integer

    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNum))
    Set sectionSetup = pgRange.Sections(1).pageSetup

    numCols = sectionSetup.TextColumns.count
    gutter = sectionSetup.gutter
    pageWidth = sectionSetup.pageWidth
    leftMargin = sectionSetup.leftMargin
    rightMargin = sectionSetup.rightMargin
    Spacing = sectionSetup.TextColumns(1).SpaceAfter
    columnWidth = (pageWidth - leftMargin - rightMargin - ((numCols - 1) * Spacing)) / numCols
    isEven = (pageNum Mod 2 = 0)
    sectionLeft = IIf(sectionSetup.MirrorMargins And isEven, gutter + leftMargin, leftMargin)

    logBuffer = "=== Layout Metrics for Page " & pageNum & " ===" & vbCrLf
    logBuffer = logBuffer & "Mirror Margins: " & sectionSetup.MirrorMargins & vbCrLf
    logBuffer = logBuffer & "Page Width: " & Format(pageWidth, "0.0") & vbCrLf
    logBuffer = logBuffer & "Left Margin: " & Format(leftMargin, "0.0") & vbCrLf
    logBuffer = logBuffer & "Right Margin: " & Format(rightMargin, "0.0") & vbCrLf
    logBuffer = logBuffer & "Gutter: " & Format(gutter, "0.0") & vbCrLf
    logBuffer = logBuffer & "Column Count: " & numCols & vbCrLf
    logBuffer = logBuffer & "Column Width: " & Format(columnWidth, "0.0") & vbCrLf
    logBuffer = logBuffer & "Spacing Between Columns: " & Format(Spacing, "0.0") & vbCrLf
    logBuffer = logBuffer & "Section Left Edge: " & Format(sectionLeft, "0.0") & vbCrLf

    For i = 0 To numCols - 1
        colStart = sectionLeft + i * (columnWidth + Spacing)
        logBuffer = logBuffer & "? Column " & (i + 1) & " starts at: " & Format(colStart, "0.0") & vbCrLf
    Next i

    logBuffer = logBuffer & "=== End of Layout Report ==="
    Debug.Print logBuffer
    MsgBox "Layout metrics for page " & pageNum & " printed to Immediate window.", vbInformation
End Sub

Sub ReportDigitAtCursor_Diagnostics()
    Dim selRange As range, ch As range, prefix As range
    Dim txt As String, style As String
    Dim posX As Single, posY As Single
    Dim prefixTxt As String, prefixStyle As String, prefixAsc As Variant
    Dim prefixX As Single, prefixY As Single

    Set selRange = Selection.range
    Set ch = ActiveDocument.range(selRange.Start, selRange.Start + 1)
    txt = ch.text
    style = ch.style.NameLocal
    posX = ch.Information(wdHorizontalPositionRelativeToPage)
    posY = ch.Information(wdVerticalPositionRelativeToPage)

    Debug.Print "=== Character at Cursor ==="
    Debug.Print "Value: '" & txt & "' | ASCII: " & AscW(txt)
    Debug.Print "Style: " & style
    Debug.Print "Font Color: " & ch.font.color & " (RGB: " & _
                RGBToString(ch.font.color) & ")"
    Debug.Print "Position: X=" & Format(posX, "0.0") & " pts, Y=" & Format(posY, "0.0") & " pts"

    If ch.Start > 0 Then
        Set prefix = ActiveDocument.range(ch.Start - 1, ch.Start)
        prefixTxt = prefix.text
        prefixStyle = prefix.style.NameLocal
        prefixAsc = AscW(prefixTxt)
        prefixX = prefix.Information(wdHorizontalPositionRelativeToPage)
        prefixY = prefix.Information(wdVerticalPositionRelativeToPage)

        Debug.Print "--- Prefix (1 char before) ---"
        Debug.Print "Value: '" & prefixTxt & "' | ASCII: " & prefixAsc
        Debug.Print "Style: " & prefixStyle
        Debug.Print "Position: X=" & Format(prefixX, "0.0") & " pts, Y=" & Format(prefixY, "0.0") & " pts"
    Else
        Debug.Print "--- No prefix (at start of document) ---"
    End If
End Sub

Function RGBToString(rgbVal As Long) As String
    RGBToString = "(" & (rgbVal And &HFF) & "," & ((rgbVal \ 256) And &HFF) & "," & ((rgbVal \ 65536) And &HFF) & ")"
End Function

Sub ReportDigitAtCursor_Diagnostics_Expanded()
    Dim rng As range
    Set rng = Selection.range
    If rng.Characters.count = 0 Then
        MsgBox "No character selected.", vbExclamation
        Exit Sub
    End If

    Dim ch As range
    Set ch = rng.Characters(1)
    Dim txt As String: txt = ch.text
    Dim ascCode As Long: ascCode = AscW(txt)
    Dim fontNameAscii As String: fontNameAscii = ch.font.NameAscii
    Dim fontNameFarEast As String: fontNameFarEast = ch.font.NameFarEast
    Dim fontNameOther As String: fontNameOther = ch.font.NameOther
    Dim fontSize As Single: fontSize = ch.font.Size
    Dim fontColor As Long: fontColor = ch.font.color
    Dim styleName As String: styleName = ch.style.NameLocal
    Dim baseStyle As String
    On Error Resume Next
    baseStyle = ch.style.baseStyle
    On Error GoTo 0

    Debug.Print "=== Character at Cursor ==="
    Debug.Print "Value: '" & txt & "' | ASCII: " & ascCode
    Debug.Print "Style: " & styleName
    Debug.Print "Base Style: " & IIf(baseStyle = "", "(none)", baseStyle)

    Debug.Print "Font Names:"
    Debug.Print "> NameAscii: " & fontNameAscii
    Debug.Print "> NameFarEast: " & fontNameFarEast
    Debug.Print "> NameOther: " & fontNameOther
    Debug.Print "Font Size: " & fontSize & " pt"
    Debug.Print "Font Color: " & fontColor & " (RGB: " & _
        (fontColor Mod 256) & "," & ((fontColor \ 256) Mod 256) & "," & (fontColor \ 65536) & ")"
    Debug.Print "Bold: " & ch.font.Bold & " | Italic: " & ch.font.Italic & " | Underline: " & ch.font.Underline

    Debug.Print "--- Prefix (1 char before) ---"
    If ch.Start > 1 Then
        Dim prefix As range
        Set prefix = ActiveDocument.range(ch.Start - 1, ch.Start)
        Debug.Print "Value: '" & prefix.text & "' | ASCII: " & AscW(prefix.text)
        Debug.Print "Style: " & prefix.style.NameLocal
        Debug.Print "Font Name: " & prefix.font.name
        Debug.Print "Font Color: " & prefix.font.color & " (RGB: " & _
            (prefix.font.color Mod 256) & "," & ((prefix.font.color \ 256) Mod 256) & "," & (prefix.font.color \ 65536) & ")"
    Else
        Debug.Print "(No character before this one.)"
    End If

    MsgBox "Expanded character diagnostics logged.", vbInformation
End Sub

Sub LogExpandedMarkerContext()
    Dim sel As range: Set sel = Selection.range
    Dim i As Long, chCount As Long
    Dim contextText As String, contextAscii As String, contextHex As String

    chCount = sel.Characters.count
    Debug.Print "=== Marker Diagnostic ==="
    Debug.Print "Selection Start=" & sel.Start & " | End=" & sel.End
    Debug.Print "Selection Text='" & Replace(sel.text, vbCr, "[CR]") & "'"

    For i = 1 To chCount
        Dim ch As String: ch = sel.Characters(i).text
        Dim ascVal As Integer: ascVal = Asc(ch)
        Dim hexVal As String: hexVal = Hex(ascVal)

        contextText = "[" & i & "] '" & Replace(ch, vbCr, "[CR]") & "'"
        contextAscii = " ASCII=" & ascVal
        contextHex = " Hex=" & hexVal

        Debug.Print contextText & contextAscii & contextHex
    Next i

    Debug.Print "Style: " & sel.style & " | Font: " & sel.font.name
    Debug.Print "=== End of Diagnostic ===" & vbCrLf
End Sub

Sub FindInvisibleFormFeeds_InPages(startPage As Long)
    Dim para As paragraph, rng As range
    Dim pgNum As Long
    Dim i As Long, pgTarget As Long

    pgTarget = startPage + 9
    Debug.Print "=== Scanning for Chr(12) from Page " & startPage & " to " & pgTarget & " ==="

    For Each para In ActiveDocument.paragraphs
        Set rng = para.range
        pgNum = rng.Information(wdActiveEndPageNumber)

        If pgNum >= startPage And pgNum <= pgTarget Then
            If InStr(rng.text, Chr(12)) > 0 Then
                Debug.Print "[Page " & pgNum & "] Chr(12) found at Start=" & rng.Start
                Debug.Print "Text='" & Replace(rng.text, Chr(12), "[FF]") & "'"
                Debug.Print "Style=" & rng.style & " | Font=" & rng.font.name
            End If
        End If
    Next para

    Debug.Print "=== End of Scan ===" & vbCrLf
End Sub

Sub AuditVerseMarkers_VerifyMergedNumberPrefix_WithContext(pageNum As Long)
    Dim pgRange As range, ch As range, scanRange As range
    Dim pageStart As Long, pageEnd As Long
    Dim logBuffer As String
    Dim chapterMarker As String, verseDigits As String, combinedNumber As String
    Dim markerStart As Long, markerEnd As Long, verseEnd As Long
    Dim digitPosX As Single, digitPosY As Single
    Dim wordRange As range, token As range, nextWords As String, wCount As Integer

    logBuffer = "=== Chapter–Verse Visual Number Check on Page " & pageNum & " ===" & vbCrLf

    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNum))
    pageStart = pgRange.Start
    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNum + 1))
    pageEnd = pgRange.Start - 1
    Set pgRange = ActiveDocument.range(pageStart, pageEnd)

    Dim i As Long
    i = pageStart
    Do While i < pageEnd
        Set ch = ActiveDocument.range(i, i + 1)
        If Len(Trim(ch.text)) = 1 And IsNumeric(ch.text) And ch.style.NameLocal = "Chapter Verse marker" And ch.font.color = RGB(255, 165, 0) Then
            chapterMarker = ch.text
            markerStart = i
            markerEnd = i + 1
            Do While markerEnd < pageEnd
                Set scanRange = ActiveDocument.range(markerEnd, markerEnd + 1)
                If Len(Trim(scanRange.text)) = 1 And IsNumeric(scanRange.text) Then
                    If scanRange.style.NameLocal = "Chapter Verse marker" And scanRange.font.color = RGB(255, 165, 0) Then
                        chapterMarker = chapterMarker & scanRange.text
                        markerEnd = markerEnd + 1
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop

            digitPosX = ch.Information(wdHorizontalPositionRelativeToPage)
            digitPosY = ch.Information(wdVerticalPositionRelativeToPage)

            verseDigits = ""
            verseEnd = markerEnd
            Do While verseEnd < pageEnd
                Set scanRange = ActiveDocument.range(verseEnd, verseEnd + 1)
                If Len(Trim(scanRange.text)) = 1 And IsNumeric(scanRange.text) Then
                    If scanRange.style.NameLocal = "Verse marker" And scanRange.font.color = RGB(80, 200, 120) Then
                        verseDigits = verseDigits & scanRange.text
                        verseEnd = verseEnd + 1
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop

            nextWords = ""
            Set wordRange = ActiveDocument.range(verseEnd, verseEnd + 80)
            wCount = 0
            For Each token In wordRange.Words
                If token.text Like "*^13*" Then Exit For
                If Trim(token.text) <> "" Then
                    nextWords = nextWords & Trim(token.text) & " "
                    wCount = wCount + 1
                    If wCount = 2 Then Exit For
                End If
            Next token

            If Len(verseDigits) = 0 Then
                logBuffer = logBuffer & "! Chapter '" & chapterMarker & "' @ X=" & Format(digitPosX, "0.0") & _
                    " | No styled Verse marker digits found | Next words: “" & Trim(nextWords) & "”" & vbCrLf
            Else
                combinedNumber = chapterMarker & verseDigits
                If Left(combinedNumber, Len(chapterMarker)) = chapterMarker Then
                    logBuffer = logBuffer & "* Chapter '" & chapterMarker & "' ? Verse '" & combinedNumber & "' @ X=" & Format(digitPosX, "0.0") & _
                        " | ? Valid | Next words: “" & Trim(nextWords) & "”" & vbCrLf
                Else
                    logBuffer = logBuffer & "! Chapter '" & chapterMarker & "' ? Verse '" & combinedNumber & "' @ X=" & Format(digitPosX, "0.0") & _
                        " | ? Mismatch | Next words: “" & Trim(nextWords) & "”" & vbCrLf
                End If
            End If

            i = verseEnd
        Else
            i = i + 1
        End If
    Loop

    logBuffer = logBuffer & "=== Audit complete ==="
    Debug.Print logBuffer
    MsgBox "Visual prefix check with context logged for page " & pageNum & ".", vbInformation
End Sub

Sub ReportAllMarkers_CondensedDiagnostics(pageNum As Long)
    Dim pgRange As range, ch As range, scanRange As range
    Dim pageStart As Long, pageEnd As Long
    Dim txt As String, styleName As String
    Dim fontName As String, fontSize As Single, fontColor As Long
    Dim charPosX As Single, charPosY As Single
    Dim digitBlock As String, blockStyle As String, blockColor As Long
    Dim blockStart As Long, blockEnd As Long, logBuffer As String

    logBuffer = "=== Marker Summary for Page " & pageNum & " ===" & vbCrLf

    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNum))
    pageStart = pgRange.Start
    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNum + 1))
    pageEnd = pgRange.Start - 1
    Set pgRange = ActiveDocument.range(pageStart, pageEnd)

    Dim i As Long
    i = pageStart
    Do While i < pageEnd
        Set ch = ActiveDocument.range(i, i + 1)
        txt = Trim(ch.text)
        styleName = ch.style.NameLocal

        If Len(txt) = 1 And IsNumeric(txt) Then
            If styleName = "Chapter Verse marker" Or styleName = "Verse marker" Then
                digitBlock = txt
                blockStyle = styleName
                blockColor = ch.font.color
                blockStart = i
                blockEnd = i + 1

                Do While blockEnd < pageEnd
                    Set scanRange = ActiveDocument.range(blockEnd, blockEnd + 1)
                    If Len(Trim(scanRange.text)) = 1 And IsNumeric(scanRange.text) Then
                        If scanRange.style.NameLocal = blockStyle And scanRange.font.color = blockColor Then
                            digitBlock = digitBlock & scanRange.text
                            blockEnd = blockEnd + 1
                        Else
                            Exit Do
                        End If
                    Else
                        Exit Do
                    End If
                Loop

                Set ch = ActiveDocument.range(blockStart, blockStart + 1)
                fontName = ch.font.name
                fontSize = ch.font.Size
                charPosX = ch.Information(wdHorizontalPositionRelativeToPage)
                charPosY = ch.Information(wdVerticalPositionRelativeToPage)

                logBuffer = logBuffer & "[" & IIf(blockStyle = "Chapter Verse marker", "Chapter", "Verse") & "] '" & digitBlock & "' @ X=" & Format(charPosX, "0.0") & ", Y=" & Format(charPosY, "0.0") & _
                    " | Font: " & fontName & " " & fontSize & "pt | RGB: (" & (blockColor Mod 256) & "," & ((blockColor \ 256) Mod 256) & "," & (blockColor \ 65536) & ")" & _
                    " | Pos: " & blockStart & "–" & (blockEnd - 1) & vbCrLf

                i = blockEnd
                GoTo ContinueLoop
            End If
        End If

        i = i + 1
ContinueLoop:
    Loop

    logBuffer = logBuffer & "=== Summary complete ==="
    Debug.Print logBuffer
    MsgBox "Condensed diagnostics logged.", vbInformation
    Selection.GoTo What:=wdGoToPage, name:=CStr(pageNum)
End Sub

Sub RepairWrappedVerseMarkers_MergedPrefix_ByColumnContext_SinglePage(pageNum As Long, ByRef fixCount As Long)
    ' Same logic as full macro, but suppresses MsgBox and passes fixCount by reference.
    ' Copy the full body from RepairWrappedVerseMarkers_MergedPrefix_ByColumnContext here
    ' And replace `MsgBox` line with: fixCount = fixCount
    Dim pgRange As range, ch As range, scanRange As range, prefixCh As range
    Dim pageStart As Long, pageEnd As Long
    Dim chapterMarker As String, verseDigits As String, combinedNumber As String
    Dim markerStart As Long, markerEnd As Long, verseEnd As Long
    Dim prefixTxt As String, prefixStyle As String, prefixAsc As Variant
    Dim prefixY As Single, digitY As Single, digitX As Single
    Dim nextWords As String, lookAhead As range, token As range, wCount As Integer
    Dim logBuffer As String

    fixCount = 0
    logBuffer = "=== Smart Prefix Repair on Page " & pageNum & " ===" & vbCrLf

    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNum))
    pageStart = pgRange.Start
    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNum + 1))
    pageEnd = pgRange.Start - 1

    Dim i As Long
    i = pageStart
    Do While i < pageEnd
        Set ch = ActiveDocument.range(i, i + 1)
        If Len(Trim(ch.text)) = 1 And IsNumeric(ch.text) And ch.style.NameLocal = "Chapter Verse marker" And ch.font.color = RGB(255, 165, 0) Then
            ' Assemble chapter marker block
            chapterMarker = ch.text
            markerStart = i
            markerEnd = i + 1
            Do While markerEnd < pageEnd
                Set scanRange = ActiveDocument.range(markerEnd, markerEnd + 1)
                If Len(Trim(scanRange.text)) = 1 And IsNumeric(scanRange.text) Then
                    If scanRange.style.NameLocal = "Chapter Verse marker" And scanRange.font.color = RGB(255, 165, 0) Then
                        chapterMarker = chapterMarker & scanRange.text
                        markerEnd = markerEnd + 1
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop

            digitY = ch.Information(wdVerticalPositionRelativeToPage)
            digitX = ch.Information(wdHorizontalPositionRelativeToPage)

            ' Assemble verse marker block
            verseDigits = ""
            verseEnd = markerEnd
            Do While verseEnd < pageEnd
                Set scanRange = ActiveDocument.range(verseEnd, verseEnd + 1)
                If Len(Trim(scanRange.text)) = 1 And IsNumeric(scanRange.text) Then
                    If scanRange.style.NameLocal = "Verse marker" And scanRange.font.color = RGB(80, 200, 120) Then
                        verseDigits = verseDigits & scanRange.text
                        verseEnd = verseEnd + 1
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop

            If Len(verseDigits) > 0 Then
                combinedNumber = chapterMarker & verseDigits

                ' Prefix check
                If markerStart > pageStart Then
                    Set prefixCh = ActiveDocument.range(markerStart - 1, markerStart)
                    prefixTxt = prefixCh.text
                    prefixStyle = prefixCh.style.NameLocal
                    prefixAsc = AscW(prefixTxt)
                    prefixY = prefixCh.Information(wdVerticalPositionRelativeToPage)

                    If (prefixAsc = 32 Or prefixAsc = 160) And prefixStyle = "Normal" Then
                        If Abs(prefixY - digitY) < 25 Then
                            nextWords = ""
                            Set lookAhead = ActiveDocument.range(verseEnd, verseEnd + 80)
                            wCount = 0
                            For Each token In lookAhead.Words
                                If token.text Like "*^13*" Then Exit For
                                If Trim(token.text) <> "" Then
                                    nextWords = nextWords & Trim(token.text) & " "
                                    wCount = wCount + 1
                                    If wCount = 2 Then Exit For
                                End If
                            Next token

                            ' Column edge logic
                            If digitX < 50 Then
                                prefixCh.text = vbCr
                                logBuffer = logBuffer & "? Repaired prefix before '" & combinedNumber & "' @ X=" & Format(digitX, "0.0") & " | Break inserted | Next words: “" & Trim(nextWords) & "”" & vbCrLf
                            Else
                                prefixCh.text = ""
                                logBuffer = logBuffer & "? Removed space before '" & combinedNumber & "' @ X=" & Format(digitX, "0.0") & " | No break | Next words: “" & Trim(nextWords) & "”" & vbCrLf
                            End If

                            fixCount = fixCount + 1
                        End If
                    End If
                End If

                i = verseEnd
            Else
                i = markerEnd
            End If
        Else
            i = i + 1
        End If
    Loop

    logBuffer = logBuffer & "=== " & fixCount & " markers repaired on page " & pageNum & " ==="
    Debug.Print logBuffer
    'MsgBox fixCount & " marker(s) repaired on page " & pageNum & ".", vbInformation
    fixCount = fixCount
    Selection.GoTo What:=wdGoToPage, name:=CStr(pageNum)
End Sub

Sub SmartPrefixRepairOnPage_WithDiagnostics(pgNum As Long, ByRef spaceCount As Long, ByRef breakCount As Long)
    ' Simulates repair logic for testing
    ' This assumes every page has 3 space repairs and 2 break repairs
    Dim j As Long

    For j = 1 To 3
        ' Simulate space-only marker fix
        spaceCount = spaceCount + 1
        ' Replace with: Selection.Delete or other spacing logic
    Next j

    For j = 1 To 2
        ' Simulate prefix insertion (vbCr)
        breakCount = breakCount + 1
        ' Replace with: Selection.InsertBefore vbCr or similar
    Next j
End Sub

Sub SmartyOne()
    Dim sCount As Long, bCount As Long
    Call SmartPrefixRepairOnPage(235, sCount, bCount)
End Sub

Sub SmartPrefixRepairOnPage(pgNum As Long, ByRef spaceCount As Long, ByRef breakCount As Long)
    Dim para As paragraph
    Dim rng As range
    Dim markerText As String
    Dim didRepair As Boolean
    Dim paraStyle As String
    Dim ascii12Count As Long
    Dim missing160Count As Long

    Debug.Print "=== Smart Prefix Repair on Page " & pgNum & " ==="

    For Each para In ActiveDocument.paragraphs
        Set rng = para.range
        paraStyle = rng.style

        ' Only process green "Verse marker" paragraphs
        If InStr(paraStyle, "Verse marker") > 0 Then
            markerText = rng.text

            ' Skip and count layout wrappers: lone Chr(12)
            If Len(markerText) = 1 And Asc(markerText) = 12 Then
                ascii12Count = ascii12Count + 1
                GoTo NextPara
            End If

            didRepair = False

            ' Column-aware prefix repair: if verse starts in left column (1)
            If rng.Information(wdStartOfRangeColumnNumber) = 1 Then
                rng.InsertBefore vbCr
                breakCount = breakCount + 1
                didRepair = True
                Debug.Print "? Repaired prefix (wrapped to left column) before '" & Trim(markerText) & "'"
            End If

            ' Remove space if leading whitespace present
            If Left(markerText, 1) = " " Then
                rng.Characters(1).Delete
                spaceCount = spaceCount + 1
                didRepair = True
                Debug.Print "? Removed space before '" & Trim(markerText) & "'"
            End If

            ' Diagnostics for skipped cases
            If Not didRepair Then
                Debug.Print "- Skipped marker '" & Trim(markerText) & "'"
                Debug.Print "  ASCII codes for first few characters:"
                Dim i As Long, limit As Long
                If Len(markerText) < 5 Then
                    limit = Len(markerText)
                Else
                    limit = 5
                End If
                For i = 1 To limit
                    Dim ch As String
                    Dim ascVal As Integer
                    ch = Mid(markerText, i, 1)
                    ascVal = Asc(ch)
                    Debug.Print "    Char " & i & ": '" & Replace(ch, vbCr, "[CR]") & "' | ASCII=" & ascVal & " | Hex=" & Hex(ascVal)
                Next i

                ' Check if last character is missing expected Chr(160)
                If Len(markerText) = 0 Or Asc(Right(markerText, 1)) <> 160 Then
                    missing160Count = missing160Count + 1
                End If

                Debug.Print "  Style=" & rng.style & " | Font=" & rng.font.name & " | Start=" & rng.Start & " | Page=" & pgNum
            End If
        End If
NextPara:
    Next para

    Debug.Print "Chr(12) marker count on Page " & pgNum & ": " & ascii12Count
    Debug.Print "Missing Chr(160) count on Page " & pgNum & ": " & missing160Count
    Debug.Print "=== End of Repairs for Page " & pgNum & " ==="
End Sub

Sub RunRepairWrappedVerseMarkers_Across10Pages_From(StartPageNum As Long)
    Const ForecastFile As String = "RepairRunnerForecast.txt"
    Dim SessionID As String: SessionID = "Session_" & Format(Now, "yyyymmdd_HHMMSS")
    Dim filePath As String: filePath = ThisDocument.Path & "\" & ForecastFile

    Dim fs As Object, ts As Object
    Set fs = CreateObject("Scripting.FileSystemObject")

    If fs.FileExists(filePath) = False Then
        Set ts = fs.CreateTextFile(filePath, True)
        ts.WriteLine "SessionID,StartPageNum," & _
            "Time1,Time2,Time3,Time4,Time5,Time6,Time7,Time8,Time9,Time10," & _
            "TotalTime,ForecastTime," & _
            "Space1,Space2,Space3,Space4,Space5,Space6,Space7,Space8,Space9,Space10," & _
            "Break1,Break2,Break3,Break4,Break5,Break6,Break7,Break8,Break9,Break10"
    Else
        Set ts = fs.OpenTextFile(filePath, 8, True)
    End If
    
    Dim i As Long, t0 As Single, t1 As Single
    Dim timeStamps(1 To 10) As Single
    Dim spaceRepairs(1 To 10) As Long, breakRepairs(1 To 10) As Long
    Dim totalTime As Single, forecastTime As Single, forecastPages As Long

    t0 = Timer
    For i = 1 To 10
        t1 = Timer
        Dim pageNum As Long: pageNum = StartPageNum + (i - 1)

        ' Replace this with your actual repair logic
        ' Ensure spaceRepairs(i) and breakRepairs(i) are incremented during the repair
        Call SmartPrefixRepairOnPage(pageNum, spaceRepairs(i), breakRepairs(i))

        timeStamps(i) = Round(Timer - t1, 2)
    Next i
    totalTime = Round(Timer - t0, 2)

    For i = 1 To 10
        If timeStamps(i) >= 0.1 Then
            forecastTime = forecastTime + timeStamps(i)
            forecastPages = forecastPages + 1
        End If
    Next i
    If forecastPages > 0 Then
        forecastTime = Round(forecastTime / forecastPages * 10, 2)
    Else
        forecastTime = 0
    End If

    Dim resultRow As String
    resultRow = SessionID & "," & StartPageNum & ","
    For i = 1 To 10: resultRow = resultRow & timeStamps(i) & ",": Next i
    resultRow = resultRow & totalTime & "," & forecastTime & ","
    For i = 1 To 10: resultRow = resultRow & spaceRepairs(i) & ",": Next i
    For i = 1 To 10
        resultRow = resultRow & breakRepairs(i)
        If i < 10 Then resultRow = resultRow & ","
    Next i
    Debug.Print resultRow
    ts.WriteLine resultRow
    ts.Close
End Sub

Sub RunRepairWrappedVerseMarkers_ForOnePage(pgNum As Long)
    Const ForecastFile As String = "RepairRunnerForecast.txt"
    Dim SessionID As String: SessionID = "Session_" & Format(Now, "yyyymmdd_HHMMSS")
    Dim filePath As String: filePath = ThisDocument.Path & "\" & ForecastFile

    Dim fs As Object, ts As Object
    Set fs = CreateObject("Scripting.FileSystemObject")

    If fs.FileExists(filePath) = False Then
        Set ts = fs.CreateTextFile(filePath, True)
        ts.WriteLine "SessionID,PageNum," & _
            "Time," & "SpaceRepairs," & "BreakRepairs"
    Else
        Set ts = fs.OpenTextFile(filePath, 8, True)
    End If

    Dim tStart As Single, tElapsed As Single
    Dim spaceCount As Long, breakCount As Long

    tStart = Timer
    Call SmartPrefixRepairOnPage(pgNum, spaceCount, breakCount)
    tElapsed = Round(Timer - tStart, 2)

    Dim resultRow As String
    resultRow = SessionID & "," & pgNum & "," & tElapsed & "," & spaceCount & "," & breakCount
    Debug.Print resultRow
    ts.WriteLine resultRow
    ts.Close
End Sub

Sub UnlinkHeadingNumbering()
    Dim para As paragraph

    For Each para In ActiveDocument.paragraphs
        If para.style = ActiveDocument.Styles("Heading 1") _
        Or para.style = ActiveDocument.Styles("Heading 2") Then
            para.range.ListFormat.RemoveNumbers
        End If
    Next

    MsgBox "Numbering removed from Heading 1 and Heading 2 paragraphs.", vbInformation
End Sub

Sub StartRepairTimingSession(StartPageNum As Long)
    Const ForecastFile As String = "RepairRunnerForecast.txt"
    Dim SessionID As String: SessionID = "Session_" & Format(Now, "yyyymmdd_HHMMSS")

    Dim fs As Object, ts As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ts = fs.OpenTextFile(ThisDocument.Path & "\" & ForecastFile, 8, True) ' Append mode

    Dim i As Long, t0 As Single, t1 As Single
    Dim timeStamps(1 To 10) As Single
    Dim totalTime As Single, forecastTime As Single
    Dim forecastPages As Long

    t0 = Timer
    For i = 1 To 10
        t1 = Timer
        ' Placeholder call: actual repair code will go here later
        Call DummyRepairPageTimerOnly(StartPageNum + i - 1)
        timeStamps(i) = Round(Timer - t1, 2)
    Next i
    totalTime = Round(Timer - t0, 2)

    ' Forecast excludes pages likely to be headers (low timing)
    For i = 1 To 10
        If timeStamps(i) >= 0.1 Then
            forecastTime = forecastTime + timeStamps(i)
            forecastPages = forecastPages + 1
        End If
    Next i
    If forecastPages > 0 Then
        forecastTime = Round(forecastTime / forecastPages * 10, 2)
    Else
        forecastTime = 0
    End If

    ' Compile debug string
    Dim resultRow As String
    resultRow = SessionID & "," & StartPageNum & ","
    For i = 1 To 10: resultRow = resultRow & timeStamps(i) & ",": Next i
    resultRow = resultRow & totalTime & "," & forecastTime

    Debug.Print resultRow
    ts.WriteLine resultRow
    ts.Close
End Sub

Sub DummyRepairPageTimerOnly(pgNum As Long)
    ' Placeholder page-level logic for timing only
    Dim dummyWait As Single
    dummyWait = Timer
    Do While Timer < dummyWait + 0.05: DoEvents: Loop
End Sub

Sub ReapplyTheFootersToAllFooters()
    Dim sec As section
    Dim hf As HeaderFooter
    Dim p As paragraph
    Dim prevStyle As String
    Dim asciiVal As Long
    Dim paraText As String

    Debug.Print "=== Reapply 'TheFooters' Style Start ==="

    For Each sec In ActiveDocument.Sections
        Debug.Print "SECTION " & sec.Index

        For Each hf In sec.Footers
            If hf.Exists Then
                For Each p In hf.range.paragraphs
                    paraText = p.range.text
                    asciiVal = AscW(Left(paraText, 1))
                    prevStyle = p.style.NameLocal
                    p.style = "TheFooters"

                    Debug.Print "  Footer paragraph updated: " & _
                                "PrevStyle='" & prevStyle & "' | ASCII=" & asciiVal & " | HEX=" & Hex(asciiVal)
                Next p
            End If
        Next hf

        Debug.Print "----------------------------------------"
    Next sec

    Debug.Print "=== Style Reapplication Complete ==="
End Sub

'------------------------------------------------------------------------------
' Macro Name : GetHeadingDefinitionsWithDescriptions
' Author     : Peter + Copilot
' Description:
'   Retrieves style properties for Heading 1 and Heading 2 from the active document.
'   Includes font, paragraph formatting, outline level, and full color diagnostics.
'   Color output includes raw Word color value (Long), RGB breakdown, and Hex string.
'
' Output:
'   Printed to Immediate Window (Ctrl+G) for audit purposes.
'   Example line: Color: -16777216, RGB(0,0,0), #000000
'
' Dependencies:
'   Requires Word constants (e.g., wdAlignParagraphCenter) to be available.
'   All style names must exist in the document or error handling should be added.
'
' Future Extensions:
'   - Export to CSV or Markdown
'   - Include suffix tracking, style inheritance, or font audit flags
'   - Integrate session-aware tracking or timing metrics
'------------------------------------------------------------------------------
Sub GetHeadingDefinitionsWithDescriptions()
    Dim headingStyles As Variant
    headingStyles = Array("Heading 1", "Heading 2")
    
    Dim s As style
    Dim info As String
    Dim styleName As Variant
    Dim alignValue As Integer
    Dim alignText As String
    
    Dim clr As Long
    Dim r As Long, g As Long, b As Long
    Dim hexColor As String

    For Each styleName In headingStyles
        Set s = ActiveDocument.Styles(styleName)
        alignValue = s.ParagraphFormat.Alignment
        
        Select Case alignValue
            Case wdAlignParagraphLeft: alignText = "Left"
            Case wdAlignParagraphCenter: alignText = "Center"
            Case wdAlignParagraphRight: alignText = "Right"
            Case wdAlignParagraphJustify: alignText = "Justified"
            Case wdAlignParagraphDistribute: alignText = "Distributed"
            Case wdAlignParagraphThaiJustify: alignText = "Thai Distributed"
            Case Else: alignText = "Unknown"
        End Select
        
        clr = s.font.color
        r = clr Mod 256
        g = (clr \ 256) Mod 256
        b = (clr \ 65536) Mod 256
        hexColor = "#" & Right$("0" & Hex(r), 2) & Right$("0" & Hex(g), 2) & Right$("0" & Hex(b), 2)

        info = "Style: " & styleName & vbCrLf
        info = info & "  Font Name: " & s.font.name & vbCrLf
        info = info & "  Font Size: " & s.font.Size & vbCrLf
        info = info & "  Bold: " & s.font.Bold & vbCrLf
        info = info & "  Italic: " & s.font.Italic & vbCrLf
        info = info & "  Color: " & clr & ", RGB(" & r & "," & g & "," & b & "), " & hexColor & vbCrLf
        info = info & "  Alignment: " & alignValue & " (" & alignText & ")" & vbCrLf
        info = info & "  Space Before: " & s.ParagraphFormat.SpaceBefore & vbCrLf
        info = info & "  Space After: " & s.ParagraphFormat.SpaceAfter & vbCrLf
        info = info & "  Line Spacing: " & s.ParagraphFormat.LineSpacing & vbCrLf
        info = info & "  Outline Level: " & s.ParagraphFormat.OutlineLevel & vbCrLf
        info = info & "  Keep With Next: " & s.ParagraphFormat.KeepWithNext & vbCrLf
        info = info & String(40, "-") & vbCrLf
        
        Debug.Print info
    Next styleName
End Sub

Sub UpdateHeading2KeepWithNext()
    Dim s As style
    Set s = ActiveDocument.Styles("Heading 2")
    
    ' Apply KeepWithNext to paragraph formatting
    s.ParagraphFormat.KeepWithNext = True

    Debug.Print "Heading 2 style updated: KeepWithNext = True"
End Sub

Sub EnforceHeading2ParagraphWidowOrphan()
    Dim para As paragraph
    For Each para In ActiveDocument.paragraphs
        If para.style = ActiveDocument.Styles("Heading 2") Then '
            With para
                .WidowControl = True    ' enforces both widow and orphan control for that paragraph
                '.OrphanControl = True - Not needed,
            End With
        End If
    Next para
    Debug.Print "[repair] Widow/Orphan enforced at paragraph level for Heading 2"
End Sub

Sub DisableKeepLinesTogetherForHeading2()
    Dim para As paragraph
    For Each para In ActiveDocument.paragraphs
        If para.style = ActiveDocument.Styles("Heading 2") Then
            para.KeepTogether = False
        End If
    Next para
    Debug.Print "[repair] KeepLinesTogether disabled for Heading 2"
End Sub


