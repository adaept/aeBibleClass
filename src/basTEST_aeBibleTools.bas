Attribute VB_Name = "basTEST_aeBibleTools"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Private Sections1Col As Integer
Private Sections2Col As Integer
Private SectionsOddPageBreaks As Integer
Private SectionsEvenPageBreaks As Integer
Private SectionsContinuousBreaks As Integer
Private SectionsNewPageBreaks As Integer
' Module-level buffer
Dim HeadingBuffer As Object ' Late-bound Dictionary

Public Sub ListCustomXMLParts()
    On Error GoTo PROC_ERR
    Dim xmlPart As customXMLPart
    Dim i As Integer
    i = 1
    For Each xmlPart In ThisDocument.CustomXMLParts
        Debug.Print "Custom XML Part " & i & ": " & xmlPart.XML
        i = i + 1
    Next xmlPart

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ListCustomXMLParts of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub ListCustomXMLSchemas()
    On Error GoTo PROC_ERR
    Dim xmlPart As customXMLPart
    For Each xmlPart In ActiveDocument.CustomXMLParts
        Debug.Print xmlPart.NamespaceURI
    Next xmlPart

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ListCustomXMLSchemas of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub AddCustomUIXML()
    On Error GoTo PROC_ERR
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

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure AddCustomUIXML of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub RemoveDuplicateCustomXMLParts()
    On Error GoTo PROC_ERR
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
    For i = 1 To xmlParts.Count
        partName = xmlParts(i).NamespaceURI
        If Not IsPartInCollection(essentialParts, partName) Then
            essentialParts.Add xmlParts(i), partName
        Else
            duplicateParts.Add xmlParts(i), partName
        End If
    Next i

    ' Remove duplicate parts
    For j = 1 To duplicateParts.Count
        duplicateParts(j).Delete
    Next j

    ' Print names of essential and duplicate parts
    Debug.Print "Essential CustomXML Parts:"
    For i = 1 To essentialParts.Count
        Debug.Print essentialParts(i).NamespaceURI
    Next i

    Debug.Print "Duplicate CustomXML Parts:"
    For j = 1 To duplicateParts.Count
        Debug.Print duplicateParts(j).NamespaceURI
    Next j

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RemoveDuplicateCustomXMLParts of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Private Function IsPartInCollection(col As Collection, partName As String) As Boolean
    On Error GoTo PROC_ERR
    Dim i As Integer
    IsPartInCollection = False
    For i = 1 To col.Count
        If col(i).NamespaceURI = partName Then
            IsPartInCollection = True
            GoTo PROC_EXIT
        End If
    Next i

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsPartInCollection of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Function

Public Sub DeleteCustomUIXML()
    On Error GoTo PROC_ERR
    Dim xmlPart As customXMLPart
    Dim xmlParts As CustomXMLParts
    Dim i As Integer

    Set xmlParts = ActiveDocument.CustomXMLParts

    ' Loop through all CustomXMLParts to find and delete the customUI parts
    For i = xmlParts.Count To 1 Step -1
        Set xmlPart = xmlParts(i)
        If xmlPart.NamespaceURI = "http://schemas.microsoft.com/office/2006/01/customui" Or _
                xmlPart.NamespaceURI = "http://schemas.microsoft.com/office/2009/07/customui" Then
            xmlPart.Delete
        End If
    Next i

    MsgBox "CustomUI XML parts deleted successfully!"

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure DeleteCustomUIXML of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

' ========================================================================================
' Function:     GetColorNameFromHex
' Purpose:      Translates a hexadecimal color string (e.g., "#FF0000") into a human-readable
'               color name. Useful for diagnostics, audit logs, or UI labeling in scripts.
' Inputs:       hexColor [String] - A color code in hexadecimal format, e.g., "#FF0000"
' Returns:      [String] - The corresponding color name, or "Unknown Color" if not matched.
' Author:       Peter
' Last Updated: 2025-08-02
' Notes:        - Hex code is normalized to uppercase for consistent comparison.
'               - Expand CASE block as needed for additional named colors.
' ========================================================================================
Private Function GetColorNameFromHex(hexColor As String) As String
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

' =================================================================================================
' Subroutine:   ListAndCountFontColors
' Purpose:      Iterates over all words in the active Word document, extracts the RGB font color,
'               and tallies occurrences per unique color. Outputs formatted results to the console
'               including color name via GetColorNameFromHex.
' Inputs:       None (operates on ActiveDocument)
' Outputs:      Debug.Print output of RGB, Hex, Count, and resolved color name
' Dependencies: Requires GetColorNameFromHex(hexColor As String) function to be present
' Author:       Peter
' Last Updated: 2025-08-02
' Notes:        - Hex keys are zero-padded for consistency
'               - Font.Color property is bitmasked and decomposed manually
'               - Does not account for style inheritance or partial selections
'               - Expansion possible to handle suffix-aware grouping or paragraph-level aggregation
' =================================================================================================
Public Sub ListAndCountFontColors()
    On Error GoTo PROC_ERR
    Dim rng As Word.Range
    Dim colorDict As Object
    Dim colorKey As Variant
    Dim colorCount As Long
    Dim r As Long, g As Long, b As Long

    ' Create a dictionary to store color counts
    Set colorDict = CreateObject("Scripting.Dictionary")

    ' Loop through each word in the document
    For Each rng In ActiveDocument.words
        ' Get the RGB values of the font color
        r = (rng.Font.color And &HFF)
        g = (rng.Font.color \ &H100 And &HFF)
        b = (rng.Font.color \ &H10000 And &HFF)

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
        g = CLng("&H" & Mid$(colorKey, 3, 2))
        b = CLng("&H" & Right(colorKey, 2))

        Debug.Print "Color: RGB(" & r & ", " & g & ", " & b & ") - Hex: #" & colorKey & " - Count: " & colorCount & " - " & GetColorNameFromHex("#" & colorKey)
    Next colorKey

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ListAndCountFontColors of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub GetVerticalPositionOfCursorParagraph()
' Get the position of the para where the cursor is
    On Error GoTo PROC_ERR
    Dim doc As Document
    Dim rng As Word.Range
    Dim paraPos As Single

    Set doc = ActiveDocument
    Set rng = Selection.Paragraphs(1).Range

    ' Get the vertical position of the paragraph relative to the page
    paraPos = rng.Information(wdVerticalPositionRelativeToPage)

    ' Display the vertical position
    MsgBox "Vertical Position of the paragraph with the cursor: " & paraPos & " points"

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GetVerticalPositionOfCursorParagraph of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub FindFirstSectionWithDifferentFirstPage()
    On Error GoTo PROC_ERR
    Dim sec As Word.Section
    Dim i As Long

    For i = 1 To ActiveDocument.Sections.Count
        Set sec = ActiveDocument.Sections(i)

        ' Check if Different First Page is enabled
        If sec.PageSetup.DifferentFirstPageHeaderFooter = True Then
            ' Select the header of the first page in this section
            sec.Headers(wdHeaderFooterFirstPage).Range.Select

            MsgBox "Found in Section " & i & ": 'Different First Page' is enabled.", vbInformation
            GoTo PROC_EXIT
        End If
    Next i

    MsgBox "No sections with 'Different First Page' found.", vbInformation

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FindFirstSectionWithDifferentFirstPage of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub FindFirstPageWithEmptyHeader()
    On Error GoTo PROC_ERR
    Dim sec As Word.Section
    Dim hdr As HeaderFooter
    Dim hdrText As String
    Dim i As Long
    Dim hdrType As Variant  ' Must be Variant for Array() to work

    For i = 1 To ActiveDocument.Sections.Count
        Set sec = ActiveDocument.Sections(i)

        For Each hdrType In Array(wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages)
            Set hdr = sec.Headers(hdrType)

            If hdr.Exists And Not hdr.LinkToPrevious Then
                hdrText = Trim(hdr.Range.Text)

                If Right(hdrText, 1) = Chr(13) Then
                    hdrText = Left(hdrText, Len(hdrText) - 1)
                End If

                If hdrText = "" Then
                    hdr.Range.Select
                    MsgBox "Found empty header in Section " & i & " (" & HeaderTypeName(hdrType) & ").", vbInformation
                    GoTo PROC_EXIT
                End If
            End If
        Next hdrType
    Next i

    MsgBox "No empty headers found.", vbInformation

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FindFirstPageWithEmptyHeader of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Private Function HeaderTypeName(hdrType As Variant) As String
    Select Case hdrType
        Case wdHeaderFooterPrimary: HeaderTypeName = "Primary"
        Case wdHeaderFooterFirstPage: HeaderTypeName = "First Page"
        Case wdHeaderFooterEvenPages: HeaderTypeName = "Even Pages"
        Case Else: HeaderTypeName = "Unknown"
    End Select
End Function

Public Sub OptimizedListFontsInDocument()
    On Error GoTo PROC_ERR
    Dim fontList As New Collection
    Dim doc As Document
    Dim para As Word.Paragraph
    Dim rng As Word.Range
    Dim fontName As String
    Dim i As Integer

    Set doc = ActiveDocument

    ' Loop through each paragraph in the document
    For Each para In doc.Paragraphs
        Set rng = para.Range
        fontName = rng.Font.Name
        On Error Resume Next
        ' Add unique fonts to the collection
        fontList.Add fontName, fontName
        On Error GoTo 0
        On Error GoTo PROC_ERR
    Next para

    ' Display the fonts in a message box
    Dim fontOutput As String
    fontOutput = "Fonts used in the document:" & vbCrLf
    For i = 1 To fontList.Count
        fontOutput = fontOutput & "- " & fontList(i) & vbCrLf
    Next i
    'MsgBox fontOutput, vbInformation, "Fonts in Document"
    Debug.Print fontOutput

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OptimizedListFontsInDocument of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub FindGentiumFromParagraph()
    On Error GoTo PROC_ERR
    Dim startParaNum As Long
    Dim para As Word.Paragraph
    Dim rng As Word.Range
    Dim charRange As Word.Range
    Dim i As Long, p As Long
    Dim totalParas As Long

    ' Ask user where to start
    startParaNum = val(InputBox("Enter paragraph number to start from:", "Start From Paragraph", 1))
    If startParaNum < 1 Then GoTo PROC_EXIT

    totalParas = ActiveDocument.Paragraphs.Count
    If startParaNum > totalParas Then
        MsgBox "There are only " & totalParas & " paragraphs in the document.", vbExclamation
        GoTo PROC_EXIT
    End If

    p = 0
    For Each para In ActiveDocument.Paragraphs
        p = p + 1
        If p < startParaNum Then GoTo NextPara

        Set rng = para.Range
        rng.End = rng.End - 1 ' Exclude paragraph mark

        For i = 1 To rng.Characters.Count Step 10 ' Check every 10 chars
            Set charRange = rng.Characters(i)
            If charRange.Font.Name = "Gentium" Then
                charRange.Select
                MsgBox "Found Gentium font at paragraph " & p, vbInformation
                Application.StatusBar = False
                GoTo PROC_EXIT
            End If
        Next i

        If p Mod 100 = 0 Then
            Application.StatusBar = "Scanning paragraph " & p & " of " & totalParas & "..."
            DoEvents
        End If

NextPara:
    Next para

    MsgBox "Gentium font not found starting from paragraph " & startParaNum & ".", vbExclamation

PROC_EXIT:
    Application.StatusBar = False
    Exit Sub
PROC_ERR:
    Application.StatusBar = False
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FindGentiumFromParagraph of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub GoToParagraph()
    On Error GoTo PROC_ERR
    Dim paraNum As Integer
    paraNum = (InputBox("Enter paragraph number:", "Goto Paragraph Number", 1))
    ActiveDocument.Paragraphs(paraNum).Range.Select

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GoToParagraph of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub ListNonMainFonts_ByParagraph()
    On Error GoTo PROC_ERR
    Dim fontDict As Object
    Set fontDict = CreateObject("Scripting.Dictionary")

    Dim storyRange As Word.Range
    Dim para As Word.Paragraph
    Dim fontName As String
    Dim fontCount As Long
    Dim scannedParas As Long

    Application.ScreenUpdating = False
    Application.StatusBar = "Scanning fonts outside main text..."

    For Each storyRange In ActiveDocument.StoryRanges
        If storyRange.StoryType <> wdMainTextStory Then
            Do
                For Each para In storyRange.Paragraphs
                    scannedParas = scannedParas + 1
                    fontName = para.Range.Font.Name
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

    If fontDict.Count = 0 Then
        MsgBox "No fonts found outside main text.", vbInformation
    Else
        Dim Output As String, key As Variant
        Output = "Fonts outside main document text:" & vbCrLf & vbCrLf
        For Each key In fontDict.Keys
            Output = Output & "- " & key & vbCrLf
        Next key
        'MsgBox output, vbInformation, "Non-Main Fonts"
        Debug.Print Output
    End If

PROC_EXIT:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub
PROC_ERR:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ListNonMainFonts_ByParagraph of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub TestComp()
    On Error GoTo PROC_ERR
    CompareDocuments "C:\adaept\aeBibleClass\Peter-USE REFINED English Bible CONTENTS.docx", "C:\Users\peter\OneDrive\Documents\Peter-USE REFINED English Bible CONTENTS - Copy (49).docx"

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TestComp of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Private Sub CompareDocuments(original As String, modified As String)
' e.g. original = "C:\Path\To\Original.docx"
' e.g. "C:\Path\To\Modified.docx"
' - Original Document - The initial version of the document before changes were made.
' - Modified Document - The updated version that includes changes.
' - Comparison Document - The newly generated document that highlights differences between the original and modified versions.
' - The **comparison document** is a completely **new document** that shows changes such as insertions, deletions, and formatting modifications.
' - The **original** and **modified** documents remain **unchanged**-Word does **not** alter them.
' wdGranularityWordLevel
' - CompareFormatting (True) - Marks differences in formatting (e.g., font changes, bold/italic modifications).
' - CompareCaseChanges (True) - Highlights changes in letter case (e.g., "word" vs. "Word").
' - CompareWhitespace (True) - Tracks differences in spaces, paragraph breaks, and other whitespace variations.
' - CompareTables (True) - Compares changes within tables, including cell modifications.
' These options allow for a detailed comparison of documents, ensuring that even subtle changes are detected.
    On Error GoTo PROC_ERR
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

    ' Save comparison Result
    docComparison.SaveAs filePath & "\Comparison.docx"

    MsgBox "Comparison complete! See the document for tracked changes."

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure CompareDocuments of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub CountSearchHits()
    On Error GoTo PROC_ERR
    Dim searchTerm As String
    Dim Count As Long
    Dim rng As Word.Range

    searchTerm = InputBox("Enter the text to search for:")
    If Len(searchTerm) = 0 Then GoTo PROC_EXIT

    Count = 0
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = searchTerm
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False

        Do While .Execute
            Count = Count + 1
            rng.Collapse wdCollapseEnd
        Loop
    End With

    MsgBox "Found " & Count & " instance(s) of '" & searchTerm & "'.", vbInformation

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure CountSearchHits of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub PrintHeading1sByLogicalPage()
    On Error GoTo PROC_ERR
    Dim i As Long
    Dim maxPage As Long
    Dim pageRange As Word.Range
    Dim para As Word.Paragraph
    Dim headingText As String
    Dim foundHeading As Boolean

    maxPage = ActiveDocument.Range.Information(wdNumberOfPagesInDocument)

    Debug.Print "=== Heading 1s by Logical Page (GoTo ^H) ==="

    For i = 1 To maxPage
        Set pageRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=i)
        Set pageRange = pageRange.GoTo(What:=wdGoToBookmark, name:="\page") ' Get full page range

        foundHeading = False

        For Each para In pageRange.Paragraphs
            If para.style = "Heading 1" Then
                headingText = Replace(para.Range.Text, vbCr, "")
                Debug.Print "Logical Page " & i & ": " & headingText
                foundHeading = True
                Exit For ' Only report first Heading 1 per page
            End If
        Next para

        If Not foundHeading Then
            ' Optional: Debug.Print "Logical Page " & i & ": No Heading 1"
        End If
    Next i

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure PrintHeading1sByLogicalPage of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub FixAndDiagnoseFootnoteReferences()
    On Error GoTo PROC_ERR
    Dim doc As Document
    Dim fn As footnote
    Dim fnRef As Word.Range
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
            fnRef.Font.Reset
            fnRef.style = doc.Styles("Footnote Reference")
            With fnRef.Font
                .Name = "Segoe UI"
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
        Selection.MoveRight Unit:=wdCharacter, Count:=firstIncorrectPos
        Selection.Select
    Else
        MsgBox "All footnote reference formatting is correct."
    End If

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FixAndDiagnoseFootnoteReferences of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Private Function IsCorrectFootnoteFormat(rng As Word.Range, ByRef mismatch As String) As Boolean
    On Error GoTo PROC_ERR
    mismatch = ""
    IsCorrectFootnoteFormat = True
    With rng.Font
        If rng.style <> "Footnote Reference" Then
            mismatch = mismatch & " - Style: " & rng.style & vbCrLf
            IsCorrectFootnoteFormat = False
        End If
        If .Name <> "Segoe UI" Then
            mismatch = mismatch & " - Font Name: " & .Name & vbCrLf
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

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsCorrectFootnoteFormat of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Function

Public Sub FixFootnoteNumberStyleInText()
    On Error GoTo PROC_ERR
    Dim fn As footnote
    Dim paraRange As Word.Range
    Dim firstRun As Word.Range

    For Each fn In ActiveDocument.Footnotes
        Set paraRange = fn.Range.Paragraphs(1).Range
        Set firstRun = paraRange.words(1) ' Usually the footnote number

        ' Apply Footnote Reference style
        firstRun.style = ActiveDocument.Styles("Footnote Reference")
    Next fn

    MsgBox "Footnote Reference style reapplied to footnote numbers in footnote text.", vbInformation

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FixFootnoteNumberStyleInText of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Private Sub ReportPageLayoutMetrics(pageNum As Long)
    On Error GoTo PROC_ERR
    Dim pgRange As Word.Range
    Dim sectionSetup As PageSetup
    Dim numCols As Integer, isEven As Boolean
    Dim gutter As Single, pageWidth As Single
    Dim leftMargin As Single, rightMargin As Single
    Dim Spacing As Single, columnWidth As Single
    Dim sectionLeft As Single, colStart As Single
    Dim logBuffer As String, i As Integer

    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNum))
    Set sectionSetup = pgRange.Sections(1).PageSetup

    numCols = sectionSetup.TextColumns.Count
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
        logBuffer = logBuffer & "> Column " & (i + 1) & " starts at: " & Format(colStart, "0.0") & vbCrLf
    Next i

    logBuffer = logBuffer & "=== End of Layout Report ==="
    Debug.Print logBuffer
    MsgBox "Layout metrics for page " & pageNum & " printed to Immediate window.", vbInformation

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ReportPageLayoutMetrics of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub ReportDigitAtCursor_Diagnostics()
    On Error GoTo PROC_ERR
    Dim selRange As Word.Range, ch As Word.Range, prefix As Word.Range
    Dim txt As String, style As String
    Dim posX As Single, posY As Single
    Dim prefixTxt As String, prefixStyle As String, prefixAsc As Variant
    Dim prefixX As Single, prefixY As Single

    Set selRange = Selection.Range
    Set ch = ActiveDocument.Range(selRange.Start, selRange.Start + 1)
    txt = ch.Text
    style = ch.style.NameLocal
    posX = ch.Information(wdHorizontalPositionRelativeToPage)
    posY = ch.Information(wdVerticalPositionRelativeToPage)

    Debug.Print "=== Character at Cursor ==="
    Debug.Print "Value: '" & txt & "' | ASCII: " & AscW(txt)
    Debug.Print "Style: " & style
    Debug.Print "Font Color: " & ch.Font.color & " (RGB: " & _
                RGBToString(ch.Font.color) & ")"
    Debug.Print "Position: X=" & Format(posX, "0.0") & " pts, Y=" & Format(posY, "0.0") & " pts"

    If ch.Start > 0 Then
        Set prefix = ActiveDocument.Range(ch.Start - 1, ch.Start)
        prefixTxt = prefix.Text
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

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ReportDigitAtCursor_Diagnostics of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Private Function RGBToString(rgbVal As Long) As String
    RGBToString = "(" & (rgbVal And &HFF) & "," & ((rgbVal \ 256) And &HFF) & "," & ((rgbVal \ 65536) And &HFF) & ")"
End Function

Public Sub ReportDigitAtCursor_Diagnostics_Expanded()
    On Error GoTo PROC_ERR
    Dim rng As Word.Range
    Set rng = Selection.Range
    If rng.Characters.Count = 0 Then
        MsgBox "No character selected.", vbExclamation
        GoTo PROC_EXIT
    End If

    Dim ch As Word.Range
    Set ch = rng.Characters(1)
    Dim txt As String: txt = ch.Text
    Dim ascCode As Long: ascCode = AscW(txt)
    Dim fontNameAscii As String: fontNameAscii = ch.Font.NameAscii
    Dim fontNameFarEast As String: fontNameFarEast = ch.Font.NameFarEast
    Dim fontNameOther As String: fontNameOther = ch.Font.NameOther
    Dim fontSize As Single: fontSize = ch.Font.Size
    Dim fontColor As Long: fontColor = ch.Font.color
    Dim StyleName As String: StyleName = ch.style.NameLocal
    Dim baseStyle As String
    On Error Resume Next
    baseStyle = ch.style.baseStyle
    On Error GoTo 0
    On Error GoTo PROC_ERR

    Debug.Print "=== Character at Cursor ==="
    Debug.Print "Value: '" & txt & "' | ASCII: " & ascCode
    Debug.Print "Style: " & StyleName
    Debug.Print "Base Style: " & IIf(baseStyle = "", "(none)", baseStyle)

    Debug.Print "Font Names:"
    Debug.Print "> NameAscii: " & fontNameAscii
    Debug.Print "> NameFarEast: " & fontNameFarEast
    Debug.Print "> NameOther: " & fontNameOther
    Debug.Print "Font Size: " & fontSize & " pt"
    Debug.Print "Font Color: " & fontColor & " (RGB: " & _
        (fontColor Mod 256) & "," & ((fontColor \ 256) Mod 256) & "," & (fontColor \ 65536) & ")"
    Debug.Print "Bold: " & ch.Font.Bold & " | Italic: " & ch.Font.Italic & " | Underline: " & ch.Font.Underline

    Debug.Print "--- Prefix (1 char before) ---"
    If ch.Start > 1 Then
        Dim prefix As Word.Range
        Set prefix = ActiveDocument.Range(ch.Start - 1, ch.Start)
        Debug.Print "Value: '" & prefix.Text & "' | ASCII: " & AscW(prefix.Text)
        Debug.Print "Style: " & prefix.style.NameLocal
        Debug.Print "Font Name: " & prefix.Font.Name
        Debug.Print "Font Color: " & prefix.Font.color & " (RGB: " & _
            (prefix.Font.color Mod 256) & "," & ((prefix.Font.color \ 256) Mod 256) & "," & (prefix.Font.color \ 65536) & ")"
    Else
        Debug.Print "(No character before this one.)"
    End If

    MsgBox "Expanded character diagnostics logged.", vbInformation

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ReportDigitAtCursor_Diagnostics_Expanded of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub LogExpandedMarkerContext()
    On Error GoTo PROC_ERR
    Dim sel As Word.Range: Set sel = Selection.Range
    Dim i As Long, chCount As Long
    Dim contextText As String, contextAscii As String, contextHex As String

    chCount = sel.Characters.Count
    Debug.Print "=== Marker Diagnostic ==="
    Debug.Print "Selection Start=" & sel.Start & " | End=" & sel.End
    Debug.Print "Selection Text='" & Replace(sel.Text, vbCr, "[CR]") & "'"

    For i = 1 To chCount
        Dim ch As String: ch = sel.Characters(i).Text
        Dim ascVal As Integer: ascVal = Asc(ch)
        Dim hexVal As String: hexVal = Hex(ascVal)

        contextText = "[" & i & "] '" & Replace(ch, vbCr, "[CR]") & "'"
        contextAscii = " ASCII=" & ascVal
        contextHex = " Hex=" & hexVal

        Debug.Print contextText & contextAscii & contextHex
    Next i

    Debug.Print "Style: " & sel.style & " | Font: " & sel.Font.Name
    Debug.Print "=== End of Diagnostic ===" & vbCrLf

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LogExpandedMarkerContext of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub FindInvisibleFormFeeds_InPages(startPage As Long)
    On Error GoTo PROC_ERR
    Dim para As Word.Paragraph, rng As Word.Range
    Dim pgNum As Long
    Dim i As Long, pgTarget As Long

    pgTarget = startPage + 9
    Debug.Print "=== Scanning for Chr(12) from Page " & startPage & " to " & pgTarget & " ==="

    For Each para In ActiveDocument.Paragraphs
        Set rng = para.Range
        pgNum = rng.Information(wdActiveEndPageNumber)

        If pgNum >= startPage And pgNum <= pgTarget Then
            If InStr(rng.Text, Chr(12)) > 0 Then
                Debug.Print "[Page " & pgNum & "] Chr(12) found at Start=" & rng.Start
                Debug.Print "Text='" & Replace(rng.Text, Chr(12), "[FF]") & "'"
                Debug.Print "Style=" & rng.style & " | Font=" & rng.Font.Name
            End If
        End If
    Next para

    Debug.Print "=== End of Scan ===" & vbCrLf

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FindInvisibleFormFeeds_InPages of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub AuditVerseMarkers_VerifyMergedNumberPrefix_WithContext(pageNum As Long)
    On Error GoTo PROC_ERR
    Dim pgRange As Word.Range, ch As Word.Range, scanRange As Word.Range
    Dim pageStart As Long, pageEnd As Long
    Dim logBuffer As String
    Dim chapterMarker As String, verseDigits As String, combinedNumber As String
    Dim markerStart As Long, markerEnd As Long, verseEnd As Long
    Dim digitPosX As Single, digitPosY As Single
    Dim wordRange As Word.Range, token As Word.Range, nextWords As String, wCount As Integer

    logBuffer = "=== Chapter Verse Visual Number Check on Page " & pageNum & " ===" & vbCrLf

    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNum))
    pageStart = pgRange.Start
    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNum + 1))
    pageEnd = pgRange.Start - 1
    Set pgRange = ActiveDocument.Range(pageStart, pageEnd)

    Dim i As Long
    i = pageStart
    Do While i < pageEnd
        Set ch = ActiveDocument.Range(i, i + 1)
        If Len(Trim(ch.Text)) = 1 And IsNumeric(ch.Text) And ch.style.NameLocal = "Chapter Verse marker" And ch.Font.color = RGB(255, 165, 0) Then
            chapterMarker = ch.Text
            markerStart = i
            markerEnd = i + 1
            Do While markerEnd < pageEnd
                Set scanRange = ActiveDocument.Range(markerEnd, markerEnd + 1)
                If Len(Trim(scanRange.Text)) = 1 And IsNumeric(scanRange.Text) Then
                    If scanRange.style.NameLocal = "Chapter Verse marker" And scanRange.Font.color = RGB(255, 165, 0) Then
                        chapterMarker = chapterMarker & scanRange.Text
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
                Set scanRange = ActiveDocument.Range(verseEnd, verseEnd + 1)
                If Len(Trim(scanRange.Text)) = 1 And IsNumeric(scanRange.Text) Then
                    If scanRange.style.NameLocal = "Verse marker" And scanRange.Font.color = RGB(80, 200, 120) Then
                        verseDigits = verseDigits & scanRange.Text
                        verseEnd = verseEnd + 1
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop

            nextWords = ""
            Set wordRange = ActiveDocument.Range(verseEnd, verseEnd + 80)
            wCount = 0
            For Each token In wordRange.words
                If token.Text Like "*^13*" Then Exit For
                If Trim(token.Text) <> "" Then
                    nextWords = nextWords & Trim(token.Text) & " "
                    wCount = wCount + 1
                    If wCount = 2 Then Exit For
                End If
            Next token

            If Len(verseDigits) = 0 Then
                logBuffer = logBuffer & "! Chapter '" & chapterMarker & "' @ X=" & Format(digitPosX, "0.0") & _
                    " | No styled Verse marker digits found | Next words: �" & Trim(nextWords) & "�" & vbCrLf
            Else
                combinedNumber = chapterMarker & verseDigits
                If Left(combinedNumber, Len(chapterMarker)) = chapterMarker Then
                    logBuffer = logBuffer & "* Chapter '" & chapterMarker & "' ? Verse '" & combinedNumber & "' @ X=" & Format(digitPosX, "0.0") & _
                        " | ? Valid | Next words: �" & Trim(nextWords) & "�" & vbCrLf
                Else
                    logBuffer = logBuffer & "! Chapter '" & chapterMarker & "' ? Verse '" & combinedNumber & "' @ X=" & Format(digitPosX, "0.0") & _
                        " | ? Mismatch | Next words: �" & Trim(nextWords) & "�" & vbCrLf
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

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure AuditVerseMarkers_VerifyMergedNumberPrefix_WithContext of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub ReportAllMarkers_CondensedDiagnostics(pageNum As Long)
    On Error GoTo PROC_ERR
    Dim pgRange As Word.Range, ch As Word.Range, scanRange As Word.Range
    Dim pageStart As Long, pageEnd As Long
    Dim txt As String, StyleName As String
    Dim fontName As String, fontSize As Single, fontColor As Long
    Dim charPosX As Single, charPosY As Single
    Dim digitBlock As String, blockStyle As String, blockColor As Long
    Dim blockStart As Long, blockEnd As Long, logBuffer As String

    logBuffer = "=== Marker Summary for Page " & pageNum & " ===" & vbCrLf

    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNum))
    pageStart = pgRange.Start
    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNum + 1))
    pageEnd = pgRange.Start - 1
    Set pgRange = ActiveDocument.Range(pageStart, pageEnd)

    Dim i As Long
    i = pageStart
    Do While i < pageEnd
        Set ch = ActiveDocument.Range(i, i + 1)
        txt = Trim(ch.Text)
        StyleName = ch.style.NameLocal

        If Len(txt) = 1 And IsNumeric(txt) Then
            If StyleName = "Chapter Verse marker" Or StyleName = "Verse marker" Then
                digitBlock = txt
                blockStyle = StyleName
                blockColor = ch.Font.color
                blockStart = i
                blockEnd = i + 1

                Do While blockEnd < pageEnd
                    Set scanRange = ActiveDocument.Range(blockEnd, blockEnd + 1)
                    If Len(Trim(scanRange.Text)) = 1 And IsNumeric(scanRange.Text) Then
                        If scanRange.style.NameLocal = blockStyle And scanRange.Font.color = blockColor Then
                            digitBlock = digitBlock & scanRange.Text
                            blockEnd = blockEnd + 1
                        Else
                            Exit Do
                        End If
                    Else
                        Exit Do
                    End If
                Loop

                Set ch = ActiveDocument.Range(blockStart, blockStart + 1)
                fontName = ch.Font.Name
                fontSize = ch.Font.Size
                charPosX = ch.Information(wdHorizontalPositionRelativeToPage)
                charPosY = ch.Information(wdVerticalPositionRelativeToPage)

                logBuffer = logBuffer & "[" & IIf(blockStyle = "Chapter Verse marker", "Chapter", "Verse") & "] '" & digitBlock & "' @ X=" & Format(charPosX, "0.0") & ", Y=" & Format(charPosY, "0.0") & _
                    " | Font: " & fontName & " " & fontSize & "pt | RGB: (" & (blockColor Mod 256) & "," & ((blockColor \ 256) Mod 256) & "," & (blockColor \ 65536) & ")" & _
                    " | Pos: " & blockStart & "�" & (blockEnd - 1) & vbCrLf

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

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ReportAllMarkers_CondensedDiagnostics of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Private Sub RepairWrappedVerseMarkers_MergedPrefix_ByColumnContext_SinglePage(pageNum As Long, ByRef fixCount As Long)
    On Error GoTo PROC_ERR
    ' Same logic as full macro, but suppresses MsgBox and passes fixCount by reference.
    ' Copy the full body from RepairWrappedVerseMarkers_MergedPrefix_ByColumnContext here
    ' And replace `MsgBox` line with: fixCount = fixCount
    Dim pgRange As Word.Range, ch As Word.Range, scanRange As Word.Range, prefixCh As Word.Range
    Dim pageStart As Long, pageEnd As Long
    Dim chapterMarker As String, verseDigits As String, combinedNumber As String
    Dim markerStart As Long, markerEnd As Long, verseEnd As Long
    Dim prefixTxt As String, prefixStyle As String, prefixAsc As Variant
    Dim prefixY As Single, digitY As Single, digitX As Single
    Dim nextWords As String, lookAhead As Word.Range, token As Word.Range, wCount As Integer
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
        Set ch = ActiveDocument.Range(i, i + 1)
        If Len(Trim(ch.Text)) = 1 And IsNumeric(ch.Text) And ch.style.NameLocal = "Chapter Verse marker" And ch.Font.color = RGB(255, 165, 0) Then
            ' Assemble chapter marker block
            chapterMarker = ch.Text
            markerStart = i
            markerEnd = i + 1
            Do While markerEnd < pageEnd
                Set scanRange = ActiveDocument.Range(markerEnd, markerEnd + 1)
                If Len(Trim(scanRange.Text)) = 1 And IsNumeric(scanRange.Text) Then
                    If scanRange.style.NameLocal = "Chapter Verse marker" And scanRange.Font.color = RGB(255, 165, 0) Then
                        chapterMarker = chapterMarker & scanRange.Text
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
                Set scanRange = ActiveDocument.Range(verseEnd, verseEnd + 1)
                If Len(Trim(scanRange.Text)) = 1 And IsNumeric(scanRange.Text) Then
                    If scanRange.style.NameLocal = "Verse marker" And scanRange.Font.color = RGB(80, 200, 120) Then
                        verseDigits = verseDigits & scanRange.Text
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
                    Set prefixCh = ActiveDocument.Range(markerStart - 1, markerStart)
                    prefixTxt = prefixCh.Text
                    prefixStyle = prefixCh.style.NameLocal
                    prefixAsc = AscW(prefixTxt)
                    prefixY = prefixCh.Information(wdVerticalPositionRelativeToPage)

                    If (prefixAsc = 32 Or prefixAsc = 160) And prefixStyle = "Normal" Then
                        If Abs(prefixY - digitY) < 25 Then
                            nextWords = ""
                            Set lookAhead = ActiveDocument.Range(verseEnd, verseEnd + 80)
                            wCount = 0
                            For Each token In lookAhead.words
                                If token.Text Like "*^13*" Then Exit For
                                If Trim(token.Text) <> "" Then
                                    nextWords = nextWords & Trim(token.Text) & " "
                                    wCount = wCount + 1
                                    If wCount = 2 Then Exit For
                                End If
                            Next token

                            ' Column edge logic
                            If digitX < 50 Then
                                prefixCh.Text = vbCr
                                logBuffer = logBuffer & "- Repaired prefix before '" & combinedNumber & "' @ X=" & Format(digitX, "0.0") & " | Break inserted | Next words: >" & Trim(nextWords) & "<" & vbCrLf
                            Else
                                prefixCh.Text = ""
                                logBuffer = logBuffer & "- Removed space before '" & combinedNumber & "' @ X=" & Format(digitX, "0.0") & " | No break | Next words: >" & Trim(nextWords) & "<" & vbCrLf
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

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RepairWrappedVerseMarkers_MergedPrefix_ByColumnContext_SinglePage of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Private Sub SmartPrefixRepairOnPage_WithDiagnostics(pgNum As Long, ByRef spaceCount As Long, ByRef breakCount As Long)
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

Public Sub SmartyOne()
    Dim sCount As Long, bCount As Long
    Call SmartPrefixRepairOnPage(235, sCount, bCount)
End Sub

Private Sub SmartPrefixRepairOnPage(pgNum As Long, ByRef spaceCount As Long, ByRef breakCount As Long)
    On Error GoTo PROC_ERR
    Dim para As Word.Paragraph
    Dim rng As Word.Range
    Dim markerText As String
    Dim didRepair As Boolean
    Dim paraStyle As String
    Dim ascii12Count As Long
    Dim missing160Count As Long

    Debug.Print "=== Smart Prefix Repair on Page " & pgNum & " ==="

    For Each para In ActiveDocument.Paragraphs
        Set rng = para.Range
        paraStyle = rng.style

        ' Only process green "Verse marker" paragraphs
        If InStr(paraStyle, "Verse marker") > 0 Then
            markerText = rng.Text

            ' Skip and Count layout wrappers: lone Chr(12)
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
                    ch = Mid$(markerText, i, 1)
                    ascVal = Asc(ch)
                    Debug.Print "    Char " & i & ": '" & Replace(ch, vbCr, "[CR]") & "' | ASCII=" & ascVal & " | Hex=" & Hex(ascVal)
                Next i

                ' Check if last character is missing expected Chr(160)
                If Len(markerText) = 0 Or Asc(Right(markerText, 1)) <> 160 Then
                    missing160Count = missing160Count + 1
                End If

                Debug.Print "  Style=" & rng.style & " | Font=" & rng.Font.Name & " | Start=" & rng.Start & " | Page=" & pgNum
            End If
        End If
NextPara:
    Next para

    Debug.Print "Chr(12) marker Count on Page " & pgNum & ": " & ascii12Count
    Debug.Print "Missing Chr(160) Count on Page " & pgNum & ": " & missing160Count
    Debug.Print "=== End of Repairs for Page " & pgNum & " ==="

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure SmartPrefixRepairOnPage of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub RunRepairWrappedVerseMarkers_Across10Pages_From(StartPageNum As Long)
    On Error GoTo PROC_ERR
    Const ForecastFile As String = "RepairRunnerForecast.txt"
    Dim sessionID As String: sessionID = "Session_" & Format(Now, "yyyymmdd_HHMMSS")
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
    resultRow = sessionID & "," & StartPageNum & ","
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

PROC_EXIT:
    Exit Sub
PROC_ERR:
    If Not ts Is Nothing Then On Error Resume Next: ts.Close: On Error GoTo 0
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RunRepairWrappedVerseMarkers_Across10Pages_From of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub RunRepairWrappedVerseMarkers_ForOnePage(pgNum As Long)
    On Error GoTo PROC_ERR
    Const ForecastFile As String = "RepairRunnerForecast.txt"
    Dim sessionID As String: sessionID = "Session_" & Format(Now, "yyyymmdd_HHMMSS")
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
    resultRow = sessionID & "," & pgNum & "," & tElapsed & "," & spaceCount & "," & breakCount
    Debug.Print resultRow
    ts.WriteLine resultRow
    ts.Close

PROC_EXIT:
    Exit Sub
PROC_ERR:
    If Not ts Is Nothing Then On Error Resume Next: ts.Close: On Error GoTo 0
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RunRepairWrappedVerseMarkers_ForOnePage of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub UnlinkHeadingNumbering()
    On Error GoTo PROC_ERR
    Dim para As Word.Paragraph

    For Each para In ActiveDocument.Paragraphs
        If para.style = ActiveDocument.Styles("Heading 1") _
        Or para.style = ActiveDocument.Styles("Heading 2") Then
            para.Range.ListFormat.RemoveNumbers
        End If
    Next

    MsgBox "Numbering removed from Heading 1 and Heading 2 paragraphs.", vbInformation

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure UnlinkHeadingNumbering of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Sub StartRepairTimingSession(StartPageNum As Long)
    On Error GoTo PROC_ERR
    Const ForecastFile As String = "RepairRunnerForecast.txt"
    Dim sessionID As String: sessionID = "Session_" & Format(Now, "yyyymmdd_HHMMSS")

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
    resultRow = sessionID & "," & StartPageNum & ","
    For i = 1 To 10: resultRow = resultRow & timeStamps(i) & ",": Next i
    resultRow = resultRow & totalTime & "," & forecastTime

    Debug.Print resultRow
    ts.WriteLine resultRow
    ts.Close

PROC_EXIT:
    Exit Sub
PROC_ERR:
    If Not ts Is Nothing Then On Error Resume Next: ts.Close: On Error GoTo 0
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure StartRepairTimingSession of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Private Sub DummyRepairPageTimerOnly(pgNum As Long)
    ' Placeholder page-level logic for timing only
    Dim dummyWait As Single
    dummyWait = Timer
    Do While Timer < dummyWait + 0.05: DoEvents: Loop
End Sub

Public Sub ReapplyTheFootersToAllFooters()
    On Error GoTo PROC_ERR
    Dim sec As Word.Section

    Dim hf As HeaderFooter
    Dim p As Word.Paragraph
    Dim prevStyle As String
    Dim asciiVal As Long
    Dim paraText As String

    Debug.Print "=== Reapply 'TheFooters' Style Start ==="

    For Each sec In ActiveDocument.Sections
        Debug.Print "SECTION " & sec.index

        For Each hf In sec.Footers
            If hf.Exists Then
                For Each p In hf.Range.Paragraphs
                    paraText = p.Range.Text
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

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ReapplyTheFootersToAllFooters of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

' =============================================================================
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
' =============================================================================
Public Sub GetHeadingDefinitionsWithDescriptions()
    On Error GoTo PROC_ERR
    Dim headingStyles As Variant
    headingStyles = Array("Heading 1", "Heading 2")

    Dim s As Word.Style
    Dim info As String
    Dim StyleName As Variant
    Dim alignValue As Integer
    Dim alignText As String

    Dim clr As Long
    Dim r As Long, g As Long, b As Long
    Dim hexColor As String

    For Each StyleName In headingStyles
        Set s = ActiveDocument.Styles(StyleName)
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

        clr = s.Font.color
        r = clr Mod 256
        g = (clr \ 256) Mod 256
        b = (clr \ 65536) Mod 256
        hexColor = "#" & Right$("0" & Hex(r), 2) & Right$("0" & Hex(g), 2) & Right$("0" & Hex(b), 2)

        info = "Style: " & StyleName & vbCrLf
        info = info & "  Font Name: " & s.Font.Name & vbCrLf
        info = info & "  Font Size: " & s.Font.Size & vbCrLf
        info = info & "  Bold: " & s.Font.Bold & vbCrLf
        info = info & "  Italic: " & s.Font.Italic & vbCrLf
        info = info & "  Color: " & clr & ", RGB(" & r & "," & g & "," & b & "), " & hexColor & vbCrLf
        info = info & "  Alignment: " & alignValue & " (" & alignText & ")" & vbCrLf
        info = info & "  Space Before: " & s.ParagraphFormat.SpaceBefore & vbCrLf
        info = info & "  Space After: " & s.ParagraphFormat.SpaceAfter & vbCrLf
        info = info & "  Line Spacing: " & s.ParagraphFormat.LineSpacing & vbCrLf
        info = info & "  Outline Level: " & s.ParagraphFormat.OutlineLevel & vbCrLf
        info = info & "  Keep With Next: " & s.ParagraphFormat.KeepWithNext & vbCrLf
        info = info & String(40, "-") & vbCrLf

        Debug.Print info
    Next StyleName

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GetHeadingDefinitionsWithDescriptions of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

'==============================================
' ValidateTaskInChangelogModule
'----------------------------------------------
' Purpose:
'   Validates that a permalink pointing to a line in the ChangeLog.bas
'   corresponds to a task tag (e.g. #293) that appears inside a delimited
'   block within the same module.
'
' Inputs:
'   - GitHub permalink with fragment ID (e.g. https://...#L17)
'   - Hardcoded path to ChangeLog.bas
'
' Behavior:
'   - Extracts line number from permalink
'   - Reads line N from the ChangeLog.bas file
'   - Parses the first #NNN task tag found on that line
'   - Scans ChangeLog.bas for `=============` block boundaries
'   - Confirms that the tag appears within one of those blocks
'
' Output:
'   - [OK] #NNN found within module block
'   - [FAIL] with specific reason (tag missing, not found, bad line number)
'
' Audit Notes:
'   - Logs only via Debug.Print (ASCII-safe, no UI interference)
'   - No assumptions about module layout beyond `=============` fences
'   - Does not modify any content � purely read-only audit
'   - Suitable for changelog integrity checks in version-controlled macros
'==============================================
Public Sub ValidateTaskInChangelogModule()
    On Error GoTo PROC_ERR
    Dim permalink As String
    permalink = InputBox("Paste permalink (with #Lnn):")

    Dim lineNum As Long
    lineNum = val(Split(permalink, "#L")(1))

    Dim changelogModulePath As String
    changelogModulePath = "C:\adaept\aeBibleClass\src\basChangeLog_aeBibleClass.bas" ' Change as needed

    Dim fs As Object: Set fs = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object: Set ts = fs.OpenTextFile(changelogModulePath, 1)

    Dim i As Long, lineText As String
    For i = 1 To lineNum
        If ts.AtEndOfStream Then
            Debug.Print "[FAIL] Line number too large"
            ts.Close
            GoTo PROC_EXIT
        End If
        lineText = ts.ReadLine
    Next
    ts.Close
    Set ts = Nothing

    ' Extract #NNN tag from line
    Dim tag As String, w
    For Each w In Split(lineText, " ")
        If Left(w, 1) = "#" And IsNumeric(Mid$(w, 2)) Then tag = w: Exit For
    Next

    If tag = "" Then
        Debug.Print "[FAIL] No #NNN task tag found at line " & lineNum
        GoTo PROC_EXIT
    End If

    ' Re-open and scan for matching task inside ========= block
    Set ts = fs.OpenTextFile(changelogModulePath, 1)
    Dim insideBlock As Boolean: insideBlock = False
    Dim found As Boolean: found = False

    Do While Not ts.AtEndOfStream
        Dim ln As String: ln = ts.ReadLine
        If InStr(ln, "===") > 0 Then
            insideBlock = Not insideBlock
        ElseIf insideBlock And InStr(ln, tag) > 0 Then
            found = True: Exit Do
        End If
    Loop
    ts.Close
    Set ts = Nothing

    If found Then
        Debug.Print "[OK] " & tag & " found within module block"
    Else
        Debug.Print "[FAIL] " & tag & " not found in any block"
    End If

PROC_EXIT:
    Exit Sub
PROC_ERR:
    If Not ts Is Nothing Then On Error Resume Next: ts.Close: On Error GoTo 0
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ValidateTaskInChangelogModule of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Private Function GetPaperSizeName(paperSizeValue As WdPaperSize) As String
    Select Case paperSizeValue
        Case wdPaperA4: GetPaperSizeName = "A4"
        Case wdPaperLetter: GetPaperSizeName = "Letter"
        Case wdPaperLegal: GetPaperSizeName = "Legal"
        Case 11: GetPaperSizeName = "B5 (JIS)" 'wdPaperB5Jis
        Case Else: GetPaperSizeName = "Other (" & paperSizeValue & ")"
    End Select
End Function

'==============================================================
' PrintCompactSectionLayoutInfo
' --------------------------------------------------------------
' Purpose : Generates a detailed layout report of all sections in
'           the active Word document, including orientation,
'           page size, column Count, margin settings, borders,
'           and section break types.
'
' Outputs : ASCII text file summarizing layout characteristics,
'           written to: C:\adaept\aeBibleClass\rpt\DocumentLayoutReport.txt
'
' Behavior:
' - Iterates through all sections using ActiveDocument.Sections.
' - Aggregates column counts, break types, border styles, and
'   various layout metrics per section.
' - Converts margin and spacing values from points to inches.
' - Uses silent output with Debug.Print for diagnostic logging.
' - Tracks aggregate layout counts (e.g., number of sections with
'   two columns or odd page breaks) and assigns them to globals.
'
' Notes   :
' - Requires external functions: GetPaperSizeName, GetBorderStyle,
'   and PointsToInches � these must be present for successful output.
' - Assumes target directory exists; no path creation is performed.
' - File output is line-oriented and portable for audit pipelines.
' - No content modification occurs; strictly layout reporting.
'
' Author  : Peter
' Last Modified : 20250731
'==============================================================
Public Sub PrintCompactSectionLayoutInfo()
    On Error GoTo PROC_ERR
    Dim sec As Word.Section
    Dim i As Long
    Dim nOneCol As Long, nTwoCol As Long
    Dim nEvenPageBreak As Long, nOddPageBreak As Long
    Dim nContinuousBreak As Long, nNewPageBreak As Long
    Dim outputFile As String
    Dim outputText As String
    Dim fileNum As Integer
    outputFile = "C:\adaept\aeBibleClass\rpt\DocumentLayoutReport.txt"  ' Change to desired path

    ' Open the text file to write
    fileNum = FreeFile
    Open outputFile For Output As #fileNum
    
    ' Write Header to the file
    outputText = "=== Layout Report ===" & vbCrLf
    outputText = outputText & "Doc: " & ActiveDocument.Name & vbCrLf
    outputText = outputText & "Total Sections: " & ActiveDocument.Sections.Count & vbCrLf & vbCrLf
    Print #fileNum, outputText

    For i = 1 To ActiveDocument.Sections.Count
        Set sec = ActiveDocument.Sections(i)

        outputText = "Section " & i & ": " & vbCrLf
        outputText = outputText & "Page: " & IIf(sec.PageSetup.Orientation = wdOrientPortrait, "Portrait", "Landscape") & ", " & _
                    "Size: " & GetPaperSizeName(sec.PageSetup.paperSize) & ", " & _
                    "Columns: " & sec.PageSetup.TextColumns.Count & vbCrLf
        If sec.PageSetup.TextColumns.Count > 1 Then nTwoCol = nTwoCol + 1 Else nOneCol = nOneCol + 1

        ' Margins
        outputText = outputText & "Margins (inches): " & _
                    "Top: " & PointsToInches(sec.PageSetup.TopMargin) & ", " & _
                    "Bottom: " & PointsToInches(sec.PageSetup.BottomMargin) & ", " & _
                    "Left: " & PointsToInches(sec.PageSetup.leftMargin) & ", " & _
                    "Right: " & PointsToInches(sec.PageSetup.rightMargin) & ", " & _
                    "Gutter: " & PointsToInches(sec.PageSetup.gutter) & vbCrLf

        ' Line Numbering
        If sec.PageSetup.LineNumbering.Active Then
            outputText = outputText & "Line Numbers: " & sec.PageSetup.LineNumbering.StartingNumber & ", " & _
                        "Increment: " & sec.PageSetup.LineNumbering.CountBy & vbCrLf
        End If

        ' Header/Footer settings
        outputText = outputText & "Header Distance: " & PointsToInches(sec.PageSetup.HeaderDistance) & ", " & _
                    "Footer Distance: " & PointsToInches(sec.PageSetup.FooterDistance) & vbCrLf

        ' Borders (if any)
        outputText = outputText & "Borders: " & _
                    "Top: " & GetBorderStyle(sec.Borders(wdBorderTop)) & ", " & _
                    "Bottom: " & GetBorderStyle(sec.Borders(wdBorderBottom)) & ", " & _
                    "Left: " & GetBorderStyle(sec.Borders(wdBorderLeft)) & ", " & _
                    "Right: " & GetBorderStyle(sec.Borders(wdBorderRight)) & vbCrLf

        ' Section Break Type
        Select Case sec.PageSetup.sectionStart
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
        Print #fileNum, outputText
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
    Print #fileNum, outputText

    ' Close the file
    Close #fileNum

    'MsgBox "Layout report saved to: " & outputFile, vbInformation
    Debug.Print "Layout report saved to: " & outputFile

PROC_EXIT:
    Exit Sub
PROC_ERR:
    If fileNum > 0 Then Close #fileNum
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure PrintCompactSectionLayoutInfo of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

' =========================
' Subroutine:   FlagEarlyBindingRoutines_LateBound
' Purpose:      Scans modules in .docm to flag early-bound object declarations.
'               Classifies hits as [EXTERNAL], [CUSTOM], [ENUM], [WORD], etc.
' Inputs:       IncludeWordTypes [Boolean], IncludeEnums [Boolean]
' Outputs:      Debug.Print log entries per flagged declaration
' Dependencies: Late binding only - safe for all environments
' Notes:        Suppresses primitive types, ambiguous declarations, and optionally Word enums
' Author:       Peter (collab w/ Copilot)
' Last Updated: 2025-08-02
' =========================
Public Sub FlagEarlyBindingRoutines_LateBound()
    On Error GoTo PROC_ERR
    Const IncludeWordTypes As Boolean = False
    Const IncludeEnums As Boolean = False

    Dim vbProj As Object, vbComps As Object, comp As Object, codeMod As Object
    Dim lineNum As Long, procName As String, procType As Long
    Dim codeLine As String, i As Long

    Dim baseTypes As Variant, wordTypes As Variant, knownEnums As Variant

    baseTypes = Array("As String", "As Integer", "As Long", "As Double", "As Boolean", "As Variant", _
                      "As Byte", "As Currency", "As Date", "As Object", "As Single")

    wordTypes = Array("As Word.Range", "As Word.Paragraph", "As Word.Section", "As Word.Style", "As Shape", "As Field", _
                      "As HeaderFooter", "As Footnote", "As EndNote", "As Table", "As Bookmark", _
                      "As Document", "As Collection", "As New Collection", "As Selection", _
                      "As customXMLPart", "As CustomXMLParts")

    knownEnums = Array("As WdColor", "As WdHeaderFooterIndex", "As WdStoryType", "As VbMsgBoxResult", _
                       "As MsoTriState", "As MsoShapeType", "As MsoTextOrientation", "As WdCollapseDirection")

    Set vbProj = ThisDocument.VBProject
    Set vbComps = vbProj.VBComponents

    For Each comp In vbComps
        Set codeMod = comp.CodeModule
        lineNum = 1
        Do While lineNum < codeMod.CountOfLines
            procName = codeMod.ProcOfLine(lineNum, procType)
            If procName <> "" Then
                For i = lineNum To lineNum + codeMod.ProcCountLines(procName, procType) - 1
                    codeLine = Trim(codeMod.lines(i, 1))
                    If InStr(codeLine, "Dim ") > 0 Or InStr(codeLine, "ReDim ") > 0 Then
                        If ShouldFlag(codeLine, baseTypes, wordTypes, knownEnums, IncludeWordTypes, IncludeEnums) Then
                            Debug.Print FlagLabel(codeLine) & " " & comp.Name & "::" & procName & " | Line " & i & ": " & codeLine
                        End If
                    End If
                Next i
                lineNum = lineNum + codeMod.ProcCountLines(procName, procType)
            Else
                lineNum = lineNum + 1
            End If
        Loop
    Next comp

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FlagEarlyBindingRoutines_LateBound of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

' =========================
' Function:     ShouldFlag
' Purpose:      Determines whether a code line should be flagged for early binding
' Inputs:       codeLine [String] - a trimmed code line
'               baseTypes [Array] - primitive suffixes
'               wordTypes [Array] - suppressible Word object suffixes
'               knownEnums [Array] - suppressible Word/VBA enums
'               IncludeWord [Boolean] - flag Word types
'               IncludeEnums [Boolean] - flag enums
' Returns:      [Boolean] - True to flag, False to suppress
' =========================
Private Function ShouldFlag(codeLine As String, baseTypes As Variant, wordTypes As Variant, knownEnums As Variant, _
                    IncludeWord As Boolean, IncludeEnums As Boolean) As Boolean
    If IsPrimitiveType(codeLine, baseTypes) Then Exit Function
    If IsWordNative(codeLine, wordTypes) And Not IncludeWord Then Exit Function
    If IsEnumType(codeLine, knownEnums) And Not IncludeEnums Then Exit Function
    ShouldFlag = True
End Function

' =========================
' Function:     FlagLabel
' Purpose:      Returns a classification label for early-bound declarations
' Inputs:       codeLine [String] - line to evaluate
' Returns:      [String] - e.g. "[EXTERNAL]", "[CUSTOM]", "[ENUM]", "[WORD]"
' =========================
Private Function FlagLabel(codeLine As String) As String
    Dim lowered As String: lowered = LCase(codeLine)
    If lowered Like "*as excel.*" Or lowered Like "*as filesystem*" Or lowered Like "*as scripting.*" Then
        FlagLabel = "[EXTERNAL]"
    ElseIf lowered Like "*as ae*" Or lowered Like "*as xae*" Then
        FlagLabel = "[CUSTOM]"
    ElseIf lowered Like "*as wd*" Or lowered Like "*as vbmsgbox*" Or lowered Like "*as mso*" Then
        FlagLabel = "[ENUM]"
    ElseIf lowered Like "*As Word.Range*" Or lowered Like "*As Word.Paragraph*" Then
        FlagLabel = "[WORD]"
    Else
        FlagLabel = "[UNCLASSIFIED]"
    End If
End Function

' =========================
' Function:     IsPrimitiveType
' Purpose:      Checks if a declaration is one of the base VBA types
' Inputs:       lineText [String] - Dim line
'               baseTypes [Array] - primitive suffixes
' Returns:      [Boolean]
' =========================
Private Function IsPrimitiveType(lineText As String, baseTypes As Variant) As Boolean
    Dim suffix As Variant
    For Each suffix In baseTypes
        If InStr(lineText, suffix) > 0 Then
            IsPrimitiveType = True
            Exit Function
        End If
    Next suffix
    IsPrimitiveType = False
End Function

' =========================
' Function:     IsWordNative
' Purpose:      Identifies declarations using Word-native object types
' Inputs:       lineText [String], wordTypes [Array]
' Returns:      [Boolean]
' =========================
Private Function IsWordNative(lineText As String, wordTypes As Variant) As Boolean
    Dim suffix As Variant
    For Each suffix In wordTypes
        If InStr(lineText, suffix) > 0 Then
            IsWordNative = True
            Exit Function
        End If
    Next suffix
    IsWordNative = False
End Function

' =========================
' Function:     IsEnumType
' Purpose:      Identifies declarations using Word or VBA enum types
' Inputs:       lineText [String], knownEnums [Array]
' Returns:      [Boolean]
' =========================
Private Function IsEnumType(lineText As String, knownEnums As Variant) As Boolean
    Dim suffix As Variant
    For Each suffix In knownEnums
        If InStr(lineText, suffix) > 0 Then
            IsEnumType = True
            Exit Function
        End If
    Next suffix
    IsEnumType = False
End Function

'====================================================================
' BuildHeadingIndexToCSV
' Scans document for Heading 1 and Heading 2 styles and writes index
' to a CSV-compatible text file with cleaned paragraph text.
' Author: Peter | Date: 20250807
'====================================================================
Public Sub BuildHeadingIndexToCSV()
    On Error GoTo PROC_ERR
    Dim para As Word.Paragraph
    Dim paraIndex As Long
    Dim headingLevel As String
    Dim csvPath As String
    Dim fileNum As Integer

    'csvPath = Environ("USERPROFILE") & "\Desktop\HeadingIndex.csv"
    csvPath = "C:\adaept\aeBibleClass\rpt\HeadingIndex.txt"
    fileNum = FreeFile

    Open csvPath For Output As #fileNum
    Print #fileNum, "Index,Style,Text"

    Dim CleanText As String
    paraIndex = 1
    For Each para In ActiveDocument.Paragraphs

        headingLevel = para.style

        If headingLevel = "Heading 1" Or headingLevel = "Heading 2" Then
            CleanText = Trim(Replace(para.Range.Text, vbCr, ""))
            CleanText = Trim(Replace(CleanText, vbLf, ""))
            CleanText = Replace(CleanText, """", "'") ' Escape quotes for CSV
            Print #fileNum, paraIndex & "," & headingLevel & ",""" & CleanText & """"
        End If

        paraIndex = paraIndex + 1
    Next para

    Close #fileNum
    MsgBox "Heading index written to CSV: " & csvPath

PROC_EXIT:
    Exit Sub
PROC_ERR:
    If fileNum > 0 Then Close #fileNum
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure BuildHeadingIndexToCSV of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

'====================================================================
' LoadHeadingIndexFromCSV
' Loads previously saved heading index into memory buffer (Dictionary)
' for fast lookup and navigation.
' Author: Peter | Date: 20250807
'====================================================================
Public Sub LoadHeadingIndexFromCSV()
    On Error GoTo PROC_ERR
    Dim csvPath As String
    Dim fileNum As Integer
    Dim line As String
    Dim parts() As String

    Set HeadingBuffer = CreateObject("Scripting.Dictionary")
    'csvPath = Environ("USERPROFILE") & "\Desktop\HeadingIndex.csv"
    csvPath = "C:\adaept\aeBibleClass\rpt\HeadingIndex.txt"

    If Dir(csvPath) = "" Then
        MsgBox "CSV file not found: " & csvPath
        GoTo PROC_EXIT
    End If

    fileNum = FreeFile
    Open csvPath For Input As #fileNum
    Line Input #fileNum, line ' Skip header

    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        parts = Split(line, ",")
        If UBound(parts) >= 2 Then
            HeadingBuffer(parts(0)) = parts(2)
        End If
    Loop

    Close #fileNum
    MsgBox "Heading index loaded into memory. " & HeadingBuffer.Count & " entries."

PROC_EXIT:
    Exit Sub
PROC_ERR:
    If fileNum > 0 Then Close #fileNum
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LoadHeadingIndexFromCSV of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

'====================================================================
' GoToHeadingByIndex
' Prompts user for paragraph index and navigates to that location.
' Requires HeadingBuffer to be loaded.
' Author: Peter | Date: 20250807
'====================================================================
Public Sub GoToHeadingByIndex()
    On Error GoTo PROC_ERR
    Dim targetIndex As String
    Dim paraIndex As Long

    If HeadingBuffer Is Nothing Then
        MsgBox "Heading buffer not loaded. Run LoadHeadingIndexFromCSV first."
        GoTo PROC_EXIT
    End If

    targetIndex = InputBox("Enter the paragraph index to jump to:")
    If IsNumeric(targetIndex) Then
        paraIndex = CLng(targetIndex)
        If paraIndex > 0 And paraIndex <= ActiveDocument.Paragraphs.Count Then
            ActiveDocument.Paragraphs(paraIndex).Range.Select
        Else
            MsgBox "Invalid index. Must be between 1 and " & ActiveDocument.Paragraphs.Count
        End If
    Else
        MsgBox "Please enter a numeric value."
    End If

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GoToHeadingByIndex of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

' ==================================================================================================
' Routine:      ShowUnicodeOfSingleCharacterSelection
'
' Purpose:      Identifies the Unicode character currently selected in Word and prints a complete,
'               Immediate-Window-safe diagnostic report. Supports both:
'                   - Single UTF-16 code units (BMP characters)
'                   - Valid surrogate pairs (Unicode > U+FFFF)
'
' Behavior:     - Rejects selections of zero characters.
'               - Rejects selections of more than one logical character.
'               - Accepts:
'                     (1) Exactly one UTF-16 code unit, OR
'                     (2) Exactly two UTF-16 code units forming a valid surrogate pair.
'
' Output:       Prints to the Immediate Window:
'                   - Unicode code point (U+XXXX or U+XXXXX)
'                   - Decimal value
'                   - UTF-16 code units
'                   - Escape sequence (\uXXXX or \UXXXXXXXX)
'                   - Word Special Character description (if applicable)
'
' Notes:        - The Immediate Window cannot display Unicode glyphs, so the routine prints only
'                 descriptive representations.
'               - Calls WordSpecialCharacterName() to map characters to the names used in
'                 Word's Insert > Symbol > Special Characters dialog.
'
' Author:       Peter Ennis
' Last Updated: 20260130
' ==================================================================================================
Public Sub ShowUnicodeOfSingleCharacterSelection()
    On Error GoTo PROC_ERR
    Dim r As Word.Range
    Dim Count As Long
    Dim s As String
    Dim codeUnit1 As Long, codeUnit2 As Long
    Dim scalar As Long
    Dim escapeSeq As String
    Dim desc As String

    Set r = Selection.Range
    Count = r.Characters.Count

    ' Enforce exactly one logical character
    If Count = 0 Then
        Debug.Print "Error: No character selected."
        GoTo PROC_EXIT
    ElseIf Count > 2 Then
        Debug.Print "Error: Selection contains more than one character (" & Count & ")."
        GoTo PROC_EXIT
    End If

    s = r.Text
    codeUnit1 = AscW(Mid$(s, 1, 1))

    ' --- BMP CHARACTER ---
    If Count = 1 Then
        scalar = codeUnit1
        escapeSeq = "\u" & Right$("0000" & Hex$(scalar), 4)

        desc = WordSpecialCharacterName(scalar)

        Debug.Print "BMP character"
        Debug.Print "Unicode code point: U+" & Hex$(scalar)
        Debug.Print "Decimal value: " & scalar
        Debug.Print "UTF-16 unit: " & Hex$(codeUnit1)
        Debug.Print "Escape sequence: " & escapeSeq
        If desc <> "" Then Debug.Print "Description: " & desc
        GoTo PROC_EXIT
    End If

    ' --- POSSIBLE SURROGATE PAIR ---
    codeUnit2 = AscW(Mid$(s, 2, 1))

    If codeUnit1 >= &HD800 And codeUnit1 <= &HDBFF _
       And codeUnit2 >= &HDC00 And codeUnit2 <= &HDFFF Then

        scalar = &H10000 + ((codeUnit1 - &HD800) * &H400) + (codeUnit2 - &HDC00)
        escapeSeq = "\U" & Right$("00000000" & Hex$(scalar), 8)

        Debug.Print "Surrogate pair (valid)"
        Debug.Print "Unicode code point: U+" & Hex$(scalar)
        Debug.Print "Decimal value: " & scalar
        Debug.Print "UTF-16 units: " & Hex$(codeUnit1) & " " & Hex$(codeUnit2)
        Debug.Print "Escape sequence: " & escapeSeq
        Debug.Print "Description: (no Word Special Character name for non-BMP characters)"
    Else
        Debug.Print "Error: Two-character selection is not a valid surrogate pair."
    End If

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ShowUnicodeOfSingleCharacterSelection of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

' ==================================================================================================
' Function:     WordSpecialCharacterName
'
' Purpose:      Returns the descriptive name of a character *exactly as it appears* in the
'               Word Insert > Symbol > Special Characters dialog.
'
' Input:        codePoint - Unicode code point (Long) for the selected character.
'
' Output:       String    - The Word UI name for the character, or "" if the character does not
'                           appear in the Special Characters list.
'
' Behavior:     - Matches only characters that Word exposes in the Special Characters tab.
'               - Order of Select Case blocks matches the order shown in the Word dialog for
'                 easy comparison and maintenance.
'               - Includes both Unicode characters and Word control characters (line break,
'                 column break, page break, etc.) where applicable.
'
' Notes:        - This function does not attempt to name characters outside the Special
'                 Characters list.
'               - Surrogate-pair characters do not appear in the Special Characters dialog.
'
' Author:       Peter Ennis
' Last Updated: 20260130
' ==================================================================================================
Public Function WordSpecialCharacterName(codepoint As Long) As String
    ' Order matches Word: Insert > Symbol > More Symbols > Special Characters

    Select Case codepoint
        ' 1. Em Dash
        Case &H2014
            WordSpecialCharacterName = "Em Dash"
        ' 2. En Dash
        Case &H2013
            WordSpecialCharacterName = "En Dash"
        ' 3. Optional Hyphen (Soft Hyphen)
        Case &HAD
            WordSpecialCharacterName = "Optional Hyphen"
        ' 4. Nonbreaking Hyphen
        Case &H2011
            WordSpecialCharacterName = "Nonbreaking Hyphen"
        ' 5. Nonbreaking Space
        Case &HA0
            WordSpecialCharacterName = "Nonbreaking Space"
        ' 6. Copyright Symbol
        Case &HA9
            WordSpecialCharacterName = "Copyright Symbol"
        ' 7. Registered Trademark Symbol
        Case &HAE
            WordSpecialCharacterName = "Registered Trademark Symbol"
        ' 8. Trademark Symbol
        Case &H2122
            WordSpecialCharacterName = "Trademark Symbol"
        ' 9. Ellipsis
        Case &H2026
            WordSpecialCharacterName = "Ellipsis"
        ' 10. Single Opening Quote
        Case &H2018
            WordSpecialCharacterName = "Single Opening Quote"
        ' 11. Single Closing Quote
        Case &H2019
            WordSpecialCharacterName = "Single Closing Quote"
        ' 12. Double Opening Quote
        Case &H201C
            WordSpecialCharacterName = "Double Opening Quote"
        ' 13. Double Closing Quote
        Case &H201D
            WordSpecialCharacterName = "Double Closing Quote"
        ' 14. Paragraph Mark
        ' (also U+00B6 as a symbol, but Word's special character is the control char)
        Case 13
            WordSpecialCharacterName = "Paragraph Mark"
        ' 15. Section Mark
        Case &HA7
            WordSpecialCharacterName = "Section Mark"
        ' 16. En Space
        Case &H2002
            WordSpecialCharacterName = "En Space"
        ' 17. Em Space
        Case &H2003
            WordSpecialCharacterName = "Em Space"
        ' 18. 1/4 Em Space (Four-per-em space)
        Case &H2005
            WordSpecialCharacterName = "1/4 Em Space"
        ' 19. No-Width Optional Break (zero-width space)
        Case &H200B
            WordSpecialCharacterName = "No-Width Optional Break"
        ' 20. No-Width Non Joiner (ZWNJ)
        Case &H200C
            WordSpecialCharacterName = "No-Width Non Joiner"
        ' 21. No-Width Joiner (ZWJ)
        Case &H200D
            WordSpecialCharacterName = "No-Width Joiner"
        ' 22. Line Break
        Case 11
            WordSpecialCharacterName = "Line Break"
        ' 23. Column Break
        Case 14
            WordSpecialCharacterName = "Column Break"
        ' 24. Page Break
        Case 12
            WordSpecialCharacterName = "Page Break"
        ' (If your build of Word shows any additional Items, we can append them here.)

        Case Else
            WordSpecialCharacterName = ""
    End Select
End Function

Public Sub FindTabsInAllFooters()
    On Error GoTo PROC_ERR
    Dim sec As Word.Section
    Dim hdrFtr As HeaderFooter
    Dim rng As Word.Range
    Dim secStartPage As Long
    Dim secEndPage As Long
    Dim secRange As Word.Range

    For Each sec In ActiveDocument.Sections

        ' Get section page range in document page numbers
        Set secRange = sec.Range

        secRange.Collapse wdCollapseStart
        secStartPage = secRange.Information(wdActiveEndPageNumber)

        secRange.Collapse wdCollapseEnd
        secEndPage = secRange.Information(wdActiveEndPageNumber)

        For Each hdrFtr In sec.Footers
            If hdrFtr.Exists Then
                Set rng = hdrFtr.Range

                With rng.Find
                    .ClearFormatting
                    .Text = "^t"
                    .Forward = True
                    .Wrap = wdFindStop
                    .Execute

                    Do While .found
                        Debug.Print "Section " & sec.index & _
                                    ", Footer type " & hdrFtr.index & _
                                    ", applies to pages " & secStartPage & _
                                    "�" & secEndPage

                        rng.Select
                        Stop

                        rng.Collapse wdCollapseEnd
                        .Execute
                    Loop
                End With
            End If
        Next hdrFtr
    Next sec

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FindTabsInAllFooters of Module basTEST_aeBibleTools"
    Resume PROC_EXIT
End Sub

Public Function CountParagraphMarksWithDarkRedFormatting()
    On Error GoTo PROC_ERR

    Dim p As Word.Paragraph
    Dim Count As Long
    Count = 0
    For Each p In ActiveDocument.Paragraphs
        With p.Range.Characters.Last.Font
            If .color = wdColorDarkRed Then
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
    Resume PROC_EXIT
End Function

Sub CleanPollutedParagraphMarks()

    Dim p As Word.Paragraph
    Dim pm As Word.Range

    For Each p In ActiveDocument.Paragraphs

        ' The paragraph mark is always the last character in the paragraph range
        Set pm = p.Range.Characters.Last

        ' If the paragraph mark has ANY direct formatting, clear it
        If pm.Font.color <> wdColorAutomatic _
        Or pm.Font.Bold <> False _
        Or pm.Font.Italic <> False _
        Or pm.Font.Underline <> wdUnderlineNone _
        Or pm.Font.Name <> "" _
        Or pm.Font.Size <> 0 Then

            ' Reset the paragraph mark to paragraph style defaults
            pm.Font.Reset
        End If

    Next p

    MsgBox "Paragraph mark cleanup complete."

End Sub

Sub CleanDarkRedParagraphMarks()
    Dim p As Word.Paragraph
    Dim pm As Word.Range

    For Each p In ActiveDocument.Paragraphs
        ' Paragraph mark = last character of the paragraph range
        Set pm = p.Range.Characters.Last

        ' Check ONLY for Dark Red formatting
        If pm.Font.color = wdColorDarkRed Then
            pm.Font.Reset   ' Remove character formatting pollution
        End If
    Next p

    MsgBox "DarkRed paragraph mark cleanup complete."
End Sub

