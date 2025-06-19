Attribute VB_Name = "basTESTaeBibleTools"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

' Use Windows API to change cursor
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" ( _
    ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

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

Sub GoToVerseSBL()
    On Error GoTo ErrHandler
    Application.StatusBar = "Searching for verse..."
    
    Dim userInput As String
    userInput = InputBox("Enter verse (e.g. 1 Sam 1:1):", "Go to Verse (SBL Format)")
    userInput = UCase(userInput)
    If Trim(userInput) = "" Then Exit Sub
    
    Dim bookAbbr As String, chapNum As String, verseNum As String
    Dim parts() As String, subParts() As String
    
    Dim hWaitCursor As Long
    ' Set spinning cursor manually
    hWaitCursor = LoadCursor(0, 32514) ' 32514 = Busy (Hourglass)
    SetCursor hWaitCursor
    Application.ScreenUpdating = False  ' Prevent flickering

    ' Parse the input
    parts = Split(userInput, ":")
    'Debug.Print "UBound(parts) = " & UBound(parts)
    If UBound(parts) = 0 Then   ' Only the ~chapter~ number was provided
        verseNum = 1
        GoTo Chapter
    ElseIf UBound(parts) <> 1 Then
        Application.ScreenUpdating = True   ' Restore normal UI
        SetCursor LoadCursor(0, 32512)      ' Restore default arrow cursor
        MsgBox "Invalid format. Use format like '1 Sam 1:1'", vbExclamation
        Exit Sub
    End If
    verseNum = Trim(parts(1))
Chapter:
    subParts = Split(Trim(parts(0)))
    If UBound(subParts) = 0 Then
        bookAbbr = Trim(parts(0))
        chapNum = "1"
    ElseIf (subParts(0) = "1" Or subParts(0) = "2") And UBound(subParts) = 1 Then
        bookAbbr = subParts(0) & " " & subParts(1)
        'Debug.Print "a:", bookAbbr
        chapNum = 1
    Else
        Dim i As Long
        bookAbbr = ""
        For i = 0 To UBound(subParts) - 1
            bookAbbr = bookAbbr & subParts(i) & " "
            'Debug.Print "b:", bookAbbr
        Next i
        bookAbbr = Trim(bookAbbr)
        'Debug.Print ">", bookAbbr
        chapNum = Trim(subParts(UBound(subParts)))
        'Debug.Print ">>", chapNum
    End If
    
    Dim fullBookName As String
    fullBookName = GetFullBookName(bookAbbr)
    'Debug.Print ">>>", fullBookName
    If fullBookName = "" Then
        Application.ScreenUpdating = True   ' Restore normal UI
        SetCursor LoadCursor(0, 32512)      ' Restore default arrow cursor
        MsgBox "Book not found: " & bookAbbr, vbExclamation
        Exit Sub
    End If

    ' Find the Heading 1 for the book
    Dim theBook As String
    Dim para As paragraph, foundBook As Boolean
    For Each para In ActiveDocument.paragraphs
        If para.style = "Heading 1" Then
            theBook = Trim(para.range.text)
            theBook = UCase(Replace(para.range.text, vbCr, ""))
            'Debug.Print theBook
            If theBook Like "*" & fullBookName & "*" Then
                para.range.Select
                foundBook = True
                'Debug.Print bookAbbr, theBook, fullBookName
                'MsgBox "Book found. Searching for chapter or verse " & chapNum, vbInformation
                Exit For
            End If
        End If
    Next para
    If Not foundBook Then
        Application.ScreenUpdating = True   ' Restore normal UI
        SetCursor LoadCursor(0, 32512)      ' Restore default arrow cursor
        MsgBox "Book heading not found: " & fullBookName, vbExclamation
        Exit Sub
    End If
    
    ' Find the Heading 2 for the chapter or psalm
    Dim theChapter As String
    Dim chapFound As Boolean
    For Each para In ActiveDocument.paragraphs
        'Debug.Print para.range.Start, Selection.range.Start
        If para.range.Start < Selection.range.Start Then GoTo SkipChapter
        If para.style = "Heading 2" Then
            Select Case theBook     ' Books of only one chapter
                Case "OBADIAH", "PHILEMON", "2 JOHN", "3 JOHN", "JUDE"
                    verseNum = chapNum
                    chapNum = 1
                    chapFound = True
                    'MsgBox "The name " & theBook & " is in the list!", vbInformation
                Case Else
                    'MsgBox "The name is NOT in the list!", vbExclamation
            End Select
            If Trim(para.range.text) Like "*Chapter " & chapNum & "*" _
                    Or Trim(para.range.text) Like "*Psalm " & chapNum & "*" Then
                para.range.Select
                chapFound = True
                Exit For
            End If
        End If
SkipChapter:
    Next para
    If Not chapFound Then
        Application.ScreenUpdating = True   ' Restore normal UI
        SetCursor LoadCursor(0, 32512)      ' Restore default arrow cursor
        MsgBox "Chapter not found: " & chapNum, vbExclamation
        Exit Sub
    End If

    ' Limit search range to current chapter
    Dim chapStart As Long, chapEnd As Long
    chapStart = Selection.range.Start
    chapEnd = ActiveDocument.content.End
    For Each para In ActiveDocument.paragraphs
        If para.range.Start > chapStart And para.style = "Heading 2" Then
            chapEnd = para.range.Start
            Exit For
        End If
    Next para

    ' Search for verse number within "Verse marker" style
    Dim r As range
    Dim charCount As Long, found As Boolean
    charCount = chapStart
    Do While charCount < chapEnd
        Set r = ActiveDocument.range(charCount, charCount + 1)
        If r.Characters(1).style = "Verse marker" Then
            Dim verseStr As String, j As Long
            verseStr = ""
            For j = 0 To 2 ' Check up to 3 digits
                If charCount + j >= chapEnd Then Exit For
                If IsNumeric(ActiveDocument.range(charCount + j, charCount + j + 1).text) Then
                    verseStr = verseStr & ActiveDocument.range(charCount + j, charCount + j + 1).text
                Else
                    Exit For
                End If
            Next j
            If verseStr = verseNum Then
                ActiveDocument.range(charCount, charCount + Len(verseStr)).Select
                found = True
                Exit Do
            End If
        End If
        charCount = charCount + 1
        If charCount Mod 1000 = 0 Then DoEvents
    Loop
    If Not found Then
        MsgBox "Verse not found: " & verseNum, vbExclamation
    End If

Cleanup:
    Application.ScreenUpdating = True   ' Restore normal UI
    SetCursor LoadCursor(0, 32512)      ' Restore default arrow cursor
    Application.StatusBar = False
    Exit Sub

ErrHandler:
    MsgBox "Error during verse search: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Function GetFullBookName(abbr As String) As String
    Dim bookMap As Object
    Set bookMap = CreateObject("Scripting.Dictionary")
    
    bookMap.Add UCase("Gen"), "Genesis"
    bookMap.Add UCase("Ge"), "Genesis"
    bookMap.Add UCase("Exod"), "Exodus"
    bookMap.Add UCase("Ex"), "Exodus"
    bookMap.Add UCase("Lev"), "Leviticus"
    bookMap.Add UCase("Le"), "Leviticus"
    bookMap.Add UCase("Num"), "Numbers"
    bookMap.Add UCase("Nu"), "Numbers"
    bookMap.Add UCase("Deut"), "Deuteronomy"
    bookMap.Add UCase("De"), "Deuteronomy"
    bookMap.Add UCase("Josh"), "Joshua"
    bookMap.Add UCase("Jos"), "Joshua"
    bookMap.Add UCase("Judg"), "Judges"
    bookMap.Add UCase("Ruth"), "Ruth"
    bookMap.Add UCase("Ru"), "Ruth"
    bookMap.Add UCase("1 Sam"), "1 Samuel"
    bookMap.Add UCase("1 S"), "1 Samuel"
    bookMap.Add UCase("2 Sam"), "2 Samuel"
    bookMap.Add UCase("2 S"), "2 Samuel"
    bookMap.Add UCase("1 Kgs"), "1 Kings"
    bookMap.Add UCase("1 K"), "1 Kings"
    bookMap.Add UCase("2 Kgs"), "2 Kings"
    bookMap.Add UCase("2 K"), "2 Kings"
    bookMap.Add UCase("1 Chr"), "1 Chronicles"
    bookMap.Add UCase("1 Ch"), "1 Chronicles"
    bookMap.Add UCase("2 Chr"), "2 Chronicles"
    bookMap.Add UCase("2 Ch"), "2 Chronicles"
    bookMap.Add UCase("Ezra"), "Ezra"
    bookMap.Add UCase("Ezr"), "Ezra"
    bookMap.Add UCase("Neh"), "Nehemiah"
    bookMap.Add UCase("Ne"), "Nehemiah"
    bookMap.Add UCase("Esth"), "Esther"
    bookMap.Add UCase("Es"), "Esther"
    bookMap.Add UCase("Job"), "Job"
    bookMap.Add UCase("Ps"), "Psalms"
    bookMap.Add UCase("Prov"), "Proverbs"
    bookMap.Add UCase("Pr"), "Proverbs"
    bookMap.Add UCase("Eccl"), "Ecclesiastes"
    bookMap.Add UCase("Ec"), "Ecclesiastes"
    bookMap.Add UCase("Ecc"), "Ecclesiastes"
    bookMap.Add UCase("Song"), "Solomon"
    bookMap.Add UCase("S"), "Solomon"
    bookMap.Add UCase("Isa"), "Isaiah"
    bookMap.Add UCase("Is"), "Isaiah"
    bookMap.Add UCase("I"), "Isaiah"
    bookMap.Add UCase("Jer"), "Jeremiah"
    bookMap.Add UCase("Je"), "Jeremiah"
    bookMap.Add UCase("Lam"), "Lamentations"
    bookMap.Add UCase("La"), "Lamentations"
    bookMap.Add UCase("Ezek"), "Ezekiel"
    bookMap.Add UCase("Eze"), "Ezekiel"
    bookMap.Add UCase("Dan"), "Daniel"
    bookMap.Add UCase("Da"), "Daniel"
    bookMap.Add UCase("Hos"), "Hosea"
    bookMap.Add UCase("Ho"), "Hosea"
    bookMap.Add UCase("Joel"), "Joel"
    bookMap.Add UCase("Joe"), "Joel"
    bookMap.Add UCase("Amos"), "Amos"
    bookMap.Add UCase("Am"), "Amos"
    bookMap.Add UCase("Obad"), "Obadiah"
    bookMap.Add UCase("O"), "Obadiah"
    bookMap.Add UCase("Jonah"), "Jonah"
    bookMap.Add UCase("Jon"), "Jonah"
    bookMap.Add UCase("Mic"), "Micah"
    bookMap.Add UCase("Mi"), "Micah"
    bookMap.Add UCase("Nah"), "Nahum"
    bookMap.Add UCase("Na"), "Nahum"
    bookMap.Add UCase("Hab"), "Habakkuk"
    bookMap.Add UCase("Zeph"), "Zephaniah"
    bookMap.Add UCase("Zep"), "Zephaniah"
    bookMap.Add UCase("Hag"), "Haggai"
    bookMap.Add UCase("Zech"), "Zechariah"
    bookMap.Add UCase("Zec"), "Zechariah"
    bookMap.Add UCase("Mal"), "Malachi"
    bookMap.Add UCase("Matt"), "Matthew"
    bookMap.Add UCase("Mat"), "Matthew"
    bookMap.Add UCase("Mark"), "Mark"
    bookMap.Add UCase("Mar"), "Mark"
    bookMap.Add UCase("Luke"), "Luke"
    bookMap.Add UCase("Lu"), "Luke"
    bookMap.Add UCase("John"), "John"
    bookMap.Add UCase("Joh"), "John"
    bookMap.Add UCase("Acts"), "Acts"
    bookMap.Add UCase("Ac"), "Acts"
    bookMap.Add UCase("Rom"), "Romans"
    bookMap.Add UCase("Ro"), "Romans"
    bookMap.Add UCase("1 Cor"), "1 Corinthians"
    bookMap.Add UCase("1 Co"), "1 Corinthians"
    bookMap.Add UCase("2 Cor"), "2 Corinthians"
    bookMap.Add UCase("2 Co"), "2 Corinthians"
    bookMap.Add UCase("Gal"), "Galatians"
    bookMap.Add UCase("Ga"), "Galatians"
    bookMap.Add UCase("Eph"), "Ephesians"
    bookMap.Add UCase("Ep"), "Ephesians"
    bookMap.Add UCase("Phil"), "Philippians"
    bookMap.Add UCase("Phili"), "Philippians"
    bookMap.Add UCase("Col"), "Colossians"
    bookMap.Add UCase("C"), "Colossians"
    bookMap.Add UCase("1 Thess"), "1 Thessalonians"
    bookMap.Add UCase("1 Th"), "1 Thessalonians"
    bookMap.Add UCase("2 Thess"), "2 Thessalonians"
    bookMap.Add UCase("2 Th"), "2 Thessalonians"
    bookMap.Add UCase("1 Tim"), "1 Timothy"
    bookMap.Add UCase("1 Ti"), "1 Timothy"
    bookMap.Add UCase("2 Tim"), "2 Timothy"
    bookMap.Add UCase("2 Ti"), "2 Timothy"
    bookMap.Add UCase("Titus"), "Titus"
    bookMap.Add UCase("T"), "Titus"
    bookMap.Add UCase("Phlm"), "Philemon"
    bookMap.Add UCase("Phile"), "Philemon"
    bookMap.Add UCase("Heb"), "Hebrews"
    bookMap.Add UCase("He"), "Hebrews"
    bookMap.Add UCase("Jas"), "James"
    bookMap.Add UCase("Ja"), "James"
    bookMap.Add UCase("1 Pet"), "1 Peter"
    bookMap.Add UCase("1 P"), "1 Peter"
    bookMap.Add UCase("2 Pet"), "2 Peter"
    bookMap.Add UCase("2 P"), "2 Peter"
    bookMap.Add UCase("1 John"), "1 John"
    bookMap.Add UCase("1 J"), "1 John"
    bookMap.Add UCase("2 John"), "2 John"
    bookMap.Add UCase("2 J"), "2 John"
    bookMap.Add UCase("3 John"), "3 John"
    bookMap.Add UCase("3 J"), "3 John"
    bookMap.Add UCase("Jude"), "Jude"
    bookMap.Add UCase("Rev"), "Revelation"
    bookMap.Add UCase("Re"), "Revelation"
    
    abbr = UCase(Trim(abbr))
    If bookMap.Exists(abbr) Then
        GetFullBookName = bookMap(abbr)
    Else
        GetFullBookName = ""
    End If
End Function

