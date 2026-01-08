Attribute VB_Name = "basUSFMExport"
' Prompt:
' The current code for Word Bible audit is here: https://github.com/adaept/aeBibleClass/blob/main/src/basWordRepairRunner.bas#L31
' Use the code as the base for style information etc. as needed to create an initial USFM export routine for VBA.
' Identify areas that need expansion and more detail as comments in the code. Use the same function calls that allow
' a page or range of pages for exporting. Provide a detailed comment header and sample test harness for validation
' purposes with an audit log file.

' ============================================================================================
'  MODULE: basUSFMExport
'  AUTHOR: Peter - reproducible, audit-traceable architecture
'  PURPOSE:
'       Initial USFM export routine built on the same structural patterns as
'       basWordRepairRunner. This module:
'           - Exports a page or page-range from Word to USFM text
'           - Uses style-aware mapping hooks (expandable)
'           - Logs all actions to an audit file
'           - Benchmarks execution time
'           - Provides deterministic, reproducible output
'
'  DESIGN PRINCIPLES:
'       - No UI-bound operations inside loops
'       - Paragraph scanning optimized for speed
'       - Style mapping isolated for maintainability
'       - All transformations logged
'       - Output is pure text (UTF-8 recommended)
'
'  REQUIRED:
'       - basWordRepairRunner (for page-range scanning patterns)
'       - A consistent style schema in the Word document
'
'  EXPANSION NEEDED:
'       - Full style -> USFM mapping table
'       - Footnote and cross-reference extraction
'       - Verse-number detection logic
'       - Poetry, lists, introductions, study notes
'       - Multi-column or side-bar content handling
'
' ============================================================================================

Option Explicit

' -----------------------------
' CONFIGURATION
' -----------------------------
Private Const LOG_FILE As String = "C:\adaept\aeBibleClass\rpt\USFM_Export_Log.txt"
Private Const OUTPUT_FILE As String = "C:\adaept\aeBibleClass\rpt\ExportedBible.usfm"
Private Const VALIDATOR_LOG As String = "C:\adaept\aeBibleClass\rpt\USFM_Validator_Log.txt"
Private currentChapter As Long
Private bookTitleLevel As Long   ' 0 = none, 1 = next is mt2, 2 = next is mt3

' ============================================================================================
' PUBLIC ENTRY POINT
' ============================================================================================
Public Sub ExportUSFM_PageRange(ByVal startPage As Long, ByVal endPage As Long)
    Dim t0 As Double: t0 = Timer
    currentChapter = 0
    bookTitleLevel = 0
    
    LogEvent "=== USFM EXPORT START ==="
    LogEvent "Page range: " & startPage & " to " & endPage

    Dim rng As range
    Set rng = GetRangeForPages(startPage, endPage)

    If rng Is Nothing Then
        LogEvent "ERROR: Page range returned no content."
        Exit Sub
    End If

    Dim usfm As String
    usfm = ConvertRangeToUSFM(rng)

    WriteTextFile OUTPUT_FILE, usfm
    ValidateUSFMFile OUTPUT_FILE
    
    LogEvent "USFM written to: " & OUTPUT_FILE

    LogEvent "Execution time: " & Format(Timer - t0, "0.00") & " seconds"
    LogEvent "=== USFM EXPORT END ==="
End Sub

' ============================================================================================
' CORE CONVERSION
' ============================================================================================
Private Function ConvertRangeToUSFM(ByVal rng As range) As String
    Dim p As paragraph
    Dim sb As String
    Dim line As String

    LogEvent "Beginning paragraph scan..."

    For Each p In rng.paragraphs
        line = ConvertParagraphToUSFM(p)
        If Len(line) > 0 Then
            sb = sb & line & vbCrLf
        End If
    Next p

    ConvertRangeToUSFM = sb
End Function

' ============================================================================================
' PARAGRAPH -> USFM
' ============================================================================================
Private Function ConvertParagraphToUSFM(ByVal p As paragraph) As String
    Dim styleName As String
    Dim txt As String
    Dim chapNum As Long
    Dim verseNum As Long
    Dim verseText As String

    styleName = Trim$(p.style.NameLocal)
    txt = CleanTextForUTF8(Trim$(p.range.text))

    'LogEvent "STYLE=[" & styleName & "] RAW=[" & txt & "]"
    'LogEvent "CHARSTYLE=[" & p.range.Characters(1).style & "]"

    '===========================================================
    ' 0. FORM FEED / WHITESPACE HANDLING
    '===========================================================
    If txt = Chr(12) Then
        ConvertParagraphToUSFM = "\pb"
        GoTo LogAndExit
    End If

    ' --- EFFECTIVE EMPTY CHECK ---
    If IsEffectivelyEmpty(txt) Then
        ConvertParagraphToUSFM = ""
        LogEvent "Ignored effectively-empty paragraph"
        GoTo LogAndExit
    End If

    '===========================================================
    ' 1. CHARACTER-STYLE SEMANTICS (highest priority)
    '===========================================================

    ' --- Book Title (character style) ---
    If ParagraphHasCharStyle(p, "Book Title") Then
        ConvertParagraphToUSFM = "\mt1 " & txt
        bookTitleLevel = 1
        GoTo LogAndExit
    End If

    ' --- Chapter marker (character style) ---
    If ParagraphHasCharStyle(p, "Chapter Verse marker") Then
        Dim chapTxt As String
        chapTxt = ExtractCharStyleText(p, "Chapter Verse marker")

        If IsNumeric(chapTxt) Then
            currentChapter = CLng(chapTxt)
            ConvertParagraphToUSFM = "\c " & currentChapter
        Else
            ConvertParagraphToUSFM = "\rem INVALID CHAPTER MARKER: " & chapTxt
        End If

        GoTo LogAndExit
    End If

    ' --- Verse marker (character style) ---
    If ParagraphHasCharStyle(p, "Verse marker") Then
        Dim vTxt As String
        vTxt = ExtractCharStyleText(p, "Verse marker")

        If IsNumeric(vTxt) Then
            verseNum = CLng(vTxt)
            verseText = Trim$(Replace(txt, vTxt, ""))
            ConvertParagraphToUSFM = "\v " & verseNum & " " & verseText
        Else
            ConvertParagraphToUSFM = "\rem INVALID VERSE MARKER: " & vTxt
        End If

        GoTo LogAndExit
    End If

    '===========================================================
    ' 2. PARAGRAPH-STYLE SEMANTICS (your existing logic)
    '===========================================================
    Select Case styleName

        Case "Book Title"
            ConvertParagraphToUSFM = "\mt1 " & txt
            bookTitleLevel = 1

        Case "Heading 1"
            ConvertParagraphToUSFM = "\mt1 " & txt
            bookTitleLevel = 1

        Case "CustomParaAfterH1"
            ConvertParagraphToUSFM = "\mt2 " & txt
            bookTitleLevel = 0   ' explicitly end any title sequence

        Case "Heading 2"
            chapNum = ExtractTrailingNumber(txt)
            If chapNum > 0 Then
                currentChapter = chapNum
                ConvertParagraphToUSFM = "\cl " & txt & vbCrLf & "\c " & chapNum
            Else
                ConvertParagraphToUSFM = "\cl " & txt
            End If

        Case "DatAuthRef"
            If Right$(txt, 1) = ":" Then
                ConvertParagraphToUSFM = "\is2 " & Left$(txt, Len(txt) - 1)
            Else
                ConvertParagraphToUSFM = "\ip " & txt
            End If

        Case "Plain Text", "Normal"
            If bookTitleLevel = 1 Then
                ConvertParagraphToUSFM = "\mt2 " & txt
                bookTitleLevel = 2

            ElseIf bookTitleLevel = 2 Then
                ConvertParagraphToUSFM = "\mt3 " & txt
                bookTitleLevel = 0

            ElseIf TryParseChapterVerseFromStyles(p, chapNum, verseNum, verseText) Then
                If chapNum > 0 Then currentChapter = chapNum
                ConvertParagraphToUSFM = "\v " & verseNum & " " & verseText

            Else
                ConvertParagraphToUSFM = "\p " & txt
            End If

        Case Else
            If TryParseChapterVerseFromStyles(p, chapNum, verseNum, verseText) Then
                If chapNum > 0 Then currentChapter = chapNum
                ConvertParagraphToUSFM = "\v " & verseNum & " " & verseText
            Else
                ConvertParagraphToUSFM = "\p " & txt
            End If

    End Select

LogAndExit:
    LogEvent "Converted (" & styleName & "): " & Left$(ConvertParagraphToUSFM, 80)
End Function

Function IsEffectivelyEmpty(txt As String) As Boolean
    Dim t As String
    t = Trim$(txt)

    ' Remove tabs, NBSP, CR, LF, Unicode spaces
    t = Replace(t, vbTab, "")
    t = Replace(t, ChrW(160), "")   ' NBSP
    t = Replace(t, ChrW(8203), "")  ' zero-width space
    t = Replace(t, ChrW(65279), "") ' BOM
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")

    IsEffectivelyEmpty = (Len(t) = 0)
End Function

Private Function ParagraphHasCharStyle(p As paragraph, styleName As String) As Boolean
    Dim r As range
    For Each r In p.range.words
        If r.style = styleName Then
            ParagraphHasCharStyle = True
            Exit Function
        End If
    Next
End Function

Private Function ExtractCharStyleText(p As paragraph, styleName As String) As String
    Dim r As range
    Dim buf As String
    For Each r In p.range.words
        If r.style = styleName Then
            buf = buf & r.text
        End If
    Next
    ExtractCharStyleText = Trim$(buf)
End Function

Private Function TryParseChapterVerseFromStyles( _
    ByVal p As paragraph, _
    ByRef chapNum As Long, _
    ByRef verseNum As Long, _
    ByRef verseText As String) As Boolean

    Dim rChap As range
    Dim rVerse As range
    Dim rText As range

    chapNum = 0
    verseNum = 0
    verseText = ""

    '------------------------------------------------------------
    ' 1. Chapter number run (character style: "Chapter Verse marker")
    '------------------------------------------------------------
    Set rChap = p.range.Duplicate
    rChap.Collapse wdCollapseStart
    rChap.MoveEnd wdCharacter, 1

    If rChap.style <> "Chapter Verse marker" Then
        ' If the paragraph doesn't start with the chapter marker style,
        ' we can't parse it as a verse line.
        TryParseChapterVerseFromStyles = False
        Exit Function
    End If

    ' Extend rChap to include all contiguous chars with that style
    Do While rChap.End < p.range.End And rChap.style = "Chapter Verse marker"
        rChap.MoveEnd wdCharacter, 1
    Loop
    rChap.MoveEnd wdCharacter, -1 ' step back one char after overshoot

    chapNum = CLng(Trim$(CleanTextForUTF8(rChap.text)))

    '------------------------------------------------------------
    ' 2. Verse number run (character style: "Verse marker")
    '------------------------------------------------------------
    Set rVerse = p.range.Duplicate
    rVerse.Start = rChap.End
    rVerse.Collapse wdCollapseStart
    rVerse.MoveEnd wdCharacter, 1

    If rVerse.style <> "Verse marker" Then
        TryParseChapterVerseFromStyles = False
        Exit Function
    End If

    ' Extend rVerse to include all contiguous chars with that style
    Do While rVerse.End < p.range.End And rVerse.style = "Verse marker"
        rVerse.MoveEnd wdCharacter, 1
    Loop
    rVerse.MoveEnd wdCharacter, -1

    verseNum = CLng(Trim$(CleanTextForUTF8(rVerse.text)))

    '------------------------------------------------------------
    ' 3. Remaining text = verse content
    '------------------------------------------------------------
    Set rText = p.range.Duplicate
    rText.Start = rVerse.End
    verseText = CleanTextForUTF8(Trim$(rText.text))

    If verseText = "" Then
        TryParseChapterVerseFromStyles = False
    Else
        TryParseChapterVerseFromStyles = True
    End If
End Function

Private Function ExtractTrailingNumber(ByVal s As String) As Long
    Dim i As Long
    Dim digits As String
    s = Trim$(s)

    For i = Len(s) To 1 Step -1
        If mid$(s, i, 1) >= "0" And mid$(s, i, 1) <= "9" Then
            digits = mid$(s, i, 1) & digits
        ElseIf digits <> "" Then
            Exit For
        End If
    Next i

    If digits <> "" Then
        ExtractTrailingNumber = CLng(digits)
    Else
        ExtractTrailingNumber = 0
    End If
End Function

Private Function CleanTextForUTF8(ByVal s As String) As String
    ' Remove soft hyphens and other invisible Unicode artifacts
    s = Replace(s, ChrW(&HAD), "")      ' Soft hyphen
    s = Replace(s, ChrW(&H2011), "-")   ' Non-breaking hyphen ? normal hyphen
    s = Replace(s, ChrW(&H200B), "")    ' Zero-width space
    s = Replace(s, ChrW(&H200C), "")    ' Zero-width non-joiner
    s = Replace(s, ChrW(&H200D), "")    ' Zero-width joiner

    ' Remove any leftover control characters except CR/LF/TAB
    Dim i As Long, out As String, ch As String
    For i = 1 To Len(s)
        ch = mid$(s, i, 1)
        If AscW(ch) >= 32 Or AscW(ch) = 9 Or AscW(ch) = 10 Or AscW(ch) = 13 Then
            out = out & ch
        End If
    Next i

    CleanTextForUTF8 = out
End Function

Private Sub LogValidator(ByVal msg As String)
    Dim stm As Object
    Dim existing As String
    Dim logLine As String
    Dim logPath As String

    logPath = VALIDATOR_LOG
    logLine = CleanTextForUTF8(Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & msg & vbCrLf)

    Set stm = CreateObject("ADODB.Stream")

    On Error GoTo ErrHandler

    ' Read existing UTF-8 content if present
    If Dir(logPath) <> "" Then
        stm.Type = 2
        stm.Charset = "UTF-8"
        stm.Open
        stm.LoadFromFile logPath
        existing = stm.ReadText
        stm.Close
    Else
        existing = ""
    End If

    ' Write combined content back as UTF-8
    stm.Open
    stm.Type = 2
    stm.Charset = "UTF-8"
    stm.WriteText existing & logLine
    stm.Position = 0
    stm.saveToFile logPath, 2
    stm.Close

    Set stm = Nothing
    Exit Sub

ErrHandler:
    Debug.Print "Validator UTF-8 ERROR: "; Err.Number; Err.Description
End Sub

Public Sub ValidateUSFMFile(ByVal filePath As String)
    Dim stm As Object
    Dim content As String
    Dim lines() As String
    Dim i As Long
    Dim line As String
    Dim marker As String

    'LogValidator "=== USFM VALIDATION START ==="
    LogValidator "Validating file: " & filePath

    On Error GoTo ErrHandler

    ' Load UTF-8 file
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "UTF-8"
    stm.Open
    stm.LoadFromFile filePath
    content = stm.ReadText
    stm.Close

    lines = Split(content, vbCrLf)

    For i = 0 To UBound(lines)
        line = CleanTextForUTF8(lines(i))

        ' Skip blank lines
        If Trim(line) = "" Then
            GoTo NextLine
        End If

        ' 1. Must start with "\" unless it's a continuation line
        If Left$(Trim(line), 1) <> "\" Then
            LogValidator "Line " & (i + 1) & ": Missing USFM marker ? " & line
            GoTo NextLine
        End If

        ' Extract marker
        marker = ExtractUSFMMarker(line)

        ' 2. Validate marker name
        If Not IsKnownUSFMMarker(marker) Then
            LogValidator "Line " & (i + 1) & ": Unknown marker '" & marker & "' ? " & line
        End If

        ' 3. Marker must be followed by space unless it's a standalone marker
        If Not MarkerAllowsNoSpace(marker) Then
            If Len(line) > Len(marker) + 1 Then
                If mid$(line, Len(marker) + 1, 1) <> " " Then
                    LogValidator "Line " & (i + 1) & ": Missing space after marker '" & marker & "' ? " & line
                End If
            End If
        End If

        ' 4. Check for empty content after markers that require content
        If MarkerRequiresContent(marker) Then
            If Len(Trim(mid$(line, Len(marker) + 2))) = 0 Then
                LogValidator "Line " & (i + 1) & ": Marker '" & marker & "' missing content ? " & line
            End If
        End If

NextLine:
    Next i
    
    LogValidator "=== USFM VALIDATION END ==="
    Exit Sub

ErrHandler:
    LogValidator "ERROR validating USFM: " & Err.Number & " - " & Err.Description
End Sub

Private Function ExtractUSFMMarker(ByVal line As String) As String
    Dim i As Long
    If Left$(line, 1) <> "\" Then
        ExtractUSFMMarker = ""
        Exit Function
    End If

    For i = 2 To Len(line)
        If mid$(line, i, 1) = " " Or mid$(line, i, 1) = "*" Then
            ExtractUSFMMarker = Left$(line, i - 1)
            Exit Function
        End If
    Next i

    ExtractUSFMMarker = line
End Function

Private Function IsKnownUSFMMarker(ByVal m As String) As Boolean
    Select Case m
        Case "\p", "\m", "\q", "\q1", "\q2", "\q3", _
             "\s1", "\s2", "\s3", _
             "\ip", "\ipi", "\im", _
             "\is1", "\is2", "\is3", _
             "\mt1", "\mt2", "\mt3", _
             "\d", _
             "\pb", _
             "\c", "\cl", "\v", _
             "\r"
            IsKnownUSFMMarker = True
        Case Else
            IsKnownUSFMMarker = False
    End Select
End Function

Private Function MarkerAllowsNoSpace(ByVal m As String) As Boolean
    Select Case m
        Case "\pb", "\c"
            MarkerAllowsNoSpace = True
        Case Else
            MarkerAllowsNoSpace = False
    End Select
End Function

Private Function MarkerRequiresContent(ByVal m As String) As Boolean
    Select Case m
        Case "\p", "\m", "\q", "\q1", "\q2", "\q3", _
             "\s1", "\s2", "\s3", _
             "\ip", "\ipi", "\im", "\is1", "\is2", "\is3", _
             "\mt1", "\mt2", "\mt3", "\d", _
             "\r"
            MarkerRequiresContent = True
        Case Else
            MarkerRequiresContent = False
    End Select
End Function

Private Function OLD_ConvertParagraphToUSFM(ByVal p As paragraph) As String
    Dim styleName As String
    styleName = p.style

    ' ---------------------------------------------------------
    ' EXPAND HERE:
    '   Map your Word styles to USFM markers.
    '   This is the core of the exporter.
    ' ---------------------------------------------------------
    Select Case styleName

        Case "Verse"
            ' EXPAND: Add verse-number extraction
            OLD_ConvertParagraphToUSFM = "\v " & Trim(p.range.text)

        Case "Paragraph"
            OLD_ConvertParagraphToUSFM = "\p " & Trim(p.range.text)

        Case "Heading 1"
            OLD_ConvertParagraphToUSFM = "\s1 " & Trim(p.range.text)

        Case "Heading 2"
            OLD_ConvertParagraphToUSFM = "\s2 " & Trim(p.range.text)

        Case "Poetry 1"
            OLD_ConvertParagraphToUSFM = "\q1 " & Trim(p.range.text)

        Case "Poetry 2"
            OLD_ConvertParagraphToUSFM = "\q2 " & Trim(p.range.text)

        Case Else
            ' EXPAND: Add more mappings
            OLD_ConvertParagraphToUSFM = Trim(p.range.text)

    End Select

    LogEvent "Converted paragraph (" & styleName & "): " & Left(OLD_ConvertParagraphToUSFM, 80)
End Function

' ============================================================================================
' PAGE RANGE EXTRACTION (same pattern as basWordRepairRunner)
' ============================================================================================
Private Function GetRangeForPages(ByVal startPage As Long, ByVal endPage As Long) As range
    Dim doc As Document
    Dim rStartPage As range
    Dim rEndPage As range
    Dim fullRange As range

    On Error GoTo ErrHandler

    Set doc = ActiveDocument

    ' -------------------------------------------------------------
    ' Go to the START PAGE using the printed page number
    ' (same pattern as RunRepairWrappedVerseMarkers_Across_Pages_From)
    ' -------------------------------------------------------------
    Application.Selection.GoTo What:=wdGoToPage, name:=CStr(startPage)
    Set rStartPage = Application.Selection.Bookmarks("\Page").range

    ' -------------------------------------------------------------
    ' Go to the END PAGE using the printed page number
    ' -------------------------------------------------------------
    Application.Selection.GoTo What:=wdGoToPage, name:=CStr(endPage)
    Set rEndPage = Application.Selection.Bookmarks("\Page").range

    ' -------------------------------------------------------------
    ' Build a range from the start of startPage to the end of endPage
    ' -------------------------------------------------------------
    Set fullRange = doc.range(Start:=rStartPage.Start, End:=rEndPage.End)

    Set GetRangeForPages = fullRange
    Exit Function

ErrHandler:
    LogEvent "ERROR in GetRangeForPages: " & Err.Number & " - " & Err.Description
    Set GetRangeForPages = Nothing
End Function

' ============================================================================================
' LOGGING
' ============================================================================================
Private Sub LogEvent(ByVal msg As String)
    Dim stm As Object
    Dim existing As String
    Dim logLine As String
    Dim logPath As String

    logPath = LOG_FILE
    logLine = CleanTextForUTF8(Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & msg & vbCrLf)
    
    Set stm = CreateObject("ADODB.Stream")

    On Error GoTo ErrHandler

    ' ------------------------------------------------------------
    ' If the log file exists, read it first (UTF-8), then append.
    ' ------------------------------------------------------------
    If Dir(logPath) <> "" Then
        stm.Type = 2                ' adTypeText
        stm.Charset = "UTF-8"
        stm.Open
        stm.LoadFromFile logPath
        existing = stm.ReadText
        stm.Close
    Else
        existing = ""
    End If

    ' ------------------------------------------------------------
    ' Write combined content back as UTF-8
    ' ------------------------------------------------------------
    stm.Open
    stm.Type = 2
    stm.Charset = "UTF-8"
    stm.WriteText existing & logLine
    stm.Position = 0
    stm.saveToFile logPath, 2       ' adSaveCreateOverWrite
    stm.Close

    Set stm = Nothing
    Exit Sub

ErrHandler:
    ' Fallback: at least try to write something
    Debug.Print "LogEvent UTF-8 ERROR: "; Err.Number; Err.Description
End Sub

' ============================================================================================
' FILE OUTPUT
' ============================================================================================
Private Sub WriteTextFile(ByVal filePath As String, ByVal content As String)
    ' Writes UTF-8 without BOM (Paratext-safe)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    On Error GoTo ErrHandler

    ' Configure stream for UTF-8 text
    stm.Type = 2                 ' adTypeText
    stm.Charset = "UTF-8"
    stm.Open

    stm.WriteText content
    stm.Position = 0

    ' Save to file (overwrite)
    stm.saveToFile filePath, 2   ' adSaveCreateOverWrite

    stm.Close
    Set stm = Nothing
    Exit Sub

ErrHandler:
    LogEvent "ERROR writing UTF-8 file: " & Err.Number & " - " & Err.Description
End Sub

' ============================================================================================
' END MODULE
' ============================================================================================

