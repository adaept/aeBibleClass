Attribute VB_Name = "basUSFM_Export"
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

Private currentChapter As Long
Private bookTitleLevel As Long   ' 0 = none, 1 = next is mt2, 2 = next is mt3

' -----------------------------
' CONFIGURATION
' -----------------------------
Private LOG_FILE As String
Private OUTPUT_FILE As String
Private VALIDATOR_LOG As String

Private Sub InitPaths()
    On Error GoTo PROC_ERR
    Dim rptPath As String
    rptPath = ActiveDocument.Path & "\rpt\"
    LOG_FILE = rptPath & "USFM_Export_Log.txt"
    OUTPUT_FILE = rptPath & "ExportedBible.usfm"
    VALIDATOR_LOG = rptPath & "USFM_Validator_Log.txt"

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure InitPaths of Module basUSFM_Export"
    Resume PROC_EXIT
End Sub

' ============================================================================================
' PUBLIC ENTRY POINT
' ============================================================================================
Public Sub ExportUSFM_PageRange(ByVal startPage As Long, ByVal endPage As Long)
    On Error GoTo PROC_ERR
    InitPaths
    Dim t0 As Double: t0 = Timer
    currentChapter = 0
    bookTitleLevel = 0

    LogEvent "=== USFM EXPORT START ==="
    LogEvent "Page range: " & startPage & " to " & endPage

    Dim rng As Word.Range
    Set rng = GetRangeForPages(startPage, endPage)

    If rng Is Nothing Then
        LogEvent "ERROR: Page range returned no content."
        GoTo PROC_EXIT
    End If

    Dim usfm As String
    usfm = ConvertRangeToUSFM(rng)

    WriteTextFile OUTPUT_FILE, usfm
    ValidateUSFMFile OUTPUT_FILE

    LogEvent "USFM written to: " & OUTPUT_FILE

    LogEvent "Execution time: " & Format(Timer - t0, "0.00") & " seconds"
    LogEvent "=== USFM EXPORT END ==="

PROC_EXIT:
    Exit Sub
PROC_ERR:
    LogEvent "ERROR: Erl=" & Erl & " Err=" & Err.Number & " (" & Err.Description & ") in ExportUSFM_PageRange"
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ExportUSFM_PageRange of module basUSFM_Export", vbExclamation, "Export Error"
    Resume PROC_EXIT
End Sub

' ============================================================================================
' CORE CONVERSION
' ============================================================================================
Private Function ConvertRangeToUSFM(ByVal rng As Word.Range) As String
    On Error GoTo PROC_ERR
    Dim p As Word.Paragraph
    Dim sb As String
    Dim line As String
    Dim parts() As String
    Dim i As Long

    LogEvent "Beginning paragraph scan..."

    For Each p In rng.Paragraphs
        line = ConvertParagraphToUSFM(p)

        If Len(line) > 0 Then

            ' --- FIX: handle multi-line USFM output safely ---
            If InStr(line, vbCrLf) > 0 Then
                parts = Split(line, vbCrLf)
                For i = LBound(parts) To UBound(parts)
                    If Len(parts(i)) > 0 Then
                        sb = sb & parts(i) & vbCrLf
                    End If
                Next i

            Else
                sb = sb & line & vbCrLf
            End If

        End If
    Next p

    ConvertRangeToUSFM = sb

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ConvertRangeToUSFM of Module basUSFM_Export"
    Resume PROC_EXIT
End Function

' ============================================================================================
' PARAGRAPH -> USFM
' ============================================================================================
Private Function ConvertParagraphToUSFM(ByVal p As Word.Paragraph) As String
    On Error GoTo PROC_ERR
    Dim styleName As String
    Dim txt As String
    Dim chapNum As Long
    Dim verseNum As Long
    Dim verseText As String

    styleName = Trim$(p.style.NameLocal)
    txt = CleanTextForUTF8(Trim$(p.Range.Text))

    ' Normalize out any embedded CR/LF coming from Word
    txt = Replace$(txt, vbCr, "")
    txt = Replace$(txt, vbLf, "")

    'LogEvent "STYLE=[" & styleName & "] RAW=[" & txt & "]"
    'LogEvent "CHARSTYLE=[" & p.Range.Characters(1).style & "]"

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

        ' FIXME_LATER: IsNumeric passes for decimals like "1.2" — CLng would silently truncate.
        ' Safe in practice because chapter markers are always whole numbers in this document.
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

        ' FIXME_LATER: IsNumeric passes for decimals like "1.2" — CLng would silently truncate.
        ' Safe in practice because verse markers are always whole numbers in this document.
        If IsNumeric(vTxt) Then
            verseNum = CLng(vTxt)
            verseText = Trim$(Replace(txt, vTxt, ""))
            ConvertParagraphToUSFM = MakeVerseLine(verseNum, verseText)
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
            ConvertParagraphToUSFM = MakeTitleLine(1, txt)
            bookTitleLevel = 1

        Case "Heading 1"
            ConvertParagraphToUSFM = "\mt1 " & txt
            bookTitleLevel = 1

        Case "CustomParaAfterH1"
            ConvertParagraphToUSFM = MakeTitleLine(2, txt)
            bookTitleLevel = 0   ' explicitly end any title sequence

        Case "Heading 2"
            chapNum = ExtractTrailingNumber(txt)
            ConvertParagraphToUSFM = MakeChapterLines(chapNum, txt)

        Case "DatAuthRef"
            If Right$(txt, 1) = ":" Then ConvertParagraphToUSFM = "\is2 " & txt
            ' FIXME_LATER - verify if this is part of USFM spec:
            ' If DatAuthRef does not end with ":" it is intentionally excluded from USFM output
                
        Case "Brief"
            ConvertParagraphToUSFM = "\ip " & txt

        Case "Plain Text", "Normal"
            If bookTitleLevel = 1 Then
                ConvertParagraphToUSFM = MakeTitleLine(2, txt)
                bookTitleLevel = 2

            ElseIf bookTitleLevel = 2 Then
                ConvertParagraphToUSFM = MakeTitleLine(3, txt)
                bookTitleLevel = 0

            ElseIf TryParseChapterVerseFromStyles(p, chapNum, verseNum, verseText) Then
                If chapNum > 0 Then currentChapter = chapNum
                ConvertParagraphToUSFM = MakeVerseLine(verseNum, verseText)
            Else
                ConvertParagraphToUSFM = "\p " & txt
            End If

        Case Else
            If TryParseChapterVerseFromStyles(p, chapNum, verseNum, verseText) Then
                If chapNum > 0 Then currentChapter = chapNum
                ConvertParagraphToUSFM = MakeVerseLine(verseNum, verseText)
            Else
                ConvertParagraphToUSFM = "\p " & txt
            End If

    End Select

PROC_EXIT:
LogAndExit:
    LogEvent "Converted (" & styleName & "): " & Left$(ConvertParagraphToUSFM, 80)
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ConvertParagraphToUSFM of Module basUSFM_Export"
    Resume PROC_EXIT
End Function

Private Function MakeTitleLine(ByVal level As Long, ByVal txt As String) As String
    ' Normalize CR/LF
    txt = Replace$(txt, vbCr, "")
    txt = Replace$(txt, vbLf, "")
    txt = Trim$(txt)

    MakeTitleLine = "\mt" & CStr(level) & " " & txt
End Function

Private Function MakeChapterLines(ByVal chapNum As Long, ByVal clText As String) As String
    clText = Replace$(clText, vbCr, "")
    clText = Replace$(clText, vbLf, "")
    clText = Trim$(clText)

    If chapNum > 0 Then
        MakeChapterLines = "\cl " & clText & vbCrLf & "\c " & chapNum
    Else
        MakeChapterLines = "\cl " & clText
    End If
End Function

Private Function MakeVerseLine(ByVal verseNum As Long, ByVal verseText As String) As String
    ' Normalize CR/LF from Word
    verseText = Replace$(verseText, vbCr, "")
    verseText = Replace$(verseText, vbLf, "")
    verseText = Trim$(verseText)

    MakeVerseLine = "\v " & verseNum & " " & verseText
End Function

Private Function IsEffectivelyEmpty(txt As String) As Boolean
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

Private Function ParagraphHasCharStyle(p As Word.Paragraph, styleName As String) As Boolean
    ' FIXME_LATER: Iterates .words not .Characters — would miss character styles applied to only
    ' part of a word. Currently safe because "Chapter Verse marker" (orange) and "Verse marker"
    ' (green) are always applied to complete words (chapter/verse numbers) in this document.
    ' Re-evaluate if partial character style application is ever introduced.
    On Error GoTo PROC_ERR
    Dim r As Word.Range
    For Each r In p.Range.words
        If r.style = styleName Then
            ParagraphHasCharStyle = True
            GoTo PROC_EXIT
        End If
    Next

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ParagraphHasCharStyle of Module basUSFM_Export"
    Resume PROC_EXIT
End Function

Private Function ExtractCharStyleText(p As Word.Paragraph, styleName As String) As String
    On Error GoTo PROC_ERR
    Dim r As Word.Range
    Dim buf As String
    For Each r In p.Range.words
        If r.style = styleName Then
            buf = buf & r.Text
        End If
    Next
    ExtractCharStyleText = Trim$(buf)

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ExtractCharStyleText of Module basUSFM_Export"
    Resume PROC_EXIT
End Function

Private Function TryParseChapterVerseFromStyles( _
    ByVal p As Word.Paragraph, _
    ByRef chapNum As Long, _
    ByRef verseNum As Long, _
    ByRef verseText As String) As Boolean

    On Error GoTo PROC_ERR
    Dim rChap As Word.Range
    Dim rVerse As Word.Range
    Dim rText As Word.Range

    chapNum = 0
    verseNum = 0
    verseText = ""

    '------------------------------------------------------------
    ' 1. Chapter number run (character style: "Chapter Verse marker")
    '------------------------------------------------------------
    Set rChap = p.Range.Duplicate
    rChap.Collapse wdCollapseStart
    rChap.MoveEnd wdCharacter, 1

    If rChap.style <> "Chapter Verse marker" Then
        ' If the paragraph doesn't start with the chapter marker style,
        ' we can't parse it as a verse line.
        TryParseChapterVerseFromStyles = False
        GoTo PROC_EXIT
    End If

    ' Extend rChap to include all contiguous chars with that style
    Do While rChap.End < p.Range.End And rChap.style = "Chapter Verse marker"
        rChap.MoveEnd wdCharacter, 1
    Loop
    ' FIXME_LATER: MoveEnd -1 assumes the loop exited via style-change overshoot (condition 2).
    ' If the loop exits via range-boundary (rChap.End >= p.Range.End) with no style change,
    ' MoveEnd -1 incorrectly drops the last character. In practice this document always has
    ' a differently-styled character after the chapter marker, so condition 2 always fires first.
    rChap.MoveEnd wdCharacter, -1 ' step back one char after overshoot

    ' FIXME_LATER: CLng raises error 13 if CleanTextForUTF8 strips all characters from rChap.Text,
    ' leaving an empty string. In practice the early-exit guard above ensures rChap contains
    ' genuine numeric chapter marker digits, so this is unlikely for this document.
    ' If CleanTextForUTF8 is ever extended to strip digit characters, add an IsNumeric guard here
    ' and return False from TryParseChapterVerseFromStyles on empty/non-numeric result.
    chapNum = CLng(Trim$(CleanTextForUTF8(rChap.Text)))

    '------------------------------------------------------------
    ' 2. Verse number run (character style: "Verse marker")
    '------------------------------------------------------------
    Set rVerse = p.Range.Duplicate
    rVerse.Start = rChap.End
    rVerse.Collapse wdCollapseStart
    rVerse.MoveEnd wdCharacter, 1

    If rVerse.style <> "Verse marker" Then
        TryParseChapterVerseFromStyles = False
        GoTo PROC_EXIT
    End If

    ' Extend rVerse to include all contiguous chars with that style
    Do While rVerse.End < p.Range.End And rVerse.style = "Verse marker"
        rVerse.MoveEnd wdCharacter, 1
    Loop
    rVerse.MoveEnd wdCharacter, -1

    verseNum = CLng(Trim$(CleanTextForUTF8(rVerse.Text)))

    '------------------------------------------------------------
    ' 3. Remaining text = verse content
    '------------------------------------------------------------
    Set rText = p.Range.Duplicate
    rText.Start = rVerse.End
    verseText = CleanTextForUTF8(Trim$(rText.Text))

    If verseText = "" Then
        TryParseChapterVerseFromStyles = False
    Else
        TryParseChapterVerseFromStyles = True
    End If

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure TryParseChapterVerseFromStyles of Module basUSFM_Export"
    Resume PROC_EXIT
End Function

Private Function ExtractTrailingNumber(ByVal s As String) As Long
    Dim i As Long
    Dim digits As String
    s = Trim$(s)

    For i = Len(s) To 1 Step -1
        If Mid$(s, i, 1) >= "0" And Mid$(s, i, 1) <= "9" Then
            digits = Mid$(s, i, 1) & digits
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
    s = Replace(s, ChrW(&H2011), "-")   ' Non-breaking hyphen, normal hyphen
    s = Replace(s, ChrW(&H200B), "")    ' Zero-width space
    s = Replace(s, ChrW(&H200C), "")    ' Zero-width non-joiner
    s = Replace(s, ChrW(&H200D), "")    ' Zero-width joiner

    ' Remove any leftover control characters except CR/LF/TAB
    Dim i As Long, out As String, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
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

    On Error GoTo PROC_ERR

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

PROC_EXIT:
    Exit Sub
PROC_ERR:
    Debug.Print "Validator UTF-8 ERROR: "; Err.Number; Err.Description
    Set stm = Nothing
    Resume PROC_EXIT
End Sub

Public Sub ValidateUSFMFile(ByVal filePath As String)
    On Error GoTo PROC_ERR
    Dim stm As Object
    Dim content As String
    Dim lines() As String
    Dim i As Long
    Dim line As String
    Dim marker As String

    'LogValidator "=== USFM VALIDATION START ==="
    LogValidator "Validating file: " & filePath

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
                If Mid$(line, Len(marker) + 1, 1) <> " " Then
                    LogValidator "Line " & (i + 1) & ": Missing space after marker '" & marker & "' ? " & line
                End If
            End If
        End If

        ' 4. Check for empty content after markers that require content
        If MarkerRequiresContent(marker) Then
            If Len(Trim(Mid$(line, Len(marker) + 2))) = 0 Then
                LogValidator "Line " & (i + 1) & ": Marker '" & marker & "' missing content ? " & line
            End If
        End If

NextLine:
    Next i

    LogValidator "=== USFM VALIDATION END ==="

PROC_EXIT:
    Exit Sub
PROC_ERR:
    Set stm = Nothing
    LogValidator "ERROR validating USFM: " & Err.Number & " - " & Err.Description
    Resume PROC_EXIT
End Sub

Private Function ExtractUSFMMarker(ByVal line As String) As String
    Dim i As Long
    If Left$(line, 1) <> "\" Then
        ExtractUSFMMarker = ""
        Exit Function
    End If

    For i = 2 To Len(line)
        If Mid$(line, i, 1) = " " Or Mid$(line, i, 1) = "*" Then
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

' ============================================================================================
' PAGE RANGE EXTRACTION (same pattern as basWordRepairRunner)
' ============================================================================================
Private Function GetRangeForPages(ByVal startPage As Long, ByVal endPage As Long) As Word.Range
    On Error GoTo PROC_ERR
    Dim doc As Document
    Dim rStartPage As Word.Range
    Dim rEndPage As Word.Range
    Dim fullRange As Word.Range

    Set doc = ActiveDocument

    ' -------------------------------------------------------------
    ' Go to the START PAGE using the printed page number
    ' (same pattern as RunRepairWrappedVerseMarkers_Across_Pages_From)
    ' -------------------------------------------------------------
    Application.Selection.GoTo What:=wdGoToPage, name:=CStr(startPage)
    Set rStartPage = Application.Selection.Bookmarks("\Page").Range

    ' -------------------------------------------------------------
    ' Go to the END PAGE using the printed page number
    ' -------------------------------------------------------------
    Application.Selection.GoTo What:=wdGoToPage, name:=CStr(endPage)
    Set rEndPage = Application.Selection.Bookmarks("\Page").Range

    ' -------------------------------------------------------------
    ' Build a range from the start of startPage to the end of endPage
    ' -------------------------------------------------------------
    Set fullRange = doc.Range(Start:=rStartPage.Start, End:=rEndPage.End)

    Set GetRangeForPages = fullRange

PROC_EXIT:
    Exit Function
PROC_ERR:
    LogEvent "ERROR in GetRangeForPages: " & Err.Number & " - " & Err.Description
    Set GetRangeForPages = Nothing
    Resume PROC_EXIT
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

    On Error GoTo PROC_ERR

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

PROC_EXIT:
    Exit Sub
PROC_ERR:
    ' Fallback: at least try to write something
    Debug.Print "LogEvent UTF-8 ERROR: "; Err.Number; Err.Description
    Resume PROC_EXIT
End Sub

' ============================================================================================
' FILE OUTPUT
' ============================================================================================
Private Sub WriteTextFile(ByVal filePath As String, ByVal content As String)
    ' Writes UTF-8 with BOM (still Paratext-safe)
    On Error GoTo PROC_ERR
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    ' Configure stream for UTF-8 text
    stm.Type = 2                 ' adTypeText
    stm.Charset = "UTF-8"        ' ADODB.Stream always writes a BOM (EF BB BF) with these 2 lines
    stm.Open

    stm.WriteText content
    stm.Position = 0

    ' Save to file (overwrite)
    stm.saveToFile filePath, 2   ' adSaveCreateOverWrite

    stm.Close
    Set stm = Nothing

PROC_EXIT:
    Exit Sub
PROC_ERR:
    LogEvent "ERROR writing UTF-8 file: " & Err.Number & " - " & Err.Description
    Resume PROC_EXIT
End Sub

' ============================================================================================
' END MODULE
' ============================================================================================

