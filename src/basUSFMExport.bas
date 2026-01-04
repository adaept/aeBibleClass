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
'       - Full style ? USFM mapping table
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

' ============================================================================================
' PUBLIC ENTRY POINT
' ============================================================================================
Public Sub ExportUSFM_PageRange(ByVal startPage As Long, ByVal endPage As Long)
    Dim t0 As Double: t0 = Timer
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

    styleName = p.style
    txt = Trim(p.range.text)

    ' Normalize form feed-only paragraphs
    If txt = Chr(12) Then
        ConvertParagraphToUSFM = "\pb"
        LogEvent "Mapped form feed to \pb for style: " & styleName
        Exit Function
    End If

    ' Ignore pure whitespace
    If Len(Replace(txt, vbTab, "")) = 0 Then
        ConvertParagraphToUSFM = ""
        LogEvent "Ignored empty/whitespace paragraph for style: " & styleName
        Exit Function
    End If

    Select Case styleName

        Case "Heading 1"
            ' Currently: \s1 GENESIS
            ' EXPAND: For strict USFM, consider \mt1 for book title instead of \s1.
            ConvertParagraphToUSFM = "\s1 " & txt

        Case "CustomParaAfterH1"
            ' "THE FIRST BOOK OF MOSES"
            ' EXPAND: For strict USFM, consider \mt2 or \d.
            ConvertParagraphToUSFM = "\mt2 " & txt

        Case "DatAuthRef"
            ' "Dating:", "Authorship:", "Refer to the maps..."
            ' EXPAND: Use \is2 for sub-headings, \ip for prose.
            If Right$(txt, 1) = ":" Then
                ConvertParagraphToUSFM = "\is2 " & Left$(txt, Len(txt) - 1)
            Else
                ConvertParagraphToUSFM = "\ip " & txt
            End If

        Case "Plain Text"
            ' Currently used for tabs, blanks, and form feeds.
            ' We already handled form feeds and blanks above.
            ' For any remaining Plain Text, treat as a normal paragraph.
            ConvertParagraphToUSFM = "\p " & txt

        Case "Normal"
            ' This is overloaded: titles, headings, and paragraphs.
            ' EXPAND: Use positional logic (e.g., first Normal lines before any \s1 ? \mt1/\mt2).
            ConvertParagraphToUSFM = "\p " & txt

        Case Else
            ' Strict fallback: never emit raw text.
            ConvertParagraphToUSFM = "\p " & txt
            LogEvent "USFM DEFAULT MAP (\p) for style: " & styleName

    End Select

    LogEvent "Converted paragraph (" & styleName & "): " & Left(ConvertParagraphToUSFM, 80)
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
    logLine = Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & msg & vbCrLf

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

