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
    styleName = p.style

    ' ---------------------------------------------------------
    ' EXPAND HERE:
    '   Map your Word styles to USFM markers.
    '   This is the core of the exporter.
    ' ---------------------------------------------------------
    Select Case styleName

        Case "Verse"
            ' EXPAND: Add verse-number extraction
            ConvertParagraphToUSFM = "\v " & Trim(p.range.text)

        Case "Paragraph"
            ConvertParagraphToUSFM = "\p " & Trim(p.range.text)

        Case "Heading 1"
            ConvertParagraphToUSFM = "\s1 " & Trim(p.range.text)

        Case "Heading 2"
            ConvertParagraphToUSFM = "\s2 " & Trim(p.range.text)

        Case "Poetry 1"
            ConvertParagraphToUSFM = "\q1 " & Trim(p.range.text)

        Case "Poetry 2"
            ConvertParagraphToUSFM = "\q2 " & Trim(p.range.text)

        Case Else
            ' EXPAND: Add more mappings
            ConvertParagraphToUSFM = Trim(p.range.text)

    End Select

    LogEvent "Converted paragraph (" & styleName & "): " & Left(ConvertParagraphToUSFM, 80)
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
    Dim f As Integer: f = FreeFile
    Open LOG_FILE For Append As #f
    Print #f, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & msg
    Close #f
End Sub

' ============================================================================================
' FILE OUTPUT
' ============================================================================================
Private Sub WriteTextFile(ByVal filePath As String, ByVal content As String)
    Dim f As Integer: f = FreeFile
    Open filePath For Output As #f
    Print #f, content
    Close #f
End Sub

' ============================================================================================
' END MODULE
' ============================================================================================

