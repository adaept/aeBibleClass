Attribute VB_Name = "basWordRepairRunner"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString
Private OneVersePerParaRepair As Boolean

Private mRevSuspects()  As String
Private mRevIdx         As Long
Private mRevTotal       As Long
Private mRevLoaded      As Boolean

'==============================================================================
' Soft-hyphen sweep - module constants (REFERENCE ONLY)
' Reference layout: "JUDE - Sample.docm" (JIS B5, w:code=13), assuming
' mirrored pages with the binding gutter on the inside edge of each leaf.
'   Page size      516.25 pt x 728.65 pt  (twips 10325 x 14573)
'   Margins        T/B 54.7, inside 54.7, outside 43.2
'   Binding gutter 10.1 pt (on inside of each page)
'   Two columns    14.4 pt gap, each ~196.925 pt wide
' These constants are NOT consumed by the sweep at runtime - SoftHyphen_*
' routines read PageSetup at runtime via GetColumnBoundsForPage. The
' constants are kept as the expected-baseline reference; if the live
' document drifts from these values SoftHyphen_DiagnoseLayout will surface
' the difference for review.
'
' Per-page geometry (mirrored, JIS B5 reference):
'   ODD pages (recto): inside on LEFT
'     Left col   64.8  ..  261.725
'     Gutter    261.725 ..  276.125
'     Right col 276.125 ..  473.05
'   EVEN pages (verso): inside on RIGHT
'     Left col   43.2  ..  240.125
'     Gutter    240.125 ..  254.525
'     Right col 254.525 ..  451.45
'==============================================================================
Private Const PAGE_BODY_X_MIN       As Single = 64.8     ' odd page only
Private Const COL_LEFT_X_MAX        As Single = 261.725
Private Const GUTTER_X_MIN          As Single = 261.725
Private Const GUTTER_X_MAX          As Single = 276.125
Private Const COL_RIGHT_X_MIN       As Single = 276.125
Private Const PAGE_BODY_X_MAX       As Single = 473.05
Private Const PAGE_BODY_Y_MIN       As Single = 54.7
Private Const PAGE_BODY_Y_MAX       As Single = 673.95

' Active vs Stray classification.
' Y(charAfter) - Y(softHyphen) > LINE_HEIGHT_TOLERANCE => active (line break,
' renders visibly, must be preserved). Otherwise => stray, removal candidate.
Private Const LINE_HEIGHT_TOLERANCE As Single = 4#

' Word "optional hyphen" (Ctrl+Hyphen). Find code "^-".
Private Const SOFT_HYPHEN_CODE      As Long = 31

' Soft-hyphen sweep mode (Q5: SH_RemoveAll dropped per design).
Public Enum SoftHyphenMode
    SH_PromptEach = 0   ' Yes / No / Cancel per Stray find
    SH_DryRunOnly = 1   ' Log only, no prompt, no removal
End Enum

'===============================================================
' Returns True if the active document filename starts with "v59"
'===============================================================
Private Function FileNameStartsWithV59() As Boolean
    Dim fileName As String

    ' Get filename only (no path)
    fileName = ActiveDocument.Name
    FileNameStartsWithV59 = (LCase$(Left$(fileName, 3)) = "v59")
End Function

Public Sub SaveAsPDF_NoOpen()
    ' Overwrite the existing PDF file silently - without prompting or warning
    On Error GoTo PROC_ERR
    Dim startTime As Single
    Dim endTime As Single
    Dim duration As Single
    Dim pdfPath As String

    ' Start timer
    startTime = Timer
    Debug.Print "Expected time ~130 seconds"

    pdfPath = "C:\adaept\aeBibleClass\Peter-USE REFINED English Bible CONTENTS.pdf"
    ActiveDocument.ExportAsFixedFormat OutputFileName:=pdfPath, _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False

    ' End timer
    endTime = Timer
    duration = endTime - startTime
    ' Print duration to Immediate Window
    Debug.Print "PDF export completed in " & Format(duration, "0.00") & " seconds."

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure SaveAsPDF_NoOpen of Module basWordRepairRunner"
    Resume PROC_EXIT
End Sub

Public Sub RunRepairWrappedVerseMarkers_Across_Pages_From(startPage As Long)
    On Error GoTo PROC_ERR
    Dim totalFixes As Long, pgFixCount As Long
    Dim numPages As Long: numPages = 0 ' Adjust if scanning more than one page

    Dim sessionID As String
    sessionID = Format(Now, "yyyyMMdd_HHmmss")

    Dim logPath As String
    logPath = "C:\adaept\aeBibleClass\rpt\RepairLog.txt"

    Dim logFile As Integer
    logFile = FreeFile

    ' Create file with header if it doesn't exist
    If Dir(logPath) = "" Then
        Open logPath For Output As #logFile
        Print #logFile, "SessionID,PageNum,Repairs"
        Close #logFile
    End If

    If FileNameStartsWithV59 Then
        OneVersePerParaRepair = False
    Else
        OneVersePerParaRepair = True
    End If
    Debug.Print "OneVersePerParaRepair = " & OneVersePerParaRepair

    ' Append results
    Open logPath For Append As #logFile
    Dim p As Long
    For p = startPage To startPage + numPages
        pgFixCount = 0
        RepairWrappedVerseMarkers_MergedPrefix_ByColumnContext_SinglePage p, pgFixCount
        Print #logFile, sessionID & "," & p & "," & pgFixCount
        totalFixes = totalFixes + pgFixCount
    Next p
    Close #logFile

    'MsgBox "Repair complete. CSV log updated at:" & vbCrLf & logPath, vbInformation
    Selection.GoTo What:=wdGoToPage, name:=CStr(startPage)

PROC_EXIT:
    Exit Sub
PROC_ERR:
    If logFile > 0 Then Close #logFile
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RunRepairWrappedVerseMarkers_Across_Pages_From of Module basWordRepairRunner"
    Resume PROC_EXIT
End Sub

Public Sub RepairWrappedVerseMarkers_MergedPrefix_ByColumnContext_SinglePage(pageNum As Long, ByRef fixCount As Long)
    ' Same logic as full macro, but suppresses MsgBox and passes fixCount by reference.
    ' Copy the full body from RepairWrappedVerseMarkers_MergedPrefix_ByColumnContext here
    ' And replace `MsgBox` line with: fixCount = fixCount
    On Error GoTo PROC_ERR
    Dim pgRange As Word.Range, ch As Word.Range, ScanRange As Word.Range, prefixCh As Word.Range
    Dim pageStart As Long, pageEnd As Long
    Dim chapterMarker As String, verseDigits As String, combinedNumber As String
    Dim markerStart As Long, markerEnd As Long, verseEnd As Long
    Dim prefixTxt As String, prefixStyle As String, prefixAsc As Variant
    Dim prefixY As Single, digitY As Single, digitX As Single
    Dim nextWords As String, lookAhead As Word.Range, token As Word.Range, wCount As Integer
    Dim logBuffer As String
    Const ASCII_FORMFEED As Long = 12   ' Chr(12) — form feed; appears in malformed verse markers
    Dim ascii12Count As Long
    Dim ascii160MissingCount As Long
    Dim suffix160Count As Long
    Dim suffixHairSpaceCount As Long
    Dim suffixSpaceCount As Long
    Dim suffixOtherCount As Long
    Dim ascii13InsertCount As Long
    
    fixCount = 0
    logBuffer = "=== Smart Prefix Repair on Page " & pageNum & " ===" & vbCrLf

    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNum))
    pageStart = pgRange.Start
    ' FIXME_LATER: ActiveDocument.Pages.Count forces full pagination on large documents
    ' (800+ pages) and may cause noticeable delay if this routine runs per-page.
    ' Consider caching the page Count outside this function if performance is an issue.
    If pageNum >= ActiveDocument.Pages.Count Then
        pageEnd = ActiveDocument.Content.End
    Else
        Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNum + 1))
        pageEnd = pgRange.Start - 1
    End If

    Dim i As Long
    i = pageStart
    Dim headerText As String
    headerText = GetPageHeaderText(pageNum)
    'Debug.Print "Page " & pageNum & " header: " & headerText
    logBuffer = logBuffer & "Header for page " & pageNum & ": " & headerText & vbCrLf
    
    Do While i < pageEnd
        Set ch = ActiveDocument.Range(i, i + 1)
        If Len(Trim(ch.Text)) = 1 And IsNumeric(ch.Text) And ch.style.NameLocal = "Chapter Verse marker" And ch.Font.color = RGB(255, 165, 0) Then
            ' Assemble chapter marker block
            chapterMarker = ch.Text
            markerStart = i
            markerEnd = i + 1
            Do While markerEnd < pageEnd
                Set ScanRange = ActiveDocument.Range(markerEnd, markerEnd + 1)
                If Len(Trim(ScanRange.Text)) = 1 And IsNumeric(ScanRange.Text) Then
                    If ScanRange.style.NameLocal = "Chapter Verse marker" And ScanRange.Font.color = RGB(255, 165, 0) Then
                        chapterMarker = chapterMarker & ScanRange.Text
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
                Set ScanRange = ActiveDocument.Range(verseEnd, verseEnd + 1)
                If Len(Trim(ScanRange.Text)) = 1 And IsNumeric(ScanRange.Text) Then
                    If ScanRange.style.NameLocal = "Verse marker" And ScanRange.Font.color = RGB(80, 200, 120) Then
                        verseDigits = verseDigits & ScanRange.Text
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
    
                ' NEW: get verse text via helper function
                Dim verseText As String
                verseText = GetVerseText(pageEnd, verseEnd)
    
                'Dim chInfo As Word.Range
                'Set chInfo = ActiveDocument.Range(verseEnd, verseEnd + 1)
                'Debug.Print "Hair space font: " & chInfo.Font.Name & " | Size=" & chInfo.Font.Size & " | Style=" & chInfo.style.NameLocal & " | ASCII=" & AscW(chInfo.Text)
                
                Dim suffixCh As Word.Range
                Set suffixCh = ActiveDocument.Range(verseEnd, verseEnd + 1)
                Dim suffixAsc As Long
                suffixAsc = AscW(suffixCh.Text)

                Select Case suffixAsc
                    Case 160: suffix160Count = suffix160Count + 1
                    Case 8239: suffixHairSpaceCount = suffixHairSpaceCount + 1
                    Case 32: suffixSpaceCount = suffixSpaceCount + 1
                    Case Else: suffixOtherCount = suffixOtherCount + 1
                End Select

                ' Optional diagnostic
                'Debug.Print "Suffix [" & combinedNumber & "] ASCII=" & suffixAsc & " Style=" & suffixCh.style.NameLocal & " Font=" & suffixCh.Font.Name & " Size=" & suffixCh.Font.Size
                
                ' Chr(12) audit
                If Len(combinedNumber) = 1 And AscW(combinedNumber) = ASCII_FORMFEED Then
                    ascii12Count = ascii12Count + 1
                    i = verseEnd
                    Exit Do
                End If
                
                ' Prefix check
                If markerStart > pageStart Then
                    Set prefixCh = ActiveDocument.Range(markerStart - 1, markerStart)
                    prefixTxt = prefixCh.Text
                    prefixStyle = prefixCh.style.NameLocal
                    prefixAsc = AscW(prefixTxt)
                    Debug.Print headerText & " " & chapterMarker & ":" & verseDigits & vbTab & Replace(verseText, Chr(13), " ")   ',prefixAsc, combinedNumber

                    prefixY = prefixCh.Information(wdVerticalPositionRelativeToPage)

                    If (prefixAsc = 32 Or prefixAsc = 160) And Trim(prefixStyle) = "Normal" Then
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
                                logBuffer = logBuffer & "> Repaired prefix before '" & combinedNumber & "' @ X=" & Format(digitX, "0.0") & " | Break inserted | Next words:  " & Trim(nextWords) & " " & vbCrLf
                            Else
                                prefixCh.Text = ""
                                logBuffer = logBuffer & "> Removed space before '" & combinedNumber & "' @ X=" & Format(digitX, "0.0") & " | No break | Next words:  " & Trim(nextWords) & " " & vbCrLf
                            End If

                            fixCount = fixCount + 1
                        End If
                    End If
                
                    ' --- NEW: Ensure each verse starts on its own line (after repair logic) ---
                    'If markerStart > pageStart Then
                    Dim versePrefix As Word.Range
                    Set versePrefix = ActiveDocument.Range(markerStart - 1, markerStart)
    
                    If OneVersePerParaRepair Then
                        ' If the char before the marker is not already a CR, insert one
                        If AscW(versePrefix.Text) <> 13 Then
                            versePrefix.Text = versePrefix.Text & Chr(13)
                            ascii13InsertCount = ascii13InsertCount + 1
                            fixCount = fixCount + 1
                            'Debug.Print "> Inserted CR before " & combinedNumber & " on page " & pageNum
                            'logBuffer = logBuffer & "> Inserted CR before " & combinedNumber & " on page " & pageNum & vbCrLf
                        End If
                    End If
                ElseIf markerStart = pageStart Then
                    logBuffer = logBuffer & "Marker '" & combinedNumber & "' is at the very start of page " & pageNum & vbCrLf
                    Debug.Print headerText & " " & chapterMarker & ":" & verseDigits & vbTab & Trim(Replace(verseText, Chr(13), " "))    ',"SoP", combinedNumber
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
    logBuffer = logBuffer & vbCrLf & "ASCII 12 audit: " & ascii12Count & " marker(s) on page " & pageNum & " contain Chr(12)"
    logBuffer = logBuffer & vbCrLf & "ASCII 160 audit: " & ascii160MissingCount & " marker(s) on page " & pageNum & " missing Chr(160) suffix"
    logBuffer = logBuffer & vbCrLf & "ASCII 13 audit: " & ascii13InsertCount & " marker(s) on page " & pageNum & " inserted Chr(13)" & vbCrLf
    Debug.Print logBuffer
    'MsgBox fixCount & " marker(s) repaired on page " & pageNum & ".", vbInformation
    Selection.GoTo What:=wdGoToPage, name:=CStr(pageNum)

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RepairWrappedVerseMarkers_MergedPrefix_ByColumnContext_SinglePage of Module basWordRepairRunner"
    Resume PROC_EXIT
End Sub

Private Function GetPageHeaderText(pgNum As Long) As String
    On Error GoTo PROC_ERR
    Dim rng As Word.Range
    Dim sec As Word.Section
    Dim hdr As HeaderFooter

    ' Get range for the page
    Set rng = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pgNum))
    Set sec = rng.Sections(1)   ' Page belongs to exactly one Section

    ' Default to primary header
    Set hdr = sec.Headers(wdHeaderFooterPrimary)

    ' Note: Does not apply in this Bible doc
    ' If primary is empty, check for first-page or even-page headers
    'If Len(hdr.Range.Text) = 0 Then
    '    If sec.Headers(wdHeaderFooterFirstPage).Exists Then
    '        Set hdr = sec.Headers(wdHeaderFooterFirstPage)
    '    ElseIf sec.Headers(wdHeaderFooterEvenPages).Exists Then
    '        Set hdr = sec.Headers(wdHeaderFooterEvenPages)
    '    End If
    'End If

    ' Clean up the header text (Word stores an end-of-cell marker)
    GetPageHeaderText = TitleCase(Trim(Replace(hdr.Range.Text, Chr(13), " ")))

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GetPageHeaderText of Module basWordRepairRunner"
    Resume PROC_EXIT
End Function

Private Function TitleCase(ByVal txt As String) As String
    Dim words() As String
    Dim i As Integer

    ' Split the sentence into words
    words = Split(LCase(txt), " ")

    ' Capitalize each word
    For i = 0 To UBound(words)
        If Len(words(i)) > 0 Then
            words(i) = UCase(Left(words(i), 1)) & Mid$(words(i), 2)
        End If
    Next i

    ' Recombine the words into a sentence
    TitleCase = Join(words, " ")
End Function

Private Function GetVerseText(pageEnd As Long, verseContentStart As Long) As String
    On Error GoTo PROC_ERR
    Dim verseContentEnd As Long
    Dim nextPos As Long
    Dim scanCh As Word.Range
    Dim txt As String

    verseContentEnd = pageEnd
    nextPos = verseContentStart

    Do While nextPos < pageEnd
        Set scanCh = ActiveDocument.Range(nextPos, nextPos + 1)

        If Len(Trim(scanCh.Text)) = 1 And IsNumeric(scanCh.Text) Then
            If (scanCh.style.NameLocal = "Chapter Verse marker" And scanCh.Font.color = RGB(255, 165, 0)) _
               Or (scanCh.style.NameLocal = "Verse marker" And scanCh.Font.color = RGB(80, 200, 120)) Then
                verseContentEnd = nextPos
                Exit Do
            End If
        End If

        nextPos = nextPos + 1
    Loop

    txt = Trim(ActiveDocument.Range(verseContentStart, verseContentEnd).Text)

    Dim pos As Long
    pos = InStrRev(txt, "CHAPTER ")
    If pos > 0 Then
        txt = Trim$(Left$(txt, pos - 1))
    End If

    GetVerseText = txt

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GetVerseText of Module basWordRepairRunner"
    Resume PROC_EXIT
End Function

'==============================================================================
' ClassifyColumnAt
' Returns the column band for a horizontal position, using bounds resolved
' for the current page (which may differ between odd/even when mirrored).
' Bands: "OutsideLeft" | "Left" | "Gutter" | "Right" | "OutsideRight"
'==============================================================================
Private Function ClassifyColumnAt(ByVal xPos As Single, _
                                  ByVal bodyXMin As Single, _
                                  ByVal leftColMax As Single, _
                                  ByVal gutterMax As Single, _
                                  ByVal bodyXMax As Single) As String
    Select Case True
        Case xPos < bodyXMin:    ClassifyColumnAt = "OutsideLeft"
        Case xPos < leftColMax:  ClassifyColumnAt = "Left"
        Case xPos < gutterMax:   ClassifyColumnAt = "Gutter"
        Case xPos < bodyXMax:    ClassifyColumnAt = "Right"
        Case Else:               ClassifyColumnAt = "OutsideRight"
    End Select
End Function

'==============================================================================
' GetColumnBoundsForPage
' Resolves the actual column-X boundaries for a given page, reading PageSetup
' at runtime so the Result is correct under mirrored margins (where odd pages
' have the binding gutter on the LEFT and even pages on the RIGHT) and under
' arbitrary column widths.
' Convention: odd pageNum = recto = inside-on-left.
'==============================================================================
Private Sub GetColumnBoundsForPage(ByVal pageNum As Long, _
                                   ByRef bodyXMin As Single, _
                                   ByRef bodyXMax As Single, _
                                   ByRef leftColMax As Single, _
                                   ByRef gutterMin As Single, _
                                   ByRef gutterMax As Single, _
                                   ByRef rightColMin As Single)
    Dim ps             As Word.PageSetup
    Dim pgRange        As Word.Range
    Dim insideMargin   As Single
    Dim outsideMargin  As Single
    Dim col1Width      As Single
    Dim colGap         As Single
    Dim insideOnLeft   As Boolean

    ' Resolve PageSetup for THIS page's section, not the document default.
    ' Document.PageSetup is ambiguous when sections differ; reading per-page
    ' avoids that and the pagination stalls it can trigger.
    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNum))
    Set ps = pgRange.Sections(1).PageSetup

    insideMargin = ps.leftMargin + ps.gutter
    outsideMargin = ps.rightMargin

    If ps.MirrorMargins Then
        insideOnLeft = ((pageNum Mod 2) = 1)   ' odd = recto = inside on left
    Else
        insideOnLeft = True
    End If

    If insideOnLeft Then
        bodyXMin = insideMargin
        bodyXMax = ps.pageWidth - outsideMargin
    Else
        bodyXMin = outsideMargin
        bodyXMax = ps.pageWidth - insideMargin
    End If

    If ps.TextColumns.Count >= 2 Then
        col1Width = ps.TextColumns(1).Width
        colGap = ps.TextColumns(1).SpaceAfter
        leftColMax = bodyXMin + col1Width
        gutterMin = leftColMax
        gutterMax = leftColMax + colGap
        rightColMin = gutterMax
    Else
        ' Single-column layout: collapse gutter to nothing.
        leftColMax = bodyXMax
        gutterMin = bodyXMax
        gutterMax = bodyXMax
        rightColMin = bodyXMax
    End If
End Sub

'==============================================================================
' SoftHyphen_DiagnoseLayout
' Read-only. Prints PageSetup for the active document plus the computed
' column-X boundaries for an odd page and an even page. Use this to confirm
' the mirroring state and column geometry before running the calibration or
' production sweep. The constants at the top of this module document the
' expected JUDE / JIS B5 reference values; this routine surfaces any drift.
'==============================================================================
Public Sub SoftHyphen_DiagnoseLayout()
    On Error GoTo PROC_ERR
    Dim oDoc           As Word.Document
    Dim sec            As Word.Section
    Dim ps             As Word.PageSetup
    Dim tc             As Word.TextColumns
    Dim i              As Long
    Dim secIdx         As Long
    Dim secCount       As Long
    Dim startPage      As Long
    Dim oneColCount    As Long
    Dim twoColCount    As Long
    Dim otherColCount  As Long
    Dim mirroredCount  As Long
    Dim anomalyCount   As Long
    Dim anomalyList    As String
    Dim sPath          As String
    Dim f              As Integer
    Const NL           As String = vbCrLf

    ' Reference geometry for "standard" 2-column Bible-body sections.
    Const STD_COL_WIDTH As Single = 196.9
    Const STD_COL_GAP   As Single = 14.4
    Const TOL           As Single = 0.5

    Set oDoc = ActiveDocument
    secCount = oDoc.Sections.Count

    Dim docDir As String
    docDir = oDoc.Path
    If Len(docDir) = 0 Then docDir = Environ$("TEMP")
    sPath = docDir & "\rpt\SoftHyphen_Layout.log"
    f = FreeFile
    Open sPath For Output As #f

    Print #f, "=== Sections in " & oDoc.Name & "  (" & secCount & " total) ==="
    Print #f, "Run: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    Print #f, "First and last sections are typically blank-template holders."
    Print #f, "Bible body sections (VerseText) should report 2 columns + Mirror=True."

    For Each sec In oDoc.Sections
        secIdx = secIdx + 1
        Set ps = sec.PageSetup
        Set tc = ps.TextColumns

        On Error Resume Next
        startPage = sec.Range.Information(wdActiveEndPageNumber)
        If Err.Number <> 0 Then startPage = -1
        Err.Clear
        On Error GoTo PROC_ERR

        ' Tallies.
        Select Case tc.Count
            Case 1: oneColCount = oneColCount + 1
            Case 2: twoColCount = twoColCount + 1
            Case Else: otherColCount = otherColCount + 1
        End Select
        If ps.MirrorMargins Then mirroredCount = mirroredCount + 1

        ' Anomaly check for 2-column sections only.
        Dim anomFlag As String
        anomFlag = ""
        If tc.Count = 2 Then
            If Abs(tc(1).Width - STD_COL_WIDTH) > TOL Or _
               Abs(tc(1).SpaceAfter - STD_COL_GAP) > TOL Or _
               Abs(tc(2).Width - STD_COL_WIDTH) > TOL Then
                anomFlag = "  [ANOMALY: non-standard 2-col geometry]"
                anomalyCount = anomalyCount + 1
                anomalyList = anomalyList & "  Section " & secIdx & " (page " & _
                              startPage & "): Col1=" & Format(tc(1).Width, "0.0") & _
                              "/" & Format(tc(1).SpaceAfter, "0.0") & _
                              " Col2=" & Format(tc(2).Width, "0.0") & NL
            End If
        End If

        Print #f, ""
        Print #f, "-- Section " & secIdx & " of " & secCount & _
                  "  (starts on page " & startPage & ") --" & anomFlag
        Print #f, "  PageSize  : " & Format(ps.pageWidth, "0.0") & " x " & _
                  Format(ps.PageHeight, "0.0") & " pt"
        Print #f, "  Margins   : T=" & Format(ps.TopMargin, "0.0") & _
                  "  B=" & Format(ps.BottomMargin, "0.0") & _
                  "  L=" & Format(ps.leftMargin, "0.0") & _
                  "  R=" & Format(ps.rightMargin, "0.0")
        Print #f, "  Gutter    : " & Format(ps.gutter, "0.0") & _
                  "  GutterPos=" & ps.GutterPos & _
                  "  Mirror=" & ps.MirrorMargins
        Print #f, "  Columns   : " & tc.Count & _
                  "  EvenlySpaced=" & tc.EvenlySpaced & _
                  "  LineBetween=" & tc.LineBetween
        For i = 1 To tc.Count
            If i < tc.Count Then
                Print #f, "    Col " & i & ": Width=" & Format(tc(i).Width, "0.0") & _
                          "  SpaceAfter=" & Format(tc(i).SpaceAfter, "0.0")
            Else
                Print #f, "    Col " & i & ": Width=" & Format(tc(i).Width, "0.0")
            End If
        Next i
    Next sec

    Print #f, ""
    Print #f, "-- Summary --"
    Print #f, "  Total sections : " & secCount
    Print #f, "  1-column       : " & oneColCount
    Print #f, "  2-column       : " & twoColCount
    Print #f, "  Other columns  : " & otherColCount
    Print #f, "  Mirrored       : " & mirroredCount
    Print #f, "  Anomalies      : " & anomalyCount & " (2-col sections deviating from std " & _
              Format(STD_COL_WIDTH, "0.0") & "/" & Format(STD_COL_GAP, "0.0") & "/" & _
              Format(STD_COL_WIDTH, "0.0") & ")"
    If anomalyCount > 0 Then
        Print #f, ""
        Print #f, "-- Anomalies --"
        Print #f, anomalyList;
    End If

    Print #f, ""
    Print #f, "-- Reference baseline (JUDE / JIS B5, ODD page) --"
    Print #f, "  Body  " & Format(PAGE_BODY_X_MIN, "0.0") & " .. " & _
              Format(PAGE_BODY_X_MAX, "0.0")
    Print #f, "  Left  " & Format(PAGE_BODY_X_MIN, "0.0") & " .. " & _
              Format(COL_LEFT_X_MAX, "0.0")
    Print #f, "  Gut   " & Format(GUTTER_X_MIN, "0.0") & " .. " & _
              Format(GUTTER_X_MAX, "0.0")
    Print #f, "  Right " & Format(COL_RIGHT_X_MIN, "0.0") & " .. " & _
              Format(PAGE_BODY_X_MAX, "0.0")

    Close #f
    f = 0

    Debug.Print "SoftHyphen_DiagnoseLayout: " & secCount & " section(s) - " & _
                oneColCount & " single-col, " & twoColCount & " two-col, " & _
                mirroredCount & " mirrored, " & anomalyCount & " anomal(ies) -> " & sPath
    MsgBox "Layout diagnostic written:" & NL & _
           "  rpt\SoftHyphen_Layout.log" & NL & NL & _
           "Total sections : " & secCount & NL & _
           "  1-column     : " & oneColCount & NL & _
           "  2-column     : " & twoColCount & NL & _
           "  Other        : " & otherColCount & NL & _
           "  Mirrored     : " & mirroredCount & " of " & secCount & NL & _
           "  Anomalies    : " & anomalyCount, _
           vbInformation, "SoftHyphen_DiagnoseLayout"

PROC_EXIT:
    If f > 0 Then
        On Error Resume Next
        Close #f
        On Error GoTo 0
    End If
    Exit Sub
PROC_ERR:
    If f > 0 Then
        On Error Resume Next
        Close #f
        On Error GoTo 0
    End If
    MsgBox "Section " & secIdx & ": Error " & Err.Number & " (" & Err.Description & _
           ") in procedure SoftHyphen_DiagnoseLayout of Module basWordRepairRunner"
    Resume PROC_EXIT
End Sub

'==============================================================================
' SoftHyphen_CalibrateColumns
' PURPOSE:
'   Calibration helper. Walks one page, locates every soft hyphen
'   (Chr(31), Word "optional hyphen"), records X / Y / next-char-Y, classifies
'   the column (Left / Gutter / Right / Outside) and the disposition
'   (Active = at a real line break, Stray = invisible inside a line, OutsideBody
'   = in gutter or page margin). Writes one row per find to
'   rpt\SoftHyphenCalibration.csv. No removals - read-only.
'
'   Run this once against a representative page in the active document, review
'   the CSV, and confirm the column-X constants and LINE_HEIGHT_TOLERANCE match
'   the live layout before running the production sweep.
'
' Usage:
'   SoftHyphen_CalibrateColumns 42
'==============================================================================
Public Sub SoftHyphen_CalibrateColumns(ByVal pageNum As Long)
    Dim currentStep   As String
    On Error GoTo PROC_ERR
    Dim oDoc          As Word.Document
    Dim pgRange       As Word.Range
    Dim searchRng     As Word.Range
    Dim nextCh        As Word.Range
    Dim ctxRng        As Word.Range
    Dim pageStart     As Long, pageEnd As Long
    Dim ctxStart      As Long, ctxEnd As Long
    Dim xPos          As Single, yShy As Single, yNext As Single, yDelta As Single
    Dim col           As String, disposition As String, ctx As String
    Dim findCount     As Long
    Dim activeCount   As Long
    Dim strayCount    As Long
    Dim outsideCount  As Long
    Dim csvPath       As String
    Dim logPath       As String
    Dim f             As Integer
    Dim logF          As Integer
    Const NL          As String = vbCrLf

    currentStep = "Set ActiveDocument"
    Set oDoc = ActiveDocument

    ' Compute page bounds without touching Pages.Count.
    ' GoTo(wdGoToPage, pageNum+1) returns end-of-document when pageNum is
    ' the last page, so we can derive pageEnd unconditionally.
    currentStep = "GoTo page " & pageNum
    Set pgRange = oDoc.GoTo(What:=wdGoToPage, name:=CStr(pageNum))
    pageStart = pgRange.Start

    currentStep = "GoTo page " & (pageNum + 1)
    Set pgRange = oDoc.GoTo(What:=wdGoToPage, name:=CStr(pageNum + 1))
    Dim pageNextStart As Long
    pageNextStart = pgRange.Start

    If pageNextStart > pageStart Then
        pageEnd = pageNextStart - 1
        ' Cap at content end in case GoTo returned a position past it.
        If pageEnd > oDoc.Content.End Then pageEnd = oDoc.Content.End
    Else
        ' pageNum is the last (or only) page.
        pageEnd = oDoc.Content.End
    End If

    ' Resolve column bounds for this specific page (handles mirrored layout).
    currentStep = "Resolve column bounds for page " & pageNum
    Dim bodyXMin As Single, bodyXMax As Single
    Dim leftColMax As Single, gutterMin As Single, gutterMax As Single, rightColMin As Single
    GetColumnBoundsForPage pageNum, bodyXMin, bodyXMax, _
                           leftColMax, gutterMin, gutterMax, rightColMin

    Dim pageSide As String
    Dim isMirrored As Boolean
    Dim sectionPS As Word.PageSetup
    Set sectionPS = pgRange.Sections(1).PageSetup
    isMirrored = sectionPS.MirrorMargins
    If isMirrored Then
        pageSide = IIf(pageNum Mod 2 = 1, "Recto", "Verso")
    Else
        pageSide = "Single"
    End If

    ' Open report files. Both append-mode so successive per-page calls
    ' accumulate. CSV header is written only on first run (file empty).
    ' Use literal "\" - Application.PathSeparator avoided for compatibility.
    currentStep = "Open report files"
    Dim docDir As String
    docDir = oDoc.Path
    If Len(docDir) = 0 Then docDir = Environ$("TEMP")
    csvPath = docDir & "\rpt\SoftHyphenCalibration.csv"
    logPath = docDir & "\rpt\SoftHyphenCalibration.log"

    Dim writeCsvHeader As Boolean
    writeCsvHeader = (Len(Dir(csvPath)) = 0)

    f = FreeFile
    Open csvPath For Append As #f
    If writeCsvHeader Then
        Print #f, "PageNum,Side,FindNum,Position,X,Y,YNext,YDelta,Column,Disposition,Context"
    End If

    logF = FreeFile
    Open logPath For Append As #logF
    Print #logF, String(72, "=")
    Print #logF, "SoftHyphen_CalibrateColumns  page=" & pageNum & _
                 "  run=" & Format(Now, "yyyy-mm-dd hh:mm:ss")
    Print #logF, "Document : " & oDoc.Name
    Print #logF, "Section  : Mirrored=" & isMirrored & "  Side=" & pageSide & _
                 "  Cols=" & sectionPS.TextColumns.Count
    Print #logF, "Bounds   : Body=[" & Format(bodyXMin, "0.0") & ".." & _
                 Format(bodyXMax, "0.0") & "]  L=[" & _
                 Format(bodyXMin, "0.0") & ".." & Format(leftColMax, "0.0") & "]  G=[" & _
                 Format(gutterMin, "0.0") & ".." & Format(gutterMax, "0.0") & "]  R=[" & _
                 Format(rightColMin, "0.0") & ".." & Format(bodyXMax, "0.0") & "]"
    Print #logF, "PageRange: [" & pageStart & ".." & pageEnd & "]"

    ' Build search range and configure Find.
    currentStep = "Build search range"
    Set searchRng = oDoc.Range(pageStart, pageEnd)
    currentStep = "Configure Find"
    With searchRng.Find
        .ClearFormatting
        .Text = Chr(SOFT_HYPHEN_CODE)
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWholeWord = False
        .MatchWildcards = False
    End With

    ' Iteration: after each Execute success, searchRng IS the match.
    ' Collapse to end and re-extend End to pageEnd so the next Execute
    ' is bounded.
    Do
        currentStep = "Find.Execute (find #" & (findCount + 1) & ")"
        If Not searchRng.Find.Execute Then Exit Do
        If searchRng.Start >= pageEnd Then Exit Do

        findCount = findCount + 1

        currentStep = "Information(X) at find #" & findCount
        xPos = CSng(searchRng.Information(wdHorizontalPositionRelativeToPage))
        currentStep = "Information(Y) at find #" & findCount
        yShy = CSng(searchRng.Information(wdVerticalPositionRelativeToPage))

        col = ClassifyColumnAt(xPos, bodyXMin, leftColMax, gutterMax, bodyXMax)

        ' Y of the next character (1-char range past the match).
        currentStep = "Information(Y next) at find #" & findCount
        If searchRng.End < oDoc.Content.End Then
            Set nextCh = oDoc.Range(searchRng.End, searchRng.End + 1)
            yNext = CSng(nextCh.Information(wdVerticalPositionRelativeToPage))
        Else
            yNext = yShy
        End If
        yDelta = yNext - yShy

        ' Disposition.
        If col <> "Left" And col <> "Right" Then
            disposition = "OutsideBody"
            outsideCount = outsideCount + 1
        ElseIf yDelta > LINE_HEIGHT_TOLERANCE Then
            disposition = "Active"
            activeCount = activeCount + 1
        Else
            disposition = "Stray"
            strayCount = strayCount + 1
        End If

        ' Context window: 30 chars before, 30 chars after, control chars
        ' and quotes sanitised for CSV.
        currentStep = "Context window at find #" & findCount
        ctxStart = searchRng.Start - 30
        If ctxStart < pageStart Then ctxStart = pageStart
        ctxEnd = searchRng.End + 30
        If ctxEnd > pageEnd Then ctxEnd = pageEnd
        Set ctxRng = oDoc.Range(ctxStart, ctxEnd)
        ctx = ctxRng.Text
        ctx = Replace(ctx, Chr(13), " ")
        ctx = Replace(ctx, Chr(11), " ")
        ctx = Replace(ctx, Chr(10), " ")
        ctx = Replace(ctx, Chr(9), " ")
        ctx = Replace(ctx, Chr(SOFT_HYPHEN_CODE), "[SHY]")
        ctx = Replace(ctx, """", """""")

        currentStep = "Write CSV row at find #" & findCount
        Print #f, pageNum & "," & pageSide & "," & findCount & "," & _
                  searchRng.Start & "," & _
                  Format(xPos, "0.0") & "," & Format(yShy, "0.0") & "," & _
                  Format(yNext, "0.0") & "," & Format(yDelta, "0.0") & "," & _
                  col & "," & disposition & ",""" & ctx & """"

        ' Advance: collapse to end of match, re-extend to pageEnd so Find
        ' stays bounded on the next iteration.
        currentStep = "Advance after find #" & findCount
        searchRng.Collapse wdCollapseEnd
        If searchRng.Start >= pageEnd Then Exit Do
        searchRng.End = pageEnd

        If findCount Mod 50 = 0 Then DoEvents
    Loop

    Print #logF, "Result   : " & findCount & " find(s) - " & _
                 activeCount & " Active, " & strayCount & " Stray, " & _
                 outsideCount & " OutsideBody"
    Close #logF
    logF = 0
    Close #f
    f = 0

    ' One-line summary in Immediate so the user sees completion at a glance.
    Debug.Print "SoftHyphen_CalibrateColumns p" & pageNum & " (" & pageSide & "): " & _
                findCount & " find(s) - " & activeCount & " Active, " & _
                strayCount & " Stray, " & outsideCount & " OutsideBody"
    MsgBox "Soft Hyphen Calibration on page " & pageNum & " (" & pageSide & "):" & NL & _
           findCount & " total find(s)" & NL & _
           activeCount & " Active (line-breaking, would be KEPT)" & NL & _
           strayCount & " Stray (in-line, removal candidate)" & NL & _
           outsideCount & " OutsideBody (Gutter/Margin, skipped)" & NL & NL & _
           "rpt\SoftHyphenCalibration.csv  (per-find rows, append)" & NL & _
           "rpt\SoftHyphenCalibration.log  (per-page detail, append)", _
           vbInformation, "SoftHyphen_CalibrateColumns"

PROC_EXIT:
    If logF > 0 Then
        On Error Resume Next
        Close #logF
        On Error GoTo 0
    End If
    If f > 0 Then
        On Error Resume Next
        Close #f
        On Error GoTo 0
    End If
    Exit Sub
PROC_ERR:
    If logF > 0 Then
        On Error Resume Next
        Print #logF, "ABORTED at step [" & currentStep & "] - Err " & _
                     Err.Number & ": " & Err.Description
        Close #logF
        On Error GoTo 0
    End If
    If f > 0 Then
        On Error Resume Next
        Close #f
        On Error GoTo 0
    End If
    MsgBox "Step: [" & currentStep & "]" & vbCrLf & _
           "Error " & Err.Number & " (" & Err.Description & ")" & vbCrLf & _
           "in procedure SoftHyphen_CalibrateColumns of Module basWordRepairRunner"
    Resume PROC_EXIT
End Sub

'==============================================================================
' SoftHyphenSweep_ByColumnContext_SinglePage
' PURPOSE:
'   Production worker: scans one page for soft hyphens (Chr(31)), classifies
'   each as Active (line-breaking, kept), Stray (in-line, removal candidate),
'   or OutsideBody (gutter / margin, skipped). For each Stray, when mode is
'   SH_PromptEach, scrolls the find into view and asks Yes / No / Cancel.
'   Yes deletes the soft hyphen; No skips; Cancel sets userCancelled=True
'   and aborts. SH_DryRunOnly classifies and logs but does not prompt or
'   delete.
'
' OUTPUT (both append-mode):
'   rpt\SoftHyphenSweep.csv  (machine-readable, header on first run)
'   rpt\SoftHyphenSweep.log  (human-readable, one block per page)
'
' BYREF accumulators (driver passes the same vars across pages):
'   strayCum    - cumulative Stray-classified finds
'   removedCum  - cumulative removals
'   userCancelled - True iff user clicked Cancel; driver should stop
'==============================================================================
Public Sub SoftHyphenSweep_ByColumnContext_SinglePage( _
        ByVal pageNum As Long, _
        ByVal mode As SoftHyphenMode, _
        ByRef strayCum As Long, _
        ByRef removedCum As Long, _
        ByRef userCancelled As Boolean)
    Dim currentStep   As String
    On Error GoTo PROC_ERR
    Dim oDoc          As Word.Document
    Dim pgRange       As Word.Range
    Dim searchRng     As Word.Range
    Dim nextCh        As Word.Range
    Dim ctxRng        As Word.Range
    Dim m             As Word.Range
    Dim sectionPS     As Word.PageSetup
    Dim pageStart     As Long, pageEnd As Long, pageNextStart As Long
    Dim ctxStart      As Long, ctxEnd As Long
    Dim xPos          As Single, yShy As Single, yNext As Single, yDelta As Single
    Dim col           As String, ctx As String, disp As String, action As String
    Dim findIdx       As Long
    Dim activeCount   As Long, outsideCount As Long
    Dim pageStrayCount As Long, pageRemovedCount As Long, pageSkippedCount As Long
    Dim isMirrored    As Boolean
    Dim pageSide      As String
    Dim csvPath       As String, logPath As String
    Dim csvF          As Integer, logF As Integer
    Dim writeCsvHeader As Boolean
    Const NL          As String = vbCrLf

    ' Per-page Stray collection. Capture full record so pass 2 can write
    ' the CSV row after the user's action is known.
    Dim strayPos()    As Long
    Dim strayX()      As Single, strayY() As Single, strayYDelta() As Single
    Dim strayCol()    As String, strayCtx() As String
    ReDim strayPos(0 To 1023)
    ReDim strayX(0 To 1023), strayY(0 To 1023), strayYDelta(0 To 1023)
    ReDim strayCol(0 To 1023), strayCtx(0 To 1023)
    Dim nStray As Long

    currentStep = "Set ActiveDocument"
    Set oDoc = ActiveDocument

    ' Resolve page bounds without Pages.Count (multi-section safe).
    currentStep = "GoTo page " & pageNum
    Set pgRange = oDoc.GoTo(What:=wdGoToPage, name:=CStr(pageNum))
    pageStart = pgRange.Start

    currentStep = "GoTo page " & (pageNum + 1)
    Set pgRange = oDoc.GoTo(What:=wdGoToPage, name:=CStr(pageNum + 1))
    pageNextStart = pgRange.Start
    If pageNextStart > pageStart Then
        pageEnd = pageNextStart - 1
        If pageEnd > oDoc.Content.End Then pageEnd = oDoc.Content.End
    Else
        pageEnd = oDoc.Content.End
    End If

    ' Resolve column bounds for THIS page's section.
    currentStep = "Resolve column bounds for page " & pageNum
    Dim bodyXMin As Single, bodyXMax As Single
    Dim leftColMax As Single, gutterMin As Single, gutterMax As Single, rightColMin As Single
    GetColumnBoundsForPage pageNum, bodyXMin, bodyXMax, _
                           leftColMax, gutterMin, gutterMax, rightColMin

    Set sectionPS = pgRange.Sections(1).PageSetup
    isMirrored = sectionPS.MirrorMargins
    If isMirrored Then
        pageSide = IIf(pageNum Mod 2 = 1, "Recto", "Verso")
    Else
        pageSide = "Single"
    End If

    ' Open report files (append).
    currentStep = "Open report files"
    Dim docDir As String
    docDir = oDoc.Path
    If Len(docDir) = 0 Then docDir = Environ$("TEMP")
    csvPath = docDir & "\rpt\SoftHyphenSweep.csv"
    logPath = docDir & "\rpt\SoftHyphenSweep.log"

    writeCsvHeader = (Len(Dir(csvPath)) = 0)
    csvF = FreeFile
    Open csvPath For Append As #csvF
    If writeCsvHeader Then
        Print #csvF, "PageNum,Side,FindNum,Position,X,Y,YDelta,Column,Disposition,Action,Context"
    End If

    logF = FreeFile
    Open logPath For Append As #logF
    Print #logF, String(72, "=")
    Print #logF, "SoftHyphenSweep page=" & pageNum & "  side=" & pageSide & _
                 "  mode=" & IIf(mode = SH_DryRunOnly, "DryRun", "PromptEach") & _
                 "  run=" & Format(Now, "yyyy-mm-dd hh:mm:ss")
    Print #logF, "Bounds   : Body=[" & Format(bodyXMin, "0.0") & ".." & _
                 Format(bodyXMax, "0.0") & "]  L=[" & _
                 Format(bodyXMin, "0.0") & ".." & Format(leftColMax, "0.0") & "]  G=[" & _
                 Format(gutterMin, "0.0") & ".." & Format(gutterMax, "0.0") & "]  R=[" & _
                 Format(rightColMin, "0.0") & ".." & Format(bodyXMax, "0.0") & "]"

    ' ---- Pass 1: scan, classify, write Active/OutsideBody rows, capture Strays.
    currentStep = "Pass 1 - configure Find"
    Set searchRng = oDoc.Range(pageStart, pageEnd)
    With searchRng.Find
        .ClearFormatting
        .Text = Chr(SOFT_HYPHEN_CODE)
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWholeWord = False
        .MatchWildcards = False
    End With

    Do
        currentStep = "Pass 1 Find.Execute (find #" & (findIdx + 1) & ")"
        If Not searchRng.Find.Execute Then Exit Do
        If searchRng.Start >= pageEnd Then Exit Do

        findIdx = findIdx + 1

        currentStep = "Pass 1 Information at find #" & findIdx
        xPos = CSng(searchRng.Information(wdHorizontalPositionRelativeToPage))
        yShy = CSng(searchRng.Information(wdVerticalPositionRelativeToPage))
        col = ClassifyColumnAt(xPos, bodyXMin, leftColMax, gutterMax, bodyXMax)

        If searchRng.End < oDoc.Content.End Then
            Set nextCh = oDoc.Range(searchRng.End, searchRng.End + 1)
            yNext = CSng(nextCh.Information(wdVerticalPositionRelativeToPage))
        Else
            yNext = yShy
        End If
        yDelta = yNext - yShy

        ' Context window.
        currentStep = "Pass 1 context at find #" & findIdx
        ctxStart = searchRng.Start - 30
        If ctxStart < pageStart Then ctxStart = pageStart
        ctxEnd = searchRng.End + 30
        If ctxEnd > pageEnd Then ctxEnd = pageEnd
        Set ctxRng = oDoc.Range(ctxStart, ctxEnd)
        ctx = ctxRng.Text
        ctx = Replace(ctx, Chr(13), " ")
        ctx = Replace(ctx, Chr(11), " ")
        ctx = Replace(ctx, Chr(10), " ")
        ctx = Replace(ctx, Chr(9), " ")
        ctx = Replace(ctx, Chr(SOFT_HYPHEN_CODE), "[SHY]")
        ctx = Replace(ctx, """", """""")

        ' Disposition.
        If col <> "Left" And col <> "Right" Then
            disp = "OutsideBody"
            action = "Skipped"
            outsideCount = outsideCount + 1
            Print #csvF, pageNum & "," & pageSide & "," & findIdx & "," & _
                searchRng.Start & "," & Format(xPos, "0.0") & "," & _
                Format(yShy, "0.0") & "," & Format(yDelta, "0.0") & "," & _
                col & "," & disp & "," & action & ",""" & ctx & """"
        ElseIf yDelta > LINE_HEIGHT_TOLERANCE Then
            disp = "Active"
            action = "Kept"
            activeCount = activeCount + 1
            Print #csvF, pageNum & "," & pageSide & "," & findIdx & "," & _
                searchRng.Start & "," & Format(xPos, "0.0") & "," & _
                Format(yShy, "0.0") & "," & Format(yDelta, "0.0") & "," & _
                col & "," & disp & "," & action & ",""" & ctx & """"
        Else
            ' Stray - capture for pass 2; CSV row written after action known.
            If nStray > UBound(strayPos) Then
                ReDim Preserve strayPos(0 To UBound(strayPos) + 1024)
                ReDim Preserve strayX(0 To UBound(strayX) + 1024)
                ReDim Preserve strayY(0 To UBound(strayY) + 1024)
                ReDim Preserve strayYDelta(0 To UBound(strayYDelta) + 1024)
                ReDim Preserve strayCol(0 To UBound(strayCol) + 1024)
                ReDim Preserve strayCtx(0 To UBound(strayCtx) + 1024)
            End If
            strayPos(nStray) = searchRng.Start
            strayX(nStray) = xPos
            strayY(nStray) = yShy
            strayYDelta(nStray) = yDelta
            strayCol(nStray) = col
            strayCtx(nStray) = ctx
            nStray = nStray + 1
        End If

        searchRng.Collapse wdCollapseEnd
        If searchRng.Start >= pageEnd Then Exit Do
        searchRng.End = pageEnd

        If findIdx Mod 50 = 0 Then DoEvents
    Loop

    pageStrayCount = nStray

    ' ---- Pass 2: prompt + remove for each Stray.
    ' Cumulative deletion offset keeps captured positions accurate as we
    ' iterate forward and remove characters.
    Dim cumDelOffset As Long
    Dim i As Long, pos As Long, response As Long

    For i = 0 To pageStrayCount - 1
        If userCancelled Then Exit For

        currentStep = "Pass 2 stray #" & (i + 1) & " of " & pageStrayCount
        pos = strayPos(i) - cumDelOffset

        ' Defensive: confirm a soft hyphen still sits at the resolved position.
        Set m = oDoc.Range(pos, pos + 1)
        If AscW(m.Text) <> SOFT_HYPHEN_CODE Then
            action = "SkippedDrift"
            pageSkippedCount = pageSkippedCount + 1
        ElseIf mode = SH_DryRunOnly Then
            action = "DryRun"
            pageSkippedCount = pageSkippedCount + 1
        Else
            ' Show the find in the document.
            currentStep = "Pass 2 stray #" & (i + 1) & " - select+scroll"
            Selection.SetRange m.Start, m.End
            On Error Resume Next
            ActiveWindow.ScrollIntoView Selection.Range, True
            On Error GoTo PROC_ERR

            currentStep = "Pass 2 stray #" & (i + 1) & " - prompt"
            response = MsgBox( _
                "Soft hyphen find " & (i + 1) & " of " & pageStrayCount & _
                " on page " & pageNum & " (" & pageSide & ", " & strayCol(i) & " col)" & NL & _
                "X=" & Format(strayX(i), "0.0") & "  Y=" & Format(strayY(i), "0.0") & _
                "  YDelta=" & Format(strayYDelta(i), "0.0") & NL & NL & _
                "Context:" & NL & strayCtx(i) & NL & NL & _
                "Remove this soft hyphen?", _
                vbYesNoCancel + vbQuestion + vbDefaultButton1, _
                "SoftHyphenSweep p" & pageNum)

            If response = vbYes Then
                m.Delete
                cumDelOffset = cumDelOffset + 1
                action = "Removed"
                pageRemovedCount = pageRemovedCount + 1
            ElseIf response = vbNo Then
                action = "Skipped"
                pageSkippedCount = pageSkippedCount + 1
            Else
                action = "Cancelled"
                userCancelled = True
                Print #csvF, pageNum & "," & pageSide & ",stray" & (i + 1) & "," & _
                    pos & "," & Format(strayX(i), "0.0") & "," & _
                    Format(strayY(i), "0.0") & "," & Format(strayYDelta(i), "0.0") & "," & _
                    strayCol(i) & ",Stray," & action & ",""" & strayCtx(i) & """"
                Print #logF, "ABORTED at stray #" & (i + 1) & " of " & pageStrayCount
                Exit For
            End If
        End If

        Print #csvF, pageNum & "," & pageSide & ",stray" & (i + 1) & "," & _
            pos & "," & Format(strayX(i), "0.0") & "," & _
            Format(strayY(i), "0.0") & "," & Format(strayYDelta(i), "0.0") & "," & _
            strayCol(i) & ",Stray," & action & ",""" & strayCtx(i) & """"
    Next i

    Print #logF, "Page " & pageNum & " Result: " & findIdx & " find(s) - " & _
                 activeCount & " Active(Kept), " & pageStrayCount & " Stray (" & _
                 pageRemovedCount & " Removed, " & pageSkippedCount & " Skipped), " & _
                 outsideCount & " OutsideBody"

    Close #logF
    logF = 0
    Close #csvF
    csvF = 0

    ' Update driver-side accumulators.
    strayCum = strayCum + pageStrayCount
    removedCum = removedCum + pageRemovedCount

    Debug.Print "SoftHyphenSweep p" & pageNum & " (" & pageSide & "): " & _
                findIdx & " find(s) - " & activeCount & " Active, " & _
                pageStrayCount & " Stray (" & pageRemovedCount & " Removed, " & _
                pageSkippedCount & " Skipped), " & outsideCount & " OutsideBody" & _
                IIf(userCancelled, "  [CANCELLED]", "")

PROC_EXIT:
    If logF > 0 Then
        On Error Resume Next
        Close #logF
        On Error GoTo 0
    End If
    If csvF > 0 Then
        On Error Resume Next
        Close #csvF
        On Error GoTo 0
    End If
    Exit Sub
PROC_ERR:
    If logF > 0 Then
        On Error Resume Next
        Print #logF, "ABORTED at step [" & currentStep & "] - Err " & _
                     Err.Number & ": " & Err.Description
        Close #logF
        On Error GoTo 0
    End If
    If csvF > 0 Then
        On Error Resume Next
        Close #csvF
        On Error GoTo 0
    End If
    MsgBox "Step: [" & currentStep & "]" & vbCrLf & _
           "Error " & Err.Number & " (" & Err.Description & ")" & vbCrLf & _
           "in procedure SoftHyphenSweep_ByColumnContext_SinglePage of Module basWordRepairRunner"
    Resume PROC_EXIT
End Sub

'==============================================================================
' RunSoftHyphenSweep_Across_Pages_From
' PURPOSE:
'   Driver: invoke SoftHyphenSweep_ByColumnContext_SinglePage across a page
'   range. All three args required (no Optional - per design Q4, dryRun must
'   be passed explicitly to force an intentional choice).
'
' Usage:
'   RunSoftHyphenSweep_Across_Pages_From 911, 1, True    ' dry-run page 911
'   RunSoftHyphenSweep_Across_Pages_From 911, 5, False   ' live, pages 911-915
'==============================================================================
Public Sub RunSoftHyphenSweep_Across_Pages_From( _
        ByVal startPage As Long, _
        ByVal pageCount As Long, _
        ByVal dryRun As Boolean)
    On Error GoTo PROC_ERR

    If pageCount < 1 Then
        MsgBox "pageCount must be >= 1.", vbExclamation, "RunSoftHyphenSweep"
        Exit Sub
    End If

    Dim mode          As SoftHyphenMode
    Dim totalStray    As Long
    Dim totalRemoved  As Long
    Dim cancelled     As Boolean
    Dim p             As Long
    Dim endPage       As Long
    Const NL          As String = vbCrLf

    mode = IIf(dryRun, SH_DryRunOnly, SH_PromptEach)
    endPage = startPage + pageCount - 1

    Debug.Print "RunSoftHyphenSweep: pages " & startPage & ".." & endPage & _
                "  mode=" & IIf(dryRun, "DryRun", "PromptEach")

    For p = startPage To endPage
        If cancelled Then Exit For
        SoftHyphenSweep_ByColumnContext_SinglePage p, mode, _
            totalStray, totalRemoved, cancelled
    Next p

    Dim trailer As String
    If cancelled Then
        trailer = "  [CANCELLED at page " & p & "]"
    Else
        trailer = ""
    End If

    Debug.Print "RunSoftHyphenSweep: done - " & totalStray & " stray total, " & _
                totalRemoved & " removed" & trailer

    MsgBox "Sweep complete." & NL & _
           "Pages       : " & startPage & " .. " & endPage & NL & _
           "Mode        : " & IIf(dryRun, "DryRun (no removals)", "PromptEach (live)") & NL & _
           "Stray total : " & totalStray & NL & _
           "Removed     : " & totalRemoved & NL & _
           IIf(cancelled, "Status      : CANCELLED at page " & p, "Status      : Completed") & NL & NL & _
           "rpt\SoftHyphenSweep.csv  (per-find rows)" & NL & _
           "rpt\SoftHyphenSweep.log  (per-page summary)", _
           vbInformation, "RunSoftHyphenSweep"

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Error " & Err.Number & " (" & Err.Description & _
           ") in procedure RunSoftHyphenSweep_Across_Pages_From of Module basWordRepairRunner"
    Resume PROC_EXIT
End Sub

'==============================================================================
' RowCharCountSurvey_SinglePage
' PURPOSE:
'   Read-only survey: walk the body of one page paragraph by paragraph,
'   group characters into visual rows by Y position, and emit one CSV row
'   per visual row. Companion to SoftHyphenSweep, but for diagnosing
'   excessive inter-word spacing in justified two-column body text.
'
' Document assumption: one verse per paragraph (normalized).
'   - Single-row verses: their only row contains the paragraph mark, so
'     IsParagraphEnd=True and the histogram excludes them.
'   - Multi-row verses: only the final row contains the paragraph mark;
'     earlier rows are fully justified and ARE measured.
'
' OUTPUTS (append-mode):
'   rpt\RowCharCount.csv (one row per visual line)
'   rpt\RowCharCount.log (per-page summary)
'
' CSV columns:
'   PageNum,Side,RowIndex,Y,LeftX,RightX,CharCount,Pitch,
'   LastCharCode,EndsWithSoftHyphen,IsParagraphEnd,
'   RangeStart,RangeEnd,FirstChars
'
' Pitch = (RightX - LeftX) / max(CharCount-1, 1)  pt per char (pen advance)
' Side  = ClassifyColumnAt(LeftX) restricted to Left/Right; rows whose
'         LeftX falls outside the body columns are written with Side set
'         to the band name (OutsideLeft/Gutter/OutsideRight) so the
'         histogram pass can filter them.
'
' ARGS:
'   pageNum       - page to scan (1-based)
'   rowsCum       - byref: total rows emitted (incremented)
'   userCancelled - byref: True if user cancels mid-scan via Esc
'==============================================================================
Public Sub RowCharCountSurvey_SinglePage( _
        ByVal pageNum As Long, _
        ByRef rowsCum As Long, _
        ByRef userCancelled As Boolean)
    Dim currentStep   As String
    On Error GoTo PROC_ERR
    Dim oDoc          As Word.Document
    Dim pgRange       As Word.Range
    Dim ch            As Word.Range
    Dim sectionPS     As Word.PageSetup
    Dim para          As Word.Paragraph
    Dim paraStart     As Long, paraEnd As Long
    Dim pageStart     As Long, pageEnd As Long, pageNextStart As Long
    Dim isMirrored    As Boolean
    Dim pageSide      As String
    Dim csvPath       As String, logPath As String
    Dim csvF          As Integer, logF As Integer
    Dim writeCsvHeader As Boolean
    Const NL          As String = vbCrLf

    Dim bodyXMin As Single, bodyXMax As Single
    Dim leftColMax As Single, gutterMin As Single, gutterMax As Single, rightColMin As Single

    ' Per-row accumulators.
    Dim rowAnchorY    As Single
    Dim rowFirstX     As Single, rowLastX As Single
    Dim rowFirstPos   As Long, rowLastPos As Long
    Dim rowCharCount  As Long
    Dim rowFirstChars As String
    Dim lastCharCode  As Long
    Dim endsWithShy   As Boolean
    Dim rowOpen       As Boolean
    Dim rowIndex      As Long
    Dim pageRows      As Long, pageRowsBody As Long, pageRowsParaEnd As Long
    Dim pageRowsShy   As Long, pageRowsOutside As Long

    Dim p             As Long
    Dim chText        As String
    Dim asciiCode     As Long
    Dim curY          As Single, curX As Single

    currentStep = "Set ActiveDocument"
    Set oDoc = ActiveDocument

    currentStep = "GoTo page " & pageNum
    Set pgRange = oDoc.GoTo(What:=wdGoToPage, name:=CStr(pageNum))
    pageStart = pgRange.Start

    currentStep = "GoTo page " & (pageNum + 1)
    Set pgRange = oDoc.GoTo(What:=wdGoToPage, name:=CStr(pageNum + 1))
    pageNextStart = pgRange.Start
    If pageNextStart > pageStart Then
        pageEnd = pageNextStart - 1
        If pageEnd > oDoc.Content.End Then pageEnd = oDoc.Content.End
    Else
        pageEnd = oDoc.Content.End
    End If

    currentStep = "Resolve column bounds for page " & pageNum
    GetColumnBoundsForPage pageNum, bodyXMin, bodyXMax, _
                           leftColMax, gutterMin, gutterMax, rightColMin

    Set sectionPS = oDoc.Range(pageStart, pageStart + 1).Sections(1).PageSetup
    isMirrored = sectionPS.MirrorMargins
    If isMirrored Then
        pageSide = IIf(pageNum Mod 2 = 1, "Recto", "Verso")
    Else
        pageSide = "Single"
    End If

    currentStep = "Open report files"
    Dim docDir As String
    docDir = oDoc.Path
    If Len(docDir) = 0 Then docDir = Environ$("TEMP")
    csvPath = docDir & "\rpt\RowCharCount.csv"
    logPath = docDir & "\rpt\RowCharCount.log"

    writeCsvHeader = (Len(Dir(csvPath)) = 0)
    csvF = FreeFile
    Open csvPath For Append As #csvF
    If writeCsvHeader Then
        Print #csvF, "PageNum,PageSide,RowIndex,Side,Y,LeftX,RightX,CharCount,Pitch," & _
                     "LastCharCode,EndsWithSoftHyphen,IsParagraphEnd,RangeStart,RangeEnd,FirstChars"
    End If

    logF = FreeFile
    Open logPath For Append As #logF
    Print #logF, String(72, "=")
    Print #logF, "RowCharCountSurvey page=" & pageNum & "  side=" & pageSide & _
                 "  run=" & Format(Now, "yyyy-mm-dd hh:mm:ss")
    Print #logF, "Bounds   : Body=[" & Format(bodyXMin, "0.0") & ".." & _
                 Format(bodyXMax, "0.0") & "]  L=[" & _
                 Format(bodyXMin, "0.0") & ".." & Format(leftColMax, "0.0") & "]  G=[" & _
                 Format(gutterMin, "0.0") & ".." & Format(gutterMax, "0.0") & "]  R=[" & _
                 Format(rightColMin, "0.0") & ".." & Format(bodyXMax, "0.0") & "]"

    ' Walk paragraphs that overlap [pageStart, pageEnd]. Limit story to main
    ' text - this skips headers, footers, footnotes by construction.
    currentStep = "Locate first paragraph on page"
    Dim startPara As Word.Range
    Set startPara = oDoc.Range(pageStart, pageStart)
    Set startPara = startPara.Paragraphs(1).Range

    rowOpen = False
    rowIndex = 0
    Application.StatusBar = "RowCharCountSurvey p" & pageNum & " - scanning..."

    For Each para In oDoc.Range(startPara.Start, pageEnd).Paragraphs
        If userCancelled Then Exit For
        paraStart = para.Range.Start
        paraEnd = para.Range.End
        If paraStart >= pageEnd Then Exit For
        If para.Range.StoryType <> wdMainTextStory Then GoTo NextPara

        currentStep = "Walk para " & paraStart & ".." & paraEnd
        For p = paraStart To paraEnd - 1
            If p < pageStart Then GoTo NextChar
            If p >= pageEnd Then Exit For

            Set ch = oDoc.Range(p, p + 1)
            chText = ch.Text
            If Len(chText) = 0 Then GoTo NextChar
            asciiCode = AscW(chText)

            ' Skip vertical-tab line-break characters within row tracking;
            ' treat them like a forced row break (rare in this doc).
            curY = CSng(ch.Information(wdVerticalPositionRelativeToPage))
            curX = CSng(ch.Information(wdHorizontalPositionRelativeToPage))

            If Not rowOpen Then
                rowAnchorY = curY
                rowFirstX = curX
                rowLastX = curX
                rowFirstPos = p
                rowLastPos = p
                rowCharCount = 0
                rowFirstChars = ""
                endsWithShy = False
                lastCharCode = 0
                rowOpen = True
            ElseIf Abs(curY - rowAnchorY) > LINE_HEIGHT_TOLERANCE Then
                ' Y jumped: flush the row that just ended (not a paragraph end).
                FlushRowCharCountRow csvF, pageNum, pageSide, rowIndex, _
                    bodyXMin, leftColMax, gutterMax, bodyXMax, _
                    rowAnchorY, rowFirstX, rowLastX, rowCharCount, _
                    lastCharCode, endsWithShy, False, _
                    rowFirstPos, rowLastPos, rowFirstChars, _
                    pageRowsBody, pageRowsShy, pageRowsOutside
                pageRows = pageRows + 1
                rowIndex = rowIndex + 1
                rowsCum = rowsCum + 1

                rowAnchorY = curY
                rowFirstX = curX
                rowLastX = curX
                rowFirstPos = p
                rowLastPos = p
                rowCharCount = 0
                rowFirstChars = ""
                endsWithShy = False
                lastCharCode = 0
            End If

            rowLastX = curX
            rowLastPos = p
            rowCharCount = rowCharCount + 1
            lastCharCode = asciiCode
            endsWithShy = (asciiCode = SOFT_HYPHEN_CODE)

            If Len(rowFirstChars) < 30 Then
                If asciiCode >= 32 And asciiCode < 127 Then
                    rowFirstChars = rowFirstChars & chText
                ElseIf asciiCode = SOFT_HYPHEN_CODE Then
                    rowFirstChars = rowFirstChars & "-"
                End If
            End If

            ' Paragraph mark = last row of this paragraph. Flush with IsParagraphEnd=True.
            If asciiCode = 13 Then
                FlushRowCharCountRow csvF, pageNum, pageSide, rowIndex, _
                    bodyXMin, leftColMax, gutterMax, bodyXMax, _
                    rowAnchorY, rowFirstX, rowLastX, rowCharCount, _
                    lastCharCode, endsWithShy, True, _
                    rowFirstPos, rowLastPos, rowFirstChars, _
                    pageRowsBody, pageRowsShy, pageRowsOutside
                pageRows = pageRows + 1
                pageRowsParaEnd = pageRowsParaEnd + 1
                rowIndex = rowIndex + 1
                rowsCum = rowsCum + 1
                rowOpen = False
            End If
NextChar:
            If (p Mod 200) = 0 Then
                DoEvents
                If userCancelled Then Exit For
            End If
        Next p
NextPara:
    Next para

    ' Flush trailing row if the page ended mid-paragraph (no Chr(13) seen).
    If rowOpen Then
        FlushRowCharCountRow csvF, pageNum, pageSide, rowIndex, _
            bodyXMin, leftColMax, gutterMax, bodyXMax, _
            rowAnchorY, rowFirstX, rowLastX, rowCharCount, _
            lastCharCode, endsWithShy, False, _
            rowFirstPos, rowLastPos, rowFirstChars, _
            pageRowsBody, pageRowsShy, pageRowsOutside
        pageRows = pageRows + 1
        rowIndex = rowIndex + 1
        rowsCum = rowsCum + 1
        rowOpen = False
    End If

    Print #logF, "Page " & pageNum & " Result: " & pageRows & " row(s) - " & _
                 pageRowsBody & " body (Left/Right), " & _
                 pageRowsOutside & " outside-body, " & _
                 pageRowsParaEnd & " paragraph-end (excluded), " & _
                 pageRowsShy & " end-with-soft-hyphen (excluded)"

    Close #logF
    logF = 0
    Close #csvF
    csvF = 0

    Application.StatusBar = False
    Debug.Print "RowCharCountSurvey p" & pageNum & " (" & pageSide & "): " & _
                pageRows & " row(s) - body=" & pageRowsBody & _
                " outside=" & pageRowsOutside & _
                " paraEnd=" & pageRowsParaEnd & _
                " endShy=" & pageRowsShy

PROC_EXIT:
    Application.StatusBar = False
    If logF > 0 Then
        On Error Resume Next
        Close #logF
        On Error GoTo 0
    End If
    If csvF > 0 Then
        On Error Resume Next
        Close #csvF
        On Error GoTo 0
    End If
    Exit Sub
PROC_ERR:
    If logF > 0 Then
        On Error Resume Next
        Print #logF, "ABORTED at step [" & currentStep & "] - Err " & _
                     Err.Number & ": " & Err.Description
        Close #logF
        On Error GoTo 0
    End If
    If csvF > 0 Then
        On Error Resume Next
        Close #csvF
        On Error GoTo 0
    End If
    MsgBox "Step: [" & currentStep & "]" & vbCrLf & _
           "Error " & Err.Number & " (" & Err.Description & ")" & vbCrLf & _
           "in procedure RowCharCountSurvey_SinglePage of Module basWordRepairRunner"
    Resume PROC_EXIT
End Sub

'------------------------------------------------------------------------------
' FlushRowCharCountRow - emit one CSV record for a completed row and update
' per-page counters. Called only by RowCharCountSurvey_SinglePage.
'------------------------------------------------------------------------------
Private Sub FlushRowCharCountRow( _
        ByVal csvF As Integer, _
        ByVal pageNum As Long, _
        ByVal pageSide As String, _
        ByVal rowIndex As Long, _
        ByVal bodyXMin As Single, _
        ByVal leftColMax As Single, _
        ByVal gutterMax As Single, _
        ByVal bodyXMax As Single, _
        ByVal anchorY As Single, _
        ByVal leftX As Single, _
        ByVal rightX As Single, _
        ByVal charCount As Long, _
        ByVal lastCharCode As Long, _
        ByVal endsWithShy As Boolean, _
        ByVal isParaEnd As Boolean, _
        ByVal rangeStart As Long, _
        ByVal rangeEnd As Long, _
        ByVal firstChars As String, _
        ByRef cntBody As Long, _
        ByRef cntShy As Long, _
        ByRef cntOutside As Long)

    Dim side As String
    side = ClassifyColumnAt(leftX, bodyXMin, leftColMax, gutterMax, bodyXMax)

    Dim pitch As Single
    If charCount > 1 Then
        pitch = (rightX - leftX) / (charCount - 1)
    Else
        pitch = 0
    End If

    Dim cleanFirst As String
    cleanFirst = Replace(firstChars, """", """""")
    cleanFirst = Replace(cleanFirst, Chr(13), " ")
    cleanFirst = Replace(cleanFirst, Chr(11), " ")
    cleanFirst = Replace(cleanFirst, Chr(10), " ")
    cleanFirst = Replace(cleanFirst, Chr(9), " ")

    Print #csvF, pageNum & "," & pageSide & "," & rowIndex & "," & _
                 side & "," & Format(anchorY, "0.0") & "," & _
                 Format(leftX, "0.0") & "," & Format(rightX, "0.0") & "," & _
                 charCount & "," & Format(pitch, "0.000") & "," & _
                 lastCharCode & "," & IIf(endsWithShy, "True", "False") & "," & _
                 IIf(isParaEnd, "True", "False") & "," & _
                 rangeStart & "," & rangeEnd & "," & _
                 """" & cleanFirst & """"

    If side = "Left" Or side = "Right" Then
        cntBody = cntBody + 1
    Else
        cntOutside = cntOutside + 1
    End If
    If endsWithShy Then cntShy = cntShy + 1
End Sub

'==============================================================================
' RunRowCharCountSurvey_Across_Pages_From
' PURPOSE:
'   Driver: invoke RowCharCountSurvey_SinglePage across a page range. Mirrors
'   RunSoftHyphenSweep_Across_Pages_From's shape. Read-only - never edits the
'   document. Output appends to rpt\RowCharCount.csv and rpt\RowCharCount.log.
'
' Usage:
'   RunRowCharCountSurvey_Across_Pages_From 100, 10
'   RunRowCharCountSurvey_Across_Pages_From 250, 10
'==============================================================================
Public Sub RunRowCharCountSurvey_Across_Pages_From( _
        ByVal startPage As Long, _
        ByVal pageCount As Long)
    On Error GoTo PROC_ERR

    If pageCount < 1 Then
        MsgBox "pageCount must be >= 1.", vbExclamation, "RunRowCharCountSurvey"
        Exit Sub
    End If

    Dim totalRows As Long
    Dim cancelled As Boolean
    Dim p         As Long
    Dim endPage   As Long
    Const NL      As String = vbCrLf

    endPage = startPage + pageCount - 1

    Debug.Print "RunRowCharCountSurvey: pages " & startPage & ".." & endPage

    For p = startPage To endPage
        If cancelled Then Exit For
        RowCharCountSurvey_SinglePage p, totalRows, cancelled
    Next p

    Dim trailer As String
    If cancelled Then
        trailer = "  [CANCELLED at page " & p & "]"
    Else
        trailer = ""
    End If

    Debug.Print "RunRowCharCountSurvey: done - " & totalRows & " row(s) total" & trailer

    MsgBox "Survey complete." & NL & _
           "Pages       : " & startPage & " .. " & endPage & NL & _
           "Rows total  : " & totalRows & NL & _
           IIf(cancelled, "Status      : CANCELLED at page " & p, "Status      : Completed") & NL & NL & _
           "rpt\RowCharCount.csv  (per-row records)" & NL & _
           "rpt\RowCharCount.log  (per-page summary)", _
           vbInformation, "RunRowCharCountSurvey"

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Error " & Err.Number & " (" & Err.Description & _
           ") in procedure RunRowCharCountSurvey_Across_Pages_From of Module basWordRepairRunner"
    Resume PROC_EXIT
End Sub

'==============================================================================
' BuildRowCharCountHistogram
' PURPOSE:
'   Phase C of the row-char-Count diagnostic. Reads rpt\RowCharCount.csv,
'   filters out rows that should not be measured (paragraph-end rows,
'   soft-hyphen-terminated rows, non-body rows), buckets the remainder by
'   CharCount and Pitch per Side, computes median pitch per side, and writes:
'
'     rpt\RowCharCountHistogram.csv  (Side, Metric, Bin, Frequency)
'     rpt\RowCharCountSuspects.csv   (rows with Pitch > median + threshold)
'     rpt\RowCharCount.log           (appended summary block)
'
' ARGS:
'   thresholdPt - Optional pitch excess (pt) above per-side median that
'                 marks a row as a suspect. Default 1.0 pt. Tune from the
'                 histogram shape.
'
' Usage:
'   BuildRowCharCountHistogram                ' default threshold 1.0 pt
'   BuildRowCharCountHistogram 0.8            ' tighter
'   BuildRowCharCountHistogram 1.5            ' looser
'==============================================================================
Public Sub BuildRowCharCountHistogram(Optional ByVal thresholdPt As Single = 1#)
    Dim currentStep As String
    On Error GoTo PROC_ERR

    Dim oDoc As Word.Document
    Dim docDir As String
    Dim inPath As String, histPath As String, suspPath As String, logPath As String
    Dim inF As Integer, histF As Integer, suspF As Integer, logF As Integer
    Const NL As String = vbCrLf

    currentStep = "Resolve paths"
    Set oDoc = ActiveDocument
    docDir = oDoc.Path
    If Len(docDir) = 0 Then docDir = Environ$("TEMP")
    inPath = docDir & "\rpt\RowCharCount.csv"
    histPath = docDir & "\rpt\RowCharCountHistogram.csv"
    suspPath = docDir & "\rpt\RowCharCountSuspects.csv"
    logPath = docDir & "\rpt\RowCharCount.log"

    If Len(Dir(inPath)) = 0 Then
        MsgBox "Input not found:" & NL & inPath & NL & NL & _
               "Run RunRowCharCountSurvey_Across_Pages_From first.", _
               vbExclamation, "BuildRowCharCountHistogram"
        Exit Sub
    End If

    ' Per-side accumulators. Bins:
    '   CharCount   : 0..199 (chars)
    '   Pitch       : 0..199 (in 0.1 pt buckets => 0.0 to 19.9 pt)
    Dim histLeftCC(0 To 199)  As Long
    Dim histRightCC(0 To 199) As Long
    Dim histLeftPB(0 To 199)  As Long
    Dim histRightPB(0 To 199) As Long

    ' Pitch values for median, kept per side (resize as needed).
    Dim pitchLeft()  As Single, nLeft As Long
    Dim pitchRight() As Single, nRight As Long
    ReDim pitchLeft(0 To 1023)
    ReDim pitchRight(0 To 1023)

    ' Parsed eligible rows held for the suspect pass (so we can compare each
    ' row's pitch against the per-side median once it is known).
    Dim rowSide()    As String
    Dim rowPitch()   As Single
    Dim rowLine()    As String
    Dim nRows        As Long
    ReDim rowSide(0 To 4095)
    ReDim rowPitch(0 To 4095)
    ReDim rowLine(0 To 4095)

    Dim totalScanned As Long, totalEligible As Long
    Dim skipParaEnd As Long, skipShy As Long, skipOutside As Long

    currentStep = "Open input " & inPath
    inF = FreeFile
    Open inPath For Input As #inF

    Dim line As String, parts() As String
    Dim sideStr As String, charCount As Long, pitch As Single
    Dim pitchBin As Long, ccBin As Long
    Dim isFirst As Boolean: isFirst = True

    Do While Not EOF(inF)
        Line Input #inF, line
        If isFirst Then
            isFirst = False  ' skip header
        Else
            If Len(line) = 0 Then GoTo NextLine
            totalScanned = totalScanned + 1

            ' parts indices (see RowCharCountSurvey_SinglePage CSV header):
            '  0=PageNum  1=PageSide  2=RowIndex  3=Side  4=Y
            '  5=LeftX  6=RightX  7=CharCount  8=Pitch
            '  9=LastCharCode  10=EndsWithSoftHyphen  11=IsParagraphEnd
            ' 12=RangeStart  13=RangeEnd  14+=FirstChars (may contain commas)
            parts = Split(line, ",")
            If UBound(parts) < 13 Then GoTo NextLine

            sideStr = parts(3)
            If sideStr <> "Left" And sideStr <> "Right" Then
                skipOutside = skipOutside + 1
                GoTo NextLine
            End If
            If parts(11) = "True" Then
                skipParaEnd = skipParaEnd + 1
                GoTo NextLine
            End If
            If parts(10) = "True" Then
                skipShy = skipShy + 1
                GoTo NextLine
            End If

            charCount = CLng(parts(7))
            pitch = CSng(parts(8))
            ccBin = charCount
            If ccBin < 0 Then ccBin = 0
            If ccBin > 199 Then ccBin = 199
            pitchBin = CLng(Int(pitch * 10#))
            If pitchBin < 0 Then pitchBin = 0
            If pitchBin > 199 Then pitchBin = 199

            If sideStr = "Left" Then
                histLeftCC(ccBin) = histLeftCC(ccBin) + 1
                histLeftPB(pitchBin) = histLeftPB(pitchBin) + 1
                If nLeft > UBound(pitchLeft) Then ReDim Preserve pitchLeft(0 To UBound(pitchLeft) + 1024)
                pitchLeft(nLeft) = pitch: nLeft = nLeft + 1
            Else
                histRightCC(ccBin) = histRightCC(ccBin) + 1
                histRightPB(pitchBin) = histRightPB(pitchBin) + 1
                If nRight > UBound(pitchRight) Then ReDim Preserve pitchRight(0 To UBound(pitchRight) + 1024)
                pitchRight(nRight) = pitch: nRight = nRight + 1
            End If

            If nRows > UBound(rowSide) Then
                ReDim Preserve rowSide(0 To UBound(rowSide) + 4096)
                ReDim Preserve rowPitch(0 To UBound(rowPitch) + 4096)
                ReDim Preserve rowLine(0 To UBound(rowLine) + 4096)
            End If
            rowSide(nRows) = sideStr
            rowPitch(nRows) = pitch
            rowLine(nRows) = line
            nRows = nRows + 1

            totalEligible = totalEligible + 1
        End If
NextLine:
    Loop
    Close #inF
    inF = 0

    ' --- Medians per side.
    currentStep = "Compute medians"
    Dim medianLeft As Single, medianRight As Single
    medianLeft = MedianOfSingles(pitchLeft, nLeft)
    medianRight = MedianOfSingles(pitchRight, nRight)

    Dim modeLeftCC As Long, modeRightCC As Long
    Dim modeLeftCount As Long, modeRightCount As Long
    Dim i As Long
    For i = 0 To 199
        If histLeftCC(i) > modeLeftCount Then modeLeftCount = histLeftCC(i): modeLeftCC = i
        If histRightCC(i) > modeRightCount Then modeRightCount = histRightCC(i): modeRightCC = i
    Next i

    ' --- Write histogram CSV (overwrite).
    currentStep = "Write histogram"
    histF = FreeFile
    Open histPath For Output As #histF
    Print #histF, "Side,Metric,Bin,Frequency"
    For i = 0 To 199
        If histLeftCC(i) > 0 Then Print #histF, "Left,CharCount," & i & "," & histLeftCC(i)
    Next i
    For i = 0 To 199
        If histRightCC(i) > 0 Then Print #histF, "Right,CharCount," & i & "," & histRightCC(i)
    Next i
    For i = 0 To 199
        If histLeftPB(i) > 0 Then _
            Print #histF, "Left,Pitch," & Format(i / 10#, "0.0") & "," & histLeftPB(i)
    Next i
    For i = 0 To 199
        If histRightPB(i) > 0 Then _
            Print #histF, "Right,Pitch," & Format(i / 10#, "0.0") & "," & histRightPB(i)
    Next i
    Close #histF
    histF = 0

    ' --- Write suspects CSV (overwrite).
    currentStep = "Write suspects"
    suspF = FreeFile
    Open suspPath For Output As #suspF
    Print #suspF, "PageNum,PageSide,RowIndex,Side,Y,LeftX,RightX,CharCount,Pitch," & _
                  "LastCharCode,EndsWithSoftHyphen,IsParagraphEnd,RangeStart,RangeEnd," & _
                  "FirstChars,MedianPitchSide,PitchExcess"

    Dim threshLeft As Single, threshRight As Single
    threshLeft = medianLeft + thresholdPt
    threshRight = medianRight + thresholdPt

    Dim suspectCount As Long
    Dim med As Single, thresh As Single
    For i = 0 To nRows - 1
        If rowSide(i) = "Left" Then
            med = medianLeft: thresh = threshLeft
        Else
            med = medianRight: thresh = threshRight
        End If
        If rowPitch(i) > thresh Then
            Print #suspF, rowLine(i) & "," & Format(med, "0.000") & "," & _
                          Format(rowPitch(i) - med, "0.000")
            suspectCount = suspectCount + 1
        End If
    Next i
    Close #suspF
    suspF = 0

    ' --- Append summary to the survey log so the run history stays in one file.
    currentStep = "Append log summary"
    logF = FreeFile
    Open logPath For Append As #logF
    Print #logF, String(72, "-")
    Print #logF, "BuildRowCharCountHistogram run=" & Format(Now, "yyyy-mm-dd hh:mm:ss") & _
                 "  threshold=+" & Format(thresholdPt, "0.0") & " pt over median"
    Print #logF, "Scanned    : " & totalScanned & " row(s) from " & inPath
    Print #logF, "Excluded   : paraEnd=" & skipParaEnd & "  endShy=" & skipShy & _
                 "  outside-body=" & skipOutside
    Print #logF, "Eligible   : " & totalEligible & "  (Left=" & nLeft & "  Right=" & nRight & ")"
    Print #logF, "Mode CC    : Left=" & modeLeftCC & " (n=" & modeLeftCount & ")  " & _
                 "Right=" & modeRightCC & " (n=" & modeRightCount & ")"
    Print #logF, "Median Pt  : Left=" & Format(medianLeft, "0.000") & _
                 "  Right=" & Format(medianRight, "0.000")
    Print #logF, "Threshold  : Left>" & Format(threshLeft, "0.000") & _
                 "  Right>" & Format(threshRight, "0.000")
    Print #logF, "Suspects   : " & suspectCount
    Print #logF, "Outputs    : " & histPath
    Print #logF, "             " & suspPath
    Close #logF
    logF = 0

    Debug.Print "BuildRowCharCountHistogram: scanned=" & totalScanned & _
                " eligible=" & totalEligible & " suspects=" & suspectCount & _
                " medianL=" & Format(medianLeft, "0.000") & _
                " medianR=" & Format(medianRight, "0.000")

    MsgBox "Histogram built." & NL & _
           "Scanned     : " & totalScanned & NL & _
           "Excluded    : paraEnd=" & skipParaEnd & "  endShy=" & skipShy & _
                          "  outside=" & skipOutside & NL & _
           "Eligible    : " & totalEligible & NL & _
           "Mode CC     : L=" & modeLeftCC & "  R=" & modeRightCC & NL & _
           "Median pitch: L=" & Format(medianLeft, "0.000") & _
                          "  R=" & Format(medianRight, "0.000") & " pt/char" & NL & _
           "Threshold   : median + " & Format(thresholdPt, "0.0") & " pt" & NL & _
           "Suspects    : " & suspectCount & NL & NL & _
           "rpt\RowCharCountHistogram.csv" & NL & _
           "rpt\RowCharCountSuspects.csv" & NL & _
           "rpt\RowCharCount.log  (summary appended)", _
           vbInformation, "BuildRowCharCountHistogram"

PROC_EXIT:
    If inF > 0 Then On Error Resume Next: Close #inF: On Error GoTo 0
    If histF > 0 Then On Error Resume Next: Close #histF: On Error GoTo 0
    If suspF > 0 Then On Error Resume Next: Close #suspF: On Error GoTo 0
    If logF > 0 Then On Error Resume Next: Close #logF: On Error GoTo 0
    Exit Sub
PROC_ERR:
    MsgBox "Step: [" & currentStep & "]" & vbCrLf & _
           "Error " & Err.Number & " (" & Err.Description & ")" & vbCrLf & _
           "in procedure BuildRowCharCountHistogram of Module basWordRepairRunner"
    Resume PROC_EXIT
End Sub

'==============================================================================
' ReviewRowCharCountSuspects
' PURPOSE:
'   Phase B navigator. Each invocation jumps to the next suspect row in
'   rpt\RowCharCountSuspects.csv: selects the row in the document and
'   scrolls it into view, then dismisses with a status MsgBox that names
'   the suspect, its pitch, and how to advance.
'
'   The MsgBox is modal (Word edits are blocked while it is open), so the
'   pattern is: dismiss -> the row remains selected -> add a soft hyphen
'   (Ctrl+Hyphen) where appropriate -> re-invoke this macro for the next
'   suspect. Bind to a keyboard shortcut for fastest cycling.
'
'   No decision logging by design: the survey itself is the ledger -
'   re-running RunRowCharCountSurvey + BuildRowCharCountHistogram on the
'   same range shows fewer suspects (rows that received a soft hyphen
'   are now end-shy and excluded). For an explicit "review later"
'   marker, just leave the suspect untouched and re-run the histogram.
'
'   State persists in module-private variables for the VBA session only.
'   Use ReviewRowCharCountSuspects_Reset to start over without restarting
'   Word, or after re-running BuildRowCharCountHistogram.
'==============================================================================
Public Sub ReviewRowCharCountSuspects()
    Dim currentStep As String
    On Error GoTo PROC_ERR

    Dim oDoc      As Word.Document
    Dim docDir    As String, suspPath As String
    Dim f         As Integer
    Dim line      As String
    Dim isFirst   As Boolean
    Const NL      As String = vbCrLf

    currentStep = "Set ActiveDocument"
    Set oDoc = ActiveDocument
    docDir = oDoc.Path
    If Len(docDir) = 0 Then docDir = Environ$("TEMP")
    suspPath = docDir & "\rpt\RowCharCountSuspects.csv"

    If Not mRevLoaded Then
        currentStep = "Load suspects CSV"
        If Len(Dir(suspPath)) = 0 Then
            MsgBox "Suspects CSV not found:" & NL & suspPath & NL & NL & _
                   "Run BuildRowCharCountHistogram first.", _
                   vbExclamation, "ReviewRowCharCountSuspects"
            Exit Sub
        End If

        ReDim mRevSuspects(0 To 1023)
        mRevTotal = 0
        f = FreeFile
        Open suspPath For Input As #f
        isFirst = True
        Do While Not EOF(f)
            Line Input #f, line
            If isFirst Then
                isFirst = False
            Else
                If Len(line) > 0 Then
                    If mRevTotal > UBound(mRevSuspects) Then
                        ReDim Preserve mRevSuspects(0 To UBound(mRevSuspects) + 1024)
                    End If
                    mRevSuspects(mRevTotal) = line
                    mRevTotal = mRevTotal + 1
                End If
            End If
        Loop
        Close #f

        If mRevTotal = 0 Then
            MsgBox "No suspects in:" & NL & suspPath & NL & NL & _
                   "Build was clean (suspects=0). Lower the threshold" & NL & _
                   "or expand the survey range to find candidates.", _
                   vbInformation, "ReviewRowCharCountSuspects"
            Exit Sub
        End If

        mRevIdx = 0
        mRevLoaded = True
        MsgBox "Loaded " & mRevTotal & " suspect(s) from:" & NL & suspPath & NL & NL & _
               "Each invocation selects the next suspect's row" & NL & _
               "and scrolls it into view. Dismiss the prompt and" & NL & _
               "edit the row directly; re-run for the next suspect.", _
               vbInformation, "ReviewRowCharCountSuspects"
    End If

    If mRevIdx >= mRevTotal Then
        MsgBox "Review complete: " & mRevTotal & " suspect(s) traversed." & NL & NL & _
               "Run ReviewRowCharCountSuspects_Reset to restart at suspect 1," & NL & _
               "or re-run RunRowCharCountSurvey + BuildRowCharCountHistogram" & NL & _
               "to refresh the suspects list against the current document.", _
               vbInformation, "ReviewRowCharCountSuspects"
        Exit Sub
    End If

    ' Parse current suspect line. Columns from BuildRowCharCountHistogram:
    '   0=PageNum  1=PageSide  2=RowIndex  3=Side  4=Y
    '   5=LeftX  6=RightX  7=CharCount  8=Pitch
    '   9=LastCharCode  10=EndsWithSoftHyphen  11=IsParagraphEnd
    '  12=RangeStart  13=RangeEnd
    '  14+=FirstChars (may contain embedded commas inside quotes)
    '  trailing 2 fields = MedianPitchSide, PitchExcess
    currentStep = "Parse suspect at index " & mRevIdx
    Dim parts() As String
    parts = Split(mRevSuspects(mRevIdx), ",")
    If UBound(parts) < 14 Then
        mRevIdx = mRevIdx + 1
        MsgBox "Skipped malformed suspect at line " & mRevIdx & "." & NL & _
               "Re-run for the next suspect.", _
               vbExclamation, "ReviewRowCharCountSuspects"
        Exit Sub
    End If

    Dim pageNum As Long, pageSideStr As String, sideStr As String
    Dim charCount As Long, pitch As Single
    Dim rangeStart As Long, rangeEnd As Long
    Dim pitchExcess As String, medianPitch As String
    pageNum = CLng(parts(0))
    pageSideStr = parts(1)
    sideStr = parts(3)
    charCount = CLng(parts(7))
    pitch = CSng(parts(8))
    rangeStart = CLng(parts(12))
    rangeEnd = CLng(parts(13))
    pitchExcess = parts(UBound(parts))
    medianPitch = parts(UBound(parts) - 1)

    currentStep = "Select range " & rangeStart & ".." & rangeEnd
    Selection.SetRange rangeStart, rangeEnd
    On Error Resume Next
    ActiveWindow.ScrollIntoView Selection.Range, True
    On Error GoTo PROC_ERR

    MsgBox "Suspect " & (mRevIdx + 1) & " of " & mRevTotal & NL & _
           "Page " & pageNum & " (" & pageSideStr & "), " & sideStr & " column" & NL & _
           "CharCount = " & charCount & "    Pitch = " & Format(pitch, "0.000") & " pt/char" & NL & _
           "Median(side) = " & medianPitch & "    Excess = +" & pitchExcess & " pt" & NL & NL & _
           "The row is selected and scrolled into view." & NL & _
           "Dismiss this dialog, then add a soft hyphen" & NL & _
           "(Ctrl+Hyphen) where you want the line to break," & NL & _
           "or leave the row as-is if no break works." & NL & NL & _
           "Re-run ReviewRowCharCountSuspects for next.", _
           vbInformation, _
           "Suspect " & (mRevIdx + 1) & "/" & mRevTotal & "  p" & pageNum

    mRevIdx = mRevIdx + 1
    Exit Sub

PROC_ERR:
    MsgBox "Step: [" & currentStep & "]" & vbCrLf & _
           "Error " & Err.Number & " (" & Err.Description & ")" & vbCrLf & _
           "in procedure ReviewRowCharCountSuspects of Module basWordRepairRunner"
End Sub

'------------------------------------------------------------------------------
' ReviewRowCharCountSuspects_Reset - clear the navigator's session state so the
' next ReviewRowCharCountSuspects call reloads the suspects CSV from disk and
' restarts at suspect 1. Use after re-running BuildRowCharCountHistogram.
'------------------------------------------------------------------------------
Public Sub ReviewRowCharCountSuspects_Reset()
    mRevLoaded = False
    mRevIdx = 0
    mRevTotal = 0
    Erase mRevSuspects
    MsgBox "Review state cleared." & vbCrLf & _
           "Next ReviewRowCharCountSuspects will reload" & vbCrLf & _
           "rpt\RowCharCountSuspects.csv from disk.", _
           vbInformation, "ReviewRowCharCountSuspects_Reset"
End Sub

'------------------------------------------------------------------------------
' MedianOfSingles - return the median of the first n entries of arr. Sorts a
' copy in place; returns 0 if n=0. Simple insertion sort (n is at most a few
' thousand for this diagnostic).
'------------------------------------------------------------------------------
Private Function MedianOfSingles(ByRef arr() As Single, ByVal n As Long) As Single
    If n <= 0 Then Exit Function
    Dim copy() As Single
    ReDim copy(0 To n - 1)
    Dim i As Long, j As Long
    Dim key As Single
    For i = 0 To n - 1
        copy(i) = arr(i)
    Next i
    For i = 1 To n - 1
        key = copy(i)
        j = i - 1
        Do While j >= 0
            If copy(j) <= key Then Exit Do
            copy(j + 1) = copy(j)
            j = j - 1
        Loop
        copy(j + 1) = key
    Next i
    If (n Mod 2) = 1 Then
        MedianOfSingles = copy(n \ 2)
    Else
        MedianOfSingles = (copy(n \ 2 - 1) + copy(n \ 2)) / 2#
    End If
End Function

