Attribute VB_Name = "basWordRepairRunner"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString
Private OneVersePerParaRepair As Boolean

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
    Dim pgRange As Word.Range, ch As Word.Range, scanRange As Word.Range, prefixCh As Word.Range
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
' at runtime so the result is correct under mirrored margins (where odd pages
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
    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, Name:=CStr(pageNum))
    Set ps = pgRange.Sections(1).PageSetup

    insideMargin = ps.LeftMargin + ps.Gutter
    outsideMargin = ps.RightMargin

    If ps.MirrorMargins Then
        insideOnLeft = ((pageNum Mod 2) = 1)   ' odd = recto = inside on left
    Else
        insideOnLeft = True
    End If

    If insideOnLeft Then
        bodyXMin = insideMargin
        bodyXMax = ps.PageWidth - outsideMargin
    Else
        bodyXMin = outsideMargin
        bodyXMax = ps.PageWidth - insideMargin
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
        Print #f, "  PageSize  : " & Format(ps.PageWidth, "0.0") & " x " & _
                  Format(ps.PageHeight, "0.0") & " pt"
        Print #f, "  Margins   : T=" & Format(ps.TopMargin, "0.0") & _
                  "  B=" & Format(ps.BottomMargin, "0.0") & _
                  "  L=" & Format(ps.LeftMargin, "0.0") & _
                  "  R=" & Format(ps.RightMargin, "0.0")
        Print #f, "  Gutter    : " & Format(ps.Gutter, "0.0") & _
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
    Set pgRange = oDoc.GoTo(What:=wdGoToPage, Name:=CStr(pageNum))
    pageStart = pgRange.Start

    currentStep = "GoTo page " & (pageNum + 1)
    Set pgRange = oDoc.GoTo(What:=wdGoToPage, Name:=CStr(pageNum + 1))
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

