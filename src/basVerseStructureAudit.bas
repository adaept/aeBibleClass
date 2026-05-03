Attribute VB_Name = "basVerseStructureAudit"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

' ==========================================================================
' AuditVerseMarkerStructure
' ==========================================================================
' Walk the Bible body and verify chapter / verse counts against
' aeBibleCitationClass canonical data. Read-only; produces a report file
' plus an Immediate-window summary.
'
' Invariants verified (core three; advanced 4-6 deferred):
'   1. Books present     - every canonical 66-book has a Heading 1.
'   2. Chapter counts    - per-book Heading 2 Count matches ChaptersInBook.
'   3. Verse counts      - per-chapter Verse marker Count matches VersesInChapter.
'
' Output: rpt\VerseStructureAudit.txt (when bWriteFile = True) plus
' Immediate-window summary.
'
' Usage:
'   AuditVerseMarkerStructure              ' default writes file
'   AuditVerseMarkerStructure False        ' Immediate only, no file
' ==========================================================================
Public Sub AuditVerseMarkerStructure(Optional ByVal bWriteFile As Boolean = True)
    Dim t As Double
    StartTimer "AuditVerseMarkerStructure", t

    Dim oDoc As Object
    Set oDoc = ActiveDocument

    ' Canonical 66-book reference data sourced directly from the citation
    ' class (single source of truth). books(BookID) = Array(BookID, name, chapters).
    Dim books As Object
    Set books = aeBibleCitationClass.GetCanonicalBookTable

    ' Walk Heading 1 paragraphs in document order
    Dim h1Names(1 To 200) As String
    Dim h1Starts(1 To 200) As Long
    Dim nH1 As Long
    nH1 = 0

    Dim oPara As Object
    For Each oPara In oDoc.Paragraphs
        If oPara.style.NameLocal = "Heading 1" Then
            nH1 = nH1 + 1
            If nH1 > 200 Then Exit For
            h1Names(nH1) = Trim(Replace(oPara.Range.Text, vbCr, ""))
            h1Starts(nH1) = oPara.Range.Start
        End If
    Next oPara

    ' Build report
    Dim sOut As String
    Const NL As String = vbCrLf
    Dim totalExpected As Long, totalFound As Long
    Dim issuesCount As Long
    Dim issues As String
    Dim seenBookID(1 To 66) As Boolean

    sOut = "---- AuditVerseMarkerStructure: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ----" & NL & NL
    sOut = sOut & nH1 & " Heading 1 paragraphs found in " & oDoc.Sections.Count & " sections." & NL & NL

    Dim i As Long
    Dim bookEndPos As Long
    Dim docEnd As Long
    Dim foundChapters As Long
    Dim chapterReport As String
    Dim bookIssues As Long
    Dim bookIssueDetail As String
    docEnd = oDoc.Content.End

    For i = 1 To nH1
        ' Reset per-book accumulators (VBA Dim-in-loop is procedure-scoped,
        ' so these would otherwise leak across iterations).
        chapterReport = vbNullString
        bookIssues = 0
        bookIssueDetail = vbNullString
        foundChapters = 0

        If i < nH1 Then
            bookEndPos = h1Starts(i + 1) - 1
        Else
            bookEndPos = docEnd
        End If
        If bookEndPos > docEnd Then bookEndPos = docEnd

        Dim h1Text As String
        h1Text = h1Names(i)

        Dim BookID As Long
        BookID = LookupBookID(h1Text)

        If BookID = 0 Then
            sOut = sOut & "?? UNKNOWN H1 [" & h1Text & "] - skip" & NL
            issues = issues & "  Unknown H1 text: [" & h1Text & "]" & NL
            issuesCount = issuesCount + 1
        Else
            seenBookID(BookID) = True
            Dim expectedChapters As Long
            expectedChapters = CLng(books(BookID)(2))

            AuditOneBook oDoc, h1Starts(i), bookEndPos, CStr(books(BookID)(1)), _
                          expectedChapters, foundChapters, chapterReport, _
                          bookIssues, bookIssueDetail, totalExpected, totalFound

            Dim bookStatus As String
            If foundChapters = expectedChapters And bookIssues = 0 Then
                bookStatus = "OK"
            Else
                bookStatus = "ISSUES"
            End If

            sOut = sOut & PadRight(CStr(books(BookID)(1)), 22) & _
                   "expected chapters=" & PadLeft(CStr(expectedChapters), 3) & _
                   "  found=" & PadLeft(CStr(foundChapters), 3) & _
                   "  " & bookStatus & NL
            sOut = sOut & chapterReport
            issuesCount = issuesCount + bookIssues
            If Len(bookIssueDetail) > 0 Then issues = issues & bookIssueDetail
        End If
    Next i

    ' Missing-book check
    Dim missing As String
    Dim k As Long
    For k = 1 To 66
        If Not seenBookID(k) Then
            missing = missing & "  Missing book: " & CStr(books(k)(1)) & " (BookID " & k & ")" & NL
            issuesCount = issuesCount + 1
        End If
    Next k

    sOut = sOut & NL
    If Len(missing) > 0 Then
        sOut = sOut & "MISSING BOOKS:" & NL & missing & NL
    End If
    If Len(issues) > 0 Then
        sOut = sOut & "ISSUES FOUND:" & NL & issues & NL
    End If

    sOut = sOut & "SUMMARY: " & totalFound & " / " & totalExpected & _
           " verses found, " & issuesCount & " structural issue(s)." & NL

    Debug.Print sOut
    If bWriteFile Then WriteAuditFile sOut

    EndTimer "AuditVerseMarkerStructure", t
End Sub

' --------------------------------------------------------------------------
' AuditOneBook - chapter and verse counts for a single book range
' --------------------------------------------------------------------------
Private Sub AuditOneBook(ByVal oDoc As Object, _
                         ByVal bookStart As Long, ByVal bookEnd As Long, _
                         ByVal bookName As String, _
                         ByVal expectedChapters As Long, _
                         ByRef foundChapters As Long, _
                         ByRef chapterReport As String, _
                         ByRef bookIssues As Long, _
                         ByRef bookIssueDetail As String, _
                         ByRef totalExpected As Long, _
                         ByRef totalFound As Long)
    Const NL As String = vbCrLf

    Dim h2Starts(1 To 250) As Long
    Dim nH2 As Long
    nH2 = 0

    Dim oRng As Object
    Set oRng = oDoc.Range(bookStart, bookEnd)
    Dim oPara As Object
    For Each oPara In oRng.Paragraphs
        If oPara.style.NameLocal = "Heading 2" Then
            nH2 = nH2 + 1
            If nH2 > 250 Then Exit For
            h2Starts(nH2) = oPara.Range.Start
        End If
    Next oPara

    foundChapters = nH2

    If nH2 <> expectedChapters Then
        bookIssues = bookIssues + 1
        bookIssueDetail = bookIssueDetail & "  " & bookName & ": chapter Count mismatch (expected " & _
                          expectedChapters & ", found " & nH2 & ")" & NL
    End If

    Dim chIdx As Long
    Dim chEnd As Long
    Dim foundVerses As Long
    Dim expectedVerses As Long
    Dim status As String

    For chIdx = 1 To nH2
        If chIdx < nH2 Then
            chEnd = h2Starts(chIdx + 1) - 1
        Else
            chEnd = bookEnd
        End If

        foundVerses = CountVerseMarkers(oDoc, h2Starts(chIdx), chEnd)
        expectedVerses = aeBibleCitationClass.VersesInChapter(bookName, chIdx)

        totalExpected = totalExpected + expectedVerses
        totalFound = totalFound + foundVerses

        If foundVerses = expectedVerses Then
            status = "OK"
        Else
            status = "MISMATCH"
            bookIssues = bookIssues + 1
            bookIssueDetail = bookIssueDetail & "  " & bookName & " " & chIdx & _
                              ": expected verses=" & expectedVerses & _
                              "  found=" & foundVerses & NL
        End If
        chapterReport = chapterReport & _
            "  ch " & PadLeft(CStr(chIdx), 3) & ": expected verses=" & PadLeft(CStr(expectedVerses), 3) & _
            "  found=" & PadLeft(CStr(foundVerses), 3) & "  " & status & NL
    Next chIdx
End Sub

' --------------------------------------------------------------------------
' CountVerseMarkers - Count Verse-marker character-style runs in a range
' --------------------------------------------------------------------------
Private Function CountVerseMarkers(ByVal oDoc As Object, _
                                    ByVal startPos As Long, _
                                    ByVal endPos As Long) As Long
    If endPos <= startPos Then
        CountVerseMarkers = 0
        Exit Function
    End If

    Dim oRng As Object
    Set oRng = oDoc.Range(startPos, endPos)

    Dim Count As Long
    Dim safety As Long
    Count = 0
    safety = 0

    With oRng.Find
        .ClearFormatting
        .style = oDoc.Styles("Verse marker")
        .Text = ""
        .Forward = True
        .Wrap = 0     ' wdFindStop
        .Format = True
        .MatchWildcards = False
        Do While .Execute
            Count = Count + 1
            safety = safety + 1
            If safety > 20000 Then Exit Do
            ' Advance past the match; oRng has collapsed to the matched run
            oRng.Start = oRng.End
            If oRng.Start >= endPos Then Exit Do
            oRng.End = endPos
        Loop
    End With

    CountVerseMarkers = Count
End Function

' --------------------------------------------------------------------------
' LookupBookID - map H1 text to canonical BookID 1-66 via the citation
' class alias map (accepts canonical names and all SBL aliases). Returns
' 0 if the text is not a recognised book alias.
' --------------------------------------------------------------------------
Private Function LookupBookID(ByVal h1Text As String) As Long
    Dim bID As Long
    bID = 0
    On Error Resume Next
    aeBibleCitationClass.ResolveAlias h1Text, bID
    On Error GoTo 0
    LookupBookID = bID
End Function

' --------------------------------------------------------------------------
' Padding helpers
' --------------------------------------------------------------------------
Private Function PadRight(ByVal s As String, ByVal n As Long) As String
    If Len(s) >= n Then
        PadRight = Left(s, n)
    Else
        PadRight = s & Space(n - Len(s))
    End If
End Function

Private Function PadLeft(ByVal s As String, ByVal n As Long) As String
    If Len(s) >= n Then
        PadLeft = Right(s, n)
    Else
        PadLeft = Space(n - Len(s)) & s
    End If
End Function

' --------------------------------------------------------------------------
' WriteAuditFile - write the report to rpt\VerseStructureAudit.txt
' --------------------------------------------------------------------------
Private Sub WriteAuditFile(ByVal sContent As String)
    Dim oFSO As Object
    Dim oStream As Object
    Dim sPath As String
    sPath = ActiveDocument.Path & "\rpt\VerseStructureAudit.txt"
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.CreateTextFile(sPath, True, False)   ' ASCII
    oStream.Write sContent
    oStream.Close
End Sub

' ==========================================================================
' AuditSelahUsage
' ==========================================================================
' Read-only diagnostic. Walks the main story for every "Selah" character-
' style run, reports the enclosing paragraph's properties and whether
' Phase 2 of the VerseText rollout (ConvertBodyTextVersesToVerseText)
' would convert it.
'
' The Phase 2 conversion rule converts a paragraph when both:
'   1. paragraph.Style.NameLocal = "BodyText"
'   2. paragraph.Range.Characters(1).Style.NameLocal = "Chapter Verse marker"
'
' This audit surfaces:
'   - Selah runs inside verse paragraphs (will convert cleanly with Phase 2)
'   - Selah-only paragraphs or other edge cases (need a policy decision
'     before Phase 2 locks in the conversion rule)
'
' Output: rpt\SelahUsageAudit.txt (when bWriteFile = True) plus Immediate
' window summary.
'
' Usage:
'   AuditSelahUsage
'   AuditSelahUsage False        ' Immediate only, no file
' ==========================================================================
Public Sub AuditSelahUsage(Optional ByVal bWriteFile As Boolean = True)
    On Error GoTo PROC_ERR
    Dim t As Double
    StartTimer "AuditSelahUsage", t

    Dim oDoc As Object
    Set oDoc = ActiveDocument

    Dim sOut As String
    Const NL As String = vbCrLf
    sOut = "---- AuditSelahUsage: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ----" & NL & NL

    Dim oRng As Object
    Set oRng = oDoc.Content

    Dim totalCount As Long
    Dim convertCount As Long
    Dim keepCount As Long
    Dim policyFlagCount As Long
    Dim safety As Long

    With oRng.Find
        .ClearFormatting
        .style = oDoc.Styles("Selah")
        .Text = ""
        .Forward = True
        .Wrap = 0     ' wdFindStop
        .Format = True
        .MatchWildcards = False

        Do While .Execute
            Dim oPara As Object
            Set oPara = oRng.Paragraphs(1)

            Dim paraStyle As String
            paraStyle = oPara.style.NameLocal

            Dim firstCharStyle As String
            firstCharStyle = oPara.Range.Characters(1).style.NameLocal

            Dim qualifies As Boolean
            qualifies = (paraStyle = "BodyText" And firstCharStyle = "Chapter Verse marker")

            ' Position of Selah within paragraph: START / END / MID
            Dim selahOffset As Long
            selahOffset = oRng.Start - oPara.Range.Start
            Dim paraTextLen As Long
            paraTextLen = oPara.Range.End - oPara.Range.Start - 1   ' exclude paragraph mark
            Dim posLabel As String
            If selahOffset = 0 Then
                posLabel = "START"
            ElseIf selahOffset >= paraTextLen - 8 Then
                posLabel = "END"
            Else
                posLabel = "MID"
            End If

            ' Excerpt: first 80 chars of paragraph, vbCr stripped
            Dim excerpt As String
            excerpt = Left$(Replace(oPara.Range.Text, vbCr, ""), 80)

            totalCount = totalCount + 1
            Dim phase2 As String
            If qualifies Then
                phase2 = "CONVERT"
                convertCount = convertCount + 1
            Else
                phase2 = "KEEP-AS-" & paraStyle
                keepCount = keepCount + 1
            End If

            sOut = sOut & "Run #" & totalCount & " | ParaStart=" & oPara.Range.Start & _
                   " | Style=" & paraStyle & " | first-char-style=" & firstCharStyle & _
                   " | Phase2: " & phase2 & NL
            sOut = sOut & "  Selah at " & posLabel & " of paragraph (offset " & _
                   selahOffset & " of " & paraTextLen & ")" & NL
            sOut = sOut & "  Excerpt: """ & excerpt & """" & NL

            ' Flag Selah-only or BodyText-with-Selah-as-first-char as policy candidates
            If Not qualifies And paraStyle = "BodyText" Then
                sOut = sOut & "  ** POLICY DECISION: BodyText paragraph not caught by Phase 2 rule." & NL
                policyFlagCount = policyFlagCount + 1
            End If
            sOut = sOut & NL

            ' Advance past this Selah run
            oRng.Start = oRng.End
            safety = safety + 1
            If safety > 5000 Then
                sOut = sOut & "*** Safety limit (5000 runs) reached, abort scan ***" & NL
                Exit Do
            End If
            If oRng.Start >= oDoc.Content.End Then Exit Do
            oRng.End = oDoc.Content.End
        Loop
    End With

    sOut = sOut & "---- Summary ----" & NL
    sOut = sOut & "Total Selah character runs: " & totalCount & NL
    sOut = sOut & "  CONVERT (verse paragraph, Phase 2 will reassign to VerseText): " & convertCount & NL
    sOut = sOut & "  KEEP-AS-other (paragraph not caught by Phase 2 rule): " & keepCount & NL
    sOut = sOut & "  Policy decision flags (BodyText paragraph not converted): " & policyFlagCount & NL

    Debug.Print sOut
    If bWriteFile Then WriteSelahUsageFile sOut

    EndTimer "AuditSelahUsage", t
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure AuditSelahUsage of Module basVerseStructureAudit"
    Resume PROC_EXIT
End Sub

' --------------------------------------------------------------------------
' WriteSelahUsageFile - write the report to rpt\SelahUsageAudit.txt
' --------------------------------------------------------------------------
Private Sub WriteSelahUsageFile(ByVal sContent As String)
    Dim oFSO As Object
    Dim oStream As Object
    Dim sPath As String
    sPath = ActiveDocument.Path & "\rpt\SelahUsageAudit.txt"
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.CreateTextFile(sPath, True, False)   ' ASCII
    oStream.Write sContent
    oStream.Close
End Sub

' ==========================================================================
' AuditOrphanBodyTextParagraphs
' ==========================================================================
' Read-only diagnostic. Walks the main story to find true verse-continuation
' orphans: BodyText paragraphs that sit BETWEEN two verse paragraphs in the
' same chapter, but do not themselves begin with a "Chapter Verse marker"
' character-style run. These are typically created by a stray paragraph
' mark mid-verse.
'
' Detection algorithm (single-pass with a buffer):
'   - Track current book (last H1) and "seen first verse in current chapter".
'   - When we see a BodyText paragraph with CVM first-char (verse) AFTER
'     having already seen one in the same chapter, FLUSH the buffer of
'     non-CVM BodyText paragraphs accumulated since the previous verse -
'     these are confirmed orphans.
'   - On H2 (new chapter) or H1 (new book), DISCARD any pending buffer -
'     they were post-last-verse content (chapter-end or book-end), not
'     orphans.
'   - Non-CVM BodyText paragraphs BEFORE the first verse of a chapter
'     are chapter intros and are excluded.
'   - End of document: any remaining buffer is post-last-verse-of-last-book
'     content, also discarded.
'
' Excluded as legitimate non-verse BodyText:
'   - Front matter (before first H1)
'   - Book introductions (between H1 and first H2 of book)
'   - Chapter intros (between H2 and first verse of chapter)
'   - Chapter-end content (after last verse, before next H2 or H1)
'   - Whitespace-only paragraphs in any of the above zones
'
' Per-orphan report:
'   - Book name (last H1 seen at the time of flush)
'   - Paragraph start char position (clickable via Word Ctrl+G)
'   - First character's style name ("Selah", "Verse marker", etc., or
'     "(empty)" for empty paragraphs)
'   - Size category: EMPTY / SHORT (<30 chars) / LONG (>=30 chars)
'   - 80-char excerpt for visual identification
'
' Summary also reports excluded counts (chapter intros / chapter-end
' content) so the noise-vs-signal ratio is visible.
'
' Output: rpt\OrphanBodyTextAudit.txt (when bWriteFile = True) plus
' Immediate window summary.
'
' Usage:
'   AuditOrphanBodyTextParagraphs
'   AuditOrphanBodyTextParagraphs False        ' Immediate only, no file
' ==========================================================================
Public Sub AuditOrphanBodyTextParagraphs(Optional ByVal bWriteFile As Boolean = True)
    On Error GoTo PROC_ERR
    Dim t As Double
    StartTimer "AuditOrphanBodyTextParagraphs", t

    Const BUF_CAP As Long = 1000

    Dim oDoc As Object
    Set oDoc = ActiveDocument

    Dim sOut As String
    Const NL As String = vbCrLf
    sOut = "---- AuditOrphanBodyTextParagraphs: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ----" & NL & NL
    sOut = sOut & "Scope: BodyText paragraphs that sit BETWEEN two verse paragraphs in the" & NL
    sOut = sOut & "       same chapter and do not begin with a 'Chapter Verse marker' run." & NL
    sOut = sOut & "       Chapter intros (before first verse) and chapter-end content" & NL
    sOut = sOut & "       (after last verse) are excluded." & NL & NL

    Dim oPara As Object
    Dim currentBook As String
    currentBook = "(front matter)"
    Dim seenFirstVerseInCurrentChapter As Boolean
    seenFirstVerseInCurrentChapter = False

    ' Buffer columns: 1=ParaStart, 2=firstCharStyle, 3=sizeCat, 4=paraLen, 5=excerpt
    Dim potentialBuffer() As String
    ReDim potentialBuffer(1 To BUF_CAP, 1 To 5)
    Dim bufCount As Long
    bufCount = 0

    Dim totalCount As Long
    Dim emptyCount As Long
    Dim shortCount As Long
    Dim longCount As Long
    Dim chapterIntroCount As Long
    Dim chapterEndCount As Long

    For Each oPara In oDoc.Paragraphs
        Dim StyleName As String
        StyleName = oPara.style.NameLocal

        Select Case StyleName
            Case "Heading 1"
                ' Discard pending buffer (post-last-verse of previous book)
                chapterEndCount = chapterEndCount + bufCount
                bufCount = 0

                currentBook = Trim$(Replace(oPara.Range.Text, vbCr, ""))
                seenFirstVerseInCurrentChapter = False

            Case "Heading 2"
                ' Discard pending buffer (post-last-verse of previous chapter)
                chapterEndCount = chapterEndCount + bufCount
                bufCount = 0

                seenFirstVerseInCurrentChapter = False

            Case "BodyText"
                Dim paraText As String
                paraText = Replace(oPara.Range.Text, vbCr, "")
                Dim firstCharStyle As String
                If Len(paraText) = 0 Then
                    firstCharStyle = "(empty)"
                Else
                    On Error Resume Next
                    firstCharStyle = oPara.Range.Characters(1).style.NameLocal
                    If Err.Number <> 0 Then firstCharStyle = "(error)"
                    Err.Clear
                    On Error GoTo PROC_ERR
                End If

                If firstCharStyle = "Chapter Verse marker" Then
                    ' Verse paragraph
                    If seenFirstVerseInCurrentChapter And bufCount > 0 Then
                        ' Flush buffer - confirmed orphans (between verses)
                        Dim i As Long
                        For i = 1 To bufCount
                            totalCount = totalCount + 1
                            Select Case potentialBuffer(i, 3)
                                Case "EMPTY":  emptyCount = emptyCount + 1
                                Case "SHORT":  shortCount = shortCount + 1
                                Case "LONG":   longCount = longCount + 1
                            End Select

                            Dim sizeLabel As String
                            If potentialBuffer(i, 3) = "EMPTY" Then
                                sizeLabel = "EMPTY"
                            Else
                                sizeLabel = potentialBuffer(i, 3) & " (" & potentialBuffer(i, 4) & " chars)"
                            End If

                            sOut = sOut & "Orphan #" & totalCount & " | Book: " & currentBook & _
                                   " | ParaStart=" & potentialBuffer(i, 1) & _
                                   " | first-char-style=" & potentialBuffer(i, 2) & _
                                   " | Size: " & sizeLabel & NL
                            sOut = sOut & "  Excerpt: """ & potentialBuffer(i, 5) & """" & NL & NL
                        Next i
                        bufCount = 0
                    End If
                    seenFirstVerseInCurrentChapter = True
                Else
                    ' Non-verse BodyText
                    If seenFirstVerseInCurrentChapter Then
                        ' After first verse - add to potential orphan buffer
                        If bufCount < BUF_CAP Then
                            bufCount = bufCount + 1
                            Dim paraLen As Long
                            paraLen = Len(paraText)

                            potentialBuffer(bufCount, 1) = CStr(oPara.Range.Start)
                            potentialBuffer(bufCount, 2) = firstCharStyle
                            If paraLen = 0 Then
                                potentialBuffer(bufCount, 3) = "EMPTY"
                            ElseIf paraLen < 30 Then
                                potentialBuffer(bufCount, 3) = "SHORT"
                            Else
                                potentialBuffer(bufCount, 3) = "LONG"
                            End If
                            potentialBuffer(bufCount, 4) = CStr(paraLen)
                            potentialBuffer(bufCount, 5) = Left$(paraText, 80)
                        End If
                        ' (silently drop overflow beyond BUF_CAP - extremely unlikely)
                    Else
                        ' Before first verse of chapter - chapter intro, exclude
                        chapterIntroCount = chapterIntroCount + 1
                    End If
                End If
        End Select
    Next oPara

    ' End of doc: discard remaining buffer (post-last-verse of last book)
    chapterEndCount = chapterEndCount + bufCount

    sOut = sOut & "---- Summary ----" & NL
    sOut = sOut & "Confirmed orphans (BodyText between two verses in same chapter): " & totalCount & NL
    sOut = sOut & "  EMPTY (0 chars): " & emptyCount & NL
    sOut = sOut & "  SHORT (<30 chars): " & shortCount & NL
    sOut = sOut & "  LONG (>=30 chars): " & longCount & NL
    sOut = sOut & NL
    sOut = sOut & "Excluded as legitimate non-verse content:" & NL
    sOut = sOut & "  Chapter intros (before first verse of chapter): " & chapterIntroCount & NL
    sOut = sOut & "  Chapter-end content (after last verse, before next H2/H1): " & chapterEndCount & NL

    Debug.Print sOut
    If bWriteFile Then WriteOrphanFile sOut

    EndTimer "AuditOrphanBodyTextParagraphs", t
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure AuditOrphanBodyTextParagraphs of Module basVerseStructureAudit"
    Resume PROC_EXIT
End Sub

' --------------------------------------------------------------------------
' WriteOrphanFile - write the report to rpt\OrphanBodyTextAudit.txt
' --------------------------------------------------------------------------
Private Sub WriteOrphanFile(ByVal sContent As String)
    Dim oFSO As Object
    Dim oStream As Object
    Dim sPath As String
    sPath = ActiveDocument.Path & "\rpt\OrphanBodyTextAudit.txt"
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.CreateTextFile(sPath, True, False)   ' ASCII
    oStream.Write sContent
    oStream.Close
End Sub
