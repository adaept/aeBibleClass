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
