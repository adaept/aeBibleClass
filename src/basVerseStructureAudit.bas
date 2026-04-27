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

    ' Canonical 66-book reference data
    Dim canonNames(1 To 66) As String
    Dim canonChapters(1 To 66) As Long
    PopulateCanonical canonNames, canonChapters

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
    docEnd = oDoc.Content.End

    For i = 1 To nH1
        If i < nH1 Then
            bookEndPos = h1Starts(i + 1) - 1
        Else
            bookEndPos = docEnd
        End If
        If bookEndPos > docEnd Then bookEndPos = docEnd

        Dim h1Text As String
        h1Text = h1Names(i)

        Dim BookID As Long
        BookID = LookupBookID(h1Text, canonNames)

        If BookID = 0 Then
            sOut = sOut & "?? UNKNOWN H1 [" & h1Text & "] - skip" & NL
            issues = issues & "  Unknown H1 text: [" & h1Text & "]" & NL
            issuesCount = issuesCount + 1
        Else
            seenBookID(BookID) = True
            Dim expectedChapters As Long
            expectedChapters = canonChapters(BookID)

            Dim foundChapters As Long
            Dim chapterReport As String
            Dim bookIssues As Long
            Dim bookIssueDetail As String

            AuditOneBook oDoc, h1Starts(i), bookEndPos, canonNames(BookID), _
                          expectedChapters, foundChapters, chapterReport, _
                          bookIssues, bookIssueDetail, totalExpected, totalFound

            Dim bookStatus As String
            If foundChapters = expectedChapters And bookIssues = 0 Then
                bookStatus = "OK"
            Else
                bookStatus = "ISSUES"
            End If

            sOut = sOut & PadRight(canonNames(BookID), 22) & _
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
            missing = missing & "  Missing book: " & canonNames(k) & " (BookID " & k & ")" & NL
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
' LookupBookID - map H1 text to canonical book index 1-66, or 0 if unknown
' --------------------------------------------------------------------------
Private Function LookupBookID(ByVal h1Text As String, ByRef canonNames() As String) As Long
    Dim cleaned As String
    cleaned = UCase(Trim(h1Text))
    Dim k As Long
    For k = 1 To 66
        If UCase(canonNames(k)) = cleaned Then
            LookupBookID = k
            Exit Function
        End If
    Next k
    LookupBookID = 0
End Function

' --------------------------------------------------------------------------
' PopulateCanonical - fill the 66-book reference (mirrors basTEST_aeBibleCitationClass)
' --------------------------------------------------------------------------
Private Sub PopulateCanonical(ByRef names() As String, ByRef chapters() As Long)
    names(1) = "Genesis":          chapters(1) = 50
    names(2) = "Exodus":           chapters(2) = 40
    names(3) = "Leviticus":        chapters(3) = 27
    names(4) = "Numbers":          chapters(4) = 36
    names(5) = "Deuteronomy":      chapters(5) = 34
    names(6) = "Joshua":           chapters(6) = 24
    names(7) = "Judges":           chapters(7) = 21
    names(8) = "Ruth":             chapters(8) = 4
    names(9) = "1 Samuel":         chapters(9) = 31
    names(10) = "2 Samuel":        chapters(10) = 24
    names(11) = "1 Kings":         chapters(11) = 22
    names(12) = "2 Kings":         chapters(12) = 25
    names(13) = "1 Chronicles":    chapters(13) = 29
    names(14) = "2 Chronicles":    chapters(14) = 36
    names(15) = "Ezra":            chapters(15) = 10
    names(16) = "Nehemiah":        chapters(16) = 13
    names(17) = "Esther":          chapters(17) = 10
    names(18) = "Job":             chapters(18) = 42
    names(19) = "Psalms":          chapters(19) = 150
    names(20) = "Proverbs":        chapters(20) = 31
    names(21) = "Ecclesiastes":    chapters(21) = 12
    names(22) = "Solomon":         chapters(22) = 8     ' project canonical; SBL output is "Song"
    names(23) = "Isaiah":          chapters(23) = 66
    names(24) = "Jeremiah":        chapters(24) = 52
    names(25) = "Lamentations":    chapters(25) = 5
    names(26) = "Ezekiel":         chapters(26) = 48
    names(27) = "Daniel":          chapters(27) = 12
    names(28) = "Hosea":           chapters(28) = 14
    names(29) = "Joel":            chapters(29) = 3
    names(30) = "Amos":            chapters(30) = 9
    names(31) = "Obadiah":         chapters(31) = 1
    names(32) = "Jonah":           chapters(32) = 4
    names(33) = "Micah":           chapters(33) = 7
    names(34) = "Nahum":           chapters(34) = 3
    names(35) = "Habakkuk":        chapters(35) = 3
    names(36) = "Zephaniah":       chapters(36) = 3
    names(37) = "Haggai":          chapters(37) = 2
    names(38) = "Zechariah":       chapters(38) = 14
    names(39) = "Malachi":         chapters(39) = 4
    names(40) = "Matthew":         chapters(40) = 28
    names(41) = "Mark":            chapters(41) = 16
    names(42) = "Luke":            chapters(42) = 24
    names(43) = "John":            chapters(43) = 21
    names(44) = "Acts":            chapters(44) = 28
    names(45) = "Romans":          chapters(45) = 16
    names(46) = "1 Corinthians":   chapters(46) = 16
    names(47) = "2 Corinthians":   chapters(47) = 13
    names(48) = "Galatians":       chapters(48) = 6
    names(49) = "Ephesians":       chapters(49) = 6
    names(50) = "Philippians":     chapters(50) = 4
    names(51) = "Colossians":      chapters(51) = 4
    names(52) = "1 Thessalonians": chapters(52) = 5
    names(53) = "2 Thessalonians": chapters(53) = 3
    names(54) = "1 Timothy":       chapters(54) = 6
    names(55) = "2 Timothy":       chapters(55) = 4
    names(56) = "Titus":           chapters(56) = 3
    names(57) = "Philemon":        chapters(57) = 1
    names(58) = "Hebrews":         chapters(58) = 13
    names(59) = "James":           chapters(59) = 5
    names(60) = "1 Peter":         chapters(60) = 5
    names(61) = "2 Peter":         chapters(61) = 3
    names(62) = "1 John":          chapters(62) = 5
    names(63) = "2 John":          chapters(63) = 1
    names(64) = "3 John":          chapters(64) = 1
    names(65) = "Jude":            chapters(65) = 1
    names(66) = "Revelation":      chapters(66) = 22
End Sub

' --------------------------------------------------------------------------
' Padding helpers
' --------------------------------------------------------------------------
Private Function PadRight(ByVal s As String, ByVal n As Long) As String
    If Len(s) >= n Then
        PadRight = Left(s, n)
    Else
        PadRight = s & space(n - Len(s))
    End If
End Function

Private Function PadLeft(ByVal s As String, ByVal n As Long) As String
    If Len(s) >= n Then
        PadLeft = Right(s, n)
    Else
        PadLeft = space(n - Len(s)) & s
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
