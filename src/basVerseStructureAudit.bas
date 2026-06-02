Attribute VB_Name = "basVerseStructureAudit"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

' Module-level cache for GetMarkerTotals - persists across aeBibleClass
' instances so slot 83 reuses slot 82's walk in single-test (OneTest) mode.
' Invalidated when ActiveDocument.FullName changes.
Private m_cachedDocName As String
Private m_cachedVMTotal As Long
Private m_cachedCVMTotal As Long
Private m_cacheValid As Boolean

' ==========================================================================
' AuditVerseMarkerStructure
' ==========================================================================
' Walk the Bible body and verify chapter / verse counts against
' aeBibleCitationClass canonical data. Read-only; produces a report file
' plus an Immediate-window summary.
'
' Project verse-marker rule (the structural contract):
'   Every verse paragraph (now styled VerseText, formerly BodyText) leads
'   with a "Chapter Verse marker" character-style run (chapter number)
'   IMMEDIATELY FOLLOWED BY a "Verse marker" character-style run (verse
'   number). One CVM + one VM per verse - no exceptions.
'
' Invariants verified (core four; advanced 5-6 deferred):
'   1. Books present     - every canonical 66-book has a Heading 1.
'   2. Chapter counts    - per-book Heading 2 Count matches ChaptersInBook.
'   3. Verse counts      - per-chapter Verse marker Count matches VersesInChapter.
'   4. CVM coverage      - per-chapter Chapter Verse marker Count equals
'                          Verse marker Count (one CVM+VM per verse rule).
'                          Catches verses missing a leading CVM run that
'                          would otherwise pass the VM Count check.
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
    Dim foundCVMs As Long
    Dim expectedVerses As Long
    Dim status As String
    Dim cvmStatus As String

    For chIdx = 1 To nH2
        If chIdx < nH2 Then
            chEnd = h2Starts(chIdx + 1) - 1
        Else
            chEnd = bookEnd
        End If

        foundVerses = CountVerseMarkers(oDoc, h2Starts(chIdx), chEnd)
        foundCVMs = CountChapterVerseMarkers(oDoc, h2Starts(chIdx), chEnd)
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

        ' Project rule: one Chapter Verse marker + one Verse marker per
        ' verse. CVM Count must equal VM Count in every chapter.
        If foundCVMs = foundVerses Then
            cvmStatus = "OK"
        Else
            cvmStatus = "CVM-MISMATCH"
            bookIssues = bookIssues + 1
            bookIssueDetail = bookIssueDetail & "  " & bookName & " " & chIdx & _
                              ": CVM Count " & foundCVMs & " <> VM Count " & _
                              foundVerses & " (one CVM+VM per verse rule)" & NL
        End If

        chapterReport = chapterReport & _
            "  ch " & PadLeft(CStr(chIdx), 3) & ": expected verses=" & PadLeft(CStr(expectedVerses), 3) & _
            "  found VM=" & PadLeft(CStr(foundVerses), 3) & _
            "  found CVM=" & PadLeft(CStr(foundCVMs), 3) & _
            "  " & status & "/" & cvmStatus & NL
    Next chIdx
End Sub

' ==========================================================================
' GetMarkerTotals
' ==========================================================================
' Walks ActiveDocument.Paragraphs once. For every VerseText paragraph:
'   - increments cvmTotal if Characters(1).Style is "Chapter Verse marker"
'   - increments vmTotal if any of Characters(1..12) is "Verse marker"
'
' Semantics: counts VerseText paragraphs satisfying the design rule
' (CVM at start, VM in the leading marker run). Does NOT see CVM/VM runs
' applied OUTSIDE VerseText paragraphs - that drift is caught by the
' presence audit in aeBibleClass.CountAuditCharacterStyles_ToFile (slot 81).
'
' Why this shape: the per-chapter Find pattern (CountVerseMarkers et al.)
' is correct but slow (300-2700 s) because Word's Find degenerates on
' character-style runs in a large document. Characters(i).Style.NameLocal
' is unambiguous (Range.Words can fall back to paragraph style on mixed
' spans) and a single pass through 35k paragraphs completes in seconds.
'
' Cache: results memoized at module scope, keyed by ActiveDocument.FullName,
' so slot 83 reuses slot 82's walk even across separate aeBibleClass
' instances (OneTest mode reinstantiates the class per test).
'
' Consumers: aeBibleClass.EnsureVerseMarkerCounts (test slots 82 + 83).
' ==========================================================================
Public Sub GetMarkerTotals(ByRef vmTotal As Long, ByRef cvmTotal As Long)
    Dim currentDoc As String
    currentDoc = ActiveDocument.FullName

    If m_cacheValid And m_cachedDocName = currentDoc Then
        vmTotal = m_cachedVMTotal
        cvmTotal = m_cachedCVMTotal
        Exit Sub
    End If

    vmTotal = 0
    cvmTotal = 0

    Dim oPara As Object
    Dim numChars As Long
    Dim maxScan As Long
    Dim j As Long
    Dim charStyle As String
    Dim firstCharStyle As String

    For Each oPara In ActiveDocument.Paragraphs
        If oPara.style.NameLocal = "VerseText" Then
            firstCharStyle = oPara.Range.Characters(1).style.NameLocal
            If firstCharStyle = "Chapter Verse marker" Then
                cvmTotal = cvmTotal + 1
            End If
            numChars = oPara.Range.Characters.Count
            maxScan = 12
            If numChars < maxScan Then maxScan = numChars
            For j = 1 To maxScan
                charStyle = oPara.Range.Characters(j).style.NameLocal
                If charStyle = "Verse marker" Then
                    vmTotal = vmTotal + 1
                    Exit For
                End If
            Next j
        End If
    Next oPara

    m_cachedDocName = currentDoc
    m_cachedVMTotal = vmTotal
    m_cachedCVMTotal = cvmTotal
    m_cacheValid = True
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
' CountChapterVerseMarkers - Count Chapter-Verse-marker character-style
' runs in a range. Per project rule, one CVM appears at the start of every
' verse paragraph (immediately followed by one Verse marker), so this
' Count must equal CountVerseMarkers for the same range.
' --------------------------------------------------------------------------
Private Function CountChapterVerseMarkers(ByVal oDoc As Object, _
                                           ByVal startPos As Long, _
                                           ByVal endPos As Long) As Long
    If endPos <= startPos Then
        CountChapterVerseMarkers = 0
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
        .style = oDoc.Styles("Chapter Verse marker")
        .Text = ""
        .Forward = True
        .Wrap = 0     ' wdFindStop
        .Format = True
        .MatchWildcards = False
        Do While .Execute
            Count = Count + 1
            safety = safety + 1
            If safety > 20000 Then Exit Do
            oRng.Start = oRng.End
            If oRng.Start >= endPos Then Exit Do
            oRng.End = endPos
        Loop
    End With

    CountChapterVerseMarkers = Count
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
' AuditCharStyleUsage
' ==========================================================================
' Read-only diagnostic. Walks the main story for every character-style
' run of the named style, reports the enclosing paragraph's properties
' and whether Phase 2 of the VerseText rollout (ConvertBodyTextVersesToVerseText)
' would convert it.
'
' The Phase 2 conversion rule converts a paragraph when both:
'   1. paragraph.Style.NameLocal = "BodyText"
'   2. paragraph.Range.Characters(1).Style.NameLocal = "Chapter Verse marker"
'
' This audit surfaces:
'   - Character runs inside verse paragraphs (will convert cleanly with Phase 2)
'   - Edge cases (need a policy decision before Phase 2 locks in the rule)
'
' Pre-flight: aborts cleanly if StyleName is not present or not a
' character style.
'
' Output: rpt\<SafeFileName(StyleName)>UsageAudit.txt (when bWriteFile = True)
' plus Immediate window summary.
'
' Usage:
'   AuditCharStyleUsage "Selah"
'   AuditCharStyleUsage "EmphasisBlack"
'   AuditCharStyleUsage "Words of Jesus", False             ' Immediate only, no file
'   AuditCharStyleUsage "Chapter Verse marker", True, True  ' anomalies-only mode
' ==========================================================================
Public Sub AuditCharStyleUsage(ByVal StyleName As String, _
                                Optional ByVal bWriteFile As Boolean = True, _
                                Optional ByVal bAnomaliesOnly As Boolean = False)
    On Error GoTo PROC_ERR
    Dim t As Double
    StartTimer "AuditCharStyleUsage(" & StyleName & ")", t

    Dim oDoc As Object
    Set oDoc = ActiveDocument

    ' Pre-flight: style exists?
    Dim oStyle As Object
    On Error Resume Next
    Set oStyle = oDoc.Styles(StyleName)
    On Error GoTo PROC_ERR
    If oStyle Is Nothing Then
        MsgBox "AuditCharStyleUsage: style """ & StyleName & """ not found.", _
               vbExclamation, "AuditCharStyleUsage"
        Exit Sub
    End If
    ' Pre-flight: must be a character style.
    If oStyle.Type <> wdStyleTypeCharacter Then
        MsgBox "AuditCharStyleUsage: style """ & StyleName & """ is not a " & _
               "character style (Type=" & oStyle.Type & "). Aborting.", _
               vbExclamation, "AuditCharStyleUsage"
        Exit Sub
    End If

    Dim startTime As Date
    Dim startTick As Double
    startTime = Now
    startTick = Timer

    Dim sOut As String
    Const NL As String = vbCrLf
    sOut = "---- AuditCharStyleUsage(""" & StyleName & """): " & _
           Format(startTime, "yyyy-mm-dd hh:nn:ss") & " ----" & NL
    If bAnomaliesOnly Then
        sOut = sOut & "Mode: ANOMALIES ONLY (suppress runs where paraStyle=VerseText AND position=START)" & NL
    End If
    sOut = sOut & NL

    Dim oRng As Object
    Set oRng = oDoc.Content

    Dim totalCount As Long
    Dim convertCount As Long
    Dim keepCount As Long
    Dim policyFlagCount As Long
    Dim anomalyCount As Long
    Dim safety As Long
    Dim safetyCap As Long
    If bAnomaliesOnly Then
        safetyCap = 100000
    Else
        safetyCap = 5000
    End If

    With oRng.Find
        .ClearFormatting
        .style = oDoc.Styles(StyleName)
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

            ' Position of run within paragraph: START / END / MID
            Dim runOffset As Long
            runOffset = oRng.Start - oPara.Range.Start
            Dim paraTextLen As Long
            paraTextLen = oPara.Range.End - oPara.Range.Start - 1   ' exclude paragraph mark
            Dim posLabel As String
            If runOffset = 0 Then
                posLabel = "START"
            ElseIf runOffset >= paraTextLen - 8 Then
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

            Dim isAnomaly As Boolean
            isAnomaly = Not (paraStyle = "VerseText" And posLabel = "START")
            If isAnomaly Then anomalyCount = anomalyCount + 1

            If (Not bAnomaliesOnly) Or isAnomaly Then
                sOut = sOut & "Run #" & totalCount & " | ParaStart=" & oPara.Range.Start & _
                       " | Style=" & paraStyle & " | first-char-style=" & firstCharStyle & _
                       " | Phase2: " & phase2 & NL
                sOut = sOut & "  " & StyleName & " at " & posLabel & " of paragraph (offset " & _
                       runOffset & " of " & paraTextLen & ")" & NL
                sOut = sOut & "  Excerpt: """ & excerpt & """" & NL

                ' Flag BodyText paragraphs not caught by Phase 2 rule as policy candidates
                If Not qualifies And paraStyle = "BodyText" Then
                    sOut = sOut & "  ** POLICY DECISION: BodyText paragraph not caught by Phase 2 rule." & NL
                    policyFlagCount = policyFlagCount + 1
                End If
                sOut = sOut & NL
            ElseIf Not qualifies And paraStyle = "BodyText" Then
                ' Still count policy flags even when suppressed (anomalies-only would emit anyway since paraStyle <> VerseText)
                policyFlagCount = policyFlagCount + 1
            End If

            ' Advance past this run
            oRng.Start = oRng.End
            safety = safety + 1

            ' Progress heartbeat every 1000 runs; lets caller see the scan
            ' is alive and Ctrl+Break cleanly via DoEvents.
            If safety Mod 1000 = 0 Then
                Debug.Print "AuditCharStyleUsage(" & StyleName & "): " & safety & _
                            " runs scanned, " & anomalyCount & " anomalies, " & _
                            Format(Timer - startTick, "0.0") & " sec elapsed"
                DoEvents
            End If

            If safety > safetyCap Then
                sOut = sOut & "*** Safety limit (" & safetyCap & " runs) reached, abort scan ***" & NL
                Exit Do
            End If
            If oRng.Start >= oDoc.Content.End Then Exit Do
            oRng.End = oDoc.Content.End
        Loop
    End With

    sOut = sOut & "---- Summary ----" & NL
    sOut = sOut & "Total " & StyleName & " character runs: " & totalCount & NL
    sOut = sOut & "  Anomalies (paraStyle<>VerseText OR position<>START): " & anomalyCount & NL
    sOut = sOut & "  CONVERT (verse paragraph, Phase 2 will reassign to VerseText): " & convertCount & NL
    sOut = sOut & "  KEEP-AS-other (paragraph not caught by Phase 2 rule): " & keepCount & NL
    sOut = sOut & "  Policy decision flags (BodyText paragraph not converted): " & policyFlagCount & NL
    sOut = sOut & NL
    sOut = sOut & "Started:  " & Format(startTime, "yyyy-mm-dd hh:nn:ss") & NL
    sOut = sOut & "Finished: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & NL
    sOut = sOut & "Duration: " & Format(Timer - startTick, "0.00") & " sec" & NL

    Debug.Print sOut
    If bWriteFile Then WriteCharStyleUsageFile StyleName, sOut

    EndTimer "AuditCharStyleUsage(" & StyleName & ")", t
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure AuditCharStyleUsage of Module basVerseStructureAudit"
    Resume PROC_EXIT
End Sub

' --------------------------------------------------------------------------
' AuditSelahUsage - thin wrapper for backwards compatibility.
' --------------------------------------------------------------------------
Public Sub AuditSelahUsage(Optional ByVal bWriteFile As Boolean = True)
    AuditCharStyleUsage "Selah", bWriteFile
End Sub

' --------------------------------------------------------------------------
' WriteCharStyleUsageFile - write the report to
'   rpt\<SafeFileName(StyleName)>UsageAudit.txt
' Uses the public SafeFileName helper from basStyleInspector to sanitise
' the style name (spaces -> underscores, etc.) for path safety.
' --------------------------------------------------------------------------
Private Sub WriteCharStyleUsageFile(ByVal StyleName As String, _
                                    ByVal sContent As String)
    Dim oFSO As Object
    Dim oStream As Object
    Dim sPath As String
    sPath = ActiveDocument.Path & "\rpt\" & SafeFileName(StyleName) & "UsageAudit.txt"
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

' ==========================================================================
' GoToPos
' ==========================================================================
' Navigation helper. Selects the cursor at the given character offset and
' scrolls the document so the position is visible.
'
' Word 365's Ctrl+G (Go To) dialog does NOT support character-offset
' navigation - only Page / Section / Line / Bookmark / etc. This helper
' fills the gap so audit reports that emit ParaStart offsets (e.g.,
' AuditOrphanBodyTextParagraphs) can be navigated quickly from VBA.
'
' Usage (from Immediate window):
'   GoToPos 2231693
'   ' then Backspace in the document to merge an orphan into the
'   ' preceding paragraph.
' ==========================================================================
Public Sub GoToPos(ByVal pos As Long)
    On Error GoTo PROC_ERR
    ActiveDocument.Range(pos, pos).Select
    ActiveWindow.ScrollIntoView Selection.Range, True
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure GoToPos of Module basVerseStructureAudit"
    Resume PROC_EXIT
End Sub

' ==========================================================================
' AuditUnconvertedVerseParagraphs
' ==========================================================================
' Read-only diagnostic. Walks the main story to find paragraphs whose
' first character is "Chapter Verse marker" (verse paragraphs by
' definition) but whose paragraph style is neither "BodyText" nor
' "VerseText". These are verses that the Phase 2 conversion did not
' touch because the paragraph-style filter ("BodyText only") missed
' them. Most likely candidates are BodyTextIndent (Psalms / prophetic
' poetry continuations).
'
' Output groups results by paragraph style and reports up to 3 samples
' per group, plus total Count. Lets the user decide whether to extend
' the Phase 2 conversion rule, add a VerseTextIndent variant, or
' leave the styles alone.
'
' Output: rpt\UnconvertedVerseAudit.txt (when bWriteFile = True) plus
' Immediate window summary.
'
' Usage:
'   AuditUnconvertedVerseParagraphs
' ==========================================================================
Public Sub AuditUnconvertedVerseParagraphs(Optional ByVal bWriteFile As Boolean = True)
    On Error GoTo PROC_ERR

    Dim oDoc As Object
    Set oDoc = ActiveDocument

    Dim sOut As String
    Const NL As String = vbCrLf
    sOut = "---- AuditUnconvertedVerseParagraphs: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ----" & NL & NL
    sOut = sOut & "Scope: paragraphs whose first character has 'Chapter Verse marker'" & NL
    sOut = sOut & "       character style AND whose paragraph style is neither" & NL
    sOut = sOut & "       'BodyText' nor 'VerseText'." & NL & NL

    Dim styleNames(1 To 50) As String
    Dim styleCounts(1 To 50) As Long
    Dim styleSamples(1 To 50, 1 To 3) As String
    Dim styleCount As Long

    Dim oPara As Object
    Dim totalUnconverted As Long
    For Each oPara In oDoc.Paragraphs
        Dim psName As String
        psName = oPara.style.NameLocal
        If psName <> "BodyText" And psName <> "VerseText" Then
            If oPara.Range.End - oPara.Range.Start > 1 Then
                On Error Resume Next
                Dim fcs As String
                fcs = oPara.Range.Characters(1).style.NameLocal
                On Error GoTo PROC_ERR
                If fcs = "Chapter Verse marker" Then
                    totalUnconverted = totalUnconverted + 1

                    Dim k As Long
                    Dim found As Long
                    found = 0
                    For k = 1 To styleCount
                        If styleNames(k) = psName Then
                            found = k
                            Exit For
                        End If
                    Next k
                    If found = 0 Then
                        styleCount = styleCount + 1
                        If styleCount > 50 Then
                            ' Safety: too many distinct styles
                            sOut = sOut & "*** Safety: more than 50 distinct paragraph styles found." & NL
                            Exit For
                        End If
                        styleNames(styleCount) = psName
                        found = styleCount
                    End If
                    styleCounts(found) = styleCounts(found) + 1

                    If styleCounts(found) <= 3 Then
                        Dim excerpt As String
                        excerpt = Left$(Replace(oPara.Range.Text, vbCr, ""), 80)
                        styleSamples(found, styleCounts(found)) = _
                            "ParaStart=" & oPara.Range.Start & "  Excerpt: """ & excerpt & """"
                    End If
                End If
            End If
        End If
    Next oPara

    sOut = sOut & "Total unconverted verse paragraphs: " & totalUnconverted & NL & NL
    Dim i As Long, j As Long
    For i = 1 To styleCount
        sOut = sOut & "Paragraph style: " & styleNames(i) & "  (Count=" & styleCounts(i) & ")" & NL
        For j = 1 To 3
            If j <= styleCounts(i) Then
                sOut = sOut & "  Sample " & j & ": " & styleSamples(i, j) & NL
            End If
        Next j
        sOut = sOut & NL
    Next i

    Debug.Print sOut
    If bWriteFile Then WriteUnconvertedFile sOut
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure AuditUnconvertedVerseParagraphs of Module basVerseStructureAudit"
    Resume PROC_EXIT
End Sub

' --------------------------------------------------------------------------
' WriteUnconvertedFile - write the report to rpt\UnconvertedVerseAudit.txt
' --------------------------------------------------------------------------
Private Sub WriteUnconvertedFile(ByVal sContent As String)
    Dim oFSO As Object
    Dim oStream As Object
    Dim sPath As String
    sPath = ActiveDocument.Path & "\rpt\UnconvertedVerseAudit.txt"
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.CreateTextFile(sPath, True, False)   ' ASCII
    oStream.Write sContent
    oStream.Close
End Sub

' ==========================================================================
' AuditHeaderFooterStyles
' ==========================================================================
' Report the paragraph style of every header/footer story in the document
' and flag any not on the approved "TheHeaders" / "TheFooters" styles.
' Read-only; produces a report file plus an Immediate-window summary.
'
' Project rule (the structural contract):
'   Header paragraphs use the "TheHeaders" paragraph style and nothing
'   else; footer paragraphs use "TheFooters" and nothing else. Word auto-
'   applies the built-in "Header"/"Footer" styles to every header/footer
'   story it creates, so a story whose text was never restyled shows up
'   here as a violation.
'
' Enumeration (changed 2026-06-01):
'   Walks ActiveDocument.StoryRanges + NextStoryRange filtered to the six
'   header/footer story types (wd*Header/FooterStory 6-11) - the SAME basis
'   as CountAuditStyles_ToFile / "Style Usage Distribution.txt". This makes
'   the audit a superset of the distribution: it now ENUMERATES ORPHANED
'   first-page / even-page stories (content that persists after the
'   PageSetup toggle was switched off) that the earlier Sections + .Exists
'   walk was blind to. NextStoryRange yields each distinct owned story once
'   (linked sections share and are not re-counted), so the totals reconcile
'   with the distribution's built-in Header/Footer counts exactly.
'
'   Each row is classified:
'     ACTIVE   - currently rendered (Primary always; FirstPage/EvenPages
'                when the owning section's PageSetup toggle is on).
'     ORPHANED - FirstPage/EvenPages content present but toggle off.
'   Section index and ACTIVE/ORPHANED are best-effort (guarded); the
'   violation flag is the authoritative signal.
'
' Output: rpt\HeaderFooterStyleAudit.txt (when bWriteFile = True) plus
' Immediate-window summary.
'
' Usage:
'   AuditHeaderFooterStyles            ' default writes file
'   AuditHeaderFooterStyles False      ' Immediate only, no file
' ==========================================================================
Public Sub AuditHeaderFooterStyles(Optional ByVal bWriteFile As Boolean = True)
    On Error GoTo PROC_ERR
    Dim t As Double
    StartTimer "AuditHeaderFooterStyles", t

    Const APPROVED_HEADER As String = "TheHeaders"
    Const APPROVED_FOOTER As String = "TheFooters"

    Dim oDoc As Object
    Set oDoc = ActiveDocument

    Dim sOut As String
    Const NL As String = vbCrLf
    sOut = "---- AuditHeaderFooterStyles: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ----" & NL & NL
    sOut = sOut & "Rule: header paragraphs use '" & APPROVED_HEADER & "', footer paragraphs use '" & APPROVED_FOOTER & "'." & NL
    sOut = sOut & "Enumerated via StoryRanges/NextStoryRange (same basis as the Style Usage" & NL
    sOut = sOut & "Distribution report), so ORPHANED first-page/even-page stories are included." & NL
    sOut = sOut & "Stories shared via Link to Previous appear once, under their owning section." & NL & NL

    Dim hdrParas As Long, hdrViol As Long
    Dim ftrParas As Long, ftrViol As Long

    Dim rng As Object
    For Each rng In oDoc.StoryRanges
        If IsHeaderFooterStory(rng.StoryType) Then
            WalkHFChain sOut, rng, APPROVED_HEADER, APPROVED_FOOTER, _
                        hdrParas, hdrViol, ftrParas, ftrViol, True
        End If
    Next rng

    sOut = sOut & NL & "---- Summary ----" & NL
    sOut = sOut & "Header paragraphs: " & hdrParas & "   violations (style <> " & APPROVED_HEADER & "): " & hdrViol & NL
    sOut = sOut & "Footer paragraphs: " & ftrParas & "   violations (style <> " & APPROVED_FOOTER & "): " & ftrViol & NL
    sOut = sOut & "TOTAL violations: " & (hdrViol + ftrViol) & NL

    Debug.Print sOut
    If bWriteFile Then WriteHeaderFooterStyleFile sOut

    EndTimer "AuditHeaderFooterStyles", t
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure AuditHeaderFooterStyles of Module basVerseStructureAudit"
    Resume PROC_EXIT
End Sub

' --------------------------------------------------------------------------
' WalkHFChain - walk one header/footer story chain (the head range from
' StoryRanges, then NextStoryRange to the end) and fold every paragraph into
' the header/footer running totals. When bReport is True, also append a
' classified row per paragraph to sOut. Shared by AuditHeaderFooterStyles
' (report) and GetHeaderFooterStyleTotals (count-only) so the two can never
' diverge.
' --------------------------------------------------------------------------
Private Sub WalkHFChain(ByRef sOut As String, ByVal headRng As Object, _
                        ByVal approvedHeader As String, ByVal approvedFooter As String, _
                        ByRef hdrParas As Long, ByRef hdrViol As Long, _
                        ByRef ftrParas As Long, ByRef ftrViol As Long, _
                        ByVal bReport As Boolean)
    Const NL As String = vbCrLf

    Dim rng As Object
    Set rng = headRng
    Do
        Dim sType As Long
        sType = rng.StoryType
        Dim isHeader As Boolean
        isHeader = HFIsHeader(sType)
        Dim approvedStyle As String
        If isHeader Then approvedStyle = approvedHeader Else approvedStyle = approvedFooter

        Dim rowPrefix As String
        If bReport Then
            rowPrefix = "Sec " & HFSectionIndex(rng) & " | " & _
                        IIf(isHeader, "Header", "Footer") & "/" & HFPositionLabel(sType) & _
                        " [" & HFStateLabel(rng, sType) & "]"
        End If

        Dim oPara As Object
        Dim StyleName As String
        Dim excerpt As String
        Dim flag As String
        For Each oPara In rng.Paragraphs
            StyleName = oPara.style.NameLocal
            flag = ""
            If isHeader Then
                hdrParas = hdrParas + 1
                If StyleName <> approvedStyle Then
                    hdrViol = hdrViol + 1
                    flag = "  *** VIOLATION"
                End If
            Else
                ftrParas = ftrParas + 1
                If StyleName <> approvedStyle Then
                    ftrViol = ftrViol + 1
                    flag = "  *** VIOLATION"
                End If
            End If

            If bReport Then
                sOut = sOut & rowPrefix & " | style=" & StyleName & flag & NL
                excerpt = Left$(Replace(oPara.Range.Text, vbCr, ""), 60)
                If Len(excerpt) > 0 Then sOut = sOut & "    text: """ & excerpt & """" & NL
            End If
        Next oPara

        Set rng = rng.NextStoryRange
    Loop Until rng Is Nothing
End Sub

' --------------------------------------------------------------------------
' Header/footer story-type helpers (WdStoryType is late-bound here):
'   6 EvenPagesHeader   7 PrimaryHeader   10 FirstPageHeader
'   8 EvenPagesFooter   9 PrimaryFooter   11 FirstPageFooter
' --------------------------------------------------------------------------
Private Function IsHeaderFooterStory(ByVal sType As Long) As Boolean
    IsHeaderFooterStory = (sType >= 6 And sType <= 11)
End Function

Private Function HFIsHeader(ByVal sType As Long) As Boolean
    HFIsHeader = (sType = 6 Or sType = 7 Or sType = 10)
End Function

Private Function HFPositionLabel(ByVal sType As Long) As String
    Select Case sType
        Case 7, 9:   HFPositionLabel = "Primary"
        Case 10, 11: HFPositionLabel = "FirstPage"
        Case 6, 8:   HFPositionLabel = "EvenPages"
        Case Else:   HFPositionLabel = "Story" & sType
    End Select
End Function

Private Function HFSectionIndex(ByVal rng As Object) As String
    Dim s As String
    s = "?"
    On Error Resume Next
    s = CStr(rng.Sections(1).Index)
    On Error GoTo 0
    HFSectionIndex = s
End Function

' ACTIVE = currently rendered; ORPHANED = content present but PageSetup
' toggle off. Primary is always rendered; FirstPage/EvenPages depend on the
' owning section's toggle. Best-effort - returns "?" if the lookup fails.
Private Function HFStateLabel(ByVal rng As Object, ByVal sType As Long) As String
    Select Case sType
        Case 7, 9
            HFStateLabel = "ACTIVE"
        Case 10, 11        ' FirstPage
            Dim onFP As Boolean
            onFP = False
            On Error Resume Next
            onFP = rng.Sections(1).PageSetup.DifferentFirstPageHeaderFooter
            On Error GoTo 0
            HFStateLabel = IIf(onFP, "ACTIVE", "ORPHANED")
        Case 6, 8          ' EvenPages
            Dim onEv As Boolean
            onEv = False
            On Error Resume Next
            onEv = rng.Sections(1).PageSetup.OddAndEvenPagesHeaderFooter
            On Error GoTo 0
            HFStateLabel = IIf(onEv, "ACTIVE", "ORPHANED")
        Case Else
            HFStateLabel = "?"
    End Select
End Function

' ==========================================================================
' GetHeaderFooterStyleTotals
' ==========================================================================
' Count-only single source of truth for test slot 84. Same StoryRanges
' enumeration as AuditHeaderFooterStyles (via WalkHFChain, bReport:=False),
' so the gate number reconciles with both the report and the Style Usage
' Distribution. Returns paragraph totals and the count NOT on the approved
' style for headers and footers separately - orphaned stories included.
' ==========================================================================
Public Sub GetHeaderFooterStyleTotals(ByRef hdrParas As Long, ByRef hdrViolations As Long, _
                                      ByRef ftrParas As Long, ByRef ftrViolations As Long)
    Const APPROVED_HEADER As String = "TheHeaders"
    Const APPROVED_FOOTER As String = "TheFooters"

    hdrParas = 0: hdrViolations = 0
    ftrParas = 0: ftrViolations = 0

    Dim oDoc As Object
    Set oDoc = ActiveDocument

    Dim sink As String      ' unused when bReport:=False
    Dim rng As Object
    For Each rng In oDoc.StoryRanges
        If IsHeaderFooterStory(rng.StoryType) Then
            WalkHFChain sink, rng, APPROVED_HEADER, APPROVED_FOOTER, _
                        hdrParas, hdrViolations, ftrParas, ftrViolations, False
        End If
    Next rng
End Sub

' --------------------------------------------------------------------------
' WriteHeaderFooterStyleFile - write the report to
' rpt\HeaderFooterStyleAudit.txt
' --------------------------------------------------------------------------
Private Sub WriteHeaderFooterStyleFile(ByVal sContent As String)
    Dim oFSO As Object
    Dim oStream As Object
    Dim sPath As String
    sPath = ActiveDocument.Path & "\rpt\HeaderFooterStyleAudit.txt"
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.CreateTextFile(sPath, True, False)   ' ASCII
    oStream.Write sContent
    oStream.Close
End Sub
