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

' High-resolution timer for the GetMarkerTotals measurement (VBA's Timer is ~10 ms,
' far too coarse for single COM calls). The Currency type carries the 64-bit counter
' scaled by 10,000 on both the counter and the frequency, so their ratio is seconds.
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long

' Power state for the slowdown hunt (2026-10-18: an unattended full run on BATTERY went into Modern
' Standby three times, one stall 9,140 s, and ran 3x slower than on AC). SystemStatusFlag = 1 is
' Windows battery saver. Read-only.
Private Type SYSTEM_POWER_STATUS
    ACLineStatus As Byte
    BatteryFlag As Byte
    BatteryLifePercent As Byte
    SystemStatusFlag As Byte
    BatteryLifeTime As Long
    BatteryFullLifeTime As Long
End Type
Private Declare PtrSafe Function GetSystemPowerStatus Lib "kernel32" (ByRef lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long

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
    Dim issueCount As Long
    issueCount = AuditVerseMarkerStructureIssueCount(bWriteFile)
End Sub

' ==========================================================================
' AuditVerseMarkerStructureIssueCount (RUN_THE_TESTS(94))
' ==========================================================================
' The audit above as a function returning the number of structural issues
' (0 = pass), so it runs as an automated test and the release gate can read
' the Result from rpt\VerseStructureAudit.txt instead of relying on a manual
' step (2026-10-16).
' ==========================================================================
Public Function AuditVerseMarkerStructureIssueCount(Optional ByVal bWriteFile As Boolean = True) As Long
    Dim t As Double
    StartTimer "AuditVerseMarkerStructure", t

    Dim oDoc As Object
    Set oDoc = ActiveDocument

    ' Canonical 66-book reference data sourced directly from the citation
    ' class (single source of truth). books(BookID) = Array(BookID, name, chapters).
    Dim books As Object
    Set books = aeBibleCitationClass.GetCanonicalBookTable

    ' Collect Heading 1 / Heading 2 paragraph starts and Verse marker / Chapter Verse marker run starts
    ' with ONE whole-document Find each (was: a Paragraphs walk per level plus two Finds per chapter,
    ' about 1189 x 2 Finds; 2026-10-19 speed-up). Per-book / per-chapter counts are then pure arithmetic
    ' on the sorted start lists (see CountStartsInRange), with the same range rule as the old per-chapter Finds.
    Dim h1Names() As String
    Dim h1Starts() As Long
    Dim nH1 As Long
    CollectParaStyleStarts oDoc, "Heading 1", h1Starts, h1Names, nH1

    Dim h2All() As Long
    Dim h2Names() As String
    Dim nH2All As Long
    CollectParaStyleStarts oDoc, "Heading 2", h2All, h2Names, nH2All

    Dim vmAll() As Long
    Dim nVM As Long
    CollectCharStyleStarts oDoc, "Verse marker", vmAll, nVM

    Dim cvmAll() As Long
    Dim nCVM As Long
    CollectCharStyleStarts oDoc, "Chapter Verse marker", cvmAll, nCVM

    ' Build report
    Dim sOut As String
    Const NL As String = vbCrLf
    Dim totalExpected As Long, totalFound As Long
    Dim issuesCount As Long
    Dim issues As String
    Dim seenBookID(1 To 66) As Boolean

    sOut = "---- AuditVerseMarkerStructure: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ----" & NL & NL
    sOut = sOut & nH1 & " Heading 1 paragraphs found in " & oDoc.Sections.Count & " sections." & NL & NL
    Dim sHeader As String
    sHeader = sOut

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

            AuditOneBook h1Starts(i), bookEndPos, CStr(books(BookID)(1)), _
                          expectedChapters, foundChapters, chapterReport, _
                          bookIssues, bookIssueDetail, totalExpected, totalFound, _
                          h2All, nH2All, vmAll, nVM, cvmAll, nCVM

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

    Dim sSummary As String
    sSummary = "SUMMARY: " & totalFound & " / " & totalExpected & _
               " verses found, " & issuesCount & " structural issue(s)." & NL
    sOut = sOut & sSummary

    ' Brief Immediate-window output (2026-10-17): with the file written, the window gets only the header,
    ' any MISSING BOOKS / ISSUES FOUND lines and the SUMMARY; the per-chapter detail is in the file.
    ' Without the file (bWriteFile = False) the full detail still goes to the Immediate window.
    If bWriteFile Then
        Dim sBrief As String
        sBrief = sHeader
        If Len(missing) > 0 Then sBrief = sBrief & "MISSING BOOKS:" & NL & missing & NL
        If Len(issues) > 0 Then sBrief = sBrief & "ISSUES FOUND:" & NL & issues & NL
        Debug.Print sBrief & sSummary & "Per-chapter detail: rpt\VerseStructureAudit.txt"
        WriteAuditFile sOut
    Else
        Debug.Print sOut
    End If

    EndTimer "AuditVerseMarkerStructure", t
    AuditVerseMarkerStructureIssueCount = issuesCount
End Function

' --------------------------------------------------------------------------
' AuditOneBook - chapter and verse counts for a single book range
' --------------------------------------------------------------------------
Private Sub AuditOneBook(ByVal bookStart As Long, ByVal bookEnd As Long, _
                         ByVal bookName As String, _
                         ByVal expectedChapters As Long, _
                         ByRef foundChapters As Long, _
                         ByRef chapterReport As String, _
                         ByRef bookIssues As Long, _
                         ByRef bookIssueDetail As String, _
                         ByRef totalExpected As Long, _
                         ByRef totalFound As Long, _
                         ByRef h2All() As Long, ByVal nH2All As Long, _
                         ByRef vmAll() As Long, ByVal nVM As Long, _
                         ByRef cvmAll() As Long, ByVal nCVM As Long)
    Const NL As String = vbCrLf

    ' This book's Heading 2 starts are the slice of the sorted document-wide list inside [bookStart, bookEnd).
    Dim h2First As Long, nH2 As Long
    h2First = LowerBound(h2All, nH2All, bookStart)
    nH2 = LowerBound(h2All, nH2All, bookEnd) - h2First
    Dim h2Starts() As Long
    ReDim h2Starts(1 To IIf(nH2 > 0, nH2, 1))
    Dim hk As Long
    For hk = 1 To nH2
        h2Starts(hk) = h2All(h2First + hk - 1)
    Next hk

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

        foundVerses = CountStartsInRange(vmAll, nVM, h2Starts(chIdx), chEnd)
        foundCVMs = CountStartsInRange(cvmAll, nCVM, h2Starts(chIdx), chEnd)
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
' Why this shape: the former per-chapter Find pattern (AuditVerseMarkerStructure, before 2026-10-19)
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

    ' Phase 1 measurement (review section 28 task 1) - the counting logic below is
    ' UNCHANGED; the Q-stamps only time each COM-call group with the high-resolution
    ' counter and a lap line goes to rpt\MarkerTotalsLaps.txt every MARKER_LAP_EVERY
    ' VerseText paragraphs. See MarkerLapLog for how to read it.
    Const MARKER_LAP_EVERY As Long = 250
    Dim qpf As Currency, qMark As Currency, qNow As Currency, qWalk As Currency
    Dim tPara As Double, tFirst As Double, tCount As Double, tScan As Double  ' this lap
    Dim sPara As Double, sFirst As Double, sCount As Double, sScan As Double  ' whole walk
    Dim nParas As Long, nVerse As Long, nIter As Long, nIterLap As Long
    Dim lapSecs As Double, walkSecs As Double
    Dim nParaTotal As Long

    QueryPerformanceFrequency qpf
    MarkerLapLog MarkerLapHeader(nParaTotal), True
    ' Experiment (2026-10-17): in the full suite the style-read canary is ~0.8 ms until Test 74,
    ' ~4.6 ms after Tests 74 and 80, and back to ~0.8 ms after the Find-based Tests 79 and 88.
    ' A cheap whole-document Find before the walk tests whether a Find resets the slow state;
    ' the log shows the canary before and after. Set False to run the walk unprimed.
    Const MARKER_PRIME_WITH_FIND As Boolean = True
    If MARKER_PRIME_WITH_FIND Then MarkerPrime
    QueryPerformanceCounter qWalk
    qMark = qWalk

    For Each oPara In ActiveDocument.Paragraphs
        nParas = nParas + 1
        ' bucket "para": enumerator advance + the paragraph-style read
        If oPara.style.NameLocal = "VerseText" Then
            QueryPerformanceCounter qNow: tPara = tPara + (qNow - qMark) / qpf: qMark = qNow
            nVerse = nVerse + 1
            firstCharStyle = oPara.Range.Characters(1).style.NameLocal
            If firstCharStyle = "Chapter Verse marker" Then
                cvmTotal = cvmTotal + 1
            End If
            QueryPerformanceCounter qNow: tFirst = tFirst + (qNow - qMark) / qpf: qMark = qNow
            numChars = oPara.Range.Characters.Count
            QueryPerformanceCounter qNow: tCount = tCount + (qNow - qMark) / qpf: qMark = qNow
            maxScan = 12
            If numChars < maxScan Then maxScan = numChars
            For j = 1 To maxScan
                nIter = nIter + 1
                nIterLap = nIterLap + 1
                charStyle = oPara.Range.Characters(j).style.NameLocal
                If charStyle = "Verse marker" Then
                    vmTotal = vmTotal + 1
                    Exit For
                End If
            Next j
            QueryPerformanceCounter qNow: tScan = tScan + (qNow - qMark) / qpf: qMark = qNow

            If nVerse Mod MARKER_LAP_EVERY = 0 Then
                lapSecs = tPara + tFirst + tCount + tScan
                MarkerLapLog "VT " & PadLeft(CStr(nVerse), 6) & " (para " & PadLeft(CStr(nParas), 6) & _
                    ", pos " & PadLeft(CStr(oPara.Range.Start), 8) & ") | lap " & Format$(lapSecs, "0.00") & " s = " & _
                    "para " & Format$(tPara, "0.00") & " + first " & Format$(tFirst, "0.00") & _
                    " + Count " & Format$(tCount, "0.00") & " + scan " & Format$(tScan, "0.00") & _
                    " | scan chars/VT " & Format$(nIterLap / MARKER_LAP_EVERY, "0.0") & _
                    " | " & MemSnapshot() & " | " & PowerStateLabel()
                sPara = sPara + tPara: sFirst = sFirst + tFirst: sCount = sCount + tCount: sScan = sScan + tScan
                tPara = 0: tFirst = 0: tCount = 0: tScan = 0: nIterLap = 0
                QueryPerformanceCounter qMark   ' the log write itself is not charged to any bucket
            ElseIf nVerse = 1 Or nVerse = 10 Or nVerse = 100 Then
                ' Early checkpoints: if the walk stalls inside the first lap, the last of these
                ' lines says how far it got and how long it took (an in-suite stall has no lap line).
                QueryPerformanceCounter qNow
                MarkerLapLog "      reached VerseText #" & nVerse & " (para " & nParas & ") after " & _
                    Format$((qNow - qWalk) / qpf, "0.000") & " s"
                QueryPerformanceCounter qMark
            End If
        End If
    Next oPara

    sPara = sPara + tPara: sFirst = sFirst + tFirst: sCount = sCount + tCount: sScan = sScan + tScan
    QueryPerformanceCounter qNow
    walkSecs = (qNow - qWalk) / qpf
    MarkerLapLog "DONE | paragraphs " & nParas & " (Paragraphs.Count said " & nParaTotal & ") | VerseText " & nVerse & _
        " | VM " & vmTotal & " | CVM " & cvmTotal & " | walk " & Format$(walkSecs, "0.0") & " s = para " & _
        Format$(sPara, "0.0") & " + first " & Format$(sFirst, "0.0") & " + Count " & Format$(sCount, "0.0") & _
        " + scan " & Format$(sScan, "0.0") & " (rest = log writes) | scan chars total " & nIter & " | " & MemSnapshot()

    m_cachedDocName = currentDoc
    m_cachedVMTotal = vmTotal
    m_cachedCVMTotal = cvmTotal
    m_cacheValid = True
End Sub

' ==========================================================================
' MarkerLapHeader / MarkerLapLog - measurement for the Tests 82/83 slowness
' ==========================================================================
' rpt\MarkerTotalsLaps.txt, one line per MARKER_LAP_EVERY (250) VerseText paragraphs, written and
' closed at once so it survives a hang or a crash. How to read it:
'   - Compare the lap seconds down the file. Flat = constant cost per call;
'     rising with the paragraph/pos columns = the cost grows with position or
'     with accumulated state inside Word.
'   - The four buckets say WHICH call is slow: para = enumerator advance plus
'     the paragraph-style read; first = Characters(1) style (CVM check);
'     Count = Characters.Count; scan = the up-to-12 Characters(j) style reads.
'   - scan chars/VT is how far into the paragraph the Verse marker sits; if it is
'     near 12 the scan is paying for the full 12 iterations on every paragraph.
'   - The header line shows the Word state (view, ScreenUpdating, pagination)
'     and the cost of Paragraphs.Count; compare a standalone run with a run
'     after Tests 1-81 to see which of those differs.
' Monitor only: file and WMI errors are swallowed.
' ==========================================================================
' PowerStateLabel - AC or battery, charge, battery saver (see the SYSTEM_POWER_STATUS Note above).
Public Function PowerStateLabel() As String
    On Error GoTo PROC_ERR
    Dim sps As SYSTEM_POWER_STATUS
    Dim s As String
    If GetSystemPowerStatus(sps) = 0 Then GoTo PROC_ERR
    Select Case sps.ACLineStatus
        Case 1: s = "AC"
        Case 0: s = "BATTERY"
        Case Else: s = "AC unknown"
    End Select
    If sps.BatteryLifePercent <> 255 Then s = s & " " & sps.BatteryLifePercent & "%"
    If sps.SystemStatusFlag = 1 Then s = s & " battery-saver ON"
    PowerStateLabel = "power " & s
    Exit Function
PROC_ERR:
    PowerStateLabel = "power n/a"
End Function


' --------------------------------------------------------------------------
' StyleReadCanary - fixed micro-benchmark for the "style reads got 8-10x slower
' after Tests 1-81" finding (in-suite Test 82 walk 1256 s vs 157 s standalone).
' Called after every test by aeBibleClass.LogTestLap: it times 100 repeats of the
' exact call GetMarkerTotals makes (Range.Characters(1).Style.NameLocal) on one
' fixed character in the verse text, and 100 repeats of Characters.Count as a
' control. The first test after which the style-read figure jumps is the one
' that leaves Word in the slow state. Read-only, ~0.1-1 s. Returns "canary n/a" on error.
' --------------------------------------------------------------------------
Public Function StyleReadCanary() As String
    On Error GoTo PROC_ERR
    Const REPEATS As Long = 100
    Dim qpf As Currency, q0 As Currency, q1 As Currency, q2 As Currency
    Dim rng As Object
    Dim i As Long
    Dim s As String

    Set rng = ActiveDocument.Range(150000, 150001)
    QueryPerformanceFrequency qpf
    QueryPerformanceCounter q0
    For i = 1 To REPEATS
        s = rng.Characters(1).style.NameLocal
    Next i
    QueryPerformanceCounter q1
    For i = 1 To REPEATS
        s = rng.Characters.Count
    Next i
    QueryPerformanceCounter q2
    StyleReadCanary = "canary style " & Format$((q1 - q0) / qpf / REPEATS * 1000, "0.00") & " ms/read" & _
        ", Count " & Format$((q2 - q1) / qpf / REPEATS * 1000, "0.00") & " ms/read"
    Exit Function
PROC_ERR:
    StyleReadCanary = "canary n/a"
End Function

Private Sub MarkerPrime()
    On Error Resume Next
    Dim qpf As Currency, q0 As Currency, q1 As Currency
    QueryPerformanceFrequency qpf
    MarkerLapLog "      canary before prime: " & StyleReadCanary()
    QueryPerformanceCounter q0
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Text = "zzqxjvzz"
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .Execute
    End With
    QueryPerformanceCounter q1
    MarkerLapLog "      prime Find took " & Format$((q1 - q0) / qpf, "0.0") & " s; canary after: " & StyleReadCanary()
End Sub

Private Function MarkerLapHeader(ByRef paraCount As Long) As String
    On Error Resume Next
    Dim qpf As Currency, q0 As Currency, q1 As Currency
    Dim s As String
    QueryPerformanceFrequency qpf
    QueryPerformanceCounter q0
    paraCount = ActiveDocument.Paragraphs.Count
    QueryPerformanceCounter q1
    s = "MarkerTotalsLaps " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & ActiveDocument.Name & _
        " | Paragraphs.Count " & paraCount & " took " & Format$((q1 - q0) / qpf, "0.000") & " s" & _
        " | view type " & ActiveWindow.View.Type & " | ScreenUpdating " & Application.ScreenUpdating & _
        " | Pagination " & Application.Options.Pagination & " | ShowAll " & ActiveWindow.View.ShowAll & _
        " | TrackRevisions " & ActiveDocument.TrackRevisions & " | documents open " & Documents.Count & _
        vbCrLf & "start: " & MemSnapshot()
    MarkerLapHeader = s
End Function

Private Sub MarkerLapLog(ByVal sLine As String, Optional ByVal bNewFile As Boolean = False)
    On Error Resume Next
    Dim fileNum As Integer
    fileNum = FreeFile
    If bNewFile Then
        Open ActiveDocument.Path & "\rpt\MarkerTotalsLaps.txt" For Output As #fileNum
    Else
        Open ActiveDocument.Path & "\rpt\MarkerTotalsLaps.txt" For Append As #fileNum
    End If
    ' Clock time on every line so a stall can be matched to the Windows event log (Modern Standby).
    Print #fileNum, Format$(Now, "hh:nn:ss") & " " & sLine
    Close #fileNum
End Sub

' ==========================================================================
' FindMarkerStyleOutsideVerseText
' ==========================================================================
' Read-only diagnostic, complementary to GetMarkerTotals. GetMarkerTotals
' only inspects paragraphs styled "VerseText" - it has no way to see a
' "Chapter Verse marker"/"Verse marker" character-style run sitting inside
' a paragraph styled something else. AuditVerseMarkerStructure's per-chapter
' Find-based counts (CollectCharStyleStarts) search by
' character style only, with no paragraph-style filter, so such a paragraph
' would be counted there but invisible to ExportDocmVersesToRWBFormat
' (which, like GetMarkerTotals, only ever visits "VerseText" paragraphs).
' This routine tests exactly that gap directly.
'
' Deliberately cheap: paragraph-level iteration (proven fast, seconds, by
' GetMarkerTotals/ExportDocmVersesToRWBFormat) plus the same first-char/
' first-12-chars character-style checks GetMarkerTotals already uses -
' but restricted to the small minority of paragraphs NOT styled
' "VerseText", so no whole-document character-level Find loop is needed.
'
' Output: rpt\MarkerStyleOutsideVerseText.txt plus Immediate window summary.
' hitCount (ByRef, same convention as GetMarkerTotals's vmTotal/cvmTotal)
' lets aeBibleClass wire this into a test slot expecting 0.
' ==========================================================================
Public Sub FindMarkerStyleOutsideVerseText(ByRef hitCount As Long, _
                                            Optional ByVal bWriteFile As Boolean = True, _
                                            Optional ByRef sHitStarts As String = "")
    Dim t As Double
    StartTimer "FindMarkerStyleOutsideVerseText", t

    Dim oPara As Object
    Dim numChars As Long
    Dim maxScan As Long
    Dim j As Long
    Dim charStyle As String
    Dim firstCharStyle As String
    Dim paraStyle As String
    Dim scannedCount As Long
    hitCount = 0
    sHitStarts = ""

    Dim sOut As String
    Const NL As String = vbCrLf
    sOut = "---- FindMarkerStyleOutsideVerseText: " & _
           Format(Now, "yyyy-mm-dd hh:nn:ss") & " ----" & NL & NL

    ' Bucket timing (2026-10-19, measurement only - the hit logic is unchanged): QueryPerformanceCounter
    ' stamps split the walk into para (enumerator advance + paragraph-style read, all paragraphs),
    ' first (Characters(1) style), Count (Characters.Count), scan (Characters(j) style reads) and
    ' hit (building the hit text). A lap line every MOV_LAP_EVERY paragraphs shows whether the cost
    ' per paragraph grows down the document. Written to rpt\MarkerOutsideLaps.txt after the walk.
    Const MOV_LAP_EVERY As Long = 5000
    Dim qpf As Currency, qMark As Currency, qNow As Currency, qStart As Currency
    Dim tPara As Double, tFirst As Double, tCount As Double, tScan As Double, tHit As Double
    Dim lPara As Double, lFirst As Double, lCount As Double, lScan As Double, lHit As Double
    Dim nParas As Long, nChars As Long, nCharsLap As Long, nFast As Long
    Dim sUniform As String, oCS As Object
    Dim sLaps As String
    QueryPerformanceFrequency qpf
    QueryPerformanceCounter qStart
    qMark = qStart

    For Each oPara In ActiveDocument.Paragraphs
        nParas = nParas + 1
        paraStyle = oPara.style.NameLocal
        QueryPerformanceCounter qNow: lPara = lPara + (qNow - qMark) / qpf: qMark = qNow
        If paraStyle <> "VerseText" Then
            scannedCount = scannedCount + 1
            Dim isHit As Boolean
            isHit = False

            ' Merged fast path (2026-10-19): ONE style read over the leading up-to-12 characters decides
            ' most paragraphs. CharacterStyle is the style when that range has a single character style and
            ' Nothing when it is mixed. Single style "Chapter Verse marker" or "Verse marker" = hit; any other
            ' single style = no marker among the 12 and the first character is not a CVM = no hit, and both
            ' the Characters(1) read and the per-character reads are skipped. Mixed (or an error) falls back
            ' to the exact original logic: Characters(1) is a CVM -> hit, else scan Characters(1..12) for a VM.
            numChars = oPara.Range.Characters.Count
            QueryPerformanceCounter qNow: lCount = lCount + (qNow - qMark) / qpf: qMark = qNow
            maxScan = 12
            If numChars < maxScan Then maxScan = numChars
            sUniform = ""
            On Error Resume Next
            Set oCS = Nothing
            Set oCS = ActiveDocument.Range(oPara.Range.Start, oPara.Range.Start + maxScan).CharacterStyle
            sUniform = oCS.NameLocal
            On Error GoTo 0
            QueryPerformanceCounter qNow: lScan = lScan + (qNow - qMark) / qpf: qMark = qNow

            If Len(sUniform) > 0 Then
                nFast = nFast + 1
                If sUniform = "Chapter Verse marker" Or sUniform = "Verse marker" Then
                    isHit = True
                    firstCharStyle = sUniform
                End If
            Else
                firstCharStyle = oPara.Range.Characters(1).style.NameLocal
                If firstCharStyle = "Chapter Verse marker" Then isHit = True
                QueryPerformanceCounter qNow: lFirst = lFirst + (qNow - qMark) / qpf: qMark = qNow
                If Not isHit Then
                    For j = 1 To maxScan
                        nChars = nChars + 1
                        nCharsLap = nCharsLap + 1
                        charStyle = oPara.Range.Characters(j).style.NameLocal
                        If charStyle = "Verse marker" Then
                            isHit = True
                            Exit For
                        End If
                    Next j
                    QueryPerformanceCounter qNow: lScan = lScan + (qNow - qMark) / qpf: qMark = qNow
                End If
            End If

            If isHit Then
                hitCount = hitCount + 1
                sHitStarts = sHitStarts & oPara.Range.Start & ","
                sOut = sOut & "Hit #" & hitCount & " | ParaStart=" & oPara.Range.Start & _
                       " | Style=" & paraStyle & " | first-char-style=" & firstCharStyle & NL
                sOut = sOut & "  Excerpt: """ & Left$(Replace(oPara.Range.Text, vbCr, ""), 80) & """" & NL & NL
                QueryPerformanceCounter qNow: lHit = lHit + (qNow - qMark) / qpf: qMark = qNow
            End If
        End If

        If nParas Mod MOV_LAP_EVERY = 0 Then
            sLaps = sLaps & "para " & PadLeft(CStr(nParas), 6) & " (non-VT scanned " & PadLeft(CStr(scannedCount), 5) & _
                ") | lap " & Format$(lPara + lFirst + lCount + lScan + lHit, "0.00") & " s = para " & Format$(lPara, "0.00") & _
                " + first " & Format$(lFirst, "0.00") & " + Count " & Format$(lCount, "0.00") & _
                " + scan " & Format$(lScan, "0.00") & " + hit " & Format$(lHit, "0.00") & _
                " | scan chars " & nCharsLap & vbCrLf
            tPara = tPara + lPara: tFirst = tFirst + lFirst: tCount = tCount + lCount: tScan = tScan + lScan: tHit = tHit + lHit
            lPara = 0: lFirst = 0: lCount = 0: lScan = 0: lHit = 0: nCharsLap = 0
            QueryPerformanceCounter qMark
        End If
    Next oPara

    tPara = tPara + lPara: tFirst = tFirst + lFirst: tCount = tCount + lCount: tScan = tScan + lScan: tHit = tHit + lHit
    QueryPerformanceCounter qNow
    If bWriteFile Then MarkerOutsideLapsWrite "FindMarkerStyleOutsideVerseText " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & ActiveDocument.Name & _
        " | " & PowerStateLabel() & " | " & StyleReadCanary() & vbCrLf & sLaps & _
        "DONE | paragraphs " & nParas & " | non-VT scanned " & scannedCount & " | hits " & hitCount & _
        " | walk " & Format$((qNow - qStart) / qpf, "0.0") & " s = para " & Format$(tPara, "0.0") & _
        " + first " & Format$(tFirst, "0.0") & " + Count " & Format$(tCount, "0.0") & _
        " + scan " & Format$(tScan, "0.0") & " + hit " & Format$(tHit, "0.0") & _
        " (rest = lap-line writes) | scan chars total " & nChars & " | fast-path paragraphs " & nFast & vbCrLf

    sOut = sOut & "---- Summary ----" & NL
    sOut = sOut & "Non-VerseText paragraphs scanned: " & scannedCount & NL
    sOut = sOut & "Hits (marker char-style found outside VerseText): " & hitCount & NL

    Debug.Print sOut
    If bWriteFile Then
        Dim oFSO As Object
        Dim oStream As Object
        Dim sPath As String
        sPath = ActiveDocument.Path & "\rpt\MarkerStyleOutsideVerseText.txt"
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        Set oStream = oFSO.CreateTextFile(sPath, True, False)
        oStream.Write sOut
        oStream.Close
    End If

    EndTimer "FindMarkerStyleOutsideVerseText", t
End Sub

' ==========================================================================
' ControlMarkerOutsideVerseText - positive control for FindMarkerStyleOutsideVerseText (Test 87)
' ==========================================================================
' The real document has 0 hits, so a clean Test 87 run cannot show that the scan still DETECTS a
' stray marker. This builds a disposable blank document (never saved, closed on exit) with known
' cases, runs FindMarkerStyleOutsideVerseText on it and compares the hit paragraph starts with the
' expected list. Run from the Immediate window: ?ControlMarkerOutsideVerseText  (0 = PASS).
' Cases (all "Normal" unless stated; CVM = Chapter Verse marker, VM = Verse marker):
'   1 CVM at char 1                  -> hit      2 VM at char 5            -> hit
'   3 VM at char 12 (boundary)       -> hit      4 VM at char 13           -> no hit
'   5 plain                          -> no hit   6 VerseText with CVM+VM   -> no hit (inside VerseText)
'   7 one-char paragraph, all VM     -> hit      8 whole paragraph VM      -> hit
'   9 CVM at char 3, no VM           -> no hit (CVM is only checked at char 1)
'  10 14 chars all VM                -> hit (uniform leading range)
'  11 15 chars all in another char style -> no hit (uniform non-marker style)
' Expected hits: paragraphs 1, 2, 3, 7, 8, 10.
' ==========================================================================
Public Function ControlMarkerOutsideVerseText() As Long
    Dim prevDoc As Object, d As Object
    Dim sTexts(1 To 11) As String
    Dim starts(1 To 11) As Long
    Dim i As Long, pos As Long
    Dim hitCount As Long, sHits As String, sExpected As String

    Set prevDoc = ActiveDocument
    Set d = Documents.Add
    On Error GoTo PROC_ERR

    d.Styles.Add "VerseText", 1                    ' wdStyleTypeParagraph
    d.Styles.Add "Verse marker", 2                 ' wdStyleTypeCharacter
    d.Styles.Add "Chapter Verse marker", 2
    d.Styles.Add "OtherCharStyle", 2

    sTexts(1) = "12 first"
    sTexts(2) = "abcd5 tail text"
    sTexts(3) = "abcdefghijk9"
    sTexts(4) = "abcdefghijkl9"
    sTexts(5) = "plain paragraph"
    sTexts(6) = "11 In the beginning"
    sTexts(7) = "3"
    sTexts(8) = "7 8 9"
    sTexts(9) = "abcdef"
    sTexts(10) = "VMVMVMVMVMVMVM"
    sTexts(11) = "OTHEROTHEROTHER"

    Dim sAll As String
    pos = 0
    For i = 1 To 11
        starts(i) = pos
        sAll = sAll & sTexts(i)
        If i < 11 Then sAll = sAll & vbCr
        pos = pos + Len(sTexts(i)) + 1
    Next i
    d.Content.Text = sAll

    d.Paragraphs(6).style = "VerseText"
    d.Range(starts(1), starts(1) + 2).style = "Chapter Verse marker"
    d.Range(starts(2) + 4, starts(2) + 5).style = "Verse marker"
    d.Range(starts(3) + 11, starts(3) + 12).style = "Verse marker"
    d.Range(starts(4) + 12, starts(4) + 13).style = "Verse marker"
    d.Range(starts(6), starts(6) + 1).style = "Chapter Verse marker"
    d.Range(starts(6) + 1, starts(6) + 2).style = "Verse marker"
    d.Range(starts(7), starts(7) + 1).style = "Verse marker"
    d.Range(starts(8), starts(8) + 5).style = "Verse marker"
    d.Range(starts(9) + 2, starts(9) + 3).style = "Chapter Verse marker"
    d.Range(starts(10), starts(10) + 14).style = "Verse marker"
    d.Range(starts(11), starts(11) + 15).style = "OtherCharStyle"

    sExpected = starts(1) & "," & starts(2) & "," & starts(3) & "," & starts(7) & "," & starts(8) & "," & starts(10) & ","

    FindMarkerStyleOutsideVerseText hitCount, False, sHits

    Dim bad As Long
    If hitCount <> 6 Or sHits <> sExpected Then bad = 1
    Debug.Print "ControlMarkerOutsideVerseText: hits " & hitCount & " (expected 6) | starts [" & sHits & "] expected [" & sExpected & "] -> " & IIf(bad = 0, "PASS", "FAIL")
    ControlMarkerOutsideVerseText = bad

PROC_EXIT:
    On Error Resume Next
    d.Close 0                                      ' wdDoNotSaveChanges
    prevDoc.Activate
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in ControlMarkerOutsideVerseText | Err: " & Err.Number & " | " & Err.Description
    ControlMarkerOutsideVerseText = -1
    Resume PROC_EXIT
End Function

' --------------------------------------------------------------------------
' MarkerOutsideLapsWrite - overwrite rpt\MarkerOutsideLaps.txt with the bucket
' timing from FindMarkerStyleOutsideVerseText (Test 87). Monitor only: errors swallowed.
' --------------------------------------------------------------------------
Private Sub MarkerOutsideLapsWrite(ByVal sContent As String)
    On Error Resume Next
    Dim oFSO As Object
    Dim oStream As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.CreateTextFile(ActiveDocument.Path & "\rpt\MarkerOutsideLaps.txt", True, False)
    oStream.Write sContent
    oStream.Close
End Sub

' --------------------------------------------------------------------------
' CollectCharStyleStarts - start positions (ascending) of every run of a
' character style in the whole document, from one Find. Replaces the old
' per-chapter CountVerseMarkers / CountChapterVerseMarkers Finds.
' --------------------------------------------------------------------------
Private Sub CollectCharStyleStarts(ByVal oDoc As Object, ByVal StyleName As String, _
                                   ByRef starts() As Long, ByRef n As Long)
    Dim cap As Long
    cap = 4096
    ReDim starts(1 To cap)
    n = 0

    Dim oRng As Object
    Set oRng = oDoc.Content
    With oRng.Find
        .ClearFormatting
        .style = oDoc.Styles(StyleName)
        .Text = ""
        .Forward = True
        .Wrap = 0     ' wdFindStop
        .Format = True
        .MatchWildcards = False
        Do While .Execute
            n = n + 1
            If n > cap Then
                cap = cap * 2
                ReDim Preserve starts(1 To cap)
            End If
            starts(n) = oRng.Start
            If n > 500000 Then Exit Do   ' safety
            oRng.Collapse 0               ' wdCollapseEnd: continue after the match
        Loop
    End With
End Sub

' --------------------------------------------------------------------------
' CollectParaStyleStarts - start position and trimmed text of every paragraph
' of a paragraph style, in document order, from one Find. A match covering
' several adjacent paragraphs is split into its paragraphs.
' --------------------------------------------------------------------------
Private Sub CollectParaStyleStarts(ByVal oDoc As Object, ByVal StyleName As String, _
                                   ByRef starts() As Long, ByRef texts() As String, _
                                   ByRef n As Long)
    Dim cap As Long
    cap = 256
    ReDim starts(1 To cap)
    ReDim texts(1 To cap)
    n = 0

    Dim oRng As Object
    Dim oPara As Object
    Dim pStart As Long
    Dim bRecord As Boolean
    Set oRng = oDoc.Content
    With oRng.Find
        .ClearFormatting
        .style = oDoc.Styles(StyleName)
        .Text = ""
        .Forward = True
        .Wrap = 0     ' wdFindStop
        .Format = True
        .MatchWildcards = False
        Do While .Execute
            For Each oPara In oRng.Paragraphs
                pStart = oPara.Range.Start
                bRecord = True
                If n > 0 Then
                    If pStart <= starts(n) Then bRecord = False   ' already recorded (overlapping match)
                End If
                If bRecord Then
                    n = n + 1
                    If n > cap Then
                        cap = cap * 2
                        ReDim Preserve starts(1 To cap)
                        ReDim Preserve texts(1 To cap)
                    End If
                    starts(n) = pStart
                    texts(n) = Trim(Replace(oPara.Range.Text, vbCr, ""))
                End If
            Next oPara
            If n > 20000 Then Exit Do   ' safety
            oRng.Collapse 0               ' wdCollapseEnd
        Loop
    End With
End Sub

' --------------------------------------------------------------------------
' LowerBound - index of the first element of the ascending list arr(1..n)
' that is >= x (n + 1 if none).
' --------------------------------------------------------------------------
Private Function LowerBound(ByRef arr() As Long, ByVal n As Long, ByVal x As Long) As Long
    Dim lo As Long, hi As Long, mid As Long
    lo = 1
    hi = n + 1
    Do While lo < hi
        mid = (lo + hi) \ 2
        If arr(mid) < x Then
            lo = mid + 1
        Else
            hi = mid
        End If
    Loop
    LowerBound = lo
End Function

' --------------------------------------------------------------------------
' CountStartsInRange - number of run starts in [startPos, endPos). Same range
' rule as the old per-chapter Find over Range(startPos, endPos).
' --------------------------------------------------------------------------
Private Function CountStartsInRange(ByRef arr() As Long, ByVal n As Long, _
                                    ByVal startPos As Long, ByVal endPos As Long) As Long
    If endPos <= startPos Then Exit Function
    CountStartsInRange = LowerBound(arr, n, endPos) - LowerBound(arr, n, startPos)
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
                ' Still Count policy flags even when suppressed (anomalies-only would emit anyway since paraStyle <> VerseText)
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

    ' End of Doc: discard remaining buffer (post-last-verse of last book)
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
' (report) and GetHeaderFooterStyleTotals (Count-only) so the two can never
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
    s = CStr(rng.Sections(1).index)
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
' Distribution. Returns paragraph totals and the Count NOT on the approved
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

' ==========================================================================
' FindSequentialVerseNumberViolations  (backs RUN_THE_TESTS test slot 91)
' ==========================================================================
' Fifth CVM/VM integrity invariant - validates marker CONTENT, not just
' presence/Count. GetMarkerTotals (slots 82/83) and
' CollectCharStyleStarts (AuditVerseMarkerStructure, Test 94) both only check whether a CVM/VM-styled
' run EXISTS once per verse paragraph - neither can see a stray character
' sharing the correct style glued onto a real marker, because a styled
' Find match swallows adjacent same-styled runs into one hit regardless of
' length. See rvw/Bug_jeremiah_37_10_stray_cvm_marker_2026-10-04.md for the
' real defect (Jeremiah 37:10 parsed as verse "910") that exposed this.
'
' This check instead tracks, per chapter, an EXPECTED NEXT VERSE NUMBER
' starting at 1, and asserts every VerseText paragraph's parsed verse
' number matches it exactly - a strictly stronger, content-aware check.
' A stray/missing digit breaks the sequence immediately and loudly instead
' of passing a presence/Count check unnoticed.
'
' Deliberately i18n-ready (rvw/Plan_cvm_content_validation_2026-10-04.md
' Addendum.4a): the expected RUN_THE_TESTS Result is 0 ("zero violations"),
' not a hardcoded per-edition verse-Count constant like Tests 82/83's
' 31102 - this runs correctly against any future docm following the same
' CVM/VM/VerseText convention, no per-language code change needed.
'
' PERFORMANCE - same proven-cheap pattern as ExportDocmVersesToRWBFormat
' (basRWBTextExport.bas): ONE Range.Text read per VerseText paragraph,
' pure string parsing (FirstNumberInText/LeadingDigits below), ZERO
' character- or word-level style lookups. That module's own header
' documents two earlier implementations in this exact area blowing past
' 2 GB of memory from per-character/word COM calls across ~35k paragraphs
' - FirstNumberInText/LeadingDigits are intentionally duplicated here
' (not called cross-module) rather than risk coupling this invariant's
' correctness to basRWBTextExport.bas's own, differently-scoped contract;
' same "copy for a self-contained module" choice this project's own
' Test-73 precedent documents (md/Adding_To_Bible_Test_Class.md).
'
' After ANY parse of a VerseText paragraph (match or mismatch), the
' expected-next-verse counter resyncs to parsedVerseNumber + 1 - so one
' bad paragraph flags itself (and, usually, exactly one follow-on "jumped
' back down" flag on the next verse) without cascading a false mismatch
' through the rest of the chapter.
' ==========================================================================
Public Sub FindSequentialVerseNumberViolations(ByRef violationCount As Long, _
                                                 ByRef firstViolationHint As String, _
                                                 Optional ByVal bWriteFile As Boolean = True)
    Dim t As Double
    StartTimer "FindSequentialVerseNumberViolations", t

    Dim canonBooks As Object
    Set canonBooks = aeBibleCitationClass.GetCanonicalBookTable()

    Dim oPara As Object
    Dim bookIndex As Long
    Dim bookName As String
    Dim chapNum As Long
    Dim expectedVerse As Long
    Dim StyleName As String
    Dim paraTxt As String
    Dim digitRun As String
    Dim chapStr As String
    Dim verseStr As String
    Dim verseNum As Long
    violationCount = 0
    firstViolationHint = ""

    Dim sOut As String
    Const NL As String = vbCrLf
    sOut = "---- FindSequentialVerseNumberViolations: " & _
           Format(Now, "yyyy-mm-dd hh:nn:ss") & " ----" & NL & NL

    bookName = ""
    chapNum = 0
    expectedVerse = 1

    For Each oPara In ActiveDocument.Paragraphs
        StyleName = oPara.style.NameLocal

        If StyleName = "Heading 1" Then
            bookIndex = bookIndex + 1
            If canonBooks.Exists(bookIndex) Then
                bookName = canonBooks.Item(bookIndex)(1)
                If bookName = "Psalms" Then
                    bookName = "Psalm"
                ElseIf bookName = "Song of Songs" Then
                    bookName = "Song of Solomon"
                End If
            Else
                bookName = "UNKNOWN_BOOK_" & bookIndex
            End If
            chapNum = 0

        ElseIf StyleName = "Heading 2" Then
            Dim n As Long
            n = FirstNumberInText(oPara.Range.Text)
            If n > 0 Then
                chapNum = n
                expectedVerse = 1
            End If

        ElseIf StyleName = "VerseText" Then
            If bookName <> "" And chapNum > 0 Then
                paraTxt = oPara.Range.Text
                digitRun = LeadingDigits(paraTxt)
                chapStr = CStr(chapNum)

                If Left$(digitRun, Len(chapStr)) = chapStr Then
                    verseStr = Mid$(digitRun, Len(chapStr) + 1)
                    If verseStr <> "" And IsNumeric(verseStr) Then
                        verseNum = CLng(verseStr)
                        If verseNum <> expectedVerse Then
                            violationCount = violationCount + 1
                            sOut = sOut & "Violation #" & violationCount & " | " & bookName & _
                                   " " & chapNum & " | expected verse " & expectedVerse & _
                                   " but parsed " & verseNum & NL
                            sOut = sOut & "  Excerpt: """ & Left$(paraTxt, 80) & """" & NL & NL
                            If firstViolationHint = "" Then firstViolationHint = bookName & " " & chapNum & _
                                ": expected verse " & expectedVerse & " but parsed " & verseNum & _
                                " - text=""" & Left$(paraTxt, 60) & """"
                        End If
                        expectedVerse = verseNum + 1
                    Else
                        violationCount = violationCount + 1
                        sOut = sOut & "Violation #" & violationCount & " | " & bookName & _
                               " " & chapNum & " | expected verse " & expectedVerse & _
                               " but verse number unparseable (digit run """ & digitRun & _
                               """)" & NL
                        sOut = sOut & "  Excerpt: """ & Left$(paraTxt, 80) & """" & NL & NL
                        If firstViolationHint = "" Then firstViolationHint = bookName & " " & chapNum & _
                            ": expected verse " & expectedVerse & _
                            " but verse number unparseable - text=""" & Left$(paraTxt, 60) & """"
                    End If
                Else
                    violationCount = violationCount + 1
                    sOut = sOut & "Violation #" & violationCount & " | " & bookName & _
                           " " & chapNum & " | expected verse " & expectedVerse & _
                           " but digit run """ & digitRun & """ does not start with chapter """ & _
                           chapStr & """" & NL
                    sOut = sOut & "  Excerpt: """ & Left$(paraTxt, 80) & """" & NL & NL
                    If firstViolationHint = "" Then firstViolationHint = bookName & " " & chapNum & _
                        ": digit run """ & digitRun & """ does not start with chapter """ & _
                        chapStr & """ - text=""" & Left$(paraTxt, 60) & """"
                End If
            End If
        End If
    Next oPara

    sOut = sOut & "---- Summary ----" & NL
    sOut = sOut & "Sequential verse-number violations: " & violationCount & NL

    Debug.Print sOut
    If bWriteFile Then
        Dim oFSO As Object
        Dim oStream As Object
        Dim sPath As String
        sPath = ActiveDocument.Path & "\rpt\SequentialVerseNumberViolations.txt"
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        Set oStream = oFSO.CreateTextFile(sPath, True, False)
        oStream.Write sOut
        oStream.Close
    End If

    EndTimer "FindSequentialVerseNumberViolations", t
End Sub

' Pure string scan (no COM) - duplicated intentionally from
' basRWBTextExport.bas's identically-named/identically-behaved helper (see
' this routine's own header comment for why) - the first run of digit
' characters anywhere in s, as a Long. Returns 0 if no digits are found.
Private Function FirstNumberInText(ByVal s As String) As Long
    Dim i As Long, ch As String, digits As String
    Dim started As Boolean
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
            started = True
        ElseIf started Then
            Exit For
        End If
    Next i
    If digits <> "" Then FirstNumberInText = CLng(digits) Else FirstNumberInText = 0
End Function

' Pure string scan (no COM) - duplicated intentionally from
' basRWBTextExport.bas's identically-named/identically-behaved helper - the
' leading run of digit characters in s, or "" if s doesn't start with a
' digit.
Private Function LeadingDigits(ByVal s As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit For
    Next i
    LeadingDigits = Left$(s, i - 1)
End Function
