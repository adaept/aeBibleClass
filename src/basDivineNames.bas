Attribute VB_Name = "basDivineNames"
Option Explicit

' Classification: docm clean-up (list-then-fix pair). Not used by runtime ribbon navigation.
'
' basDivineNames - the divine-names edits of the RWB 1.1.1 release (task RT-86; follow-ups FU-49, FU-52, FU-53; owner decisions 2026-11-13):
'   1. Exodus 6:3     "by my name Jehovah I was not known"          becomes  "by my name the Lord I was not known"
'   2. Psalm 69:6     "through me, Lord the Lord God of Hosts"      becomes  "through me, O Lord God of Hosts"
'   3. 14 verses (15 places) read "the Lord, God of Hosts" and lose the comma: "the Lord God of Hosts".
'      Isaiah 3:15, 10:23, 10:24, 22:5, 22:12, 22:14, 22:15, 28:22; Jeremiah 2:19, 46:10 (twice), 49:5, 50:25, 50:31; Amos 9:5.
'      The five verses Isaiah 1:24, 3:1, 10:16, 10:33 and 19:4 keep the comma (Hebrew ha-adon YHWH tsevaot; the census ratifies them).
'
' How a place is found: not by chapter and verse (that needs the book and chapter walk of the export), but by a short piece of the verse
' text around it (the "anchor", about 30 characters, taken from rpt\docm-verses.txt of 2026-10-04 and unique in it). A place is only
' changed when its anchor is found EXACTLY ONCE in the document and the hit is in a VerseText paragraph; anything else is reported and
' nothing is written for it. Only the characters that differ are replaced (the comma alone; the word; the words), so the run formatting
' of the rest of the phrase is kept.
'
' Entry points (Immediate window, one statement each; the document name is optional, default ActiveDocument):
'   ListDivineNameFixes                 read-only: for each of the 17 places: to do / done / PROBLEM (anchor not found once)
'   ? CountDivineNameFixesPending      the number of places still to do (0 = finished), -1 = a place is not found exactly once
'   ? FixDivineNames "<Doc name>"       writes the edits, returns the number changed, -1 = failed (DivineNamesLastError). Places that
'                                       are already done are skipped, so a second run changes 0. After the edits it checks the text
'                                       length, that "Lord, God of Hosts" stands in exactly 5 places, and that "Jehovah" and
'                                       "Lord the Lord" are gone. Run it on a COPY of the docm first.
' Reference: rvw\RWB 1.1.1 Follow-up Issues.md FU-49, FU-52, FU-53 (aeBibleAddin), release train rows RT-86 and RT-91.

Private Const VERSE_STYLE As String = "VerseText"
Private Const FIND_SAFETY_LIMIT As Long = 1000
Private Const EXPECTED_COMMA_PLACES As Long = 5

Private m_LastError As String
Private m_Ref() As String
Private m_Anchor() As String
Private m_From() As String
Private m_To() As String
Private m_Count As Long

Public Property Get DivineNamesLastError() As String
    DivineNamesLastError = m_LastError
End Property

' The same text as the property, as a function, so that the Word COM harness can read it (Run cannot call a property).
Public Function DivineNamesLastErrorText(Optional ByVal sIgnored As String = "") As String
    DivineNamesLastErrorText = m_LastError
End Function

Private Sub AddTarget(ByVal sRef As String, ByVal sAnchor As String, ByVal sFrom As String, ByVal sTo As String)
    m_Count = m_Count + 1
    ReDim Preserve m_Ref(1 To m_Count)
    ReDim Preserve m_Anchor(1 To m_Count)
    ReDim Preserve m_From(1 To m_Count)
    ReDim Preserve m_To(1 To m_Count)
    m_Ref(m_Count) = sRef
    ' #L# and #R# stand for the curly double quotes, so that this file stays plain ASCII
    m_Anchor(m_Count) = Replace(Replace(sAnchor, "#L#", ChrW$(&H201C)), "#R#", ChrW$(&H201D))
    m_From(m_Count) = sFrom
    m_To(m_Count) = sTo
End Sub

Private Sub LoadTargets()
    Const COMMA_FROM As String = "Lord, God of Hosts"
    Const COMMA_TO As String = "Lord God of Hosts"
    m_Count = 0
    AddTarget "Exodus 6:3", "but by my name Jehovah", "Jehovah", "the Lord"
    AddTarget "Psalm 69:6", "shamed through me, Lord the Lord God of Hosts.", "Lord the Lord", "O Lord"
    AddTarget "Isaiah 3:15", "poor?#R# says the Lord, God of Hosts.", COMMA_FROM, COMMA_TO
    AddTarget "Isaiah 10:23", "For the Lord, God of Hosts, will make", COMMA_FROM, COMMA_TO
    AddTarget "Isaiah 10:24", "Therefore the Lord, God of Hosts, says #L#My", COMMA_FROM, COMMA_TO
    AddTarget "Isaiah 22:5", "perplexity, from the Lord, God of Hosts, in the", COMMA_FROM, COMMA_TO
    AddTarget "Isaiah 22:12", "that day, the Lord, God of Hosts, called", COMMA_FROM, COMMA_TO
    AddTarget "Isaiah 22:14", "die,#R# says the Lord, God of Hosts.", COMMA_FROM, COMMA_TO
    AddTarget "Isaiah 22:15", "The Lord, God of Hosts says, #L#Go,", COMMA_FROM, COMMA_TO
    AddTarget "Isaiah 28:22", "destruction from the Lord, God of Hosts, on the", COMMA_FROM, COMMA_TO
    AddTarget "Jeremiah 2:19", "in you,#R# says the Lord, God of Hosts.", COMMA_FROM, COMMA_TO
    AddTarget "Jeremiah 46:10 (first)", "day is of the Lord, God of Hosts, a day of", COMMA_FROM, COMMA_TO
    AddTarget "Jeremiah 46:10 (second)", "blood; for the Lord, God of Hosts, has a sacrifice", COMMA_FROM, COMMA_TO
    AddTarget "Jeremiah 49:5", "you,#R# says the Lord, God of Hosts, #L#from all", COMMA_FROM, COMMA_TO
    AddTarget "Jeremiah 50:25", "indignation; for the Lord, God of Hosts, has a work", COMMA_FROM, COMMA_TO
    AddTarget "Jeremiah 50:31", "one,#R# says the Lord, God of Hosts; #L#for your", COMMA_FROM, COMMA_TO
    AddTarget "Amos 9:5", "For the Lord, God of Hosts, is he who", COMMA_FROM, COMMA_TO
End Sub

Private Function ResolveDocument(ByVal sDocName As String) As Word.Document
    If Len(sDocName) = 0 Then
        Set ResolveDocument = ActiveDocument
    Else
        Set ResolveDocument = Documents(sDocName)
    End If
End Function

' Number of hits of a plain text in the document; the first hit goes to oFirst (Nothing when there is none). nInVerse counts the hits
' that sit in a VerseText paragraph.
Private Function CountHits(ByVal oDoc As Word.Document, ByVal sText As String, ByRef oFirst As Word.Range, ByRef nInVerse As Long) As Long
    Dim oRng As Word.Range
    Dim lStart As Long
    Dim lGuard As Long
    Dim n As Long

    Set oFirst = Nothing
    nInVerse = 0
    lStart = 0
    Do
        lGuard = lGuard + 1
        If lGuard > FIND_SAFETY_LIMIT Then Err.Raise vbObjectError + 1, "basDivineNames.CountHits", "safety limit reached"
        Set oRng = oDoc.Range(lStart, oDoc.Content.End)
        With oRng.Find
            .ClearFormatting
            .Text = sText
            .MatchWildcards = False
            .MatchCase = True
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
        End With
        If Not oRng.Find.Execute Then Exit Do
        n = n + 1
        If n = 1 Then Set oFirst = oRng.Duplicate
        If oRng.Paragraphs(1).Range.Style = VERSE_STYLE Then nInVerse = nInVerse + 1
        lStart = oRng.End
    Loop
    CountHits = n
End Function

' The anchor after the edit: the old part replaced by the new part
Private Function DoneAnchor(ByVal i As Long) As String
    DoneAnchor = Replace(m_Anchor(i), m_From(i), m_To(i), 1, 1, vbBinaryCompare)
End Function

' 0 = to do (old anchor found once), 1 = done (new anchor found once), -1 = problem
Private Function StateOf(ByVal oDoc As Word.Document, ByVal i As Long, ByRef oHit As Word.Range) As Long
    Dim nHits As Long
    Dim nVerse As Long

    nHits = CountHits(oDoc, m_Anchor(i), oHit, nVerse)
    If nHits = 1 And nVerse = 1 Then
        StateOf = 0
        Exit Function
    End If
    If nHits = 0 Then
        nHits = CountHits(oDoc, DoneAnchor(i), oHit, nVerse)
        If nHits = 1 And nVerse = 1 Then
            StateOf = 1
            Exit Function
        End If
    End If
    StateOf = -1
End Function

' The report behind ListDivineNameFixes (one line per place, then a summary line); a function so that the Word COM harness can read it.
Public Function DivineNameFixesReport(Optional ByVal sDocName As String = "") As String
    Dim oDoc As Word.Document
    Dim oHit As Word.Range
    Dim i As Long
    Dim lState As Long
    Dim sState As String
    Dim sOut As String
    Dim nTodo As Long, nDone As Long, nProblem As Long
    Dim nHits As Long, nVerse As Long

    Set oDoc = ResolveDocument(sDocName)
    LoadTargets
    For i = 1 To m_Count
        lState = StateOf(oDoc, i, oHit)
        Select Case lState
            Case 0: sState = "to do": nTodo = nTodo + 1
            Case 1: sState = "done": nDone = nDone + 1
            Case Else
                nHits = CountHits(oDoc, m_Anchor(i), oHit, nVerse)
                sState = "PROBLEM (old anchor: " & nHits & " hit(s), " & nVerse & " in " & VERSE_STYLE & "; new anchor: " & _
                    CountHits(oDoc, DoneAnchor(i), oHit, nVerse) & " hit(s))"
                nProblem = nProblem + 1
        End Select
        sOut = sOut & m_Ref(i) & " | " & sState & " | " & m_Anchor(i) & vbCrLf
    Next i
    sOut = sOut & "ListDivineNameFixes: " & m_Count & " places: " & nTodo & " to do, " & nDone & " done, " & nProblem & " problem(s). " & _
        """Lord, God of Hosts"" now in " & CountPlain(oDoc, "Lord, God of Hosts") & " place(s) (expected " & EXPECTED_COMMA_PLACES & " when all is done), " & _
        """Jehovah"" in " & CountPlain(oDoc, "Jehovah") & ", ""Lord the Lord"" in " & CountPlain(oDoc, "Lord the Lord") & " (both expected 0)."
    DivineNameFixesReport = sOut
End Function

Public Sub ListDivineNameFixes(Optional ByVal sDocName As String = "")
    Debug.Print DivineNameFixesReport(sDocName)
End Sub

' The number of places still to do (0 when the docm is finished); -1 when any place is a PROBLEM. For the Word COM harness and for a test.
Public Function CountDivineNameFixesPending(Optional ByVal sDocName As String = "") As Long
    Dim oDoc As Word.Document
    Dim oHit As Word.Range
    Dim i As Long
    Dim lState As Long
    Dim nTodo As Long

    Set oDoc = ResolveDocument(sDocName)
    LoadTargets
    For i = 1 To m_Count
        lState = StateOf(oDoc, i, oHit)
        If lState < 0 Then
            CountDivineNameFixesPending = -1
            Exit Function
        End If
        If lState = 0 Then nTodo = nTodo + 1
    Next i
    CountDivineNameFixesPending = nTodo
End Function

Private Function CountPlain(ByVal oDoc As Word.Document, ByVal sText As String) As Long
    Dim oFirst As Word.Range
    Dim nVerse As Long
    CountPlain = CountHits(oDoc, sText, oFirst, nVerse)
End Function

' The part that really differs between from and to: the common start and the common end are left untouched.
Private Sub MinimalChange(ByVal sFrom As String, ByVal sTo As String, ByRef lPrefix As Long, ByRef sOldCore As String, ByRef sNewCore As String)
    Dim lSuffix As Long
    Dim lMax As Long

    lMax = IIf(Len(sFrom) < Len(sTo), Len(sFrom), Len(sTo))
    lPrefix = 0
    Do While lPrefix < lMax
        If Mid$(sFrom, lPrefix + 1, 1) <> Mid$(sTo, lPrefix + 1, 1) Then Exit Do
        lPrefix = lPrefix + 1
    Loop
    lSuffix = 0
    Do While lSuffix < lMax - lPrefix
        If Mid$(sFrom, Len(sFrom) - lSuffix, 1) <> Mid$(sTo, Len(sTo) - lSuffix, 1) Then Exit Do
        lSuffix = lSuffix + 1
    Loop
    sOldCore = Mid$(sFrom, lPrefix + 1, Len(sFrom) - lPrefix - lSuffix)
    sNewCore = Mid$(sTo, lPrefix + 1, Len(sTo) - lPrefix - lSuffix)
End Sub

Public Function FixDivineNames(Optional ByVal sDocName As String = "") As Long
    On Error GoTo PROC_ERR
    Dim oDoc As Word.Document
    Dim oHit As Word.Range
    Dim oCore As Word.Range
    Dim i As Long
    Dim lState As Long
    Dim lPrefix As Long
    Dim oFrom As Word.Range
    Dim sOldCore As String
    Dim sNewCore As String
    Dim nChanged As Long
    Dim nProblem As Long
    Dim lDelta As Long
    Dim nLenBefore As Long
    Dim nComma As Long
    Dim bScreen As Boolean

    m_LastError = ""
    Set oDoc = ResolveDocument(sDocName)
    If oDoc.TrackRevisions Then
        m_LastError = "FixDivineNames: Track Changes is on in " & oDoc.Name & "; switch it off first (the fix would write revisions)."
        Debug.Print "ERROR in basDivineNames." & m_LastError
        FixDivineNames = -1
        Exit Function
    End If
    LoadTargets
    ' pass 1, read only: every place must be to do or done, otherwise nothing is written
    For i = 1 To m_Count
        If StateOf(oDoc, i, oHit) < 0 Then
            nProblem = nProblem + 1
            Debug.Print "PROBLEM: " & m_Ref(i) & " | " & m_Anchor(i)
        End If
    Next i
    If nProblem > 0 Then
        m_LastError = "FixDivineNames: " & nProblem & " place(s) not found exactly once; nothing was changed. Run ListDivineNameFixes."
        Debug.Print "ERROR in basDivineNames." & m_LastError
        FixDivineNames = -1
        Exit Function
    End If

    bScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    nLenBefore = Len(oDoc.Content.Text)
    ' pass 2: write
    For i = 1 To m_Count
        lState = StateOf(oDoc, i, oHit)
        If lState = 0 Then
            MinimalChange m_From(i), m_To(i), lPrefix, sOldCore, sNewCore
            ' the old text is located with Find, not InStr: Word's Find ignores the optional hyphens (Chr(31)) that the text holds
            ' (for example "Ho-sts"), so the position in the hit's Text would be wrong
            Set oFrom = oHit.Duplicate
            With oFrom.Find
                .ClearFormatting
                .Text = m_From(i)
                .MatchWildcards = False
                .MatchCase = True
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
            End With
            If Not oFrom.Find.Execute Then
                m_LastError = "FixDivineNames: " & m_Ref(i) & ": the text to change is not in the hit. Do not save this document."
                Debug.Print "ERROR in basDivineNames." & m_LastError
                FixDivineNames = -1
                GoTo PROC_EXIT
            End If
            Set oCore = oDoc.Range(oFrom.Start + lPrefix, oFrom.Start + lPrefix + Len(sOldCore))
            If oCore.Text <> sOldCore Then
                m_LastError = "FixDivineNames: " & m_Ref(i) & ": the characters at the position are """ & oCore.Text & """, expected """ & sOldCore & """. Do not save this document."
                Debug.Print "ERROR in basDivineNames." & m_LastError
                FixDivineNames = -1
                GoTo PROC_EXIT
            End If
            If Len(sNewCore) = 0 Then
                oCore.Delete
            Else
                oCore.Text = sNewCore
            End If
            nChanged = nChanged + 1
            lDelta = lDelta + Len(sNewCore) - Len(sOldCore)
            Debug.Print "changed: " & m_Ref(i) & " | " & DoneAnchor(i)
        End If
    Next i

    If Len(oDoc.Content.Text) - nLenBefore <> lDelta Then
        m_LastError = "FixDivineNames: the text length changed by " & (Len(oDoc.Content.Text) - nLenBefore) & ", expected " & lDelta & ". Do not save this document."
        Debug.Print "ERROR in basDivineNames." & m_LastError
        FixDivineNames = -1
        GoTo PROC_EXIT
    End If
    nComma = CountPlain(oDoc, "Lord, God of Hosts")
    If nComma <> EXPECTED_COMMA_PLACES Or CountPlain(oDoc, "Jehovah") <> 0 Or CountPlain(oDoc, "Lord the Lord") <> 0 Then
        m_LastError = "FixDivineNames: after the edits ""Lord, God of Hosts"" stands in " & nComma & " place(s) (expected " & EXPECTED_COMMA_PLACES & _
            "), ""Jehovah"" in " & CountPlain(oDoc, "Jehovah") & " and ""Lord the Lord"" in " & CountPlain(oDoc, "Lord the Lord") & " (both expected 0). Do not save this document."
        Debug.Print "ERROR in basDivineNames." & m_LastError
        FixDivineNames = -1
        GoTo PROC_EXIT
    End If
    Debug.Print "FixDivineNames: " & nChanged & " changed, " & (m_Count - nChanged) & " already done. Comma form stands in " & nComma & " places."
    FixDivineNames = nChanged
PROC_EXIT:
    Application.ScreenUpdating = bScreen
    Exit Function
PROC_ERR:
    m_LastError = "FixDivineNames: " & Err.Number & " " & Err.Description
    Debug.Print "ERROR in basDivineNames." & m_LastError
    FixDivineNames = -1
    Resume PROC_EXIT
End Function
