Attribute VB_Name = "basFormalVoice"
Option Explicit

' Classification: docm clean-up and test (list-then-fix pair plus Test 100). Not used by runtime ribbon navigation.
'
' basFormalVoice - the RWB "formal voice" statement: the verse text has no contractions (owner decision 2026-11-13, FU-51, task RT-84).
'
' What counts as a contraction: a word of letters that ends in an apostrophe (U+2019) plus one of  ll  re  ve  d  m  or  n't, and the
' word "let's". The possessive 's is NOT a contraction and is never looked at (God's, king's). Only paragraphs in the style VerseText
' are counted, because the rule is about the verse text; front matter, headings and notes are not read.
' The straight apostrophe (U+0027) is not searched: the Bible text is normalised to U+2019 (Tests 52 to 65), and Word's Find treats
' the two quote marks as equal in a plain search, so a second pass could count the same word twice.
'
' Entry points (Immediate window, one statement each; the document name is optional, default ActiveDocument):
'   ListContractions                 read-only: every contraction in a VerseText paragraph, with the start of its paragraph
'   ? CountContractions              the number of them (Test 100: expected 0)
'   ? FixContractions "<Doc name>"   writes the full form ("you'll" to "you will", "we're" to "we are", "you've" to "you have", "I'm" to
'                                    "I am", "let's" to "let us", "n't" to " not", can't / won't / shan't to cannot / will not /
'                                    shall not) and returns the number changed, -1 = failed (FormalVoiceLastError). It changes
'                                    only the ending of the word, so the run formatting of the word is kept. It leaves "'d"
'                                    (would or had) and "ain't" alone and reports them: the reader of each verse decides.
'                                    Run it on a COPY of the docm first, then ListContractions (expected: nothing, or only the
'                                    undecided forms).
' Reference: rvw\RWB 1.1.1 Follow-up Issues.md FU-51 (aeBibleAddin), release train row RT-84.

Private Const VERSE_STYLE As String = "VerseText"
Private Const FIND_SAFETY_LIMIT As Long = 100000
Private Const CONTEXT_CHARS As Long = 40
' Whole words (lower case, "|" between them) that end like a contraction but are not one. Empty on purpose: add a word here only after
' the owner has looked at it (names such as "Dan'd" do not exist in the text today).
Private Const ALLOWLIST As String = ""

Private m_LastError As String

Public Property Get FormalVoiceLastError() As String
    FormalVoiceLastError = m_LastError
End Property

' The Find patterns (wildcards). Q stands for the apostrophe. Every pattern ends in [!A-Za-z] so that "'ll" does not match inside a
' longer word; the let's patterns start with [!A-Za-z] so that "outlet's" is not a hit. Word's wildcard Find is case sensitive, so the
' capital Let's has its own pattern.
Private Function ContractionPatterns() As Variant
    ContractionPatterns = Array( _
        "[A-Za-z]Qll[!A-Za-z]", _
        "[A-Za-z]Qre[!A-Za-z]", _
        "[A-Za-z]Qve[!A-Za-z]", _
        "[A-Za-z]Qd[!A-Za-z]", _
        "[A-Za-z]Qm[!A-Za-z]", _
        "[A-Za-z]nQt[!A-Za-z]", _
        "[!A-Za-z]letQs[!A-Za-z]", _
        "[!A-Za-z]LetQs[!A-Za-z]")
End Function

Private Function ResolveDocument(ByVal sDocName As String) As Word.Document
    If Len(sDocName) = 0 Then
        Set ResolveDocument = ActiveDocument
    Else
        Set ResolveDocument = Documents(sDocName)
    End If
End Function

Private Function IsLetter(ByVal sChar As String) As Boolean
    IsLetter = (Len(sChar) = 1) And (sChar Like "[A-Za-z]")
End Function

Private Function IsAllowlisted(ByVal sWord As String) As Boolean
    IsAllowlisted = (InStr(1, "|" & ALLOWLIST & "|", "|" & LCase$(sWord) & "|", vbBinaryCompare) > 0)
End Function

' The word around a hit: the pattern match has one boundary character (the first for the let's patterns, the last for all patterns) that is
' not part of the word, and only one letter before the apostrophe, so the start of the word is found by walking back.
Private Sub WordOfHit(ByVal oDoc As Word.Document, ByVal oHit As Word.Range, ByRef lWordStart As Long, ByRef lWordEnd As Long, ByRef sWord As String)
    Dim sText As String
    Dim lFirst As Long
    Dim lLast As Long

    sText = oHit.Text
    lFirst = 1
    lLast = Len(sText)
    If Not IsLetter(Mid$(sText, lFirst, 1)) Then lFirst = lFirst + 1
    If Not IsLetter(Mid$(sText, lLast, 1)) Then lLast = lLast - 1
    lWordStart = oHit.Start + lFirst - 1
    lWordEnd = oHit.Start + lLast
    ' the pattern holds only one letter before the apostrophe: walk back to the start of the word ("on" + "n" + "t" is "don" + "t")
    Do While lWordStart > 0
        If Not IsLetter(oDoc.Range(lWordStart - 1, lWordStart).Text) Then Exit Do
        lWordStart = lWordStart - 1
    Loop
    sWord = oDoc.Range(lWordStart, lWordEnd).Text
End Sub

' What FixContractions writes for a word: the first character index (1-based, in the word) of the part that is replaced and the
' replacement text. Returns False for the forms left alone ("'d", ain't).
Private Function FixFor(ByVal sWord As String, ByRef lReplaceFrom As Long, ByRef sNew As String) As Boolean
    Dim sLower As String
    Dim sStem As String
    Dim lApos As Long

    sLower = LCase$(sWord)
    lApos = InStr(1, sWord, ChrW$(&H2019))
    If lApos = 0 Then Exit Function

    Select Case True
        Case sLower = "let" & ChrW$(&H2019) & "s"
            lReplaceFrom = lApos: sNew = " us"
        Case Right$(sLower, 3) = ChrW$(&H2019) & "ll"
            lReplaceFrom = lApos: sNew = " will"
        Case Right$(sLower, 3) = ChrW$(&H2019) & "re"
            lReplaceFrom = lApos: sNew = " are"
        Case Right$(sLower, 3) = ChrW$(&H2019) & "ve"
            lReplaceFrom = lApos: sNew = " have"
        Case Right$(sLower, 2) = ChrW$(&H2019) & "m"
            lReplaceFrom = lApos: sNew = " am"
        Case Right$(sLower, 3) = "n" & ChrW$(&H2019) & "t"
            sStem = Left$(sLower, Len(sLower) - 3)
            Select Case sStem
                Case "ca"
                    lReplaceFrom = 1: sNew = IIf(Left$(sWord, 1) = "C", "Cannot", "cannot")
                Case "wo"
                    lReplaceFrom = 1: sNew = IIf(Left$(sWord, 1) = "W", "Will not", "will not")
                Case "sha"
                    lReplaceFrom = 1: sNew = IIf(Left$(sWord, 1) = "S", "Shall not", "shall not")
                Case "ai"
                    Exit Function
                Case Else
                    lReplaceFrom = lApos - 1: sNew = " not"
            End Select
        Case Else
            Exit Function      ' 'd: would or had, decided by reading the verse
    End Select
    FixFor = True
End Function

' The one scan behind the list, the count and the fix. bFix False = read only. nFound counts the contractions in VerseText paragraphs that
' are not allowlisted; nFixed the ones written; nUndecided the ones the fix leaves alone; nDelta the expected change of the text length.
Private Sub ScanContractions(ByVal oDoc As Word.Document, ByVal bFix As Boolean, ByRef nFound As Long, ByRef nFixed As Long, _
        ByRef nUndecided As Long, ByRef nDelta As Long, ByRef sReport As String)
    Dim vPatterns As Variant
    Dim vP As Variant
    Dim sPattern As String
    Dim oRng As Word.Range
    Dim lStart As Long
    Dim lGuard As Long
    Dim lWordStart As Long
    Dim lWordEnd As Long
    Dim sWord As String
    Dim lReplaceFrom As Long
    Dim sNew As String
    Dim oPara As Word.Range
    Dim sContext As String
    Dim oTail As Word.Range
    Dim lOldLen As Long

    nFound = 0: nFixed = 0: nUndecided = 0: nDelta = 0: sReport = ""
    vPatterns = ContractionPatterns()
    For Each vP In vPatterns
        sPattern = Replace(CStr(vP), "Q", ChrW$(&H2019))
        lStart = 0
        Do
            lGuard = lGuard + 1
            If lGuard > FIND_SAFETY_LIMIT Then Err.Raise vbObjectError + 1, "basFormalVoice.ScanContractions", "safety limit reached"
            Set oRng = oDoc.Range(lStart, oDoc.Content.End)
            With oRng.Find
                .ClearFormatting
                .Text = sPattern
                .MatchWildcards = True
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
            End With
            If Not oRng.Find.Execute Then Exit Do
            WordOfHit oDoc, oRng, lWordStart, lWordEnd, sWord
            lStart = oRng.End - 1               ' keep the closing boundary character available to the next hit
            Set oPara = oRng.Paragraphs(1).Range
            If oPara.Style = VERSE_STYLE And Not IsAllowlisted(sWord) Then
                nFound = nFound + 1
                sContext = Replace(Mid$(oPara.Text, 1, CONTEXT_CHARS + Len(sWord) + CONTEXT_CHARS), vbCr, "")
                sReport = sReport & sWord & " | " & oPara.Start & " | " & sContext & vbCrLf
                If bFix Then
                    If FixFor(sWord, lReplaceFrom, sNew) Then
                        Set oTail = oDoc.Range(lWordStart + lReplaceFrom - 1, lWordEnd)
                        lOldLen = oTail.End - oTail.Start
                        oTail.Text = sNew
                        nFixed = nFixed + 1
                        nDelta = nDelta + Len(sNew) - lOldLen
                        lStart = lWordStart + lReplaceFrom - 1 + Len(sNew)
                    Else
                        nUndecided = nUndecided + 1
                    End If
                End If
            End If
        Loop
    Next vP
End Sub

Public Sub ListContractions(Optional ByVal sDocName As String = "")
    Dim nFound As Long, nFixed As Long, nUndecided As Long, nDelta As Long
    Dim sReport As String

    ScanContractions ResolveDocument(sDocName), False, nFound, nFixed, nUndecided, nDelta, sReport
    Debug.Print "ListContractions: " & nFound & " contraction(s) in " & VERSE_STYLE & " paragraphs (word | paragraph start | text of the paragraph)"
    If nFound > 0 Then Debug.Print sReport
End Sub

' Test 100: the number of contractions in the verse text (expected 0).
Public Function CountContractions(Optional ByVal sDocName As String = "") As Long
    Dim nFound As Long, nFixed As Long, nUndecided As Long, nDelta As Long
    Dim sReport As String

    ScanContractions ResolveDocument(sDocName), False, nFound, nFixed, nUndecided, nDelta, sReport
    CountContractions = nFound
End Function

Public Function FixContractions(Optional ByVal sDocName As String = "") As Long
    On Error GoTo PROC_ERR
    Dim oDoc As Word.Document
    Dim nFound As Long, nFixed As Long, nUndecided As Long, nDelta As Long
    Dim nLeft As Long
    Dim sReport As String
    Dim nLenBefore As Long
    Dim bScreen As Boolean

    m_LastError = ""
    Set oDoc = ResolveDocument(sDocName)
    If oDoc.TrackRevisions Then
        m_LastError = "FixContractions: Track Changes is on in " & oDoc.Name & "; switch it off first (the fix would write revisions)."
        Debug.Print "ERROR in basFormalVoice." & m_LastError
        FixContractions = -1
        Exit Function
    End If
    bScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    nLenBefore = Len(oDoc.Content.Text)
    ScanContractions oDoc, True, nFound, nFixed, nUndecided, nDelta, sReport
    Debug.Print "FixContractions: " & nFound & " found, " & nFixed & " changed, " & nUndecided & " left for a decision."
    Debug.Print sReport
    If Len(oDoc.Content.Text) - nLenBefore <> nDelta Then
        m_LastError = "FixContractions: the text length changed by " & (Len(oDoc.Content.Text) - nLenBefore) & ", expected " & nDelta & ". Do not save this document."
        Debug.Print "ERROR in basFormalVoice." & m_LastError
        FixContractions = -1
        GoTo PROC_EXIT
    End If
    nLeft = CountContractions(oDoc.Name)
    If nLeft <> nUndecided Then
        m_LastError = "FixContractions: " & nLeft & " contraction(s) left, expected " & nUndecided & " (the undecided forms). Do not save this document."
        Debug.Print "ERROR in basFormalVoice." & m_LastError
        FixContractions = -1
        GoTo PROC_EXIT
    End If
    FixContractions = nFixed
PROC_EXIT:
    Application.ScreenUpdating = bScreen
    Exit Function
PROC_ERR:
    m_LastError = "FixContractions: " & Err.Number & " " & Err.Description
    Debug.Print "ERROR in basFormalVoice." & m_LastError
    FixContractions = -1
    Resume PROC_EXIT
End Function
