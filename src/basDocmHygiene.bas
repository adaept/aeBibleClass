Attribute VB_Name = "basDocmHygiene"
Option Explicit
Option Compare Text

' Classification: docm clean-up (list-then-fix pairs). Not used by runtime ribbon navigation.
'
' basDocmHygiene - clean the dev docm before the RWB 1.1.1 export (tasks RT-13 to RT-17, owner decisions S-2, 2026-11-09)
'
' Pattern: every fix has a read-only LIST twin. Run the list, read it, run the fix on a COPY of the docm, run the list again
' (expected: nothing listed), and only then run the fix on the real docm. A fix never changes the text: it compares the
' length of the document text before and after and reports a mismatch.
'
' This file so far (RT-13 only; RT-14 to RT-17 are added to this module as they are written):
'   ListDirectFontRuns "<Doc name>"   runs whose font resolves outside Carlito / Noto Sans / Liberation Serif (reuses the
'                                     audit of basFontRunAudit); read-only
'   FixDirectFontRuns  "<Doc name>"   clears the direct font setting (character Reset) on exactly those runs, EXCEPT runs in
'                                     the paragraph style BodyText: those 3 non-Latin runs (Cambria Math, Arial) are looked
'                                     at by the owner first. Pass bIncludeBodyText:=True only after that look.
'
' Entry points (Immediate window, one statement each; the document name is optional, default ActiveDocument):
'   ListDirectFontRuns
'   FixDirectFontRuns
'   ? CountDirectFontRunsOutsideAllowlist
' Reference: basFontRunAudit (the allowlist and the report), rvw\Plan_rwb_1_1_1_release_train_2026-11-08.md rows RT-13 to RT-17.

Private Const BODY_STYLE As String = "BodyText"
Private Const FIND_SAFETY_LIMIT As Long = 100000
Private Const BODY_TOKEN As String = "+Body"
Private Const HEADINGS_TOKEN As String = "+Headings"
Private Const LANG_ENGLISH_US As Long = 1033
Private Const VERSE_STYLE As String = "VerseText"
Private Const MAX_LISTED_UNITS As Long = 200
Private m_LastError As String

' The text of the last error ("" = none), for callers that cannot see the Immediate window.
Public Function DocmHygieneLastError() As String
    DocmHygieneLastError = m_LastError
End Function

'==============================================================================
' RT-13 - direct font runs
'==============================================================================
Public Sub ListDirectFontRuns(Optional ByVal sDocName As String = "")
    Dim hitCount As Long
    Dim sReport As String

    FindFontRunsOutsideAllowlist hitCount, sReport, ResolveDocument(sDocName)
    Debug.Print sReport
    Debug.Print "ListDirectFontRuns: " & hitCount & " run(s) outside the allowlist (the " & BODY_STYLE & " ones are skipped by FixDirectFontRuns)."
End Sub

Public Function CountDirectFontRunsOutsideAllowlist(Optional ByVal sDocName As String = "") As Long
    CountDirectFontRunsOutsideAllowlist = CountFontRunsOutsideAllowlist(sDocName)
End Function

' Returns the number of runs fixed (-1 = failed, see DocmHygieneLastError). Text is never changed.
Public Function FixDirectFontRuns(Optional ByVal sDocName As String = "", Optional ByVal bIncludeBodyText As Boolean = False) As Long
    On Error GoTo PROC_ERR

    Dim oDoc As Word.Document
    Dim dictFonts As Object
    Dim vFont As Variant
    Dim nFixed As Long
    Dim nSkipped As Long
    Dim nLenBefore As Long
    Dim bScreen As Boolean

    m_LastError = ""
    Set oDoc = ResolveDocument(sDocName)
    bScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    nLenBefore = Len(oDoc.Content.Text)

    Set dictFonts = CreateObject("Scripting.Dictionary")
    dictFonts.CompareMode = vbTextCompare
    CollectOutsideFonts oDoc, dictFonts

    For Each vFont In dictFonts.Keys
        Application.StatusBar = "FixDirectFontRuns: " & CStr(vFont) & "..."
        ResetRunsInFont oDoc, CStr(vFont), bIncludeBodyText, nFixed, nSkipped
    Next vFont

    If Len(oDoc.Content.Text) <> nLenBefore Then
        m_LastError = "FixDirectFontRuns: the document text length changed (" & nLenBefore & " -> " & Len(oDoc.Content.Text) & "); undo or close without saving."
        Debug.Print "ERROR in basDocmHygiene." & m_LastError
        nFixed = -1
    Else
        Debug.Print "FixDirectFontRuns: " & nFixed & " run(s) reset, " & nSkipped & " skipped (" & BODY_STYLE & "). Text length unchanged (" & nLenBefore & ")."
    End If
    FixDirectFontRuns = nFixed

PROC_EXIT:
    Application.StatusBar = False
    Application.ScreenUpdating = bScreen
    Exit Function
PROC_ERR:
    m_LastError = "FixDirectFontRuns: " & Err.Number & " " & Err.Description
    Debug.Print "ERROR in basDocmHygiene." & m_LastError
    FixDirectFontRuns = -1
    Resume PROC_EXIT
End Function

' The fonts to look for: the two theme tokens and every font a style names that is outside the allowlist, plus the fonts the
' audit report names (read from the same places the audit reads: styles and the usual Microsoft fonts as a fallback list).
Private Sub CollectOutsideFonts(ByVal oDoc As Word.Document, ByVal dict As Object)
    Dim oStyle As Word.Style
    Dim v As Variant
    Dim sName As String

    dict.Add BODY_TOKEN, True
    dict.Add HEADINGS_TOKEN, True
    For Each oStyle In oDoc.Styles
        sName = ""
        On Error Resume Next
        sName = oStyle.Font.Name
        On Error GoTo 0
        AddOutsideFont dict, sName
    Next oStyle
    For Each v In Array("Aptos", "Aptos Display", "Calibri", "Calibri Light", "Cambria", "Cambria Math", "Times New Roman", "Arial", _
                        "Arial Unicode MS", "MS Mincho", "Segoe UI", "Courier New", "Symbol", "Wingdings")
        AddOutsideFont dict, CStr(v)
    Next v
End Sub

Private Sub AddOutsideFont(ByVal dict As Object, ByVal sName As String)
    sName = Trim$(sName)
    If Len(sName) = 0 Then Exit Sub
    If IsAllowedRunFont(sName) Then Exit Sub
    If Not dict.Exists(sName) Then dict.Add sName, True
End Sub

' Character Reset on every run found in sFont whose RESOLVED font is outside the allowlist (the same test as the audit).
Private Sub ResetRunsInFont(ByVal oDoc As Word.Document, ByVal sFont As String, ByVal bIncludeBodyText As Boolean, _
                            ByRef nFixed As Long, ByRef nSkipped As Long)
    Dim oStory As Word.Range
    Dim oRng As Word.Range
    Dim r As Word.Range
    Dim lastStart As Long
    Dim guard As Long
    Dim sResolved As String
    Dim sParaStyle As String

    For Each oStory In oDoc.StoryRanges
        Set oRng = oStory
        Do While Not oRng Is Nothing
            Set r = oRng.Duplicate
            lastStart = -1
            guard = 0
            With r.Find
                .ClearFormatting
                .Text = ""
                .Font.Name = sFont
                .Format = True
                .Forward = True
                .Wrap = wdFindStop
                .MatchWildcards = False
                Do While .Execute
                    guard = guard + 1
                    If guard > FIND_SAFETY_LIMIT Then Exit Do
                    If r.Start <= lastStart Then Exit Do
                    lastStart = r.Start
                    sResolved = ""
                    sParaStyle = ""
                    On Error Resume Next
                    sResolved = r.Font.Name
                    sParaStyle = r.Paragraphs(1).style.NameLocal
                    On Error GoTo 0
                    If Not IsAllowedRunFont(sResolved) Then
                        If StrComp(sParaStyle, BODY_STYLE, vbTextCompare) = 0 And Not bIncludeBodyText Then
                            nSkipped = nSkipped + 1
                        Else
                            ResetKeepingCharacterStyles r, oDoc
                            nFixed = nFixed + 1
                        End If
                    End If
                    r.Collapse wdCollapseEnd
                Loop
            End With
            Set oRng = oRng.NextStoryRange
        Loop
    Next oStory
End Sub

' Font.Reset removes the character style as well as the direct formatting (measured 2026-11-09: it stripped "Words of Jesus" and its
' colour from 4 verse runs). So the character style of every character is noted first and put back after the reset.
Private Sub ResetKeepingCharacterStyles(ByVal r As Word.Range, ByVal oDoc As Word.Document)
    Dim nChars As Long
    Dim k As Long
    Dim sDefault As String
    Dim saved() As String

    sDefault = oDoc.Styles(wdStyleDefaultParagraphFont).NameLocal
    nChars = r.End - r.Start
    If nChars < 1 Then Exit Sub
    ReDim saved(1 To nChars)
    For k = 1 To nChars
        saved(k) = sDefault
        On Error Resume Next
        saved(k) = oDoc.Range(r.Start + k - 1, r.Start + k).CharacterStyle.NameLocal
        On Error GoTo 0
    Next k

    r.Font.Reset

    For k = 1 To nChars
        If StrComp(saved(k), sDefault, vbTextCompare) <> 0 Then
            On Error Resume Next
            oDoc.Range(r.Start + k - 1, r.Start + k).style = saved(k)
            On Error GoTo 0
        End If
    Next k
End Sub

'==============================================================================
' RT-14 - language of the verse text (F1): English (US), no direct de-DE tags
'==============================================================================
' A "unit" is one stretch of verse text (or one paragraph mark) whose language is not English (US). Only paragraphs in the style
' VerseText are looked at. Entry points (the document name is optional):
'   ListVerseRunsNotEnglish         read-only list to the Immediate window (position, language, kind, text)
'   ? CountVerseRunsNotEnglish      the number of units (Test 96: expected 0)
'   ? FixVerseTextLanguage          sets the VerseText style language to English (US), then sets exactly the listed units to
'                                   English (US); returns the number of units fixed (-1 = failed, see DocmHygieneLastError)

Public Sub ListVerseRunsNotEnglish(Optional ByVal sDocName As String = "")
    Dim oDoc As Word.Document
    Dim nUnits As Long
    Dim sReport As String

    Set oDoc = ResolveDocument(sDocName)
    ScanVerseLanguage oDoc, False, nUnits, sReport
    Debug.Print sReport
    Debug.Print "ListVerseRunsNotEnglish: " & nUnits & " unit(s) outside English (US); style " & VERSE_STYLE & " language " & oDoc.Styles(VERSE_STYLE).LanguageID & "."
End Sub

Public Function CountVerseRunsNotEnglish(Optional ByVal sDocName As String = "") As Long
    Dim nUnits As Long
    Dim sReport As String
    ScanVerseLanguage ResolveDocument(sDocName), False, nUnits, sReport
    CountVerseRunsNotEnglish = nUnits
End Function

Public Function FixVerseTextLanguage(Optional ByVal sDocName As String = "") As Long
    On Error GoTo PROC_ERR
    Dim oDoc As Word.Document
    Dim nUnits As Long
    Dim sReport As String
    Dim nLenBefore As Long

    m_LastError = ""
    Set oDoc = ResolveDocument(sDocName)
    nLenBefore = Len(oDoc.Content.Text)
    oDoc.Styles(VERSE_STYLE).LanguageID = LANG_ENGLISH_US
    ScanVerseLanguage oDoc, True, nUnits, sReport
    If Len(oDoc.Content.Text) <> nLenBefore Then
        m_LastError = "FixVerseTextLanguage: the document text length changed (" & nLenBefore & " -> " & Len(oDoc.Content.Text) & ")."
        Debug.Print "ERROR in basDocmHygiene." & m_LastError
        FixVerseTextLanguage = -1
    Else
        Debug.Print "FixVerseTextLanguage: style language set; " & nUnits & " unit(s) set to English (US). Text length unchanged (" & nLenBefore & ")."
        FixVerseTextLanguage = nUnits
    End If
    Exit Function
PROC_ERR:
    m_LastError = "FixVerseTextLanguage: " & Err.Number & " " & Err.Description
    Debug.Print "ERROR in basDocmHygiene." & m_LastError
    FixVerseTextLanguage = -1
End Function

' Finds every character whose language is not English (US) by bisection: Range.LanguageID of a range is English (US) only when every
' character of it is, so uniform stretches are skipped in one call (measured 2026-11-10: 15 s for the whole docm; walking the 33,826
' paragraphs one by one was still running after 10 minutes). Characters outside VerseText paragraphs (for example the German tags in
' BibleIndexList) are counted but left alone.
Private Sub ScanVerseLanguage(ByVal oDoc As Word.Document, ByVal bFix As Boolean, ByRef nUnits As Long, ByRef sReport As String)
    Dim cFound As Collection
    Dim v As Variant
    Dim lPos As Long
    Dim lLang As Long
    Dim lParaEnd As Long
    Dim bVerse As Boolean
    Dim lUnitStart As Long
    Dim lUnitEnd As Long
    Dim lUnitLang As Long
    Dim bBusy As Boolean
    Dim nOutside As Long
    Dim bScreen As Boolean
    Dim oChar As Word.Range

    bScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    nUnits = 0
    Set cFound = New Collection
    CollectNonEnglishChars oDoc, 0, oDoc.Content.End, cFound
    sReport = "VerseText language scan - " & oDoc.Name & " - " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
              "Position | Language id | Kind | Text" & vbCrLf & String$(60, "-") & vbCrLf
    lParaEnd = -1
    For Each v In cFound
        lPos = CLng(Split(CStr(v), ":")(0))
        lLang = CLng(Split(CStr(v), ":")(1))
        If lPos >= lParaEnd Then
            Set oChar = oDoc.Range(lPos, lPos + 1)
            lParaEnd = oChar.Paragraphs(1).Range.End
            bVerse = (StrComp(oChar.Paragraphs(1).style.NameLocal, VERSE_STYLE, vbTextCompare) = 0)
        End If
        If Not bVerse Then
            nOutside = nOutside + 1
        Else
            Set oChar = oDoc.Range(lPos, lPos + 1)
            If oChar.Text = vbCr Then
                If bBusy Then FinishUnit oDoc, lUnitStart, lUnitEnd, lUnitLang, "text", bFix, nUnits, sReport: bBusy = False
                FinishUnit oDoc, lPos, lPos + 1, lLang, "paragraph mark", bFix, nUnits, sReport
            ElseIf bBusy And lPos = lUnitEnd And lLang = lUnitLang Then
                lUnitEnd = lPos + 1
            Else
                If bBusy Then FinishUnit oDoc, lUnitStart, lUnitEnd, lUnitLang, "text", bFix, nUnits, sReport
                lUnitStart = lPos: lUnitEnd = lPos + 1: lUnitLang = lLang: bBusy = True
            End If
        End If
    Next v
    If bBusy Then FinishUnit oDoc, lUnitStart, lUnitEnd, lUnitLang, "text", bFix, nUnits, sReport
    Application.ScreenUpdating = bScreen
    sReport = sReport & "TOTAL units outside English (US) in " & VERSE_STYLE & ": " & nUnits & "   (characters outside " & VERSE_STYLE & ", not touched: " & nOutside & ")" & vbCrLf
End Sub

' Appends "position:languageId" to cFound, in document order, for every character of [lStart, lEnd) that is not English (US).
Private Sub CollectNonEnglishChars(ByVal oDoc As Word.Document, ByVal lStart As Long, ByVal lEnd As Long, ByVal cFound As Collection)
    Dim lLang As Long
    Dim lMid As Long

    lLang = oDoc.Range(lStart, lEnd).LanguageID
    If lLang = LANG_ENGLISH_US Then Exit Sub
    If lEnd - lStart <= 1 Then
        cFound.Add lStart & ":" & lLang
        Exit Sub
    End If
    lMid = lStart + (lEnd - lStart) \ 2
    CollectNonEnglishChars oDoc, lStart, lMid, cFound
    CollectNonEnglishChars oDoc, lMid, lEnd, cFound
End Sub

Private Sub FinishUnit(ByVal oDoc As Word.Document, ByVal lStart As Long, ByVal lEnd As Long, ByVal lLang As Long, _
                       ByVal sKind As String, ByVal bFix As Boolean, ByRef nUnits As Long, ByRef sReport As String)
    nUnits = nUnits + 1
    If nUnits <= MAX_LISTED_UNITS Then
        sReport = sReport & lStart & "-" & lEnd & " | " & lLang & " | " & sKind & " | """ & Left$(CleanText(oDoc.Range(lStart, lEnd).Text), 30) & """" & vbCrLf
    End If
    If bFix Then oDoc.Range(lStart, lEnd).LanguageID = LANG_ENGLISH_US
End Sub

' ASCII-only, single line (the report is ASCII like the other rpt\ files).
Private Function CleanText(ByVal s As String) As String
    Dim i As Long
    Dim c As Long
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536
        If c >= 32 And c <= 126 Then CleanText = CleanText & Mid$(s, i, 1) Else CleanText = CleanText & "?"
    Next i
End Function

'==============================================================================
' Small helpers
'==============================================================================
Private Function ResolveDocument(ByVal sDocName As String) As Word.Document
    If Len(sDocName) = 0 Then
        Set ResolveDocument = ActiveDocument
    Else
        Set ResolveDocument = Documents(sDocName)
    End If
End Function
