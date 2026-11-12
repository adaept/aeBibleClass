Attribute VB_Name = "basFontRunAudit"
Option Explicit
Option Compare Text

' Classification: font audit, READ-ONLY. Not used by runtime ribbon navigation.
'
' basFontRunAudit - which font does each RUN of text actually resolve to? (task PB-20, 2026-11-07)
'
' Why this exists: AuditFontsUsage (basTEST_aeBibleFonts) reads oPara.Range.Font.Name per PARAGRAPH and skips the paragraph
' when the name is empty. Word returns an empty name when a paragraph mixes fonts, and every verse paragraph does (Noto Sans
' verse numbers, Carlito text), so a single run set in another font inside a verse is invisible to that audit. This module asks
' Word a different question: for each font that could be in play, which stretches of text are set in it? (Find with Font.Name
' matches the font of the text however it got there: direct formatting, character style, paragraph style or theme.)
'
' Rule (owner, 2026-11-07): the RWB text is set in free fonts only - Carlito, Noto Sans, Liberation Serif - because the fonts
' that ship with Office and Windows (Calibri, Times New Roman, Arial, Aptos ...) restrict printing and distribution.
' Theme fonts are reported under their token: Word's Find sees "+Body" (the theme body font, Aptos in this document's theme)
' and "+Headings", while Range.Font.Name resolves them to the theme font's name. A run found through a token is counted only
' when the font it resolves to is outside the allowlist.
'
' What it does, in order (nothing is changed in the document; no Selection, no clipboard, no save):
'   1. Candidate fonts = theme tokens +Body/+Headings, every style's font name, and every font named in a w:rFonts element
'      (read once from Content.WordOpenXML). Any font in play is named in one of those places; if the XML cannot be read, a
'      built-in list of the usual Microsoft fonts is tried as well.
'   2. For every candidate OUTSIDE the allowlist, Find walks every story (body, headers, footers, footnotes ...) and records
'      each run: where (book chapter:verse for verse text), paragraph style, character style, resolved font, a text snippet.
'   3. Reports the total. 0 = clean.
'
' Entry points (Immediate window, one statement each):
'   FontRunAudit                          lists to the Immediate window and writes rpt\FontRunAudit.txt when that folder exists
'   FontRunAuditToFile "C:\path\x.txt"    same, to the file you name
'   ? CountFontRunsOutsideAllowlist       the Count only (for a test: expected 0)
' Both file/Count forms take an optional second/first argument: the NAME of an open document to audit instead of the active one.
' Runs against ActiveDocument by default. About a minute on the full docm; progress is shown in the status bar.

Private Const ALLOWED_FONTS As String = "|Carlito|Noto Sans|Liberation Serif|"
Private Const MAX_LISTED_HITS As Long = 300
Private Const FIND_SAFETY_LIMIT As Long = 100000
Private Const BODY_TOKEN As String = "+Body"
Private m_Step As String
Private m_LastError As String   ' set by the handlers; read with FontRunAuditLastError
Private Const HEADINGS_TOKEN As String = "+Headings"

' The text of the last error the audit hit ("" = none): lets a caller that cannot see the Immediate window find out why a run failed.
Public Function FontRunAuditLastError() As String
    FontRunAuditLastError = m_LastError
End Function

Public Function IsAllowedRunFont(ByVal sFontName As String) As Boolean
    If Len(sFontName) = 0 Then Exit Function
    IsAllowedRunFont = (InStr(1, ALLOWED_FONTS, "|" & sFontName & "|", vbTextCompare) > 0)
End Function

'==============================================================================
' Entry points
'==============================================================================
Public Sub FontRunAudit()
    Dim hitCount As Long
    Dim sReport As String
    Dim sPath As String

    FindFontRunsOutsideAllowlist hitCount, sReport
    Debug.Print sReport
    On Error Resume Next
    If Len(Dir(ActiveDocument.Path & "\rpt", vbDirectory)) > 0 Then sPath = ActiveDocument.Path & "\rpt\FontRunAudit.txt"
    On Error GoTo 0
    If Len(sPath) > 0 Then
        WriteReportFile sPath, sReport
        Debug.Print "Report written: " & sPath
    Else
        Debug.Print "No rpt folder next to the document: report printed only."
    End If
End Sub

Public Sub FontRunAuditToFile(ByVal sReportPath As String, Optional ByVal sDocName As String = "")
    On Error GoTo PROC_ERR
    Dim hitCount As Long
    Dim sReport As String

    FindFontRunsOutsideAllowlist hitCount, sReport, ResolveDocument(sDocName)
    WriteReportFile sReportPath, sReport
    Debug.Print "FontRunAudit: " & hitCount & " run(s) outside the allowlist. Report: " & sReportPath
    Exit Sub
PROC_ERR:
    m_LastError = "FontRunAuditToFile: " & Err.Number & " " & Err.Description
    Debug.Print "ERROR in basFontRunAudit." & m_LastError
End Sub

Public Function CountFontRunsOutsideAllowlist(Optional ByVal sDocName As String = "") As Long
    Dim hitCount As Long
    Dim sReport As String
    FindFontRunsOutsideAllowlist hitCount, sReport, ResolveDocument(sDocName)
    CountFontRunsOutsideAllowlist = hitCount
End Function

' The document to audit: the one named (an open document's name), else the active document.
Private Function ResolveDocument(ByVal sDocName As String) As Word.Document
    If Len(sDocName) = 0 Then
        Set ResolveDocument = ActiveDocument
    Else
        Set ResolveDocument = Documents(sDocName)
    End If
End Function

'==============================================================================
' FindFontRunsOutsideAllowlist
' hitCount : number of runs found in a font outside the allowlist (0 = clean; -1 = the audit itself failed, see FontRunAuditLastError)
' sReport  : the full report text
' oTarget  : the document to audit (default ActiveDocument)
'==============================================================================
Public Sub FindFontRunsOutsideAllowlist(ByRef hitCount As Long, ByRef sReport As String, _
                                        Optional ByVal oTarget As Word.Document)
    On Error GoTo PROC_ERR

    Dim oDoc As Word.Document
    Dim dictCandidates As Object
    Dim vCand As Variant
    Dim sFont As String
    Dim nBooks As Long
    Dim bookStarts() As Long
    Dim bookNames() As String
    Dim sHits As String
    Dim nListed As Long
    Dim nCandidate As Long
    Dim nPerFont As Long
    Dim nResolvedOk As Long
    Dim sPerFont As String
    Dim sCandList As String
    Dim t0 As Double
    Dim bScreen As Boolean

    t0 = Timer
    If oTarget Is Nothing Then Set oDoc = ActiveDocument Else Set oDoc = oTarget
    hitCount = 0
    bScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False

    m_LastError = "": m_Step = "candidate fonts"
    Application.StatusBar = "FontRunAudit: collecting candidate fonts..."
    Set dictCandidates = CreateObject("Scripting.Dictionary")
    dictCandidates.CompareMode = vbTextCompare
    CollectCandidateFonts oDoc, dictCandidates

    m_Step = "book headings"
    Application.StatusBar = "FontRunAudit: locating the book headings..."
    CollectBookStarts oDoc, nBooks, bookStarts, bookNames

    For Each vCand In dictCandidates.Keys
        sFont = CStr(vCand)
        If Not IsAllowedRunFont(sFont) Then
            nCandidate = nCandidate + 1
            m_Step = "searching " & sFont
            Application.StatusBar = "FontRunAudit: searching for text set in " & sFont & " (" & nCandidate & ")..."
            nPerFont = 0
            nResolvedOk = 0
            FindRunsInFont oDoc, sFont, nBooks, bookStarts, bookNames, nPerFont, nResolvedOk, sHits, nListed
            hitCount = hitCount + nPerFont
            sPerFont = sPerFont & "  " & Left$(sFont & space$(26), 26) & nPerFont & " run(s)" & _
                       IIf(nResolvedOk > 0, "   (+" & nResolvedOk & " more that resolve to an allowed font: not counted)", "") & vbCrLf
        End If
        sCandList = sCandList & sFont & "; "
    Next vCand

    sReport = "FontRunAudit - " & oDoc.Name & " - " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
              "Allowed fonts: Carlito, Noto Sans, Liberation Serif (owner rule 2026-11-07: free fonts only, no print restrictions)" & vbCrLf & _
              "Fonts considered (" & dictCandidates.Count & "): " & sCandList & vbCrLf & vbCrLf & _
              "Runs found in a font OUTSIDE the allowlist, by font:" & vbCrLf & _
              IIf(Len(sPerFont) = 0, "  (no font outside the allowlist is in play)" & vbCrLf, sPerFont) & vbCrLf
    If hitCount > 0 Then
        sReport = sReport & "Location | Story | Paragraph style | Character style | Resolved font | Text" & vbCrLf & String$(80, "-") & vbCrLf & sHits
        If nListed < hitCount Then sReport = sReport & "... " & (hitCount - nListed) & " more not listed (limit " & MAX_LISTED_HITS & ")" & vbCrLf
    End If
    sReport = sReport & vbCrLf & "TOTAL runs outside the allowlist: " & hitCount & "   (" & Format$(Timer - t0, "0.0") & " s)" & vbCrLf

PROC_EXIT:
    Application.StatusBar = False
    Application.ScreenUpdating = bScreen
    Exit Sub

PROC_ERR:
    m_LastError = "FindFontRunsOutsideAllowlist (" & m_Step & "): " & Err.Number & " " & Err.Description
    Debug.Print "ERROR in basFontRunAudit." & m_LastError
    sReport = sReport & vbCrLf & "ERROR " & Err.Number & ": " & Err.Description & vbCrLf
    hitCount = -1
    Resume PROC_EXIT
End Sub

'==============================================================================
' Candidate fonts
'==============================================================================
Private Sub CollectCandidateFonts(ByVal oDoc As Word.Document, ByVal dict As Object)
    Dim v As Variant
    Dim oStyle As Word.Style
    Dim sName As String
    Dim sXml As String

    AddCandidate dict, BODY_TOKEN
    AddCandidate dict, HEADINGS_TOKEN

    For Each oStyle In oDoc.Styles
        sName = ""
        On Error Resume Next
        sName = oStyle.Font.Name
        On Error GoTo 0
        AddCandidate dict, sName
    Next oStyle

    sXml = ""
    On Error Resume Next
    sXml = oDoc.Content.WordOpenXML
    On Error GoTo 0
    If Len(sXml) > 0 Then
        HarvestRFontsFromXml sXml, dict
    Else
        ' Fallback only: the XML could not be read, so also try the usual Microsoft fonts by name.
        For Each v In Array("Aptos", "Aptos Display", "Calibri", "Calibri Light", "Cambria", "Times New Roman", "Arial", "Arial Black", _
                            "Arial Narrow", "Arial Unicode MS", "MS Mincho", "Segoe UI", "Verdana", "Tahoma", "Courier New", "Symbol", _
                            "Wingdings", "Georgia", "Consolas", "Liberation Sans Narrow")
            AddCandidate dict, CStr(v)
        Next v
    End If
End Sub

Private Sub AddCandidate(ByVal dict As Object, ByVal sName As String)
    sName = Trim$(sName)
    If Len(sName) = 0 Then Exit Sub
    If Not dict.Exists(sName) Then dict.Add sName, True
End Sub

' Every font named by a w:rFonts element: w:ascii / w:hAnsi names, and the theme attributes mapped to the Find tokens.
Private Sub HarvestRFontsFromXml(ByVal sXml As String, ByVal dict As Object)
    Dim pos As Long
    Dim endPos As Long
    Dim sTag As String

    pos = 1
    Do
        pos = InStr(pos, sXml, "<w:rFonts ", vbBinaryCompare)
        If pos = 0 Then Exit Do
        endPos = InStr(pos, sXml, "/>", vbBinaryCompare)
        If endPos = 0 Then Exit Do
        sTag = Mid$(sXml, pos, endPos - pos)
        AddCandidate dict, XmlAttr(sTag, "w:ascii")
        AddCandidate dict, XmlAttr(sTag, "w:hAnsi")
        If InStr(1, sTag, "Theme=""minor", vbBinaryCompare) > 0 Then AddCandidate dict, BODY_TOKEN
        If InStr(1, sTag, "Theme=""major", vbBinaryCompare) > 0 Then AddCandidate dict, HEADINGS_TOKEN
        pos = endPos
    Loop
End Sub

Private Function XmlAttr(ByVal sTag As String, ByVal sAttr As String) As String
    Dim p As Long
    Dim q As Long
    p = InStr(1, sTag, sAttr & "=""", vbBinaryCompare)
    If p = 0 Then Exit Function
    p = p + Len(sAttr) + 2
    q = InStr(p, sTag, """", vbBinaryCompare)
    If q > p Then XmlAttr = Mid$(sTag, p, q - p)
End Function

'==============================================================================
' Book headings (to say "Leviticus 2:11" instead of a character position)
'==============================================================================
Private Sub CollectBookStarts(ByVal oDoc As Word.Document, ByRef nBooks As Long, _
                              ByRef bookStarts() As Long, ByRef bookNames() As String)
    Dim rng As Word.Range
    Dim lastStart As Long

    ReDim bookStarts(1 To 200)
    ReDim bookNames(1 To 200)
    nBooks = 0
    lastStart = -1
    Set rng = oDoc.Content
    With rng.Find
        .ClearFormatting
        .Text = ""
        .Style = oDoc.Styles(wdStyleHeading1)
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            If rng.Start <= lastStart Then Exit Do
            lastStart = rng.Start
            If nBooks < 200 Then
                nBooks = nBooks + 1
                bookStarts(nBooks) = rng.Paragraphs(1).Range.Start
                bookNames(nBooks) = CleanSnippet(rng.Paragraphs(1).Range.Text, 40)
            End If
            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub

'==============================================================================
' Find every run set in one font, in every story
'==============================================================================
Private Sub FindRunsInFont(ByVal oDoc As Word.Document, ByVal sFont As String, ByVal nBooks As Long, _
                           ByRef bookStarts() As Long, ByRef bookNames() As String, _
                           ByRef nFound As Long, ByRef nResolvedOk As Long, ByRef sHits As String, ByRef nListed As Long)
    Dim oStory As Word.Range
    Dim oRng As Word.Range
    Dim r As Word.Range
    Dim lastStart As Long
    Dim guard As Long
    Dim sLoc As String
    Dim sParaStyle As String
    Dim sCharStyle As String
    Dim sResolved As String

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
                    On Error Resume Next
                    sResolved = r.Font.Name
                    On Error GoTo 0
                    If IsAllowedRunFont(sResolved) Then
                        nResolvedOk = nResolvedOk + 1
                    Else
                      nFound = nFound + 1
                      If nListed < MAX_LISTED_HITS Then
                        nListed = nListed + 1
                        sParaStyle = "": sCharStyle = ""
                        On Error Resume Next
                        sParaStyle = r.Paragraphs(1).Style.NameLocal
                        sCharStyle = r.CharacterStyle.NameLocal
                        On Error GoTo 0
                        sLoc = ""
                        If oRng.StoryType = wdMainTextStory Then sLoc = VerseRefOf(r.Paragraphs(1), nBooks, bookStarts, bookNames)
                        If Len(sLoc) = 0 Then sLoc = "pos " & r.Start
                        sHits = sHits & sLoc & " | story " & oRng.StoryType & " | " & sParaStyle & " | " & _
                                IIf(Len(sCharStyle) = 0, "(none)", sCharStyle) & " | " & sResolved & " (found as " & sFont & ") | """ & _
                                CleanSnippet(r.Text, 30) & """" & vbCrLf
                      End If
                    End If
                    r.Collapse wdCollapseEnd
                Loop
            End With
            Set oRng = oRng.NextStoryRange
        Loop
    Next oStory
End Sub

' "Leviticus 2:11" from the heading list and the marker-styled characters at the start of a verse paragraph; "" when it is not one.
Private Function VerseRefOf(ByVal oPara As Word.Paragraph, ByVal nBooks As Long, _
                            ByRef bookStarts() As Long, ByRef bookNames() As String) As String
    Dim ps As Long
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim sBook As String
    Dim sCh As String
    Dim sVs As String
    Dim sCs As String
    Dim rc As Word.Range

    If oPara.Style.NameLocal <> "VerseText" Then Exit Function
    ps = oPara.Range.Start
    For i = nBooks To 1 Step -1
        If bookStarts(i) <= ps Then sBook = bookNames(i): Exit For
    Next i
    n = oPara.Range.Characters.Count
    If n > 14 Then n = 14
    For j = 1 To n
        Set rc = oPara.Range.Document.Range(ps + j - 1, ps + j)
        sCs = ""
        On Error Resume Next
        sCs = rc.CharacterStyle.NameLocal
        On Error GoTo 0
        If sCs = "Chapter Verse marker" Then
            sCh = sCh & DigitsOnly(rc.Text)
        ElseIf sCs = "Verse marker" Then
            sVs = sVs & DigitsOnly(rc.Text)
        ElseIf Len(sCh) > 0 Or Len(sVs) > 0 Then
            Exit For
        End If
    Next j
    If Len(sCh) > 0 And Len(sVs) > 0 Then VerseRefOf = Trim$(sBook & " " & sCh & ":" & sVs)
End Function

'==============================================================================
' Small helpers
'==============================================================================
Private Function DigitsOnly(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then DigitsOnly = DigitsOnly & ch
    Next i
End Function

' ASCII-only, single-line, trimmed, at most nMax characters (the report file is ASCII like the other rpt\ files).
Private Function CleanSnippet(ByVal s As String, ByVal nMax As Long) As String
    Dim i As Long
    Dim c As Long
    Dim sOut As String

    If Len(s) > nMax Then s = Left$(s, nMax)
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536
        If c >= 32 And c <= 126 Then
            sOut = sOut & Mid$(s, i, 1)
        ElseIf c = 13 Or c = 10 Or c = 11 Or c = 9 Then
            sOut = sOut & " "
        Else
            sOut = sOut & "?"
        End If
    Next i
    CleanSnippet = Trim$(sOut)
End Function

Private Sub WriteReportFile(ByVal sPath As String, ByVal sContent As String)
    Dim oFSO As Object
    Dim oStream As Object

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.CreateTextFile(sPath, True, False)    ' ASCII
    oStream.Write sContent
    oStream.Close
End Sub
