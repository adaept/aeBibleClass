Attribute VB_Name = "basRwbFrontMatter"
Option Explicit
Option Compare Text

' Classification: docm clean-up and export taxonomy (RT-22a). Not used by runtime ribbon navigation.
'
' basRwbFrontMatter - the styles and the text of the RWB front matter (title page, epigraph, Introduction, acknowledgments, signature,
' copyright page with the CC0 statement) of the scripture-only publication. Owner decisions (2026-11-11): RT-20 route A (the styles are
' KEPT by the scripture-only export, GetScriptureOnlyStyles; the front matter ships in 1.1.1), a character style for italics, the CC0
' paragraph on the copyright page, three sections (title page | Introduction to signature | copyright page) so the title and copyright
' pages can be centred vertically with Word's own Page Setup > Layout > Vertical alignment, no page numbers on these pages.
' The text comes from "RWB Introduction.docx" (accepted by the owner, no tracked changes), converted to frontmatter\rwb-front-matter.txt:
'   @section<TAB>center|top       starts a section (a Next Page break between sections) with that vertical alignment
'   STYLE<TAB>text                one paragraph; {i}...{/i} inside the text = character style RWBFrontItalic;
'                                 {justify} at the start of the text = direct justified alignment of that paragraph (the CC0 paragraph,
'                                 as the owner set it in the docm on 2026-11-11)
'   # ...                         a comment
' Its single style ("BODY") and its direct Times New Roman runs are NOT carried over: InsertFrontMatter inserts the text paragraph by
' paragraph into these styles.
'
' Six paragraph styles and one character style, all starting with RWBFront so that they never clash with the Study Bible styles (Title,
' Acknowledgments, Introduction ... are all on the STRIP list). Heading 1 is not used on purpose: the export counts 66 Heading 1 books and
' the Bible Index depends on it.
'
' Fonts: Liberation Serif for the six paragraph styles (the free metric twin of Times New Roman, the look of the draft and of the other
' front styles). Latin, Other and Complex-script slots are set to the same font; the East Asian slot is set by FixStyleSlotFonts of
' basDocmHygiene (Word refuses a Latin-only font there), so run that after CreateFrontMatterStyles. The character style RWBFrontItalic
' sets only Italic: it keeps the font of the paragraph it is used in.
'
' Page layout (FU-34): a section break copies the page setup of the section it splits, which here is the Study Bible's own US Letter title
' page. InsertFrontMatter therefore copies the page setup of the Bible Index section (the first section holding a BibleIndexEyebrow
' paragraph), the layout of the scripture-only pages, onto every front section.
'
' Entry points (Immediate window, one statement each; the document name is optional, default ActiveDocument):
'   ? ListFrontMatterStyles            read-only: prints one line per style (present or missing, font, size); returns the number missing
'   ? CreateFrontMatterStyles          creates the missing styles with their specification; returns the number created (-1 = failed);
'                                      never touches text (compares the text length before and after)
'   ? CountFrontMatterStylesOutsideSpec   styles that exist but differ from the specification (expected 0 once created)
'   ? InsertFrontMatter("C:\...\frontmatter\rwb-front-matter.txt")
'                                      inserts the text at the very top as new leading sections; refuses when section 1 already holds front
'                                      matter or a style is missing or differs; returns the number of paragraphs inserted (-1 = failed).
'                                      Run it on a COPY first.
' Reference: basTEST_aeBibleConfig.RUN_TAXONOMY_STYLES carries the same specification as AuditOneStyle rows;
' rvw\Plan_rwb_1_1_1_release_train_2026-11-08.md rows RT-20 to RT-24.

Public Const STYLE_FRONT_TITLE As String = "RWBFrontTitle"
Public Const STYLE_FRONT_EPIGRAPH As String = "RWBFrontEpigraph"
Public Const STYLE_FRONT_HEADING As String = "RWBFrontHeading"
Public Const STYLE_FRONT_BODY As String = "RWBFrontBody"
Public Const STYLE_FRONT_SIGNATURE As String = "RWBFrontSignature"
Public Const STYLE_FRONT_COPYRIGHT As String = "RWBFrontCopyright"
Public Const STYLE_FRONT_ITALIC As String = "RWBFrontItalic"

' Which version of this module is loaded (changes with every edit; printed by DescribeFrontMatterStyles).
Public Const FRONT_MATTER_VERSION As String = "2026-11-11 k ({justify} marker)"

Private Const FRONT_FONT As String = "Liberation Serif"
Private Const BASE_CHARACTER_STYLE As String = "Default Paragraph Font"
Private Const INDEX_HOST_STYLE As String = "BibleIndexEyebrow"
Private Const LANG_ENGLISH_US As Long = 1033
Private Const ALIGN_CENTER As Long = 1
Private Const ALIGN_RIGHT As Long = 2
Private Const ALIGN_JUSTIFY As Long = 3
Private Const MARK_ITALIC_OPEN As String = "{i}"
Private Const MARK_ITALIC_CLOSE As String = "{/i}"
Private Const MARK_JUSTIFY As String = "{justify}"
Private Const DIRECTIVE_SECTION As String = "@section"
Private Const VALIGN_TOP_NAME As String = "top"
Private Const VALIGN_CENTER_NAME As String = "center"
Private Const MAX_ITALIC_SPANS As Long = 64
Private Const MAX_FRONT_SECTIONS As Long = 12
Private Const MAX_FOREIGN_HITS As Long = 1000
Private m_LastError As String

' The text of the last error ("" = none), for callers that cannot see the Immediate window.
Public Function FrontMatterLastError() As String
    FrontMatterLastError = m_LastError
End Function

' The six paragraph styles, in the order of the Word style gallery (the order used in GetApprovedStyles).
Public Function GetFrontMatterStyles() As Variant
    GetFrontMatterStyles = Array(STYLE_FRONT_TITLE, STYLE_FRONT_EPIGRAPH, STYLE_FRONT_HEADING, _
                                 STYLE_FRONT_BODY, STYLE_FRONT_SIGNATURE, STYLE_FRONT_COPYRIGHT)
End Function

' The character styles of the front matter (the one for italics).
Public Function GetFrontMatterCharacterStyles() As Variant
    GetFrontMatterCharacterStyles = Array(STYLE_FRONT_ITALIC)
End Function

' One specification per paragraph style: Array(size, alignment, space before, space after, bold, italic, Next-style name).
' Single line spacing, no first-line indent and no page break before for all six (the same values the taxonomy audit checks; the page
' breaks are the section breaks). The title has no space before: its page is centred vertically by the section, so a gap would shift it.
' The epigraph is not italic itself: only its quotation is, through RWBFrontItalic ("(Prov 4:18)" stays roman). The Introduction heading
' has a gap of 72 points (one inch) above it: the Introduction runs over several pages, so it is not centred, it just starts lower.
Private Function SpecFor(ByVal sStyle As String) As Variant
    Select Case sStyle
        Case STYLE_FRONT_TITLE:     SpecFor = Array(36, ALIGN_CENTER, 0, 12, False, False, STYLE_FRONT_EPIGRAPH)
        Case STYLE_FRONT_EPIGRAPH:  SpecFor = Array(12, ALIGN_CENTER, 12, 30, False, False, STYLE_FRONT_HEADING)
        Case STYLE_FRONT_HEADING:   SpecFor = Array(22, ALIGN_CENTER, 72, 12, False, False, STYLE_FRONT_BODY)
        Case STYLE_FRONT_BODY:      SpecFor = Array(12, ALIGN_JUSTIFY, 0, 12, False, False, STYLE_FRONT_BODY)
        Case STYLE_FRONT_SIGNATURE: SpecFor = Array(12, ALIGN_RIGHT, 12, 0, False, False, STYLE_FRONT_BODY)
        Case STYLE_FRONT_COPYRIGHT: SpecFor = Array(10, ALIGN_CENTER, 0, 6, False, False, STYLE_FRONT_COPYRIGHT)
        Case Else:                  SpecFor = Empty
    End Select
End Function

Public Function ListFrontMatterStyles(Optional ByVal sDocName As String = "") As Long
    Dim oDoc As Word.Document
    Dim vNames As Variant
    Dim i As Long
    Dim oStyle As Word.Style
    Dim nMissing As Long

    On Error GoTo PROC_ERR
    m_LastError = ""
    Set oDoc = ResolveDoc(sDocName)
    vNames = AllFrontMatterStyleNames()
    For i = LBound(vNames) To UBound(vNames)
        Set oStyle = FindStyle(oDoc, CStr(vNames(i)))
        If oStyle Is Nothing Then
            nMissing = nMissing + 1
            Debug.Print vNames(i) & " | MISSING"
        Else
            Debug.Print vNames(i) & " | " & oStyle.Font.Name & " | " & oStyle.Font.Size & " | align " & oStyle.ParagraphFormat.Alignment
        End If
    Next i
    Debug.Print "ListFrontMatterStyles: " & nMissing & " of " & (UBound(vNames) - LBound(vNames) + 1) & " missing in " & oDoc.Name
    ListFrontMatterStyles = nMissing
    Exit Function
PROC_ERR:
    m_LastError = "ListFrontMatterStyles: " & Err.Number & " " & Err.Description
    Debug.Print m_LastError
    ListFrontMatterStyles = -1
End Function

Public Function CreateFrontMatterStyles(Optional ByVal sDocName As String = "") As Long
    Dim oDoc As Word.Document
    Dim vNames As Variant
    Dim i As Long
    Dim nCreated As Long
    Dim nLenBefore As Long
    Dim oPrevious As Word.Document

    On Error GoTo PROC_ERR
    m_LastError = ""
    Set oDoc = ResolveDoc(sDocName)
    ' Word resolves a Next style in the ACTIVE document (error 5171 when another document is active, as in the COM harness
    ' where a scratch document holds the modules), so the target is activated for the run and the previous one restored.
    Set oPrevious = ActiveDocument
    If Not oDoc Is oPrevious Then oDoc.Activate
    nLenBefore = Len(oDoc.Content.Text)
    vNames = GetFrontMatterStyles()
    ' All styles are created before any specification is applied: a style's Next style may be one created later in the list
    ' (error 5834 on the first live run, 2026-11-11).
    For i = LBound(vNames) To UBound(vNames)
        If FindStyle(oDoc, CStr(vNames(i))) Is Nothing Then
            oDoc.Styles.Add name:=CStr(vNames(i)), Type:=wdStyleTypeParagraph
            nCreated = nCreated + 1
        End If
    Next i
    If FindStyle(oDoc, STYLE_FRONT_ITALIC) Is Nothing Then
        oDoc.Styles.Add name:=STYLE_FRONT_ITALIC, Type:=wdStyleTypeCharacter
        nCreated = nCreated + 1
    End If
    For i = LBound(vNames) To UBound(vNames)
        ApplySpec oDoc, oDoc.Styles(CStr(vNames(i)))
    Next i
    oDoc.Styles(STYLE_FRONT_ITALIC).Font.Italic = True
    If Not oDoc Is oPrevious Then oPrevious.Activate
    If Len(oDoc.Content.Text) <> nLenBefore Then
        m_LastError = "CreateFrontMatterStyles: the document text length changed (" & nLenBefore & " -> " & Len(oDoc.Content.Text) & ")."
        Debug.Print m_LastError
        CreateFrontMatterStyles = -1
        Exit Function
    End If
    Debug.Print "CreateFrontMatterStyles: " & nCreated & " style(s) created, the rest only re-set. Text length unchanged (" & nLenBefore & ")."
    CreateFrontMatterStyles = nCreated
    Exit Function
PROC_ERR:
    m_LastError = "CreateFrontMatterStyles: " & Err.Number & " " & Err.Description
    On Error Resume Next
    If Not oPrevious Is Nothing Then If Not oDoc Is oPrevious Then oPrevious.Activate
    On Error GoTo 0
    Debug.Print m_LastError
    CreateFrontMatterStyles = -1
End Function

Public Function CountFrontMatterStylesOutsideSpec(Optional ByVal sDocName As String = "") As Long
    Dim oDoc As Word.Document
    Dim vNames As Variant
    Dim i As Long
    Dim oStyle As Word.Style
    Dim nOutside As Long

    On Error GoTo PROC_ERR
    m_LastError = ""
    Set oDoc = ResolveDoc(sDocName)
    vNames = GetFrontMatterStyles()
    For i = LBound(vNames) To UBound(vNames)
        Set oStyle = FindStyle(oDoc, CStr(vNames(i)))
        If oStyle Is Nothing Then
            nOutside = nOutside + 1
        ElseIf Not MatchesSpec(oStyle, SpecFor(CStr(vNames(i)))) Then
            nOutside = nOutside + 1
            Debug.Print vNames(i) & " | differs from the specification"
        End If
    Next i
    Set oStyle = FindStyle(oDoc, STYLE_FRONT_ITALIC)
    If oStyle Is Nothing Then
        nOutside = nOutside + 1
    ElseIf Not MatchesItalicSpec(oStyle) Then
        nOutside = nOutside + 1
        Debug.Print STYLE_FRONT_ITALIC & " | differs from the specification"
    End If
    Debug.Print "CountFrontMatterStylesOutsideSpec: " & nOutside
    CountFrontMatterStylesOutsideSpec = nOutside
    Exit Function
PROC_ERR:
    m_LastError = "CountFrontMatterStylesOutsideSpec: " & Err.Number & " " & Err.Description
    Debug.Print m_LastError
    CountFrontMatterStylesOutsideSpec = -1
End Function

' Read-only diagnostic: for each paragraph style, every property that MatchesSpec compares, with the wanted value beside it.
' Use it when CountFrontMatterStylesOutsideSpec lists a style: it shows which property is the one that differs.
Public Sub DescribeFrontMatterStyles(Optional ByVal sDocName As String = "")
    Dim oDoc As Word.Document
    Dim vNames As Variant
    Dim i As Long
    Dim oStyle As Word.Style
    Dim vSpec As Variant

    On Error GoTo PROC_ERR
    Set oDoc = ResolveDoc(sDocName)
    Debug.Print "basRwbFrontMatter version " & FRONT_MATTER_VERSION & " - " & oDoc.Name
    vNames = GetFrontMatterStyles()
    For i = LBound(vNames) To UBound(vNames)
        Set oStyle = FindStyle(oDoc, CStr(vNames(i)))
        If oStyle Is Nothing Then
            Debug.Print vNames(i) & " | MISSING"
        Else
            vSpec = SpecFor(CStr(vNames(i)))
            With oStyle
                Debug.Print vNames(i) & " | fonts " & .Font.NameAscii & "/" & .Font.NameOther & "/" & .Font.NameBi & _
                    " | size " & .Font.Size & " (want " & vSpec(0) & ")" & _
                    " | align " & .ParagraphFormat.Alignment & " (want " & vSpec(1) & ")" & _
                    " | before " & .ParagraphFormat.SpaceBefore & " (want " & vSpec(2) & ")" & _
                    " | after " & .ParagraphFormat.SpaceAfter & " (want " & vSpec(3) & ")" & _
                    " | bold " & .Font.Bold & " | italic " & .Font.Italic & " (want " & vSpec(5) & ")" & _
                    " | page break before " & .ParagraphFormat.PageBreakBefore & " (want False)" & _
                    " | base [" & CStr(.BaseStyle) & "] | rule " & .ParagraphFormat.LineSpacingRule & _
                    " | first indent " & .ParagraphFormat.FirstLineIndent & " | spec match " & MatchesSpec(oStyle, vSpec)
            End With
        End If
    Next i
    Set oStyle = FindStyle(oDoc, STYLE_FRONT_ITALIC)
    If oStyle Is Nothing Then
        Debug.Print STYLE_FRONT_ITALIC & " | MISSING"
    Else
        Debug.Print STYLE_FRONT_ITALIC & " | italic " & oStyle.Font.Italic & " | base [" & CStr(oStyle.BaseStyle) & "] | spec match " & MatchesItalicSpec(oStyle)
    End If
    Exit Sub
PROC_ERR:
    Debug.Print "DescribeFrontMatterStyles: " & Err.Number & " " & Err.Description
End Sub

' True when any paragraph of the section is in one of the six RWBFront paragraph styles. Pass 5 of the export uses it to find the front
' sections; InsertFrontMatter uses it to refuse a second insert.
Public Function SectionHoldsFrontStyle(ByVal oSection As Object) As Boolean
    Dim vNames As Variant
    Dim oPara As Object
    Dim i As Long

    vNames = GetFrontMatterStyles()
    For Each oPara In oSection.Range.Paragraphs
        For i = LBound(vNames) To UBound(vNames)
            If StrComp(oPara.Style.NameLocal, CStr(vNames(i)), vbTextCompare) = 0 Then
                SectionHoldsFrontStyle = True
                Exit Function
            End If
        Next i
    Next oPara
End Function

' The number of leading sections that hold an RWBFront paragraph (0 = the document has no RWB front matter).
Public Function CountLeadingFrontSections(ByVal oDoc As Object) As Long
    Dim n As Long

    Do While n < oDoc.Sections.Count And n < MAX_FRONT_SECTIONS
        If Not SectionHoldsFrontStyle(oDoc.Sections(n + 1)) Then Exit Do
        n = n + 1
    Loop
    CountLeadingFrontSections = n
End Function

' Inserts the front matter at the very top of the document as new leading sections. Returns the number of paragraphs inserted, -1 on
' failure (see FrontMatterLastError). Refuses (and changes nothing) when section 1 already holds front matter, when a style is missing or
' differs from its specification, or when the text file is not as expected. Each section ends with a Next Page section break that REPLACES
' the paragraph mark of its last paragraph (no empty carrier paragraph). The page setup of the Bible Index section is copied onto the front
' sections, their vertical alignment comes from the file, they get no page-number footer, and the inserted text gets English (US) and "do
' not check spelling" (owner decision A, FU-32). The document text must grow by exactly the inserted length, else the routine reports it.
Public Function InsertFrontMatter(ByVal sTextFile As String, Optional ByVal sDocName As String = "") As Long
    Dim oDoc As Word.Document
    Dim oPrevious As Word.Document
    Dim oPara As Word.Paragraph
    Dim vLines As Variant
    Dim aStarts(0 To MAX_ITALIC_SPANS - 1) As Long
    Dim aEnds(0 To MAX_ITALIC_SPANS - 1) As Long
    Dim aSectionAlign() As Long
    Dim aStyle() As String
    Dim aPlain() As String
    Dim aSpans() As String
    Dim aSectionOf() As Long
    Dim aJustify() As Boolean
    Dim nParas As Long
    Dim nSections As Long
    Dim nSpans As Long
    Dim i As Long
    Dim k As Long
    Dim lPos As Long
    Dim lStart As Long
    Dim nLenBefore As Long
    Dim nGrew As Long
    Dim bLast As Boolean
    Dim sStyle As String
    Dim sText As String
    Dim sArg As String

    On Error GoTo PROC_ERR
    m_LastError = ""
    Set oDoc = ResolveDoc(sDocName)
    Set oPrevious = ActiveDocument
    If Not oDoc Is oPrevious Then oDoc.Activate

    If SectionHoldsFrontStyle(oDoc.Sections(1)) Then
        m_LastError = "InsertFrontMatter: section 1 of " & oDoc.Name & " already holds front matter; nothing inserted."
        GoTo PROC_FAIL
    End If
    If CountFrontMatterStylesOutsideSpec(oDoc.Name) <> 0 Then
        m_LastError = "InsertFrontMatter: a front-matter style is missing or differs from its specification (run CreateFrontMatterStyles); nothing inserted."
        GoTo PROC_FAIL
    End If
    m_LastError = ""

    ' Phase 1: read and check the whole text file before the document is touched.
    vLines = Split(Replace(ReadUtf8Text(sTextFile), vbCrLf, vbLf), vbLf)
    ReDim aSectionAlign(0 To MAX_FRONT_SECTIONS - 1)
    ReDim aStyle(0 To UBound(vLines))
    ReDim aPlain(0 To UBound(vLines))
    ReDim aSpans(0 To UBound(vLines))
    ReDim aSectionOf(0 To UBound(vLines))
    ReDim aJustify(0 To UBound(vLines))
    For i = LBound(vLines) To UBound(vLines)
        If Len(vLines(i)) = 0 Or Left$(vLines(i), 1) = "#" Then
            ' blank line or comment
        ElseIf Left$(vLines(i), Len(DIRECTIVE_SECTION)) = DIRECTIVE_SECTION Then
            If nSections >= MAX_FRONT_SECTIONS Then
                m_LastError = "InsertFrontMatter: more than " & MAX_FRONT_SECTIONS & " sections in the text file."
                GoTo PROC_FAIL
            End If
            k = InStr(1, vLines(i), vbTab, vbBinaryCompare)
            sArg = IIf(k = 0, "", LCase$(Trim$(Mid$(vLines(i), k + 1))))
            If sArg = VALIGN_CENTER_NAME Then
                aSectionAlign(nSections) = wdAlignVerticalCenter
            ElseIf sArg = VALIGN_TOP_NAME Then
                aSectionAlign(nSections) = wdAlignVerticalTop
            Else
                m_LastError = "InsertFrontMatter: line " & (i + 1) & ": " & DIRECTIVE_SECTION & " needs a tab and " & VALIGN_TOP_NAME & " or " & VALIGN_CENTER_NAME & "."
                GoTo PROC_FAIL
            End If
            nSections = nSections + 1
        Else
            If nSections = 0 Then
                m_LastError = "InsertFrontMatter: line " & (i + 1) & ": a paragraph before the first " & DIRECTIVE_SECTION & " line."
                GoTo PROC_FAIL
            End If
            k = InStr(1, vLines(i), vbTab, vbBinaryCompare)
            If k = 0 Then
                m_LastError = "InsertFrontMatter: line " & (i + 1) & " of the text file has no tab after the style name."
                GoTo PROC_FAIL
            End If
            sStyle = Left$(vLines(i), k - 1)
            sText = Mid$(vLines(i), k + 1)
            aJustify(nParas) = (Left$(sText, Len(MARK_JUSTIFY)) = MARK_JUSTIFY)
            If aJustify(nParas) Then sText = Mid$(sText, Len(MARK_JUSTIFY) + 1)
            If Not IsFrontMatterParagraphStyle(sStyle) Then
                m_LastError = "InsertFrontMatter: line " & (i + 1) & " names the style """ & sStyle & """, which is not a front-matter paragraph style."
                GoTo PROC_FAIL
            End If
            aStyle(nParas) = sStyle
            ' The owner's hand-set optional hyphens are U+00AD in the text file; Word stores an optional hyphen as Chr(31) in a range's text (the docx
            ' element is w:softHyphen). The export keeps them in the front matter (Pass4b skips the front sections, 2026-11-12).
            aPlain(nParas) = Replace(StripItalicMarkers(sText, aStarts, aEnds, nSpans), ChrW(173), Chr$(31))
            aSpans(nParas) = ""
            For k = 0 To nSpans - 1
                aSpans(nParas) = aSpans(nParas) & aStarts(k) & "-" & aEnds(k) & ";"
            Next k
            aSectionOf(nParas) = nSections
            nParas = nParas + 1
        End If
    Next i
    If nParas = 0 Then
        m_LastError = "InsertFrontMatter: the text file holds no paragraph; nothing inserted."
        GoTo PROC_FAIL
    End If

    ' Phase 2: insert. The last paragraph of a section is inserted WITHOUT a paragraph mark and the Next Page section break is put right
    ' after its text, so the break terminates that paragraph (a break inserted at a paragraph mark leaves the mark behind: an empty
    ' paragraph at the start of the next section, found on the first run, 2026-11-11).
    nLenBefore = Len(oDoc.Content.Text)
    lPos = 0
    For i = 0 To nParas - 1
        bLast = (i = nParas - 1)
        If Not bLast Then bLast = (aSectionOf(i + 1) <> aSectionOf(i))
        lStart = lPos
        If bLast Then
            oDoc.Range(lPos, lPos).InsertAfter aPlain(i)
            oDoc.Range(lPos + Len(aPlain(i)), lPos + Len(aPlain(i))).InsertBreak wdSectionBreakNextPage
        Else
            oDoc.Range(lPos, lPos).InsertAfter aPlain(i) & vbCr
        End If
        lPos = lPos + Len(aPlain(i)) + 1
        Set oPara = oDoc.Range(lStart, lStart).Paragraphs(1)
        oPara.Style = oDoc.Styles(aStyle(i))
        If aJustify(i) Then oPara.Alignment = wdAlignParagraphJustify
        ApplyItalicSpans oDoc, lStart, aSpans(i)
    Next i
    If oDoc.Sections.Count < nSections + 1 Then
        m_LastError = "InsertFrontMatter: " & nSections & " section break(s) were expected to give at least " & (nSections + 1) & " sections, the document has " & oDoc.Sections.Count & "."
        GoTo PROC_FAIL
    End If

    ApplyFrontSectionLayout oDoc, nSections, aSectionAlign
    If m_LastError <> "" Then GoTo PROC_FAIL

    With oDoc.Range(0, lPos)
        .LanguageID = LANG_ENGLISH_US
        .NoProofing = True
        .LanguageDetected = False
    End With

    nGrew = Len(oDoc.Content.Text) - nLenBefore
    If nGrew <> lPos Then
        m_LastError = "InsertFrontMatter: the document text grew by " & nGrew & " characters, expected " & lPos & " (the text; each section break terminates its last paragraph)."
        GoTo PROC_FAIL
    End If
    If Not oDoc Is oPrevious Then oPrevious.Activate
    Debug.Print "InsertFrontMatter: " & nParas & " paragraph(s) in " & nSections & " section(s) inserted, " & lPos & " characters, now " & oDoc.Sections.Count & " section(s) in " & oDoc.Name & "."
    InsertFrontMatter = nParas
    Exit Function
PROC_ERR:
    m_LastError = "InsertFrontMatter: " & Err.Number & " " & Err.Description
PROC_FAIL:
    On Error Resume Next
    If Not oPrevious Is Nothing Then If Not oDoc Is oPrevious Then oPrevious.Activate
    On Error GoTo 0
    Debug.Print m_LastError
    InsertFrontMatter = -1
End Function

' Applies RWBFrontItalic to the "start-end;" spans (offsets within the paragraph that starts at lParaStart).
Private Sub ApplyItalicSpans(ByVal oDoc As Word.Document, ByVal lParaStart As Long, ByVal sSpans As String)
    Dim vSpan As Variant
    Dim vPair As Variant

    For Each vSpan In Split(sSpans, ";")
        If Len(vSpan) > 0 Then
            vPair = Split(vSpan, "-")
            oDoc.Range(lParaStart + CLng(vPair(0)), lParaStart + CLng(vPair(1))).Style = oDoc.Styles(STYLE_FRONT_ITALIC)
        End If
    Next vSpan
End Sub

' Gives the nFront leading sections the page setup of the Bible Index section, a New Page start, one column, no different first page,
' the vertical alignment from the text file, and no page-number footer (the footers of the first section are emptied; the later front
' sections link to it). Sets m_LastError when the Bible Index section cannot be found.
Private Sub ApplyFrontSectionLayout(ByVal oDoc As Word.Document, ByVal nFront As Long, ByRef aSectionAlign() As Long)
    Dim oHost As Word.Section
    Dim i As Long
    Dim idx As Long

    Set oHost = FindIndexHostSection(oDoc, nFront)
    If oHost Is Nothing Then
        m_LastError = "InsertFrontMatter: no section with a " & INDEX_HOST_STYLE & " paragraph after the " & nFront & " front sections; the page setup cannot be copied."
        Exit Sub
    End If
    For i = 1 To nFront
        basSectionPageSetup.CopySectionPageSetup oHost, oDoc.Sections(i)
        With oDoc.Sections(i).PageSetup
            .SectionStart = wdSectionNewPage
            .DifferentFirstPageHeaderFooter = False
            .TextColumns.SetCount 1
            .TextColumns.LineBetween = False
            .VerticalAlignment = aSectionAlign(i - 1)
        End With
    Next i
    For idx = wdHeaderFooterPrimary To wdHeaderFooterEvenPages
        oDoc.Sections(1).Footers(idx).Range.Delete
        For i = 2 To nFront
            oDoc.Sections(i).Footers(idx).LinkToPrevious = True
        Next i
    Next idx
End Sub

' The first section after the nFront front sections that holds a BibleIndexEyebrow paragraph (found by style, not by number).
Private Function FindIndexHostSection(ByVal oDoc As Word.Document, ByVal nFront As Long) As Word.Section
    Dim k As Long
    Dim oPara As Word.Paragraph

    For k = nFront + 1 To oDoc.Sections.Count
        If k > nFront + MAX_FRONT_SECTIONS Then Exit For
        For Each oPara In oDoc.Sections(k).Range.Paragraphs
            If StrComp(oPara.Style.NameLocal, INDEX_HOST_STYLE, vbTextCompare) = 0 Then
                Set FindIndexHostSection = oDoc.Sections(k)
                Exit Function
            End If
        Next oPara
    Next k
End Function

' The paragraph style names and the character style name together (for the list and the read-only checks).
Private Function AllFrontMatterStyleNames() As Variant
    Dim vNames As Variant
    Dim vAll() As String
    Dim i As Long

    vNames = GetFrontMatterStyles()
    ReDim vAll(LBound(vNames) To UBound(vNames) + 1)
    For i = LBound(vNames) To UBound(vNames)
        vAll(i) = vNames(i)
    Next i
    vAll(UBound(vNames) + 1) = STYLE_FRONT_ITALIC
    AllFrontMatterStyleNames = vAll
End Function

Private Function IsFrontMatterParagraphStyle(ByVal sName As String) As Boolean
    Dim vNames As Variant
    Dim i As Long

    vNames = GetFrontMatterStyles()
    For i = LBound(vNames) To UBound(vNames)
        If StrComp(sName, CStr(vNames(i)), vbTextCompare) = 0 Then IsFrontMatterParagraphStyle = True: Exit Function
    Next i
End Function

' The text without the {i}...{/i} markers; aStarts/aEnds receive the 0-based [start, end) offsets of each italic span in the plain text.
Private Function StripItalicMarkers(ByVal sText As String, ByRef aStarts() As Long, ByRef aEnds() As Long, ByRef nSpans As Long) As String
    Dim sPlain As String
    Dim lOpen As Long
    Dim lClose As Long

    nSpans = 0
    Do
        lOpen = InStr(1, sText, MARK_ITALIC_OPEN, vbBinaryCompare)
        If lOpen = 0 Then
            sPlain = sPlain & sText
            Exit Do
        End If
        lClose = InStr(lOpen + Len(MARK_ITALIC_OPEN), sText, MARK_ITALIC_CLOSE, vbBinaryCompare)
        If lClose = 0 Then Err.Raise vbObjectError + 1, "StripItalicMarkers", "an " & MARK_ITALIC_OPEN & " has no " & MARK_ITALIC_CLOSE
        If nSpans >= MAX_ITALIC_SPANS Then Err.Raise vbObjectError + 2, "StripItalicMarkers", "more than " & MAX_ITALIC_SPANS & " italic spans in one paragraph"
        sPlain = sPlain & Left$(sText, lOpen - 1)
        aStarts(nSpans) = Len(sPlain)
        sPlain = sPlain & Mid$(sText, lOpen + Len(MARK_ITALIC_OPEN), lClose - lOpen - Len(MARK_ITALIC_OPEN))
        aEnds(nSpans) = Len(sPlain)
        nSpans = nSpans + 1
        sText = Mid$(sText, lClose + Len(MARK_ITALIC_CLOSE))
    Loop
    StripItalicMarkers = sPlain
End Function

Private Function ReadUtf8Text(ByVal sPath As String) As String
    Dim oStream As Object

    Set oStream = CreateObject("ADODB.Stream")
    oStream.Type = 2
    oStream.Charset = "utf-8"
    oStream.Open
    oStream.LoadFromFile sPath
    ReadUtf8Text = oStream.ReadText
    oStream.Close
End Function

Private Sub ApplySpec(ByVal oDoc As Word.Document, ByVal oStyle As Word.Style)
    Dim vSpec As Variant

    vSpec = SpecFor(oStyle.NameLocal)
    ' No base style, like the other front styles (the taxonomy audit expects "").
    oStyle.BaseStyle = ""
    ' Next style BEFORE the size: setting it afterwards resets a size of 12 to Normal's 9 (probe, live run 2026-11-11).
    ' The Style OBJECT of the target document, not its name: a name is resolved in the ACTIVE document (error 5171 when another
    ' document is active, as in the COM harness).
    oStyle.NextParagraphStyle = oDoc.Styles(CStr(vSpec(6)))
    With oStyle.Font
        .Name = FRONT_FONT
        .NameAscii = FRONT_FONT
        .NameOther = FRONT_FONT
        .NameBi = FRONT_FONT
        .Size = vSpec(0)
        .Bold = vSpec(4)
        .Italic = vSpec(5)
    End With
    With oStyle.ParagraphFormat
        .Alignment = vSpec(1)
        .FirstLineIndent = 0
        .LeftIndent = 0
        .LineSpacingRule = wdLineSpaceSingle
        .SpaceBefore = vSpec(2)
        .SpaceAfter = vSpec(3)
        .PageBreakBefore = False
    End With
    oStyle.LanguageID = LANG_ENGLISH_US
    oStyle.NoProofing = True
End Sub

Private Function MatchesSpec(ByVal oStyle As Word.Style, ByVal vSpec As Variant) As Boolean
    With oStyle
        MatchesSpec = StrComp(.Font.NameAscii, FRONT_FONT, vbTextCompare) = 0 _
            And StrComp(.Font.NameOther, FRONT_FONT, vbTextCompare) = 0 _
            And StrComp(.Font.NameBi, FRONT_FONT, vbTextCompare) = 0 _
            And .Font.Size = vSpec(0) _
            And .ParagraphFormat.Alignment = vSpec(1) _
            And .ParagraphFormat.SpaceBefore = vSpec(2) _
            And .ParagraphFormat.SpaceAfter = vSpec(3) _
            And CBool(.Font.Bold) = CBool(vSpec(4)) _
            And CBool(.Font.Italic) = CBool(vSpec(5)) _
            And Not CBool(.ParagraphFormat.PageBreakBefore) _
            And Len(CStr(.BaseStyle)) = 0 _
            And .ParagraphFormat.FirstLineIndent = 0 _
            And .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
    End With
End Function

' The character style: italic, based on Default Paragraph Font, nothing else (it keeps the font of the paragraph it is used in).
Private Function MatchesItalicSpec(ByVal oStyle As Word.Style) As Boolean
    MatchesItalicSpec = CBool(oStyle.Font.Italic) And StrComp(CStr(oStyle.BaseStyle), BASE_CHARACTER_STYLE, vbTextCompare) = 0
End Function

Private Function FindStyle(ByVal oDoc As Word.Document, ByVal sName As String) As Word.Style
    On Error Resume Next
    Set FindStyle = oDoc.Styles(sName)
    On Error GoTo 0
End Function

Private Function ResolveDoc(ByVal sDocName As String) As Word.Document
    If Len(sDocName) = 0 Then
        Set ResolveDoc = ActiveDocument
    Else
        Set ResolveDoc = Documents(sDocName)
    End If
End Function

' ==========================================================================
' Test 99 (RT-24, 2026-11-12): the RWB front matter and the Study Bible are kept apart
' ==========================================================================
' Returns the number of violations (expected 0; -1 when the check itself failed). A document without front matter has none. It counts:
'   - a paragraph in a leading front section that is not in one of the six RWBFront paragraph styles (Study Bible text in the Foreword);
'   - a footnote, endnote, field, hyperlink, table, inline shape, shape or comment in a leading front section (Study Bible apparatus);
'   - an RWBFront paragraph style or the RWBFrontItalic character style used after the leading front sections (Foreword in the Study Bible).
' The export relies on exactly this split: Pass 4b and Pass 5 treat the leading RWBFront sections as the Foreword and everything after them as
' the Study Bible, and the Node export reads the front matter from the RWBFront styles alone. `ListFrontMatterSeparation` prints each hit.
Public Function CountFrontMatterSeparationViolations(Optional ByVal sDocName As String = "") As Long
    On Error GoTo PROC_ERR
    CountFrontMatterSeparationViolations = SeparationViolations(ResolveDoc(sDocName), False)
    Exit Function
PROC_ERR:
    m_LastError = "CountFrontMatterSeparationViolations: " & Err.Number & " " & Err.Description
    Debug.Print m_LastError
    CountFrontMatterSeparationViolations = -1
End Function

' Read-only: prints every violation CountFrontMatterSeparationViolations counts.
Public Sub ListFrontMatterSeparation(Optional ByVal sDocName As String = "")
    Debug.Print "ListFrontMatterSeparation: " & SeparationViolations(ResolveDoc(sDocName), True) & " violation(s)"
End Sub

Private Function SeparationViolations(ByVal oDoc As Word.Document, ByVal bList As Boolean) As Long
    Dim nFront As Long

    nFront = CountLeadingFrontSections(oDoc)
    SeparationViolations = FrontSectionViolations(oDoc, nFront, bList) + ForeignFrontStyleViolations(oDoc, nFront, bList)
End Function

' Study Bible content inside the leading front sections: a paragraph outside the six styles, or any note, field, link, table, shape or comment.
Private Function FrontSectionViolations(ByVal oDoc As Word.Document, ByVal nFront As Long, ByVal bList As Boolean) As Long
    Dim vKinds As Variant
    Dim oRng As Word.Range
    Dim oPara As Word.Paragraph
    Dim i As Long
    Dim k As Long
    Dim nFound As Long
    Dim nHits As Long

    vKinds = Array("Footnotes", "Endnotes", "Fields", "Hyperlinks", "Tables", "InlineShapes", "ShapeRange", "Comments")
    For i = 1 To nFront
        Set oRng = oDoc.Sections(i).Range
        For Each oPara In oRng.Paragraphs
            If Not IsFrontMatterParagraphStyle(oPara.Style.NameLocal) Then
                nHits = nHits + 1
                If bList Then Debug.Print "front section " & i & " | paragraph in style " & oPara.Style.NameLocal & " | " & Left$(oPara.Range.Text, 40)
            End If
        Next oPara
        For k = LBound(vKinds) To UBound(vKinds)
            nFound = CLng(CallByName(CallByName(oRng, CStr(vKinds(k)), VbGet), "Count", VbGet))
            If nFound > 0 Then
                nHits = nHits + nFound
                If bList Then Debug.Print "front section " & i & " | " & nFound & " x " & vKinds(k)
            End If
        Next k
    Next i
    FrontSectionViolations = nHits
End Function

' Foreword styles used after the leading front sections. Find on the style is fast; walking the paragraphs of the whole docm is not.
Private Function ForeignFrontStyleViolations(ByVal oDoc As Word.Document, ByVal nFront As Long, ByVal bList As Boolean) As Long
    Dim vNames As Variant
    Dim oStyle As Word.Style
    Dim oRng As Word.Range
    Dim lFrom As Long
    Dim i As Long
    Dim nHits As Long
    Dim nGuard As Long

    If nFront >= oDoc.Sections.Count Then Exit Function
    lFrom = 0
    If nFront > 0 Then lFrom = oDoc.Sections(nFront + 1).Range.Start
    vNames = AllFrontMatterStyleNames()
    For i = LBound(vNames) To UBound(vNames)
        Set oStyle = FindStyle(oDoc, CStr(vNames(i)))
        If Not oStyle Is Nothing Then
            Set oRng = oDoc.Range(lFrom, oDoc.Content.End)
            With oRng.Find
                .ClearFormatting
                .Text = ""
                .Style = oStyle
                .Format = True
                .Forward = True
                .Wrap = wdFindStop
                .MatchWildcards = False
                nGuard = 0
                Do While .Execute And nGuard < MAX_FOREIGN_HITS
                    nHits = nHits + 1
                    nGuard = nGuard + 1
                    If bList Then Debug.Print "after the front sections | style " & vNames(i) & " at " & oRng.Start & " | " & Left$(oRng.Text, 40)
                    oRng.Collapse wdCollapseEnd
                Loop
            End With
        End If
    Next i
    ForeignFrontStyleViolations = nHits
End Function
