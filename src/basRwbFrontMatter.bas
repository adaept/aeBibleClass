Attribute VB_Name = "basRwbFrontMatter"
Option Explicit
Option Compare Text

' Classification: docm clean-up and export taxonomy (RT-22a). Not used by runtime ribbon navigation.
'
' basRwbFrontMatter - the six styles of the RWB front matter (title page, epigraph, Introduction, acknowledgments, CC0 statement,
' signature, copyright page) of the scripture-only publication. Owner decision RT-20 (2026-11-11): route A, the styles are KEPT by the
' scripture-only export (GetScriptureOnlyStyles), the front matter ships in 1.1.1. The draft is "RWB Introduction.docx"; its single style
' ("BODY") and its direct Times New Roman runs are NOT carried over - the text is inserted paragraph by paragraph into these styles.
'
' The names start with RWBFront so that they never clash with the Study Bible styles (Title, Acknowledgments, Introduction ... are all on
' the STRIP list). Heading 1 is not used on purpose: the export counts 66 Heading 1 books and the Bible Index depends on it.
'
' Fonts: Liberation Serif for all six (the free metric twin of Times New Roman, the look of the draft and of the other front styles).
' Latin, Other and Complex-script slots are set to the same font; the East Asian slot is set by FixStyleSlotFonts of basDocmHygiene
' (Word refuses a Latin-only font there), so run that after CreateFrontMatterStyles (expected return 0 or the few slots it sets).
'
' Entry points (Immediate window, one statement each; the document name is optional, default ActiveDocument):
'   ? ListFrontMatterStyles            read-only: prints one line per style (present or missing, font, size); returns the number missing
'   ? CreateFrontMatterStyles          creates the missing styles with their specification; returns the number created (-1 = failed);
'                                      never touches text (compares the text length before and after)
'   ? CountFrontMatterStylesOutsideSpec   styles that exist but differ from the specification (expected 0 once created)
' Reference: basTEST_aeBibleConfig.RUN_TAXONOMY_STYLES carries the same specification as AuditOneStyle rows;
' rvw\Plan_rwb_1_1_1_release_train_2026-11-08.md rows RT-20 to RT-24.

Public Const STYLE_FRONT_TITLE As String = "RWBFrontTitle"
Public Const STYLE_FRONT_EPIGRAPH As String = "RWBFrontEpigraph"
Public Const STYLE_FRONT_HEADING As String = "RWBFrontHeading"
Public Const STYLE_FRONT_BODY As String = "RWBFrontBody"
Public Const STYLE_FRONT_SIGNATURE As String = "RWBFrontSignature"
Public Const STYLE_FRONT_COPYRIGHT As String = "RWBFrontCopyright"

' Which version of this module is loaded (changes with every edit; printed by DescribeFrontMatterStyles).
Public Const FRONT_MATTER_VERSION As String = "2026-11-11 h (target document activated)"

Private Const FRONT_FONT As String = "Liberation Serif"
Private Const LANG_ENGLISH_US As Long = 1033
Private Const ALIGN_CENTER As Long = 1
Private Const ALIGN_RIGHT As Long = 2
Private Const ALIGN_JUSTIFY As Long = 3
Private m_LastError As String

' The text of the last error ("" = none), for callers that cannot see the Immediate window.
Public Function FrontMatterLastError() As String
    FrontMatterLastError = m_LastError
End Function

' The six names, in the order of the Word style gallery (the order used in GetApprovedStyles).
Public Function GetFrontMatterStyles() As Variant
    GetFrontMatterStyles = Array(STYLE_FRONT_TITLE, STYLE_FRONT_EPIGRAPH, STYLE_FRONT_HEADING, _
                                 STYLE_FRONT_BODY, STYLE_FRONT_SIGNATURE, STYLE_FRONT_COPYRIGHT)
End Function

' One specification per style: Array(size, alignment, space before, space after, bold, italic, Next-style name).
' Single line spacing and no first-line indent for all six (the same values the taxonomy audit checks).
Private Function SpecFor(ByVal sStyle As String) As Variant
    Select Case sStyle
        Case STYLE_FRONT_TITLE:     SpecFor = Array(36, ALIGN_CENTER, 144, 12, False, False, STYLE_FRONT_EPIGRAPH)
        Case STYLE_FRONT_EPIGRAPH:  SpecFor = Array(12, ALIGN_CENTER, 12, 30, False, True, STYLE_FRONT_HEADING)
        Case STYLE_FRONT_HEADING:   SpecFor = Array(22, ALIGN_CENTER, 12, 12, False, False, STYLE_FRONT_BODY)
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
    vNames = GetFrontMatterStyles()
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
    Dim oStyle As Word.Style
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
    ' All six are created before any specification is applied: a style's Next style may be one created later in the list
    ' (error 5834 on the first live run, 2026-11-11).
    For i = LBound(vNames) To UBound(vNames)
        If FindStyle(oDoc, CStr(vNames(i))) Is Nothing Then
            oDoc.Styles.Add name:=CStr(vNames(i)), Type:=wdStyleTypeParagraph
            nCreated = nCreated + 1
        End If
    Next i
    For i = LBound(vNames) To UBound(vNames)
        ApplySpec oDoc, oDoc.Styles(CStr(vNames(i)))
    Next i
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
    Dim vSpec As Variant
    Dim nOutside As Long

    On Error GoTo PROC_ERR
    m_LastError = ""
    Set oDoc = ResolveDoc(sDocName)
    vNames = GetFrontMatterStyles()
    For i = LBound(vNames) To UBound(vNames)
        Set oStyle = FindStyle(oDoc, CStr(vNames(i)))
        vSpec = SpecFor(CStr(vNames(i)))
        If oStyle Is Nothing Then
            nOutside = nOutside + 1
        ElseIf Not MatchesSpec(oStyle, vSpec) Then
            nOutside = nOutside + 1
            Debug.Print vNames(i) & " | differs from the specification"
        End If
    Next i
    Debug.Print "CountFrontMatterStylesOutsideSpec: " & nOutside
    CountFrontMatterStylesOutsideSpec = nOutside
    Exit Function
PROC_ERR:
    m_LastError = "CountFrontMatterStylesOutsideSpec: " & Err.Number & " " & Err.Description
    Debug.Print m_LastError
    CountFrontMatterStylesOutsideSpec = -1
End Function

' Read-only diagnostic: for each style, every property that MatchesSpec compares, with the wanted value beside it and a mark where it differs.
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
                    " | base [" & CStr(.BaseStyle) & "] | rule " & .ParagraphFormat.LineSpacingRule & _
                    " | first indent " & .ParagraphFormat.FirstLineIndent & " | spec match " & MatchesSpec(oStyle, vSpec)
            End With
        End If
    Next i
    Exit Sub
PROC_ERR:
    Debug.Print "DescribeFrontMatterStyles: " & Err.Number & " " & Err.Description
End Sub

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
            And Len(CStr(.BaseStyle)) = 0 _
            And .ParagraphFormat.FirstLineIndent = 0 _
            And .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
    End With
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
