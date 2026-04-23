Attribute VB_Name = "basFixDocxRoutines"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

'====================================================================
' DefineBodyTextStyle
' PURPOSE:
'   Creates the BodyText style in the active document if it does not
'   already exist.  BodyText is the semantic replacement for Normal
'   and Plain Text paragraph usage throughout the document.
'
' FORMATTING (matches Normal as measured 2026-04-22):
'   Font:           Carlito 9pt, not bold, not italic
'   Alignment:      Left (wdAlignParagraphLeft)
'   Line spacing:   Single (wdLineSpaceSingle)
'   First indent:   14.4 pt (0.2 inches)
'   Left indent:    0
'   Space before:   0pt
'   Space after:    0pt
'
' USFM mapping:  \p  (standard body paragraph)
'
' RERUN SAFE:
'   If BodyText already exists the routine reports and exits without
'   modifying the existing style definition.
'====================================================================
Public Sub DefineBodyTextStyle()
    On Error GoTo PROC_ERR
    Dim oDoc    As Document
    Dim oStyle  As Word.Style

    Set oDoc = ActiveDocument

    ' Check if style already exists
    On Error Resume Next
    Set oStyle = oDoc.Styles("BodyText")
    On Error GoTo PROC_ERR

    If Not oStyle Is Nothing Then
        Debug.Print "DefineBodyTextStyle: BodyText already exists -- no changes made."
        MsgBox "BodyText style already exists in this document.", _
               vbInformation, "DefineBodyTextStyle"
        GoTo PROC_EXIT
    End If

    ' Create new style -- based on no existing style to avoid Normal cascade
    Set oStyle = oDoc.Styles.Add(name:="BodyText", Type:=wdStyleTypeParagraph)

    With oStyle
        .baseStyle = ""                         ' no cascade from Normal
        .NextParagraphStyle = oStyle            ' BodyText follows BodyText

        With .Font
            .Name = "Carlito"
            .Size = 9
            .Bold = False
            .Italic = False
        End With

        With .ParagraphFormat
            .Alignment = wdAlignParagraphJustify
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing = 10                   ' Exactly 10pt — matches original docm
            .FirstLineIndent = 0                ' no first-line indent on Bible body text
            .LeftIndent = 0
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
    End With

    Debug.Print "DefineBodyTextStyle: BodyText created successfully."
    MsgBox "BodyText style created successfully." & vbCrLf & _
           "Font: Carlito 9pt | No indent | Justified | Spacing: Single", _
           vbInformation, "DefineBodyTextStyle"

PROC_EXIT:
    Set oStyle = Nothing
    Set oDoc = Nothing
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure DefineBodyTextStyle of Module basFixDocxRoutines"
    Resume PROC_EXIT
End Sub

'====================================================================
' DefineBodyTextIndentStyle
' PURPOSE:
'   Creates the BodyTextIndent style in the active document if it does
'   not already exist.  BodyTextIndent is for indented body paragraphs
'   (quoted or subordinate text) — identical to BodyText except for
'   the 0.2" first-line indent.
'
' FORMATTING:
'   Font:           Carlito 9pt, not bold, not italic
'   Alignment:      Justified (wdAlignParagraphJustify)
'   Line spacing:   Single (wdLineSpaceSingle)
'   First indent:   14.4 pt (0.2 inches)
'   Left indent:    0
'   Space before:   0pt
'   Space after:    0pt
'
' USFM mapping:  \pi  (paragraph indented)
'
' RERUN SAFE:
'   If BodyTextIndent already exists the routine reports and exits.
'====================================================================
Public Sub DefineBodyTextIndentStyle()
    On Error GoTo PROC_ERR
    Dim oDoc    As Document
    Dim oStyle  As Word.Style

    Set oDoc = ActiveDocument

    On Error Resume Next
    Set oStyle = oDoc.Styles("BodyTextIndent")
    On Error GoTo PROC_ERR

    If Not oStyle Is Nothing Then
        Debug.Print "DefineBodyTextIndentStyle: BodyTextIndent already exists -- no changes made."
        MsgBox "BodyTextIndent style already exists in this document.", _
               vbInformation, "DefineBodyTextIndentStyle"
        GoTo PROC_EXIT
    End If

    Set oStyle = oDoc.Styles.Add(name:="BodyTextIndent", Type:=wdStyleTypeParagraph)

    With oStyle
        .baseStyle = ""
        .NextParagraphStyle = oDoc.Styles("BodyText")   ' returns to BodyText after indent para

        With .Font
            .Name = "Carlito"
            .Size = 9
            .Bold = False
            .Italic = False
        End With

        With .ParagraphFormat
            .Alignment = wdAlignParagraphJustify
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing = 10                           ' Exactly 10pt — matches BodyText
            .FirstLineIndent = 14.4                     ' 0.2 inches in points
            .LeftIndent = 0
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
    End With

    Debug.Print "DefineBodyTextIndentStyle: BodyTextIndent created successfully."
    MsgBox "BodyTextIndent style created successfully." & vbCrLf & _
           "Font: Carlito 9pt | Indent: 0.2"" | Justified | Spacing: Single", _
           vbInformation, "DefineBodyTextIndentStyle"

PROC_EXIT:
    Set oStyle = Nothing
    Set oDoc = Nothing
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure DefineBodyTextIndentStyle of Module basFixDocxRoutines"
    Resume PROC_EXIT
End Sub

'====================================================================
' AddBookNameHeaders
' PURPOSE:
'   Walk document sections and apply running book-name headers based
'   on Heading 1 (book title) and Heading 2 (first chapter) markers.
' BEHAVIOR:
'   Heading 1 section:
'       Header cleared (book title page)
'   Heading 2 section:
'       Header populated with most recent Heading 1 text (book name)
'   All other sections:
'       Header linked to previous section
' IMPORTANT IMPLEMENTATION DETAIL (RERUN SAFE):
'   Word headers always contain a terminating paragraph mark. Writing
'   directly to Header.Range.Text inserts an additional paragraph,
'   producing a second blank line each time the macro is executed.
'   BAD (creates extra blank lines on every run):
'       oHeader.Range.Text = sBookName
'   FIX:
'       Write only into the first paragraph range while excluding the
'       terminating paragraph mark. This overwrites existing content
'       without adding new paragraphs and makes the macro idempotent.
'   Additionally, any extra paragraphs are removed so the macro can
'   safely repair documents created by earlier buggy runs.
' Result:
'   - Only one header line is present
'   - No blank lines added
'   - Macro is safe to rerun
'   - Existing documents are auto-corrected
'====================================================================
Public Sub AddBookNameHeaders()
    On Error GoTo PROC_ERR
    Dim oDoc        As Document
    Dim oSections   As Sections
    Dim oSection    As Word.Section
    Dim oHeader     As HeaderFooter
    Dim oClassPara  As Word.Paragraph
    Dim lStartSect  As Long
    Dim lIdx        As Long
    Dim sBookName   As String
    Dim lResponse   As Long
    Dim bFoundH1    As Boolean
    Dim bFoundH2    As Boolean

    lResponse = MsgBox("Place cursor in the section to start header labelling. Do you want to start?", _
                       vbYesNo + vbDefaultButton2 + vbQuestion, _
                       "AddBookNameHeaders")
    If lResponse = vbNo Then GoTo PROC_EXIT

    Set oDoc = ActiveDocument
    Set oSections = oDoc.Sections

    ' -- Find the section containing the cursor --------------------------------
    lStartSect = 0
    For lIdx = 1 To oSections.Count
        If oSections(lIdx).Range.End >= Selection.Range.Start Then
            lStartSect = lIdx
            Exit For
        End If
    Next lIdx

    If lStartSect = 0 Then
        MsgBox "Could not determine the current section. " & _
               "Please place the cursor in the document body and try again.", _
               vbExclamation, "AddBookNameHeaders"
        GoTo PROC_EXIT
    End If

    sBookName = ""

    ' -- Walk sections from cursor to end -------------------------------------
    For lIdx = lStartSect To oSections.Count

        Set oSection = oSections(lIdx)
        Set oHeader = oSection.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary)

        ' Classify the section by scanning ALL paragraphs for H1 or H2.
        ' Checking only Paragraphs(1) fails when a blank paragraph precedes
        ' the heading (e.g. section-break artifact at a book title page).
        bFoundH1 = False
        bFoundH2 = False
        For Each oClassPara In oSection.Range.Paragraphs
            If oClassPara.style = oDoc.Styles("Heading 1") Then
                bFoundH1 = True
                Exit For
            ElseIf oClassPara.style = oDoc.Styles("Heading 2") Then
                bFoundH2 = True
                Exit For
            End If
        Next oClassPara

        If bFoundH1 Then
            ' Book title page: capture book name here (eliminates need for
            ' backward search later) then clear the header.
            ' Empty header spec: TheHeaders style, center-aligned, one tab
            ' character as intentional marker (default 0.1" tab, no other stops).
            sBookName = Trim$(Replace(oClassPara.Range.Text, vbCr, ""))
            oHeader.LinkToPrevious = False
            oHeader.Range.Delete
            With oHeader.Range.Paragraphs(1)
                .style = oDoc.Styles("TheHeaders")
                .Range.InsertBefore vbTab
            End With
            Debug.Print "Title page cleared: " & sBookName

        ElseIf bFoundH2 Then
            ' First chapter section: sBookName already set from H1 branch above.
            ' No backward search needed.
            oHeader.LinkToPrevious = False
            Do While oHeader.Range.Paragraphs.Count > 1
                oHeader.Range.Paragraphs.Last.Range.Delete
            Loop
            With oHeader.Range.Paragraphs(1).Range
                .End = .End - 1
                .Text = sBookName
                .style = oDoc.Styles("TheHeaders")
            End With
            Debug.Print "Header added: " & sBookName

        Else
            ' All other sections inherit the header from the section above
            oHeader.LinkToPrevious = True
        End If

    Next lIdx

    Debug.Print "Done. Book name headers applied from section " & _
           lStartSect & " through section " & oSections.Count & "."
    MsgBox "Done. Book name headers applied from section " & _
           lStartSect & " through section " & oSections.Count & ".", _
           vbInformation, "AddBookNameHeaders"

PROC_EXIT:
    Set oClassPara = Nothing
    Set oHeader = Nothing
    Set oSection = Nothing
    Set oSections = Nothing
    Set oDoc = Nothing
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure AddBookNameHeaders of Module basFixDocxRoutines"
    Resume PROC_EXIT
End Sub

Public Sub FixTheFooters()
    On Error GoTo PROC_ERR
    Dim lResponse As Long

    lResponse = MsgBox("Put the cursor in the section to commence renumbering of the footers.", _
                       vbYesNo + vbDefaultButton2 + vbQuestion, _
                       "FixTheFooters")

    If lResponse = vbYes Then
        Call AddConsecutiveFootersFromCursor
        Call LinkFootersToPrevious
    End If

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FixTheFooters of Module basFixDocxRoutines"
    Resume PROC_EXIT
End Sub

Private Sub AddConsecutiveFootersFromCursor()
    ' Adds a footer with consecutive page numbering starting at 1
    ' from the section containing the cursor through to the end of the document.
    ' Footer text is styled with the paragraph style "TheFooters".
    On Error GoTo PROC_ERR
    Dim oDoc        As Document
    Dim oSections   As Sections
    Dim oSection    As Word.Section
    Dim oFooter     As HeaderFooter
    Dim oRange      As Word.Range
    Dim oPara As Word.Paragraph
    Dim lStartSect  As Long
    Dim lIdx        As Long

    Set oDoc = ActiveDocument
    Set oSections = oDoc.Sections

    ' Use Selection.Range to locate the cursor section index
    lStartSect = 0
    For lIdx = 1 To oSections.Count
        If oSections(lIdx).Range.End >= Selection.Range.Start Then
            lStartSect = lIdx
            Exit For
        End If
    Next lIdx

    If lStartSect = 0 Then
        MsgBox "Could not determine the current section. " & _
               "Please place the cursor in the document body and try again.", _
               vbExclamation, "AddConsecutiveFootersFromCursor"
        GoTo PROC_EXIT
    End If

    ' Process every section from the cursor section to the end
    For lIdx = lStartSect To oSections.Count

        Set oSection = oSections(lIdx)

        Set oFooter = oSection.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)

        ' Break the link to the previous section's footer so we can set our own
        oFooter.LinkToPrevious = False

        ' Clear whatever is currently in this footer
        oFooter.Range.Delete

        ' Build the footer content
        Set oRange = oFooter.Range

        ' Apply the named paragraph style
        oRange.ParagraphFormat.style = oDoc.Styles("TheFooters")

        ' Insert the PAGE field so Word tracks the absolute page number.
        ' Using wdFieldPage gives the true physical page number, which is
        ' already consecutive across sections when NumPages restarts are off.
        oRange.Fields.Add Range:=oRange, _
                          Type:=WdFieldType.wdFieldPage, _
                          PreserveFormatting:=True

        ' Ensure page numbering does NOT restart in this section
        If lIdx = lStartSect Then
            oFooter.PageNumbers.StartingNumber = 1
            oFooter.PageNumbers.RestartNumberingAtSection = True
        Else
            oFooter.PageNumbers.RestartNumberingAtSection = False
        End If

        ' Re-apply the style to the paragraph that now contains the field
        ' (Fields.Add may reset paragraph formatting)
        For Each oPara In oFooter.Range.Paragraphs
            oPara.style = oDoc.Styles("TheFooters")
        Next oPara

    Next lIdx

    MsgBox "Done. Footers with consecutive page numbers (starting at 1) " & _
           "have been added from section " & lStartSect & _
           " through section " & oSections.Count & ".", _
           vbInformation, "AddConsecutiveFootersFromCursor"

PROC_EXIT:
    ' Clean up
    Set oPara = Nothing
    Set oRange = Nothing
    Set oFooter = Nothing
    Set oSection = Nothing
    Set oSections = Nothing
    Set oDoc = Nothing
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure AddConsecutiveFootersFromCursor of Module basFixDocxRoutines"
    Resume PROC_EXIT
End Sub

Private Sub LinkFootersToPrevious()
    On Error GoTo PROC_ERR
    Dim oDoc        As Document
    Dim oSections   As Sections
    Dim lStartSect  As Long
    Dim lIdx        As Long

    Set oDoc = ActiveDocument
    Set oSections = oDoc.Sections

    ' Find the section containing the cursor - same logic as AddConsecutiveFootersFromCursor
    lStartSect = 0
    For lIdx = 1 To oSections.Count
        If oSections(lIdx).Range.End >= Selection.Range.Start Then
            lStartSect = lIdx
            Exit For
        End If
    Next lIdx

    If lStartSect = 0 Then
        MsgBox "Could not determine the current section.", _
               vbExclamation, "LinkFootersToPrevious"
        GoTo PROC_EXIT
    End If

    ' Link from the section AFTER the cursor section to the end
    For lIdx = lStartSect + 1 To oSections.Count
        oSections(lIdx).Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = True
    Next lIdx

    MsgBox "Done. Sections " & lStartSect + 1 & " through " & oSections.Count & _
           " footers are now linked to previous.", _
           vbInformation, "LinkFootersToPrevious"

PROC_EXIT:
    Set oSections = Nothing
    Set oDoc = Nothing
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LinkFootersToPrevious of Module basFixDocxRoutines"
    Resume PROC_EXIT
End Sub

Public Sub FixHeader7()
    Dim sec As Word.Section
    Dim hdr As HeaderFooter
    Dim para As Word.Paragraph
    For Each sec In ActiveDocument.Sections
        Set hdr = sec.Headers(wdHeaderFooterPrimary)
        If hdr.Exists Then
            For Each para In hdr.Range.Paragraphs
                If para.style.NameLocal <> "TheHeaders" Then
                    Debug.Print "Section " & sec.index & ": style='" & para.style.NameLocal & "' linked=" & hdr.LinkToPrevious & _
                                    " | " & Left(para.Range.Text, 40)
                End If
            Next para
        End If
    Next sec
End Sub

Public Sub FixFrontMatterHeaders()
    ' Section 1 has LinkToPrevious=True which defers to Word's Normal template
    ' default "Header" style.  Break the link and apply the empty-header spec
    ' (TheHeaders + vbTab).  Sections 2-N that chain to Section 1 via
    ' LinkToPrevious will automatically inherit the corrected style.
    On Error GoTo PROC_ERR
    Dim oDoc    As Document
    Dim oHdr    As HeaderFooter

    Set oDoc = ActiveDocument
    Set oHdr = oDoc.Sections(1).Headers(wdHeaderFooterPrimary)

    oHdr.LinkToPrevious = False
    oHdr.Range.Delete
    With oHdr.Range.Paragraphs(1)
        .style = oDoc.Styles("TheHeaders")
        .Range.InsertBefore vbTab
    End With
    Debug.Print "FixFrontMatterHeaders: Section 1 header reset to TheHeaders."
PROC_EXIT:
    Set oHdr = Nothing
    Set oDoc = Nothing
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FixFrontMatterHeaders of Module basFixDocxRoutines"
    Resume PROC_EXIT
End Sub

'====================================================================
' ReplaceNormalWithBodyText
' PURPOSE:
'   Replaces every paragraph whose style is EXACTLY "Normal" with
'   BodyText.  This is the primary fix for Bible text paragraphs —
'   the author used Normal throughout; BodyText is the semantic
'   replacement (USFM \p).
'
' SCOPE:
'   doc.Content only (main body story).  Headers, footers, and
'   footnotes are not affected — they carry their own styles.
'
' SAFETY — EXACT MATCH ONLY:
'   Uses paragraph iteration with NameLocal = "Normal" exact match.
'   Find/Replace must NOT be used here — Word's Find/Replace with a
'   style also matches child styles (styles based on Normal such as
'   Words of Jesus, EmphasisRed, EmphasisBlack) and would destroy
'   their semantic assignments.
'
' RERUN SAFE:
'   After the first run Normal Count = 0; subsequent runs report
'   "0 replaced" and exit cleanly.
'
' PREREQUISITE:
'   DefineBodyTextStyle must have been run first.
'====================================================================
Public Sub ReplaceNormalWithBodyText()
    On Error GoTo PROC_ERR
    Dim oDoc        As Document
    Dim oPara       As Word.Paragraph
    Dim lBefore     As Long
    Dim lReplaced   As Long
    Dim lResponse   As Long

    Set oDoc = ActiveDocument

    ' Verify BodyText exists before proceeding
    Dim oCheck As Word.Style
    On Error Resume Next
    Set oCheck = oDoc.Styles("BodyText")
    On Error GoTo PROC_ERR
    If oCheck Is Nothing Then
        MsgBox "BodyText style not found. Run DefineBodyTextStyle first.", _
               vbExclamation, "ReplaceNormalWithBodyText"
        GoTo PROC_EXIT
    End If

    ' Count exact Normal paragraphs (NameLocal match — child styles excluded)
    lBefore = 0
    For Each oPara In oDoc.Content.Paragraphs
        If oPara.style.NameLocal = "Normal" Then lBefore = lBefore + 1
    Next oPara

    If lBefore = 0 Then
        Debug.Print "ReplaceNormalWithBodyText: No Normal paragraphs found -- nothing to do."
        MsgBox "No Normal paragraphs found in document body. Nothing replaced.", _
               vbInformation, "ReplaceNormalWithBodyText"
        GoTo PROC_EXIT
    End If

    lResponse = MsgBox(lBefore & " Normal paragraphs found. Replace all with BodyText?", _
                       vbYesNo + vbDefaultButton2 + vbQuestion, _
                       "ReplaceNormalWithBodyText")
    If lResponse = vbNo Then GoTo PROC_EXIT

    ' Iterate and replace — exact NameLocal match only
    lReplaced = 0
    For Each oPara In oDoc.Content.Paragraphs
        If oPara.style.NameLocal = "Normal" Then
            oPara.style = oDoc.Styles("BodyText")
            lReplaced = lReplaced + 1
        End If
    Next oPara

    Debug.Print "ReplaceNormalWithBodyText: " & lReplaced & " replaced."
    MsgBox "Done. " & lReplaced & " paragraphs changed from Normal to BodyText.", _
           vbInformation, "ReplaceNormalWithBodyText"

PROC_EXIT:
    Set oPara = Nothing
    Set oCheck = Nothing
    Set oDoc = Nothing
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ReplaceNormalWithBodyText of Module basFixDocxRoutines"
    Resume PROC_EXIT
End Sub

'====================================================================
' DefineAppendixTitleStyle
' PURPOSE:
'   Creates the AppendixTitle style if it does not already exist.
'   Used for section titles within appendix content (e.g. the
'   Concordance section heading).
'
' FORMATTING:
'   Font:           Carlito 10pt, Bold
'   Alignment:      Left (wdAlignParagraphLeft)
'   Line spacing:   Single
'   First indent:   0
'   Space before:   6pt
'   Space after:    0pt
'
' USFM mapping:  \imt  (introduction major title)
'====================================================================
Public Sub DefineAppendixTitleStyle()
    On Error GoTo PROC_ERR
    Dim oDoc    As Document
    Dim oStyle  As Word.Style

    Set oDoc = ActiveDocument

    On Error Resume Next
    Set oStyle = oDoc.Styles("AppendixTitle")
    On Error GoTo PROC_ERR

    If Not oStyle Is Nothing Then
        Debug.Print "DefineAppendixTitleStyle: AppendixTitle already exists -- no changes made."
        MsgBox "AppendixTitle style already exists in this document.", _
               vbInformation, "DefineAppendixTitleStyle"
        GoTo PROC_EXIT
    End If

    Set oStyle = oDoc.Styles.Add(name:="AppendixTitle", Type:=wdStyleTypeParagraph)

    With oStyle
        .baseStyle = ""
        .NextParagraphStyle = oDoc.Styles("AppendixBody")

        With .Font
            .Name = "Carlito"
            .Size = 10
            .Bold = True
            .Italic = False
        End With

        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpaceSingle
            .FirstLineIndent = 0
            .LeftIndent = 0
            .SpaceBefore = 6
            .SpaceAfter = 0
        End With
    End With

    Debug.Print "DefineAppendixTitleStyle: AppendixTitle created successfully."
    MsgBox "AppendixTitle style created successfully." & vbCrLf & _
           "Font: Carlito 10pt Bold | Left | 6pt before", _
           vbInformation, "DefineAppendixTitleStyle"

PROC_EXIT:
    Set oStyle = Nothing
    Set oDoc = Nothing
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure DefineAppendixTitleStyle of Module basFixDocxRoutines"
    Resume PROC_EXIT
End Sub

'====================================================================
' DefineAppendixBodyStyle
' PURPOSE:
'   Creates the AppendixBody style if it does not already exist.
'   Used for body paragraphs within appendix content (e.g. the
'   Concordance explanatory text and software links).
'   Visually identical to BodyText; semantically distinct for USFM
'   export (\ip vs \p).
'
' FORMATTING:
'   Font:           Carlito 9pt
'   Alignment:      Justified
'   Line spacing:   Single
'   First indent:   0
'   Space before:   0pt
'   Space after:    0pt
'
' USFM mapping:  \ip  (introduction paragraph / appendix body)
'====================================================================
Public Sub DefineAppendixBodyStyle()
    On Error GoTo PROC_ERR
    Dim oDoc    As Document
    Dim oStyle  As Word.Style

    Set oDoc = ActiveDocument

    On Error Resume Next
    Set oStyle = oDoc.Styles("AppendixBody")
    On Error GoTo PROC_ERR

    If Not oStyle Is Nothing Then
        Debug.Print "DefineAppendixBodyStyle: AppendixBody already exists -- no changes made."
        MsgBox "AppendixBody style already exists in this document.", _
               vbInformation, "DefineAppendixBodyStyle"
        GoTo PROC_EXIT
    End If

    Set oStyle = oDoc.Styles.Add(name:="AppendixBody", Type:=wdStyleTypeParagraph)

    With oStyle
        .baseStyle = ""
        .NextParagraphStyle = oStyle

        With .Font
            .Name = "Carlito"
            .Size = 9
            .Bold = False
            .Italic = False
        End With

        With .ParagraphFormat
            .Alignment = wdAlignParagraphJustify
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing = 10                   ' Exactly 10pt — matches BodyText
            .FirstLineIndent = 0
            .LeftIndent = 0
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
    End With

    Debug.Print "DefineAppendixBodyStyle: AppendixBody created successfully."
    MsgBox "AppendixBody style created successfully." & vbCrLf & _
           "Font: Carlito 9pt | Justified | No indent", _
           vbInformation, "DefineAppendixBodyStyle"

PROC_EXIT:
    Set oStyle = Nothing
    Set oDoc = Nothing
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure DefineAppendixBodyStyle of Module basFixDocxRoutines"
    Resume PROC_EXIT
End Sub

'====================================================================
' ReplacePlainTextStyles
' PURPOSE:
'   Replaces all 26 Plain Text paragraphs with the correct semantic
'   style based on position in the document:
'     - Front matter (position < 1 000 000) -> BodyText
'     - Concordance appendix (position >= 1 000 000) -> AppendixBody
'
' MANUAL FOLLOW-UP:
'   The concordance contains section-letter headings (e.g. "A") and
'   sub-section titles that may warrant AppendixTitle style. Review
'   the Immediate Window output and reclassify manually as needed.
'
' PREREQUISITES:
'   DefineBodyTextStyle, DefineAppendixBodyStyle must have been run.
'====================================================================
Public Sub ReplacePlainTextStyles()
    On Error GoTo PROC_ERR
    Dim oDoc        As Document
    Dim oPara       As Word.Paragraph
    Dim lCount      As Long
    Dim lBody       As Long
    Dim lAppendix   As Long

    Set oDoc = ActiveDocument

    ' Verify required styles exist
    Dim oCheck As Word.Style
    On Error Resume Next
    Set oCheck = oDoc.Styles("AppendixBody")
    On Error GoTo PROC_ERR
    If oCheck Is Nothing Then
        MsgBox "AppendixBody style not found. Run DefineAppendixBodyStyle first.", _
               vbExclamation, "ReplacePlainTextStyles"
        GoTo PROC_EXIT
    End If

    ' Count Plain Text paragraphs first
    lCount = 0
    For Each oPara In oDoc.Content.Paragraphs
        If oPara.style.NameLocal = "Plain Text" Then lCount = lCount + 1
    Next oPara

    If lCount = 0 Then
        Debug.Print "ReplacePlainTextStyles: No Plain Text paragraphs found -- nothing to do."
        MsgBox "No Plain Text paragraphs found. Nothing replaced.", _
               vbInformation, "ReplacePlainTextStyles"
        GoTo PROC_EXIT
    End If

    Dim lResponse As Long
    lResponse = MsgBox(lCount & " Plain Text paragraphs found. Replace all?", _
                       vbYesNo + vbDefaultButton2 + vbQuestion, _
                       "ReplacePlainTextStyles")
    If lResponse = vbNo Then GoTo PROC_EXIT

    lBody = 0
    lAppendix = 0
    For Each oPara In oDoc.Content.Paragraphs
        If oPara.style.NameLocal = "Plain Text" Then
            If oPara.Range.Start < 1000000 Then
                oPara.style = oDoc.Styles("BodyText")
                lBody = lBody + 1
                Debug.Print "BodyText     p" & oPara.Range.Start & " | " & _
                            Left(Trim(oPara.Range.Text), 40)
            Else
                oPara.style = oDoc.Styles("AppendixBody")
                lAppendix = lAppendix + 1
                Debug.Print "AppendixBody p" & oPara.Range.Start & " | " & _
                            Left(Trim(oPara.Range.Text), 40)
            End If
        End If
    Next oPara

    Debug.Print "ReplacePlainTextStyles: " & lBody & " -> BodyText, " & _
                lAppendix & " -> AppendixBody."
    MsgBox "Done." & vbCrLf & _
           lBody & " paragraphs -> BodyText" & vbCrLf & _
           lAppendix & " paragraphs -> AppendixBody" & vbCrLf & _
           "Review Immediate Window for concordance entries that may need AppendixTitle.", _
           vbInformation, "ReplacePlainTextStyles"

PROC_EXIT:
    Set oCheck = Nothing
    Set oPara = Nothing
    Set oDoc = Nothing
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ReplacePlainTextStyles of Module basFixDocxRoutines"
    Resume PROC_EXIT
End Sub

'====================================================================
' DefineBookIntroStyle
' PURPOSE:
'   Creates the BookIntro style if it does not already exist.
'   Used for the centered background-summary paragraph that follows
'   each DatAuthRef paragraph — one paragraph per book using manual
'   line breaks (Shift+Enter) to separate Dating, Authorship, etc.
'
' FORMATTING:
'   Font:           Carlito 9pt
'   Alignment:      Center (wdAlignParagraphCenter)
'   Line spacing:   Single
'   First indent:   0
'   Space before:   6pt  (visual separation from DatAuthRef)
'   Space after:    6pt  (visual separation from following content)
'
' USFM mapping:  \ip  (introduction paragraph)
'
' SIDE EFFECT:
'   Sets DatAuthRef.NextParagraphStyle = BookIntro so Word
'   automatically applies BookIntro when Enter is pressed after
'   a DatAuthRef paragraph.
'====================================================================
Public Sub DefineBookIntroStyle()
    On Error GoTo PROC_ERR
    Dim oDoc    As Document
    Dim oStyle  As Word.Style

    Set oDoc = ActiveDocument

    On Error Resume Next
    Set oStyle = oDoc.Styles("BookIntro")
    On Error GoTo PROC_ERR

    If Not oStyle Is Nothing Then
        Debug.Print "DefineBookIntroStyle: BookIntro already exists -- no changes made."
        MsgBox "BookIntro style already exists in this document.", _
               vbInformation, "DefineBookIntroStyle"
        GoTo PROC_EXIT
    End If

    Set oStyle = oDoc.Styles.Add(name:="BookIntro", Type:=wdStyleTypeParagraph)

    With oStyle
        .baseStyle = ""
        .NextParagraphStyle = oDoc.Styles("BodyText")

        With .Font
            .Name = "Carlito"
            .Size = 9
            .Bold = False
            .Italic = False
        End With

        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing = 10                   ' Exactly 10pt — matches BodyText
            .FirstLineIndent = 0
            .LeftIndent = 0
            .SpaceBefore = 6
            .SpaceAfter = 6
        End With
    End With

    ' Set DatAuthRef to flow into BookIntro on Enter
    Dim oDatAuth As Word.Style
    On Error Resume Next
    Set oDatAuth = oDoc.Styles("DatAuthRef")
    On Error GoTo PROC_ERR
    If Not oDatAuth Is Nothing Then
        oDatAuth.NextParagraphStyle = oDoc.Styles("BookIntro")
        Debug.Print "DefineBookIntroStyle: DatAuthRef.NextParagraphStyle set to BookIntro."
    End If

    Debug.Print "DefineBookIntroStyle: BookIntro created successfully."
    MsgBox "BookIntro style created successfully." & vbCrLf & _
           "Font: Carlito 9pt | Centered | 6pt before/after" & vbCrLf & _
           "DatAuthRef next-style set to BookIntro.", _
           vbInformation, "DefineBookIntroStyle"

PROC_EXIT:
    Set oDatAuth = Nothing
    Set oStyle = Nothing
    Set oDoc = Nothing
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure DefineBookIntroStyle of Module basFixDocxRoutines"
    Resume PROC_EXIT
End Sub

'====================================================================
' ApplyBookIntroAfterDatAuthRef
' PURPOSE:
'   Finds every paragraph immediately following a DatAuthRef paragraph
'   and reclassifies it as BookIntro.  The BookIntro style definition
'   supplies center alignment — this routine does not select by
'   existing alignment.
'
'   Rule: all book introduction summary paragraphs live on the same
'   page as Heading 1 (book title) and follow DatAuthRef.  Applying
'   BookIntro makes centering consistent across all 66 books regardless
'   of what direct formatting the author may have applied.
'
' RERUN SAFE: paragraphs already styled BookIntro are skipped.
'====================================================================
Public Sub ApplyBookIntroAfterDatAuthRef()
    On Error GoTo PROC_ERR
    Dim oDoc    As Document
    Dim oPara   As Word.Paragraph
    Dim oNext   As Word.Paragraph
    Dim lCount  As Long
    Dim lBefore As Long

    Set oDoc = ActiveDocument

    ' Verify BookIntro exists
    Dim oCheck As Word.Style
    On Error Resume Next
    Set oCheck = oDoc.Styles("BookIntro")
    On Error GoTo PROC_ERR
    If oCheck Is Nothing Then
        MsgBox "BookIntro style not found. Run DefineBookIntroStyle first.", _
               vbExclamation, "ApplyBookIntroAfterDatAuthRef"
        GoTo PROC_EXIT
    End If

    ' Count candidates — paragraphs that follow DatAuthRef and are not yet BookIntro
    lBefore = 0
    For Each oPara In oDoc.Content.Paragraphs
        If oPara.style.NameLocal = "DatAuthRef" Then
            On Error Resume Next
            Set oNext = oPara.Next
            On Error GoTo PROC_ERR
            If Not oNext Is Nothing Then
                If oNext.style.NameLocal <> "BookIntro" Then lBefore = lBefore + 1
            End If
        End If
    Next oPara

    If lBefore = 0 Then
        Debug.Print "ApplyBookIntroAfterDatAuthRef: All DatAuthRef paragraphs already followed by BookIntro."
        MsgBox "Nothing to do — all DatAuthRef paragraphs already followed by BookIntro.", _
               vbInformation, "ApplyBookIntroAfterDatAuthRef"
        GoTo PROC_EXIT
    End If

    Dim lResponse As Long
    lResponse = MsgBox(lBefore & " paragraph(s) following DatAuthRef will be reclassified as BookIntro." & _
                       vbCrLf & "Proceed?", _
                       vbYesNo + vbDefaultButton2 + vbQuestion, _
                       "ApplyBookIntroAfterDatAuthRef")
    If lResponse = vbNo Then GoTo PROC_EXIT

    lCount = 0
    For Each oPara In oDoc.Content.Paragraphs
        If oPara.style.NameLocal = "DatAuthRef" Then
            On Error Resume Next
            Set oNext = oPara.Next
            On Error GoTo PROC_ERR
            If Not oNext Is Nothing Then
                If oNext.style.NameLocal <> "BookIntro" Then
                    oNext.style = oDoc.Styles("BookIntro")
                    lCount = lCount + 1
                    Debug.Print "BookIntro p" & oNext.Range.Start & " | " & _
                                Left(Trim(oNext.Range.Text), 50)
                End If
            End If
        End If
    Next oPara

    Debug.Print "ApplyBookIntroAfterDatAuthRef: " & lCount & " paragraphs reclassified as BookIntro."
    MsgBox "Done. " & lCount & " paragraphs reclassified as BookIntro.", _
           vbInformation, "ApplyBookIntroAfterDatAuthRef"

PROC_EXIT:
    Set oCheck = Nothing
    Set oNext = Nothing
    Set oPara = Nothing
    Set oDoc = Nothing
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ApplyBookIntroAfterDatAuthRef of Module basFixDocxRoutines"
    Resume PROC_EXIT
End Sub

'====================================================================
' DefineAuthorStyles
' PURPOSE:
'   Creates all four author-text styles in one call:
'     AuthorBodyText    - paragraph style for author body text
'     AuthorSectionHead - paragraph style for author section headings
'     AuthorQuote       - character style for red italic quotes
'     AuthorRef         - character style for bold book references
'
'   Font: Liberation Serif (free, metric-compatible with Times New Roman)
'
' RERUN SAFE: skips any style that already exists.
'====================================================================
Public Sub DefineAuthorStyles()
    On Error GoTo PROC_ERR
    Dim oDoc As Document
    Set oDoc = ActiveDocument

    ' --- AuthorBodyText ---------------------------------------------------------------------------
    ' Liberation Serif 12pt, Justified, 0.33" first indent,
    ' 12pt space after, Single spacing, Widow/Orphan control.
    ' USFM: \ip
    Dim oStyle As Word.Style
    On Error Resume Next
    Set oStyle = oDoc.Styles("AuthorBodyText")
    On Error GoTo PROC_ERR
    If oStyle Is Nothing Then
        Set oStyle = oDoc.Styles.Add(name:="AuthorBodyText", Type:=wdStyleTypeParagraph)
        With oStyle
            .baseStyle = ""
            .NextParagraphStyle = oStyle
            With .Font
                .Name = "Liberation Serif"
                .Size = 12
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            With .ParagraphFormat
                .Alignment = wdAlignParagraphJustify
                .LineSpacingRule = wdLineSpaceSingle
                .FirstLineIndent = 23.76      ' 0.33 inches in points
                .LeftIndent = 0
                .SpaceBefore = 0
                .SpaceAfter = 12
                .WidowControl = True
            End With
        End With
        Debug.Print "DefineAuthorStyles: AuthorBodyText created."
    Else
        Debug.Print "DefineAuthorStyles: AuthorBodyText already exists -- skipped."
    End If

    ' --- AuthorSectionHead ---------------------------------------------------------------------------
    ' Liberation Serif 14pt, plain (bold/italic applied word by word
    ' as direct formatting).  Space before 12pt, after 6pt.
    ' USFM: \is
    Set oStyle = Nothing
    On Error Resume Next
    Set oStyle = oDoc.Styles("AuthorSectionHead")
    On Error GoTo PROC_ERR
    If oStyle Is Nothing Then
        Set oStyle = oDoc.Styles.Add(name:="AuthorSectionHead", Type:=wdStyleTypeParagraph)
        With oStyle
            .baseStyle = ""
            .NextParagraphStyle = oDoc.Styles("AuthorBodyText")
            With .Font
                .Name = "Liberation Serif"
                .Size = 14
                .Bold = False
                .Italic = False
                .color = wdColorAutomatic
            End With
            With .ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .LineSpacingRule = wdLineSpaceSingle
                .FirstLineIndent = 0
                .LeftIndent = 0
                .SpaceBefore = 12
                .SpaceAfter = 6
                .WidowControl = False
                .PageBreakBefore = True
            End With
        End With
        Debug.Print "DefineAuthorStyles: AuthorSectionHead created."
    Else
        Debug.Print "DefineAuthorStyles: AuthorSectionHead already exists -- skipped."
    End If

    ' --- AuthorQuote ---------------------------------------------------------------------------
    ' Red italic - quotes of Jesus in author text.
    ' USFM: \wj
    Set oStyle = Nothing
    On Error Resume Next
    Set oStyle = oDoc.Styles("AuthorQuote")
    On Error GoTo PROC_ERR
    If oStyle Is Nothing Then
        Set oStyle = oDoc.Styles.Add(name:="AuthorQuote", Type:=wdStyleTypeCharacter)
        With oStyle.Font
            .Italic = True
            .color = wdColorRed
        End With
        Debug.Print "DefineAuthorStyles: AuthorQuote created."
    Else
        Debug.Print "DefineAuthorStyles: AuthorQuote already exists -- skipped."
    End If

    ' Note: AuthorRef (character style) removed — replaced by AuthorBookRef (paragraph style).
    '       See DefineAuthorBookRefStyle.

    MsgBox "AuthorBodyText, AuthorSectionHead, AuthorQuote - done." & vbCrLf & _
           "Check Immediate Window for details.", vbInformation, "DefineAuthorStyles"

PROC_EXIT:
    Set oStyle = Nothing
    Set oDoc = Nothing
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure DefineAuthorStyles of Module basFixDocxRoutines"
    Resume PROC_EXIT
End Sub

'====================================================================
' DefineListItemStyle
' PURPOSE:
'   Creates the ListItem paragraph style.
'   Used for bold-italic headed entries in study lists and
'   Concordance sections.
'
' SPEC:
'   Font:           Carlito 11pt, Bold, Italic
'   Alignment:      Left
'   LeftIndent:     36pt (0.5 inch)
'   FirstLineIndent:0
'   LineSpacing:    Single
'   WidowControl:   True
'   NextStyle:      ListItemBody
'
' USFM mapping:  \li1  (list item level 1)
'
' RERUN SAFE: skips if style already exists.
'====================================================================
Public Sub DefineListItemStyle()
    On Error GoTo PROC_ERR
    Dim oDoc   As Document
    Dim oStyle As Word.Style
    Dim oNext  As Word.Style

    Set oDoc = ActiveDocument

    On Error Resume Next
    Set oStyle = oDoc.Styles("ListItem")
    On Error GoTo PROC_ERR

    If Not oStyle Is Nothing Then
        Debug.Print "DefineListItemStyle: ListItem already exists -- skipped."
        GoTo PROC_EXIT
    End If

    ' ListItemBody must exist first (referenced as NextParagraphStyle)
    On Error Resume Next
    Set oNext = oDoc.Styles("ListItemBody")
    On Error GoTo PROC_ERR

    Set oStyle = oDoc.Styles.Add(name:="ListItem", Type:=wdStyleTypeParagraph)
    With oStyle
        .baseStyle = "List Number"    ' inherits autonumbering list template
        If Not oNext Is Nothing Then
            .NextParagraphStyle = oNext
        End If
        With .Font
            .Name = "Carlito"
            .Size = 11
            .Bold = True
            .Italic = True
            .color = wdColorAutomatic
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpaceSingle
            .FirstLineIndent = -18    ' number at 0.25" (LeftIndent 0.5" minus 0.25")
            .LeftIndent = 36          ' text wraps at 0.5"
            .SpaceBefore = 0
            .SpaceAfter = 0
            .WidowControl = True
            .KeepWithNext = True
        End With
    End With

    Debug.Print "DefineListItemStyle: ListItem created."

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure DefineListItemStyle of Module basFixDocxRoutines"
    Resume PROC_EXIT
End Sub

'====================================================================
' DefineListItemBodyStyle
' PURPOSE:
'   Creates the ListItemBody paragraph style.
'   Continuation paragraph following a ListItem — plain body text
'   at the same indent level.
'
' SPEC:
'   Font:           Carlito 11pt, not Bold, not Italic
'   Alignment:      Left (not justified)
'   LeftIndent:     36pt (0.5 inch — aligns under ListItem)
'   FirstLineIndent:0
'   LineSpacing:    Single
'   SpaceAfter:     11pt
'   WidowControl:   True
'   NextStyle:      ListItem  (cycles back for next entry)
'
' USFM mapping:  \lim1  (list item continuation level 1)
'
' Note: Run this BEFORE DefineListItemStyle so ListItem can
'       reference ListItemBody as its NextParagraphStyle.
'
' RERUN SAFE: skips if style already exists.
'====================================================================
Public Sub DefineListItemBodyStyle()
    On Error GoTo PROC_ERR
    Dim oDoc   As Document
    Dim oStyle As Word.Style

    Set oDoc = ActiveDocument

    On Error Resume Next
    Set oStyle = oDoc.Styles("ListItemBody")
    On Error GoTo PROC_ERR

    If Not oStyle Is Nothing Then
        Debug.Print "DefineListItemBodyStyle: ListItemBody already exists -- skipped."
        GoTo PROC_EXIT
    End If

    Set oStyle = oDoc.Styles.Add(name:="ListItemBody", Type:=wdStyleTypeParagraph)
    With oStyle
        .baseStyle = ""
        ' NextParagraphStyle cycles back to ListItem for the next entry.
        ' Wire after DefineListItemStyle runs; here we set it if ListItem exists.
        Dim oNext As Word.Style
        On Error Resume Next
        Set oNext = oDoc.Styles("ListItem")
        On Error GoTo PROC_ERR
        If Not oNext Is Nothing Then
            .NextParagraphStyle = oNext
        End If
        With .Font
            .Name = "Carlito"
            .Size = 11
            .Bold = False
            .Italic = False
            .color = wdColorAutomatic
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpaceSingle
            .FirstLineIndent = 0
            .LeftIndent = 36          ' 0.5 inch x 72 points
            .SpaceBefore = 0
            .SpaceAfter = 11
            .WidowControl = True
        End With
    End With

    Debug.Print "DefineListItemBodyStyle: ListItemBody created."

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure DefineListItemBodyStyle of Module basFixDocxRoutines"
    Resume PROC_EXIT
End Sub

'====================================================================
' DefineAuthorBookRefStyle
' PURPOSE:
'   Creates the AuthorBookRef paragraph style.
'   Cross-reference lookup entries — book name with tab-leader
'   dot leader to right-aligned page number.
'   Replaces the former AuthorRef character style.
'
' SPEC:
'   Font:           Carlito 11pt, Bold, not Italic
'   Base style:     List Number (inherits autonumbering)
'   LeftIndent:     36pt (0.5")
'   FirstLineIndent:-18pt (number at 0.25")
'   KeepWithNext:   False
'   WidowControl:   True
'   Tab stop:       Right-aligned at 5.3" (381.6pt), dot leader
'
' USFM mapping:  \xt  (cross-reference target)
'
' RERUN SAFE: skips if style already exists.
'====================================================================
Public Sub DefineAuthorBookRefStyle()
    On Error GoTo PROC_ERR
    Dim oDoc   As Document
    Dim oStyle As Word.Style

    Set oDoc = ActiveDocument

    On Error Resume Next
    Set oStyle = oDoc.Styles("AuthorBookRef")
    On Error GoTo PROC_ERR

    If Not oStyle Is Nothing Then
        Debug.Print "DefineAuthorBookRefStyle: AuthorBookRef already exists -- skipped."
        GoTo PROC_EXIT
    End If

    Set oStyle = oDoc.Styles.Add(name:="AuthorBookRef", Type:=wdStyleTypeParagraph)
    With oStyle
        .baseStyle = "List Number"
        .NextParagraphStyle = oStyle
        With .Font
            .Name = "Carlito"
            .Size = 11
            .Bold = True
            .Italic = False
            .color = wdColorAutomatic
        End With
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .LineSpacingRule = wdLineSpaceSingle
            .FirstLineIndent = -18        ' number at 0.25"
            .LeftIndent = 36              ' text at 0.5"
            .SpaceBefore = 0
            .SpaceAfter = 11
            .WidowControl = True
            .KeepWithNext = False
            .TabStops.Add Position:=381.6, _
                           Alignment:=wdAlignTabRight, _
                           Leader:=wdTabLeaderDots
        End With
    End With

    Debug.Print "DefineAuthorBookRefStyle: AuthorBookRef created."

PROC_EXIT:
    Set oStyle = Nothing
    Set oDoc = Nothing
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure DefineAuthorBookRefStyle of Module basFixDocxRoutines"
    Resume PROC_EXIT
End Sub
