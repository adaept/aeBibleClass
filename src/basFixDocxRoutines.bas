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
    Dim oStyle  As Word.style

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
            .LineSpacingRule = wdLineSpaceSingle
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
    Dim oStyle  As Word.style

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
            .LineSpacingRule = wdLineSpaceSingle
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
'   Replaces every paragraph styled Normal in the document body with
'   BodyText.  This is the primary fix for Bible text paragraphs -
'   the author used Normal throughout; BodyText is the semantic
'   replacement (USFM \p).
'
' SCOPE:
'   doc.Content only (main body story).  Headers, footers, and
'   footnotes are not affected — they carry their own styles.
'
' PERFORMANCE:
'   Uses Find/Replace with Format:=True rather than iterating
'   paragraphs - safe and fast on 16 000+ paragraphs.
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
    Dim oRange      As Word.Range
    Dim lBefore     As Long
    Dim lAfter      As Long
    Dim lReplaced   As Long

    Set oDoc = ActiveDocument

    ' Verify BodyText exists before proceeding
    Dim oCheck As Word.style
    On Error Resume Next
    Set oCheck = oDoc.Styles("BodyText")
    On Error GoTo PROC_ERR
    If oCheck Is Nothing Then
        MsgBox "BodyText style not found. Run DefineBodyTextStyle first.", _
               vbExclamation, "ReplaceNormalWithBodyText"
        GoTo PROC_EXIT
    End If

    ' Count Normal paragraphs before replacement
    lBefore = 0
    Dim oPara As Word.Paragraph
    For Each oPara In oDoc.Content.Paragraphs
        If oPara.style.NameLocal = "Normal" Then lBefore = lBefore + 1
    Next oPara

    If lBefore = 0 Then
        Debug.Print "ReplaceNormalWithBodyText: No Normal paragraphs found -- nothing to do."
        MsgBox "No Normal paragraphs found in document body. Nothing replaced.", _
               vbInformation, "ReplaceNormalWithBodyText"
        GoTo PROC_EXIT
    End If

    ' Confirm before proceeding
    Dim lResponse As Long
    lResponse = MsgBox(lBefore & " Normal paragraphs found. Replace all with BodyText?", _
                       vbYesNo + vbDefaultButton2 + vbQuestion, _
                       "ReplaceNormalWithBodyText")
    If lResponse = vbNo Then GoTo PROC_EXIT

    ' Use Find/Replace for performance
    Set oRange = oDoc.Content
    With oRange.Find
        .ClearFormatting
        .style = oDoc.Styles("Normal")
        .Text = ""
        .Replacement.ClearFormatting
        .Replacement.style = oDoc.Styles("BodyText")
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .Execute Replace:=wdReplaceAll
    End With

    ' Count remaining Normal paragraphs to confirm
    lAfter = 0
    For Each oPara In oDoc.Content.Paragraphs
        If oPara.style.NameLocal = "Normal" Then lAfter = lAfter + 1
    Next oPara
    lReplaced = lBefore - lAfter

    Debug.Print "ReplaceNormalWithBodyText: " & lReplaced & " replaced, " & lAfter & " remaining."
    MsgBox "Done. " & lReplaced & " paragraphs changed from Normal to BodyText." & vbCrLf & _
           lAfter & " Normal paragraphs remaining (check Immediate Window for details).", _
           vbInformation, "ReplaceNormalWithBodyText"

PROC_EXIT:
    Set oPara = Nothing
    Set oCheck = Nothing
    Set oRange = Nothing
    Set oDoc = Nothing
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ReplaceNormalWithBodyText of Module basFixDocxRoutines"
    Resume PROC_EXIT
End Sub
