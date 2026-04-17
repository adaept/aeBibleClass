Attribute VB_Name = "basAddHeaderFooter"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

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
    Dim oPara As Word.Paragraph
    Dim oRange      As Word.Range
    Dim lStartSect  As Long
    Dim lIdx        As Long
    Dim sBookName   As String
    Dim lResponse   As Long

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
        Set oPara = oSection.Range.Paragraphs(1)

        If oPara.style = oDoc.Styles("Heading 1") Then
            ' Book title page - clear the header and leave it empty
            oHeader.LinkToPrevious = False
            oHeader.Range.Delete
        ElseIf oPara.style = oDoc.Styles("Heading 2") Then
            ' First chapter section - capture book name from Heading 1 text
            ' Search backwards from this section for the nearest Heading 1
            Dim oSearch As Word.Range
            Set oSearch = oDoc.Range(0, oSection.Range.Start)
            Dim oFound  As Word.Range
            Set oFound = Nothing

            Dim pIdx    As Long
            For pIdx = oSearch.Paragraphs.Count To 1 Step -1
                If oSearch.Paragraphs(pIdx).style = oDoc.Styles("Heading 1") Then
                    sBookName = Trim$(Replace(oSearch.Paragraphs(pIdx).Range.Text, vbCr, ""))
                    Exit For
                End If
            Next pIdx

            ' Write the book name into the header
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

    Debug.Print "Done. Book name headers have been added from section " & _
           lStartSect & " through section " & oSections.Count & "."
    MsgBox "Done. Book name headers have been added from section " & _
           lStartSect & " through section " & oSections.Count & ".", _
           vbInformation, "AddBookNameHeaders"

PROC_EXIT:
    Set oRange = Nothing
    Set oPara = Nothing
    Set oHeader = Nothing
    Set oSection = Nothing
    Set oSections = Nothing
    Set oDoc = Nothing
    Set oSearch = Nothing
    Set oFound = Nothing
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure AddBookNameHeaders of Module basAddHeaderFooter"
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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FixTheFooters of Module basAddHeaderFooter"
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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure AddConsecutiveFootersFromCursor of Module basAddHeaderFooter"
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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LinkFootersToPrevious of Module basAddHeaderFooter"
    Resume PROC_EXIT
End Sub

