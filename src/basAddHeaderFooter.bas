Attribute VB_Name = "basAddHeaderFooter"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Public Sub FixTheFooters()
    Dim lResponse As Long

    lResponse = MsgBox("Put the cursor in the section to commence renumbering of the footers.", _
                       vbYesNo + vbDefaultButton2 + vbQuestion, _
                       "FixTheFooters")

    If lResponse = vbYes Then
        Call AddConsecutiveFootersFromCursor
        Call LinkFootersToPrevious
    End If
End Sub

Private Sub AddConsecutiveFootersFromCursor()
    ' Adds a footer with consecutive page numbering starting at 1
    ' from the section containing the cursor through to the end of the document.
    ' Footer text is styled with the paragraph style "TheFooters".

    Dim oDoc        As Document
    Dim oSections   As Sections
    Dim oSection    As section
    Dim oFooter     As HeaderFooter
    Dim oRange      As Range
    Dim oPara As Word.Paragraph
    Dim lStartSect  As Long
    Dim lIdx        As Long

    Set oDoc = ActiveDocument
    Set oSections = oDoc.Sections

    ' Use Selection.Range to locate the cursor section index
    lStartSect = 0
    For lIdx = 1 To oSections.count
        If oSections(lIdx).Range.End >= Selection.Range.Start Then
            lStartSect = lIdx
            Exit For
        End If
    Next lIdx

    If lStartSect = 0 Then
        MsgBox "Could not determine the current section. " & _
               "Please place the cursor in the document body and try again.", _
               vbExclamation, "AddConsecutiveFootersFromCursor"
        Exit Sub
    End If

    ' Process every section from the cursor section to the end
    For lIdx = lStartSect To oSections.count

        Set oSection = oSections(lIdx)

        Set oFooter = oSection.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary) ' placeholder
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
        oRange.Fields.Add range:=oRange, _
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
           " through section " & oSections.count & ".", _
           vbInformation, "AddConsecutiveFootersFromCursor"

    ' Clean up
    Set oPara = Nothing
    Set oRange = Nothing
    Set oFooter = Nothing
    Set oSection = Nothing
    Set oSections = Nothing
    Set oDoc = Nothing
End Sub

Private Sub LinkFootersToPrevious()
    Dim oDoc      As Document
    Dim oSections As Sections
    Dim lIdx      As Long

    Set oDoc = ActiveDocument
    Set oSections = oDoc.Sections

    For lIdx = 2 To oSections.count
        oSections(lIdx).Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = True
    Next lIdx

    MsgBox "Done. Sections 2 through " & oSections.count & _
           " footers are now linked to previous.", _
           vbInformation, "LinkFootersToPrevious"

    Set oSections = Nothing
    Set oDoc = Nothing
End Sub


