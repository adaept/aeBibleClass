Attribute VB_Name = "basAuditDocument"
Option Explicit
Option Compare Text
Option Private Module

Public Sub AddBookNameHeaders()
    Dim oDoc        As Document
    Dim oSections   As Sections
    Dim oSection    As section
    Dim oHeader     As HeaderFooter
    Dim oPara As Word.Paragraph
    Dim oRange      As Range
    Dim lStartSect  As Long
    Dim lIdx        As Long
    Dim sBookName   As String
    Dim lResponse   As Long

    lResponse = MsgBox("Place cursor in the section to start header labelling. Do you want to start?", _
                       vbYesNo + vbDefaultButton2 + vbQuestion, _
                       "AddBookNameHeaders")

    If lResponse = vbNo Then Exit Sub

    Set oDoc = ActiveDocument
    Set oSections = oDoc.Sections

    ' -- Find the section containing the cursor --------------------------------
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
               vbExclamation, "AddBookNameHeaders"
        Exit Sub
    End If

    sBookName = ""

    ' -- Walk sections from cursor to end -------------------------------------
    For lIdx = lStartSect To oSections.count

        Set oSection = oSections(lIdx)
        Set oHeader = oSection.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        Set oPara = oSection.Range.Paragraphs(1)

        If oPara.style = oDoc.Styles("Heading 1") Then

            ' Book title page � clear the header and leave it empty
            oHeader.LinkToPrevious = False
            oHeader.Range.Delete

        ElseIf oPara.style = oDoc.Styles("Heading 2") Then

            ' First chapter section � capture book name from Heading 1 text
            ' Search backwards from this section for the nearest Heading 1
            Dim oSearch As Range
            Set oSearch = oDoc.Range(0, oSection.Range.Start)
            Dim oFound  As Range
            Set oFound = Nothing

            Dim pIdx    As Long
            For pIdx = oSearch.Paragraphs.count To 1 Step -1
                If oSearch.Paragraphs(pIdx).style = oDoc.Styles("Heading 1") Then
                    sBookName = Trim(oSearch.Paragraphs(pIdx).Range.Text)
                    Exit For
                End If
            Next pIdx

            ' Write the book name into the header
            oHeader.LinkToPrevious = False
            oHeader.Range.Delete

            Set oRange = oHeader.Range
            oRange.Text = sBookName
            oRange.ParagraphFormat.style = oDoc.Styles("TheHeaders")

        Else

            ' All other sections inherit the header from the section above
            oHeader.LinkToPrevious = True

        End If

    Next lIdx

    MsgBox "Done. Book name headers have been added from section " & _
           lStartSect & " through section " & oSections.count & ".", _
           vbInformation, "AddBookNameHeaders"

    Set oRange = Nothing
    Set oPara = Nothing
    Set oHeader = Nothing
    Set oSection = Nothing
    Set oSections = Nothing
    Set oDoc = Nothing
End Sub

Public Sub CountOrphanFooters()
    Dim oDoc        As Document
    Dim oSection    As section
    Dim oFooter     As HeaderFooter
    Dim lOrphan     As Long
    Dim lTotal      As Long
    Dim lIndependent As Long
    Dim sOrphanNames As String

    Set oDoc = ActiveDocument
    lOrphan = 0
    lTotal = 0
    lIndependent = 0
    sOrphanNames = ""

    For Each oSection In oDoc.Sections
        Set oFooter = oSection.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        lTotal = lTotal + 1
        If oFooter.LinkToPrevious = False Then
            lIndependent = lIndependent + 1
            If Len(Trim(oFooter.Range.Text)) = 0 And _
               oFooter.Range.Fields.count = 0 Then
                lOrphan = lOrphan + 1
                sOrphanNames = sOrphanNames & "  footer" & lIndependent & ".xml" & vbCrLf
            End If
        End If
    Next oSection

    Debug.Print "Total footer slots: " & lTotal & vbCrLf & _
           "Independent footers: " & lIndependent & vbCrLf & _
           "Orphaned footers (unlinked and empty): " & lOrphan & vbCrLf & vbCrLf & _
           IIf(lOrphan > 0, "Orphaned XML files:" & vbCrLf & sOrphanNames, "")

    MsgBox "Total footer slots: " & lTotal & vbCrLf & _
           "Independent footers: " & lIndependent & vbCrLf & _
           "Orphaned footers (unlinked and empty): " & lOrphan & vbCrLf & vbCrLf & _
           IIf(lOrphan > 0, "Orphaned XML files:" & vbCrLf & sOrphanNames, ""), _
           vbInformation, "CountOrphanFooters"

    Set oFooter = Nothing
    Set oSection = Nothing
    Set oDoc = Nothing
End Sub

Public Sub CountOrphanHeaders()
    Dim oDoc         As Document
    Dim oSection     As section
    Dim oHeader      As HeaderFooter
    Dim lOrphan      As Long
    Dim lTotal       As Long
    Dim lIndependent As Long
    Dim sOrphanNames As String

    Set oDoc = ActiveDocument
    lOrphan = 0
    lTotal = 0
    lIndependent = 0
    sOrphanNames = ""

    For Each oSection In oDoc.Sections
        Set oHeader = oSection.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        lTotal = lTotal + 1
        If oHeader.LinkToPrevious = False Then
            lIndependent = lIndependent + 1
            If Len(Trim(oHeader.Range.Text)) = 0 And _
               oHeader.Range.Fields.count = 0 Then
                lOrphan = lOrphan + 1
                sOrphanNames = sOrphanNames & "  header" & lIndependent & ".xml" & vbCrLf
            End If
        End If
    Next oSection

    Debug.Print "Total header slots: " & lTotal & vbCrLf & _
           "Independent headers: " & lIndependent & vbCrLf & _
           "Orphaned headers (unlinked and empty): " & lOrphan & vbCrLf & vbCrLf & _
           IIf(lOrphan > 0, "Orphaned XML files:" & vbCrLf & sOrphanNames, "")

    MsgBox "Total header slots: " & lTotal & vbCrLf & _
           "Independent headers: " & lIndependent & vbCrLf & _
           "Orphaned headers (unlinked and empty): " & lOrphan & vbCrLf & vbCrLf & _
           IIf(lOrphan > 0, "Orphaned XML files:" & vbCrLf & sOrphanNames, ""), _
           vbInformation, "CountOrphanHeaders"

    Set oHeader = Nothing
    Set oSection = Nothing
    Set oDoc = Nothing
End Sub

'==========================================
' Entry Points
'==========================================
Public Sub AuditDoc_Original()
    WriteAuditToFile "AuditDoc_Original.txt", "ORIGINAL"
End Sub

Public Sub AuditDoc_New()
    WriteAuditToFile "AuditDoc_New.txt", "NEW"
End Sub

'==========================================
' Core Writer
'==========================================
Private Sub WriteAuditToFile(ByVal fileName As String, ByVal label As String)
    Dim filePath As String
    filePath = GetRptPath() & fileName
    
    Dim f As Integer
    f = FreeFile
    
    Open filePath For Output As #f
    
    WriteHeader f, label
    WriteDocumentStats f
    WriteSectionAudit f
    WriteSignature f, label
    
    Close #f
    
    MsgBox "Audit written to:" & vbCrLf & filePath, vbInformation
End Sub

'==========================================
' Path Helper
'==========================================
Private Function GetRptPath() As String
    Dim basePath As String
    basePath = ActiveDocument.Path
    
    If basePath = "" Then
        Err.Raise vbObjectError + 1, , "Document must be saved before auditing."
    End If
    
    Dim rptPath As String
    rptPath = basePath & "\rpt\"
    
    ' Create folder if it does not exist
    If Dir(rptPath, vbDirectory) = "" Then MkDir rptPath
    
    GetRptPath = rptPath
End Function

'==========================================
' Writers
'==========================================
Private Sub WriteHeader(ByVal f As Integer, ByVal label As String)
    Print #f, String(60, "=")
    Print #f, "DOCUMENT AUDIT: " & label
    Print #f, "File: " & ActiveDocument.name
    Print #f, "Path: " & ActiveDocument.FullName
    Print #f, "Timestamp: " & Now
    Print #f, String(60, "=")
End Sub

Private Sub WriteDocumentStats(ByVal f As Integer)
    With ActiveDocument
        Print #f, "Total Sections: " & .Sections.count
        Print #f, "Total Paragraphs: " & .Paragraphs.count
        Print #f, "Total Words: " & .words.count
        Print #f, "Total Characters: " & .Characters.count
        Print #f, "Total Footnotes: " & .Footnotes.count
        Print #f, "Total Endnotes: " & .Endnotes.count
    End With
    
    Print #f, String(40, "-")
End Sub

Private Sub WriteSectionAudit(ByVal f As Integer)
    Dim sec As section
    
    Dim i As Long
    i = 1
    
    For Each sec In ActiveDocument.Sections
        
        Print #f, "Section " & i
        
        With sec.PageSetup
            ' Basic section properties
            Print #f, "  Page Size: " & .pageWidth & " x " & .PageHeight
            Print #f, "  Margins (T/B/L/R): " & _
                .TopMargin & "/" & .BottomMargin & "/" & _
                .leftMargin & "/" & .rightMargin
            Print #f, "  Orientation: " & _
                IIf(.Orientation = wdOrientPortrait, "Portrait", "Landscape")
            
            ' Columns (use only properties exposed on the collection)
            With .TextColumns
                Print #f, "  Columns: " & .count
                Print #f, "  EvenlySpaced: " & .EvenlySpaced
                Print #f, "  LineBetween: " & .LineBetween
                Print #f, "  Width: " & .Width
            End With
        End With
        
        ' Header/Footer linkage
        Print #f, "  Header LinkToPrevious: " & _
            sec.Headers(wdHeaderFooterPrimary).LinkToPrevious
        Print #f, "  Footer LinkToPrevious: " & _
            sec.Footers(wdHeaderFooterPrimary).LinkToPrevious
        
        Print #f, String(30, "-")
        
        i = i + 1
    Next sec
End Sub

Private Sub WriteSignature(ByVal f As Integer, ByVal label As String)
    Dim sec As section
    Dim i As Long
    Dim sig As String
    
    Print #f, ""
    Print #f, "SIGNATURE: " & label
    
    i = 1
    For Each sec In ActiveDocument.Sections
        With sec.PageSetup
            sig = "S" & i & "|" & _
                  .pageWidth & "x" & .PageHeight & "|" & _
                  .TopMargin & "," & .BottomMargin & "," & _
                  .leftMargin & "," & .rightMargin & "|" & _
                  "Cols=" & .TextColumns.count
        End With
        
        Print #f, sig
        i = i + 1
    Next sec
    
    Print #f, "END SIGNATURE"
    Print #f, String(60, "=")
End Sub

