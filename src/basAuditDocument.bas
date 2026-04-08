Attribute VB_Name = "basAuditDocument"
Option Explicit
Option Compare Text
Option Private Module

Public Sub ReplaceTimesInStyles()
    On Error GoTo PROC_ERR
    Dim oDoc   As Document
    Dim oStyle As style
    Dim lCount As Long

    Set oDoc = ActiveDocument
    lCount = 0

    For Each oStyle In oDoc.Styles
        On Error Resume Next
        If Trim(oStyle.Font.name) = "Times" Then
            oStyle.Font.name = "Times New Roman"
            lCount = lCount + 1
        End If
        On Error GoTo 0
        On Error GoTo PROC_ERR
    Next oStyle

    Debug.Print "Done. Replaced Times with Times New Roman in " & _
           lCount & " style definitions."

    MsgBox "Done. Replaced Times with Times New Roman in " & _
           lCount & " style definitions.", _
           vbInformation, "ReplaceTimesInStyles"

    Set oDoc = Nothing

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ReplaceTimesInStyles of Module basAuditDocument"
    Resume PROC_EXIT
End Sub

Public Sub FindFontUsage()
    On Error GoTo PROC_ERR
    Dim oDoc     As Document
    Dim oPara    As Word.Paragraph
    Dim oSection As Word.Section
    Dim oHF      As HeaderFooter
    Dim oStyle   As style
    Dim sTarget  As String
    Dim lBody    As Long
    Dim lHF      As Long
    Dim oStyles  As New Collection
    Dim bFound   As Boolean
    Dim i        As Long

    Set oDoc = ActiveDocument
    sTarget = "Times"           'Arial Unicode MS"
    lBody = 0
    lHF = 0

    ' -- Body paragraphs ------------------------------------------------------
    For Each oPara In oDoc.Paragraphs
        If InStr(1, ResolveFont(oPara), sTarget, vbTextCompare) > 0 Then
            lBody = lBody + 1
            bFound = False
            For i = 1 To oStyles.Count
                If oStyles(i) = "(body) " & oPara.style.NameLocal Then
                    bFound = True
                    Exit For
                End If
            Next i
            If Not bFound Then oStyles.Add "(body) " & oPara.style.NameLocal
        End If
    Next oPara

    ' -- Header and footer paragraphs -----------------------------------------
    For Each oSection In oDoc.Sections
        For Each oHF In oSection.Headers
            If oHF.Exists Then
                For Each oPara In oHF.Range.Paragraphs
                    If InStr(1, ResolveFont(oPara), sTarget, vbTextCompare) > 0 Then
                        lHF = lHF + 1
                        bFound = False
                        For i = 1 To oStyles.Count
                            If oStyles(i) = "(header) " & oPara.style.NameLocal Then
                                bFound = True
                                Exit For
                            End If
                        Next i
                        If Not bFound Then oStyles.Add "(header) " & oPara.style.NameLocal
                    End If
                Next oPara
            End If
        Next oHF
        For Each oHF In oSection.Footers
            If oHF.Exists Then
                For Each oPara In oHF.Range.Paragraphs
                    If InStr(1, ResolveFont(oPara), sTarget, vbTextCompare) > 0 Then
                        lHF = lHF + 1
                        bFound = False
                        For i = 1 To oStyles.Count
                            If oStyles(i) = "(footer) " & oPara.style.NameLocal Then
                                bFound = True
                                Exit For
                            End If
                        Next i
                        If Not bFound Then oStyles.Add "(footer) " & oPara.style.NameLocal
                    End If
                Next oPara
            End If
        Next oHF
    Next oSection

    ' -- Style definitions -----------------------------------------------------
    Dim lStyleDef As Long
    lStyleDef = 0
    For Each oStyle In oDoc.Styles
        On Error Resume Next
        Dim sStyleFont As String
        sStyleFont = Trim(oStyle.Font.name)
        On Error GoTo 0
        On Error GoTo PROC_ERR
        If InStr(1, sStyleFont, sTarget, vbTextCompare) > 0 Then
            lStyleDef = lStyleDef + 1
            bFound = False
            For i = 1 To oStyles.Count
                If oStyles(i) = "(style def) " & oStyle.NameLocal Then
                    bFound = True
                    Exit For
                End If
            Next i
            If Not bFound Then oStyles.Add "(style def) " & oStyle.NameLocal
        End If
    Next oStyle

    Dim sList As String
    sList = ""
    For i = 1 To oStyles.Count
        sList = sList & "  " & oStyles(i) & vbCrLf
    Next i

    Debug.Print "Font searched: " & sTarget & vbCrLf & vbCrLf & _
           "Body paragraphs: " & lBody & vbCrLf & _
           "Header/footer paragraphs: " & lHF & vbCrLf & _
           "Style definitions: " & lStyleDef & vbCrLf & vbCrLf & _
           IIf(oStyles.Count > 0, "Found in:" & vbCrLf & sList, "Not found anywhere.")

    MsgBox "Font searched: " & sTarget & vbCrLf & vbCrLf & _
           "Body paragraphs: " & lBody & vbCrLf & _
           "Header/footer paragraphs: " & lHF & vbCrLf & _
           "Style definitions: " & lStyleDef & vbCrLf & vbCrLf & _
           IIf(oStyles.Count > 0, "Found in:" & vbCrLf & sList, "Not found anywhere."), _
           vbInformation, "FindFontUsage"

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FindFontUsage of Module basAuditDocument"
    Resume PROC_EXIT
End Sub

Public Sub CountParagraphsAndFonts()
    On Error GoTo PROC_ERR
    Dim oDoc       As Document
    Dim oPara      As Word.Paragraph
    Dim oSection   As Word.Section
    Dim oHF        As HeaderFooter
    Dim oFonts     As New Collection
    Dim lBody      As Long
    Dim lHF        As Long
    Dim sFont      As String
    Dim bFound     As Boolean
    Dim i          As Long

    Set oDoc = ActiveDocument
    lBody = 0
    lHF = 0

    ' -- Body paragraphs ------------------------------------------------------
    For Each oPara In oDoc.Paragraphs
        lBody = lBody + 1
        sFont = ResolveFont(oPara)
        AddToCollection oFonts, sFont
    Next oPara

    ' -- Header and footer paragraphs -----------------------------------------
    For Each oSection In oDoc.Sections
        For Each oHF In oSection.Headers
            If oHF.Exists Then
                For Each oPara In oHF.Range.Paragraphs
                    lHF = lHF + 1
                    sFont = ResolveFont(oPara)
                    AddToCollection oFonts, sFont
                Next oPara
            End If
        Next oHF
        For Each oHF In oSection.Footers
            If oHF.Exists Then
                For Each oPara In oHF.Range.Paragraphs
                    lHF = lHF + 1
                    sFont = ResolveFont(oPara)
                    AddToCollection oFonts, sFont
                Next oPara
            End If
        Next oHF
    Next oSection

    Dim sFontList As String
    sFontList = ""
    For i = 1 To oFonts.Count
        sFontList = sFontList & "  " & oFonts(i) & vbCrLf
    Next i

    Debug.Print "Body paragraphs: " & lBody & vbCrLf & _
           "Header/footer paragraphs: " & lHF & vbCrLf & vbCrLf & _
           "Fonts used (" & oFonts.Count & "):" & vbCrLf & sFontList

    MsgBox "Body paragraphs: " & lBody & vbCrLf & _
           "Header/footer paragraphs: " & lHF & vbCrLf & vbCrLf & _
           "Fonts used (" & oFonts.Count & "):" & vbCrLf & sFontList, _
           vbInformation, "CountParagraphsAndFonts"

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure CountParagraphsAndFonts of Module basAuditDocument"
    Resume PROC_EXIT
End Sub

Private Function ResolveFont(ByVal oPara As Word.Paragraph) As String
    On Error GoTo PROC_ERR
    Dim sFont As String
    sFont = Trim(oPara.Range.Font.name)
    If Len(sFont) = 0 Or Left(sFont, 1) = "+" Then
        On Error Resume Next
        sFont = Trim(oPara.style.Font.name)
        On Error GoTo 0
        On Error GoTo PROC_ERR
    End If
    If Len(sFont) = 0 Or Left(sFont, 1) = "+" Then
        sFont = "(inherited/theme)"
    End If
    ResolveFont = sFont

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ResolveFont of Module basAuditDocument"
    Resume PROC_EXIT
End Function

Private Sub AddToCollection(ByRef oCol As Collection, ByVal sValue As String)
    Dim i As Long
    For i = 1 To oCol.Count
        If oCol(i) = sValue Then Exit Sub
    Next i
    oCol.Add sValue
End Sub

Public Sub CountFields()
    On Error GoTo PROC_ERR
    Dim lCount As Long
    lCount = ActiveDocument.Fields.Count
    Debug.Print "Total fields in document: " & lCount
    MsgBox "Total fields in document: " & lCount, vbInformation, "CountFields"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure CountFields of Module basAuditDocument"
    Resume PROC_EXIT
End Sub

Public Sub CountCodeLines()
    On Error GoTo PROC_ERR
    Dim oComp       As Object
    Dim oModule     As Object
    Dim lTotalCode  As Long
    Dim lTotalComment As Long
    Dim lTotalEmpty As Long
    Dim lCode       As Long
    Dim lComment    As Long
    Dim lEmpty      As Long
    Dim lIdx        As Long
    Dim sLine       As String
    Dim sTrimmed    As String

    lTotalCode = 0
    lTotalComment = 0
    lTotalEmpty = 0

    Debug.Print String(70, "-")
    Debug.Print PadRight("Module", 35) & _
                PadRight("Code", 10) & _
                PadRight("Comments", 10) & _
                PadRight("Empty", 10) & _
                "Total"
    Debug.Print String(70, "-")

    For Each oComp In ThisDocument.VBProject.VBComponents

        Set oModule = oComp.CodeModule
        lCode = 0
        lComment = 0
        lEmpty = 0

        For lIdx = 1 To oModule.CountOfLines
            sLine = oModule.lines(lIdx, 1)
            sTrimmed = Trim(sLine)

            If Len(sTrimmed) = 0 Then
                lEmpty = lEmpty + 1
            ElseIf Left(sTrimmed, 1) = "'" Then
                lComment = lComment + 1
            Else
                lCode = lCode + 1
            End If

        Next lIdx

        lTotalCode = lTotalCode + lCode
        lTotalComment = lTotalComment + lComment
        lTotalEmpty = lTotalEmpty + lEmpty

        Debug.Print PadRight(oComp.name, 35) & _
                    PadRight(lCode, 10) & _
                    PadRight(lComment, 10) & _
                    PadRight(lEmpty, 10) & _
                    (lCode + lComment + lEmpty)

    Next oComp

    Debug.Print String(70, "-")
    Debug.Print PadRight("TOTAL", 35) & _
                PadRight(lTotalCode, 10) & _
                PadRight(lTotalComment, 10) & _
                PadRight(lTotalEmpty, 10) & _
                (lTotalCode + lTotalComment + lTotalEmpty)
    Debug.Print String(70, "-")

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure CountCodeLines of Module basAuditDocument"
    Resume PROC_EXIT
End Sub

Private Function PadRight(ByVal sValue As Variant, ByVal lWidth As Long) As String
    Dim sStr As String
    sStr = CStr(sValue)
    If Len(sStr) < lWidth Then
        PadRight = sStr & space(lWidth - Len(sStr))
    Else
        PadRight = Left(sStr, lWidth)
        End If
End Function

Public Sub CountOrphanFooters()
    On Error GoTo PROC_ERR
    Dim oDoc        As Document
    Dim oSection    As Word.Section
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
               oFooter.Range.Fields.Count = 0 Then
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

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure CountOrphanFooters of Module basAuditDocument"
    Resume PROC_EXIT
End Sub

Public Sub CountOrphanHeaders()
    On Error GoTo PROC_ERR
    Dim oDoc         As Document
    Dim oSection     As Word.Section
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
               oHeader.Range.Fields.Count = 0 Then
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

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure CountOrphanHeaders of Module basAuditDocument"
    Resume PROC_EXIT
End Sub

'==========================================
' Entry Points
'==========================================
Public Sub AuditDoc_Original()
    On Error GoTo PROC_ERR
    WriteAuditToFile "AuditDoc_Original.txt", "ORIGINAL"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure AuditDoc_Original of Module basAuditDocument"
    Resume PROC_EXIT
End Sub

Public Sub AuditDoc_New()
    On Error GoTo PROC_ERR
    WriteAuditToFile "AuditDoc_New.txt", "NEW"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure AuditDoc_New of Module basAuditDocument"
    Resume PROC_EXIT
End Sub

'==========================================
' Core Writer
'==========================================
Private Sub WriteAuditToFile(ByVal fileName As String, ByVal label As String)
    On Error GoTo PROC_ERR
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

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure WriteAuditToFile of Module basAuditDocument"
    Resume PROC_EXIT
End Sub

'==========================================
' Path Helper
'==========================================
Private Function GetRptPath() As String
    On Error GoTo PROC_ERR
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

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GetRptPath of Module basAuditDocument"
    Resume PROC_EXIT
End Function

'==========================================
' Writers
'==========================================
Private Sub WriteHeader(ByVal f As Integer, ByVal label As String)
    On Error GoTo PROC_ERR
    Print #f, String(60, "=")
    Print #f, "DOCUMENT AUDIT: " & label
    Print #f, "File: " & ActiveDocument.name
    Print #f, "Path: " & ActiveDocument.FullName
    Print #f, "Timestamp: " & Now
    Print #f, String(60, "=")
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure WriteHeader of Module basAuditDocument"
    Resume PROC_EXIT
End Sub

Private Sub WriteDocumentStats(ByVal f As Integer)
    On Error GoTo PROC_ERR
    With ActiveDocument
        Print #f, "Total Sections: " & .Sections.Count
        Print #f, "Total Paragraphs: " & .Paragraphs.Count
        Print #f, "Total Words: " & .words.Count
        Print #f, "Total Characters: " & .Characters.Count
        Print #f, "Total Footnotes: " & .Footnotes.Count
        Print #f, "Total Endnotes: " & .Endnotes.Count
    End With

    Print #f, String(40, "-")
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure WriteDocumentStats of Module basAuditDocument"
    Resume PROC_EXIT
End Sub

Private Sub WriteSectionAudit(ByVal f As Integer)
    On Error GoTo PROC_ERR
    Dim sec As Word.Section

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
                Print #f, "  Columns: " & .Count
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

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure WriteSectionAudit of Module basAuditDocument"
    Resume PROC_EXIT
End Sub

Private Sub WriteSignature(ByVal f As Integer, ByVal label As String)
    On Error GoTo PROC_ERR
    Dim sec As Word.Section
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
                  "Cols=" & .TextColumns.Count
        End With
        
        Print #f, sig
        i = i + 1
    Next sec
    
    Print #f, "END SIGNATURE"
    Print #f, String(60, "=")

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure WriteSignature of Module basAuditDocument"
    Resume PROC_EXIT
End Sub

