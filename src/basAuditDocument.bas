Attribute VB_Name = "basAuditDocument"
Option Explicit
Option Compare Text
Option Private Module

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
            Print #f, "  Page Size: " & .pageWidth & " x " & .PageHeight
            
            Print #f, "  Margins (T/B/L/R): " & _
                .TopMargin & "/" & .BottomMargin & "/" & _
                .leftMargin & "/" & .rightMargin
            
            Print #f, "  Orientation: " & _
                IIf(.Orientation = wdOrientPortrait, "Portrait", "Landscape")
            
            Print #f, "  Columns: " & .TextColumns.count
            
            If .TextColumns.count > 1 Then
                Dim c As Long
                For c = 1 To .TextColumns.count
                    Print #f, "    Column " & c & _
                        " Width: " & .TextColumns(c).Width & _
                        " Space: " & .TextColumns(c).SpaceAfter
                Next c
            End If
        End With
        
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

