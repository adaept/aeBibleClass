Attribute VB_Name = "basScripturePdfExport"
Option Explicit

'Copyright (c) 2025-2026 Peter F. Ennis
'SPDX-License-Identifier: LGPL-3.0-or-later OR LicenseRef-adaept-Commercial
'DUAL-LICENSED. You may use this file under EITHER of:
'  (1) the GNU Lesser General Public License, version 3.0 or (at your option)
'      any later version  -  https://www.gnu.org/licenses/lgpl-3.0.txt ; OR
'  (2) a commercial / proprietary license available from adaept (Peter Ennis),
'      permitting use in closed-source / proprietary works WITHOUT the LGPL
'      copyleft obligations. Contact the copyright holder for commercial terms.
'As the sole copyright holder, adaept may license this file under either option.
'This library is distributed WITHOUT ANY WARRANTY; without even the implied
'warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.

' ==========================================================================
' basScripturePdfExport  (RT-33, 2026-11-12)
' ==========================================================================
' The two PDFs of a Radiant Word Bible release, made from the FINISHED scripture-only .docx (the file that the export
' ExportScriptureOnlyDocx made and the Node steps rwb.fontsubset and rwb.fonts finished), never from the dev .docm:
'
'     rwb-print.pdf   print-ready: Word's "Standard (publishing online and printing)" quality (OptimizeFor print)
'     rwb.pdf         screen: Word's "Minimum size (publishing online)" quality (OptimizeFor on screen)
'
' Both carry heading bookmarks (the navigation pane of a PDF reader), document structure tags (tagged PDF) and the document
' properties. Fonts are embedded in the PDF by Word (the three free families only: the run stops first if one is not installed,
' RT-35). A PDF is NOT reproducible byte for byte (creation date and file id differ on every export): the manifest records its
' SHA-256 and its identity (pages, fonts), and rwb.verify checks the identity, not the bytes (RT-02).
'
' The .docx is opened READ-ONLY and closed without saving: it is never changed. The routine writes only the two PDFs and the
' report rpt\ScripturePdfExport.txt (beside the docx, or in the PDF folder when it has no rpt folder).
'
' Entry point (Immediate window, one statement):
'     ExportScripturePdf "C:\adaept\aeTest\x\RadiantWordBible.docx"            PDFs next to the docx
'     ExportScripturePdf "C:\adaept\aeTest\x\RadiantWordBible.docx", "C:\adaept\aeTest\x\pdf"    PDFs in that folder
' Run it on a COPY under C:\adaept\aeTest\ first (tools\word-com\run-pdf.vbs does that in a hidden Word). About 100 seconds a PDF.
' ==========================================================================

Private Const PRINT_PDF_NAME As String = "rwb-print.pdf"
Private Const SCREEN_PDF_NAME As String = "rwb.pdf"
Private Const REPORT_NAME As String = "ScripturePdfExport.txt"
Private Const EXPORT_FORMAT_PDF As Long = 17
Private Const OPTIMIZE_FOR_PRINT As Long = 0
Private Const OPTIMIZE_FOR_SCREEN As Long = 1
Private Const BOOKMARKS_FROM_HEADINGS As Long = 1
Private Const REPORT_BYTES_PER_KB As Long = 1024

Private m_report As String
Private m_issues As Long
Private m_reportFolder As String

Private Sub PdfLog(ByVal sLine As String)
    m_report = m_report & sLine & vbCrLf
    Debug.Print sLine
End Sub

' One PDF of the open document. Returns True when the file exists afterwards and is not empty.
Private Function ExportOnePdf(ByVal oDoc As Object, ByVal pdfPath As String, ByVal qualityMode As Long, _
                              ByVal stepName As String) As Boolean
    Dim startTime As Single
    startTime = Timer
    If Len(Dir(pdfPath)) > 0 Then Kill pdfPath
    oDoc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=EXPORT_FORMAT_PDF, OpenAfterExport:=False, _
        OptimizeFor:=qualityMode, Range:=wdExportAllDocument, Item:=wdExportDocumentContent, _
        IncludeDocProps:=True, CreateBookmarks:=BOOKMARKS_FROM_HEADINGS, _
        DocStructureTags:=True, BitmapMissingFonts:=False, UseISO19005_1:=False
    If Len(Dir(pdfPath)) = 0 Then
        PdfLog stepName & ": FAILED - no file was written: " & pdfPath
        Exit Function
    End If
    If FileLen(pdfPath) = 0 Then
        PdfLog stepName & ": FAILED - the file is empty: " & pdfPath
        Exit Function
    End If
    PdfLog stepName & ": wrote " & pdfPath & " (" & Format(FileLen(pdfPath) / REPORT_BYTES_PER_KB, "#,##0") & " KB, " & _
        Format(Timer - startTime, "0") & " s)"
    ExportOnePdf = True
End Function

Private Sub WritePdfReport()
    Dim reportPath As String
    Dim fileNumber As Integer
    If Len(m_reportFolder) = 0 Then Exit Sub
    reportPath = m_reportFolder & "\" & REPORT_NAME
    fileNumber = FreeFile
    Open reportPath For Output As #fileNumber
    Print #fileNumber, m_report;
    Print #fileNumber, "Result: " & IIf(m_issues = 0, "CLEAN (0 issues)", "NOT CLEAN (" & m_issues & " issue(s))")
    Close #fileNumber
End Sub

Public Sub ExportScripturePdf(ByVal docxPath As String, Optional ByVal pdfFolder As String)
    On Error GoTo PROC_ERR

    Dim oDoc As Object
    Dim missingNames As String
    Dim keepAwakeOn As Boolean
    Dim docxFolder As String

    m_report = ""
    m_issues = 0
    m_reportFolder = ""

    If LCase$(Right$(docxPath, 5)) <> ".docx" Then
        Debug.Print "ExportScripturePdf: give the path of the finished .docx, not a .docm or another file: " & docxPath
        Exit Sub
    End If
    If Len(Dir(docxPath)) = 0 Then
        Debug.Print "ExportScripturePdf: file not found: " & docxPath
        Exit Sub
    End If
    docxFolder = Left$(docxPath, InStrRev(docxPath, "\") - 1)
    If Len(pdfFolder) = 0 Then pdfFolder = docxFolder
    If Len(Dir(pdfFolder, vbDirectory)) = 0 Then MkDir pdfFolder
    If Len(Dir(docxFolder & "\rpt", vbDirectory)) > 0 Then m_reportFolder = docxFolder & "\rpt" Else m_reportFolder = pdfFolder

    PdfLog "---- RadiantWordBible PDF export " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ----"
    PdfLog "source docx=" & docxPath
    PdfLog "pdf folder=" & pdfFolder

    ' RT-35 preflight: without the three free fonts Word would substitute silently in the PDF.
    missingNames = basFontRunAudit.MissingAllowedFonts()
    If Len(missingNames) > 0 Then
        m_issues = m_issues + 1
        PdfLog "Pass0_FontsInstalled: STOPPED - not installed on this machine: " & missingNames & ". Nothing was exported."
        GoTo PROC_EXIT
    End If
    PdfLog "Pass0_FontsInstalled: Carlito, Noto Sans and Liberation Serif are installed."

    keepAwakeOn = StartKeepAwake
    PdfLog "keep-awake: " & IIf(keepAwakeOn, "ON", "FAILED (the machine may sleep on its idle timer)")

    ' Read-only, not added to the recent files: the finished docx must stay byte for byte as it is.
    Set oDoc = Documents.Open(FileName:=docxPath, ReadOnly:=True, AddToRecentFiles:=False, Visible:=False, _
        ConfirmConversions:=False)
    PdfLog "opened read-only: " & oDoc.Name & " (macros: " & IIf(oDoc.HasVBProject, "PRESENT - not a finished docx", "none") & ")"
    If oDoc.HasVBProject Then
        m_issues = m_issues + 1
        PdfLog "STOPPED - the file carries a VBA project; the PDFs come from the finished .docx only."
        GoTo PROC_CLOSE
    End If

    If Not ExportOnePdf(oDoc, pdfFolder & "\" & SCREEN_PDF_NAME, OPTIMIZE_FOR_SCREEN, "Pass1_ScreenPdf") Then m_issues = m_issues + 1
    If Not ExportOnePdf(oDoc, pdfFolder & "\" & PRINT_PDF_NAME, OPTIMIZE_FOR_PRINT, "Pass2_PrintPdf") Then m_issues = m_issues + 1

PROC_CLOSE:
    oDoc.Close SaveChanges:=wdDoNotSaveChanges
    Set oDoc = Nothing
    If m_issues = 0 Then
        PdfLog "ExportScripturePdf: COMPLETE. Next: node tools\release-harness\pdf-check.mjs on each PDF (aeRWB), then rwb.verify."
    Else
        PdfLog "ExportScripturePdf: FINISHED WITH " & m_issues & " ISSUE(S). Do not use these PDFs."
    End If

PROC_EXIT:
    If keepAwakeOn Then StopKeepAwake
    WritePdfReport
    Exit Sub
PROC_ERR:
    m_issues = m_issues + 1
    PdfLog "ERROR in basScripturePdfExport.ExportScripturePdf | Err: " & Err.Number & " | " & Err.Description
    On Error Resume Next
    If Not oDoc Is Nothing Then oDoc.Close SaveChanges:=wdDoNotSaveChanges
    Resume PROC_EXIT
End Sub
