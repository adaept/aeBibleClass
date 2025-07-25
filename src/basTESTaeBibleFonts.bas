Attribute VB_Name = "basTESTaeBibleFonts"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Sub CheckOpenFontsWithDownloads()
    Dim fontList As Variant
    Dim fontName As Variant
    Dim fontStatus As String
    Dim InstalledFonts As String
    Dim MissingFonts As String
    Dim DownloadLinks As String
    Dim FontInstalled As Boolean
    Dim FontURL As String

    ' Array of open-source fonts and their download URLs
    fontList = Array( _
        Array("Libre Franklin", "https://fonts.google.com/specimen/Libre+Franklin"), _
        Array("Noto Sans", "https://fonts.google.com/specimen/Noto+Sans"), _
        Array("Roboto", "https://fonts.google.com/specimen/Roboto"), _
        Array("Libre Baskerville", "https://fonts.google.com/specimen/Libre+Baskerville"), _
        Array("Source Sans 3", "https://fonts.google.com/specimen/Source+Sans+3") _
    )

    InstalledFonts = ""
    MissingFonts = ""
    DownloadLinks = ""

    Dim i As Integer
    For i = LBound(fontList) To UBound(fontList)
        fontName = fontList(i)(0)
        FontURL = fontList(i)(1)
        FontInstalled = IsFontInstalled(CStr(fontName))

        If FontInstalled Then
            InstalledFonts = InstalledFonts & "> " & fontName & vbCrLf
        Else
            MissingFonts = MissingFonts & "X " & fontName & vbCrLf
            DownloadLinks = DownloadLinks & fontName & ": " & FontURL & vbCrLf
        End If
    Next i

    'MsgBox "Open Font Availability Report:" & vbCrLf & vbCrLf & _
           "Installed Fonts:" & vbCrLf & InstalledFonts & vbCrLf & _
           "Missing Fonts:" & vbCrLf & MissingFonts & _
           IIf(DownloadLinks <> "", vbCrLf & "Download Missing Fonts:" & vbCrLf & DownloadLinks, ""), _
           vbInformation, "Open Font Check"
    Debug.Print "Open Font Availability Report:" & vbCrLf & vbCrLf & _
           "Installed Fonts:" & vbCrLf & InstalledFonts & vbCrLf & _
           "Missing Fonts:" & vbCrLf & MissingFonts & _
           IIf(DownloadLinks <> "", vbCrLf & "Download Missing Fonts:" & vbCrLf & DownloadLinks, "")
End Sub

Function IsFontInstalled(fontName As String) As Boolean
    Dim TestDoc As Document
    Dim testRange As range
    On Error Resume Next
    Set TestDoc = Documents.Add(Visible:=False)
    Set testRange = TestDoc.content
    testRange.text = "Test"
    testRange.font.name = fontName
    IsFontInstalled = (testRange.font.name = fontName)
    TestDoc.Close SaveChanges:=False
    On Error GoTo 0
End Function

Sub CreateEmphasisBlackStyle()
    Dim charStyle As style
    
    ' Check if the style already exists
    On Error Resume Next
    Set charStyle = ActiveDocument.Styles("EmphasisBlack")
    On Error GoTo 0

    ' If the style doesn't exist, create it
    If charStyle Is Nothing Then
        Set charStyle = ActiveDocument.Styles.Add(name:="EmphasisBlack", Type:=wdStyleTypeCharacter)
    End If

    ' Apply formatting
    With charStyle.font
        .name = "Arial Black"
        .Size = 8
        .Bold = True
    End With

    ' Add to style gallery
    charStyle.Priority = 1
    
    charStyle.QuickStyle = True

    MsgBox "Character style 'EmphasisBlack' created and added to the Style Gallery.", vbInformation
End Sub

Sub AuditStyleUsage_Footnote()
    Dim r As range, hitCount As Long
    Dim logBuffer As String

    logBuffer = "=== Audit: Style Usage for ->Footnote<- ===" & vbCrLf

    Set r = ActiveDocument.content
    With r.Find
        .ClearFormatting
        .style = ActiveDocument.Styles("Footnote")
        .text = ""
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            hitCount = hitCount + 1
            logBuffer = logBuffer & "* Style hit at Char " & r.Start & " ? ->" & Left(r.text, 40) & "...<-" & vbCrLf
            r.Start = r.Start + 1
            r.End = ActiveDocument.content.End
        Loop
    End With

    logBuffer = logBuffer & vbCrLf & "Total ->Footnote<- style instances: " & hitCount
    Debug.Print logBuffer
    MsgBox "Audit complete. See Immediate Window for details.", vbInformation
End Sub

Sub RedefineFootnoteStyle_NotoSans()
    Dim s As style
    Set s = ActiveDocument.Styles("Footnote")

    With s.font
        .name = "Noto Sans"
        .Size = 7
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .color = wdColorAutomatic
    End With

    MsgBox "->Footnote<- style updated to Noto Sans, 7pt.", vbInformation
End Sub

Sub AuditStyleUsage_FootnoteNormal()
    Dim para As paragraph
    Dim hitCount As Long
    Dim logBuffer As String

    logBuffer = "=== Audit: Paragraph Style Usage for 'Footnote normal' ===" & vbCrLf

    For Each para In ActiveDocument.paragraphs
        If para.style = ActiveDocument.Styles("Footnote normal") Then
            hitCount = hitCount + 1
            logBuffer = logBuffer & "* Paragraph at Char " & para.range.Start & " -> """ & _
            Replace(Left(para.range.text, 40), vbCr, "") & "...""" & vbCrLf
        End If
    Next para

    logBuffer = logBuffer & vbCrLf & "Total 'Footnote normal' style instances: " & hitCount
    Debug.Print logBuffer
    MsgBox "Audit complete. See Immediate Window for details.", vbInformation
End Sub

Sub RedefineFootnoteNormalStyle_NotoSans()
    Dim s As style
    Set s = ActiveDocument.Styles("Footnote normal")

    With s.font
        .name = "Noto Sans"
        .Size = 7
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .color = wdColorAutomatic
    End With

    MsgBox "'Footnote normal' style updated to Noto Sans, 7pt.", vbInformation
End Sub

Sub AuditStyleUsage_PictureCaption()
    Dim para As paragraph
    Dim hitCount As Long
    Dim logBuffer As String
    Dim s As style

    logBuffer = "=== Audit: Paragraph Style Usage for 'Picture Caption' ===" & vbCrLf

    On Error Resume Next
    Set s = ActiveDocument.Styles("Picture Caption")
    If s Is Nothing Then
        MsgBox "Style 'Picture Caption' not found in this document.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    For Each para In ActiveDocument.paragraphs
        If para.style = s Then
            hitCount = hitCount + 1
            logBuffer = logBuffer & "* Paragraph at Char " & para.range.Start & " -> """ & _
                Replace(Left(para.range.text, 40), vbCr, "") & "...""" & vbCrLf
        End If
    Next para

    logBuffer = logBuffer & vbCrLf & "Total 'Picture Caption' style instances: " & hitCount
    Debug.Print logBuffer
    MsgBox "Audit complete. See Immediate Window for details.", vbInformation
End Sub

Sub RedefinePictureCaptionStyle_NotoSans()
    Dim s As style
    On Error Resume Next
    Set s = ActiveDocument.Styles("Picture Caption")
    If s Is Nothing Then
        MsgBox "'Picture Caption' style not found in this document.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    With s.font
        .name = "Noto Sans"
        .Size = 9
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .color = wdColorAutomatic
    End With

    Debug.Print "? 'Picture Caption' style updated to Noto Sans, 9pt."
    MsgBox "'Picture Caption' style redefined to Noto Sans, 9pt.", vbInformation
End Sub

Sub Identify_ArialUnicodeMS_Paragraphs()
    Dim para As paragraph
    Dim paraIndex As Long
    Dim secIndex As Long
    Dim hfIndex As Long
    Dim hfTypes As Variant
    Dim hfKind As Variant
    Dim logBuffer As String
    Dim sec As section
    Dim hf As HeaderFooter
    Dim fontName As String

    logBuffer = "=== Arial Unicode MS Paragraph Identification ===" & vbCrLf

    ' Scan body
    paraIndex = 0
    For Each para In ActiveDocument.paragraphs
        paraIndex = paraIndex + 1
        fontName = para.range.Characters(1).font.name
        If fontName = "Arial Unicode MS" Then
            logBuffer = logBuffer & "[Body] Para #" & paraIndex & " - Style: " & para.style & vbCrLf
            logBuffer = logBuffer & "Text: " & Left(para.range.text, 120) & vbCrLf & vbCrLf
        End If
    Next para

    ' Header/footer types
    hfTypes = Array(wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages)

    ' Scan headers/footers
    secIndex = 0
    For Each sec In ActiveDocument.Sections
        secIndex = secIndex + 1
        For Each hfKind In hfTypes
            Set hf = sec.Headers(hfKind)
            If hf.Exists Then
                paraIndex = 0
                For Each para In hf.range.paragraphs
                    paraIndex = paraIndex + 1
                    fontName = para.range.Characters(1).font.name
                    If fontName = "Arial Unicode MS" Then
                        logBuffer = logBuffer & "[Header] Sec " & secIndex & ", Type " & hfKind & ", Para #" & paraIndex & vbCrLf
                        logBuffer = logBuffer & "Style: " & para.style & vbCrLf
                        logBuffer = logBuffer & "Text: " & Left(para.range.text, 120) & vbCrLf & vbCrLf
                    End If
                Next
            End If
            Set hf = sec.Footers(hfKind)
            If hf.Exists Then
                paraIndex = 0
                For Each para In hf.range.paragraphs
                    paraIndex = paraIndex + 1
                    fontName = para.range.Characters(1).font.name
                    If fontName = "Arial Unicode MS" Then
                        logBuffer = logBuffer & "[Footer] Sec " & secIndex & ", Type " & hfKind & ", Para #" & paraIndex & vbCrLf
                        logBuffer = logBuffer & "Style: " & para.style & vbCrLf
                        logBuffer = logBuffer & "Text: " & Left(para.range.text, 120) & vbCrLf & vbCrLf
                    End If
                Next
            End If
        Next hfKind
    Next sec

    Debug.Print logBuffer
    MsgBox "Arial Unicode MS detection complete. See Immediate Window.", vbInformation
End Sub

