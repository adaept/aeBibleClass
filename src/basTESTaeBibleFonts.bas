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



