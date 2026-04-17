Attribute VB_Name = "basTEST_aeBibleFonts"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Public Sub CheckOpenFontsWithDownloads()
    On Error GoTo PROC_ERR
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

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure CheckOpenFontsWithDownloads of Module basTEST_aeBibleFonts"
    Resume PROC_EXIT
End Sub

Private Function IsFontInstalled(fontName As String) As Boolean
    On Error GoTo PROC_ERR
    Dim TestDoc As Document
    Dim testRange As Word.Range
    On Error Resume Next
    Set TestDoc = Documents.Add(Visible:=False)
    Set testRange = TestDoc.Content
    testRange.Text = "Test"
    testRange.Font.Name = fontName
    IsFontInstalled = (testRange.Font.Name = fontName)
    TestDoc.Close SaveChanges:=False
    On Error GoTo 0
    On Error GoTo PROC_ERR

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure IsFontInstalled of Module basTEST_aeBibleFonts"
    Resume PROC_EXIT
End Function

Public Sub CreateEmphasisBlackStyle()
    On Error GoTo PROC_ERR
    Dim charStyle As style

    ' Check if the style already exists
    On Error Resume Next
    Set charStyle = ActiveDocument.Styles("EmphasisBlack")
    On Error GoTo 0
    On Error GoTo PROC_ERR

    ' If the style doesn't exist, create it
    If charStyle Is Nothing Then
        Set charStyle = ActiveDocument.Styles.Add(name:="EmphasisBlack", Type:=wdStyleTypeCharacter)
    End If

    ' Apply formatting
    With charStyle.Font
        .Name = "Arial Black"
        .Size = 8
        .Bold = True
    End With

    ' Add to style gallery
    charStyle.Priority = 1

    charStyle.QuickStyle = True

    MsgBox "Character style 'EmphasisBlack' created and added to the Style Gallery.", vbInformation

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure CreateEmphasisBlackStyle of Module basTEST_aeBibleFonts"
    Resume PROC_EXIT
End Sub

Public Sub AuditStyleUsage_Footnote()
    On Error GoTo PROC_ERR
    Dim r As Word.Range, hitCount As Long
    Dim logBuffer As String

    logBuffer = "=== Audit: Style Usage for ->Footnote<- ===" & vbCrLf

    Set r = ActiveDocument.Content
    With r.Find
        .ClearFormatting
        .style = ActiveDocument.Styles("Footnote")
        .Text = ""
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            hitCount = hitCount + 1
            logBuffer = logBuffer & "* Style hit at Char " & r.Start & " - ->" & Left(r.Text, 40) & "...<-" & vbCrLf
            r.Start = r.Start + 1
            r.End = ActiveDocument.Content.End
        Loop
    End With

    logBuffer = logBuffer & vbCrLf & "Total ->Footnote<- style instances: " & hitCount
    Debug.Print logBuffer
    MsgBox "Audit complete. See Immediate Window for details.", vbInformation

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure AuditStyleUsage_Footnote of Module basTEST_aeBibleFonts"
    Resume PROC_EXIT
End Sub

Public Sub RedefineFootnoteStyle_NotoSans()
    On Error GoTo PROC_ERR
    Dim s As style
    Set s = ActiveDocument.Styles("Footnote")

    With s.Font
        .Name = "Noto Sans"
        .Size = 8
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineNone
        .color = wdColorAutomatic
    End With

    MsgBox "->Footnote<- style updated to Noto Sans, 7pt.", vbInformation

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RedefineFootnoteStyle_NotoSans of Module basTEST_aeBibleFonts"
    Resume PROC_EXIT
End Sub

Public Sub AuditStyleUsage_FootnoteNormal()
    On Error GoTo PROC_ERR
    Dim para As Word.Paragraph
    Dim hitCount As Long
    Dim logBuffer As String

    logBuffer = "=== Audit: Paragraph Style Usage for 'Footnote normal' ===" & vbCrLf

    For Each para In ActiveDocument.Paragraphs
        If para.style = ActiveDocument.Styles("Footnote normal") Then
            hitCount = hitCount + 1
            logBuffer = logBuffer & "* Paragraph at Char " & para.Range.Start & " -> """ & _
            Replace(Left(para.Range.Text, 40), vbCr, "") & "...""" & vbCrLf
        End If
    Next para

    logBuffer = logBuffer & vbCrLf & "Total 'Footnote normal' style instances: " & hitCount
    Debug.Print logBuffer
    MsgBox "Audit complete. See Immediate Window for details.", vbInformation

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure AuditStyleUsage_FootnoteNormal of Module basTEST_aeBibleFonts"
    Resume PROC_EXIT
End Sub

Public Sub RedefineFootnoteNormalStyle_NotoSans()
    On Error GoTo PROC_ERR
    Dim s As style
    Set s = ActiveDocument.Styles("Footnote normal")

    With s.Font
        .Name = "Noto Sans"
        .Size = 7
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .color = wdColorAutomatic
    End With

    MsgBox "'Footnote normal' style updated to Noto Sans, 7pt.", vbInformation

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RedefineFootnoteNormalStyle_NotoSans of Module basTEST_aeBibleFonts"
    Resume PROC_EXIT
End Sub

Public Sub AuditStyleUsage_PictureCaption()
    On Error GoTo PROC_ERR
    Dim para As Word.Paragraph
    Dim hitCount As Long
    Dim logBuffer As String
    Dim s As style

    logBuffer = "=== Audit: Paragraph Style Usage for 'Picture Caption' ===" & vbCrLf

    On Error Resume Next
    Set s = ActiveDocument.Styles("Picture Caption")
    On Error GoTo 0
    On Error GoTo PROC_ERR

    If s Is Nothing Then
        MsgBox "Style 'Picture Caption' not found in this document.", vbExclamation
        GoTo PROC_EXIT
    End If

    For Each para In ActiveDocument.Paragraphs
        If para.style = s Then
            hitCount = hitCount + 1
            logBuffer = logBuffer & "* Paragraph at Char " & para.Range.Start & " -> """ & _
                Replace(Left(para.Range.Text, 40), vbCr, "") & "...""" & vbCrLf
        End If
    Next para

    logBuffer = logBuffer & vbCrLf & "Total 'Picture Caption' style instances: " & hitCount
    Debug.Print logBuffer
    MsgBox "Audit complete. See Immediate Window for details.", vbInformation

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure AuditStyleUsage_PictureCaption of Module basTEST_aeBibleFonts"
    Resume PROC_EXIT
End Sub

Public Sub RedefinePictureCaptionStyle_NotoSans()
    On Error GoTo PROC_ERR
    Dim s As style
    On Error Resume Next
    Set s = ActiveDocument.Styles("Picture Caption")
    On Error GoTo 0
    On Error GoTo PROC_ERR

    If s Is Nothing Then
        MsgBox "'Picture Caption' style not found in this document.", vbExclamation
        GoTo PROC_EXIT
    End If

    With s.Font
        .Name = "Noto Sans"
        .Size = 9
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .color = wdColorAutomatic
    End With

    Debug.Print "'Picture Caption' style updated to Noto Sans, 9pt."
    MsgBox "'Picture Caption' style redefined to Noto Sans, 9pt.", vbInformation

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RedefinePictureCaptionStyle_NotoSans of Module basTEST_aeBibleFonts"
    Resume PROC_EXIT
End Sub

Public Sub Identify_ArialUnicodeMS_Paragraphs()
    On Error GoTo PROC_ERR
    Dim para As Word.Paragraph
    Dim paraIndex As Long
    Dim secIndex As Long
    Dim hfIndex As Long
    Dim hfTypes As Variant
    Dim hfKind As Variant
    Dim logBuffer As String
    Dim sec As Word.Section
    Dim hf As HeaderFooter
    Dim fontName As String

    logBuffer = "=== Arial Unicode MS Paragraph Identification ===" & vbCrLf

    ' Scan body
    paraIndex = 0
    For Each para In ActiveDocument.Paragraphs
        paraIndex = paraIndex + 1
        fontName = para.Range.Characters(1).Font.Name
        If fontName = "Arial Unicode MS" Then
            logBuffer = logBuffer & "[Body] Para #" & paraIndex & " - Style: " & para.style & vbCrLf
            logBuffer = logBuffer & "Text: " & Left(para.Range.Text, 120) & vbCrLf & vbCrLf
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
                For Each para In hf.Range.Paragraphs
                    paraIndex = paraIndex + 1
                    fontName = para.Range.Characters(1).Font.Name
                    If fontName = "Arial Unicode MS" Then
                        logBuffer = logBuffer & "[Header] Sec " & secIndex & ", Type " & hfKind & ", Para #" & paraIndex & vbCrLf
                        logBuffer = logBuffer & "Style: " & para.style & vbCrLf
                        logBuffer = logBuffer & "Text: " & Left(para.Range.Text, 120) & vbCrLf & vbCrLf
                    End If
                Next
            End If
            Set hf = sec.Footers(hfKind)
            If hf.Exists Then
                paraIndex = 0
                For Each para In hf.Range.Paragraphs
                    paraIndex = paraIndex + 1
                    fontName = para.Range.Characters(1).Font.Name
                    If fontName = "Arial Unicode MS" Then
                        logBuffer = logBuffer & "[Footer] Sec " & secIndex & ", Type " & hfKind & ", Para #" & paraIndex & vbCrLf
                        logBuffer = logBuffer & "Style: " & para.style & vbCrLf
                        logBuffer = logBuffer & "Text: " & Left(para.Range.Text, 120) & vbCrLf & vbCrLf
                    End If
                Next
            End If
        Next hfKind
    Next sec

    Debug.Print logBuffer
    MsgBox "Arial Unicode MS detection complete. See Immediate Window.", vbInformation

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Identify_ArialUnicodeMS_Paragraphs of Module basTEST_aeBibleFonts"
    Resume PROC_EXIT
End Sub

Private Function StoryTypeName(StoryType As WdStoryType) As String
    Select Case StoryType
        Case wdMainTextStory: StoryTypeName = "Body"
        Case wdPrimaryHeaderStory: StoryTypeName = "Primary Header"
        Case wdFirstPageHeaderStory: StoryTypeName = "First Page Header"
        Case wdEvenPagesHeaderStory: StoryTypeName = "Even Pages Header"
        Case wdPrimaryFooterStory: StoryTypeName = "Primary Footer"
        Case wdFirstPageFooterStory: StoryTypeName = "First Page Footer"
        Case wdEvenPagesFooterStory: StoryTypeName = "Even Pages Footer"
        Case wdFootnotesStory: StoryTypeName = "Footnotes"
        Case wdEndnotesStory: StoryTypeName = "Endnotes"
        Case wdTextFrameStory: StoryTypeName = "Textboxes"
        Case wdCommentsStory: StoryTypeName = "Comments"
        Case 8: StoryTypeName = "TOC"   ' 8 is wdTOCStory in newer versions
        Case Else: StoryTypeName = "Other Story (" & StoryType & ")"
    End Select
End Function

Public Sub AuditFontUsage_ParagraphsAndHeadersFooters()
    On Error GoTo PROC_ERR
    Dim para As Word.Paragraph
    Dim fontMap As Object
    Dim fName As String
    Dim keyVar As Variant
    Dim logBuffer As String
    Dim sec As Word.Section
    Dim hf As HeaderFooter
    Dim hfTypes As Variant
    Dim hfKind As Variant

    Set fontMap = CreateObject("Scripting.Dictionary")

    ' Scan body paragraphs
    For Each para In ActiveDocument.Paragraphs
        fName = para.Range.Characters(1).Font.Name
        If Not fontMap.Exists(fName) Then
            fontMap.Add fName, 1
        Else
            fontMap(fName) = fontMap(fName) + 1
        End If
    Next para

    ' Define header/footer types
    hfTypes = Array(wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages)

    ' Scan header/footer paragraphs
    For Each sec In ActiveDocument.Sections
        For Each hfKind In hfTypes
            Set hf = sec.Headers(hfKind)
            If hf.Exists Then
                For Each para In hf.Range.Paragraphs
                    fName = para.Range.Characters(1).Font.Name
                    If Not fontMap.Exists(fName) Then
                        fontMap.Add fName, 1
                    Else
                        fontMap(fName) = fontMap(fName) + 1
                    End If
                Next para
            End If

            Set hf = sec.Footers(hfKind)
            If hf.Exists Then
                For Each para In hf.Range.Paragraphs
                    fName = para.Range.Characters(1).Font.Name
                    If Not fontMap.Exists(fName) Then
                        fontMap.Add fName, 1
                    Else
                        fontMap(fName) = fontMap(fName) + 1
                    End If
                Next para
            End If
        Next hfKind
    Next sec

    ' Output results
    logBuffer = "=== Font Usage Across Body, Headers, and Footers ===" & vbCrLf
    For Each keyVar In fontMap.Keys
        logBuffer = logBuffer & "- " & keyVar & ": " & fontMap(keyVar) & " paragraph(s)" & vbCrLf
    Next

    Debug.Print logBuffer
    MsgBox "Full font audit complete. See Immediate Window.", vbInformation

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure AuditFontUsage_ParagraphsAndHeadersFooters of Module Module1"
    Resume PROC_EXIT
End Sub

Public Sub FindParagraphsByFirstCharFont_BodyHeadersFooters(targetFont As String)
    Dim para As Word.Paragraph
    Dim sec As Word.Section
    Dim hf As HeaderFooter
    Dim hfTypes As Variant
    Dim hfKind As Variant
    Dim count As Long
    Dim fName As String

    Debug.Print "=== Paragraphs whose FIRST CHARACTER is font: " & targetFont & " ==="

    ' Body paragraphs
    For Each para In ActiveDocument.Paragraphs
        fName = para.Range.Characters(1).Font.Name
        If StrComp(fName, targetFont, vbTextCompare) = 0 Then
            count = count + 1
            Debug.Print count & ". [Body] Sec " & para.Range.Sections(1).index & _
                        " | """ & Left$(Trim$(para.Range.Text), 80) & """"
        End If
    Next para

    ' Header/footer types
    hfTypes = Array(wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages)

    ' Header/footer paragraphs
    For Each sec In ActiveDocument.Sections
        For Each hfKind In hfTypes

            Set hf = sec.Headers(hfKind)
            If hf.Exists Then
                For Each para In hf.Range.Paragraphs
                    fName = para.Range.Characters(1).Font.Name
                    If StrComp(fName, targetFont, vbTextCompare) = 0 Then
                        count = count + 1
                        Debug.Print count & ". [Header " & HeaderFooterLabel(hfKind) & _
                                    "] Sec " & sec.index & _
                                    " | """ & Left$(Trim$(para.Range.Text), 80) & """"
                    End If
                Next para
            End If

            Set hf = sec.Footers(hfKind)
            If hf.Exists Then
                For Each para In hf.Range.Paragraphs
                    fName = para.Range.Characters(1).Font.Name
                    If StrComp(fName, targetFont, vbTextCompare) = 0 Then
                        count = count + 1
                        Debug.Print count & ". [Footer " & HeaderFooterLabel(hfKind) & _
                                    "] Sec " & sec.index & _
                                    " | """ & Left$(Trim$(para.Range.Text), 80) & """"
                    End If
                Next para
            End If
        Next hfKind
    Next sec

    Debug.Print "=== Total paragraphs found: " & count & " ==="
End Sub

Private Function HeaderFooterLabel(ByVal kind As WdHeaderFooterIndex) As String
    Select Case kind
        Case wdHeaderFooterPrimary:   HeaderFooterLabel = "Primary"
        Case wdHeaderFooterFirstPage: HeaderFooterLabel = "FirstPage"
        Case wdHeaderFooterEvenPages: HeaderFooterLabel = "EvenPages"
        Case Else:                    HeaderFooterLabel = "Other"
    End Select
End Function

