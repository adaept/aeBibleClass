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

    MsgBox "Open Font Availability Report:" & vbCrLf & vbCrLf & _
           "Installed Fonts:" & vbCrLf & InstalledFonts & vbCrLf & _
           "Missing Fonts:" & vbCrLf & MissingFonts & _
           IIf(DownloadLinks <> "", vbCrLf & "Download Missing Fonts:" & vbCrLf & DownloadLinks, ""), _
           vbInformation, "Open Font Check"
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

Function CountFindNotEmphasisBlack()
    CountFindNotEmphasisBlack = FindNotEmphasisBlackRed("Automatic")
End Function

Function CountFindNotEmphasisRed()
    CountFindNotEmphasisRed = FindNotEmphasisBlackRed("Dark Red")
End Function

Function FindNotEmphasisBlackRed(fontColor As String) As Integer
' fontColor = Automatic - Arial Black, 8pt, "Normal" Style
' fontColor = Dark Red - Arial Black, 8pt, "Words of Jesus" Style
' This is a fast that will skip over text that is of the style EmphasisBlack or EmphasisRed
' When count is 0 there is no more text that needs setting for style EmphasisBlack or EmphasisRed

    Dim rng As range
    Set rng = ActiveDocument.content
    Dim totalCount As Integer
    Dim wordColor As Long
    Dim wordStyle As style
    Dim myStyle As style

    If fontColor = "Automatic" Then
        wordColor = wdColorAutomatic
        Set wordStyle = ActiveDocument.Styles("Normal")
        Set myStyle = ActiveDocument.Styles("EmphasisBlack")
    ElseIf fontColor = "Dark Red" Then
        wordColor = wdColorDarkRed
        Set wordStyle = ActiveDocument.Styles("Words of Jesus")
        Set myStyle = ActiveDocument.Styles("EmphasisRed")
    Else
        MsgBox "Incorrect parameter. Automatic or Dark Red expected!", vbCritical, "FindNotEmphasisBlackRed"
        Stop
    End If

    'Debug.Print "wordStyle = " & wordStyle, "wordColor = " & wordColor, "wdColorDarkRed = " & wdColorDarkRed, _
                    "myStyle = " & myStyle
    totalCount = 0
    With rng.Find
        .ClearFormatting
        .style = ActiveDocument.Styles(wordStyle)
        .font.name = "Arial Black"
        .font.Size = 8
        .font.color = wordColor
        .text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True

        'Debug.Print ".style = " & rng.Find.style
        Do While .Execute
            ' Check if character style is NOT EmphasisBlack or EmphasisRed
            If Not rng.Characters(1).style = myStyle Then
                rng.Select
                totalCount = totalCount + 1
                MsgBox "Found matching text (not " & myStyle & ").", vbInformation
                Exit Function
            End If
            rng.Collapse wdCollapseEnd
        Loop

        'If totalCount = 0 Then MsgBox "totalCount = 0, No matching text found (excluding EmphasisBlack).", vbExclamation
    End With
    'Debug.Print "totalCount = " & totalCount
End Function



