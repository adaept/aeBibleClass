Attribute VB_Name = "basBibleRibbon_OLD"
' DEPRECATED: This module is superseded by aeRibbonClass.cls.
' Do not add new code here. Public declarations below shadow aeRibbonClass names
' and are retained only to avoid breaking any legacy callers. Use aeRibbonClass instead.
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString
Public headingData(1 To 66, 0 To 1) As Variant
Private savedPos As Long
Private bookAbbr As String

' Use Windows API to change cursor
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" ( _
    ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
' Global ribbon object and button state flags
Public ribbonUI As IRibbonUI
Public ribbonIsReady As Boolean
Public BtnNextEnabled As Boolean
Dim bookmarkIndex As Long

'Private Function GetParaIndexSafe(rng As Word.Range) As Long
'' Search Isa 23:42 (intentional false verse number) scanned nearly 9,000 paragraphs in under a quarter second,
'' with full interruptibility and no layout lock
'    Dim r As Word.Range
'    Set r = ActiveDocument.Range(0, 0)
'
'    Dim idx As Long: idx = 1
'    Dim tickStart As Single: tickStart = Timer
'    Dim tickNow As Single
'
'    Do While r.Start < rng.Start
'        Set r = r.Next(Unit:=wdParagraph)
'        idx = idx + 1
'
'        If idx Mod 500 = 0 Then
'            tickNow = Timer
'            'Debug.Print "Step " & idx & " ? Range.Start=" & r.Start & " | Elapsed=" & Format(tickNow - tickStart, "0.00") & "s"
'            If tickNow - tickStart > 5 Then
'                Debug.Print "> Timeout: Paragraph scan exceeded 5 seconds. Breaking out."
'                GetParaIndexSafe = -2
'                Exit Function
'            End If
'        End If
'    Loop
'
'    If r.Start = rng.Start Then
'        'Debug.Print ">> Found match at paragraph #" & idx
'        GetParaIndexSafe = idx
'    Else
'        'Debug.Print ">>> No exact match. Closest index: " & idx
'        GetParaIndexSafe = -1
'    End If
'End Function

Private Function StyleTypeLabel(st As WdStyleType) As String
    Select Case st
        Case wdStyleTypeParagraph: StyleTypeLabel = "Paragraph"
        Case wdStyleTypeCharacter: StyleTypeLabel = "Character"
        Case wdStyleTypeTable:     StyleTypeLabel = "Table"
        Case wdStyleTypeList:      StyleTypeLabel = "List"
        Case Else:                 StyleTypeLabel = "Unknown"
    End Select
End Function

Private Function LeftUntilLastSpace(ByVal txt As String) As String
    Dim lastSpacePos As Long

    ' Find the first space from the right
    lastSpacePos = InStrRev(txt, " ")

    If lastSpacePos > 0 Then
        LeftUntilLastSpace = Left(txt, lastSpacePos - 1)
    Else
        LeftUntilLastSpace = txt  ' No space found, return full string
    End If
End Function

Private Function ExtractTrailingDigits(ByVal txt As String) As String
    Dim i As Long, ch As String, Result As String
    Result = ""

    ' Scan backwards, collecting up to 3 digits
    For i = Len(txt) To 1 Step -1
        ch = Mid$(txt, i, 1)
        If ch Like "#" Then
            Result = ch & Result
            If Len(Result) = 3 Then Exit For
        Else
            Exit For  ' Stop at first non-digit
        End If
    Next i

    ExtractTrailingDigits = Result
End Function

Private Function IsOneChapterBook(book As String) As Boolean
    Select Case book  ' Books of only one chapter
        Case "OBADIAH", "PHILEMON", "2 JOHN", "3 JOHN", "JUDE"
            IsOneChapterBook = True
        Case Else
            IsOneChapterBook = False
    End Select
End Function

Private Function SaveCursor() As Long
    SaveCursor = Selection.Start
End Function

Private Sub RestoreCursor(ByVal savedPos As Long)
    Selection.SetRange savedPos, savedPos
    Selection.Collapse Direction:=wdCollapseStart
End Sub

Private Sub FindBookH1(fullBookName As String, ByRef paraIndex As Long, _
                               Optional ByVal chapNum As String = "1", Optional ByVal verseNum As String = "1")
    On Error GoTo PROC_ERR
    Debug.Print "FindBookH1: >> chapNum = " & chapNum, "verseNum = " & verseNum
    savedPos = SaveCursor()

    Dim r As Word.Range
    Set r = ActiveDocument.Paragraphs(1).Range

    Dim paraText As String, bookFound As Boolean
    Dim paraCount As Long: paraCount = 1
    bookFound = False

    Do While Not r Is Nothing
        If r.Paragraphs(1).style = "Heading 1" Then
            paraText = UCase(Replace(Trim$(r.Text), vbCr, ""))
            If paraText = UCase(fullBookName) Then
                bookFound = True
                paraIndex = paraCount
                Debug.Print "FindBookH1: >> Book found", "'" & paraText & "'", "#" & paraIndex, "bookFound = " & bookFound

                ' Move cursor safely
                With ActiveDocument.Paragraphs(paraIndex).Range
                    .Select
                    Selection.Collapse Direction:=wdCollapseStart
                End With

                ' Call next routine
                FindChapterH2 fullBookName, paraIndex, chapNum, verseNum
                GoTo PROC_EXIT
            End If
        End If
        paraCount = paraCount + 1
        Set r = r.Next(Unit:=wdParagraph)
    Loop

    If Not bookFound Then RestoreCursor savedPos
    Debug.Print "FindBookH1: >> Book not found: '" & fullBookName & "'", "bookFound = " & bookFound
    MsgBox "Book not found: '" & fullBookName & "'", vbExclamation, "Bible"

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FindBookH1 of Module basBibleRibbon_OLD"
    Resume PROC_EXIT
End Sub

Private Sub FindChapterH2(fullBookName As String, ByRef paraIndex As Long, _
    Optional ByVal chapNum As String = "1", Optional ByVal verseNum As String = "1")
    On Error GoTo PROC_ERR
    Dim chapTag1 As String, chapTag2 As String
    Dim rng As Word.Range
    Dim paraText As String
    Dim count As Long

    chapTag1 = "Chapter " & chapNum
    chapTag2 = "PSALM " & chapNum

    Set rng = ActiveDocument.Paragraphs(paraIndex).Range
    count = 0

    Do While Not rng Is Nothing
        If rng.style = "Heading 2" Then
            paraText = Trim$(rng.Text)
            If InStr(1, paraText, chapTag1, vbTextCompare) > 0 Or _
                InStr(1, paraText, chapTag2, vbTextCompare) > 0 Then
                paraIndex = paraIndex + count
                With ActiveDocument.Paragraphs(paraIndex).Range
                    .Select
                    Selection.Collapse Direction:=wdCollapseStart
                End With
                Debug.Print "FindChapterH2: >>>", "Cursor moved to paraIndex = #" & paraIndex; ""
                GoTo PROC_EXIT
            End If
        End If
        count = count + 1
        Set rng = rng.Next(Unit:=wdParagraph, count:=1)
    Loop

    MsgBox "Chapter not found: '" & fullBookName & "' Chapter = " & chapNum, vbExclamation, "Bible"

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FindChapterH2 of Module basBibleRibbon_OLD"
    Resume PROC_EXIT
End Sub

Private Function ParseParts(ByVal userInput As String, Optional ByVal delimiter As String = ":") As String()
    On Error GoTo PROC_ERR
    Dim parts() As String
    Dim i As Long

    parts = Split(userInput, delimiter)

    Debug.Print "ParseParts: Input: """ & userInput & """"
    Debug.Print "ParseParts: Delimiter: """ & delimiter & """"
    Debug.Print "ParseParts: Parts found: " & UBound(parts) - LBound(parts) + 1

    For i = LBound(parts) To UBound(parts)
        Debug.Print "Part " & i & ": " & parts(i)
    Next i

    ParseParts = parts

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ParseParts of Module basBibleRibbon_OLD"
    Resume PROC_EXIT
End Function

Sub GoToSection()
    On Error GoTo PROC_ERR
    'NavigateToNextBookmark()
    Dim bmList As Collection
    Set bmList = GetBookmarkList()

    If bmList.count = 0 Then
        MsgBox "No bookmarks found.", vbExclamation
        GoTo PROC_EXIT
    End If

    bookmarkIndex = bookmarkIndex + 1
    If bookmarkIndex > bmList.count Then bookmarkIndex = 1

    bmList.Item(bookmarkIndex).Range.Select

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GoToSection of Module basBibleRibbon_OLD"
    Resume PROC_EXIT
End Sub

Private Function GetBookmarkList() As Collection
    On Error GoTo PROC_ERR
    Dim bmColl As New Collection
    Dim bm As Bookmark

    For Each bm In ActiveDocument.Bookmarks
        bmColl.Add bm
    Next bm

    Set GetBookmarkList = bmColl

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GetBookmarkList of Module basBibleRibbon_OLD"
    Resume PROC_EXIT
End Function

Private Sub GoToH1()
    On Error GoTo PROC_ERR
    Dim pattern As String
    Dim para As Word.Paragraph
    Dim paraText As String
    Dim matchFound As Boolean

    pattern = InputBox("Enter a Book Name (Heading 1) abbreviation:", "Go To Bible Book")
    If pattern = "" Then GoTo PROC_EXIT ' User canceled
    matchFound = False

    ' Disable UI updates for speed
    Application.ScreenUpdating = False

    For Each para In ActiveDocument.Paragraphs
        If para.style = "Heading 1" Then
            paraText = Trim$(para.Range.Text)
            If paraText Like "*" & UCase(pattern) & "*" Then
                para.Range.Select
                ' Move insertion point (cursor) without selecting text
                ActiveDocument.Range(para.Range.Start, para.Range.Start).Select
                matchFound = True
                Exit For
            End If
        End If
    Next para

    Application.ScreenUpdating = True
    Selection.Range.Select  ' Re-select current range to restore cursor
    DoEvents  ' Allows UI refresh

    If Not matchFound Then
        MsgBox "Book not found! No Heading 1 matches pattern: '" & pattern & "'", vbExclamation, "Bible"
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    Application.ScreenUpdating = True
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GoToH1 of Module basBibleRibbon_OLD"
    Resume PROC_EXIT
End Sub

Private Sub NextButton()
    On Error GoTo PROC_ERR
    Dim doc As Document
    Dim searchRange As Word.Range
    Dim paraEnd As Long
    Dim found As Boolean

    Set doc = ActiveDocument
    found = False

    ' Move start past current paragraph to avoid re-matching
    paraEnd = Selection.Paragraphs(1).Range.End
    Set searchRange = doc.Range(paraEnd, doc.content.End)

    With searchRange.Find
        .ClearFormatting
        .style = doc.Styles("Heading 1")
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Text = ""
        found = .Execute
    End With

    ' If not found, wrap: from beginning to current paragraph start
    If Not found Then
        Set searchRange = doc.Range(0, paraEnd)
        With searchRange.Find
            .ClearFormatting
            .style = doc.Styles("Heading 1")
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .Text = ""
            found = .Execute
        End With
    End If

    ' If found, move cursor to start of heading
    If found Then
        Selection.SetRange searchRange.Start, searchRange.Start
        ActiveWindow.ScrollIntoView Selection.Range, True
    Else
        MsgBox "No Heading 1 found in the document.", vbInformation
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure NextButton of Module basBibleRibbon_OLD"
    Resume PROC_EXIT
End Sub

Private Function GetExactVerticalScroll() As Double
' Return the scroll percentage rounded to three decimal places
    On Error GoTo PROC_ERR
    Dim visibleStart As Long
    Dim totalLength As Long
    Dim scrollPercentage As Double

    ' Get the starting position of the visible content
    visibleStart = ActiveWindow.Selection.Start

    ' Get the total document length
    totalLength = ActiveDocument.content.End

    ' Calculate the exact scroll percentage
    If totalLength > 0 Then
        scrollPercentage = (visibleStart / totalLength) * 100
    Else
        scrollPercentage = 0
    End If

    ' Round to 3 decimal places
    GetExactVerticalScroll = Round(scrollPercentage, 3)

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GetExactVerticalScroll of Module basBibleRibbon_OLD"
    Resume PROC_EXIT
End Function

