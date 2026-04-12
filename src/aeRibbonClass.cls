VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aeRibbonClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'Copyright (c) 2025-2026 Peter F. Ennis
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation;
'version 3.0, or (at your option) any later version.
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'Lesser General Public License for more details.
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, visit
'https://www.gnu.org/licenses/lgpl-3.0.txt

' ================================================================
' Author:   Peter F. Ennis
' Date:     March 10, 2025
' Comment:  Bible Ribbon Class for Word 365 VBA
' History:  See comment details, basChangeLog_aeBibleClass, commit messages on GitHub
' GitHub:   https://github.com/adaept/aeBibleClass
' ================================================================

' -- Private State -------------------------------------------------------------
Private m_ribbon         As IRibbonUI
Private m_ribbonReady    As Boolean
Private m_btnNextEnabled As Boolean
Private m_btnPrevEnabled As Boolean
Private m_bookmarkIndex  As Long
Private m_savedPos       As Long
Private m_bookAbbr       As String
Private headingData(1 To 66, 0 To 1) As Variant
Private m_currentBookIndex As Long   ' 1-based index into headingData; 0 = not set
Private m_currentBookPos   As Long   ' character position of book H1 in document
Private m_currentChapter   As Long   ' current chapter number; 0 = not set
Private m_currentVerse     As Long   ' current verse number; 0 = not set
Private m_pendingChapter   As Long   ' chapter number staged for deferred GoToChapter call

' Windows API - cursor control
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" ( _
    ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

' -- Lifecycle -----------------------------------------------------------------

Private Sub Class_Initialize()
    Debug.Print "RibbonController: Class_Initialize at " & Format(Now, "hh:nn:ss")
    m_btnNextEnabled = False
    m_btnPrevEnabled = False
    m_ribbonReady = False
    m_bookmarkIndex = 0
    m_savedPos = 0
    m_bookAbbr = vbNullString
    m_currentBookIndex = 0
    m_currentBookPos = 0
    m_currentChapter = 0
    m_currentVerse = 0
    m_pendingChapter = 0
End Sub

Private Sub Class_Terminate()
    Debug.Print "RibbonController: Class_Terminate at " & Format(Now, "hh:nn:ss")
    Set m_ribbon = Nothing
    m_ribbonReady = False
End Sub

' -- Properties ----------------------------------------------------------------

Public Property Get RibbonReady() As Boolean
    RibbonReady = m_ribbonReady
End Property

Public Property Get BtnNextEnabled() As Boolean
    BtnNextEnabled = m_btnNextEnabled
End Property

Public Property Let BtnNextEnabled(val As Boolean)
    m_btnNextEnabled = val
End Property

Public Property Get BtnPrevEnabled() As Boolean
    BtnPrevEnabled = m_btnPrevEnabled
End Property

Public Property Let BtnPrevEnabled(val As Boolean)
    m_btnPrevEnabled = val
End Property

' -- Ribbon Load ---------------------------------------------------------------

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    On Error GoTo PROC_ERR
    Set m_ribbon = ribbon
    m_ribbonReady = True
    Debug.Print "RibbonController: Ribbon ready at " & Format(Now, "hh:nn:ss")
    m_ribbon.InvalidateControl "NextBookButton"
    m_ribbon.InvalidateControl "PrevBookButton"
    Call EnableButtonsRoutine
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OnRibbonLoad of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Invalidation --------------------------------------------------------------

Public Sub InvalidateAll()
    On Error GoTo PROC_ERR
    If Not m_ribbon Is Nothing Then m_ribbon.Invalidate
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure InvalidateAll of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

Public Sub InvalidateControl(controlId As String)
    On Error GoTo PROC_ERR
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl controlId
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure InvalidateControl of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Ribbon Callbacks ----------------------------------------------------------

Public Sub OnGoToVerseSblClick(control As IRibbonControl)
    Call GoToVerseSBL
End Sub

Public Sub OnPrevButtonClick(control As IRibbonControl)
    Call PrevButton
End Sub

Public Sub OnGoToH1ButtonClick(control As IRibbonControl)
    Call GoToH1
End Sub

' Test wrapper — calls GoToH1 outside the ribbon callback mechanism (Alt+F8)
Public Sub GoToH1Direct()
    Call GoToH1
End Sub

Public Sub OnNextButtonClick(control As IRibbonControl)
    Call NextButton
End Sub

Public Sub OnAdaeptAboutClick(control As IRibbonControl)
    MsgBox "Hello, adaept World!" & vbCrLf & _
           "adaeptMsg  = " & AdaeptMsg, vbInformation, "About adaept"
End Sub

' -- Enabled State Callback ----------------------------------------------------

Public Function GetNextEnabled(control As IRibbonControl) As Boolean
    GetNextEnabled = m_btnNextEnabled
End Function

Public Function GetPrevEnabled(control As IRibbonControl) As Boolean
    GetPrevEnabled = m_btnPrevEnabled
End Function

' -- Initialization ------------------------------------------------------------

Private Sub EnableButtonsRoutine()
    On Error GoTo PROC_ERR
    Debug.Print "RibbonController: EnableButtonsRoutine"
    CaptureHeading1s
    LogHeadingData
    ' WarmLayoutCache disabled: causes ~50s freeze at document open and brings
    ' File Explorer to the foreground. The first GoTo Book will warm the layout
    ' on demand (~12s, once per session). Re-enable here if the pre-warm strategy
    ' is revisited.
    ' Application.OnTime Now + TimeValue("00:00:05"), _
    '     ActiveDocument.VBProject.Name & ".basRibbonDeferred.WarmLayoutCacheDeferred"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure EnableButtonsRoutine of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- GoTo Verse (stub - parser to be implemented) ------------------------------

Private Sub GoToVerseSBL()
    MsgBox "GoToVerseSBL - Parser not yet implemented."
End Sub

' -- GoTo Book (H1) ------------------------------------------------------------

Private Sub GoToH1()
    On Error GoTo PROC_ERR
    Dim pattern  As String
    Dim i        As Long
    Dim k        As Long
    Dim foundPos As Long
    Dim matchFound As Boolean

    ' Re-capture headings if array is empty (e.g. after a VBA IDE reset clears
    ' module state and the singleton is recreated without OnRibbonLoad firing)
    If IsEmpty(headingData(1, 0)) Then CaptureHeading1s

    pattern = InputBox("Enter a Book Name (Heading 1) abbreviation:", "Go To Bible Book")
    If pattern = "" Then GoTo PROC_EXIT

    matchFound = False
    For i = 1 To 66
        If Not IsEmpty(headingData(i, 0)) Then
            If CStr(headingData(i, 0)) Like "*" & UCase(pattern) & "*" Then
                matchFound = True
                Exit For
            End If
        End If
    Next i

    If Not matchFound Then
        MsgBox "Book not found! No Heading 1 matches: '" & pattern & "'", vbExclamation, "Bible"
        GoTo PROC_EXIT
    End If

    foundPos = CLng(headingData(i, 1))
    ActiveDocument.Range(foundPos, foundPos).Select
    m_btnPrevEnabled = True
    m_btnNextEnabled = True
    If Not IsEmpty(headingData(1, 1)) Then
        If foundPos = CLng(headingData(1, 1)) Then m_btnPrevEnabled = False
    End If
    For k = 66 To 1 Step -1
        If Not IsEmpty(headingData(k, 1)) And CLng(headingData(k, 1)) > 0 Then
            If foundPos = CLng(headingData(k, 1)) Then m_btnNextEnabled = False
            Exit For
        End If
    Next k

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GoToH1 of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Next Book (H1) ------------------------------------------------------------

Private Sub NextButton()
    On Error GoTo PROC_ERR
    Dim found    As Boolean
    Dim foundPos As Long
    Dim k        As Long
    Dim curParaEnd As Long
    curParaEnd = Selection.Paragraphs(1).Range.End
    Application.ScreenUpdating = False
    Selection.SetRange curParaEnd, curParaEnd
    Selection.Find.ClearFormatting
    Selection.Find.Text = ""
    Selection.Find.style = ActiveDocument.Styles("Heading 1")
    Selection.Find.Forward = True
    Selection.Find.Wrap = wdFindStop
    found = Selection.Find.Execute
    If found Then
        Selection.Collapse Direction:=wdCollapseStart
        foundPos = Selection.Range.Start
        m_btnPrevEnabled = True
        m_btnNextEnabled = True
        For k = 66 To 1 Step -1
            If Not IsEmpty(headingData(k, 1)) And CLng(headingData(k, 1)) > 0 Then
                If foundPos = CLng(headingData(k, 1)) Then m_btnNextEnabled = False
                Exit For
            End If
        Next k
    Else
        m_btnNextEnabled = False
    End If
    Application.ScreenUpdating = True
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "NextBookButton"
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "PrevBookButton"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    Application.ScreenUpdating = True
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure NextButton of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Prev Book (H1) ------------------------------------------------------------

Private Sub PrevButton()
    On Error GoTo PROC_ERR
    Dim found    As Boolean
    Dim foundPos As Long
    Application.ScreenUpdating = False
    Selection.Find.ClearFormatting
    Selection.Find.Text = ""
    Selection.Find.style = ActiveDocument.Styles("Heading 1")
    Selection.Find.Forward = False
    Selection.Find.Wrap = wdFindStop
    found = Selection.Find.Execute
    If found Then
        Selection.Collapse Direction:=wdCollapseStart
        foundPos = Selection.Range.Start
        m_btnNextEnabled = True
        m_btnPrevEnabled = True
        If Not IsEmpty(headingData(1, 1)) Then
            If foundPos = CLng(headingData(1, 1)) Then m_btnPrevEnabled = False
        End If
    Else
        m_btnPrevEnabled = False
    End If
    Application.ScreenUpdating = True
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "NextBookButton"
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "PrevBookButton"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    Application.ScreenUpdating = True
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure PrevButton of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Heading Data --------------------------------------------------------------

'===========================================================================================
' Purpose: Capture Heading 1 text and paragraph index into a static array (1 to 66, 0 to 1)
'   - Only runs once per session
'   - Heading text excludes vbCrLf
'   - Optimized for speed and reproducibility
'   - Audit-traceable: logs execution status to Immediate Window
'===========================================================================================
Public Sub CaptureHeading1s()
    On Error GoTo PROC_ERR
    Static hasRun As Boolean
    Dim para As Word.Paragraph
    Dim i        As Long
    Dim paraText As String

    If hasRun Then
        Debug.Print "CaptureHeading1s: Already executed this session."
        GoTo PROC_EXIT
    End If

    i = 1
    For Each para In ActiveDocument.Paragraphs
        If para.style = "Heading 1" Then
            If i > 66 Then Exit For
            paraText = Trim(Replace(para.Range.Text, vbCr, ""))
            headingData(i, 0) = paraText
            headingData(i, 1) = para.Range.Start
            i = i + 1
        End If
    Next para

    hasRun = True
    Debug.Print "CaptureHeading1s: Stored " & i - 1 & " Heading 1 entries."

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure CaptureHeading1s of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

'=============================================================================================
' Maintain a persistent 66-row CSV file + header, with session-aware ParaIndex updates
' sessionID,Heading1,BookName,ParaIndex
'   - Only updates ParaIndex and timestamp if changed
'   - Entire file is rewritten each time
'   - Prior values are preserved for unchanged rows
'   - Git-friendly: diffs show timestamped changes only
'=============================================================================================
Public Sub LogHeadingData()
    On Error GoTo PROC_ERR
    Const ROWS        As Long = 66
    Dim fPath         As String
    Dim sessionStamp  As String
    Dim i             As Long
    Dim fNum          As Integer
    Dim csvLines(1 To ROWS) As String
    Static lastParaStarts(1 To ROWS) As Long
    Static lastBookNames(1 To ROWS)  As String
    Static lastTimestamps(1 To ROWS) As String
    Dim changed As Boolean

    fPath = ActiveDocument.Path & "\rpt\HeadingLog.txt"
    sessionStamp = Format(Now, "yyyy-mm-dd hh:nn:ss")

    For i = 1 To ROWS
        If Not IsEmpty(headingData(i, 0)) Then
            If headingData(i, 1) <> lastParaStarts(i) Then
                lastParaStarts(i) = headingData(i, 1)
                lastBookNames(i) = headingData(i, 0)
                lastTimestamps(i) = sessionStamp
                changed = True
            End If
        End If
        csvLines(i) = lastTimestamps(i) & ",H1[" & i & "]," & lastBookNames(i) & "," & lastParaStarts(i)
    Next i

    fNum = FreeFile
    Open fPath For Output As #fNum
    Print #fNum, "sessionID,Heading1,BookName,ParaIndex"
    For i = 1 To ROWS
        Print #fNum, csvLines(i)
    Next i
    Close #fNum

    If changed Then
        Debug.Print "LogHeadingData: File updated at " & sessionStamp
    Else
        Debug.Print "LogHeadingData: No changes - file preserved"
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure LogHeadingData of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Cursor Helpers ------------------------------------------------------------

Private Function SaveCursor() As Long
    SaveCursor = Selection.Start
End Function

Private Sub RestoreCursor(ByVal pos As Long)
    Selection.SetRange pos, pos
    Selection.Collapse Direction:=wdCollapseStart
End Sub

' -- About Message -------------------------------------------------------------

Public Sub WarmLayoutCache()
    On Error GoTo PROC_ERR
    Dim k        As Long
    Dim lastPos  As Long
    Dim savedPos As Long
    For k = 66 To 1 Step -1
        If Not IsEmpty(headingData(k, 1)) And CLng(headingData(k, 1)) > 0 Then
            lastPos = CLng(headingData(k, 1))
            Exit For
        End If
    Next k
    If lastPos = 0 Then
        Debug.Print "WarmLayoutCache: no heading data, skipping"
        GoTo PROC_EXIT
    End If
    savedPos = Selection.Start
    Debug.Print "WarmLayoutCache: warming to pos=" & lastPos & " at " & Format(Now, "hh:nn:ss")
    Application.StatusBar = "Bible: building navigation index..."
    ActiveDocument.Range(lastPos, lastPos).Select
    ActiveDocument.Range(savedPos, savedPos).Select
    Application.StatusBar = False
    Debug.Print "WarmLayoutCache: complete at " & Format(Now, "hh:nn:ss")
PROC_EXIT:
    Exit Sub
PROC_ERR:
    Application.StatusBar = False
    Debug.Print "WarmLayoutCache: error " & Err.Number & " " & Err.Description
    Resume PROC_EXIT
End Sub

' -- About Message -------------------------------------------------------------

Private Function AdaeptMsg() As String
    AdaeptMsg = """...the truth shall make you free.""" & " John 8:32 (KJV)"
End Function

' -- Book comboBox (Step 1) ----------------------------------------------------

Private Function NormalizeBookInput(ByVal raw As String) As String
    Dim s As String
    s = Trim(UCase(raw))
    If Len(s) >= 2 Then
        If s Like "[0-9][A-Z]*" Then s = Left$(s, 1) & " " & Mid$(s, 2)
    End If
    NormalizeBookInput = s
End Function

Public Function GetBookEnabled(control As IRibbonControl) As Boolean
    GetBookEnabled = True
End Function

Public Function GetBookCount(control As IRibbonControl) As Long
    GetBookCount = 0   ' Step 1: delegate to aeBibleCitationClass
End Function

Public Function GetBookItemLabel(control As IRibbonControl, index As Long) As String
    GetBookItemLabel = ""   ' Step 1
End Function

Public Function GetBookItemID(control As IRibbonControl, index As Long) As String
    GetBookItemID = CStr(index)   ' Step 1: return "SEP" for OT_NT_SEPARATOR_INDEX
End Function

Public Function GetBookText(control As IRibbonControl) As String
    If m_currentBookIndex > 0 Then
        If Not IsEmpty(headingData(m_currentBookIndex, 0)) Then
            GetBookText = CStr(headingData(m_currentBookIndex, 0))
            Exit Function
        End If
    End If
    GetBookText = ""
End Function

Public Sub OnBookChanged(control As IRibbonControl, text As String)
    On Error GoTo PROC_ERR

    Dim pattern As String
    pattern = NormalizeBookInput(text)
    If Len(pattern) < 2 Then GoTo PROC_EXIT   ' ignore single-character input

    If IsEmpty(headingData(1, 0)) Then CaptureHeading1s

    Dim i As Long
    Dim matchFound As Boolean
    For i = 1 To 66
        If Not IsEmpty(headingData(i, 0)) Then
            If CStr(headingData(i, 0)) Like "*" & pattern & "*" Then
                matchFound = True
                Exit For
            End If
        End If
    Next i

    If Not matchFound Then GoTo PROC_EXIT

    m_currentBookIndex = i
    m_currentBookPos = CLng(headingData(i, 1))
    m_currentChapter = 0
    m_currentVerse = 0

    m_btnPrevEnabled = (m_currentBookIndex > 1)
    Dim k As Long
    m_btnNextEnabled = False
    For k = 66 To 1 Step -1
        If Not IsEmpty(headingData(k, 1)) And CLng(headingData(k, 1)) > 0 Then
            m_btnNextEnabled = (m_currentBookIndex < k)
            Exit For
        End If
    Next k

    ' Bug 1 fix: full invalidate clears user-typed comboBox text in Chapter/Verse rows
    If Not m_ribbon Is Nothing Then m_ribbon.Invalidate

    ' Bug 2b fix: defer document selection so the key event is not in flight
    Dim projName As String
    projName = Application.ActiveDocument.VBProject.Name
    Application.OnTime Now, projName & ".basRibbonDeferred.GoToBookDeferred"

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OnBookChanged of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' Called from basRibbonDeferred.GoToBookDeferred — executes after key event clears
Public Sub NavigateToCurrentBook()
    If m_currentBookPos > 0 Then
        ActiveDocument.Range(m_currentBookPos, m_currentBookPos).Select
    End If
End Sub

' -- Chapter row (Step 3) ------------------------------------------------------

Private Function FindChapterPos(ByVal chNum As Long) As Long
    ' Returns character position of the Nth Heading 2 after the current book H1.
    ' Returns 0 if not found.
    On Error GoTo PROC_ERR
    Dim r As Word.Range
    Set r = ActiveDocument.Range(m_currentBookPos, ActiveDocument.content.End)
    Dim j As Long
    For j = 1 To chNum
        r.Find.ClearFormatting
        r.Find.style = ActiveDocument.Styles("Heading 2")
        r.Find.Text = ""
        r.Find.Forward = True
        r.Find.Wrap = wdFindStop
        If Not r.Find.Execute Then
            FindChapterPos = 0
            Exit Function
        End If
        If j < chNum Then
            r.Start = r.End
            r.End = ActiveDocument.content.End
        End If
    Next j
    FindChapterPos = r.Start
    Exit Function
PROC_ERR:
    FindChapterPos = 0
End Function

Private Sub GoToChapter(ByVal chNum As Long)
    On Error GoTo PROC_ERR
    If m_currentBookIndex = 0 Then GoTo PROC_EXIT   ' assert: select book first

    Dim bookName As String
    bookName = CStr(headingData(m_currentBookIndex, 0))

    Dim maxCh As Long
    maxCh = aeBibleCitationClass.ChaptersInBook(bookName)
    If chNum < 1 Or chNum > maxCh Then GoTo PROC_EXIT

    Dim chPos As Long
    chPos = FindChapterPos(chNum)
    If chPos = 0 Then GoTo PROC_EXIT

    ActiveDocument.Range(chPos, chPos).Select
    m_currentChapter = chNum
    m_currentVerse = 0

    If Not m_ribbon Is Nothing Then
        m_ribbon.InvalidateControl "cmbChapter"
        m_ribbon.InvalidateControl "PrevChapterButton"
        m_ribbon.InvalidateControl "NextChapterButton"
        m_ribbon.InvalidateControl "cmbVerse"
        m_ribbon.InvalidateControl "PrevVerseButton"
        m_ribbon.InvalidateControl "NextVerseButton"
    End If

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GoToChapter of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

Public Function GetChapterEnabled(control As IRibbonControl) As Boolean
    GetChapterEnabled = (m_currentBookIndex > 0)
End Function

Public Function GetPrevChapterEnabled(control As IRibbonControl) As Boolean
    GetPrevChapterEnabled = (m_currentChapter > 1)
End Function

Public Function GetNextChapterEnabled(control As IRibbonControl) As Boolean
    If m_currentChapter = 0 Or m_currentBookIndex = 0 Then
        GetNextChapterEnabled = False
        Exit Function
    End If
    Dim bookName As String
    bookName = CStr(headingData(m_currentBookIndex, 0))
    GetNextChapterEnabled = (m_currentChapter < aeBibleCitationClass.ChaptersInBook(bookName))
End Function

Public Function GetChapterText(control As IRibbonControl) As String
    If m_currentChapter > 0 Then
        GetChapterText = CStr(m_currentChapter)
    Else
        GetChapterText = ""
    End If
End Function

Public Sub OnChapterChanged(control As IRibbonControl, text As String)
    ' editBox has no onAction — onChange fires on Enter and on each keystroke.
    ' Navigation is guarded by numeric + range validation; only a valid chapter
    ' number triggers the deferred GoToChapter call.
    On Error GoTo PROC_ERR
    If Not IsNumeric(Trim(text)) Then GoTo PROC_EXIT
    If m_currentBookIndex = 0 Then GoTo PROC_EXIT
    Dim chNum As Long
    chNum = CLng(Trim(text))
    Dim bookName As String
    bookName = CStr(headingData(m_currentBookIndex, 0))
    If chNum < 1 Or chNum > aeBibleCitationClass.ChaptersInBook(bookName) Then GoTo PROC_EXIT
    m_pendingChapter = chNum
    Dim projName As String
    projName = Application.ActiveDocument.VBProject.Name
    Application.OnTime Now, projName & ".basRibbonDeferred.GoToChapterDeferred"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OnChapterChanged of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

Public Sub OnChapterAction(control As IRibbonControl, text As String)
    ' editBox does not support onAction in the customUI14 schema — this stub
    ' is retained in case the control type is changed back to comboBox.
End Sub

' Called from basRibbonDeferred.GoToChapterDeferred — executes after key event clears
Public Sub ExecutePendingChapter()
    If m_pendingChapter > 0 Then
        GoToChapter m_pendingChapter
        m_pendingChapter = 0
    End If
End Sub

Public Sub OnPrevChapterClick(control As IRibbonControl)
    If m_currentChapter > 1 Then GoToChapter m_currentChapter - 1
End Sub

Public Sub OnNextChapterClick(control As IRibbonControl)
    On Error GoTo PROC_ERR
    If m_currentChapter = 0 Or m_currentBookIndex = 0 Then GoTo PROC_EXIT
    Dim bookName As String
    bookName = CStr(headingData(m_currentBookIndex, 0))
    If m_currentChapter < aeBibleCitationClass.ChaptersInBook(bookName) Then
        GoToChapter m_currentChapter + 1
    End If
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OnNextChapterClick of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Verse row (skeleton — Step 5 wires navigation logic) ----------------------

Public Function GetPrevVerseEnabled(control As IRibbonControl) As Boolean
    GetPrevVerseEnabled = False
End Function

Public Function GetNextVerseEnabled(control As IRibbonControl) As Boolean
    GetNextVerseEnabled = False
End Function

Public Function GetVerseEnabled(control As IRibbonControl) As Boolean
    GetVerseEnabled = False
End Function

Public Function GetVerseText(control As IRibbonControl) As String
    GetVerseText = ""   ' Step 5: return CStr(m_currentVerse) when set
End Function

Public Sub OnVerseChanged(control As IRibbonControl, text As String)
    ' onChange fires on every keystroke — no navigation here.
    ' Navigation is triggered by Enter (OnVerseAction).
End Sub

Public Sub OnVerseAction(control As IRibbonControl, text As String)
    ' Step 5 — validate range, navigate by paragraph count or verse marker scan
End Sub

Public Sub OnPrevVerseClick(control As IRibbonControl)
    ' Step 5
End Sub

Public Sub OnNextVerseClick(control As IRibbonControl)
    ' Step 5
End Sub

' -- New Search ----------------------------------------------------------------

Public Function GetNewSearchEnabled(control As IRibbonControl) As Boolean
    GetNewSearchEnabled = (m_currentBookIndex > 0)
End Function

Public Sub OnNewSearchClick(control As IRibbonControl)
    m_currentBookIndex = 0
    m_currentBookPos = 0
    m_currentChapter = 0
    m_currentVerse = 0
    If Not m_ribbon Is Nothing Then m_ribbon.Invalidate
End Sub

