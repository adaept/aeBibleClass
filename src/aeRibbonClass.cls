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
Private m_bookmarkIndex  As Long
Private m_savedPos       As Long
Private m_bookAbbr       As String
Private headingData(1 To 66, 0 To 1) As Variant
Private chapterData(1 To 66, 1 To 150) As Long   ' H2 char positions; populated by CaptureHeading1s
Private m_currentBookIndex As Long    ' 1-based index into headingData; 0 = not set
Private m_currentBookPos   As Long    ' character position of book H1 in document
Private m_currentChapter    As Long   ' current chapter number; 0 = not set
Private m_currentVerse      As Long   ' current verse number; 0 = not set
Private m_currentChapterPos As Long   ' cached char position of current chapter H2; 0 = not cached
Private m_pendingChapter   As Long    ' chapter number staged for deferred GoToChapter call
Private m_pendingVerse     As Long    ' verse number staged for deferred GoToVerse call
Private m_bookTextValid    As Boolean ' True when last OnBookChanged resolved an alias; False on no-match
Private m_chapterTextValid As Boolean ' True when last OnChapterChanged accepted the value; False on rejection
Private m_chapterMax       As Long    ' max chapter Count for current book at last rejection; used in OnGoClick error message
Private m_verseTextValid   As Boolean ' True when last OnVerseChanged accepted the value; False on rejection
Private m_verseMax         As Long    ' max verse Count for current chapter at last rejection; used in OnGoClick error message
Private m_studyVersionSet  As Boolean ' True after IsStudyVersion has been evaluated
Private m_studyVersionVal  As Boolean ' cached Result of IsStudyVersion

' Windows API - cursor control
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" ( _
    ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

' -- Lifecycle -----------------------------------------------------------------

Private Sub Class_Initialize()
    Debug.Print "RibbonController: Class_Initialize at " & Format(Now, "hh:nn:ss")
    m_ribbonReady = False
    m_bookmarkIndex = 0
    m_savedPos = 0
    m_bookAbbr = vbNullString
    m_currentBookIndex = 0
    m_currentBookPos = 0
    m_currentChapter = 0
    m_currentVerse = 0
    m_currentChapterPos = 0
    m_pendingChapter = 0
    m_pendingVerse = 0
    m_bookTextValid = True
    m_chapterTextValid = True
    m_chapterMax = 0
    m_verseTextValid = True
    m_verseMax = 0
    m_studyVersionSet = False
    m_studyVersionVal = False
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

' -- Ribbon Load ---------------------------------------------------------------

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    On Error GoTo PROC_ERR
    Set m_ribbon = ribbon
    m_ribbonReady = True
    Debug.Print "RibbonController: Ribbon ready at " & Format(Now, "hh:nn:ss")
    Call EnableButtonsRoutine
    ' Full Invalidate after EnableButtonsRoutine so the initial cache reflects the
    ' post-load field values. Without this, the cache is from the very first ribbon
    ' render (before OnRibbonLoad fired), and deferred Invalidate calls from onChange
    ' callbacks may not update it in time for Tab routing on first load.
    m_ribbon.Invalidate
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

' Called deferred (Application.OnTime) after invalid chapter input — resets display
' to current valid value without disrupting Tab routing (safe from OnTime context).
Public Sub ResetChapterDisplay()
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "cmbChapter"
End Sub

' Called deferred (Application.OnTime) after invalid verse input — same pattern.
Public Sub ResetVerseDisplay()
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "cmbVerse"
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

Public Function GetNextBkEnabled(control As IRibbonControl) As Boolean
    GetNextBkEnabled = True   ' always-enable (#599); OnNextBkClick guards m_currentBookIndex = 0 or >= 66
End Function

Public Function GetPrevBkEnabled(control As IRibbonControl) As Boolean
    GetPrevBkEnabled = True   ' always-enable (#599); OnPrevBkClick guards m_currentBookIndex <= 1
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

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GoToH1 of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Next Book (H1) ------------------------------------------------------------

Private Sub NextButton()
    ' Navigate to the next book using ribbon state (m_currentBookIndex), not Selection.
    ' Using Selection caused Bug 19: ScrollIntoView does not move the document cursor,
    ' so after OnBookChanged the cursor lagged behind the ribbon state. Next/Prev were
    ' navigating from the stale cursor position instead of from the current book.
    On Error GoTo PROC_ERR
    If m_currentBookIndex <= 0 Then GoTo PROC_EXIT
    If m_currentBookIndex >= 66 Then
        Application.StatusBar = SB_ALREADY_LAST_BOOK
        GoTo PROC_EXIT
    End If
    Dim nextIdx As Long
    nextIdx = m_currentBookIndex + 1
    If IsEmpty(headingData(nextIdx, 1)) Then GoTo PROC_EXIT
    Dim nextPos As Long
    nextPos = CLng(headingData(nextIdx, 1))
    m_currentBookIndex = nextIdx
    m_currentBookPos = nextPos
    m_currentChapter = 1   ' Rule 2a
    m_currentVerse = 1     ' Rule 2a
    m_currentChapterPos = 0
    ' Range.Select moves the document cursor to the new book heading (not ScrollIntoView).
    ' Called from onAction — focus returns to document after button activation regardless.
    ' Cursor must follow viewport so Enter-on-button does not snap back to the previous
    ' verse position (Bug 22b).
    ActiveDocument.Range(nextPos, nextPos).Select
    If Not m_ribbon Is Nothing Then m_ribbon.Invalidate
    Application.OnTime Now, Application.ActiveDocument.VBProject.Name & ".basRibbonDeferred.UpdateStatusBarDeferred"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure NextButton of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Prev Book (H1) ------------------------------------------------------------

Private Sub PrevButton()
    ' Navigate to the previous book using ribbon state (m_currentBookIndex), not Selection.
    ' See NextButton comment — same root cause (Bug 19).
    On Error GoTo PROC_ERR
    If m_currentBookIndex <= 0 Then GoTo PROC_EXIT
    If m_currentBookIndex = 1 Then
        Application.StatusBar = SB_ALREADY_FIRST_BOOK
        GoTo PROC_EXIT
    End If
    Dim prevIdx As Long
    prevIdx = m_currentBookIndex - 1
    If IsEmpty(headingData(prevIdx, 1)) Then GoTo PROC_EXIT
    Dim prevPos As Long
    prevPos = CLng(headingData(prevIdx, 1))
    m_currentBookIndex = prevIdx
    m_currentBookPos = prevPos
    m_currentChapter = 1   ' Rule 2a
    m_currentVerse = 1     ' Rule 2a
    m_currentChapterPos = 0
    ' Range.Select moves the document cursor (see NextButton comment — same rationale).
    ActiveDocument.Range(prevPos, prevPos).Select
    If Not m_ribbon Is Nothing Then m_ribbon.Invalidate
    Application.OnTime Now, Application.ActiveDocument.VBProject.Name & ".basRibbonDeferred.UpdateStatusBarDeferred"
PROC_EXIT:
    Exit Sub
PROC_ERR:
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
    ' Single-pass scan: captures all H1 (book) and H2 (chapter) heading positions.
    ' H1 positions go to headingData(1..66, 0..1).
    ' H2 positions go to chapterData(bookIdx, chNum) — O(1) lookup replaces FindChapterPos loop.
    On Error GoTo PROC_ERR
    Static hasRun As Boolean
    Dim para As Word.Paragraph
    Dim i        As Long   ' H1 slot index (1-based); incremented after each H1 is stored
    Dim j        As Long   ' chapter Count within current book; reset on each H1
    Dim paraText As String

    If hasRun Then
        Debug.Print "CaptureHeading1s: Already executed this session."
        GoTo PROC_EXIT
    End If

    i = 1
    j = 0
    For Each para In ActiveDocument.Paragraphs
        If para.style = "Heading 1" Then
            If i > 66 Then Exit For
            j = 0
            paraText = Trim(Replace(para.Range.Text, vbCr, ""))
            headingData(i, 0) = paraText
            headingData(i, 1) = para.Range.Start
            i = i + 1
        ElseIf para.style = "Heading 2" Then
            ' i > 1 means at least one H1 has been stored; current book index is i-1
            If i > 1 Then
                j = j + 1
                If j <= 150 Then
                    chapterData(i - 1, j) = para.Range.Start
                End If
            End If
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
    Application.StatusBar = SB_WARM_CACHE
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

' -- Book row -----------------------------------------------------------------

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
    GetBookCount = 0   ' Dropdown population intentionally deferred — not used in current navigation flow
End Function

Public Function GetBookItemLabel(control As IRibbonControl, index As Long) As String
    GetBookItemLabel = ""   ' Dropdown population intentionally deferred — not used in current navigation flow
End Function

Public Function GetBookItemID(control As IRibbonControl, index As Long) As String
    GetBookItemID = CStr(index)   ' Reserved for future dropdown item IDs
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

    If Not matchFound Then
        Application.StatusBar = SB_INVALID_BOOK
        m_bookTextValid = False
        GoTo PROC_EXIT
    End If
    m_bookTextValid = True
    m_chapterTextValid = True   ' new book resets chapter to 1 — always valid
    m_verseTextValid = True     ' new book resets verse to 1 — always valid

    m_currentBookIndex = i
    m_currentBookPos = CLng(headingData(i, 1))
    m_currentChapter = 1   ' default chapter 1 — enables Prev/Next Chapter buttons (#598)
    m_currentVerse = 1     ' default verse 1 — enables Prev/Next Verse buttons (#598)
    m_currentChapterPos = 0

    ' Invalidate so ribbon re-queries enabled state from current field values.
    ' GetChapterEnabled = (m_currentBookIndex > 0) now returns True.
    ' No ScrollIntoView here — the view scrolls when GoToVerseByScan executes
    ' (ScrollIntoView is called there after cursor placement). Removing the call here
    ' eliminates the 4-10s first-load page-layout block from this callback. When
    ' ScrollIntoView was here, Tab key presses during the block were routed through
    ' stale ribbon state, sending Tab to the document instead of cmbVerse (Bug 25a).
    ' The double-Invalidate workaround (Fix 5a) is therefore also removed — a single
    ' Invalidate is sufficient now that there is no blocking call to bracket.
    If Not m_ribbon Is Nothing Then m_ribbon.Invalidate

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OnBookChanged of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Chapter row --------------------------------------------------------------

Private Function FindChapterPos(ByVal chNum As Long) As Long
    ' Direct lookup from pre-built chapterData index (O(1)).
    ' chapterData is populated by CaptureHeading1s at load time — same pass as H1 scan.
    ' Old path: chNum consecutive Range.Find.Execute calls (O(n), up to 119 for Psalm 119).
    ' New path: one array read regardless of chapter number or book.
    If m_currentBookIndex < 1 Or m_currentBookIndex > 66 Then
        FindChapterPos = 0
        Exit Function
    End If
    If chNum < 1 Or chNum > 150 Then
        FindChapterPos = 0
        Exit Function
    End If
    FindChapterPos = chapterData(m_currentBookIndex, chNum)
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

    ' Range.Select moves the document cursor to the chapter heading.
    ' GoToChapter is called only from onAction button callbacks (Prev/Next Chapter)
    ' and from ExecutePendingChapter (currently a no-op). In button click context,
    ' focus returns to the document after activation regardless — Range.Select is
    ' correct here. ScrollIntoView (no cursor update) caused snap-back to the
    ' previous verse when Enter returned focus to the document (Bug 22b).
    Application.StatusBar = SB_NAVIGATING   ' visible during Range.Select layout cost
    ActiveDocument.Range(chPos, chPos).Select
    m_currentChapter = chNum
    m_currentVerse = 1     ' Rule 2a
    m_currentChapterPos = chPos   ' cache for GoToVerse — avoids repeat FindChapterPos calls
    Dim projNameCt As String
    projNameCt = Application.ActiveDocument.VBProject.Name
    Application.OnTime Now, projNameCt & ".basRibbonDeferred.UpdateStatusBarDeferred"

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
    ' Always True — cmbChapter must be an enabled tab stop from initial render.
    ' m_ribbon.Invalidate called from within onChange is deferred (fires after Tab
    ' routing). If GetChapterEnabled depended on m_currentBookIndex, the initial
    ' render cache would show DISABLED and Tab from cmbBook would fall to the document
    ' on first load (Bug 25a root cause). OnChapterChanged guards against no-book input.
    GetChapterEnabled = True
End Function

Public Function GetPrevChapterEnabled(control As IRibbonControl) As Boolean
    GetPrevChapterEnabled = True   ' always-enable (#599); OnPrevChapterClick guards m_currentChapter > 1
End Function

Public Function GetNextChapterEnabled(control As IRibbonControl) As Boolean
    GetNextChapterEnabled = True   ' always-enable (#599); OnNextChapterClick guards m_currentChapter = 0 or m_currentBookIndex = 0
End Function

Public Function GetChapterCount(control As IRibbonControl) As Long
    GetChapterCount = 0   ' Dropdown population not currently used; chapter validation uses aeBibleCitationClass.ChaptersInBook
End Function

Public Function GetChapterItemLabel(control As IRibbonControl, index As Long) As String
    GetChapterItemLabel = ""   ' Dropdown population not currently used
End Function

Public Function GetChapterItemID(control As IRibbonControl, index As Long) As String
    GetChapterItemID = CStr(index)
End Function

Public Function GetChapterText(control As IRibbonControl) As String
    If m_currentChapter > 0 Then
        GetChapterText = CStr(m_currentChapter)
    ElseIf m_currentBookIndex > 0 Then
        GetChapterText = "1"
    Else
        GetChapterText = ""
    End If
End Function

Public Sub OnChapterChanged(control As IRibbonControl, text As String)
    ' onChange fires on each keystroke and on Tab/Enter commit.
    ' If text is empty or out of range, retain the current value (1 by default from
    ' Rule 2a; never leaves the field in an unknown state).
    ' Navigation fires on each valid chapter value — same deferred pattern as OnVerseChanged.
    ' Multiple deferreds from rapid typing all read m_pendingVerse at execution time;
    ' only the first fires a navigation, subsequent ones are no-ops (m_pendingVerse = 0).
    On Error GoTo PROC_ERR
    If m_currentBookIndex = 0 Then GoTo PROC_EXIT
    Dim projNameCh As String
    projNameCh = Application.ActiveDocument.VBProject.Name
    Dim bookNameCh As String
    bookNameCh = CStr(headingData(m_currentBookIndex, 0))
    Dim maxChNum As Long
    maxChNum = aeBibleCitationClass.ChaptersInBook(bookNameCh)
    If Not IsNumeric(Trim(text)) Then
        m_chapterTextValid = False
        m_chapterMax = maxChNum
        Application.OnTime Now, projNameCh & ".basRibbonDeferred.ResetChapterDisplayDeferred"
        GoTo PROC_EXIT
    End If
    Dim chNum As Long
    chNum = CLng(Trim(text))
    If chNum < 1 Or chNum > maxChNum Then
        m_chapterTextValid = False
        m_chapterMax = maxChNum
        Application.OnTime Now, projNameCh & ".basRibbonDeferred.ResetChapterDisplayDeferred"
        GoTo PROC_EXIT
    End If
    m_chapterTextValid = True
    m_currentChapter = chNum
    m_currentChapterPos = 0
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OnChapterChanged of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' Called from basRibbonDeferred.GoToChapterDeferred after the key event clears.
' Does NOT call GoToChapter or ScrollIntoView — those steal ribbon focus from OnTime
' context (Bug 21). Navigation occurs via GoToVerse (uses FindChapterPos) or via the
' Prev/Next Chapter buttons (GoToChapter direct).
Public Sub ExecutePendingChapter()
    ' Fires from basRibbonDeferred.GoToChapterDeferred after the key event clears.
    ' Verse controls are invalidated here rather than in OnChapterChanged so that
    ' they update AFTER Tab routing has already moved focus to cmbVerse (Bug 26).
    m_pendingChapter = 0
    If Not m_ribbon Is Nothing Then
        m_ribbon.InvalidateControl "cmbVerse"
        m_ribbon.InvalidateControl "PrevVerseButton"
        m_ribbon.InvalidateControl "NextVerseButton"
    End If
End Sub

Public Sub OnPrevChapterClick(control As IRibbonControl)
    If m_currentChapter > 1 Then
        GoToChapter m_currentChapter - 1
    ElseIf m_currentBookIndex > 0 Then
        Dim bookNamePC As String
        bookNamePC = CStr(headingData(m_currentBookIndex, 0))
        Application.StatusBar = FormatMsg(SB_ALREADY_FIRST_CHAPTER, bookNamePC, aeBibleCitationClass.ChaptersInBook(bookNamePC))
    End If
End Sub

Public Sub OnNextChapterClick(control As IRibbonControl)
    On Error GoTo PROC_ERR
    If m_currentChapter = 0 Or m_currentBookIndex = 0 Then GoTo PROC_EXIT
    Dim bookName As String
    bookName = CStr(headingData(m_currentBookIndex, 0))
    Dim maxCh As Long
    maxCh = aeBibleCitationClass.ChaptersInBook(bookName)
    If m_currentChapter < maxCh Then
        GoToChapter m_currentChapter + 1
    Else
        Application.StatusBar = FormatMsg(SB_ALREADY_LAST_CHAPTER, bookName, maxCh)
    End If
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OnNextChapterClick of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Verse row ----------------------------------------------------------------

Public Function GetPrevVerseEnabled(control As IRibbonControl) As Boolean
    GetPrevVerseEnabled = True   ' always-enable (#599); OnPrevVerseClick guards m_currentVerse > 1
End Function

Public Function GetNextVerseEnabled(control As IRibbonControl) As Boolean
    GetNextVerseEnabled = True   ' always-enable (#599); OnNextVerseClick guards m_currentVerse = 0 or m_currentChapter = 0 or m_currentBookIndex = 0
End Function

Public Function GetVerseEnabled(control As IRibbonControl) As Boolean
    ' Always True — same reasoning as GetChapterEnabled. cmbVerse must be an enabled
    ' tab stop from initial render so that Tab from cmbChapter reaches it regardless
    ' of when the deferred Invalidate from OnChapterChanged fires. OnVerseChanged
    ' guards against no-chapter input.
    GetVerseEnabled = True
End Function

Public Function GetVerseCount(control As IRibbonControl) As Long
    GetVerseCount = 0   ' Dropdown population not currently used; verse validation uses aeBibleCitationClass.VersesInChapter
End Function

Public Function GetVerseItemLabel(control As IRibbonControl, index As Long) As String
    GetVerseItemLabel = ""   ' Dropdown population not currently used
End Function

Public Function GetVerseItemID(control As IRibbonControl, index As Long) As String
    GetVerseItemID = CStr(index)
End Function

Public Function GetVerseText(control As IRibbonControl) As String
    If m_currentVerse > 0 Then
        GetVerseText = CStr(m_currentVerse)
    ElseIf m_currentBookIndex > 0 Then
        GetVerseText = "1"
    Else
        GetVerseText = ""
    End If
End Function

Public Sub OnVerseChanged(control As IRibbonControl, text As String)
    ' onChange fires on Enter and on each keystroke.
    ' Validates and stores m_currentVerse. Navigation fires from OnGoClick only (#600).
    On Error GoTo PROC_ERR
    Dim projNameVs As String
    projNameVs = Application.ActiveDocument.VBProject.Name
    If m_currentChapter = 0 Or m_currentBookIndex = 0 Then GoTo PROC_EXIT
    Dim bookNameVs As String
    bookNameVs = CStr(headingData(m_currentBookIndex, 0))
    Dim maxVsNum As Long
    maxVsNum = aeBibleCitationClass.VersesInChapter(bookNameVs, m_currentChapter)
    If Not IsNumeric(Trim(text)) Then
        m_verseTextValid = False
        m_verseMax = maxVsNum
        Application.OnTime Now, projNameVs & ".basRibbonDeferred.ResetVerseDisplayDeferred"
        GoTo PROC_EXIT
    End If
    Dim vsNum As Long
    vsNum = CLng(Trim(text))
    If vsNum < 1 Or vsNum > maxVsNum Then
        m_verseTextValid = False
        m_verseMax = maxVsNum
        Application.OnTime Now, projNameVs & ".basRibbonDeferred.ResetVerseDisplayDeferred"
        GoTo PROC_EXIT
    End If
    m_verseTextValid = True
    m_currentVerse = vsNum   ' store confirmed value; GoButton fires navigation
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OnVerseChanged of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

Public Sub UpdateStatusBar()
    ' Called from basRibbonDeferred.UpdateStatusBarDeferred via Application.OnTime.
    ' Fires after Word's own post-onAction status bar refresh, overwriting "Page X of Y"
    ' with the SBL citation for the current position.
    If m_currentBookIndex = 0 Or m_currentChapter = 0 Then Exit Sub
    Application.StatusBar = aeBibleCitationClass.ToSBLShortForm( _
        CStr(headingData(m_currentBookIndex, 0)) & " " & m_currentChapter & ":" & m_currentVerse)
End Sub

Public Sub OnPrevVerseClick(control As IRibbonControl)
    If m_currentVerse > 1 Then
        GoToVerse m_currentVerse - 1
    ElseIf m_currentBookIndex > 0 And m_currentChapter > 0 Then
        Dim bookNamePV As String
        bookNamePV = CStr(headingData(m_currentBookIndex, 0))
        Dim maxVsPV As Long
        maxVsPV = aeBibleCitationClass.VersesInChapter(bookNamePV, m_currentChapter)
        Application.StatusBar = FormatMsg(SB_ALREADY_FIRST_VERSE, bookNamePV, m_currentChapter, maxVsPV)
    End If
End Sub

Public Sub OnNextVerseClick(control As IRibbonControl)
    On Error GoTo PROC_ERR
    If m_currentVerse = 0 Or m_currentChapter = 0 Or m_currentBookIndex = 0 Then GoTo PROC_EXIT
    Dim bookName As String
    bookName = CStr(headingData(m_currentBookIndex, 0))
    Dim maxVs As Long
    maxVs = aeBibleCitationClass.VersesInChapter(bookName, m_currentChapter)
    If m_currentVerse < maxVs Then
        GoToVerse m_currentVerse + 1
    Else
        Application.StatusBar = FormatMsg(SB_ALREADY_LAST_VERSE, bookName, m_currentChapter, maxVs)
    End If
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OnNextVerseClick of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Verse row navigation -----------------------------------------------------

Private Function IsStudyVersion() As Boolean
    ' Study version has one verse per paragraph; paragraph Count > 30,000.
    ' Print candidate has multiple verses per paragraph; Count <= 30,000.
    ' Result is cached after first call — Paragraphs.Count on a 33,857-para document
    ' is non-trivial overhead to repeat on every GoToVerse invocation.
    If Not m_studyVersionSet Then
        m_studyVersionVal = (ActiveDocument.Paragraphs.Count > 30000)
        m_studyVersionSet = True
    End If
    IsStudyVersion = m_studyVersionVal
End Function

Private Sub GoToVerse(ByVal vsNum As Long)
    On Error GoTo PROC_ERR
    If m_currentBookIndex = 0 Then GoTo PROC_EXIT   ' assert: select book first
    If m_currentChapter = 0 Then GoTo PROC_EXIT     ' assert: confirm chapter first

    Dim bookName As String
    bookName = CStr(headingData(m_currentBookIndex, 0))

    Dim maxVs As Long
    maxVs = aeBibleCitationClass.VersesInChapter(bookName, m_currentChapter)
    If vsNum < 1 Or vsNum > maxVs Then GoTo PROC_EXIT

    Dim chPos As Long
    If m_currentChapterPos > 0 Then
        chPos = m_currentChapterPos
    Else
        chPos = FindChapterPos(m_currentChapter)
        If chPos = 0 Then GoTo PROC_EXIT
        m_currentChapterPos = chPos
    End If

    ' Both study and print versions mark each verse with inline "Chapter Verse marker"
    ' (chapter number) + "Verse marker" (verse number) character style runs.
    ' GoToVerseByScan finds the Nth "Verse marker" run — correct for both versions.
    ' GoToVerseByCount (MoveDown paragraph-counting) was study-only and broke on
    ' chapters with section-heading paragraphs (e.g. Psalm 119 has 22 Aleph-Taw
    ' headings → MoveDown(176) landed at verse 155, Bug 25b). IsStudyVersion branch
    ' for verse navigation is therefore obsolete.
    Application.StatusBar = SB_NAVIGATING   ' visible during GoToVerseByScan layout cost
    GoToVerseByScan chPos, vsNum

    m_currentVerse = vsNum
    Dim projNameSt As String
    projNameSt = Application.ActiveDocument.VBProject.Name
    Application.OnTime Now, projNameSt & ".basRibbonDeferred.UpdateStatusBarDeferred"

    If Not m_ribbon Is Nothing Then
        m_ribbon.InvalidateControl "cmbVerse"
        m_ribbon.InvalidateControl "PrevVerseButton"
        m_ribbon.InvalidateControl "NextVerseButton"
        m_ribbon.InvalidateControl "PrevChapterButton"
        m_ribbon.InvalidateControl "NextChapterButton"
    End If

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GoToVerse of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

Private Sub GoToVerseByCount(ByVal chPos As Long, ByVal vsNum As Long)
    ' Dead stub — GoToVerse now always calls GoToVerseByScan (Bug 25b).
    ' Paragraph-move counting broke on chapters with section-heading paragraphs
    ' (Psalm 119: 22 Hebrew-letter headings caused MoveDown(176) → verse 155).
    ' Retained for reference; not called.
End Sub

Private Sub GoToVerseByScan(ByVal chPos As Long, ByVal vsNum As Long)
    ' Find the Nth "Verse marker" inline style run from chPos forward.
    ' Every verse in both study and print versions begins with a "Chapter Verse marker"
    ' run (chapter number) immediately followed by a "Verse marker" run (verse number).
    ' Searching for "Verse marker" skips all non-verse paragraphs (H1, H2, section
    ' headings like Psalm 119's Aleph-Taw headings) without any paragraph counting.
    On Error GoTo PROC_ERR
    Dim r As Word.Range
    Set r = ActiveDocument.Range(chPos, ActiveDocument.Content.End)
    Dim Count As Long
    Count = 0
    Do
        r.Find.ClearFormatting
        r.Find.style = ActiveDocument.Styles("Verse marker")
        r.Find.Text = ""
        r.Find.Forward = True
        r.Find.Wrap = wdFindStop
        If Not r.Find.Execute Then Exit Do
        Count = Count + 1
        If Count = vsNum Then
            ActiveDocument.Range(r.Start, r.Start).Select
            ' ScrollIntoView ensures the viewport updates — on first load the page
            ' layout delay means the view may not scroll without an explicit call.
            ActiveWindow.ScrollIntoView ActiveDocument.Range(r.Start, r.Start), True
            Exit Do
        End If
        r.Start = r.End
        r.End = ActiveDocument.Content.End
    Loop
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GoToVerseByScan of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- Go (navigate) ------------------------------------------------------------

Public Function GetGoEnabled(control As IRibbonControl) As Boolean
    GetGoEnabled = (m_currentBookIndex <> 0)   ' disabled until book selected; consistent with GetNewSearchEnabled
End Function

' Current navigation rule: Book/Chapter/Verse onChange validates and stages state; Go performs verse navigation.
Public Sub OnGoClick(control As IRibbonControl)
    ' Explicit navigation trigger (GoButton, #600).
    ' Reads m_currentBookIndex, m_currentChapter, m_currentVerse — all set by
    ' OnBookChanged / OnChapterChanged / OnVerseChanged before this fires.
    On Error GoTo PROC_ERR
    If m_currentBookIndex = 0 Then GoTo PROC_EXIT
    If Not m_bookTextValid Then
        Application.StatusBar = SB_INVALID_BOOK
        If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "BookComboBox"
        GoTo PROC_EXIT
    End If
    If m_currentChapter = 0 Then GoTo PROC_EXIT
    If Not m_chapterTextValid Then
        Application.StatusBar = FormatMsg(SB_INVALID_CHAPTER, m_chapterMax)
        If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "ChapterComboBox"
        GoTo PROC_EXIT
    End If
    If Not m_verseTextValid Then
        Application.StatusBar = FormatMsg(SB_INVALID_VERSE, m_verseMax)
        If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "VerseComboBox"
        GoTo PROC_EXIT
    End If
    Dim vsNum As Long
    vsNum = m_currentVerse
    If vsNum < 1 Then vsNum = 1
    GoToVerse vsNum
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OnGoClick of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub

' -- New Search ----------------------------------------------------------------

Public Function GetNewSearchEnabled(control As IRibbonControl) As Boolean
    GetNewSearchEnabled = m_currentBookIndex <> 0
End Function

Public Sub OnNewSearchClick(control As IRibbonControl)
    m_currentBookIndex = 0
    m_currentBookPos = 0
    m_currentChapter = 0
    m_currentVerse = 0
    m_currentChapterPos = 0
    m_bookTextValid = True
    m_chapterTextValid = True
    m_verseTextValid = True
    If Not m_ribbon Is Nothing Then m_ribbon.Invalidate
    ' Bug #597 — focus cmbBook after reset.
    ' SendKeys approach attempted (Option A) but rejected: VBA SendKeys %X sends
    ' Alt+X as a chord; ribbon keytip mode requires Alt tapped alone then plain
    ' key presses — not expressible in SendKeys syntax. Characters leaked into
    ' the document. Accepted as known limitation (Option C): after New Search,
    ' user re-focuses cmbBook manually (click or Alt,Y2,B).
End Sub

' -- Archived reference paths (not used by current ribbon flow) ----------------
' These methods are retained for XML/control-type fallback reference only.
' NavigateToCurrentBook: navigation removed from OnBookChanged (Bug 9);
'   book-level navigation now occurs in GoToChapter via m_currentBookPos.
' OnChapterAction / OnVerseAction: onAction is not fired by comboBox controls
'   in the customUI14 schema; retained in case control type reverts to editBox.
' ExecutePendingVerse: navigation trigger moved to OnGoClick (#600);
'   m_pendingVerse is never set so this is a permanent no-op.

Public Sub NavigateToCurrentBook()
    ' ActiveDocument.Range(m_currentBookPos, m_currentBookPos).Select
End Sub

Public Sub OnChapterAction(control As IRibbonControl, text As String)
End Sub

Public Sub OnVerseAction(control As IRibbonControl, text As String)
End Sub

Public Sub ExecutePendingVerse()
    If m_pendingVerse > 0 Then
        GoToVerse m_pendingVerse
        m_pendingVerse = 0
    End If
End Sub

