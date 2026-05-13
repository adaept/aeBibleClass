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
Private m_lastScanWasSaved As Boolean ' ActiveDocument.Saved value at end of last CaptureHeading1s; cache valid only when both this and current Saved are True

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
    m_lastScanWasSaved = False
End Sub
Private Sub Class_Terminate()
    Debug.Print "RibbonController: Class_Terminate at " & Format(Now, "hh:nn:ss")
    Set m_ribbon = Nothing
    m_ribbonReady = False
End Sub
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
Public Sub InvalidateControl(controlId As String)
    On Error GoTo PROC_ERR
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl controlId
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure InvalidateControl of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub
Public Sub ResetChapterDisplay()
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl CTRL_CHAPTER
End Sub
Public Sub ResetVerseDisplay()
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl CTRL_VERSE
End Sub
Public Sub OnPrevButtonClick(control As IRibbonControl)
    Call PrevButton
End Sub
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
Public Function GetNextBkEnabled(control As IRibbonControl) As Boolean
    GetNextBkEnabled = True   ' always-enable (#599); OnNextBkClick guards m_currentBookIndex = 0 or >= 66
End Function
Public Function GetPrevBkEnabled(control As IRibbonControl) As Boolean
    GetPrevBkEnabled = True   ' always-enable (#599); OnPrevBkClick guards m_currentBookIndex <= 1
End Function
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
Private Sub GoToH1()
    On Error GoTo PROC_ERR
    Dim pattern  As String
    Dim i        As Long
    Dim foundPos As Long
    Dim matchFound As Boolean

    ' Refresh headings if cache is empty OR document was edited since last scan.
    ' CaptureHeading1s self-decides via its m_lastScanWasSaved / ActiveDocument.Saved
    ' check; calling unconditionally is cheap when the cache is still valid.
    CaptureHeading1s

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
Private Sub NextButton()
    ' Navigate to the next book using ribbon state (m_currentBookIndex), not Selection.
    ' Using Selection caused Bug 19: ScrollIntoView does not move the document cursor,
    ' so after OnBookChanged the cursor lagged behind the ribbon state. Next/Prev were
    ' navigating from the stale cursor position instead of from the current book.
    On Error GoTo PROC_ERR
    ' Refresh heading cache if document was edited since last scan; cheap no-op otherwise.
    CaptureHeading1s
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
    ' Force H1 to top of viewport with no double-scroll flicker.
    '   1. ScrollIntoView a non-zero range (H1 .. end-of-doc) with Start:=True so
    '      Word has a real upper-left corner (the H1) to align to the top. A
    '      zero-width (pos, pos) range degenerates and falls back to "ensure
    '      visible" - same as Word's default minimal scroll, leaves Next at
    '      bottom and Prev wobbling.
    '   2. Selection.SetRange after the scroll places the cursor WITHOUT
    '      triggering Word's "scroll-to-cursor" heuristic that .Select fires.
    '      That heuristic was the source of the prior caret-flash artifact.
    ' ScreenUpdating off suppresses any intermediate paint - one render per click.
    Application.ScreenUpdating = False
    Dim rNextView As Word.Range
    Set rNextView = ActiveDocument.Range(nextPos, ActiveDocument.Content.End)
    ActiveWindow.ScrollIntoView rNextView, True
    Selection.SetRange Start:=nextPos, End:=nextPos
    Application.ScreenUpdating = True
    If Not m_ribbon Is Nothing Then m_ribbon.Invalidate
    Application.OnTime Now, Application.ActiveDocument.VBProject.Name & ".basRibbonDeferred.UpdateStatusBarDeferred"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    Application.ScreenUpdating = True
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure NextButton of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub
Private Sub PrevButton()
    ' Navigate to the previous book using ribbon state (m_currentBookIndex), not Selection.
    ' See NextButton comment — same root cause (Bug 19).
    On Error GoTo PROC_ERR
    ' Refresh heading cache if document was edited since last scan; cheap no-op otherwise.
    CaptureHeading1s
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
    ' Same pattern as NextButton - see NextButton comment for the why.
    Application.ScreenUpdating = False
    Dim rPrevView As Word.Range
    Set rPrevView = ActiveDocument.Range(prevPos, ActiveDocument.Content.End)
    ActiveWindow.ScrollIntoView rPrevView, True
    Selection.SetRange Start:=prevPos, End:=prevPos
    Application.ScreenUpdating = True
    If Not m_ribbon Is Nothing Then m_ribbon.Invalidate
    Application.OnTime Now, Application.ActiveDocument.VBProject.Name & ".basRibbonDeferred.UpdateStatusBarDeferred"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    Application.ScreenUpdating = True
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure PrevButton of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub
Public Sub CaptureHeading1s(Optional ByVal bForce As Boolean = False)
    ' Single-pass scan: captures all H1 (book) and H2 (chapter) heading positions.
    ' H1 positions go to headingData(1..66, 0..1).
    ' H2 positions go to chapterData(bookIdx, chNum) - O(1) lookup replaces FindChapterPos loop.
    '
    ' Cache validity (Option 2 - saved-flag invalidation):
    '   Skip rescan only when BOTH the previous scan and the current state report
    '   ActiveDocument.Saved = True. If either is False, the document may have been
    '   edited since the last scan (cached char positions go stale on any insertion
    '   / deletion); rescan to be safe. Pass bForce:=True to override (e.g., for a
    '   manual Refresh button).
    On Error GoTo PROC_ERR
    Static hasRun As Boolean
    Dim para As Word.Paragraph
    Dim i        As Long   ' H1 slot index (1-based); incremented after each H1 is stored
    Dim j        As Long   ' chapter Count within current book; reset on each H1
    Dim paraText As String

    If hasRun And Not bForce And m_lastScanWasSaved And ActiveDocument.Saved Then
        Debug.Print "CaptureHeading1s: cache valid (no edits since last scan)."
        GoTo PROC_EXIT
    End If

    ' Erase before rescan: if the new scan finds fewer H1/H2 entries than the
    ' previous one, stale tail entries would otherwise remain.
    Erase headingData
    Erase chapterData

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
    m_lastScanWasSaved = ActiveDocument.Saved
    Debug.Print "CaptureHeading1s: Stored " & i - 1 & " Heading 1 entries (saved=" & ActiveDocument.Saved & ")."

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure CaptureHeading1s of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub
Public Sub LogHeadingData()
    On Error GoTo PROC_ERR
    ' Diagnostic CSV; skip silently when no rpt\ folder exists beside the
    ' active document (production template / docx hosts will not have one).
    If Dir(ActiveDocument.Path & "\rpt", vbDirectory) = "" Then Exit Sub
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
Private Function AdaeptMsg() As String
    AdaeptMsg = """...the truth shall make you free.""" & " John 8:32 (KJV)"
End Function
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

    ' Refresh headings if the document was edited since last scan; cheap no-op
    ' when cache is still valid.
    CaptureHeading1s

    ' Two-layer resolution:
    '   Layer 1 - alias-authoritative: try aeBibleCitationClass.ResolveAlias
    '     against the typed text. If it returns a canonical name (e.g. "Nahum"),
    '     match THAT against headingData. This makes every alias in aliasMap
    '     work and eliminates the first-substring-wins collisions of the
    '     legacy single-layer scan (e.g. "NAH" -> Jonah because JONAH is
    '     scanned before NAHUM and contains "NAH").
    '   Layer 2 - legacy substring fallback: if the typed text is not a
    '     known alias, fall back to the substring Like scan against
    '     headingData. Preserves partial-typing UX (e.g. "Genesi" + Tab).
    Dim i As Long
    Dim matchFound As Boolean
    Dim canonical As String

    On Error Resume Next
    canonical = aeBibleCitationClass.ResolveAlias(text)
    On Error GoTo PROC_ERR

    If Len(canonical) > 0 Then
        For i = 1 To 66
            If Not IsEmpty(headingData(i, 0)) Then
                If CStr(headingData(i, 0)) Like "*" & canonical & "*" Then
                    matchFound = True
                    Exit For
                End If
            End If
        Next i
    End If

    If Not matchFound Then
        For i = 1 To 66
            If Not IsEmpty(headingData(i, 0)) Then
                If CStr(headingData(i, 0)) Like "*" & pattern & "*" Then
                    matchFound = True
                    Exit For
                End If
            End If
        Next i
    End If

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
    ' Refresh heading cache if document was edited since last scan; cheap no-op otherwise.
    ' chapterData is repopulated by CaptureHeading1s in the same pass as headingData.
    CaptureHeading1s
    If m_currentBookIndex = 0 Then GoTo PROC_EXIT   ' assert: select book first

    Dim bookName As String
    bookName = CStr(headingData(m_currentBookIndex, 0))

    Dim maxCh As Long
    maxCh = aeBibleCitationClass.ChaptersInBook(bookName)
    If chNum < 1 Or chNum > maxCh Then GoTo PROC_EXIT

    Dim chPos As Long
    chPos = FindChapterPos(chNum)
    If chPos = 0 Then GoTo PROC_EXIT

    ' Force chapter H2 to top of viewport with no double-scroll wobble.
    ' Same three-step pattern as NextButton / PrevButton (book nav):
    '   1. ScrollIntoView a non-zero range (chapter .. EOD) with Start:=True so
    '      Word has a real upper-left corner (the H2) to align to the top. A
    '      zero-width (chPos, chPos) range degenerates to "ensure visible".
    '   2. Selection.SetRange places the cursor WITHOUT firing Word's
    '      "scroll-to-cursor" heuristic that .Select triggers (which was the
    '      Bug 22b snap-back source).
    ' ScreenUpdating off suppresses any intermediate paint - one render per click.
    Application.StatusBar = SB_NAVIGATING   ' visible during scroll layout cost
    Application.ScreenUpdating = False
    Dim rChView As Word.Range
    Set rChView = ActiveDocument.Range(chPos, ActiveDocument.Content.End)
    ActiveWindow.ScrollIntoView rChView, True
    Selection.SetRange Start:=chPos, End:=chPos
    Application.ScreenUpdating = True
    m_currentChapter = chNum
    m_currentVerse = 1     ' Rule 2a
    m_currentChapterPos = chPos   ' cache for GoToVerse — avoids repeat FindChapterPos calls
    Dim projNameCt As String
    projNameCt = Application.ActiveDocument.VBProject.Name
    Application.OnTime Now, projNameCt & ".basRibbonDeferred.UpdateStatusBarDeferred"

    If Not m_ribbon Is Nothing Then
        m_ribbon.InvalidateControl CTRL_CHAPTER
        m_ribbon.InvalidateControl CTRL_PREV_CHAPTER
        m_ribbon.InvalidateControl CTRL_NEXT_CHAPTER
        m_ribbon.InvalidateControl CTRL_VERSE
        m_ribbon.InvalidateControl CTRL_PREV_VERSE
        m_ribbon.InvalidateControl CTRL_NEXT_VERSE
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
Public Sub ExecutePendingChapter()
    ' Fires from basRibbonDeferred.GoToChapterDeferred after the key event clears.
    ' Verse controls are invalidated here rather than in OnChapterChanged so that
    ' they update AFTER Tab routing has already moved focus to cmbVerse (Bug 26).
    m_pendingChapter = 0
    If Not m_ribbon Is Nothing Then
        m_ribbon.InvalidateControl CTRL_VERSE
        m_ribbon.InvalidateControl CTRL_PREV_VERSE
        m_ribbon.InvalidateControl CTRL_NEXT_VERSE
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
        m_ribbon.InvalidateControl CTRL_VERSE
        m_ribbon.InvalidateControl CTRL_PREV_VERSE
        m_ribbon.InvalidateControl CTRL_NEXT_VERSE
        m_ribbon.InvalidateControl CTRL_PREV_CHAPTER
        m_ribbon.InvalidateControl CTRL_NEXT_CHAPTER
    End If

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GoToVerse of Class aeRibbonClass"
    Resume PROC_EXIT
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
            ' Force verse marker to top of viewport - same three-step pattern as
            ' NextButton / GoToChapter. ScrollIntoView with a non-zero range
            ' (verse .. EOD) gives Word a real upper-left corner; SetRange places
            ' the cursor without firing .Select's scroll-to-cursor heuristic.
            Application.ScreenUpdating = False
            Dim rVsView As Word.Range
            Set rVsView = ActiveDocument.Range(r.Start, ActiveDocument.Content.End)
            ActiveWindow.ScrollIntoView rVsView, True
            Selection.SetRange Start:=r.Start, End:=r.Start
            ' Transfer focus from the ribbon to the document body so Word
            ' renders the caret. Without this the Selection is updated but
            ' the I-beam stays invisible until another action moves focus.
            ActiveDocument.ActiveWindow.Activate
            Application.ScreenUpdating = True
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
Public Function GetGoEnabled(control As IRibbonControl) As Boolean
    GetGoEnabled = (m_currentBookIndex <> 0)   ' disabled until book selected; consistent with GetNewSearchEnabled
End Function
Public Sub OnGoClick(control As IRibbonControl)
    ' Explicit navigation trigger (GoButton, #600).
    ' Reads m_currentBookIndex, m_currentChapter, m_currentVerse — all set by
    ' OnBookChanged / OnChapterChanged / OnVerseChanged before this fires.
    On Error GoTo PROC_ERR
    If m_currentBookIndex = 0 Then GoTo PROC_EXIT
    If Not m_bookTextValid Then
        Application.StatusBar = SB_INVALID_BOOK
        If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl CTRL_BOOK
        GoTo PROC_EXIT
    End If
    If m_currentChapter = 0 Then GoTo PROC_EXIT
    If Not m_chapterTextValid Then
        Application.StatusBar = FormatMsg(SB_INVALID_CHAPTER, m_chapterMax)
        If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl CTRL_CHAPTER
        GoTo PROC_EXIT
    End If
    If Not m_verseTextValid Then
        Application.StatusBar = FormatMsg(SB_INVALID_VERSE, m_verseMax)
        If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl CTRL_VERSE
        GoTo PROC_EXIT
    End If
    Dim vsNum As Long
    vsNum = m_currentVerse
    If vsNum < 1 Then vsNum = 1
    GoToVerse vsNum
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl CTRL_BOOK
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OnGoClick of Class aeRibbonClass"
    Resume PROC_EXIT
End Sub
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
    ' Bug #597 — focusing cmbBook after reset is a known VBA limitation.
    ' See rvw/Code_review 2026-04-20.md § 12 for analysis and VSTO resolution.
End Sub

