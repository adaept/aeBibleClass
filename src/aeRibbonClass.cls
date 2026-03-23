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

' -- Private State -------------------------------------------------------------
Private m_ribbon         As IRibbonUI
Private m_ribbonReady    As Boolean
Private m_btnNextEnabled As Boolean
Private m_bookmarkIndex  As Long
Private m_savedPos       As Long
Private m_bookAbbr       As String
Private headingData(1 To 66, 0 To 1) As Variant

' Windows API - cursor control
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" ( _
    ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

' -- Lifecycle -----------------------------------------------------------------

Private Sub Class_Initialize()
    Debug.Print "RibbonController: Class_Initialize at " & Format(Now, "hh:nn:ss")
    m_btnNextEnabled = True
    m_ribbonReady = False
    m_bookmarkIndex = 0
    m_savedPos = 0
    m_bookAbbr = vbNullString
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

' -- Ribbon Load ---------------------------------------------------------------

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set m_ribbon = ribbon
    m_ribbonReady = True
    Debug.Print "RibbonController: Ribbon ready at " & Format(Now, "hh:nn:ss")
    m_ribbon.InvalidateControl "GoToNextButton"
    Call EnableButtonsRoutine
End Sub

' -- Invalidation --------------------------------------------------------------

Public Sub InvalidateAll()
    If Not m_ribbon Is Nothing Then m_ribbon.Invalidate
End Sub

Public Sub InvalidateControl(controlId As String)
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl controlId
End Sub

' -- Ribbon Callbacks ----------------------------------------------------------

Public Sub OnGoToVerseSblClick(control As IRibbonControl)
    Call GoToVerseSBL
End Sub

Public Sub OnGoToH1ButtonClick(control As IRibbonControl)
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

' -- Initialization ------------------------------------------------------------

Private Sub EnableButtonsRoutine()
    Debug.Print "RibbonController: EnableButtonsRoutine"
    m_btnNextEnabled = True
    If Not m_ribbon Is Nothing Then m_ribbon.InvalidateControl "GoToNextButton"
    CaptureHeading1s
    LogHeadingData
End Sub

' -- GoTo Verse (stub - parser to be implemented) ------------------------------

Private Sub GoToVerseSBL()
    MsgBox "GoToVerseSBL - Parser not yet implemented."
End Sub

' -- GoTo Book (H1) ------------------------------------------------------------

Private Sub GoToH1()
    Dim pattern  As String
    Dim para As Word.Paragraph
    Dim paraText As String
    Dim matchFound As Boolean

    pattern = InputBox("Enter a Book Name (Heading 1) abbreviation:", "Go To Bible Book")
    If pattern = "" Then Exit Sub

    matchFound = False
    Application.ScreenUpdating = False

    For Each para In ActiveDocument.Paragraphs
        If para.style = "Heading 1" Then
            paraText = Trim$(para.Range.Text)
            If paraText Like "*" & UCase(pattern) & "*" Then
                para.Range.Select
                ActiveDocument.Range(para.Range.Start, para.Range.Start).Select
                matchFound = True
                Exit For
            End If
        End If
    Next para

    Application.ScreenUpdating = True
    Selection.Range.Select
    DoEvents

    If Not matchFound Then
        MsgBox "Book not found! No Heading 1 matches: '" & pattern & "'", vbExclamation, "Bible"
    End If
End Sub

' -- Next Book (H1) ------------------------------------------------------------

Private Sub NextButton()
    Dim doc         As Document
    Dim searchRange As Word.Range
    Dim paraEnd     As Long
    Dim found       As Boolean

    Set doc = ActiveDocument
    found = False

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

    If found Then
        Selection.SetRange searchRange.Start, searchRange.Start
        ActiveWindow.ScrollIntoView Selection.Range, True
    Else
        MsgBox "No Heading 1 found in the document.", vbInformation
    End If
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
    Static hasRun As Boolean
    Dim para As Word.Paragraph
    Dim i        As Long
    Dim paraText As String

    If hasRun Then
        Debug.Print "CaptureHeading1s: Already executed this session."
        Exit Sub
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

    fPath = "C:\adaept\aeBibleClass\rpt\HeadingLog.txt"
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

Private Function AdaeptMsg() As String
    AdaeptMsg = """...the truth shall make you free.""" & " John 8:32 (KJV)"
End Function

