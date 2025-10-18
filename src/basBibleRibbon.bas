Attribute VB_Name = "basBibleRibbon"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString
Private savedPos As Long
Private bookAbbr As String

' Use Windows API to change cursor
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" ( _
    ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
' Global ribbon object and button state flags
Public ribbonUI As IRibbonUI
Public ribbonIsReady As Boolean
Public btnNextEnabled As Boolean
Dim bookmarkIndex As Long

' Initialize early
Public Sub AutoExec()
    Debug.Print "In AutoExec routine"
    btnNextEnabled = True
End Sub

'Public Function UserConfirmed(promptText As String, Optional promptTitle As String = "Hide User Interface?") As Boolean
'    Dim response As VbMsgBoxResult
'    response = MsgBox(promptText, vbYesNo + vbQuestion, promptTitle)
'    UserConfirmed = (response = vbYes)
'End Function

Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set ribbonUI = ribbon
    ribbonIsReady = True
    Debug.Print "Ribbon ready at: " & Format(Now, "hh:nn:ss")
    ' Invalidate a specific control only at initialization
    ribbonUI.InvalidateControl "GoToNextButton"
    ' Optional: force ribbon refresh after init
'    ribbonUI.Invalidate
    Call EnableButtonsRoutine
End Sub

Sub EnableButtonsRoutine()
    Debug.Print "In EnableButtonsRoutine"
    btnNextEnabled = True
    ribbonUI.InvalidateControl "GoToNextButton"
End Sub

' Callback to dynamically enable or disable buttons
Public Function GetNextEnabled(control As IRibbonControl) As Boolean
    ' Make sure var is set to account for ribbon load timing mismatch
    If isEmpty(btnNextEnabled) Then btnNextEnabled = True
    GetNextEnabled = btnNextEnabled
End Function

Public Sub OnGoToVerseSblClick(control As IRibbonControl)
    Call GoToVerseSBL
End Sub

Public Sub OnHelloWorldButtonClick(control As IRibbonControl)
   MsgBox "Hello, SILAS World!" & vbCrLf & _
                "GetVScroll  = " & GetExactVerticalScroll
End Sub

Public Sub OnSectionButtonClick(control As IRibbonControl)
    Call GoToSection
End Sub

Public Sub OnGoToH1ButtonClick(control As IRibbonControl)
    Call GoToH1
End Sub

Public Sub OnNextButtonClick(control As IRibbonControl)
    Call NextButton
End Sub

Public Sub OnAdaeptAboutClick(control As IRibbonControl)
    MsgBox "Hello, adaept World!" & vbCrLf & _
                "adaeptMsg  = " & adaeptMsg, vbInformation, "About adaept"
End Sub

Private Function adaeptMsg() As String
    adaeptMsg = """...the truth shall make you free.""" & " John 8:32 (KJV)"
End Function

Private Function GetParaIndexSafe(rng As range) As Long
' Search Isa 23:42 (intentional false verse number) scanned nearly 9,000 paragraphs in under a quarter second,
' with full interruptibility and no layout lock
    Dim r As range
    Set r = ActiveDocument.range(0, 0)

    Dim idx As Long: idx = 1
    Dim tickStart As Single: tickStart = Timer
    Dim tickNow As Single

    Do While r.Start < rng.Start
        Set r = r.Next(Unit:=wdParagraph)
        idx = idx + 1

        If idx Mod 500 = 0 Then
            tickNow = Timer
            'Debug.Print "Step " & idx & " ? Range.Start=" & r.Start & " | Elapsed=" & Format(tickNow - tickStart, "0.00") & "s"
            If tickNow - tickStart > 5 Then
                Debug.Print "> Timeout: Paragraph scan exceeded 5 seconds. Breaking out."
                GetParaIndexSafe = -2
                Exit Function
            End If
        End If
    Loop

    If r.Start = rng.Start Then
        'Debug.Print ">> Found match at paragraph #" & idx
        GetParaIndexSafe = idx
    Else
        'Debug.Print ">>> No exact match. Closest index: " & idx
        GetParaIndexSafe = -1
    End If
End Function

Private Function StyleTypeLabel(st As WdStyleType) As String
    Select Case st
        Case wdStyleTypeParagraph: StyleTypeLabel = "Paragraph"
        Case wdStyleTypeCharacter: StyleTypeLabel = "Character"
        Case wdStyleTypeTable:     StyleTypeLabel = "Table"
        Case wdStyleTypeList:      StyleTypeLabel = "List"
        Case Else:                 StyleTypeLabel = "Unknown"
    End Select
End Function

Public Sub GoToVerseSBL()
    On Error GoTo ErrHandler
    Application.StatusBar = "Searching for verse..."
    
    Dim userInput As String, i As Long
    userInput = InputBox("Enter verse (e.g. 1 Sam 1:1):", "Go to Verse (SBL Format)")
    userInput = Trim(UCase(userInput))
    If Trim(userInput) = "" Then Exit Sub

    Dim tickStartSBL As Single: tickStartSBL = Timer
    Dim tickNowSBL As Single

    Dim chapNum As String, verseNum As String, paraIndex As Long
    Dim parts() As String, subParts() As String
    
    Dim hWaitCursor As Long
    ' Set spinning cursor manually
    hWaitCursor = LoadCursor(0, 32514) ' 32514 = Busy (Hourglass)
    SetCursor hWaitCursor
    Application.ScreenUpdating = False  ' Prevent flickering

    ' Parse the input
    parts = ParseParts(userInput, ":")
    Debug.Print "UBound(parts) = " & UBound(parts)
   
    Dim fullBookName As String
    If UBound(parts) = 0 Then   ' No ":" delimeter is used
        Select Case Left(parts(0), 1)  ' Book starts with a number
            Case "1", "2", "3"
                'Debug.Print "Starts with 1, 2, or 3 " & "'" & parts(0) & "'"
                ' If the rightmost character is not a digit then we have a Book name only
                bookAbbr = Trim(parts(0))
                If Not (Right(userInput, 1) Like "#") Then
                    fullBookName = GetFullBookName(bookAbbr)
                    ' Optional default chapNum = "1" and verseNum = "1"
                    Debug.Print "Starts with 1, 2, or 3 " & "fullBookName = " & fullBookName
                    FindBookH1 fullBookName, paraIndex
                End If
            Case Else
                'Debug.Print "Does not start with 1, 2, or 3 " & "'" & parts(0) & "'"
                bookAbbr = Trim(parts(0))
                ' If the rightmost character is not a digit then we have a Book name only
                If Not (Right(userInput, 1) Like "#") Then
                    fullBookName = GetFullBookName(bookAbbr)
                    ' Optional default chapNum = "1" and verseNum = "1"
                    Debug.Print "Does Not Start with 1, 2, or 3 " & "fullBookName = " & fullBookName
                    FindBookH1 fullBookName, paraIndex
                Else    ' Found digits indicate a chapter number then set verseNum = "1"
                     
                End If
        End Select
        Debug.Print "paraIndex = " & paraIndex

Selection.range.Select  ' re-activate the cursor
GoTo Cleanup    ' for Exit Sub temp stop

        verseNum = 1
        GoTo Chapter
'    ElseIf UBound(parts) <> 1 Then
'        Application.ScreenUpdating = True   ' Restore normal UI
'        SetCursor LoadCursor(0, 32512)      ' Restore default arrow cursor
'        MsgBox "Invalid format. Use format like '1 Sam 1:1'", vbExclamation
'        Exit Sub
    End If
    verseNum = Trim(parts(1))
Chapter:
    subParts = Split(Trim(parts(0)))
    If UBound(subParts) = 0 Then
        bookAbbr = Trim(parts(0))
        chapNum = "1"
    ElseIf (subParts(0) = "1" Or subParts(0) = "2") And UBound(subParts) = 1 Then
        bookAbbr = subParts(0) & " " & subParts(1)
        'Debug.Print "a:", bookAbbr
        chapNum = "1"
    Else
        bookAbbr = ""
        For i = 0 To UBound(subParts) - 1
            bookAbbr = bookAbbr & subParts(i) & " "
            'Debug.Print "b:", bookAbbr
        Next i
        bookAbbr = Trim(bookAbbr)
        Debug.Print ">", bookAbbr
        chapNum = Trim(subParts(UBound(subParts)))
        Debug.Print ">>", chapNum
    End If
    
    fullBookName = GetFullBookName(bookAbbr)
    Debug.Print "a>>>", fullBookName
    If fullBookName = "" Then
        Application.ScreenUpdating = True   ' Restore normal UI
        SetCursor LoadCursor(0, 32512)      ' Restore default arrow cursor
        MsgBox "Book not found: " & bookAbbr, vbExclamation
        Exit Sub
    End If
 
    ' Find the Heading 2 for the chapter or psalm
    Dim theChapter As String, chapFound As Boolean, chapIdx As Long, paraIdx As Long, bookIdx As Long, j As Long

    ' chapIdx is set to the found book paragraph index
    Dim para As paragraph, theBook As String
    chapIdx = bookIdx
    For j = chapIdx To ActiveDocument.paragraphs.count
        Set para = ActiveDocument.paragraphs(j)
        If para.range.Start < Selection.range.Start Then GoTo SkipChapter
        
        
'        'Debug.Print "Paragraph #" & j & ": " & Left(para.range.text, 40)
'
        If para.style = "Heading 2" Then
'            Select Case theBook  ' Books of only one chapter
'                Case "OBADIAH", "PHILEMON", "2 JOHN", "3 JOHN", "JUDE"
'                    Dim tmp As String
'                    Debug.Print "verse = " & verseNum, "chapter = " & chapNum
'                    If chapNum > 1 Then verseNum = chapNum
'                    chapNum = 1
'                    'Debug.Print "verse = " & verseNum, "chapter = " & chapNum
'                    para.range.Select
'                    Application.Selection.range.GoTo
'                    Selection.Collapse Direction:=wdCollapseEnd
'                    chapFound = True
'                    chapIdx = j
'                    Debug.Print "!Paragraph #" & chapIdx, chapNum, verseNum
'                    'Stop
'                    Exit For
'                Case Else
'                    ' Multi-chapter books—continue
'            End Select

            If Trim(para.range.text) Like "*Chapter " & chapNum & "*" _
                    Or Trim(para.range.text) Like "*Psalm " & chapNum & "*" Then

                para.range.Select

                Dim idx As Long
                idx = GetParaIndexSafe(para.range)
                'Debug.Print "idx = " & idx

                Select Case idx
                Case Is >= 1
                    Debug.Print "Jumped to paragraph #" & idx
                Case -1
                    'Debug.Print "Paragraph not found."
                Case -2
                    'Debug.Print "Scan timed out—possible layout stall."
                End Select
                
                Dim styleName As String
                styleName = Trim(para.style.NameLocal)
                Dim s As style

                styleName = para.style ' This is a Variant containing the name
                Set s = ActiveDocument.Styles(styleName)

                'Debug.Print "Style: " & s.NameLocal & _
                    " | Type=" & StyleTypeLabel(s.Type) & _
                    " | OutlineLevel=" & s.ParagraphFormat.OutlineLevel & _
                    " | Content=" & Trim(Replace(para.range.text, vbCr, ""))

                Dim suffixChar As String
                Dim suffixCode As Integer
                suffixChar = Right(Trim(para.range.text), 1)

                If Len(suffixChar) = 1 Then
                    suffixCode = Asc(suffixChar)
    
                    Select Case suffixCode
                    Case 0 To 31, 127
                        'Debug.Print "Suffix: [ASCII " & suffixCode & "]"
                    Case Else
                        'Debug.Print "Suffix: '" & suffixChar & "' [ASCII " & suffixCode & "]"
                    End Select
                Else
                    'Debug.Print "Suffix: [None]"
                End If

                chapFound = True
                chapIdx = idx
                Debug.Print "chapIdx = " & idx
                Application.Selection.range.GoTo
                Stop
                Exit For
            End If
        End If
SkipChapter:
    chapIdx = j
    'xxxNext para
    Next j

    If Not chapFound Then
        Application.ScreenUpdating = True   ' Restore normal UI
        SetCursor LoadCursor(0, 32512)      ' Restore default arrow cursor
        MsgBox "Chapter not found: " & chapNum, vbExclamation
        Exit Sub
    End If

    Dim p As paragraph
    Dim v As Long
    Dim targetVerse As String: targetVerse = verseNum
    Dim charStyleName As String: charStyleName = "Verse marker"
    Dim tickStart As Single: tickStart = Timer
    Dim tickLimit As Single: tickLimit = tickStart + 5
    Dim maxScan As Long: maxScan = 5000
    Dim found As Boolean
    Dim normalized As String

    idx = chapIdx
    Debug.Print "Starting verse marker scan from paragraph #" & idx

    For v = idx + 1 To idx + maxScan
        If Timer > tickLimit Then
            Debug.Print "Timeout: Scan exceeded 5 seconds. Aborting."
            Exit For
        End If

        Set p = ActiveDocument.paragraphs(v)
        Dim styleNameH2 As String: styleNameH2 = Trim(p.style.NameLocal)

        If InStr(styleNameH2, "Heading 2") > 0 Then
            Dim pageNum As Long
            paraIndex = v ' your known paragraph index
            pageNum = ActiveDocument.paragraphs(paraIndex).range.Information(wdActiveEndPageNumber) - 2 ' to get actual page number of doc using ^H
            'MsgBox "Paragraph " & paraIndex & " is on page " & pageNum
            Debug.Print "Error: Reached next chapter at paragraph #" & v & " (style: '" & styleNameH2 & "')", Left(para.range.text, 40), "Page " & pageNum
            MsgBox "No verse " & verseNum & " found in Chapter " & chapNum, vbCritical
            Exit For
        End If

        Dim rng As range: Set rng = p.range.Duplicate
        With rng.Find
            .ClearFormatting
            .style = charStyleName
            .text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            Do While .Execute
                normalized = Replace(rng.text, ChrW(8239), "")
                If normalized = targetVerse Then
                    Debug.Print "Found verse '" & targetVerse & "' at paragraph #" & v
                    rng.Select
                    found = True
                    Exit For
                End If
                rng.Start = rng.End ' Move to next match
                rng.End = p.range.End
            Loop
        End With

        If found Then Exit For
    Next v
    Debug.Print "Scan complete. Elapsed time: " & Format(Timer - tickStart, "0.00") & " seconds."

Cleanup:
    tickNowSBL = Timer
    Debug.Print "GoToVerseSBL complete. Elapsed time: " & Format(tickNowSBL - tickStartSBL, "0.00") & " seconds."

    Application.ScreenUpdating = True   ' Restore normal UI
    SetCursor LoadCursor(0, 32512)      ' Restore default arrow cursor
    Application.StatusBar = False
    Exit Sub

ErrHandler:
    MsgBox "Erl = " & Erl & " Err = " & Err & vbCrLf & Err.Description, vbCritical, "Error during Bible verse search "
    Resume Cleanup
End Sub

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
    Debug.Print "chapNum = " & chapNum, "verseNum = " & verseNum
    savedPos = SaveCursor()
    ' Find the Heading 1 for the book
    Dim theBook As String
    theBook = ""
    Dim para As paragraph, foundBook As Boolean, bookIdx As Integer
    bookIdx = 0
    For Each para In ActiveDocument.paragraphs
        bookIdx = bookIdx + 1
        If para.style = "Heading 1" Then
            theBook = Trim(para.range.text)
            theBook = UCase(Replace(para.range.text, vbCr, ""))
            'Debug.Print "FindBookH1: >", theBook, fullBookName
            If theBook = fullBookName Then
                para.range.Select
                Application.Selection.range.GoTo
                Selection.Collapse Direction:=wdCollapseEnd
                foundBook = True
                Debug.Print "FindBookH1: >>", "Book found '" & fullBookName & "'", "#" & bookIdx
                paraIndex = bookIdx
                'Stop
                Exit Sub
            Else
                'Debug.Print "FindBookH1: >> Book not found: " & "'" & fullBookName & "'"
            End If
        End If
    Next para
    RestoreCursor savedPos
    Debug.Print "FindBookH1: >> Book not found: " & "'" & fullBookName & "'" & " for '" & bookAbbr & "'"
    MsgBox "FindBookH1: >> Book not found: " & "'" & fullBookName & "'" & " for '" & bookAbbr & "'", vbExclamation, "Bible"
End Sub

Private Function ParseParts(ByVal userInput As String, Optional ByVal delimiter As String = ":") As String()
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
End Function

Private Function GetFullBookName(abbr As String) As String
    Dim bookMap As Object
    Set bookMap = CreateObject("Scripting.Dictionary")
    
    bookMap.Add UCase("Gen"), "Genesis"
    bookMap.Add UCase("Ge"), "Genesis"
    bookMap.Add UCase("Exod"), "Exodus"
    bookMap.Add UCase("Ex"), "Exodus"
    bookMap.Add UCase("Lev"), "Leviticus"
    bookMap.Add UCase("Le"), "Leviticus"
    bookMap.Add UCase("Num"), "Numbers"
    bookMap.Add UCase("Nu"), "Numbers"
    bookMap.Add UCase("Deut"), "Deuteronomy"
    bookMap.Add UCase("De"), "Deuteronomy"
    bookMap.Add UCase("Josh"), "Joshua"
    bookMap.Add UCase("Jos"), "Joshua"
    bookMap.Add UCase("Judg"), "Judges"
    bookMap.Add UCase("Ruth"), "Ruth"
    bookMap.Add UCase("Ru"), "Ruth"
    bookMap.Add UCase("1 Sam"), "1 Samuel"
    bookMap.Add UCase("1 S"), "1 Samuel"
    bookMap.Add UCase("2 Sam"), "2 Samuel"
    bookMap.Add UCase("2 S"), "2 Samuel"
    bookMap.Add UCase("1 Kgs"), "1 Kings"
    bookMap.Add UCase("1 K"), "1 Kings"
    bookMap.Add UCase("2 Kgs"), "2 Kings"
    bookMap.Add UCase("2 K"), "2 Kings"
    bookMap.Add UCase("1 Chr"), "1 Chronicles"
    bookMap.Add UCase("1 Ch"), "1 Chronicles"
    bookMap.Add UCase("2 Chr"), "2 Chronicles"
    bookMap.Add UCase("2 Ch"), "2 Chronicles"
    bookMap.Add UCase("Ezra"), "Ezra"
    bookMap.Add UCase("Ezr"), "Ezra"
    bookMap.Add UCase("Neh"), "Nehemiah"
    bookMap.Add UCase("Ne"), "Nehemiah"
    bookMap.Add UCase("Esth"), "Esther"
    bookMap.Add UCase("Es"), "Esther"
    bookMap.Add UCase("Job"), "Job"
    bookMap.Add UCase("Ps"), "Psalms"
    bookMap.Add UCase("Prov"), "Proverbs"
    bookMap.Add UCase("Pr"), "Proverbs"
    bookMap.Add UCase("Eccl"), "Ecclesiastes"
    bookMap.Add UCase("Ec"), "Ecclesiastes"
    bookMap.Add UCase("Ecc"), "Ecclesiastes"
    bookMap.Add UCase("Song"), "Solomon"
    bookMap.Add UCase("S"), "Solomon"
    bookMap.Add UCase("Isa"), "Isaiah"
    bookMap.Add UCase("Is"), "Isaiah"
    bookMap.Add UCase("I"), "Isaiah"
    bookMap.Add UCase("Jer"), "Jeremiah"
    bookMap.Add UCase("Je"), "Jeremiah"
    bookMap.Add UCase("Lam"), "Lamentations"
    bookMap.Add UCase("La"), "Lamentations"
    bookMap.Add UCase("Ezek"), "Ezekiel"
    bookMap.Add UCase("Eze"), "Ezekiel"
    bookMap.Add UCase("Dan"), "Daniel"
    bookMap.Add UCase("Da"), "Daniel"
    bookMap.Add UCase("Hos"), "Hosea"
    bookMap.Add UCase("Ho"), "Hosea"
    bookMap.Add UCase("Joel"), "Joel"
    bookMap.Add UCase("Joe"), "Joel"
    bookMap.Add UCase("Amos"), "Amos"
    bookMap.Add UCase("Am"), "Amos"
    bookMap.Add UCase("Obad"), "Obadiah"
    bookMap.Add UCase("O"), "Obadiah"
    bookMap.Add UCase("Jonah"), "Jonah"
    bookMap.Add UCase("Jon"), "Jonah"
    bookMap.Add UCase("Mic"), "Micah"
    bookMap.Add UCase("Mi"), "Micah"
    bookMap.Add UCase("Nah"), "Nahum"
    bookMap.Add UCase("Na"), "Nahum"
    bookMap.Add UCase("Hab"), "Habakkuk"
    bookMap.Add UCase("Zeph"), "Zephaniah"
    bookMap.Add UCase("Zep"), "Zephaniah"
    bookMap.Add UCase("Hag"), "Haggai"
    bookMap.Add UCase("Zech"), "Zechariah"
    bookMap.Add UCase("Zec"), "Zechariah"
    bookMap.Add UCase("Mal"), "Malachi"
    bookMap.Add UCase("Matt"), "Matthew"
    bookMap.Add UCase("Mat"), "Matthew"
    bookMap.Add UCase("Mark"), "Mark"
    bookMap.Add UCase("Mar"), "Mark"
    bookMap.Add UCase("Luke"), "Luke"
    bookMap.Add UCase("Lu"), "Luke"
    bookMap.Add UCase("John"), "John"
    bookMap.Add UCase("Joh"), "John"
    bookMap.Add UCase("Acts"), "Acts"
    bookMap.Add UCase("Ac"), "Acts"
    bookMap.Add UCase("Rom"), "Romans"
    bookMap.Add UCase("Ro"), "Romans"
    bookMap.Add UCase("1 Cor"), "1 Corinthians"
    bookMap.Add UCase("1 Co"), "1 Corinthians"
    bookMap.Add UCase("2 Cor"), "2 Corinthians"
    bookMap.Add UCase("2 Co"), "2 Corinthians"
    bookMap.Add UCase("Gal"), "Galatians"
    bookMap.Add UCase("Ga"), "Galatians"
    bookMap.Add UCase("Eph"), "Ephesians"
    bookMap.Add UCase("Ep"), "Ephesians"
    bookMap.Add UCase("Phil"), "Philippians"
    bookMap.Add UCase("Phili"), "Philippians"
    bookMap.Add UCase("Col"), "Colossians"
    bookMap.Add UCase("C"), "Colossians"
    bookMap.Add UCase("1 Thess"), "1 Thessalonians"
    bookMap.Add UCase("1 The"), "1 Thessalonians"
    bookMap.Add UCase("1 Th"), "1 Thessalonians"
    bookMap.Add UCase("2 Thess"), "2 Thessalonians"
    bookMap.Add UCase("2 The"), "1 Thessalonians"
    bookMap.Add UCase("2 Th"), "2 Thessalonians"
    bookMap.Add UCase("1 Tim"), "1 Timothy"
    bookMap.Add UCase("1 Ti"), "1 Timothy"
    bookMap.Add UCase("2 Tim"), "2 Timothy"
    bookMap.Add UCase("2 Ti"), "2 Timothy"
    bookMap.Add UCase("Titus"), "Titus"
    bookMap.Add UCase("T"), "Titus"
    bookMap.Add UCase("Phlm"), "Philemon"
    bookMap.Add UCase("Phile"), "Philemon"
    bookMap.Add UCase("Heb"), "Hebrews"
    bookMap.Add UCase("He"), "Hebrews"
    bookMap.Add UCase("Jas"), "James"
    bookMap.Add UCase("Ja"), "James"
    bookMap.Add UCase("1 Pet"), "1 Peter"
    bookMap.Add UCase("1 P"), "1 Peter"
    bookMap.Add UCase("2 Pet"), "2 Peter"
    bookMap.Add UCase("2 P"), "2 Peter"
    bookMap.Add UCase("1 John"), "1 John"
    bookMap.Add UCase("1 J"), "1 John"
    bookMap.Add UCase("2 John"), "2 John"
    bookMap.Add UCase("2 J"), "2 John"
    bookMap.Add UCase("3 John"), "3 John"
    bookMap.Add UCase("3 J"), "3 John"
    bookMap.Add UCase("Jude"), "Jude"
    bookMap.Add UCase("Rev"), "Revelation"
    bookMap.Add UCase("Re"), "Revelation"
    
    abbr = UCase(Trim(abbr))
    Debug.Print "GetFullBookName: ", "abbr = " & abbr
    If bookMap.Exists(abbr) Then
        GetFullBookName = UCase(bookMap(abbr))
    Else
        GetFullBookName = ""
    End If
End Function

Sub GoToSection()
    'NavigateToNextBookmark()
    Dim bmList As Collection
    Set bmList = GetBookmarkList()
    
    If bmList.count = 0 Then
        MsgBox "No bookmarks found.", vbExclamation
        Exit Sub
    End If

    bookmarkIndex = bookmarkIndex + 1
    If bookmarkIndex > bmList.count Then bookmarkIndex = 1

    bmList.item(bookmarkIndex).range.Select
End Sub

Private Function GetBookmarkList() As Collection
    Dim bmColl As New Collection
    Dim bm As Bookmark

    For Each bm In ActiveDocument.Bookmarks
        bmColl.Add bm
    Next bm

    Set GetBookmarkList = bmColl
End Function

Private Sub GoToH1()
    Dim pattern As String
    Dim para As paragraph
    Dim paraText As String
    Dim matchFound As Boolean

    pattern = InputBox("Enter a Book Name (Heading 1) abbreviation:", "Go To Bible Book")
    If pattern = "" Then Exit Sub ' User canceled
    matchFound = False

    ' Disable UI updates for speed
    Application.ScreenUpdating = False

    For Each para In ActiveDocument.paragraphs
        If para.style = "Heading 1" Then
            paraText = Trim$(para.range.text)
            If paraText Like "*" & UCase(pattern) & "*" Then
                para.range.Select
                ' Move insertion point (cursor) without selecting text
                ActiveDocument.range(para.range.Start, para.range.Start).Select
                matchFound = True
                Exit For
            End If
        End If
    Next para

    Application.ScreenUpdating = True
    Selection.range.Select  ' Re-select current range to restore cursor
    DoEvents  ' Allows UI refresh
    
    If Not matchFound Then
        MsgBox "Book not found! No Heading 1 matches pattern: '" & pattern & "'", vbExclamation, "Bible"
    End If
End Sub

Private Sub NextButton()
    Dim doc As Document
    Dim searchRange As range
    Dim paraEnd As Long
    Dim found As Boolean

    Set doc = ActiveDocument
    found = False

    ' Move start past current paragraph to avoid re-matching
    paraEnd = Selection.paragraphs(1).range.End
    Set searchRange = doc.range(paraEnd, doc.content.End)

    With searchRange.Find
        .ClearFormatting
        .style = doc.Styles("Heading 1")
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .text = ""
        found = .Execute
    End With

    ' If not found, wrap: from beginning to current paragraph start
    If Not found Then
        Set searchRange = doc.range(0, paraEnd)
        With searchRange.Find
            .ClearFormatting
            .style = doc.Styles("Heading 1")
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .text = ""
            found = .Execute
        End With
    End If

    ' If found, move cursor to start of heading
    If found Then
        Selection.SetRange searchRange.Start, searchRange.Start
        ActiveWindow.ScrollIntoView Selection.range, True
    Else
        MsgBox "No Heading 1 found in the document.", vbInformation
    End If
End Sub

Private Function GetExactVerticalScroll() As Double
' Return the scroll percentage rounded to three decimal places
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
End Function

