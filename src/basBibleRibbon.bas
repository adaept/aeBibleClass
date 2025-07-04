Attribute VB_Name = "basBibleRibbon"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

' Use Windows API to change cursor
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" ( _
    ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
' Global ribbon object and button state flags
Public ribbonUI As IRibbonUI
Public ribbonIsReady As Boolean
Public btnNextEnabled As Boolean

' Initialize early
Public Sub AutoExec()
    Debug.Print "In AutoExec routine"
    btnNextEnabled = True
End Sub

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

Sub OnGoToVerseSblClick(control As IRibbonControl)
    Call GoToVerseSBL
End Sub

Sub OnHelloWorldButtonClick(control As IRibbonControl)
   MsgBox "Hello, SILAS World!" & vbCrLf & _
                "GetVScroll  = " & GetExactVerticalScroll
End Sub

Sub OnGoToH1ButtonClick(control As IRibbonControl)
    Call GoToH1
End Sub

Public Sub OnNextButtonClick(control As IRibbonControl)
    Call NextButton
End Sub

Sub OnAdaeptAboutClick(control As IRibbonControl)
    MsgBox "Hello, adaept World!" & vbCrLf & _
                "adaeptMsg  = " & adaeptMsg, vbInformation, "About adaept"
End Sub

Function adaeptMsg() As String
    adaeptMsg = """...the truth shall make you free.""" & " John 8:32 (KJV)"
End Function

Sub GoToVerseSBL()
    On Error GoTo ErrHandler
    Application.StatusBar = "Searching for verse..."
    
    Dim userInput As String
    userInput = InputBox("Enter verse (e.g. 1 Sam 1:1):", "Go to Verse (SBL Format)")
    userInput = UCase(userInput)
    If Trim(userInput) = "" Then Exit Sub
    
    Dim bookAbbr As String, chapNum As String, verseNum As String
    Dim parts() As String, subParts() As String
    
    Dim hWaitCursor As Long
    ' Set spinning cursor manually
    hWaitCursor = LoadCursor(0, 32514) ' 32514 = Busy (Hourglass)
    SetCursor hWaitCursor
    Application.ScreenUpdating = False  ' Prevent flickering

    ' Parse the input
    parts = Split(userInput, ":")
    'Debug.Print "UBound(parts) = " & UBound(parts)
    If UBound(parts) = 0 Then   ' Only the ~chapter~ number was provided
        verseNum = 1
        GoTo Chapter
    ElseIf UBound(parts) <> 1 Then
        Application.ScreenUpdating = True   ' Restore normal UI
        SetCursor LoadCursor(0, 32512)      ' Restore default arrow cursor
        MsgBox "Invalid format. Use format like '1 Sam 1:1'", vbExclamation
        Exit Sub
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
        chapNum = 1
    Else
        Dim i As Long
        bookAbbr = ""
        For i = 0 To UBound(subParts) - 1
            bookAbbr = bookAbbr & subParts(i) & " "
            'Debug.Print "b:", bookAbbr
        Next i
        bookAbbr = Trim(bookAbbr)
        'Debug.Print ">", bookAbbr
        chapNum = Trim(subParts(UBound(subParts)))
        'Debug.Print ">>", chapNum
    End If
    
    Dim fullBookName As String
    fullBookName = GetFullBookName(bookAbbr)
    'Debug.Print ">>>", fullBookName
    If fullBookName = "" Then
        Application.ScreenUpdating = True   ' Restore normal UI
        SetCursor LoadCursor(0, 32512)      ' Restore default arrow cursor
        MsgBox "Book not found: " & bookAbbr, vbExclamation
        Exit Sub
    End If

    ' Find the Heading 1 for the book
    Dim theBook As String
    Dim para As paragraph, foundBook As Boolean
    For Each para In ActiveDocument.paragraphs
        If para.style = "Heading 1" Then
            theBook = Trim(para.range.text)
            theBook = UCase(Replace(para.range.text, vbCr, ""))
            'Debug.Print theBook
            If theBook Like "*" & fullBookName & "*" Then
                para.range.Select
                foundBook = True
                'Debug.Print bookAbbr, theBook, fullBookName
                'MsgBox "Book found. Searching for chapter or verse " & chapNum, vbInformation
                Exit For
            End If
        End If
    Next para
    If Not foundBook Then
        Application.ScreenUpdating = True   ' Restore normal UI
        SetCursor LoadCursor(0, 32512)      ' Restore default arrow cursor
        MsgBox "Book heading not found: " & fullBookName, vbExclamation
        Exit Sub
    End If
    
    ' Find the Heading 2 for the chapter or psalm
    Dim theChapter As String
    Dim chapFound As Boolean
    For Each para In ActiveDocument.paragraphs
        'Debug.Print para.range.Start, Selection.range.Start
        If para.range.Start < Selection.range.Start Then GoTo SkipChapter
        If para.style = "Heading 2" Then
            Select Case theBook     ' Books of only one chapter
                Case "OBADIAH", "PHILEMON", "2 JOHN", "3 JOHN", "JUDE"
                    verseNum = chapNum
                    chapNum = 1
                    chapFound = True
                    'MsgBox "The name " & theBook & " is in the list!", vbInformation
                Case Else
                    'MsgBox "The name is NOT in the list!", vbExclamation
            End Select
            If Trim(para.range.text) Like "*Chapter " & chapNum & "*" _
                    Or Trim(para.range.text) Like "*Psalm " & chapNum & "*" Then
                para.range.Select
                chapFound = True
                Exit For
            End If
        End If
SkipChapter:
    Next para
    If Not chapFound Then
        Application.ScreenUpdating = True   ' Restore normal UI
        SetCursor LoadCursor(0, 32512)      ' Restore default arrow cursor
        MsgBox "Chapter not found: " & chapNum, vbExclamation
        Exit Sub
    End If

    ' Limit search range to current chapter
    Dim chapStart As Long, chapEnd As Long
    chapStart = Selection.range.Start
    chapEnd = ActiveDocument.content.End
    For Each para In ActiveDocument.paragraphs
        If para.range.Start > chapStart And para.style = "Heading 2" Then
            chapEnd = para.range.Start
            Exit For
        End If
    Next para

    ' Search for verse number within "Verse marker" style
    Dim r As range
    Dim charCount As Long, found As Boolean
    charCount = chapStart
    Do While charCount < chapEnd
        Set r = ActiveDocument.range(charCount, charCount + 1)
        If r.Characters(1).style = "Verse marker" Then
            Dim verseStr As String, j As Long
            verseStr = ""
            For j = 0 To 2 ' Check up to 3 digits
                If charCount + j >= chapEnd Then Exit For
                If IsNumeric(ActiveDocument.range(charCount + j, charCount + j + 1).text) Then
                    verseStr = verseStr & ActiveDocument.range(charCount + j, charCount + j + 1).text
                Else
                    Exit For
                End If
            Next j
            If verseStr = verseNum Then
                ActiveDocument.range(charCount, charCount + Len(verseStr)).Select
                found = True
                Exit Do
            End If
        End If
        charCount = charCount + 1
        If charCount Mod 1000 = 0 Then DoEvents
    Loop
    If Not found Then
        MsgBox "Verse not found: " & verseNum, vbExclamation
    End If

Cleanup:
    Application.ScreenUpdating = True   ' Restore normal UI
    SetCursor LoadCursor(0, 32512)      ' Restore default arrow cursor
    Application.StatusBar = False
    Exit Sub

ErrHandler:
    MsgBox "Error during verse search: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Function GetFullBookName(abbr As String) As String
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
    bookMap.Add UCase("1 Th"), "1 Thessalonians"
    bookMap.Add UCase("2 Thess"), "2 Thessalonians"
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
    If bookMap.Exists(abbr) Then
        GetFullBookName = bookMap(abbr)
    Else
        GetFullBookName = ""
    End If
End Function

Sub NextButton()
    'GoToNextHeading1Circular()
    Dim doc As Document
    Dim searchRange As range
    Dim selEnd As Long
    Dim found As Boolean

    Set doc = ActiveDocument
    selEnd = Selection.End
    found = False

    ' Search forward: from current position to end
    Set searchRange = doc.range(selEnd, doc.content.End)
    With searchRange.Find
        .ClearFormatting
        .style = doc.Styles("Heading 1")
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .text = ""
        found = .Execute
    End With

    ' If not found, wrap: from beginning to current position
    If Not found Then
        Set searchRange = doc.range(0, selEnd)
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

    ' If found, position cursor at end of heading to prepare for next search
    If found Then
        Selection.SetRange searchRange.End, searchRange.End
        ActiveWindow.ScrollIntoView Selection.range, True
    Else
        MsgBox "No Heading 1 found in the document.", vbInformation
    End If
End Sub

Sub GoToH1()
    Dim pattern As String
    Dim para As paragraph
    Dim paraText As String
    Dim matchFound As Boolean

    pattern = InputBox("Enter a Heading 1 pattern to match (use * and ? wildcards):", "Go To Bible Book")
    If pattern = "" Then Exit Sub ' User canceled
    matchFound = False

    ' Disable UI updates for speed
    Application.ScreenUpdating = False

    For Each para In ActiveDocument.paragraphs
        If para.style = "Heading 1" Then
            paraText = Trim$(para.range.text)
            If paraText Like pattern Then
                para.range.Select
                ' Move insertion point (cursor) without selecting text
                ActiveDocument.range(para.range.Start, para.range.Start).Select
                matchFound = True
                Exit For
            End If
        End If
    Next para

    Application.ScreenUpdating = True

    If Not matchFound Then
        MsgBox "No Heading 1 matches pattern: " & pattern, vbExclamation
    End If
End Sub

Function GetExactVerticalScroll() As Double
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


