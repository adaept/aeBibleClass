Attribute VB_Name = "XbasTESTaeBibleClass_SLOW"
'==============================================================================
' XbasTESTaeBibleClass_SLOW - Slow Diagnostic Tests (DEFERRED)
' ----------------------------------------------------------------------------
' X-prefix convention: excluded from the normal test run. Contains tests that
' iterate all 31,102 Bible verses and are too slow for routine execution.
' Run manually from the Immediate Window when deep diagnosis is needed.
'==============================================================================
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Public Sub FindAnyNumberWithStyleAndPrintNextCharASCII()
' Interactive search of 31,102 Bible verses, at 1000 per run,
' to find any spaces after "Verse marker" style.
' It takes 2 minutes per run of one thousand.
' Found 33 in Copy (32).docx
'
    On Error GoTo PROC_ERR
    Dim searchText As String
    Dim StyleName As String
    Dim found As Boolean
    Dim firstFound As Boolean
    Dim count As Integer
    Dim nextChar As String

    searchText = "[0-9]{1,}" ' Pattern to find any number (one or more digits)
    StyleName = "Verse marker" ' Replace with your specific character style name
    firstFound = False
    count = 0

    ' Set the search parameters
    With Selection.Find
        .Text = searchText
        .style = StyleName
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

    ' Execute the search
    Do While Selection.Find.Execute
        'Debug.Print "Found number: " & Selection.Text

        ' Move to the next character and print its ASCII value
        Selection.Collapse Direction:=wdCollapseEnd
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        nextChar = Selection.Text

        If Len(nextChar) > 0 Then
            If Asc(nextChar) = 32 Then
                Debug.Print "Next character: " & nextChar & " (ASCII: " & Asc(nextChar) & ")"
                Stop
            End If
        Else
            Debug.Print "No next character found."
        End If

        ' Move back to the end of the found number to continue search
        Selection.MoveLeft Unit:=wdCharacter, Count:=1

        firstFound = True
        count = count + 1
        If count >= 1000 Then  ' Safety limit: process in batches of 1000 to keep runtime manageable
            Exit Do
        End If
    Loop
    
    ' Notify the user in the Immediate Window if no numbers were found
    If Not firstFound Then
        Debug.Print "No numbers with the specified style found."
    Else
        Debug.Print "Total numbers found: " & count
    End If

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FindAnyNumberWithStyleAndPrintNextCharASCII of Module XbasTESTaeBibleClass_SLOW"
    Resume PROC_EXIT
End Sub

Public Sub PrintBibleHeading1Info()
' This will print the count, heading text, page number, and document position of each Heading 1 in your document to the Immediate Window
' (press `Ctrl + G` to view the Immediate Window if it's not already visible).

    On Error GoTo PROC_ERR
    Dim para As Word.Paragraph
    Dim headingText As String
    Dim pageNumber As Long
    Dim docPosition As Long
    Dim count As Integer

    count = 0
    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        ' Check if the paragraph style is Heading 1
        If para.style = ActiveDocument.Styles(wdStyleHeading1) Then
            count = count + 1
            headingText = para.Range.Text
            pageNumber = para.Range.Information(wdActiveEndPageNumber)
            docPosition = para.Range.Start
            
            ' Print the heading text, page number, and document position to the console
            Debug.Print count & ": " & "Heading: " & Replace(headingText, vbCr, "") & " | Page: " & pageNumber & " | Position: " & docPosition
        End If
    Next para

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure PrintBibleHeading1Info of Module XbasTESTaeBibleClass_SLOW"
    Resume PROC_EXIT
End Sub

Public Sub PrintBibleBookHeadings()
' Find Heading 1, then all Heading 2 until the next Heading 1, and print the heading names to the console.
    On Error GoTo PROC_ERR
    Dim headingLabel As String
    Dim para As Word.Paragraph
    Dim foundHeading1 As Boolean

    ' Prompt the user to enter the Heading 1 label
    headingLabel = InputBox("Enter the Heading 1 label:")
    headingLabel = UCase(headingLabel)

    foundHeading1 = False

    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        If para.style = ActiveDocument.Styles(wdStyleHeading1) Then
            ' Check if the Heading 1 matches the input label
            If para.Range.Text = headingLabel & vbCr Then
                ' Get the text of the Heading 1 without the extra carriage return
                Debug.Print Replace(para.Range.Text, vbCr, "")
                foundHeading1 = True
            ElseIf foundHeading1 Then
                ' Stop when the next Heading 1 is found
                GoTo PROC_EXIT
            End If
        End If

        ' If Heading 1 is found, start processing
        If foundHeading1 Then
            If para.style = ActiveDocument.Styles(wdStyleHeading2) Then
                ' Get the text of the Heading 2 without the extra carriage return
                Debug.Print Replace(para.Range.Text, vbCr, "")
            End If
        End If
    Next para

    ' Display a message if no headings are found
    If Not foundHeading1 Then
        MsgBox "No headings found with the specified label.", vbExclamation
    End If

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure PrintBibleBookHeadings of Module XbasTESTaeBibleClass_SLOW"
    Resume PROC_EXIT
End Sub

Sub ListAndReviewAscii12Characters()
' Ascii 12 is Form Feed, FF, Page Break
    On Error GoTo PROC_ERR
    Dim rng As Word.Range
    Dim count As Long
    Dim startPos As Long
    Dim response As VbMsgBoxResult

    ' Set the range to the entire document
    Set rng = ActiveDocument.Content

    ' Initialize the count
    count = 0

    ' Find all ASCII 12 characters and record their positions
    With rng.Find
        .Text = Chr(12) ' Chr(12) represents the ASCII 12 character
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        Do While .Execute
            count = count + 1
            startPos = rng.Start
            Debug.Print "Position " & count & ": " & startPos
            rng.Collapse wdCollapseEnd

            ' Navigate to the position in the document
            ActiveDocument.Range(startPos, startPos).Select

            ' Ask if the user wants to continue
            response = MsgBox("ASCII 12 character found at position " & startPos & ". Do you want to continue?", vbYesNo + vbQuestion, "Review ASCII 12 Characters")
            If response = vbNo Then
                GoTo PROC_EXIT
            End If
        Loop
    End With

    ' Display a message if no ASCII 12 characters are found
    If count = 0 Then
        MsgBox "No ASCII 12 characters found in the document.", vbInformation, "ASCII 12 Characters"
    End If

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ListAndReviewAscii12Characters of Module XbasTESTaeBibleClass_SLOW"
    Resume PROC_EXIT
End Sub

Sub CountParagraphsTypes()
' Slow running routine ~10+ minutes
    On Error GoTo PROC_ERR
    Dim doc As Document
    Dim para As Word.Paragraph
    Dim totalParagraphs As Long
    Dim emptyParagraphs As Long
    Dim pageBreakParagraphs As Long
    Dim columnBreakParagraphs As Long
    Dim textWrappingBreakParagraphs As Long
    Dim nextPageSectionBreakParagraphs As Long
    Dim continuousSectionBreakParagraphs As Long
    Dim evenPageSectionBreakParagraphs As Long
    Dim oddPageSectionBreakParagraphs As Long
    Dim paraIndex As Long
    Dim debugFile As String
    Dim fileNum As Integer
    Dim continueProcessing As VbMsgBoxResult
    Dim pageBreakIndices As String
    Dim columnBreakIndices As String
    Dim textWrappingBreakIndices As String
    Dim nextPageSectionBreakIndices As String
    Dim continuousSectionBreakIndices As String
    Dim evenPageSectionBreakIndices As String
    Dim oddPageSectionBreakIndices As String
    
    ' Initialize counts and indices
    totalParagraphs = 0
    emptyParagraphs = 0
    pageBreakParagraphs = 0
    columnBreakParagraphs = 0
    textWrappingBreakParagraphs = 0
    nextPageSectionBreakParagraphs = 0
    continuousSectionBreakParagraphs = 0
    evenPageSectionBreakParagraphs = 0
    oddPageSectionBreakParagraphs = 0
    paraIndex = 0
    pageBreakIndices = ""
    columnBreakIndices = ""
    textWrappingBreakIndices = ""
    nextPageSectionBreakIndices = ""
    continuousSectionBreakIndices = ""
    evenPageSectionBreakIndices = ""
    oddPageSectionBreakIndices = ""
    
    ' Set the document to the active document
    Set doc = ActiveDocument
    
    ' Set the debug file path to the current document directory
    debugFile = doc.Path & "\ParagraphsCountDebugTestFile.txt"
    
    ' Delete the old debug file if it exists
    If Dir(debugFile) <> "" Then
        On Error Resume Next
        Kill debugFile
        If Err.Number <> 0 Then
            Debug.Print "Could not delete debug file: " & Err.Description & " - " & debugFile
            Err.Clear
        End If
        On Error GoTo 0
        On Error GoTo PROC_ERR
    End If

    ' Open the debug file for writing
    fileNum = FreeFile
    Open debugFile For Output As fileNum
    Close fileNum
    
    ' Loop through each paragraph in the document
    For Each para In doc.Paragraphs
        paraIndex = paraIndex + 1
        totalParagraphs = totalParagraphs + 1
        
        ' Check if the paragraph is empty
        If Len(para.Range.Text) = 1 And para.Range.Text = vbCr Then
            emptyParagraphs = emptyParagraphs + 1
        End If
        
        ' Check for different types of breaks using Find method
        With para.Range.Find
            .ClearFormatting
            .Text = "^m"
            If .Execute Then
                textWrappingBreakParagraphs = textWrappingBreakParagraphs + 1
                textWrappingBreakIndices = textWrappingBreakIndices & paraIndex & ", "
            End If
            .Text = "^b"
            If .Execute Then
                columnBreakParagraphs = columnBreakParagraphs + 1
                columnBreakIndices = columnBreakIndices & paraIndex & ", "
            End If
            .Text = "^p"
            If .Execute Then
                pageBreakParagraphs = pageBreakParagraphs + 1
                pageBreakIndices = pageBreakIndices & paraIndex & ", "
            End If
        End With
        
        ' Check for different types of section breaks
        If para.Range.Sections.Count > 0 Then
            Select Case para.Range.Sections(1).PageSetup.sectionStart
                Case wdSectionNewPage
                    nextPageSectionBreakParagraphs = nextPageSectionBreakParagraphs + 1
                    nextPageSectionBreakIndices = nextPageSectionBreakIndices & paraIndex & ", "
                Case wdSectionContinuous
                    continuousSectionBreakParagraphs = continuousSectionBreakParagraphs + 1
                    continuousSectionBreakIndices = continuousSectionBreakIndices & paraIndex & ", "
                Case wdSectionEvenPage
                    evenPageSectionBreakParagraphs = evenPageSectionBreakParagraphs + 1
                    evenPageSectionBreakIndices = evenPageSectionBreakIndices & paraIndex & ", "
                Case wdSectionOddPage
                    oddPageSectionBreakParagraphs = oddPageSectionBreakParagraphs + 1
                    oddPageSectionBreakIndices = oddPageSectionBreakIndices & paraIndex & ", "
            End Select
        End If
        
'        ' Prompt user to continue processing after every 100 paragraphs
'        If paraIndex Mod 100 = 0 Then
'            continueProcessing = MsgBox("Continue processing?", vbYesNo + vbQuestion, "Continue?")
'            If continueProcessing = vbNo Then
'                Exit For
'            End If
'        End If
        
        ' Allow the system to process other events
        DoEvents
    Next para
    
    ' Remove trailing commas and spaces
    If Len(pageBreakIndices) > 0 Then pageBreakIndices = Left(pageBreakIndices, Len(pageBreakIndices) - 2)
    If Len(columnBreakIndices) > 0 Then columnBreakIndices = Left(columnBreakIndices, Len(columnBreakIndices) - 2)
    If Len(textWrappingBreakIndices) > 0 Then textWrappingBreakIndices = Left(textWrappingBreakIndices, Len(textWrappingBreakIndices) - 2)
    If Len(nextPageSectionBreakIndices) > 0 Then nextPageSectionBreakIndices = Left(nextPageSectionBreakIndices, Len(nextPageSectionBreakIndices) - 2)
    If Len(continuousSectionBreakIndices) > 0 Then continuousSectionBreakIndices = Left(continuousSectionBreakIndices, Len(continuousSectionBreakIndices) - 2)
    If Len(evenPageSectionBreakIndices) > 0 Then evenPageSectionBreakIndices = Left(evenPageSectionBreakIndices, Len(evenPageSectionBreakIndices) - 2)
    If Len(oddPageSectionBreakIndices) > 0 Then oddPageSectionBreakIndices = Left(oddPageSectionBreakIndices, Len(oddPageSectionBreakIndices) - 2)
    
    ' Append the final results to the debug file
    AppendToFile debugFile, "Paragraphs with Page Break: " & pageBreakIndices
    AppendToFile debugFile, "Paragraphs with Column Break: " & columnBreakIndices
    AppendToFile debugFile, "Paragraphs with Text Wrapping Break: " & textWrappingBreakIndices
    AppendToFile debugFile, "Paragraphs with Section Break (Next Page): " & nextPageSectionBreakIndices
    AppendToFile debugFile, "Paragraphs with Section Break (Continuous): " & continuousSectionBreakIndices
    AppendToFile debugFile, "Paragraphs with Section Break (Even Page): " & evenPageSectionBreakIndices
    AppendToFile debugFile, "Paragraphs with Section Break (Odd Page): " & oddPageSectionBreakIndices
    
    ' Print the counts to the console (Immediate Window)
    Debug.Print "Total Paragraphs: " & totalParagraphs
    Debug.Print "Empty Paragraphs: " & emptyParagraphs
    Debug.Print "Paragraphs with Page Break: " & pageBreakParagraphs
    Debug.Print "Paragraphs with Column Break: " & columnBreakParagraphs
    Debug.Print "Paragraphs with Text Wrapping Break: " & textWrappingBreakParagraphs
    Debug.Print "Paragraphs with Section Break (Next Page): " & nextPageSectionBreakParagraphs
    Debug.Print "Paragraphs with Section Break (Continuous): " & continuousSectionBreakParagraphs
    Debug.Print "Paragraphs with Section Break (Even Page): " & evenPageSectionBreakParagraphs
    Debug.Print "Paragraphs with Section Break (Odd Page): " & oddPageSectionBreakParagraphs

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure CountParagraphsTypes of Module XbasTESTaeBibleClass_SLOW"
    Resume PROC_EXIT
End Sub

Sub AppendToFile(filePath As String, text As String)
    On Error GoTo PROC_ERR
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Append As fileNum
    Print #fileNum, text
    Close fileNum

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure AppendToFile of Module XbasTESTaeBibleClass_SLOW"
    Resume PROC_EXIT
End Sub

Sub FindNextVerseMarkerSequence()
' Search for char style "Chapter Verse marker" followed by char style "Verse marker"
' with space of "Normal" style before and after.
' ~200 secs and there should be no matches.
    On Error GoTo PROC_ERR
    Dim doc As Document
    Dim searchRange As Word.Range
    Dim chapterRng As Word.Range, nextRng As Word.Range
    Dim found As Boolean
    Dim progressCount As Long
    Dim tStart As Single

    Application.ScreenUpdating = False
    Application.StatusBar = "Starting search..."

    Set doc = ActiveDocument
    found = False
    tStart = Timer

    Set searchRange = doc.Range(0, doc.Content.End)

    ' Begin search for Chapter Verse marker
    With searchRange.Find
        .ClearFormatting
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .style = "Chapter Verse marker"
        .Execute
    End With

    Do While searchRange.Find.found
        Set chapterRng = searchRange.Duplicate

        ' Attempt to get the next character styled as Verse marker
        If chapterRng.End + 1 <= doc.Content.End Then
            Set nextRng = doc.Range(Start:=chapterRng.End, End:=chapterRng.End + 1)
        Else
            searchRange.Start = chapterRng.End
            searchRange.End = doc.Content.End
            searchRange.Find.Execute
            GoTo ContinueLoop
        End If

        If nextRng.Characters.Count = 1 Then
            If nextRng.style = "Verse marker" Then
                Dim beforeChar As Word.Range, afterChar As Word.Range

                ' Before chapter
                If chapterRng.Start > 0 Then
                    Set beforeChar = doc.Range(Start:=chapterRng.Start - 1, End:=chapterRng.Start)
                Else
                    GoTo ContinueLoop
                End If

                ' After verse
                If nextRng.End + 1 <= doc.Content.End Then
                    Set afterChar = doc.Range(Start:=nextRng.End, End:=nextRng.End + 1)
                Else
                    GoTo ContinueLoop
                End If

                ' Safety checks
                If beforeChar.Characters.Count < 1 Or afterChar.Characters.Count < 1 Then
                    Debug.Print "Invalid character count at " & chapterRng.Start
                    chapterRng.Select
                    MsgBox "Cannot access one of the surrounding characters. Stopping for inspection.", vbExclamation
                    GoTo PROC_EXIT
                End If

                ' Check styles and spaces
                If Trim(beforeChar.Text) = "" And beforeChar.style = "Normal" Then
                    If Trim(afterChar.Text) = "" And afterChar.style = "Normal" Then
                        ' Found match
                        chapterRng.Start = beforeChar.Start
                        nextRng.End = afterChar.End
                        doc.Range(chapterRng.Start, nextRng.End).Select
                        MsgBox "Match found at position " & chapterRng.Start, vbInformation
                        found = True
                        Stop
                        Exit Do
                    End If
                End If
            End If
        End If

ContinueLoop:
        ' Continue search
        searchRange.Start = chapterRng.End
        searchRange.End = doc.Content.End
        searchRange.Find.Execute

        progressCount = progressCount + 1
        If progressCount Mod 100 = 0 Then
            Application.StatusBar = "Searching... character " & searchRange.Start
            DoEvents
        End If
    Loop

    If Not found Then
        MsgBox "No more matches found.", vbInformation
    End If

    Debug.Print "Elapsed: " & Format(Timer - tStart, "0.00") & " sec"

PROC_EXIT:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub
PROC_ERR:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FindNextVerseMarkerSequence of Module XbasTESTaeBibleClass_SLOW"
    Resume PROC_EXIT
End Sub

