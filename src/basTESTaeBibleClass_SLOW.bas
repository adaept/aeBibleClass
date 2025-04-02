Attribute VB_Name = "basTESTaeBibleClass_SLOW"
Option Explicit
Option Compare Text
Option Private Module

Public Sub FindAnyNumberWithStyleAndPrintNextCharASCII()
' Interactive search of 31,102 Bible verses, at 1000 per run,
' to find any spaces after "Verse marker" style.
' It takes 2 minutes per run of one thousand.
' Found 33 in Copy (32).docx
'
    Dim searchText As String
    Dim styleName As String
    Dim found As Boolean
    Dim firstFound As Boolean
    Dim count As Integer
    Dim nextChar As String

    searchText = "[0-9]{1,}" ' Pattern to find any number (one or more digits)
    styleName = "Verse marker" ' Replace with your specific character style name
    firstFound = False
    count = 0

    ' Set the search parameters
    With Selection.Find
        .text = searchText
        .style = styleName
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
        'Debug.Print "Found number: " & Selection.text

        ' Move to the next character and print its ASCII value
        Selection.Collapse Direction:=wdCollapseEnd
        Selection.MoveRight Unit:=wdCharacter, count:=1
        nextChar = Selection.text

        If Len(nextChar) > 0 Then
            If Asc(nextChar) = 32 Then
                Debug.Print "Next character: " & nextChar & " (ASCII: " & Asc(nextChar) & ")"
                Stop
            End If
        Else
            Debug.Print "No next character found."
        End If

        ' Move back to the end of the found number to continue search
        Selection.MoveLeft Unit:=wdCharacter, count:=1

        firstFound = True
        count = count + 1
        If count >= 1000 Then
            Exit Do
        End If
    Loop
    
    ' Notify the user in the Immediate Window if no numbers were found
    If Not firstFound Then
        Debug.Print "No numbers with the specified style found."
    Else
        Debug.Print "Total numbers found: " & count
    End If
End Sub

Public Sub PrintBibleHeading1Info()
' This will print the count, heading text, page number, and document position of each Heading 1 in your document to the Immediate Window
' (press `Ctrl + G` to view the Immediate Window if it's not already visible).

    Dim para As paragraph
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
            headingText = para.Range.text
            pageNumber = para.Range.Information(wdActiveEndPageNumber)
            docPosition = para.Range.Start
            
            ' Print the heading text, page number, and document position to the console
            Debug.Print count & ": " & "Heading: " & Replace(headingText, vbCr, "") & " | Page: " & pageNumber & " | Position: " & docPosition
        End If
    Next para
End Sub

Public Sub PrintBibleBookHeadings()
' Find Heading 1, then all Heading 2 until the next Heading 1, and print the heading names to the console.
    
    Dim headingLabel As String
    Dim para As paragraph
    Dim foundHeading1 As Boolean
    
    ' Prompt the user to enter the Heading 1 label
    headingLabel = InputBox("Enter the Heading 1 label:")
    headingLabel = UCase(headingLabel)
    
    foundHeading1 = False
    
    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        If para.style = ActiveDocument.Styles(wdStyleHeading1) Then
            ' Check if the Heading 1 matches the input label
            If para.Range.text = headingLabel & vbCr Then
                ' Get the text of the Heading 1 without the extra carriage return
                Debug.Print Replace(para.Range.text, vbCr, "")
                foundHeading1 = True
            ElseIf foundHeading1 Then
                ' Stop when the next Heading 1 is found
                Exit For
            End If
        End If
        
        ' If Heading 1 is found, start processing
        If foundHeading1 Then
            If para.style = ActiveDocument.Styles(wdStyleHeading2) Then
                ' Get the text of the Heading 2 without the extra carriage return
                Debug.Print Replace(para.Range.text, vbCr, "")
            End If
        End If
    Next para
    
    ' Display a message if no headings are found
    If Not foundHeading1 Then
        MsgBox "No headings found with the specified label.", vbExclamation
    End If
End Sub

Public Sub PrintBibleBookHeadingsVerseNumbers()
' Find Heading 1, then all Heading 2 until the next Heading 1, and print the heading names to the console.
' Updated to write out verse numbers.
' Before processing each paragraph, check if it contains a continuous page break and handle it accordingly.
    
    On Error GoTo ErrorHandler

    Dim headingLabel As String
    Dim para As paragraph
    Dim paraText As String
    Dim foundHeading1 As Boolean
    Dim char As String
    Dim asciiValue As Integer
    Dim hexValue As String
   
    ' Prompt the user to enter the Heading 1 label
    headingLabel = InputBox("Enter the Heading 1 label:")
    headingLabel = UCase(headingLabel)
    
    foundHeading1 = False
    
    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        paraText = para.Range.text
        ' Remove formatting characters
        paraText = Replace(paraText, vbCr, "") ' Paragraph mark
        paraText = Replace(paraText, vbTab, "") ' Tab character
        paraText = Replace(paraText, "^b", "") ' Section break
        paraText = Replace(paraText, "^m", "") ' Continuous section break
          
        If Len(paraText) = 0 Then
            'hexValue = "00" ' Hex value for an empty paragraph
            'Debug.Print "> Paragraph is empty. Hex value: " & hexValue
        ElseIf Len(paraText) < 3 Then
            Debug.Print "> Len(paraText) = " & Len(paraText)
            char = Mid(paraText, 1, 1)
            asciiValue = Asc(char)
            hexValue = Hex(asciiValue)
            Debug.Print "1> Character: " & char & " ASCII value: " & asciiValue & " Hex value: " & hexValue
            If Len(paraText) = 2 Then
                char = Mid(paraText, 2, 1)
                asciiValue = Asc(char)
                hexValue = Hex(asciiValue)
                Debug.Print "2> Character: " & char & " ASCII value: " & asciiValue & " Hex value: " & hexValue
            End If
        End If

        If para.style = ActiveDocument.Styles(wdStyleHeading1) Then         ' Process paragraph
            ' Check if the Heading 1 matches the input label
            If paraText = headingLabel Then
                Debug.Print para.Range.text
                foundHeading1 = True
            ElseIf foundHeading1 Then
                ' Stop when the next Heading 1 is found
                Debug.Print para.Range.text
                Stop
                Exit For
            End If
        End If
        
        ' If Heading 1 is found, start processing
        If foundHeading1 Then
            If para.style = ActiveDocument.Styles(wdStyleHeading2) Then
                Debug.Print
                ' Get the text of the Heading 2 without the extra carriage return
                Debug.Print Replace(para.Range.text, vbCr, "")
                ' Get numbers from character style
                'ExtractNumbersFromParagraph para, "Verse marker"
                ExtractNumbersFromParagraph2 para, "cvmarker"
            End If
        End If
        DoEvents ' Allow Word to process other events
    Next para

    ' Display a message if no headings are found
    If Not foundHeading1 Then
        MsgBox "No headings found with the specified label.", vbExclamation
    End If

ErrorHandler:
    MsgBox "Err = " & Err.Number & " Erl = " & Erl & " An error occurred: " & Err.Description, vbCritical
    ' Optionally close the document or perform other cleanup
    'ThisDocument.Close SaveChanges:=wdDoNotSaveChanges
    End
End Sub

Private Sub ExtractNumbersFromParagraph(para As paragraph, styleName As String)
' The regex pattern `[0-9]{1,}` is used to match numbers of any length,
' then check if each match has the specified character style and collect the numbers.
' To ensure the style information is preserved when calling the routine from another subroutine,
' we need to pass the paragraph and the style name as parameters.
' The `Selection.Find` method is used to search for numbers in the specified character style within the paragraph.
' The `MatchWildcards` property is set to `True` to enable regex-like searching.
' The routine loops through all matches and collects the numbers formatted with the specified character style.

    Dim rng As Range
    Dim foundNumbers As Collection
    Dim num As String
    Dim Result As String
    Dim arr() As String
    Dim i As Integer
    
    Set foundNumbers = New Collection
    
    ' Set the range to the paragraph
    Set rng = para.Range
    
    ' Initialize the find object
    With rng.Find
        .ClearFormatting
        .text = "[0-9]{1,}"
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchWildcards = True
        .style = styleName
    End With
    
    ' Find all matches in the paragraph
    Do While rng.Find.Execute
        ' Check if the match is formatted with the specified character style
        If rng.style = styleName Then
            num = Trim(rng.text)
            ' Add the number to the collection
            foundNumbers.Add num
        End If
        ' Move the range to the next character to continue the search
        rng.Start = rng.End + 1
        rng.End = para.Range.End
    Loop
    
    ' Convert the collection to a comma-separated string
    If foundNumbers.count > 0 Then
        ReDim arr(1 To foundNumbers.count)
        For i = 1 To foundNumbers.count
            arr(i) = foundNumbers(i)
        Next i
        Result = Join(arr, ", ")
        Debug.Print Result
    End If
End Sub

Private Sub ExtractNumbersFromParagraph2(para As paragraph, styleName As String)
' The `rng.Find` method is used to search for ranges with the specified character style within the paragraph.
' A regex object is used to find numbers within the styled ranges.
' The numbers are collected and printed as a comma-separated list.
    
    Dim rng As Range
    Dim foundNumbers As Collection
    Dim num As String
    Dim Result As String
    Dim arr() As String
    Dim i As Integer
    
    DoEvents ' Allows the system to process other events
    Set foundNumbers = New Collection
    
    ' Set the range to the paragraph
    Set rng = para.Range
    
    ' Initialize the find object
    With rng.Find
        .ClearFormatting
        .style = styleName
        .text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
    End With
    
    ' Find all ranges with the specified style
    Do While rng.Find.Execute
        'Debug.Print "Found styled range: " & rng.text, rng.Style.NameLocal
        ' Check if the range contains numbers
        If rng.style = styleName Then
            ' Create a regex object to find numbers within the styled range
            Dim regex As Object
            Dim matches As Object
            Dim match As Variant
            
            Set regex = CreateObject("VBScript.RegExp")
            regex.pattern = "[0-9]{1,}" ' Pattern to match numbers
            regex.Global = True
            
            ' Find all matches in the styled range text
            Set matches = regex.Execute(rng.text)
            
            ' Loop through each match
            For Each match In matches
                num = Trim(match.value)
                ' Add the number to the collection
                foundNumbers.Add num
            Next match
        End If
        ' Move the range to the next character to continue the search
        rng.Start = rng.End + 1
        rng.End = para.Range.End
    Loop
    
    ' Convert the collection to a comma-separated string
    If foundNumbers.count > 0 Then
        ReDim arr(1 To foundNumbers.count)
        For i = 1 To foundNumbers.count
            arr(i) = foundNumbers(i)
        Next i
        Result = Join(arr, ", ")
        Debug.Print Result
    End If
End Sub

Private Sub ListAscii12Characters()
    Dim rng As Range
    Dim count As Long
    Dim positions As String
    Dim startPos As Long
    
    ' Set the range to the entire document
    Set rng = ActiveDocument.Content
    
    ' Initialize the count and positions
    count = 0
    positions = ""
    
    ' Find all ASCII 12 characters and record their positions
    With rng.Find
        .text = Chr(12) ' Chr(12) represents the ASCII 12 character
        .Replacement.text = ""
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
            positions = positions & "Position " & count & ": " & startPos & vbCrLf
            rng.Collapse wdCollapseEnd
        Loop
    End With
    
    ' Display the positions of ASCII 12 characters
    If count > 0 Then
        Debug.Print "There are " & count & " ASCII 12 characters in the document at the following positions:" & vbCrLf & positions
    Else
        MsgBox "No ASCII 12 characters found in the document.", vbInformation, "ASCII 12 Characters"
    End If
End Sub

Private Sub ListAndReviewAscii12Characters()
    Dim rng As Range
    Dim count As Long
    Dim startPos As Long
    Dim response As VbMsgBoxResult
    
    ' Set the range to the entire document
    Set rng = ActiveDocument.Content
    
    ' Initialize the count
    count = 0
    
    ' Find all ASCII 12 characters and record their positions
    With rng.Find
        .text = Chr(12) ' Chr(12) represents the ASCII 12 character
        .Replacement.text = ""
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
                Exit Sub
            End If
        Loop
    End With
    
    ' Display a message if no ASCII 12 characters are found
    If count = 0 Then
        MsgBox "No ASCII 12 characters found in the document.", vbInformation, "ASCII 12 Characters"
    End If
End Sub

Sub CountParagraphsTypes()
' Slow running routine ~10+ minutes

    Dim doc As Document
    Dim para As paragraph
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
        Kill debugFile
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
        If Len(para.Range.text) = 1 And para.Range.text = vbCr Then
            emptyParagraphs = emptyParagraphs + 1
        End If
        
        ' Check for different types of breaks using Find method
        With para.Range.Find
            .ClearFormatting
            .text = "^m"
            If .Execute Then
                textWrappingBreakParagraphs = textWrappingBreakParagraphs + 1
                textWrappingBreakIndices = textWrappingBreakIndices & paraIndex & ", "
            End If
            .text = "^b"
            If .Execute Then
                columnBreakParagraphs = columnBreakParagraphs + 1
                columnBreakIndices = columnBreakIndices & paraIndex & ", "
            End If
        End With
        
        ' Check for different types of section breaks
        If para.Range.Sections.count > 0 Then
            Select Case para.Range.Sections(1).PageSetup.SectionStart
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
End Sub

Sub AppendToFile(filePath As String, text As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Append As fileNum
    Print #fileNum, text
    Close fileNum
End Sub
 
Sub CountNumberDashNumberInFootnotes()
' This is slow
    Dim doc As Document
    Dim footnote As footnote
    Dim rng As Range
    Dim pattern As String
    Dim foundCount As Long
    Dim debugInfo As String
    
    ' Initialize variables
    Set doc = ActiveDocument
    foundCount = 0
    pattern = "[0-9]{1,}-[0-9]{1,}" ' Pattern for number-dash-number
    debugInfo = "Footnotes with Number-Dash-Number Patterns:" & vbCrLf
    
    ' Loop through each footnote in the document
    For Each footnote In doc.Footnotes
        Set rng = footnote.Range
        rng.Find.ClearFormatting
        
        ' Configure the Find object
        With rng.Find
            .text = pattern
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchWildcards = True
        End With
        
        ' Perform the search within the footnote
        Do While rng.Find.Execute
            foundCount = foundCount + 1
            debugInfo = debugInfo & "Footnote " & footnote.Index & ": '" & rng.text & "' at position " & rng.Start & vbCrLf
            rng.Collapse wdCollapseEnd
        Loop
    Next footnote
    
    ' Display the count of number-dash-number patterns in footnotes
    MsgBox "Number of number-dash-number patterns in footnotes: " & foundCount
    ' Display debug information
    MsgBox debugInfo
    
    ' Print the result to the console
    Debug.Print "Number of number-dash-number patterns in footnotes: " & foundCount
    Debug.Print debugInfo
End Sub
