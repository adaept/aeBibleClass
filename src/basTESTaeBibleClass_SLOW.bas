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
        .Style = styleName
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

    Dim para As Paragraph
    Dim headingText As String
    Dim pageNumber As Long
    Dim docPosition As Long
    Dim count As Integer
    
    count = 0
    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        ' Check if the paragraph style is Heading 1
        If para.Style = ActiveDocument.Styles(wdStyleHeading1) Then
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
    Dim para As Paragraph
    Dim foundHeading1 As Boolean
    
    ' Prompt the user to enter the Heading 1 label
    headingLabel = InputBox("Enter the Heading 1 label:")
    headingLabel = UCase(headingLabel)
    
    foundHeading1 = False
    
    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        If para.Style = ActiveDocument.Styles(wdStyleHeading1) Then
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
            If para.Style = ActiveDocument.Styles(wdStyleHeading2) Then
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
    Dim para As Paragraph
    Dim paraText As String
    Dim foundHeading1, foundHeading2 As Boolean
    Dim char As String
    Dim asciiValue As Integer
    Dim hexValue As String
   
    ' Prompt the user to enter the Heading 1 label
    headingLabel = InputBox("Enter the Heading 1 label:")
    headingLabel = UCase(headingLabel)
    
    foundHeading1 = False
    foundHeading2 = False
    
    ' Loop through all paragraphs in the document
10    For Each para In ActiveDocument.Paragraphs
20        paraText = para.Range.text
        ' Remove formatting characters
30        paraText = Replace(paraText, vbCr, "") ' Paragraph mark
40        paraText = Replace(paraText, vbTab, "") ' Tab character
50        paraText = Replace(paraText, "^b", "") ' Section break
60        paraText = Replace(paraText, "^m", "") ' Continuous section break
          
70        If Len(paraText) = 0 Then
71            hexValue = "00" ' Hex value for an empty paragraph
72            Debug.Print "> Paragraph is empty. Hex value: " & hexValue
73        ElseIf Len(paraText) < 3 Then
74            Debug.Print "> Len(paraText) = " & Len(paraText)
75            char = Mid(paraText, 1, 1)
76            asciiValue = Asc(char)
77            hexValue = Hex(asciiValue)
78            Debug.Print "1> Character: " & char & " ASCII value: " & asciiValue & " Hex value: " & hexValue
79            ', Asc(Mid(paraText, 2, 1))
80        End If

110        If para.Style = ActiveDocument.Styles(wdStyleHeading1) Then         ' Process paragraph
            ' Check if the Heading 1 matches the input label
120            If paraText = headingLabel Then
130                foundHeading1 = True
140                foundHeading2 = False
150            ElseIf foundHeading1 Then
                ' Stop when the next Heading 1 is found
160                Stop
170                Exit For
180            End If
190        End If
        
        ' If Heading 1 is found, start processing
200        If foundHeading1 Then
210            If para.Style = ActiveDocument.Styles(wdStyleHeading2) Then
220                Debug.Print
                ' Get the text of the Heading 2 without the extra carriage return
230                Debug.Print Replace(para.Range.text, vbCr, "")
240                foundHeading2 = True
250            ElseIf foundHeading2 And foundHeading1 Then
                ' Get numbers from character style
                'ExtractNumbersFromParagraph para, "Verse marker"
260                ExtractNumbersFromParagraph2 para, "cvmarker"
270            End If
280        End If
290        DoEvents ' Allow Word to process other events
300    Next para

    ' Display a message if no headings are found
310    If Not foundHeading1 Then
320        MsgBox "No headings found with the specified label.", vbExclamation
330    End If

ErrorHandler:
340    MsgBox "Err = " & Err.Number & " Erl = " & Erl & " An error occurred: " & Err.Description, vbCritical
    ' Optionally close the document or perform other cleanup
    'ThisDocument.Close SaveChanges:=wdDoNotSaveChanges
350    End
End Sub

Private Sub ExtractNumbersFromParagraph(para As Paragraph, styleName As String)
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
    Dim result As String
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
        .Style = styleName
    End With
    
    ' Find all matches in the paragraph
    Do While rng.Find.Execute
        ' Check if the match is formatted with the specified character style
        If rng.Style = styleName Then
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
        result = Join(arr, ", ")
        Debug.Print result
    End If
End Sub

Private Sub ExtractNumbersFromParagraph2(para As Paragraph, styleName As String)
' The `rng.Find` method is used to search for ranges with the specified character style within the paragraph.
' A regex object is used to find numbers within the styled ranges.
' The numbers are collected and printed as a comma-separated list.
    
    Dim rng As Range
    Dim foundNumbers As Collection
    Dim num As String
    Dim result As String
    Dim arr() As String
    Dim i As Integer
    
    DoEvents ' Allows the system to process other events
    Set foundNumbers = New Collection
    
    ' Set the range to the paragraph
    Set rng = para.Range
    
    ' Initialize the find object
    With rng.Find
        .ClearFormatting
        .Style = styleName
        .text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
    End With
    
    ' Find all ranges with the specified style
    Do While rng.Find.Execute
        'Debug.Print "Found styled range: " & rng.text, rng.Style.NameLocal
        ' Check if the range contains numbers
        If rng.Style = styleName Then
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
                num = Trim(match.Value)
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
        result = Join(arr, ", ")
        Debug.Print result
    End If
End Sub

