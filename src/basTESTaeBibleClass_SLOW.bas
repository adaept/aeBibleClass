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
            Debug.Print count & ": " & "Heading: " & headingText & " | Page: " & pageNumber & " | Position: " & docPosition
        End If
    Next para
End Sub

