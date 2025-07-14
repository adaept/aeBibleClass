Attribute VB_Name = "basTESTaeBibleDOCVARIABLE"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Public lastFoundLocation As range
Private Const wdHeaderStory As Integer = 6
Private Const wdFooterStory As Integer = 7
Private Const wdFootnoteStory As Integer = 4
Private Const wdEndnoteStory As Integer = 5

Function FindNextHeading1OnVisiblePage(bookPage As Integer, textH1 As String, Optional ByVal restartVal As Variant) As Boolean
    Dim doc As Document
    Dim para As paragraph
    Dim paraPageNum As Integer
    Dim textFound As Boolean
    Dim startRange As range
    Dim headingText As String
    Dim counter As Long

    ' Set the active document
    Set doc = ActiveDocument

    If Not IsMissing(restartVal) Then
        ' Set the range to start from the specified location
        Debug.Print "Restarting from location " & restartVal
        Set startRange = doc.range(Start:=restartVal, End:=doc.content.End)
    ElseIf lastFoundLocation Is Nothing Then
        ' Check if we have a previously found location to continue from
        ' Start at the beginning of the specified page
        Debug.Print ">bookPage = " & bookPage, "textH1 = " & textH1
        Set startRange = doc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, count:=bookPage)
    Else
        ' Continue searching from the last found location
        Debug.Print ">lastFoundLocation = " & Replace(lastFoundLocation, vbCr, "")
        Set startRange = lastFoundLocation
    End If

    ' Initialize the flag for text found
    textFound = False
    ' Initialize an empty string to store Heading 1 text
    headingText = ""
    ' Initialize the counter
    counter = 0

    ' Iterate through paragraphs starting from the specified range
    For Each para In doc.paragraphs
        ' Get the visible page number of the current paragraph
        paraPageNum = para.range.Information(wdActiveEndAdjustedPageNumber)

        ' Check if the paragraph is on the specified or subsequent pages
        If para.range.Start >= startRange.Start Then
            'Debug.Print "paraPageNum = " & paraPageNum
            If paraPageNum >= bookPage Then
                ' Verify if the paragraph is styled as Heading 1
                If para.style = "Heading 1" Then
                    textFound = True
                    headingText = Trim(para.range.text) ' Get the text of the Heading 1
                    ' Clean up the text by removing all newline, carriage return, and control characters
                    headingText = Replace(headingText, vbCrLf, "")
                    headingText = Replace(headingText, vbLf, "")
                    headingText = Replace(headingText, vbCr, "")
                    headingText = Replace(headingText, Chr(11), "") ' Remove vertical tab if present
                    headingText = Replace(headingText, Chr(12), "") ' Remove form feed if present
                    headingText = Trim(headingText) ' Finally, trim spaces
                    
                    ' Remember the location for the next search
                    Set lastFoundLocation = para.range
                    'MsgBox "Found Heading 1 on visible page " & paraPageNum & " at location: " & para.range.Start, _
                        vbInformation, "Heading 1 Found"
                    Debug.Print "Found Heading 1 " & headingText & " on visible page " & paraPageNum & " at location: " & para.range.Start
                    
                    FindNextHeading1OnVisiblePage = textFound
                    Exit Function ' Exit after finding the first Heading 1
                End If
            End If
        End If
        ' Increment the counter
        counter = counter + 1

        ' Call DoEvents intermittently (e.g., every 100 iterations)
        If counter Mod 100 = 0 Then
            DoEvents
        End If
    Next para

    ' Handle the case where no Heading 1 is found
    If Not textFound Then
        MsgBox "No Heading 1 found starting from page " & bookPage & ".", vbExclamation, "Search Complete"
        Set lastFoundLocation = Nothing ' Reset the tracking if nothing is found
        FindNextHeading1OnVisiblePage = textFound
    End If
    
End Function

Sub VerifyBookNameFromDocVariable(docVar As String, theTextOfH1 As String)
    Dim doc As Document
    Dim bookNum As Integer
    Dim textFoundHere As Boolean

    ' Set the active document
    Set doc = ActiveDocument

    ' Get the value of the DOCVARIABLE
    On Error Resume Next
    bookNum = doc.Variables(docVar).value
    On Error GoTo 0
    Debug.Print "BookNum = " & bookNum

    ' Check if the DOCVARIABLE is valid
    If bookNum <= 0 Then
        MsgBox docVar & " DOCVARIABLE is not set or has an invalid value.", vbExclamation, "Error"
        bookNum = InputBox("Enter the correct page number for '" & docVar & "':", "Correct Page Number")
        If bookNum > 0 Then
            doc.Variables(docVar).value = bookNum
            doc.Fields.Update
        Else
            'MsgBox "No valid page number entered. Exiting process.", vbCritical, "Process Canceled"
            Debug.Print "No valid page number entered. Exiting process."
            Exit Sub
        End If
    End If

    textFoundHere = False
    
    ' Define the text to search for in theTextOfH1
    ' Using "GENESIS" as theTextOfH1 for the Heading 1 test
    textFoundHere = FindNextHeading1OnVisiblePage(bookNum, UCase(theTextOfH1))
    
    ' Evaluate the search result
    If textFoundHere Then
        'MsgBox "Success: The text Heading 1 '" & UCase(theTextOfH1) & "' was found on page " & bookNum & ".", vbInformation, "Verification Complete"
        Debug.Print "Success: The text Heading 1 '" & UCase(theTextOfH1) & "' was found on page " & bookNum & "."
    Else
        ' If text not found, throw an error and prompt for correction
        Set lastFoundLocation = Nothing ' Reset the tracking if nothing is found
        Err.Raise 1000, "VerifyBookNameFromDocVariable", "The text Heading 1 '" & UCase(theTextOfH1) & "' was NOT found on page " & bookNum & "."
    End If

    Exit Sub

ErrorHandler:
    ' Handle the error and prompt for a new page number
    MsgBox Err.Description, vbExclamation, "Error"
    bookNum = InputBox("The text '" & UCase(theTextOfH1) & "' was not found. Enter the correct page number:", "Correct Page Number")
    If bookNum > 0 Then
        doc.Variables(docVar).value = bookNum
        doc.Fields.Update
        Resume
    Else
        MsgBox "No valid page number entered. Exiting process.", vbCritical, "Process Canceled"
    End If
End Sub

Sub FindDocVariableByName(docVar As String)
    Dim doc As Document
    Dim variableExists As Boolean
    Dim variableValue As String

    ' Set the active document
    Set doc = ActiveDocument

    ' Initialize the flag for existence and variable value
    variableExists = False
    variableValue = ""

    ' Search for the DOCVARIABLE
    On Error Resume Next
    variableValue = doc.Variables(docVar).value
    If Err.Number = 0 Then
        variableExists = True
    End If
    On Error GoTo 0

    ' Display the result
    If variableExists Then
        'MsgBox "DOCVARIABLE '" & docVar & "' exists with the value: " & vbCrLf & "'" & variableValue & "'", _
               vbInformation, "Variable Found"
        Debug.Print "DOCVARIABLE '" & docVar & "' exists with the value: " & "'" & variableValue & "'"
    Else
        'MsgBox "DOCVARIABLE '" & docVar & "' does not exist in this document.", vbExclamation, "Variable Not Found"
        Debug.Print "DOCVARIABLE '" & docVar & "' does not exist in this document."
    End If
End Sub

Sub FindDocVariableEverywhere()
    Dim doc As Document
    Dim variableName As String
    Dim variableFound As Boolean
    Dim field As field
    Dim shape As shape
    Dim section As section
    Dim note As footnote
    Dim endNote As endNote

    ' Set the active document
    Set doc = ActiveDocument

    ' Prompt the user to enter the name of the DOCVARIABLE to locate
    variableName = InputBox("Enter the name of the DOCVARIABLE to locate:", "Search DOCVARIABLE")

    ' Check if the variable name is valid
    If Trim(variableName) = "" Then
        MsgBox "Invalid variable name. Please enter a valid name.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Initialize the flag for existence
    variableFound = False

    ' First: Search for DOCVARIABLE in shapes (including nested shapes)
    For Each shape In doc.Shapes
        variableFound = SearchShapeForVariable(shape, variableName)
        If variableFound Then Exit Sub ' Exit once found
    Next shape

    ' Second: Search for DOCVARIABLE in the main document body
    For Each field In doc.Fields
        If field.Type = wdFieldDocVariable Then
            If InStr(1, field.Code.text, variableName, vbTextCompare) > 0 Then
                ' Select the field and stop at its location
                field.Select
                variableFound = True
                MsgBox "DOCVARIABLE '" & variableName & "' found in the document body.", vbInformation, "Variable Found"
                Exit Sub
            End If
        End If
    Next field

    ' Third: Search for DOCVARIABLE in headers and footers
    For Each section In doc.Sections
        ' Check headers
        For Each field In section.Headers(wdHeaderFooterPrimary).range.Fields
            If field.Type = wdFieldDocVariable Then
                If InStr(1, field.Code.text, variableName, vbTextCompare) > 0 Then
                    field.Select
                    variableFound = True
                    MsgBox "DOCVARIABLE '" & variableName & "' found in a header.", vbInformation, "Variable Found"
                    Exit Sub
                End If
            End If
        Next field

        ' Check footers
        For Each field In section.Footers(wdHeaderFooterPrimary).range.Fields
            If field.Type = wdFieldDocVariable Then
                If InStr(1, field.Code.text, variableName, vbTextCompare) > 0 Then
                    field.Select
                    variableFound = True
                    MsgBox "DOCVARIABLE '" & variableName & "' found in a footer.", vbInformation, "Variable Found"
                    Exit Sub
                End If
            End If
        Next field
    Next section

    ' Fourth: Search for DOCVARIABLE in footnotes
    For Each note In doc.Footnotes
        For Each field In note.range.Fields
            If field.Type = wdFieldDocVariable Then
                If InStr(1, field.Code.text, variableName, vbTextCompare) > 0 Then
                    field.Select
                    variableFound = True
                    MsgBox "DOCVARIABLE '" & variableName & "' found in a footnote.", vbInformation, "Variable Found"
                    Exit Sub
                End If
            End If
        Next field
    Next note

    ' Fifth: Search for DOCVARIABLE in endnotes
    For Each endNote In doc.Endnotes
        For Each field In endNote.range.Fields
            If field.Type = wdFieldDocVariable Then
                If InStr(1, field.Code.text, variableName, vbTextCompare) > 0 Then
                    field.Select
                    variableFound = True
                    MsgBox "DOCVARIABLE '" & variableName & "' found in an endnote.", vbInformation, "Variable Found"
                    Exit Sub
                End If
            End If
        Next field
    Next endNote

    ' Handle the case where the DOCVARIABLE is not found
    If Not variableFound Then
        MsgBox "DOCVARIABLE '" & variableName & "' does not exist in this document.", vbExclamation, "Variable Not Found"
    End If
End Sub

Function SearchShapeForVariable(shape As shape, variableName As String) As Boolean
    Dim childShape As shape
    Dim field As field
    Dim textFrameRange As range

    ' Initialize return value
    SearchShapeForVariable = False

    ' Check if the shape has text and search its fields
    If shape.TextFrame.HasText Then
        Set textFrameRange = shape.TextFrame.textRange
        For Each field In textFrameRange.Fields
            If field.Type = wdFieldDocVariable Then
                If InStr(1, field.Code.text, variableName, vbTextCompare) > 0 Then
                    ' Select the shape and notify the user
                    shape.Select
                    MsgBox "DOCVARIABLE '" & variableName & "' found in a nested shape.", vbInformation, "Variable Found"
                    SearchShapeForVariable = True
                    Exit Function
                End If
            End If
        Next field
    End If

    ' Check for nested shapes within this shape
    If shape.Type = msoGroup Then
        For Each childShape In shape.GroupItems
            SearchShapeForVariable = SearchShapeForVariable(childShape, variableName)
            If SearchShapeForVariable Then Exit Function ' Exit once found
        Next childShape
    End If
End Function

Sub SetDocVariables()
    'GoTo NewTestament
    
    ' Old Testament
    ActiveDocument.Variables("Gen").value = 20
    ActiveDocument.Variables("Exod").value = 57
    ActiveDocument.Variables("Lev").value = 87
    ActiveDocument.Variables("Num").value = 109
    ActiveDocument.Variables("Deut").value = 140
    ActiveDocument.Variables("Josh").value = 166
    ActiveDocument.Variables("Judg").value = 184
    ActiveDocument.Variables("Ruth").value = 202
    ActiveDocument.Variables("1Sam").value = 206
    ActiveDocument.Variables("2Sam").value = 229
    ActiveDocument.Variables("1Kgs").value = 248
    ActiveDocument.Variables("2Kgs").value = 271
    ActiveDocument.Variables("1Chr").value = 293
    ActiveDocument.Variables("2Chr").value = 313
    ActiveDocument.Variables("Ezra").value = 338
    ActiveDocument.Variables("Neh").value = 346
    ActiveDocument.Variables("Esth").value = 357
    ActiveDocument.Variables("Job").value = 364
    ActiveDocument.Variables("Ps").value = 383
    ActiveDocument.Variables("Prov").value = 427
    ActiveDocument.Variables("Eccl").value = 443
    ActiveDocument.Variables("Song").value = 449
    ActiveDocument.Variables("Isa").value = 454
    ActiveDocument.Variables("Jer").value = 491
    ActiveDocument.Variables("Lam").value = 531
    ActiveDocument.Variables("Ezek").value = 537
    ActiveDocument.Variables("Dan").value = 574
    ActiveDocument.Variables("Hos").value = 586
    ActiveDocument.Variables("Joel").value = 593
    ActiveDocument.Variables("Amos").value = 596
    ActiveDocument.Variables("Obad").value = 601
    ActiveDocument.Variables("Jonah").value = 603
    ActiveDocument.Variables("Mic").value = 606
    ActiveDocument.Variables("Nah").value = 611
    ActiveDocument.Variables("Hab").value = 614
    ActiveDocument.Variables("Zeph").value = 617
    ActiveDocument.Variables("Hag").value = 620
    ActiveDocument.Variables("Zech").value = 622
    ActiveDocument.Variables("Mal").value = 629

NewTestament:
    ActiveDocument.Variables("Matt").value = 634
    ActiveDocument.Variables("Mark").value = 661
    ActiveDocument.Variables("Luke").value = 678
    ActiveDocument.Variables("John").value = 706
    ActiveDocument.Variables("Acts").value = 728
    ActiveDocument.Variables("Rom").value = 753
    ActiveDocument.Variables("1Cor").value = 764
    ActiveDocument.Variables("2Cor").value = 775
    ActiveDocument.Variables("Gal").value = 783
    ActiveDocument.Variables("Eph").value = 788
    ActiveDocument.Variables("Phil").value = 793
    ActiveDocument.Variables("Col").value = 797
    ActiveDocument.Variables("1Thess").value = 801
    ActiveDocument.Variables("2Thess").value = 805
    ActiveDocument.Variables("1Tim").value = 807
    ActiveDocument.Variables("2Tim").value = 811
    ActiveDocument.Variables("Tit").value = 814
    ActiveDocument.Variables("Phlm").value = 817
    ActiveDocument.Variables("Heb").value = 819
    ActiveDocument.Variables("Jas").value = 828
    ActiveDocument.Variables("1Pet").value = 832
    ActiveDocument.Variables("2Pet").value = 836
    ActiveDocument.Variables("1John").value = 839
    ActiveDocument.Variables("2John").value = 843
    ActiveDocument.Variables("3John").value = 845
    ActiveDocument.Variables("Jude").value = 847
    ActiveDocument.Variables("Rev").value = 849
    'MsgBox "DOCVARIABLE values set successfully!"
    Debug.Print "DOCVARIABLE values set successfully!"
End Sub

Sub ListMyDocVariables()
    Dim v As Variant
    For Each v In ActiveDocument.Variables
        Debug.Print v.name & ": " & v.value
    Next
End Sub

Sub DeleteDocVariable()
    Dim varName As String
    varName = "xxxJos" ' Change this to the name of your variable

    On Error Resume Next
    ActiveDocument.Variables(varName).Delete
    On Error GoTo 0

    ActiveDocument.Fields.Update ' Refresh any DOCVARIABLE fields
End Sub

Sub TestPageNumbers()
    GoTo NewTestament
    
    ' Old Testament
    VerifyBookNameFromDocVariable "Gen", "Genesis"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Exod", "Exodus"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Lev", "Leviticus"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Num", "Numbers"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Deut", "Deuteronomy"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Josh", "Joshua"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Judg", "Judges"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Ruth", "Ruth"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "1Sam", "1 Samuel"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "2Sam", "2 Samuel"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "1Kgs", "1 Kings"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "2Kgs", "2 Kings"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "1Chr", "1 Chronicles"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "2Chr", "2 Chronicles"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Ezra", "Ezra"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Neh", "Nehemiah"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Esth", "Esther"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Job", "Job"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Ps", "Psalms"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Prov", "Proverbs"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Eccl", "Ecclesiastes"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Song", "Solomon"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Isa", "Isaiah"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Jer", "Jeremiah"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Lam", "Lamentations"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Ezek", "Ezekiel"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Dan", "Daniel"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Hos", "Hoseah"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Joel", "Joel"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Amos", "Amos"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Obad", "Obadiah"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Jonah", "Jonah"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Mic", "Micah"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Nah", "Nahum"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Hab", "Habakkuk"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Zeph", "Zephaniah"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Hag", "Haggai"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Zech", "Zechariah"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Mal", "Malachi"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    Debug.Print "Done Old Testament !!!"

NewTestament:
    Dim rng As range
    Set rng = ActiveDocument.GoTo(What:=1, Which:=1, name:="630")
    rng.Select

    VerifyBookNameFromDocVariable "Matt", "Matthew"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Mark", "Mark"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Luke", "Luke"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "John", "John"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Acts", "Acts"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Rom", "Romans"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "1Cor", "1 Corinthians"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "2Cor", "2 Corinthians"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Gal", "Galations"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Eph", "Ephesians"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Phil", "Philippians"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Col", "Colossians"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "1Thess", "1 Thessalonians"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "2Thess", "2 Thessalonians"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "1Tim", "1 Timothy"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "2Tim", "2 Timothy"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Tit", "Titus"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Phlm", "Philemon"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Heb", "Hebrews"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Jas", "James"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "1Pet", "1 Peter"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "2Pet", "2 Peter"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "1John", "1 John"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "2John", "2 John"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "3John", "3 John"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Jude", "Jude"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Rev", "Revelation"
    Debug.Print ">>" & Replace(lastFoundLocation, vbCr, "")
    Debug.Print "Done New Testament !!!"
    Debug.Print "Done!!!"
End Sub
