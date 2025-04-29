Attribute VB_Name = "basTESTaeBibleDOCVARIABLE"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Private lastFoundLocation As range
Const wdHeaderStory As Integer = 6
Const wdFooterStory As Integer = 7
Const wdFootnoteStory As Integer = 4
Const wdEndnoteStory As Integer = 5

Function FindNextHeading1OnVisiblePage(bookPage As Integer, textH1 As String) As Boolean
    Dim doc As Document
    Dim para As paragraph
    Dim paraPageNum As Integer
    Dim textFound As Boolean
    Dim startRange As range
    Dim headingText As String

    ' Set the active document
    Set doc = ActiveDocument

    ' Check if we have a previously found location to continue from
    If lastFoundLocation Is Nothing Then
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
    ' Old Testament
    ActiveDocument.Variables("Gen").value = 20
    ActiveDocument.Variables("Exod").value = 58
    ActiveDocument.Variables("Lev").value = 87
    ActiveDocument.Variables("Num").value = 109
    ActiveDocument.Variables("Deut").value = 140
    ActiveDocument.Variables("Josh").value = 166
    ActiveDocument.Variables("Judg").value = 187
    ActiveDocument.Variables("Ruth").value = 202
    ActiveDocument.Variables("1Sam").value = 206
    'MsgBox "DOCVARIABLE values set successfully!"
    Debug.Print "DOCVARIABLE values set successfully!"
End Sub

Sub ListMyDocVariables()
    Dim v As Variant
    For Each v In ActiveDocument.Variables
        Debug.Print v.name & ": " & v.value
    Next
End Sub

Sub TestPageNumbers()
    Set lastFoundLocation = Nothing
    VerifyBookNameFromDocVariable "Gen", "Genesis"
    Debug.Print ">>lastFoundLocation = " & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Exod", "Exodus"
    Debug.Print ">>lastFoundLocation = " & Replace(lastFoundLocation, vbCr, "")
    VerifyBookNameFromDocVariable "Lev", "Leviticus"
    Debug.Print ">>lastFoundLocation = " & Replace(lastFoundLocation, vbCr, "")
    Debug.Print "Done!!!"
End Sub
