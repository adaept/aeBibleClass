Attribute VB_Name = "basTESTaeBibleDOCVARIABLE"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Private lastFoundLocation As range

Sub FindNextHeading1OnVisiblePage()
    Dim doc As Document
    Dim pageNum As Integer
    Dim para As paragraph
    Dim paraPageNum As Integer
    Dim textFound As Boolean
    Dim startRange As range
    Dim headingText As String

    ' Set the active document
    Set doc = ActiveDocument

    ' Prompt the user for the visible page number to start the search
    pageNum = InputBox("Enter the visible page number to start searching for Heading 1:", "Page Number")
    If pageNum <= 0 Then
        MsgBox "Invalid page number entered. Exiting process.", vbCritical, "Error"
        Exit Sub
    End If

    ' Check if we have a previously found location to continue from
    If lastFoundLocation Is Nothing Then
        ' Start at the beginning of the specified page
        Set startRange = doc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, count:=pageNum)
    Else
        ' Continue searching from the last found location
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
            If paraPageNum >= pageNum Then
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

                    Exit Sub ' Exit after finding the first Heading 1
                End If
            End If
        End If
    Next para

    ' Handle the case where no Heading 1 is found
    If Not textFound Then
        MsgBox "No Heading 1 found starting from page " & pageNum & ".", vbExclamation, "Search Complete"
        Set lastFoundLocation = Nothing ' Reset the tracking if nothing is found
    End If
End Sub

