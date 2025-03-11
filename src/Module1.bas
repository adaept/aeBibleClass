Attribute VB_Name = "Module1"
Sub ViewCodeDetails()
    Dim selectedText As String
    Dim msg As String
    Dim i As Integer

    ' Get the selected text
    selectedText = Selection.text

    ' Initialize the message string
    msg = "Code details for the selected text:" & vbCrLf & vbCrLf

    ' Loop through each character in the selected text
    For i = 1 To Len(selectedText)
        msg = msg & "Character " & i & ": " & Mid(selectedText, i, 1) & " (ASCII: " & Asc(Mid(selectedText, i, 1)) & ")" & vbCrLf
    Next i

    ' Display the code details in a message box
    MsgBox msg
End Sub

Sub PrintFontProperties()
    Dim sel As Selection
    Set sel = Selection
    With sel.font
        Debug.Print "Name: " & .name
        Debug.Print "Size: " & .Size
        Debug.Print "Bold: " & .Bold
        Debug.Print "Italic: " & .Italic
        Debug.Print "Underline: " & .Underline
        Debug.Print "Color: " & .Color
        Debug.Print "StrikeThrough: " & .StrikeThrough
        Debug.Print "DoubleStrikeThrough: " & .DoubleStrikeThrough
        Debug.Print "Subscript: " & .Subscript
        Debug.Print "Superscript: " & .Superscript
        Debug.Print "Shadow: " & .Shadow
        Debug.Print "Outline: " & .Outline
        Debug.Print "Emboss: " & .Emboss
        Debug.Print "Engrave: " & .Engrave
        Debug.Print "AllCaps: " & .AllCaps
        Debug.Print "Hidden: " & .Hidden
        Debug.Print "SmallCaps: " & .SmallCaps
        Debug.Print "Kerning: " & .Kerning
        Debug.Print "Spacing: " & .Spacing
        Debug.Print "Scaling: " & .Scaling
        Debug.Print "Position: " & .Position
        Debug.Print "Ligatures: " & .Ligatures
        Debug.Print "NumberForm: " & .NumberForm
        Debug.Print "NumberSpacing: " & .NumberSpacing
        Debug.Print "StylisticSet: " & .StylisticSet
        Debug.Print "ContextualAlternates: " & .ContextualAlternates
    End With
End Sub

Sub PrintBibleBookHeadingsVerseNumbers()
' Find Heading 1, then all Heading 2 until the next Heading 1, and print the heading names to the console.
' Update to write out verse numbers
    
    Dim headingLabel As String
    Dim para As Paragraph
    Dim foundHeading1, foundHeading2 As Boolean
    
    ' Prompt the user to enter the Heading 1 label
    headingLabel = InputBox("Enter the Heading 1 label:")
    headingLabel = UCase(headingLabel)
    
    foundHeading1 = False
    foundHeading2 = False
    
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
                Debug.Print
                ' Get the text of the Heading 2 without the extra carriage return
                Debug.Print Replace(para.Range.text, vbCr, "")
                foundHeading2 = True
            ElseIf foundHeading2 Then
                ' Update here and add routine to get numbers from character style
                Debug.Print "Num",
            End If
        End If
    Next para
    
    ' Display a message if no headings are found
    If Not foundHeading1 Then
        MsgBox "No headings found with the specified label.", vbExclamation
    End If
End Sub


