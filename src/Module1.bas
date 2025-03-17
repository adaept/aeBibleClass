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

Sub ListUsedStylesWithCount()
    Dim doc As Document
    Dim style As style
    Dim usedStyles As Collection
    Dim para As Paragraph
    Dim rng As Range
    Dim charRange As Range
    Dim i As Integer
    Dim styleName As String
    
    Set doc = ActiveDocument
    Set usedStyles = New Collection
    
    ' Loop through all paragraphs to find used paragraph styles
    For Each para In doc.Paragraphs
        styleName = para.style.NameLocal
        On Error Resume Next
        usedStyles.Add styleName, styleName
        On Error GoTo 0
    Next para
    
    ' Loop through all story ranges to find used character styles and font styles
    For Each rng In doc.StoryRanges
        Do While Not rng Is Nothing
            ' Check for character styles
            For Each charRange In rng.Characters
                If charRange.style.Type = wdStyleTypeCharacter Then
                    styleName = charRange.style.NameLocal
                    On Error Resume Next
                    usedStyles.Add styleName, styleName
                    On Error GoTo 0
                End If
                ' Check for font styles
                If charRange.font.name <> "" Then
                    styleName = "Font: " & charRange.font.name
                    On Error Resume Next
                    usedStyles.Add styleName, styleName
                    On Error GoTo 0
                End If
            Next charRange
            Set rng = rng.NextStoryRange
        Loop
    Next rng
    
    ' Print styles to the Immediate Window
    Debug.Print "Styles in Use:"
    For i = 1 To usedStyles.count
        Debug.Print usedStyles(i)
    Next i
    
    ' Print total count of styles
    Debug.Print "Total Styles in Use: " & usedStyles.count
End Sub



