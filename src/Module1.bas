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

Sub ListCustomXMLParts()
    Dim xmlPart As CustomXMLPart
    Dim i As Integer
    
    i = 1
    For Each xmlPart In ThisDocument.CustomXMLParts
        Debug.Print "Custom XML Part " & i & ": " & xmlPart.XML
        i = i + 1
    Next xmlPart
End Sub

Sub RemoveDuplicateCustomXMLParts()
    Dim xmlPart As CustomXMLPart
    Dim xmlParts As CustomXMLParts
    Dim essentialParts As Collection
    Dim duplicateParts As Collection
    Dim partName As String
    Dim i As Integer, j As Integer
    
    Set xmlParts = ActiveDocument.CustomXMLParts
    Set essentialParts = New Collection
    Set duplicateParts = New Collection
    
    ' Identify essential and duplicate parts
    For i = 1 To xmlParts.count
        partName = xmlParts(i).NamespaceURI
        If Not IsPartInCollection(essentialParts, partName) Then
            essentialParts.Add xmlParts(i), partName
        Else
            duplicateParts.Add xmlParts(i), partName
        End If
    Next i
    
    ' Remove duplicate parts
    For j = 1 To duplicateParts.count
        duplicateParts(j).Delete
    Next j
    
    ' Print names of essential and duplicate parts
    Debug.Print "Essential CustomXML Parts:"
    For i = 1 To essentialParts.count
        Debug.Print essentialParts(i).NamespaceURI
    Next i
    
    Debug.Print "Duplicate CustomXML Parts:"
    For j = 1 To duplicateParts.count
        Debug.Print duplicateParts(j).NamespaceURI
    Next j
End Sub

Function IsPartInCollection(col As Collection, partName As String) As Boolean
    Dim i As Integer
    IsPartInCollection = False
    For i = 1 To col.count
        If col(i).NamespaceURI = partName Then
            IsPartInCollection = True
            Exit Function
        End If
    Next i
End Function

Sub DeleteCustomUIXML()
    Dim xmlPart As CustomXMLPart
    Dim xmlParts As CustomXMLParts
    Dim i As Integer
    
    Set xmlParts = ActiveDocument.CustomXMLParts
    
    ' Loop through all CustomXMLParts to find and delete the customUI parts
    For i = xmlParts.count To 1 Step -1
        Set xmlPart = xmlParts(i)
        If xmlPart.NamespaceURI = "http://schemas.microsoft.com/office/2006/01/customui" Or _
           xmlPart.NamespaceURI = "http://schemas.microsoft.com/office/2009/07/customui" Then
            xmlPart.Delete
        End If
    Next i
    
    MsgBox "CustomUI XML parts deleted successfully!"
End Sub

Sub PrintBibleBook()
    Dim heading1Name As String
    Dim para As paragraph
    Dim startProcessing As Boolean
    Dim heading1Found As Boolean
    Dim heading2Found As Boolean
    
    ' Prompt user to enter the name of Heading 1
    heading1Name = InputBox("Enter the name of Heading 1:")
    heading1Name = UCase(heading1Name)
    
    startProcessing = False
    heading1Found = False
    heading2Found = False
    
    ' Loop through all paragraphs in the document
    For Each para In ActiveDocument.Paragraphs
        If para.style = "Heading 1" Then
            If InStr(para.Range.text, heading1Name) > 0 Then
                Debug.Print "Heading 1: " & para.Range.text
                startProcessing = True
                heading1Found = True
            Else
                startProcessing = False
                heading1Found = False
            End If
        End If
        
        If startProcessing Then
            If para.style = "Heading 2" Then
                Debug.Print "Heading 2: " & para.Range.text
                heading2Found = True
            ElseIf heading2Found Then
                Debug.Print para.Range.text
            End If
        End If
        
        If heading1Found And para.style = "Heading 1" And InStr(para.Range.text, heading1Name) = 0 Then
            Exit For
        End If
    Next para
End Sub

Function IsParagraphEmpty(paragraph As Range) As Boolean
    ' Check if the paragraph is empty
    If Len(paragraph.text) = 1 And paragraph.text = vbCr Then
        IsParagraphEmpty = True
    Else
        IsParagraphEmpty = False
    End If
End Function

Sub GoToParagraphIndex()
    Dim para As paragraph
    Dim paraIndex As Integer
    Dim targetIndex As Integer
    
    ' Prompt user to enter the index of the paragraph
    targetIndex = InputBox("Enter the index of the paragraph you want to go to:")
    
    ' Validate the entered index
    If targetIndex > 0 And targetIndex <= ActiveDocument.Paragraphs.count Then
        paraIndex = 1
        For Each para In ActiveDocument.Paragraphs
            If paraIndex = targetIndex Then
                para.Range.Select
                Exit Sub
            End If
            paraIndex = paraIndex + 1
        Next para
    Else
        MsgBox "Invalid index entered. Please enter a valid index between 1 and " & ActiveDocument.Paragraphs.count & "."
    End If
End Sub


