Attribute VB_Name = "basTESTaeBibleTools"
Option Explicit
Option Compare Text
Option Private Module

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

Function GetColorNameFromHex(hexColor As String) As String
    Dim colorName As String
    
    ' Convert hex to uppercase for consistency
    hexColor = UCase(hexColor)
    
    ' Determine the color name based on the hex value
    Select Case hexColor
        Case "#FF0000"
            colorName = "Red"
        Case "#00FF00"
            colorName = "Green"
        Case "#0000FF"
            colorName = "Blue"
        Case "#FFD700"
            colorName = "Gold"
        Case "#FFA500"
            colorName = "Orange"
        Case "#663399"
            colorName = "Purple"
        Case "#FFFFFF"
            colorName = "White"
        Case "#000000"
            colorName = "Black"
        Case "#800000"
            colorName = "Dark Red"
        Case "#808080"
            colorName = "Gray"
        Case Else
            colorName = "Unknown Color"
    End Select
    
    ' Return the color name
    GetColorNameFromHex = colorName
End Function

Sub ListAndCountFontColors()
    Dim rng As Range
    Dim colorDict As Object
    Dim colorKey As Variant
    Dim colorCount As Long
    Dim r As Long, g As Long, b As Long
    
    ' Create a dictionary to store color counts
    Set colorDict = CreateObject("Scripting.Dictionary")
    
    ' Loop through each word in the document
    For Each rng In ActiveDocument.Words
        ' Get the RGB values of the font color
        r = (rng.font.color And &HFF)
        g = (rng.font.color \ &H100 And &HFF)
        b = (rng.font.color \ &H10000 And &HFF)
        
        ' Create a key for the color in hex format
        colorKey = Right("0" & Hex(r), 2) & Right("0" & Hex(g), 2) & Right("0" & Hex(b), 2)
        
        ' Count the color occurrences
        If colorDict.Exists(colorKey) Then
            colorDict(colorKey) = colorDict(colorKey) + 1
        Else
            colorDict.Add colorKey, 1
        End If
    Next rng
    
    ' Print the results to the console
    For Each colorKey In colorDict.Keys
        colorCount = colorDict(colorKey)
        r = CLng("&H" & Left(colorKey, 2))
        g = CLng("&H" & Mid(colorKey, 3, 2))
        b = CLng("&H" & Right(colorKey, 2))
        
        Debug.Print "Color: RGB(" & r & ", " & g & ", " & b & ") - Hex: #" & colorKey & " - Count: " & colorCount & " - " & GetColorNameFromHex("#" & colorKey)
    Next colorKey
End Sub

