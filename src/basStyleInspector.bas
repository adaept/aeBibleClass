Attribute VB_Name = "basStyleInspector"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

'==============================================================================
' DumpStyleProperties
'==============================================================================
' Prints a named style's properties in a form that can be pasted into a
' Define<Style> routine. Output goes to the Immediate window; optionally also
' writes to rpt\style_<name>.txt for diffing / style-guide generation.
'
' Usage:
'   DumpStyleProperties "FrontPageBodyText"         ' Immediate only
'   DumpStyleProperties "FrontPageBodyText", True   ' also rpt\ file
'==============================================================================
Public Sub DumpStyleProperties(ByVal sStyleName As String, _
                               Optional ByVal bWriteFile As Boolean = False)

    Dim oDoc   As Object
    Dim oStyle As Object
    Dim oFont  As Object
    Dim oPF    As Object
    Dim sOut   As String
    Const NL   As String = vbCrLf

    Set oDoc = ActiveDocument

    On Error Resume Next
    Set oStyle = oDoc.Styles(sStyleName)
    On Error GoTo 0
    If oStyle Is Nothing Then
        Debug.Print "ERROR: style '" & sStyleName & "' not found."
        Exit Sub
    End If

    sOut = "'--- " & sStyleName & "  (Type=" & StyleTypeName(oStyle.Type) & _
           ", Priority=" & oStyle.Priority & ") ---" & NL
    sOut = sOut & ".BaseStyle = """ & CStr(oStyle.baseStyle) & """" & NL
    sOut = sOut & ".NextParagraphStyle = """ & CStr(oStyle.NextParagraphStyle) & """" & NL
    sOut = sOut & ".AutomaticallyUpdate = " & oStyle.AutomaticallyUpdate & NL
    sOut = sOut & ".QuickStyle = " & oStyle.QuickStyle & NL

    Set oFont = oStyle.Font
    sOut = sOut & ".Font.Name = """ & oFont.Name & """" & NL
    sOut = sOut & ".Font.Size = " & oFont.Size & NL
    sOut = sOut & ".Font.Bold = " & oFont.Bold & NL
    sOut = sOut & ".Font.Italic = " & oFont.Italic & NL
    sOut = sOut & ".Font.Underline = " & oFont.Underline & NL
    sOut = sOut & ".Font.Color = " & oFont.color & NL
    sOut = sOut & ".Font.SmallCaps = " & oFont.SmallCaps & NL
    sOut = sOut & ".Font.AllCaps = " & oFont.AllCaps & NL

    If oStyle.Type = 1 Then        ' wdStyleTypeParagraph
        Set oPF = oStyle.ParagraphFormat
        sOut = sOut & ".ParagraphFormat.Alignment = " & oPF.Alignment & NL
        sOut = sOut & ".ParagraphFormat.LeftIndent = " & oPF.LeftIndent & NL
        sOut = sOut & ".ParagraphFormat.RightIndent = " & oPF.RightIndent & NL
        sOut = sOut & ".ParagraphFormat.FirstLineIndent = " & oPF.FirstLineIndent & NL
        sOut = sOut & ".ParagraphFormat.SpaceBefore = " & oPF.SpaceBefore & NL
        sOut = sOut & ".ParagraphFormat.SpaceAfter = " & oPF.SpaceAfter & NL
        sOut = sOut & ".ParagraphFormat.LineSpacing = " & oPF.LineSpacing & NL
        sOut = sOut & ".ParagraphFormat.LineSpacingRule = " & oPF.LineSpacingRule & NL
        sOut = sOut & ".ParagraphFormat.WidowControl = " & oPF.WidowControl & NL
        sOut = sOut & ".ParagraphFormat.KeepTogether = " & oPF.KeepTogether & NL
        sOut = sOut & ".ParagraphFormat.KeepWithNext = " & oPF.KeepWithNext & NL
        sOut = sOut & ".ParagraphFormat.PageBreakBefore = " & oPF.PageBreakBefore & NL
        sOut = sOut & ".ParagraphFormat.OutlineLevel = " & oPF.OutlineLevel & NL
    End If

    Debug.Print sOut
    If bWriteFile Then WriteStyleDump sStyleName, sOut
End Sub

Private Function StyleTypeName(ByVal lType As Long) As String
    Select Case lType
        Case 1: StyleTypeName = "Paragraph"
        Case 2: StyleTypeName = "Character"
        Case 3: StyleTypeName = "Table"
        Case 4: StyleTypeName = "List"
        Case Else: StyleTypeName = "Unknown(" & lType & ")"
    End Select
End Function

Private Sub WriteStyleDump(ByVal sStyleName As String, ByVal sContent As String)
    Dim oFSO    As Object
    Dim oStream As Object
    Dim sPath   As String
    sPath = ActiveDocument.Path & "\rpt\style_" & SafeFileName(sStyleName) & ".txt"
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.CreateTextFile(sPath, True, False)   ' ASCII
    oStream.Write sContent
    oStream.Close
End Sub

Private Function SafeFileName(ByVal s As String) As String
    SafeFileName = Replace(Replace(Replace(s, " ", "_"), "/", "_"), "\", "_")
End Function
