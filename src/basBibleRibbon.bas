Attribute VB_Name = "basBibleRibbon"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Sub OnGoToVerseSblClick(control As IRibbonControl)
    Call GoToVerseSBL
End Sub

Sub OnHelloWorldButtonClick(control As IRibbonControl)
    MsgBox "Hello, SILAS World!" & vbCrLf & _
                "GetVScroll  = " & GetExactVerticalScroll
End Sub

Sub OnGoToH1ButtonClick(control As IRibbonControl)
    Call GoToH1
End Sub

Sub OnAdaeptAboutClick(control As IRibbonControl)
    MsgBox "Hello, adaept World!" & vbCrLf & _
                "adaeptMsg  = " & adaeptMsg, vbInformation, "About adaept"
End Sub

Function adaeptMsg() As String
    adaeptMsg = """...the truth shall make you free.""" & " John 8:32 (KJV)"
End Function

Sub GoToH1()
    Dim pattern As String
    Dim para As paragraph
    Dim paraText As String
    Dim matchFound As Boolean

    pattern = InputBox("Enter a Heading 1 pattern to match (use * and ? wildcards):", "Go To Bible Book")
    If pattern = "" Then Exit Sub ' User canceled
    matchFound = False

    ' Disable UI updates for speed
    Application.ScreenUpdating = False

    For Each para In ActiveDocument.paragraphs
        If para.style = "Heading 1" Then
            paraText = Trim$(para.range.text)
            If paraText Like pattern Then
                para.range.Select
                ' Move insertion point (cursor) without selecting text
                ActiveDocument.range(para.range.Start, para.range.Start).Select
                matchFound = True
                Exit For
            End If
        End If
    Next para

    Application.ScreenUpdating = True

    If Not matchFound Then
        MsgBox "No Heading 1 matches pattern: " & pattern, vbExclamation
    End If
End Sub

Function GetExactVerticalScroll() As Double
' Return the scroll percentage rounded to three decimal places
    Dim visibleStart As Long
    Dim totalLength As Long
    Dim scrollPercentage As Double

    ' Get the starting position of the visible content
    visibleStart = ActiveWindow.Selection.Start

    ' Get the total document length
    totalLength = ActiveDocument.content.End

    ' Calculate the exact scroll percentage
    If totalLength > 0 Then
        scrollPercentage = (visibleStart / totalLength) * 100
    Else
        scrollPercentage = 0
    End If

    ' Round to 3 decimal places
    GetExactVerticalScroll = Round(scrollPercentage, 3)
End Function

