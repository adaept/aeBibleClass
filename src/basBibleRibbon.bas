Attribute VB_Name = "basBibleRibbon"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Sub OnGoToVerseSblClick(control As IRibbonControl)
    Call GoToVerseSBL
End Sub

Sub OnHelloWorldButtonClick(control As IRibbonControl)
    MsgBox "Hello SILAS World!" & vbCrLf & _
                "GetVScroll  = " & GetExactVerticalScroll
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

