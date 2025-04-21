Attribute VB_Name = "basHelloWorld"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Sub OnHelloWorldButtonClick(control As IRibbonControl)
    MsgBox "Hello SILAS World!"
End Sub

