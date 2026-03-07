Attribute VB_Name = "basSBL_Types"
Option Explicit

Public Type ParsedReference
    ' Only structure needed for test harness
    RawInput   As String
    BookAlias  As String   ' e.g. "JUDE", "ROM"
    Chapter    As Long     ' 0 if omitted
    VerseSpec  As String   ' always string ("5", "1-3", "3,5")
End Type

Public Type LexTokens
    RawAlias As String
    Num1     As Long
    Num2     As Long
    HasColon As Boolean
End Type

Public Type ListTokens
    IsList As Boolean
    Segments() As String
End Type
