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

'===========================================================
' RangeTokens
' Stage 9 output structure
'===========================================================
Public Type RangeTokens
    IsRange As Boolean
    LeftRaw As String
    RightRaw As String
End Type

'===========================================================
' ScriptureRef
'   Represents a single canonical Bible reference.
' Examples
'   John 3:16
'   Jude 1:5
'===========================================================
Public Type ScriptureRef
    BookID As Integer
    Chapter As Integer
    Verse As Integer
End Type

'===========================================================
' ScriptureRange
'   Represents a continuous reference range.
' Examples
'   John 3:16–18
'   John 3–5
'   John 3:16–4:2
'===========================================================
Public Type ScriptureRange
    StartRef As ScriptureRef
    EndRef As ScriptureRef
    ErrorCode As Long
    ErrorText As String
End Type

'===========================================================
' ScriptureList
'   Represents a list of references or ranges.
' Example
'   John 3:16-18,20; 4:1
'===========================================================
Public Type ScriptureList
    IsValid As Boolean
    Items() As Variant   ' ScriptureRef or ScriptureRange
                         ' Using Variant allows nested structures
    ErrorCode As Long
    ErrorText As String
End Type
