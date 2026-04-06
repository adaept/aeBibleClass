VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aeBibleCitationClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'Copyright (c) 2025-2026 Peter F. Ennis
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation;
'version 3.0, or (at your option) any later version.
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'Lesser General Public License for more details.
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, visit
'https://www.gnu.org/licenses/lgpl-3.0.txt

' ================================================================
' Author:   Peter F. Ennis
' Date:     March 10, 2025
' Comment:  Bible Citation Parser for Word 365 VBA
'           - This is a Deterministic Structural Parser (DSP)
' History:  See comment details, basChangeLog_aeBibleClass, commit messages on GitHub
' GitHub:   https://github.com/adaept/aeBibleClass
' ================================================================

'===========================================================
' Type Declarations
' Note: VBA class modules require all UDTs to be Private.
' Functions using these types in their signatures are
' therefore also Private (internal to this class).
'===========================================================

Private Type ParsedReference
    ' Only structure needed for test harness
    rawInput   As String
    BookAlias  As String   ' e.g. "JUDE", "ROM"
    Chapter    As Long     ' 0 if omitted
    VerseSpec  As String   ' always string ("5", "1-3", "3,5")
End Type

Private Type LexTokens
    RawAlias As String
    Num1     As Long
    Num2     As Long
    HasColon As Boolean
    VerseSpec As String
End Type

Private Type ListTokens
    IsList As Boolean
    Segments() As String
End Type

'===========================================================
' RangeTokens
' Stage 9 output structure
'===========================================================
Private Type RangeTokens
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
Private Type ScriptureRef
    BookID As Integer
    Chapter As Integer
    Verse As Integer
End Type

'===========================================================
' ScriptureRange
'   Represents a continuous reference range.
' Examples
'   John 3:16-18
'   John 3-5
'   John 3:16-4:2
'===========================================================
Private Type ScriptureRange
    StartRef  As ScriptureRef
    EndRef    As ScriptureRef
    ErrorCode As Long
    ErrorText As String
End Type

'===========================================================
' ScriptureList
'   Represents a list of references or ranges.
' Example
'   John 3:16-18,20; 4:1
'===========================================================
Private Type ScriptureList
    IsValid As Boolean
    ItemType() As Integer
        ' 1 = ScriptureRef
        ' 2 = ScriptureRange
    Refs() As ScriptureRef
    Ranges() As ScriptureRange
    ErrorCode As Long
    ErrorText As String
End Type

Public Enum CitationMode
    ModeGeneric = 0   ' Accept common abbreviations
    ModeSBL = 1       ' Enforce SBL Study Bible rules
End Enum

'===========================================================
' ContextState (Stage 13 internal only)
'   - Tracks the inherited book/chapter while resolving
'     contextual shorthand.
' This type is NOT returned outside Stage 13.
'===========================================================
Private Type ContextState
    BookID As Integer
    Chapter As Integer
End Type

Private aliasMap As Object

Private Sub Class_Initialize()
    On Error GoTo PROC_ERR
    Debug.Print "aeSBL_Citation_Class.Class_Initialize"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    Debug.Print "ERROR in aeSBL_Citation_Class.Class_Initialize | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    Resume PROC_EXIT
End Sub

Private Sub Class_Terminate()
    On Error GoTo PROC_ERR
    Debug.Print
    Debug.Print "aeSBL_Citation_Class.Class_Terminate"
    Set aliasMap = Nothing
PROC_EXIT:
    Exit Sub
PROC_ERR:
    Debug.Print "ERROR in aeSBL_Citation_Class.Class_Terminate | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    Resume PROC_EXIT
End Sub


Public Sub ResetBookAliasMap()
    Set aliasMap = Nothing
End Sub

Private Function CanonicalFromRef(r As ScriptureRef) As String
    Dim books As Object
    Dim b As Variant
    Dim name As String

    Set books = GetCanonicalBookTable
    b = books.Item(r.BookID)
    name = b(1)   ' canonical book name

    If r.Verse = 0 Then
        CanonicalFromRef = name & " " & r.Chapter
    Else
        CanonicalFromRef = name & " " & r.Chapter & ":" & r.Verse
    End If
End Function

Private Function PackRef(r As ScriptureRef) As Variant
    PackRef = Array("REF", r.BookID, r.Chapter, r.Verse)
End Function

Private Function PackRange(r As ScriptureRange) As Variant
    Dim s As ScriptureRef
    Dim e As ScriptureRef
    s = r.StartRef
    e = r.EndRef
    PackRange = Array("RANGE", _
                      s.BookID, s.Chapter, s.Verse, _
                      e.BookID, e.Chapter, e.Verse)
End Function

Private Function CanonicalFromRange(rg As ScriptureRange) As String
    Dim startText As String, endText As String
    startText = CanonicalFromRef(rg.StartRef)
    If rg.StartRef.BookID = rg.EndRef.BookID Then
        If rg.StartRef.Chapter = rg.EndRef.Chapter Then
            ' Same chapter: suppress repeated chapter number
            endText = CStr(rg.EndRef.Verse)
        Else
            endText = CStr(rg.EndRef.Chapter)
            If rg.EndRef.Verse <> 0 Then
                endText = endText & ":" & rg.EndRef.Verse
            End If
        End If
    Else
        endText = CanonicalFromRef(rg.EndRef)
    End If
    CanonicalFromRange = startText & "-" & endText
End Function

Public Function ComposeList(raw As String) As Collection
    Dim list As ScriptureList
    Dim Result As New Collection
    Dim i As Long

    raw = NormalizeRawInput(raw)
    ComposeList_Internal raw, list
    '------------------------------------------
    ' Stage 13 - Contextual Shorthand Expansion
    '------------------------------------------
    ApplyContextShorthand list
    For i = LBound(list.ItemType) To UBound(list.ItemType)
        Select Case list.ItemType(i)
        Case 1
            Result.Add CanonicalFromRef(list.Refs(i))
        Case 2
            Result.Add CanonicalFromRange(list.Ranges(i))
        End Select
    Next i

    Set ComposeList = Result
End Function

'=====================================================
' ComposeList_Internal
' Stage 11
'   Parses a list of scripture references.
'
' Items in the resulting ScriptureList.Items() array
' may contain either:
'   ScriptureRef
'   ScriptureRange
' A Variant array is used because VBA Collections
' cannot store UDT structures. Variant arrays *can*
' store copies of UDT values safely.
' Examples
'   "John 3:16, 18"
'   "John 3:16-18, 20"
' Contextual shorthand supported:
'   "John 3:16, 18, 20-22"
'       inherits book/chapter from previous item
' Pipeline stages used
'   Stage 8  - ListDetection
'   Stage 9  - IsRangeSegment
'   Stage 10 - ParseReferenceRef
'   Stage 11 - ComposeRange
'=====================================================
Private Sub ComposeList_Internal(raw As String, ByRef Result As ScriptureList)
    Dim t As ListTokens
    Dim i As Long
    Dim count As Long
    Dim capacity As Long
    Dim prevRef As ScriptureRef
    Dim havePrev As Boolean
    Dim seg As String
    Dim p As Long
    Dim cp13 As Long    ' Stage 13a: colon position
    Dim ch13 As String  ' Stage 13a: chapter string
    Dim vr13 As String  ' Stage 13a: verse or verse-range string
    Dim dp13 As Long    ' Stage 13a: dash position within verse range

    capacity = 8
    ReDim Result.ItemType(0 To capacity - 1)
    ReDim Result.Refs(0 To capacity - 1)
    ReDim Result.Ranges(0 To capacity - 1)
    '------------------------------------------
    ' Stage 8 - list detection
    '------------------------------------------
    t = ListDetection(raw)
    For i = LBound(t.Segments) To UBound(t.Segments)
        '------------------------------------------
        ' grow arrays if needed
        '------------------------------------------
        If count >= capacity Then
            capacity = capacity * 2
            ReDim Preserve Result.ItemType(0 To capacity - 1)
            ReDim Preserve Result.Refs(0 To capacity - 1)
            ReDim Preserve Result.Ranges(0 To capacity - 1)
        End If
        seg = Trim$(t.Segments(i))
        '------------------------------------------
        ' Stage 9 - range detection
        '------------------------------------------
        If IsRangeSegment(seg) Then
            '------------------------------------------
            ' Stage 13 shorthand numeric range
            '------------------------------------------
            If havePrev And IsNumericRange(seg) Then
                p = InStr(seg, "-")
                Result.ItemType(count) = 2
                Result.Ranges(count).StartRef.BookID = prevRef.BookID
                Result.Ranges(count).StartRef.Chapter = prevRef.Chapter
                Result.Ranges(count).StartRef.Verse = CLng(Left$(seg, p - 1))
                Result.Ranges(count).EndRef.BookID = prevRef.BookID
                Result.Ranges(count).EndRef.Chapter = prevRef.Chapter
                Result.Ranges(count).EndRef.Verse = CLng(Mid$(seg, p + 1))
            '------------------------------------------
            ' Stage 13a: chapter:verse-range with inherited book
            ' e.g. "103:8-11" after "Ps 19:1"
            '------------------------------------------
            ElseIf havePrev And IsBooklessChapRef(seg) Then
                cp13 = InStr(seg, ":")
                ch13 = Left$(seg, cp13 - 1)
                vr13 = Mid$(seg, cp13 + 1)
                dp13 = InStr(vr13, "-")
                Result.ItemType(count) = 2
                Result.Ranges(count).StartRef.BookID = prevRef.BookID
                Result.Ranges(count).StartRef.Chapter = CLng(ch13)
                Result.Ranges(count).StartRef.Verse = CLng(Left$(vr13, dp13 - 1))
                Result.Ranges(count).EndRef.BookID = prevRef.BookID
                Result.Ranges(count).EndRef.Chapter = CLng(ch13)
                Result.Ranges(count).EndRef.Verse = CLng(Mid$(vr13, dp13 + 1))
            Else
                Result.ItemType(count) = 2
                Result.Ranges(count) = ComposeRange(seg)
            End If
            prevRef = Result.Ranges(count).EndRef
            havePrev = True
        Else
            '------------------------------------------
            ' Stage 13 shorthand single verse
            '------------------------------------------
            If havePrev And IsNumeric(seg) And InStr(seg, ":") = 0 Then
                Result.ItemType(count) = 1
                Result.Refs(count).BookID = prevRef.BookID
                If prevRef.Verse = 0 Then
                    ' prev was chapter-only (e.g. "Romans 8") -> bare number is next chapter
                    Result.Refs(count).Chapter = CLng(seg)
                    Result.Refs(count).Verse = 0
                Else
                    ' prev had a verse -> bare number is next verse in same chapter
                    Result.Refs(count).Chapter = prevRef.Chapter
                    Result.Refs(count).Verse = CLng(seg)
                End If
            '------------------------------------------
            ' Stage 13a: chapter:verse with inherited book
            ' e.g. "23:1" after "Ps 19:1"
            '------------------------------------------
            ElseIf havePrev And IsBooklessChapRef(seg) Then
                cp13 = InStr(seg, ":")
                Result.ItemType(count) = 1
                Result.Refs(count).BookID = prevRef.BookID
                Result.Refs(count).Chapter = CLng(Left$(seg, cp13 - 1))
                Result.Refs(count).Verse = CLng(Mid$(seg, cp13 + 1))
            Else
                Result.ItemType(count) = 1
                Result.Refs(count) = ParseReferenceRef(seg)
            End If
            prevRef = Result.Refs(count)
            havePrev = True
        End If
        count = count + 1
    Next i
    '------------------------------------------
    ' trim arrays to actual size
    '------------------------------------------
    If count > 0 Then
        ReDim Preserve Result.ItemType(0 To count - 1)
        ReDim Preserve Result.Refs(0 To count - 1)
        ReDim Preserve Result.Ranges(0 To count - 1)
    End If
    Result.IsValid = True
End Sub

'=====================================================
' ApplyContextShorthand
' Stage 13
'   Expands contextual shorthand by inheriting
'   BookID and Chapter from the previous reference.
' Examples:
'   John 3:16, 18
'       -> John 3:16
'       -> John 3:18
'   John 3:16-4:2, 5
'       -> John 3:16-4:2
'       -> John 4:5
'=====================================================
Private Sub ApplyContextShorthand(ByRef list As ScriptureList)
    Dim i As Long
    Dim currentBook As Integer
    Dim currentChapter As Integer

    For i = LBound(list.ItemType) To UBound(list.ItemType)
        Select Case list.ItemType(i)
        '------------------------------------------
        ' Single reference
        '------------------------------------------
        Case 1
            If list.Refs(i).BookID = 0 Then
                list.Refs(i).BookID = currentBook
            End If
            If list.Refs(i).Chapter = 0 Then
                list.Refs(i).Chapter = currentChapter
            End If
            currentBook = list.Refs(i).BookID
            currentChapter = list.Refs(i).Chapter
        '------------------------------------------
        ' Range
        '------------------------------------------
        Case 2
            With list.Ranges(i)
                If .StartRef.BookID = 0 Then
                    .StartRef.BookID = currentBook
                End If
                If .StartRef.Chapter = 0 Then
                    .StartRef.Chapter = currentChapter
                End If
                If .EndRef.BookID = 0 Then
                    .EndRef.BookID = .StartRef.BookID
                End If
                If .EndRef.Chapter = 0 Then
                    .EndRef.Chapter = .StartRef.Chapter
                End If
                currentBook = .EndRef.BookID
                currentChapter = .EndRef.Chapter
            End With
        End Select
    Next i
End Sub

Private Function ListDetection(ByVal rawInput As String) As ListTokens
    Dim Result As ListTokens
    Dim parts() As String

    If InStr(rawInput, ",") > 0 Then
        parts = Split(rawInput, ",")
        Result.IsList = True
        Result.Segments = parts
        ListDetection = Result
        Exit Function
    End If

    If InStr(rawInput, ";") > 0 Then
        parts = Split(rawInput, ";")
        Result.IsList = True
        Result.Segments = parts
        ListDetection = Result
        Exit Function
    End If

    Result.IsList = False
    ListDetection = Result
End Function

'===========================================================
' RangeDetection Stage 9
' Detects and tokenizes reference ranges
' Example
'   John 3:16-18
'       LeftRaw  = "John 3:16"
'       RightRaw = "18"
'===========================================================
Private Function RangeDetection(ByVal segment As String) As RangeTokens
    Dim r As RangeTokens

    If Not IsRangeSegment(segment) Then
        r.IsRange = False
        RangeDetection = r
        Exit Function
    End If

    Dim dashPos As Long
    Dim enDash As String
    enDash = ChrW(8211)

    dashPos = InStr(segment, "-")
    If dashPos = 0 Then
        dashPos = InStr(segment, enDash)
    End If

    r.IsRange = True
    r.LeftRaw = Trim(Left(segment, dashPos - 1))
    r.RightRaw = Trim(Mid$(segment, dashPos + 1))

    RangeDetection = r
End Function

'===========================================================
' IsRangeSegment
'   Determines if a segment contains a range separator.
' Supported separators:
'   Hyphen  (-) ASCII 45
'   En dash (?) Unicode U+2013
' Note
' Use ChrW(8211) so the source code does not depend
' on editors correctly displaying the en dash.
'===========================================================
Public Function IsRangeSegment(ByVal segment As String) As Boolean
    Dim enDash As String
    enDash = ChrW(8211)   ' Unicode U+2013
    If InStr(segment, "-") > 0 Then
        IsRangeSegment = True
        Exit Function
    End If
    If InStr(segment, enDash) > 0 Then
        IsRangeSegment = True
        Exit Function
    End If
End Function

Private Function ComposeRange(ByVal raw As String) As ScriptureRange
    Dim tokens As RangeTokens
    Dim r As ScriptureRange
    Dim StartRef As ScriptureRef
    Dim EndRef As ScriptureRef
    Dim normalizedRight As String

    tokens = RangeDetection(raw)
    StartRef = ParseReferenceRef(tokens.LeftRaw)
    normalizedRight = ExpandRightSide(tokens.LeftRaw, tokens.RightRaw)
    EndRef = ParseReferenceRef(normalizedRight)
    r.StartRef = StartRef
    r.EndRef = EndRef

    ComposeRange = r   ' returns Variant instead of UDT
End Function

Public Function ParseReference( _
        ByVal inputRef As String, _
        Optional ByVal mode As CitationMode = ModeGeneric _
    ) As String
' ParseReference()  -> canonical string
' ParseReferenceRef() -> structured reference

    Dim t As LexTokens
    Dim BookID As Long
    Dim canonical As String
    Dim Chapter As Long
    Dim VerseSpec As String
    Dim Verse As Long
    Dim formatted As String

    '------------------------------------------
    ' Stage 2 - Lexical scan
    '------------------------------------------
    t = LexicalScan(inputRef)
    '------------------------------------------
    ' Stage 3 - Resolve alias
    '------------------------------------------
    canonical = ResolveBookStrict(t.RawAlias, BookID, mode)
    '------------------------------------------
    ' Stage 4 - Interpret structure
    '------------------------------------------
    Dim maxCh As Long
    maxCh = GetMaxChapter(BookID)
    If t.HasColon Then
        ' Book Chapter:Verse
        Chapter = t.Num1
        VerseSpec = CStr(t.Num2)
    Else
        If maxCh = 1 Then
            ' Single chapter book (Jude 5 -> 1:5)
            Chapter = 0
            VerseSpec = CStr(t.Num1)
        Else
            ' Multi chapter book (Romans 8)
            Chapter = t.Num1
            VerseSpec = ""
        End If
    End If
    '------------------------------------------
    ' Stage 5 - Canonical validation
    '------------------------------------------
    If Not ValidateSBLReference(BookID, canonical, Chapter, VerseSpec, mode) Then
        Err.Raise vbObjectError + 600, , "Invalid scripture reference"
    End If
    '------------------------------------------
    ' Stage 6 - Canonical formatting
    '------------------------------------------
    If VerseSpec = "" Then
        If Chapter = 0 Then
            formatted = RewriteSingleChapterRef(BookID, 0, 0)
        Else
            formatted = CStr(Chapter)
        End If
    Else
        formatted = RewriteSingleChapterRef(BookID, Chapter, CLng(VerseSpec))
    End If
    ParseReference = canonical & " " & formatted
End Function

'=====================================================
' ExpandRightSide
'   Expands shorthand range tokens so they become valid
'   canonical references for the atomic parser.
' Examples
'   John 3:16-18   -> John 3:18
'   Matt 5:3-12    -> Matt 5:12
'   John 3-5       -> John 5
'   Gen 1:31-2:3   -> Gen 2:3
'=====================================================
Private Function ExpandRightSide( _
        ByVal LeftRaw As String, _
        ByVal RightRaw As String _
    ) As String

    Dim leftParts() As String
    leftParts = Split(LeftRaw, " ")

    Dim book As String
    book = leftParts(0)
    '------------------------------------------
    ' Case 1: Right side already includes book
    '------------------------------------------
    If InStr(RightRaw, " ") > 0 Then
        ExpandRightSide = RightRaw
        Exit Function
    End If
    '------------------------------------------
    ' Case 2: Cross chapter range (2:3)
    '------------------------------------------
    If InStr(RightRaw, ":") > 0 Then
        ExpandRightSide = book & " " & RightRaw
        Exit Function
    End If
    '------------------------------------------
    ' Case 3: Verse shorthand
    '------------------------------------------
    If InStr(leftParts(1), ":") > 0 Then
        Dim cv() As String
        cv = Split(leftParts(1), ":")

        ExpandRightSide = book & " " & cv(0) & ":" & RightRaw
        Exit Function
    End If
    '------------------------------------------
    ' Case 4: Chapter shorthand
    '------------------------------------------
    ExpandRightSide = book & " " & RightRaw
End Function

'=====================================================
' ParseReferenceRef
' Adapter for extension stages (Stage 10+).
'   The atomic parser (ParseReference) returns canonical
'   reference text. This adapter converts that canonical
'   string into a ScriptureRef structure.
' Example
'   ParseReference("John 3:16")
'       -> "John 3:16"
'   ParseReferenceRef("John 3:16")
'       -> ScriptureRef(BookID=43, Chapter=3, Verse=16)
'=====================================================
Private Function ParseReferenceRef( _
        ByVal inputRef As String, _
        Optional ByVal mode As CitationMode = ModeGeneric _
    ) As ScriptureRef
' ParseReference()  -> canonical string
' ParseReferenceRef() -> structured reference

    Dim canon As String
    canon = ParseReference(inputRef, mode)

    Dim r As ScriptureRef
    Dim parts() As String
    Dim cv() As String
    Dim BookID As Long

    parts = Split(canon, " ")
    '------------------------------------------
    ' Resolve book ID again from canonical name
    '------------------------------------------
    Call ResolveBookStrict(parts(0), BookID, mode)
    r.BookID = BookID
    '------------------------------------------
    ' Parse chapter / verse
    '------------------------------------------
    If InStr(parts(1), ":") > 0 Then
        cv = Split(parts(1), ":")
        r.Chapter = CLng(cv(0))
        r.Verse = CLng(cv(1))
    Else
        r.Chapter = CLng(parts(1))
        r.Verse = 0
    End If

    ParseReferenceRef = r
End Function

Public Function IsListSegment(ByVal raw As String) As Boolean
    ' A list is indicated by comma or semicolon
    If InStr(raw, ",") > 0 Then
        IsListSegment = True
        Exit Function
    End If

    If InStr(raw, ";") > 0 Then
        IsListSegment = True
        Exit Function
    End If

    IsListSegment = False
End Function

Public Function ParseScripture(ByVal raw As String) As Variant
    raw = NormalizeRawInput(raw)

    '------------------------------------------
    ' List
    '------------------------------------------
    If IsListSegment(raw) Then
        Set ParseScripture = ComposeList(raw)
        Exit Function
    End If
    '------------------------------------------
    ' Range
    '------------------------------------------
    If IsRangeSegment(raw) Then
        Dim r As ScriptureRange
        r = ComposeRange(raw)

        Dim startText As String
        Dim endText As String
        startText = CanonicalFromRef(r.StartRef)

        If r.StartRef.BookID = r.EndRef.BookID Then
            If r.StartRef.Chapter = r.EndRef.Chapter Then
                endText = CStr(r.EndRef.Verse)
            Else
                endText = CStr(r.EndRef.Chapter)
                If r.EndRef.Verse <> 0 Then
                    endText = endText & ":" & r.EndRef.Verse
                End If
            End If
        Else
            endText = CanonicalFromRef(r.EndRef)
        End If
        ParseScripture = startText & "-" & endText
        Exit Function
    End If
    '------------------------------------------
    ' Single reference
    '------------------------------------------
    ParseScripture = ParseReference(raw)
End Function

'=====================================================
' LexicalScan
'    Stage 2 of the SBL parser pipeline.
' Purpose:
'   Splits a normalized scripture reference into:
'       RawAlias  -> book alias
'       Num1      -> chapter
'       Num2      -> verse
'       HasColon  -> indicates verse present
' Special Handling:
'   Book aliases beginning with a number (e.g. "1 John",
'   "2 Tim", "3 Jn") may appear without a numeric
'   reference (e.g. "1 Jn").
'   In this case the entire input is treated as the
'   RawAlias and numeric parsing is skipped. Later
'   stages will expand this to Chapter 1 Verse 1.
' Examples:
'   "Genesis"            -> RawAlias="Genesis"
'   "Romans 8"           -> RawAlias="Romans", Num1=8
'   "John 3:16"          -> RawAlias="John", Num1=3, Num2=16
'   "1 Jn"               -> RawAlias="1 Jn"
'   "1 Corinthians 13:4" -> RawAlias="1 Corinthians"
'=====================================================
Private Function LexicalScan(ByVal normalizedInput As String) As LexTokens
    Dim t As LexTokens
    Dim parts() As String
    Dim refPart As String

    parts = Split(Trim$(normalizedInput), " ")
    Select Case UBound(parts)
        Case 0
            ' Example: "Genesis"
            t.RawAlias = parts(0)
            LexicalScan = t
            Exit Function
        Case 1
            ' Example: "Romans 8"
            ' Example: "1 Jn"
            If IsNumeric(parts(0)) Then
                t.RawAlias = parts(0) & " " & parts(1)
                LexicalScan = t
                Exit Function
            Else
                t.RawAlias = parts(0)
                refPart = parts(1)
            End If
        Case Else
            ' Example: "3 John 4"
            ' Example: "1 Corinthians 13:4"

            t.RawAlias = parts(0) & " " & parts(1)
            refPart = parts(2)
    End Select
    '----------------------------------
    ' Parse numeric portion
    '----------------------------------
    If InStr(refPart, ":") > 0 Then
        Dim subParts() As String
        subParts = Split(refPart, ":")
        ' Chapter must be numeric
        If Not IsNumeric(subParts(0)) Then
            Err.Raise vbObjectError + 1000, , _
                "Invalid chapter reference: " & refPart
        End If
        t.Num1 = CLng(subParts(0))
        ' Preserve verse specification exactly
        t.VerseSpec = subParts(1)
        ' Populate Num2 only for simple verse
        If IsNumeric(subParts(1)) Then
            t.Num2 = CLng(subParts(1))
        Else
            t.Num2 = 0
        End If
        t.HasColon = True
    Else
        ' Single numeric reference
        If Not IsNumeric(refPart) Then
            Err.Raise vbObjectError + 1001, , _
                "Invalid numeric reference: " & refPart
        End If
        t.Num1 = CLng(refPart)
        t.Num2 = 0
        t.HasColon = False
    End If

    LexicalScan = t
End Function


Public Function IsValidSBLAlias(BookID As Long, aliasText As String) As Boolean
    Dim canonical As String
    Dim books As Object
    Dim expected As String

    Set books = GetCanonicalBookTable
    canonical = books(BookID)(1)    ' e.g. "1 John"

    ' Normalize both sides
    expected = UCase$(canonical)
    aliasText = UCase$(Trim$(aliasText))

    IsValidSBLAlias = (aliasText = expected)
End Function

Public Function ResolveBookStrict( _
        abbr As String, _
        Optional BookID As Long, _
        Optional mode As CitationMode = ModeGeneric _
    ) As String

    Dim canonical As String

    ' Step 1: Resolve (existing logic)
    canonical = ResolveAlias(abbr, BookID)

    ' Step 2: Validate (NEW)
    If mode = ModeSBL Then
        If Not IsValidSBLAlias(BookID, abbr) Then
            Err.Raise vbObjectError + 20, , _
                "Non-SBL book form: '" & abbr & _
                "'. Expected '" & canonical & "'"
        End If
    End If

    ResolveBookStrict = canonical
End Function

'========================================================
' Important distinctions - Canonical vs SBL tables
'========================================================
' | Aspect   | Canonical Table | SBL Table              |
' | -------- | --------------- | ---------------------- |
' | Purpose  | Identity        | Style enforcement      |
' | Case     | Mixed           | **Uppercase required** |
' | Variants | Allowed         | **Exactly one**        |
' | Usage    | Output          | Validation             |
'========================================================
Public Function GetCanonicalBookTable() As Object
    Static books As Object

    If books Is Nothing Then
        Set books = CreateObject("Scripting.Dictionary")

        books.Add 1, Array(1, "Genesis", 50)
        books.Add 2, Array(2, "Exodus", 40)
        books.Add 3, Array(3, "Leviticus", 27)
        books.Add 4, Array(4, "Numbers", 36)
        books.Add 5, Array(5, "Deuteronomy", 34)
        books.Add 6, Array(6, "Joshua", 24)
        books.Add 7, Array(7, "Judges", 21)
        books.Add 8, Array(8, "Ruth", 4)
        books.Add 9, Array(9, "1 Samuel", 31)
        books.Add 10, Array(10, "2 Samuel", 24)
        books.Add 11, Array(11, "1 Kings", 22)
        books.Add 12, Array(12, "2 Kings", 25)
        books.Add 13, Array(13, "1 Chronicles", 29)
        books.Add 14, Array(14, "2 Chronicles", 36)
        books.Add 15, Array(15, "Ezra", 10)
        books.Add 16, Array(16, "Nehemiah", 13)
        books.Add 17, Array(17, "Esther", 10)
        books.Add 18, Array(18, "Job", 42)
        books.Add 19, Array(19, "Psalms", 150)
        books.Add 20, Array(20, "Proverbs", 31)
        books.Add 21, Array(21, "Ecclesiastes", 12)
        books.Add 22, Array(22, "Solomon", 8)
        books.Add 23, Array(23, "Isaiah", 66)
        books.Add 24, Array(24, "Jeremiah", 52)
        books.Add 25, Array(25, "Lamentations", 5)
        books.Add 26, Array(26, "Ezekiel", 48)
        books.Add 27, Array(27, "Daniel", 12)
        books.Add 28, Array(28, "Hosea", 14)
        books.Add 29, Array(29, "Joel", 3)
        books.Add 30, Array(30, "Amos", 9)
        books.Add 31, Array(31, "Obadiah", 1)
        books.Add 32, Array(32, "Jonah", 4)
        books.Add 33, Array(33, "Micah", 7)
        books.Add 34, Array(34, "Nahum", 7)
        books.Add 35, Array(35, "Habakkuk", 3)
        books.Add 36, Array(36, "Zephaniah", 3)
        books.Add 37, Array(37, "Haggai", 2)
        books.Add 38, Array(38, "Zechariah", 14)
        books.Add 39, Array(39, "Malachi", 4)
        books.Add 40, Array(40, "Matthew", 28)
        books.Add 41, Array(41, "Mark", 16)
        books.Add 42, Array(42, "Luke", 24)
        books.Add 43, Array(43, "John", 21)
        books.Add 44, Array(44, "Acts", 28)
        books.Add 45, Array(45, "Romans", 16)
        books.Add 46, Array(46, "1 Corinthians", 16)
        books.Add 47, Array(47, "2 Corinthians", 13)
        books.Add 48, Array(48, "Galatians", 6)
        books.Add 49, Array(49, "Ephesians", 6)
        books.Add 50, Array(50, "Philippians", 4)
        books.Add 51, Array(51, "Colossians", 4)
        books.Add 52, Array(52, "1 Thessalonians", 5)
        books.Add 53, Array(53, "2 Thessalonians", 3)
        books.Add 54, Array(54, "1 Timothy", 6)
        books.Add 55, Array(55, "2 Timothy", 4)
        books.Add 56, Array(56, "Titus", 3)
        books.Add 57, Array(57, "Philemon", 1)
        books.Add 58, Array(58, "Hebrews", 13)
        books.Add 59, Array(59, "James", 5)
        books.Add 60, Array(60, "1 Peter", 5)
        books.Add 61, Array(61, "2 Peter", 3)
        books.Add 62, Array(62, "1 John", 5)
        books.Add 63, Array(63, "2 John", 1)
        books.Add 64, Array(64, "3 John", 1)
        books.Add 65, Array(65, "Jude", 1)
        books.Add 66, Array(66, "Revelation", 22)
    End If

    Set GetCanonicalBookTable = books
End Function

Public Function GetSBLCanonicalBookTable() As Object
    Static sbl As Object

    If sbl Is Nothing Then
        Set sbl = CreateObject("Scripting.Dictionary")

        sbl.Add 1, "GENESIS"
        sbl.Add 2, "EXODUS"
        sbl.Add 3, "LEVITICUS"
        sbl.Add 4, "NUMBERS"
        sbl.Add 5, "DEUTERONOMY"
        sbl.Add 6, "JOSHUA"
        sbl.Add 7, "JUDGES"
        sbl.Add 8, "RUTH"
        sbl.Add 9, "1 SAMUEL"
        sbl.Add 10, "2 SAMUEL"
        sbl.Add 11, "1 KINGS"
        sbl.Add 12, "2 KINGS"
        sbl.Add 13, "1 CHRONICLES"
        sbl.Add 14, "2 CHRONICLES"
        sbl.Add 15, "EZRA"
        sbl.Add 16, "NEHEMIAH"
        sbl.Add 17, "ESTHER"
        sbl.Add 18, "JOB"
        sbl.Add 19, "PSALMS"
        sbl.Add 20, "PROVERBS"
        sbl.Add 21, "ECCLESIASTES"
        sbl.Add 22, "SOLOMON"
        sbl.Add 23, "ISAIAH"
        sbl.Add 24, "JEREMIAH"
        sbl.Add 25, "LAMENTATIONS"
        sbl.Add 26, "EZEKIEL"
        sbl.Add 27, "DANIEL"
        sbl.Add 28, "HOSEA"
        sbl.Add 29, "JOEL"
        sbl.Add 30, "AMOS"
        sbl.Add 31, "OBADIAH"
        sbl.Add 32, "JONAH"
        sbl.Add 33, "MICAH"
        sbl.Add 34, "NAHUM"
        sbl.Add 35, "HABAKKUK"
        sbl.Add 36, "ZEPHANIAH"
        sbl.Add 37, "HAGGAI"
        sbl.Add 38, "ZECHARIAH"
        sbl.Add 39, "MALACHI"
        sbl.Add 40, "MATTHEW"
        sbl.Add 41, "MARK"
        sbl.Add 42, "LUKE"
        sbl.Add 43, "JOHN"
        sbl.Add 44, "ACTS"
        sbl.Add 45, "ROMANS"
        sbl.Add 46, "1 CORINTHIANS"
        sbl.Add 47, "2 CORINTHIANS"
        sbl.Add 48, "GALATIONS"
        sbl.Add 49, "EPHESIANS"
        sbl.Add 50, "PHILIPPIANS"
        sbl.Add 51, "COLOSSIANS"
        sbl.Add 52, "1 THESSALONIANS"
        sbl.Add 53, "2 THESSALONIANS"
        sbl.Add 54, "1 TIMOTHY"
        sbl.Add 55, "2 TIMOTHY"
        sbl.Add 56, "TITUS"
        sbl.Add 57, "PHILEMON"
        sbl.Add 58, "HEBREWS"
        sbl.Add 59, "JAMES"
        sbl.Add 60, "1 PETER"
        sbl.Add 61, "2 PETER"
        sbl.Add 62, "1 JOHN"
        sbl.Add 63, "2 JOHN"
        sbl.Add 64, "3 JOHN"
        sbl.Add 65, "JUDE"
        sbl.Add 66, "REVELATION"
    End If

    Set GetSBLCanonicalBookTable = sbl
End Function

Public Sub ValidateBookSBL( _
        ByVal BookID As Long, _
        ByVal InputBook As String)

    Dim sbl As Object
    Dim expected As String

    Set sbl = GetSBLCanonicalBookTable

    If Not sbl.Exists(BookID) Then
        Err.Raise vbObjectError + 400, , _
            "Book not defined in SBL canon: " & BookID
    End If

    expected = sbl(BookID)

    If UCase(Trim(InputBook)) <> expected Then
        Err.Raise vbObjectError + 401, , _
            "Non-SBL book form. Expected '" & expected & _
            "', got '" & InputBook & "'"
    End If
End Sub

Public Function GetMaxChapter(BookID As Long) As Long
    Dim books As Object
    Dim b As Variant

    Set books = GetCanonicalBookTable

    If Not books.Exists(BookID) Then
        Err.Raise vbObjectError + 20, , "Unknown BookID in GetMaxChapter"
    End If

    b = books(BookID)
    GetMaxChapter = CLng(b(2))
End Function

Public Sub AssertOneBased(arr As Variant, context As String)
    If LBound(arr) <> 1 Then
        Err.Raise vbObjectError + 900, , _
            "0-Based array detected in " & context
    End If
End Sub

Public Function GetMaxVerse(BookID As Long, Chapter As Long) As Long
    Dim maps As Variant
    maps = GetChapterVerseMap()   ' actually raw map now

    AssertOneBased maps(BookID), "GetMaxVerse chapters"

    If BookID < 1 Or BookID > 66 Then
        Err.Raise vbObjectError + 41, , "Invalid BookID"
    End If

    If Not IsArray(maps(BookID)) Then
        Err.Raise vbObjectError + 42, , "Chapter map missing"
    End If

    If Chapter < 1 Or Chapter > UBound(maps(BookID)) + 1 Then
        Err.Raise vbObjectError + 40, , "Invalid chapter for book"
    End If

    GetMaxVerse = maps(BookID)(Chapter)
End Function

Public Function GetChapterVerseMap() As Variant

    Static cached As Variant

    If IsEmpty(cached) Then

        Dim d As Object
        Set d = GetVerseCounts()

        Dim maps(1 To 66) As Variant
        Dim BookID As Long

        For BookID = 1 To 66
            maps(BookID) = d(BookID)
            AssertOneBased maps(BookID), "GetChapterVerseMap chapters"
        Next BookID

        AssertOneBased maps, "GetChapterVerseMap books"

        cached = maps
    End If

    GetChapterVerseMap = cached

End Function

Public Function ValidateSBLReference( _
        BookID As Long, _
        canonicalName As String, _
        Chapter As Long, _
        VerseSpec As String, _
        mode As CitationMode, _
        Optional silent As Boolean = False) As Boolean

    ' Generic mode: always valid at this layer
    If mode = ModeGeneric Then
        ValidateSBLReference = True
        Exit Function
    End If

    '------------------------------------------
    ' Normalize single-chapter shorthand
    '------------------------------------------
    If Chapter = 0 And VerseSpec <> "" Then
        If GetMaxChapter(BookID) = 1 Then
            ' Single-chapter book
            ' Example: Jude 5 -> Jude 1:5
            Chapter = 1
        Else
            ' Multi-chapter book
            ' Example: Romans 8 -> Romans 8
            Chapter = CLng(VerseSpec)
            VerseSpec = ""
        End If
    End If

    ' ---------- SBL MODE BELOW ----------
    ' 1. Book must exist in canonical table
    If Not GetCanonicalBookTable.Exists(BookID) Then
        Debug.Print "SBL FAIL: Unknown BookID " & BookID
        Exit Function
    End If

    ' 2. Canonical name must match SBL form EXACTLY
    ' (SBL is case-insensitive in print, but normalized internally)
    Dim canon As Variant
    canon = GetCanonicalBookTable(BookID)

    If UCase(canon(1)) <> UCase(canonicalName) Then
        Debug.Print "SBL FAIL: Non-canonical book name"
        Exit Function
    End If

    ' 3. Chapter rules
    If Chapter < 1 Then
        If Not silent Then Debug.Print "SBL FAIL: Chapter must be >= 1"
        ValidateSBLReference = False
        Exit Function
    End If

    Dim maxCh As Long
    maxCh = GetMaxChapter(BookID)
    If Chapter > maxCh Then
        If Not silent Then Debug.Print "SBL FAIL: Chapter exceeds max (" & maxCh & ")"
        ValidateSBLReference = False
        Exit Function
    End If

    ' 4. Single-chapter book rules
    If GetSingleChapterBookSet.Exists(BookID) Then
        If Chapter > 1 Then
            Debug.Print "SBL FAIL: Chapter > 1 for single-chapter book"
            Exit Function
        End If
    End If

    ' 5. Whole-chapter reference (VerseSpec empty after normalization) is valid
    If Len(VerseSpec) = 0 Then
        ValidateSBLReference = True
        Exit Function
    End If

    Dim v As Long

    If Not IsNumeric(VerseSpec) Then
        Debug.Print "SBL FAIL: Non-numeric verse"
        Exit Function
    End If

    v = CLng(VerseSpec)

    If v < 1 Then
        If Not silent Then Debug.Print "SBL FAIL: Verse must be >= 1"
        ValidateSBLReference = False
        Exit Function
    End If

    Dim maxV As Long
    maxV = GetMaxVerse(BookID, Chapter)

    If v > maxV Then
        Debug.Print "SBL FAIL: Verse exceeds max (" & maxV & ")"
        Exit Function
    End If

    ' Note:
    ' Verse range bounds (max verse per chapter)
    ' are intentionally NOT enforced here yet.
    ' That is a later enhancement.

    ValidateSBLReference = True
End Function

Public Function GetSingleChapterBookSet() As Object
    Static sc As Object

    If sc Is Nothing Then
        Set sc = CreateObject("Scripting.Dictionary")
        ' Old Testament
        sc.Add 31, True   ' Obadiah
        ' New Testament
        sc.Add 57, True   ' Philemon
        sc.Add 63, True   ' 2 John
        sc.Add 64, True   ' 3 John
        sc.Add 65, True   ' Jude
    End If

    Set GetSingleChapterBookSet = sc
End Function

Private Function InterpretStructure(ByRef tokens As LexTokens) As ParsedReference
    Dim ref As ParsedReference
    '------------------------------------------
    ' Chapter : Verse structure
    '------------------------------------------
    If tokens.HasColon Then
        ref.Chapter = tokens.Num1
        ' Preserve verse specification
        If Len(tokens.VerseSpec) > 0 Then
            ref.VerseSpec = tokens.VerseSpec
        ElseIf tokens.Num2 > 0 Then
            ref.VerseSpec = CStr(tokens.Num2)
        Else
            ref.VerseSpec = ""
        End If
    Else
        '------------------------------------------
        ' Single numeric reference (ambiguous)
        '------------------------------------------
        If tokens.Num1 > 0 Then
            ref.Chapter = 0
            ref.VerseSpec = CStr(tokens.Num1)
        Else
            ref.Chapter = 0
            ref.VerseSpec = ""
        End If
    End If

    InterpretStructure = ref
End Function

Public Function RewriteSingleChapterRef( _
        ByVal BookID As Long, _
        ByVal Chapter As Long, _
        ByVal Verse As Long) As String

    Dim sc As Object
    Set sc = GetSingleChapterBookSet
    '-----------------------------------------
    ' BOOK ONLY
    ' Example: John -> 1:1
    '-----------------------------------------
    If Chapter = 0 And Verse = 0 Then
        RewriteSingleChapterRef = "1:1"
        Exit Function
    End If
    '-----------------------------------------
    ' Single-chapter books
    ' Example: Jude 5 -> 1:5
    '-----------------------------------------
    If sc.Exists(BookID) And Chapter = 0 Then
        RewriteSingleChapterRef = "1:" & Verse
        Exit Function
    End If
    '-----------------------------------------
    ' Normal reference
    '-----------------------------------------
    If Verse > 0 Then
        RewriteSingleChapterRef = Chapter & ":" & Verse
    Else
        ' Chapter-only reference OR book-only reference
        If Chapter = 0 Then
            RewriteSingleChapterRef = "1:1"
        Else
            RewriteSingleChapterRef = CStr(Chapter)
        End If
    End If
End Function

Public Function AliasCoverage( _
        Optional ByRef report As String = "" _
    ) As Boolean
' Validate the Alias Coverage

    Dim books As Object
    Dim aMap As Object
    Dim missing As Collection
    Dim k As Variant
    Dim canon As String

    Set books = GetCanonicalBookTable
    Set aMap = GetBookAliasMap
    Set missing = New Collection

    For Each k In books.Keys
        canon = UCase$(books(k)(1))   ' Canonical name

        If Not aMap.Exists(canon) Then
            missing.Add canon
        End If
    Next k

    If missing.Count > 0 Then
        Dim i As Long
        report = "Missing canonical aliases:" & vbCrLf
        For i = 1 To missing.Count
            report = report & "  - " & missing(i) & vbCrLf
        Next i

        AliasCoverage = False
    Else
        report = "Alias coverage complete (canonical names present)."
        AliasCoverage = True
    End If
End Function

Public Function GetBookAliasMap() As Object
    ' Single-letter aliases are not allowed due to potential false positives
    ' Sort form allowed, common in Europe, (International / Critical Apparatus Style)

    If aliasMap Is Nothing Then
        Set aliasMap = CreateObject("Scripting.Dictionary")

        ' Genesis
        aliasMap.Add "GENESIS", 1
        aliasMap.Add "GEN", 1
        aliasMap.Add "GE", 1
        aliasMap.Add "GN", 1
        ' Exodus
        aliasMap.Add "EXODUS", 2
        aliasMap.Add "EXOD", 2
        aliasMap.Add "EXO", 2
        aliasMap.Add "EX", 2
        ' Leviticus
        aliasMap.Add "LEVITICUS", 3
        aliasMap.Add "LEV", 3
        aliasMap.Add "LE", 3
        aliasMap.Add "LV", 3
        ' Numbers
        aliasMap.Add "NUMBERS", 4
        aliasMap.Add "NUM", 4
        aliasMap.Add "NU", 4
        aliasMap.Add "NM", 4
        ' Deuteronomy
        aliasMap.Add "DEUTERONOMY", 5
        aliasMap.Add "DEUT", 5
        aliasMap.Add "DEU", 5
        aliasMap.Add "DE", 5
        aliasMap.Add "DT", 5
        ' Joshua
        aliasMap.Add "JOSHUA", 6
        aliasMap.Add "JOSH", 6
        aliasMap.Add "JOS", 6
        ' Judges
        aliasMap.Add "JUDGES", 7
        aliasMap.Add "JUDGE", 7
        aliasMap.Add "JUDG", 7
        aliasMap.Add "JGS", 7
        ' Ruth
        aliasMap.Add "RUTH", 8
        aliasMap.Add "RUT", 8
        aliasMap.Add "RU", 8
        ' 1 Samuel
        aliasMap.Add "1 SAMUEL", 9
        aliasMap.Add "1 SAM", 9
        aliasMap.Add "1 SA", 9
        aliasMap.Add "1 SM", 9
        ' 2 Samuel
        aliasMap.Add "2 SAMUEL", 10
        aliasMap.Add "2 SAM", 10
        aliasMap.Add "2 SA", 10
        aliasMap.Add "2 SM", 10
        ' 1 Kings
        aliasMap.Add "1 KINGS", 11
        aliasMap.Add "1 KGS", 11
        aliasMap.Add "1 KING", 11
        aliasMap.Add "1 KIN", 11
        aliasMap.Add "1 KI", 11
        ' 2 Kings
        aliasMap.Add "2 KINGS", 12
        aliasMap.Add "2 KGS", 12
        aliasMap.Add "2 KING", 12
        aliasMap.Add "2 KIN", 12
        aliasMap.Add "2 KI", 12
        ' 1 Chronicles
        aliasMap.Add "1 CHRONICLES", 13
        aliasMap.Add "1 CHRON", 13
        aliasMap.Add "1 CHRO", 13
        aliasMap.Add "1 CHR", 13
        aliasMap.Add "1 CH", 13
        ' 2 Chronicles
        aliasMap.Add "2 CHRONICLES", 14
        aliasMap.Add "2 CHRON", 14
        aliasMap.Add "2 CHRO", 14
        aliasMap.Add "2 CHR", 14
        aliasMap.Add "2 CH", 14
        ' Ezra
        aliasMap.Add "EZRA", 15
        aliasMap.Add "EZR", 15
        ' Nehemiah
        aliasMap.Add "NEHEMIAH", 16
        aliasMap.Add "NEH", 16
        aliasMap.Add "NE", 16
        ' Esther
        aliasMap.Add "ESTHER", 17
        aliasMap.Add "ESTH", 17
        aliasMap.Add "EST", 17
        aliasMap.Add "ES", 17
        ' Job
        aliasMap.Add "JOB", 18
        aliasMap.Add "JB", 18
        ' Psalms
        aliasMap.Add "PSALMS", 19
        aliasMap.Add "PSALM", 19
        aliasMap.Add "PSA", 19
        aliasMap.Add "PS", 19
        ' Proverbs
        aliasMap.Add "PROVERBS", 20
        aliasMap.Add "PROV", 20
        aliasMap.Add "PRO", 20
        aliasMap.Add "PR", 20
        aliasMap.Add "PRV", 20
        ' Ecclesiastes
        aliasMap.Add "ECCLESIASTES", 21
        aliasMap.Add "ECCL", 21
        aliasMap.Add "ECC", 21
        aliasMap.Add "EC", 21
        ' Solomon
        aliasMap.Add "SOLOMON", 22
        aliasMap.Add "SOLO", 22
        aliasMap.Add "SOL", 22
        aliasMap.Add "SO", 22
        aliasMap.Add "SONG", 22
        aliasMap.Add "SG", 22
        ' Isaiah
        aliasMap.Add "ISAIAH", 23
        aliasMap.Add "ISA", 23
        aliasMap.Add "IS", 23
        ' Jeremiah
        aliasMap.Add "JEREMIAH", 24
        aliasMap.Add "JER", 24
        aliasMap.Add "JE", 24
        ' Lamentations
        aliasMap.Add "LAMENTATIONS", 25
        aliasMap.Add "LAM", 25
        aliasMap.Add "LA", 25
        ' Ezekiel
        aliasMap.Add "EZEKIEL", 26
        aliasMap.Add "EZEK", 26
        aliasMap.Add "EZE", 26
        aliasMap.Add "EZ", 26
        ' Daniel
        aliasMap.Add "DANIEL", 27
        aliasMap.Add "DAN", 27
        aliasMap.Add "DA", 27
        aliasMap.Add "DN", 27
        ' Hosea
        aliasMap.Add "HOSEA", 28
        aliasMap.Add "HOS", 28
        aliasMap.Add "HO", 28
        ' Joel
        aliasMap.Add "JOEL", 29
        aliasMap.Add "JOE", 29
        aliasMap.Add "JL", 29
        ' Amos
        aliasMap.Add "AMOS", 30
        aliasMap.Add "AMO", 30
        aliasMap.Add "AM", 30
        ' Obadiah
        aliasMap.Add "OBADIAH", 31
        aliasMap.Add "OBAD", 31
        aliasMap.Add "OBA", 31
        aliasMap.Add "OB", 31
        ' Jonah
        aliasMap.Add "JONAH", 32
        aliasMap.Add "JONA", 32
        aliasMap.Add "JON", 32
        ' Micah
        aliasMap.Add "MICAH", 33
        aliasMap.Add "MIC", 33
        aliasMap.Add "MI", 33
        ' Nahum
        aliasMap.Add "NAHUM", 34
        aliasMap.Add "NAH", 34
        aliasMap.Add "NA", 34
        ' Habakkuk
        aliasMap.Add "HABAKKUK", 35
        aliasMap.Add "HAB", 35
        aliasMap.Add "HA", 35
        aliasMap.Add "HB", 35
        ' Zephaniah
        aliasMap.Add "ZEPHANIAH", 36
        aliasMap.Add "ZEPH", 36
        aliasMap.Add "ZEP", 36
        ' Haggai
        aliasMap.Add "HAGGAI", 37
        aliasMap.Add "HAG", 37
        aliasMap.Add "HG", 37
        ' Zechariah
        aliasMap.Add "ZECHARIAH", 38
        aliasMap.Add "ZECH", 38
        aliasMap.Add "ZEC", 38
        ' Malachi
        aliasMap.Add "MALACHI", 39
        aliasMap.Add "MAL", 39
        ' Matthew
        aliasMap.Add "MATTHEW", 40
        aliasMap.Add "MATT", 40
        aliasMap.Add "MAT", 40
        aliasMap.Add "MT", 40
        ' Mark
        aliasMap.Add "MARK", 41
        aliasMap.Add "MAR", 41
        aliasMap.Add "MK", 41
        ' Luke
        aliasMap.Add "LUKE", 42
        aliasMap.Add "LUK", 42
        aliasMap.Add "LU", 42
        aliasMap.Add "LK", 42
        ' John
        aliasMap.Add "JOHN", 43
        aliasMap.Add "JOH", 43
        aliasMap.Add "JN", 43
        ' Acts
        aliasMap.Add "ACTS", 44
        aliasMap.Add "ACT", 44
        aliasMap.Add "AC", 44
        ' Romans
        aliasMap.Add "ROMANS", 45
        aliasMap.Add "ROM", 45
        aliasMap.Add "RO", 45
        ' 1 Corinthians
        aliasMap.Add "1 CORINTHIANS", 46
        aliasMap.Add "1 COR", 46
        aliasMap.Add "1 CO", 46
        ' 2 Corinthians
        aliasMap.Add "2 CORINTHIANS", 47
        aliasMap.Add "2 COR", 47
        aliasMap.Add "2 CO", 47
        ' Galatians
        aliasMap.Add "GALATIANS", 48
        aliasMap.Add "GAL", 48
        aliasMap.Add "GA", 48
        ' Ephesians
        aliasMap.Add "EPHESIANS", 49
        aliasMap.Add "EPH", 49
        aliasMap.Add "EP", 49
        ' Philippians
        aliasMap.Add "PHILIPPIANS", 50
        aliasMap.Add "PHILI", 50
        aliasMap.Add "PHIL", 50
        ' Colossians
        aliasMap.Add "COLOSSIANS", 51
        aliasMap.Add "COL", 51
        aliasMap.Add "CO", 51
        ' 1 Thessalonians
        aliasMap.Add "1 THESSALONIANS", 52
        aliasMap.Add "1 THESS", 52
        aliasMap.Add "1 THES", 52
        aliasMap.Add "1 THE", 52
        aliasMap.Add "1 THS", 52
        ' 2 Thessalonians
        aliasMap.Add "2 THESSALONIANS", 53
        aliasMap.Add "2 THESS", 53
        aliasMap.Add "2 THES", 53
        aliasMap.Add "2 THE", 53
        aliasMap.Add "2 THS", 53
        ' 1 Timothy
        aliasMap.Add "1 TIMOTHY", 54
        aliasMap.Add "1 TIM", 54
        aliasMap.Add "1 TI", 54
        aliasMap.Add "1 TM", 54
        ' 2 Timothy
        aliasMap.Add "2 TIMOTHY", 55
        aliasMap.Add "2 TIM", 55
        aliasMap.Add "2 TI", 55
        aliasMap.Add "2 TM", 55
        ' Titus
        aliasMap.Add "TITUS", 56
        aliasMap.Add "TIT", 56
        aliasMap.Add "TI", 56
        ' Philemon
        aliasMap.Add "PHILEMON", 57
        aliasMap.Add "PHILE", 57
        aliasMap.Add "PHLM", 57
        ' Hebrews
        aliasMap.Add "HEBREWS", 58
        aliasMap.Add "HEB", 58
        aliasMap.Add "HE", 58
        ' James
        aliasMap.Add "JAMES", 59
        aliasMap.Add "JAM", 59
        aliasMap.Add "JA", 59
        aliasMap.Add "JAS", 59
        ' 1 Peter
        aliasMap.Add "1 PETER", 60
        aliasMap.Add "1 PET", 60
        aliasMap.Add "1 PE", 60
        aliasMap.Add "1 PT", 60
        ' 2 Peter
        aliasMap.Add "2 PETER", 61
        aliasMap.Add "2 PET", 61
        aliasMap.Add "2 PE", 61
        aliasMap.Add "2 PT", 61
        ' 1 John
        aliasMap.Add "1 JOHN", 62
        aliasMap.Add "1 JOH", 62
        aliasMap.Add "1 JN", 62
        ' 2 John
        aliasMap.Add "2 JOHN", 63
        aliasMap.Add "2 JOH", 63
        aliasMap.Add "2 JN", 63
        ' 3 John
        aliasMap.Add "3 JOHN", 64
        aliasMap.Add "3 JOH", 64
        aliasMap.Add "3 JN", 64
        ' Jude
        aliasMap.Add "JUDE", 65
        aliasMap.Add "JUD", 65
        ' Revelation
        aliasMap.Add "REVELATION", 66
        aliasMap.Add "REV", 66
        aliasMap.Add "RE", 66
        aliasMap.Add "RV", 66
    End If

    Set GetBookAliasMap = aliasMap
End Function

Private Function IsNumericRange(ByVal seg As String) As Boolean
    Dim p As Long
    p = InStr(seg, "-")
    If p = 0 Then Exit Function
    IsNumericRange = _
        IsNumeric(Left$(seg, p - 1)) And _
        IsNumeric(Mid$(seg, p + 1))
End Function

'=====================================================
' IsBooklessChapRef  (Stage 13a)
' Returns True when seg begins with a numeric chapter
' and a colon — meaning it is a chapter:verse or
' chapter:verse-range segment with no book alias.
' Examples: "23:1", "103:8-11"
'=====================================================
Private Function IsBooklessChapRef(ByVal seg As String) As Boolean
    Dim p As Long
    p = InStr(seg, ":")
    If p = 0 Then Exit Function
    IsBooklessChapRef = IsNumeric(Left$(seg, p - 1))
End Function

'=====================================================
' Stage 14 - Canonical Compression
'=====================================================
' Purpose:
'   Collapse adjacent consecutive same-chapter verses
'   into ranges in a resolved canonical reference list.
' Input:
'   A Collection of canonical reference strings,
'   such as the output of ComposeList / CanonicalFromRef.
'   Example: "John 3:16", "John 3:17", "John 3:18", "Romans 8:1"
' Output:
'   A Collection of compressed canonical strings.
'   Example: "John 3:16-3:18", "Romans 8:1"
' Compression rules:
'   1. Two refs are adjacent when same book, same chapter,
'      and second verse = first verse + 1.
'   2. A run of 2+ adjacent verses becomes a range:
'         BookName startChapter:startVerse-endChapter:endVerse
'   3. Cross-chapter adjacency is never collapsed.
'   4. Cross-book adjacency is never collapsed.
'   5. Order is preserved.
'   6. Non-canonical strings (chapter-only refs) are passed
'      through unchanged.
'=====================================================
Public Function CompressCanonical(ByVal Refs As Collection) As Collection
    On Error GoTo PROC_ERR
    Dim Result As New Collection
    Dim i As Long
    Dim n As Long

    n = Refs.Count
    If n = 0 Then
        Set CompressCanonical = Result
        GoTo PROC_EXIT
    End If
    '------------------------------------------
    ' Parse all refs into parallel arrays
    '------------------------------------------
    Dim bookNames()  As String
    Dim chapters()   As Long
    Dim verses()     As Long
    Dim parsedOk()   As Boolean

    ReDim bookNames(1 To n)
    ReDim chapters(1 To n)
    ReDim verses(1 To n)
    ReDim parsedOk(1 To n)

    Dim refStr As String
    For i = 1 To n
        refStr = Refs(i)
        parsedOk(i) = ParseCanonicalRef( _
                          refStr, _
                          bookNames(i), _
                          chapters(i), _
                          verses(i))
    Next i
    '------------------------------------------
    ' Walk through refs, accumulating runs
    '------------------------------------------
    i = 1
    Do While i <= n
        '------------------------------------------
        ' If this ref could not be parsed as a
        ' chapter:verse ref, emit it unchanged
        '------------------------------------------
        If Not parsedOk(i) Then
            Result.Add Refs(i)
            i = i + 1
        Else
            '------------------------------------------
            ' Start a potential run
            '------------------------------------------
            Dim runStart As Long
            runStart = i
            '------------------------------------------
            ' Extend the run while adjacent.
            ' Note: VBA And is not short-circuit - guard
            ' the array accesses with a Do While i < n
            ' outer check and inner If/Exit Do tests.
            '------------------------------------------
            Do While i < n
                If Not parsedOk(i + 1) Then Exit Do
                If bookNames(i + 1) <> bookNames(i) Then Exit Do
                If chapters(i + 1) <> chapters(i) Then Exit Do
                If verses(i + 1) <> verses(i) + 1 Then Exit Do
                i = i + 1
            Loop
            '------------------------------------------
            ' Emit the run as a range or single ref
            '------------------------------------------
            If i > runStart Then
                ' Collapsed range: BookName startCh:startV-endCh:endV
                Result.Add bookNames(runStart) & " " & _
                           chapters(runStart) & ":" & verses(runStart) & _
                           "-" & _
                           chapters(i) & ":" & verses(i)
            Else
                ' Single ref: emit as-is (canonical form preserved)
                Result.Add Refs(i)
            End If

            i = i + 1
        End If
    Loop

    Set CompressCanonical = Result

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure CompressCanonical of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

'=====================================================
' ParseCanonicalRef
'   Parses a canonical reference string into its
'   component parts: book name, chapter, verse.
' Accepts strings of the form:
'   "BookName chapter:verse"   e.g. "John 3:16"
'   "BookName chapter"         e.g. "Romans 8"
'                              (verse returns 0; parsedOk=False)
' Multi-word book names ("1 John", "Song of Solomon")
'   are handled by treating everything up to the last
'   space-delimited token that contains ":" (or is purely
'   numeric) as the book name.
' Returns True when a chapter:verse pair was successfully
'   extracted. Returns False for chapter-only refs or
'   strings that cannot be parsed.
'=====================================================
Private Function ParseCanonicalRef( _
        ByVal canonical As String, _
        ByRef bookName As String, _
        ByRef Chapter As Long, _
        ByRef Verse As Long _
    ) As Boolean
    On Error GoTo PROC_ERR

    bookName = ""
    Chapter = 0
    Verse = 0

    canonical = Trim$(canonical)
    If Len(canonical) = 0 Then GoTo PROC_EXIT

    '------------------------------------------
    ' Find the last space ? everything after it
    ' is the chapter[:verse] token.
    '------------------------------------------
    Dim lastSpace As Long
    lastSpace = 0
    Dim k As Long
    For k = Len(canonical) To 1 Step -1
        If Mid$(canonical, k, 1) = " " Then
            lastSpace = k
            Exit For
        End If
    Next k

    If lastSpace = 0 Then GoTo PROC_EXIT   ' no space found ? unparseable

    Dim refToken As String
    refToken = Mid$(canonical, lastSpace + 1)
    bookName = Left$(canonical, lastSpace - 1)

    '------------------------------------------
    ' Parse chapter[:verse] token
    '------------------------------------------
    Dim colonPos As Long
    colonPos = InStr(refToken, ":")

    If colonPos > 0 Then
        Dim chStr As String
        Dim vStr As String
        chStr = Left$(refToken, colonPos - 1)
        vStr = Mid$(refToken, colonPos + 1)
        If IsNumeric(chStr) And IsNumeric(vStr) Then
            Chapter = CLng(chStr)
            Verse = CLng(vStr)
            ParseCanonicalRef = True   ' chapter:verse successfully parsed
        End If
    End If
    ' If no colon (chapter-only ref), we return False
    ' (chapter-only refs are emitted unchanged by CompressCanonical)

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ParseCanonicalRef of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

Public Function ResolveAlias(abbr As String, _
                            Optional BookID As Long) As String
    Dim key As String
    Dim aMap As Object
    Dim books As Object
    Dim b As Variant

    key = UCase$(Trim$(abbr))
    Set aMap = GetBookAliasMap

    If Not aMap.Exists(key) Then
        Debug.Print "BAD  > ResolveBook(" & key & ")"
        Err.Raise vbObjectError + 10, , "Unknown book alias: " & abbr
    End If

    BookID = aMap(key)
    Set books = GetCanonicalBookTable

    b = books.Item(BookID)     ' SAFE
    ResolveAlias = b(1)         ' Canonical name
End Function

'=====================================================
' Stage 15 - Canonical Validation
'=====================================================
' Purpose:
'     Remove or trim canonical reference strings that fall
'     outside the valid scripture bounds of their book.
' Input:
'     A Collection of canonical strings ? output of Stage 14.
'     Strings may be:
'         Single ref:   "John 3:16"
'         Range:        "John 3:16-3:18"
'         Chapter-only: "Romans 8"   (passed through unchanged)
' Output:
'     A new Collection with invalid refs removed or trimmed.
' Processing rules per item:
'     Single ref "Book ch:v"
'         - ch > ChaptersInBook  -> remove
'         - v  > VersesInChapter -> remove
'         - otherwise            -> keep
'     Range "Book startCh:startV-endCh:endV"
'         - startCh > ChaptersInBook        -> remove
'         - startV  > VersesInChapter(start) -> remove
'         - endCh: clamp to ChaptersInBook
'         - endV:  clamp to VersesInChapter(endCh)
'         - after clamping, end < start      -> remove
'         - after clamping, end = start      -> emit single ref
'         - otherwise                        -> emit clamped range
'     Chapter-only: pass through unchanged
'=====================================================

'-----------------------------------------------------
' ChaptersInBook
'   Returns the number of chapters in the named book.
'   Returns 0 if the book name cannot be resolved.
'-----------------------------------------------------
Private Function ChaptersInBook(ByVal bookName As String) As Long
    On Error GoTo PROC_ERR
    Dim aMap As Object
    Dim key As String
    Dim BookID As Long

    key = UCase$(Trim$(bookName))
    Set aMap = GetBookAliasMap

    If Not aMap.Exists(key) Then
        ChaptersInBook = 0
        GoTo PROC_EXIT
    End If
    BookID = aMap(key)
    ChaptersInBook = GetMaxChapter(BookID)

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ChaptersInBook of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

'-----------------------------------------------------
' VersesInChapter
'   Returns the number of verses in the specified
'   chapter of the named book.
'   Returns 0 if the book name or chapter is invalid.
'-----------------------------------------------------
Private Function VersesInChapter(ByVal bookName As String, ByVal Chapter As Long) As Long
    On Error GoTo PROC_ERR

    Dim aMap As Object
    Dim key As String
    Dim BookID As Long
    Dim maxCh As Long

    key = UCase$(Trim$(bookName))
    Set aMap = GetBookAliasMap

    If Not aMap.Exists(key) Then
        VersesInChapter = 0
        GoTo PROC_EXIT
    End If
    BookID = aMap(key)
    maxCh = GetMaxChapter(BookID)

    If Chapter < 1 Or Chapter > maxCh Then
        VersesInChapter = 0
        GoTo PROC_EXIT
    End If
    VersesInChapter = GetMaxVerse(BookID, Chapter)

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure VersesInChapter of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

'-----------------------------------------------------
' ParseCanonicalRange
'   Parses a range string of the form:
'       "BookName startCh:startV-endCh:endV"
'   into its component parts.
'   Returns True if successfully parsed as a range.
'   Returns False for single refs, chapter-only refs,
'   or strings that cannot be parsed.
'-----------------------------------------------------
Private Function ParseCanonicalRange( _
        ByVal s As String, _
        ByRef bookName As String, _
        ByRef startCh As Long, ByRef startV As Long, _
        ByRef endCh As Long, ByRef endV As Long _
    ) As Boolean
    On Error GoTo PROC_ERR

    bookName = ""
    startCh = 0
    startV = 0
    endCh = 0
    endV = 0

    s = Trim$(s)
    If Len(s) = 0 Then GoTo PROC_EXIT

    '------------------------------------------
    ' Find the last space to isolate the
    ' reference token "startCh:startV-endCh:endV"
    '------------------------------------------
    Dim lastSpace As Long
    Dim k As Long
    lastSpace = 0
    For k = Len(s) To 1 Step -1
        If Mid$(s, k, 1) = " " Then
            lastSpace = k
            Exit For
        End If
    Next k

    If lastSpace = 0 Then GoTo PROC_EXIT

    Dim refToken As String
    refToken = Mid$(s, lastSpace + 1)
    bookName = Left$(s, lastSpace - 1)
    '------------------------------------------
    ' The reference token must contain a dash
    ' separating start ch:v from end ch:v.
    ' Find the dash that follows a digit.
    '------------------------------------------
    Dim dashPos As Long
    dashPos = 0
    Dim c As String
    For k = 2 To Len(refToken)
        c = Mid$(refToken, k, 1)
        If c = "-" Then
            ' Verify the preceding character is a digit
            If Mid$(refToken, k - 1, 1) >= "0" And Mid$(refToken, k - 1, 1) <= "9" Then
                dashPos = k
                Exit For
            End If
        End If
    Next k

    If dashPos = 0 Then GoTo PROC_EXIT

    Dim leftPart As String
    Dim rightPart As String
    leftPart = Left$(refToken, dashPos - 1)
    rightPart = Mid$(refToken, dashPos + 1)
    '------------------------------------------
    ' Parse left side: startCh:startV
    '------------------------------------------
    Dim colonLeft As Long
    colonLeft = InStr(leftPart, ":")
    If colonLeft = 0 Then GoTo PROC_EXIT

    Dim chStrL As String
    Dim vStrL As String
    chStrL = Left$(leftPart, colonLeft - 1)
    vStrL = Mid$(leftPart, colonLeft + 1)
    If Not IsNumeric(chStrL) Or Not IsNumeric(vStrL) Then GoTo PROC_EXIT
    startCh = CLng(chStrL)
    startV = CLng(vStrL)
    '------------------------------------------
    ' Parse right side: endCh:endV
    '------------------------------------------
    Dim colonRight As Long
    colonRight = InStr(rightPart, ":")
    If colonRight = 0 Then GoTo PROC_EXIT

    Dim chStrR As String
    Dim vStrR As String
    chStrR = Left$(rightPart, colonRight - 1)
    vStrR = Mid$(rightPart, colonRight + 1)
    If Not IsNumeric(chStrR) Or Not IsNumeric(vStrR) Then GoTo PROC_EXIT
    endCh = CLng(chStrR)
    endV = CLng(vStrR)

    ParseCanonicalRange = True
PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ParseCanonicalRange of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

'=====================================================
' ValidateCanonical
' Stage 15
'   Validates a Collection of canonical reference
'   strings against scripture metadata.
'   Invalid references are removed.
'   Range boundaries are clamped to valid limits.
' Input:
'   Refs -  Collection of canonical strings from Stage 14
' Output:
'   New Collection with only valid references.
'=====================================================
Public Function ValidateCanonical(ByVal Refs As Collection) As Collection
    On Error GoTo PROC_ERR

    Dim Result As New Collection
    Dim Item As Variant
    Dim s As String

    For Each Item In Refs
        s = CStr(Item)

        '------------------------------------------
        ' Attempt to parse as a range first
        '------------------------------------------
        Dim bkName As String
        Dim sCh As Long, sV As Long
        Dim eCh As Long, eV As Long

        If ParseCanonicalRange(s, bkName, sCh, sV, eCh, eV) Then
            '------------------------------------------
            ' It is a range: validate and clamp
            '------------------------------------------
            Dim maxChBook As Long
            maxChBook = ChaptersInBook(bkName)

            ' Start chapter invalid -> remove
            If sCh < 1 Or sCh > maxChBook Then GoTo NextItem

            ' Start verse invalid -> remove
            Dim maxVStart As Long
            maxVStart = VersesInChapter(bkName, sCh)
            If sV < 1 Or sV > maxVStart Then GoTo NextItem

            ' Clamp end chapter
            If eCh > maxChBook Then
                eCh = maxChBook
            End If

            ' Clamp end verse
            Dim maxVEnd As Long
            maxVEnd = VersesInChapter(bkName, eCh)
            If eV > maxVEnd Then
                eV = maxVEnd
            End If

            ' After clamping, check end >= start canonically
            ' (same chapter: verse comparison; cross-chapter: chapter wins)
            Dim endBeforeStart As Boolean
            If eCh < sCh Then
                endBeforeStart = True
            ElseIf eCh = sCh And eV < sV Then
                endBeforeStart = True
            Else
                endBeforeStart = False
            End If

            If endBeforeStart Then GoTo NextItem

            ' Check if clamped range collapses to single ref
            If eCh = sCh And eV = sV Then
                ' Emit as single ref
                Result.Add bkName & " " & sCh & ":" & sV
            Else
                ' Emit clamped range
                Result.Add bkName & " " & sCh & ":" & sV & "-" & eCh & ":" & eV
            End If

        Else
            '------------------------------------------
            ' Not a range - try to parse as single ref
            '------------------------------------------
            Dim singleBook As String
            Dim singleCh As Long
            Dim singleV As Long
            Dim isParsed As Boolean

            isParsed = ParseCanonicalRef(s, singleBook, singleCh, singleV)

            If Not isParsed Then
                ' Chapter-only ref or unparseable - pass through unchanged
                Result.Add s
            Else
                ' Single ch:v ref ? validate
                Dim maxChSingle As Long
                maxChSingle = ChaptersInBook(singleBook)

                If maxChSingle = 0 Then GoTo NextItem
                If singleCh < 1 Or singleCh > maxChSingle Then GoTo NextItem

                Dim maxVSingle As Long
                maxVSingle = VersesInChapter(singleBook, singleCh)
                If singleV < 1 Or singleV > maxVSingle Then GoTo NextItem

                ' Valid - keep as-is
                Result.Add s
            End If
        End If

NextItem:
    Next Item

    Set ValidateCanonical = Result
PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ValidateCanonical of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

'=====================================================
' BuildCanonicalRanges
' Stage 16
'   Groups adjacent consecutive same-chapter verses
'   from a validated verse-level canonical reference
'   list into contiguous canonical ranges.
' Input:
'   Refs - Ordered Collection of validated verse-level
'          canonical strings from Stage 15.
'          Each item is of the form "BookName ch:v".
' Output:
'   New Collection of canonical range strings.
'   Adjacent verses within same book and chapter are
'   grouped; non-adjacent, cross-chapter, and
'   cross-book verses remain as separate Items.
'=====================================================
Public Function BuildCanonicalRanges(ByVal Refs As Collection) As Collection
    On Error GoTo PROC_ERR

    Dim Result As New Collection
    Dim i As Long
    Dim n As Long

    n = Refs.Count
    If n = 0 Then
        Set BuildCanonicalRanges = Result
        GoTo PROC_EXIT
    End If
    '------------------------------------------
    ' Parse all refs into parallel arrays
    '------------------------------------------
    Dim bookNames()  As String
    Dim chapters()   As Long
    Dim verses()     As Long
    Dim parsedOk()   As Boolean

    ReDim bookNames(1 To n)
    ReDim chapters(1 To n)
    ReDim verses(1 To n)
    ReDim parsedOk(1 To n)

    Dim refStr As String
    For i = 1 To n
        refStr = Refs(i)
        parsedOk(i) = ParseCanonicalRef( _
                          refStr, _
                          bookNames(i), _
                          chapters(i), _
                          verses(i))
    Next i
    '------------------------------------------
    ' Walk through refs, accumulating runs
    '------------------------------------------
    i = 1
    Do While i <= n
        '------------------------------------------
        ' If this ref could not be parsed as a
        ' chapter:verse ref, emit it unchanged
        '------------------------------------------
        If Not parsedOk(i) Then
            Result.Add Refs(i)
            i = i + 1
        Else
            '------------------------------------------
            ' Start a potential run
            '------------------------------------------
            Dim runStart As Long
            runStart = i
            '------------------------------------------
            ' Extend the run while adjacent.
            ' Note: VBA And is not short-circuit - guard
            ' the array accesses with a Do While i < n
            ' outer check and inner If/Exit Do tests.
            '------------------------------------------
            Do While i < n
                If Not parsedOk(i + 1) Then Exit Do
                If bookNames(i + 1) <> bookNames(i) Then Exit Do
                If chapters(i + 1) <> chapters(i) Then Exit Do
                If verses(i + 1) <> verses(i) + 1 Then Exit Do
                i = i + 1
            Loop
            '------------------------------------------
            ' Emit the run as a range or single ref
            '------------------------------------------
            If i > runStart Then
                Result.Add bookNames(runStart) & " " & _
                           chapters(runStart) & ":" & verses(runStart) & _
                           "-" & _
                           chapters(i) & ":" & verses(i)
            Else
                Result.Add Refs(i)
            End If

            i = i + 1
        End If
    Loop

    Set BuildCanonicalRanges = Result

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure BuildCanonicalRanges of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

'=====================================================
' FormatCanonicalString
' Stage 17
'   Converts a validated canonical range collection
'   into a single formatted SBL-style reference string.
' Input:
'   Refs - Collection of canonical range strings
'          from Stage 16 (or Stage 15).
'          Items are "BookName ch:v" or
'          "BookName ch:v-ch:v".
' Output:
'   Single formatted SBL reference string.
'   Book names suppressed after first occurrence per
'   contiguous group.  Chapter numbers suppressed when
'   same as previous item.  Commas separate same-chapter
'   Items; semicolons separate chapter and book changes.
'=====================================================
Public Function FormatCanonicalString(ByVal Refs As Collection) As String
    On Error GoTo PROC_ERR

    If Refs.Count = 0 Then
        FormatCanonicalString = ""
        GoTo PROC_EXIT
    End If

    Dim Result    As String
    Dim prevBook  As String
    Dim prevEndCh As Long
    Dim Item      As Variant
    Dim s         As String
    Dim bk        As String
    Dim sCh       As Long
    Dim sV        As Long
    Dim eCh       As Long
    Dim eV        As Long
    Dim ch        As Long
    Dim v         As Long
    Dim IsRange   As Boolean
    Dim sep       As String
    Dim showBook  As Boolean
    Dim showCh    As Boolean
    Dim fragment  As String

    Result = ""
    prevBook = ""
    prevEndCh = 0

    For Each Item In Refs
        s = CStr(Item)
        bk = ""
        sCh = 0: sV = 0: eCh = 0: eV = 0

        IsRange = ParseCanonicalRange(s, bk, sCh, sV, eCh, eV)

        If Not IsRange Then
            ch = 0: v = 0
            If ParseCanonicalRef(s, bk, ch, v) Then
                sCh = ch: sV = v
                eCh = ch: eV = v
            Else
                ' Unparseable (chapter-only ref) - pass through
                If Len(Result) > 0 Then Result = Result & "; "
                Result = Result & s
                GoTo NextItem
            End If
        End If

        ' Determine separator and what context to suppress
        If Len(prevBook) = 0 Then
            sep = ""
            showBook = True
            showCh = True
        ElseIf bk <> prevBook Then
            sep = "; "
            showBook = True
            showCh = True
        ElseIf sCh <> prevEndCh Then
            sep = "; "
            showBook = False
            showCh = True
        Else
            sep = ", "
            showBook = False
            showCh = False
        End If

        ' Build the reference fragment
        fragment = ""
        If showBook Then fragment = bk & " "
        If showCh Then
            fragment = fragment & sCh & ":" & sV
        Else
            fragment = fragment & sV
        End If

        ' Append range end when this item spans more than one verse
        If eCh <> sCh Or eV <> sV Then
            If eCh = sCh Then
                ' Same-chapter range: suppress repeated chapter
                fragment = fragment & "-" & eV
            Else
                ' Cross-chapter range: show full end ch:v
                fragment = fragment & "-" & eCh & ":" & eV
            End If
        End If

        Result = Result & sep & fragment

        prevBook = bk
        prevEndCh = eCh

NextItem:
    Next Item

    FormatCanonicalString = Result

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FormatCanonicalString of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

'=====================================================
' NormalizeRawInput  (Public)
' Strips CR/LF, collapses multiple spaces, and
' replaces en-dash (U+2013) with ASCII hyphen.
' Called at the entry of ComposeList, ParseScripture,
' and ParseCitationBlock before any parsing begins.
'=====================================================
Public Function NormalizeRawInput(ByVal raw As String) As String
    On Error GoTo PROC_ERR
    Dim s As String
    s = raw
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, ChrW(8211), "-")   ' en-dash -> ASCII hyphen
    s = Replace(s, Chr(11), " ")   ' Word forced line break (Shift+Enter)
    Dim prev As String
    Do
        prev = s
        s = Replace(s, "  ", " ")
    Loop While s <> prev
    NormalizeRawInput = Trim$(s)
PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure NormalizeRawInput of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

'=====================================================
' ParseCitationBlock  (Public)
' Two-level parser for study Bible citation blocks.
'   Outer split: ";" — major segments; book context
'                      propagates across segments.
'   Inner split: "," — verse sub-Items within segment.
' Calls NormalizeRawInput first (strips line breaks,
' normalizes en-dash to ASCII hyphen).
' Raises on unresolved book alias.
' Returns a Collection of canonical reference strings.
' Example input:
'   "Ps 19:1-2; 23:1; 103:8-11; 145:8-9,17; Isa 40:28"
' Example output Collection:
'   "Psalms 19:1-2", "Psalms 23:1", "Psalms 103:8-11",
'   "Psalms 145:8-9", "Psalms 145:17", "Isaiah 40:28"
'=====================================================
Public Function ParseCitationBlock(ByVal rawBlock As String) As Collection
    On Error GoTo PROC_ERR
    Dim normalized As String
    normalized = NormalizeRawInput(rawBlock)

    If Len(normalized) > 10000 Then
        Err.Raise vbObjectError + 1002, , "Citation block too long (max 10000 chars)"
    End If

    Dim Result As New Collection
    Dim segs() As String
    segs = Split(normalized, ";")

    Dim ctxBookID  As Long:   ctxBookID = 0
    Dim ctxCanon   As String: ctxCanon = ""
    Dim ctxChapter As Long:   ctxChapter = 0

    Dim si As Long
    For si = LBound(segs) To UBound(segs)
        Dim seg As String
        seg = Trim$(segs(si))
        If seg = "" Then GoTo NEXT_SEG

        Dim parts() As String
        parts = Split(seg, " ")
        Dim p0 As String
        p0 = parts(0)
        Dim refPart As String
        refPart = ""
        Dim newBook As Boolean
        newBook = False

        '------------------------------------------
        ' Detect book alias
        ' Case 1: single digit -> two-token alias
        '------------------------------------------
        If Len(p0) = 1 And p0 >= "1" And p0 <= "9" Then
            If UBound(parts) >= 1 Then
                Dim tt As String
                tt = p0 & " " & parts(1)
                Dim ttID As Long
                Dim ttCan As String
                On Error Resume Next
                ttCan = ResolveAlias(tt, ttID)
                On Error GoTo PROC_ERR
                If ttCan <> "" Then
                    ctxBookID = ttID
                    ctxCanon = ttCan
                    newBook = True
                    If Len(seg) > Len(tt) + 1 Then
                        refPart = Trim$(Mid$(seg, Len(tt) + 2))
                    End If
                End If
            End If
        End If

        '------------------------------------------
        ' Case 2: letter prefix -> single-token alias
        '------------------------------------------
        If Not newBook And p0 Like "[A-Za-z]*" Then
            Dim stID As Long
            Dim stCan As String
            On Error Resume Next
            stCan = ResolveAlias(p0, stID)
            On Error GoTo PROC_ERR
            If stCan <> "" Then
                ctxBookID = stID
                ctxCanon = stCan
                newBook = True
                If Len(seg) > Len(p0) + 1 Then
                    refPart = Trim$(Mid$(seg, Len(p0) + 2))
                End If
            Else
                Err.Raise vbObjectError + 1001, , _
                    "Cannot resolve alias: """ & p0 & """"
            End If
        End If

        '------------------------------------------
        ' Case 3: no new book — bare chapter:verse
        '------------------------------------------
        If Not newBook Then refPart = seg

        ' Non-ASCII token: not a digit, not ASCII alpha — raise descriptive error
        ' rather than silently misattributing the reference (Case 3 fall-through).
        If Not newBook And Len(p0) > 0 Then
            If AscW(Left$(p0, 1)) > 127 Then
                Err.Raise vbObjectError + 1003, , _
                    "Unrecognised token (non-ASCII character): """ & p0 & """"
            End If
        End If

        ' Reset chapter context on every book switch so that single-chapter
        ' book shorthand (e.g. "Jude 6") is not contaminated by the previous
        ' book's chapter number.
        If newBook Then ctxChapter = 0

        '------------------------------------------
        ' Parse chapter from refPart
        '------------------------------------------
        Dim colonPos As Long
        colonPos = InStr(refPart, ":")
        Dim vsStr As String
        If colonPos > 0 Then
            Dim chStr As String
            chStr = Trim$(Left$(refPart, colonPos - 1))
            vsStr = Trim$(Mid$(refPart, colonPos + 1))
            If IsNumeric(chStr) Then ctxChapter = CLng(chStr)
        Else
            vsStr = Trim$(refPart)
        End If

        ' Single-chapter book with no colon (e.g. "Jude 6"): promote chapter to 1
        ' so the canonical string is "Jude 1:6" not "Jude 0:6".
        If ctxChapter = 0 And GetMaxChapter(ctxBookID) = 1 Then ctxChapter = 1

        ' Whole-chapter reference on multi-chapter book (e.g. "Ezek 16"):
        ' no colon, ctxChapter still 0 after book switch.
        ' Emit "Ezekiel 16" (no verse) and move to next segment.
        If ctxChapter = 0 And colonPos = 0 And IsNumeric(vsStr) Then
            ctxChapter = CLng(vsStr)
            Result.Add ctxCanon & " " & ctxChapter
            GoTo NEXT_SEG
        End If

        '------------------------------------------
        ' Split verse spec on "," for "8-9,17"
        '------------------------------------------
        Dim vss() As String
        vss = Split(vsStr, ",")

        Dim vi As Long
        For vi = LBound(vss) To UBound(vss)
            Dim vsRaw As String
            vsRaw = Trim$(vss(vi))
            If vsRaw = "" Then GoTo NEXT_VS

            Dim sV As Long, eV As Long, isRng As Boolean
            Dim dp As Long
            dp = InStr(vsRaw, "-")
            If dp > 0 Then
                If IsNumeric(Left$(vsRaw, dp - 1)) And IsNumeric(Mid$(vsRaw, dp + 1)) Then
                    sV = CLng(Left$(vsRaw, dp - 1))
                    eV = CLng(Mid$(vsRaw, dp + 1))
                    isRng = True
                Else
                    ' malformed range (e.g. ":-11") — sentinel verse 0 produces FAIL
                    Result.Add ctxCanon & " " & ctxChapter & ":0"
                    GoTo NEXT_VS
                End If
            ElseIf IsNumeric(vsRaw) Then
                sV = CLng(vsRaw)
                eV = sV
                isRng = False
            Else
                ' malformed verse spec — sentinel verse 0 produces FAIL
                Result.Add ctxCanon & " " & ctxChapter & ":0"
                GoTo NEXT_VS
            End If

            Dim canon As String
            canon = ctxCanon & " " & ctxChapter & ":" & sV
            If isRng Then canon = canon & "-" & eV
            Result.Add canon
NEXT_VS:
        Next vi

NEXT_SEG:
    Next si

    Set ParseCitationBlock = Result
PROC_EXIT:
    Exit Function
PROC_ERR:
    ' User-input errors (non-ASCII token, block too long): re-raise so the caller
    ' (VerifyCitationBlockReport) can set failCount = -1 and show a friendly message.
    If Err.Number = vbObjectError + 1002 Or Err.Number = vbObjectError + 1003 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Debug.Print "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ParseCitationBlock of Class aeSBL_Citation_Class"
    'MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ParseCitationBlock of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

'=====================================================
' RenderEnDash  (Public)
' Stage 17 option: replaces ASCII hyphen with en-dash
' in a canonical reference string for Study Bible
' document output. Safe as a simple Replace because
' SBL canonical book names never contain hyphens;
' the only hyphens in a canonical string are verse
' range separators.
' Example: "Psalms 103:8-11" -> "Psalms 103:8" & ChrW(8211) & "11"
'=====================================================
Public Function RenderEnDash(ByVal canon As String) As String
    On Error GoTo PROC_ERR
    RenderEnDash = Replace(canon, "-", ChrW(8211))
PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RenderEnDash of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

'=====================================================
' SortCitationBlock  (Public)
' Stage 13b: sorts a Collection of canonical reference
' strings by (BookID, Chapter, StartVerse) — canonical
' book order Gen=1 through Rev=66. Required before
' Stage 14 Canonical Compression, which assumes
' adjacency ordering.
' Input: Collection of strings from ParseCitationBlock.
' Returns a new Collection in canonical order.
' Sort key: BookID * 1000000 + Chapter * 10000 + StartVerse
' Long max ~2.1B; max key = 66*1000000 + 150*10000 + 176 = 67,500,176 (safe)
'=====================================================
Public Function SortCitationBlock(ByVal Items As Collection) As Collection
    On Error GoTo PROC_ERR

    Dim n As Long
    n = Items.Count
    If n = 0 Then Set SortCitationBlock = New Collection: GoTo PROC_EXIT

    Dim keys() As Long
    Dim vals() As String
    ReDim keys(1 To n)
    ReDim vals(1 To n)

    Dim i As Long
    i = 1
    Dim Item As Variant
    For Each Item In Items
        Dim canon As String
        canon = CStr(Item)

        ' Find last space: left = book name, right = Chapter:Verse[-End]
        Dim lastSp As Long
        lastSp = 0
        Dim k As Long
        For k = Len(canon) To 1 Step -1
            If Mid$(canon, k, 1) = " " Then lastSp = k: Exit For
        Next k

        Dim bookName As String
        Dim numPart As String
        bookName = Left$(canon, lastSp - 1)
        numPart = Mid$(canon, lastSp + 1)

        ' Resolve BookID
        Dim bID As Long
        bID = 0
        On Error Resume Next
        ResolveAlias bookName, bID
        On Error GoTo PROC_ERR

        ' Parse Chapter and StartVerse from numPart
        Dim cpPos As Long
        cpPos = InStr(numPart, ":")
        Dim ch As Long
        Dim vPart As String
        Dim sV As Long
        If cpPos > 0 Then
            ch = CLng(Left$(numPart, cpPos - 1))
            vPart = Mid$(numPart, cpPos + 1)
            Dim dpPos As Long
            dpPos = InStr(vPart, "-")
            If dpPos > 0 Then
                sV = CLng(Left$(vPart, dpPos - 1))
            Else
                sV = CLng(vPart)
            End If
        Else
            ' Whole-chapter reference (no colon): numPart is the chapter
            ch = CLng(numPart)
            sV = 0
        End If

        keys(i) = bID * 1000000 + ch * 10000 + sV
        vals(i) = canon
        i = i + 1
    Next Item

    ' Insertion sort on keys/vals arrays
    Dim j As Long
    Dim tmpKey As Long
    Dim tmpVal As String
    For i = 2 To n
        tmpKey = keys(i)
        tmpVal = vals(i)
        j = i - 1
        Do While j >= 1
            If keys(j) <= tmpKey Then Exit Do
            keys(j + 1) = keys(j)
            vals(j + 1) = vals(j)
            j = j - 1
        Loop
        keys(j + 1) = tmpKey
        vals(j + 1) = tmpVal
    Next i

    ' Build result Collection
    Dim Result As New Collection
    For i = 1 To n
        Result.Add vals(i)
    Next i
    Set SortCitationBlock = Result

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure SortCitationBlock of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

'=====================================================
' ToSBLShortForm  (Public)
' Converts a canonical reference string (full book name)
' to SBL abbreviated form.
' Example: "Psalms 19:1-2" -> "Ps 19:1-2"
'          "1 Chronicles 29:10-13" -> "1 Chr 29:10-13"
' If the BookID cannot be resolved the canonical string
' is returned unchanged.
'=====================================================
Public Function ToSBLShortForm(ByVal canon As String) As String
    On Error GoTo PROC_ERR

    ' Find last space: left = book name, right = Chapter:Verse[-End]
    Dim lastSp As Long
    Dim k As Long
    For k = Len(canon) To 1 Step -1
        If Mid$(canon, k, 1) = " " Then lastSp = k: Exit For
    Next k
    If lastSp = 0 Then ToSBLShortForm = canon: GoTo PROC_EXIT

    Dim bookName As String
    Dim numPart As String
    bookName = Left$(canon, lastSp - 1)
    numPart = Mid$(canon, lastSp + 1)

    Dim bID As Long
    On Error Resume Next
    ResolveAlias bookName, bID
    On Error GoTo PROC_ERR

    Dim abbr As String
    Select Case bID
        Case 1:  abbr = "Gen"
        Case 2:  abbr = "Exod"
        Case 3:  abbr = "Lev"
        Case 4:  abbr = "Num"
        Case 5:  abbr = "Deut"
        Case 6:  abbr = "Josh"
        Case 7:  abbr = "Judg"
        Case 8:  abbr = "Ruth"
        Case 9:  abbr = "1 Sam"
        Case 10: abbr = "2 Sam"
        Case 11: abbr = "1 Kgs"
        Case 12: abbr = "2 Kgs"
        Case 13: abbr = "1 Chr"
        Case 14: abbr = "2 Chr"
        Case 15: abbr = "Ezra"
        Case 16: abbr = "Neh"
        Case 17: abbr = "Esth"
        Case 18: abbr = "Job"
        Case 19: abbr = "Ps"
        Case 20: abbr = "Prov"
        Case 21: abbr = "Eccl"
        Case 22: abbr = "Song"
        Case 23: abbr = "Isa"
        Case 24: abbr = "Jer"
        Case 25: abbr = "Lam"
        Case 26: abbr = "Ezek"
        Case 27: abbr = "Dan"
        Case 28: abbr = "Hos"
        Case 29: abbr = "Joel"
        Case 30: abbr = "Amos"
        Case 31: abbr = "Obad"
        Case 32: abbr = "Jonah"
        Case 33: abbr = "Mic"
        Case 34: abbr = "Nah"
        Case 35: abbr = "Hab"
        Case 36: abbr = "Zeph"
        Case 37: abbr = "Hag"
        Case 38: abbr = "Zech"
        Case 39: abbr = "Mal"
        Case 40: abbr = "Matt"
        Case 41: abbr = "Mark"
        Case 42: abbr = "Luke"
        Case 43: abbr = "John"
        Case 44: abbr = "Acts"
        Case 45: abbr = "Rom"
        Case 46: abbr = "1 Cor"
        Case 47: abbr = "2 Cor"
        Case 48: abbr = "Gal"
        Case 49: abbr = "Eph"
        Case 50: abbr = "Phil"
        Case 51: abbr = "Col"
        Case 52: abbr = "1 Thess"
        Case 53: abbr = "2 Thess"
        Case 54: abbr = "1 Tim"
        Case 55: abbr = "2 Tim"
        Case 56: abbr = "Titus"
        Case 57: abbr = "Phlm"
        Case 58: abbr = "Heb"
        Case 59: abbr = "Jas"
        Case 60: abbr = "1 Pet"
        Case 61: abbr = "2 Pet"
        Case 62: abbr = "1 John"
        Case 63: abbr = "2 John"
        Case 64: abbr = "3 John"
        Case 65: abbr = "Jude"
        Case 66: abbr = "Rev"
        Case Else: abbr = bookName
    End Select

    ' SBL shorthand for single-chapter books omits the chapter number.
    ' Canonical form is "Jude 1:6"; SBL output is "Jude 6".
    If GetMaxChapter(bID) = 1 Then
        Dim cpPos As Long
        cpPos = InStr(numPart, ":")
        If cpPos > 0 Then numPart = Mid$(numPart, cpPos + 1)
    End If

    ToSBLShortForm = abbr & " " & numPart
PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ToSBLShortForm of Class aeBibleCitationClass"
    Resume PROC_EXIT
End Function

Public Sub Test_AllBookAliases_STRICT()
    On Error GoTo PROC_ERR
    Dim aMap As Object
    Dim books As Object
    Dim k As Variant
    Dim BookID As Long
    Dim canonicalActual As String
    Dim failures As Long

    Set aMap = GetBookAliasMap
    Set books = GetCanonicalBookTable

    Debug.Print "=== STRICT Alias Validation ==="

    For Each k In aMap.Keys
        On Error GoTo AliasFail

        canonicalActual = ResolveAlias(CStr(k), BookID)

        ' 1. BookID must be in canonical range
        If BookID < 1 Or BookID > 66 Then
            Debug.Print "FAIL (INVALID BOOK ID): "; k; " ? "; BookID
            failures = failures + 1
            GoTo NextAlias
        End If

        ' 2. BookID must exist in canonical table
        If Not books.Exists(BookID) Then
            Debug.Print "FAIL (MISSING CANONICAL): "; k; " ? "; BookID
            failures = failures + 1
            GoTo NextAlias
        End If

        ' 3. Canonical name must match
        If canonicalActual <> books(BookID)(1) Then
            Debug.Print "FAIL (NAME MISMATCH): "; k; _
                        " ? "; canonicalActual; _
                        " (expected "; books(BookID)(1); ")"
            failures = failures + 1
        End If

NextAlias:
        On Error GoTo PROC_ERR
    Next k

    If failures = 0 Then
        Debug.Print "PASS: All aliases are canonical and valid."
    Else
        Debug.Print "FAILURES: "; failures
    End If
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_AllBookAliases_STRICT of Module aeBibleCitationClass"
    Resume PROC_EXIT

AliasFail:
    Debug.Print "FAIL (ERROR): "; k; " > "; Err.Description
    Err.Clear
    failures = failures + 1
    Resume NextAlias
End Sub


Private Function ParseReferenceStub(ByVal inputText As String) As ParsedReference
    On Error GoTo PROC_ERR
    Dim p As ParsedReference
    p.rawInput = inputText

    Debug.Print "  [Stub] Raw input = >" & inputText & "<"

    Dim parts() As String
    parts = Split(Trim$(inputText), " ")
    '----------------------------------
    ' Book alias (first token)
    '----------------------------------
    p.BookAlias = UCase$(parts(0))
    Debug.Print "  [Stub] Parsed alias = >" & p.BookAlias & "<"
    ' Semantic guard to check for null failures
    Debug.Assert p.BookAlias <> vbNullString
    '----------------------------------
    ' No chapter/verse provided
    '----------------------------------
    If UBound(parts) = 0 Then
        p.Chapter = 0
        p.VerseSpec = ""
        ParseReferenceStub = p
        Exit Function
    End If
    '----------------------------------
    ' Chapter or Chapter:Verse
    '----------------------------------
    Dim refPart As String
    refPart = parts(1)

    If InStr(refPart, ":") > 0 Then
        Dim cvParts() As String
        cvParts = Split(refPart, ":")
        p.Chapter = CLng(cvParts(0))
        p.VerseSpec = cvParts(1)
    Else
        ' Single-chapter book case (e.g. Jude 5)
        p.Chapter = 0
        p.VerseSpec = refPart
    End If

    ParseReferenceStub = p
PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ParseReferenceStub of Class aeBibleCitationClass"
    Resume PROC_EXIT
End Function


Public Sub Test_SemanticFlow_WithParserStub()
    On Error GoTo PROC_ERR
    ResetBookAliasMap

    Debug.Print "======================================"
    Debug.Print " Test_SemanticFlow_WithParserStub"
    Debug.Print " (Uses lightweight parser stub)"
    Debug.Print "======================================"

    Dim failures As Long
    failures = 0

    Dim tests
    tests = Array( _
        Array("Jude 5", 65, True, True), _
        Array("Jude 1:5", 65, True, True), _
        Array("Obadiah 3", 31, True, True), _
        Array("Romans 8:1", 45, True, False), _
        Array("Genesis 1:1", 1, True, False), _
        Array("Jude 5-7", 65, False, False) _
    )

    Dim i As Long
    For i = LBound(tests) To UBound(tests)
        Debug.Print ""
        Debug.Print "INPUT: "; tests(i)(0)

        Dim testFailed As Boolean
        testFailed = False

        Dim parsed As ParsedReference
        parsed = ParseReferenceStub(tests(i)(0))
        Debug.Print "  ParseReferenceStub:"
        Debug.Print "    -> Alias:   "; parsed.BookAlias
        Debug.Print "    -> Chapter: "; parsed.Chapter
        Debug.Print "    -> Verse:   "; parsed.VerseSpec

        Dim bookName As String
        Dim BookID As Long

        On Error Resume Next
        bookName = ResolveAlias(parsed.BookAlias, BookID)
        Dim errNum As Long
        errNum = Err.Number
        Err.Clear
        On Error GoTo 0

        If errNum <> 0 Then
            Debug.Print "  ERROR: ResolveBook failed"
            failures = failures + 1
            GoTo NextTest
        End If
        Debug.Print "  Resolver:"
        Debug.Print "    -> BookID:    "; BookID
        Debug.Print "    -> Canonical: "; bookName
        If BookID <> tests(i)(1) Then
            Debug.Print "  FAIL: BookID mismatch"
            testFailed = True
        End If

        Dim semanticMsg As String
        Dim IsValid As Boolean

        IsValid = ValidateSBLReference( _
                    BookID, _
                    bookName, _
                    parsed.Chapter, _
                    parsed.VerseSpec, _
                    ModeSBL)
        Debug.Print "  ValidateSBLReference:"
        Debug.Print "    -> Valid: "; IsValid
        If IsValid <> tests(i)(2) Then
            Debug.Print "  FAIL: semantic validity mismatch"
            testFailed = True
        End If

        If IsValid Then
            If IsNumeric(parsed.VerseSpec) Then
                Dim rewritten As String
                rewritten = RewriteSingleChapterRef( _
                                BookID, _
                                parsed.Chapter, _
                                CLng(parsed.VerseSpec))
                Debug.Print "  Output: "; rewritten
                If tests(i)(3) Then
                    If Left$(rewritten, 2) <> "1:" Then
                        Debug.Print "  FAIL: expected single-chapter rewrite"
                        testFailed = True
                    End If
                Else
                    If rewritten <> parsed.Chapter & ":" & parsed.VerseSpec Then
                        Debug.Print "  FAIL: unexpected rewrite"
                        testFailed = True
                    End If
                End If
            Else
                Debug.Print "  Skipped: RewriteSingleChapterRef (VerseSpec is non-numeric: >" & parsed.VerseSpec & "<)"
            End If
        End If
        If testFailed Then
            failures = failures + 1
        Else
            Debug.Print "  RESULT: PASS"
        End If

NextTest:
    Next i

    Debug.Print ""
    Debug.Print "======================================"
    Debug.Print "FAILURES: "; failures
    Debug.Print "======================================"
    Report_TODOs
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_SemanticFlow_WithParserStub of Class aeBibleCitationClass"
    Resume PROC_EXIT
End Sub


Public Sub Test_SemanticFlow_WithParserStub_Negative()
    On Error GoTo PROC_ERR
    ResetBookAliasMap

    Debug.Print "=========================================="
    Debug.Print " Test_SemanticFlow_WithParserStub_Negative"
    Debug.Print "=========================================="

    Dim failReason As String
    failReason = ""

    Dim failures As Long
    failures = 0

    Dim tests As Variant
    tests = Array( _
        Array("Jude 0", FailNone), _
        Array("Jude 999", FailNone), _
        Array("Jude 1:0", FailNone), _
        Array("Romans 0:1", FailNone), _
        Array("Romans 999:1", FailNone), _
        Array("Genesis 1:999", FailNone) _
    )

    Dim i As Long
    For i = LBound(tests) To UBound(tests)

        Debug.Print ""
        Debug.Print "INPUT: "; tests(i)(0)

        Dim parsed As ParsedReference
        parsed = ParseReferenceStub(tests(i)(0))

        Debug.Print "  ParseReferenceStub:"
        Debug.Print "    -> Alias:   "; parsed.BookAlias
        Debug.Print "    -> Chapter: "; parsed.Chapter
        Debug.Print "    -> Verse:   "; parsed.VerseSpec

        Dim bookName As String
        Dim BookID As Long

        On Error Resume Next
        bookName = ResolveAlias(parsed.BookAlias, BookID)

        If Err.Number <> 0 Then
            Debug.Print "  ResolveBook ERROR: "; Err.Description

            If tests(i)(1) = FailResolveBook Then
                Debug.Print "  RESULT: PASS (ResolveBook failed as expected)"
            Else
                Debug.Print "  RESULT: FAIL (ResolveBook failed unexpectedly)"
                failures = failures + 1
            End If

            Err.Clear
            On Error GoTo 0
            GoTo NextTest
        End If

        On Error GoTo 0

        Debug.Print "  Resolver:"
        Debug.Print "    -> BookID:     "; BookID
        Debug.Print "    -> Canonical:  "; bookName

        Dim IsValid As Boolean

        IsValid = ValidateSBLReference( _
                    BookID, _
                    bookName, _
                    parsed.Chapter, _
                    parsed.VerseSpec, _
                    ModeSBL)

        Debug.Print "  ValidateSBLReference:"
        Debug.Print "    -> Valid: "; IsValid

        If IsValid <> tests(i)(1) Then
            Debug.Print "  FAIL: expected validity = "; tests(i)(1)
            failures = failures + 1
        Else
            Debug.Print "  RESULT: PASS"
        End If

NextTest:
    Next i

    Debug.Print ""
    Debug.Print "======================================"
    Debug.Print "FAILURES: "; failures
    Debug.Print "======================================"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_SemanticFlow_WithParserStub_Negative of Class aeBibleCitationClass"
    Resume PROC_EXIT
End Sub


Public Sub Test_Stage2_LexicalScan()
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage2_LexicalScan"
    Debug.Print "------------------------------------------"

    Dim t As LexTokens
    t = LexicalScan("Jude 1:5")
    aeAssert.AssertEqual "Jude", t.RawAlias, "Alias parsed"
    aeAssert.AssertEqual 1, t.Num1, "Chapter parsed"
    aeAssert.AssertEqual 5, t.Num2, "Verse parsed"
    aeAssert.AssertTrue t.HasColon, "Colon detected"
    t = LexicalScan("Romans 8")
    aeAssert.AssertEqual "Romans", t.RawAlias, "Alias parsed"
    aeAssert.AssertEqual 8, t.Num1, "Number parsed"
    aeAssert.AssertTrue Not t.HasColon, "No colon detected"
End Sub


Public Sub Test_Stage3_ResolveAlias()
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage3_ResolveAlias"
    Debug.Print "------------------------------------------"

    Dim tokens As LexTokens
    Dim BookID As Long
    Dim canonical As String

    tokens = LexicalScan("Jude 1:5")
    canonical = ResolveAlias(tokens.RawAlias, BookID)
    aeAssert.AssertEqual 65, BookID, "Jude BookID"
    aeAssert.AssertEqual "Jude", canonical, "Jude canonical"
    tokens = LexicalScan("Genesis 1:1")
    canonical = ResolveAlias(tokens.RawAlias, BookID)
    aeAssert.AssertEqual 1, BookID, "Genesis BookID"
    aeAssert.AssertEqual "Genesis", canonical, "Genesis canonical"
End Sub


Public Sub Test_Stage4_InterpretStructure()
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage4_InterpretStructure"
    Debug.Print "------------------------------------------"

    Dim tokens As LexTokens
    Dim ref As ParsedReference
    '------------------------------------------
    ' Single-chapter book (implicit verse)
    '------------------------------------------
    tokens = LexicalScan("Jude 5")
    ref = InterpretStructure(tokens)
    aeAssert.AssertEqual 0, ref.Chapter, "Jude 5 chapter interpreted"
    aeAssert.AssertEqual "5", ref.VerseSpec, "Jude 5 verse interpreted"
    '------------------------------------------
    ' Standard chapter:verse
    '------------------------------------------
    tokens = LexicalScan("Romans 8:1")
    ref = InterpretStructure(tokens)
    aeAssert.AssertEqual 8, ref.Chapter, "Romans chapter interpreted"
    aeAssert.AssertEqual "1", ref.VerseSpec, "Romans verse interpreted"
    '------------------------------------------
    ' Ambiguous single number (no colon)
    '------------------------------------------
    tokens = LexicalScan("Genesis 1")
    ref = InterpretStructure(tokens)
    aeAssert.AssertEqual 0, ref.Chapter, "Genesis 1 chapter interpreted"
    aeAssert.AssertEqual "1", ref.VerseSpec, "Genesis 1 verse interpreted"
    '------------------------------------------
    ' Verse range
    '------------------------------------------
    tokens = LexicalScan("John 3:16-18")
    ref = InterpretStructure(tokens)
    aeAssert.AssertEqual 3, ref.Chapter, "John chapter interpreted"
    aeAssert.AssertEqual "16-18", ref.VerseSpec, "John verse range interpreted"
    '------------------------------------------
    ' Verse list
    '------------------------------------------
    tokens = LexicalScan("Psalm 23:1,4,6")
    ref = InterpretStructure(tokens)
    aeAssert.AssertEqual 23, ref.Chapter, "Psalm chapter interpreted"
    aeAssert.AssertEqual "1,4,6", ref.VerseSpec, "Psalm verse list interpreted"
    '------------------------------------------
    ' Mixed list and range
    '------------------------------------------
    tokens = LexicalScan("Matthew 5:3-5,9")
    ref = InterpretStructure(tokens)
    aeAssert.AssertEqual 5, ref.Chapter, "Matthew chapter interpreted"
    aeAssert.AssertEqual "3-5,9", ref.VerseSpec, "Matthew mixed verse spec interpreted"
End Sub


Private Sub PrintScriptureList(list As ScriptureList)
    On Error GoTo PROC_ERR
    Dim i As Long
    Dim refIndex As Long
    Dim rangeIndex As Long

    refIndex = 0
    rangeIndex = 0
    For i = LBound(list.ItemType) To UBound(list.ItemType)
        Select Case list.ItemType(i)
            Case 1  ' ScriptureRef
                Debug.Print CanonicalFromRef(list.Refs(refIndex))
                refIndex = refIndex + 1
            Case 2  ' ScriptureRange
                Debug.Print CanonicalFromRef(list.Ranges(rangeIndex).StartRef) & _
                            "-" & _
                            CanonicalFromRef(list.Ranges(rangeIndex).EndRef)
                rangeIndex = rangeIndex + 1
            Case Else
                Debug.Print "Unknown item type at index "; i
        End Select
    Next i
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure PrintScriptureList of Class aeBibleCitationClass"
    Resume PROC_EXIT
End Sub


Public Sub Test_Stage8_ListDetection()
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage8_ListDetection"
    Debug.Print "------------------------------------------"

    Dim t As ListTokens
    '------------------------------------------
    ' Test 1 - comma list
    '------------------------------------------
    t = ListDetection("John 3:16,18,20")
    aeAssert.AssertTrue t.IsList, "comma list detected"
    aeAssert.AssertEqual 2, UBound(t.Segments), "comma list segment count"
    '------------------------------------------
    ' Test 2 - semicolon list
    '------------------------------------------
    t = ListDetection("John 3:16; 4:1")
    aeAssert.AssertTrue t.IsList, "semicolon list detected"
    aeAssert.AssertEqual 1, UBound(t.Segments), "semicolon list segment count"
    '------------------------------------------
    ' Test 3 - single reference
    '------------------------------------------
    t = ListDetection("John 3:16")
    aeAssert.AssertFalse t.IsList, "single reference not list"
    '------------------------------------------
    ' Test 4 - list containing range
    '------------------------------------------
    t = ListDetection("John 3:16-18,20")
    aeAssert.AssertTrue t.IsList, "range preserved inside list"
    aeAssert.AssertEqual 1, UBound(t.Segments), "range list segment count"
    aeAssert.AssertEqual "John 3:16-18", t.Segments(0), "range first segment"
    aeAssert.AssertEqual "20", t.Segments(1), "range second segment"
    '------------------------------------------
    ' Test 5 - mixed whitespace
    '------------------------------------------
    t = ListDetection("John 3:16 , 18 , 20")
    aeAssert.AssertTrue t.IsList, "whitespace tolerated"
    aeAssert.AssertEqual 2, UBound(t.Segments), "whitespace segment count"
End Sub


Public Sub Test_Stage9_RangeDetection()
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage9_RangeDetection"
    Debug.Print "------------------------------------------"

    Dim r As RangeTokens
    '------------------------------------------
    ' Test 1 - verse range
    '------------------------------------------
    aeAssert.AssertTrue IsRangeSegment("John 3:16-18"), _
        "IsRangeSegment verse range"
    r = RangeDetection("John 3:16-18")
    aeAssert.AssertTrue r.IsRange, "verse range detected"
    aeAssert.AssertEqual r.LeftRaw, "John 3:16", "range left token"
    aeAssert.AssertEqual r.RightRaw, "18", "range right token"
    '------------------------------------------
    ' Test 2 - chapter range
    '------------------------------------------
    r = RangeDetection("John 3-5")
    aeAssert.AssertTrue r.IsRange, "chapter range detected"
    aeAssert.AssertEqual r.LeftRaw, "John 3", "chapter range left"
    aeAssert.AssertEqual r.RightRaw, "5", "chapter range right"
    '------------------------------------------
    ' Test 3 - cross chapter range
    '------------------------------------------
    r = RangeDetection("John 3:16-4:2")
    aeAssert.AssertTrue r.IsRange, "cross chapter range detected"
    aeAssert.AssertEqual r.LeftRaw, "John 3:16", "cross chapter left"
    aeAssert.AssertEqual r.RightRaw, "4:2", "cross chapter right"
    '------------------------------------------
    ' Test 4 - en dash
    '------------------------------------------
    r = RangeDetection("John 3:16" & ChrW(8211) & "18")   ' en dash character (U+2013)
    aeAssert.AssertTrue r.IsRange, "en dash range detected"
    aeAssert.AssertEqual r.LeftRaw, "John 3:16", "en dash left"
    aeAssert.AssertEqual r.RightRaw, "18", "en dash right"
    '------------------------------------------
    ' Test 5 - not a range
    '------------------------------------------
    aeAssert.AssertFalse IsRangeSegment("John 3:16"), _
        "IsRangeSegment single reference"
    r = RangeDetection("John 3:16")
    aeAssert.AssertFalse r.IsRange, "single reference not range"
End Sub


Public Sub Test_Stage10_RangeComposition()
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage10_RangeComposition"
    Debug.Print "------------------------------------------"

    Dim r As ScriptureRange
    '------------------------------------------
    ' Verse shorthand
    '------------------------------------------
    r = ComposeRange("John 3:16-18")
    aeAssert.AssertEqual 3, r.StartRef.Chapter, "start chapter"
    aeAssert.AssertEqual 16, r.StartRef.Verse, "start verse"
    aeAssert.AssertEqual 3, r.EndRef.Chapter, "end chapter"
    aeAssert.AssertEqual 18, r.EndRef.Verse, "end verse"
    '------------------------------------------
    ' Chapter shorthand
    '------------------------------------------
    r = ComposeRange("John 3-5")
    aeAssert.AssertEqual 3, r.StartRef.Chapter, "chapter start"
    aeAssert.AssertEqual 5, r.EndRef.Chapter, "chapter end"
    '------------------------------------------
    ' Cross-chapter range
    '------------------------------------------
    r = ComposeRange("Genesis 1:31-2:3")
    aeAssert.AssertEqual 1, r.StartRef.Chapter, "cross start chapter"
    aeAssert.AssertEqual 31, r.StartRef.Verse, "cross start verse"
    aeAssert.AssertEqual 2, r.EndRef.Chapter, "cross end chapter"
    aeAssert.AssertEqual 3, r.EndRef.Verse, "cross end verse"
End Sub

