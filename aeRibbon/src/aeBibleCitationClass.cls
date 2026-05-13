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
        books.Add 22, Array(22, "Song of Songs", 8)   ' SBL output abbreviation is "Song" (see ToSBLShortForm Case 22).
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
        books.Add 34, Array(34, "Nahum", 3)
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
        ' Song of Songs, aka Song of Solomon (SBL output abbreviation is "Song")
        ' Accepts "Song" and variants as input aliases so SBL-typed citations resolve correctly.
        aliasMap.Add "SONG OF SONGS", 22
        aliasMap.Add "SONG", 22
        aliasMap.Add "SON", 22
        aliasMap.Add "SG", 22
        aliasMap.Add "SOLOMON", 22
        aliasMap.Add "SOLO", 22
        aliasMap.Add "SOL", 22
        aliasMap.Add "SO", 22
        aliasMap.Add "CANTICLES", 22
        aliasMap.Add "CANT", 22
        aliasMap.Add "CAN", 22
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
Public Function ChaptersInBook(ByVal bookName As String) As Long
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
Public Function VersesInChapter(ByVal bookName As String, ByVal Chapter As Long) As Long
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
Public Function ToSBLShortForm(ByVal canon As String) As String
    On Error GoTo PROC_ERR

    ' Find last space: left = book name, right = Chapter:Verse[-End]
    Dim lastSp As Long
    Dim k As Long
    For k = Len(canon) To 1 Step -1
        'Debug.Print "> " & Mid$(canon, k, 1)
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

