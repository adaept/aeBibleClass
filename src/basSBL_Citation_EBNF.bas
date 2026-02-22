Attribute VB_Name = "basSBL_Citation_EBNF"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Private aliasMap As Object

'=======================================
' SBL Scripture Citation - Unified EBNF
'=======================================
' Citation
'    ::= WS? Reference (WS? RefSep WS? Reference)* WS?
' Reference
'    ::= BookRef (WS ChapterSpec)?
' BookRef
'    ::= Prefix? WS? BookName
' Prefix
'    ::= ArabicPrefix | RomanPrefix
' ArabicPrefix
'    ::= "1" | "2" | "3"
' RomanPrefix
'    ::= "I" | "II" | "III"
' NOTE: Prefix may be adjacent to BookWord (e.g., "1John", "IJohn")
' BookName
'    ::= BookWord (WS BookWord)*
' BookWord
'    ::= Letter+ ("." )?
' NOTE: BOOK_WORD may include a trailing . but never internal punctuation
' ChapterSpec
'    ::= Chapter
'     | Chapter ":" VerseSpec
'     | ChapterRange
'     | Chapter ":" VerseRangeSpec
' ChapterRange
'    ::= Chapter "-" Chapter
' VerseSpec
'    ::= VerseItem ("," VerseItem)*
' VerseRangeSpec
'    ::= VerseRange ("," VerseRange)*
' VerseItem
'    ::= Verse | VerseRange
' VerseRange
'    ::= Verse "-" Verse
' Verse
'    ::= Digit+ VerseSuffix?
' NOTE: VerseSuffix letters (e.g., "a", "b") are captured
'       during tokenization and validated in post-processing
' VerseSuffix
'    ::= Letter
' Chapter
'    ::= Digit+
' RefSep
'    ::= ";" | ","
' WS
'    ::= " " { " " }
' Letter
'    ::= "A"..."Z" | "a"..."z"
' Digit
'    ::= "0"..."9"
' NOTE: This DFA validates structural syntax only.
'       Semantic constraints are enforced post-parse.
'=====================================================
' Embedded Extension Hooks (Implicit but Intentional)
'=====================================================
' The grammar is designed to allow future expansion without structural change:
' Single-chapter books > semantic rewrite (Jude 5 ? Jude 1:5)
' Abbreviations / aliases > BookWord resolution table
' Verse lists & ranges > already supported
' Multiple references > ; and ,
' Roman numeral normalization > Prefix
' Language variants > alternate BookName lexemes
' Pericope titles / version tags > append after Reference
'=====================================================
' Canonical Normal Form (Post-Parse Contract)
'=====================================================
' <BookName> <Chapter>:<VerseSpec>   (lists and ranges preserved)

'=====================================================
' 1. Token Stream Definition
' 1.1 Token Types
' | Token           | Description                            | Examples          |
' | --------------- | -------------------------------------- | ----------------- |
' | BOOK_WORD       | Alphabetic word, optional trailing `.` | Genesis, Gen.     |
' | PREFIX_ARABIC   | Arabic numeric prefix                  | 1, 2, 3           |
' | PREFIX_ROMAN    | Roman numeral prefix                   | I, II, III        |
' | DIGITS          | One or more digits                     | 1, 23, 150        |
' | COLON           | Chapter-verse separator                | :                 |
' | DASH            | Range separator                        | -                 |
' | COMMA           | List separator                         | ,                 |
' | SEMICOLON       | Reference separator                    | ;                 |
' | WS              | One or more spaces (collapsed)         | " "               |
' | EOF             | displayed as <END> in debug output     |                   |

'=====================================================
' 1.2 Tokenization Rules (Critical)
' Collapse whitespace ? emit a single WS
' Case-insensitive for BOOK_WORD, PREFIX_ROMAN
' BOOK_WORD may include a trailing . but never internal punctuation
' DIGITS is greedy
' : - , ; are single-character tokens
' Whitespace is significant only between book and chapter

'=====================================================
' 1.3 Example Token Streams
' Input:
' I Cor. 13:1-3,5; Rom 8:1
' Tokens:
' PREFIX_ROMAN ("I")
' WS
' BOOK_WORD ("Cor.")
' WS
' DIGITS ("13")
' COLON
' DIGITS ("1")
' DASH
' DIGITS ("3")
' COMMA
' DIGITS ("5")
' SEMICOLON
' WS
' BOOK_WORD ("Rom")
' WS
' DIGITS ("8")
' COLON
' DIGITS ("1")
' EOF

'=====================================================
' 2. Deterministic State Machine
'=====================================================
' NOTE ON STATE NUMBERING
' State S5 is intentionally unused.
' An earlier grammar revision included an intermediate
' post-chapter state that was eliminated during DFA
' minimization. State numbers were preserved to keep
' historical continuity with test data, debug traces,
' and documentation.
'
' State numbering is symbolic and not ordinal.
' This is a single-pass, left-to-right DFA.
'=====================================================
' 2.1 State Definitions
' | State | Meaning                | Accepting |
' | ----- | ---------------------- | --------- |
' | S0    | Start                  | X         |
' | S1    | Reading numeric prefix | X         |
' | S2    | Reading book name      | X         |
' | S3    | Expecting chapter      | X         |
' | S4    | Reading chapter        | ^         |
' | S6    | Reading verse          | ^         |
' | S7    | After dash (range)     | X         |
' | S8    | After comma (list)     | X         |
' | SX    | Error                  | X         |

'=====================================================
' 2.2 State Transition Table
' Legend:
' >Sx   transition to state Sx
' X     non-accepting
' ^     conditionally accepting (see Acceptance Rules)
' SX    error state (terminal)
'----------------------------------------------
' Acceptance Rules(EXPLICIT)
' A state marked ^ is accepting only if the next token is:
' <END> => ACCEPT (end of citation)
' SEMICOLON => >S0 (start next reference)
' Any other token from an accepting state => SX.
' NOTE: Transitions to >ACCEPT and >S0 are shown explicitly
' for readability; acceptance is governed by the rules above.
'----------------------------------------------
' S0 - Start:
' | Token           | Action |
' | --------------- | ------ |
' | WS              | >S0    |
' | PREFIX_ARABIC   | >S1    |
' | PREFIX_ROMAN    | >S1    |
' | BOOK_WORD       | >S2    |
' | otherwise       | >SX    |
' S1 - Prefix:
' | Token       | Action |
' | ----------- | ------ |
' | WS          | >S2    |
' | BOOK_WORD   | >S2    |
' | otherwise   | >SX    |
' S2 - Book Name:
' | Token       | Action |
' | ----------- | ------ |
' | BOOK_WORD   | >S2    |
' | WS          | >S3    |
' | otherwise   | >SX    |
' S3 - Expect Chapter:
' | Token     | Action |
' | --------- | ------ |
' | DIGITS    | >S4    |
' | otherwise | >SX    |
' S4 - Chapter(^):
' | Token       | Action   |
' | ----------- | -------- |
' | DIGITS      | >S4      |
' | COLON       | >S6      |
' | DASH        | >S7      |
' | <END>       | >ACCEPT  |
' | SEMICOLON   | >S0      |
' | otherwise   | >SX      |
' S6 - Verse(^):
' | Token       | Action   |
' | ----------- | -------- |
' | DIGITS      | >S6      |
' | DASH        | >S7      |
' | COMMA       | >S8      |
' | <END>       | >ACCEPT  |
' | SEMICOLON   | >S0      |
' | otherwise   | >SX      |
' S7 - After Dash (Range)
' | Token     | Action |
' | --------- | ------ |
' | DIGITS    | >S6    |
' | otherwise | >SX    |
' S8 - After Comma (List)
' | Token     | Action |
' | --------- | ------ |
' | DIGITS    | >S6    |
' | otherwise | >SX    |
' SX - Error
' | Token | Action |
' | ----- | ------ |
' | any   | >SX    |
' NOTE: <END> represents the EOF token in debug output

'============================================================================
' 3. Semantic Post-Processing is outside Deterministic Finite Automaton (DFA)
' Handled after a successful parse:
' Normalize prefixes > I > 1
' Collapse whitespace > single space
' Validate book name via SBL alias table
' Resolve single-chapter books
'   ----------------------------------------------
'   Single-Chapter Book Chapter Inference Rule
'   ----------------------------------------------
'   If a citation targets a single-chapter book and the input omits
'   an explicit chapter number, the chapter is inferred as 1.
'
'   Implementation Convention:
'   - chapter = 0  => chapter omitted by user
'   - chapter = 1  => chapter explicitly provided
'
'   Rewrite Rule:
'   - <Book> <Verse>            => <Book> 1:<Verse>
'   - <Book> <Chapter>          => unchanged
'   - <Book> <Chapter>:<Verse>  => unchanged
'
'   This inference is applied ONLY during semantic post-processing
'   after successful DFA parsing.
' Single-chapter rewrite rule:
'   If a reference targets a single-chapter book AND no chapter
'   was specified in the citation, rewrite <Book> <Verse>
'   as <Book> 1:<Verse>.
'   If a chapter is explicitly provided, no rewrite occurs.
' Enforce chapter/verse bounds
' Normalize output (Book Chapter:VerseSpec)

Public Type BibleBook
    bookID As Long      ' 1-66
    canonical As String ' "Genesis"
End Type

Public Enum CitationMode
    ModeGeneric = 0   ' Accept common abbreviations
    ModeSBL = 1       ' Enforce SBL Study Bible rules
End Enum

Public Sub ResetBookAliasMap()
    Set aliasMap = Nothing
End Sub

Public Function IsValidSBLAlias(bookID As Long, aliasText As String) As Boolean
    Dim canonical As String
    Dim books As Object
    Dim expected As String

    Set books = GetCanonicalBookTable
    canonical = books(bookID)(1)    ' e.g. "1 John"

    ' Normalize both sides
    expected = UCase$(canonical)
    aliasText = UCase$(Trim$(aliasText))

    IsValidSBLAlias = (aliasText = expected)
End Function

Public Function ResolveBookStrict( _
        abbr As String, _
        Optional bookID As Long, _
        Optional mode As CitationMode = ModeGeneric _
    ) As String

    Dim canonical As String

    ' Step 1: Resolve (existing logic)
    canonical = ResolveBook(abbr, bookID)

    ' Step 2: Validate (NEW)
    If mode = ModeSBL Then
        If Not IsValidSBLAlias(bookID, abbr) Then
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

        books.Add 1, Array(1, "Genesis")
        books.Add 2, Array(2, "Exodus")
        books.Add 3, Array(3, "Leviticus")
        books.Add 4, Array(4, "Numbers")
        books.Add 5, Array(5, "Deuteronomy")
        books.Add 6, Array(6, "Joshua")
        books.Add 7, Array(7, "Judges")
        books.Add 8, Array(8, "Ruth")
        books.Add 9, Array(9, "1 Samuel")
        books.Add 10, Array(10, "2 Samuel")
        books.Add 11, Array(11, "1 Kings")
        books.Add 12, Array(12, "2 Kings")
        books.Add 13, Array(13, "1 Chronicles")
        books.Add 14, Array(14, "2 Chronicles")
        books.Add 15, Array(15, "Ezra")
        books.Add 16, Array(16, "Nehemiah")
        books.Add 17, Array(17, "Esther")
        books.Add 18, Array(18, "Job")
        books.Add 19, Array(19, "Psalms")
        books.Add 20, Array(20, "Proverbs")
        books.Add 21, Array(21, "Ecclesiastes")
        books.Add 22, Array(22, "Solomon")
        books.Add 23, Array(23, "Isaiah")
        books.Add 24, Array(24, "Jeremiah")
        books.Add 25, Array(25, "Lamentations")
        books.Add 26, Array(26, "Ezekiel")
        books.Add 27, Array(27, "Daniel")
        books.Add 28, Array(28, "Hosea")
        books.Add 29, Array(29, "Joel")
        books.Add 30, Array(30, "Amos")
        books.Add 31, Array(31, "Obadiah")
        books.Add 32, Array(32, "Jonah")
        books.Add 33, Array(33, "Micah")
        books.Add 34, Array(34, "Nahum")
        books.Add 35, Array(35, "Habakkuk")
        books.Add 36, Array(36, "Zephaniah")
        books.Add 37, Array(37, "Haggai")
        books.Add 38, Array(38, "Zechariah")
        books.Add 39, Array(39, "Malachi")
        books.Add 40, Array(40, "Matthew")
        books.Add 41, Array(41, "Mark")
        books.Add 42, Array(42, "Luke")
        books.Add 43, Array(43, "John")
        books.Add 44, Array(44, "Acts")
        books.Add 45, Array(45, "Romans")
        books.Add 46, Array(46, "1 Corinthians")
        books.Add 47, Array(47, "2 Corinthians")
        books.Add 48, Array(48, "Galatians")
        books.Add 49, Array(49, "Ephesians")
        books.Add 50, Array(50, "Philippians")
        books.Add 51, Array(51, "Colossians")
        books.Add 52, Array(52, "1 Thessalonians")
        books.Add 53, Array(53, "2 Thessalonians")
        books.Add 54, Array(54, "1 Timothy")
        books.Add 55, Array(55, "2 Timothy")
        books.Add 56, Array(56, "Titus")
        books.Add 57, Array(57, "Philemon")
        books.Add 58, Array(58, "Hebrews")
        books.Add 59, Array(59, "James")
        books.Add 60, Array(60, "1 Peter")
        books.Add 61, Array(61, "2 Peter")
        books.Add 62, Array(62, "1 John")
        books.Add 63, Array(63, "2 John")
        books.Add 64, Array(64, "3 John")
        books.Add 65, Array(65, "Jude")
        books.Add 66, Array(66, "Revelation")
    End If

    Set GetCanonicalBookTable = books
End Function

Public Function GetSBLCanonicalBookTable() As Object
    Static sbl As Object

    If sbl Is Nothing Then
        Set sbl = CreateObject("Scripting.Dictionary")

        sbl.Add 1, "GENESIS"
        sbl.Add 2, "EXODUS"
        sbl.Add 4, "Numbers"
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
        sbl.Add 62, "1 JOHN"
        sbl.Add 63, "2 JOHN"
        sbl.Add 64, "3 JOHN"
        sbl.Add 64, "JUDE"
        sbl.Add 66, "REVELATION"
    End If

    Set GetSBLCanonicalBookTable = sbl
End Function

Public Sub ValidateBookSBL( _
        ByVal bookID As Long, _
        ByVal InputBook As String)

    Dim sbl As Object
    Dim expected As String

    Set sbl = GetSBLCanonicalBookTable

    If Not sbl.Exists(bookID) Then
        Err.Raise vbObjectError + 400, , _
            "Book not defined in SBL canon: " & bookID
    End If

    expected = sbl(bookID)

    If UCase(Trim(InputBook)) <> expected Then
        Err.Raise vbObjectError + 401, , _
            "Non-SBL book form. Expected '" & expected & _
            "', got '" & InputBook & "'"
    End If
End Sub

Public Function ValidateSBLReference( _
        bookID As Long, _
        canonicalName As String, _
        Chapter As Long, _
        VerseSpec As String, _
        mode As CitationMode) As Boolean

    ' Generic mode: always valid at this layer
    If mode = ModeGeneric Then
        ValidateSBLReference = True
        Exit Function
    End If

    ' ---------- SBL MODE BELOW ----------

    ' 1. Book must exist in canonical table
    If Not GetCanonicalBookTable.Exists(bookID) Then
        Debug.Print "SBL FAIL: Unknown BookID " & bookID
        Exit Function
    End If

    ' 2. Canonical name must match SBL form EXACTLY
    ' (SBL is case-insensitive in print, but normalized internally)
    Dim canon As Variant
    canon = GetCanonicalBookTable(bookID)

    If UCase(canon(1)) <> UCase(canonicalName) Then
        Debug.Print "SBL FAIL: Non-canonical book name"
        Exit Function
    End If

    ' 3. Chapter rules
    If Chapter < 0 Then
        Debug.Print "SBL FAIL: Invalid chapter number"
        Exit Function
    End If

    ' 4. Single-chapter book rules
    If GetSingleChapterBookSet.Exists(bookID) Then
        If Chapter > 1 Then
            Debug.Print "SBL FAIL: Chapter > 1 for single-chapter book"
            Exit Function
        End If
    Else
        If Chapter = 0 Then
            Debug.Print "SBL FAIL: Chapter required for multi-chapter book"
            Exit Function
        End If
    End If

    ' 5. Verse spec must exist
    If Len(VerseSpec) = 0 Then
        Debug.Print "SBL FAIL: Missing verse specification"
        Exit Function
    End If

    ' NOTE:
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

Public Function RewriteSingleChapterRef( _
        ByVal bookID As Long, _
        ByVal Chapter As Long, _
        ByVal verse As Long) As String

    Dim sc As Object
    Set sc = GetSingleChapterBookSet

    ' Only rewrite when:
    ' 1) book is single-chapter
    ' 2) chapter was omitted (chapter = 0)
    If sc.Exists(bookID) And Chapter = 0 Then
        RewriteSingleChapterRef = "1:" & verse
    ElseIf verse > 0 Then
        RewriteSingleChapterRef = Chapter & ":" & verse
    Else
        RewriteSingleChapterRef = CStr(Chapter)
    End If

End Function

Public Function ValidateAliasCoverage( _
        Optional ByRef report As String = "" _
    ) As Boolean

    Dim books As Object
    Dim aliasMap As Object
    Dim missing As Collection
    Dim k As Variant
    Dim canon As String

    Set books = GetCanonicalBookTable
    Set aliasMap = GetBookAliasMap
    Set missing = New Collection

    For Each k In books.Keys
        canon = UCase$(books(k)(1))   ' Canonical name

        If Not aliasMap.Exists(canon) Then
            missing.Add canon
        End If
    Next k

    If missing.count > 0 Then
        Dim i As Long
        report = "Missing canonical aliases:" & vbCrLf
        For i = 1 To missing.count
            report = report & "  - " & missing(i) & vbCrLf
        Next i

        ValidateAliasCoverage = False
    Else
        report = "Alias coverage complete (canonical names present)."
        ValidateAliasCoverage = True
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

Public Function ResolveBook(abbr As String, _
                            Optional bookID As Long) As String
    Dim key As String
    Dim aliasMap As Object
    Dim books As Object
    Dim b As Variant

    key = UCase$(Trim$(abbr))
    Set aliasMap = GetBookAliasMap

    If Not aliasMap.Exists(key) Then
        Debug.Print "BAD  > ResolveBook(" & key & ")"
        Err.Raise vbObjectError + 10, , "Unknown book alias: " & abbr
    End If

    bookID = aliasMap(key)
    Set books = GetCanonicalBookTable

    b = books.item(bookID)     ' SAFE
    ResolveBook = b(1)         ' Canonical name
End Function

Public Sub Test_ResolveBook()
    Debug.Print "GEN  > "; ResolveBook("GEN")
    Debug.Print "Gn   > "; ResolveBook("Gn")
    Debug.Print "1 JN > "; ResolveBook("1 JN")
    Debug.Print "1 JOH> "; ResolveBook("1 JOH")
    Debug.Print "REV  > "; ResolveBook("REV")

    Debug.Print "BAD  > "; ResolveBook("XYZ")
End Sub

Public Sub Test_ResolveBook_Strict()
    Debug.Assert ResolveBook("GEN") = "Genesis"
    Debug.Assert ResolveBook("GN") = "Genesis"
    Debug.Assert ResolveBook("1 JN") = "1 John"
    Debug.Assert ResolveBook("1 JOH") = "1 John"
    Debug.Assert ResolveBook("REV") = "Revelation"
    Debug.Assert ResolveBook("XYZ") = ""
End Sub

Public Sub Test_AllBookAliases_STRICT()
    Dim aliasMap As Object
    Dim books As Object
    Dim k As Variant
    Dim bookID As Long
    Dim canonicalActual As String
    Dim failures As Long

    Set aliasMap = GetBookAliasMap
    Set books = GetCanonicalBookTable

    Debug.Print "=== STRICT Alias Validation ==="

    For Each k In aliasMap.Keys
        On Error GoTo AliasFail

        canonicalActual = ResolveBook(CStr(k), bookID)

        ' 1. BookID must be in canonical range
        If bookID < 1 Or bookID > 66 Then
            Debug.Print "FAIL (INVALID BOOK ID): "; k; " ? "; bookID
            failures = failures + 1
            GoTo NextAlias
        End If

        ' 2. BookID must exist in canonical table
        If Not books.Exists(bookID) Then
            Debug.Print "FAIL (MISSING CANONICAL): "; k; " ? "; bookID
            failures = failures + 1
            GoTo NextAlias
        End If

        ' 3. Canonical name must match
        If canonicalActual <> books(bookID)(1) Then
            Debug.Print "FAIL (NAME MISMATCH): "; k; _
                        " ? "; canonicalActual; _
                        " (expected "; books(bookID)(1); ")"
            failures = failures + 1
        End If

NextAlias:
        On Error GoTo 0
    Next k

    If failures = 0 Then
        Debug.Print "PASS: All aliases are canonical and valid."
    Else
        Debug.Print "FAILURES: "; failures
    End If
    Exit Sub

AliasFail:
    Debug.Print "FAIL (ERROR): "; k; " > "; Err.Description
    Err.Clear
    failures = failures + 1
    Resume NextAlias
End Sub


