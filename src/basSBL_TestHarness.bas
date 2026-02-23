Attribute VB_Name = "basSBL_TestHarness"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Public Enum ExpectedFailureStage
    FailNone = 0
    FailResolveBook = 1
    FailSemantic = 2
End Enum

Public Function ParseReferenceStub(ByVal inputText As String) As ParsedReference
    Dim p As ParsedReference
    p.RawInput = inputText

    Debug.Print "  [Stub] Raw input = >" & inputText & "<"

    Dim parts() As String
    parts = Split(Trim$(inputText), " ")

    '----------------------------------
    ' Book alias (first token)
    '----------------------------------
    p.BookAlias = UCase$(parts(0))
    Debug.Print "  [Stub] Parsed alias = >" & p.BookAlias & "<"
    ' Semantic guardd to check for null failures
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
        p.Chapter = CLng(Split(refPart, ":")(0))
        p.VerseSpec = Split(refPart, ":")(1)
    Else
        ' Single-chapter book case (e.g. Jude 5)
        p.Chapter = 0
        p.VerseSpec = refPart
    End If

    ParseReferenceStub = p
End Function

Public Sub Test_AliasCoverage()
' Assert that every canonical book name (upper-cased) exists as a key in the alias map
'   Uses GetCanonicalBookTable
'   Uses GetBookAliasMap
'   Canonical name is normalized as UCase(Canonical)
'   Does not mutate state
'   Emits diagnostics
    Debug.Print ""
    Debug.Print "=== Alias Coverage Validation ==="

    Dim msg As String
    Dim ok As Boolean

    ok = ValidateAliasCoverage(msg)

    Debug.Print msg

    If Not ok Then
        Debug.Print "RESULT: FAIL"
    Else
        Debug.Print "RESULT: PASS"
    End If

    Debug.Print "==============================="
End Sub

Public Sub Test_SemanticFlow_WithParserStub()
    ResetBookAliasMap

    Debug.Print "======================================"
    Debug.Print " Test_SemanticFlow_WithParserStub"
    Debug.Print " (Uses lightweight parser stub)"
    Debug.Print "======================================"

    Dim failures As Long
    failures = 0

    ' Test cases:
    ' RawInput, ExpectedBookID, ExpectValid, ExpectRewrite
    ' NOTE:
    '  - Alias / chapter / verse are now derived via ParseReferenceStub
    '  - This simulates a real parser without implementing DFA/tokenizer
    Dim tests
    tests = Array( _
        Array("Jude 5", 65, True, True), _
        Array("Jude 1:5", 65, True, True), _
        Array("Obadiah 3", 31, True, True), _
        Array("Romans 8:1", 45, True, False), _
        Array("Genesis 1:1", 1, True, False) _
    )

    Dim i As Long
    For i = LBound(tests) To UBound(tests)

        Debug.Print ""
        Debug.Print "INPUT: "; tests(i)(0)

        Dim testFailed As Boolean
        testFailed = False

        '---------------------------------------
        ' Parser Stub Phase
        '---------------------------------------
        ' This replaces manual extraction.
        ' Later, this call will be replaced by a real parser.
        Dim parsed As ParsedReference
        parsed = ParseReferenceStub(tests(i)(0))

        Debug.Print "  ParseReferenceStub:"
        Debug.Print "    -> Alias:   "; parsed.BookAlias
        Debug.Print "    -> Chapter: "; parsed.Chapter
        Debug.Print "    -> Verse:   "; parsed.VerseSpec

        '---------------------------------------
        ' Resolver Phase
        '---------------------------------------
        Dim bookName As String
        Dim bookID As Long

        On Error Resume Next
        bookName = ResolveBook(parsed.BookAlias, bookID)
        If Err.Number <> 0 Then
            Debug.Print "  ERROR: ResolveBook failed"
            failures = failures + 1
            Err.Clear
            GoTo NextTest
        End If
        On Error GoTo 0

        Debug.Print "  Resolver:"
        Debug.Print "    -> BookID:    "; bookID
        Debug.Print "    -> Canonical: "; bookName

        If bookID <> tests(i)(1) Then
            Debug.Print "  FAIL: BookID mismatch"
            testFailed = True
        End If

        '---------------------------------------
        ' Semantic Validation Phase (SBL)
        '---------------------------------------
        Dim semanticMsg As String
        Dim isValid As Boolean

        isValid = ValidateSBLReference( _
                    bookID, _
                    bookName, _
                    parsed.Chapter, _
                    parsed.VerseSpec, _
                    ModeSBL)

        Debug.Print "  ValidateSBLReference:"
        Debug.Print "    -> Valid: "; isValid

        If isValid <> tests(i)(2) Then
            Debug.Print "  FAIL: semantic validity mismatch"
            testFailed = True
        End If

        '---------------------------------------
        ' Rewrite Phase (single-chapter books)
        '---------------------------------------
        If isValid Then
            Dim rewritten As String
            rewritten = RewriteSingleChapterRef( _
                            bookID, _
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
End Sub


Public Sub Report_TODOs()
    Debug.Print "=== NOT IMPLEMENTED / TODO ============================"
    Debug.Print "- Replace ParseReferenceStub with real tokenizer + DFA"
    Debug.Print "- Multi-token book names (1 John, Song of Songs)"
    Debug.Print "- Roman numeral prefixes"
    Debug.Print "- Verse list/range parsing"
    Debug.Print "- Structured parse errors"
    Debug.Print "- Optional future validator hardening"
    Debug.Print "    This validator can be extended to assert:"
    Debug.Print "      - each book has >= 1 non-canonical alias"
    Debug.Print "      - no alias maps to multiple books"
    Debug.Print "      - SBL-strict aliases form a closed subset"
    Debug.Print "      - alias casing normalization consistency"
    Debug.Print "======================================================="
End Sub

Public Sub Test_SemanticFlow_WithParserStub_Negative()
    ResetBookAliasMap

    Debug.Print "=========================================="
    Debug.Print " Test_SemanticFlow_WithParserStub_Negative"
    Debug.Print "=========================================="

    Dim failReason As String
    failReason = ""
    
    Dim failures As Long
    failures = 0

    ' RawInput, ExpectValid
    ' Jude 0        => Verse 0 invalid
    ' Jude 999      => Verse out of range
    ' Jude 1:0      => Explicit verse 0
    ' Romans 0:1    => Chapter 0 invalid for multi-chapter book
    ' Romans 999:1  => Chapter out of range
    ' Genesis 1:999 => Verse out of range
    
    Dim tests As Variant
    tests = Array( _
        Array("Jude 0", FailSemantic), _
        Array("Jude 999", FailSemantic), _
        Array("Jude 1:0", FailSemantic), _
        Array("Romans 0:1", FailSemantic), _
        Array("Romans 999:1", FailSemantic), _
        Array("Genesis 1:999", FailSemantic) _
    )
    
    Dim i As Long
    For i = LBound(tests) To UBound(tests)

        Debug.Print ""
        Debug.Print "INPUT: "; tests(i)(0)

'x        Dim BookAlias As String
'x        Dim chapter As Long
'x        Dim verseSpec As String

        ' -----------------------------
        ' Parser stub phase
        ' -----------------------------
        Dim parsed As ParsedReference
        parsed = ParseReferenceStub(tests(i)(0))

        Debug.Print "  ParseReferenceStub:"
        Debug.Print "    -> Alias:   "; parsed.BookAlias
        Debug.Print "    -> Chapter: "; parsed.Chapter
        Debug.Print "    -> Verse:   "; parsed.VerseSpec

        ' -----------------------------
        ' Resolver phase
        ' -----------------------------
        Dim bookName As String
        Dim bookID As Long
        
        On Error Resume Next
        bookName = ResolveBook(parsed.BookAlias, bookID)
        
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
        Debug.Print "    -> BookID:     "; bookID
        Debug.Print "    -> Canonical:  "; bookName

        ' -----------------------------
        ' Semantic validation phase
        ' -----------------------------
        Dim isValid As Boolean

        isValid = ValidateSBLReference( _
                    bookID, _
                    bookName, _
                    parsed.Chapter, _
                    parsed.VerseSpec, _
                    ModeSBL)

        Debug.Print "  ValidateSBLReference:"
        Debug.Print "    -> Valid: "; isValid

        If isValid <> tests(i)(1) Then
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
End Sub

Public Sub Run_All_SBL_Tests()
    ResetBookAliasMap
    Test_AliasCoverage
    Test_SemanticFlow_WithParserStub
End Sub

