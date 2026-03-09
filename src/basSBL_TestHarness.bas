Attribute VB_Name = "basSBL_TestHarness"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Private Const RUN_FAILURE_DEMOS As Boolean = False

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
    Debug.Print "------------------------------------------"
    Debug.Print "   Alias Coverage Validation"
    Debug.Print "------------------------------------------"

    Dim msg As String
    Dim ok As Boolean
    ok = AliasCoverage(msg)
    Debug.Print msg

    If Not ok Then
        Debug.Print "RESULT: FAIL"
    Else
        Debug.Print "RESULT: PASS"
    End If
End Sub

Public Sub Test_TokenizeReference()
    Dim t As LexTokens
    
    t = LexicalScan("Jude 1:5")
    Debug.Assert t.RawAlias = "Jude"
    Debug.Assert t.Num1 = 1
    Debug.Assert t.Num2 = 5
    Debug.Assert t.HasColon = True
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
        Dim BookID As Long

        On Error Resume Next
        bookName = ResolveAlias(parsed.BookAlias, BookID)
        If Err.Number <> 0 Then
            Debug.Print "  ERROR: ResolveBook failed"
            failures = failures + 1
            Err.Clear
            GoTo NextTest
        End If
        On Error GoTo 0

        Debug.Print "  Resolver:"
        Debug.Print "    -> BookID:    "; BookID
        Debug.Print "    -> Canonical: "; bookName

        If BookID <> tests(i)(1) Then
            Debug.Print "  FAIL: BookID mismatch"
            testFailed = True
        End If

        '---------------------------------------
        ' Semantic Validation Phase (SBL)
        '---------------------------------------
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

        '---------------------------------------
        ' Rewrite Phase (single-chapter books)
        '---------------------------------------
        If IsValid Then
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
        Array("Jude 0", False), _
        Array("Jude 999", False), _
        Array("Jude 1:0", False), _
        Array("Romans 0:1", False), _
        Array("Romans 999:1", False), _
        Array("Genesis 1:999", False) _
    )
    
    Dim i As Long
    For i = LBound(tests) To UBound(tests)

        Debug.Print ""
        Debug.Print "INPUT: "; tests(i)(0)

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

        ' -----------------------------
        ' Semantic validation phase
        ' -----------------------------
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
End Sub

Public Sub Test_GetMaxVerse()
    Dim failCount As Long
    Dim Result As Long
    
    Debug.Print ""
    Debug.Print "---- Test_GetMaxVerse ----"
    ' ========================
    ' POSITIVE TESTS
    ' ========================
    On Error GoTo FailHandler
    Result = GetMaxVerse(1, 1)          ' Genesis 1
    If Result <> 31 Then FailTest failCount, "Genesis 1", 31, Result
    Result = GetMaxVerse(19, 119)       ' Psalms 119
    If Result <> 176 Then FailTest failCount, "Psalms 119", 176, Result
    Result = GetMaxVerse(65, 1)         ' Jude 1
    If Result <> 25 Then FailTest failCount, "Jude 1", 25, Result
    Result = GetMaxVerse(66, 22)        ' Revelation 22
    If Result <> 21 Then FailTest failCount, "Revelation 22", 21, Result
    ' ========================
    ' NEGATIVE TESTS
    ' ========================
    On Error Resume Next
    Err.Clear
    Result = GetMaxVerse(1, 999)
    If Err.Number = 0 Then
        Debug.Print "FAIL: Invalid chapter not rejected"
        failCount = failCount + 1
    End If
    Err.Clear
    Result = GetMaxVerse(999, 1)
    If Err.Number = 0 Then
        Debug.Print "FAIL: Invalid book not rejected"
        failCount = failCount + 1
    End If
    Err.Clear
    Result = GetMaxVerse(19, 0)
    If Err.Number = 0 Then
        Debug.Print "FAIL: Chapter zero not rejected"
        failCount = failCount + 1
    End If
    Err.Clear
    On Error GoTo 0
    ' ========================
    ' SUMMARY
    ' ========================
    If failCount = 0 Then
        Debug.Print "Test_GetMaxVerse: PASS"
    Else
        Debug.Print "Test_GetMaxVerse: FAIL (" & failCount & " errors)"
    End If
    Exit Sub
    
FailHandler:
    Debug.Print "Unexpected runtime error in Test_GetMaxVerse"
    Debug.Print "Error: "; Err.Number; Err.Description
    failCount = failCount + 1
    Resume Next
End Sub

Private Sub FailTest(ByRef failCount As Long, _
                     ByVal label As String, _
                     ByVal expected As Long, _
                     ByVal actual As Long)
    Debug.Print "FAIL: "; label; _
                " Expected="; expected; _
                " Actual="; actual
    
    failCount = failCount + 1
End Sub

Public Sub Run_All_SBL_Tests()
    TestStart

    If Not VerifyPackedVerseMap() Then
        Debug.Print "ABORT: Packed verse map invalid"
        Exit Sub
    End If

    ResetBookAliasMap
    Test_AliasCoverage
    Test_Stage2_LexicalScan
    Test_Stage3_ResolveAlias
    Test_Stage4_InterpretStructure
    Test_Stage5_ValidateCanonical
    Test_Stage6_FormatCanonical
    Test_Stage6_FormatCanonical_FailureDemo
    Test_Stage7_EndToEnd
    Test_GetMaxVerse
    '-----------------------------
    ' Embedded Extension Hooks
    '-----------------------------
    'The order is:
    '   - Update DFA documentation
    '   - Define Stage-8 contract
    '   - Write Test_Stage8_ListDetection
    '   - Implement Stage-8 code
    '   - Enable test in runner
    '   Stages 9 and 10 will follow the same sequence.
    Test_Stage8_ListDetection
    Test_Stage9_RangeDetection
    Test_Stage10_RangeComposition
    Test_Stage11_ListComposition
    Test_Stage12_FinalParser
    TestSummary
End Sub

Public Sub Test_Stage2_LexicalScan()
    Debug.Print ""
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage2_LexicalScan"
    Debug.Print "------------------------------------------"

    Dim t As LexTokens
    t = LexicalScan("Jude 1:5")
    AssertEqual "Jude", t.RawAlias, "Alias parsed"
    AssertEqual 1, t.Num1, "Chapter parsed"
    AssertEqual 5, t.Num2, "Verse parsed"
    AssertTrue t.HasColon, "Colon detected"

    t = LexicalScan("Romans 8")
    AssertEqual "Romans", t.RawAlias, "Alias parsed"
    AssertEqual 8, t.Num1, "Number parsed"
    AssertTrue Not t.HasColon, "No colon detected"
End Sub

Public Sub Test_Stage3_ResolveAlias()
    Debug.Print ""
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage3_ResolveAlias"
    Debug.Print "------------------------------------------"

    Dim tokens As LexTokens
    Dim BookID As Long
    Dim canonical As String

    tokens = LexicalScan("Jude 1:5")
    canonical = ResolveAlias(tokens.RawAlias, BookID)
    AssertEqual 65, BookID, "Jude BookID"
    AssertEqual "Jude", canonical, "Jude canonical"

    tokens = LexicalScan("Genesis 1:1")
    canonical = ResolveAlias(tokens.RawAlias, BookID)
    AssertEqual 1, BookID, "Genesis BookID"
    AssertEqual "Genesis", canonical, "Genesis canonical"
End Sub

Public Function InterpretStructure(ByRef t As LexTokens) As ParsedReference
    Dim p As ParsedReference

    '----------------------------------------
    ' Propagate alias
    '----------------------------------------
    p.BookAlias = UCase$(t.RawAlias)
    Debug.Assert t.RawAlias <> vbNullString
    '----------------------------------------
    ' Structural interpretation
    '----------------------------------------
    If t.HasColon Then
        ' Book Chapter:Verse
        p.Chapter = t.Num1
        p.VerseSpec = CStr(t.Num2)
    Else
        If t.Num1 > 0 Then
            ' Ambiguous case (chapter or verse)
            p.Chapter = 0
            p.VerseSpec = CStr(t.Num1)
        Else
            ' Book-only reference
            p.Chapter = 0
            p.VerseSpec = ""
        End If
    End If
    InterpretStructure = p
End Function

Public Sub Test_Stage4_InterpretStructure()
    Debug.Print ""
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage4_InterpretStructure"
    Debug.Print "------------------------------------------"

    Dim tokens As LexTokens
    Dim ref As ParsedReference

    tokens = LexicalScan("Jude 5")
    ref = InterpretStructure(tokens)
    AssertEqual 0, ref.Chapter, "Jude 5 chapter interpreted"
    AssertEqual "5", ref.VerseSpec, "Jude 5 verse interpreted"

    tokens = LexicalScan("Romans 8:1")
    ref = InterpretStructure(tokens)
    AssertEqual 8, ref.Chapter, "Romans chapter interpreted"
    AssertEqual "1", ref.VerseSpec, "Romans verse interpreted"
End Sub

Public Sub Test_Stage5_ValidateCanonical()
    Debug.Print ""
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage5_ValidateCanonical"
    Debug.Print "------------------------------------------"

    Dim valid As Boolean
    valid = ValidateSBLReference(65, "Jude", 0, "5", ModeSBL)
    AssertTrue valid, "Jude 5 valid"
    
    valid = ValidateSBLReference(65, "Jude", 1, "0", ModeSBL, True)
    AssertTrue Not valid, "Jude 1:0 rejected"
    
    valid = ValidateSBLReference(45, "Romans", 999, "1", ModeSBL, True)
    AssertTrue Not valid, "Romans 999:1 rejected"
End Sub

Public Sub Test_Stage6_FormatCanonical()
    Debug.Print ""
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage6_FormatCanonical"
    Debug.Print "------------------------------------------"

    Dim Result As String
    Result = RewriteSingleChapterRef(65, 0, 5)
    AssertEqual "1:5", Result, "Jude single-chapter rewrite"

    Result = RewriteSingleChapterRef(45, 8, 1)
    AssertEqual "8:1", Result, "Romans unchanged"
End Sub

Public Sub Test_Stage6_FormatCanonical_FailureDemo()
    If RUN_FAILURE_DEMOS Then
        Debug.Print ""
        Debug.Print "------------------------------------------"
        Debug.Print " Test_Stage6_FormatCanonical (Failure Demo)"
        Debug.Print "------------------------------------------"
    
        Dim Result As String
        ' Call the real formatter
        Result = RewriteSingleChapterRef(65, 0, 5)   ' Actual: "1:5"
    
        ' Deliberate wrong expected value
        AssertEqual "5", Result, "Canonical rewrite for Jude"
    End If
End Sub

Public Sub Test_Stage7_EndToEnd()
    Dim Result As String

    Debug.Print ""
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage7_EndToEnd"
    Debug.Print "------------------------------------------"

    Result = ParseReference("Jude 5")
    AssertEqual "Jude 1:5", Result, "Jude single-chapter expansion"

    Result = ParseReference("Romans 8")
    AssertEqual "Romans 8", Result, "Romans chapter reference"

    Result = ParseReference("3 John 4")
    AssertEqual "3 John 1:4", Result, "3 John expansion"

    Result = ParseReference("Genesis 1:1")
    AssertEqual "Genesis 1:1", Result, "Genesis unchanged"
End Sub

Public Sub Test_Stage8_ListDetection()
    Debug.Print
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage8_ListDetection"
    Debug.Print "------------------------------------------"

    Dim t As ListTokens
    '------------------------------------------
    ' Test 1 - comma list
    '------------------------------------------
    t = ListDetection("John 3:16,18,20")
    If t.IsList And UBound(t.Segments) = 2 Then
        Debug.Print "PASS: comma list detected"
    Else
        Debug.Print "FAIL: comma list detection"
    End If
    '------------------------------------------
    ' Test 2 - semicolon list
    '------------------------------------------
    t = ListDetection("John 3:16; 4:1")
    If t.IsList And UBound(t.Segments) = 1 Then
        Debug.Print "PASS: semicolon list detected"
    Else
        Debug.Print "FAIL: semicolon list detection"
    End If
    '------------------------------------------
    ' Test 3 - single reference
    '------------------------------------------
    t = ListDetection("John 3:16")
    If Not t.IsList Then
        Debug.Print "PASS: single reference not list"
    Else
        Debug.Print "FAIL: false list detection"
    End If
    '------------------------------------------
    ' Test 4 - list containing range
    '------------------------------------------
    t = ListDetection("John 3:16-18,20")
    If t.IsList _
       And UBound(t.Segments) = 1 _
       And t.Segments(0) = "John 3:16-18" _
       And t.Segments(1) = "20" Then
        Debug.Print "PASS: range preserved inside list"
    Else
        Debug.Print "FAIL: range incorrectly split"
    End If
    '------------------------------------------
    ' Test 5 - mixed whitespace
    '------------------------------------------
    ' Optional as whitespace normalization should already be handled in Stage 1
    t = ListDetection("John 3:16 , 18 , 20")
    If t.IsList And UBound(t.Segments) = 2 Then
        Debug.Print "PASS: whitespace tolerated"
    Else
        Debug.Print "FAIL: whitespace handling"
    End If
End Sub

Public Sub Test_Stage9_RangeDetection()
    Debug.Print
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage9_RangeDetection"
    Debug.Print "------------------------------------------"

    Dim r As RangeTokens
    '------------------------------------------
    ' Test 1 - verse range
    '------------------------------------------
    AssertTrue IsRangeSegment("John 3:16-18"), _
        "IsRangeSegment verse range"

    r = RangeDetection("John 3:16-18")
    AssertTrue r.IsRange, "verse range detected"
    AssertEqual r.LeftRaw, "John 3:16", "range left token"
    AssertEqual r.RightRaw, "18", "range right token"
    '------------------------------------------
    ' Test 2 - chapter range
    '------------------------------------------
    r = RangeDetection("John 3-5")
    AssertTrue r.IsRange, "chapter range detected"
    AssertEqual r.LeftRaw, "John 3", "chapter range left"
    AssertEqual r.RightRaw, "5", "chapter range right"
    '------------------------------------------
    ' Test 3 - cross chapter range
    '------------------------------------------
    r = RangeDetection("John 3:16-4:2")
    AssertTrue r.IsRange, "cross chapter range detected"
    AssertEqual r.LeftRaw, "John 3:16", "cross chapter left"
    AssertEqual r.RightRaw, "4:2", "cross chapter right"
    '------------------------------------------
    ' Test 4 - en dash
    '------------------------------------------
    r = RangeDetection("John 3:16–18")
    AssertTrue r.IsRange, "en dash range detected"
    AssertEqual r.LeftRaw, "John 3:16", "en dash left"
    AssertEqual r.RightRaw, "18", "en dash right"
    '------------------------------------------
    ' Test 5 - not a range
    '------------------------------------------
    AssertFalse IsRangeSegment("John 3:16"), _
        "IsRangeSegment single reference"

    r = RangeDetection("John 3:16")
    AssertFalse r.IsRange, "single reference not range"
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
    AssertEqual 3, r.StartRef.Chapter, "start chapter"
    AssertEqual 16, r.StartRef.Verse, "start verse"
    AssertEqual 3, r.EndRef.Chapter, "end chapter"
    AssertEqual 18, r.EndRef.Verse, "end verse"
    '------------------------------------------
    ' Chapter shorthand
    '------------------------------------------
    r = ComposeRange("John 3-5")
    AssertEqual 3, r.StartRef.Chapter, "chapter start"
    AssertEqual 5, r.EndRef.Chapter, "chapter end"
    '------------------------------------------
    ' Cross-chapter range
    '------------------------------------------
    r = ComposeRange("Genesis 1:31-2:3")
    AssertEqual 1, r.StartRef.Chapter, "cross start chapter"
    AssertEqual 31, r.StartRef.Verse, "cross start verse"
    AssertEqual 2, r.EndRef.Chapter, "cross end chapter"
    AssertEqual 3, r.EndRef.Verse, "cross end verse"
End Sub

Public Sub PrintScriptureList(list As ScriptureList)
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
End Sub

Public Sub Test_Stage11_ListComposition()
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage11_ListComposition"
    Debug.Print "------------------------------------------"

    Dim Items As Collection
    '------------------------------------------
    ' Simple reference list
    '------------------------------------------
    Set Items = ComposeList("John 3:16, John 3:18")
    AssertEqual 2, Items.count, "two references parsed"
    AssertEqual "John 3:16", Items(1), "first reference"
    AssertEqual "John 3:18", Items(2), "second reference"
    '------------------------------------------
    ' Range inside list
    '------------------------------------------
    Set Items = ComposeList("John 3:16-18, John 3:20")
    AssertEqual 2, Items.count, "range + reference"
    AssertEqual "John 3:16-3:18", Items(1), "range canonical"
    AssertEqual "John 3:20", Items(2), "second reference"
    '------------------------------------------
    ' Semicolon separation
    '------------------------------------------
    Set Items = ComposeList("Romans 8; Romans 9")
    AssertEqual 2, Items.count, "semicolon list"
End Sub

Public Sub Test_Stage12_FinalParser()
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage12_FinalParser"
    Debug.Print "------------------------------------------"

    Dim Result As Variant
    Dim Items As Collection
    '------------------------------------------
    ' Single reference
    '------------------------------------------
    Result = ParseScripture("John 3:16")
    AssertEqual "John 3:16", Result, "single reference"
    '------------------------------------------
    ' Range
    '------------------------------------------
    Result = ParseScripture("John 3:16-18")
    AssertEqual "John 3:16-3:18", Result, "range parsed"
    '------------------------------------------
    ' List
    '------------------------------------------
    Set Items = ParseScripture("John 3:16, John 3:18")
    AssertEqual 2, Items.count, "list parsed"
    '------------------------------------------
    ' Mixed list + range
    '------------------------------------------
    Set Items = ParseScripture("John 3:16-18, John 3:20")
    AssertEqual 2, Items.count, "mixed parsed"
End Sub
