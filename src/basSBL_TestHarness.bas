Attribute VB_Name = "basSBL_TestHarness"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Private Const RUN_FAILURE_DEMOS As Boolean = False  ' Set True to run intentional-failure test cases that demonstrate error detection

Public Enum ExpectedFailureStage
    FailNone = 0
    FailResolveBook = 1
    FailSemantic = 2
End Enum

Public Function ParseReferenceStub(ByVal inputText As String) As ParsedReference
    On Error GoTo PROC_ERR
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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ParseReferenceStub of Module basSBL_TestHarness"
    Resume PROC_EXIT
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
    On Error GoTo PROC_ERR
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
    '  - This simulates a real parser without implementing DSP/tokenizer
    Dim tests
    ' "Jude 5-7": Range spec — ValidateSBLReference currently rejects non-numeric VerseSpec (ExpectValid=False).
    ' When Stage 8-12 range support is added, ExpectValid becomes True.
    ' Without the IsNumeric guard in the rewrite phase, that transition crashes with error 13.
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
            ' FIXME_LATER: When Stage 8-12 range extensions update ValidateSBLReference to
            ' accept range VerseSpec values (e.g. "5-7"), IsValid will be True for range inputs.
            ' CLng(parsed.VerseSpec) raises error 13 (Type Mismatch) on a range string.
            ' The IsNumeric guard below prevents the crash; the Else branch handles the range path.
            ' See test case "Jude 5-7" below — currently ExpectValid=False (range not yet supported),
            ' but that expectation will change when Stage 8-12 is complete.
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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_SemanticFlow_WithParserStub of Module basSBL_TestHarness"
    Resume PROC_EXIT
End Sub

Public Sub Report_TODOs()
    Debug.Print "=== NOT IMPLEMENTED / TODO ============================"
    Debug.Print "- Replace ParseReferenceStub with real tokenizer + DSP"
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
    On Error GoTo PROC_ERR
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
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_SemanticFlow_WithParserStub_Negative of Module basSBL_TestHarness"
    Resume PROC_EXIT
End Sub

Public Sub Test_GetMaxVerse()
    On Error GoTo PROC_ERR
    Dim failCount As Long
    Dim Result As Long

    Debug.Print ""
    Debug.Print "---- Test_GetMaxVerse ----"
    ' ========================
    ' POSITIVE TESTS
    ' ========================
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
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_GetMaxVerse of Module basSBL_TestHarness"
    Resume PROC_EXIT
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
    On Error GoTo PROC_ERR
    TestStart

    If Not VerifyPackedVerseMap() Then
        Debug.Print "ABORT: Packed verse map invalid"
        GoTo PROC_EXIT
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
    '   - Update DSP documentation
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
    Test_Stage13_ContextShorthand
    TestSummary
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Run_All_SBL_Tests of Module basSBL_TestHarness"
    Resume PROC_EXIT
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

Public Sub Test_Stage4_InterpretStructure()
    Debug.Print ""
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
    AssertEqual 0, ref.Chapter, "Jude 5 chapter interpreted"
    AssertEqual "5", ref.VerseSpec, "Jude 5 verse interpreted"
    '------------------------------------------
    ' Standard chapter:verse
    '------------------------------------------
    tokens = LexicalScan("Romans 8:1")
    ref = InterpretStructure(tokens)
    AssertEqual 8, ref.Chapter, "Romans chapter interpreted"
    AssertEqual "1", ref.VerseSpec, "Romans verse interpreted"
    '------------------------------------------
    ' Ambiguous single number (no colon)
    '------------------------------------------
    tokens = LexicalScan("Genesis 1")
    ref = InterpretStructure(tokens)
    AssertEqual 0, ref.Chapter, "Genesis 1 chapter interpreted"
    AssertEqual "1", ref.VerseSpec, "Genesis 1 verse interpreted"
    '------------------------------------------
    ' Verse range
    '------------------------------------------
    tokens = LexicalScan("John 3:16-18")
    ref = InterpretStructure(tokens)
    AssertEqual 3, ref.Chapter, "John chapter interpreted"
    AssertEqual "16-18", ref.VerseSpec, "John verse range interpreted"
    '------------------------------------------
    ' Verse list
    '------------------------------------------
    tokens = LexicalScan("Psalm 23:1,4,6")
    ref = InterpretStructure(tokens)
    AssertEqual 23, ref.Chapter, "Psalm chapter interpreted"
    AssertEqual "1,4,6", ref.VerseSpec, "Psalm verse list interpreted"
    '------------------------------------------
    ' Mixed list and range
    '------------------------------------------
    tokens = LexicalScan("Matthew 5:3-5,9")
    ref = InterpretStructure(tokens)
    AssertEqual 5, ref.Chapter, "Matthew chapter interpreted"
    AssertEqual "3-5,9", ref.VerseSpec, "Matthew mixed verse spec interpreted"
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
    '------------------------------------------
    ' Book-only expansion
    '------------------------------------------
    Result = ParseReference("John")
    AssertEqual "John 1:1", Result, "John book expansion"
    Result = ParseReference("1 Jn")
    AssertEqual "1 John 1:1", Result, "1 Jn expansion"
    Result = ParseReference("Jude")
    AssertEqual "Jude 1:1", Result, "Jude book expansion"
    Result = ParseReference("Romans")
    AssertEqual "Romans 1:1", Result, "Romans book expansion"
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
    AssertTrue t.IsList, "comma list detected"
    AssertEqual 2, UBound(t.Segments), "comma list segment count"
    '------------------------------------------
    ' Test 2 - semicolon list
    '------------------------------------------
    t = ListDetection("John 3:16; 4:1")
    AssertTrue t.IsList, "semicolon list detected"
    AssertEqual 1, UBound(t.Segments), "semicolon list segment count"
    '------------------------------------------
    ' Test 3 - single reference
    '------------------------------------------
    t = ListDetection("John 3:16")
    AssertFalse t.IsList, "single reference not list"
    '------------------------------------------
    ' Test 4 - list containing range
    '------------------------------------------
    t = ListDetection("John 3:16-18,20")
    AssertTrue t.IsList, "range preserved inside list"
    AssertEqual 1, UBound(t.Segments), "range list segment count"
    AssertEqual "John 3:16-18", t.Segments(0), "range first segment"
    AssertEqual "20", t.Segments(1), "range second segment"
    '------------------------------------------
    ' Test 5 - mixed whitespace
    '------------------------------------------
    ' Optional as whitespace normalization should already be handled in Stage 1
    t = ListDetection("John 3:16 , 18 , 20")
    AssertTrue t.IsList, "whitespace tolerated"
    AssertEqual 2, UBound(t.Segments), "whitespace segment count"
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
    r = RangeDetection("John 3:16" & ChrW(8211) & "18")   ' en dash character (U+2013)
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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure PrintScriptureList of Module basSBL_TestHarness"
    Resume PROC_EXIT
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

Public Sub Test_Stage13_ContextShorthand()
    Dim c As Collection
    Dim v

    Debug.Print "====================================="
    Debug.Print "Stage 13 Contextual Shorthand Tests"
    Debug.Print "====================================="
    '------------------------------------------
    ' Test 1
    '------------------------------------------
    Set c = ComposeList("John 3:16, 18, 20-22")
    Debug.Print "Test 1"
    For Each v In c
        Debug.Print v
    Next
    'Expected
    'John 3:16
    'John 3:18
    'John 3:20-22
    '------------------------------------------
    ' Test 2
    '------------------------------------------
    Set c = ComposeList("John 3:16-4:2, 5")
    Debug.Print
    Debug.Print "Test 2"
    For Each v In c
        Debug.Print v
    Next
    'Expected
    'John 3:16-4:2
    'John 4:5
    '------------------------------------------
    ' Test 3
    '------------------------------------------
    Set c = ComposeList("Romans 8; 9")
    Debug.Print
    Debug.Print "Test 3"
    For Each v In c
        Debug.Print v
    Next
    'Expected
    'Romans 8
    'Romans 9
    Debug.Print
    Debug.Print "Stage 13 tests complete."
End Sub
