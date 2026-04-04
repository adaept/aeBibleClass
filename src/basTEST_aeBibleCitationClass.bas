Attribute VB_Name = "basTEST_aeBibleCitationClass"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Public aeAssert As aeAssertClass
Private Const RUN_FAILURE_DEMOS As Boolean = False  ' Set True to run intentional-failure test cases that demonstrate error detection

Public Enum ExpectedFailureStage
    FailNone = 0
    FailResolveBook = 1
    FailSemantic = 2
End Enum

Public Sub Test_Stage1_AliasCoverage()
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

    ok = aeBibleCitationClass.AliasCoverage(msg)
    Debug.Print msg

    aeAssert.AssertTrue ok, "Alias coverage validation", True, ok
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

Public Sub Test_GetMaxVerse()
    On Error GoTo PROC_ERR
    Dim failCount As Long
    Dim Result As Long

    Debug.Print ""
    Debug.Print "---- Test_GetMaxVerse ----"
    ' ========================
    ' POSITIVE TESTS
    ' ========================
    Result = aeBibleCitationClass.GetMaxVerse(1, 1)          ' Genesis 1
    If Result <> 31 Then FailTest failCount, "Genesis 1", 31, Result
    Result = aeBibleCitationClass.GetMaxVerse(19, 119)       ' Psalms 119
    If Result <> 176 Then FailTest failCount, "Psalms 119", 176, Result
    Result = aeBibleCitationClass.GetMaxVerse(65, 1)         ' Jude 1
    If Result <> 25 Then FailTest failCount, "Jude 1", 25, Result
    Result = aeBibleCitationClass.GetMaxVerse(66, 22)        ' Revelation 22
    If Result <> 21 Then FailTest failCount, "Revelation 22", 21, Result
    ' ========================
    ' NEGATIVE TESTS
    ' ========================
    On Error Resume Next
    Err.Clear
    Result = aeBibleCitationClass.GetMaxVerse(1, 999)
    If Err.Number = 0 Then
        Debug.Print "FAIL: Invalid chapter not rejected"
        failCount = failCount + 1
    End If
    Err.Clear
    Result = aeBibleCitationClass.GetMaxVerse(999, 1)
    If Err.Number = 0 Then
        Debug.Print "FAIL: Invalid book not rejected"
        failCount = failCount + 1
    End If
    Err.Clear
    Result = aeBibleCitationClass.GetMaxVerse(19, 0)
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
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_GetMaxVerse of Module basTEST_aeBibleCitationClass"
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

    If Not VerifyPackedVerseMap() Then
        Debug.Print "ABORT: Packed verse map invalid"
        GoTo PROC_EXIT
    End If

    Set aeAssert = New aeAssertClass
    aeAssert.Initialize

    aeBibleCitationClass.ResetBookAliasMap
    Test_Stage1_AliasCoverage
    aeBibleCitationClass.Test_Stage2_LexicalScan
    aeBibleCitationClass.Test_Stage3_ResolveAlias
    aeBibleCitationClass.Test_Stage4_InterpretStructure
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
    aeBibleCitationClass.Test_Stage8_ListDetection
    aeBibleCitationClass.Test_Stage9_RangeDetection
    aeBibleCitationClass.Test_Stage10_RangeComposition
    Test_Stage11_ListComposition
    Test_Stage12_FinalParser
    Test_Stage13_ContextShorthand
    Test_Stage13a_BookContextPropagation
    Test_Stage14_CanonicalCompression
    Test_Stage15_CanonicalValidation
    Test_Stage16_CanonicalRangeBuilder
    Test_Stage17_CanonicalStringFormatter

    aeAssert.Terminate
    Set aeAssert = Nothing

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Run_All_SBL_Tests of Module basTEST_aeBibleCitationClass"
    Resume PROC_EXIT
End Sub

Public Sub Test_Stage5_ValidateCanonical()
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage5_ValidateCanonical"
    Debug.Print "------------------------------------------"

    Dim valid As Boolean
    valid = aeBibleCitationClass.ValidateSBLReference(65, "Jude", 0, "5", ModeSBL)
    aeAssert.AssertTrue valid, "Jude 5 valid"
    valid = aeBibleCitationClass.ValidateSBLReference(65, "Jude", 1, "0", ModeSBL, True)
    aeAssert.AssertTrue Not valid, "Jude 1:0 rejected"
    valid = aeBibleCitationClass.ValidateSBLReference(45, "Romans", 999, "1", ModeSBL, True)
    aeAssert.AssertTrue Not valid, "Romans 999:1 rejected"
End Sub

Public Sub Test_Stage6_FormatCanonical()
    Debug.Print ""
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage6_FormatCanonical"
    Debug.Print "------------------------------------------"

    Dim Result As String
    Result = aeBibleCitationClass.RewriteSingleChapterRef(65, 0, 5)
    aeAssert.AssertEqual "1:5", Result, "Jude single-chapter rewrite"
    Result = aeBibleCitationClass.RewriteSingleChapterRef(45, 8, 1)
    aeAssert.AssertEqual "8:1", Result, "Romans unchanged"
End Sub

Public Sub Test_Stage6_FormatCanonical_FailureDemo()
    If RUN_FAILURE_DEMOS Then
        Debug.Print "------------------------------------------"
        Debug.Print " Test_Stage6_FormatCanonical (Failure Demo)"
        Debug.Print "------------------------------------------"

        Dim Result As String
        ' Call the real formatter
        Result = aeBibleCitationClass.RewriteSingleChapterRef(65, 0, 5)   ' Actual: "1:5"
        ' Deliberate wrong expected value
        aeAssert.AssertEqual "5", Result, "Canonical rewrite for Jude"
    End If
End Sub

Public Sub Test_Stage7_EndToEnd()
    Dim Result As String
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage7_EndToEnd"
    Debug.Print "------------------------------------------"

    Result = aeBibleCitationClass.ParseReference("Jude 5")
    aeAssert.AssertEqual "Jude 1:5", Result, "Jude single-chapter expansion"
    Result = aeBibleCitationClass.ParseReference("Romans 8")
    aeAssert.AssertEqual "Romans 8", Result, "Romans chapter reference"
    Result = aeBibleCitationClass.ParseReference("3 John 4")
    aeAssert.AssertEqual "3 John 1:4", Result, "3 John expansion"
    Result = aeBibleCitationClass.ParseReference("Genesis 1:1")
    aeAssert.AssertEqual "Genesis 1:1", Result, "Genesis unchanged"
    '------------------------------------------
    ' Book-only expansion
    '------------------------------------------
    Result = aeBibleCitationClass.ParseReference("John")
    aeAssert.AssertEqual "John 1:1", Result, "John book expansion"
    Result = aeBibleCitationClass.ParseReference("1 Jn")
    aeAssert.AssertEqual "1 John 1:1", Result, "1 Jn expansion"
    Result = aeBibleCitationClass.ParseReference("Jude")
    aeAssert.AssertEqual "Jude 1:1", Result, "Jude book expansion"
    Result = aeBibleCitationClass.ParseReference("Romans")
    aeAssert.AssertEqual "Romans 1:1", Result, "Romans book expansion"
End Sub

Public Sub Test_Stage11_ListComposition()
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage11_ListComposition"
    Debug.Print "------------------------------------------"

    Dim Items As Collection
    '------------------------------------------
    ' Simple reference list
    '------------------------------------------
    Set Items = aeBibleCitationClass.ComposeList("John 3:16, John 3:18")
    aeAssert.AssertEqual 2, Items.count, "two references parsed"
    aeAssert.AssertEqual "John 3:16", Items(1), "first reference"
    aeAssert.AssertEqual "John 3:18", Items(2), "second reference"
    '------------------------------------------
    ' Range inside list
    '------------------------------------------
    Set Items = aeBibleCitationClass.ComposeList("John 3:16-18, John 3:20")
    aeAssert.AssertEqual 2, Items.count, "range + reference"
    aeAssert.AssertEqual "John 3:16-18", Items(1), "range canonical"
    aeAssert.AssertEqual "John 3:20", Items(2), "second reference"
    '------------------------------------------
    ' Semicolon separation
    '------------------------------------------
    Set Items = aeBibleCitationClass.ComposeList("Romans 8; Romans 9")
    aeAssert.AssertEqual 2, Items.count, "semicolon list"
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
    Result = aeBibleCitationClass.ParseScripture("John 3:16")
    aeAssert.AssertEqual "John 3:16", Result, "single reference"
    '------------------------------------------
    ' Range
    '------------------------------------------
    Result = aeBibleCitationClass.ParseScripture("John 3:16-18")
    aeAssert.AssertEqual "John 3:16-18", Result, "range parsed"
    '------------------------------------------
    ' List
    '------------------------------------------
    Set Items = aeBibleCitationClass.ParseScripture("John 3:16, John 3:18")
    aeAssert.AssertEqual 2, Items.count, "list parsed"
    '------------------------------------------
    ' Mixed list + range
    '------------------------------------------
    Set Items = aeBibleCitationClass.ParseScripture("John 3:16-18, John 3:20")
    aeAssert.AssertEqual 2, Items.count, "mixed parsed"
End Sub

Public Sub Test_Stage13_ContextShorthand()

    Dim c As Collection
    Dim v As Variant
    Dim i As Long

    Debug.Print "------------------------------------------"
    Debug.Print "Stage 13 Contextual Shorthand Tests"
    Debug.Print "------------------------------------------"

    '------------------------------------------
    ' Test 1
    '------------------------------------------
    Set c = aeBibleCitationClass.ComposeList("John 3:16, 18, 20-22")
    aeAssert.AssertEqual 3, c.count, "Stage13 Test1 count"
    aeAssert.AssertEqual "John 3:16", c(1), "Stage13 Test1 item 1"
    aeAssert.AssertEqual "John 3:18", c(2), "Stage13 Test1 item 2"
    aeAssert.AssertEqual "John 3:20-22", c(3), "Stage13 Test1 item 3"
    'Expected
    'John 3:16
    'John 3:18
    'John 3:20-22
    '------------------------------------------
    ' Test 2
    '------------------------------------------
    Set c = aeBibleCitationClass.ComposeList("John 3:16-4:2, 5")
    aeAssert.AssertEqual 2, c.count, "Stage13 Test2 count"
    aeAssert.AssertEqual "John 3:16-4:2", c(1), "Stage13 Test2 item 1"
    aeAssert.AssertEqual "John 4:5", c(2), "Stage13 Test2 item 2"
    'Expected
    'John 3:16-4:2
    'John 4:5
    '------------------------------------------
    ' Test 3
    '------------------------------------------
    Set c = aeBibleCitationClass.ComposeList("Romans 8; 9")
    aeAssert.AssertEqual 2, c.count, "Stage13 Test3 count"
    aeAssert.AssertEqual "Romans 8", c(1), "Stage13 Test3 item 1"
    aeAssert.AssertEqual "Romans 9", c(2), "Stage13 Test3 item 2"
    'Expected
    'Romans 8
    'Romans 9
End Sub

Public Sub Test_Stage13a_BookContextPropagation()
    On Error GoTo PROC_ERR
    Dim c As Collection
    Dim valid As Boolean
    Dim ok As Boolean

    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage13a_BookContextPropagation"
    Debug.Print "------------------------------------------"

    '------------------------------------------
    ' Positive: single-book propagation
    '------------------------------------------
    Set c = aeBibleCitationClass.ComposeList("Ps 19:1; 23:1; 28:7")
    aeAssert.AssertEqual 3, c.count, "Stage13a: 3 Psalm refs"
    aeAssert.AssertEqual "Psalms 19:1", c(1), "Stage13a: Ps 19:1"
    aeAssert.AssertEqual "Psalms 23:1", c(2), "Stage13a: Ps 23:1 inherited"
    aeAssert.AssertEqual "Psalms 28:7", c(3), "Stage13a: Ps 28:7 inherited"

    '------------------------------------------
    ' Positive: cross-book transition
    '------------------------------------------
    Set c = aeBibleCitationClass.ComposeList("Ps 103:8; Isa 40:28; 63:16")
    aeAssert.AssertEqual 3, c.count, "Stage13a: cross-book count"
    aeAssert.AssertEqual "Psalms 103:8", c(1), "Stage13a: Ps 103:8"
    aeAssert.AssertEqual "Isaiah 40:28", c(2), "Stage13a: Isa 40:28"
    aeAssert.AssertEqual "Isaiah 63:16", c(3), "Stage13a: Isa 63:16 inherited"

    '------------------------------------------
    ' Positive: range with inherited book
    '------------------------------------------
    Set c = aeBibleCitationClass.ComposeList("Ps 19:1-2; 103:8-11")
    aeAssert.AssertEqual 2, c.count, "Stage13a: Psalm range count"
    aeAssert.AssertEqual "Psalms 19:1-2", c(1), "Stage13a: Ps 19:1-2"
    aeAssert.AssertEqual "Psalms 103:8-11", c(2), "Stage13a: Ps 103:8-11 inherited"

    '------------------------------------------
    ' Negative: bad alias ("Jerimiah" misspelling)
    '------------------------------------------
    ok = False
    On Error Resume Next
    aeBibleCitationClass.ComposeList "Gen 1:1; Jerimiah 33:11; Mal 1:1"
    ok = (Err.Number <> 0)
    Err.Clear
    On Error GoTo PROC_ERR
    aeAssert.AssertTrue ok, "Stage13a neg: bad alias (Jerimiah) rejected"

    '------------------------------------------
    ' Negative: verse out of range (Ps 103:200)
    '------------------------------------------
    valid = aeBibleCitationClass.ValidateSBLReference(19, "Psalms", 103, "200", ModeSBL, True)
    aeAssert.AssertTrue Not valid, "Stage13a neg: Ps 103:200 rejected"

    '------------------------------------------
    ' Negative: chapter out of range (Jer 99:1)
    '------------------------------------------
    valid = aeBibleCitationClass.ValidateSBLReference(24, "Jeremiah", 99, "1", ModeSBL, True)
    aeAssert.AssertTrue Not valid, "Stage13a neg: Jer 99:1 rejected"

    '------------------------------------------
    ' Negative: Jude 99 — single-chapter book;
    ' Chapter=0 normalized to 1; verse 99 > max (25)
    '------------------------------------------
    valid = aeBibleCitationClass.ValidateSBLReference(65, "Jude", 0, "99", ModeSBL, True)
    aeAssert.AssertTrue Not valid, "Stage13a neg: Jude 99 rejected (max verse 25)"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_Stage13a_BookContextPropagation of Module basTEST_aeBibleCitationClass"
    Resume PROC_EXIT
End Sub

Public Sub Test_Stage14_CanonicalCompression()
    On Error GoTo PROC_ERR
    Dim Refs As Collection
    Dim Result As Collection

    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage14_CanonicalCompression"
    Debug.Print "------------------------------------------"
    '------------------------------------------
    ' Test 1 - two adjacent verses collapse to range
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "John 3:17"
    Set Result = aeBibleCitationClass.CompressCanonical(Refs)
    aeAssert.AssertEqual 1, Result.count, "Test 1: two adjacent -> one range"
    aeAssert.AssertEqual "John 3:16-3:17", Result(1), "Test 1: range value"
    '------------------------------------------
    ' Test 2 - three adjacent verses collapse to range
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "John 3:17"
    Refs.Add "John 3:18"
    Set Result = aeBibleCitationClass.CompressCanonical(Refs)
    aeAssert.AssertEqual 1, Result.count, "Test 2: three adjacent -> one range"
    aeAssert.AssertEqual "John 3:16-3:18", Result(1), "Test 2: range value"
    '------------------------------------------
    ' Test 3 - non-adjacent verses not collapsed
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "John 3:18"
    Set Result = aeBibleCitationClass.CompressCanonical(Refs)
    aeAssert.AssertEqual 2, Result.count, "Test 3: non-adjacent -> two refs"
    aeAssert.AssertEqual "John 3:16", Result(1), "Test 3: first ref"
    aeAssert.AssertEqual "John 3:18", Result(2), "Test 3: second ref"
    '------------------------------------------
    ' Test 4 - adjacent run then gap
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "John 3:17"
    Refs.Add "John 3:19"
    Set Result = aeBibleCitationClass.CompressCanonical(Refs)
    aeAssert.AssertEqual 2, Result.count, "Test 4: run then gap -> range + single"
    aeAssert.AssertEqual "John 3:16-3:17", Result(1), "Test 4: range"
    aeAssert.AssertEqual "John 3:19", Result(2), "Test 4: single after gap"
    '------------------------------------------
    ' Test 5 - cross-book not collapsed
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "Romans 8:1"
    Set Result = aeBibleCitationClass.CompressCanonical(Refs)
    aeAssert.AssertEqual 2, Result.count, "Test 5: cross-book -> two refs"
    aeAssert.AssertEqual "John 3:16", Result(1), "Test 5: John ref"
    aeAssert.AssertEqual "Romans 8:1", Result(2), "Test 5: Romans ref"
    '------------------------------------------
    ' Test 6 - cross-chapter not collapsed
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:36"
    Refs.Add "John 4:1"
    Set Result = aeBibleCitationClass.CompressCanonical(Refs)
    aeAssert.AssertEqual 2, Result.count, "Test 6: cross-chapter -> two refs"
    aeAssert.AssertEqual "John 3:36", Result(1), "Test 6: end of chapter 3"
    aeAssert.AssertEqual "John 4:1", Result(2), "Test 6: start of chapter 4"
    '------------------------------------------
    ' Test 7 - single ref passthrough
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Romans 8:1"
    Set Result = aeBibleCitationClass.CompressCanonical(Refs)
    aeAssert.AssertEqual 1, Result.count, "Test 7: single ref passthrough count"
    aeAssert.AssertEqual "Romans 8:1", Result(1), "Test 7: single ref value"
    '------------------------------------------
    ' Test 8 - multi-book mixed compression
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "John 3:17"
    Refs.Add "Romans 8:1"
    Refs.Add "Romans 8:2"
    Set Result = aeBibleCitationClass.CompressCanonical(Refs)
    aeAssert.AssertEqual 2, Result.count, "Test 8: multi-book mixed -> two ranges"
    aeAssert.AssertEqual "John 3:16-3:17", Result(1), "Test 8: John range"
    aeAssert.AssertEqual "Romans 8:1-8:2", Result(2), "Test 8: Romans range"

    Debug.Print "------------------------------------------"
    Debug.Print " Stage 14 tests complete."
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_Stage14_CanonicalCompression of Module basTEST_aeBibleCitationClass"
    Resume PROC_EXIT
End Sub

Public Sub Test_Stage15_CanonicalValidation()
    On Error GoTo PROC_ERR
    Dim Refs As Collection
    Dim Result As Collection

    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage15_CanonicalValidation"
    Debug.Print "------------------------------------------"
    '------------------------------------------
    ' Test 1 - valid single ref passes through unchanged
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 1, Result.count, "Test 1: valid single ref count"
    aeAssert.AssertEqual "John 3:16", Result(1), "Test 1: valid single ref value"
    '------------------------------------------
    ' Test 2 - invalid chapter removed
    ' Matthew has 28 chapters; ch 29 is invalid
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Matt 29:1"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 0, Result.count, "Test 2: invalid chapter removed"
    '------------------------------------------
    ' Test 3 - invalid verse removed
    ' Jude 1 has 25 verses; v 50 is invalid
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Jude 1:50"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 0, Result.count, "Test 3: invalid verse removed"
    '------------------------------------------
    ' Test 4 - valid verse at boundary kept
    ' Jude 1:25 is the last verse of Jude
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Jude 1:25"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 1, Result.count, "Test 4: boundary verse kept count"
    aeAssert.AssertEqual "Jude 1:25", Result(1), "Test 4: boundary verse kept value"
    '------------------------------------------
    ' Test 5 - range with valid bounds passes unchanged
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Gen 1:1-1:5"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 1, Result.count, "Test 5: valid range count"
    aeAssert.AssertEqual "Gen 1:1-1:5", Result(1), "Test 5: valid range value"
    '------------------------------------------
    ' Test 6 - range end verse clamped
    ' Gen 1 has 31 verses; end verse 999 clamped to 31
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Gen 1:1-1:999"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 1, Result.count, "Test 6: clamped end verse count"
    aeAssert.AssertEqual "Gen 1:1-1:31", Result(1), "Test 6: clamped end verse value"
    '------------------------------------------
    ' Test 7 - range end chapter clamped
    ' Gen has 50 chapters; end ch 999 clamped to 50.
    ' Gen 50 has 26 verses; end v 1 is within bounds.
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Gen 1:1-999:1"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 1, Result.count, "Test 7: clamped end chapter count"
    aeAssert.AssertEqual "Gen 1:1-50:1", Result(1), "Test 7: clamped end chapter value"
    '------------------------------------------
    ' Test 8 - range start chapter invalid -> removed
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Matt 29:1-29:5"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 0, Result.count, "Test 8: invalid start chapter removed"
    '------------------------------------------
    ' Test 9 - range collapses to single ref after clamping
    ' Gen 1:31-1:999 -> end clamped to 1:31 = start -> single ref
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Gen 1:31-1:999"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 1, Result.count, "Test 9: collapsed range count"
    aeAssert.AssertEqual "Gen 1:31", Result(1), "Test 9: collapsed range value"
    '------------------------------------------
    ' Test 10 - mixed valid and invalid
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Gen 1:1"
    Refs.Add "Matt 29:1"
    Refs.Add "John 3:16"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 2, Result.count, "Test 10: mixed count"
    aeAssert.AssertEqual "Gen 1:1", Result(1), "Test 10: first valid ref"
    aeAssert.AssertEqual "John 3:16", Result(2), "Test 10: second valid ref"

    Debug.Print "------------------------------------------"
    Debug.Print " Stage 15 tests complete."
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_Stage15_CanonicalValidation of Module basTEST_aeBibleCitationClass"
    Resume PROC_EXIT
End Sub

Public Sub Test_Stage16_CanonicalRangeBuilder()
    On Error GoTo PROC_ERR
    Dim Refs As Collection
    Dim Result As Collection

    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage16_CanonicalRangeBuilder"
    Debug.Print "------------------------------------------"
    '------------------------------------------
    ' Test 1 - two adjacent verses grouped into range
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "John 3:17"
    Set Result = aeBibleCitationClass.BuildCanonicalRanges(Refs)
    aeAssert.AssertEqual 1, Result.count, "Test 1: two adjacent -> one range"
    aeAssert.AssertEqual "John 3:16-3:17", Result(1), "Test 1: range value"
    '------------------------------------------
    ' Test 2 - three adjacent verses grouped into range
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "John 3:17"
    Refs.Add "John 3:18"
    Set Result = aeBibleCitationClass.BuildCanonicalRanges(Refs)
    aeAssert.AssertEqual 1, Result.count, "Test 2: three adjacent -> one range"
    aeAssert.AssertEqual "John 3:16-3:18", Result(1), "Test 2: range value"
    '------------------------------------------
    ' Test 3 - non-adjacent verses not grouped
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "John 3:18"
    Set Result = aeBibleCitationClass.BuildCanonicalRanges(Refs)
    aeAssert.AssertEqual 2, Result.count, "Test 3: non-adjacent -> two refs"
    aeAssert.AssertEqual "John 3:16", Result(1), "Test 3: first ref"
    aeAssert.AssertEqual "John 3:18", Result(2), "Test 3: second ref"
    '------------------------------------------
    ' Test 4 - adjacent run then gap
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "John 3:17"
    Refs.Add "John 3:19"
    Set Result = aeBibleCitationClass.BuildCanonicalRanges(Refs)
    aeAssert.AssertEqual 2, Result.count, "Test 4: run then gap -> range + single"
    aeAssert.AssertEqual "John 3:16-3:17", Result(1), "Test 4: range"
    aeAssert.AssertEqual "John 3:19", Result(2), "Test 4: single after gap"
    '------------------------------------------
    ' Test 5 - cross-chapter boundary not grouped
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:36"
    Refs.Add "John 4:1"
    Set Result = aeBibleCitationClass.BuildCanonicalRanges(Refs)
    aeAssert.AssertEqual 2, Result.count, "Test 5: cross-chapter -> two refs"
    aeAssert.AssertEqual "John 3:36", Result(1), "Test 5: end of chapter 3"
    aeAssert.AssertEqual "John 4:1", Result(2), "Test 5: start of chapter 4"
    '------------------------------------------
    ' Test 6 - cross-book boundary not grouped
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "Romans 8:1"
    Set Result = aeBibleCitationClass.BuildCanonicalRanges(Refs)
    aeAssert.AssertEqual 2, Result.count, "Test 6: cross-book -> two refs"
    aeAssert.AssertEqual "John 3:16", Result(1), "Test 6: John ref"
    aeAssert.AssertEqual "Romans 8:1", Result(2), "Test 6: Romans ref"
    '------------------------------------------
    ' Test 7 - single ref passthrough
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Romans 8:1"
    Set Result = aeBibleCitationClass.BuildCanonicalRanges(Refs)
    aeAssert.AssertEqual 1, Result.count, "Test 7: single ref count"
    aeAssert.AssertEqual "Romans 8:1", Result(1), "Test 7: single ref value"
    '------------------------------------------
    ' Test 8 - empty collection
    '------------------------------------------
    Set Refs = New Collection
    Set Result = aeBibleCitationClass.BuildCanonicalRanges(Refs)
    aeAssert.AssertEqual 0, Result.count, "Test 8: empty collection"
    '------------------------------------------
    ' Test 9 - multi-book mixed grouping
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "John 3:17"
    Refs.Add "Romans 8:1"
    Refs.Add "Romans 8:2"
    Set Result = aeBibleCitationClass.BuildCanonicalRanges(Refs)
    aeAssert.AssertEqual 2, Result.count, "Test 9: multi-book mixed -> two ranges"
    aeAssert.AssertEqual "John 3:16-3:17", Result(1), "Test 9: John range"
    aeAssert.AssertEqual "Romans 8:1-8:2", Result(2), "Test 9: Romans range"
    '------------------------------------------
    ' Test 10 - four consecutive verses -> single range
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Gen 1:1"
    Refs.Add "Gen 1:2"
    Refs.Add "Gen 1:3"
    Refs.Add "Gen 1:4"
    Set Result = aeBibleCitationClass.BuildCanonicalRanges(Refs)
    aeAssert.AssertEqual 1, Result.count, "Test 10: four adjacent -> one range"
    aeAssert.AssertEqual "Gen 1:1-1:4", Result(1), "Test 10: range value"

    Debug.Print "------------------------------------------"
    Debug.Print " Stage 16 tests complete."
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_Stage16_CanonicalRangeBuilder of Module basTEST_aeBibleCitationClass"
    Resume PROC_EXIT
End Sub

Public Sub Test_Stage17_CanonicalStringFormatter()
    On Error GoTo PROC_ERR
    Dim Refs As Collection
    Dim Result As String

    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage17_CanonicalStringFormatter"
    Debug.Print "------------------------------------------"
    '------------------------------------------
    ' Test 1 - single ref
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Result = aeBibleCitationClass.FormatCanonicalString(Refs)
    aeAssert.AssertEqual "John 3:16", Result, "Test 1: single ref"
    '------------------------------------------
    ' Test 2 - same-chapter range: suppress repeated chapter
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Gen 1:1-1:3"
    Result = aeBibleCitationClass.FormatCanonicalString(Refs)
    aeAssert.AssertEqual "Gen 1:1-3", Result, "Test 2: same-chapter range"
    '------------------------------------------
    ' Test 3 - two same-book same-chapter refs: comma + verse only
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Gen 1:1"
    Refs.Add "Gen 1:3"
    Result = aeBibleCitationClass.FormatCanonicalString(Refs)
    aeAssert.AssertEqual "Gen 1:1, 3", Result, "Test 3: same-chapter comma"
    '------------------------------------------
    ' Test 4 - same-book different-chapter: semicolon + ch:v, no book
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Gen 1:1"
    Refs.Add "Gen 2:1"
    Result = aeBibleCitationClass.FormatCanonicalString(Refs)
    aeAssert.AssertEqual "Gen 1:1; 2:1", Result, "Test 4: same-book chapter break"
    '------------------------------------------
    ' Test 5 - different books: semicolon + full new book ref
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Gen 1:1"
    Refs.Add "Exod 1:1"
    Result = aeBibleCitationClass.FormatCanonicalString(Refs)
    aeAssert.AssertEqual "Gen 1:1; Exod 1:1", Result, "Test 5: book break"
    '------------------------------------------
    ' Test 6 - full pipeline example from doc
    '   John 3:16-3:18, John 4:1-4:2, Romans 8:1-8:2
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16-3:18"
    Refs.Add "John 4:1-4:2"
    Refs.Add "Romans 8:1-8:2"
    Result = aeBibleCitationClass.FormatCanonicalString(Refs)
    aeAssert.AssertEqual "John 3:16-18; 4:1-2; Romans 8:1-2", Result, "Test 6: full pipeline example"
    '------------------------------------------
    ' Test 7 - empty collection returns empty string
    '------------------------------------------
    Set Refs = New Collection
    Result = aeBibleCitationClass.FormatCanonicalString(Refs)
    aeAssert.AssertEqual "", Result, "Test 7: empty collection"
    '------------------------------------------
    ' Test 8 - two same-chapter ranges: comma + verse range only
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16-3:17"
    Refs.Add "John 3:19-3:20"
    Result = aeBibleCitationClass.FormatCanonicalString(Refs)
    aeAssert.AssertEqual "John 3:16-17, 19-20", Result, "Test 8: two same-chapter ranges"
    '------------------------------------------
    ' Test 9 - same-chapter range followed by single verse
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16-3:17"
    Refs.Add "John 3:19"
    Result = aeBibleCitationClass.FormatCanonicalString(Refs)
    aeAssert.AssertEqual "John 3:16-17, 19", Result, "Test 9: range then single same chapter"
    '------------------------------------------
    ' Test 10 - three books
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Gen 1:1"
    Refs.Add "John 3:16"
    Refs.Add "Romans 8:1"
    Result = aeBibleCitationClass.FormatCanonicalString(Refs)
    aeAssert.AssertEqual "Gen 1:1; John 3:16; Romans 8:1", Result, "Test 10: three books"

    Debug.Print "------------------------------------------"
    Debug.Print " Stage 17 tests complete."
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_Stage17_CanonicalStringFormatter of Module basTEST_aeBibleCitationClass"
    Resume PROC_EXIT
End Sub
