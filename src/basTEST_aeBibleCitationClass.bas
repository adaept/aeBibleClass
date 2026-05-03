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

Public Sub Run_All_SBL_Tests()
    On Error GoTo PROC_ERR

    Dim log As New aeLoggerClass
    log.Log_Init ActiveDocument.Path & "\rpt\SBL_Tests.UTF8.txt"

    If Not VerifyPackedVerseMap() Then
        Debug.Print "ABORT: Packed verse map invalid"
        log.Log_Write "ABORT: Packed verse map invalid"
        log.Log_Close
        GoTo PROC_EXIT
    End If

    Set aeAssert = New aeAssertClass
    aeAssert.SetLogger log
    aeAssert.Initialize

    aeBibleCitationClass.ResetBookAliasMap
    Test_Stage1_AliasCoverage
    Test_SongOfSongs_AllAliases
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

    Run_Extra_Tests

    log.Log_Close
    Set log = Nothing

PROC_EXIT:
    Exit Sub
PROC_ERR:
    If Not log Is Nothing Then log.Log_Close
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Run_All_SBL_Tests of Module basTEST_aeBibleCitationClass"
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

Public Sub Test_SongOfSongs_AllAliases()
' Focused coverage for book 22 after the 2026-04-30 rename of project canonical
' from "Solomon" to "Song of Songs". Verifies every documented alias resolves to
' BookID 22 with canonical name "Song of Songs"; verifies the WEB-aligned
' verse-Count data; verifies ToSBLShortForm yields the SBL "Song N:V" output.
'
' Uses the suite-level aeAssert (initialized by Run_All_SBL_Tests).
'
    Debug.Print "------------------------------------------"
    Debug.Print " Test_SongOfSongs_AllAliases"
    Debug.Print "------------------------------------------"

    ' All documented aliases resolve to BookID 22 with canonical name.
    Dim aliases As Variant
    aliases = Array( _
        "Song of Songs", "song of songs", "SONG OF SONGS", _
        "Song", "Son", "Sg", _
        "Solomon", "Solo", "Sol", "So", _
        "Canticles", "Cant", "Can")

    Dim i As Long, bID As Long
    Dim canonName As String
    For i = LBound(aliases) To UBound(aliases)
        bID = 0
        canonName = aeBibleCitationClass.ResolveAlias(CStr(aliases(i)), bID)
        aeAssert.AssertEqual 22, bID, _
            "ResolveAlias(""" & aliases(i) & """) BookID"
        aeAssert.AssertEqual "Song of Songs", canonName, _
            "ResolveAlias(""" & aliases(i) & """) canonical name"
    Next i

    ' Negative: "Song of Solomon" (multi-word) is not in the alias map.
    On Error Resume Next
    bID = 0
    canonName = aeBibleCitationClass.ResolveAlias("Song of Solomon", bID)
    Dim raised As Boolean
    raised = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    aeAssert.AssertTrue raised, _
        "ResolveAlias(""Song of Solomon"") raises (multi-word not in alias map)"

    ' ChaptersInBook via canonical and two aliases.
    aeAssert.AssertEqual 8, aeBibleCitationClass.ChaptersInBook("Song of Songs"), _
        "ChaptersInBook(""Song of Songs"")"
    aeAssert.AssertEqual 8, aeBibleCitationClass.ChaptersInBook("Song"), _
        "ChaptersInBook(""Song"") via alias"
    aeAssert.AssertEqual 8, aeBibleCitationClass.ChaptersInBook("Solomon"), _
        "ChaptersInBook(""Solomon"") via alias"

    ' VersesInChapter - WEB-aligned counts: 17, 17, 11, 16, 16, 13, 13, 14.
    Dim verseCounts As Variant
    verseCounts = Array(17, 17, 11, 16, 16, 13, 13, 14)
    Dim ch As Long
    For ch = 1 To 8
        aeAssert.AssertEqual CLng(verseCounts(ch - 1)), _
            aeBibleCitationClass.VersesInChapter("Song of Songs", ch), _
            "VersesInChapter(""Song of Songs"", " & ch & ")"
    Next ch

    ' ToSBLShortForm.
    aeAssert.AssertEqual "Song 2:4", _
        aeBibleCitationClass.ToSBLShortForm("Song of Songs 2:4"), _
        "ToSBLShortForm(""Song of Songs 2:4"")"
    aeAssert.AssertEqual "Song 1:1", _
        aeBibleCitationClass.ToSBLShortForm("Song of Songs 1:1"), _
        "ToSBLShortForm(""Song of Songs 1:1"")"

    Debug.Print " Test_SongOfSongs_AllAliases complete."
    Debug.Print "------------------------------------------"
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

Public Sub Test_Stage11_ListComposition()
    Debug.Print "------------------------------------------"
    Debug.Print " Test_Stage11_ListComposition"
    Debug.Print "------------------------------------------"

    Dim Items As Collection
    '------------------------------------------
    ' Simple reference list
    '------------------------------------------
    Set Items = aeBibleCitationClass.ComposeList("John 3:16, John 3:18")
    aeAssert.AssertEqual 2, Items.Count, "two references parsed"
    aeAssert.AssertEqual "John 3:16", Items(1), "first reference"
    aeAssert.AssertEqual "John 3:18", Items(2), "second reference"
    '------------------------------------------
    ' Range inside list
    '------------------------------------------
    Set Items = aeBibleCitationClass.ComposeList("John 3:16-18, John 3:20")
    aeAssert.AssertEqual 2, Items.Count, "range + reference"
    aeAssert.AssertEqual "John 3:16-18", Items(1), "range canonical"
    aeAssert.AssertEqual "John 3:20", Items(2), "second reference"
    '------------------------------------------
    ' Semicolon separation
    '------------------------------------------
    Set Items = aeBibleCitationClass.ComposeList("Romans 8; Romans 9")
    aeAssert.AssertEqual 2, Items.Count, "semicolon list"
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
    aeAssert.AssertEqual 2, Items.Count, "list parsed"
    '------------------------------------------
    ' Mixed list + range
    '------------------------------------------
    Set Items = aeBibleCitationClass.ParseScripture("John 3:16-18, John 3:20")
    aeAssert.AssertEqual 2, Items.Count, "mixed parsed"
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
    aeAssert.AssertEqual 3, c.Count, "Stage13 Test1 Count"
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
    aeAssert.AssertEqual 2, c.Count, "Stage13 Test2 Count"
    aeAssert.AssertEqual "John 3:16-4:2", c(1), "Stage13 Test2 item 1"
    aeAssert.AssertEqual "John 4:5", c(2), "Stage13 Test2 item 2"
    'Expected
    'John 3:16-4:2
    'John 4:5
    '------------------------------------------
    ' Test 3
    '------------------------------------------
    Set c = aeBibleCitationClass.ComposeList("Romans 8; 9")
    aeAssert.AssertEqual 2, c.Count, "Stage13 Test3 Count"
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
    aeAssert.AssertEqual 3, c.Count, "Stage13a: 3 Psalm refs"
    aeAssert.AssertEqual "Psalms 19:1", c(1), "Stage13a: Ps 19:1"
    aeAssert.AssertEqual "Psalms 23:1", c(2), "Stage13a: Ps 23:1 inherited"
    aeAssert.AssertEqual "Psalms 28:7", c(3), "Stage13a: Ps 28:7 inherited"

    '------------------------------------------
    ' Positive: cross-book transition
    '------------------------------------------
    Set c = aeBibleCitationClass.ComposeList("Ps 103:8; Isa 40:28; 63:16")
    aeAssert.AssertEqual 3, c.Count, "Stage13a: cross-book Count"
    aeAssert.AssertEqual "Psalms 103:8", c(1), "Stage13a: Ps 103:8"
    aeAssert.AssertEqual "Isaiah 40:28", c(2), "Stage13a: Isa 40:28"
    aeAssert.AssertEqual "Isaiah 63:16", c(3), "Stage13a: Isa 63:16 inherited"

    '------------------------------------------
    ' Positive: range with inherited book
    '------------------------------------------
    Set c = aeBibleCitationClass.ComposeList("Ps 19:1-2; 103:8-11")
    aeAssert.AssertEqual 2, c.Count, "Stage13a: Psalm range Count"
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
    aeAssert.AssertEqual 1, Result.Count, "Test 1: two adjacent -> one range"
    aeAssert.AssertEqual "John 3:16-3:17", Result(1), "Test 1: range value"
    '------------------------------------------
    ' Test 2 - three adjacent verses collapse to range
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "John 3:17"
    Refs.Add "John 3:18"
    Set Result = aeBibleCitationClass.CompressCanonical(Refs)
    aeAssert.AssertEqual 1, Result.Count, "Test 2: three adjacent -> one range"
    aeAssert.AssertEqual "John 3:16-3:18", Result(1), "Test 2: range value"
    '------------------------------------------
    ' Test 3 - non-adjacent verses not collapsed
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "John 3:18"
    Set Result = aeBibleCitationClass.CompressCanonical(Refs)
    aeAssert.AssertEqual 2, Result.Count, "Test 3: non-adjacent -> two refs"
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
    aeAssert.AssertEqual 2, Result.Count, "Test 4: run then gap -> range + single"
    aeAssert.AssertEqual "John 3:16-3:17", Result(1), "Test 4: range"
    aeAssert.AssertEqual "John 3:19", Result(2), "Test 4: single after gap"
    '------------------------------------------
    ' Test 5 - cross-book not collapsed
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "Romans 8:1"
    Set Result = aeBibleCitationClass.CompressCanonical(Refs)
    aeAssert.AssertEqual 2, Result.Count, "Test 5: cross-book -> two refs"
    aeAssert.AssertEqual "John 3:16", Result(1), "Test 5: John ref"
    aeAssert.AssertEqual "Romans 8:1", Result(2), "Test 5: Romans ref"
    '------------------------------------------
    ' Test 6 - cross-chapter not collapsed
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:36"
    Refs.Add "John 4:1"
    Set Result = aeBibleCitationClass.CompressCanonical(Refs)
    aeAssert.AssertEqual 2, Result.Count, "Test 6: cross-chapter -> two refs"
    aeAssert.AssertEqual "John 3:36", Result(1), "Test 6: end of chapter 3"
    aeAssert.AssertEqual "John 4:1", Result(2), "Test 6: start of chapter 4"
    '------------------------------------------
    ' Test 7 - single ref passthrough
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Romans 8:1"
    Set Result = aeBibleCitationClass.CompressCanonical(Refs)
    aeAssert.AssertEqual 1, Result.Count, "Test 7: single ref passthrough Count"
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
    aeAssert.AssertEqual 2, Result.Count, "Test 8: multi-book mixed -> two ranges"
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
    aeAssert.AssertEqual 1, Result.Count, "Test 1: valid single ref Count"
    aeAssert.AssertEqual "John 3:16", Result(1), "Test 1: valid single ref value"
    '------------------------------------------
    ' Test 2 - invalid chapter removed
    ' Matthew has 28 chapters; ch 29 is invalid
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Matt 29:1"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 0, Result.Count, "Test 2: invalid chapter removed"
    '------------------------------------------
    ' Test 3 - invalid verse removed
    ' Jude 1 has 25 verses; v 50 is invalid
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Jude 1:50"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 0, Result.Count, "Test 3: invalid verse removed"
    '------------------------------------------
    ' Test 4 - valid verse at boundary kept
    ' Jude 1:25 is the last verse of Jude
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Jude 1:25"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 1, Result.Count, "Test 4: boundary verse kept Count"
    aeAssert.AssertEqual "Jude 1:25", Result(1), "Test 4: boundary verse kept value"
    '------------------------------------------
    ' Test 5 - range with valid bounds passes unchanged
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Gen 1:1-1:5"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 1, Result.Count, "Test 5: valid range Count"
    aeAssert.AssertEqual "Gen 1:1-1:5", Result(1), "Test 5: valid range value"
    '------------------------------------------
    ' Test 6 - range end verse clamped
    ' Gen 1 has 31 verses; end verse 999 clamped to 31
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Gen 1:1-1:999"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 1, Result.Count, "Test 6: clamped end verse Count"
    aeAssert.AssertEqual "Gen 1:1-1:31", Result(1), "Test 6: clamped end verse value"
    '------------------------------------------
    ' Test 7 - range end chapter clamped
    ' Gen has 50 chapters; end ch 999 clamped to 50.
    ' Gen 50 has 26 verses; end v 1 is within bounds.
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Gen 1:1-999:1"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 1, Result.Count, "Test 7: clamped end chapter Count"
    aeAssert.AssertEqual "Gen 1:1-50:1", Result(1), "Test 7: clamped end chapter value"
    '------------------------------------------
    ' Test 8 - range start chapter invalid -> removed
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Matt 29:1-29:5"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 0, Result.Count, "Test 8: invalid start chapter removed"
    '------------------------------------------
    ' Test 9 - range collapses to single ref after clamping
    ' Gen 1:31-1:999 -> end clamped to 1:31 = start -> single ref
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Gen 1:31-1:999"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 1, Result.Count, "Test 9: collapsed range Count"
    aeAssert.AssertEqual "Gen 1:31", Result(1), "Test 9: collapsed range value"
    '------------------------------------------
    ' Test 10 - mixed valid and invalid
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Gen 1:1"
    Refs.Add "Matt 29:1"
    Refs.Add "John 3:16"
    Set Result = aeBibleCitationClass.ValidateCanonical(Refs)
    aeAssert.AssertEqual 2, Result.Count, "Test 10: mixed Count"
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
    aeAssert.AssertEqual 1, Result.Count, "Test 1: two adjacent -> one range"
    aeAssert.AssertEqual "John 3:16-3:17", Result(1), "Test 1: range value"
    '------------------------------------------
    ' Test 2 - three adjacent verses grouped into range
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "John 3:17"
    Refs.Add "John 3:18"
    Set Result = aeBibleCitationClass.BuildCanonicalRanges(Refs)
    aeAssert.AssertEqual 1, Result.Count, "Test 2: three adjacent -> one range"
    aeAssert.AssertEqual "John 3:16-3:18", Result(1), "Test 2: range value"
    '------------------------------------------
    ' Test 3 - non-adjacent verses not grouped
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "John 3:18"
    Set Result = aeBibleCitationClass.BuildCanonicalRanges(Refs)
    aeAssert.AssertEqual 2, Result.Count, "Test 3: non-adjacent -> two refs"
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
    aeAssert.AssertEqual 2, Result.Count, "Test 4: run then gap -> range + single"
    aeAssert.AssertEqual "John 3:16-3:17", Result(1), "Test 4: range"
    aeAssert.AssertEqual "John 3:19", Result(2), "Test 4: single after gap"
    '------------------------------------------
    ' Test 5 - cross-chapter boundary not grouped
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:36"
    Refs.Add "John 4:1"
    Set Result = aeBibleCitationClass.BuildCanonicalRanges(Refs)
    aeAssert.AssertEqual 2, Result.Count, "Test 5: cross-chapter -> two refs"
    aeAssert.AssertEqual "John 3:36", Result(1), "Test 5: end of chapter 3"
    aeAssert.AssertEqual "John 4:1", Result(2), "Test 5: start of chapter 4"
    '------------------------------------------
    ' Test 6 - cross-book boundary not grouped
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "Romans 8:1"
    Set Result = aeBibleCitationClass.BuildCanonicalRanges(Refs)
    aeAssert.AssertEqual 2, Result.Count, "Test 6: cross-book -> two refs"
    aeAssert.AssertEqual "John 3:16", Result(1), "Test 6: John ref"
    aeAssert.AssertEqual "Romans 8:1", Result(2), "Test 6: Romans ref"
    '------------------------------------------
    ' Test 7 - single ref passthrough
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "Romans 8:1"
    Set Result = aeBibleCitationClass.BuildCanonicalRanges(Refs)
    aeAssert.AssertEqual 1, Result.Count, "Test 7: single ref Count"
    aeAssert.AssertEqual "Romans 8:1", Result(1), "Test 7: single ref value"
    '------------------------------------------
    ' Test 8 - empty collection
    '------------------------------------------
    Set Refs = New Collection
    Set Result = aeBibleCitationClass.BuildCanonicalRanges(Refs)
    aeAssert.AssertEqual 0, Result.Count, "Test 8: empty collection"
    '------------------------------------------
    ' Test 9 - multi-book mixed grouping
    '------------------------------------------
    Set Refs = New Collection
    Refs.Add "John 3:16"
    Refs.Add "John 3:17"
    Refs.Add "Romans 8:1"
    Refs.Add "Romans 8:2"
    Set Result = aeBibleCitationClass.BuildCanonicalRanges(Refs)
    aeAssert.AssertEqual 2, Result.Count, "Test 9: multi-book mixed -> two ranges"
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
    aeAssert.AssertEqual 1, Result.Count, "Test 10: four adjacent -> one range"
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

'=====================================================
' Test_CanonicalNamesAndSBLTable
'   Validates all 66 canonical book names via
'   ChaptersInBook and all 66 SBL abbreviations via
'   ToSBLShortForm against known expected values.
'   Run via Alt+F8 or from Run_All_SBL_Tests.
'=====================================================
Public Sub Test_CanonicalNamesAndSBLTable()
    On Error GoTo PROC_ERR

    Dim assert As New aeAssertClass
    Dim log As New aeLoggerClass
    log.Log_Init ActiveDocument.Path & "\rpt\TestReport.txt"
    assert.SetLogger log

    Debug.Print "------------------------------------------"
    Debug.Print " Test_CanonicalNamesAndSBLTable"

    ' Expected: canonical name -> chapter Count (66 books)
    Dim canonNames(1 To 66) As String
    Dim canonChapters(1 To 66) As Long

    canonNames(1) = "Genesis":        canonChapters(1) = 50
    canonNames(2) = "Exodus":         canonChapters(2) = 40
    canonNames(3) = "Leviticus":      canonChapters(3) = 27
    canonNames(4) = "Numbers":        canonChapters(4) = 36
    canonNames(5) = "Deuteronomy":    canonChapters(5) = 34
    canonNames(6) = "Joshua":         canonChapters(6) = 24
    canonNames(7) = "Judges":         canonChapters(7) = 21
    canonNames(8) = "Ruth":           canonChapters(8) = 4
    canonNames(9) = "1 Samuel":       canonChapters(9) = 31
    canonNames(10) = "2 Samuel":      canonChapters(10) = 24
    canonNames(11) = "1 Kings":       canonChapters(11) = 22
    canonNames(12) = "2 Kings":       canonChapters(12) = 25
    canonNames(13) = "1 Chronicles":  canonChapters(13) = 29
    canonNames(14) = "2 Chronicles":  canonChapters(14) = 36
    canonNames(15) = "Ezra":          canonChapters(15) = 10
    canonNames(16) = "Nehemiah":      canonChapters(16) = 13
    canonNames(17) = "Esther":        canonChapters(17) = 10
    canonNames(18) = "Job":           canonChapters(18) = 42
    canonNames(19) = "Psalms":        canonChapters(19) = 150
    canonNames(20) = "Proverbs":      canonChapters(20) = 31
    canonNames(21) = "Ecclesiastes":  canonChapters(21) = 12
    canonNames(22) = "Song of Songs": canonChapters(22) = 8   ' Project canonical; SBL output is "Song"
    canonNames(23) = "Isaiah":        canonChapters(23) = 66
    canonNames(24) = "Jeremiah":      canonChapters(24) = 52
    canonNames(25) = "Lamentations":  canonChapters(25) = 5
    canonNames(26) = "Ezekiel":       canonChapters(26) = 48
    canonNames(27) = "Daniel":        canonChapters(27) = 12
    canonNames(28) = "Hosea":         canonChapters(28) = 14
    canonNames(29) = "Joel":          canonChapters(29) = 3
    canonNames(30) = "Amos":          canonChapters(30) = 9
    canonNames(31) = "Obadiah":       canonChapters(31) = 1
    canonNames(32) = "Jonah":         canonChapters(32) = 4
    canonNames(33) = "Micah":         canonChapters(33) = 7
    canonNames(34) = "Nahum":         canonChapters(34) = 3
    canonNames(35) = "Habakkuk":      canonChapters(35) = 3
    canonNames(36) = "Zephaniah":     canonChapters(36) = 3
    canonNames(37) = "Haggai":        canonChapters(37) = 2
    canonNames(38) = "Zechariah":     canonChapters(38) = 14
    canonNames(39) = "Malachi":       canonChapters(39) = 4
    canonNames(40) = "Matthew":       canonChapters(40) = 28
    canonNames(41) = "Mark":          canonChapters(41) = 16
    canonNames(42) = "Luke":          canonChapters(42) = 24
    canonNames(43) = "John":          canonChapters(43) = 21
    canonNames(44) = "Acts":          canonChapters(44) = 28
    canonNames(45) = "Romans":        canonChapters(45) = 16
    canonNames(46) = "1 Corinthians": canonChapters(46) = 16
    canonNames(47) = "2 Corinthians": canonChapters(47) = 13
    canonNames(48) = "Galatians":     canonChapters(48) = 6
    canonNames(49) = "Ephesians":     canonChapters(49) = 6
    canonNames(50) = "Philippians":   canonChapters(50) = 4
    canonNames(51) = "Colossians":    canonChapters(51) = 4
    canonNames(52) = "1 Thessalonians": canonChapters(52) = 5
    canonNames(53) = "2 Thessalonians": canonChapters(53) = 3
    canonNames(54) = "1 Timothy":     canonChapters(54) = 6
    canonNames(55) = "2 Timothy":     canonChapters(55) = 4
    canonNames(56) = "Titus":         canonChapters(56) = 3
    canonNames(57) = "Philemon":      canonChapters(57) = 1
    canonNames(58) = "Hebrews":       canonChapters(58) = 13
    canonNames(59) = "James":         canonChapters(59) = 5
    canonNames(60) = "1 Peter":       canonChapters(60) = 5
    canonNames(61) = "2 Peter":       canonChapters(61) = 3
    canonNames(62) = "1 John":        canonChapters(62) = 5
    canonNames(63) = "2 John":        canonChapters(63) = 1
    canonNames(64) = "3 John":        canonChapters(64) = 1
    canonNames(65) = "Jude":          canonChapters(65) = 1
    canonNames(66) = "Revelation":    canonChapters(66) = 22

    ' Expected SBL abbreviation for each book (index 1-66)
    Dim sblExpected(1 To 66) As String
    sblExpected(1) = "Gen":   sblExpected(2) = "Exod":  sblExpected(3) = "Lev"
    sblExpected(4) = "Num":   sblExpected(5) = "Deut":  sblExpected(6) = "Josh"
    sblExpected(7) = "Judg":  sblExpected(8) = "Ruth":  sblExpected(9) = "1 Sam"
    sblExpected(10) = "2 Sam": sblExpected(11) = "1 Kgs": sblExpected(12) = "2 Kgs"
    sblExpected(13) = "1 Chr": sblExpected(14) = "2 Chr": sblExpected(15) = "Ezra"
    sblExpected(16) = "Neh":  sblExpected(17) = "Esth": sblExpected(18) = "Job"
    sblExpected(19) = "Ps":   sblExpected(20) = "Prov": sblExpected(21) = "Eccl"
    sblExpected(22) = "Song": sblExpected(23) = "Isa":  sblExpected(24) = "Jer"
    sblExpected(25) = "Lam":  sblExpected(26) = "Ezek": sblExpected(27) = "Dan"
    sblExpected(28) = "Hos":  sblExpected(29) = "Joel": sblExpected(30) = "Amos"
    sblExpected(31) = "Obad": sblExpected(32) = "Jonah": sblExpected(33) = "Mic"
    sblExpected(34) = "Nah":  sblExpected(35) = "Hab":  sblExpected(36) = "Zeph"
    sblExpected(37) = "Hag":  sblExpected(38) = "Zech": sblExpected(39) = "Mal"
    sblExpected(40) = "Matt": sblExpected(41) = "Mark": sblExpected(42) = "Luke"
    sblExpected(43) = "John": sblExpected(44) = "Acts": sblExpected(45) = "Rom"
    sblExpected(46) = "1 Cor": sblExpected(47) = "2 Cor": sblExpected(48) = "Gal"
    sblExpected(49) = "Eph":  sblExpected(50) = "Phil": sblExpected(51) = "Col"
    sblExpected(52) = "1 Thess": sblExpected(53) = "2 Thess": sblExpected(54) = "1 Tim"
    sblExpected(55) = "2 Tim": sblExpected(56) = "Titus": sblExpected(57) = "Phlm"
    sblExpected(58) = "Heb":  sblExpected(59) = "Jas":  sblExpected(60) = "1 Pet"
    sblExpected(61) = "2 Pet": sblExpected(62) = "1 John": sblExpected(63) = "2 John"
    sblExpected(64) = "3 John": sblExpected(65) = "Jude": sblExpected(66) = "Rev"

    Dim i As Long
    For i = 1 To 66
        ' Validate ChaptersInBook resolves each canonical name to the correct chapter Count
        Dim gotCh As Long
        gotCh = aeBibleCitationClass.ChaptersInBook(canonNames(i))
        assert.AssertEqual canonChapters(i), gotCh, _
            "ChaptersInBook(" & canonNames(i) & ") expected " & canonChapters(i)

        ' Validate ToSBLShortForm produces the correct abbreviation for book i chapter 1 verse 1
        ' Single-chapter books omit the chapter; use chapter 1 verse 1 for all others.
        Dim sblInput As String
        Dim sblResult As String
        Dim sblResultBook As String
        sblInput = canonNames(i) & " 1:1"
        sblResult = aeBibleCitationClass.ToSBLShortForm(sblInput)
        ' Extract book abbreviation: everything before the first digit in Result
        Dim spacePos As Long
        spacePos = InStr(sblResult, " ")
        If spacePos > 0 Then
            sblResultBook = Left$(sblResult, spacePos - 1)
        Else
            sblResultBook = sblResult   ' single-chapter book with verse only (should not occur for 1:1)
        End If
        assert.AssertEqual sblExpected(i), sblResultBook, _
            "ToSBLShortForm book abbr for " & canonNames(i) & " expected " & sblExpected(i)
    Next i

    Debug.Print " Test_CanonicalNamesAndSBLTable complete."
    Debug.Print "------------------------------------------"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_CanonicalNamesAndSBLTable of Module basTEST_aeBibleCitationClass"
    Resume PROC_EXIT
End Sub
