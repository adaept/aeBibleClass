Attribute VB_Name = "basSBL_TestHarness"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Public Sub Test_SemanticFlow_NoParser()
    Debug.Print "======================================"
    Debug.Print " Test_SemanticFlow_NoParser"
    Debug.Print " (Parser intentionally NOT implemented)"
    Debug.Print "======================================"

    Dim failures As Long
    failures = 0

    ' Test cases:
    ' RawInput, Alias, ExpectedBookID, Chapter, Verse, ExpectRewrite
    Dim tests
    tests = Array( _
        Array("Jude 5", "JUDE", 65, 0, 5, True), _
        Array("Jude 1:5", "JUDE", 65, 1, 5, True), _
        Array("Obadiah 3", "OBAD", 31, 0, 3, True), _
        Array("Romans 8:1", "ROM", 45, 8, 1, False), _
        Array("Genesis 1:1", "GEN", 1, 1, 1, False) _
    )

    Dim i As Long
    For i = LBound(tests) To UBound(tests)

        Debug.Print ""
        Debug.Print "INPUT: "; tests(i)(0)

        Dim alias As String
        alias = tests(i)(1)

        Debug.Print "  ResolveBook("; alias; ")"

        Dim bookName As String
        Dim bookID As Long

        On Error GoTo ResolveFail
        bookName = ResolveBook(alias, bookID)
        On Error GoTo 0

        Debug.Print "    -> BookID:", bookID
        Debug.Print "    -> Canonical:", bookName

        Dim rewritten As String
        rewritten = RewriteSingleChapterRef( _
                        bookID, _
                        tests(i)(3), _
                        tests(i)(4))

        Debug.Print "  Chapter:", tests(i)(3)
        Debug.Print "  Verse:  ", tests(i)(4)
        Debug.Print "  Output: ", rewritten

        If tests(i)(5) Then
            If Left$(rewritten, 2) <> "1:" Then
                Debug.Print "  FAIL: expected single-chapter rewrite"
                failures = failures + 1
            End If
        Else
            If rewritten <> tests(i)(3) & ":" & tests(i)(4) Then
                Debug.Print "  FAIL: unexpected rewrite"
                failures = failures + 1
            End If
        End If

ContinueLoop:
        Debug.Print "  RESULT: PASS"
        GoTo NextTest

ResolveFail:
        Debug.Print "  ERROR: ResolveBook failed"
        failures = failures + 1
        Resume NextTest

NextTest:
    Next i

    Debug.Print ""
    Debug.Print "======================================"
    Debug.Print "FAILURES:", failures
    Debug.Print "======================================"
End Sub

Public Sub Report_TODOs()
    Debug.Print "=== NOT IMPLEMENTED / TODO ==="
    Debug.Print "- Input parser (string > book/chapter/verse)"
    Debug.Print "- Tokenization / DFA"
    Debug.Print "- Chapter bounds validation"
    Debug.Print "- Verse bounds validation"
    Debug.Print "- Punctuation normalization"
    Debug.Print "- Multi-verse ranges (e.g. John 3:16–18)"
    Debug.Print "- Error recovery / diagnostics"
    Debug.Print "=============================="
End Sub


