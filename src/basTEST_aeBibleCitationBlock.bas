Attribute VB_Name = "basTEST_aeBibleCitationBlock"
Option Explicit
Option Compare Text
Option Private Module

' =============================================================================
' basTEST_aeBibleCitationBlock
' Verifies SBL correctness of a study Bible citation block.
' Implements book-context propagation: after "Ps", references like "23:1; 28:7"
' inherit Ps from the preceding segment (standard SBL citation format).
' See md/basTEST_aeBibleCitationBlock.md for full design.
' =============================================================================

' --- Error constants ---------------------------------------------------------
Public Const E_ALIAS_UNRESOLVED As Long = 1001  ' ResolveAlias raised error
Public Const E_CHAPTER_MISSING  As Long = 1002  ' No chapter could be inferred
Public Const E_VERSE_MALFORMED  As Long = 1003  ' VerseSpec not numeric and not range
Public Const E_SBL_FAIL         As Long = 1006  ' ValidateSBLReference returned False

' --- Data structure ----------------------------------------------------------
Private Type BlockToken
    InputAlias  As String   ' e.g. "Ps", "1 Cor" -- empty string if inherited from context
    BookID      As Long     ' 0 if unresolved
    CanonName   As String   ' Canonical name from ResolveAlias
    Chapter     As Long
    StartVerse  As Long     ' after DecomposeVerseSpec
    EndVerse    As Long     ' = StartVerse if not a range
    IsRange     As Boolean
    SegText     As String   ' original segment text for error messages
    ErrorCode   As Long     ' 0 = ok; see constants above
    ErrorText   As String
End Type

' =============================================================================
' NormalizeBlockInput
' Replace CR/LF/CRLF with space; collapse multiple spaces to one.
' =============================================================================
Private Function NormalizeBlockInput(raw As String) As String
    Dim s As String
    s = raw
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    ' Collapse multiple spaces
    Dim prev As String
    Do
        prev = s
        s = Replace(s, "  ", " ")
    Loop While s <> prev
    NormalizeBlockInput = Trim$(s)
End Function

' =============================================================================
' TryResolveAlias
' Error-safe wrapper around aeBibleCitationClass.ResolveAlias.
' Returns False on any error (unresolved alias).
' =============================================================================
Private Function TryResolveAlias(alias As String, ByRef BookID As Long, ByRef CanonName As String) As Boolean
    On Error GoTo RESOLVE_FAIL
    CanonName = aeBibleCitationClass.ResolveAlias(alias, BookID)
    If CanonName = "" Then GoTo RESOLVE_FAIL
    TryResolveAlias = True
    Exit Function
RESOLVE_FAIL:
    BookID = 0
    CanonName = ""
    TryResolveAlias = False
End Function

' =============================================================================
' DetectBookAliasInSegment
' Determines whether a segment begins with a new book alias.
' Returns True if a new book was found; sets alias and refPart.
' Returns False if the segment is a bare chapter:verse inheriting context.
' =============================================================================
Private Function DetectBookAliasInSegment(seg As String, contextBookID As Long, _
    ByRef alias As String, ByRef refPart As String) As Boolean

    Dim parts() As String
    parts = Split(Trim$(seg), " ")
    If UBound(parts) < 0 Then
        DetectBookAliasInSegment = False
        Exit Function
    End If

    Dim p0 As String
    p0 = parts(0)

    ' Case 1: Single digit prefix -> try two-token alias (e.g. "1 Sam", "2 Pet")
    If Len(p0) = 1 And p0 >= "1" And p0 <= "9" Then
        If UBound(parts) >= 1 Then
            Dim twoToken As String
            twoToken = p0 & " " & parts(1)
            Dim bid1 As Long, can1 As String
            If TryResolveAlias(twoToken, bid1, can1) Then
                alias = twoToken
                If UBound(parts) >= 2 Then
                    refPart = Join(SliceArray(parts, 2), " ")
                Else
                    refPart = ""
                End If
                DetectBookAliasInSegment = True
                Exit Function
            End If
        End If
        ' Fall through — treat as bare reference inheriting context
        alias = ""
        refPart = seg
        DetectBookAliasInSegment = False
        Exit Function
    End If

    ' Case 2: Starts with letter -> try single-token alias (e.g. "Gen", "Ps", "Isa")
    If p0 Like "[A-Za-z]*" Then
        Dim bid2 As Long, can2 As String
        If TryResolveAlias(p0, bid2, can2) Then
            alias = p0
            If UBound(parts) >= 1 Then
                refPart = Join(SliceArray(parts, 1), " ")
            Else
                refPart = ""
            End If
            DetectBookAliasInSegment = True
            Exit Function
        End If
        ' Alias detection failed for a letter-prefix — fall through to bare reference
        alias = ""
        refPart = seg
        DetectBookAliasInSegment = False
        Exit Function
    End If

    ' Case 3: Multi-digit number -> bare chapter:verse continuation; inherit context
    alias = ""
    refPart = seg
    DetectBookAliasInSegment = False
End Function

' =============================================================================
' SliceArray
' Returns elements of arr from startIdx to UBound(arr).
' Helper for DetectBookAliasInSegment to avoid Join(Array()) issues.
' =============================================================================
Private Function SliceArray(arr() As String, startIdx As Long) As String()
    Dim count As Long
    count = UBound(arr) - startIdx + 1
    If count <= 0 Then
        SliceArray = Split(vbNullString)  ' returns Array(""); Join gives ""
        Exit Function
    End If
    Dim Result() As String
    ReDim Result(0 To count - 1)
    Dim i As Long
    For i = 0 To count - 1
        Result(i) = arr(startIdx + i)
    Next i
    SliceArray = Result
End Function

' =============================================================================
' DecomposeVerseSpec
' Splits "8-9" or "8{en-dash}9" into startV=8, endV=9; returns True (is range).
' For a single verse, sets startV = endV = CLng(spec); returns False.
' =============================================================================
Private Function DecomposeVerseSpec(spec As String, ByRef startV As Long, ByRef endV As Long) As Boolean
    Dim i As Long, ch As String
    For i = 1 To Len(spec)
        ch = Mid$(spec, i, 1)
        If ch = Chr(45) Or AscW(ch) = 8211 Then   ' ASCII hyphen or en dash
            Dim leftPart As String, rightPart As String
            leftPart = Left$(spec, i - 1)
            rightPart = Mid$(spec, i + 1)
            If IsNumeric(leftPart) And IsNumeric(rightPart) Then
                startV = CLng(leftPart)
                endV = CLng(rightPart)
                DecomposeVerseSpec = True
            Else
                startV = 0
                endV = 0
                DecomposeVerseSpec = False
            End If
            Exit Function
        End If
    Next i
    ' No dash found — single verse
    If IsNumeric(spec) Then
        startV = CLng(spec)
        endV = startV
    Else
        startV = 0
        endV = 0
    End If
    DecomposeVerseSpec = False
End Function

' =============================================================================
' TokenizeCitationBlock
' Stage 0: Produces flat array of BlockToken from raw citation block.
' Propagates book and chapter context across semicolon-separated segments.
' =============================================================================
Private Function TokenizeCitationBlock(raw As String) As BlockToken()
    Dim normalized As String
    normalized = NormalizeBlockInput(raw)

    Dim Segments() As String
    Segments = Split(normalized, ";")

    Dim tokens() As BlockToken
    Dim tokenCount As Long
    tokenCount = 0
    ReDim tokens(0 To 0)   ' will grow dynamically

    Dim contextBookID  As Long:   contextBookID = 0
    Dim contextCanon   As String: contextCanon = ""
    Dim contextChapter As Long:   contextChapter = 0

    Dim segIdx As Long
    For segIdx = LBound(Segments) To UBound(Segments)
        Dim seg As String
        seg = Trim$(Segments(segIdx))
        If seg = "" Then GoTo NEXT_SEG

        ' --- Detect book alias ---
        Dim detectedAlias As String, refPart As String
        Dim newBook As Boolean
        newBook = DetectBookAliasInSegment(seg, contextBookID, detectedAlias, refPart)

        If newBook Then
            Dim resolvedID As Long, resolvedCanon As String
            If TryResolveAlias(detectedAlias, resolvedID, resolvedCanon) Then
                contextBookID = resolvedID
                contextCanon = resolvedCanon
            Else
                ' Unresolved alias — emit error token and continue
                Dim errTok As BlockToken
                errTok.InputAlias = detectedAlias
                errTok.SegText = seg
                errTok.ErrorCode = E_ALIAS_UNRESOLVED
                errTok.ErrorText = "Cannot resolve alias: """ & detectedAlias & """"
                AppendToken tokens, tokenCount, errTok
                GoTo NEXT_SEG
            End If
        End If

        ' --- Parse refPart for chapter:verse ---
        Dim colonPos As Long
        colonPos = InStr(refPart, ":")
        Dim verseSpecStr As String
        If colonPos > 0 Then
            Dim chStr As String
            chStr = Trim$(Left$(refPart, colonPos - 1))
            verseSpecStr = Trim$(Mid$(refPart, colonPos + 1))
            If IsNumeric(chStr) Then contextChapter = CLng(chStr)
        Else
            ' No colon — segment is a bare verse or the refPart is something else
            ' Treat entire refPart as verse spec; chapter remains from context
            verseSpecStr = Trim$(refPart)
        End If

        If contextChapter = 0 Then
            Dim chErrTok As BlockToken
            chErrTok.InputAlias = detectedAlias
            chErrTok.BookID = contextBookID
            chErrTok.CanonName = contextCanon
            chErrTok.SegText = seg
            chErrTok.ErrorCode = E_CHAPTER_MISSING
            chErrTok.ErrorText = "No chapter could be inferred for segment: """ & seg & """"
            AppendToken tokens, tokenCount, chErrTok
            GoTo NEXT_SEG
        End If

        ' --- Split verseSpec on "," to handle "8-9,17" ---
        Dim verseSubSegs() As String
        verseSubSegs = Split(verseSpecStr, ",")

        Dim vsIdx As Long
        For vsIdx = LBound(verseSubSegs) To UBound(verseSubSegs)
            Dim vsRaw As String
            vsRaw = Trim$(verseSubSegs(vsIdx))
            If vsRaw = "" Then GoTo NEXT_VS

            Dim tok As BlockToken
            tok.InputAlias = detectedAlias
            tok.BookID = contextBookID
            tok.CanonName = contextCanon
            tok.Chapter = contextChapter
            tok.SegText = seg

            Dim sV As Long, eV As Long
            Dim isRng As Boolean
            isRng = DecomposeVerseSpec(vsRaw, sV, eV)

            If sV = 0 And Not IsNumeric(vsRaw) Then
                tok.ErrorCode = E_VERSE_MALFORMED
                tok.ErrorText = "VerseSpec not numeric and not a valid range: """ & vsRaw & """"
            Else
                tok.StartVerse = sV
                tok.EndVerse = eV
                tok.IsRange = isRng
            End If

            AppendToken tokens, tokenCount, tok
NEXT_VS:
        Next vsIdx

NEXT_SEG:
    Next segIdx

    ' Return trimmed array
    If tokenCount = 0 Then
        ReDim tokens(0 To -1)
    Else
        ReDim Preserve tokens(0 To tokenCount - 1)
    End If
    TokenizeCitationBlock = tokens
End Function

' =============================================================================
' AppendToken
' Grows the token array by one and stores tok at tokenCount; increments count.
' =============================================================================
Private Sub AppendToken(ByRef tokens() As BlockToken, ByRef tokenCount As Long, tok As BlockToken)
    If tokenCount = 0 Then
        ReDim tokens(0 To 0)
    Else
        ReDim Preserve tokens(0 To tokenCount)
    End If
    tokens(tokenCount) = tok
    tokenCount = tokenCount + 1
End Sub

' =============================================================================
' FormatTokenRef
' Formats a BlockToken as a canonical reference string for display.
' =============================================================================
Private Function FormatTokenRef(t As BlockToken) As String
    Dim s As String
    s = t.CanonName & " " & t.Chapter & ":" & t.StartVerse
    If t.IsRange Then s = s & "-" & t.EndVerse
    FormatTokenRef = s
End Function

' =============================================================================
' VerifyCitationBlock  (Public)
' Tokenizes rawBlock; validates each atomic verse endpoint via
' aeBibleCitationClass.ValidateSBLReference(ModeSBL); prints PASS/FAIL.
' =============================================================================
Public Sub VerifyCitationBlock(rawBlock As String)
    Dim tokens() As BlockToken
    tokens = TokenizeCitationBlock(rawBlock)

    Dim passCount As Long, failCount As Long
    passCount = 0
    failCount = 0

    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        Dim t As BlockToken
        t = tokens(i)

        ' Pre-tokenization error
        If t.ErrorCode <> 0 Then
            Debug.Print "FAIL [" & t.ErrorCode & "]: " & t.ErrorText & " (segment: """ & t.SegText & """)"
            failCount = failCount + 1
            GoTo NEXT_TOK
        End If

        ' Validate start verse
        Dim okStart As Boolean
        okStart = aeBibleCitationClass.ValidateSBLReference( _
            t.BookID, t.CanonName, t.Chapter, CStr(t.StartVerse), ModeSBL, True)
        If Not okStart Then
            Debug.Print "FAIL [" & E_SBL_FAIL & "]: " & FormatTokenRef(t) & " (start verse failed ValidateSBLReference)"
            failCount = failCount + 1
            GoTo NEXT_TOK
        End If

        ' Validate end verse if range
        If t.IsRange Then
            Dim okEnd As Boolean
            okEnd = aeBibleCitationClass.ValidateSBLReference( _
                t.BookID, t.CanonName, t.Chapter, CStr(t.EndVerse), ModeSBL, True)
            If Not okEnd Then
                Debug.Print "FAIL [" & E_SBL_FAIL & "]: " & FormatTokenRef(t) & " (end verse " & t.EndVerse & " failed ValidateSBLReference)"
                failCount = failCount + 1
                GoTo NEXT_TOK
            End If
        End If

        Debug.Print "PASS: " & FormatTokenRef(t)
        passCount = passCount + 1

NEXT_TOK:
    Next i

    Debug.Print "--- " & passCount & " passed, " & failCount & " failed. ---"
End Sub

' =============================================================================
' Test_VerifyCitationBlock  (Public)
' Positive test — full 35-token citation block; expected: all tokens pass.
' En dashes constructed via Chr(8211) to keep source file ASCII-safe.
' =============================================================================
Public Sub Test_VerifyCitationBlock()
    Dim rawBlock As String
    rawBlock = "Gen 1:27; Num 14:18; Deut 32:6; Josh 1:9; 1 Sam 2:2; " & _
               "1 Chr 29:10" & ChrW(8211) & "13; " & _
               "Ps 19:1" & ChrW(8211) & "2; 23:1; 28:7; 68:5; " & _
               "103:8" & ChrW(8211) & "11; 111:3" & ChrW(8211) & "5; " & _
               "145:8" & ChrW(8211) & "9,17; Isa 40:28; 63:16; 64:8; " & _
               "Jer 33:11; Nah 1:3; Mal 2:10" & ChrW(8211) & "15; " & _
               "Matt 6:9; 7:11; 23:9; John 3:16; 4:24; " & _
               "Rom 1:20; 8:15; 1 Cor 8:6; 14:33; Gal 3:20; Eph 4:6; " & _
               "Heb 13:6; 1 Pet 1:17; 2 Pet 3:9; 1 John 4:16"
    Debug.Print "=== Test_VerifyCitationBlock (positive, 35 tokens expected) ==="
    VerifyCitationBlock rawBlock
End Sub

' =============================================================================
' Test_VerifyCitationBlock_Negative  (Public)
' 3-case negative test: bad alias, verse out of range, chapter out of range.
' Expected: 3 failures.
' =============================================================================
Public Sub Test_VerifyCitationBlock_Negative()
    Debug.Print "=== Test_VerifyCitationBlock_Negative (3 failures expected) ==="

    ' Case 1: Bad alias — "Jerimiah" is a misspelling; confirmed absent from alias map
    Dim rawBadAlias As String
    rawBadAlias = "Gen 1:1; Jerimiah 33:11; Mal 1:1"
    Debug.Print "--- Case 1: Bad alias (Jerimiah) ---"
    VerifyCitationBlock rawBadAlias

    ' Case 2: Verse out of range — Ps 103 has 22 verses; verse 200 is invalid
    Dim rawBadVerse As String
    rawBadVerse = "Ps 103:8" & ChrW(8211) & "200"
    Debug.Print "--- Case 2: Verse out of range (Ps 103:8-200) ---"
    VerifyCitationBlock rawBadVerse

    ' Case 3: Chapter out of range — Jeremiah has 52 chapters; chapter 99 is invalid
    Dim rawBadChapter As String
    rawBadChapter = "Jer 99:1"
    Debug.Print "--- Case 3: Chapter out of range (Jer 99:1) ---"
    VerifyCitationBlock rawBadChapter
End Sub
