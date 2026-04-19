Attribute VB_Name = "basTEST_aeBibleCitationBlock"
Option Explicit
Option Compare Text
Option Private Module

' =============================================================================
' basTEST_aeBibleCitationBlock
' Integration tests for study Bible citation blocks.
' Block parsing and context propagation are handled by
' aeBibleCitationClass.ParseCitationBlock.
' See md/basTEST_aeBibleCitationBlock.md for design background.
' =============================================================================

' =============================================================================
' VerifyCitationBlock  (Public)
' Calls ParseCitationBlock to resolve the block, then validates each canonical
' reference via ValidateSBLReference(ModeSBL). Prints PASS/FAIL per item.
' Returns failCount. Raises on unresolved book alias (propagated from class).
' =============================================================================
Public Function VerifyCitationBlock(rawBlock As String) As Long
    On Error GoTo PROC_ERR
    Dim Items As Collection
    Dim passCount As Long
    Dim failCount As Long

    Set Items = aeBibleCitationClass.SortCitationBlock( _
        aeBibleCitationClass.ParseCitationBlock(rawBlock))

    Dim Item As Variant
    For Each Item In Items
        Dim canon As String
        canon = CStr(Item)

        ' Find last space to split "Book Name Chapter:Verse[-EndVerse]"
        Dim lastSp As Long
        Dim k As Long
        For k = Len(canon) To 1 Step -1
            If Mid$(canon, k, 1) = " " Then lastSp = k: Exit For
        Next k

        Dim bookName As String
        bookName = Left$(canon, lastSp - 1)
        Dim numPart As String
        numPart = Mid$(canon, lastSp + 1)

        ' Resolve BookID from canonical name
        Dim bID As Long
        Dim bCanon As String
        bCanon = aeBibleCitationClass.ResolveAlias(bookName, bID)

        ' Parse chapter:startVerse[-endVerse]
        Dim cpPos As Long
        cpPos = InStr(numPart, ":")
        Dim ch As Long
        ch = CLng(Left$(numPart, cpPos - 1))
        Dim vPart As String
        vPart = Mid$(numPart, cpPos + 1)

        Dim dpPos As Long
        dpPos = InStr(vPart, "-")
        Dim startV As Long
        Dim endV As Long
        Dim isRng As Boolean
        If dpPos > 0 Then
            startV = CLng(Left$(vPart, dpPos - 1))
            endV = CLng(Mid$(vPart, dpPos + 1))
            isRng = True
        Else
            startV = CLng(vPart)
            endV = startV
            isRng = False
        End If

        ' Validate start verse
        Dim okStart As Boolean
        okStart = aeBibleCitationClass.ValidateSBLReference( _
            bID, bCanon, ch, CStr(startV), ModeSBL, True)
        If Not okStart Then
            Debug.Print "FAIL: " & canon & " (start verse failed)"
            failCount = failCount + 1
        ElseIf isRng Then
            Dim okEnd As Boolean
            okEnd = aeBibleCitationClass.ValidateSBLReference( _
                bID, bCanon, ch, CStr(endV), ModeSBL, True)
            If Not okEnd Then
                Debug.Print "FAIL: " & canon & " (end verse " & endV & " failed)"
                failCount = failCount + 1
            Else
                Debug.Print "PASS: " & canon
                passCount = passCount + 1
            End If
        Else
            Debug.Print "PASS: " & canon
            passCount = passCount + 1
        End If
    Next Item

    Debug.Print "--- " & passCount & " passed, " & failCount & " failed. ---"
    VerifyCitationBlock = failCount
PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure VerifyCitationBlock of Module basTEST_aeBibleCitationBlock"
    Resume PROC_EXIT
End Function

' =============================================================================
' Test_VerifyCitationBlock  (Public)
' Integration test — 35-token study Bible citation block.
' Input is deliberately out of canonical order and contains one malformed
' verse spec (103:-11). Expected: 34 PASS, 1 FAIL; output in canonical order.
' En dashes use ChrW(8211); NormalizeRawInput converts them to ASCII hyphen.
' Expected Result after sort and fix:
'   "Gen 1:27; Num 14:18; Deut 32:6; Josh 1:9; 1 Sam 2:2; " & _
'   "1 Chr 29:10-13; " & _
'   "Ps 19:1-2; 23:1; 28:7; 68:5; " & _
'   "103:8-11; 111:3-5; " & _
'   "145:8-9,17; Isa 40:28; 63:16; 64:8; " & _
'   "Jer 33:11; Nah 1:3; Mal 2:10-15; " & _
'   "Matt 6:9; 7:11; 23:9; John 3:16; 4:24; " & _
'   "Rom 1:20; 8:15; 1 Cor 8:6; 14:33; Gal 3:20; Eph 4:6; " & _
'   "Heb 13:6; 1 Pet 1:17; 2 Pet 3:9; 1 John 4:16"
' =============================================================================
Public Sub Test_VerifyCitationBlock()
    Dim rawBlock As String
    rawBlock = "Gen 1:27; Num 14:18; Deut 32:6; Josh 1:9; 1 Sam 2:2; " & _
               "1 Chr 29:10" & ChrW(8211) & "13; " & _
               "Matt 6:9; 7:11; 23:9; John 3:16; 4:24; " & _
               "Ps 19:1" & ChrW(8211) & "2; 23:1; 28:7; 68:5; " & _
               "103:-11; 111:3" & ChrW(8211) & "5; " & _
               "145:8" & ChrW(8211) & "9,17; Isa 40:28; 63:16; 64:8; " & _
               "1 John 4:16; Jer 33:11; Nah 1:3; Mal 2:10" & ChrW(8211) & "15; " & _
               "Rom 1:20; 8:15; 1 Cor 8:6; 14:33; Gal 3:20; Eph 4:6; " & _
               "Heb 13:6; 1 Pet 1:17; 2 Pet 3:9"
    Debug.Print "=== Test_VerifyCitationBlock (35 tokens: 34 pass, 1 fail expected) ==="
    VerifyCitationBlock rawBlock
End Sub

' =============================================================================
' Test_RenderEnDash  (Public)
' Stage 17 option: verifies that RenderEnDash replaces ASCII hyphen with
' en-dash in range strings and leaves non-range strings unchanged.
' =============================================================================
Public Sub Test_RenderEnDash()
    On Error GoTo PROC_ERR
    Dim ownAssert As Boolean
    ownAssert = (aeAssert Is Nothing)
    If ownAssert Then
        Set aeAssert = New aeAssertClass
        aeAssert.Initialize
    End If

    Debug.Print "=== Test_RenderEnDash ==="

    Dim rendered As String

    ' Range entry: hyphen becomes en-dash
    rendered = aeBibleCitationClass.RenderEnDash("Psalms 103:8-11")
    aeAssert.AssertEqual "Psalms 103:8" & ChrW(8211) & "11", rendered, _
        "RenderEnDash: range gets en-dash"

    ' Multi-word book name with range
    rendered = aeBibleCitationClass.RenderEnDash("1 Chronicles 29:10-13")
    aeAssert.AssertEqual "1 Chronicles 29:10" & ChrW(8211) & "13", rendered, _
        "RenderEnDash: multi-word book range"

    ' Non-range entry: unchanged
    rendered = aeBibleCitationClass.RenderEnDash("Psalms 23:1")
    aeAssert.AssertEqual "Psalms 23:1", rendered, _
        "RenderEnDash: non-range unchanged"

    If ownAssert Then
        aeAssert.Terminate
        Set aeAssert = Nothing
    End If
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_RenderEnDash of Module basTEST_aeBibleCitationBlock"
    Resume PROC_EXIT
End Sub

' =============================================================================
' Test_SortCitationBlock  (Public)
' Stage 13b: verifies that SortCitationBlock returns a Collection in canonical
' book order (BookID 1-66), then chapter, then start verse.
' =============================================================================
Public Sub Test_SortCitationBlock()
    On Error GoTo PROC_ERR
    Dim ownAssert As Boolean
    ownAssert = (aeAssert Is Nothing)
    If ownAssert Then
        Set aeAssert = New aeAssertClass
        aeAssert.Initialize
    End If

    Debug.Print "=== Test_SortCitationBlock ==="

    Dim sorted As Collection

    ' Cross-book: out-of-order input (John, Genesis, Psalms)
    Set sorted = aeBibleCitationClass.SortCitationBlock( _
        aeBibleCitationClass.ParseCitationBlock("John 3:16; Gen 1:1; Ps 23:1"))
    aeAssert.AssertEqual 3, sorted.Count, "Sort: Count preserved"
    aeAssert.AssertEqual "Genesis 1:1", sorted(1), "Sort: Gen first"
    aeAssert.AssertEqual "Psalms 23:1", sorted(2), "Sort: Ps second"
    aeAssert.AssertEqual "John 3:16", sorted(3), "Sort: John third"

    ' Same-book: chapter order within Psalms
    Set sorted = aeBibleCitationClass.SortCitationBlock( _
        aeBibleCitationClass.ParseCitationBlock("Ps 103:8; Ps 19:1; Ps 68:5"))
    aeAssert.AssertEqual 3, sorted.Count, "Sort: same-book Count"
    aeAssert.AssertEqual "Psalms 19:1", sorted(1), "Sort: Ps 19 before 68"
    aeAssert.AssertEqual "Psalms 68:5", sorted(2), "Sort: Ps 68 before 103"
    aeAssert.AssertEqual "Psalms 103:8", sorted(3), "Sort: Ps 103 last"

    If ownAssert Then
        aeAssert.Terminate
        Set aeAssert = Nothing
    End If
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_SortCitationBlock of Module basTEST_aeBibleCitationBlock"
    Resume PROC_EXIT
End Sub

' =============================================================================
' VerifyCitationBlockReport  (Public)
' Same validation logic as VerifyCitationBlock but returns a formatted String
' report suitable for display in a MsgBox. passCount and failCount are returned
' ByRef so the caller can branch on failCount without parsing the string.
' =============================================================================
Public Function VerifyCitationBlockReport(rawBlock As String, _
                                          ByRef passCount As Long, _
                                          ByRef failCount As Long) As String
    On Error GoTo PROC_ERR
    Dim Items As Collection
    Dim report As String

    Set Items = aeBibleCitationClass.SortCitationBlock( _
        aeBibleCitationClass.ParseCitationBlock(rawBlock))

    Dim Item As Variant
    For Each Item In Items
        Dim canon As String
        canon = CStr(Item)

        Dim lastSp As Long
        Dim k As Long
        For k = Len(canon) To 1 Step -1
            If Mid$(canon, k, 1) = " " Then lastSp = k: Exit For
        Next k

        Dim bookName As String
        bookName = Left$(canon, lastSp - 1)
        Dim numPart As String
        numPart = Mid$(canon, lastSp + 1)

        Dim bID As Long
        Dim bCanon As String
        bCanon = aeBibleCitationClass.ResolveAlias(bookName, bID)

        Dim cpPos As Long
        cpPos = InStr(numPart, ":")
        Dim ch As Long
        Dim vPart As String

        If cpPos > 0 Then
            ch = CLng(Left$(numPart, cpPos - 1))
            vPart = Mid$(numPart, cpPos + 1)
        Else
            ' Whole-chapter reference (no colon): numPart is the chapter
            ch = CLng(numPart)
            vPart = ""
        End If

        If vPart = "" Then
            ' Whole-chapter: validate chapter only
            Dim okChap As Boolean
            okChap = aeBibleCitationClass.ValidateSBLReference( _
                bID, bCanon, ch, "", ModeSBL, True)
            If Not okChap Then
                report = report & "FAIL: " & canon & " (chapter failed)" & vbCrLf
                failCount = failCount + 1
            Else
                report = report & "PASS: " & canon & vbCrLf
                passCount = passCount + 1
            End If
        Else

        Dim dpPos As Long
        dpPos = InStr(vPart, "-")
        Dim startV As Long
        Dim endV As Long
        Dim isRng As Boolean
        If dpPos > 0 Then
            startV = CLng(Left$(vPart, dpPos - 1))
            endV = CLng(Mid$(vPart, dpPos + 1))
            isRng = True
        Else
            startV = CLng(vPart)
            endV = startV
            isRng = False
        End If

        Dim okStart As Boolean
        okStart = aeBibleCitationClass.ValidateSBLReference( _
            bID, bCanon, ch, CStr(startV), ModeSBL, True)
        If Not okStart Then
            report = report & "FAIL: " & canon & " (start verse failed)" & vbCrLf
            failCount = failCount + 1
        ElseIf isRng Then
            Dim okEnd As Boolean
            okEnd = aeBibleCitationClass.ValidateSBLReference( _
                bID, bCanon, ch, CStr(endV), ModeSBL, True)
            If Not okEnd Then
                report = report & "FAIL: " & canon & " (end verse " & endV & " failed)" & vbCrLf
                failCount = failCount + 1
            Else
                report = report & "PASS: " & canon & vbCrLf
                passCount = passCount + 1
            End If
        Else
            report = report & "PASS: " & canon & vbCrLf
            passCount = passCount + 1
        End If
        End If  ' vPart = ""
    Next Item

    report = report & "--- " & passCount & " passed, " & failCount & " failed. ---"
    VerifyCitationBlockReport = report
PROC_EXIT:
    Exit Function
PROC_ERR:
    ' Parse errors (non-ASCII token, block too long) signal bad input — let the
    ' caller display a user-friendly message rather than a raw error box.
    If Err.Number = vbObjectError + 1002 Or Err.Number = vbObjectError + 1003 Then
        failCount = -1
        Resume PROC_EXIT
    End If
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure VerifyCitationBlockReport of Module basTEST_aeBibleCitationBlock"
    Resume PROC_EXIT
End Function

' =============================================================================
' basRepairCitationBlock
' Interactive citation block repair procedure for Study Bible documents.
' Place cursor anywhere in a paragraph containing a citation block, then run
' RepairCitationBlockInParagraph.
' See rvw/Code_review - 2026-04-04a.md for design background.
' =============================================================================

Public Sub RepairCitationBlockInParagraph()
    On Error GoTo PROC_ERR

    ' --- Task 1: Capture working range BEFORE confirm dialog (preserves selection) ---
    Dim workRng As Object
    If Selection.Type = wdSelectionNormal Then
        Set workRng = Selection.Range
        ' Word often extends a drag selection to include the trailing paragraph mark.
        ' Exclude it so that replacing workRng.Text does not delete the mark and
        ' merge this paragraph with the next.
        If Right$(workRng.Text, 1) = Chr(13) Then
            workRng.End = workRng.End - 1
        End If
    Else
        Set workRng = Selection.Paragraphs(1).Range
        workRng.End = workRng.End - 1   ' exclude paragraph mark
    End If
    'Debug.Print "workRng = " & workRng.Text

    ' --- Task 2: Confirm intent (default No) ---
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Repair citation block in the current paragraph?", _
                    vbYesNo + vbDefaultButton2 + vbQuestion, _
                    "Repair Citation Block")
    If answer <> vbYes Then Exit Sub

    ' --- Task 3: Validate ---
    Dim rawBlock As String
    rawBlock = workRng.Text
    If Right$(rawBlock, 1) = Chr(13) Then
        rawBlock = Left$(rawBlock, Len(rawBlock) - 1)
    End If

    Dim report As String
    Dim passCount As Long
    Dim failCount As Long
    passCount = 0
    failCount = 0
    report = VerifyCitationBlockReport(rawBlock, passCount, failCount)
    'Debug.Print "report = " & report

    If failCount = -1 Then
        MsgBox "The selected text contains non-citation content." & vbCrLf & vbCrLf & _
               "Select only the citations to validate, or insert the cursor in a " & _
               "paragraph that contains only citations.", _
               vbOKOnly + vbExclamation, "Invalid Selection"
        Exit Sub
    End If

    If failCount > 0 Then
        MsgBox report & vbCrLf & vbCrLf & _
               "Fix the errors above in the paragraph, then run the command again.", _
               vbOKOnly + vbExclamation, _
               "Citation Block Errors (" & failCount & " failed)"
        Exit Sub
    End If

    ' --- Task 4: Sort into canonical order ---
    Dim Items As Collection
    Set Items = aeBibleCitationClass.SortCitationBlock( _
        aeBibleCitationClass.ParseCitationBlock(rawBlock))

    ' --- Task 5: Render SBL short form with en-dash; suppress repeated book/chapter ---
    Dim Result As String
    Dim Item As Variant
    Dim prevBook As String
    Dim prevChap As String
    prevBook = ""
    prevChap = ""
    For Each Item In Items
        Dim canonStr As String
        canonStr = CStr(Item)

        ' Split canonical string at last space: left = book name, right = ch:verse[-end]
        Dim lastSp As Long
        Dim jj As Long
        lastSp = 0
        For jj = Len(canonStr) To 1 Step -1
            If Mid$(canonStr, jj, 1) = " " Then lastSp = jj: Exit For
        Next jj

        Dim canonBook As String
        Dim numPart As String
        canonBook = Left$(canonStr, lastSp - 1)
        numPart = Mid$(canonStr, lastSp + 1)

        ' Extract chapter from numPart (before ":" if present)
        Dim colonPos As Long
        Dim thisChap As String
        Dim versePart As String
        colonPos = InStr(numPart, ":")
        If colonPos > 0 Then
            thisChap = Left$(numPart, colonPos - 1)
            versePart = Mid$(numPart, colonPos + 1)
        Else
            ' Whole-chapter reference: numPart is the chapter number, no verse
            thisChap = numPart
            versePart = ""
        End If

        Dim seg As String
        Dim sep As String
        If canonBook = prevBook And thisChap = prevChap And thisChap <> "" Then
            ' Same book, same chapter — comma-separated verse only
            seg = aeBibleCitationClass.RenderEnDash(versePart)
            sep = ", "
        ElseIf canonBook = prevBook Then
            ' Same book, different chapter — semicolon, ch:verse
            seg = aeBibleCitationClass.RenderEnDash(numPart)
            sep = "; "
        Else
            ' New book — full SBL short form
            seg = aeBibleCitationClass.RenderEnDash( _
                aeBibleCitationClass.ToSBLShortForm(canonStr))
            sep = "; "
            prevBook = canonBook
        End If
        prevChap = thisChap

        If Len(Result) > 0 Then Result = Result & sep
        Result = Result & seg
    Next Item

    Dim dataObj As Object
    Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.SetText Result
    dataObj.PutInClipboard
    Set dataObj = Nothing

    ' --- Task 6: Prompt to replace selection ---
    Dim replaceIt As VbMsgBoxResult
    replaceIt = MsgBox("Corrected block copied to clipboard:" & vbCrLf & vbCrLf & _
                       Result & vbCrLf & vbCrLf & _
                       "Replace the original paragraph or selection with the corrected version?", _
                       vbYesNo + vbDefaultButton1 + vbQuestion, _
                       "Replace Citation Block")
    If replaceIt = vbYes Then
        workRng.Text = Result
        workRng.Select
        Selection.Collapse wdCollapseEnd
    End If

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RepairCitationBlockInParagraph of Module basRepairCitationBlock"
    Resume PROC_EXIT
End Sub

Public Sub Test_ParseCitationBlock_EdgeCases()
    On Error GoTo PROC_ERR
    Dim Items As Collection

    ' Empty string - should return empty Collection
    Set Items = aeBibleCitationClass.ParseCitationBlock("")
    aeAssert.AssertEqual 0, Items.Count, "EdgeCase: empty string yields 0 Items"

    ' All-whitespace - should return empty Collection
    Set Items = aeBibleCitationClass.ParseCitationBlock("   ")
    aeAssert.AssertEqual 0, Items.Count, "EdgeCase: whitespace-only yields 0 Items"

    ' Single reference
    Set Items = aeBibleCitationClass.ParseCitationBlock("John 3:16")
    aeAssert.AssertEqual 1, Items.Count, "EdgeCase: single ref item Count"
    aeAssert.AssertEqual "John 3:16", CStr(Items(1)), "EdgeCase: single ref value"

    ' Trailing semicolon - should not produce a spurious extra item
    Set Items = aeBibleCitationClass.ParseCitationBlock("John 3:16;")
    aeAssert.AssertEqual 1, Items.Count, "EdgeCase: trailing semicolon item Count"
    aeAssert.AssertEqual "John 3:16", CStr(Items(1)), "EdgeCase: trailing semicolon value"

    Debug.Print "Test_ParseCitationBlock_EdgeCases: all assertions passed"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_ParseCitationBlock_EdgeCases of Module basTEST_aeBibleCitationBlock"
    Resume PROC_EXIT
End Sub

Public Sub Test_SingleChapterBooks()
    On Error GoTo PROC_ERR
    Dim Items As Collection

    ' Obadiah — BookID 31
    Set Items = aeBibleCitationClass.ParseCitationBlock("Obad 3")
    aeAssert.AssertEqual 1, Items.Count, "SingleChapter: Obad item Count"
    aeAssert.AssertEqual "Obadiah 1:3", CStr(Items(1)), "SingleChapter: Obad 3 -> Obadiah 1:3"

    ' Philemon — BookID 57
    Set Items = aeBibleCitationClass.ParseCitationBlock("Phlm 10")
    aeAssert.AssertEqual 1, Items.Count, "SingleChapter: Phlm item Count"
    aeAssert.AssertEqual "Philemon 1:10", CStr(Items(1)), "SingleChapter: Phlm 10 -> Philemon 1:10"

    ' 2 John — BookID 63
    Set Items = aeBibleCitationClass.ParseCitationBlock("2 John 5")
    aeAssert.AssertEqual 1, Items.Count, "SingleChapter: 2 John item Count"
    aeAssert.AssertEqual "2 John 1:5", CStr(Items(1)), "SingleChapter: 2 John 5 -> 2 John 1:5"

    Debug.Print "Test_SingleChapterBooks: all assertions passed"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_SingleChapterBooks of Module basTEST_aeBibleCitationBlock"
    Resume PROC_EXIT
End Sub

Public Sub Test_NormalizeRawInput_Chr11()
    On Error GoTo PROC_ERR
    ' Chr(11) is Word forced line break (Shift+Enter); must be treated as whitespace
    Dim raw As String
    raw = "1 Chr 29:10-13;" & Chr(11) & "Ps 19:1-2"
    Dim Items As Collection
    Set Items = aeBibleCitationClass.ParseCitationBlock(raw)
    aeAssert.AssertEqual 2, Items.Count, "Chr(11) normalization: item Count"
    aeAssert.AssertEqual "1 Chronicles 29:10-13", CStr(Items(1)), "Chr(11) normalization: first item"
    aeAssert.AssertEqual "Psalms 19:1-2", CStr(Items(2)), "Chr(11) normalization: Ps must not inherit 1 Chr book"
    Debug.Print "Test_NormalizeRawInput_Chr11: all assertions passed"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_NormalizeRawInput_Chr11 of Module basTEST_aeBibleCitationBlock"
    Resume PROC_EXIT
End Sub

Public Sub Test_ctxChapter_Reset()
    On Error GoTo PROC_ERR
    ' "2 Pet 2:4; Jude 6" — Jude must not inherit chapter 2 from 2 Peter
    Dim Items As Collection
    Set Items = aeBibleCitationClass.ParseCitationBlock("2 Pet 2:4; Jude 6")
    aeAssert.AssertEqual 2, Items.Count, "ctxChapter reset: item Count"
    aeAssert.AssertEqual "2 Peter 2:4", CStr(Items(1)), "ctxChapter reset: 2 Peter item"
    aeAssert.AssertEqual "Jude 1:6", CStr(Items(2)), "ctxChapter reset: Jude must not inherit chapter 2"
    Debug.Print "Test_ctxChapter_Reset: all assertions passed"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_ctxChapter_Reset of Module basTEST_aeBibleCitationBlock"
    Resume PROC_EXIT
End Sub

Public Sub Test_VerifyCitationBlockReport()
    On Error GoTo PROC_ERR
    Dim report As String
    Dim passCount As Long, failCount As Long

    ' Known-good block: 2 valid references
    passCount = 0: failCount = 0
    report = VerifyCitationBlockReport("John 3:16; Rev 22:1", passCount, failCount)
    aeAssert.AssertEqual 2, passCount, "VerifyCitationBlockReport: passCount good block"
    aeAssert.AssertEqual 0, failCount, "VerifyCitationBlockReport: failCount good block"
    aeAssert.AssertTrue Len(report) > 0, "VerifyCitationBlockReport: report non-empty good block"

    ' Block with one invalid reference (chapter out of range)
    passCount = 0: failCount = 0
    report = VerifyCitationBlockReport("John 3:16; Rev 99:1", passCount, failCount)
    aeAssert.AssertEqual 1, passCount, "VerifyCitationBlockReport: passCount bad block"
    aeAssert.AssertEqual 1, failCount, "VerifyCitationBlockReport: failCount bad block"
    aeAssert.AssertTrue Len(report) > 0, "VerifyCitationBlockReport: report non-empty bad block"

    Debug.Print "Test_VerifyCitationBlockReport: all assertions passed"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_VerifyCitationBlockReport of Module basTEST_aeBibleCitationBlock"
    Resume PROC_EXIT
End Sub

Public Sub Test_ToSBLShortForm()
    On Error GoTo PROC_ERR
    ' Multi-word book
    aeAssert.AssertEqual "1 Chr 29:10-13", _
        aeBibleCitationClass.ToSBLShortForm("1 Chronicles 29:10-13"), _
        "ToSBLShortForm: 1 Chronicles"
    ' Single-chapter book — chapter number omitted
    aeAssert.AssertEqual "Jude 6", _
        aeBibleCitationClass.ToSBLShortForm("Jude 1:6"), _
        "ToSBLShortForm: Jude single-chapter"
    ' Single-chapter range
    aeAssert.AssertEqual "Obad 3-5", _
        aeBibleCitationClass.ToSBLShortForm("Obadiah 1:3-5"), _
        "ToSBLShortForm: Obadiah range"
    ' Standard multi-chapter
    aeAssert.AssertEqual "Ps 23:1", _
        aeBibleCitationClass.ToSBLShortForm("Psalms 23:1"), _
        "ToSBLShortForm: Psalms"
    Debug.Print "Test_ToSBLShortForm: all assertions passed"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_ToSBLShortForm of Module basTEST_aeBibleCitationBlock"
    Resume PROC_EXIT
End Sub


Public Sub Test_WholeChapterReference()
    On Error GoTo PROC_ERR
    Dim Items As Collection

    ' Single whole-chapter reference
    Set Items = aeBibleCitationClass.ParseCitationBlock("Ezek 16")
    aeAssert.AssertEqual 1, Items.Count, "WholeChapter: Ezek 16 item Count"
    aeAssert.AssertEqual "Ezekiel 16", CStr(Items(1)), "WholeChapter: Ezek 16 canonical"

    ' Whole chapter mixed with verse references
    Set Items = aeBibleCitationClass.ParseCitationBlock("Gen 6:6; Ezek 16; Luke 15:4-32")
    aeAssert.AssertEqual 3, Items.Count, "WholeChapter: mixed block item Count"
    aeAssert.AssertEqual "Genesis 6:6", CStr(Items(1)), "WholeChapter: Gen 6:6"
    aeAssert.AssertEqual "Ezekiel 16", CStr(Items(2)), "WholeChapter: Ezek 16 in mixed block"
    aeAssert.AssertEqual "Luke 15:4-32", CStr(Items(3)), "WholeChapter: Luke 15:4-32"

    ' Verify report passes for whole-chapter block
    Dim report As String
    Dim passCount As Long, failCount As Long
    passCount = 0: failCount = 0
    report = VerifyCitationBlockReport("Gen 6:6; Ezek 16; Luke 15:4-32", passCount, failCount)
    aeAssert.AssertEqual 3, passCount, "WholeChapter: verify passCount"
    aeAssert.AssertEqual 0, failCount, "WholeChapter: verify failCount"

    ' Sort key: whole-chapter sorts before verse refs in same chapter
    Set Items = aeBibleCitationClass.SortCitationBlock( _
        aeBibleCitationClass.ParseCitationBlock("Ezek 16:5; Ezek 16"))
    aeAssert.AssertEqual "Ezekiel 16", CStr(Items(1)), "WholeChapter: sorts before verse in same chapter"
    aeAssert.AssertEqual "Ezekiel 16:5", CStr(Items(2)), "WholeChapter: verse after whole-chapter"

    ' ToSBLShortForm for whole-chapter
    aeAssert.AssertEqual "Ezek 16", _
        aeBibleCitationClass.ToSBLShortForm("Ezekiel 16"), _
        "WholeChapter: ToSBLShortForm Ezek 16"

    ' Chapter-switch within same book: bare number after semicolon is a chapter, not a verse
    ' "Isa 45:17; 60" — 60 is chapter 60 of Isaiah, not verse 60 of chapter 45
    Set Items = aeBibleCitationClass.ParseCitationBlock("Isa 45:17; 60")
    aeAssert.AssertEqual 2, Items.Count, "WholeChapter: Isa 45:17; 60 item Count"
    aeAssert.AssertEqual "Isaiah 45:17", CStr(Items(1)), "WholeChapter: Isa 45:17"
    aeAssert.AssertEqual "Isaiah 60", CStr(Items(2)), "WholeChapter: Isa 60 via same-book chapter switch"

    ' Verify report passes for same-book chapter switch
    Dim passCount2 As Long, failCount2 As Long, report2 As String
    passCount2 = 0: failCount2 = 0
    report2 = VerifyCitationBlockReport("Isa 45:17; 60; 62:4", passCount2, failCount2)
    aeAssert.AssertEqual 3, passCount2, "WholeChapter: Isa mixed passCount"
    aeAssert.AssertEqual 0, failCount2, "WholeChapter: Isa mixed failCount"

    Debug.Print "Test_WholeChapterReference: all assertions passed"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_WholeChapterReference of Module basTEST_aeBibleCitationBlock"
    Resume PROC_EXIT
End Sub

Public Sub Run_Extra_Tests()
    On Error GoTo PROC_ERR

    Set aeAssert = New aeAssertClass
    aeAssert.Initialize

    Test_ParseCitationBlock_EdgeCases
    Test_SingleChapterBooks
    Test_NormalizeRawInput_Chr11
    Test_ctxChapter_Reset
    Test_VerifyCitationBlockReport
    Test_ToSBLShortForm
    Test_WholeChapterReference

    aeAssert.Terminate
    Set aeAssert = Nothing

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Run_Extra_Tests of Module basTEST_aeBibleCitationBlock"
    Resume PROC_EXIT
End Sub
