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

    Dim item As Variant
    For Each item In Items
        Dim canon As String
        canon = CStr(item)

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
    Next item

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
' Expected result after sort and fix:
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
    aeAssert.AssertEqual 3, sorted.count, "Sort: count preserved"
    aeAssert.AssertEqual "Genesis 1:1", sorted(1), "Sort: Gen first"
    aeAssert.AssertEqual "Psalms 23:1", sorted(2), "Sort: Ps second"
    aeAssert.AssertEqual "John 3:16", sorted(3), "Sort: John third"

    ' Same-book: chapter order within Psalms
    Set sorted = aeBibleCitationClass.SortCitationBlock( _
        aeBibleCitationClass.ParseCitationBlock("Ps 103:8; Ps 19:1; Ps 68:5"))
    aeAssert.AssertEqual 3, sorted.count, "Sort: same-book count"
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
