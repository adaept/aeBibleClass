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

    Set Items = aeBibleCitationClass.ParseCitationBlock(rawBlock)

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
' Positive integration test - full 35-token study Bible citation block.
' All tokens expected to pass. En dashes use ChrW(8211); NormalizeRawInput
' in ParseCitationBlock converts them to ASCII hyphen before parsing.
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
