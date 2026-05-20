Attribute VB_Name = "basApplyDateRule_2026_05_19"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

' Path to the editorial example file. Status block at the end of this
' file is the authority for which pairs are still pending.
Private Const STATUS_FILE As String = _
    "C:\adaept\aeBibleClass\Date_Example.txt"

'=======================================================================
' ApplyDateRule_2026_05_19
' Purpose : Apply the 2026-05-19 date-formatting rule to the active
'           document. Replaces century-form date descriptions with
'           en-dash year ranges, matching the rewrite already applied
'           to Date_Example.txt.
'
' Interactive: for each pair the user is shown Find/Replace text and
'              prompted Yes / No / Cancel.
'                Yes    - apply, mark pair done.
'                No     - skip this run, leave pending.
'                Cancel - abort the macro at this pair.
'
' Idempotent : status block at the end of Date_Example.txt records
'              completion per pair. Pairs marked done are skipped on
'              subsequent runs. A pair whose Find string is no longer
'              present in the document is auto-marked done.
'
' Scope      : 20 example passages from Date_Example.txt (23 pairs).
'              Book-number ordinals (1st Samuel, etc.) are out of scope.
'
' Usage      : open the target document, then run from the VBE
'              Immediate window:  ApplyDateRule_2026_05_19
'              Re-run Test_NoSuperscriptOrdinals afterwards.
'=======================================================================
Public Sub ApplyDateRule_2026_05_19()

    On Error GoTo PROC_ERR

    Dim ndash As String
    Dim apos As String
    ndash = ChrW(8211)
    apos = ChrW(8217)

    Dim aFind(1 To 23) As String
    Dim aRepl(1 To 23) As String
    Dim aExam(1 To 23) As String
    BuildPairs aFind, aRepl, aExam, ndash, apos

    Dim aStatus(1 To 23) As String
    Dim strFileText As String
    strFileText = ReadAllUTF8(STATUS_FILE)
    ParseStatuses strFileText, aStatus

    Dim lngApplied As Long
    Dim lngSkippedDone As Long
    Dim lngSkippedNo As Long
    Dim lngAutoDone As Long
    Dim lngTotalHits As Long
    Dim blnCancelled As Boolean
    Dim i As Long

    For i = 1 To 23

        If aStatus(i) = "done" Then
            Debug.Print "Pair " & Format(i, "00") & _
                        " (" & aExam(i) & ") - already done, skipped."
            lngSkippedDone = lngSkippedDone + 1
            GoTo ContinueLoop
        End If

        If Not FindExists(aFind(i)) Then
            Debug.Print "Pair " & Format(i, "00") & _
                        " (" & aExam(i) & _
                        ") - Find string not present, auto-marking done."
            aStatus(i) = "done"
            lngAutoDone = lngAutoDone + 1
            GoTo ContinueLoop
        End If

        Dim strMsg As String
        strMsg = "Pair " & Format(i, "00") & " of 23  (" & aExam(i) & ")" & _
                 vbCrLf & vbCrLf & _
                 "FIND:" & vbCrLf & _
                 Truncate(aFind(i), 400) & vbCrLf & vbCrLf & _
                 "REPLACE:" & vbCrLf & _
                 Truncate(aRepl(i), 400) & vbCrLf & vbCrLf & _
                 "Apply this replacement?" & vbCrLf & _
                 "  Yes    - apply and mark done" & vbCrLf & _
                 "  No     - skip for now (stay pending)" & vbCrLf & _
                 "  Cancel - abort the macro"

        Dim lngAnswer As Long
        lngAnswer = MsgBox(strMsg, _
                           vbYesNoCancel + vbQuestion + vbDefaultButton1, _
                           "ApplyDateRule 2026-05-19")

        Select Case lngAnswer
            Case vbYes
                Dim lngHits As Long
                lngHits = DoReplace(aFind(i), aRepl(i))
                Debug.Print "Pair " & Format(i, "00") & _
                            " applied, hits = " & lngHits
                lngTotalHits = lngTotalHits + lngHits
                lngApplied = lngApplied + 1
                aStatus(i) = "done"
            Case vbNo
                Debug.Print "Pair " & Format(i, "00") & _
                            " skipped (No), remains pending."
                lngSkippedNo = lngSkippedNo + 1
            Case vbCancel
                Debug.Print "Pair " & Format(i, "00") & _
                            " - macro cancelled by user."
                blnCancelled = True
                Exit For
        End Select

ContinueLoop:
    Next i

    ' Persist updated statuses back to the file.
    Dim strNewFile As String
    strNewFile = RewriteStatuses(strFileText, aStatus)
    WriteAllUTF8 STATUS_FILE, strNewFile

    Debug.Print "-----------------------------------------------"
    Debug.Print "ApplyDateRule_2026_05_19 summary:"
    Debug.Print "  applied (Yes)            = " & lngApplied
    Debug.Print "  total replacements       = " & lngTotalHits
    Debug.Print "  auto-marked done         = " & lngAutoDone
    Debug.Print "  skipped (already done)   = " & lngSkippedDone
    Debug.Print "  skipped (No this run)    = " & lngSkippedNo
    Debug.Print "  cancelled                = " & CStr(blnCancelled)
    Debug.Print "-----------------------------------------------"
    Debug.Print "Status block in Date_Example.txt has been updated."
    Debug.Print "Next: re-run Test_NoSuperscriptOrdinals."

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & _
           " (" & Err.Description & _
           ") in procedure ApplyDateRule_2026_05_19 of Module " & _
           "basApplyDateRule_2026_05_19"
    Resume PROC_EXIT
End Sub

'-----------------------------------------------------------------------
' Populate the three pair arrays. Kept in its own routine so the main
' routine reads top-to-bottom.
'-----------------------------------------------------------------------
Private Sub BuildPairs(ByRef aFind() As String, _
                       ByRef aRepl() As String, _
                       ByRef aExam() As String, _
                       ByVal ndash As String, _
                       ByVal apos As String)

    aExam(1) = "example 2"
    aFind(1) = "the middle of the 5th century BC"
    aRepl(1) = "the middle of 500" & ndash & "400 BC"

    aExam(2) = "example 2"
    aFind(2) = "in the 6th century BC, during the Babylonian exile"
    aRepl(2) = "in 600" & ndash & "500 BC, during the Babylonian exile"

    aExam(3) = "example 2"
    aFind(3) = "and the 5th century post-exilic period"
    aRepl(3) = "and the 500" & ndash & "400 BC post-exilic period"

    aExam(4) = "example 5"
    aFind(4) = "to the southern kingdom of Judah in the 8th century BC, " & _
               "then adapted in King Josiah" & apos & "s era in the 7th " & _
               "century and finally polished to its current form in " & _
               "about the 6th century BC"
    aRepl(4) = "to the southern kingdom of Judah in 800" & ndash & "700 BC, " & _
               "then adapted in King Josiah" & apos & "s era in 700" & ndash & _
               "600 BC and finally polished to its current form in " & _
               "about 600" & ndash & "500 BC"

    aExam(5) = "examples 6 and 7"
    aFind(5) = "in the time of King Hezekiah in the 8th century BC, and an " & _
               "early version is attributed to his grandson, Josiah, at " & _
               "the end of the 7th century BC"
    aRepl(5) = "in the time of King Hezekiah in 800" & ndash & "700 BC, and an " & _
               "early version is attributed to his grandson, Josiah, at " & _
               "the end of 700" & ndash & "600 BC"

    aExam(6) = "example 8"
    aFind(6) = "between the 6th to the 4th centuries AD"
    aRepl(6) = "between 300" & ndash & "600 AD"

    aExam(7) = "example 16"
    aFind(7) = "originated in about the 4th century BC"
    aRepl(7) = "originated in about 400" & ndash & "300 BC"

    aExam(8) = "example 17"
    aFind(8) = "sometime in the 3rd or 4th century BC"
    aRepl(8) = "sometime in 400" & ndash & "200 BC"

    aExam(9) = "example 18"
    aFind(9) = "written sometime in the 6th century BC"
    aRepl(9) = "written sometime in 600" & ndash & "500 BC"

    aExam(10) = "example 21"
    aFind(10) = "during the last half of the 3rd century BC"
    aRepl(10) = "during the last half of 300" & ndash & "200 BC"

    aExam(11) = "example 24"
    aFind(11) = "during the 6th and 5th centuries BC"
    aRepl(11) = "during 600" & ndash & "400 BC"

    aExam(12) = "example 24"
    aFind(12) = "until the 2nd century BC"
    aRepl(12) = "until 200" & ndash & "100 BC"

    aExam(13) = "example 28"
    aFind(13) = "lived during the 8th century BC"
    aRepl(13) = "lived during 800" & ndash & "700 BC"

    aExam(14) = "example 29"
    aFind(14) = "from the 9th century BC to the 5th century BC"
    aRepl(14) = "from 900" & ndash & "400 BC"

    aExam(15) = "example 32"
    aFind(15) = "between the late 5th and early 4th century BC"
    aRepl(15) = "between 450" & ndash & "350 BC"

    aExam(16) = "example 33"
    aFind(16) = "first three chapters were written in the 8th century BC"
    aRepl(16) = "first three chapters were written in 800" & ndash & "700 BC"

    aExam(17) = "example 33"
    aFind(17) = "added in the early 5th century BC"
    aRepl(17) = "added in 500" & ndash & "450 BC"

    aExam(18) = "example 34"
    aFind(18) = "probably written in the 7th century BC"
    aRepl(18) = "probably written in 700" & ndash & "600 BC"

    aExam(19) = "example 35"
    aFind(19) = "lived during the late 7th century BC"
    aRepl(19) = "lived during 650" & ndash & "600 BC"

    aExam(20) = "example 36"
    aFind(20) = "date this book in the 7th century BC"
    aRepl(20) = "date this book in 700" & ndash & "600 BC"

    aExam(21) = "example 38"
    aFind(21) = "written in the 6th century BC, with possible other " & _
                "additions later in the 5th century BC"
    aRepl(21) = "written in 600" & ndash & "500 BC, with possible other " & _
                "additions later in 500" & ndash & "400 BC"

    aExam(22) = "example 39"
    aFind(22) = "written in the mid-5th century BC"
    aRepl(22) = "written in 475" & ndash & "425 BC"

    aExam(23) = "example 59"
    aFind(23) = "perhaps toward the end of the first century or into the " & _
                "early 2nd century AD"
    aRepl(23) = "perhaps 90" & ndash & "120 AD"
End Sub

'-----------------------------------------------------------------------
' Returns True if strFind occurs at least once in ActiveDocument.
'-----------------------------------------------------------------------
Private Function FindExists(ByVal strFind As String) As Boolean
    Dim rng As Word.Range
    Set rng = ActiveDocument.Content
    rng.Collapse wdCollapseStart
    With rng.Find
        .ClearFormatting
        .Text = strFind
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
    End With
    FindExists = rng.Find.Execute
End Function

'-----------------------------------------------------------------------
' Replace every occurrence of strFind with strReplace; clear superscript
' on each new range; return the number of replacements made.
'-----------------------------------------------------------------------
Private Function DoReplace(ByVal strFind As String, _
                           ByVal strReplace As String) As Long
    On Error GoTo PROC_ERR

    Dim rng As Word.Range
    Set rng = ActiveDocument.Content
    rng.Collapse wdCollapseStart

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = strFind
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
    End With

    Dim lngHits As Long
    Do While rng.Find.Execute
        rng.Text = strReplace
        rng.Font.Superscript = False
        rng.Collapse wdCollapseEnd
        lngHits = lngHits + 1
    Loop
    DoReplace = lngHits

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & _
           " (" & Err.Description & _
           ") in procedure DoReplace of Module basApplyDateRule_2026_05_19"
    Resume PROC_EXIT
End Function

'-----------------------------------------------------------------------
' Truncate a long string for display in the prompt.
'-----------------------------------------------------------------------
Private Function Truncate(ByVal s As String, ByVal lngMax As Long) As String
    If Len(s) <= lngMax Then
        Truncate = s
    Else
        Truncate = Left$(s, lngMax) & " ..."
    End If
End Function

'-----------------------------------------------------------------------
' Parse "pair NN ... pending|done" lines out of the file body.
' Lines that don't match leave aStatus unchanged (default "").
'-----------------------------------------------------------------------
Private Sub ParseStatuses(ByVal strFileText As String, _
                          ByRef aStatus() As String)
    Dim aLines() As String
    aLines = Split(strFileText, vbLf)

    Dim i As Long
    Dim strLine As String
    Dim strBody As String
    Dim lngPair As Long
    For i = LBound(aLines) To UBound(aLines)
        strLine = aLines(i)
        ' Tolerate stray CR from CRLF-encoded files.
        If Right$(strLine, 1) = vbCr Then strLine = Left$(strLine, Len(strLine) - 1)

        If Left$(strLine, 5) = "pair " Then
            lngPair = CLng(val(Mid$(strLine, 6)))
            If lngPair >= 1 And lngPair <= 23 Then
                strBody = LCase$(Trim$(strLine))
                If Right$(strBody, 4) = "done" Then
                    aStatus(lngPair) = "done"
                Else
                    aStatus(lngPair) = "pending"
                End If
            End If
        End If
    Next i
End Sub

'-----------------------------------------------------------------------
' Re-emit the file body, swapping the "pending"/"done" token on each
' parsed "pair NN" line to reflect aStatus. Preserves the rest of the
' file untouched (other than line-ending normalisation to LF, which is
' what WriteAllUTF8 then writes back).
'-----------------------------------------------------------------------
Private Function RewriteStatuses(ByVal strFileText As String, _
                                 ByRef aStatus() As String) As String
    Dim aLines() As String
    aLines = Split(strFileText, vbLf)

    Dim i As Long
    Dim strLine As String
    Dim strLineNoCR As String
    Dim lngPair As Long
    Dim strNew As String

    For i = LBound(aLines) To UBound(aLines)
        strLine = aLines(i)
        strLineNoCR = strLine
        If Right$(strLineNoCR, 1) = vbCr Then _
            strLineNoCR = Left$(strLineNoCR, Len(strLineNoCR) - 1)

        If Left$(strLineNoCR, 5) = "pair " Then
            lngPair = CLng(val(Mid$(strLineNoCR, 6)))
            If lngPair >= 1 And lngPair <= 23 Then
                Dim lngPos As Long
                lngPos = InStrRev(strLineNoCR, " ")
                If lngPos > 0 Then
                    aLines(i) = Left$(strLineNoCR, lngPos) & aStatus(lngPair)
                End If
            End If
        End If
    Next i

    RewriteStatuses = Join(aLines, vbLf)
End Function

'-----------------------------------------------------------------------
' Read a UTF-8 file (with or without BOM) as a VBA String. Late-bound
' ADODB.Stream; no project references required.
'-----------------------------------------------------------------------
Private Function ReadAllUTF8(ByVal strPath As String) As String
    Dim oStream As Object
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Type = 2 ' adTypeText
    oStream.Charset = "utf-8"
    oStream.Open
    oStream.LoadFromFile strPath
    ReadAllUTF8 = oStream.ReadText
    oStream.Close
End Function

'-----------------------------------------------------------------------
' Write a VBA String to disk as UTF-8 without BOM. Uses an intermediate
' binary stream to strip the BOM ADODB.Stream prepends by default.
'-----------------------------------------------------------------------
Private Sub WriteAllUTF8(ByVal strPath As String, ByVal strContent As String)
    Dim oText As Object
    Dim oBin As Object

    Set oText = CreateObject("ADODB.Stream")
    oText.Type = 2 ' adTypeText
    oText.Charset = "utf-8"
    oText.Open
    oText.WriteText strContent
    oText.Position = 0
    oText.Type = 1 ' adTypeBinary
    oText.Position = 3 ' skip 3-byte UTF-8 BOM

    Set oBin = CreateObject("ADODB.Stream")
    oBin.Type = 1
    oBin.Open
    oText.CopyTo oBin
    oText.Close

    oBin.saveToFile strPath, 2 ' adSaveCreateOverWrite
    oBin.Close
End Sub
