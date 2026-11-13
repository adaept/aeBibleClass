Attribute VB_Name = "ProbeBuiltInHeaderFooter"
Option Explicit

' ProbeBuiltInHeaderFooter - FU-47 / Test 84, 2026-11-13. READ-ONLY (changes nothing, saves nothing).
' Test 84 returned 1 (header viol=0, footer viol=1) in the full run of 2026-11-13 10:11.
' This lists every paragraph of every header and footer story whose style is not
' TheHeaders (headers) or TheFooters (footers): section, story type, whether the story
' exists / is linked, the style and the text. It walks the stories the same way Test 84 does
' (StoryRanges + NextStoryRange) and then, for comparison, the way Tests 36 and 50 do
' (Sections(i).Headers/Footers(1..3)), so the two views can be compared.
' Run in the open Blank Bible Copy.docm: ProbeBuiltInHeaderFooter.Run  (output: Immediate window, also
' C:\adaept\aeTest\probe-builtin-hf.txt). Do not save the document afterwards.

Public Sub Run()
    Dim sOut As String
    Dim rng As Object
    Dim oPara As Object
    Dim nStoryParas As Long, nBad As Long
    Dim expected As String
    Dim idx As Long
    Dim sStep As String

    On Error GoTo Fail
    sStep = "start"
    sOut = "ProbeBuiltInHeaderFooter " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    sOut = sOut & "Document: " & ActiveDocument.FullName & vbCrLf
    sOut = sOut & "Sections: " & ActiveDocument.Sections.Count & vbCrLf & vbCrLf
    sOut = sOut & "--- View 1: StoryRanges walk (what Test 84 reads) ---" & vbCrLf

    For Each rng In ActiveDocument.StoryRanges
        If rng.StoryType >= 6 And rng.StoryType <= 11 Then
            Dim cur As Object
            Set cur = rng
            Do
                If cur.StoryType = 6 Or cur.StoryType = 7 Or cur.StoryType = 10 Then
                    expected = "TheHeaders"
                Else
                    expected = "TheFooters"
                End If
                sStep = "view 1, story type " & cur.StoryType
                For Each oPara In cur.Paragraphs
                    nStoryParas = nStoryParas + 1
                    sStep = "view 1, story type " & cur.StoryType & ", reading style of paragraph " & nStoryParas
                    If oPara.Style.NameLocal <> expected Then
                        nBad = nBad + 1
                        sStep = "view 1, BAD line, story type " & cur.StoryType & " (cur.Sections(1).Index)"
                        sOut = sOut & "BAD  section " & cur.Sections(1).index & " | " & storyName(cur.StoryType) & _
                            " | style=" & oPara.Style.NameLocal & " (expected " & expected & ") | text=""" & _
                            Left$(Replace(Replace(oPara.Range.Text, vbCr, "<CR>"), vbTab, "<TAB>"), 60) & """" & vbCrLf
                    End If
                Next oPara
                sStep = "view 1, NextStoryRange after story type " & cur.StoryType
                Set cur = cur.NextStoryRange
            Loop Until cur Is Nothing
        End If
    Next rng
    sOut = sOut & "View 1 total: " & nStoryParas & " paragraphs, " & nBad & " not on the approved style" & vbCrLf & vbCrLf

    sOut = sOut & "--- View 2: Sections(i).Headers/Footers(1..3) (what Tests 36 and 50 read) ---" & vbCrLf
    Dim s As Long, k As Long, hf As Object, bad2 As Long, kind As String
    For s = 1 To ActiveDocument.Sections.Count
        For k = 1 To 3
            Dim h As Long
            For h = 1 To 2
                sStep = "view 2, section " & s & ", k=" & k & ", h=" & h
                If h = 1 Then Set hf = ActiveDocument.Sections(s).Headers(k) Else Set hf = ActiveDocument.Sections(s).Footers(k)
                kind = IIf(h = 1, "Header", "Footer") & "(" & Choose(k, "Primary", "FirstPage", "EvenPages") & ")"
                If hf.Exists Then
                    For Each oPara In hf.Range.Paragraphs
                        If oPara.Style.NameLocal <> IIf(h = 1, "TheHeaders", "TheFooters") Then
                            bad2 = bad2 + 1
                            sOut = sOut & "BAD  section " & s & " | " & kind & " | LinkToPrevious=" & hf.LinkToPrevious & _
                                " | style=" & oPara.Style.NameLocal & " | text=""" & _
                                Left$(Replace(Replace(oPara.Range.Text, vbCr, "<CR>"), vbTab, "<TAB>"), 60) & """" & vbCrLf
                        End If
                    Next oPara
                End If
            Next h
        Next k
    Next s
    sOut = sOut & "View 2 total: " & bad2 & " paragraph(s) not on the approved style (existing stories only)" & vbCrLf & vbCrLf

    sOut = sOut & "--- Section 1 to 4 detail: every story, existing or not ---" & vbCrLf
    For s = 1 To 4
        If s <= ActiveDocument.Sections.Count Then
            For k = 1 To 3
                For h = 1 To 2
                    sStep = "detail, section " & s & ", k=" & k & ", h=" & h
                    If h = 1 Then Set hf = ActiveDocument.Sections(s).Headers(k) Else Set hf = ActiveDocument.Sections(s).Footers(k)
                    kind = IIf(h = 1, "Header", "Footer") & "(" & Choose(k, "Primary", "FirstPage", "EvenPages") & ")"
                    sOut = sOut & "section " & s & " | " & kind & " | Exists=" & hf.Exists & " | Link=" & hf.LinkToPrevious & _
                        " | style=" & hf.Range.Paragraphs(1).Style.NameLocal & " | paras=" & hf.Range.Paragraphs.Count & vbCrLf
                Next h
            Next k
        End If
    Next s
    sOut = sOut & vbCrLf & "Reading it: a BAD line in View 1 only = an orphan story (Exists False, FU-47 item 1) that Test 84 sees and Tests 36/50 do not; " & _
        "in both views = a real paragraph on screen (extend RepairFrontHeaderFooterStyles)." & vbCrLf

Done:
    Debug.Print sOut
    Dim ff As Integer
    ff = FreeFile
    Open "C:\adaept\aeTest\probe-builtin-hf.txt" For Output As #ff
    Print #ff, sOut
    Close #ff
    Exit Sub

Fail:
    sOut = sOut & vbCrLf & "*** ERROR " & Err.Number & " (" & Err.Description & ") at step: " & sStep & vbCrLf
    Err.Clear
    On Error GoTo 0
    GoTo Done
End Sub

Private Function storyName(ByVal t As Long) As String
    Select Case t
        Case 6: storyName = "EvenPagesHeader"
        Case 7: storyName = "PrimaryHeader"
        Case 8: storyName = "EvenPagesFooter"
        Case 9: storyName = "PrimaryFooter"
        Case 10: storyName = "FirstPageHeader"
        Case 11: storyName = "FirstPageFooter"
        Case Else: storyName = "Story" & t
    End Select
End Function
