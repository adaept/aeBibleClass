' docm-text-edit.vbs - exact text edits in a docm, in a hidden Word of its own, then SAVE the document (Pilot 3 docm session, 2026-11-26).
'
'   cscript //nologo docm-text-edit.vbs <docm> <edits.txt> <result.txt> [body|footnotes]
'
' The last argument says WHICH TEXT is edited: body (the default) or footnotes (Word's body search does not reach the footnotes).
' <edits.txt>: one edit per line, FIVE TAB-separated fields: label, left context, target, right context, replacement (ASCII; {U+2018} writes a
' character by its hex code point; lines starting with # are skipped). The find text is left + target + right: it must occur EXACTLY ONCE in that text
' (case sensitive, no wildcards) or the edit is skipped and reported. Only the TARGET is replaced, so the character style of the context is never touched.
' Keep the find text short and free of soft hyphens (the docm has 7,513; Word counts each as one character, so a find text spanning one makes the
' narrowed range one character off and the edit is refused). The edit file is made by rvw/divine-names/gen_docm_edits.py in aeBibleAddin and the saved
' file is checked by rvw/divine-names/verify_docm_edits.py. Macros stay off (AutomationSecurity 3); Track Changes is switched off; the document is saved
' only when every edit succeeded. Close every Word first (the script refuses otherwise). Back up the docm and rehearse on a copy first.
' After each replacement the edited range with its context is read back and compared with the expected text (no global counting, which Word's Find cannot do for common words).
' It replaces apply-edits-ctx.vbs and apply-edits-ctx-footnotes.vbs. The older apply-edits.vbs (three fields, replaces the whole find text) stays.
' Original header of apply-edits.vbs follows.
' apply-edits.vbs - make a list of exact text replacements in a docm, in a hidden Word of its own, and SAVE the document.
'
'   cscript //nologo apply-edits.vbs <docm> <edits.txt> <result.txt>
'
' <edits.txt>: one edit per line, three TAB-separated fields: label, find text, replacement text (ASCII only; lines starting with # are skipped).
' A character outside ASCII is written as {U+2018} (hex code point, 4 digits), for example a curly quote; the script turns it into the character (2026-11-17, E-10).
' For every edit the find text must occur EXACTLY ONCE in the document (whole body, case sensitive, no wildcards); otherwise that edit is skipped
' and reported, nothing is guessed. The replacement goes through Find.Execute Replace:=wdReplaceOne, so the formatting of the found text stays.
' After the replacement the new text must occur exactly ONE time more than before (it may already occur elsewhere) and the old text not at all (or, when the replacement contains the find text, as often as the replacement contains it: an insertion). Track Changes is switched off. Macros stay off (AutomationSecurity 3).
' Rules (see README.md): close every Word first (this script refuses otherwise); make a backup of the docm before you point this at the real file; run
' the SAME edits on a copy first and compare (the verification reads the saved XML, for example tools/verify-docm-edits.mjs). Written for FU-50 (2026-11-14).
Option Explicit
Dim gWord, gDoc, gOut, gFso, gErrText, gStory

' The text to edit: the body, or the footnotes story (wdFootnotesStory = 2).
Function StoryRange()
  If gStory = "footnotes" Then
    Set StoryRange = gDoc.StoryRanges(2)
  Else
    Set StoryRange = gDoc.Content
  End If
End Function

Sub RefuseIfWordRunning()
  Dim svc, procs
  Set svc = GetObject("winmgmts:\\.\root\cimv2")
  Set procs = svc.ExecQuery("select ProcessId from Win32_Process where Name = 'WINWORD.EXE'")
  If procs.Count > 0 Then
    WScript.Echo "A Word is already running (" & procs.Count & " process(es)). Close it first: CreateObject may hand this script that Word, and Quit would close it."
    WScript.Quit 2
  End If
End Sub

Function Unesc(text)
  Dim pos, code
  Unesc = text
  pos = InStr(Unesc, "{U+")
  Do While pos > 0
    code = Mid(Unesc, pos + 3, 4)
    Unesc = Left(Unesc, pos - 1) & ChrW(CLng("&H" & code)) & Mid(Unesc, pos + 8)
    pos = InStr(Unesc, "{U+")
  Loop
End Function

' How often the find text stands inside the replacement: 0 for a true replacement, 1 when the edit only adds text next to the find text
' (an insertion), so that after the edit the find text is expected exactly that many times (2026-11-17, E-7).
' The text of the edited range with its left and right context, read back from the document (the proof that the edit did what the edit file says).
Function ChunkText(r, leftT, rightT)
  Dim c
  Set c = StoryRange()
  c.SetRange r.Start - Len(leftT), r.End + Len(rightT)
  ChunkText = c.Text
End Function

Function OccIn(text, part)
  OccIn = Int((Len(text) - Len(Replace(text, part, ""))) / Len(part))
End Function

Function CountOf(text)
  Dim r, n
  n = 0
  Set r = StoryRange()
  With r.Find
    .ClearFormatting
    .Text = text
    .Forward = True
    .Wrap = 0
    .MatchCase = True
    .MatchWildcards = False
    .MatchWholeWord = False
    .Format = False
  End With
  Do While r.Find.Execute
    n = n + 1
    r.Collapse 0
    If n > 200 Then Exit Do
  Loop
  CountOf = n
End Function

Sub Main(target, editsFile)
  Dim ts, line, parts, label, leftT, targT, rightT, newT, findT, repT, n, nNew, ok, skipped, r, paraText
  On Error Resume Next
  Set gWord = CreateObject("Word.Application")
  If Err.Number <> 0 Then gErrText = "cannot start Word": Exit Sub
  gWord.Visible = False
  gWord.DisplayAlerts = 0
  gWord.AutomationSecurity = 3
  Set gDoc = gWord.Documents.Open(target, False, False)
  If Err.Number <> 0 Then gErrText = "cannot open " & target & " (" & Err.Description & ")": Exit Sub
  gDoc.TrackRevisions = False
  gOut.WriteLine "story " & gStory & " | opened " & gDoc.FullName & " | paragraphs " & gDoc.Paragraphs.Count & " | text length " & Len(gDoc.Content.Text) & " | TrackRevisions " & gDoc.TrackRevisions & " | revisions " & gDoc.Revisions.Count
  ok = 0: skipped = 0
  Set ts = gFso.OpenTextFile(editsFile, 1, False)
  Do While Not ts.AtEndOfStream
    line = ts.ReadLine
    If Len(line) > 0 And Left(line, 1) <> "#" Then
      parts = Split(line, vbTab)
      If UBound(parts) <> 4 Then
        gOut.WriteLine "SKIPPED bad line (need 5 TAB fields): " & Left(line, 60): skipped = skipped + 1
      Else
        label = parts(0): leftT = Unesc(parts(1)): targT = Unesc(parts(2)): rightT = Unesc(parts(3)): repT = Unesc(parts(4))
        findT = leftT & targT & rightT
        newT = leftT & repT & rightT
        n = CountOf(findT)
        If n <> 1 Then
          gOut.WriteLine "SKIPPED " & label & ": the find text occurs " & n & " time(s), expected exactly 1"
          skipped = skipped + 1
        Else
          Set r = StoryRange()
          With r.Find
            .ClearFormatting
            .Text = findT
            .Forward = True
            .Wrap = 0
            .MatchCase = True
            .MatchWildcards = False
            .Format = False
          End With
          If Not r.Find.Execute Then
            gOut.WriteLine "FAILED " & label & ": find did not run": skipped = skipped + 1
          Else
            r.SetRange r.Start + Len(leftT), r.End - Len(rightT)
            If r.Text <> targT Then
              gOut.WriteLine "FAILED " & label & ": the narrowed range reads [" & r.Text & "] not the target": skipped = skipped + 1
            Else
              r.Text = repT
              If Err.Number <> 0 Then
                gOut.WriteLine "FAILED " & label & ": " & Err.Number & " " & Err.Description: Err.Clear: skipped = skipped + 1
              ElseIf ChunkText(r, leftT, rightT) = newT Then
                gOut.WriteLine "OK      " & label
                ok = ok + 1
              Else
                gOut.WriteLine "CHECK   " & label & ": after the replace the text reads [" & ChunkText(r, leftT, rightT) & "] and not [" & newT & "]"
                skipped = skipped + 1
              End If
            End If
          End If
        End If
      End If
    End If
  Loop
  ts.Close
  gOut.WriteLine "edits made: " & ok & " | skipped or failed: " & skipped & " | revisions now " & gDoc.Revisions.Count & " | paragraphs " & gDoc.Paragraphs.Count
  If skipped = 0 Then
    Err.Clear
    gDoc.Save
    If Err.Number <> 0 Then gErrText = "save failed (" & Err.Description & ")": Exit Sub
    gOut.WriteLine "saved " & target
  Else
    gOut.WriteLine "NOT SAVED: some edits were skipped or failed; the document on disk is unchanged"
  End If
End Sub

If WScript.Arguments.Count < 3 Then
  WScript.Echo "usage: cscript //nologo docm-text-edit.vbs <docm> <edits.txt> <result.txt> [body|footnotes]"
  WScript.Quit 1
End If
RefuseIfWordRunning
Set gFso = CreateObject("Scripting.FileSystemObject")
Set gOut = gFso.CreateTextFile(WScript.Arguments(2), True, False)
gStory = "body"
If WScript.Arguments.Count >= 4 Then gStory = LCase(WScript.Arguments(3))
If gStory <> "body" And gStory <> "footnotes" Then
  WScript.Echo "the fourth argument must be body or footnotes"
  WScript.Quit 1
End If
Main WScript.Arguments(0), WScript.Arguments(1)
If gErrText <> "" Then gOut.WriteLine "FAILED: " & gErrText
On Error Resume Next
If Not gDoc Is Nothing Then gDoc.Close 0
If Not gWord Is Nothing Then gWord.Quit 0
gOut.WriteLine "done"
gOut.Close
