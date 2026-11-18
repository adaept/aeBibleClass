' apply-edits.vbs - make a list of exact text replacements in a docm, in a hidden Word of its own, and SAVE the document.
'
'   cscript //nologo apply-edits.vbs <docm> <edits.txt> <result.txt>
'
' <edits.txt>: one edit per line, three TAB-separated fields: label, find text, replacement text (ASCII only; lines starting with # are skipped).
' A character outside ASCII is written as {U+2018} (hex code point, 4 digits), for example a curly quote; the script turns it into the character (2026-11-17, E-10).
' For every edit the find text must occur EXACTLY ONCE in the document (whole body, case sensitive, no wildcards); otherwise that edit is skipped
' and reported, nothing is guessed. The replacement goes through Find.Execute Replace:=wdReplaceOne, so the formatting of the found text stays.
' After the replacement the new text must occur exactly ONE time more than before (it may already occur elsewhere) and the old text not at all. Track Changes is switched off. Macros stay off (AutomationSecurity 3).
' Rules (see README.md): close every Word first (this script refuses otherwise); make a backup of the docm before you point this at the real file; run
' the SAME edits on a copy first and compare (the verification reads the saved XML, for example tools/verify-docm-edits.mjs). Written for FU-50 (2026-11-14).
Option Explicit
Dim gWord, gDoc, gOut, gFso, gErrText

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

Function CountOf(text)
  Dim r, n
  n = 0
  Set r = gDoc.Content
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
    If n > 5 Then Exit Do
  Loop
  CountOf = n
End Function

Sub Main(target, editsFile)
  Dim ts, line, parts, label, findT, repT, n, nNew, ok, skipped, r, paraText
  On Error Resume Next
  Set gWord = CreateObject("Word.Application")
  If Err.Number <> 0 Then gErrText = "cannot start Word": Exit Sub
  gWord.Visible = False
  gWord.DisplayAlerts = 0
  gWord.AutomationSecurity = 3
  Set gDoc = gWord.Documents.Open(target, False, False)
  If Err.Number <> 0 Then gErrText = "cannot open " & target & " (" & Err.Description & ")": Exit Sub
  gDoc.TrackRevisions = False
  gOut.WriteLine "opened " & gDoc.FullName & " | paragraphs " & gDoc.Paragraphs.Count & " | text length " & Len(gDoc.Content.Text) & " | TrackRevisions " & gDoc.TrackRevisions & " | revisions " & gDoc.Revisions.Count
  ok = 0: skipped = 0
  Set ts = gFso.OpenTextFile(editsFile, 1, False)
  Do While Not ts.AtEndOfStream
    line = ts.ReadLine
    If Len(line) > 0 And Left(line, 1) <> "#" Then
      parts = Split(line, vbTab)
      If UBound(parts) <> 2 Then
        gOut.WriteLine "SKIPPED bad line (need 3 TAB fields): " & Left(line, 60): skipped = skipped + 1
      Else
        label = parts(0): findT = Unesc(parts(1)): repT = Unesc(parts(2))
        n = CountOf(findT)
        nNew = CountOf(repT)
        If n <> 1 Then
          gOut.WriteLine "SKIPPED " & label & ": the find text occurs " & n & " time(s), expected exactly 1"
          skipped = skipped + 1
        Else
          Set r = gDoc.Content
          With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = findT
            .Replacement.Text = repT
            .Forward = True
            .Wrap = 0
            .MatchCase = True
            .MatchWildcards = False
            .Format = False
          End With
          r.Find.Execute , , , , , , , , , , 1
          If Err.Number <> 0 Then
            gOut.WriteLine "FAILED " & label & ": " & Err.Number & " " & Err.Description: Err.Clear: skipped = skipped + 1
          ElseIf CountOf(repT) = nNew + 1 And CountOf(findT) = 0 Then
            gOut.WriteLine "OK      " & label
            ok = ok + 1
          Else
            gOut.WriteLine "CHECK   " & label & ": after the replace the new text occurs " & CountOf(repT) & " time(s) (expected " & (nNew + 1) & ") and the old text " & CountOf(findT)
            skipped = skipped + 1
          End If
        End If
      End If
    End If
  Loop
  ts.Close
  gOut.WriteLine "edits made: " & ok & " | skipped or failed: " & skipped & " | revisions now " & gDoc.Revisions.Count & " | paragraphs " & gDoc.Paragraphs.Count
  If skipped = 0 Then
    gDoc.Save
    If Err.Number <> 0 Then gErrText = "save failed (" & Err.Description & ")": Exit Sub
    gOut.WriteLine "saved " & target
  Else
    gOut.WriteLine "NOT SAVED: some edits were skipped or failed; the document on disk is unchanged"
  End If
End Sub

If WScript.Arguments.Count < 3 Then
  WScript.Echo "usage: cscript //nologo apply-edits.vbs <docm> <edits.txt> <result.txt>"
  WScript.Quit 1
End If
RefuseIfWordRunning
Set gFso = CreateObject("Scripting.FileSystemObject")
Set gOut = gFso.CreateTextFile(WScript.Arguments(2), True, False)
Main WScript.Arguments(0), WScript.Arguments(1)
If gErrText <> "" Then gOut.WriteLine "FAILED: " & gErrText
On Error Resume Next
If Not gDoc Is Nothing Then gDoc.Close 0
If Not gWord Is Nothing Then gWord.Quit 0
gOut.WriteLine "done"
gOut.Close
