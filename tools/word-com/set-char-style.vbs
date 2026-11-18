' set-char-style.vbs - set a CHARACTER style on the FIRST occurrence of each of some words inside ONE paragraph of the first 80 paragraphs of a docm (the front
' matter), in a hidden Word of its own, and SAVE the document. Written for the italic titles in the new Introduction sentence (E-6, option 3, 2026-11-17).
'
'   cscript //nologo set-char-style.vbs <docm> <paragraph-start> <style> <words, separated by ;> <result.txt>
'
' <paragraph-start>: the paragraph that starts with this text (found exactly once among the first 80 paragraphs).
' <style>: the character style, for example RWBFrontItalic. <words>: each word or phrase must occur in that paragraph; the first occurrence gets the style.
' Characters outside ASCII are written as {U+2019} (hex, 4 digits). Rules as for apply-edits.vbs: Word closed first (refused otherwise), a COPY first, verify the
' saved XML afterwards. Macros off (AutomationSecurity 3). Nothing is saved unless every word was found and now carries the style.
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

Sub Main(target, startText, styleName, wordList)
  Dim i, hit, hits, para, words, w, r, found, done
  On Error Resume Next
  Set gWord = CreateObject("Word.Application")
  If Err.Number <> 0 Then gErrText = "cannot start Word": Exit Sub
  gWord.Visible = False
  gWord.DisplayAlerts = 0
  gWord.AutomationSecurity = 3
  Set gDoc = gWord.Documents.Open(target, False, False)
  If Err.Number <> 0 Then gErrText = "cannot open " & target & " (" & Err.Description & ")": Exit Sub
  gDoc.TrackRevisions = False
  gOut.WriteLine "opened " & gDoc.FullName & " | paragraphs " & gDoc.Paragraphs.Count
  hits = 0: hit = 0
  For i = 1 To 80
    If Left(gDoc.Paragraphs(i).Range.Text, Len(startText)) = startText Then hits = hits + 1: hit = i
  Next
  If hits <> 1 Then gErrText = "the paragraph that starts with the given text was found " & hits & " time(s) in the first 80 paragraphs, expected exactly 1": Exit Sub
  Set para = gDoc.Paragraphs(hit)
  gOut.WriteLine "paragraph " & hit & " (style " & para.Style & "): " & Left(para.Range.Text, 60)
  words = Split(wordList, ";")
  done = 0
  For Each w In words
    Set r = para.Range
    With r.Find
      .ClearFormatting
      .Text = w
      .Forward = True
      .Wrap = 0
      .MatchCase = True
      .MatchWildcards = False
      .Format = False
    End With
    found = r.Find.Execute
    If Not found Then gErrText = "not found in the paragraph: " & w: Exit Sub
    r.Style = styleName
    If Err.Number <> 0 Then gErrText = "styling failed for " & w & " (" & Err.Number & " " & Err.Description & ")": Exit Sub
    If CStr(r.CharacterStyle) <> styleName Then gErrText = "after styling, " & w & " has character style " & r.CharacterStyle: Exit Sub
    gOut.WriteLine "OK      " & w & " (offset " & (r.Start - para.Range.Start) & ") now " & r.CharacterStyle
    done = done + 1
  Next
  gDoc.Save
  If Err.Number <> 0 Then gErrText = "save failed (" & Err.Description & ")": Exit Sub
  gOut.WriteLine "styled " & done & " word(s); saved " & target
End Sub

If WScript.Arguments.Count < 5 Then
  WScript.Echo "usage: cscript //nologo set-char-style.vbs <docm> <paragraph-start> <style> <words;words> <result.txt>"
  WScript.Quit 1
End If
RefuseIfWordRunning
Set gFso = CreateObject("Scripting.FileSystemObject")
Set gOut = gFso.CreateTextFile(WScript.Arguments(4), True, False)
Main WScript.Arguments(0), Unesc(WScript.Arguments(1)), WScript.Arguments(2), Unesc(WScript.Arguments(3))
If gErrText <> "" Then gOut.WriteLine "FAILED: " & gErrText
On Error Resume Next
If Not gDoc Is Nothing Then gDoc.Close 0
If Not gWord Is Nothing Then gWord.Quit 0
gOut.WriteLine "done"
gOut.Close
