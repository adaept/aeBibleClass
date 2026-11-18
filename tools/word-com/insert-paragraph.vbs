' insert-paragraph.vbs - insert ONE new paragraph after an existing paragraph of the first 80 paragraphs of a docm (the front matter), in a hidden
' Word of its own, and SAVE the document. Written for E-8 (the Foreword item "Formal voice.", 2026-11-17).
'
'   cscript //nologo insert-paragraph.vbs <docm> <after-text> <paragraph style> <lead-in> <body> <result.txt>
'
' <after-text>: the new paragraph goes after the FIRST of the first 80 paragraphs whose text starts with it (it must be found exactly once there).
' <paragraph style>: the paragraph style of the new paragraph, for example RWBFrontBody.
' <lead-in>: the first words of the new paragraph, set in the character style RWBFrontItalic (empty = no lead-in).
' <body>: the rest of the text (it starts with a space after a lead-in). Characters outside ASCII are written as {U+201C} (hex, 4 digits).
' Rules as for apply-edits.vbs: close every Word first (this refuses otherwise), run it on a COPY first, verify the saved XML afterwards. Macros stay off
' (AutomationSecurity 3), Track Changes is off. Nothing is saved unless the new paragraph reads back exactly as asked.
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

Sub Main(target, afterText, styleName, leadIn, body)
  Dim i, hit, hits, para, newPara, fullText, lead, body2, markR
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
    If Left(gDoc.Paragraphs(i).Range.Text, Len(afterText)) = afterText Then hits = hits + 1: hit = i
  Next
  If hits <> 1 Then gErrText = "the paragraph that starts with the after-text was found " & hits & " time(s) in the first 80 paragraphs, expected exactly 1": Exit Sub
  gOut.WriteLine "after paragraph " & hit & " (style " & gDoc.Paragraphs(hit).Style & "): " & Left(gDoc.Paragraphs(hit).Range.Text, 60)
  Set para = gDoc.Paragraphs(hit)
  para.Range.InsertParagraphAfter
  Set newPara = gDoc.Paragraphs(hit + 1)
  fullText = leadIn & body
  newPara.Range.InsertBefore fullText
  newPara.Style = styleName
  If Err.Number <> 0 Then gErrText = "insert failed (" & Err.Number & " " & Err.Description & ")": Exit Sub
  ' The new text may inherit a character style from its neighbours (seen 2026-11-17: the whole paragraph came out in RWBFrontItalic), so the body is
  ' set back to "no character style" (Default Paragraph Font, built-in -66: Word writes no w:rStyle for it) and only the lead-in gets RWBFrontItalic.
  Set body2 = gDoc.Range(newPara.Range.Start + Len(leadIn), newPara.Range.End - 1)
  body2.Style = gDoc.Styles(-66)
  Set markR = gDoc.Range(newPara.Range.End - 1, newPara.Range.End)
  markR.Style = gDoc.Styles(-66)
  If Len(leadIn) > 0 Then
    Set lead = gDoc.Range(newPara.Range.Start, newPara.Range.Start + Len(leadIn))
    lead.Style = "RWBFrontItalic"
  End If
  If Err.Number <> 0 Then gErrText = "styling failed (" & Err.Number & " " & Err.Description & ")": Exit Sub
  Set newPara = gDoc.Paragraphs(hit + 1)
  gOut.WriteLine "new paragraph " & (hit + 1) & " (style " & newPara.Style & "): " & Left(newPara.Range.Text, 80)
  gOut.WriteLine "paragraphs now " & gDoc.Paragraphs.Count
  If newPara.Range.Text <> fullText & vbCr Then gErrText = "the new paragraph does not read back as asked: " & newPara.Range.Text: Exit Sub
  If CStr(newPara.Style) <> styleName Then gErrText = "the new paragraph has style " & newPara.Style: Exit Sub
  gDoc.Save
  If Err.Number <> 0 Then gErrText = "save failed (" & Err.Description & ")": Exit Sub
  gOut.WriteLine "saved " & target
End Sub

If WScript.Arguments.Count < 6 Then
  WScript.Echo "usage: cscript //nologo insert-paragraph.vbs <docm> <after-text> <paragraph style> <lead-in> <body> <result.txt>"
  WScript.Quit 1
End If
RefuseIfWordRunning
Set gFso = CreateObject("Scripting.FileSystemObject")
Set gOut = gFso.CreateTextFile(WScript.Arguments(5), True, False)
Main WScript.Arguments(0), Unesc(WScript.Arguments(1)), WScript.Arguments(2), Unesc(WScript.Arguments(3)), Unesc(WScript.Arguments(4))
If gErrText <> "" Then gOut.WriteLine "FAILED: " & gErrText
On Error Resume Next
If Not gDoc Is Nothing Then gDoc.Close 0
If Not gWord Is Nothing Then gWord.Quit 0
gOut.WriteLine "done"
gOut.Close
