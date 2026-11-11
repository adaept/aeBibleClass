' RT-13 style-preservation check on a fresh COPY. args: 0 = copy path, 1 = basFontRunAudit.bas, 2 = basDocmHygiene.bas
Option Explicit
Dim word, scratch, doc, dn, t0, nf, hits(), n, i, r, f, line, a, b, fso, rep, diffs
Set fso = CreateObject("Scripting.FileSystemObject")
Set rep = fso.CreateTextFile("C:\adaept\aeTest\vba-test\hyg-compare.txt", True, False)
Set word = CreateObject("Word.Application")
word.Visible = False: word.DisplayAlerts = 0
Set scratch = word.Documents.Add()
scratch.VBProject.VBComponents.Import WScript.Arguments(1)
scratch.VBProject.VBComponents.Import WScript.Arguments(2)
scratch.SaveAs2 "C:\adaept\aeTest\vba-test\scratch-hyg2.docm", 13
word.AutomationSecurity = 3
Set doc = word.Documents.Open(WScript.Arguments(0), False, False)
scratch.Activate
dn = doc.Name

Function Props(rg)
  Props = "cs=" & rg.CharacterStyle & "|b=" & rg.Font.Bold & "|i=" & rg.Font.Italic & "|sz=" & rg.Font.Size & "|col=" & rg.Font.Color & "|u=" & rg.Font.Underline & "|lang=" & rg.LanguageID & "|ps=" & rg.Paragraphs(1).Style & "|txt=" & rg.Text
End Function
Sub SaveText(path)
  Dim st
  Set st = CreateObject("ADODB.Stream")
  st.Type = 2: st.Charset = "utf-8": st.Open
  st.WriteText doc.Content.Text
  st.SaveToFile path, 2
  st.Close
End Sub

' 1. collect the +Body runs that resolve outside the allowlist (main text + footnotes), outside BodyText
ReDim hits(2000)
n = 0
Dim story, rng, rr, last, res, ps
For Each story In doc.StoryRanges
  Set rng = story
  Do While Not rng Is Nothing
    Set rr = rng.Duplicate
    last = -1
    With rr.Find
      .ClearFormatting: .Text = "": .Font.Name = "+Body": .Format = True: .Forward = True: .Wrap = 0
      Do While .Execute
        If rr.Start <= last Then Exit Do
        last = rr.Start
        res = rr.Font.Name
        If res <> "Carlito" And res <> "Noto Sans" And res <> "Liberation Serif" Then
          hits(n) = rng.StoryType & ";" & rr.Start & ";" & rr.End & ";" & Props(rr)
          n = n + 1
        End If
        rr.Collapse 0
      Loop
    End With
    Set rng = rng.NextStoryRange
  Loop
Next
rep.WriteLine "runs recorded before the fix: " & n
SaveText "C:\adaept\aeTest\vba-test\text-before.txt"

' 2. fix
nf = word.Run("basDocmHygiene.FixDirectFontRuns", dn, False)
rep.WriteLine "FixDirectFontRuns returned " & nf & " lastErr=[" & word.Run("basDocmHygiene.DocmHygieneLastError") & "]"
SaveText "C:\adaept\aeTest\vba-test\text-after.txt"

' 3. compare each recorded run: same properties except the font (and positions are unchanged by a reset)
diffs = 0
For i = 0 To n - 1
  a = Split(hits(i), ";")
  Dim sty, st2, kind
  For Each story In doc.StoryRanges
    If story.StoryType = CLng(a(0)) Then Set rng = story
  Next
  Set r = rng.Duplicate
  r.SetRange CLng(a(1)), CLng(a(2))
  line = a(0) & ";" & a(1) & ";" & a(2) & ";" & Props(r)
  If line <> hits(i) Then
    diffs = diffs + 1
    rep.WriteLine "DIFF  before: " & hits(i)
    rep.WriteLine "      after : " & line
  End If
  rep.WriteLine "font now: " & r.Font.Name & " | " & Left(a(3), 90)
Next
rep.WriteLine "runs with a property difference other than font: " & diffs
rep.Close
doc.Close False
scratch.Close False
word.Quit
WScript.Echo "done"
