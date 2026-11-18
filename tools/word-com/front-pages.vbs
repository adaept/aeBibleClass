' front-pages.vbs - read-only: print the page number of the start and of the end of each of the first N paragraphs of a docm (the front matter), in a hidden
' Word of its own. The docm is opened read-only and is NOT saved. Written for the page-fit check of the Foreword (E-8, 2026-11-17).
'
'   cscript //nologo front-pages.vbs <docm> <N> <result.txt>
'
' A hidden Word may paginate a little differently from the owner's visible Word (printer, fonts): treat the answer as a pre-check and look at the page in Word.
Option Explicit
Dim gWord, gDoc, gOut, gFso

Sub RefuseIfWordRunning()
  Dim svc, procs
  Set svc = GetObject("winmgmts:\\.\root\cimv2")
  Set procs = svc.ExecQuery("select ProcessId from Win32_Process where Name = 'WINWORD.EXE'")
  If procs.Count > 0 Then
    WScript.Echo "A Word is already running (" & procs.Count & " process(es)). Close it first."
    WScript.Quit 2
  End If
End Sub

If WScript.Arguments.Count < 3 Then
  WScript.Echo "usage: cscript //nologo front-pages.vbs <docm> <N> <result.txt>"
  WScript.Quit 1
End If
RefuseIfWordRunning
Set gFso = CreateObject("Scripting.FileSystemObject")
Set gOut = gFso.CreateTextFile(WScript.Arguments(2), True, False)
On Error Resume Next
Set gWord = CreateObject("Word.Application")
gWord.Visible = False
gWord.DisplayAlerts = 0
gWord.AutomationSecurity = 3
Set gDoc = gWord.Documents.Open(WScript.Arguments(0), False, True)
If Err.Number <> 0 Then
  gOut.WriteLine "FAILED: cannot open (" & Err.Description & ")"
Else
  gDoc.Repaginate
  Dim i, n, para, r1, r2, txt
  n = CLng(WScript.Arguments(1))
  For i = 1 To n
    Set para = gDoc.Paragraphs(i)
    Set r1 = gDoc.Range(para.Range.Start, para.Range.Start)
    Set r2 = gDoc.Range(para.Range.End - 1, para.Range.End - 1)
    txt = Left(para.Range.Text, 40)
    gOut.WriteLine i & " | " & para.Style & " | page " & r1.Information(3) & " to " & r2.Information(3) & " | " & Replace(txt, vbCr, "")
  Next
  gOut.WriteLine "pages in the document: " & gDoc.ComputeStatistics(2)
End If
If Not gDoc Is Nothing Then gDoc.Close 0
If Not gWord Is Nothing Then gWord.Quit 0
gOut.WriteLine "done"
gOut.Close
