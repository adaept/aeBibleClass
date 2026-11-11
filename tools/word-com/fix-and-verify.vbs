' fix-and-verify.vbs - test a count/fix pair of a VBA module on a COPY of a document, in a hidden Word of its own.
'
'   cscript //nologo fix-and-verify.vbs <copy.docm> <modules.bas;more.cls> <Module.CountFunction> <Module.FixFunction> <result.txt>
'
' Runs: count, fix, count, fix again (a repeatable fix returns 0). It also saves the document text before and after to
' <result.txt>.before.txt and <result.txt>.after.txt (UTF-8) so the text can be compared with sha256sum: a fix must not change
' it. The copy is opened read-write and closed WITHOUT saving; never point this at the real document. For property
' comparisons (character style, bold, size, colour of the touched runs before and after: the check that found Font.Reset
' stripping character styles) see examples\hyg_rt13_compare.vbs. Rules and cautions: README.md.
Option Explicit

Dim gWord, gScratch, gDoc, gOut, gFso, gErrText

Sub RefuseIfWordRunning()
  Dim svc, procs
  Set svc = GetObject("winmgmts:\\.\root\cimv2")
  Set procs = svc.ExecQuery("select ProcessId from Win32_Process where Name = 'WINWORD.EXE'")
  If procs.Count > 0 Then
    WScript.Echo "A Word is already running (" & procs.Count & " process(es)). Close it first: CreateObject may hand this script that Word, and Quit would close it."
    WScript.Quit 2
  End If
End Sub

Sub Fail(text)
  gErrText = text & " (error " & Err.Number & ": " & Err.Description & ")"
  Err.Clear
End Sub

Sub SaveText(path)
  Dim stream
  Set stream = CreateObject("ADODB.Stream")
  stream.Type = 2
  stream.Charset = "utf-8"
  stream.Open
  stream.WriteText gDoc.Content.Text
  stream.SaveToFile path, 2
  stream.Close
End Sub

Sub Main(target, modules, countFn, fixFn, outPath)
  Dim m, tmp, dn, t0
  On Error Resume Next
  Set gWord = CreateObject("Word.Application")
  If Err.Number <> 0 Then Fail "cannot start Word": Exit Sub
  gWord.Visible = False
  gWord.DisplayAlerts = 0
  Set gScratch = gWord.Documents.Add()
  For Each m In Split(modules, ";")
    gScratch.VBProject.VBComponents.Import Trim(m)
    If Err.Number <> 0 Then Fail "import failed: " & m: Exit Sub
  Next
  tmp = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\scratch-fix-and-verify.docm"
  If gFso.FileExists(tmp) Then gFso.DeleteFile tmp
  gScratch.SaveAs2 tmp, 13
  gWord.AutomationSecurity = 3
  Set gDoc = gWord.Documents.Open(target, False, False)
  If Err.Number <> 0 Then Fail "cannot open " & target: Exit Sub
  gScratch.Activate
  dn = gDoc.Name
  gOut.WriteLine "document " & dn & ", text length " & Len(gDoc.Content.Text)
  t0 = Timer
  gOut.WriteLine "BEFORE " & countFn & " = " & gWord.Run(countFn, dn) & " (" & FormatNumber(Timer - t0, 0) & " s)"
  If Err.Number <> 0 Then Fail "count failed": Exit Sub
  SaveText outPath & ".before.txt"
  t0 = Timer
  gOut.WriteLine "FIX " & fixFn & " = " & gWord.Run(fixFn, dn) & " (" & FormatNumber(Timer - t0, 0) & " s)"
  If Err.Number <> 0 Then Fail "fix failed": Exit Sub
  SaveText outPath & ".after.txt"
  gOut.WriteLine "AFTER " & countFn & " = " & gWord.Run(countFn, dn)
  gOut.WriteLine "SECOND FIX (expected 0) = " & gWord.Run(fixFn, dn)
  If Err.Number <> 0 Then Fail "second pass failed": Exit Sub
End Sub

If WScript.Arguments.Count < 5 Then
  WScript.Echo "usage: cscript //nologo fix-and-verify.vbs <copy.docm> <modules.bas;more.cls> <Module.CountFunction> <Module.FixFunction> <result.txt>"
  WScript.Quit 1
End If
RefuseIfWordRunning
Set gFso = CreateObject("Scripting.FileSystemObject")
Set gOut = gFso.CreateTextFile(WScript.Arguments(4), True, False)
Main WScript.Arguments(0), WScript.Arguments(1), WScript.Arguments(2), WScript.Arguments(3), WScript.Arguments(4)
If gErrText <> "" Then gOut.WriteLine "FAILED: " & gErrText
gOut.Close
On Error Resume Next
If Not gDoc Is Nothing Then gDoc.Close False
If Not gScratch Is Nothing Then gScratch.Close False
If Not gWord Is Nothing Then gWord.Quit
On Error GoTo 0
If gErrText <> "" Then
  WScript.Echo "FAILED: " & gErrText
  WScript.Quit 1
End If
WScript.Echo "done: " & WScript.Arguments(4)
