' run-counts.vbs - run public functions of VBA modules against a COPY of a document, in a hidden Word of its own.
'
'   cscript //nologo run-counts.vbs <copy.docm> <modules.bas;more.cls> <Module.Function;Module.Other> <result.txt>
'
' The modules are imported into a scratch document (the copy is never modified: it is opened read-only with macros off), then
' each "Module.Function" is run with the copy's file name as its argument (the hygiene functions take an optional document
' name). Example, the five clean-up counts of the RWB docm:
'   run-counts.vbs C:\adaept\aeTest\vba-test\copy.docm "C:\adaept\aeBibleClass\src\basFontRunAudit.bas;C:\adaept\aeBibleClass\src\basDocmHygiene.bas"
'     "basDocmHygiene.CountDirectFontRunsOutsideAllowlist;basDocmHygiene.CountVerseRunsNotEnglish;basDocmHygiene.CountThemeFontsOutsideAllowlist;
'      basDocmHygiene.CountStyleSlotFontsOutsideAllowlist;basDocmHygiene.CountStrayDirectFormattingRuns" result.txt
' Rules (see README.md): close every Word first (this script refuses otherwise), work on copies under C:\adaept\aeTest, a run
' of the full docm takes minutes (start it in the background and poll the result file), the Immediate window is not visible.
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

Sub Main(target, modules, funcs)
  Dim m, fn, tmp, dn
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
  tmp = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\scratch-run-counts.docm"
  If gFso.FileExists(tmp) Then gFso.DeleteFile tmp
  gScratch.SaveAs2 tmp, 13
  gWord.AutomationSecurity = 3
  Set gDoc = gWord.Documents.Open(target, False, True)
  If Err.Number <> 0 Then Fail "cannot open " & target: Exit Sub
  gScratch.Activate
  dn = gDoc.Name
  gOut.WriteLine "document " & dn & ", text length " & Len(gDoc.Content.Text)
  For Each fn In Split(funcs, ";")
    gOut.WriteLine Trim(fn) & " = " & gWord.Run(Trim(fn), dn)
    If Err.Number <> 0 Then Fail "run failed: " & fn: Exit Sub
  Next
End Sub

If WScript.Arguments.Count < 4 Then
  WScript.Echo "usage: cscript //nologo run-counts.vbs <copy.docm> <modules.bas;more.cls> <Module.Function;...> <result.txt>"
  WScript.Quit 1
End If
RefuseIfWordRunning
Set gFso = CreateObject("Scripting.FileSystemObject")
Set gOut = gFso.CreateTextFile(WScript.Arguments(3), True, False)
Main WScript.Arguments(0), WScript.Arguments(1), WScript.Arguments(2)
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
WScript.Echo "done: " & WScript.Arguments(3)
