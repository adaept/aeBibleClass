' run-export.vbs - run the scripture-only export (basBibleOnlyExport.ExportScriptureOnlyDocx) on a COPY, in a hidden Word of its own.
'
'   cscript //nologo run-export.vbs <src folder> <source.docm> <result.txt>
'
' Imports every module of <src folder> into a scratch document (as compile-check.vbs does), then runs
' ExportScriptureOnlyDocx with <source.docm>. The export derives every path from the folder of <source.docm>: the working copy
' (RadiantWordBible_work.docm), the result (RadiantWordBible.docx), the report folder rpt\ and the python steps py\strip_ribbon.py and
' py\verify_char_style_change.py. So put the source copy in its OWN new folder under C:\adaept\aeTest\ together with a py\ folder
' (those two scripts) and an empty rpt\ folder; the baseline docx (Bible\v0.0.RadiantWordBible.docx) is optional, its check is skipped
' and logged when absent. Never point this at the real docm. A run takes minutes (Run blocks until the export ends): start it in
' the background and poll <result.txt>. Rules and cautions: README.md.
Option Explicit

Dim gWord, gScratch, gOut, gFso, gErrText, gOutPath

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

Sub Main(srcFolder, source)
  Dim folder, file, nm, n, tmp, t0, vbc
  On Error Resume Next
  Set gWord = CreateObject("Word.Application")
  If Err.Number <> 0 Then Fail "cannot start Word": Exit Sub
  gWord.Visible = False
  gWord.DisplayAlerts = 0
  Set gScratch = gWord.Documents.Add()
  Set folder = gFso.GetFolder(srcFolder)
  gScratch.VBProject.VBComponents.Import folder.Files.Item("IaeLongProcessClass.cls").Path
  For Each file In folder.Files
    nm = file.Name
    If (LCase(Right(nm, 4)) = ".bas" Or LCase(Right(nm, 4)) = ".cls") And nm <> "ThisDocument.cls" And nm <> "Module1.bas" And Left(nm, 1) <> "X" And nm <> "IaeLongProcessClass.cls" Then
      gScratch.VBProject.VBComponents.Import file.Path
      If Err.Number <> 0 Then Fail "import failed: " & nm: Exit Sub
      n = n + 1
    End If
  Next
  gOut.WriteLine "imported " & n & " modules"
  ' A wrapper with its own error trap, so a runtime error inside the export is written to <result.txt>.err.txt instead of reaching
  ' this script as the opaque "Unknown runtime error" (40040).
  Set vbc = gScratch.VBProject.VBComponents.Add(1)
  vbc.Name = "zRunExport"
  vbc.CodeModule.AddFromString "Public Sub zRun(ByVal src As String, ByVal errFile As String)" & vbCrLf & _
    "    On Error GoTo E" & vbCrLf & _
    "    Dim g As Integer: g = FreeFile" & vbCrLf & _
    "    Open errFile & "".started"" For Output As #g" & vbCrLf & _
    "    Print #g, ""entered, active document: "" & ActiveDocument.Name & "", documents: "" & Documents.Count" & vbCrLf & _
    "    Close #g" & vbCrLf & _
    "    basBibleOnlyExport.ExportScriptureOnlyDocx src" & vbCrLf & _
    "    Exit Sub" & vbCrLf & _
    "E:" & vbCrLf & _
    "    Dim f As Integer: f = FreeFile" & vbCrLf & _
    "    Open errFile For Output As #f" & vbCrLf & _
    "    Print #f, Err.Number & "" | "" & Err.Description & "" | "" & Err.Source" & vbCrLf & _
    "    Close #f" & vbCrLf & _
    "End Sub" & vbCrLf
  tmp = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\scratch-run-export.docm"
  If gFso.FileExists(tmp) Then gFso.DeleteFile tmp
  gScratch.SaveAs2 tmp, 13
  ' AutomationSecurity is NOT forced to 3 here (unlike the other scripts): the export opens its own working copy from VBA, and
  ' with macros force-disabled Word answers that Documents.Open with error 40040 "Unknown runtime error" after about 6 s (probe,
  ' 2026-11-11). The copy's Document_Open then runs (it only wires events and checks for a conflicting ribbon add-in).
  gScratch.Activate
  gOut.WriteLine "start " & Now & " source " & source
  gOut.Flush
  t0 = Timer
  Err.Clear
  gWord.Run "zRunExport.zRun", source, gOutPath & ".err.txt"
  If Err.Number <> 0 Then Fail "export raised an error": Exit Sub
  gOut.WriteLine "export returned after " & FormatNumber(Timer - t0, 0) & " s (the verdict is in rpt\RadiantWordBibleExport.txt beside the source)"
End Sub

If WScript.Arguments.Count < 3 Then
  WScript.Echo "usage: cscript //nologo run-export.vbs <src folder> <source.docm> <result.txt>"
  WScript.Quit 1
End If
RefuseIfWordRunning
Set gFso = CreateObject("Scripting.FileSystemObject")
gOutPath = WScript.Arguments(2)
Set gOut = gFso.CreateTextFile(WScript.Arguments(2), True, False)
Main WScript.Arguments(0), WScript.Arguments(1)
If gErrText <> "" Then gOut.WriteLine "FAILED: " & gErrText
gOut.Close
On Error Resume Next
If Not gScratch Is Nothing Then gScratch.Close False
If Not gWord Is Nothing Then gWord.Quit
On Error GoTo 0
If gErrText <> "" Then
  WScript.Echo "FAILED: " & gErrText
  WScript.Quit 1
End If
WScript.Echo "done: " & WScript.Arguments(2)
