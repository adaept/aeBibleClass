' run-pdf.vbs - run the scripture PDF export (basScripturePdfExport.ExportScripturePdf) on a COPY of the finished docx, in a hidden Word.
'
'   cscript //nologo run-pdf.vbs <src folder> <finished.docx> <pdf folder> <result.txt>
'
' Imports every module of <src folder> into a scratch document (as run-export.vbs does), then runs ExportScripturePdf, which opens the
' docx READ-ONLY and writes rwb.pdf (screen) and rwb-print.pdf (print-ready) into <pdf folder>. Put the docx copy in its own new
' folder under C:\adaept\aeTest\ (never the release folder). A run takes about 3 to 4 minutes: start it in the background and poll
' the report <pdf folder>\ScripturePdfExport.txt or <docx folder>\rpt\ScripturePdfExport.txt (the verdict is its last line).
' Runtime errors land in <result.txt>.err.txt. Refuses to run while any Word is open (README.md).
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

Sub Main(srcFolder, docx, pdfFolder)
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
  ' A wrapper with its own error trap, so a runtime error inside the export is written to <result.txt>.err.txt.
  Set vbc = gScratch.VBProject.VBComponents.Add(1)
  vbc.Name = "zRunPdf"
  vbc.CodeModule.AddFromString "Public Sub zRun(ByVal docx As String, ByVal pdfFolder As String, ByVal errFile As String)" & vbCrLf & _
    "    On Error GoTo E" & vbCrLf & _
    "    basScripturePdfExport.ExportScripturePdf docx, pdfFolder" & vbCrLf & _
    "    Exit Sub" & vbCrLf & _
    "E:" & vbCrLf & _
    "    Dim f As Integer: f = FreeFile" & vbCrLf & _
    "    Open errFile For Output As #f" & vbCrLf & _
    "    Print #f, Err.Number & "" | "" & Err.Description & "" | "" & Err.Source" & vbCrLf & _
    "    Close #f" & vbCrLf & _
    "End Sub" & vbCrLf
  tmp = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\scratch-run-pdf.docm"
  If gFso.FileExists(tmp) Then gFso.DeleteFile tmp
  gScratch.SaveAs2 tmp, 13
  gScratch.Activate
  gOut.WriteLine "start " & Now & " docx " & docx
  gOut.Flush
  t0 = Timer
  Err.Clear
  gWord.Run "zRunPdf.zRun", docx, pdfFolder, gOutPath & ".err.txt"
  If Err.Number <> 0 Then Fail "pdf export raised an error": Exit Sub
  gOut.WriteLine "export returned after " & FormatNumber(Timer - t0, 0) & " s (the verdict is in ScripturePdfExport.txt)"
End Sub

If WScript.Arguments.Count < 4 Then
  WScript.Echo "usage: cscript //nologo run-pdf.vbs <src folder> <finished.docx> <pdf folder> <result.txt>"
  WScript.Quit 1
End If
RefuseIfWordRunning
Set gFso = CreateObject("Scripting.FileSystemObject")
gOutPath = WScript.Arguments(3)
Set gOut = gFso.CreateTextFile(WScript.Arguments(3), True, False)
Main WScript.Arguments(0), WScript.Arguments(1), WScript.Arguments(2)
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
WScript.Echo "done: " & WScript.Arguments(3)
