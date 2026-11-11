' compile-check.vbs - import every module of src into a scratch document and force the compile of one class.
'
'   cscript //nologo compile-check.vbs <src folder> <ClassName> <test numbers 94;95;96> <result.txt>
'
' Used before importing an edited aeBibleClass.cls into the real docm: a compile error there stops the whole project. The
' script imports all .bas and .cls files of the folder (not ThisDocument.cls, Module1.bas, files starting with X; the interface
' class IaeLongProcessClass.cls first), adds a probe module that instantiates <ClassName>, and reads GetTestDescription(n) for
' each number. A hidden compile-error dialog would make the call hang: run it in the background and poll the result file; if no
' line appears after two minutes kill ONLY the Word this script started (README.md). The wiring test of Tests 95 to 98 used it.
Option Explicit

Dim gWord, gScratch, gOut, gFso, gErrText

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

Sub Main(srcFolder, className, numbers)
  Dim folder, file, nm, n, vbc, tmp, num, text
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
  Set vbc = gScratch.VBProject.VBComponents.Add(1)
  vbc.Name = "zChk"
  vbc.CodeModule.AddFromString "Public Function zChkDesc(ByVal n As Long) As String" & vbCrLf & "    Dim c As " & className & vbCrLf & "    Set c = New " & className & vbCrLf & "    zChkDesc = c.GetTestDescription(n)" & vbCrLf & "End Function" & vbCrLf
  tmp = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\scratch-compile-check.docm"
  If gFso.FileExists(tmp) Then gFso.DeleteFile tmp
  gScratch.SaveAs2 tmp, 13
  gOut.WriteLine "probe module added; running (a hang here means a compile error dialog)"
  For Each num In Split(numbers, ";")
    text = gWord.Run("zChk.zChkDesc", CLng(Trim(num)))
    If Err.Number <> 0 Then Fail "run failed for test " & num: Exit Sub
    gOut.WriteLine "Test " & Trim(num) & ": " & Left(text, 100)
  Next
  gOut.WriteLine "compiled without a dialog"
End Sub

If WScript.Arguments.Count < 4 Then
  WScript.Echo "usage: cscript //nologo compile-check.vbs <src folder> <ClassName> <numbers 94;95> <result.txt>"
  WScript.Quit 1
End If
RefuseIfWordRunning
Set gFso = CreateObject("Scripting.FileSystemObject")
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
