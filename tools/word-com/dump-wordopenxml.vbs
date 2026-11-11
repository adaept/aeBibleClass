' dump-wordopenxml.vbs - write Content.WordOpenXML of a COPY of a document to a UTF-8 file (about 31 MB, about 10 s for the RWB docm).
'
'   cscript //nologo dump-wordopenxml.vbs <copy.docm> <out.xml>
'
' The result is the same flat-OPC package that Office.js getOoxml() returns, so the Python tools in ..\..\py (docm_*.py) and
' the add-in's hygiene port can read it. For the SAVED file on disk use "unzip -p <docm> word/document.xml" instead (no Word).
' Refuses to run while any Word is open (see README.md).
Option Explicit

Dim gWord, gDoc, gErrText

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

Sub Main(target, outPath)
  Dim stream, xml
  On Error Resume Next
  Set gWord = CreateObject("Word.Application")
  If Err.Number <> 0 Then Fail "cannot start Word": Exit Sub
  gWord.Visible = False
  gWord.DisplayAlerts = 0
  gWord.AutomationSecurity = 3
  Set gDoc = gWord.Documents.Open(target, False, True)
  If Err.Number <> 0 Then Fail "cannot open " & target: Exit Sub
  xml = gDoc.Content.WordOpenXML
  If Err.Number <> 0 Then Fail "WordOpenXML failed": Exit Sub
  Set stream = CreateObject("ADODB.Stream")
  stream.Type = 2
  stream.Charset = "utf-8"
  stream.Open
  stream.WriteText xml
  stream.SaveToFile outPath, 2
  stream.Close
  WScript.Echo "WordOpenXML length " & Len(xml)
End Sub

If WScript.Arguments.Count < 2 Then
  WScript.Echo "usage: cscript //nologo dump-wordopenxml.vbs <copy.docm> <out.xml>"
  WScript.Quit 1
End If
RefuseIfWordRunning
Main WScript.Arguments(0), WScript.Arguments(1)
On Error Resume Next
If Not gDoc Is Nothing Then gDoc.Close False
If Not gWord Is Nothing Then gWord.Quit
On Error GoTo 0
If gErrText <> "" Then
  WScript.Echo "FAILED: " & gErrText
  WScript.Quit 1
End If
WScript.Echo "done: " & WScript.Arguments(1)
