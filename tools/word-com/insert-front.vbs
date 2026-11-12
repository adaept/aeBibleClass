' insert-front.vbs - RT-22a: create the RWB front-matter styles and insert the front-matter text into a COPY of the docm, in a hidden Word of its own.
'
'   cscript //nologo insert-front.vbs <src folder> <copy.docm> <front-matter.txt> <result.txt>
'
' Imports every module of <src folder> (as compile-check.vbs does), opens the COPY, then runs CreateFrontMatterStyles, FixStyleSlotFonts,
' CountFrontMatterStylesOutsideSpec and InsertFrontMatter(<front-matter.txt>), tries a second insert (expected -1: it refuses), writes the section
' page setup (size, margins, vertical alignment, footers) and, for the first 20 paragraphs, style, page, vertical position and text to <result.txt>,
' and SAVES the copy. A hidden Word repaginates slowly: the vertical positions are indicative, look at the pages in a visible Word.
' The proof used in RT-22a: py\docm_front_verify.py on the saved copy, py\docm_front_compare.py against the owner's saved docm, then run-export.vbs on it.
' Never point this at the real docm. Rules and cautions: README.md.
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

Sub Main(srcFolder, target, textFile)
  Dim folder, file, nm, n, tmp, dn, i, p
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
  tmp = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\scratch-insert-front.docm"
  If gFso.FileExists(tmp) Then gFso.DeleteFile tmp
  gScratch.SaveAs2 tmp, 13
  gWord.AutomationSecurity = 3
  Set gDoc = gWord.Documents.Open(target, False, False)
  If Err.Number <> 0 Then Fail "cannot open " & target: Exit Sub
  gScratch.Activate
  dn = gDoc.Name
  gOut.WriteLine "sections before " & gDoc.Sections.Count & ", text length " & Len(gDoc.Content.Text)
  gOut.WriteLine "CreateFrontMatterStyles = " & gWord.Run("basRwbFrontMatter.CreateFrontMatterStyles", dn) & " | last error [" & gWord.Run("basRwbFrontMatter.FrontMatterLastError") & "]"
  gOut.WriteLine "FixStyleSlotFonts = " & gWord.Run("basDocmHygiene.FixStyleSlotFonts", dn)
  gOut.WriteLine "CountFrontMatterStylesOutsideSpec = " & gWord.Run("basRwbFrontMatter.CountFrontMatterStylesOutsideSpec", dn)
  gOut.WriteLine "InsertFrontMatter = " & gWord.Run("basRwbFrontMatter.InsertFrontMatter", textFile, dn) & " | last error [" & gWord.Run("basRwbFrontMatter.FrontMatterLastError") & "]"
  gOut.WriteLine "sections after " & gDoc.Sections.Count & ", text length " & Len(gDoc.Content.Text)
  gDoc.Repaginate
  For i = 1 To 6
    With gDoc.Sections(i).PageSetup
      gOut.WriteLine "section " & i & " | " & .PageWidth & " x " & .PageHeight & " pt | margins T" & .TopMargin & " B" & .BottomMargin & " L" & .LeftMargin & " R" & .RightMargin & " gutter " & .Gutter & " | hdr/ftr " & .HeaderDistance & "/" & .FooterDistance & " | start " & .SectionStart & " | valign " & .VerticalAlignment & " | first-page-different " & .DifferentFirstPageHeaderFooter & " | cols " & .TextColumns.Count & " | primary footer fields " & gDoc.Sections(i).Footers(1).Range.Fields.Count & " | paragraphs " & gDoc.Sections(i).Range.Paragraphs.Count
    End With
  Next
  For i = 1 To 20
    Set p = gDoc.Paragraphs(i)
    gOut.WriteLine i & " | " & p.Style & " | page " & p.Range.Information(3) & " | y " & FormatNumber(p.Range.Information(6), 1) & " pt | section " & p.Range.Sections(1).Index & " | " & Left(Replace(Replace(p.Range.Text, vbCr, "<CR>"), Chr(12), "<FF>"), 50)
  Next
  gOut.WriteLine "second InsertFrontMatter (expected -1, refuses) = " & gWord.Run("basRwbFrontMatter.InsertFrontMatter", textFile, dn) & " | last error [" & gWord.Run("basRwbFrontMatter.FrontMatterLastError") & "]"
  gDoc.Save
  If Err.Number <> 0 Then Fail "save failed": Exit Sub
  gOut.WriteLine "saved " & target
End Sub

If WScript.Arguments.Count < 4 Then
  WScript.Echo "usage: cscript //nologo insert-front.vbs <src folder> <copy.docm> <front-matter.txt> <result.txt>"
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
