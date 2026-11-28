' selftest-nonbmp.vbs - makes the scratch document for the non-BMP self test of docm-text-edit.vbs (I18N-5, 2026-11-27).
'   cscript //nologo selftest-nonbmp.vbs <new.docx>      then run docm-text-edit.vbs on it with selftest-nonbmp-edits.txt (body); expected: OK B3, OK B1, saved.
' The text holds CJK extension B (U+20000) and an emoji (U+1F600), each a surrogate pair: Len counts two units, and the narrowing of the find range by Len(left) must still land on the target.
Option Explicit
Dim w, d, s
Set w = CreateObject("Word.Application")
w.Visible = False
w.DisplayAlerts = 0
Set d = w.Documents.Add
s = "Plain start " & ChrW(&HD840) & ChrW(&HDC00) & " CJK-B left TARGET right " & ChrW(&HD83D) & ChrW(&HDE00) & " emoji end." & vbCr & "Second paragraph with " & ChrW(&HD83D) & ChrW(&HDE00) & " one TARGET2 here."
d.Content.Text = s
d.SaveAs2 WScript.Arguments(0), 16
d.Close 0
w.Quit 0
