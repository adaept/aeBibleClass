Attribute VB_Name = "basRWBTextExport"
Option Explicit

'Copyright (c) 2025-2026 Peter F. Ennis
'SPDX-License-Identifier: LGPL-3.0-or-later OR LicenseRef-adaept-Commercial
'DUAL-LICENSED. You may use this file under EITHER of:
'  (1) the GNU Lesser General Public License, version 3.0 or (at your option)
'      any later version  -  https://www.gnu.org/licenses/lgpl-3.0.txt ; OR
'  (2) a commercial / proprietary license available from adaept (Peter Ennis),
'      permitting use in closed-source / proprietary works WITHOUT the LGPL
'      copyleft obligations. Contact the copyright holder for commercial terms.
'As the sole copyright holder, adaept may license this file under either option.
'This library is distributed WITHOUT ANY WARRANTY; without even the implied
'warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.

' ============================================================================
' basRWBTextExport
' ----------------------------------------------------------------------------
' Dumps the docm's VerseText paragraphs to a plain-text file in aeRWB's
' rwb.txt/web.txt format ("Book Chapter:Verse<TAB>text", one verse per line),
' so the docm's CURRENT text can be compared against aeRWB/rwb.txt and the
' WEBU reference corpus (aeRWB/tools/web-diff). See
' rvw/Plan_engwebu_baseline_sync_2026-09-14.md ("New goal" / item 11).
'
' "text != text" unless the comparison is explicitly defined (operator,
' 2026-09-14) - see the plan doc's "Text equality is not automatic" section.
' This module's specific contract:
'
'   - Encoding: UTF-8 WITH a leading BOM, matching web.txt/rwb.txt's
'     documented encoding (aeRWB rvw 2026-06-16 SS4). Written via
'     ADODB.Stream (Type=2, Charset="utf-8") - NOT FileSystemObject.
'     CreateTextFile's "Unicode" flag, which writes UTF-16LE, not UTF-8,
'     and would fail every byte-for-byte comparison against rwb.txt.
'   - One verse per line: any manual line break (Chr(11)) or stray
'     paragraph mark found within a paragraph's extracted verse text is
'     collapsed to a single space (NormalizeForSingleLine) - a raw embedded
'     control character would corrupt the one-verse-per-line format for
'     every downstream line-based parser. This collapsing IS a content
'     decision, not a no-op - documented here so it isn't mistaken for a
'     byte-exact capture.
'   - Field separator: a literal Tab between reference and verse text; any
'     stray Tab within the verse text itself is replaced with a space so it
'     can never be mistaken for the field separator.
'   - What is captured verbatim, deliberately NOT normalized: curly vs.
'     straight quotes and the exact quote-nesting codepoints (U+201C/2018/
'     201D/2019) - that's the entire subject of the Test 70/71 comparison
'     this export exists for. basUSFM_Export.CleanTextForUTF8 (a pure string
'     transform, no COM calls - reused directly, not via
'     TryParseChapterVerseFromStyles) is applied to the prose text and was
'     checked line-by-line before reuse: it strips soft hyphens/zero-width
'     characters/control characters below Chr(32) other than tab/CR/LF, and
'     does not touch quote characters - safe for this purpose. Re-check this
'     if CleanTextForUTF8 is ever extended.
'
' Book-name spelling: uses aeBibleCitationClass.GetCanonicalBookTable() (the
' project's existing canonical-name SSOT, e.g. "1 Kings", "Song of Solomon")
' keyed by a running Heading-1 counter through the document's canonical book
' order, rather than attempting to reformat the docm's ALL-CAPS H1 heading
' text - avoids a lossy ad hoc title-casing pass (e.g. "SONG OF SOLOMON" ->
' naive title case would wrongly capitalize "Of").
'
' PERFORMANCE - NO character- or word-level style lookups at all (2026-09-14):
' two earlier versions of this routine both blew Word's memory up into the
' multiple-GB range before completing:
'   v1 called basUSFM_Export's TryParseChapterVerseFromStyles per VerseText
'      paragraph, which extends a Range one CHARACTER at a time via fresh
'      Document.Range(...) objects - tens of thousands of short-lived Range
'      objects across ~31,102 paragraphs. Same class of COM-heavy
'      per-character-property-get cost already documented for
'      GetMarkerTotals / Test 82 (rvw/Code_review 2026-09-13.md item 7).
'   v2 replaced that with basUSFM_Export.ParagraphHasCharStyle/
'      ExtractCharStyleText (word-level, via the paragraph's native .words
'      collection) to read the verse number from the "Verse marker" style -
'      still blew up, AND a `maxVerses` testing safety net gated on
'      successful writes never engaged, because the word-level style match
'      was silently failing (word-splitting across a chapter/verse marker
'      boundary did not behave as assumed), so the loop ran the full
'      document regardless of the limit while iterating every word of every
'      failing verse looking for a match that never came.
' Word's internal range-tracking overhead scales with the TOTAL NUMBER of
' Range/word-iteration objects ever created in a session, not just how many
' are alive at once - so the fix is to not query character OR word style at
' all in the hot path, not just to query it more cheaply.
'
' This version instead uses, per VerseText paragraph, ONLY:
'   - chapNum, already tracked from "Heading 2" paragraph text (paragraph-
'     level Range.Text + pure string parsing - the same mechanism
'     basUSFM_Export.ConvertParagraphToUSFM already uses for chapter
'     tracking).
'   - ONE Range.Text read for the paragraph's own plain text.
'   - Pure VBA string comparison (LeadingDigits, Left$/Mid$) to find where
'     the leading digit run (chapNum immediately followed by the verse
'     number, no separator - e.g. chapter 31 verse 35 -> "3135") splits:
'     since chapNum is already known, checking whether the digit run starts
'     with chapNum's own decimal string and taking the remainder as the
'     verse number needs no style information at all.
' Net COM cost per VerseText paragraph: one .style.NameLocal check, one
' Range.Text read - both paragraph-level, zero character- or word-level
' calls. Matches the cost profile of CaptureHeading1s (a full 33k-paragraph
' scan that has never shown this memory issue).
' A `maxVerses` testing limiter counts paragraphs VISITED, not successful
' writes, so it engages even if every match fails - the exact gap that let
' v2's testing safety net silently not fire.
' ============================================================================

Public Sub ExportDocmVersesToRWBFormat(Optional ByVal outputPath As String, Optional ByVal maxVerses As Long = 0)
    On Error GoTo PROC_ERR

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If outputPath = "" Then
        If Dir(ActiveDocument.Path & "\rpt", vbDirectory) = "" Then
            fso.CreateFolder ActiveDocument.Path & "\rpt"
        End If
        outputPath = ActiveDocument.Path & "\rpt\docm-verses.txt"
    End If

    Dim canonBooks As Object
    Set canonBooks = aeBibleCitationClass.GetCanonicalBookTable()

    Dim oPara As Word.Paragraph
    Dim bookIndex As Long
    Dim bookName As String
    Dim chapNum As Long
    Dim chapStr As String
    Dim verseNum As Long
    Dim paraTxt As String
    Dim digitRun As String
    Dim verseStr As String
    Dim prose As String
    Dim ref As String
    Dim lineCount As Long, skipCount As Long, dupCount As Long, unknownBookCount As Long
    Dim visitedCount As Long
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim StyleName As String

    Dim buf As String
    buf = "RWB-DOCM" & vbLf & _
          "Radiant Word Bible - exported from the live .docm (" & Format(Now, "yyyy-mm-dd") & ")" & vbLf

    bookName = ""
    chapNum = 0
    For Each oPara In ActiveDocument.Paragraphs
        StyleName = oPara.style.NameLocal

        If StyleName = "Heading 1" Then
            bookIndex = bookIndex + 1
            If canonBooks.Exists(bookIndex) Then
                bookName = canonBooks.Item(bookIndex)(1)
            Else
                bookName = "UNKNOWN_BOOK_" & bookIndex
                unknownBookCount = unknownBookCount + 1
            End If
            chapNum = 0

        ElseIf StyleName = "Heading 2" Then
            Dim headTxt As String
            headTxt = oPara.Range.Text
            Dim n As Long
            n = FirstNumberInText(headTxt)
            If n > 0 Then chapNum = n

        ElseIf StyleName = "VerseText" Then
            visitedCount = visitedCount + 1
            ' No character/word-style lookup at all here (that's what drove the
            ' earlier memory blowups - see the module header). chapNum is
            ' already known from Heading 2 tracking; the leading digit run in
            ' the paragraph's own plain text is chapNum immediately followed
            ' by the verse number with no separator (e.g. chapter 31 verse 35
            ' -> "3135") - a pure string comparison against chapNum's own
            ' digit string finds the split, zero COM calls beyond the one
            ' Range.Text already read.
            If bookName = "" Or chapNum = 0 Then
                skipCount = skipCount + 1
            Else
                paraTxt = oPara.Range.Text
                digitRun = LeadingDigits(paraTxt)
                chapStr = CStr(chapNum)
                If Left$(digitRun, Len(chapStr)) = chapStr Then
                    verseStr = Mid$(digitRun, Len(chapStr) + 1)
                    If verseStr <> "" And IsNumeric(verseStr) Then
                        verseNum = CLng(verseStr)
                        ref = bookName & " " & chapNum & ":" & verseNum
                        If seen.Exists(ref) Then
                            dupCount = dupCount + 1
                        Else
                            seen.Add ref, True
                            prose = NormalizeForSingleLine(basUSFM_Export.CleanTextForUTF8(Mid$(paraTxt, Len(digitRun) + 1)))
                            buf = buf & ref & vbTab & prose & vbLf
                            lineCount = lineCount + 1
                        End If
                    Else
                        skipCount = skipCount + 1
                    End If
                Else
                    skipCount = skipCount + 1
                End If
            End If
        End If

        ' Counts paragraphs VISITED, not just successful writes - a limiter
        ' gated on lineCount alone would never fire if matches are failing
        ' (exactly what happened testing the character/word-style version).
        If maxVerses > 0 And visitedCount >= maxVerses Then Exit For
    Next oPara

    WriteUtf8WithBom outputPath, buf

    Debug.Print "ExportDocmVersesToRWBFormat: wrote " & lineCount & " verses to " & outputPath
    Debug.Print "  skipped=" & skipCount & " duplicates=" & dupCount & " unknownBookHeadings=" & unknownBookCount

PROC_EXIT:
    Exit Sub
PROC_ERR:
    Debug.Print "ERROR in basRWBTextExport.ExportDocmVersesToRWBFormat | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    Resume PROC_EXIT
End Sub

' Pure string scan (no COM) - the first run of digit characters anywhere in
' s, as a Long. Used to pull the chapter number out of a "Heading 2"
' paragraph's raw text (e.g. "CHAPTER 5", "5"), whatever its exact wording.
' Returns 0 if no digits are found.
Private Function FirstNumberInText(ByVal s As String) As Long
    Dim i As Long, ch As String, digits As String
    Dim started As Boolean
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
            started = True
        ElseIf started Then
            Exit For
        End If
    Next i
    If digits <> "" Then FirstNumberInText = CLng(digits) Else FirstNumberInText = 0
End Function

' Pure string scan (no COM) - the leading run of digit characters in s (the
' Chapter Verse marker + Verse marker text, always plain digit characters in
' Range.Text regardless of their character style), or "" if s doesn't start
' with a digit. Does NOT touch anything else in the string.
Private Function LeadingDigits(ByVal s As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit For
    Next i
    LeadingDigits = Left$(s, i - 1)
End Function

' Collapses characters that would corrupt the one-verse-per-line TSV format
' into a single space. Deliberately does NOT touch quote characters or any
' other content character - see the module header's "text != text" Note.
Private Function NormalizeForSingleLine(ByVal s As String) As String
    s = Replace(s, Chr$(11), " ")  ' manual line break
    s = Replace(s, vbTab, " ")     ' stray tab - would collide with the field separator
    s = Replace(s, vbCr, " ")      ' stray paragraph mark
    s = Replace(s, vbLf, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeForSingleLine = Trim$(s)
End Function

' UTF-8 WITH a leading BOM, matching web.txt/rwb.txt's documented encoding.
' FileSystemObject.CreateTextFile's "Unicode" flag writes UTF-16LE, not
' UTF-8 - do not use it for this output. ADODB.Stream (Type=2 text,
' Charset="utf-8") writes UTF-8 with a BOM on SaveToFile by default.
Private Sub WriteUtf8WithBom(ByVal path As String, ByVal content As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText content
    stm.saveToFile path, 2 ' adSaveCreateOverWrite
    stm.Close
End Sub
