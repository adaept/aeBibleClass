Attribute VB_Name = "basBibleOnlyExport"
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
' basBibleOnlyExport
' ----------------------------------------------------------------------------
' DRAFT - 2026-10-07. Written against
' adaept5tudio/rvw/Plan_radiant_word_bible_fork_2026-10-06.md (Passes 1-8).
' This is hand-written VBA source that has never been compiled - I cannot
' run the VBE's compiler. Per this codebase's manual-cycle convention
' (see EDSG and [[feedback_docm_manual_vba_import_convention]] in memory),
' the only valid path from here is: import via ImportAllVBAFiles, Debug ->
' Compile Project, fix anything that doesn't compile, export the
' now-compiling project back out (that export - not this file - becomes
' the authoritative source), then run against a DISPOSABLE DUPLICATE of
' the production docm, never the original directly.
'
' Produces a scripture-only fork of the full Study Bible manuscript:
' strips footnotes, front matter, and per-book author commentary; keeps
' the Bible text itself plus the Bible Index (regenerated with corrected
' page numbers). See the plan Doc's "Iteration methodology: break / pause
' / add-test" section - this WILL halt on its first run or two; that is
' expected, not a bug to route around.
'
' Halt discipline: every pass checks a specific condition and, on failure,
' calls HaltExport (sets m_haltFired = True, Debug.Prints the reason) and
' returns immediately. ExportScriptureOnlyDocx checks m_haltFired after
' every pass and stops the whole run without closing or saving the
' working duplicate, so it stays open mid-state for inspection - never
' catch-and-continue, per the plan's explicit methodology.
'
' KEEP / STRIP taxonomy (GetScriptureOnlyStyles / GetIndexRegenStyles /
' GetScriptureStripStyles below) is deliberately NOT sourced from
' basTEST_aeBibleConfig.GetApprovedStyles() - that array is the
' project-wide taxonomy and contains most of the STRIP targets; reusing
' it here would conflate "approved somewhere in aeBibleClass" with
' "survives this fork," which are different questions.
'
' Bible Index handling (Task A/B, 2026-10-07 - see the plan Doc): both
' Bible Index tables (Old Testament, New Testament; 66 BibleIndexList
' rows total) live inside Section 5 of the original docm, alongside
' ~170 OTHER paragraphs of legitimately-STRIP Introduction/Author*/
' Topical-Index content sharing that same section. So this does NOT
' simply keep-or-strip Section 5 wholesale - Pass 3 keeps
' BibleIndexEyebrow/BibleIndex/BibleIndexList by style name, plus (via a
' table-membership rule) any OTHER paragraph that happens to sit inside
' a Table containing at least one BibleIndexList row (that table's own
' structural BodyText spacer paragraphs - BodyText is legitimately STRIP
' everywhere else in the document, so this has to be membership-based,
' not a blanket style exception). Section 5 survives Pass 5's section
' collapse, shrunk down to just the two tables - no new section or
' Table.Add is needed; Pass 6 patches the 66 existing rows' page numbers
' in place.
'
' DocVariable audit (Task C, 2026-10-07): confirmed via direct read-only
' inspection of the live docm's OOXML parts that there is no DocVariable
' usage anywhere in it (zero w:docVar entries, zero DOCVARIABLE field
' codes, across all 160 XML parts). Nothing to migrate or remove here;
' this module's own page-number patching (Pass 6/8) does not read or
' write ActiveDocument.Variables at all.
' ============================================================================

Private m_haltFired As Boolean
Private m_report As String          ' accumulated run log, written to rpt\ at exit
Private m_issues As Long            ' Count of halts/errors/failed post-save steps
Private m_reportFolder As String    ' folder whose rpt\ receives the report

' ==========================================================================
' GetScriptureOnlyStyles
' ==========================================================================
' The actual scripture-text KEEP list - the styles that make this a Bible
' at all. Does not include the Bible Index styles (see
' GetIndexRegenStyles) - those are front matter, not scripture, kept for
' a different reason (regenerated page numbers, not narrative content).
' ==========================================================================
Public Function GetScriptureOnlyStyles() As Variant
    GetScriptureOnlyStyles = Array( _
        "Heading 1", "Heading 2", "VerseText", _
        "Chapter Verse marker", "Verse marker", _
        "Words of Jesus", "EmphasisBlack", "EmphasisRed", _
        "BookHyperlink", "Psalms BOOK", "PsalmSuperscription", _
        "Selah", "PsalmAcrostic", "SpeakerLabel")
End Function

' ==========================================================================
' GetIndexRegenStyles
' ==========================================================================
' Front-matter content that survives Pass 3 not because it's scripture,
' but because Pass 6 regenerates it in place. See the plan Doc's Task
' A/B (2026-10-07): both Bible Index tables live in Section 5 alongside
' ~170 paragraphs of legitimately-STRIP content; these three style names
' are the only Section-5 survivors by name - everything else in that
' section strips normally. The tables' own structural BodyText spacer
' paragraphs are handled separately, by table membership, not by name
' (BodyText is legitimately STRIP everywhere else) - see
' BuildIndexTableRanges / Pass3_ParagraphSweep.
' ==========================================================================
Public Function GetIndexRegenStyles() As Variant
    GetIndexRegenStyles = Array("BibleIndexEyebrow", "BibleIndex", "BibleIndexList")
End Function

' ==========================================================================
' GetScriptureStripStyles
' ==========================================================================
' Explicit STRIP list, matching the plan Doc's scope item 1 verbatim
' (minus BibleIndexEyebrow/BibleIndex/BibleIndexList, moved to
' GetIndexRegenStyles per Task A/B). Deliberately an explicit list, not
' "everything not in KEEP" - Pass 3 halts on any paragraph style in
' neither this list nor the KEEP lists, which is the point: an
' unclassified style (drift in the taxonomy, or a genuine new defect -
' see the "Plain Text" precedent that blocked this plan before Pass 1
' could even be written) must stop the run, not get silently deleted or
' silently kept.
' ==========================================================================
Public Function GetScriptureStripStyles() As Variant
    GetScriptureStripStyles = Array( _
        "FrontPageTopLine", "TitleEyebrow", "Title", "TitleVersion", "FrontPageBodyText", _
        "BodyTextTopLineCPBB", "Acknowledgments", "AuthorBodyText", "Contents", "ContentsRef", _
        "Introduction", "TitleOnePage", _
        "AuthorListItem", "AuthorListItemBody", "AuthorListItemTab", _
        "AuthorBookRefHeader", "AuthorBookRef", "AuthorBookSections", "AuthorSectionHead", _
        "CenterSubText", "CustomParaAfterH1", "DatAuthRef", "Brief", "BodyText", _
        "Footnote Text", "Footnote Reference")
End Function

' ==========================================================================
' ExportScriptureOnlyDocx
' ==========================================================================
' Orchestrates Passes 1-8 end to end, in auto mode (no prompts). Halts
' immediately at the first failed check, leaving the working duplicate
' open and unsaved for inspection - see the plan Doc's break/pause/
' add-test methodology. Never run against the production docm directly;
' always against a disposable duplicate (sourcePath defaults to
' ActiveDocument.FullName - make that duplicate the active document
' before running this with no arguments). This Sub then makes its OWN
' internal working copy of whatever sourcePath is, so even that
' operator-made duplicate is never touched - only the internal working
' copy is mutated.
'
' Usage:
'   ExportScriptureOnlyDocx                    ' uses ActiveDocument as source
'   ExportScriptureOnlyDocx "C:\...\dup.docm"   ' explicit source
' ==========================================================================
Public Sub ExportScriptureOnlyDocx(Optional ByVal sourcePath As String, _
                                    Optional ByVal workingPath As String, _
                                    Optional ByVal destPath As String)
    On Error GoTo PROC_ERR

    m_haltFired = False
    m_report = ""
    m_issues = 0
    m_reportFolder = ""

    If sourcePath = "" Then sourcePath = ActiveDocument.FullName
    Dim SourceFolder As String
    SourceFolder = Left$(sourcePath, InStrRev(sourcePath, "\"))
    If workingPath = "" Then workingPath = SourceFolder & "RadiantWordBible_work.docm"
    If destPath = "" Then destPath = SourceFolder & "RadiantWordBible.docx"
    m_reportFolder = Left$(SourceFolder, Len(SourceFolder) - 1)

    Dim oSettings As Object
    Set oSettings = GetExportSettings()
    ExportLog "---- RadiantWordBible export " & Format(Now, "yyyy-mm-dd hh:nn:ss") & _
        " | ExportVersion " & oSettings("ExportVersion") & " ----"
    Dim sKey As Variant
    For Each sKey In oSettings.Keys
        ExportLog "  setting " & sKey & " = " & oSettings(sKey)
    Next sKey

    ExportLog "ExportScriptureOnlyDocx: source=" & sourcePath
    ExportLog "  working copy=" & workingPath
    ExportLog "  destination=" & destPath

    ' Pass 1 - duplicate and open. All following passes operate on oDoc
    ' only; sourcePath is never touched.
    If Dir(workingPath) <> "" Then
        On Error Resume Next
        Kill workingPath
        On Error GoTo PROC_ERR
        If Dir(workingPath) <> "" Then
            ExportLog "ExportScriptureOnlyDocx: could not remove stale working copy at " & _
                workingPath & " - likely still open in Word from a previous halted run. " & _
                "Close it first, inspect/fix per the manual cycle, then re-run."
            GoTo PROC_EXIT
        End If
    End If

    ' Word's Document object has no SaveCopyAs method at all (that's an
    ' Excel/PowerPoint-only method - calling it on a Word Document raises
    ' "That method is not available on that object", Err 5892, confirmed
    ' 2026-10-07). VBA's native FileCopy statement is documented to refuse
    ' to copy a file that is currently open (the real cause of the earlier
    ' Err 70 "Permission denied" - sourcePath is the active document this
    ' very macro is running from). Scripting.FileSystemObject.CopyFile
    ' goes through the OS copy API directly and tolerates copying a file
    ' Word has open under its normal (non-exclusive) share mode - use that
    ' unconditionally instead of branching on whether the source is open.
    Dim oFSO As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")

    Dim diagErrNum As Long
    Dim diagErrDesc As String

    On Error Resume Next
    oFSO.CopyFile sourcePath, workingPath, True
    diagErrNum = Err.Number
    diagErrDesc = Err.Description
    Err.Clear
    On Error GoTo PROC_ERR
    If diagErrNum <> 0 Then
        ExportLog "ExportScriptureOnlyDocx: FileSystemObject.CopyFile failed - Err " & _
            diagErrNum & ": " & diagErrDesc
        GoTo PROC_EXIT
    End If

    Dim oDoc As Object
    On Error Resume Next
    Set oDoc = Documents.Open(workingPath)
    diagErrNum = Err.Number
    diagErrDesc = Err.Description
    Err.Clear
    On Error GoTo PROC_ERR
    If diagErrNum <> 0 Or oDoc Is Nothing Then
        ExportLog "ExportScriptureOnlyDocx: Documents.Open failed - Err " & diagErrNum & _
            ": " & diagErrDesc
        GoTo PROC_EXIT
    End If
    ExportLog "Pass1_DuplicateAndOpen: opened working copy."

    Dim screenWas As Boolean
    screenWas = Application.ScreenUpdating
    Application.ScreenUpdating = False

    ' Pass 2 - footnote sweep
    Pass2_FootnoteSweep oDoc
    If m_haltFired Then GoTo PROC_HALT

    ' Pass 3 - paragraph sweep (captures the pre-deletion section Count,
    ' used by Pass 5 to confirm the expected section collapse happened)
    Dim preDeleteSectionCount As Long
    preDeleteSectionCount = oDoc.Sections.Count
    Pass3_ParagraphSweep oDoc
    If m_haltFired Then GoTo PROC_HALT

    ' Pass 4 - verify strip (intermediate state)
    Dim violations As Long
    violations = VerifyScriptureOnlyStrip(oDoc)
    If violations <> 0 Then
        HaltExport "Pass4_VerifyScriptureOnlyStrip", _
            violations & " violation(s) found - see rpt\VerifyScriptureOnlyStrip.txt"
        GoTo PROC_HALT
    End If

    ' Pass 4b - settings-driven cleanup (v1.0: manual hyphen removal)
    If oSettings("StripOptionalHyphens") Then
        Dim hyphensLeft As Long
        hyphensLeft = Pass4b_StripOptionalHyphens(oDoc)
        If hyphensLeft <> 0 Then
            HaltExport "Pass4b_StripOptionalHyphens", _
                hyphensLeft & " optional hyphen(s) remain after removal"
            GoTo PROC_HALT
        End If
    End If

    ' Pass 4c - emphasis character styles (B2/B3): EmphasisBlack removed so
    ' the text takes VerseText, EmphasisRed -> Words of Jesus
    If oSettings("ReplaceEmphasisStyles") Then
        Dim emphViolations As Long
        emphViolations = Pass4c_ReplaceEmphasisStyles(oDoc)
        If emphViolations <> 0 Then
            HaltExport "Pass4c_ReplaceEmphasisStyles", _
                emphViolations & " violation(s) - see Immediate window"
            GoTo PROC_HALT
        End If
    End If

    ' Pass 4d - VerseText alignment (v1.0: left; the docm is Justified)
    If oSettings("VerseTextAlignment") <> "" Then
        Dim alignViolations As Long
        alignViolations = Pass4d_SetVerseTextAlignment(oDoc, CStr(oSettings("VerseTextAlignment")))
        If alignViolations <> 0 Then
            HaltExport "Pass4d_SetVerseTextAlignment", _
                alignViolations & " violation(s) - see the report"
            GoTo PROC_HALT
        End If
    End If

    ' Pass 5 - section surgery
    Pass5_SectionSurgery oDoc, preDeleteSectionCount
    If m_haltFired Then GoTo PROC_HALT

    ' Pass 6 - Bible Index page-number regeneration
    Pass6_RegenerateBibleIndex oDoc
    If m_haltFired Then GoTo PROC_HALT

    ' Pass 7 - save
    Application.ScreenUpdating = screenWas
    oDoc.SaveAs2 fileName:=destPath, FileFormat:=wdFormatXMLDocument
    ExportLog "Pass7_Save: saved " & destPath
    oDoc.Close SaveChanges:=False
    Set oDoc = Nothing

    ' Pass 7b - strip the orphaned ribbon part (py\strip_ribbon.py via WSL).
    ' The document is closed in Word at this point, as the script requires.
    If oSettings("StripRibbonAfterSave") Then
        Dim ribbonRc As Long
        ribbonRc = RunPythonStep("Pass7b_StripRibbon", SourceFolder & "py\strip_ribbon.py", _
            Array(destPath))
        If ribbonRc <> 0 Then
            m_issues = m_issues + 1
            ExportLog "Pass7b_StripRibbon: FAILED (exit code " & ribbonRc & "). " & _
                "Artefact saved but still carries the ribbon part."
        End If
    End If

    ' Pass 8 - final verification, fresh open, independent of Pass 6's
    ' in-memory state
    Dim mismatches As Long
    mismatches = VerifyBibleIndexPageNumbers(destPath)
    If mismatches <> 0 Then
        m_issues = m_issues + 1
        ExportLog "Pass8: " & mismatches & " mismatch(es) - see rpt\VerifyBibleIndexPageNumbers.txt."
    Else
        ExportLog "Pass8: Bible Index page numbers verified, 0 mismatches."
    End If

    ' Pass 9 - character-style change verifier vs. the baseline .docx
    ' (py\verify_char_style_change.py): nothing but the two emphasis
    ' styles may have changed.
    If oSettings("ReplaceEmphasisStyles") Then
        Dim baselinePath As String
        baselinePath = SourceFolder & oSettings("BaselineDocxRelPath")
        If Dir(baselinePath) = "" Then
            ExportLog "Pass9_VerifyCharStyles: SKIPPED - no baseline at " & baselinePath
        Else
            Dim verifyRc As Long
            If oSettings("VerseTextAlignment") <> "" Then
                verifyRc = RunPythonStep("Pass9_VerifyCharStyles", _
                    SourceFolder & "py\verify_char_style_change.py", _
                    Array(baselinePath, destPath, "--verse-align", oSettings("VerseTextAlignment")))
            Else
                verifyRc = RunPythonStep("Pass9_VerifyCharStyles", _
                    SourceFolder & "py\verify_char_style_change.py", Array(baselinePath, destPath))
            End If
            If verifyRc <> 0 Then m_issues = m_issues + 1
        End If
    End If

    If m_issues = 0 Then
        ExportLog "ExportScriptureOnlyDocx: COMPLETE. " & destPath & " - all checks passed."
    Else
        ExportLog "ExportScriptureOnlyDocx: FINISHED WITH " & m_issues & " ISSUE(S). " & _
            "Artefact saved but NOT clean - do not treat as done. See the report."
    End If

    GoTo PROC_EXIT

PROC_HALT:
    Application.ScreenUpdating = screenWas
    ExportLog "ExportScriptureOnlyDocx: HALTED. Working copy left open and unsaved " & _
        "for inspection: " & workingPath
    m_haltFired = False

PROC_EXIT:
    WriteExportReport
    Exit Sub
PROC_ERR:
    Application.ScreenUpdating = True
    m_issues = m_issues + 1
    ExportLog "ERROR in basBibleOnlyExport.ExportScriptureOnlyDocx | Erl: " & Erl & _
        " | Err: " & Err.Number & " | " & Err.Description
    ExportLog "  Working copy (if open) left as-is for inspection: " & workingPath
    Resume PROC_EXIT
End Sub

' ==========================================================================
' Pass2_FootnoteSweep
' ==========================================================================
' Reverse-iterate Footnotes and delete each - matches this codebase's
' existing reverse-delete precedent (Hyperlinks.Count To 1 Step -1 in
' basTEST_aeBibleTools.bas). Removes the reference mark and footnote-story
' text together.
'
' Halt condition: Footnotes.Count <> 0 after the loop.
' ==========================================================================
Private Sub Pass2_FootnoteSweep(ByVal oDoc As Object)
    Dim i As Long
    For i = oDoc.Footnotes.Count To 1 Step -1
        oDoc.Footnotes(i).Delete
    Next i

    If oDoc.Footnotes.Count <> 0 Then
        HaltExport "Pass2_FootnoteSweep", _
            "Footnotes.Count = " & oDoc.Footnotes.Count & " after sweep, expected 0."
        Exit Sub
    End If

    ExportLog "Pass2_FootnoteSweep: all footnotes removed."
End Sub

' ==========================================================================
' Pass3_ParagraphSweep
' ==========================================================================
' Single forward walk of ActiveDocument.Paragraphs (main body only -
' matches the "proven fast, seconds" idiom documented in
' basVerseStructureAudit.GetMarkerTotals, not the per-style Range.Find
' approach documented elsewhere as 300-2700s on this document's size).
'
' Classification per paragraph, in order:
'   1. KEEP (GetScriptureOnlyStyles + GetIndexRegenStyles) - never deleted.
'   2. Index-table membership (BuildIndexTableRanges) - a paragraph of ANY
'      style sitting inside a Table that contains at least one
'      BibleIndexList row is preserved. This is what correctly keeps each
'      Index table's own structural BodyText spacer paragraphs without
'      blanket-keeping BodyText, which is legitimately STRIP everywhere
'      else in the document (Task A/B, 2026-10-07).
'   3. STRIP (GetScriptureStripStyles) - marked for deletion.
'   4. Anything else - HALT. An unclassified style is the single most
'      likely place for the taxonomy to drift out from under this tool;
'      the cheapest point to catch it is here, not after deletion.
'
' Deletion is a SECOND, reverse-order pass. deleteStarts/deleteEnds are
' collected during the forward walk, so they are already in ascending
' document order - no sort needed, just iterate backwards (highest Start
' first) so earlier offsets never go stale from a later deletion.
'
' Halt condition: any unclassified paragraph style.
' ==========================================================================
Private Sub Pass3_ParagraphSweep(ByVal oDoc As Object)
    Dim keepDict As Object
    Set keepDict = BuildKeepDict()
    Dim stripDict As Object
    Set stripDict = BuildStripDict()

    Dim tblStarts() As Long
    Dim tblEnds() As Long
    Dim nTbl As Long
    BuildIndexTableRanges oDoc, tblStarts, tblEnds, nTbl

    Dim cap As Long
    cap = oDoc.Paragraphs.Count
    Dim deleteStarts() As Long
    Dim deleteEnds() As Long
    ReDim deleteStarts(1 To cap)
    ReDim deleteEnds(1 To cap)
    Dim nDelete As Long
    nDelete = 0

    Dim oPara As Object
    Dim StyleName As String
    Dim k As Long
    Dim inIndexTable As Boolean

    For Each oPara In oDoc.Paragraphs
        StyleName = oPara.style.NameLocal

        If keepDict.Exists(StyleName) Then
            ' KEEP - scripture style, or Bible-Index-regeneration style.
        Else
            inIndexTable = False
            For k = 1 To nTbl
                If oPara.Range.Start >= tblStarts(k) And oPara.Range.Start < tblEnds(k) Then
                    inIndexTable = True
                    Exit For
                End If
            Next k

            If inIndexTable Then
                ' Table-membership rule - preserved regardless of style name.
            ElseIf stripDict.Exists(StyleName) Then
                nDelete = nDelete + 1
                deleteStarts(nDelete) = oPara.Range.Start
                deleteEnds(nDelete) = oPara.Range.End
            Else
                HaltExport "Pass3_ParagraphSweep", _
                    "Unclassified paragraph style """ & StyleName & """ at Range.Start=" & _
                    oPara.Range.Start & " - not in the KEEP, STRIP, or Index-table " & _
                    "taxonomy. Excerpt: """ & _
                    Left$(Replace(oPara.Range.Text, vbCr, ""), 80) & """"
                Exit Sub
            End If
        End If
    Next oPara

    Dim i As Long
    Dim oRng As Object
    Dim nCarriers As Long
    For i = nDelete To 1 Step -1
        Set oRng = oDoc.Range(deleteStarts(i), deleteEnds(i))
        ' A paragraph that carries a section break ends in Chr(12), not a
        ' paragraph mark. Deleting it would merge the section into its
        ' neighbour (v1.0 review: 144 of 145 breaks sit in BodyText
        ' paragraphs). Delete the text only and keep the break; Pass 5
        ' removes the 12 STRIP-only sections explicitly.
        If Right$(oRng.Text, 1) = Chr(12) Then
            nCarriers = nCarriers + 1
            oRng.End = oRng.End - 1
            If oRng.End > oRng.Start Then oRng.Delete
        Else
            oRng.Delete
        End If
    Next i

    ' Word never lets the final paragraph mark be deleted, so an empty STRIP-style
    ' final paragraph survives the sweep. It is not scripture: give it BodyText,
    ' the default outside scripture (operator decision, 2026-10-07).
    Dim oLast As Object
    Set oLast = oDoc.Paragraphs.Last
    If oLast.Range.Text = vbCr Then
        If stripDict.Exists(oLast.style.NameLocal) And oLast.style.NameLocal <> "BodyText" Then
            oLast.style = oDoc.Styles("BodyText")
        End If
    End If

    ExportLog "Pass3_ParagraphSweep: deleted " & nDelete & " STRIP-style paragraph(s) (" & _
        nCarriers & " section-break carriers kept as break-only); " & _
        nTbl & " Index table(s) and their contents preserved."
End Sub

' ==========================================================================
' VerifyScriptureOnlyStrip  (Pass 4)
' ==========================================================================
' Walks every StoryRanges entry via its NextStoryRange chain (same shape
' as aeBibleClass.CountAuditStyles_ToFile, which is Private to that class
' so its walk pattern is copied here, not called directly). Asserts:
'   - Count is 0 for every STRIP style, across all story ranges.
'   - Footnotes.Count = 0.
'   - BibleIndexList Count still equals the full canonical book Count -
'     confirms Pass 3's table-membership rule preserved every row, not
'     just most of them.
'
' Returns the violation Count (0 = clean). Writes
' rpt\VerifyScriptureOnlyStrip.txt. Public so it can also be run
' standalone from the Immediate window against an already-swept document.
' ==========================================================================
Public Function VerifyScriptureOnlyStrip(ByVal oDoc As Object) As Long
    On Error GoTo PROC_ERR

    Dim stripDict As Object
    Set stripDict = BuildStripDict()

    ' Same table-membership exemption as Pass3_ParagraphSweep: the Bible Index
    ' tables' structural spacer paragraphs may legitimately keep a STRIP style.
    Dim tblStarts() As Long
    Dim tblEnds() As Long
    Dim nTbl As Long
    Dim k As Long
    Dim inIndexTable As Boolean
    BuildIndexTableRanges oDoc, tblStarts, tblEnds, nTbl

    Dim violations As Long
    Dim rng As Object
    Dim para As Object
    Dim StyleName As String
    Dim sOut As String
    Const NL As String = vbCrLf
    sOut = "---- VerifyScriptureOnlyStrip: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ----" & NL & NL

    For Each rng In oDoc.StoryRanges
        Do
            For Each para In rng.Paragraphs
                StyleName = para.style.NameLocal
                inIndexTable = False
                If rng.StoryType = wdMainTextStory Then
                    For k = 1 To nTbl
                        If para.Range.Start >= tblStarts(k) And para.Range.Start < tblEnds(k) Then
                            inIndexTable = True
                            Exit For
                        End If
                    Next k
                End If
                ' The final empty paragraph mark cannot be deleted; Pass 3 sets it to
                ' BodyText, so exempt exactly that one paragraph here.
                If rng.StoryType = wdMainTextStory And StyleName = "BodyText" Then
                    If para.Range.End = oDoc.Content.End And para.Range.Text = vbCr Then inIndexTable = True
                End If
                ' Break-only carrier paragraphs (Pass 3 keeps the section break) are
                ' exempt; Pass 5 removes the ones in STRIP-only sections.
                If para.Range.End - para.Range.Start = 1 Then
                    If Right$(para.Range.Text, 1) = Chr(12) Then inIndexTable = True
                End If
                If stripDict.Exists(StyleName) And Not inIndexTable Then
                    violations = violations + 1
                    sOut = sOut & "STRIP-style survivor: """ & StyleName & """ at Range.Start=" & _
                        para.Range.Start & " | Excerpt: """ & _
                        Left$(Replace(para.Range.Text, vbCr, ""), 80) & """" & NL
                End If
            Next para
            Set rng = rng.NextStoryRange
        Loop Until rng Is Nothing
    Next rng

    If oDoc.Footnotes.Count <> 0 Then
        violations = violations + 1
        sOut = sOut & "Footnotes.Count = " & oDoc.Footnotes.Count & ", expected 0." & NL
    End If

    Dim expectedBooks As Long
    expectedBooks = aeBibleCitationClass.GetCanonicalBookTable().Count
    Dim idxCount As Long
    idxCount = 0
    Dim oPara As Object
    For Each oPara In oDoc.Paragraphs
        If oPara.style.NameLocal = "BibleIndexList" Then idxCount = idxCount + 1
    Next oPara
    If idxCount <> expectedBooks Then
        violations = violations + 1
        sOut = sOut & "BibleIndexList Count = " & idxCount & ", expected " & expectedBooks & "." & NL
    End If

    sOut = sOut & NL & "TOTAL violations: " & violations & NL
    ExportLog sOut
    WriteReportFileTo oDoc.Path, "VerifyScriptureOnlyStrip.txt", sOut
    VerifyScriptureOnlyStrip = violations

PROC_EXIT:
    Exit Function
PROC_ERR:
    ExportLog "ERROR in basBibleOnlyExport.VerifyScriptureOnlyStrip | Erl: " & Erl & _
        " | Err: " & Err.Number & " | " & Err.Description
    Resume PROC_EXIT
End Function

' ==========================================================================
' Pass5_SectionSurgery
' ==========================================================================
' Confirmed by code search: nothing in this codebase deletes or merges a
' Section today - this is genuinely new, unproven mechanics (see the plan
' Doc). REVISED 2026-10-07: the original design relied on Pass 3 collapsing
' the STRIP-only sections as a side effect. That was wrong - 144 of the 145
' section breaks sit in BodyText (STRIP) paragraphs, so Pass 3 collapsed
' 145 sections to 1. Pass 3 now keeps every section break (break-only
' paragraphs) and this pass deletes the 12 STRIP-only sections explicitly.
'
' Per Task A/B (2026-10-07): the removable blocks are sections 1-4 (front
' matter MINUS Section 5, which survives because it also hosts the Bible
' Index tables), 84-85 (Historical Parallels Chart), and 140-145 (back
' matter after Revelation) - 4 + 2 + 6 = 12 sections removed, not the
' originally-assumed 13.
'
' Halt condition: post-Pass-3 section Count does not equal
' preDeleteSectionCount - EXPECTED_SECTIONS_REMOVED.
'
' Reassertion is deliberately minimal - only clears a LinkToPrevious
' reference on the new Section(1) that would otherwise dangle (its
' original "previous" section no longer exists). Deeper front-matter
' page-numbering cosmetics are out of scope here; the plan's own "manual
' visual QA pass" step covers that, and Pass 8's verification does not
' depend on where page-number restarts happen - it compares whatever
' Range.Information(wdActiveEndPageNumber) actually reports against
' whatever Pass 6 wrote from that same call, so the comparison holds
' regardless of the restart scheme in effect.
' ==========================================================================
Private Sub Pass5_SectionSurgery(ByVal oDoc As Object, ByVal preDeleteSectionCount As Long)
    ' Pass 3 now keeps every section break (break-only carrier paragraphs),
    ' so the Count must be unchanged. The 12 STRIP-only sections are then
    ' removed explicitly, last block first so earlier indexes stay valid.
    ' Section numbers come from the Pass 0b map of this docm (145 sections).
    Const EXPECTED_SECTIONS As Long = 145
    Const EXPECTED_SECTIONS_REMOVED As Long = 12

    If preDeleteSectionCount <> EXPECTED_SECTIONS Or oDoc.Sections.Count <> EXPECTED_SECTIONS Then
        HaltExport "Pass5_SectionSurgery", _
            "Section Count is " & preDeleteSectionCount & " before / " & oDoc.Sections.Count & _
            " after Pass 3, expected " & EXPECTED_SECTIONS & " both times. The section map " & _
            "(1-4, 84-85, 140-145) no longer applies - re-map before proceeding."
        Exit Sub
    End If

    ' Safety: none of the target sections may hold a KEEP-style paragraph.
    Dim keepDict As Object
    Set keepDict = BuildKeepDict()
    Dim blkFirst As Variant
    Dim blkLast As Variant
    blkFirst = Array(1, 84, 140)
    blkLast = Array(4, 85, 145)
    Dim b As Long
    Dim s As Long
    Dim oPara As Object
    For b = 0 To 2
        For s = blkFirst(b) To blkLast(b)
            For Each oPara In oDoc.Sections(s).Range.Paragraphs
                If keepDict.Exists(oPara.style.NameLocal) Then
                    HaltExport "Pass5_SectionSurgery", _
                        "Section " & s & " (slated for removal) holds a KEEP-style paragraph """ & _
                        oPara.style.NameLocal & """ at Range.Start=" & oPara.Range.Start
                    Exit Sub
                End If
            Next oPara
        Next s
    Next b

    ' Back matter 140-145. The document's last paragraph mark and final
    ' sectPr cannot be deleted, so delete from the section-139 break
    ' through the end; 139 then takes the final section's setup, which is
    ' first overwritten with 139's own page setup.
    CopySectionPageSetup oDoc.Sections(139), oDoc.Sections(145)
    Dim oRng As Object
    Set oRng = oDoc.Range(oDoc.Sections(139).Range.End - 1, oDoc.Content.End - 1)
    oRng.Delete
    ' Historical Parallels Chart 84-85
    Set oRng = oDoc.Range(oDoc.Sections(84).Range.Start, oDoc.Sections(85).Range.End)
    oRng.Delete
    ' Front matter 1-4 (Section 5 survives: it hosts the Bible Index tables)
    Set oRng = oDoc.Range(oDoc.Sections(1).Range.Start, oDoc.Sections(4).Range.End)
    oRng.Delete

    Dim postCount As Long
    postCount = oDoc.Sections.Count
    If preDeleteSectionCount - postCount <> EXPECTED_SECTIONS_REMOVED Then
        HaltExport "Pass5_SectionSurgery", _
            "Section Count went " & preDeleteSectionCount & " -> " & postCount & _
            ", expected exactly " & EXPECTED_SECTIONS_REMOVED & " removed. Inspect the duplicate."
        Exit Sub
    End If

    ExportLog "Pass5_SectionSurgery: " & EXPECTED_SECTIONS_REMOVED & " sections removed (" & _
        preDeleteSectionCount & " -> " & postCount & ")."
End Sub

' ==========================================================================
' CopySectionPageSetup
' ==========================================================================
' Copies the page-setup properties that matter here from one Section to
' another. Headers and footers are NOT copied (plan task 4).
' ==========================================================================
Private Sub CopySectionPageSetup(ByVal src As Object, ByVal dst As Object)
    Dim ps As Object
    Dim pd As Object
    Set ps = src.PageSetup
    Set pd = dst.PageSetup
    pd.Orientation = ps.Orientation
    pd.pageWidth = ps.pageWidth
    pd.PageHeight = ps.PageHeight
    pd.TopMargin = ps.TopMargin
    pd.BottomMargin = ps.BottomMargin
    pd.leftMargin = ps.leftMargin
    pd.rightMargin = ps.rightMargin
    pd.gutter = ps.gutter
    pd.HeaderDistance = ps.HeaderDistance
    pd.FooterDistance = ps.FooterDistance
    pd.SectionStart = ps.SectionStart
    pd.DifferentFirstPageHeaderFooter = ps.DifferentFirstPageHeaderFooter
    pd.OddAndEvenPagesHeaderFooter = ps.OddAndEvenPagesHeaderFooter
    pd.TextColumns.SetCount ps.TextColumns.Count
    pd.TextColumns.EvenlySpaced = ps.TextColumns.EvenlySpaced
    pd.TextColumns.LineBetween = ps.TextColumns.LineBetween
    If ps.TextColumns.EvenlySpaced Then pd.TextColumns.Spacing = ps.TextColumns.Spacing
End Sub

' ==========================================================================
' Pass6_RegenerateBibleIndex
' ==========================================================================
' Self-verifying loop (up to 3 iterations), since there is no
' Repaginate-before-reading precedent anywhere in this codebase to lean
' on: Repaginate, walk Heading 1 (expect exactly the canonical book
' Count) and BibleIndexList (same Count, same positional order per Pass
' 0's confirmed finding that BibleIndexList order matches
' GetCanonicalBookTable order - no name-matching needed), patch every
' row's trailing page-number text to its book's real page, then re-walk
' Heading 1 once more. If any page number changed from the previous
' iteration's reading, repeat; otherwise stable.
'
' Patching happens on every iteration, including the first - a stale or
' placeholder value (e.g. the historical "xxxxxxx" DocVariable-testing
' marker - see the plan Doc's Task A/B correction) must never survive
' untouched into the regenerated artefact.
'
' Halt conditions: Heading 1 Count, or BibleIndexList Count, does not
' equal the canonical book Count; or the patch loop fails to stabilize
' within 3 iterations.
' ==========================================================================
Private Sub Pass6_RegenerateBibleIndex(ByVal oDoc As Object)
    On Error GoTo PROC_ERR

    Dim expectedBooks As Long
    expectedBooks = aeBibleCitationClass.GetCanonicalBookTable().Count

    Dim prevH1Pages() As Long
    ReDim prevH1Pages(1 To expectedBooks)

    Dim iteration As Long
    Dim stable As Boolean
    stable = False

    For iteration = 1 To 3
        oDoc.Repaginate

        Dim h1Paras() As Object
        Dim idxParas() As Object
        ReDim h1Paras(1 To expectedBooks)
        ReDim idxParas(1 To expectedBooks)
        Dim nH1 As Long
        Dim nIdx As Long
        nH1 = 0
        nIdx = 0

        Dim oPara As Object
        For Each oPara In oDoc.Paragraphs
            Select Case oPara.style.NameLocal
                Case "Heading 1"
                    nH1 = nH1 + 1
                    If nH1 <= expectedBooks Then Set h1Paras(nH1) = oPara
                Case "BibleIndexList"
                    nIdx = nIdx + 1
                    If nIdx <= expectedBooks Then Set idxParas(nIdx) = oPara
            End Select
        Next oPara

        If nH1 <> expectedBooks Then
            HaltExport "Pass6_RegenerateBibleIndex", _
                "Heading 1 Count = " & nH1 & ", expected " & expectedBooks & _
                " (canonical book Count)."
            Exit Sub
        End If
        If nIdx <> expectedBooks Then
            HaltExport "Pass6_RegenerateBibleIndex", _
                "BibleIndexList Count = " & nIdx & ", expected " & expectedBooks & "."
            Exit Sub
        End If

        Dim h1Pages() As Long
        ReDim h1Pages(1 To expectedBooks)
        Dim i As Long
        Dim changed As Boolean
        changed = False
        For i = 1 To expectedBooks
            h1Pages(i) = h1Paras(i).Range.Information(wdActiveEndPageNumber)
            If iteration > 1 Then
                If h1Pages(i) <> prevH1Pages(i) Then changed = True
            End If
        Next i

        For i = 1 To expectedBooks
            PatchIndexRowPageNumber idxParas(i), h1Pages(i)
            If m_haltFired Then Exit Sub
        Next i

        If iteration > 1 And Not changed Then
            stable = True
            Exit For
        End If

        For i = 1 To expectedBooks
            prevH1Pages(i) = h1Pages(i)
        Next i
    Next iteration

    If Not stable Then
        HaltExport "Pass6_RegenerateBibleIndex", _
            "Page numbers did not stabilize within 3 Repaginate/patch iterations."
        Exit Sub
    End If

    ExportLog "Pass6_RegenerateBibleIndex: " & expectedBooks & _
        " Index row(s) patched, stable after " & iteration & " iteration(s)."

PROC_EXIT:
    Exit Sub
PROC_ERR:
    ExportLog "ERROR in basBibleOnlyExport.Pass6_RegenerateBibleIndex | Erl: " & Erl & _
        " | Err: " & Err.Number & " | " & Err.Description
    HaltExport "Pass6_RegenerateBibleIndex", "Unhandled error " & Err.Number & ": " & Err.Description
    Resume PROC_EXIT
End Sub

' --------------------------------------------------------------------------
' PatchIndexRowPageNumber - replace a BibleIndexList paragraph's trailing
' page-number text (everything after its last Tab, before the terminating
' paragraph mark) with pageNum. Leaves the leading "Book (Abbrev)" text
' and the dot-leader tab stop itself untouched - only the tab's own
' trailing content changes. This also normalizes away any pre-existing
' "page " prefix inconsistency in the raw text (some rows have it, some
' don't) to a bare number, as a side effect of a clean regeneration.
' --------------------------------------------------------------------------
Private Sub PatchIndexRowPageNumber(ByVal oPara As Object, ByVal pageNum As Long)
    Dim fullText As String
    fullText = oPara.Range.Text

    Dim tabPos As Long
    tabPos = InStrRev(fullText, vbTab)
    If tabPos = 0 Then
        HaltExport "Pass6_RegenerateBibleIndex", _
            "BibleIndexList row at Range.Start=" & oPara.Range.Start & _
            " has no tab character to anchor the page-number replacement. Text: """ & _
            Replace(fullText, vbCr, "") & """"
        Exit Sub
    End If

    Dim oRepl As Object
    Set oRepl = oPara.Range.Document.Range(oPara.Range.Start + tabPos, oPara.Range.End - 1)
    oRepl.Text = CStr(pageNum)
End Sub

' ==========================================================================
' VerifyBibleIndexPageNumbers  (Pass 8)
' ==========================================================================
' Opens the saved artefact FRESH (read-only, not added to recent files) -
' deliberately independent of Pass 6's in-memory state, per the plan.
' Walks Heading 1 and BibleIndexList in document order (same positional
' correspondence as Pass 6), compares each pair's page number. Reports
' mismatches by book name. Returns the mismatch Count (0 = clean).
' Writes rpt\VerifyBibleIndexPageNumbers.txt next to the opened file.
' ==========================================================================
Public Function VerifyBibleIndexPageNumbers(ByVal docPath As String) As Long
    On Error GoTo PROC_ERR

    Dim oDoc As Object
    Set oDoc = Documents.Open(fileName:=docPath, ReadOnly:=True, AddToRecentFiles:=False)
    oDoc.Repaginate

    Dim canonBooks As Object
    Set canonBooks = aeBibleCitationClass.GetCanonicalBookTable()
    Dim expectedBooks As Long
    expectedBooks = canonBooks.Count

    Dim h1Pages() As Long
    Dim idxPages() As Long
    ReDim h1Pages(1 To expectedBooks)
    ReDim idxPages(1 To expectedBooks)
    Dim nH1 As Long
    Dim nIdx As Long
    nH1 = 0
    nIdx = 0

    Dim oPara As Object
    For Each oPara In oDoc.Paragraphs
        Select Case oPara.style.NameLocal
            Case "Heading 1"
                nH1 = nH1 + 1
                If nH1 <= expectedBooks Then
                    h1Pages(nH1) = oPara.Range.Information(wdActiveEndPageNumber)
                End If
            Case "BibleIndexList"
                nIdx = nIdx + 1
                If nIdx <= expectedBooks Then
                    Dim fullText As String
                    Dim tabPos As Long
                    Dim trailing As String
                    fullText = Replace(oPara.Range.Text, vbCr, "")
                    tabPos = InStrRev(fullText, vbTab)
                    If tabPos > 0 Then
                        trailing = Trim$(Mid$(fullText, tabPos + 1))
                    Else
                        trailing = ""
                    End If
                    If IsNumeric(trailing) Then
                        idxPages(nIdx) = CLng(trailing)
                    Else
                        idxPages(nIdx) = -1
                    End If
                End If
        End Select
    Next oPara

    Dim mismatches As Long
    Dim sOut As String
    Const NL As String = vbCrLf
    sOut = "---- VerifyBibleIndexPageNumbers: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ----" & NL & NL
    sOut = sOut & "Checked against: " & docPath & NL & NL

    If nH1 <> expectedBooks Or nIdx <> expectedBooks Then
        mismatches = mismatches + 1
        sOut = sOut & "Count MISMATCH: Heading 1 Count=" & nH1 & "  BibleIndexList Count=" & _
            nIdx & "  expected=" & expectedBooks & NL
    End If

    Dim n As Long
    n = expectedBooks
    If nH1 < n Then n = nH1
    If nIdx < n Then n = nIdx

    Dim i As Long
    Dim bookName As String
    For i = 1 To n
        If h1Pages(i) <> idxPages(i) Then
            mismatches = mismatches + 1
            bookName = "?"
            If canonBooks.Exists(i) Then bookName = canonBooks.Item(i)(1)
            sOut = sOut & "MISMATCH book #" & i & " (" & bookName & "): Heading 1 page=" & _
                h1Pages(i) & "  Index row page=" & idxPages(i) & NL
        End If
    Next i

    sOut = sOut & NL & "TOTAL mismatches: " & mismatches & NL
    ExportLog sOut
    WriteReportFileTo oDoc.Path, "VerifyBibleIndexPageNumbers.txt", sOut

    oDoc.Close SaveChanges:=False
    VerifyBibleIndexPageNumbers = mismatches

PROC_EXIT:
    Exit Function
PROC_ERR:
    ExportLog "ERROR in basBibleOnlyExport.VerifyBibleIndexPageNumbers | Erl: " & Erl & _
        " | Err: " & Err.Number & " | " & Err.Description
    If Not oDoc Is Nothing Then
        On Error Resume Next
        oDoc.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Resume PROC_EXIT
End Function

' --------------------------------------------------------------------------
' BuildKeepDict - union of GetScriptureOnlyStyles and GetIndexRegenStyles,
' as a Scripting.Dictionary (TextCompare) for O(1) membership tests.
' --------------------------------------------------------------------------
Private Function BuildKeepDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 ' TextCompare

    Dim arr As Variant
    Dim i As Long

    arr = GetScriptureOnlyStyles()
    For i = LBound(arr) To UBound(arr)
        d(arr(i)) = True
    Next i

    arr = GetIndexRegenStyles()
    For i = LBound(arr) To UBound(arr)
        d(arr(i)) = True
    Next i

    Set BuildKeepDict = d
End Function

' --------------------------------------------------------------------------
' BuildStripDict - GetScriptureStripStyles as a Scripting.Dictionary
' (TextCompare) for O(1) membership tests.
' --------------------------------------------------------------------------
Private Function BuildStripDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 ' TextCompare

    Dim arr As Variant
    Dim i As Long

    arr = GetScriptureStripStyles()
    For i = LBound(arr) To UBound(arr)
        d(arr(i)) = True
    Next i

    Set BuildStripDict = d
End Function

' --------------------------------------------------------------------------
' BuildIndexTableRanges - walk ActiveDocument.Tables once; record the
' [Start, End) Range bounds of every table that contains at least one
' BibleIndexList-styled paragraph. Used by Pass3_ParagraphSweep's
' table-membership rule. Expected to find exactly 2 (Old Testament, New
' Testament) in the current production docm, but written generally.
' --------------------------------------------------------------------------
Private Sub BuildIndexTableRanges(ByVal oDoc As Object, ByRef tblStarts() As Long, _
                                   ByRef tblEnds() As Long, ByRef nTbl As Long)
    Dim tblCount As Long
    tblCount = oDoc.Tables.Count
    nTbl = 0

    If tblCount = 0 Then
        ReDim tblStarts(1 To 1)
        ReDim tblEnds(1 To 1)
        Exit Sub
    End If

    ReDim tblStarts(1 To tblCount)
    ReDim tblEnds(1 To tblCount)

    Dim oTbl As Object
    Dim oPara As Object
    Dim isIndexTable As Boolean

    For Each oTbl In oDoc.Tables
        isIndexTable = False
        For Each oPara In oTbl.Range.Paragraphs
            If oPara.style.NameLocal = "BibleIndexList" Then
                isIndexTable = True
                Exit For
            End If
        Next oPara

        If isIndexTable Then
            nTbl = nTbl + 1
            tblStarts(nTbl) = oTbl.Range.Start
            tblEnds(nTbl) = oTbl.Range.End
        End If
    Next oTbl
End Sub

' --------------------------------------------------------------------------
' WriteReportFileTo - write a report file to <docFolderPath>\rpt\<fileName>,
' creating the rpt folder if needed. Matches this codebase's established
' FSO.CreateTextFile convention for re-runnable report writers (avoids
' Err 55/70 from leaked handles - see [[feedback_fso_file_writes]]).
' --------------------------------------------------------------------------
Private Sub WriteReportFileTo(ByVal docFolderPath As String, ByVal fileName As String, _
                               ByVal sContent As String)
    Dim oFSO As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")

    If Dir(docFolderPath & "\rpt", vbDirectory) = "" Then
        oFSO.CreateFolder docFolderPath & "\rpt"
    End If

    Dim oStream As Object
    Set oStream = oFSO.CreateTextFile(docFolderPath & "\rpt\" & fileName, True, False)
    oStream.Write sContent
    oStream.Close
End Sub

' --------------------------------------------------------------------------
' HaltExport - the single halt mechanism every pass uses. Sets
' m_haltFired so ExportScriptureOnlyDocx's orchestration loop stops
' immediately without closing or saving the working duplicate (per the
' plan's "halt immediately, never catch-and-continue" methodology).
' Deliberately Debug.Print, not MsgBox - see [[feedback_vba_errors_to_immediate]].
' --------------------------------------------------------------------------
Private Sub HaltExport(ByVal passName As String, ByVal reason As String)
    ExportLog "=== HALT in " & passName & " ==="
    ExportLog reason
    m_haltFired = True
    m_issues = m_issues + 1
End Sub

' ==========================================================================
' GetExportSettings
' ==========================================================================
' Single parameter table for the export (plan Doc, v1.0 review finding 8).
' Late-bound Scripting.Dictionary. Defaults reproduce the v1.0 baseline.
' Only keys that a pass actually reads are defined - add a key in the same
' change that adds the pass that uses it.
' ==========================================================================
Public Function GetExportSettings() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("ExportVersion") = "1.0"
    ' Remove manual (optional) hyphens and turn off automatic hyphenation.
    d("StripOptionalHyphens") = True
    ' EmphasisBlack character style removed (text takes VerseText);
    ' EmphasisRed character style replaced by Words of Jesus.
    d("ReplaceEmphasisStyles") = True
    ' VerseText paragraph alignment: "left" (v1.0), "justify" (the docm),
    ' "center", "right"; "" leaves the style as it is.
    d("VerseTextAlignment") = "left"
    ' Post-save automation (python via WSL, no manual py calls):
    ' strip the orphaned customUI ribbon part from the saved .docx, then
    ' verify against a baseline .docx (path relative to the source folder;
    ' the verify step is skipped, and logged, if the file is absent).
    d("StripRibbonAfterSave") = True
    d("BaselineDocxRelPath") = "Bible\v0.0.RadiantWordBible.docx"
    Set GetExportSettings = d
End Function

' ==========================================================================
' Pass4b_StripOptionalHyphens
' ==========================================================================
' Deletes every manual optional hyphen (Word "^-", Chr(31)) from the main
' story and turns off automatic hyphenation. Real hyphens and non-breaking
' hyphens are left alone. Returns the number of optional hyphens still
' present afterwards (expected 0). Must run before any pagination-dependent
' pass (Pass 6/8): removal changes line breaks.
' ==========================================================================
Private Function Pass4b_StripOptionalHyphens(ByVal oDoc As Object) As Long
    On Error GoTo PROC_ERR

    Dim oRng As Object
    Dim removed As Long

    Set oRng = oDoc.Content
    With oRng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^-"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With

    oDoc.AutoHyphenation = False

    ' Verify: Count survivors
    Dim remaining As Long
    Set oRng = oDoc.Content
    With oRng.Find
        .ClearFormatting
        .Text = "^-"
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
        Do While .Execute
            remaining = remaining + 1
            oRng.Collapse wdCollapseEnd
        Loop
    End With

    ExportLog "Pass4b_StripOptionalHyphens: remaining=" & remaining & _
        ", AutoHyphenation=" & oDoc.AutoHyphenation
    Pass4b_StripOptionalHyphens = remaining

PROC_EXIT:
    Exit Function
PROC_ERR:
    ExportLog "ERROR in basBibleOnlyExport.Pass4b_StripOptionalHyphens | Erl: " & Erl & _
        " | Err: " & Err.Number & " | " & Err.Description
    Pass4b_StripOptionalHyphens = -1
    Resume PROC_EXIT
End Function

' ==========================================================================
' Pass4c_ReplaceEmphasisStyles
' ==========================================================================
' B2/B3 (plan Doc, 2026-10-08). Both styles are CHARACTER styles and every
' run sits inside a VerseText paragraph (OOXML check of v0.0), so:
'   EmphasisBlack -> Default Paragraph Font (run takes VerseText formatting)
'   EmphasisRed   -> Words of Jesus
' Format-only Find/Replace on the style, no per-run iteration.
'
' In-Word checks (cheap): full text identical, paragraph Count identical,
' zero remaining Find hits for either style. The exact "nothing else
' changed" check is offline: py\verify_char_style_change.py on the saved
' .docx against a baseline.
'
' Returns the number of violations (expected 0).
' ==========================================================================
Private Function Pass4c_ReplaceEmphasisStyles(ByVal oDoc As Object) As Long
    On Error GoTo PROC_ERR

    Dim violations As Long
    Dim textBefore As String
    Dim parasBefore As Long
    textBefore = oDoc.Content.Text
    parasBefore = oDoc.Paragraphs.Count

    ReplaceCharStyle oDoc, "EmphasisBlack", oDoc.Styles(wdStyleDefaultParagraphFont)
    ReplaceCharStyle oDoc, "EmphasisRed", oDoc.Styles("Words of Jesus")

    If oDoc.Content.Text <> textBefore Then
        violations = violations + 1
        ExportLog "Pass4c: document text changed - the style swap must not alter text."
    End If
    If oDoc.Paragraphs.Count <> parasBefore Then
        violations = violations + 1
        ExportLog "Pass4c: paragraph Count " & parasBefore & " -> " & oDoc.Paragraphs.Count
    End If
    Dim remBlack As Long
    Dim remRed As Long
    remBlack = CountCharStyleHits(oDoc, "EmphasisBlack")
    remRed = CountCharStyleHits(oDoc, "EmphasisRed")
    If remBlack <> 0 Or remRed <> 0 Then
        violations = violations + 1
        ExportLog "Pass4c: EmphasisBlack hits=" & remBlack & ", EmphasisRed hits=" & remRed & _
            ", expected 0"
    End If

    ExportLog "Pass4c_ReplaceEmphasisStyles: violations=" & violations
    Pass4c_ReplaceEmphasisStyles = violations

PROC_EXIT:
    Exit Function
PROC_ERR:
    ExportLog "ERROR in basBibleOnlyExport.Pass4c_ReplaceEmphasisStyles | Erl: " & Erl & _
        " | Err: " & Err.Number & " | " & Err.Description
    Pass4c_ReplaceEmphasisStyles = violations + 1
    Resume PROC_EXIT
End Function

Private Sub ReplaceCharStyle(ByVal oDoc As Object, ByVal fromName As String, _
                             ByVal toStyle As Object)
    Dim oRng As Object
    Set oRng = oDoc.Content
    With oRng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .style = oDoc.Styles(fromName)
        .Replacement.Text = ""
        .Replacement.style = toStyle
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Private Function CountCharStyleHits(ByVal oDoc As Object, ByVal StyleName As String) As Long
    Dim oRng As Object
    Dim n As Long
    Set oRng = oDoc.Content
    With oRng.Find
        .ClearFormatting
        .Text = ""
        .style = oDoc.Styles(StyleName)
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
        Do While .Execute
            n = n + 1
            oRng.Collapse wdCollapseEnd
        Loop
    End With
    CountCharStyleHits = n
End Function

' ==========================================================================
' ExportLog / WriteExportReport
' ==========================================================================
' ExportLog replaces Debug.Print throughout this module: it still prints to
' the Immediate window and also accumulates the line, so the whole run can
' be written to rpt\RadiantWordBibleExport.txt - a permanent, git-tracked
' record outside VBA (every halt, error, and post-save step Result).
' ==========================================================================
Private Sub ExportLog(ByVal sLine As String)
    Debug.Print sLine
    m_report = m_report & sLine & vbCrLf
End Sub

Private Sub WriteExportReport()
    If m_reportFolder = "" Then Exit Sub
    Dim sResult As String
    If m_issues = 0 Then
        sResult = "Result: CLEAN (0 issues)"
    Else
        sResult = "Result: NOT CLEAN (" & m_issues & " issue(s))"
    End If
    On Error Resume Next
    WriteReportFileTo m_reportFolder, "RadiantWordBibleExport.txt", _
        m_report & vbCrLf & sResult & vbCrLf
    If Err.Number <> 0 Then Debug.Print "WriteExportReport failed: Err " & Err.Number & _
        " - " & Err.Description
    Err.Clear
End Sub

' ==========================================================================
' RunPythonStep
' ==========================================================================
' Runs a project py\ script inside WSL (python3 -I, per the project's WSL
' convention) and waits for it. Output is captured to a temp file and
' logged. Returns the script's exit code (-1 if it could not be launched).
' wsl.exe --exec passes arguments without a Shell, so spaces are safe.
' ==========================================================================
Private Function RunPythonStep(ByVal stepName As String, ByVal scriptWinPath As String, _
                               ByVal args As Variant) As Long
    On Error GoTo PROC_ERR

    Dim cmd As String
    Dim i As Long
    cmd = "wsl.exe --exec python3 -I " & Chr(34) & ToWslPath(scriptWinPath) & Chr(34)
    For i = LBound(args) To UBound(args)
        cmd = cmd & " " & Chr(34) & ToWslPath(CStr(args(i))) & Chr(34)
    Next i

    Dim outFile As String
    outFile = Environ$("TEMP") & "\rwb_export_step.txt"
    If Dir(outFile) <> "" Then Kill outFile

    ExportLog stepName & ": running " & cmd
    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")
    Dim rc As Long
    ' cmd /c wrapper (outer quotes) so the output redirection works
    rc = oShell.Run("cmd.exe /c " & Chr(34) & cmd & " > " & Chr(34) & outFile & Chr(34) & _
        " 2>&1" & Chr(34), 0, True)

    If Dir(outFile) <> "" Then
        Dim oFSO As Object
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        Dim oTS As Object
        Set oTS = oFSO.OpenTextFile(outFile, 1)
        Dim sOut As String
        If Not oTS.AtEndOfStream Then sOut = oTS.ReadAll
        oTS.Close
        If Len(sOut) > 0 Then ExportLog sOut
    End If
    ExportLog stepName & ": exit code " & rc

    RunPythonStep = rc

PROC_EXIT:
    Exit Function
PROC_ERR:
    Debug.Print "ERROR in basBibleOnlyExport.RunPythonStep | Err: " & Err.Number & " | " & _
        Err.Description
    RunPythonStep = -1
    Resume PROC_EXIT
End Function

' C:\a\b -> /mnt/c/a/b
Private Function ToWslPath(ByVal winPath As String) As String
    Dim p As String
    p = Replace(winPath, "\", "/")
    If Len(p) >= 2 And Mid$(p, 2, 1) = ":" Then
        p = "/mnt/" & LCase$(Left$(p, 1)) & Mid$(p, 3)
    End If
    ToWslPath = p
End Function

' ==========================================================================
' Pass4d_SetVerseTextAlignment
' ==========================================================================
' v1.0 task (plan Doc, 2026-10-08): the docm's VerseText style is Justified;
' the export wants it left aligned (setting VerseTextAlignment). Applied to
' the style definition of the disposable working copy only - never the
' production docm - through the object model (no Modify Style dialog).
' Any VerseText paragraph carrying a direct alignment override is also
' corrected and counted in the log.
'
' Checks: full text and paragraph Count unchanged; afterwards no VerseText
' paragraph has a different alignment. Returns the violation Count
' (expected 0). The saved .docx is re-checked offline by
' py\verify_char_style_change.py --verse-align.
' ==========================================================================
Private Function Pass4d_SetVerseTextAlignment(ByVal oDoc As Object, _
                                              ByVal alignName As String) As Long
    On Error GoTo PROC_ERR

    Dim violations As Long
    Dim want As Long
    Select Case LCase$(alignName)
        Case "left": want = wdAlignParagraphLeft
        Case "justify": want = wdAlignParagraphJustify
        Case "center": want = wdAlignParagraphCenter
        Case "right": want = wdAlignParagraphRight
        Case Else
            ExportLog "Pass4d: unknown VerseTextAlignment """ & alignName & """"
            Pass4d_SetVerseTextAlignment = 1
            Exit Function
    End Select

    Dim textBefore As String
    Dim parasBefore As Long
    textBefore = oDoc.Content.Text
    parasBefore = oDoc.Paragraphs.Count

    oDoc.Styles("VerseText").ParagraphFormat.Alignment = want

    Dim oPara As Object
    Dim nOverride As Long
    Dim nVerse As Long
    For Each oPara In oDoc.Paragraphs
        If oPara.style.NameLocal = "VerseText" Then
            nVerse = nVerse + 1
            If oPara.Alignment <> want Then
                nOverride = nOverride + 1
                oPara.Alignment = want
            End If
        End If
    Next oPara

    If oDoc.Content.Text <> textBefore Then
        violations = violations + 1
        ExportLog "Pass4d: document text changed - alignment must not alter text."
    End If
    If oDoc.Paragraphs.Count <> parasBefore Then
        violations = violations + 1
        ExportLog "Pass4d: paragraph Count " & parasBefore & " -> " & oDoc.Paragraphs.Count
    End If

    ExportLog "Pass4d_SetVerseTextAlignment: " & alignName & ", " & nVerse & _
        " VerseText paragraph(s), " & nOverride & " direct override(s) corrected, violations=" & _
        violations
    Pass4d_SetVerseTextAlignment = violations

PROC_EXIT:
    Exit Function
PROC_ERR:
    ExportLog "ERROR in basBibleOnlyExport.Pass4d_SetVerseTextAlignment | Erl: " & Erl & _
        " | Err: " & Err.Number & " | " & Err.Description
    Pass4d_SetVerseTextAlignment = violations + 1
    Resume PROC_EXIT
End Function
