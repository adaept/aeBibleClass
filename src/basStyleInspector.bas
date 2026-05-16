Attribute VB_Name = "basStyleInspector"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

' Session-scoped dictionary of last-known runtimes (seconds), keyed by
' routine name. Reset on Word restart / VBA project reset.
Private mLastRuntimes As Object

'==============================================================================
' StartTimer / EndTimer  (session-scoped routine timing)
'==============================================================================
' Bracket a long-running routine to print expected (last-run) and actual
' duration to the Immediate window. First-run-this-session prints a no-prior
' notice instead of an expected value.
'
' Usage:
'   Public Sub LongRoutine()
'       Dim t As Double
'       StartTimer "LongRoutine", t
'       ' ... work ...
'       EndTimer "LongRoutine", t
'   End Sub
'==============================================================================
Public Sub StartTimer(ByVal sName As String, ByRef startTime As Double)
    Dim d As Object
    Set d = GetRuntimeDict()
    If d.Exists(sName) Then
        Debug.Print sName & " - expected ~" & Format(d(sName), "0.00") & " sec (last run)"
    Else
        Debug.Print sName & " - first run this session, no prior timing"
    End If
    startTime = Timer
End Sub

Public Sub EndTimer(ByVal sName As String, ByVal startTime As Double)
    Dim runTime As Double
    runTime = Timer - startTime
    GetRuntimeDict()(sName) = runTime
    Debug.Print sName & " - actual " & Format(runTime, "0.00") & " sec"
End Sub

Private Function GetRuntimeDict() As Object
    If mLastRuntimes Is Nothing Then Set mLastRuntimes = CreateObject("Scripting.Dictionary")
    Set GetRuntimeDict = mLastRuntimes
End Function

'==============================================================================
' DumpStyleProperties
'==============================================================================
' Prints a named style's properties in a form that can be pasted into a
' Define<Style> routine. Output goes to the Immediate window; optionally also
' writes to rpt\Styles\style_<name>.txt for diffing / style-guide generation.
'
' Usage:
'   DumpStyleProperties "FrontPageBodyText"         ' Immediate only
'   DumpStyleProperties "FrontPageBodyText", True   ' also rpt\ file
'==============================================================================
Public Sub DumpStyleProperties(ByVal sStyleName As String, _
                               Optional ByVal bWriteFile As Boolean = False)

    Dim oDoc   As Object
    Dim oStyle As Object
    Dim oFont  As Object
    Dim oPF    As Object
    Dim sOut   As String
    Const NL   As String = vbCrLf

    Set oDoc = ActiveDocument

    On Error Resume Next
    Set oStyle = oDoc.Styles(sStyleName)
    On Error GoTo 0
    If oStyle Is Nothing Then
        Debug.Print "ERROR: style '" & sStyleName & "' not found."
        Exit Sub
    End If

    sOut = "'--- " & sStyleName & "  (Type=" & StyleTypeName(oStyle.Type) & _
           ", Priority=" & oStyle.Priority & ") ---" & NL
    sOut = sOut & ".BaseStyle = """ & CStr(oStyle.baseStyle) & """" & NL
    sOut = sOut & ".QuickStyle = " & oStyle.QuickStyle & NL

    Set oFont = oStyle.Font
    sOut = sOut & ".Font.Name = """ & oFont.Name & """" & NL
    sOut = sOut & ".Font.Size = " & oFont.Size & NL
    sOut = sOut & ".Font.Bold = " & oFont.Bold & NL
    sOut = sOut & ".Font.Italic = " & oFont.Italic & NL
    sOut = sOut & ".Font.Underline = " & oFont.Underline & NL
    sOut = sOut & ".Font.Color = " & oFont.Color & NL
    sOut = sOut & ".Font.SmallCaps = " & oFont.SmallCaps & NL
    sOut = sOut & ".Font.AllCaps = " & oFont.AllCaps & NL

    If oStyle.Type = 1 Then        ' wdStyleTypeParagraph
        ' Paragraph-only properties (error 5900 if accessed on character styles)
        sOut = sOut & ".NextParagraphStyle = """ & CStr(oStyle.NextParagraphStyle) & """" & NL
        sOut = sOut & ".AutomaticallyUpdate = " & oStyle.AutomaticallyUpdate & NL
        Set oPF = oStyle.ParagraphFormat
        sOut = sOut & ".ParagraphFormat.Alignment = " & oPF.Alignment & NL
        sOut = sOut & ".ParagraphFormat.LeftIndent = " & oPF.LeftIndent & NL
        sOut = sOut & ".ParagraphFormat.RightIndent = " & oPF.RightIndent & NL
        sOut = sOut & ".ParagraphFormat.FirstLineIndent = " & oPF.FirstLineIndent & NL
        sOut = sOut & ".ParagraphFormat.SpaceBefore = " & oPF.SpaceBefore & NL
        sOut = sOut & ".ParagraphFormat.SpaceAfter = " & oPF.SpaceAfter & NL
        sOut = sOut & ".ParagraphFormat.LineSpacing = " & oPF.LineSpacing & NL
        sOut = sOut & ".ParagraphFormat.LineSpacingRule = " & oPF.LineSpacingRule & NL
        sOut = sOut & ".ParagraphFormat.WidowControl = " & oPF.WidowControl & NL
        sOut = sOut & ".ParagraphFormat.KeepTogether = " & oPF.KeepTogether & NL
        sOut = sOut & ".ParagraphFormat.KeepWithNext = " & oPF.KeepWithNext & NL
        sOut = sOut & ".ParagraphFormat.PageBreakBefore = " & oPF.PageBreakBefore & NL
        sOut = sOut & ".ParagraphFormat.OutlineLevel = " & oPF.OutlineLevel & NL

        ' Tab stops - explicit stops only; Word's default tab grid (DefaultTabStop)
        ' is not part of this collection. Skip the block when Count = 0 so styles
        ' without custom tabs produce no extra output.
        Dim tabCount As Long
        tabCount = oPF.TabStops.Count
        If tabCount > 0 Then
            sOut = sOut & ".ParagraphFormat.TabStops.Count = " & tabCount & NL
            Dim ts As Word.TabStop
            Dim tsIdx As Long
            tsIdx = 0
            For Each ts In oPF.TabStops
                tsIdx = tsIdx + 1
                sOut = sOut & ".ParagraphFormat.TabStops(" & tsIdx & ") = " & _
                       "Position=" & ts.Position & _
                       " Align=" & TabAlignName(ts.Alignment) & _
                       " Leader=" & TabLeaderName(ts.Leader) & NL
            Next ts
        End If
    End If

    Debug.Print sOut
    If bWriteFile Then WriteStyleDump sStyleName, sOut
End Sub

Public Function TabAlignName(ByVal a As Long) As String
    Select Case a
        Case wdAlignTabLeft:    TabAlignName = "Left"
        Case wdAlignTabCenter:  TabAlignName = "Center"
        Case wdAlignTabRight:   TabAlignName = "Right"
        Case wdAlignTabDecimal: TabAlignName = "Decimal"
        Case wdAlignTabBar:     TabAlignName = "Bar"
        Case wdAlignTabList:    TabAlignName = "List"
        Case Else:              TabAlignName = "Unknown(" & a & ")"
    End Select
End Function

Public Function TabLeaderName(ByVal l As Long) As String
    Select Case l
        Case wdTabLeaderSpaces:    TabLeaderName = "Spaces"
        Case wdTabLeaderDots:      TabLeaderName = "Dots"
        Case wdTabLeaderDashes:    TabLeaderName = "Dashes"
        Case wdTabLeaderLines:     TabLeaderName = "Lines"
        Case wdTabLeaderHeavy:     TabLeaderName = "Heavy"
        Case wdTabLeaderMiddleDot: TabLeaderName = "MiddleDot"
        Case Else:                 TabLeaderName = "Unknown(" & l & ")"
    End Select
End Function

Private Function StyleTypeName(ByVal lType As Long) As String
    Select Case lType
        Case 1: StyleTypeName = "Paragraph"
        Case 2: StyleTypeName = "Character"
        Case 3: StyleTypeName = "Table"
        Case 4: StyleTypeName = "List"
        Case Else: StyleTypeName = "Unknown(" & lType & ")"
    End Select
End Function

Private Sub WriteStyleDump(ByVal sStyleName As String, ByVal sContent As String)
    Dim oFSO    As Object
    Dim oStream As Object
    Dim sPath   As String
    sPath = ActiveDocument.Path & "\rpt\Styles\style_" & SafeFileName(sStyleName) & ".txt"
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.CreateTextFile(sPath, True, False)   ' ASCII
    oStream.Write sContent
    oStream.Close
End Sub

Public Function SafeFileName(ByVal s As String) As String
    SafeFileName = Replace(Replace(Replace(s, " ", "_"), "/", "_"), "\", "_")
End Function

'==============================================================================
' DumpAllApprovedStyles
'==============================================================================
' Dumps every "approved" paragraph or character style to rpt\Styles\style_<name>.txt,
' in Priority order. "Approved" = Priority <> 99, matching the convention set
' by PromoteApprovedStyles / DumpPrioritiesSorted in basTEST_aeBibleConfig.
'
' Source of truth is the live document - whatever PromoteApprovedStyles has
' promoted is what gets dumped. No duplicate approved-list in this module.
'
' Usage:
'   DumpAllApprovedStyles
'==============================================================================
Public Sub DumpAllApprovedStyles()
    Dim oDoc   As Object
    Dim oStyle As Object
    Dim arr()  As Variant
    Dim nCount As Long
    Dim i As Long, j As Long
    Dim tmpName As String
    Dim tmpPri  As Long
    Dim t       As Double

    StartTimer "DumpAllApprovedStyles", t
    Set oDoc = ActiveDocument

    ' First pass - Count eligible approved styles
    For Each oStyle In oDoc.Styles
        If oStyle.Type = 1 Or oStyle.Type = 2 Then    ' Paragraph or Character
            If oStyle.Priority <> 99 Then
                nCount = nCount + 1
            End If
        End If
    Next oStyle

    If nCount = 0 Then
        Debug.Print "DumpAllApprovedStyles: no approved styles (Priority <> 99) found."
        Exit Sub
    End If

    ReDim arr(1 To nCount, 1 To 2)

    ' Second pass - fill array with (Name, Priority)
    nCount = 1
    For Each oStyle In oDoc.Styles
        If oStyle.Type = 1 Or oStyle.Type = 2 Then
            If oStyle.Priority <> 99 Then
                arr(nCount, 1) = oStyle.NameLocal
                arr(nCount, 2) = oStyle.Priority
                nCount = nCount + 1
            End If
        End If
    Next oStyle
    nCount = nCount - 1

    ' Bubble sort by Priority ascending (matches DumpPrioritiesSorted convention)
    For i = 1 To nCount - 1
        For j = i + 1 To nCount
            If arr(j, 2) < arr(i, 2) Then
                tmpName = arr(i, 1)
                tmpPri = arr(i, 2)
                arr(i, 1) = arr(j, 1)
                arr(i, 2) = arr(j, 2)
                arr(j, 1) = tmpName
                arr(j, 2) = tmpPri
            End If
        Next j
    Next i

    Dim nFailed As Long
    Debug.Print "---- DumpAllApprovedStyles: " & nCount & " style(s) ----"
    For i = 1 To nCount
        Debug.Print "[" & arr(i, 2) & "] " & arr(i, 1)
        On Error Resume Next
        DumpStyleProperties CStr(arr(i, 1)), True
        If Err.Number <> 0 Then
            Debug.Print "  !! FAILED: " & arr(i, 1) & " - err " & Err.Number & ": " & Err.Description
            nFailed = nFailed + 1
            Err.Clear
        End If
        On Error GoTo 0
    Next i
    Debug.Print "DumpAllApprovedStyles: Done. " & (nCount - nFailed) & " succeeded, " & nFailed & " failed."
    CleanupOrphanStyleDumps arr, nCount
    EndTimer "DumpAllApprovedStyles", t
End Sub

'==============================================================================
' CleanupOrphanStyleDumps  (helper)
'==============================================================================
' Compares the set of files just written by DumpAllApprovedStyles against the
' on-disk contents of rpt\Styles\style_*.txt. Any file that exists on disk but
' was not written this run is an "orphan" - typically left over from a style
' rename (e.g., ContentsCPBB -> Contents leaves style_ContentsCPBB.txt behind).
'
' Lists orphans to the Immediate window and prompts (single MsgBox) to delete
' them all. No prompt if there are no orphans.
'==============================================================================
Private Sub CleanupOrphanStyleDumps(ByRef arr As Variant, ByVal nCount As Long)
    Dim oFSO       As Object
    Dim oFolder    As Object
    Dim oFile      As Object
    Dim oExpected  As Object
    Dim arrOrphans() As String
    Dim nOrphan    As Long
    Dim k          As Long
    Dim sBase      As String

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oExpected = CreateObject("Scripting.Dictionary")
    oExpected.CompareMode = 1   ' TextCompare (case-insensitive)

    For k = 1 To nCount
        oExpected("style_" & SafeFileName(CStr(arr(k, 1))) & ".txt") = True
    Next k

    Set oFolder = oFSO.GetFolder(ActiveDocument.Path & "\rpt\Styles")
    If oFolder.Files.Count = 0 Then Exit Sub

    ReDim arrOrphans(1 To oFolder.Files.Count)
    nOrphan = 0
    For Each oFile In oFolder.Files
        sBase = oFile.Name
        If LCase$(Left$(sBase, 6)) = "style_" And LCase$(Right$(sBase, 4)) = ".txt" Then
            If Not oExpected.Exists(sBase) Then
                nOrphan = nOrphan + 1
                arrOrphans(nOrphan) = oFile.Path
            End If
        End If
    Next oFile

    If nOrphan = 0 Then Exit Sub

    Debug.Print "Orphan style dumps (in rpt\Styles but not in current approved list):"
    For k = 1 To nOrphan
        Debug.Print "  " & arrOrphans(k)
    Next k

    If MsgBox(nOrphan & " orphan style dump(s) found. Delete them?", _
              vbYesNo + vbQuestion, "DumpAllApprovedStyles cleanup") = vbYes Then
        For k = 1 To nOrphan
            oFSO.DeleteFile arrOrphans(k), True
            Debug.Print "  deleted: " & arrOrphans(k)
        Next k
    Else
        Debug.Print "  skipped deletion of " & nOrphan & " orphan(s)."
    End If
End Sub

'==============================================================================
' ListApprovedStylesByBookOrder
'==============================================================================
' For every approved style (Priority <> 99), finds the page number of its
' FIRST occurrence across all stories in the document (main body, headers,
' footers, footnotes, endnotes, text frames, comments) and lists them in
' page order. Unused approved styles are flagged [not used]. Sort is
' (Page ascending, Priority ascending).
'
' Output goes to the Immediate window; optionally also to
' rpt\Styles\styles_book_order.txt for QA diffing against the approved array.
'
' Usage:
'   ListApprovedStylesByBookOrder
'   ListApprovedStylesByBookOrder True     ' also writes rpt file
'==============================================================================
Public Sub ListApprovedStylesByBookOrder(Optional ByVal bWriteFile As Boolean = False)
    Dim oDoc    As Object
    Dim oStyle  As Object
    Dim oRng    As Object
    Dim arr()   As Variant
    Dim nCount  As Long, i As Long, j As Long
    Dim lPage   As Long
    Dim pi As Long, pj As Long
    Dim tmpName As String, tmpPri As Long, tmpPage As Long
    Dim sOut    As String, sLine As String
    Dim t       As Double
    Const NL    As String = vbCrLf

    StartTimer "ListApprovedStylesByBookOrder", t
    Set oDoc = ActiveDocument

    ' Count approved paragraph/character styles
    For Each oStyle In oDoc.Styles
        If oStyle.Type = 1 Or oStyle.Type = 2 Then
            If oStyle.Priority <> 99 Then nCount = nCount + 1
        End If
    Next oStyle

    If nCount = 0 Then
        Debug.Print "ListApprovedStylesByBookOrder: no approved styles (Priority <> 99) found."
        Exit Sub
    End If

    ReDim arr(1 To nCount, 1 To 3)    ' (Name, Priority, Page)

    ' Find first-occurrence page per approved style across all stories
    ' (main body, headers, footers, footnotes, etc.)
    nCount = 1
    For Each oStyle In oDoc.Styles
        If (oStyle.Type = 1 Or oStyle.Type = 2) And oStyle.Priority <> 99 Then
            lPage = FirstPageForStyle(oDoc, oStyle)
            arr(nCount, 1) = oStyle.NameLocal
            arr(nCount, 2) = oStyle.Priority
            arr(nCount, 3) = lPage
            nCount = nCount + 1
        End If
    Next oStyle
    nCount = nCount - 1

    ' Sort by Page ascending; "not used" (-1) pushed to end via MAX_LONG
    For i = 1 To nCount - 1
        For j = i + 1 To nCount
            pi = arr(i, 3)
            If pi = -1 Then pi = 2147483647
            pj = arr(j, 3)
            If pj = -1 Then pj = 2147483647
            If pj < pi Or (pj = pi And arr(j, 2) < arr(i, 2)) Then
                tmpName = arr(i, 1)
                tmpPri = arr(i, 2)
                tmpPage = arr(i, 3)
                arr(i, 1) = arr(j, 1)
                arr(i, 2) = arr(j, 2)
                arr(i, 3) = arr(j, 3)
                arr(j, 1) = tmpName
                arr(j, 2) = tmpPri
                arr(j, 3) = tmpPage
            End If
        Next j
    Next i

    sOut = "Approved styles in book order (by page of first occurrence)" & NL
    sOut = sOut & " Page | Prio | Style" & NL
    sOut = sOut & "------+------+-----------------------------" & NL
    For i = 1 To nCount
        If arr(i, 3) = -1 Then
            sLine = "    - | " & Right("    " & arr(i, 2), 4) & " | " & arr(i, 1) & "  [not used]"
        Else
            sLine = Right("    " & arr(i, 3), 5) & " | " & Right("    " & arr(i, 2), 4) & " | " & arr(i, 1)
        End If
        sOut = sOut & sLine & NL
    Next i

    Debug.Print sOut
    If bWriteFile Then WriteBookOrderFile sOut
    EndTimer "ListApprovedStylesByBookOrder", t
End Sub

Private Sub WriteBookOrderFile(ByVal sContent As String)
    Dim oFSO    As Object
    Dim oStream As Object
    Dim sPath   As String
    sPath = ActiveDocument.Path & "\rpt\Styles\styles_book_order.txt"
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.CreateTextFile(sPath, True, False)    ' ASCII
    oStream.Write sContent
    oStream.Close
End Sub

'==============================================================================
' FirstPageForStyle  (helper)
'==============================================================================
' Returns the page number of the first occurrence of oStyle across every
' story in the document (main body, headers, footers, footnotes, endnotes,
' text frames, comments). Walks each story's NextStoryRange chain to cover
' section-specific headers/footers.
'
' Returns -1 only if the style is not found in any story. If the style
' IS found but no story returns a positive page number (e.g., header /
' footer stories where Range.Information(wdActiveEndPageNumber) returns -1
' because the header tiles across many pages), returns 1 as a best-effort
' fallback - these styles "first appear" on page 1 of the single-section
' document.
'==============================================================================
Private Function FirstPageForStyle(ByVal oDoc As Object, ByVal oStyle As Object) As Long
    Dim oStory         As Object
    Dim oNext          As Object
    Dim oSection       As Object
    Dim oFindRng       As Object
    Dim bestPage       As Long
    Dim thisPage       As Long
    Dim bFoundAnywhere As Boolean
    Dim i              As Long

    bestPage = -1
    bFoundAnywhere = False

    ' Stories enumerable via For Each StoryRanges (main body, footnotes,
    ' endnotes, text frames, comments). Header/footer stories (types 6-11)
    ' are skipped here - the explicit Sections walk below handles them, and
    ' Find inside a header story returns Information() pages tied to the
    ' section anchor rather than where the header first applies.
    For Each oStory In oDoc.StoryRanges
        Select Case oStory.StoryType
            Case 6, 7, 8, 9, 10, 11    ' header / footer story types - handled by Sections walk
                ' skip
            Case Else
                Set oFindRng = oStory.Duplicate
                thisPage = FindStylePage(oFindRng, oStyle, bFoundAnywhere)
                If thisPage > 0 Then
                    If bestPage = -1 Or thisPage < bestPage Then bestPage = thisPage
                End If

                Set oNext = oStory.NextStoryRange
                Do While Not oNext Is Nothing
                    Set oFindRng = oNext.Duplicate
                    thisPage = FindStylePage(oFindRng, oStyle, bFoundAnywhere)
                    If thisPage > 0 Then
                        If bestPage = -1 Or thisPage < bestPage Then bestPage = thisPage
                    End If
                    Set oNext = oNext.NextStoryRange
                Loop
        End Select
    Next oStory

    ' Headers and Footers - iterate paragraphs directly. Find with empty text
    ' on a tab-only or paragraph-mark-only header range does not match
    ' reliably, even with a style filter. Paragraph iteration is deterministic
    ' and headers/footers are tiny so perf is fine.
    For Each oSection In oDoc.Sections
        For i = 1 To 3   ' wdHeaderFooterEvenPages=1, Primary=2, FirstPage=3
            thisPage = FirstPageInParagraphs(oSection.Headers(i).Range, oStyle, bFoundAnywhere)
            If thisPage > 0 Then
                If bestPage = -1 Or thisPage < bestPage Then bestPage = thisPage
            End If
            thisPage = FirstPageInParagraphs(oSection.Footers(i).Range, oStyle, bFoundAnywhere)
            If thisPage > 0 Then
                If bestPage = -1 Or thisPage < bestPage Then bestPage = thisPage
            End If
        Next i
    Next oSection

    If bestPage = -1 And bFoundAnywhere Then
        FirstPageForStyle = 1    ' fallback: found only in header/footer-type stories
    Else
        FirstPageForStyle = bestPage
    End If
End Function

'==============================================================================
' FindStylePage  (helper)
'==============================================================================
' Runs Find on oRng with oStyle as the style filter. Returns the page number
' of the first match, or -1 if not found. oRng is mutated by Find; pass a
' Duplicate if the caller needs the original range preserved.
'
' Sets bFoundAnywhere := True when Find succeeds, regardless of whether the
' returned page number is positive. Caller uses this to distinguish "truly
' not used" from "used in a story that can't report a single page".
'==============================================================================
Private Function FindStylePage(ByVal oRng As Object, _
                                ByVal oStyle As Object, _
                                ByRef bFoundAnywhere As Boolean) As Long
    With oRng.Find
        .ClearFormatting
        .Text = ""
        .style = oStyle
        .Forward = True
        .Wrap = 0           ' wdFindStop
        .Format = True
        .MatchWildcards = False
        If .Execute Then
            bFoundAnywhere = True
            FindStylePage = oRng.Information(3)   ' wdActiveEndPageNumber
        Else
            FindStylePage = -1
        End If
    End With
End Function

'==============================================================================
' FirstPageInParagraphs  (helper)
'==============================================================================
' Walks oRng.Paragraphs looking for any paragraph whose Style.NameLocal
' matches oStyle. Sets bFoundAnywhere := True on the first match and exits.
' Always returns -1 - the caller's page-1 fallback handles header/footer hits.
'
' Used for header / footer ranges where:
'   - Range.Find with empty text + style filter does not match tab-only or
'     paragraph-mark-only content; and
'   - Paragraph.Range.Information(wdActiveEndPageNumber) returns a misleading
'     section-anchor page (e.g., 417 instead of 1) for header paragraphs.
'
' Treating header/footer matches as "page 1" via the fallback is correct
' for headers that tile from the start of the document (the case here).
'==============================================================================
Private Function FirstPageInParagraphs(ByVal oRng As Object, _
                                        ByVal oStyle As Object, _
                                        ByRef bFoundAnywhere As Boolean) As Long
    Dim oPara      As Object
    Dim sStyleName As String

    sStyleName = oStyle.NameLocal

    For Each oPara In oRng.Paragraphs
        If oPara.style.NameLocal = sStyleName Then
            bFoundAnywhere = True
            Exit For
        End If
    Next oPara

    FirstPageInParagraphs = -1
End Function

'==============================================================================
' DumpHeaderFooterStyles
'==============================================================================
' Diagnostic. Walks every section x every header/footer slot and reports the
' first paragraph's style and a text excerpt. Read-only. Writes to
' rpt\Styles\header_footer_audit.txt and prints a summary to the Immediate
' window.
'
' Use to figure out which sections actually carry custom header/footer styles
' versus the built-in "Header" / "Footer" defaults, especially when a Find
' for a style appears to fail.
'
' Usage:
'   DumpHeaderFooterStyles
'==============================================================================
Public Sub DumpHeaderFooterStyles()
    Dim oDoc     As Object
    Dim oSection As Object
    Dim secNum   As Long
    Dim i        As Long
    Dim sOut     As String
    Dim t        As Double
    Const NL     As String = vbCrLf

    StartTimer "DumpHeaderFooterStyles", t
    Set oDoc = ActiveDocument
    sOut = "---- DumpHeaderFooterStyles: " & oDoc.Sections.Count & " sections ----" & NL

    secNum = 0
    For Each oSection In oDoc.Sections
        secNum = secNum + 1
        For i = 1 To 3   ' wdHeaderFooterEvenPages=1, Primary=2, FirstPage=3
            sOut = sOut & FormatHFLine(secNum, "Header(" & i & ")", oSection.Headers(i)) & NL
            sOut = sOut & FormatHFLine(secNum, "Footer(" & i & ")", oSection.Footers(i)) & NL
        Next i
    Next oSection

    WriteHeaderFooterAuditFile sOut
    Debug.Print "DumpHeaderFooterStyles: wrote " _
              & (oDoc.Sections.Count * 6) & " lines to rpt\Styles\header_footer_audit.txt"
    EndTimer "DumpHeaderFooterStyles", t
End Sub

Private Function FormatHFLine(ByVal secNum As Long, _
                              ByVal sKind As String, _
                              ByVal oHF As Object) As String
    Dim oRng       As Object
    Dim sLink      As String
    Dim sStyle     As String
    Dim sText      As String
    Dim nParaCount As Long

    sLink = ""
    If oHF.LinkToPrevious Then sLink = " linked"

    Set oRng = oHF.Range
    nParaCount = oRng.Paragraphs.Count

    If nParaCount = 0 Then
        sStyle = "(none)"
        sText = ""
    Else
        sStyle = oRng.Paragraphs(1).style.NameLocal
        sText = Replace(oRng.Paragraphs(1).Range.Text, vbCr, "")
        sText = Replace(sText, vbTab, "<tab>")
        If Len(sText) > 50 Then sText = Left(sText, 50) & "..."
    End If

    FormatHFLine = "Sec " & Right("000" & secNum, 3) & " " & sKind & sLink _
                 & "  paras=" & nParaCount _
                 & "  style=" & sStyle _
                 & "  text=[" & sText & "]"
End Function

Private Sub WriteHeaderFooterAuditFile(ByVal sContent As String)
    Dim oFSO    As Object
    Dim oStream As Object
    Dim sPath   As String
    sPath = ActiveDocument.Path & "\rpt\Styles\header_footer_audit.txt"
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.CreateTextFile(sPath, True, False)    ' ASCII
    oStream.Write sContent
    oStream.Close
End Sub

'==============================================================================
' AuditCharStyleBases
' PURPOSE:
'   Scan every character style explicitly used in the document and check
'   that it is based on "Default Paragraph Font". Any style with a
'   different base is printed to the Immediate window as
'       <StyleName>  ->  <baseStyleName>
'   The function returns the number of offenders. Returns 0 if all
'   in-use character styles are correctly based on the default.
'
'   "Explicitly used" = Style.InUse = True (Word marks a style as InUse
'   when at least one run carries it OR when it has been edited; for an
'   in-production docx this is the practical "is referenced anywhere"
'   signal).
'
' EXCLUSION:
'   "Default Paragraph Font" itself is the root character style. Querying
'   its BaseStyle returns empty because there is nothing above it to
'   inherit from, and it cannot be based on itself. Including it would
'   inflate the offender Count by one and report a non-fixable false
'   positive, so the loop skips it by name. The reported Count therefore
'   reflects only styles that can actually be repointed to the default.
'
' Usage from Immediate:
'   ?AuditCharStyleBases
'==============================================================================
Public Function AuditCharStyleBases() As Long
    Const DEFAULT_BASE As String = "Default Paragraph Font"
    Dim oStyle    As Word.Style
    Dim sBase     As String
    Dim nChecked  As Long
    Dim nOff      As Long

    For Each oStyle In ActiveDocument.Styles
        If oStyle.Type = wdStyleTypeCharacter Then
            If oStyle.InUse Then
                If oStyle.NameLocal = DEFAULT_BASE Then
                    ' Root character style - cannot be based on itself; skip.
                Else
                    nChecked = nChecked + 1
                    On Error Resume Next
                    sBase = ""
                    sBase = CStr(oStyle.baseStyle)
                    On Error GoTo 0
                    If sBase <> DEFAULT_BASE Then
                        Debug.Print oStyle.NameLocal & "  ->  " & _
                                    IIf(Len(sBase) = 0, "(none)", sBase)
                        nOff = nOff + 1
                    End If
                End If
            End If
        End If
    Next oStyle

    Debug.Print "AuditCharStyleBases: checked " & nChecked & _
                " in-use character style(s) (excluding """ & DEFAULT_BASE & _
                """); " & nOff & " not based on """ & DEFAULT_BASE & """."
    AuditCharStyleBases = nOff
End Function

'==============================================================================
' ScanCharStyleApplications
' PURPOSE:
'   Distinguish character styles that are merely present in the styles
'   palette from those actually carried by at least one run of text,
'   AND distinguish Word built-in styles (which cannot be deleted) from
'   custom styles (which can). Word's Style.InUse flag is True for any
'   custom character style from the moment it is created, and for any
'   built-in style that has been applied or modified - so InUse alone
'   cannot answer either "is this style live in the document?" or "is
'   this style deletable?".
'
'   For each in-use character style this function performs a Find with
'   Style filter across all primary StoryRanges (main body, footnotes,
'   endnotes, headers, footers). Output:
'       <StyleName>  ->  Applied  [Builtin|Custom]
'       <StyleName>  ->  Unapplied  [Builtin|Custom]
'
' RETURN:
'   Count of Unapplied AND Custom styles - the only deletable cruft.
'   Built-in Unapplied styles are tallied separately in the summary but
'   excluded from the return value because Word will recreate them on
'   demand (theme switches, paste from HTML, comment use, etc.) and they
'   cannot be removed. The return value is the action list size: how
'   many palette-cruft styles are candidates for Style.Delete.
'
' EXCLUSION:
'   "Default Paragraph Font" is the document-wide default and is implicit
'   on every run that has no explicit character style. Find with
'   .Style = "Default Paragraph Font" is therefore both expensive and
'   meaningless, so the loop skips it by name. This matches the
'   exclusion in AuditCharStyleBases for symmetry.
'
'   Find with .Style = oStyle reports a hit only if a run explicitly
'   carries the style. It does not infer application via inheritance
'   chains - that is the correct semantics for "live in the text".
'
' Usage from Immediate:
'   ?ScanCharStyleApplications
'==============================================================================
Public Function ScanCharStyleApplications() As Long
    Const DEFAULT_BASE As String = "Default Paragraph Font"
    Dim oDoc      As Word.Document
    Dim oStyle    As Word.Style
    Dim story     As Word.Range
    Dim probe     As Word.Range
    Dim nChecked  As Long
    Dim nAppBI    As Long, nAppCu As Long
    Dim nUnaBI    As Long, nUnaCu As Long
    Dim found     As Boolean
    Dim kind      As String

    Set oDoc = ActiveDocument

    For Each oStyle In oDoc.Styles
        If oStyle.Type = wdStyleTypeCharacter Then
            If oStyle.InUse Then
                If oStyle.NameLocal = DEFAULT_BASE Then
                    ' Implicit default; skip - see EXCLUSION in header.
                Else
                    nChecked = nChecked + 1
                    kind = IIf(oStyle.BuiltIn, "Builtin", "Custom")

                    found = False
                    For Each story In oDoc.StoryRanges
                        Set probe = story.Duplicate
                        With probe.Find
                            .ClearFormatting
                            .style = oStyle
                            .Text = ""
                            .Forward = True
                            .Wrap = wdFindStop
                            .Format = True
                            .MatchWildcards = False
                            If .Execute Then
                                found = True
                                Exit For
                            End If
                        End With
                    Next story

                    If found Then
                        Debug.Print oStyle.NameLocal & "  ->  Applied  [" & kind & "]"
                        If oStyle.BuiltIn Then nAppBI = nAppBI + 1 Else nAppCu = nAppCu + 1
                    Else
                        Debug.Print oStyle.NameLocal & "  ->  Unapplied  [" & kind & "]"
                        If oStyle.BuiltIn Then nUnaBI = nUnaBI + 1 Else nUnaCu = nUnaCu + 1
                    End If
                End If
            End If
        End If
    Next oStyle

    Debug.Print "ScanCharStyleApplications: checked " & nChecked & _
                " in-use character style(s) (excluding """ & DEFAULT_BASE & """)."
    Debug.Print "  Applied   : Builtin=" & nAppBI & "  Custom=" & nAppCu & _
                "  (total " & (nAppBI + nAppCu) & ")"
    Debug.Print "  Unapplied : Builtin=" & nUnaBI & "  Custom=" & nUnaCu & _
                "  (total " & (nUnaBI + nUnaCu) & ")"
    Debug.Print "  Deletable cruft (Unapplied & Custom): " & nUnaCu
    ScanCharStyleApplications = nUnaCu
End Function

'==============================================================================
' AuditFootnoteReferenceMarkers
'==============================================================================
' Scan the Footnotes story for "Footnote Reference"-styled runs. For each
' run, determine which footnote (if any) contains it by comparing the
' run's Start position against each footnote.Range start/end. Report runs
' that fall outside every footnote.Range - those are orphan FR markers
' (the stray we are hunting).
'
' Why this approach: the more obvious "walk each footnote, Count its FR
' markers" loop fails on this document. footnote.Range either excludes
' the auto-numbered marker or Find with .Style = "Footnote Reference"
' does not match the field-Result character, so the per-footnote
' walker returns 0 even for legitimate footnotes. Scanning the
' Footnotes story directly and classifying by position is reliable
' regardless of whether Range includes the auto-number.
'
' Output:
'   - One line per orphan FR run (if any), with story position and text.
'   - Summary: total FR runs in Footnotes story, Count inside a footnote,
'     Count of orphans.
'
' RETURN:
'   Orphan Count. Expected 0 in a clean document.
'
' Usage from Immediate:
'   ?AuditFootnoteReferenceMarkers
'==============================================================================
Public Function AuditFootnoteReferenceMarkers() As Long
    Const FR_STYLE As String = "Footnote Reference"
    Const MAX_SNIP As Long = 60
    Dim oDoc       As Word.Document
    Dim story      As Word.Range
    Dim s          As Word.Range
    Dim probe      As Word.Range
    Dim totalFn    As Long
    Dim fnStart()  As Long
    Dim fnEnd()    As Long
    Dim i          As Long
    Dim runStart   As Long
    Dim totalFR    As Long
    Dim insideCnt  As Long
    Dim orphans    As Long
    Dim inFn       As Long
    Dim snip       As String
    Dim frPerFn()  As Long
    Dim anomalies  As Long
    Dim pageNum    As Long

    Set oDoc = ActiveDocument

    ' Find the Footnotes story.
    For Each s In oDoc.StoryRanges
        If s.StoryType = wdFootnotesStory Then
            Set story = s
            Exit For
        End If
    Next s
    If story Is Nothing Then
        Debug.Print "AuditFootnoteReferenceMarkers: no Footnotes story in this document."
        AuditFootnoteReferenceMarkers = 0
        Exit Function
    End If

    ' Snapshot each footnote's range bounds for fast in-range tests.
    totalFn = oDoc.Footnotes.Count
    If totalFn = 0 Then
        Debug.Print "AuditFootnoteReferenceMarkers: no footnotes in this document."
        AuditFootnoteReferenceMarkers = 0
        Exit Function
    End If
    ReDim fnStart(1 To totalFn)
    ReDim fnEnd(1 To totalFn)
    ReDim frPerFn(1 To totalFn)
    For i = 1 To totalFn
        fnStart(i) = oDoc.Footnotes(i).Range.Start
        fnEnd(i) = oDoc.Footnotes(i).Range.End
    Next i

    Debug.Print "AuditFootnoteReferenceMarkers: " & totalFn & _
                " footnote(s); scanning Footnotes story for """ & FR_STYLE & """ runs..."

    Set probe = story.Duplicate
    With probe.Find
        .ClearFormatting
        .Text = ""
        .style = oDoc.Styles(FR_STYLE)
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchWildcards = False
    End With

    ' Word's auto-numbered footnote marker sits at Range.Start - 1 in the
    ' Footnotes story (one character before the footnote body proper).
    ' MARKER_GAP covers that boundary so the marker is classified as
    ' belonging to its footnote rather than counted as an orphan.
    Const MARKER_GAP As Long = 5
    ' Avoid flooding Immediate when classification is broken: at most
    ' MAX_PRINT orphan lines are printed verbatim. Beyond that, the
    ' summary Count still includes all of them.
    Const MAX_PRINT  As Long = 20

    Do While probe.Find.Execute
        totalFR = totalFR + 1
        runStart = probe.Start
        inFn = -1
        For i = 1 To totalFn
            If runStart >= (fnStart(i) - MARKER_GAP) And runStart < fnEnd(i) Then
                inFn = i
                Exit For
            End If
        Next i
        If inFn = -1 Then
            orphans = orphans + 1
            If orphans <= MAX_PRINT Then
                snip = probe.Text
                If Len(snip) > MAX_SNIP Then snip = Left$(snip, MAX_SNIP) & " ..."
                snip = Replace(Replace(snip, vbCr, " | "), vbLf, " | ")
                Debug.Print "  ORPHAN FR run at Story.Pos=" & runStart & _
                            "  text=[" & snip & "]"
            ElseIf orphans = MAX_PRINT + 1 Then
                Debug.Print "  ... (additional orphans suppressed; see total in summary)"
            End If
        Else
            insideCnt = insideCnt + 1
            frPerFn(inFn) = frPerFn(inFn) + 1
        End If
        probe.Collapse wdCollapseEnd
    Loop

    Debug.Print "AuditFootnoteReferenceMarkers: total FR runs in Footnotes story=" & _
                totalFR & "  (inside footnote.Range=" & insideCnt & _
                ", orphans=" & orphans & ")"

    ' Per-footnote anomaly check: any footnote whose FR-marker Count
    ' inside its bounds is not exactly 1 is anomalous. The single duplicate
    ' that produced the 2001-vs-2000 surplus surfaces here.
    For i = 1 To totalFn
        If frPerFn(i) <> 1 Then
            anomalies = anomalies + 1
            On Error Resume Next
            pageNum = -1
            pageNum = oDoc.Footnotes(i).Reference.Information(wdActiveEndPageNumber)
            snip = ""
            snip = oDoc.Footnotes(i).Range.Text
            On Error GoTo 0
            If Len(snip) > MAX_SNIP Then snip = Left$(snip, MAX_SNIP) & " ..."
            snip = Replace(Replace(snip, vbCr, " | "), vbLf, " | ")
            Debug.Print "  ANOMALY footnote(" & i & "): FR markers=" & frPerFn(i) & _
                        " (expected 1)  page=" & pageNum & _
                        "  text=[" & snip & "]"
        End If
    Next i

    Debug.Print "AuditFootnoteReferenceMarkers: per-footnote check - " & _
                anomalies & " footnote(s) with FR Count != 1."
    AuditFootnoteReferenceMarkers = orphans + anomalies
End Function

'==============================================================================
' AuditBookHyperlinkStyling
'==============================================================================
' Verify every BookHyperlink-styled run in the document carries the
' expected character properties. Walks all primary StoryRanges and Finds
' runs by character style "BookHyperlink".
'
' For each match, verify:
'   - Font.Name      = "Carlito"
'   - Font.Size      = 9
'   - Font.Color     = ColorFromName("DarkBlue")  (= 8388608 = #000080)
'   - Font.Underline = wdUnderlineSingle
'
' Anomalies are reported per-property: each mismatched property shows
' in the output. Expected 0 after LockBookHyperlinks runs.
'
' BookHyperlink replaced the built-in Hyperlink style as the doc's
' one-form hyperlink target. The built-in inherits font/size from
' paragraph context and so cannot be enforced uniformly; the custom
' style pins all four properties explicitly. See EDSG/01-styles.md
' "Companion rule: no clickable hyperlinks anywhere" for the rule.
'
' Output:
'   One line per anomaly (story, page, mismatch list, run text snippet),
'   then a summary Count.
'
' RETURN:
'   Anomaly Count.
'
' Usage from Immediate:
'   ?AuditBookHyperlinkStyling
'==============================================================================
Public Function AuditBookHyperlinkStyling() As Long
    Const EXPECTED_STYLE As String = "BookHyperlink"
    Const EXPECTED_FONT  As String = "Carlito"
    Const EXPECTED_SIZE  As Single = 9
    Const MAX_SNIP       As Long = 40
    Dim oDoc       As Word.Document
    Dim story      As Word.Range
    Dim probe      As Word.Range
    Dim runColor   As Long
    Dim runUL      As Long
    Dim runFont    As String
    Dim runSize    As Single
    Dim expected   As Long
    Dim pageNum    As Long
    Dim anomalies  As Long
    Dim total      As Long
    Dim snip       As String
    Dim storyName  As String
    Dim mismatch   As String

    Set oDoc = ActiveDocument
    expected = ColorFromName("DarkBlue")

    Debug.Print "AuditBookHyperlinkStyling: scanning all stories for """ & _
                EXPECTED_STYLE & """-styled runs; expecting font=" & EXPECTED_FONT & _
                " " & EXPECTED_SIZE & "pt + color=" & expected & " " & LongToHex(expected) & _
                " + underline single..."

    For Each story In oDoc.StoryRanges
        Set probe = story.Duplicate
        With probe.Find
            .ClearFormatting
            .Text = ""
            .style = oDoc.Styles(EXPECTED_STYLE)
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchWildcards = False
        End With
        Do While probe.Find.Execute
            total = total + 1
            runColor = probe.Font.Color
            runUL = probe.Font.Underline
            runFont = probe.Font.Name
            runSize = probe.Font.Size

            mismatch = ""
            If runFont <> EXPECTED_FONT Then mismatch = mismatch & " font=[" & runFont & "]"
            If runSize <> EXPECTED_SIZE Then mismatch = mismatch & " size=" & runSize
            If runColor <> expected Then mismatch = mismatch & " color=" & runColor & " " & LongToHex(runColor)
            If runUL <> wdUnderlineSingle Then mismatch = mismatch & " underline=" & runUL

            If Len(mismatch) > 0 Then
                anomalies = anomalies + 1
                pageNum = -1
                On Error Resume Next
                pageNum = probe.Information(wdActiveEndPageNumber)
                On Error GoTo 0
                snip = probe.Text
                If Len(snip) > MAX_SNIP Then snip = Left$(snip, MAX_SNIP) & " ..."
                snip = Replace(Replace(snip, vbCr, " | "), vbLf, " | ")
                storyName = "StoryType=" & story.StoryType
                Debug.Print "  ANOMALY " & storyName & _
                            "  page=" & pageNum & _
                            "  mismatch:" & mismatch & _
                            "  text=[" & snip & "]"
            End If
            probe.Collapse wdCollapseEnd
        Loop
    Next story

    Debug.Print "AuditBookHyperlinkStyling: " & total & " " & EXPECTED_STYLE & _
                "-styled run(s) checked, " & anomalies & " anomaly/anomalies."
    AuditBookHyperlinkStyling = anomalies
End Function

'==============================================================================
' AuditNonPaletteStyleColors
'==============================================================================
' Two-tier colour discipline (see EDSG/01-styles.md):
'   Tier 1 - default-text intent : Font.Color = wdColorAutomatic
'   Tier 2 - deliberate colour   : Font.Color = a palette-registered value
'
' Five-bucket classification:
'   1. Tier 1   - Font.Color = wdColorAutomatic
'   2. Tier 2   - Font.Color in palette (NameFromColor returns non-empty)
'   3. Theme    - Font.ObjectThemeColor <> wdThemeColorNone. Office theme
'                 colour reference. Reported separately - banned by the
'                 rule but not the same as a hand-typed off-palette
'                 anomaly. Many Word built-ins (Heading 4-9, Caption,
'                 Quote, etc.) carry theme colours by default.
'   4. Anomaly  - none of the above. The actual editorial concern: a
'                 hand-typed off-palette RGB value. Return-value
'                 contribution.
'   5. Skipped  - Table or List style (Font.Color not meaningful),
'                 Font.Color read errored, or built-in style when
'                 IncludeBuiltIn is False.
'
' Return value = Anomaly Count only (the editorial-discipline assertion).
' Theme-colour Count is informational and addressed by a separate audit
' / hide-sweep workflow on built-ins.
'
' Color display Note: when Font.Color is outside [0, 0xFFFFFF], the value
' is a sentinel or theme reference, not a real RGB. The output flags this
' explicitly rather than pretending the low bits are RGB triplet.
'
' Optional IncludeBuiltIn (default False): skip Word built-in styles
' (Heading 4-9, Caption, Quote, Mention, etc.). Most built-ins carry
' theme colours unmanaged by editorial; their Theme-bucket noise drowns
' out the custom-style signal. Pass True for a one-off diagnostic that
' covers everything.
'
' Usage from Immediate:
'   ?AuditNonPaletteStyleColors           ' custom + linked only (typical)
'   ?AuditNonPaletteStyleColors True      ' include built-ins too
'==============================================================================
Public Function AuditNonPaletteStyleColors( _
        Optional ByVal IncludeBuiltIn As Boolean = False) As Long
    Const WD_THEME_NONE As Long = -1   ' wdThemeColorNone
    Dim oDoc        As Word.Document
    Dim oStyle      As Word.Style
    Dim c           As Long
    Dim themeIdx    As Long
    Dim paletteName As String
    Dim styleType   As Long
    Dim okColor     As Boolean
    Dim okTheme     As Boolean
    Dim isBuiltIn   As Boolean
    Dim total       As Long
    Dim autoN       As Long
    Dim paletteN    As Long
    Dim themeN      As Long
    Dim anomalyN    As Long
    Dim skippedN    As Long
    Dim biSkippedN  As Long

    Set oDoc = ActiveDocument
    Debug.Print "AuditNonPaletteStyleColors: classifying Font.Color across all styles" & _
                IIf(IncludeBuiltIn, " (including built-ins)", " (custom + linked only)") & "..."
    Debug.Print "Tier 1=Automatic OK; Tier 2=Palette OK; Theme=banned-but-built-in; Anomaly=hand-typed off-palette."

    For Each oStyle In oDoc.Styles
        styleType = oStyle.Type
        If styleType = wdStyleTypeTable Or styleType = wdStyleTypeList Then
            skippedN = skippedN + 1
        Else
            On Error Resume Next
            isBuiltIn = False
            isBuiltIn = oStyle.BuiltIn
            On Error GoTo 0

            If isBuiltIn And Not IncludeBuiltIn Then
                biSkippedN = biSkippedN + 1
            Else
                c = 0
                okColor = False
                themeIdx = WD_THEME_NONE
                okTheme = False
                On Error Resume Next
                c = oStyle.Font.Color
                okColor = (Err.Number = 0)
                Err.Clear
                ' Theme colour lives on Font.TextColor.ObjectThemeColor in
                ' Word's modern object model. The direct Font.ObjectThemeColor
                ' is not a member.
                themeIdx = oStyle.Font.TextColor.ObjectThemeColor
                okTheme = (Err.Number = 0)
                On Error GoTo 0

                If Not okColor Then
                    skippedN = skippedN + 1
                Else
                    total = total + 1
                    If okTheme And themeIdx <> WD_THEME_NONE Then
                        ' Theme-coloured. Report and bucket; do NOT Count in anomalies.
                        themeN = themeN + 1
                        Debug.Print "  THEME    style=[" & oStyle.NameLocal & "]" & _
                                    "  type=[" & StyleTypeName(styleType) & "]" & _
                                    "  builtin=" & isBuiltIn & _
                                    "  ObjectThemeColor=" & themeIdx & _
                                    "  Font.Color=" & c & _
                                    "  InUse=" & oStyle.InUse
                    ElseIf c = wdColorAutomatic Then
                        autoN = autoN + 1
                    Else
                        paletteName = NameFromColor(c)
                        If Len(paletteName) > 0 Then
                            paletteN = paletteN + 1
                        Else
                            anomalyN = anomalyN + 1
                            Debug.Print "  ANOMALY  style=[" & oStyle.NameLocal & "]" & _
                                        "  type=[" & StyleTypeName(styleType) & "]" & _
                                        "  builtin=" & isBuiltIn & _
                                        "  Font.Color=" & c & _
                                        ColorDisplay(c) & _
                                        "  InUse=" & oStyle.InUse
                        End If
                    End If
                End If
            End If
        End If
    Next oStyle

    Debug.Print "---"
    Debug.Print "AuditNonPaletteStyleColors summary: " & total & " styles classified."
    Debug.Print "  Tier 1 - Automatic (default text):    " & autoN
    Debug.Print "  Tier 2 - Palette (deliberate colour): " & paletteN
    Debug.Print "  Theme  (Office theme colour, banned): " & themeN
    Debug.Print "  Anomaly (hand-typed off-palette):     " & anomalyN & " <- return value"
    Debug.Print "  Skipped (Table/List/error):           " & skippedN
    Debug.Print "  BuiltIn skipped:                      " & biSkippedN
    AuditNonPaletteStyleColors = anomalyN
End Function

' Display helper: format Font.Color for readable output. When value is
' outside the [0, 0xFFFFFF] RGB range it is a sentinel or theme reference,
' not a real colour - label it as such rather than byte-decomposing the
' low bits and pretending they are RGB.
Private Function ColorDisplay(ByVal c As Long) As String
    If c >= 0 And c <= &HFFFFFF Then
        ColorDisplay = " " & LongToHex(c)
    Else
        ColorDisplay = " (sentinel/theme-encoded)"
    End If
End Function

'==============================================================================
' ListInertHyperlinkStyledRuns
'==============================================================================
' Per-instance dump of every Hyperlink-character-styled run that is NOT
' backed by an active Hyperlink collection object. These are the
' "styled-but-inert" runs surfaced by ReportClickableHyperlinks - text
' that looks like a link (DarkBlue + underline + Hyperlink style) but
' carries no click target.
'
' For each story:
'   - snapshot [Start, End] bounds of every collection Hyperlink.
'   - Find each Hyperlink-styled run.
'   - If the run's position is not inside any collection-Hyperlink's
'     range, it is inert -> print: story, page, run text (snippet),
'     surrounding context (CTX chars on each side).
'
' Use the output to decide for each inert run:
'   - Restyle: strip the Hyperlink character style (text becomes plain
'     body text).
'   - Leave: keep the visible-as-link styling as a deliberate emphasis.
'   - Repoint: if a link is desired, add a Hyperlink object back.
'
' Usage from Immediate:
'   ListInertHyperlinkStyledRuns
' ==========================================================================
Public Sub ListInertHyperlinkStyledRuns()
    Const EXPECTED_STYLE As String = "Hyperlink"
    Const MAX_SNIP       As Long = 60
    Const CTX_PAD        As Long = 40
    Const MAX_PRINT      As Long = 50
    Dim oDoc       As Word.Document
    Dim story      As Word.Range
    Dim probe      As Word.Range
    Dim ctx        As Word.Range
    Dim hl         As Word.Hyperlink
    Dim hlStarts() As Long
    Dim hlEnds()   As Long
    Dim hlCount    As Long
    Dim runStart   As Long
    Dim i          As Long
    Dim inertN     As Long
    Dim total      As Long
    Dim isActive   As Boolean
    Dim pageNum    As Long
    Dim snip       As String
    Dim ctxStart   As Long
    Dim ctxEnd     As Long
    Dim ctxText    As String

    Set oDoc = ActiveDocument
    Debug.Print "ListInertHyperlinkStyledRuns: enumerating Hyperlink-styled runs not backed by a collection-Hyperlink..."

    For Each story In oDoc.StoryRanges
        ' Snapshot active-Hyperlink bounds in this story.
        hlCount = 0
        On Error Resume Next
        hlCount = story.Hyperlinks.Count
        On Error GoTo 0
        If hlCount > 0 Then
            ReDim hlStarts(1 To hlCount)
            ReDim hlEnds(1 To hlCount)
            i = 0
            For Each hl In story.Hyperlinks
                i = i + 1
                hlStarts(i) = hl.Range.Start
                hlEnds(i) = hl.Range.End
            Next hl
        End If

        ' Walk Hyperlink-styled runs.
        Set probe = story.Duplicate
        With probe.Find
            .ClearFormatting
            .Text = ""
            .style = oDoc.Styles(EXPECTED_STYLE)
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchWildcards = False
        End With
        Do While probe.Find.Execute
            total = total + 1
            runStart = probe.Start
            isActive = False
            For i = 1 To hlCount
                If runStart >= hlStarts(i) And runStart < hlEnds(i) Then
                    isActive = True
                    Exit For
                End If
            Next i
            If Not isActive Then
                inertN = inertN + 1
                If inertN <= MAX_PRINT Then
                    On Error Resume Next
                    pageNum = -1
                    pageNum = probe.Information(wdActiveEndPageNumber)
                    On Error GoTo 0
                    snip = probe.Text
                    If Len(snip) > MAX_SNIP Then snip = Left$(snip, MAX_SNIP) & " ..."
                    snip = Replace(Replace(snip, vbCr, " | "), vbLf, " | ")

                    ctxStart = probe.Start - CTX_PAD
                    If ctxStart < story.Start Then ctxStart = story.Start
                    ctxEnd = probe.End + CTX_PAD
                    If ctxEnd > story.End Then ctxEnd = story.End
                    Set ctx = story.Duplicate
                    ctx.SetRange ctxStart, ctxEnd
                    ctxText = ctx.Text
                    ctxText = Replace(Replace(ctxText, vbCr, " | "), vbLf, " | ")

                    Debug.Print "  INERT  story=" & StoryRangeName(story.StoryType) & _
                                "  page=" & pageNum & _
                                "  pos=" & runStart & _
                                "  text=[" & snip & "]" & vbCrLf & _
                                "         ctx=...[" & ctxText & "]..."
                ElseIf inertN = MAX_PRINT + 1 Then
                    Debug.Print "  ... (additional inert runs suppressed; summary Count still includes them)"
                End If
            End If
            probe.Collapse wdCollapseEnd
        Loop
    Next story

    Debug.Print "---"
    Debug.Print "ListInertHyperlinkStyledRuns: total Hyperlink-styled runs=" & total & _
                "  inert (no active link)=" & inertN
End Sub

'==============================================================================
' ReportClickableHyperlinks
'==============================================================================
' Read-only diagnostic for the no-clickable-hyperlinks rule.
'
' Editorial rule: every hyperlink in the doc must be non-clickable. Print
' is the primary target; online interactivity is a future-mode concern.
' "Hyperlink" in this doc means exactly one thing: a web URL pointing to
' an online concordance tool, displayed as Hyperlink-character-styled
' text + DarkBlue + underline. Some are still backed by active Hyperlink
' objects (clickable). Most are inert text-with-styling (the link object
' was removed but the styling stayed).
'
' This probe answers:
'   1. Per story: how many active Hyperlinks (clickable) vs how many
'      Hyperlink-styled runs (visible-as-link).
'   2. For each active Hyperlink: Address, SubAddress (bookmark, if
'      internal), display text, run style, and the page.
'   3. Are any internal SubAddress bookmarks dangling (target deleted)?
'      Surface them so they can be reviewed.
'   4. Any non-Hyperlink fields whose Result is styled Hyperlink (would
'      indicate REF/PAGEREF/etc. that aren't supposed to be in this doc).
'
' Use the output to decide:
'   - Which active Hyperlinks should be unlinked (text + style preserved,
'     click target removed).
'   - Whether any active Hyperlinks should be deleted entirely.
'   - The expected post-cleanup value for the test 17 audit (likely 0
'     across all stories).
'
' Usage from Immediate:
'   ReportClickableHyperlinks
'==============================================================================
Public Sub ReportClickableHyperlinks()
    Const EXPECTED_STYLE As String = "Hyperlink"
    Const MAX_SNIP       As Long = 60
    Dim oDoc        As Word.Document
    Dim story       As Word.Range
    Dim probe       As Word.Range
    Dim hl          As Word.Hyperlink
    Dim fld         As Word.Field
    Dim totalHL     As Long
    Dim totalStyled As Long
    Dim totalStyledFields As Long
    Dim totalDangling As Long
    Dim storyHL     As Long
    Dim storyStyled As Long
    Dim StyleName   As String
    Dim subTarget   As String
    Dim isInternal  As Boolean
    Dim hasTarget   As Boolean
    Dim snip        As String
    Dim pageNum     As Long
    Dim i           As Long

    Set oDoc = ActiveDocument
    Debug.Print "ReportClickableHyperlinks: scoping the no-clickable rule..."
    Debug.Print "Editorial: '" & EXPECTED_STYLE & "'-styled web-URL links only; no REF/PAGEREF expected."
    Debug.Print "---"

    For Each story In oDoc.StoryRanges
        storyHL = 0
        On Error Resume Next
        storyHL = story.Hyperlinks.Count
        On Error GoTo 0

        Set probe = story.Duplicate
        With probe.Find
            .ClearFormatting
            .Text = ""
            .style = oDoc.Styles(EXPECTED_STYLE)
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchWildcards = False
        End With
        storyStyled = 0
        Do While probe.Find.Execute
            storyStyled = storyStyled + 1
            probe.Collapse wdCollapseEnd
        Loop

        If storyHL > 0 Or storyStyled > 0 Then
            Debug.Print "Story " & StoryRangeName(story.StoryType) & _
                        "  ActiveHyperlinks=" & storyHL & _
                        "  HyperlinkStyledRuns=" & storyStyled & _
                        "  StyledButInert=" & (storyStyled - storyHL)
        End If
        totalHL = totalHL + storyHL
        totalStyled = totalStyled + storyStyled

        ' Per-active-Hyperlink detail
        If storyHL > 0 Then
            For Each hl In story.Hyperlinks
                On Error Resume Next
                pageNum = -1
                pageNum = hl.Range.Information(wdActiveEndPageNumber)
                StyleName = ""
                StyleName = CStr(hl.Range.style.NameLocal)
                On Error GoTo 0
                snip = hl.TextToDisplay
                If Len(snip) > MAX_SNIP Then snip = Left$(snip, MAX_SNIP) & " ..."
                snip = Replace(Replace(snip, vbCr, " | "), vbLf, " | ")

                subTarget = ""
                On Error Resume Next
                subTarget = hl.SubAddress
                On Error GoTo 0
                isInternal = (Len(hl.Address) = 0 And Len(subTarget) > 0)

                ' For internal links, check whether the bookmark target exists.
                hasTarget = True
                If isInternal Then
                    hasTarget = False
                    On Error Resume Next
                    hasTarget = oDoc.Bookmarks.Exists(subTarget)
                    On Error GoTo 0
                    If Not hasTarget Then totalDangling = totalDangling + 1
                End If

                Debug.Print "  HL page=" & pageNum & _
                            "  style=[" & StyleName & "]" & _
                            "  addr=[" & hl.Address & "]" & _
                            "  sub=[" & subTarget & "]" & _
                            IIf(isInternal, IIf(hasTarget, "  bookmark=OK", "  bookmark=DANGLING"), "") & _
                            "  text=[" & snip & "]"
            Next hl
        End If

        ' Probe for any Hyperlink-styled-Result fields (REF/PAGEREF/etc).
        ' Expected zero per the doc's "URL only" rule; report if any
        ' exist so they can be reviewed.
        On Error Resume Next
        For i = 1 To story.Fields.Count
            Set fld = story.Fields(i)
            On Error Resume Next
            StyleName = ""
            StyleName = CStr(fld.Result.style.NameLocal)
            On Error GoTo 0
            If StyleName = EXPECTED_STYLE And fld.Type <> wdFieldHyperlink Then
                totalStyledFields = totalStyledFields + 1
                Debug.Print "  UNEXPECTED FIELD  story=" & StoryRangeName(story.StoryType) & _
                            "  fieldType=" & fld.Type & _
                            "  Result=[" & Left$(fld.Result.Text, MAX_SNIP) & "]"
            End If
        Next i
        On Error GoTo 0
    Next story

    Debug.Print "---"
    Debug.Print "Summary:"
    Debug.Print "  Total ActiveHyperlinks across all stories : " & totalHL
    Debug.Print "  Total Hyperlink-styled runs               : " & totalStyled
    Debug.Print "  StyledButInert (= styled - active)        : " & (totalStyled - totalHL)
    Debug.Print "  Dangling internal-bookmark Hyperlinks     : " & totalDangling
    Debug.Print "  Unexpected styled-Result fields           : " & totalStyledFields
    Debug.Print "  Rule target after cleanup: ActiveHyperlinks = 0, StyledButInert = " & totalStyled
End Sub

'==============================================================================
' ReportHyperlinkStoryDistribution
'==============================================================================
' Diagnostic: for each StoryRange, print
'   (a) the Count of Hyperlinks collection entries, and
'   (b) the Count of Hyperlink-character-styled runs (via Find).
'
' The two counts can differ - real Hyperlinks objects always carry the
' Hyperlink style, but Hyperlink-styled REF / HYPERLINK field-Result
' runs (e.g. concordance navigation) carry the style without being in
' the Hyperlinks collection. The gap tells you how much of the
' style-discipline picture the Hyperlinks collection misses.
'
' Usage from Immediate:
'   ReportHyperlinkStoryDistribution
'==============================================================================
Public Sub ReportHyperlinkStoryDistribution()
    Const EXPECTED_STYLE As String = "Hyperlink"
    Dim oDoc        As Word.Document
    Dim story       As Word.Range
    Dim probe       As Word.Range
    Dim collCount   As Long
    Dim styleCount  As Long
    Dim totalColl   As Long
    Dim totalStyle  As Long
    Dim storyName   As String

    Set oDoc = ActiveDocument
    Debug.Print "ReportHyperlinkStoryDistribution: per-story counts of Hyperlinks collection vs Hyperlink-styled runs"

    For Each story In oDoc.StoryRanges
        collCount = 0
        On Error Resume Next
        collCount = story.Hyperlinks.Count
        On Error GoTo 0

        Set probe = story.Duplicate
        With probe.Find
            .ClearFormatting
            .Text = ""
            .style = oDoc.Styles(EXPECTED_STYLE)
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchWildcards = False
        End With
        styleCount = 0
        Do While probe.Find.Execute
            styleCount = styleCount + 1
            probe.Collapse wdCollapseEnd
        Loop

        If collCount > 0 Or styleCount > 0 Then
            storyName = "StoryType=" & story.StoryType
            Debug.Print "  " & Left$(storyName & String(24, " "), 24) & _
                        "Hyperlinks.Count=" & collCount & _
                        "  Hyperlink-styled runs=" & styleCount
        End If
        totalColl = totalColl + collCount
        totalStyle = totalStyle + styleCount
    Next story

    Debug.Print "  TOTAL across stories: collection=" & totalColl & _
                "  styled runs=" & totalStyle
End Sub

'==============================================================================
' HideUnapprovedBuiltInStyles  (Public)
' ----------------------------------------------------------------------------
' Hide-sweep companion to AuditNonPaletteStyleColors. Closes Item 13 / 2.1.
'
' Walks ActiveDocument.Styles. For every style where BuiltIn = True AND the
' name is NOT in the approved-styles SSOT (basTEST_aeBibleConfig.
' GetApprovedStyles), sets the three-property hide pattern:
'
'   .Priority         = 99
'   .QuickStyle       = False
'   .UnhideWhenUsed   = False
'
' The three-property pattern (not just Priority) matters because
' UnhideWhenUsed = True re-surfaces a style in the gallery the moment any
' run touches it - including paste operations.
'
' Built-in styles that ARE approved (left alone by this sweep):
'   "Normal", "Title", "Heading 1", "Heading 2",
'   "Footnote Reference", "Footnote Text"
' All other built-ins (Hyperlink, FollowedHyperlink, Caption, TOC 1-9,
' Body Text, List Paragraph, the 122+ noise styles, etc.) are hidden.
' Custom (BuiltIn=False) styles are untouched - they are governed by
' AuditNonPaletteStyleColors and the rest of the editorial discipline.
'
' Idempotent: re-running is a no-op for already-hidden built-ins (they
' move into the "already hidden" tally).
'
' Some Word built-ins reject property writes (locked / read-only). Those
' are caught per-style and counted as skipped; the sweep does not abort.
'
' Reports to Immediate:
'   n hidden (newly), n already hidden, n skipped (locked),
'   plus a list of newly-hidden names for first-run verification.
'==============================================================================
Public Sub HideUnapprovedBuiltInStyles()
    On Error GoTo PROC_ERR
    Dim approved As Variant
    Dim approvedDict As Object
    Dim i As Long
    approved = GetApprovedStyles()
    Set approvedDict = CreateObject("Scripting.Dictionary")
    approvedDict.CompareMode = 1 ' TextCompare (case-insensitive)
    For i = LBound(approved) To UBound(approved)
        If Not approvedDict.Exists(CStr(approved(i))) Then
            approvedDict.Add CStr(approved(i)), True
        End If
    Next i

    Dim s As Word.Style
    Dim nHidden As Long, nAlready As Long, nSkipped As Long
    Dim newlyHidden As String
    Dim wasHidden As Boolean
    Dim writeErr As Long

    For Each s In ActiveDocument.Styles
        If s.BuiltIn Then
            If Not approvedDict.Exists(s.NameLocal) Then
                wasHidden = (s.Priority = 99) And (s.QuickStyle = False) And (s.UnhideWhenUsed = False)
                writeErr = 0
                On Error Resume Next
                s.Priority = 99
                writeErr = writeErr Or Err.Number
                Err.Clear
                s.QuickStyle = False
                writeErr = writeErr Or Err.Number
                Err.Clear
                s.UnhideWhenUsed = False
                writeErr = writeErr Or Err.Number
                Err.Clear
                On Error GoTo PROC_ERR

                If writeErr <> 0 Then
                    nSkipped = nSkipped + 1
                ElseIf wasHidden Then
                    nAlready = nAlready + 1
                Else
                    nHidden = nHidden + 1
                    newlyHidden = newlyHidden & "  - " & s.NameLocal & vbCrLf
                End If
            End If
        End If
    Next s

    Debug.Print "HideUnapprovedBuiltInStyles:"
    Debug.Print "  newly hidden  : " & nHidden
    Debug.Print "  already hidden: " & nAlready
    Debug.Print "  skipped (locked): " & nSkipped
    If nHidden > 0 Then
        Debug.Print "Newly hidden built-in styles:"
        Debug.Print newlyHidden;
    End If

PROC_EXIT:
    Set approvedDict = Nothing
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure HideUnapprovedBuiltInStyles of Module basStyleInspector"
    Resume PROC_EXIT
End Sub
