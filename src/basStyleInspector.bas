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
    sOut = sOut & ".Font.Color = " & oFont.color & NL
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
    End If

    Debug.Print sOut
    If bWriteFile Then WriteStyleDump sStyleName, sOut
End Sub

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

Private Function SafeFileName(ByVal s As String) As String
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
    Dim t      As Double

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
    EndTimer "DumpAllApprovedStyles", t
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
