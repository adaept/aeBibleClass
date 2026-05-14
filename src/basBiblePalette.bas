Attribute VB_Name = "basBiblePalette"
Option Explicit
Option Compare Text

' ==========================================================================
' basBiblePalette
' ==========================================================================
' Single source of truth for the named colors used and allowed in the
' production Bible docx. Replaces the scattered set of color helpers
' previously living in Module1, basTEST_aeBibleTools, and aeBibleClass.
'
' Architecture
' ------------
' GetPalette() returns a Scripting.Dictionary keyed by canonical color
' Name -> nested Scripting.Dictionary carrying the seven fields below.
' Each entry is self-describing so callers can pick any representation
' they need without conversion.
'
' Per-entry fields (case-insensitive, keys are strings):
'   "Name"     - String, canonical palette name
'   "R"        - Long, red component   0..255
'   "G"        - Long, green component 0..255
'   "B"        - Long, blue component  0..255
'   "RgbLong"  - Long, = RGB(R, G, B)  (Word Font.Color value)
'   "HexCode"  - String, "#RRGGBB"
'   "Usage"    - String, where this color appears in the production doc
'
' Nested-dictionary layout (rather than a Public Type record) is required
' by VBA's late-binding rule: UDTs declared in .bas modules cannot be
' stored in a late-bound Scripting.Dictionary - they must live in a class
' module to cross that boundary. Dictionaries do cross cleanly.
'
' wdColorAutomatic is intentionally NOT in the palette. It is a sentinel
' meaning "inherit, will be black in default theme," not a color. Theme
' work depends on body text remaining wdColorAutomatic so the page-
' background inversion does the right thing.
'
' Office theme colors (ObjectThemeColor) are out of scope by design -
' too niche, too template-coupled, and not portable to non-Office
' renderers.
'
' Public helpers
' --------------
'   GetPalette(theme)             -> Dictionary of color entries
'   ColorFromName(name)           -> RgbLong   (raises if unknown)
'   NameFromColor(rgbLong)        -> Name      ("" if unknown)
'   LongToHex(rgbLong)            -> "#RRGGBB"
'   HexToLong(hex)                -> RgbLong
'   LongToRgbString(rgbLong)      -> "(R,G,B)"
'   DumpPalette                   -> diagnostic Debug.Print dump
'   CountRunsWithColor(c)         -> exact total across all stories
'   ReportRunsWithColor(c)        -> per-story breakdown + total
'   ListRunsOfColorByStyle(c)     -> per-run-style breakdown + total
'   DescribeFirstRunOfColor(c)    -> locate first explicit-override run
'   DescribeStylesCarryingColor(c)-> find styles whose Font.Color = c
'
' Theme arg is "Default" today. "Dark" and "Colorblind" raise
' "Not implemented" so call sites can be wired now and the palettes
' populated later without an API change.
'
' Late binding throughout (Scripting.Dictionary via CreateObject).
' No project references added.
' ==========================================================================

Private mPaletteCache As Object   ' Scripting.Dictionary, lazy-built

Public Function GetPalette(Optional ByVal theme As String = "Default") As Object
    On Error GoTo PROC_ERR

    Select Case LCase$(theme)
        Case "default", ""
            If mPaletteCache Is Nothing Then Set mPaletteCache = BuildDefaultPalette()
            Set GetPalette = mPaletteCache
        Case "dark", "colorblind"
            Err.Raise 5, "basBiblePalette.GetPalette", _
                "Theme '" & theme & "' not implemented yet. Only 'Default' is populated."
        Case Else
            Err.Raise 5, "basBiblePalette.GetPalette", _
                "Unknown theme '" & theme & "'. Valid: Default, Dark, Colorblind."
    End Select
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure GetPalette of Module basBiblePalette"
End Function

Private Function BuildDefaultPalette() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1   ' vbTextCompare - name lookups are case-insensitive

    AddColor d, "Black", 0, 0, 0, "Explicit black. Distinct from wdColorAutomatic (sentinel)."
    AddColor d, "White", 255, 255, 255, "Empty-paragraph detection (aeBibleClass)."
    AddColor d, "Red", 255, 0, 0, "CountRedFootnoteReferences probe color (legacy / research item)."
    AddColor d, "DarkRed", 128, 0, 0, "Words of Jesus, EmphasisRed character styles."
    AddColor d, "Green", 0, 255, 0, "Palette only - not currently applied in the production docx."
    AddColor d, "DarkGreen", 0, 100, 0, "Palette only - not currently applied in the production docx."
    AddColor d, "Emerald", 80, 200, 120, "Verse marker character style."
    AddColor d, "Blue", 0, 0, 255, "Footnote Reference character style (confirmed 2026-05-13 by live-doc probe: 296 references at this color)."
    AddColor d, "DarkBlue", 0, 0, 128, "Hyperlink + FollowedHyperlink character styles (print-locked; matches wdColorDarkBlue). Distinct from Blue so audits separate hyperlinks from Footnote References."
    AddColor d, "Gold", 255, 215, 0, "Palette only - not currently applied in the production docx."
    AddColor d, "Orange", 255, 165, 0, "Chapter Verse marker character style (semantic role: ChapterVerseOrange)."
    AddColor d, "Purple", 102, 51, 153, "Palette only - not currently applied in the production docx. Rebecca purple."
    AddColor d, "Gray", 128, 128, 128, "Palette only - not currently applied in the production docx."

    Set BuildDefaultPalette = d
End Function

Private Sub AddColor(ByVal d As Object, ByVal name As String, _
                     ByVal r As Long, ByVal g As Long, ByVal b As Long, _
                     ByVal usage As String)
    Dim entry As Object
    Set entry = CreateObject("Scripting.Dictionary")
    entry.CompareMode = 1
    entry.Add "Name", name
    entry.Add "R", r
    entry.Add "G", g
    entry.Add "B", b
    entry.Add "RgbLong", RGB(r, g, b)
    entry.Add "HexCode", "#" & PadHex(r) & PadHex(g) & PadHex(b)
    entry.Add "Usage", usage
    d.Add name, entry
End Sub

Private Function PadHex(ByVal n As Long) As String
    PadHex = Right$("00" & Hex$(n), 2)
End Function

' ==========================================================================
' Public conversion helpers
' ==========================================================================

' Name -> RgbLong. Raises if the name is not in the palette.
' Use this when applying a known palette color: Font.Color = ColorFromName("Purple").
Public Function ColorFromName(ByVal name As String) As Long
    Dim d As Object
    Set d = GetPalette()
    If Not d.Exists(name) Then
        Err.Raise 5, "basBiblePalette.ColorFromName", _
            "Unknown palette color '" & name & "'. Call DumpPalette to inspect available names."
    End If
    ColorFromName = CLng(d(name)("RgbLong"))
End Function

' RgbLong -> Name. Returns "" when the value is not in the palette.
' Use this for audit logs and histograms where unknown colors are expected
' (existing legacy content) and a missing name is information, not an error.
Public Function NameFromColor(ByVal rgbLong As Long) As String
    Dim d As Object, k As Variant
    Set d = GetPalette()
    For Each k In d.Keys
        If CLng(d(k)("RgbLong")) = rgbLong Then
            NameFromColor = CStr(d(k)("Name"))
            Exit Function
        End If
    Next k
    NameFromColor = ""
End Function

' Long -> "#RRGGBB". Byte-correct (handles Word's BGR-ordered Font.Color
' storage by extracting R, G, B explicitly rather than Hex()-ing the raw
' Long, which produces BGR-order text and is a known bug in
' aeBibleClass.ColorToHex).
Public Function LongToHex(ByVal rgbLong As Long) As String
    Dim r As Long, g As Long, b As Long
    r = rgbLong And &HFF
    g = (rgbLong \ &H100) And &HFF
    b = (rgbLong \ &H10000) And &HFF
    LongToHex = "#" & PadHex(r) & PadHex(g) & PadHex(b)
End Function

' "#RRGGBB" or "RRGGBB" -> RgbLong (via VBA RGB()).
Public Function HexToLong(ByVal hex As String) As Long
    Dim h As String
    h = Replace(hex, "#", "")
    If Len(h) <> 6 Then
        Err.Raise 5, "basBiblePalette.HexToLong", _
            "Expected 6 hex digits (optionally prefixed '#'); got '" & hex & "'."
    End If
    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Mid$(h, 1, 2))
    g = CLng("&H" & Mid$(h, 3, 2))
    b = CLng("&H" & Mid$(h, 5, 2))
    HexToLong = RGB(r, g, b)
End Function

' Long -> "(R,G,B)". Diagnostic / log format.
Public Function LongToRgbString(ByVal rgbLong As Long) As String
    Dim r As Long, g As Long, b As Long
    r = rgbLong And &HFF
    g = (rgbLong \ &H100) And &HFF
    b = (rgbLong \ &H10000) And &HFF
    LongToRgbString = "(" & r & "," & g & "," & b & ")"
End Function

' ==========================================================================
' CountRunsWithColor
' ==========================================================================
' Authoritative Count: how many runs in the document carry the given
' Font.Color? Walks all primary StoryRanges with Find and tallies each
' contiguous match. Returns the total.
'
' Use this when you need an accurate Count and the Word-level histogram
' in basTEST_aeBibleTools.ListAndCountFontColors is undercounting because
' coloured single-character runs sit inside mixed-color Words. The
' histogram is fast but approximate; this is slower but exact.
'
' Usage from Immediate:
'   ?CountRunsWithColor(ColorFromName("Blue"))         ' expect ~2000
'   ?CountRunsWithColor(ColorFromName("Orange"))       ' expect N verses
'   ?CountRunsWithColor(ColorFromName("Emerald"))      ' expect N verses
'   ?CountRunsWithColor(RGB(192, 0, 0))                ' #C00000 cleanup target
'                                                      ' (Note: NOT &HC00000 -
'                                                      ' that literal is a
'                                                      ' different colour. The
'                                                      ' histogram's "#C00000"
'                                                      ' display is RGB notation
'                                                      ' R=C0, G=00, B=00.
'                                                      ' Use RGB() or
'                                                      ' HexToLong("#C00000").)
' ==========================================================================
Public Function CountRunsWithColor(ByVal rgbLong As Long) As Long
    On Error GoTo PROC_ERR
    Dim oDoc  As Word.Document
    Dim story As Word.Range
    Dim probe As Word.Range
    Dim n     As Long

    Set oDoc = ActiveDocument
    For Each story In oDoc.StoryRanges
        Set probe = story.Duplicate
        With probe.Find
            .ClearFormatting
            .Text = ""
            .Font.Color = rgbLong
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchWildcards = False
        End With
        Do While probe.Find.Execute
            n = n + 1
            probe.Collapse wdCollapseEnd
        Loop
    Next story
    CountRunsWithColor = n
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure CountRunsWithColor of Module basBiblePalette"
End Function

' ==========================================================================
' ReportRunsWithColor
' ==========================================================================
' Authoritative Count with per-story breakdown. Same Find-based scan as
' CountRunsWithColor but prints one line per story plus a total. Useful
' when you want to see WHERE the runs are (MainText vs Footnotes vs
' Headers etc.), not just how many.
'
' Usage from Immediate:
'   ReportRunsWithColor ColorFromName("Blue")
' ==========================================================================
Public Sub ReportRunsWithColor(ByVal rgbLong As Long)
    On Error GoTo PROC_ERR
    Dim oDoc      As Word.Document
    Dim story     As Word.Range
    Dim probe     As Word.Range
    Dim n         As Long, total As Long
    Dim storyName As String

    Set oDoc = ActiveDocument
    Debug.Print "ReportRunsWithColor: " & LongToRgbString(rgbLong) & _
                " " & LongToHex(rgbLong)

    For Each story In oDoc.StoryRanges
        Set probe = story.Duplicate
        With probe.Find
            .ClearFormatting
            .Text = ""
            .Font.Color = rgbLong
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchWildcards = False
        End With
        n = 0
        Do While probe.Find.Execute
            n = n + 1
            probe.Collapse wdCollapseEnd
        Loop
        If n > 0 Then
            storyName = StoryRangeName(story.StoryType)
            Debug.Print "  " & Left$(storyName & String(20, " "), 20) & n
        End If
        total = total + n
    Next story
    Debug.Print "  " & Left$("TOTAL" & String(20, " "), 20) & total
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ReportRunsWithColor of Module basBiblePalette"
End Sub

' ==========================================================================
' ListRunsOfColorByStyle
' ==========================================================================
' Group every run of the given color by the run's character-style name and
' print the counts. Use this to identify strays: the legitimate styles
' that carry the color will appear with large counts; an outlier style
' with a Count of 1 (or any small number) is a stray candidate.
'
' Walks all primary StoryRanges with Find, just like CountRunsWithColor,
' but tracks the run's Style.NameLocal at each match.
'
' Output shape:
'   ListRunsOfColorByStyle: (0,0,255) #0000FF
'     Footnote Reference   2000
'     Hyperlink              14
'     AuthorBodyText          1   <- the stray's style
'     TOTAL                2015
'
' Usage from Immediate:
'   ListRunsOfColorByStyle ColorFromName("Blue")
' ==========================================================================
Public Sub ListRunsOfColorByStyle(ByVal rgbLong As Long)
    On Error GoTo PROC_ERR
    Dim oDoc      As Word.Document
    Dim story     As Word.Range
    Dim probe     As Word.Range
    Dim styleDict As Object
    Dim StyleName As String
    Dim total     As Long
    Dim k         As Variant

    Set oDoc = ActiveDocument
    Set styleDict = CreateObject("Scripting.Dictionary")
    styleDict.CompareMode = 1   ' case-insensitive

    Debug.Print "ListRunsOfColorByStyle: " & LongToRgbString(rgbLong) & _
                " " & LongToHex(rgbLong)

    For Each story In oDoc.StoryRanges
        Set probe = story.Duplicate
        With probe.Find
            .ClearFormatting
            .Text = ""
            .Font.Color = rgbLong
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchWildcards = False
        End With
        Do While probe.Find.Execute
            StyleName = ""
            On Error Resume Next
            StyleName = CStr(probe.style.NameLocal)
            On Error GoTo PROC_ERR
            If Len(StyleName) = 0 Then StyleName = "(no style)"

            If styleDict.Exists(StyleName) Then
                styleDict(StyleName) = styleDict(StyleName) + 1
            Else
                styleDict.Add StyleName, 1
            End If
            total = total + 1
            probe.Collapse wdCollapseEnd
        Loop
    Next story

    For Each k In styleDict.Keys
        Debug.Print "  " & Left$(CStr(k) & String(28, " "), 28) & styleDict(k)
    Next k
    Debug.Print "  " & Left$("TOTAL" & String(28, " "), 28) & total
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ListRunsOfColorByStyle of Module basBiblePalette"
End Sub

' ==========================================================================
' DescribeFirstRunOfColor
' ==========================================================================
' Diagnostic: locate the first run in the active document whose explicit
' Font.Color matches the given Long, and print its page number, paragraph
' style, character style, run text, and surrounding context to the
' Immediate window. Searches all primary StoryRanges (main body,
' footnotes, endnotes, headers, footers).
'
' Scope: explicit run-level color overrides only. Find with
' .Font.Color = X matches runs that carry the color as a direct override,
' not runs that inherit it through a character or paragraph style.
'
' Note - this is a DIFFERENT scope from
' basTEST_aeBibleTools.ListAndCountFontColors. That routine reads
' Range.Font.Color via ActiveDocument.Words, and Word resolves the style
' chain when you read Font.Color on a Range - so the histogram counts
' the RESOLVED color (effective rendered color), including style-
' inherited values. DescribeFirstRunOfColor will return NOT FOUND for
' colors that appear in the histogram only because a style carries them.
' Use DescribeStylesCarryingColor to find those.
'
' Usage from Immediate:
'   DescribeFirstRunOfColor RGB(127,150,152)   ' #7F9698
'   DescribeFirstRunOfColor RGB(192,0,0)       ' #C00000
'   DescribeFirstRunOfColor &H42495            ' or pass the Long directly
' ==========================================================================
Public Sub DescribeFirstRunOfColor(ByVal rgbLong As Long)
    On Error GoTo PROC_ERR
    Dim oDoc      As Word.Document
    Dim story     As Word.Range
    Dim probe     As Word.Range
    Dim ctx       As Word.Range
    Dim ctxStart  As Long, ctxEnd As Long
    Dim runText   As String, ctxText As String
    Dim parStyle  As String, runStyle As String
    Dim pageNum   As Long
    Dim storyName As String
    Const NL      As String = vbCrLf
    Const MAX_RUN As Long = 80
    Const CTX_PAD As Long = 40

    Set oDoc = ActiveDocument

    Debug.Print "DescribeFirstRunOfColor: searching for " & _
                LongToRgbString(rgbLong) & " " & LongToHex(rgbLong)

    For Each story In oDoc.StoryRanges
        Set probe = story.Duplicate
        With probe.Find
            .ClearFormatting
            .Text = ""
            .Font.Color = rgbLong
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchWildcards = False
        End With

        If probe.Find.Execute Then
            ' Capture before any further range operations.
            runText = probe.Text
            If Len(runText) > MAX_RUN Then runText = Left$(runText, MAX_RUN) & " ..."

            On Error Resume Next
            parStyle = ""
            parStyle = CStr(probe.Paragraphs(1).style.NameLocal)
            runStyle = ""
            runStyle = CStr(probe.style.NameLocal)
            pageNum = -1
            pageNum = probe.Information(wdActiveEndPageNumber)
            On Error GoTo PROC_ERR

            ' Surrounding context: pad on both sides up to CTX_PAD chars,
            ' clamped to story bounds.
            ctxStart = probe.Start - CTX_PAD
            If ctxStart < story.Start Then ctxStart = story.Start
            ctxEnd = probe.End + CTX_PAD
            If ctxEnd > story.End Then ctxEnd = story.End
            On Error Resume Next
            ctxText = ""
            Set ctx = story.Duplicate
            ctx.SetRange ctxStart, ctxEnd
            ctxText = ctx.Text
            On Error GoTo PROC_ERR
            ctxText = Replace(Replace(ctxText, vbCr, " | "), vbLf, " | ")

            storyName = StoryRangeName(story.StoryType)
            Debug.Print "  FOUND in story: " & storyName
            Debug.Print "    Page         : " & pageNum
            Debug.Print "    Paragraph    : style=" & parStyle
            Debug.Print "    Run          : style=" & runStyle
            Debug.Print "    Text         : [" & runText & "]"
            Debug.Print "    Context      : ..." & ctxText & "..."
            Exit Sub
        End If
    Next story

    Debug.Print "  NOT FOUND - no run in any primary story carries " & _
                "explicit Font.Color = " & rgbLong & " " & LongToHex(rgbLong)
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure DescribeFirstRunOfColor of Module basBiblePalette"
End Sub

' ==========================================================================
' DescribeStylesCarryingColor
' ==========================================================================
' Diagnostic complement to DescribeFirstRunOfColor: walks
' ActiveDocument.Styles and prints any style whose Font.Color matches the
' given Long. Use this to identify a color that the histogram counts but
' DescribeFirstRunOfColor cannot locate - i.e., a color applied via style
' chain rather than direct run-level override.
'
' Output per matching style: name, type, base style, color long / hex,
' InUse flag.
'
' Usage from Immediate:
'   DescribeStylesCarryingColor RGB(127,150,152)   ' #7F9698
' ==========================================================================
Public Sub DescribeStylesCarryingColor(ByVal rgbLong As Long)
    On Error GoTo PROC_ERR
    Dim s         As Word.Style
    Dim sColor    As Long
    Dim typeName  As String
    Dim baseName  As String
    Dim matches   As Long

    Debug.Print "DescribeStylesCarryingColor: searching for " & _
                LongToRgbString(rgbLong) & " " & LongToHex(rgbLong)

    For Each s In ActiveDocument.Styles
        On Error Resume Next
        sColor = 0
        sColor = s.Font.Color
        On Error GoTo PROC_ERR

        If sColor = rgbLong Then
            Select Case s.Type
                Case wdStyleTypeParagraph:      typeName = "Paragraph"
                Case wdStyleTypeCharacter:      typeName = "Character"
                Case wdStyleTypeTable:          typeName = "Table"
                Case wdStyleTypeList:           typeName = "List"
                Case wdStyleTypeLinked:         typeName = "Linked"
                Case Else:                      typeName = "Type=" & s.Type
            End Select

            baseName = ""
            On Error Resume Next
            baseName = CStr(s.baseStyle)
            On Error GoTo PROC_ERR
            If Len(baseName) = 0 Then baseName = "(none)"

            Debug.Print "  MATCH: " & s.NameLocal & _
                        "  [" & typeName & "]" & _
                        "  base=" & baseName & _
                        "  Color=" & sColor & " " & LongToHex(sColor) & _
                        "  InUse=" & s.InUse
            matches = matches + 1
        End If
    Next s

    If matches = 0 Then
        Debug.Print "  NOT FOUND - no style carries Font.Color = " & _
                    rgbLong & " " & LongToHex(rgbLong)
    Else
        Debug.Print "  " & matches & " matching style(s)."
    End If
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure DescribeStylesCarryingColor of Module basBiblePalette"
End Sub

Private Function StoryRangeName(ByVal st As WdStoryType) As String
    Select Case st
        Case wdMainTextStory:               StoryRangeName = "MainText"
        Case wdFootnotesStory:              StoryRangeName = "Footnotes"
        Case wdEndnotesStory:               StoryRangeName = "Endnotes"
        Case wdCommentsStory:               StoryRangeName = "Comments"
        Case wdTextFrameStory:              StoryRangeName = "TextFrame"
        Case wdEvenPagesHeaderStory:        StoryRangeName = "EvenHdr"
        Case wdPrimaryHeaderStory:          StoryRangeName = "PrimaryHdr"
        Case wdEvenPagesFooterStory:        StoryRangeName = "EvenFtr"
        Case wdPrimaryFooterStory:          StoryRangeName = "PrimaryFtr"
        Case wdFirstPageHeaderStory:        StoryRangeName = "FirstHdr"
        Case wdFirstPageFooterStory:        StoryRangeName = "FirstFtr"
        Case wdFootnoteSeparatorStory:      StoryRangeName = "FtnSep"
        Case wdFootnoteContinuationSeparatorStory: StoryRangeName = "FtnContSep"
        Case wdFootnoteContinuationNoticeStory:    StoryRangeName = "FtnContNotice"
        Case wdEndnoteSeparatorStory:       StoryRangeName = "EndSep"
        Case wdEndnoteContinuationSeparatorStory:  StoryRangeName = "EndContSep"
        Case wdEndnoteContinuationNoticeStory:     StoryRangeName = "EndContNotice"
        Case Else:                          StoryRangeName = "StoryType=" & st
    End Select
End Function

' ==========================================================================
' DumpPalette
' ==========================================================================
' Diagnostic: dump the current palette to the Immediate window so the
' editor can see what is defined without opening source. Useful before
' running any sweep that will compare run colors against the palette.
'
' Usage from Immediate:
'   DumpPalette
' ==========================================================================
Public Sub DumpPalette()
    Dim d As Object, k As Variant, entry As Object
    Set d = GetPalette()
    Debug.Print "basBiblePalette (theme=Default, " & d.Count & " colors):"
    For Each k In d.Keys
        Set entry = d(k)
        Debug.Print "  " & Left$(CStr(entry("Name")) & String(12, " "), 12) & _
                    CStr(entry("HexCode")) & "  " & _
                    Left$(LongToRgbString(CLng(entry("RgbLong"))) & String(16, " "), 16) & _
                    "Long=" & entry("RgbLong") & "  " & CStr(entry("Usage"))
    Next k
End Sub
