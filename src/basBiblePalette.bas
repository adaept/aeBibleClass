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
' Name -> PaletteColor record. The record carries every common
' representation (R, G, B, RgbLong, HexCode) so callers do not have to
' convert. A Usage field documents where each color is applied in the
' production document, so a future palette swap (light / dark /
' colorblind theme) can audit impact before committing.
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
'   GetPalette(theme)             -> Dictionary of PaletteColor
'   ColorFromName(name)           -> RgbLong   (raises if unknown)
'   NameFromColor(rgbLong)        -> Name      ("" if unknown)
'   LongToHex(rgbLong)            -> "#RRGGBB"
'   HexToLong(hex)                -> RgbLong
'   LongToRgbString(rgbLong)      -> "(R,G,B)"
'
' Theme arg is "Default" today. "Dark" and "Colorblind" raise
' "Not implemented" so call sites can be wired now and the palettes
' populated later without an API change.
'
' Late binding throughout (Scripting.Dictionary via CreateObject).
' No project references added.
' ==========================================================================

Public Type PaletteColor
    Name     As String
    R        As Long
    G        As Long
    B        As Long
    RgbLong  As Long     ' = RGB(R, G, B)
    HexCode  As String   ' = "#RRGGBB"
    Usage    As String   ' brief note on where this color appears in the doc
End Type

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

    AddColor d, "Black",     0,   0,   0,   "Explicit black. Distinct from wdColorAutomatic (sentinel)."
    AddColor d, "White",     255, 255, 255, "Empty-paragraph detection (aeBibleClass)."
    AddColor d, "Red",       255, 0,   0,   "CountRedFootnoteReferences probe color (legacy / research item)."
    AddColor d, "DarkRed",   128, 0,   0,   "Words of Jesus, EmphasisRed character styles."
    AddColor d, "Green",     0,   255, 0,   "Palette only - not currently applied in the production docx."
    AddColor d, "DarkGreen", 0,   100, 0,   "Palette only - not currently applied in the production docx."
    AddColor d, "Emerald",   80,  200, 120, "Verse marker character style."
    AddColor d, "Blue",      0,   0,   255, "Palette only. NOTE: basTEST_aeBibleConfig audits 'Footnote Reference' at 16711680 (this value) but Module1.EnsureFootnoteReferenceStyleColor sets it to Purple - see research task."
    AddColor d, "Gold",      255, 215, 0,   "Palette only - not currently applied in the production docx."
    AddColor d, "Orange",    255, 165, 0,   "Chapter Verse marker character style (semantic role: ChapterVerseOrange)."
    AddColor d, "Purple",    102, 51,  153, "Footnote Reference character style (semantic role: FootnotePurple). Rebecca purple."
    AddColor d, "Gray",      128, 128, 128, "Palette only - not currently applied in the production docx."

    Set BuildDefaultPalette = d
End Function

Private Sub AddColor(ByVal d As Object, ByVal name As String, _
                     ByVal r As Long, ByVal g As Long, ByVal b As Long, _
                     ByVal usage As String)
    Dim pc As PaletteColor
    pc.Name = name
    pc.R = r
    pc.G = g
    pc.B = b
    pc.RgbLong = RGB(r, g, b)
    pc.HexCode = "#" & PadHex(r) & PadHex(g) & PadHex(b)
    pc.Usage = usage
    d.Add name, pc
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
            "Unknown palette color '" & name & "'. Call GetPalette to inspect available names."
    End If
    Dim pc As PaletteColor
    pc = d(name)
    ColorFromName = pc.RgbLong
End Function

' RgbLong -> Name. Returns "" when the value is not in the palette.
' Use this for audit logs and histograms where unknown colors are expected
' (existing legacy content) and a missing name is information, not an error.
Public Function NameFromColor(ByVal rgbLong As Long) As String
    Dim d As Object, k As Variant
    Dim pc As PaletteColor
    Set d = GetPalette()
    For Each k In d.Keys
        pc = d(k)
        If pc.RgbLong = rgbLong Then
            NameFromColor = pc.Name
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
    Dim d As Object, k As Variant
    Dim pc As PaletteColor
    Set d = GetPalette()
    Debug.Print "basBiblePalette (theme=Default, " & d.Count & " colors):"
    For Each k In d.Keys
        pc = d(k)
        Debug.Print "  " & Left$(pc.Name & String(12, " "), 12) & _
                    pc.HexCode & "  " & _
                    Left$(LongToRgbString(pc.RgbLong) & String(16, " "), 16) & _
                    "Long=" & pc.RgbLong & "  " & pc.Usage
    Next k
End Sub
