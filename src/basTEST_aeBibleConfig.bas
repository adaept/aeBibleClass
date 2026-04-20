Attribute VB_Name = "basTEST_aeBibleConfig"
Option Explicit
Option Compare Text
Option Private Module

'==============================================================================
' basTEST_aeBibleConfig - Configuration for Editing
' ----------------------------------------------------------------------------
' Routines that setup the Word environment for editing.
' Purpose: One top routine that will call others for a consistent experience
' Run manually from the Immediate Window when needed.
'==============================================================================
Public Sub WordEditingConfig()
    ' Add other procedure call as required
    PromoteApprovedStyles
    ' Uncomment this to check priority settings
    DumpPrioritiesSorted
End Sub

Private Sub PromoteApprovedStyles()
    Dim s As style
    Dim approved As Variant
    Dim i As Long
    Dim missing As Collection
    Set missing = New Collection

    'List your approved styles in the order you want them to appear
    approved = Array("Normal", "Body Text", "Heading 1", "Heading 2", _
                     "CustomParaAfterH1", "CustomParaAfterH1-2nd", "Brief", "DatAuthRef", _
                     "Chapter Verse marker", "Verse marker", _
                     "EmphasisBlack", "EmphasisRed", "Lamentation", "Psalms BOOK", _
                     "Words of Jesus", "TheHeaders", "TheFooters", _
                     "Title", "Book Title", _
                     "Footnote Reference", "Footnote Text", "FargleBlargle")

    'Push everything else down
    For Each s In ActiveDocument.Styles
        If s.Type = wdStyleTypeParagraph Or s.Type = wdStyleTypeCharacter Then
            s.Priority = 99
        End If
    Next s

    'Promote approved styles + diagnostic guard
    For i = LBound(approved) To UBound(approved)
        On Error Resume Next
        Set s = ActiveDocument.Styles(approved(i))
        On Error GoTo 0

        If s Is Nothing Then
            missing.Add approved(i)
        Else
            s.Priority = i + 1
        End If

        Set s = Nothing
    Next i

    'Report missing styles
    If missing.Count > 0 Then
        Dim msg As String
        msg = "WARNING: The following styles were NOT found:" & vbCrLf

        For i = 1 To missing.Count
            msg = msg & " -> " & missing(i) & vbCrLf
        Next i

        'MsgBox msg, vbExclamation, "PromoteApprovedStyles Diagnostics"
        Debug.Print msg & " style is missing!"
    End If

    Debug.Print "PromoteApprovedStyles: Done!"
End Sub

Private Sub DumpPrioritiesSorted()
    Dim s As style
    Dim arr() As Variant
    Dim Count As Long
    Dim i As Long, j As Long
    Dim tmpName As String, tmpPri As Long

    'First pass: Count eligible styles
    For Each s In ActiveDocument.Styles
        If s.Type = wdStyleTypeParagraph Or s.Type = wdStyleTypeCharacter Then
            Count = Count + 1
        End If
    Next s

    'Allocate array: 1-based, 2 columns (Name, Priority)
    ReDim arr(1 To Count, 1 To 2)

    'Second pass: fill array
    Count = 1
    For Each s In ActiveDocument.Styles
        If s.Type = wdStyleTypeParagraph Or s.Type = wdStyleTypeCharacter Then
            arr(Count, 1) = s.NameLocal
            arr(Count, 2) = s.Priority
            Count = Count + 1
        End If
    Next s

    'Sort array by Priority ascending (simple bubble sort, fast enough for <500 styles)
    For i = LBound(arr, 1) To UBound(arr, 1) - 1
        For j = i + 1 To UBound(arr, 1)
            If arr(j, 2) < arr(i, 2) Then
                'swap
                tmpName = arr(i, 1)
                tmpPri = arr(i, 2)

                arr(i, 1) = arr(j, 1)
                arr(i, 2) = arr(j, 2)

                arr(j, 1) = tmpName
                arr(j, 2) = tmpPri
            End If
        Next j
    Next i

    'Print sorted results
    Debug.Print "---- Sorted by Priority ----"
    For i = LBound(arr, 1) To UBound(arr, 1)
        If arr(i, 2) <> 99 Then
            Debug.Print arr(i, 1) & "  ->  " & arr(i, 2)
        End If
    Next i
End Sub

Public Sub TestInvisible()
    Dim s As String
    s = CountInvisibleCharacters()
    Debug.Print "[" & s & "]"
    MsgBox "[" & s & "]"
End Sub

Private Function CountInvisibleCharacters(Optional doc As Document) As String
    Dim r As Word.Range
    Dim targets As Variant
    Dim labels As Variant
    Dim counts() As Long
    Dim i As Long
    Dim total As Long
    Dim report As String

    If doc Is Nothing Then Set doc = ActiveDocument

    ' Default return value
    CountInvisibleCharacters = "0"

    targets = Array(ChrW(&H200B), ChrW(&H200C), ChrW(&H200D), ChrW(&HFEFF), ChrW(&H2060))
    labels = Array( _
        "U+200B ZERO WIDTH SPACE", _
        "U+200C ZERO WIDTH NON-JOINER", _
        "U+200D ZERO WIDTH JOINER", _
        "U+FEFF ZERO WIDTH NO-BREAK SPACE", _
        "U+2060 WORD JOINER")

    ReDim counts(UBound(targets))

    ' Count per story, per character
    For Each r In doc.StoryRanges
        For i = 0 To UBound(targets)
            counts(i) = counts(i) + UBound(Split(r.Text, targets(i)))
        Next i
    Next r

    ' Sum total
    For i = 0 To UBound(counts)
        total = total + counts(i)
    Next i

    ' If nothing found, keep "0"
    If total = 0 Then Exit Function

    ' Build per-character report
    For i = 0 To UBound(counts)
        If counts(i) > 0 Then
            report = report & labels(i) & ": " & counts(i) & vbCrLf
        End If
    Next i

    CountInvisibleCharacters = Trim$(report)
End Function

Private Function CountOrphanedShapes(doc As Document) As Long
    Dim shp As shape
    Dim Count As Long
    For Each shp In doc.Shapes
        If shp.Anchor Is Nothing Then Count = Count + 1
    Next shp
    CountOrphanedShapes = Count
End Function

Private Function CountOrphanedBookmarks(doc As Document) As Long
    Dim bm As Bookmark
    Dim Count As Long
    For Each bm In doc.Bookmarks
        If bm.Range.Text = "" Then Count = Count + 1
    Next bm
    CountOrphanedBookmarks = Count
End Function

'STAGE 0 - FINAL GOAL (redefined)
'
'Stage 0 is complete when:
'1. All BODY text uses a print-safe (free) font
'2. All TITLE/SERIF text uses a print-safe (free) serif font
'3. Layout shift is minimized (hyphenation preserved as much as possible)
'4. No direct overrides interfere with future layout tuning
'5. Screen-only fonts (Word defaults) are eliminated from print paths

'-> Replacement for Calibri (Body)
'similar metrics
'similar line length
'Minimal reflow
'Best candidate:
'
'-> Carlito
'Property    Value
'License Open (Google / SIL OFL)
'Metric compatibility    - Designed to match Calibri
'Hyphenation impact  Minimal

'-> Replacement for Times New Roman (Serif)
'classic book serif
'Close metrics
'stable print behavior
'Best candidates:
'Option A (closest match):
'
'-> Liberation Serif
'Property    Value
'Metric compatibility    - Times-compatible
'License Open (SIL)
'Layout shift    Minimal

