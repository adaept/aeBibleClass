Attribute VB_Name = "basTEST_aeBibleConfig"
Option Explicit
Option Compare Text
Option Private Module

Private m_TaxFile   As Integer
Private m_TaxPass   As Long
Private m_TaxFail   As Long

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
    Dim s As Word.Style
    Dim approved As Variant
    Dim i As Long
    Dim missing As Collection
    Set missing = New Collection

    'List your approved styles in the order you want them to appear
    approved = Array( _
                     "TheHeaders", "BodyText", "TheFooters", _
                     "FrontPageTopLine", "TitleEyebrow", "Title", "TitleVersion", "FrontPageBodyText", _
                     "BodyTextTopLineCPBB", "Acknowledgments", "AuthorBodyText", _
                     "Contents", "ContentsRef", _
                     "BibleIndexEyebrow", "BibleIndex", "Introduction", _
                     "TitleOnePage", _
                     "AuthorListItem", "AuthorListItemBody", "AuthorListItemTab", _
                     "AuthorBookRefHeader", "AuthorBookRef", "CenterSubText", _
                     "Heading 1", "CustomParaAfterH1", "Brief", "DatAuthRef", _
                     "Heading 2", "Chapter Verse marker", "Verse marker", _
                     "Footnote Reference", "Footnote Text", "Psalms BOOK", _
                     "PsalmSuperscription", "Selah", "PsalmAcrostic", _
                     "SpeakerLabel", _
                     "BodyTextIndent", "BodyTextContinuation", _
                     "BookIntro", _
                     "AppendixTitle", "AppendixBody", _
                     "EmphasisBlack", "EmphasisRed", _
                     "Words of Jesus", _
                     "AuthorSectionHead", _
                     "AuthorQuote", _
                     "Normal", _
                     "FargleBlargle")

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
    Dim s As Word.Style
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

Public Function CountInvisibleCharacters(Optional doc As Document) As String
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

'==============================================================================
' RUN_TAXONOMY_STYLES / AuditOneStyle
' PURPOSE:
'   Audits 21 styles via AuditOneStyle + 2 tab-stop specs via AuditStyleTabs;
'   total 23 checks. Writes a structured report to rpt\StyleTaxonomyAudit.txt.
'   Style audit buckets (21):
'     10 fully specified (all properties verified) - BodyText, BodyTextIndent,
'                            Heading 1, Heading 2, CustomParaAfterH1, DatAuthRef,
'                            Brief, Psalms BOOK, Footnote Text, AuthorBookRef
'      8 existence-verified (full spec pending)
'      3 not yet created (expected FAIL until each Define* routine runs)
'   Tab-stop audits (2):
'      AuthorListItemTab (2 stops at 144 / 252 pt, Left, Spaces)
'      AuthorBookRef     (2 stops at 36 / 378 pt; Left+Spaces / Right+Dots)
'   Specs encoded as descriptive (capture current document state); see
'   rvw/Code_review 2026-04-25.md "Spec promotion: descriptive vs prescriptive"
'   for the decision and rationale.
'
' DESIGN:
'   AuditOneStyle (Private) checks up to 7 properties per style. Sentinel
'   values suppress individual checks where the spec is not yet defined:
'     sExpFont        = ""    -> skip font-name check
'     dExpSize        = 0     -> skip font-size check
'     lExpAlign       = -1   -> skip alignment check  (wdAlignParagraphJustify=3)
'     dExpFirstIndent = -999 -> skip first-indent check
'     lExpLineRule    = -1   -> skip line-spacing-rule check (wdLineSpaceExactly=4)
'     dExpLineSpacing = -999 -> skip line-spacing point value (pair with Exactly rule)
'     dExpSpaceBefore = -999 -> skip space-before check
'     dExpSpaceAfter  = -999 -> skip space-after check
'
' RERUN SAFE: overwrites rpt\StyleTaxonomyAudit.txt each run.
' RUN:        RUN_TAXONOMY_STYLES  (Immediate Window or Ribbon)
'==============================================================================
Public Sub RUN_TAXONOMY_STYLES()
    On Error GoTo PROC_ERR
    Dim sPath As String

    sPath = ActiveDocument.Path & "\rpt\StyleTaxonomyAudit.txt"
    m_TaxFile = FreeFile
    m_TaxPass = 0
    m_TaxFail = 0

    Open sPath For Output As #m_TaxFile

    Print #m_TaxFile, "=== Style Taxonomy Audit ==="
    Print #m_TaxFile, "Date    : " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    Print #m_TaxFile, "Document: " & ActiveDocument.Name
    Print #m_TaxFile, String(72, "=")

    ' -- Fully specified styles (all properties verified) -----------------------------------
    Print #m_TaxFile, ""
    Print #m_TaxFile, "-- Fully specified (all properties verified) --"
    '                                   Font          Sz  Align  Indent  LineRule  LineSp  SpB   SpA
    '                                                 0=skip -1=skip -999=skip
    AuditOneStyle "BodyText", "Carlito", 9, 3, 0, 4, 10, 0, 0
    AuditOneStyle "BodyTextIndent", "Carlito", 9, 3, 14.4, 4, 10, 0, 0
    AuditOneStyle "Heading 1", "Noto Sans", 24, 1, 0, 0, 12, 144, 0
    AuditOneStyle "Heading 2", "Noto Sans", 8, 1, 0, 4, 10, 12, 8
    AuditOneStyle "CustomParaAfterH1", "Noto Sans", 10, 1, 0, 4, 10, 0, 62
    AuditOneStyle "DatAuthRef", "Noto Sans", 8, 1, 0, 0, 12, 11, 0
    AuditOneStyle "Brief", "Noto Sans", 10, 1, 0, 4, 9.5, 0, 0
    AuditOneStyle "Psalms BOOK", "Carlito", 9, 0, 14.4, 4, 10, 10, 0
    AuditOneStyle "Footnote Text", "Carlito", 7, 3, 0, 4, 8, 0, 0
    AuditOneStyle "AuthorBookRef", "Carlito", 11, 0, -18, 0, 12, 0, 11

    ' -- Existence verified; full spec pending -------------------------------------------------
    ' Footnote Reference (Character style) parked here until AuditOneStyle is
    ' extended to check character-style Bold / Italic / Color (deferred follow-up).
    Print #m_TaxFile, ""
    Print #m_TaxFile, "-- Existence verified (full spec pending) --"
    AuditOneStyle "BookIntro", "Carlito", 9, 1, 0, 4, 10, 6, 6
    AuditOneStyle "AuthorListItem", "Carlito", 11, 0, 0, -1, -999, 0, 0
    AuditOneStyle "AuthorListItemBody", "Carlito", 11, 0, 0, -1, -999, 0, 11
    AuditOneStyle "AuthorListItemTab", "", 0, -1, -999, -1, -999, -999, -999
    AuditOneStyle "TheHeaders", "", 0, -1, -999, -1, -999, -999, -999
    AuditOneStyle "TheFooters", "", 0, -1, -999, -1, -999, -999, -999
    AuditOneStyle "Title", "", 0, -1, -999, -1, -999, -999, -999
    AuditOneStyle "Footnote Reference", "Carlito", 9, -1, -999, -1, -999, -999, -999

    ' -- Not yet created - expected FAIL until each Define* routine is run ----------------------
    Print #m_TaxFile, ""
    Print #m_TaxFile, "-- Not yet created (expected FAIL) --"
    AuditOneStyle "BodyTextContinuation", "", 0, -1, -999, -1, -999, -999, -999
    AuditOneStyle "AppendixTitle", "", 0, -1, -999, -1, -999, -999, -999
    AuditOneStyle "AppendixBody", "", 0, -1, -999, -1, -999, -999, -999

    ' -- Tab stops verified (per-style explicit tab-stop validation) -----------------------------
    Print #m_TaxFile, ""
    Print #m_TaxFile, "-- Tab stops verified --"
    AuditStyleTabs "AuthorListItemTab", _
        Array(144, wdAlignTabLeft, wdTabLeaderSpaces), _
        Array(252, wdAlignTabLeft, wdTabLeaderSpaces)
    AuditStyleTabs "AuthorBookRef", _
        Array(36, wdAlignTabLeft, wdTabLeaderSpaces), _
        Array(378, wdAlignTabRight, wdTabLeaderDots)

    Print #m_TaxFile, ""
    Print #m_TaxFile, String(72, "=")
    Print #m_TaxFile, "Summary: " & m_TaxPass & " PASS   " & m_TaxFail & " FAIL"
    Print #m_TaxFile, "=== End Style Taxonomy Audit ==="

    Close #m_TaxFile
    m_TaxFile = 0

    Debug.Print "RUN_TAXONOMY_STYLES: " & m_TaxPass & " PASS  " & m_TaxFail & " FAIL  -> " & sPath
    MsgBox "Style Taxonomy Audit complete." & vbCrLf & _
           m_TaxPass & " PASS   " & m_TaxFail & " FAIL" & vbCrLf & _
           "Report: rpt\StyleTaxonomyAudit.txt", vbInformation, "RUN_TAXONOMY_STYLES"

PROC_EXIT:
    If m_TaxFile > 0 Then Close #m_TaxFile
    m_TaxFile = 0
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure RUN_TAXONOMY_STYLES of Module basTEST_aeBibleConfig"
    Resume PROC_EXIT
End Sub

'------------------------------------------------------------------------------
' AuditOneStyle  (Private)
' Checks one style against expected values; writes Result to open report file.
' Called only by RUN_TAXONOMY_STYLES.
'------------------------------------------------------------------------------
Private Sub AuditOneStyle(ByVal sName As String, _
                          ByVal sExpFont As String, _
                          ByVal dExpSize As Double, _
                          ByVal lExpAlign As Long, _
                          ByVal dExpFirstIndent As Double, _
                          ByVal lExpLineRule As Long, _
                          ByVal dExpLineSpacing As Double, _
                          ByVal dExpSpaceBefore As Double, _
                          ByVal dExpSpaceAfter As Double)
    On Error GoTo PROC_ERR
    Dim oStyle  As Word.Style
    Dim bPass   As Boolean
    Dim sDetail As String

    On Error Resume Next
    Set oStyle = ActiveDocument.Styles(sName)
    On Error GoTo PROC_ERR

    If oStyle Is Nothing Then
        Print #m_TaxFile, "FAIL  " & sName
        Print #m_TaxFile, "      NOT FOUND in document"
        m_TaxFail = m_TaxFail + 1
        GoTo PROC_EXIT
    End If

    bPass = True
    sDetail = ""

    If sExpFont <> "" Then
        If oStyle.Font.Name <> sExpFont Then
            bPass = False
            sDetail = sDetail & "      Font     : expected """ & sExpFont & _
                      """ got """ & oStyle.Font.Name & """" & vbCrLf
        End If
    End If

    If dExpSize <> 0 Then
        If oStyle.Font.Size <> dExpSize Then
            bPass = False
            sDetail = sDetail & "      Size     : expected " & dExpSize & _
                      " got " & oStyle.Font.Size & vbCrLf
        End If
    End If

    If lExpAlign <> -1 Then
        If oStyle.ParagraphFormat.Alignment <> lExpAlign Then
            bPass = False
            sDetail = sDetail & "      Alignment: expected " & lExpAlign & _
                      " got " & oStyle.ParagraphFormat.Alignment & vbCrLf
        End If
    End If

    If dExpFirstIndent <> -999 Then
        If Abs(oStyle.ParagraphFormat.FirstLineIndent - dExpFirstIndent) > 0.1 Then
            bPass = False
            sDetail = sDetail & "      Indent   : expected " & dExpFirstIndent & _
                      " got " & oStyle.ParagraphFormat.FirstLineIndent & vbCrLf
        End If
    End If

    If lExpLineRule <> -1 Then
        If oStyle.ParagraphFormat.LineSpacingRule <> lExpLineRule Then
            bPass = False
            sDetail = sDetail & "      LineRule : expected " & lExpLineRule & _
                      " got " & oStyle.ParagraphFormat.LineSpacingRule & vbCrLf
        End If
    End If

    If dExpLineSpacing <> -999 Then
        If Abs(oStyle.ParagraphFormat.LineSpacing - dExpLineSpacing) > 0.1 Then
            bPass = False
            sDetail = sDetail & "      LineSpacing: expected " & dExpLineSpacing & _
                      "pt got " & oStyle.ParagraphFormat.LineSpacing & "pt" & vbCrLf
        End If
    End If

    If dExpSpaceBefore <> -999 Then
        If Abs(oStyle.ParagraphFormat.SpaceBefore - dExpSpaceBefore) > 0.1 Then
            bPass = False
            sDetail = sDetail & "      SpaceBef : expected " & dExpSpaceBefore & _
                      " got " & oStyle.ParagraphFormat.SpaceBefore & vbCrLf
        End If
    End If

    If dExpSpaceAfter <> -999 Then
        If Abs(oStyle.ParagraphFormat.SpaceAfter - dExpSpaceAfter) > 0.1 Then
            bPass = False
            sDetail = sDetail & "      SpaceAft : expected " & dExpSpaceAfter & _
                      " got " & oStyle.ParagraphFormat.SpaceAfter & vbCrLf
        End If
    End If

    If bPass Then
        Print #m_TaxFile, "PASS  " & sName
        m_TaxPass = m_TaxPass + 1
    Else
        Print #m_TaxFile, "FAIL  " & sName
        If Len(sDetail) >= 2 Then _
            Print #m_TaxFile, Left(sDetail, Len(sDetail) - 2)
        m_TaxFail = m_TaxFail + 1
    End If

PROC_EXIT:
    Set oStyle = Nothing
    Exit Sub
PROC_ERR:
    Print #m_TaxFile, "ERROR " & sName & " -- Erl=" & Erl & _
          "  Err=" & Err.Number & "  " & Err.Description
    m_TaxFail = m_TaxFail + 1
    Resume PROC_EXIT
End Sub

'------------------------------------------------------------------------------
' AuditStyleTabs  (Public)
' Validates a paragraph style's explicit tab-stop list against expected specs.
' Each ParamArray element is a 3-element Array(Position, Alignment, Leader).
' Pass with no expected entries to assert "no explicit tab stops".
'
' Output channel matches AuditOneStyle: writes via Print #m_TaxFile and
' increments m_TaxPass / m_TaxFail. Result lines tagged "(TabStops)" so
' they don't collide with AuditOneStyle entries on the same style name.
'
' Position tolerance: 0.1 pt (Word stores tab positions as Double).
'
' Called only by RUN_TAXONOMY_STYLES.
'------------------------------------------------------------------------------
Public Sub AuditStyleTabs(ByVal sName As String, ParamArray expected() As Variant)
    On Error GoTo PROC_ERR
    Dim oStyle As Word.Style
    Dim bPass As Boolean
    Dim sDetail As String

    On Error Resume Next
    Set oStyle = ActiveDocument.Styles(sName)
    On Error GoTo PROC_ERR

    If oStyle Is Nothing Then
        Print #m_TaxFile, "FAIL  " & sName & " (TabStops)"
        Print #m_TaxFile, "      NOT FOUND in document"
        m_TaxFail = m_TaxFail + 1
        GoTo PROC_EXIT
    End If

    ' Compute expected count - empty ParamArray has LBound > UBound in VBA.
    Dim expCount As Long
    If LBound(expected) > UBound(expected) Then
        expCount = 0
    Else
        expCount = UBound(expected) - LBound(expected) + 1
    End If

    Dim actCount As Long
    actCount = oStyle.ParagraphFormat.TabStops.Count

    bPass = True
    sDetail = ""

    If actCount <> expCount Then
        bPass = False
        sDetail = sDetail & "      Count    : expected " & expCount & _
                  " got " & actCount & vbCrLf
    Else
        Dim i As Long
        Dim ts As Word.TabStop
        Dim spec As Variant
        Dim expPos As Double, expAlign As Long, expLeader As Long
        For i = 0 To expCount - 1
            spec = expected(LBound(expected) + i)
            expPos = CDbl(spec(0))
            expAlign = CLng(spec(1))
            expLeader = CLng(spec(2))

            Set ts = oStyle.ParagraphFormat.TabStops(i + 1)

            If Abs(ts.Position - expPos) > 0.1 Then
                bPass = False
                sDetail = sDetail & "      Tab " & (i + 1) & _
                          " Position: expected " & expPos & _
                          " got " & ts.Position & vbCrLf
            End If
            If ts.Alignment <> expAlign Then
                bPass = False
                sDetail = sDetail & "      Tab " & (i + 1) & _
                          " Align   : expected " & TabAlignName(expAlign) & _
                          " got " & TabAlignName(ts.Alignment) & vbCrLf
            End If
            If ts.Leader <> expLeader Then
                bPass = False
                sDetail = sDetail & "      Tab " & (i + 1) & _
                          " Leader  : expected " & TabLeaderName(expLeader) & _
                          " got " & TabLeaderName(ts.Leader) & vbCrLf
            End If
        Next i
    End If

    If bPass Then
        Print #m_TaxFile, "PASS  " & sName & " (TabStops)"
        m_TaxPass = m_TaxPass + 1
    Else
        Print #m_TaxFile, "FAIL  " & sName & " (TabStops)"
        If Len(sDetail) >= 2 Then _
            Print #m_TaxFile, Left(sDetail, Len(sDetail) - 2)
        m_TaxFail = m_TaxFail + 1
    End If

PROC_EXIT:
    Set oStyle = Nothing
    Exit Sub
PROC_ERR:
    Print #m_TaxFile, "ERROR " & sName & " (TabStops) -- Erl=" & Erl & _
          "  Err=" & Err.Number & "  " & Err.Description
    m_TaxFail = m_TaxFail + 1
    Resume PROC_EXIT
End Sub

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

