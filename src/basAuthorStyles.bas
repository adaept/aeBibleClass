Attribute VB_Name = "basAuthorStyles"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

' ==========================================================================
' basAuthorStyles
' ==========================================================================
' Migration kit for the Word "List Paragraph" numbering-engine bug.
' Replaces ListItem / ListItemBody / ListItemTab (which inherit from
' "List Paragraph" and trigger the Modify-Style hang in large documents)
' with standalone Author* equivalents:
'   ListItem      -> AuthorListItem
'   ListItemBody  -> AuthorListItemBody
'   ListItemTab   -> AuthorListItemTab
'
' See EDSG/10-list-paragraph-bug.md for symptom, project policy, and the
' five-step migration recipe. Phase 0 (this Sub) is the diagnostic.
' Phase 1+ Subs (CreateAuthorStyles, TransportAuthorStyles,
' MigrateParagraphs) are added as each phase is approved.
'
' This module is self-contained and intended to be removed once the
' migration is complete and RUN_TAXONOMY_STYLES is clean.
' ==========================================================================

' ==========================================================================
' AuditListStyleRisk
' ==========================================================================
' Step 0 of the List Paragraph migration (EDSG/10-list-paragraph-bug.md).
' Inventories paragraph styles that inherit from any of Word's built-in
' list-family parents (List Paragraph, List, List Number, List Bullet,
' List Continue) - all are risk vectors for the numbering-engine hang
' on Modify Style edits in large documents.
'
' Note: Style.LinkToListTemplate is a method (not a read-only property)
' so it cannot be queried here. The list-template-attachment-without-
' inheritance case is rare in practice; if Phase 0 output looks
' incomplete, a paragraph-level fallback can be layered on (sample
' Paragraphs.Range.ListFormat.ListTemplate per style).
'
' Output:
'   (A) Flagged at-risk styles (list-family inheritance).
'   (B) Full inventory of every paragraph style with non-empty BaseStyle.
' Both sections sorted by Priority ASC, then NameLocal ASC, so approved
' styles (priorities 1..N) appear first and Word built-ins (Priority=99)
' fall to the bottom alphabetically.
'
' Default writes to rpt\ListStyleRiskAudit.txt; pass False for Immediate
' window only.
'
' Usage:
'   AuditListStyleRisk              ' default writes file
'   AuditListStyleRisk False        ' Immediate only, no file
' ==========================================================================
Public Sub AuditListStyleRisk(Optional ByVal bWriteFile As Boolean = True)
    On Error GoTo PROC_ERR

    ' Collect into 2D Variant arrays: (NameLocal, BaseStyle, Priority).
    ' 200 row cap matches AuditVerseMarkerStructure's H1 cap; ~106 rows
    ' observed in the Bible-class doc, so 200 has comfortable headroom.
    Dim flaggedArr() As Variant
    Dim allBaseArr() As Variant
    Dim flaggedCount As Long
    Dim allBaseCount As Long
    ReDim flaggedArr(1 To 200, 1 To 3)
    ReDim allBaseArr(1 To 200, 1 To 3)

    Dim s As Word.Style
    Dim base As String
    Dim baseLC As String

    For Each s In ActiveDocument.Styles
        If s.Type = wdStyleTypeParagraph Then
            On Error Resume Next
            base = s.baseStyle
            On Error GoTo PROC_ERR

            If Len(Trim$(base)) > 0 Then
                allBaseCount = allBaseCount + 1
                allBaseArr(allBaseCount, 1) = s.NameLocal
                allBaseArr(allBaseCount, 2) = base
                allBaseArr(allBaseCount, 3) = s.Priority

                baseLC = LCase$(Trim$(base))
                If baseLC = "list paragraph" Or baseLC = "list" _
                   Or baseLC Like "list number*" Or baseLC Like "list bullet*" _
                   Or baseLC Like "list continue*" Then
                    flaggedCount = flaggedCount + 1
                    flaggedArr(flaggedCount, 1) = s.NameLocal
                    flaggedArr(flaggedCount, 2) = base
                    flaggedArr(flaggedCount, 3) = s.Priority
                End If
            End If
        End If
    Next s

    SortByPriorityName flaggedArr, flaggedCount
    SortByPriorityName allBaseArr, allBaseCount

    Dim sOut As String
    Const NL As String = vbCrLf

    sOut = "---- AuditListStyleRisk: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ----" & NL & NL
    sOut = sOut & "(A) Paragraph styles whose BaseStyle is a list-family built-in" & NL
    sOut = sOut & "    (List Paragraph, List Number, List Bullet, List Continue, List):" & NL & NL
    sOut = sOut & FormatBaseStyleRows(flaggedArr, flaggedCount, "  FLAG  ")
    sOut = sOut & NL
    sOut = sOut & "(B) Full inventory: every paragraph style with non-empty BaseStyle" & NL
    sOut = sOut & "    (sorted by priority ascending, then name):" & NL & NL
    sOut = sOut & FormatBaseStyleRows(allBaseArr, allBaseCount, "  ")
    sOut = sOut & NL
    sOut = sOut & "Flagged (list-family inheritance): " & flaggedCount & NL
    sOut = sOut & "Total paragraph styles with BaseStyle: " & allBaseCount & NL

    Debug.Print sOut
    If bWriteFile Then WriteListStyleRiskFile sOut

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure AuditListStyleRisk of Module basAuthorStyles"
    Resume PROC_EXIT
End Sub

' --------------------------------------------------------------------------
' SortByPriorityName - bubble sort on 2D array (Name, BaseStyle, Priority)
' Primary key: Priority ASC. Secondary: NameLocal ASC (case-insensitive).
' --------------------------------------------------------------------------
Private Sub SortByPriorityName(ByRef arr As Variant, ByVal count As Long)
    Dim i As Long, j As Long
    Dim tmpName As String, tmpBase As String, tmpPri As Long
    For i = 1 To count - 1
        For j = i + 1 To count
            If (CLng(arr(j, 3)) < CLng(arr(i, 3))) Or _
               (CLng(arr(j, 3)) = CLng(arr(i, 3)) And _
                StrComp(CStr(arr(j, 1)), CStr(arr(i, 1)), vbTextCompare) < 0) Then
                tmpName = arr(i, 1):  arr(i, 1) = arr(j, 1):  arr(j, 1) = tmpName
                tmpBase = arr(i, 2):  arr(i, 2) = arr(j, 2):  arr(j, 2) = tmpBase
                tmpPri = arr(i, 3):   arr(i, 3) = arr(j, 3):  arr(j, 3) = tmpPri
            End If
        Next j
    Next i
End Sub

' --------------------------------------------------------------------------
' FormatBaseStyleRows - render N rows in the form
'   <prefix>Name <- "BaseStyle" | Priority=N<NL>
' --------------------------------------------------------------------------
Private Function FormatBaseStyleRows(ByRef arr As Variant, ByVal count As Long, _
                                     ByVal prefix As String) As String
    Const NL As String = vbCrLf
    Dim i As Long, sBuf As String
    For i = 1 To count
        sBuf = sBuf & prefix & arr(i, 1) & " <- """ & arr(i, 2) & _
               """ | Priority=" & arr(i, 3) & NL
    Next i
    FormatBaseStyleRows = sBuf
End Function

' --------------------------------------------------------------------------
' WriteListStyleRiskFile - write the report to rpt\ListStyleRiskAudit.txt
' --------------------------------------------------------------------------
Private Sub WriteListStyleRiskFile(ByVal sContent As String)
    Dim oFSO As Object
    Dim oStream As Object
    Dim sPath As String
    sPath = ActiveDocument.Path & "\rpt\ListStyleRiskAudit.txt"
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.CreateTextFile(sPath, True, False)   ' ASCII
    oStream.Write sContent
    oStream.Close
End Sub

' ==========================================================================
' CreateAuthorStyles
' ==========================================================================
' Step 1 of the List Paragraph migration. Defines the replacement styles
' from scratch in the *active* document - intended to be run in a fresh
' blank .docm holding file (e.g. tools/style_holding.docm), NOT in the
' live document.
'
' Critical isolation: BaseStyle = "" set BEFORE any other property; no
' LinkToListTemplate call anywhere. Specs are descriptive (capture
' current visual rendering of the styles being replaced) - read from
' rpt/Styles/style_ListItem.txt and rpt/Styles/style_AuthorBookRef.txt.
' Two QA-checklist fixes applied during creation: AuthorListItem is
' created with QuickStyle = False and AutomaticallyUpdate = False
' (current ListItem has both as True). User decision 2026-04-29.
'
' NextParagraphStyle is NOT set in Phase 1 - the holding .docm contains
' only built-in styles plus the two created here, so referencing custom
' names like ListItemBody / AuthorBookRefNew would raise error 5834
' (item with specified name does not exist) on Word's name-resolution
' check. Default "Normal" is used for transport. Phase 4 sets the
' final NextParagraphStyle values once all target names exist in the
' live document.
'
' Usage (in a fresh blank .docm):
'   CreateAuthorStyles
'   ' then save the holding file as tools/style_holding.docm
' ==========================================================================
Public Sub CreateAuthorStyles()
    On Error GoTo PROC_ERR

    DefineAuthorListItem
    DefineAuthorBookRefNew

    Debug.Print "CreateAuthorStyles: Done. AuthorListItem and AuthorBookRefNew defined."
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure CreateAuthorStyles of Module basAuthorStyles"
    Resume PROC_EXIT
End Sub

Private Sub DefineAuthorListItem()
    Dim s As Word.Style
    On Error Resume Next
    Set s = ActiveDocument.Styles("AuthorListItem")
    On Error GoTo 0
    If s Is Nothing Then
        Set s = ActiveDocument.Styles.Add("AuthorListItem", wdStyleTypeParagraph)
    End If

    s.baseStyle = ""                           ' MUST come first
    s.AutomaticallyUpdate = False               ' QA-checklist fix (was True on ListItem)
    s.QuickStyle = False                        ' QA-checklist fix (was True on ListItem)
    ' NextParagraphStyle deferred to Phase 4 - the holding .docm doesn't
    ' contain ListItemBody / AuthorListItemBody, so setting it here would
    ' raise 5834. Default ("Normal") is fine for transport.
    s.Font.Name = "Carlito"
    s.Font.Size = 11
    s.Font.Bold = True
    s.Font.Italic = True
    s.Font.Underline = wdUnderlineNone
    s.Font.color = wdColorAutomatic
    With s.ParagraphFormat
        .Alignment = wdAlignParagraphLeft
        .LeftIndent = 18
        .RightIndent = 0
        .FirstLineIndent = -18
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacing = 12
        .LineSpacingRule = wdLineSpaceSingle
        .WidowControl = True
        .KeepTogether = False
        .KeepWithNext = True
        .PageBreakBefore = False
        .OutlineLevel = wdOutlineLevelBodyText
    End With
End Sub

Private Sub DefineAuthorBookRefNew()
    Dim s As Word.Style
    On Error Resume Next
    Set s = ActiveDocument.Styles("AuthorBookRefNew")
    On Error GoTo 0
    If s Is Nothing Then
        Set s = ActiveDocument.Styles.Add("AuthorBookRefNew", wdStyleTypeParagraph)
    End If

    s.baseStyle = ""
    s.AutomaticallyUpdate = False
    s.QuickStyle = False
    ' NextParagraphStyle deferred to Phase 4 (set to AuthorBookRef after
    ' the rename). Default ("Normal") is fine for transport.
    s.Font.Name = "Carlito"
    s.Font.Size = 11
    s.Font.Bold = True
    s.Font.Italic = False
    s.Font.Underline = wdUnderlineNone
    s.Font.color = wdColorAutomatic
    With s.ParagraphFormat
        .Alignment = wdAlignParagraphLeft
        .LeftIndent = 36
        .RightIndent = 0
        .FirstLineIndent = -18
        .SpaceBefore = 0
        .SpaceAfter = 11
        .LineSpacing = 12
        .LineSpacingRule = wdLineSpaceSingle
        .WidowControl = True
        .KeepTogether = False
        .KeepWithNext = False
        .PageBreakBefore = False
        .OutlineLevel = wdOutlineLevelBodyText
    End With
End Sub

' ==========================================================================
' TransportAuthorStyles
' ==========================================================================
' Step 2 of the List Paragraph migration. Reads the AuthorListItem and
' AuthorBookRefNew style definitions from a holding .docm and creates
' equivalent standalone styles in ActiveDocument (the live .docm).
'
' Pre-flight:
'   * Source holding document is open by name (default: style_holding.docm).
'   * Source contains both AuthorListItem and AuthorBookRefNew.
'   * ActiveDocument is not the holding doc itself.
'
' Idempotency: if a destination style already exists, the Sub warns and
' skips it. To force a re-import, manually delete the destination style
' and re-run.
'
' NextParagraphStyle is intentionally not copied here - Phase 4 sets it
' after the live-doc renames complete.
'
' Usage (in the live .docm, with holding doc also open):
'   TransportAuthorStyles                       ' default source name
'   TransportAuthorStyles "my_holding.docm"     ' custom source name
' ==========================================================================
Public Sub TransportAuthorStyles(Optional ByVal sourceName As String = "style_holding.docm")
    On Error GoTo PROC_ERR

    ' Pre-flight 1: source doc open
    Dim srcDoc As Document
    On Error Resume Next
    Set srcDoc = Documents(sourceName)
    On Error GoTo PROC_ERR
    If srcDoc Is Nothing Then
        MsgBox "Source document """ & sourceName & """ is not open. " & _
               "Open it first, then re-run TransportAuthorStyles.", _
               vbExclamation, "TransportAuthorStyles"
        Exit Sub
    End If

    ' Pre-flight 2: ActiveDocument is not the source
    If ActiveDocument.Name = srcDoc.Name Then
        MsgBox "ActiveDocument is the holding file. Activate the live " & _
               ".docm before running TransportAuthorStyles.", _
               vbExclamation, "TransportAuthorStyles"
        Exit Sub
    End If

    ' Pre-flight 3: source contains both expected styles
    If Not StyleExists(srcDoc, "AuthorListItem") Then
        MsgBox "Source missing AuthorListItem. Run CreateAuthorStyles in " & _
               "the holding doc first.", vbExclamation, "TransportAuthorStyles"
        Exit Sub
    End If
    If Not StyleExists(srcDoc, "AuthorBookRefNew") Then
        MsgBox "Source missing AuthorBookRefNew. Run CreateAuthorStyles in " & _
               "the holding doc first.", vbExclamation, "TransportAuthorStyles"
        Exit Sub
    End If

    Debug.Print "TransportAuthorStyles: source=""" & srcDoc.Name & _
                """, target=""" & ActiveDocument.Name & """"

    CopyOneStyle srcDoc, ActiveDocument, "AuthorListItem"
    CopyOneStyle srcDoc, ActiveDocument, "AuthorBookRefNew"

    Debug.Print "TransportAuthorStyles: Done."
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure TransportAuthorStyles of Module basAuthorStyles"
    Resume PROC_EXIT
End Sub

Private Function StyleExists(ByVal doc As Document, ByVal styleName As String) As Boolean
    Dim s As Word.Style
    On Error Resume Next
    Set s = doc.Styles(styleName)
    On Error GoTo 0
    StyleExists = Not (s Is Nothing)
End Function

Private Sub CopyOneStyle(ByVal srcDoc As Document, ByVal dstDoc As Document, _
                         ByVal styleName As String)
    ' Idempotent skip
    If StyleExists(dstDoc, styleName) Then
        Debug.Print "  SKIP   " & styleName & " already exists in target."
        Exit Sub
    End If

    Dim src As Word.Style
    Dim dst As Word.Style
    Set src = srcDoc.Styles(styleName)
    Set dst = dstDoc.Styles.Add(styleName, wdStyleTypeParagraph)

    ' BaseStyle MUST be set first - same isolation rule as Phase 1.
    dst.baseStyle = ""

    ' QA-checklist properties.
    dst.AutomaticallyUpdate = src.AutomaticallyUpdate
    dst.QuickStyle = src.QuickStyle

    ' Font properties.
    dst.Font.Name = src.Font.Name
    dst.Font.Size = src.Font.Size
    dst.Font.Bold = src.Font.Bold
    dst.Font.Italic = src.Font.Italic
    dst.Font.Underline = src.Font.Underline
    dst.Font.color = src.Font.color
    dst.Font.SmallCaps = src.Font.SmallCaps
    dst.Font.AllCaps = src.Font.AllCaps

    ' ParagraphFormat properties.
    With dst.ParagraphFormat
        .Alignment = src.ParagraphFormat.Alignment
        .LeftIndent = src.ParagraphFormat.LeftIndent
        .RightIndent = src.ParagraphFormat.RightIndent
        .FirstLineIndent = src.ParagraphFormat.FirstLineIndent
        .SpaceBefore = src.ParagraphFormat.SpaceBefore
        .SpaceAfter = src.ParagraphFormat.SpaceAfter
        .LineSpacing = src.ParagraphFormat.LineSpacing
        .LineSpacingRule = src.ParagraphFormat.LineSpacingRule
        .WidowControl = src.ParagraphFormat.WidowControl
        .KeepTogether = src.ParagraphFormat.KeepTogether
        .KeepWithNext = src.ParagraphFormat.KeepWithNext
        .PageBreakBefore = src.ParagraphFormat.PageBreakBefore
        .OutlineLevel = src.ParagraphFormat.OutlineLevel
    End With

    ' ParagraphFormat tab stops - reproduce explicit stops from source.
    ' ClearAll first in case the destination inherited tabs from BaseStyle
    ' before BaseStyle was set to "". Cheap; safe; idempotent.
    dst.ParagraphFormat.TabStops.ClearAll
    Dim ts As Word.TabStop
    For Each ts In src.ParagraphFormat.TabStops
        dst.ParagraphFormat.TabStops.Add _
            Position:=ts.Position, _
            Alignment:=ts.Alignment, _
            Leader:=ts.Leader
    Next ts

    ' NextParagraphStyle deliberately not copied - Phase 4 sets it.

    Debug.Print "  COPY   " & styleName & " transported."
End Sub

' ==========================================================================
' MigrateParagraphs
' ==========================================================================
' Step 3 of the List Paragraph migration. Reassigns every paragraph
' currently using oldName to use newName. Walks the main story only -
' list-shaped paragraphs do not appear in footnotes / headers / footers
' in this project (user-confirmed 2026-04-30).
'
' Pre-flight:
'   * Both oldName and newName must exist in ActiveDocument.
'   * Caller is responsible for choosing the correct doc context
'     (run on the test copy first, then on production).
'
' Output: count of paragraphs migrated.
'
' Idempotent: if run twice, the second run reports 0 (no paragraphs
' still using oldName).
'
' Usage:
'   MigrateParagraphs "ListItem", "AuthorListItem"
'   MigrateParagraphs "AuthorBookRef", "AuthorBookRefNew"
' ==========================================================================
Public Sub MigrateParagraphs(ByVal oldName As String, ByVal newName As String)
    On Error GoTo PROC_ERR

    ' Pre-flight: both styles must exist.
    If Not StyleExists(ActiveDocument, oldName) Then
        MsgBox "Source style """ & oldName & """ not found in " & _
               ActiveDocument.Name & ".", vbExclamation, "MigrateParagraphs"
        Exit Sub
    End If
    If Not StyleExists(ActiveDocument, newName) Then
        MsgBox "Target style """ & newName & """ not found in " & _
               ActiveDocument.Name & ". Run TransportAuthorStyles first.", _
               vbExclamation, "MigrateParagraphs"
        Exit Sub
    End If

    Dim screenWas As Boolean
    screenWas = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim total As Long
    Dim para As Word.Paragraph

    Debug.Print "MigrateParagraphs: " & oldName & " -> " & newName & _
                " in " & ActiveDocument.Name

    For Each para In ActiveDocument.Paragraphs
        If para.style.NameLocal = oldName Then
            para.style = ActiveDocument.Styles(newName)
            total = total + 1
        End If
    Next para

    Application.ScreenUpdating = screenWas

    Debug.Print "MigrateParagraphs: " & total & " paragraph(s) migrated."
PROC_EXIT:
    Exit Sub
PROC_ERR:
    Application.ScreenUpdating = True
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure MigrateParagraphs of Module basAuthorStyles"
    Resume PROC_EXIT
End Sub

' ==========================================================================
' DecommissionAuthorStyles
' ==========================================================================
' Phase 4a of the List Paragraph migration. Performs six ordered live-doc
' operations to retire the old list-engine-entangled styles and finalise
' the new standalone styles' inter-style references.
'
' Pre-flight (all must pass before any state change):
'   * No paragraphs still using "ListItem" (would lose styling on delete).
'   * No paragraphs still using "AuthorBookRef" (same risk).
'   * AuthorListItem, AuthorBookRefNew, ListItemBody all present.
'   * AuthorListItemBody not already present (rename collision).
'
' Idempotency: detects "already decommissioned" state and exits cleanly
' rather than failing partway through a re-run.
'
' Operations:
'   1. Delete old ListItem
'   2. Delete old AuthorBookRef
'   3. Rename ListItemBody -> AuthorListItemBody
'   4. Rename AuthorBookRefNew -> AuthorBookRef
'   5. AuthorListItem.NextParagraphStyle = "AuthorListItemBody"
'   6. AuthorBookRef.NextParagraphStyle = "AuthorBookRef" (self)
'
' Run on test copy first, verify, then on production. After both,
' apply Phase 4b source-code edits (approved array, RUN_TAXONOMY_STYLES).
'
' Usage:
'   DecommissionAuthorStyles
' ==========================================================================
Public Sub DecommissionAuthorStyles()
    On Error GoTo PROC_ERR

    ' Idempotency check: already decommissioned?
    If Not StyleExists(ActiveDocument, "ListItem") _
       And StyleExists(ActiveDocument, "AuthorListItem") _
       And Not StyleExists(ActiveDocument, "ListItemBody") _
       And StyleExists(ActiveDocument, "AuthorListItemBody") _
       And Not StyleExists(ActiveDocument, "AuthorBookRefNew") _
       And StyleExists(ActiveDocument, "AuthorBookRef") Then
        MsgBox "Decommission already complete. No action taken.", _
               vbInformation, "DecommissionAuthorStyles"
        Exit Sub
    End If

    ' Pre-flight 1: no orphan paragraphs on deletion.
    Dim n As Long
    n = CountParagraphsByStyle("ListItem")
    If n > 0 Then
        MsgBox n & " paragraph(s) still use ""ListItem"". Run " & _
               "MigrateParagraphs ""ListItem"", ""AuthorListItem"" first.", _
               vbExclamation, "DecommissionAuthorStyles"
        Exit Sub
    End If
    n = CountParagraphsByStyle("AuthorBookRef")
    If n > 0 Then
        MsgBox n & " paragraph(s) still use ""AuthorBookRef"". Run " & _
               "MigrateParagraphs ""AuthorBookRef"", ""AuthorBookRefNew"" first.", _
               vbExclamation, "DecommissionAuthorStyles"
        Exit Sub
    End If

    ' Pre-flight 2: target styles present.
    If Not StyleExists(ActiveDocument, "AuthorListItem") Then
        MsgBox "AuthorListItem missing. Run TransportAuthorStyles first.", _
               vbExclamation, "DecommissionAuthorStyles"
        Exit Sub
    End If
    If Not StyleExists(ActiveDocument, "AuthorBookRefNew") Then
        MsgBox "AuthorBookRefNew missing. Run TransportAuthorStyles first.", _
               vbExclamation, "DecommissionAuthorStyles"
        Exit Sub
    End If
    If Not StyleExists(ActiveDocument, "ListItemBody") Then
        MsgBox "ListItemBody missing. Document state is unexpected; abort.", _
               vbExclamation, "DecommissionAuthorStyles"
        Exit Sub
    End If

    ' Pre-flight 3: no rename-target collisions.
    If StyleExists(ActiveDocument, "AuthorListItemBody") Then
        MsgBox "AuthorListItemBody already exists; rename collision. Abort.", _
               vbExclamation, "DecommissionAuthorStyles"
        Exit Sub
    End If

    Debug.Print "DecommissionAuthorStyles: starting on " & ActiveDocument.Name

    ' Step 1: delete old ListItem
    ActiveDocument.Styles("ListItem").Delete
    Debug.Print "  1. Deleted ListItem"

    ' Step 2: delete old AuthorBookRef
    ActiveDocument.Styles("AuthorBookRef").Delete
    Debug.Print "  2. Deleted AuthorBookRef"

    ' Step 3: rename ListItemBody -> AuthorListItemBody
    ActiveDocument.Styles("ListItemBody").NameLocal = "AuthorListItemBody"
    Debug.Print "  3. Renamed ListItemBody -> AuthorListItemBody"

    ' Step 4: rename AuthorBookRefNew -> AuthorBookRef
    ActiveDocument.Styles("AuthorBookRefNew").NameLocal = "AuthorBookRef"
    Debug.Print "  4. Renamed AuthorBookRefNew -> AuthorBookRef"

    ' Step 5: AuthorListItem.NextParagraphStyle = "AuthorListItemBody"
    ActiveDocument.Styles("AuthorListItem").NextParagraphStyle = _
        ActiveDocument.Styles("AuthorListItemBody")
    Debug.Print "  5. AuthorListItem.NextParagraphStyle = AuthorListItemBody"

    ' Step 6: AuthorBookRef.NextParagraphStyle = "AuthorBookRef" (self)
    ActiveDocument.Styles("AuthorBookRef").NextParagraphStyle = _
        ActiveDocument.Styles("AuthorBookRef")
    Debug.Print "  6. AuthorBookRef.NextParagraphStyle = AuthorBookRef (self)"

    Debug.Print "DecommissionAuthorStyles: Done."
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure DecommissionAuthorStyles of Module basAuthorStyles"
    Resume PROC_EXIT
End Sub

Private Function CountParagraphsByStyle(ByVal styleName As String) As Long
    Dim para As Word.Paragraph
    Dim n As Long
    For Each para In ActiveDocument.Paragraphs
        If para.style.NameLocal = styleName Then
            n = n + 1
        End If
    Next para
    CountParagraphsByStyle = n
End Function
