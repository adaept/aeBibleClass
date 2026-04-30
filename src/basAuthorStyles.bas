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
' Output: Immediate window only. No file written.
'   (A) Flagged at-risk styles (list-family inheritance).
'   (B) Full inventory of every paragraph style with non-empty BaseStyle.
'
' Usage:
'   AuditListStyleRisk
' ==========================================================================
Public Sub AuditListStyleRisk()
    On Error GoTo PROC_ERR
    Dim s As Word.Style
    Dim base As String
    Dim baseLC As String
    Dim flagged As Long
    Dim totalWithBase As Long

    Debug.Print "---- AuditListStyleRisk: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ----"
    Debug.Print ""
    Debug.Print "(A) Paragraph styles whose BaseStyle is a list-family built-in"
    Debug.Print "    (List Paragraph, List Number, List Bullet, List Continue, List):"
    Debug.Print ""

    For Each s In ActiveDocument.Styles
        If s.Type = wdStyleTypeParagraph Then
            On Error Resume Next
            base = s.baseStyle
            On Error GoTo PROC_ERR
            baseLC = LCase$(Trim$(base))

            If baseLC = "list paragraph" Or baseLC = "list" _
               Or baseLC Like "list number*" Or baseLC Like "list bullet*" _
               Or baseLC Like "list continue*" Then
                Debug.Print "  FLAG  " & s.NameLocal & " | BaseStyle=""" & base & _
                            """ | Priority=" & s.Priority
                flagged = flagged + 1
            End If
        End If
    Next s

    Debug.Print ""
    Debug.Print "(B) Full inventory: every paragraph style with non-empty BaseStyle:"
    Debug.Print ""

    For Each s In ActiveDocument.Styles
        If s.Type = wdStyleTypeParagraph Then
            On Error Resume Next
            base = s.baseStyle
            On Error GoTo PROC_ERR
            If Len(Trim$(base)) > 0 Then
                Debug.Print "  " & s.NameLocal & " <- """ & base & """ | Priority=" & s.Priority
                totalWithBase = totalWithBase + 1
            End If
        End If
    Next s

    Debug.Print ""
    Debug.Print "Flagged (list-family inheritance): " & flagged
    Debug.Print "Total paragraph styles with BaseStyle: " & totalWithBase
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure AuditListStyleRisk of Module basAuthorStyles"
    Resume PROC_EXIT
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
