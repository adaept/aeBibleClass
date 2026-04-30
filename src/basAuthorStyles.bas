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
