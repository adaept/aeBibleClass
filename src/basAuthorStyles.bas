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
' Inventories paragraph styles that inherit from "List Paragraph" or hold
' a LinkToListTemplate - both are at risk of triggering the Word
' numbering-engine hang on Modify Style edits.
'
' Output: Immediate window only. No file written.
'
' Usage:
'   AuditListStyleRisk
' ==========================================================================
Public Sub AuditListStyleRisk()
    On Error GoTo PROC_ERR
    Dim s As Word.Style
    Dim base As String
    Dim hasLT As Boolean
    Dim flagged As Long

    Debug.Print "---- AuditListStyleRisk: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ----"
    Debug.Print "Approved paragraph styles flagged as at-risk:"
    Debug.Print "(BaseStyle inherits List Paragraph, OR LinkToListTemplate set)"
    Debug.Print ""

    For Each s In ActiveDocument.Styles
        If s.Type = wdStyleTypeParagraph Then
            On Error Resume Next
            base = s.baseStyle
            hasLT = Not (s.LinkToListTemplate Is Nothing)
            On Error GoTo PROC_ERR

            If LCase$(Trim$(base)) = "list paragraph" Or hasLT Then
                Debug.Print s.NameLocal & " | BaseStyle=""" & base & _
                            """ | HasListTemplate=" & hasLT & _
                            " | Priority=" & s.Priority
                flagged = flagged + 1
            End If
        End If
    Next s

    Debug.Print ""
    Debug.Print "Flagged: " & flagged & " style(s)."
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure AuditListStyleRisk of Module basAuthorStyles"
    Resume PROC_EXIT
End Sub
