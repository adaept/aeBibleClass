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
    'DumpPriorities
End Sub

Public Sub PromoteApprovedStyles()

    Dim s As style
    Dim approved As Variant
    Dim i As Long
    Dim missing As Collection
    Set missing = New Collection

    'List your approved styles in the order you want them to appear
    approved = Array("Normal", "Body Text", "Heading 1", "Heading 2", _
                     "CustomParaAfterH1", "CustomParaAfterH1-2nd", "DatAuthRef", _
                     "Chapter Verse marker", "Verse marker", _
                     "EmphasisBlack", "EmphasisRed", "Lamentation", "Psalms BOOK", _
                     "Words of Jesus", "TheHeaders", "TheFooters", _
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
        msg = "WARNING: The following styles were NOT found:" & vbCrLf & vbCrLf

        For i = 1 To missing.Count
            msg = msg & "  � " & missing(i) & vbCrLf
        Next i

        MsgBox msg, vbExclamation, "PromoteApprovedStyles Diagnostics"
    End If

    Debug.Print "PromoteApprovedStyles: Done!"
End Sub

Public Sub DumpPriorities()
    Dim s As style
    For Each s In ActiveDocument.Styles
        If s.Type = wdStyleTypeParagraph Or s.Type = wdStyleTypeCharacter Then
            Debug.Print s.NameLocal & "  ->  " & s.Priority
        End If
    Next s
End Sub

