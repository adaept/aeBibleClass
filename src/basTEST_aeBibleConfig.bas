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
    Dim count As Long
    Dim i As Long, j As Long
    Dim tmpName As String, tmpPri As Long

    'First pass: count eligible styles
    For Each s In ActiveDocument.Styles
        If s.Type = wdStyleTypeParagraph Or s.Type = wdStyleTypeCharacter Then
            count = count + 1
        End If
    Next s

    'Allocate array: 1-based, 2 columns (Name, Priority)
    ReDim arr(1 To count, 1 To 2)

    'Second pass: fill array
    count = 1
    For Each s In ActiveDocument.Styles
        If s.Type = wdStyleTypeParagraph Or s.Type = wdStyleTypeCharacter Then
            arr(count, 1) = s.NameLocal
            arr(count, 2) = s.Priority
            count = count + 1
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

