Attribute VB_Name = "basTEST_NoSuperscriptOrdinals"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

'=======================================================================
' Test_NoSuperscriptOrdinals
' Purpose : Fail if the active document contains any of the ordinal
'           suffixes "st", "nd", "rd", or "th" formatted as superscript.
' Expected: 0 hits total across all four suffixes.
' Hint    : On the first failing match the test prints the page number,
'           line number, and the surrounding paragraph text so the user
'           can navigate directly to the offending location.
'=======================================================================
Public Function Test_NoSuperscriptOrdinals() As Long

    On Error GoTo PROC_ERR

    Dim oAssert As aeAssertClass
    Set oAssert = New aeAssertClass

    Dim vSuffix As Variant
    vSuffix = Array("st", "nd", "rd", "th")

    Dim lngTotal As Long
    Dim i As Long
    Dim strFirstHint As String
    strFirstHint = vbNullString

    For i = LBound(vSuffix) To UBound(vSuffix)
        lngTotal = lngTotal + _
                   CountSuperscriptToken(CStr(vSuffix(i)), strFirstHint)
    Next i

    Debug.Print "Superscript ordinal suffix Count = " & lngTotal
    If lngTotal > 0 And Len(strFirstHint) > 0 Then
        Debug.Print "HINT (first failure): " & strFirstHint
    End If

    oAssert.AssertEqual 0, lngTotal, _
        "No superscript ordinal suffixes (st/nd/rd/th) in document"

    Test_NoSuperscriptOrdinals = lngTotal

PROC_EXIT:
    Set oAssert = Nothing
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & _
           " (" & Err.Description & _
           ") in procedure Test_NoSuperscriptOrdinals of Module basTEST_NoSuperscriptOrdinals"
    Resume PROC_EXIT
End Function

'-----------------------------------------------------------------------
' Count occurrences of strToken formatted as superscript in the main
' story of ActiveDocument. On the first hit, populate strFirstHint with
' a navigation aid (page, line, paragraph snippet).
'-----------------------------------------------------------------------
Private Function CountSuperscriptToken( _
        ByVal strToken As String, _
        ByRef strFirstHint As String) As Long

    On Error GoTo PROC_ERR

    Dim rng As Word.Range
    Set rng = ActiveDocument.Content
    rng.Collapse wdCollapseStart

    With rng.Find
        .ClearFormatting
        .Font.Superscript = True
        .Text = strToken
        .Forward = True
        .MatchCase = False
        .MatchWholeWord = False
        .Wrap = wdFindStop
    End With

    Dim lngHits As Long
    Do While rng.Find.Execute
        lngHits = lngHits + 1
        If Len(strFirstHint) = 0 Then
            strFirstHint = BuildHint(rng, strToken)
        End If
        rng.Collapse wdCollapseEnd
    Loop

    CountSuperscriptToken = lngHits

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & _
           " (" & Err.Description & _
           ") in procedure CountSuperscriptToken of Module basTEST_NoSuperscriptOrdinals"
    Resume PROC_EXIT
End Function

'-----------------------------------------------------------------------
' Build a single-line hint identifying where the offending superscript
' was found. Uses Range.Information for page and line numbers.
'-----------------------------------------------------------------------
Private Function BuildHint(ByVal rng As Word.Range, _
                           ByVal strToken As String) As String
    On Error Resume Next

    Dim lngPage As Long
    Dim lngLine As Long
    Dim strPara As String

    lngPage = rng.Information(wdActiveEndPageNumber)
    lngLine = rng.Information(wdFirstCharacterLineNumber)
    strPara = Left$(rng.Paragraphs(1).Range.Text, 120)
    strPara = Replace(strPara, vbCr, " ")
    strPara = Replace(strPara, vbLf, " ")

    BuildHint = "token=""" & strToken & """ page=" & lngPage & _
                " line=" & lngLine & " para=""" & strPara & """"
End Function
