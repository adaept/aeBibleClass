Attribute VB_Name = "basSBL_TestFramework"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Private gTestsRun As Long
Private gTestsFailed As Long

Public Sub AssertTrue( _
        ByVal condition As Boolean, _
        ByVal message As String, _
        Optional ByVal expected As Variant, _
        Optional ByVal actual As Variant)

    gTestsRun = gTestsRun + 1
    If condition Then
        Debug.Print "PASS: " & message
    Else
        gTestsFailed = gTestsFailed + 1
        If IsMissing(expected) Then
            Debug.Print "FAIL: " & message
        Else
            Debug.Print "FAIL: " & message & _
                        " | Expected=" & expected & _
                        " Actual=" & actual
        End If
    End If
End Sub

Public Sub AssertEqual(expected As Variant, actual As Variant, label As String)
' Useful for numeric/string comparisons.
    gTestsRun = gTestsRun + 1
    
    If expected = actual Then
        Debug.Print "PASS: "; label
    Else
        Debug.Print "FAIL: "; label; _
                    " (expected="; expected; _
                    ", actual="; actual; ")"
        gTestsFailed = gTestsFailed + 1
    End If
End Sub

'===========================================================
' AssertFalse
'===========================================================
Public Sub AssertFalse(ByVal condition As Boolean, ByVal message As String)
    AssertTrue Not condition, message
End Sub

Public Sub TestStart()
    gTestsRun = 0
    gTestsFailed = 0
    
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print " SBL PARSER TEST HARNESS"
    Debug.Print "=========================================="
End Sub

Public Sub TestSummary()
    Debug.Print ""
    Debug.Print "------------------------------------------"
    Debug.Print " TEST SUMMARY"
    Debug.Print "------------------------------------------"
    
    Debug.Print "Tests Run: "; gTestsRun
    Debug.Print "Failures:  "; gTestsFailed
    
    If gTestsFailed = 0 Then
        Debug.Print "RESULT: PASS"
    Else
        Debug.Print "RESULT: FAIL"
    End If
End Sub

