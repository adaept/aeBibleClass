VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aeAssertClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mTestsRun As Long
Private mTestsFailed As Long
Private mInitialized As Boolean

'===========================================================
' Initialize
'===========================================================
Public Sub Initialize()
    If mInitialized Then Exit Sub   ' <-- prevents reset

    mTestsRun = 0
    mTestsFailed = 0
    mInitialized = True

    Debug.Print "=========================================="
    Debug.Print " SBL PARSER TEST HARNESS"
    Debug.Print "=========================================="
End Sub

'===========================================================
' Terminate
'===========================================================
Public Sub Terminate()
    If Not mInitialized Then Exit Sub

    Debug.Print "------------------------------------------"
    Debug.Print " TEST SUMMARY"
    Debug.Print "------------------------------------------"

    Debug.Print "Tests Run: "; mTestsRun
    Debug.Print "Failures:  "; mTestsFailed

    If mTestsFailed = 0 Then
        Debug.Print "RESULT: PASS"
    Else
        Debug.Print "RESULT: FAIL"
    End If

    mInitialized = False
End Sub

'===========================================================
' AssertTrue
'===========================================================
Public Sub AssertTrue( _
        ByVal condition As Boolean, _
        ByVal message As String, _
        Optional ByVal expected As Variant, _
        Optional ByVal actual As Variant)

    On Error GoTo PROC_ERR

    mTestsRun = mTestsRun + 1

    If condition Then
        Debug.Print "PASS: " & message
    Else
        mTestsFailed = mTestsFailed + 1

        If IsMissing(expected) Then
            Debug.Print "FAIL: " & message
        Else
            Debug.Print "FAIL: " & message & _
                        " | Expected=" & expected & _
                        " Actual=" & actual
        End If
    End If

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & _
           " Error " & Err.Number & _
           " (" & Err.Description & _
           ") in aeAssertClass.AssertTrue"
    Resume PROC_EXIT
End Sub

'===========================================================
' AssertEqual
'===========================================================
Public Sub AssertEqual( _
        ByVal expected As Variant, _
        ByVal actual As Variant, _
        ByVal label As String)

    On Error GoTo PROC_ERR

    mTestsRun = mTestsRun + 1

    If expected = actual Then
        Debug.Print "PASS: "; label
    Else
        Debug.Print "FAIL: "; label; _
                    " (expected="; expected; _
                    ", actual="; actual; ")"
        mTestsFailed = mTestsFailed + 1
    End If

PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & _
           " Error " & Err.Number & _
           " (" & Err.Description & _
           ") in aeAssertClass.AssertEqual"
    Resume PROC_EXIT
End Sub

'===========================================================
' AssertFalse
'===========================================================
Public Sub AssertFalse( _
        ByVal condition As Boolean, _
        ByVal message As String)

    AssertTrue Not condition, message
End Sub


