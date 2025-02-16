Attribute VB_Name = "basTestBibleClass"
Option Explicit
Option Compare Text
Option Private Module

' Default Usage:
' Run in immediate window:
'   RUN_THE_TESTS
' Show debug output in immediate window:
'   RUN_THE_TESTS("varDebug")
' Version is set in BibleClassVERSION As String
'   BibleClassVERSION is found in Class Modules BibleClass
'

Public Function RUN_THE_TESTS(Optional ByVal varDebug As Variant) As Boolean
    On Error GoTo 0
    If IsMissing(varDebug) Then
        BibleClassTest
    Else
        BibleClassTest varDebug:="varDebug"
    End If
End Function

Public Function BibleClassTest(Optional ByVal varDebug As Variant) As Boolean

    On Error GoTo PROC_ERR

    Dim oWordBibleObjects As BibleClass
    Set oWordBibleObjects = New BibleClass

    Dim bln1 As Boolean

    If CStr(varDebug) = "Error 448" Then
        Debug.Print , "varDebug is Not Used"
    End If

Test1:
    '=============
    ' TEST 1
    '=============
    Debug.Print
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "1. BibleClassTest => TheBibleClassTests"
    Debug.Print "BibleClassTest"
    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to TheBibleClassTests"
        Debug.Print , "DEBUGGING IS OFF"
        bln1 = oWordBibleObjects.TheBibleClassTests()
    Else
        Debug.Print , "varDebug IS NOT missing so blnDebug is set to True"
        bln1 = oWordBibleObjects.TheBibleClassTests("WithDebugging")
    End If
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print

PROC_EXIT:
    Exit Function

PROC_ERR:
    If Err = 6068 Then ' VBA Project Not Trusted - "Programmatic access to the Visual Basic Project is not trusted..."
        MsgBox "VBA Project Not Trusted", vbCritical, "BibleClassTest"
        Stop
        'Resume PROC_EXIT
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure BibleClassTest of Module basTestBibleClass"
        Resume PROC_EXIT
    End If

End Function




