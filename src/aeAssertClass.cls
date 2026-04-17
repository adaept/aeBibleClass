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

'Copyright (c) 2025-2026 Peter F. Ennis
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation;
'version 3.0, or (at your option) any later version.
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'Lesser General Public License for more details.
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, visit
'https://www.gnu.org/licenses/lgpl-3.0.txt

' ================================================================
' Author:   Peter F. Ennis
' Date:     March 10, 2025
' Comment:  Assert Class for Word 365 VBA
' History:  See comment details, basChangeLog_aeBibleClass, commit messages on GitHub
' GitHub:   https://github.com/adaept/aeBibleClass
' ================================================================
Private mTestsRun    As Long
Private mTestsFailed As Long
Private mInitialized As Boolean
Private m_log        As aeLoggerClass

'===========================================================
' SetLogger - optional; mirrors all assert output to log file
'===========================================================
Public Sub SetLogger(log As aeLoggerClass)
    Set m_log = log
End Sub

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
    If Not m_log Is Nothing Then
        m_log.Log_Write "=========================================="
        m_log.Log_Write " SBL PARSER TEST HARNESS"
        m_log.Log_Write "=========================================="
    End If
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
        Debug.Print "Result: PASS"
    Else
        Debug.Print "Result: FAIL"
    End If

    If Not m_log Is Nothing Then
        m_log.Log_Write "------------------------------------------"
        m_log.Log_Write " TEST SUMMARY"
        m_log.Log_Write "------------------------------------------"
        m_log.Log_Write "Tests Run: " & mTestsRun
        m_log.Log_Write "Failures:  " & mTestsFailed
        If mTestsFailed = 0 Then
            m_log.Log_Write "Result: PASS"
        Else
            m_log.Log_Write "Result: FAIL"
        End If
        Set m_log = Nothing
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

    Dim assertLine As String
    If condition Then
        assertLine = "PASS: " & message
    Else
        mTestsFailed = mTestsFailed + 1
        If IsMissing(expected) Then
            assertLine = "FAIL: " & message
        Else
            assertLine = "FAIL: " & message & _
                         " | Expected=" & expected & _
                         " Actual=" & actual
        End If
    End If
    Debug.Print assertLine
    If Not m_log Is Nothing Then m_log.Log_Write assertLine

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

    Dim assertLine As String
    If expected = actual Then
        assertLine = "PASS: " & label
    Else
        assertLine = "FAIL: " & label & _
                     " (expected=" & expected & _
                     ", actual=" & actual & ")"
        mTestsFailed = mTestsFailed + 1
    End If
    Debug.Print assertLine
    If Not m_log Is Nothing Then m_log.Log_Write assertLine

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


