Attribute VB_Name = "basTestaeBibleClass"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

' Default Usage:
' Run all tests in immediate window:
'   RUN_THE_TESTS
' Run a specific test in immediate window:
'   RUN_THE_TESTS(1)
' Show debug output in immediate window:
'   RUN_THE_TESTS("varDebug")
' Version is set in BibleClassVERSION As String
'   BibleClassVERSION is found in Class Modules BibleClass
'

Public Function RUN_THE_TESTS(Optional ByVal varDebug As Variant) As Boolean
    On Error GoTo 0
    If IsMissing(varDebug) Then
        aeBibleClassTest
    ElseIf varDebug = "varDebug" Then
        Debug.Print "Running in varDebug Mode!"
        aeBibleClassTest varDebug:="varDebug"
    ElseIf VarType(varDebug) = vbInteger Then
        Debug.Print "@@@ varDebug = " & varDebug
        aeBibleClassTest varDebug:=varDebug
    End If
End Function

Public Function aeBibleClassTest(Optional ByVal varDebug As Variant) As Boolean

    On Error GoTo PROC_ERR

    Dim oWordBibleObjects As aeBibleClass
    Set oWordBibleObjects = New aeBibleClass

    Dim bln1 As Boolean

    If CStr(varDebug) = "Error 448" Then
        Debug.Print , "varDebug is Not Used"
    End If

    Debug.Print
    Debug.Print "vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv"
    Debug.Print "1. aeBibleClassTest => TheBibleClassTests"
    Debug.Print "aeBibleClassTest"
    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so no parameter is passed to TheBibleClassTests"
        Debug.Print , "DEBUGGING IS OFF"
        bln1 = oWordBibleObjects.TheBibleClassTests()
    ElseIf varDebug = "varDebug" Then
        Debug.Print , "varDebug IS NOT missing so blnDebug is set to True"
        bln1 = oWordBibleObjects.TheBibleClassTests("WithDebugging")
    ElseIf VarType(varDebug) = vbInteger Then
        Debug.Print "### Running Test " & varDebug, "varDebug = " & varDebug
        bln1 = oWordBibleObjects.TheBibleClassTests(varDebug)
    Else
        Debug.Print "Unexpected Parameter !!!"
    End If
    Debug.Print "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
    Debug.Print

PROC_EXIT:
    Exit Function

PROC_ERR:
    If Err = 6068 Then ' VBA Project Not Trusted - "Programmatic access to the Visual Basic Project is not trusted..."
        MsgBox "VBA Project Not Trusted", vbCritical, "aeBibleClassTest"
        Stop
        'Resume PROC_EXIT
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeBibleClassTest of Module basTestaeBibleClass"
        Resume PROC_EXIT
    End If

End Function

'=======================================================================================
' Procedure : GitAutoTagRelease
' Purpose   : Automatically tag a release and push it to GitHub from within Word.
' Notes     :
'   - Requires Git installed and configured on the system.
'   - Adjust sRepoPath, sTag, sBranch as needed for your project.
'   - Uses WScript.Shell to execute Git CLI commands in audit-safe batch.
'   - Tag is created with annotation (-a) and pushed to remote.
'   - Output is logged to Immediate Window for audit traceability.
'   - Extendable with error handling, tag existence checks, or version bump logic.
' Audit     :
'   - Logs both [TAG] and [PUSH] output to aid changelog verification.
'   - Safe to wrap inside session-aware macro runners or audit triggers.
'=======================================================================================
Sub GitAutoTagRelease()
    Const sTag As String = "v0.1.1"
    Const sMessage As String = "Release version 0.1.1"
    Const sBranch As String = "main" ' adjust if needed
    Const sRepoPath As String = "C:\adaept\aeBibleClass" ' local repo folder
    Dim shellCmd As String, cmdOutput As String
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")

    If GitTagExists(sRepoPath, sTag) Then
        MsgBox "Tag " & sTag & " already exists. Aborting push.", vbExclamation
        Exit Sub
    End If

    ' Navigate to repo and tag release
    shellCmd = "cmd.exe /c cd /d """ & sRepoPath & """ && git tag -a " & sTag & " -m """ & sMessage & """"
    cmdOutput = wsh.Exec(shellCmd).StdOut.ReadAll
    Debug.Print "[TAG] " & cmdOutput

    ' Push the tag to GitHub
    shellCmd = "cmd.exe /c cd /d """ & sRepoPath & """ && git push origin " & sTag
    cmdOutput = wsh.Exec(shellCmd).StdOut.ReadAll
    Debug.Print "[PUSH] " & cmdOutput

    MsgBox "Git tag " & sTag & " created and pushed successfully.", vbInformation
End Sub

'=======================================================================================
' Function  : GitTagExists
' Purpose   : Check whether a Git tag (e.g., "v0.1.1") already exists in the local repo.
' Notes     :
'   - Uses WScript.Shell to run `git tag` and check for a match.
'   - Returns True if the tag exists; False otherwise.
'   - Call before creating or pushing a new release tag to avoid conflicts.
'   - Logs result to Immediate Window for audit traceability.
' Audit     :
'   - Ensures tagging is idempotent and reversible.
'   - Extendable to check remote tags via `git ls-remote --tags`.
'=======================================================================================
Function GitTagExists(sRepoPath As String, sTag As String) As Boolean
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Dim cmd As String, execObj As Object, result As String

    cmd = "cmd.exe /c cd /d """ & sRepoPath & """ && git tag"
    Set execObj = wsh.Exec(cmd)
    result = execObj.StdOut.ReadAll

    If InStr(result, sTag) > 0 Then
        Debug.Print "[TAG CHECK] Tag '" & sTag & "' already exists."
        GitTagExists = True
    Else
        Debug.Print "[TAG CHECK] Tag '" & sTag & "' does not exist."
        GitTagExists = False
    End If
End Function


