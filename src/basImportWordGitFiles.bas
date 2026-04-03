Attribute VB_Name = "basImportWordGitFiles"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Public Sub ImportAllVBAFiles(Optional ByVal varDebug As Variant)
    On Error GoTo 0
    Dim strSrcPath As String
    Dim strFile As String
    Dim intImported As Integer
    Dim intSkipped As Integer
    Dim strExt As Variant
    Dim vbCompName As String

    strSrcPath = ThisDocument.Path & "\src\"

    ' Verify src folder exists
    If Dir(strSrcPath, vbDirectory) = "" Then
        MsgBox "Source folder not found:" & vbCrLf & strSrcPath, vbCritical, "Import Aborted"
        Exit Sub
    End If

    ' Delete all modules except this one - prompts for confirmation
    If Not DeleteAllModulesExceptImporter() Then
        Debug.Print "Import aborted - deletion cancelled.", "in Sub ImportAllVBAFiles"
        Exit Sub
    End If

    ' Collect all file paths BEFORE importing - Dir() is not reentrant
    Dim colFiles As Collection
    Set colFiles = New Collection
    For Each strExt In Array("*.bas", "*.cls", "*.frm")
        strFile = Dir(strSrcPath & strExt)
        Do While strFile <> ""
            colFiles.Add strSrcPath & strFile
            strFile = Dir()
        Loop
    Next strExt

    ' Now import from the collected list
    intImported = 0
    intSkipped = 0

    Dim colSkipped As Collection
    Set colSkipped = New Collection

    Dim strFullPath As Variant
    For Each strFullPath In colFiles
        strFile = Mid$(strFullPath, InStrRev(strFullPath, "\") + 1)
        If strFile <> "ThisDocument.cls" Then
            vbCompName = Left(strFile, InStrRev(strFile, ".") - 1)
            If Not ModuleOrClassExists(vbCompName) Then
                ImportVBAFile CStr(strFullPath)
                intImported = intImported + 1
            Else
                colSkipped.Add vbCompName & " (already exists)"
                intSkipped = intSkipped + 1
            End If
        Else
            colSkipped.Add strFile & " (ThisDocument)"
            intSkipped = intSkipped + 1
        End If
    Next strFullPath

    Debug.Print "Import complete.", _
                "Imported: " & intImported, _
                "Skipped: " & intSkipped, _
                "in Sub ImportAllVBAFiles"

    Dim strSkippedItem As Variant
    Dim strSkippedList As String
    strSkippedList = ""
    For Each strSkippedItem In colSkipped
        Debug.Print "  skipped:", CStr(strSkippedItem), "in Sub ImportAllVBAFiles"
        strSkippedList = strSkippedList & vbCrLf & "  " & strSkippedItem
    Next strSkippedItem

    Dim strMsgSkipped As String
    strMsgSkipped = ""
    If intSkipped > 0 Then
        strMsgSkipped = vbCrLf & vbCrLf & "Skipped files:" & strSkippedList
    End If

    MsgBox "Import complete." & vbCrLf & vbCrLf & _
           "Imported: " & intImported & vbCrLf & _
           "Skipped:  " & intSkipped & strMsgSkipped, _
           vbInformation, "Import Complete"
End Sub

Private Sub ImportVBAFile(myCodeFile As String)
    On Error GoTo PROC_ERR
    Dim vbaModule As Object
    Dim filePath, fileName, fullPath, vbCompName As String

    ' Set the file path of the exported VBA source file
    ' fullPath = "C:\path\to\your\exported\file.bas" ' Change this to the actual path of your .bas or .cls file
    fullPath = myCodeFile
    ' Get the file name using VBA built-in functions
    fileName = Mid$(fullPath, InStrRev(fullPath, "\") + 1)
    ' Remove the extension
    vbCompName = Left(fileName, InStrRev(fileName, ".") - 1)

    ' Check if the source file exists
    If Dir(fullPath) <> "" Then
        ' Import the VBA source file into the current document
        'Debug.Print "ModuleOrClassExists(vbCompName) = " & ModuleOrClassExists(vbCompName)
        If Not ModuleOrClassExists(vbCompName) Then
            Set vbaModule = ThisDocument.VBProject.VBComponents.Import(fullPath)
            Debug.Print vbCompName, "import SUCCESS!", "in Sub ImportVBAFile"
        Else
            Debug.Print vbCompName, "import ABORTED!", "in Sub ImportVBAFile"
        End If
    End If

PROC_EXIT:
    Exit Sub
PROC_ERR:
    If Err = 6068 Then
        MsgBox "VBA Project Not Trusted" & vbCrLf & "Enable 'Trust access to the VBA project object model' in Word Trust Center.", vbCritical, "ImportVBAFile"
        Stop
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in Sub ImportVBAFile", vbCritical, "ImportVBAFile"
        Resume PROC_EXIT
    End If
End Sub

Public Function DeleteAllModulesExceptImporter() As Boolean
    On Error GoTo PROC_ERR
    Dim vbComp As Object
    Dim strProtected As String
    Dim strToDelete As String
    Dim strMsg As String
    Dim intResponse As Integer

    strProtected = "basImportWordGitFiles"

    ' Build list of modules to be deleted for confirmation prompt
    strToDelete = ""
    For Each vbComp In ThisDocument.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1, 2, 3
                If vbComp.name <> strProtected Then
                    strToDelete = strToDelete & "  " & vbComp.name & vbCrLf
                End If
        End Select
    Next vbComp

    If strToDelete = "" Then
        MsgBox "No modules found to delete.", vbInformation, "Delete Modules"
        DeleteAllModulesExceptImporter = True
        Exit Function
    End If

    ' Prompt for confirmation
    strMsg = "The following modules and classes will be deleted:" & vbCrLf & vbCrLf & _
             strToDelete & vbCrLf & _
             "'" & strProtected & "' will be preserved." & vbCrLf & vbCrLf & _
             "Proceed with deletion?"
    intResponse = MsgBox(strMsg, vbYesNo + vbExclamation, "Confirm Delete")

    If intResponse <> vbYes Then
        Debug.Print "Deletion cancelled by user.", "in Function DeleteAllModulesExceptImporter"
        DeleteAllModulesExceptImporter = False
        Exit Function
    End If

    ' Collect names to delete first - never modify a collection while iterating it
    Dim colToDelete As Collection
    Set colToDelete = New Collection
    For Each vbComp In ThisDocument.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1, 2, 3
                If vbComp.name <> strProtected Then
                    colToDelete.Add vbComp.name
                End If
        End Select
    Next vbComp

    ' Now delete using the collected names
    Dim strName As Variant
    For Each strName In colToDelete
        Set vbComp = ThisDocument.VBProject.VBComponents(strName)
        ThisDocument.VBProject.VBComponents.Remove vbComp
        Debug.Print strName, "deleted", "in Function DeleteAllModulesExceptImporter"
    Next strName

    DeleteAllModulesExceptImporter = True

PROC_EXIT:
    Exit Function
PROC_ERR:
    If Err = 6068 Then
        MsgBox "VBA Project Not Trusted" & vbCrLf & "Enable 'Trust access to the VBA project object model' in Word Trust Center.", vbCritical, "DeleteAllModulesExceptImporter"
        Stop
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in Function DeleteAllModulesExceptImporter", vbCritical, "DeleteAllModulesExceptImporter"
        Resume PROC_EXIT
    End If
End Function

Private Function ModuleOrClassExists(name As String) As Boolean
    On Error GoTo 0
    Dim vbComp As Object
    Dim found As Boolean
    
    found = False
    'Debug.Print "name = " & name, "in Function ModuleOrClassExists"
    For Each vbComp In ThisDocument.VBProject.VBComponents
        If vbComp.name = name Then
            found = True
            Exit For
        End If
    Next vbComp
    
    ModuleOrClassExists = found
    Debug.Print name, "ModuleOrClassExists = " & found, "in Function ModuleOrClassExists"
End Function
