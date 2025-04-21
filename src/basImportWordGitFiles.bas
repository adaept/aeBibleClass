Attribute VB_Name = "basImportWordGitFiles"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Public Sub ImportWordBibleFiles()
    'Call ImportVBAFile("C:\adaept\Bible\src\aeBibleClass.cls")
    'Call ImportVBAFile("C:\adaept\Bible\src\basChangeLogaeBibleClass.bas")
    'Call ImportVBAFile("C:\adaept\Bible\src\basTESTaeBibleClass.bas")
End Sub

Public Sub ImportWordGitFiles()
'    Call ImportVBAFile("C:\adaept\aewordgit\src\aewordgitClass.cls")
    'Call ImportVBAFile("C:\adaept\aewordgit\src\basChangeLogaewordgit.bas")
    'Call ImportVBAFile("C:\adaept\aewordgit\src\basImportWordGitFiles.bas")
'    Call ImportVBAFile("C:\adaept\aewordgit\src\basTESTaewordgitClass.bas")
End Sub

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

Private Sub ImportVBAFile(myCodeFile As String)
    On Error GoTo 0
    Dim vbaModule As Object
    Dim filePath, fileName, fullPath, vbCompName As String
    
    ' Set the file path of the exported VBA source file
    ' fullPath = "C:\path\to\your\exported\file.bas" ' Change this to the actual path of your .bas or .cls file
    fullPath = myCodeFile
    ' Get the file name using VBA built-in functions
    fileName = Mid(fullPath, InStrRev(fullPath, "\") + 1)
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
End Sub

