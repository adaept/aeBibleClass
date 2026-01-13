Attribute VB_Name = "basImportWordGitFiles"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

' =====================================================================================
' Procedure : DeleteAllCodeModules_LateBinding
' Author    : Peter
' Purpose   :
'   Deletes ALL user-created VBA components (standard modules, class modules, and
'   UserForms) from the current DOCM/DOTM using late binding so no VBIDE reference
'   is required.
'
'   A Y/N confirmation prompt is shown:
'       - Default is N (pressing Enter or typing anything except "Y")
'       - If N is selected, no code is deleted and a message is shown:
'             "No Code Deleted !!!"
'       - If Y is selected, all deletable components are removed.
'
' Notes:
'   - Requires: Word Trust Center ? "Trust access to the VBA project object model"
'   - Does NOT delete built-in components such as ThisDocument.
'   - Safe for DOCM and DOTM projects.
'   - Uses late binding for maximum portability and zero reference dependencies.
'
' =====================================================================================
Public Sub DeleteAllCodeModules_LateBinding()

    Dim resp As String
    Dim proj As Object
    Dim vbComps As Object
    Dim vbComp As Object
    Dim compType As Long

    ' Prompt user
    resp = InputBox( _
        "Delete ALL modules and class modules from this DOCM?" & vbCrLf & _
        "Type Y to confirm. Default is N.", _
        "Delete Code?")

    If UCase$(Trim$(resp)) <> "Y" Then
        MsgBox "No Code Deleted !!!", vbInformation, "Cancelled"
        Exit Sub
    End If

    ' Late binding to VBProject
    Set proj = ThisDocument.VBProject
    Set vbComps = proj.VBComponents

    ' Component type constants (late-bound)
    Const vbext_ct_StdModule As Long = 1
    Const vbext_ct_ClassModule As Long = 2
    Const vbext_ct_MSForm As Long = 3

    ' Delete modules, classes, and forms
    For Each vbComp In vbComps
        compType = vbComp.Type
        If compType = vbext_ct_StdModule _
        Or compType = vbext_ct_ClassModule _
        Or compType = vbext_ct_MSForm Then
            vbComps.Remove vbComp
        End If
    Next vbComp

    MsgBox "All modules and class modules deleted.", vbInformation, "Done"

End Sub

' =====================================================================================
' Procedure : ReImportAllCodeModules_LateBinding
' Author    : Peter
' Purpose   :
'   Imports all .bas, .cls, and .frm files from a user-selected folder into the
'   current Word VBA project. Uses late binding so no VBIDE reference is required.
'
'   Intended for use after deleting all modules (e.g., to force full re-tokenization).
'
' Notes:
'   - Requires: Word Trust Center ? "Trust access to the VBA project object model"
'   - Imports only user-created components (standard modules, class modules, forms)
'   - Does NOT modify ThisDocument or built-in components
'   - Safe for DOCM and DOTM projects
'
' =====================================================================================
Public Sub ReImportAllCodeModules_LateBinding()

    Dim proj As Object
    Dim vbComps As Object
    Dim fDialog As FileDialog
    Dim folderPath As String
    Dim file As String

    ' Select folder containing exported modules
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .Title = "Select folder containing exported VBA modules"
        If .Show <> -1 Then
            MsgBox "No folder selected. Import cancelled.", vbInformation, "Cancelled"
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Set proj = ThisDocument.VBProject
    Set vbComps = proj.VBComponents

    ' Import .bas (standard modules)
    file = Dir(folderPath & "*.bas")
    Do While file <> ""
        vbComps.Import folderPath & file
        file = Dir()
    Loop

    ' Import .cls (class modules)
    file = Dir(folderPath & "*.cls")
    Do While file <> ""
        vbComps.Import folderPath & file
        file = Dir()
    Loop

    ' Import .frm (UserForms)
    file = Dir(folderPath & "*.frm")
    Do While file <> ""
        vbComps.Import folderPath & file
        file = Dir()
    Loop

    MsgBox "All modules, classes, and forms imported.", vbInformation, "Import Complete"

End Sub

Public Sub ReloadAllWordBibleFilesV59()
    Call ImportVBAFile("C:\adaept\aeBibleClass\src\aeBibleRibbon.bas")
    Call ImportVBAFile("C:\adaept\aeBibleClass\src\basChangeLogaeBibleClass.bas")
    Call ImportVBAFile("C:\adaept\aeBibleClass\src\basChangeLogaewordgit.bas")
    ' PLACE HOLDER: Import the file basImportWordGitFiles.bas MANUALLY for the command
    'Call ImportVBAFile("C:\adaept\aeBibleClass\src\basImportWordGitFiles.bas")
    Call ImportVBAFile("C:\adaept\aeBibleClass\src\basTESTaeBibleClass.bas")
    Call ImportVBAFile("C:\adaept\aeBibleClass\src\basTESTaeBibleFonts.bas")
    Call ImportVBAFile("C:\adaept\aeBibleClass\src\basTESTaeBibleTools.bas")
    Call ImportVBAFile("C:\adaept\aeBibleClass\src\basTESTaewordgitClass.bas")
    Call ImportVBAFile("C:\adaept\aeBibleClass\src\basUSFMExport.bas")
    Call ImportVBAFile("C:\adaept\aeBibleClass\src\basWordRepairRunner.bas")
    Call ImportVBAFile("C:\adaept\aeBibleClass\src\basWordSettingsDiagnostic.bas")
    Call ImportVBAFile("C:\adaept\aeBibleClass\src\Module1.bas")
    Call ImportVBAFile("C:\adaept\aeBibleClass\src\Modules2.bas")
    Call ImportVBAFile("C:\adaept\aeBibleClass\src\XbasTESTaeBibleClass_SLOW.bas")
    Call ImportVBAFile("C:\adaept\aeBibleClass\src\XbasTESTaeBibleDOCVARIABLE.bas")
    Call ImportVBAFile("C:\adaept\aeBibleClass\src\XLongRunningProcessCode.bas")
    Call ImportVBAFile("C:\adaept\aeBibleClass\src\aeBibleClass.cls")
    Call ImportVBAFile("C:\adaept\aeBibleClass\src\aewordgitClass.cls")
End Sub

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
    fileName = mid(fullPath, InStrRev(fullPath, "\") + 1)
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

