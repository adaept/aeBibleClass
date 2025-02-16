VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aewordgitClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'Copyright (c) 2018-2025 Peter F. Ennis
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation;
'version 3.0.
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'Lesser General Public License for more details.
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, visit
'http://www.gnu.org/licenses/lgpl-3.0.txt

'=======================================================================
' Author:   Peter F. Ennis
' Date:     September 3, 2018
' Comment:  Create class for revision control of word VBA
' History:  See comment details, basChangeLog, commit messages on GitHub
' GitHub:   https://github.com/peterennis/aewordgit
'=======================================================================

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)

Private Const aewordgitVERSION As String = "0.0.6"
Private Const aewordgitVERSION_DATE As String = "February 10, 2025"

Private Type mySetupType
    SourceFolder As String
    UserSourceFolder As String
    ImportFolder As String
End Type

' Ref: http://www.pcreview.co.uk/forums/macro-save-macros-t1918984p2.html
Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100

Private aewordgitType As mySetupType
Private aewordgitSourceFolder As String
Private aewordgitUserSourceFolder As String

Private Sub Class_Initialize()
    ' Ref: http://www.cadalyst.com/cad/autocad/programming-with-class-part-1-5050
    ' Ref: http://www.bigresource.com/Tracker/Track-vb-cyJ1aJEyKj/
    ' Ref: http://stackoverflow.com/questions/1731052/is-there-a-way-to-overload-the-constructor-initialize-procedure-for-a-class-in

    On Error GoTo 0
    'MsgBox Application.VBE.ActiveVBProject.Name
    'MsgBox Application.Caption
    If Application.VBE.ActiveVBProject.name = "aewordgit" Then
        Application.Caption = Application.VBE.ActiveVBProject.name & " " & aewordgitVERSION
    End If
    ' Provide a default value for the SourceFolder and ImportFolder properties
    aewordgitSourceFolder = "default"
    aewordgitType.SourceFolder = "C:\adaept\aewordgit\src\"
    aewordgitType.UserSourceFolder = ActiveDocument.Path & "\src\"

    Debug.Print "Class_Initialize"
    Debug.Print , "Default for aewordgitSourceFolder = " & aewordgitSourceFolder
    Debug.Print , "Default for aewordgitType.SourceFolder = " & aewordgitType.SourceFolder
    Debug.Print , "Default for aewordgitType.UserSourceFolder = " & aewordgitType.UserSourceFolder
End Sub

Private Sub Class_Terminate()
    On Error GoTo 0
    Debug.Print
    Debug.Print "Class_Terminate"
    Debug.Print , "aewordgit VERSION: " & aewordgitVERSION
    Debug.Print , "aewordgit VERSION_DATE: " & aewordgitVERSION_DATE
End Sub

Public Property Get SourceFolder() As String
    On Error GoTo 0
    SourceFolder = aewordgitSourceFolder
End Property

Public Property Let SourceFolder(ByVal strSourceFolder As String)
    On Error GoTo 0
    ' Ref: http://www.techrepublic.com/article/build-your-skills-using-class-modules-in-an-access-database-solution/5031814
    ' Ref: http://www.utteraccess.com/wiki/index.php/Classes
    aewordgitSourceFolder = strSourceFolder
End Property

Public Property Get DocumentTheWordCode(Optional ByVal varDebug As Variant) As Boolean

    On Error GoTo PROC_ERR

    'MsgBox "DocumentTheWordCode"
    If IsMissing(varDebug) Then
        Debug.Print "Get DocumentTheWordCode"
        Debug.Print , "varDebug IS missing so no parameter is passed to aeDocumentTheWordCode"
        Debug.Print , "DEBUGGING IS OFF"
        DocumentTheWordCode = aeDocumentTheWordCode
    Else
        Debug.Print "Get DocumentTheWordCode"
        Debug.Print , "varDebug IS NOT missing so a variant parameter is passed to aeDocumentTheWordCode"
        Debug.Print , "DEBUGGING TURNED ON"
        DocumentTheWordCode = aeDocumentTheWordCode(varDebug)
    End If

PROC_EXIT:
    Exit Property

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure DocumentTheWordCode of Class aewordgitClass"
    Resume PROC_EXIT

End Property

Private Function ThisIsAnAddIn(ByVal strAddIn As String, Optional ByVal varDebug As Variant) As Boolean
    Dim rw As Byte
    Dim varAddIn As Variant
    ThisIsAnAddIn = False
    rw = 1
    For Each varAddIn In Application.AddIns
        If Not IsMissing(varDebug) Then Debug.Print varAddIn.name, varAddIn.Installed
        If varAddIn.name = strAddIn Then
            ThisIsAnAddIn = True
            Exit Function
        Else
            rw = rw + 1
        End If
    Next
End Function

Private Function aeDocumentTheWordCode(Optional ByVal varDebug As Variant) As Boolean

    Dim strExportPath As String
    Dim strFileName As String

    'MsgBox "aeDocumentTheWordCode"
    On Error GoTo PROC_ERR

    Dim objComponent As Object
    Set objComponent = ThisDocument.VBProject.VBComponents

    Debug.Print "aeDocumentTheWordCode"
    If IsMissing(varDebug) Then
        Debug.Print , "varDebug IS missing so blnDebug of aeDocumentTheWordCode is set to False"
        Debug.Print , "DEBUGGING IS OFF"
    Else
        Debug.Print , "varDebug IS NOT missing so blnDebug of aeDocumentTheWordCode is set to True"
        Debug.Print , "NOW DEBUGGING..."
    End If
    
    Dim response As VbMsgBoxResult
    
    ' Display the message box with "Yes" and "No" buttons
    response = MsgBox("Files will be deleted from" & vbCrLf & _
        FolderWithVBAProjectFiles & "*.*", vbYesNo + vbQuestion, "aeDocumentTheWordCode")
    ' Handle the user's response
    Select Case response
        Case vbYes
            'MsgBox "You chose Yes.", vbInformation, "Response"
            ' Add your code here for when the user selects Yes
            On Error Resume Next
            Kill FolderWithVBAProjectFiles & "*.*"
            On Error GoTo PROC_ERR
        Case vbNo
            'MsgBox "You chose No.", vbInformation, "Response"
            ' Add your code here for when the user selects No
            Exit Function
    End Select

    RunExportCode FolderWithVBAProjectFiles

    aeDocumentTheWordCode = True

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure aeDocumentTheWordCode of Class aewordgitClass"
    Resume PROC_EXIT

End Function

Private Sub RunExportCode(ByVal strExportPath As String)

    Dim strFileName As String
    Dim strDocName As String
    Dim docSource As Object
    Set docSource = ThisDocument

    On Error GoTo PROC_ERR

    ' NOTE: The document must be open in Word
    strDocName = ActiveDocument.name
    Debug.Print , "strDocName = " & strDocName
    'Stop

    Dim objComponent As Object
    Set objComponent = ThisDocument.VBProject.VBComponents

    If docSource.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected!" & vbCrLf & _
            "It is not possible to export the code."
        Exit Sub
    End If

    For Each objComponent In docSource.VBProject.VBComponents

        strFileName = objComponent.name
        Debug.Print , "strFileName = " & strFileName

        ' Concatenate the correct filename for export.
        Select Case objComponent.Type
            Case vbext_ct_ClassModule
                strFileName = strFileName & ".cls"
            Case vbext_ct_MSForm
                strFileName = strFileName & ".frm"
            Case vbext_ct_StdModule
                strFileName = strFileName & ".bas"
            Case vbext_ct_Document
                strFileName = strFileName & ".cls"
            Case Else
                MsgBox "Type not known: " & objComponent.Type
        End Select

        ' Export the component to a text file.
        If Left$(strFileName, 3) <> "zzz" Then
            objComponent.Export strExportPath & strFileName
        End If

        ' remove it from the project if you want
        'wkbSource.VBProject.VBComponents.Remove objComponent

    Next objComponent

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure RunExportCode of Class aewordgitClass"
    Resume PROC_EXIT

End Sub

Private Sub OutputListOfWordProperties(Optional ByVal varDebug As Variant)
    ' Ref: http://stackoverflow.com/questions/17406585/vba-set-custom-document-property
    ' Workbook.BuiltinDocumentProperties Property
    ' Ref: http://msdn.microsoft.com/en-us/library/bb220896.aspx

    On Error GoTo PROC_ERR

    Dim wdoc As Document
    Set wdoc = Application.ActiveDocument

    Dim docProp As DocumentProperty
    Dim strExportPathFileName As String
    Dim fle As Integer

    fle = FreeFile()
    strExportPathFileName = FolderWithVBAProjectFiles & "\OutputListOfWordProperties.txt"
    Open strExportPathFileName For Output As #fle

    If Not IsMissing(varDebug) Then Debug.Print "OutputListOfWordProperties"

    ' Ref: http://stackoverflow.com/questions/16642362/how-to-get-the-following-code-to-continue-on-error
    For Each docProp In wdoc.CustomDocumentProperties
        If Not IsMissing(varDebug) Then Debug.Print , docProp.name & ": " & docProp.Value
        Print #fle, docProp.name & ": " & docProp.Value
    Next

    For Each docProp In ActiveDocument.BuiltInDocumentProperties
        If Not IsMissing(varDebug) Then Debug.Print , docProp.name & ": " & docProp.Value
        Print #fle, docProp.name & ": " & docProp.Value
    Next

PROC_EXIT:
    Close fle
    Set wdoc = Nothing
    Exit Sub

PROC_ERR:
    If Err = -2147467259 Then     ' Method 'Value' of object ... failed
        'MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfWordProperties of Class aewordgitClass"
        Resume Next
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure OutputListOfWordProperties of Class aewordgitClass"
        Resume PROC_EXIT
    End If

End Sub

Private Function FolderWithVBAProjectFiles() As String

    On Error GoTo PROC_ERR
    
    Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath, fullPath, directoryName As String

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("Scripting.FileSystemObject")

    SpecialPath = WshShell.SpecialFolders("MyDocuments")

    'Debug.Print "A. FolderWithVBAProjectFiles", "aewordgitSourceFolder = " & aewordgitSourceFolder

    ' Get the full path of the active document
    fullPath = ActiveDocument.Path
    ' Extract the directory name from the full path
    directoryName = Mid(fullPath, InStrRev(fullPath, "\") + 1)

    If aewordgitSourceFolder = "default" And directoryName = "aewordgit" Then
        FolderWithVBAProjectFiles = aewordgitType.SourceFolder
        'Debug.Print "B. FolderWithVBAProjectFiles", "FolderWithVBAProjectFiles = " & aewordgitType.SourceFolder
    ElseIf aewordgitSourceFolder = "default" And directoryName <> "aewordgit" Then
        FolderWithVBAProjectFiles = aewordgitType.UserSourceFolder
        'Debug.Print "C. FolderWithVBAProjectFiles", "FolderWithVBAProjectFiles = " & FolderWithVBAProjectFiles
    ElseIf IsNull(aewordgitSourceFolder) Then
        If Right$(SpecialPath, 1) <> "\" Then SpecialPath = SpecialPath & "\"

        If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
            On Error Resume Next
            MkDir SpecialPath & "VBAProjectFiles"
            On Error GoTo 0
        End If

        If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
            FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
        Else
            FolderWithVBAProjectFiles = "Error"
        End If

    Else
        FolderWithVBAProjectFiles = aewordgitSourceFolder
    End If
    
    'MsgBox "FolderWithVBAProjectFiles=" & FolderWithVBAProjectFiles
    If Not FSO.FolderExists(FolderWithVBAProjectFiles) Then
        MsgBox "FolderWithVBAProjectFiles Error" & vbCrLf & _
            "FolderWithVBAProjectFiles = " & FolderWithVBAProjectFiles & vbCrLf & _
            "Does NOT exist!", vbCritical, "FolderWithVBAProjectFiles"
        FolderWithVBAProjectFiles = "Error"
    End If
    
    'Debug.Print "D. FolderWithVBAProjectFiles", "FolderWithVBAProjectFiles = " & FolderWithVBAProjectFiles
    'Stop

PROC_EXIT:
    Exit Function

PROC_ERR:
    If Err = -2147467259 Then     ' Method 'Value' of object ... failed
        'MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in function FolderWithVBAProjectFiles of Class aewordgitClass"
        Resume Next
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in function FolderWithVBAProjectFiles of Class aewordgitClass"
        Resume PROC_EXIT
    End If

End Function

Private Function DeleteVBAModulesAndUserForms() As Boolean

    On Error GoTo 0
    Dim objProject As Object
    Set objProject = CreateObject("VBIDE.VBProject")

    Dim objComponent As Object
    Set objComponent = CreateObject("VBIDE.VBComponent")

    Set objProject = ActiveDocument.VBProject

    For Each objComponent In objProject.VBComponents
        If objComponent.Type = vbext_ct_Document Then
            ' ThisDocument we do nothing
        Else
            objProject.VBComponents.Remove objComponent
        End If
    Next objComponent

End Function


