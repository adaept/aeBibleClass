Attribute VB_Name = "basImportWordRibbonGitFiles"
Option Explicit
Option Compare Text
Option Private Module

' =============================================================================
' basImportWordRibbonGitFiles - aeRibbon.dotm template build bootstrap
' -----------------------------------------------------------------------------
' Template-only counterpart to the dev docm's basImportWordGitFiles.bas. Lives
' permanently inside aeRibbon.dotm's VBA project (imported once, manually, via
' VBE File -> Import File - a normal file, self-registering after that first
' bootstrap, same pattern as its dev-docm counterpart protecting itself from
' its own delete-all step).
'
' Run ImportAllRibbonVBAFiles for every rebuild. It:
'   1. Deletes every other module/class in the live template (confirmation
'      prompt first) - closes the "does the doc say to delete first" gap in
'      BUILD.md; previously a manual, undocumented assumption.
'   2. Re-imports every *.bas/*.cls/*.frm found in aeRibbon/src/ (the trim
'      pipeline's output) - add a file there and it is picked up automatically,
'      no edit needed here.
'   3. Reads aeRibbon/VERSION and stamps both RIBBON_VERSION (in
'      basBibleRibbonSetup's CodeModule) and the aeRibbonVersion custom
'      document property - closes the manual copy/paste version-stamp step in
'      BUILD.md's G6.
'
' Deliberately NOT handled (same as the manual process it replaces):
'   - ThisDocument.cls: excluded per md/aeProductionRibbonPlan.md Sec 2.2/7.2.
'     If a Document_Open body is ever added to the template's ThisDocument,
'     this importer will not touch it - keep pasting that by hand per
'     BUILD.md, or extend this module to handle it the way the dev docm's
'     ImportThisDocumentFile does (CodeModule.AddFromString - VBComponents.Import
'     cannot target the built-in ThisDocument component).
'
' Path note: REPO_ROOT is a hardcoded absolute path, not derived from
' ThisDocument.Path. aeRibbon.dotm is legitimately opened from two different
' locations in normal use (the canonical aeRibbon/template/ copy, and a
' deployed copy in %APPDATA%\Microsoft\Word\STARTUP\ per BUILD.md's own
' two-copy warning) - a Path-relative source lookup would silently resolve to
' the wrong folder (or none) when run from the STARTUP copy. Every other
' build doc in this repo already hardcodes this same absolute path; this
' module follows that existing convention rather than introducing a new one.
' =============================================================================

Public Const REPO_ROOT As String = "C:\adaept\aeBibleClass\"

Public Sub ImportAllRibbonVBAFiles(Optional ByVal varDebug As Variant)
    On Error GoTo 0
    Dim strSrcPath As String
    Dim strFile As String
    Dim intImported As Integer
    Dim intSkipped As Integer
    Dim strExt As Variant
    Dim vbCompName As String

    strSrcPath = REPO_ROOT & "aeRibbon\src\"

    ' Verify src folder exists
    If Dir(strSrcPath, vbDirectory) = "" Then
        MsgBox "Source folder not found:" & vbCrLf & strSrcPath, vbCritical, "Import Aborted"
        Exit Sub
    End If

    ' Delete all modules except this one - prompts for confirmation
    If Not DeleteAllRibbonModulesExceptImporter() Then
        Debug.Print "Import aborted - deletion cancelled.", "in Sub ImportAllRibbonVBAFiles"
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
        vbCompName = Left(strFile, InStrRev(strFile, ".") - 1)
        If Not RibbonModuleOrClassExists(vbCompName) Then
            ImportRibbonVBAFile CStr(strFullPath)
            intImported = intImported + 1
        Else
            colSkipped.Add vbCompName & " (already exists)"
            intSkipped = intSkipped + 1
        End If
    Next strFullPath

    Debug.Print "Import complete.", _
                "Imported: " & intImported, _
                "Skipped: " & intSkipped, _
                "in Sub ImportAllRibbonVBAFiles"

    Dim strSkippedItem As Variant
    Dim strSkippedList As String
    strSkippedList = ""
    For Each strSkippedItem In colSkipped
        Debug.Print "  skipped:", CStr(strSkippedItem), "in Sub ImportAllRibbonVBAFiles"
        strSkippedList = strSkippedList & vbCrLf & "  " & strSkippedItem
    Next strSkippedItem

    ' Auto-stamp RIBBON_VERSION + aeRibbonVersion from aeRibbon/VERSION.
    ' Per BUILD.md's own rule: this patches the live template's in-memory
    ' copy only - never write the resolved value back into src/.
    Dim verVal As String
    verVal = ReadRibbonVersionFile()
    If verVal <> "" Then StampRibbonVersion verVal

    Dim strMsgSkipped As String
    strMsgSkipped = ""
    If intSkipped > 0 Then
        strMsgSkipped = vbCrLf & vbCrLf & "Skipped files:" & strSkippedList
    End If

    MsgBox "Import complete." & vbCrLf & vbCrLf & _
           "Imported: " & intImported & vbCrLf & _
           "Skipped:  " & intSkipped & strMsgSkipped & vbCrLf & vbCrLf & _
           "RIBBON_VERSION stamped: " & verVal & vbCrLf & vbCrLf & _
           "Next: Debug -> Compile VBAProject, verify the About dialog, then Save.", _
           vbInformation, "Import Complete"
End Sub

Private Sub ImportRibbonVBAFile(myCodeFile As String)
    On Error GoTo PROC_ERR
    Dim vbaModule As Object
    Dim filePath, fileName, fullPath, vbCompName As String

    fullPath = myCodeFile
    fileName = Mid$(fullPath, InStrRev(fullPath, "\") + 1)
    vbCompName = Left(fileName, InStrRev(fileName, ".") - 1)

    If Dir(fullPath) <> "" Then
        If Not RibbonModuleOrClassExists(vbCompName) Then
            Set vbaModule = ThisDocument.VBProject.VBComponents.Import(fullPath)
            Debug.Print vbCompName, "import SUCCESS!", "in Sub ImportRibbonVBAFile"
        Else
            Debug.Print vbCompName, "import ABORTED!", "in Sub ImportRibbonVBAFile"
        End If
    End If

PROC_EXIT:
    Exit Sub
PROC_ERR:
    If Err = 6068 Then
        MsgBox "VBA Project Not Trusted" & vbCrLf & "Enable 'Trust access to the VBA project object model' in Word Trust Center.", vbCritical, "ImportRibbonVBAFile"
        Stop
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in Sub ImportRibbonVBAFile", vbCritical, "ImportRibbonVBAFile"
        Resume PROC_EXIT
    End If
End Sub

Public Function DeleteAllRibbonModulesExceptImporter() As Boolean
    On Error GoTo PROC_ERR
    Dim vbComp As Object
    Dim strProtected As String
    Dim strToDelete As String
    Dim strMsg As String
    Dim intResponse As Integer

    strProtected = "basImportWordRibbonGitFiles"

    strToDelete = ""
    For Each vbComp In ThisDocument.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1, 2, 3
                If vbComp.Name <> strProtected Then
                    strToDelete = strToDelete & "  " & vbComp.Name & vbCrLf
                End If
        End Select
    Next vbComp

    If strToDelete = "" Then
        MsgBox "No modules found to delete.", vbInformation, "Delete Modules"
        DeleteAllRibbonModulesExceptImporter = True
        Exit Function
    End If

    strMsg = "The following modules and classes will be deleted from aeRibbon.dotm:" & vbCrLf & vbCrLf & _
             strToDelete & vbCrLf & _
             "'" & strProtected & "' will be preserved." & vbCrLf & vbCrLf & _
             "Proceed with deletion?"
    intResponse = MsgBox(strMsg, vbYesNo + vbExclamation, "Confirm Delete")

    If intResponse <> vbYes Then
        Debug.Print "Deletion cancelled by user.", "in Function DeleteAllRibbonModulesExceptImporter"
        DeleteAllRibbonModulesExceptImporter = False
        Exit Function
    End If

    Dim colToDelete As Collection
    Set colToDelete = New Collection
    For Each vbComp In ThisDocument.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1, 2, 3
                If vbComp.Name <> strProtected Then
                    colToDelete.Add vbComp.Name
                End If
        End Select
    Next vbComp

    Dim strName As Variant
    For Each strName In colToDelete
        Set vbComp = ThisDocument.VBProject.VBComponents(strName)
        ThisDocument.VBProject.VBComponents.Remove vbComp
        Debug.Print strName, "deleted", "in Function DeleteAllRibbonModulesExceptImporter"
    Next strName

    DeleteAllRibbonModulesExceptImporter = True

PROC_EXIT:
    Exit Function
PROC_ERR:
    If Err = 6068 Then
        MsgBox "VBA Project Not Trusted" & vbCrLf & "Enable 'Trust access to the VBA project object model' in Word Trust Center.", vbCritical, "DeleteAllRibbonModulesExceptImporter"
        Stop
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in Function DeleteAllRibbonModulesExceptImporter", vbCritical, "DeleteAllRibbonModulesExceptImporter"
        Resume PROC_EXIT
    End If
End Function

Private Function RibbonModuleOrClassExists(name As String) As Boolean
    On Error GoTo 0
    Dim vbComp As Object
    Dim found As Boolean

    found = False
    For Each vbComp In ThisDocument.VBProject.VBComponents
        If vbComp.Name = name Then
            found = True
            Exit For
        End If
    Next vbComp

    RibbonModuleOrClassExists = found
End Function

Private Function ReadRibbonVersionFile() As String
    On Error GoTo PROC_ERR
    Dim fso As Object
    Dim ts As Object
    Dim verPath As String
    Dim s As String

    verPath = REPO_ROOT & "aeRibbon\VERSION"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(verPath) Then
        MsgBox "VERSION file not found:" & vbCrLf & verPath & vbCrLf & _
               "RIBBON_VERSION stamp skipped.", vbExclamation, "Stamp Version"
        ReadRibbonVersionFile = ""
        Exit Function
    End If

    Set ts = fso.OpenTextFile(verPath, 1) ' 1 = ForReading
    s = Trim$(ts.ReadLine)
    ts.Close
    ReadRibbonVersionFile = s
    Exit Function

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in Function ReadRibbonVersionFile", vbCritical, "ReadRibbonVersionFile"
    ReadRibbonVersionFile = ""
End Function

Private Sub StampRibbonVersion(ByVal verVal As String)
    On Error GoTo PROC_ERR
    Dim cm As Object
    Dim i As Long
    Dim found As Boolean

    If Not RibbonModuleOrClassExists("basBibleRibbonSetup") Then
        MsgBox "basBibleRibbonSetup not found in the project - RIBBON_VERSION stamp skipped.", _
               vbExclamation, "Stamp Version"
        Exit Sub
    End If

    Set cm = ThisDocument.VBProject.VBComponents("basBibleRibbonSetup").CodeModule
    found = False
    For i = 1 To cm.CountOfLines
        If InStr(1, cm.Lines(i, 1), "Public Const RIBBON_VERSION", vbTextCompare) > 0 Then
            cm.ReplaceLine i, "Public Const RIBBON_VERSION As String = """ & verVal & """"
            found = True
            Exit For
        End If
    Next i

    If Not found Then
        MsgBox "RIBBON_VERSION constant not found in basBibleRibbonSetup - stamp skipped.", _
               vbExclamation, "Stamp Version"
        Exit Sub
    End If

    ' Custom document property, mirrors BUILD.md G6 step 2. Late-bound: 4 is
    ' msoPropertyTypeString - no Office library reference added, per this
    ' project's late-binding-only convention.
    On Error Resume Next
    Err.Clear
    ThisDocument.CustomDocumentProperties.Add Name:="aeRibbonVersion", LinkToContent:=False, Type:=4, Value:=verVal
    If Err.Number <> 0 Then
        Err.Clear
        ThisDocument.CustomDocumentProperties("aeRibbonVersion").Value = verVal
    End If
    On Error GoTo PROC_ERR

    Debug.Print "RIBBON_VERSION stamped: " & verVal, "in Sub StampRibbonVersion"
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in Sub StampRibbonVersion", vbCritical, "StampRibbonVersion"
End Sub
