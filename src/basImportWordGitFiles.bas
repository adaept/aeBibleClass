Attribute VB_Name = "basImportWordGitFiles"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

' ImportAllVBAFiles
' ------------------
' Full wipe + reimport of every module from src\ (see
' DeleteAllModulesExceptImporter/RunOneImportPass below). Skipped should
' ALWAYS come back as exactly 1 - the importer module itself
' (basImportWordGitFiles), which DeleteAllModulesExceptImporter
' deliberately leaves alive and RunOneImportPass's own ModuleOrClassExists
' check then correctly reports as "already exists". Any OTHER value has
' been observed twice (see feedback_importallvbafiles_error17 memory) and
' traced to a partial-deletion race, not a legitimate skip - some modules
' silently missing from the live project despite a reported "success".
' Retries the whole pass (including the deletion confirmation, on purpose
' - it's a destructive step and shouldn't be silently re-run) up to
' MAX_IMPORT_ATTEMPTS times when that happens, since simply re-running is
' the known, already-proven fix.
Public Sub ImportAllVBAFiles(Optional ByVal varDebug As Variant)
    On Error GoTo 0
    Const MAX_IMPORT_ATTEMPTS As Long = 3
    Dim attemptNum As Long
    Dim intImported As Integer
    Dim intSkipped As Integer
    Dim colSkipped As Collection
    Dim blnCompleted As Boolean

    For attemptNum = 1 To MAX_IMPORT_ATTEMPTS
        If Not RunOneImportPass(intImported, intSkipped, colSkipped) Then
            ' Aborted (missing src folder or deletion cancelled) - already reported.
            Exit Sub
        End If

        If intSkipped = 1 Then
            blnCompleted = True
            Exit For
        End If

        Debug.Print "ImportAllVBAFiles", _
                    "Skipped=" & intSkipped & " (expected 1) on attempt " & attemptNum & " of " & MAX_IMPORT_ATTEMPTS, _
                    "in Sub ImportAllVBAFiles"

        If attemptNum < MAX_IMPORT_ATTEMPTS Then
            If MsgBox("Import anomaly: Skipped " & intSkipped & " modules, expected exactly 1 (the importer itself)." & vbCrLf & vbCrLf & _
                      "This is the known partial-deletion race (feedback_importallvbafiles_error17) - re-running the import has always fixed it." & vbCrLf & vbCrLf & _
                      "Retry now? (attempt " & (attemptNum + 1) & " of " & MAX_IMPORT_ATTEMPTS & ")", _
                      vbExclamation + vbYesNo, "Import Anomaly - Skipped <> 1") = vbNo Then
                Exit For
            End If
        End If
    Next attemptNum

    ' A For loop that completes without Exit For leaves its counter one past
    ' the limit (attemptNum = MAX_IMPORT_ATTEMPTS + 1 here) - clamp so the
    ' report never claims an attempt that didn't happen.
    If attemptNum > MAX_IMPORT_ATTEMPTS Then attemptNum = MAX_IMPORT_ATTEMPTS

    ReportImportResult intImported, intSkipped, colSkipped, attemptNum, blnCompleted
End Sub

' RunOneImportPass
' -----------------
' One full delete-then-reimport pass. Returns False (with its own
' MsgBox already shown) if the src folder is missing or the deletion
' confirmation was cancelled - either way, ImportAllVBAFiles should stop,
' not retry. Returns True otherwise, regardless of whether intSkipped
' came back as the expected 1 - that check is the caller's job.
Private Function RunOneImportPass(ByRef intImported As Integer, ByRef intSkipped As Integer, ByRef colSkipped As Collection) As Boolean
    Dim strSrcPath As String
    Dim strFile As String
    Dim strExt As Variant
    Dim vbCompName As String

    RunOneImportPass = False
    strSrcPath = ThisDocument.Path & "\src\"

    ' Verify src folder exists
    If Dir(strSrcPath, vbDirectory) = "" Then
        MsgBox "Source folder not found:" & vbCrLf & strSrcPath, vbCritical, "Import Aborted"
        Exit Function
    End If

    ' Delete all modules except this one - prompts for confirmation
    If Not DeleteAllModulesExceptImporter() Then
        Debug.Print "Import aborted - deletion cancelled.", "in Sub ImportAllVBAFiles"
        Exit Function
    End If

    ' Diagnostic for the Skipped<>1 investigation (feedback_importallvbafiles_error17):
    ' dump exactly what the live project thinks it still has, right after deletion
    ' claims success and before anything is reimported. Tells us whether a repeat
    ' offender (basBiblePalette/basTEST_aeBibleTools/Module1) is genuinely still
    ' present (Remove silently failed/skipped it) or genuinely gone (the later
    ' ModuleOrClassExists check is reading stale state). Remove once root-caused.
    DumpLiveVBComponents "post-delete"

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
            ' ThisDocument is a built-in document module - it cannot be removed
            ' and re-imported like a normal component (VBComponents.Import errors
            ' on an existing name). Its CodeModule is replaced in place instead,
            ' so src\ThisDocument.cls (including its license header) actually
            ' reaches the live project rather than being silently skipped.
            ImportThisDocumentFile CStr(strFullPath)
            Debug.Print "ThisDocument", "code replaced in place", "in Sub ImportAllVBAFiles"
            intImported = intImported + 1
        End If
    Next strFullPath

    Debug.Print "Import pass complete.", _
                "Imported: " & intImported, _
                "Skipped: " & intSkipped, _
                "in Sub ImportAllVBAFiles"

    RunOneImportPass = True
End Function

' ReportImportResult
' -------------------
' Final summary dialog. attemptsUsed/blnCompleted let the message say
' plainly whether Skipped=1 was reached (possibly after retries) or
' whether MAX_IMPORT_ATTEMPTS was exhausted / the operator declined a
' retry - in the latter case this is the same "go fix it by hand"
' situation as before this change, just reached faster and more visibly.
Private Sub ReportImportResult(ByVal intImported As Integer, ByVal intSkipped As Integer, ByRef colSkipped As Collection, ByVal attemptsUsed As Long, ByVal blnCompleted As Boolean)
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

    Dim strHeadline As String
    If blnCompleted Then
        If attemptsUsed > 1 Then
            strHeadline = "Import complete (succeeded on attempt " & attemptsUsed & ")."
        Else
            strHeadline = "Import complete."
        End If
        MsgBox strHeadline & vbCrLf & vbCrLf & _
               "Imported: " & intImported & vbCrLf & _
               "Skipped:  " & intSkipped & strMsgSkipped, _
               vbInformation, "Import Complete"
    Else
        MsgBox "Import STILL anomalous after " & attemptsUsed & " attempt(s) - Skipped=" & intSkipped & ", expected 1." & vbCrLf & vbCrLf & _
               "Imported: " & intImported & vbCrLf & _
               "Skipped:  " & intSkipped & strMsgSkipped & vbCrLf & vbCrLf & _
               "Manual recovery needed - see feedback_importallvbafiles_error17 memory.", _
               vbExclamation, "Import Incomplete"
    End If
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

Private Sub ImportThisDocumentFile(ByVal myCodeFile As String)
    On Error GoTo PROC_ERR

    Dim fso As Object
    Dim ts As Object
    Dim strLine As String
    Dim strBody As String
    Dim blnBodyStarted As Boolean
    Dim strTrimmed As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(myCodeFile) Then Exit Sub

    ' The exported .cls text starts with a VBE-managed VERSION/BEGIN/END block
    ' and Attribute lines - those are component metadata, not CodeModule text,
    ' and cannot be written back through CodeModule.AddFromString. Skip past
    ' them and keep everything from Option Explicit onward (license header
    ' included) as the module body.
    blnBodyStarted = False
    Set ts = fso.OpenTextFile(myCodeFile, 1)  ' ForReading
    Do While Not ts.AtEndOfStream
        strLine = ts.ReadLine
        strTrimmed = LTrim$(strLine)
        If Not blnBodyStarted Then
            If LCase$(Left$(strTrimmed, 9)) <> "attribute" _
               And strLine <> "VERSION 1.0 CLASS" _
               And strLine <> "BEGIN" _
               And strLine <> "END" _
               And InStr(strTrimmed, "MultiUse") = 0 Then
                blnBodyStarted = True
            End If
        End If
        If blnBodyStarted Then
            ' Attribute lines are VBE-managed metadata wherever they appear,
            ' not just in the leading header block above - Word's export
            ' also emits one mid-file for each WithEvents variable (e.g.
            ' "Attribute oWordApp.VB_VarHelpID = -1", added 2026-09-21 for
            ' ThisDocument's first-ever WithEvents variable). Injecting that
            ' literally via AddFromString is a syntax error (confirmed live:
            ' red/error line in the VBE). Skip any such line here too, not
            ' just before blnBodyStarted flips True.
            If LCase$(Left$(strTrimmed, 9)) <> "attribute" Then
                strBody = strBody & strLine & vbCrLf
            End If
        End If
    Loop
    ts.Close

    ' strBody is built with vbCrLf as a line TERMINATOR (appended after every
    ' line, including the last), so it always ends with a trailing vbCrLf.
    ' CodeModule.AddFromString treats that trailing terminator as introducing
    ' one more (empty) line - left as-is, every import/export round trip
    ' silently grows the module by one blank line. Strip it so vbCrLf acts
    ' as a separator instead.
    If Right$(strBody, 2) = vbCrLf Then strBody = Left$(strBody, Len(strBody) - 2)

    With ThisDocument.VBProject.VBComponents("ThisDocument").CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If Len(strBody) > 0 Then .AddFromString strBody
    End With

    Debug.Print "ThisDocument", "import SUCCESS!", "in Sub ImportThisDocumentFile"

PROC_EXIT:
    Exit Sub
PROC_ERR:
    If Err = 6068 Then
        MsgBox "VBA Project Not Trusted" & vbCrLf & "Enable 'Trust access to the VBA project object model' in Word Trust Center.", vbCritical, "ImportThisDocumentFile"
        Stop
    Else
        MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in Sub ImportThisDocumentFile", vbCritical, "ImportThisDocumentFile"
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
                If vbComp.Name <> strProtected Then
                    strToDelete = strToDelete & "  " & vbComp.Name & vbCrLf
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
                If vbComp.Name <> strProtected Then
                    colToDelete.Add vbComp.Name
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

' DumpLiveVBComponents
' ---------------------
' One-off diagnostic for the Skipped<>1 investigation - prints
' VBComponents.Count plus every remaining component's Name/Type to the
' Immediate window, tagged with a caller-supplied label so multiple calls
' (e.g. "post-delete") are distinguishable in the log. Not wired into any
' RUN_THE_TESTS Case; remove the DumpLiveVBComponents call site(s) once
' the Skipped<>1 root cause is confirmed.
Private Sub DumpLiveVBComponents(ByVal label As String)
    Dim vbComp As Object
    Dim typeName As String

    Debug.Print "DumpLiveVBComponents [" & label & "]", _
                "Count=" & ThisDocument.VBProject.VBComponents.Count, _
                "in Sub DumpLiveVBComponents"

    For Each vbComp In ThisDocument.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1: typeName = "StdModule"
            Case 2: typeName = "ClassModule"
            Case 3: typeName = "MSForm"
            Case 100: typeName = "Document"
            Case Else: typeName = "Type=" & vbComp.Type
        End Select
        Debug.Print "  " & label, vbComp.Name, typeName, "in Sub DumpLiveVBComponents"
    Next vbComp
End Sub

Private Function ModuleOrClassExists(name As String) As Boolean
    On Error GoTo 0
    Dim vbComp As Object
    Dim found As Boolean
    
    found = False
    'Debug.Print "name = " & name, "in Function ModuleOrClassExists"
    For Each vbComp In ThisDocument.VBProject.VBComponents
        If vbComp.Name = name Then
            found = True
            Exit For
        End If
    Next vbComp
    
    ModuleOrClassExists = found
    Debug.Print name, "ModuleOrClassExists = " & found, "in Function ModuleOrClassExists"
End Function
