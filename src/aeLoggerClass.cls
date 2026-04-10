VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aeLoggerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

' Note: Option Private Module is intentionally absent.
' Class modules never use Option Private Module — it is a standard-module-only
' declaration. Classes are always accessible to any module that can instantiate
' them with New. This means aeLoggerClass is callable from basRibbonDeferred and
' other modules that also omit Option Private Module.

Private logFilePath As String
Private logIsOpen As Boolean

Public Sub TestLogging()
    Debug.Print "Log_Init called at "; Now
    Dim s As String
    s = ChrW(&H5D0) & ChrW(&H5B1) & ChrW(&H5DC) & ChrW(&H5D4) & ChrW(&H5D9) & ChrW(&H5DD)
    Log_Init "C:\Temp\audit.txt"
    Log_UnicodeDetail "<<<Log_Init called at " & Now & ">>>"
    Log_UnicodeDetail s
    s = ChrW(&H5D0) & ChrW(&H5B1) & ChrW(&H5DC) & ChrW(&H5B9) & ChrW(&H5D4) & ChrW(&H5B4) & ChrW(&H5DD)
    Log_UnicodeDetail s
    Log_Close
End Sub

' ================================================================
'  LOGGING CLASS - AUDIT-READY, UNICODE-SAFE, SESSION-AWARE
'  Author: Peter
'  Purpose:
'     - Produce deterministic, timestamped logs
'     - Avoid Immediate Window truncation
'     - Unicode-safe (UTF-8 with BOM)
'     - Minimal hidden state; explicit initialization
'     - Suitable for publishing pipelines and reproducible debugging
'
'  Usage:
'     Dim log As New aeLoggerClass
'     log.Log_Init "C:\path\to\log.txt"
'     log.Log_Write "Message"
'     log.Log_Close
'
'  Notes:
'     - File is opened once and kept open for performance
'     - All writes are flushed immediately for crash-safe logging
'     - Session header includes timestamp and machine/user info
' ================================================================

' ------------------------------------------------------------
' Initialize UTF-8 logging (append mode, no ANSI fallback)
' ------------------------------------------------------------
Public Sub Log_Init(ByVal path As String)
    logFilePath = path
    logIsOpen = True
    Log_Write "==============================================================="
    Log_Write "LOG SESSION START: " & Format(Now, "yyyy-mm-dd HH:nn:ss")
    Log_Write "User: " & Environ$("USERNAME")
    Log_Write "Machine: " & Environ$("COMPUTERNAME")
    Log_Write "==============================================================="
End Sub

Private Sub Log_WriteRaw(ByVal line As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2              ' adTypeText
        .Charset = "UTF-8"
        .Open
        ' If file exists, load and move to end; else write BOM
        If Dir$(logFilePath) <> "" Then
            .LoadFromFile logFilePath
            .Position = .Size
        Else
            .WriteText ChrW(&HFEFF), 0   ' BOM
        End If
        .WriteText line & vbCrLf
        .saveToFile logFilePath, 2       ' adSaveCreateOverWrite
        .Close
    End With
End Sub

Public Sub Log_Write(ByVal msg As String)
    If Not logIsOpen Then Err.Raise vbObjectError + 513, , "Log not initialized"
    Log_WriteRaw Format(Now, "HH:nn:ss") & " | " & msg
End Sub

Public Sub Log_Close()
    If logIsOpen Then
        Log_Write "LOG SESSION END"
        logIsOpen = False
    End If
End Sub

Public Sub Log_UnicodeDetail(ByVal s As String)
    Dim i As Long, cp As Long, ch As String
    Log_Write "Unicode detail for string: [" & s & "]"
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        cp = AscW(ch)
        Log_Write "  pos " & i & ": U+" & Right$("0000" & Hex$(cp), 4) & " (" & cp & ")"
    Next i
End Sub
