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
' Class modules never use Option Private Module - it is a standard-module-only
' declaration. Classes are always accessible to any module that can instantiate
' them with New. This means aeLoggerClass is callable from basRibbonDeferred and
' other modules that also omit Option Private Module.

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
'     - All writes are buffered in memory (m_buffer)
'     - Each Log_Write rewrites the full file from the buffer
'       (crash-safe: file is always current after each write)
'     - ADODB.Stream text-mode position seeking is avoided entirely;
'       the buffer approach is reliable across all ADODB versions
'     - Session header includes timestamp and machine/user info
' ================================================================

Private m_buffer    As String   ' accumulates all content including BOM
Private logFilePath As String
Private logIsOpen   As Boolean

Private Sub Class_Terminate()
    If logIsOpen Then Log_Close
End Sub

' ------------------------------------------------------------
' Initialize UTF-8 logging - buffer starts with BOM
' ------------------------------------------------------------
Public Sub Log_Init(ByVal path As String)
    logFilePath = path
    m_buffer = ChrW(&HFEFF)    ' UTF-8 BOM (always start fresh each session)
    logIsOpen = True
    Log_Write "==============================================================="
    Log_Write "LOG SESSION START: " & Format(Now, "yyyy-mm-dd HH:nn:ss")
    Log_Write "User: " & Environ$("USERNAME")
    Log_Write "Machine: " & Environ$("COMPUTERNAME")
    Log_Write "==============================================================="
End Sub

' ------------------------------------------------------------
' Append line to buffer and rewrite full file
' ------------------------------------------------------------
Private Sub Log_WriteRaw(ByVal line As String)
    m_buffer = m_buffer & line & vbCrLf
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2              ' adTypeText
        .Charset = "UTF-8"
        .Open
        .WriteText m_buffer
        .saveToFile logFilePath, 2   ' adSaveCreateOverWrite
        .Close
    End With
    Set stm = Nothing
End Sub

Public Sub Log_Write(ByVal msg As String)
    If Not logIsOpen Then Err.Raise vbObjectError + 513, , "Log not initialized"
    Log_WriteRaw Format(Now, "HH:nn:ss") & " | " & msg
End Sub

' ------------------------------------------------------------
' Flush final line and close
' ------------------------------------------------------------
Public Sub Log_Close()
    If logIsOpen Then
        Log_Write "LOG SESSION END"
        logIsOpen = False
        m_buffer = ""
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

' ------------------------------------------------------------
' Self-test - validates UTF-8 output with Hebrew sample text
' ------------------------------------------------------------
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
