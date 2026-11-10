Attribute VB_Name = "basDocmHygiene"
Option Explicit
Option Compare Text

' Classification: docm clean-up (list-then-fix pairs). Not used by runtime ribbon navigation.
'
' basDocmHygiene - clean the dev docm before the RWB 1.1.1 export (tasks RT-13 to RT-17, owner decisions S-2, 2026-11-09)
'
' Pattern: every fix has a read-only LIST twin. Run the list, read it, run the fix on a COPY of the docm, run the list again
' (expected: nothing listed), and only then run the fix on the real docm. A fix never changes the text: it compares the
' length of the document text before and after and reports a mismatch.
'
' This file so far (RT-13 only; RT-14 to RT-17 are added to this module as they are written):
'   ListDirectFontRuns "<Doc name>"   runs whose font resolves outside Carlito / Noto Sans / Liberation Serif (reuses the
'                                     audit of basFontRunAudit); read-only
'   FixDirectFontRuns  "<Doc name>"   clears the direct font setting (character Reset) on exactly those runs, EXCEPT runs in
'                                     the paragraph style BodyText: those 3 non-Latin runs (Cambria Math, Arial) are looked
'                                     at by the owner first. Pass bIncludeBodyText:=True only after that look.
'
' Entry points (Immediate window, one statement each; the document name is optional, default ActiveDocument):
'   ListDirectFontRuns
'   FixDirectFontRuns
'   ? CountDirectFontRunsOutsideAllowlist
' Reference: basFontRunAudit (the allowlist and the report), rvw\Plan_rwb_1_1_1_release_train_2026-11-08.md rows RT-13 to RT-17.

Private Const BODY_STYLE As String = "BodyText"
Private Const FIND_SAFETY_LIMIT As Long = 100000
Private Const BODY_TOKEN As String = "+Body"
Private Const HEADINGS_TOKEN As String = "+Headings"
Private m_LastError As String

' The text of the last error ("" = none), for callers that cannot see the Immediate window.
Public Function DocmHygieneLastError() As String
    DocmHygieneLastError = m_LastError
End Function

'==============================================================================
' RT-13 - direct font runs
'==============================================================================
Public Sub ListDirectFontRuns(Optional ByVal sDocName As String = "")
    Dim hitCount As Long
    Dim sReport As String

    FindFontRunsOutsideAllowlist hitCount, sReport, ResolveDocument(sDocName)
    Debug.Print sReport
    Debug.Print "ListDirectFontRuns: " & hitCount & " run(s) outside the allowlist (the " & BODY_STYLE & " ones are skipped by FixDirectFontRuns)."
End Sub

Public Function CountDirectFontRunsOutsideAllowlist(Optional ByVal sDocName As String = "") As Long
    CountDirectFontRunsOutsideAllowlist = CountFontRunsOutsideAllowlist(sDocName)
End Function

' Returns the number of runs fixed (-1 = failed, see DocmHygieneLastError). Text is never changed.
Public Function FixDirectFontRuns(Optional ByVal sDocName As String = "", Optional ByVal bIncludeBodyText As Boolean = False) As Long
    On Error GoTo PROC_ERR

    Dim oDoc As Word.Document
    Dim dictFonts As Object
    Dim vFont As Variant
    Dim nFixed As Long
    Dim nSkipped As Long
    Dim nLenBefore As Long
    Dim bScreen As Boolean

    m_LastError = ""
    Set oDoc = ResolveDocument(sDocName)
    bScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    nLenBefore = Len(oDoc.Content.Text)

    Set dictFonts = CreateObject("Scripting.Dictionary")
    dictFonts.CompareMode = vbTextCompare
    CollectOutsideFonts oDoc, dictFonts

    For Each vFont In dictFonts.Keys
        Application.StatusBar = "FixDirectFontRuns: " & CStr(vFont) & "..."
        ResetRunsInFont oDoc, CStr(vFont), bIncludeBodyText, nFixed, nSkipped
    Next vFont

    If Len(oDoc.Content.Text) <> nLenBefore Then
        m_LastError = "FixDirectFontRuns: the document text length changed (" & nLenBefore & " -> " & Len(oDoc.Content.Text) & "); undo or close without saving."
        Debug.Print "ERROR in basDocmHygiene." & m_LastError
        nFixed = -1
    Else
        Debug.Print "FixDirectFontRuns: " & nFixed & " run(s) reset, " & nSkipped & " skipped (" & BODY_STYLE & "). Text length unchanged (" & nLenBefore & ")."
    End If
    FixDirectFontRuns = nFixed

PROC_EXIT:
    Application.StatusBar = False
    Application.ScreenUpdating = bScreen
    Exit Function
PROC_ERR:
    m_LastError = "FixDirectFontRuns: " & Err.Number & " " & Err.Description
    Debug.Print "ERROR in basDocmHygiene." & m_LastError
    FixDirectFontRuns = -1
    Resume PROC_EXIT
End Function

' The fonts to look for: the two theme tokens and every font a style names that is outside the allowlist, plus the fonts the
' audit report names (read from the same places the audit reads: styles and the usual Microsoft fonts as a fallback list).
Private Sub CollectOutsideFonts(ByVal oDoc As Word.Document, ByVal dict As Object)
    Dim oStyle As Word.Style
    Dim v As Variant
    Dim sName As String

    dict.Add BODY_TOKEN, True
    dict.Add HEADINGS_TOKEN, True
    For Each oStyle In oDoc.Styles
        sName = ""
        On Error Resume Next
        sName = oStyle.Font.Name
        On Error GoTo 0
        AddOutsideFont dict, sName
    Next oStyle
    For Each v In Array("Aptos", "Aptos Display", "Calibri", "Calibri Light", "Cambria", "Cambria Math", "Times New Roman", "Arial", _
                        "Arial Unicode MS", "MS Mincho", "Segoe UI", "Courier New", "Symbol", "Wingdings")
        AddOutsideFont dict, CStr(v)
    Next v
End Sub

Private Sub AddOutsideFont(ByVal dict As Object, ByVal sName As String)
    sName = Trim$(sName)
    If Len(sName) = 0 Then Exit Sub
    If IsAllowedRunFont(sName) Then Exit Sub
    If Not dict.Exists(sName) Then dict.Add sName, True
End Sub

' Character Reset on every run found in sFont whose RESOLVED font is outside the allowlist (the same test as the audit).
Private Sub ResetRunsInFont(ByVal oDoc As Word.Document, ByVal sFont As String, ByVal bIncludeBodyText As Boolean, _
                            ByRef nFixed As Long, ByRef nSkipped As Long)
    Dim oStory As Word.Range
    Dim oRng As Word.Range
    Dim r As Word.Range
    Dim lastStart As Long
    Dim guard As Long
    Dim sResolved As String
    Dim sParaStyle As String

    For Each oStory In oDoc.StoryRanges
        Set oRng = oStory
        Do While Not oRng Is Nothing
            Set r = oRng.Duplicate
            lastStart = -1
            guard = 0
            With r.Find
                .ClearFormatting
                .Text = ""
                .Font.Name = sFont
                .Format = True
                .Forward = True
                .Wrap = wdFindStop
                .MatchWildcards = False
                Do While .Execute
                    guard = guard + 1
                    If guard > FIND_SAFETY_LIMIT Then Exit Do
                    If r.Start <= lastStart Then Exit Do
                    lastStart = r.Start
                    sResolved = ""
                    sParaStyle = ""
                    On Error Resume Next
                    sResolved = r.Font.Name
                    sParaStyle = r.Paragraphs(1).style.NameLocal
                    On Error GoTo 0
                    If Not IsAllowedRunFont(sResolved) Then
                        If StrComp(sParaStyle, BODY_STYLE, vbTextCompare) = 0 And Not bIncludeBodyText Then
                            nSkipped = nSkipped + 1
                        Else
                            ResetKeepingCharacterStyles r, oDoc
                            nFixed = nFixed + 1
                        End If
                    End If
                    r.Collapse wdCollapseEnd
                Loop
            End With
            Set oRng = oRng.NextStoryRange
        Loop
    Next oStory
End Sub

' Font.Reset removes the character style as well as the direct formatting (measured 2026-11-09: it stripped "Words of Jesus" and its
' colour from 4 verse runs). So the character style of every character is noted first and put back after the reset.
Private Sub ResetKeepingCharacterStyles(ByVal r As Word.Range, ByVal oDoc As Word.Document)
    Dim nChars As Long
    Dim k As Long
    Dim sDefault As String
    Dim saved() As String

    sDefault = oDoc.Styles(wdStyleDefaultParagraphFont).NameLocal
    nChars = r.End - r.Start
    If nChars < 1 Then Exit Sub
    ReDim saved(1 To nChars)
    For k = 1 To nChars
        saved(k) = sDefault
        On Error Resume Next
        saved(k) = oDoc.Range(r.Start + k - 1, r.Start + k).CharacterStyle.NameLocal
        On Error GoTo 0
    Next k

    r.Font.Reset

    For k = 1 To nChars
        If StrComp(saved(k), sDefault, vbTextCompare) <> 0 Then
            On Error Resume Next
            oDoc.Range(r.Start + k - 1, r.Start + k).style = saved(k)
            On Error GoTo 0
        End If
    Next k
End Sub

'==============================================================================
' Small helpers
'==============================================================================
Private Function ResolveDocument(ByVal sDocName As String) As Word.Document
    If Len(sDocName) = 0 Then
        Set ResolveDocument = ActiveDocument
    Else
        Set ResolveDocument = Documents(sDocName)
    End If
End Function
