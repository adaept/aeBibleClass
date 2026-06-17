VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aeUpdateCharStyleClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Implements IaeLongProcessClass

'Copyright (c) 2025-2026 Peter F. Ennis
'SPDX-License-Identifier: LGPL-3.0-or-later OR LicenseRef-adaept-Commercial
'DUAL-LICENSED. You may use this file under EITHER of:
'  (1) the GNU Lesser General Public License, version 3.0 or (at your option)
'      any later version  -  https://www.gnu.org/licenses/lgpl-3.0.txt ; OR
'  (2) a commercial / proprietary license available from adaept (Peter Ennis),
'      permitting use in closed-source / proprietary works WITHOUT the LGPL
'      copyleft obligations. Contact the copyright holder for commercial terms.
'As the sole copyright holder, adaept may license this file under either option.
'This library is distributed WITHOUT ANY WARRANTY; without even the implied
'warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.

' =============================================================================
' aeUpdateCharStyleClass - IaeLongProcessClass implementation
' -----------------------------------------------------------------------------
' Re-applies the "Chapter Verse marker" character style to every character
' that already carries it. This forces Word to rebuild its internal style data
' for those characters (style refresh/repair operation).
'
' One paragraph is processed per ExecuteItem call. The runner (aeLongProcessClass)
' manages batching, DoEvents, progress persistence, and logging.
'
' Usage (Immediate Window):
'   Dim t As New aeUpdateCharStyleClass
'   Dim r As New aeLongProcessClass
'   r.Run t
'
' To target a specific style, set StyleName before calling Run:
'   t.StyleName = "My Character Style"
' =============================================================================

Public StyleName       As String  ' legacy parameter - not used by ExecuteItem; kept for future single-style tasks
Public DoEventsEvery   As Long    ' call DoEvents every N characters within a paragraph (default: 50; 0 = disable)

Private Sub Class_Initialize()
    StyleName = "Chapter Verse marker"
    DoEventsEvery = 50
End Sub

' -----------------------------------------------------------------------------
' IaeLongProcessClass interface implementation
' -----------------------------------------------------------------------------

Private Property Get IaeLongProcessClass_ItemCount() As Long
    IaeLongProcessClass_ItemCount = ActiveDocument.Paragraphs.Count
End Property

Private Property Get IaeLongProcessClass_TaskName() As String
    IaeLongProcessClass_TaskName = "UpdateCharacterStyle"
End Property

Private Function IaeLongProcessClass_ExecuteItem(itemIndex As Long) As Boolean
    On Error GoTo PROC_ERR
    Dim para As Word.Paragraph
    Dim rng  As Word.Range
    Set para = ActiveDocument.Paragraphs(itemIndex)
    ' Marker runs are always the first two runs of every verse paragraph:
    '   Run 1 - "Chapter Verse marker" (chapter number digit(s))
    '   Run 2 - "Verse marker"         (verse number + narrow no-break space)
    ' Exit as soon as a non-marker character is found - no need to scan the
    ' entire paragraph. Both styles are refreshed in a single pass.
    For Each rng In para.Range.Characters
        If rng.style = "Chapter Verse marker" Then
            rng.style = "Chapter Verse marker"
        ElseIf rng.style = "Verse marker" Then
            rng.style = "Verse marker"
        Else
            Exit For
        End If
    Next rng
    IaeLongProcessClass_ExecuteItem = True   ' continue
PROC_EXIT:
    Exit Function
PROC_ERR:
    IaeLongProcessClass_ExecuteItem = False  ' stop task on error
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ExecuteItem (item " & itemIndex & ") of Class aeUpdateCharStyleClass"
    Resume PROC_EXIT
End Function
