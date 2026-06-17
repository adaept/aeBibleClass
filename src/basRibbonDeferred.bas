Attribute VB_Name = "basRibbonDeferred"
Option Explicit

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
' basRibbonDeferred - Deferred Ribbon Action Dispatch
' -----------------------------------------------------------------------------
' Public subs scheduled via Application.OnTime from ribbon callbacks in
' basBibleRibbonSetup.
'
' Why a separate module:
'   All standard modules in this project use Option Private Module. That
'   declaration prevents Application.OnTime from resolving macros by name -
'   it is NOT the same as Alt+F8 visibility (which is separately controlled
'   by whether a sub has required parameters). This module intentionally omits
'   Option Private Module so that Application.OnTime can find its public subs.
' =============================================================================

' -- Active deferred entry points ----------------------------------------------

Public Sub GoToChapterDeferred()
    Instance().ExecutePendingChapter
End Sub

Public Sub UpdateStatusBarDeferred()
    Instance().UpdateStatusBar
End Sub

Public Sub ResetChapterDisplayDeferred()
    Instance().ResetChapterDisplay
End Sub

Public Sub ResetVerseDisplayDeferred()
    Instance().ResetVerseDisplay
End Sub


' -- Archived deferred entry points retained for rollback/testing --------------

' WarmLayoutCacheDeferred: WarmLayoutCache is disabled — the OnTime call that
' would schedule this sub is commented out in aeRibbonClass.cls. The cache
' method itself is preserved there for future use.
Public Sub WarmLayoutCacheDeferred()
    ' Dim rc As aeRibbonClass
    ' Set rc = Instance()
    ' rc.WarmLayoutCache
End Sub

' GoToH1Deferred: legacy entry point from the old GoTo Book button flow.
' That button was removed from ribbon XML; Book selection now uses the Book comboBox.
' Note: this sub has no parameters and will appear in Alt+F8 — safe to run manually.
Public Sub GoToH1Deferred()
    Dim rc As aeRibbonClass
    Set rc = Instance()
    rc.GoToH1Direct
    rc.InvalidateControl CTRL_BOOK
    rc.InvalidateControl CTRL_NEXT_BOOK
    rc.InvalidateControl CTRL_PREV_BOOK
End Sub

' GoToBookDeferred: dead stub — NavigateToCurrentBook removed (Bug 9).
Public Sub GoToBookDeferred()
    ' Instance().NavigateToCurrentBook
End Sub

' GoToVerseDeferred: dead stub — navigation trigger moved to OnGoClick (GoButton, #600).
Public Sub GoToVerseDeferred()
    ' Instance().ExecutePendingVerse
End Sub
