Attribute VB_Name = "basRibbonDeferred"
Option Explicit

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

Public Sub FocusBookDeferred()
    ' Bug #597 — focus cmbBook after New Search resets state.
    ' Fires via Application.OnTime after onAction returns focus to the document.
    ' Sends the ribbon keytip sequence: Alt+Y2 (RWB tab) then B (Book comboBox).
    SendKeys "%Y2B"
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
    rc.InvalidateControl "NextBookButton"
    rc.InvalidateControl "PrevBookButton"
End Sub

' GoToBookDeferred: dead stub — NavigateToCurrentBook removed (Bug 9).
Public Sub GoToBookDeferred()
    ' Instance().NavigateToCurrentBook
End Sub

' GoToVerseDeferred: dead stub — navigation trigger moved to OnGoClick (GoButton, #600).
Public Sub GoToVerseDeferred()
    ' Instance().ExecutePendingVerse
End Sub
