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
'
'   Note: GoToH1Deferred has no parameters so it will appear in Alt+F8.
'   This is acceptable - it is safe to run manually for testing.
' =============================================================================

Public Sub WarmLayoutCacheDeferred()
    ' WarmLayoutCache disabled: entry point in EnableButtonsRoutine (aeRibbonClass.cls)
    ' is commented out so this sub will not be called at document open.
    ' WarmLayoutCache itself is preserved in aeRibbonClass.cls for future use.
    ' Dim rc As aeRibbonClass
    ' Set rc = Instance()
    ' rc.WarmLayoutCache
End Sub

Public Sub GoToH1Deferred()
    Dim rc As aeRibbonClass
    Set rc = Instance()
    rc.GoToH1Direct
    rc.InvalidateControl "NextBookButton"
    rc.InvalidateControl "PrevBookButton"
End Sub

Public Sub GoToBookDeferred()
    ' Dead stub — NavigateToCurrentBook removed (Bug 9).
    ' Instance().NavigateToCurrentBook
End Sub

Public Sub GoToChapterDeferred()
    Instance().ExecutePendingChapter
End Sub

Public Sub GoToVerseDeferred()
    Instance().ExecutePendingVerse
End Sub
