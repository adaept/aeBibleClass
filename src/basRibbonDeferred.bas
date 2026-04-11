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
    ' InvalidateControl is called after WarmLayoutCacheDeferred has run at document
    ' open (Option B, §30). With the layout cache warm, these calls are cheap.
    ' If called before the warm-up completes (within first 5 seconds of open),
    ' the cache may be cold and a brief block may occur on that first navigation only.
    rc.InvalidateControl "GoToNextButton"
    rc.InvalidateControl "GoToPrevButton"
End Sub
