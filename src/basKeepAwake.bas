Attribute VB_Name = "basKeepAwake"
Option Explicit
Option Private Module

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

' ============================================================================
' basKeepAwake
' ----------------------------------------------------------------------------
' Keeps Windows from going to sleep or hibernating on its idle timer while a
' long run is in progress (RUN_THE_TESTS, ExportScriptureOnlyDocx). Added
' 2026-10-17 after an unattended Test 94 run slept mid-run: VBA Timer is wall
' clock, so the sleep was counted as run time (1866 s instead of about 287 s).
'
' Uses the Windows API SetThreadExecutionState with ES_SYSTEM_REQUIRED. It
' changes NO power settings, so there is nothing to save or restore: Windows
' removes the request by itself when the Word process exits or crashes. If VBA
' stops on an unhandled error before StopKeepAwake runs, the request stays
' until Word closes (the machine just stays awake - harmless).
'
' What it does NOT do: it does not block closing the lid (that follows the
' lid-close action of the power plan), a manual Sleep/Hibernate/Shut down, or
' sign-out. The display may still turn off; the run continues. See
' rvw/Code_review 2026-10-09.md in aeBibleAddin, section 27.
' ============================================================================

Private Declare PtrSafe Function SetThreadExecutionState Lib "kernel32" (ByVal esFlags As Long) As Long

Private Const ES_SYSTEM_REQUIRED As Long = &H1
Private Const ES_CONTINUOUS As Long = &H80000000

Private m_active As Boolean

' True when the keep-awake request was accepted (the API returns the previous
' state, or 0 on failure).
Public Function StartKeepAwake() As Boolean
    StartKeepAwake = (SetThreadExecutionState(ES_CONTINUOUS Or ES_SYSTEM_REQUIRED) <> 0)
    m_active = StartKeepAwake
End Function

' Clears the request (back to normal idle sleep behaviour). Safe to call when
' not active.
Public Sub StopKeepAwake()
    If m_active Then
        SetThreadExecutionState ES_CONTINUOUS
        m_active = False
    End If
End Sub

Public Function IsKeepAwakeActive() As Boolean
    IsKeepAwakeActive = m_active
End Function
