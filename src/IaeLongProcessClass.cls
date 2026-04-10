VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IaeLongProcessClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

' =============================================================================
' IaeLongProcessClass - Interface for Long-Running Process Tasks
' -----------------------------------------------------------------------------
' This class defines the contract that all long-running task implementations
' must satisfy. It is not instantiated directly -- declare variables As
' IaeLongProcessClass and assign a concrete implementation (e.g. New aeUpdateCharStyleClass).
'
' Implementing a task:
'   1. Add "Implements IaeLongProcessClass" to the class module declaration.
'   2. Implement all three members with the IaeLongProcessClass_ prefix:
'        Private Function IaeLongProcessClass_ExecuteItem(itemIndex As Long) As Boolean
'        Private Property Get IaeLongProcessClass_ItemCount() As Long
'        Private Property Get IaeLongProcessClass_TaskName() As String
'
' Contract:
'   ExecuteItem  - Called once per item. Return True to continue, False to stop.
'   ItemCount    - Total number of Items to process (used for progress calculation).
'   TaskName     - Short identifier used in log output and progress sidecar filename.
' =============================================================================

' Execute one item of work. Return True to continue, False to stop the task.
Public Function ExecuteItem(itemIndex As Long) As Boolean
End Function

' Total number of Items to process.
Public Property Get ItemCount() As Long
End Property

' Short task identifier (used in log output and rpt/ sidecar filename).
Public Property Get TaskName() As String
End Property
