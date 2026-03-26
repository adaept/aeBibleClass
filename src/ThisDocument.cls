VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Document_Open: intentionally empty.
' The commented-out lines were used only for document-open timing tests.
' No startup logic is required here; all initialisation is triggered manually.
Private Sub Document_Open()
    'ActiveWindow.View.Type = wdNormalView
    'ActiveDocument.ActiveWindow.Caption = "Loaded at: " & Now()
End Sub

