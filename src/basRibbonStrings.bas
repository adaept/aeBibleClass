Attribute VB_Name = "basRibbonStrings"
Option Explicit
Option Private Module

' =============================================================================
' basRibbonStrings - Ribbon String Resources
' -----------------------------------------------------------------------------
' All user-facing strings used by the ribbon are declared here as named
' constants. Logic procedures never contain inline string literals for UI text.
'
' i18n: to localise the ribbon, edit only this module. The ribbon XML
' (customUI14.xml) uses getKeytip callbacks that read from these constants —
' it never needs to change for a localisation.
'
' VSTO port: replace this module with a .resx resource file. Constant names
' map directly to resource keys.
'
' Keytip conventions (English):
'   - ComboBox controls: first letter of the navigation level (B/C/V)
'   - Prev/Next buttons: punctuation with natural directional meaning
'   - Action buttons: first letter of the action (S=Search, A=About)
' =============================================================================

' -- KeyTips -------------------------------------------------------------------

Public Const KT_BOOK         As String = "B"   ' Book comboBox
Public Const KT_CHAPTER      As String = "C"   ' Chapter comboBox
Public Const KT_VERSE        As String = "V"   ' Verse comboBox
Public Const KT_PREV_BOOK    As String = "["   ' Previous Book button
Public Const KT_NEXT_BOOK    As String = "]"   ' Next Book button
Public Const KT_PREV_CHAPTER As String = ","   ' Previous Chapter button
Public Const KT_NEXT_CHAPTER As String = "."   ' Next Chapter button
Public Const KT_PREV_VERSE   As String = "<"   ' Previous Verse button
Public Const KT_NEXT_VERSE   As String = ">"   ' Next Verse button
Public Const KT_NEW_SEARCH   As String = "S"   ' New Search button
Public Const KT_ABOUT        As String = "A"   ' About (adaept) button
