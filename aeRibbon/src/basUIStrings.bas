Attribute VB_Name = "basUIStrings"
Option Explicit
Option Private Module

' =============================================================================
' basUIStrings - UI String Resources
' -----------------------------------------------------------------------------
' All user-facing strings used by the ribbon and status bar are declared here
' as named constants. Logic procedures never contain inline string literals
' for UI text.
'
' i18n: to localise the UI, edit only this module. The ribbon XML
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
'   Documentation in Ribbon Design.md must match these constants exactly.
'   Ribbon tab keytip (Alt, Y2) is a static keytip="Y2" attribute on the <tab> element
'   in customUI14.xml. The <tab> element does not support a getKeytip callback, so
'   this value cannot be expressed as a constant in this module.
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
Public Const KT_GO           As String = "G"   ' Go (navigate) button
Public Const KT_SEARCH       As String = "S"   ' New Search button
Public Const KT_ABOUT        As String = "A"   ' About (adaept) button

' -- Labels (i18n: ribbon labels returned by getLabel callbacks) ----------------

Public Const LBL_TAB      As String = "Radiant Word Bible"  ' <tab id="RWB"> label
Public Const LBL_GROUP    As String = "Bible Navigation"    ' <group id="NavGroup"> label
Public Const LBL_GO       As String = "Go"                  ' GoButton label
Public Const LBL_ABOUT    As String = "About"               ' adaeptButton label

' -- Control IDs (ribbon XML id= attributes) -----------------------------------
' Use these constants in all InvalidateControl calls to prevent silent mismatches.

Public Const CTRL_BOOK         As String = "cmbBook"
Public Const CTRL_CHAPTER      As String = "cmbChapter"
Public Const CTRL_VERSE        As String = "cmbVerse"
Public Const CTRL_PREV_BOOK    As String = "PrevBookButton"
Public Const CTRL_NEXT_BOOK    As String = "NextBookButton"
Public Const CTRL_PREV_CHAPTER As String = "PrevChapterButton"
Public Const CTRL_NEXT_CHAPTER As String = "NextChapterButton"
Public Const CTRL_PREV_VERSE   As String = "PrevVerseButton"
Public Const CTRL_NEXT_VERSE   As String = "NextVerseButton"

' -- Status bar messages -------------------------------------------------------
' Static messages: no runtime data — use directly as Application.StatusBar = SB_xxx
' Dynamic messages: contain {0}, {1} placeholders — use FormatMsg(SB_xxx, arg0, arg1)

Public Const SB_NAVIGATING            As String = "Navigating ..."
Public Const SB_WARM_CACHE            As String = "Bible: building navigation index..."
Public Const SB_INVALID_BOOK          As String = "Invalid input for Book - enter a book name or abbreviation"
Public Const SB_INVALID_CHAPTER       As String = "Invalid input for Chapter - out of range (1-{0})"
Public Const SB_INVALID_VERSE         As String = "Invalid input for Verse - out of range (1-{0})"
Public Const SB_ALREADY_FIRST_BOOK    As String = "Already at first book"
Public Const SB_ALREADY_LAST_BOOK     As String = "Already at last book"
Public Const SB_ALREADY_FIRST_CHAPTER As String = "Already at first chapter of {0} (1-{1})"
Public Const SB_ALREADY_LAST_CHAPTER  As String = "Already at last chapter of {0} (1-{1})"
Public Const SB_ALREADY_FIRST_VERSE   As String = "Already at first verse of {0} {1} (1-{2})"
Public Const SB_ALREADY_LAST_VERSE    As String = "Already at last verse of {0} {1} (1-{2})"

Public Function FormatMsg(ByVal template As String, ParamArray args() As Variant) As String
    Dim Result As String
    Result = template
    Dim i As Long
    For i = 0 To UBound(args)
        Result = Replace(Result, "{" & i & "}", CStr(args(i)))
    Next i
    FormatMsg = Result
End Function
