Attribute VB_Name = "basChangeLog_aeWordGitClass"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

'=============================================================================================================================
' Tasks:
' #020 -
' #019 -
' #018 -
' #009 - Add setup info to the docm source file
' #006 - Can't execute code in break mode - error after doc saved from template and opened. Use error trapping in ThisDocument
'=============================================================================================================================
'
    ' FIXED - #017 - Add Const E_FAIL As Long = -2147467259    ' Unspecified Error(E_FAIL)
' 20260315 - v008
    ' FIXED - #016 - Use aeWordGit, Note: uppercase WordGit throughout for readability
' 20250217 - v007
    ' FIXED - #015 - Delete unused code: ThisIsAnAddIn, OutputListOfWordProperties, DeleteVBAModulesAndUserForms
    ' FIXED - #014 - Update usage instruction in basTESTaeWordGitClass
    ' FIXED - #013 - Configure aeWordGit owner as adaept
' 20250210 - v006
    ' FIXED - #012 - If current folder is not aeWordGit then export to src as user default
' 20250209 - v005
    ' FIXED - #011 - Add Yes No MessageBox when deleting src *.* files so as to confirm correct setup location
' 20250207 - v004
    ' FIXED - #010 - Error 448 when running EXPORT_THE_CODE, varDebug not passed correctly
    ' FIXED - #008 - Update to use c:\adaept\aeWordGit\src\ as default - repo is now in the github adaept organization
    ' OBSOLETE - #004 - Add an About setion in Ambigram tab to show version and logo
    ' OBSOLETE - #003 - Word 2019 Preview does not show 2016 Ambigram ribbon tab, report bug to Avenius
' 20190608 - v003
    ' FIXED - #007 - Compile error for x64, needs PtrSafe
' 20180920 - v002
    ' FIXED - #005 - Zoom full screen and page for dotm and new doc
' 20180909 - v001 - FIXED - #001 - Implement simple test for dropdown code
    ' FIXED - #002 - Change project name to ambigram and export to .\src
' 20180903 - v000 - Use aexlgitClass as starting model for aeWordGitClass


