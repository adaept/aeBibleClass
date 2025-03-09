Attribute VB_Name = "basChangeLogaeBibleClass"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

'=============================================================================================================================
' Tasks:
' #020 -
' #019 - Add test to count cr/number/number/space
' #017 - Add optional variant to aeBibleClass for indicating Copy (x) under testing
' #016 - Add funtion to print pass/fail based on comparing Result with Expected
'=============================================================================================================================
'
    ' FIXED - #015 - Add test for count number dash number in footnotes only
    ' FIXED - #018 - Update Copy(???) in test runner to default Copy () as current version
    ' FIXED - #014 - Add test for count number dash number
    ' FIXED - #013 - Add test to count number of nonbreaking spaces
    ' FIXED - #012 - Add test to count number of period space left parenthesis
    ' FIXED - #011 - Add test to count style with number and space
    ' FIXED - #010 - Add copy(???) to output as placeholder for revision under test
    ' FIXED - #009 - Add test to count style with space and number
    ' FIXED - #008 - Add test to count quadruple paragraph marks
' 20250221 - v003
    ' FIXED - #007 - Add test to count space followed by carriage return with white font color
    ' FIXED - #006 - Add test to count number of double tabs
    ' FIXED - #005 - Add test to count space followed by carriage
    ' FIXED - #004 - Add tests to count double spaces in doc, and in shapes including groups
    ' FIXED - #003 - Change module name to basTESTaeBibleClass
' 20250219 - v002
    ' FIXED - #002 - Update class name to aeBibleClass
' 20250217 - v001
    ' FIXED - #001 - Create Bible Class base template, initial test module, and change log

