Attribute VB_Name = "basChangeLogaeBibleClass"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

'=============================================================================================================================
' Tasks:
' #065 -
' #064 -
' #063 -
' #062 -
' #061 -
' #060 - Add boolean test to check if any theme colors are used - Bible should use standard/defined colors, not themes
' #057 - Add ability to run only a specific test
' #055 - Update RunTest so expected gets values from Expected string array
' #053 - Add test for Footnote Reference followed by a space
' #048 - Use https://www.bibleprotector.com/editions.htm for comparison of KJV with Pure Cambridge Edition
' #047 - Research diff code that will display like GitHub for comparison with
' #046 - Update style of cv marker to be smaller than Verse marker
' #045 - Test call to SILAS from robbon using customUI.xml OnHelloWorldButtonClick routine
' #044 - Add extract to text file routine with book chapter reference - see web.txt from openbible.com as reference
' #043 - Add extract to USFM routine
' #042 - Add readme to aewordgit
' #041 - Add auto-generated TOC for maps
' #040 - Add figure headings to maps - use map vs fig?
' #039 - Replace manual TOC with auto-generated version
' #038 - Add test for no empty para after h2 heading
' #037 - Add updated maps in color
' #036 - Add test for h1 pages that have heading
' #035 - Add test for page numbers of h1 on odd or even pages
' #031 - Consider SILAS recommendation for adding pictures in text boxes to support USFM output
' #029 - Add versions of usfm_sb.sty to the SILAS folder to be able to track progress
' #024 - ExtractNumbersFromParagraph2 using DoEvents. Still unresponsive after Genesis 50, fifth para
' #017 - Add optional variant to aeBibleClass for indicating Copy (x) under testing
' #016 - Add funtion to print pass/fail based on comparing Result with Expected
'=============================================================================================================================
'
    ' FIXED - #059 - Add boolean flag to class to turn off timing for all tests
    ' FIXED - #058 - Add timer to each test and output total runtime of all tests
    ' FIXED - #054 - Add string array Expected to aeBibleClass to and initialize with RunTest expected values
    ' FIXED - #056 - Add test for white paragraph marks
' 20250323 - v005
    ' FIXED - #052 - Add vba message box with yes/no choice to continue or stop for RunTest error
    ' FIXED - #051 - Add Yes/No continuation message to RunTest error
    ' FIXED - #050 - Error Test num = 11 Function RunTest - need to fix it
    ' FIXED - #049 - Add test for count of empty paragraphs with no theme color, wdColorAutomatic = -1
    ' FIXED - #025 (Ref #034) - Check if para is continuous break or section break next page then read the next para
    ' OBSOLETE - #027 - Create SILAS dir and add Normal.dot then extract the code to GitHub - code provided by Jim
    ' FIXED - #034 - Add routine to count of all paragraphs types
    ' FIXED - #033 - Add Hello World custom menu tab as example for ribbon integration
    ' FIXED - #032 - Revert RunTest (12) as form feeds are needed in page and section breaks
    ' FIXED - #030 - Add routine to count and review Form feed char positions. Needed in docx as part of page and section breaks
' 20250317 - v004
    ' FIXED - #028 - Add test to count Hex 12 i.e. Form feed - it can cause Word not responding
    ' FIXED - #026 - Add debugging code to deal with empty paragrahs in ExtractNumbersFromParagraph2
    ' FIXED - #022 - Add routine to print book h1, chapter h2, verse number - based on #021
    ' FIXED - #023 - PrintBibleHeading1Info outputs the CR of Heading 1. Remove it so output is all on one line
    ' FIXED - #021 - Add routine to print Bible book headings
    ' FIXED - #020 - Add routine to print Bible book heading details - could be used as manual page number verification
    ' FIXED - #019 - Add module for interactive slow tests not in aeBibleClass
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

