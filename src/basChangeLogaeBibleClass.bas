Attribute VB_Name = "basChangeLogaeBibleClass"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

'=============================================================================================================================
' Tasks:
' #125 -
' #124 -
' #123 -
' #116 - Check use of Gentium font (make it unnecessary?)
' #109 - Add test for CountAllEmptyParagraphs in doc, headers, footers, footnotes, and textboxes
' #108 - Add test for all line feed to have a space before
' #106 - Fix H1 pages to use line feed in text as appropriate
' #100 - Continue check multipage view from 300 for orphans of H2
' #095 - Fix GetColorNameFromHex to match the chosen Bible RGB colors
' #088 - Add test for Footnote Reference to list those that do not have bold in the paragraph
' #083 - Update name of Bible to Refined Word Bible (RWB) - Michael
' #082 - Fix Word paragraph style so minimal empty paragraphs are needed
' #076 - Update all Arial Black emphasis to new style. It should demonstrate significance in EDP.
' #075 - Add style for Arial Black 8 pt emphasis.
' #070 - Word automatically adjusts smart quotes to match the context of the text
'        Add test for Verse marker followed by any closing quote
' #069 - Use WEB.doc to get a proper count of "'" and make sure REV is correct
'        Verify smart quotes
'        Several Bible versions use smart quotes for opening and closing quotations,
'        including the triple quote style for verses like Ezekiel 39:7
'        Additional example versions that follow this style:
'           New International Version (NIV):
'               Uses smart quotes for direct speech and often includes nested quotes for emphasis.
'           English Standard Version (ESV):
'               Employs smart quotes and follows similar formatting conventions for nested quotations.
'           New Living Translation (NLT):
'               Uses smart quotes and maintains clear distinctions between different levels of speech.
'           New King James Version (NKJV):
'               Adopts smart quotes and includes nested quotes for clarity.
'           Christian Standard Bible (CSB):
'               Utilizes smart quotes and nested quotations for direct speech.
' #060 - Add boolean test to check if any theme colors are used - Bible should use standard/defined colors, not themes
' #057 - Add ability to run only a specific test
' #048 - Use https://www.bibleprotector.com/editions.htm for comparison of KJV with Pure Cambridge Edition
' #047 - Research diff code that will display like GitHub for comparison with
' #046 - Update style of cv marker to be smaller than Verse marker
' #045 - Test call to SILAS from ribbon using customUI.xml OnHelloWorldButtonClick routine
' #044 - Add extract to text file routine with book chapter reference - see web.txt from openbible.com as reference
' #043 - Add extract to USFM routine
' #042 - Add readme to aewordgit
' #041 - Add auto-generated TOC for maps
' #040 - Add figure headings to maps - use map vs fig?
' #039 - Replace manual TOC with auto-generated version
' #037 - Add updated maps in color
' #035 - Add test for page numbers of h1 on odd or even pages
' #031 - Consider SILAS recommendation for adding pictures in text boxes to support USFM output
' #029 - Add versions of usfm_sb.sty to the SILAS folder to be able to track progress
' #024 - ExtractNumbersFromParagraph2 using DoEvents. Still unresponsive after Genesis 50, fifth para
'=============================================================================================================================
'
    ' FIXED - #122 - Add test for count linefeed and space linefeed in doc
    ' FIXED - #115 - Add style "TheFooters" based on "TheHeaders" and update all footer sections, use Noto Sans font
    ' FIXED - #121 - Update debug output of Expected1BasedArray for Test(x) to be 15 per line
    ' FIXED - #120 - Add test for "TheHeaders" style as there should be only one paragraph mark per section
    ' FIXED - #118 - Add test for use of "Header" style, should be 0 as "TheHeaders" has to be used instead
    ' FIXED - #112 - Clear all tab stops from para headers, default is 0.1", add one tab to empty headers
    ' FIXED - #117 - See #113 - Add test to count tab followed by paragraph mark in headers
    ' FIXED - #119 - See #113 - Add test to count paragraph mark in headers that does not have a tab
    ' FIXED - #114 - Add style "TheHeaders"
    ' FIXED - #107 - Fix lamentations to use manual line break (line feed) with Lamentation style
    ' FIXED - #113 - Add test for empty and non empty header paragraphs
    ' FIXED - #111 - Fix empty paragraphs in text boxes
    ' FIXED - #110 - Fix empty paragraphs in footers
    ' FIXED - #087 - Set 1st/2nd paras after H1 to CustomParaAfterH1 or CustomParaAfterH1-2nd and verify vertical pos of Bbs
    ' FIXED - #086 - Define style CustomParaAfterH1 for vertical position of Brief background summary (Bbs)
    ' FIXED - #085 - Add routine to tools to check the vertical position of a paragraph
' 20250408 - v007
    ' FIXED - #105 - Update chapter and verse markers to orange and emerald
    ' FIXED - #104 - Add routine to set Winword as high priority for vba
    ' FIXED - #103 - Use UpdateCharacterStyle in batches from a page number
    ' FIXED - #102 - Add LongRunningProcessCode skeleton to allow resume and percent completed output to console
    ' FIXED - #097 - Some footnote references reset to red - why? - fix it to be consistent style using font color automatic
    ' FIXED - #101 - Update cvmarker to Chapter Verse marker
    ' FIXED - #099 - Add test to count number of colors in Footnote Reference
    ' FIXED - #098 - Add test to count number of Footnote References
    ' FIXED - #096 - Add test for count/delete empty para before H2, related #084
    ' FIXED - #084 - Update Heading 2 style paragraph to before 12 pt and delete the previous empty para
    ' OBSOLETE - #017 - Add optional variant to aeBibleClass for indicating Copy (x) under testing
    ' FIXED - #094 - Add test to List And Count Font Colors, and print the name from a conversion function
    ' FIXED - #090 - Work through Count Spaces After Footnotes debug output and fix as appropriate, split from ch/v numbers
    ' FIXED - #016 - Add function to print pass/fail based on comparing Result with Expected
    ' FIXED - #066 - Add tests to count paragraphs, empty,
    ' FIXED - See #073 - #067 - Add test to Count Red Footnote References
    ' FIXED - See #091 - #053 - Add test for Footnote Reference followed by a space
    ' FIXED - #089 - Continue find of footnote followed by space ("^f ") from 500 on, and fix as appropriate
    ' FIXED - #093 - Add initial PassFail test for result and expected
    ' FIXED - #080 - Review all footnote references so that, as much as possible, they are at the end of a paragraph
' 20250402 - v006
    ' FIXED - #091 - Add test for CountSpacesAfterFootnotes - also shows Footnote References and Following Chars (ASCII Val)
    ' FIXED - #092 - Add test for CountFootnotesFollowedByDigit
    ' FIXED - #073 - Run test to verify count of red footnote reference is zero
    ' FIXED - #072 - Check red footnote reference from Genesis to end of Study Bible
    ' FIXED - #071 - Finish check of red footnote reference from Ezek 39 to end of Bible
    ' FIXED - #038 - Add test for no empty para after h2 headings in doc - total count should be 0
    ' FIXED - #079 - Resolve issue around name of REV Bible - see #083
    ' FIXED - #078 - Add test to count number of h1 heading, should be 66 for Bible books
    ' FIXED - #074 - Set Heading 1 to 144 points before, follows section break so each book is on a new page with
    '           no empty first para, and delete existing 144 pt empty para
    ' FIXED - #081 - Check Books with only one chapter and verify references only use verse number per SBL abbreviations
    '           Obad, Phlm, 2 John, 3 John, Jude
    ' FIXED - #077 - Check Ezek for three in a row closing quotes
    ' FIXED - #068 - Check Ezek 1 to 26 for proper use of "'" and Ezek 39 to end of book for "'"
'        Double quotes to indicate the start and end of the direct speech.
'        Single quotes within the double quotes to emphasize the words spoken by God.
'        Closing double quotes to complete the direct speech.
'           Opening double quote: “ (ASCII code: 147 or Unicode: U+201C)
'           Closing double quote: ” (ASCII code: 148 or Unicode: U+201D)
'           Opening single quote: ‘ (ASCII code: 145 or Unicode: U+2018)
'           Closing single quote: ’ (ASCII code: 146 or Unicode: U+2019)
'        These smart quotes are different from the straight quotes (" and ') which have ASCII codes 34 and 39, respectively.
'        To insert these characters manually, you can use the following key combinations in Word:
'           Opening double quote: Alt + 0147
'           Closing double quote: Alt + 0148
'           Opening single quote: Alt + 0145
'           Closing single quote: Alt + 0146    ' FIXED - #064 - When bTimeAllTests is True it does not show total time
    ' FIXED - #063 - Update RunTest so it will allow a range of tests to be run (15 tests range)
    ' FIXED - #065 - Add module basTESTaeBibleTools for routines that are useful to tests outside of the class
    ' OBSOLETE - Replaced with #062 - #036 - Add test for h1 pages that have heading
    ' FIXED - #062 - Add test for Sections With Different FirstPage selected
    ' FIXED - #055 - Update RunTest so expected gets values from Expected string array
    ' FIXED - #061 - Add variant get array function of Expected to aeBibleClass and initialize with RunTest expected values
    ' FIXED - #059 - Add boolean flag to class to turn off timing for all tests
    ' FIXED - #058 - Add timer to each test and output total runtime of all tests
    ' OBSOLETE - #054 - Add string array Expected to aeBibleClass and initialize with RunTest expected values
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

