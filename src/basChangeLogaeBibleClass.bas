Attribute VB_Name = "basChangeLogaeBibleClass"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

'====================================================================================================================================
' Tasks: [doc] [test] [bug] [perf] [audit] [disc] [feat] [idea] [impr] [flow] [cp] [code] [wip] [clean] [obso] [regr]
' #310 -
' #309 -
' #308 -
' #307 -
' #306 - Add audit log from squash #274 [doc] [audit]
' #305 - Check header writing standard, vbnet vs. vba [cp] [clean]
' #303 - Fix single RUN_THE_TESTS(x) so it does not run AppendToFile and kill the full report [bug]
' #295 - Verify use of late binding in all code base so there is no need to set references [code]
' #294 - Cut a 0.1.1 release and tag it on GitHub [doc] [cp]
' #293 - Add md doc 'Bias Guard' to reduce hallucination (h13n) [doc] [cp]
' #292 - Add md doc describing use of Copilot for documentation creation [doc] [cp]
' #291 - See #300 - Add md doc that shows clearly the workflow for GitHub integration [doc] [flow]
' #290 - Add test for count of H1 with style [test]
' #289 - Add test for count of H2 with style [test]
' #288 - Create md doc file describing use of Tasks labels [doc]
' #287 - Update labels for Tasks and retroactively link to historic issues [doc] [impr] [cp]
' #281 - Explain methodology of Test Driven Development [doc]
' #280 - Add test to count H2, "How many Chapters are in the Bible", Copilot -> 1,189
' #271 - Add routine headers for targeting github.io docs in future [doc] [cp]
' #268 - Timings of TestReport to go in csv log file with session ID for each run [impr]
' #267 - Add code for CompleteAuditPageLayout [code]
' #266 - Create design for new routine CompleteAuditPageLayout in md format - Pre, Scan, Post [doc]
' #259 - Remove old code that regressed [clean]
' #247 - see #279 - Add code to define H1 and H2 exactly and apply to all [code] [doc] [impr]
' #226 - Update CompareHeading1sWithShowHideToggle to use CheckShowHideStatus [impr]
' #221 - Add test that will compare DOCVARIABLEs with result of PrintHeading1sByLogicalPage for page verification [test]
' #214 - Fix contents page to include all bookmarked Heading_01+ numbers
' #206 - Add test for all H1 pages to verify no paragraphs have indent setting [test]
' #195 - Improve verse find - Ps 119:176 is most verses, search is 14 secs, Psalm has most chapters (150), search is 2 secs
' #191 - Add test to verify all correct Verse Marker per book [test]
' #190 - Add test to verify all correct Chapter Verse Marker per book [test]
' #170 - Check doc and use line feed instead of paragraph mark throughout where verses are divided
' #151 - Add test for PrintCompactSectionLayoutInfo, number of one and two col sections, and print layout report file
' #150 - Add module for free fonts setup and testing
' #109 - Add test for CountAllEmptyParagraphs in doc, headers, footers, footnotes, and textboxes
' #100 - Continue check multipage view from 300 for orphans of H2
' #095 - Fix GetColorNameFromHex to match the chosen Bible RGB colors
' #083 - Update name of Bible to Refined Word Bible (RWB) - Michael
' #070 - Word automatically adjusts smart quotes to match the context of the text
'        Add test for Verse marker followed by any closing quote [test]
' #069 - Use WEB.doc to get a proper count of "'" and make sure RWB is correct
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
' #060 - Add boolean test to check if any theme colors are used - Bible should use standard/defined colors, not themes [test]
' #048 - Use https://www.bibleprotector.com/editions.htm for comparison of KJV with Pure Cambridge Edition
' #047 - Research diff code that will display like GitHub for comparison with verse versions
' #044 - Add extract to text file routine with book chapter reference - see web.txt from openbible.com as reference
' #043 - Add extract to USFM routine
' #042 - Add readme to aewordgit
' #040 - Add figure headings to maps - use map vs fig?
' #037 - Add updated maps in color
' #035 - Add test for page numbers of h1 on odd or even pages [test]
' #031 - Consider SILAS recommendation for adding pictures in text boxes to support USFM output
' #029 - Add versions of usfm_sb.sty to the SILAS folder to be able to track progress
' #024 - ExtractNumbersFromParagraph2 using DoEvents. Still unresponsive after Genesis 50, fifth para
'====================================================================================================================================
'
    ' FIXED - #279 - Add routine to define H2 style and reapply it in the project, add code header [impr] [code]
'Sq ' FIXED - #300 - Add md doc to outline a Compact Strategy for Squashed Audit Commits and reduce GitHub commit log spam
    ' FIXED - #304 - Add task type [wip] - it will prepend the task commits until replaced by FIXED
    ' FIXED - #302 - Update PrintCompactSectionLayoutInfo to output in rpt folder, move to basTESTaeBibleTools and add doc header
    ' FIXED - #301 - 999 AppendToFile should be "SKIPPED" [bug]
    ' FIXED - #274 - Fix output path so 'Style Usage Distribution.txt' goes to rpt folder, add code header [bug] [doc]
    ' FIXED - #299 - Add initial README and Bias Guard md files [doc] [cp]
    ' FIXED - #298 - Use SSOT with Select Case statements for values such as num and verify with RUN_THE_TESTS
    ' FIXED - #297 - Create file to hold Audits for Commit Log
    ' FIXED - #296 - Add code for ValidateTaskInChangelogModule
    ' FIXED - #286 - Update Heading 2 with DisableKeepLinesTogetherForHeading2
    ' FIXED - #285 - Update Heading 2 with EnforceHeading2WidowOrphanControl
    ' FIXED - #284 - Update Heading 2 KeepWithNext [audit]
    ' FIXED - #283 - Add code GetHeadingDefinitionsWithDescriptions to tools [audit]
    ' FIXED - #282 - Update guide with 'Example of Tags use for Audit Clarity' [doc]
    ' FIXED - #278 - Use Single Source of Truth (SSOT) to fix multiple locations of array definition via MaxTests - see #273
    ' FIXED - #277 - Define standard for types of "Tasks" to use with git commit messages [doc]
    ' FIXED - #273 - New error: Erl = 0 Error = 9 (Subscript out of range) in procedure RunBibleClassTests of Class BibleClass [bug]
    ' FIXED - #276 - git mv TestReport to rpt/ and delete old version
    ' FIXED - #275 - Create md folder for docs - md format, target github.io in future, git mv "Editorial Design and Style Guide.md"
    ' FIXED - #272 - Add section on Architecture Overview: DOCM-Coupled Macro System
    ' FIXED - #269 - All reports to be output to rpt folder
    ' FIXED - #265 - Add SKIP option to RUN_THE_TESTS for slow tests. Return -1 in report log, and GetPassFail return SKIP!!!!
    ' FIXED - #270 - Add test for SummarizeHeaderFooterAuditToFile
    ' FIXED - #264 - Add test for Style Usage Distribution [test]
    ' FIXED - #263 - Add CountAuditStyles_ToFile [test]
    ' FIXED - #262 - Update code module names to match EDSG manifest
    ' FIXED - #261 - Add initial Editorial Design and Style Guide [doc]
    ' FIXED - #257 - Update SmartPrefixRepairOnPage to give a count of Ascii 160 chars and any other e.g. hair space
    ' FIXED - #260 - Update RepairWrappedVerseMarkers_MergedPrefix_ByColumnContext_SinglePage to give a count of Ascii 12 chars
    ' FIXED - #258 - Add RunRepairWrappedVerseMarkers_Across_Pages_From to allow per page testing
    ' [obso] [regr] - #256 - Update SmartPrefixRepairOnPage to give a count of Ascii 12 chars
    ' FIXED - #255 - Update SmartPrefixRepairOnPage for details on unprintable characters
    ' FIXED - #254 - Add code for FindInvisibleFormFeeds_InPages
    ' FIXED - #253 - Add code for LogExpandedMarkerContext
    ' FIXED - #252 - Add code SmartPrefixRepairOnPage with Diagnostic Counter
    ' FIXED - #251 - Add header to csv forecast output file
    ' FIXED - #250 - Wire up dummy repair test with stats collection logic
    ' FIXED - #249 - Add skeleton for StartRepairTimingSession
    ' FIXED - #248 - Update repair tool for 10 pages
    ' FIXED - #174 - Add tests for count tab para in headers and footers [test]
    ' FIXED - #088 - Add tests for Footnote Reference (in doc and footnote) to count those that are not bold with correct style [test]
    ' FIXED - #246 - Add test for styles using Liberation Sans Narrow [test]
    ' FIXED - #245 - Add code Identify_ArialUnicodeMS_Paragraphs
    ' FIXED - #244 - Unlink heading numbering, should not display Article... or Section... for H1 or H2
    ' FIXED - #243 - Add code RedefinePictureCaptionStyle_NotoSans, step 3 of removing Lieration Sans Narrow reference
    ' FIXED - #242 - Add code RedefineFootnoteNormalStyle_NotoSans, step 2 of removing Lieration Sans Narrow reference
    ' FIXED - #241 - Add code RedefineFootnoteStyle_NotoSans, step 1 of removing Lieration Sans Narrow reference
    ' FIXED - #240 - Update all repair code and add runner for checking 5 pages at a time
    ' FIXED - #239 - Add routine ReportDigitAtCursor_Diagnostics_Expanded
    ' FIXED - #238 - Update Chapter Verse marker repair tool with latest RepairWrappedVerseMarkerPrefixes_AdjacencyWithContext_Navigate
    ' FIXED - #237 - Add diagnostic code to get character information around the cursor position
    ' FIXED - #236 - Add routine to report Report Page Layout Metrics for a particular page
    ' FIXED - #235 - Add code to repair "Chapter Verse marker" per page, add vbCr if on column edge with space before, defrag as needed
    ' FIXED - #234 - Add test to count footers that have only a tab character [test]
    ' FIXED - #212 - Add test for CountFindNotEmphasisBlack = 0 and CountFindNotEmphasisRed = 0 when all have been set [test]
    ' FIXED - #233 - Add test for CountParagraphMarks_CalibriDarkRed [test]
    ' FIXED - #232 - Add word version into to output and test report
' 20250719 - v010
    ' FIXED - #148 - Add version info to TestReport output
    ' FIXED - #231 - Reapply explicit formatting (Segoe UI 8, Bold, Blue, Superscript) for Footnote Reference, Fix for #230
    ' [obso] - #230 - Add code to fix Footnote Reference by reapplying style
    ' FIXED - #225 - Add code to verify Show/Hide is True when tests are run else stop with error message
    ' FIXED - #229 - Add code to verify all necessary settings of Word are enabled - basWordSettingsDiagnostic
    ' FIXED - #228 - Abort tests if Show/Hide is not set
    ' FIXED - #227 - Update CheckShowHideStatus as ActiveWindow.View.ShowAll is the only reliable source of truth
    ' FIXED - #224 - Fix error in CheckShowHideStatus to make it reliable
    ' FIXED - #223 - Add routine with two different ways to check Show/Hide status
    ' FIXED - #222 - Add routine to compare Heading 1s with Show/Hide toggled
    ' FIXED - #220 - Update DOCVARIABLEs based on results of PrintHeading1sByLogicalPage
    ' FIXED - #219 - Add routine to count search hits with match case true
    ' FIXED - #218 - Add routine to print logical page numbers with Heading 1, in a list, for Bible book page check
    ' FIXED - #217 - Update "I am The lord" to "I am the Lord" x42
    ' FIXED - #210 - See #213 - WoJ emphasised is 9pt, use that in search then set to 8pt as style EmphsasisRed
    ' FIXED - #184 - See #211 - Add test for Footnote Text to count those that have any bold text [test]
    ' FIXED - #215 - Add test for paragraph mark styled - Calibri 9 Dark Red - should be color Automatic [test]
    ' FIXED - #216 - Error with H1 count of 66 vs 59 for show/hide true false
'       Problem list = "Joshua", "2 Kings", "Nehemiah", "Habakkuk", "Haggai", "Philemon", "1 Peter"
'           The issue wasn’t with the styles or outline levels themselves, but with invisible or corrupted inline content
'           (probably non-printing characters or hidden formatting) hiding in those paragraphs. When one cleaned one (Joshua),
'           it likely triggered a reflow/re-rendering in Word that corrected the others.
'       Solution - Click at the end of the word "Joshua" and press Delete Then press Enter once.
'           This clears any hidden/invisible content after the heading text that may prevent proper recognition.
'           Reselect the paragraph and reapply Heading 1 style
    ' FIXED - #211 - Add test for CountBoldFootnotesWordLevel [test]
    ' FIXED - #213 - Add test for Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed = 0 when all have been set [test]
    ' FIXED - #142 - Add routine to output book names and pages for TOC manual verification - see #039
    ' FIXED - #209 - Add Section Nav button to ribbon for bookmarked Heading_00 to Heading_12 sections
    ' FIXED - #186 See #209 - Add ribbon button for index sections - Intro etc.; OT; NT; Maps; DP;
    ' FIXED - #162 See #209 - Update routines to allow page num checks for Heading 0 sections
    ' FIXED - #161 See #209 - Create bookmarkers for other sections that are not the Bible
    ' FIXED - #208 - Make style for Dating, Authorship, and Refer to maps - with 6pt spacing before and update all H1 pages
    ' FIXED - #207 - Check H1 pages for consistent use of superscript as in 2nd etc.
    ' FIXED - #106 - Fix H1 pages to use line feed in text as appropriate
    ' FIXED - #205 - Goto next book on next- button click in constant cycle (Note: getEnabled is flaky so do not use for now)
    ' [impr] - FIXED - #204 - Add next- book button to ribbon : Refer also to customUI14backupRWB.xml
    ' [obso] - #157 - Add Word OT DOCVARIABLEs, Ctrl + F9 field brackets { }, right-click the field, select Update Field - verify
    ' FIXED - #045 - Test call to SILAS from ribbon using customUI14.xml OnHelloWorldButtonClick routine
    ' [obso] - #041 - Add auto-generated TOC for maps : auto gen too slow
    ' FIXED - See #203 - #160 - Add DOCVARIABLEs for all New Testament books
    ' FIXED - #203 - Add DOCVARIABLEs for New Testament
    ' FIXED - #202 - Move GoToVerseSBL to ribbon module
    ' FIXED - #201 - Add synch-to-onedrive.bat for adaept folders
    ' FIXED - #200 - Add search for Chapter and Verse marker styles preceded and followed by a space
    ' FIXED - #199 - Add ribbon command for GoToH1 - Bible Book
    ' FIXED - #198 - Add adaept prototype about button to ribbon
    ' FIXED - #197 - Add Is to book map for Isaiah
    ' FIXED - #196 - Add bookMap Ecc "Ecclesiastes"
    ' FIXED - #194 - Set cursor to spinning when searching and restore on completion
    ' FIXED - #185 - Add ribbon with bible search button for GoToVerseSBL
    ' FIXED - #193 - Manually export RWB ribbon xml to a file by copying from Office RibbonX Editor
    ' FIXED - #192 - Add ribbon button for GoToVerseSBL
    ' FIXED - #189 - Update map for min case and selection for single chapter verses
    ' FIXED - #188 - Create an Excel file for a list of map options to Books, add min matches
    ' FIXED - #187 - Error Ps 37:19 is finding Isaiah 37:19 - Look for PSALM as H2 instead of CHAPTER
    ' FIXED - #183 - Error "1 Sam" - book not found, also for all books starting with a number, should return e.g. "1 Sam 1:1"
    ' FIXED - #182 - GoToVerseSBL search for Jude 5 fails (only 1 chapter) - solve for Obad, Phlm, 2 John, 3 John, Jude - see #081
    ' FIXED - #181 - #180 GoToVerseSBL regression for Gen 2:2
    ' FIXED - #180 - GoToVerseSBL fails with invalid format if only chapter entered - update so it finds verse 1
    ' FIXED - #179 - Add compare documents routine
    ' FIXED - #178 - Add SBL goto verse routine
    ' FIXED - #177 - Check tabernacle references in Exodus and update footnote 103 accordingly
' 20250529 - v009
    ' FIXED - #176 - Define Normal style as Calibri 9 to fix #175
    ' FIXED - #175 - Gentium font found at para 13964 - procedure FindGentiumFromParagraph, use GoToParagraph to check
    ' FIXED - #116 - Check use of Gentium font (make it unnecessary?) - See #175
    ' FIXED - #076 - Update all Arial Black emphasis to new style. It should demonstrate significance in EDP.
    ' FIXED - #075 - Add style for Arial Black 8 pt emphasis.
    ' FIXED - #173 - Rename CountTabOnlyParagraphs to CountDocTabOnlyParagraphs
    ' FIXED - #172 - Add test for CountParagraphMarks_ArialBlackDarkRed - 8pt, Automatic or Black, paragraph marks only [test]
    ' FIXED - #171 - Add test for CountParagraphMarks_ArialBlack - 8pt, Automatic or Black, paragraph marks only [test]
    ' FIXED - #168 - Add style for emphasized Words of Jesus - EmphasisRed
    ' FIXED - #169 - Add code for FindNotEmphasisBlackRed to return 0 when completed
    ' FIXED - #167 - Rename FastFindArialBlack8ptNormalStyleSkipEmphasisBlack to FindNotEmphasisBlackRed
    ' FIXED - #166 - Update FastFindArialBlack8ptNormalStyleSkipEmphasisBlack to also check font color Auatomatic
    ' FIXED - #165 - Add code FastFindArialBlack8ptNormalStyleSkipEmphasisBlack
    ' FIXED - #164 - Create style EmphasisBlack
    ' FIXED - #163 - Add code for CheckOpenFontsWithDownloads
    ' FIXED - #159 - Run TestPageNumbers to verify page numbers stored in all Old Testament DOCVARIABLEs
    ' FIXED - #158 - Add restart capability from location to FindNextHeading1OnVisiblePage
    ' FIXED - #154 - Add test for DOCVARIABLE "Gen", give it a page value, if wrong show error for updating
    ' FIXED - #156 - Add code FindDocVariableEverywhere
    ' FIXED - #155 - Add code for FindNextHeading1OnVisiblePage and remember found location for next search
    ' FIXED - #153 - Add code for GetExactVerticalScroll - return the scroll percentage rounded to three decimal places
    ' FIXED - #152 - Update test 36 to stop if Footer style found
    ' FIXED - #143 - Clone original SILAS as Jim put it on GitHub
    ' FIXED - #057 - Add ability to run only a specific test
    ' FIXED - #147 - Add date/time output to TestReport
    ' FIXED - #145 - Add global const TestReportFlag for optional output using OutputTestReport function
    ' FIXED - #146 - TestReport FAIL!!!! message is not correctly aligned - pad PASS to be "PASS    "
    ' FIXED - #149 - Replace Expected with oneBasedExpectedArray(x) as appropriate
    ' FIXED - #144 - Add code for checking fonts used
    ' FIXED - #123 - Add file TestReport.txt output additional to console result for GitHub tracking
    ' FIXED - #046 - Update style of cv marker to be smaller than Verse marker
    ' FIXED - #082 - Fix Word paragraph style so minimal empty paragraphs are needed
    ' [obso] - #039 - Replace manual TOC with auto-generated version (this is too slow)
    ' FIXED - #141 - Update UTF8bom-Template.txt with multiple language sample of "Hello, World!" ala C style, plus phonetics
' 20250420 - v008
    ' FIXED - #140 - Set version info as global variables and assign in Class_Initialize
    ' FIXED - #139 - Add UTF8bom-Template.txt with BunnyEgg emoji for Easter using :emojisense in VS Code
    ' FIXED - #133 - Store actual result is 1 based results array for comparison without recalc
    ' FIXED - #138 - Create 1 based array for storing results
    ' FIXED - #137 - Update test to notSpaceCount CountNotSpacesAfterFootnoteReferences [test]
    ' FIXED - #108 - Add test for all line feed to have a space before (Test 32 and 33) [test]
    ' FIXED - #136 - Add back test for CountEmptyParagraphs [test]
    ' FIXED - #135 - Fix sections where different first page is selected - deselect them
    ' FIXED - #134 - Output debug formatting header to console for comma spacing
    ' FIXED - #131 - Add DoEvents to number dash number search and stop switch to doc window for ISBN
    ' FIXED - #132 - Add test for tab paragraph mark only [test]
    ' FIXED - #130 - Update CountEmptyParagraphs to CountEmptyParagraphsWithFormatting
    ' FIXED - #129 - Add DoEvents in long loops so console results are processed
    ' FIXED - #128 - Update test CountEmptyParagraphs for speed [test] [perf]
    ' FIXED - #127 - Update test CountNumberDashNumberInFootnotes with fast algorithm [test] [perf]
    ' FIXED - #126 - Update test CountDeleteEmptyParagraphsBeforeHeading2 with fast algorithm from ChatGPT [test] [perf]
    ' FIXED - #125 - Add test to count number of footers with style "Footer" [test]
    ' FIXED - #124 - Add test for count linefeed and space linefeed in footnotes [test]
    ' FIXED - #122 - Add test for count linefeed and space linefeed in doc [test]
    ' FIXED - #115 - Add style "TheFooters" based on "TheHeaders" and update all footer sections, use Noto Sans font
    ' FIXED - #121 - Update debug output of Expected1BasedArray for Test(x) to be 15 per line
    ' FIXED - #120 - Add test for "TheHeaders" style as there should be only one paragraph mark per section [test]
    ' FIXED - #118 - Add test for use of "Header" style, should be 0 as "TheHeaders" has to be used instead [test]
    ' FIXED - #112 - Clear all tab stops from para headers, default is 0.1", add one tab to empty headers
    ' FIXED - #117 - See #113 - Add test to count tab followed by paragraph mark in headers [test]
    ' FIXED - #119 - See #113 - Add test to count paragraph mark in headers that does not have a tab [test]
    ' FIXED - #114 - Add style "TheHeaders"
    ' FIXED - #107 - Fix lamentations to use manual line break (line feed) with Lamentation style
    ' FIXED - #113 - Add test for empty and non empty header paragraphs [test]
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
    ' FIXED - #099 - Add test to count number of colors in Footnote Reference [test]
    ' FIXED - #098 - Add test to count number of Footnote References [test]
    ' FIXED - #096 - Add test for count/delete empty para before H2, related #084 [test]
    ' FIXED - #084 - Update Heading 2 style paragraph to before 12 pt and delete the previous empty para
    ' [obso] - #017 - Add optional variant to aeBibleClass for indicating Copy (x) under testing
    ' FIXED - #094 - Add test to List And Count Font Colors, and print the name from a conversion function
    ' FIXED - #090 - Work through Count Spaces After Footnotes debug output and fix as appropriate, split from ch/v numbers
    ' FIXED - #016 - Add function to print pass/fail based on comparing Result with Expected
    ' FIXED - #066 - Add tests to count paragraphs, empty, [test]
    ' FIXED - See #073 - #067 - Add test to Count Red Footnote References
    ' FIXED - See #091 - #053 - Add test for Footnote Reference followed by a space
    ' FIXED - #089 - Continue find of footnote followed by space ("^f ") from 500 on, and fix as appropriate
    ' FIXED - #093 - Add initial PassFail test for result and expected
    ' FIXED - #080 - Review all footnote references so that, as much as possible, they are at the end of a paragraph
' 20250402 - v006
    ' FIXED - #091 - Add test for CountSpacesAfterFootnotes - also shows Footnote References and Following Chars (ASCII Val) [test]
    ' FIXED - #092 - Add test for CountFootnotesFollowedByDigit [test]
    ' FIXED - #073 - Run test to verify count of red footnote reference is zero [test]
    ' FIXED - #072 - Check red footnote reference from Genesis to end of Study Bible
    ' FIXED - #071 - Finish check of red footnote reference from Ezek 39 to end of Bible
    ' FIXED - #038 - Add test for no empty para after h2 headings in doc - total count should be 0 [test]
    ' FIXED - #079 - Resolve issue around name of REV Bible - see #083
    ' FIXED - #078 - Add test to count number of h1 heading, should be 66 for Bible books [test]
    ' FIXED - #074 - Set Heading 1 to 144 points before, follows section break so each book is on a new page with
    '           no empty first para, and delete existing 144 pt empty para
    ' FIXED - #081 - Check Books with only one chapter and verify references only use verse number per SBL abbreviations
    '           Obad, Phlm, 2 John, 3 John, Jude
    ' FIXED - #077 - Check Ezek for three in a row closing quotes
    ' FIXED - #068 - Check Ezek 1 to 26 for proper use of "'" and Ezek 39 to end of book for "'"
'        Double quotes to indicate the start and end of the direct speech.
'        Single quotes within the double quotes to emphasize the words spoken by God.
'        Closing double quotes to complete the direct speech.
'           Opening double quote:   (ASCII code: 147 or Unicode: U+201C)
'           Closing double quote:   (ASCII code: 148 or Unicode: U+201D)
'           Opening single quote:   (ASCII code: 145 or Unicode: U+2018)
'           Closing single quote:   (ASCII code: 146 or Unicode: U+2019)
'        These smart quotes are different from the straight quotes (" and ') which have ASCII codes 34 and 39, respectively.
'        To insert these characters manually, you can use the following key combinations in Word:
'           Opening double quote: Alt + 0147
'           Closing double quote: Alt + 0148
'           Opening single quote: Alt + 0145
'           Closing single quote: Alt + 0146    ' FIXED - #064 - When bTimeAllTests is True it does not show total time
    ' FIXED - #063 - Update RunTest so it will allow a range of tests to be run (15 tests range)
    ' FIXED - #065 - Add module basTESTaeBibleTools for routines that are useful to tests outside of the class
    ' [obso] - Replaced with #062 - #036 - Add test for h1 pages that have heading
    ' FIXED - #062 - Add test for Sections With Different FirstPage selected [test]
    ' FIXED - #055 - Update RunTest so expected gets values from Expected string array
    ' FIXED - #061 - Add variant get array function of Expected to aeBibleClass and initialize with RunTest expected values
    ' FIXED - #059 - Add boolean flag to class to turn off timing for all tests
    ' FIXED - #058 - Add timer to each test and output total runtime of all tests
    ' [obso] - #054 - Add string array Expected to aeBibleClass and initialize with RunTest expected values
    ' FIXED - #056 - Add test for white paragraph marks [test]
' 20250323 - v005
    ' FIXED - #052 - Add vba message box with yes/no choice to continue or stop for RunTest error
    ' FIXED - #051 - Add Yes/No continuation message to RunTest error
    ' FIXED - #050 - Error Test num = 11 Function RunTest - need to fix it [test]
    ' FIXED - #049 - Add test for count of empty paragraphs with no theme color, wdColorAutomatic = -1 [test]
    ' FIXED - #025 (Ref #034) - Check if para is continuous break or section break next page then read the next para
    ' [obso] - #027 - Create SILAS dir and add Normal.dot then extract the code to GitHub - code provided by Jim
    ' FIXED - #034 - Add routine to count of all paragraphs types
    ' FIXED - #033 - Add Hello World custom menu tab as example for ribbon integration
    ' FIXED - #032 - Revert RunTest (12) as form feeds are needed in page and section breaks
    ' FIXED - #030 - Add routine to count and review Form feed char positions. Needed in docx as part of page and section breaks
' 20250317 - v004
    ' FIXED - #028 - Add test to count Hex 12 i.e. Form feed - it can cause Word not responding [test]
    ' FIXED - #026 - Add debugging code to deal with empty paragrahs in ExtractNumbersFromParagraph2
    ' FIXED - #022 - Add routine to print book h1, chapter h2, verse number - based on #021
    ' FIXED - #023 - PrintBibleHeading1Info outputs the CR of Heading 1. Remove it so output is all on one line
    ' FIXED - #021 - Add routine to print Bible book headings
    ' FIXED - #020 - Add routine to print Bible book heading details - could be used as manual page number verification
    ' FIXED - #019 - Add module for interactive slow tests not in aeBibleClass
    ' FIXED - #015 - Add test for count number dash number in footnotes only [test]
    ' FIXED - #018 - Update Copy(???) in test runner to default Copy () as current version
    ' FIXED - #014 - Add test for count number dash number [test]
    ' FIXED - #013 - Add test to count number of nonbreaking spaces [test]
    ' FIXED - #012 - Add test to count number of period space left parenthesis [test]
    ' FIXED - #011 - Add test to count style with number and space [test]
    ' FIXED - #010 - Add copy(???) to output as placeholder for revision under test
    ' FIXED - #009 - Add test to count style with space and number [test]
    ' FIXED - #008 - Add test to count quadruple paragraph marks [test]
' 20250221 - v003
    ' FIXED - #007 - Add test to count space followed by carriage return with white font color [test]
    ' FIXED - #006 - Add test to count number of double tabs [test]
    ' FIXED - #005 - Add test to count space followed by carriage [test]
    ' FIXED - #004 - Add tests to count double spaces in doc, and in shapes including groups [test]
    ' FIXED - #003 - Change module name to basTESTaeBibleClass
' 20250219 - v002
    ' FIXED - #002 - Update class name to aeBibleClass
' 20250217 - v001
    ' FIXED - #001 - Create Bible Class base template, initial test module, and change log


