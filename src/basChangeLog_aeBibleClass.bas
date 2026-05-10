Attribute VB_Name = "basChangeLog_aeBibleClass"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

'====================================================================================================================================
' Tasks: [doc] [test] [bug] [perf] [audit] [disc] [feat] [idea] [impr] [flow] [code] [wip] [clean] [obso] [regr] [refac] [opt]
' #640 -
' #639 -
' #638 -
' #637 -
' #636 -
' #635 - see #634 - List all fonts used in the docx and track problem fonts Arial/Times New Roman [impr] [audit]
' #628 - aeBibleClass CountFindNotEmphasisBlack Test 45 obsolete, can be reused (see #623) [obso]
' #627 - aeBibleClass CountFindNotEmphasisRed Test 46 obsolete, can be reused (see #623) [obso]
' #625 - aeBibleClass Test 44 obsolete, can be re-used (see #623) [obso]
' #622 - Add World English Bible Updates_ChangeLog.txt (view-source:https://worldenglish.bible/webupdates.php) and work through the changes [wip]
' #621 - Add 2012-12-28 World English Bible lang_ChangeLog.txt (https://ebible.org/Scriptures/changelog.txt) and work through the changes [wip]
' #619 - Make and implement styles Poetry 1,2,3 - see email notes; no indent for this version but it allows flexibility [feat]
' #612 - **Use "Default Paragraph Font" for ALL character styles. Never use "(no style)".**
'            It's the difference between: Stable, predictable, audit-safe formatting vs.
'               Hidden inheritance, inconsistent rendering, and style-tree instability [feat]
' #609 - Soft Hyphens checked to end of Exodus - See #389 [wip]
' #606 - Add function CountInvisibleCharacters and include in BibleClass test, expected Result = 0 [test]
' #393 - Add glossary of terms used in Divine Principle from first reference in the Bible [idea]
' #391 - Create a test to Count all 1st 2nd 3rd etc. abbreviations - goal is to - 0, 1st Century ->
'           CountNumericOrdinals Numeric Ordinal Suffix Counts: st: 7 nd: 12 rd: 4 th: 44 TOTAL: 67
' #389 - Fix doc formatting using Optional Hyphen Alt+Ctrl+- (manual hyphenation) [wip]
' #365 - Map styles to USFM markers [wip]
' #314 - Add a routine to extract all the Words of Jesus into the "Jesus Document" [Idea]
' #288 - Create md doc file describing use of Tasks labels [doc]
' #109 - Add test for CountAllEmptyParagraphs in doc, headers, footers, footnotes, and textboxes [test]
' #095 - Fix GetColorNameFromHex to match the chosen Bible RGB colors
' #060 - Add boolean test to check if any theme colors are used - Bible should use standard/defined colors, not themes [test]
' #047 - Research diff code that will display like GitHub for comparison with verse versions [idea]
'====================================================================================================================================
'
'Sq
    ' [obso] - for next version - #037 - Add updated maps in color [feat]
    ' [obso] - for next version - #040 - Add figure headings to maps - use map vs fig? [idea]
    ' [obso] - Export to USFM - #336 - Gen 41:45 console output shows box for manual line break (Shift+Enter) - needs special consideration for file output [feat]
    ' [obso] - Export to USFM - #396 - Export - Psalms 110:7 He will drink of the brook on the way; therefore he will lift up his head. PSALM 111 [bug]
    ' FIXED - #601 - Build Word configuration module for consistent editing setup [feat][wip]
    ' [obso] - #069 - Use WEB.doc to get a proper Count of "'" and make sure RWB is correct
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
    ' FIXED in earlier code - #191 - Add test to verify all correct Verse Marker per book [test] [wip]
    ' FIXED in earlier code - #190 - Add test to verify all correct Chapter Verse Marker per book [test] [wip]
    ' [obso] - #150 - Add module for free fonts setup and testing [idea]
    ' [obso] #394 - Export of Psalms 72:20 to immediate windows shows BOOK 3 PSALM 73 A Psalm by Asaph at the end. [bug]
    ' [obso] - #418 - Extend the parser (SBL, UBS, NRSV, etc.) [impr] [feat] - Maybe re-open when i18n is real
    ' FIXED in earlier code - #615 - Duplication using `PopulateCanonical` - not using the DRY Principle [bug]
    ' [obso] - #400 - Check #399 & #401 with WEB/WEBU doc/USFM data [idea]
    ' FIXED in RunSoftHyphenSweep_Across_Pages_From - #620 - Make test to find stray hyphens in column text, cf. RepairWrappedVerseMarkers_MergedPrefix_ByColumnContext_SinglePage [test]
    ' FIXED - #630 - Add styles for author end matter [impr]
    ' FIXED in earlier code - #403 - See #422 - Bible text paragraph should start with Chapter/Verse styles. Verify numbers [test] [wip]
    ' FIXED in earlier code - #631 - Test for Count of "Chapter Verse marker" did not catch missing marker [bug] - see #190, #191, #403
    ' FIXED - #634 - Add AuditFonts_Fast report to list all used fonts and show usage of Arial/Times New Roman [impr]
    ' FIXED - #633 - Delete style BobyTextIndent when not used [impr]
    ' FIXED - #632 - Add ListBodyTextIndentUsage to check if style BobyTextIndent is used [impr]
    ' [obso] - #617 - SBL Canonical is Song of Songs, using Solomon causes variation so keep to standard in all parts [impr]
    ' [obso] - #484 - Store the Verse Map as a Byte Array [impr] - Revisit if server demand warrants it
    ' [obso] - #402 - Export shows "Acts of the Apostles", from Book header instead of H1. Create test "H1 text"="Book Header" [test][bug]
    ' [obso] - #618 - Add So as search alias - works for Song and Solomon [impr]
    ' [obso] - #626 - Function FindNotEmphasisBlackRed obsolete (see #623) [obso]
    ' FIXED - #624 - Change all aeBibleClass search for Calibri (now using Carlito so add global const) [bug]
    ' FIXED - #629 - Mark aeBibleClass Test 45, 46 obsolete in code [impr]
    ' FIXED - #623 - Update aeBibleClass with CountParagraphMarksWithDarkRedFormatting, expected 0 [impr]
    ' FIXED - #616 - Space function is changed to space, need normalizer fix [bug]
    ' [obso] - #596 - Psalms header not carried over with routine to auto-load from new doc clone [bug]
    ' FIXED - #395 - Add style Selah, where the word is italic (\qs for USFM) [impr]
    ' FIXED - #247 (as part of Styles config) - see also #279 - Add code to define H1 and H2 exactly and apply to all [code] [doc] [impr]
    ' [obso] - #214 - Fix contents page to include all bookmarked Heading_01+ numbers
    ' [obso] - #610 - Add Inspect_Aptos_Sources to aeBibleClass tools - expected Result 0 [test]
    ' FIXED - #614 - Use a paragrah style for Aleph, Bet etc., done in earlier commit [impr]
    ' FIXED - #613 - Add character style for Selah, small caps [impr]
    ' FIXED - #611 - Replace Calibri with Carlito (SIL OFL) 1. ReplaceCalibriInStyles; 2. ReplaceCalibriWithCarlito [feat]
    ' [obso] - #453 - Create class aeBibleDataClass to share values of Books, Chapters, Verses etc. with validation tests for arrays [feat] - Done in Ribbon code
    ' [obso] - #374 - Error search book Jeremiah, and verse Jeremiah 18:6 [bug] - Done in Ribbon code
    ' [obso] - #324 - Add index generation code to ribbon [impr] [feat] - Done in Ribbon code
    ' [obso] - #322 - Timeout on #195 (5.19 seconds), need more speed improvement [bug] [impr] [perf] - Done in Ribbon code
    ' [obso] - #357 - Search Gen 120 finds Psalms 120, error in FindChapterH2 [bug] - Done in Ribbon code
    ' [obso] - #351 - Check GoToVerseSBL input after Trim - no multi spaces; max spaces = 2; if exists ":" only 1; digit before & after ":" [impr] - Done in Ribbon code
    ' [obso] - #345 - Add routine to find chapter heading H2 based on a given H1 paragraph index [impr] [refac] - Done in Ribbon code
    ' [obso] - #340 - GoToVerseSBL 3 John 5, 3 John 1:5 - error 'No verse 5 found in Chapter 1' [bug] - Done in Ribbon code
    ' [obso] - #339 - On page 243 error Joshua not found search Josh 24:19, also check conv to UCase in verse find [bug] - Done in Ribbon code
    ' [obso] - #363 - Search Judges 15:11 Book Not Found [bug] [regr] - Done in Ribbon code
    ' [obso] - #411 - Fix reloading of all code routine to include SBL EBNF module [bug] - See basImportWordGitFiles
    ' FIXED - #417 - Add an SBL auto-corrector [feat] - See basTEST_aeBibleCitationBlock
    ' [obso] - #419 - Add typing look ahead, similar to Access combo box (see #417) [feat] - Done in Ribbon code
    ' #567 - Implement GoTo Verse using headingData in aeRibbonClass - speedup [feat][perf] - Done in Ribbon code
    ' FIXED - #597 - New Search should set the focus in cmbBook and not the docm [bug]
    ' FIXED - #599 - First load Gen tab tab tab 119 tab sets focus in docm, second use tab will go through all controls [bug]
    ' FIXED - #600 - Consider Enter button in ribbon to activate search Pros/Cons [idea] Done using Go button
    ' FIXED - #608 - Add i18n strings for status messages
    ' FIXED - #607 - Include status message for invalid input of Book/Chapter/Verse and status for Prev/Next out of bounds [impr]
    ' FIXED - #598 - Gen tab fills C/V with 1/1 but does not enable C/V Prev/Next buttons [bug]
    ' FIXED - #605 - Add ReportFootnoteComplexity - it should return 0 complex in 1000 footnotes [test]
    ' FIXED - #604 - Add FindParagraphsByFirstCharFont_BodyHeadersFooters, should matcch Result of AuditFontUsage_ParagraphsAndHeadersFooters [test]
    ' FIXED - #603 - Add find font routine and fix use of Palatino - verify by using AuditFontUsage_ParagraphsAndHeadersFooters [impr]
    ' FIXED - #602 - Only print styles with priority <> 99 [impr]
    ' [obso] - #492 - Add a step for Verse Boundary Validation - Done in latest DSP code
    ' [obso] - #491 - Add a step for Cross-Book Range Validation - Done in latest DSP code
    ' [obso] - #490 - Add a step for Chapter/Book Expansion Awareness - Done in latest DSP code
    ' [obso] - #489 - Add a step for Canonical Verse Ordering - Done in latest DSP code
    ' [obso] - #488 - Add a step called span normalization or range consolidation - Done in latest DSP code
    ' FIXED - #595 - Implement New Navigation Architecture: Default-Fill + Action-Gate [refac][feat]
    ' FIXED - #594 - Add summary notice related to RWB as trademark [feat]
    ' FIXED - #593 - Tab into Verse not working [bug]
    ' FIXED - #592 - Navigation interface bugs - use editBox for Chapter and Verse [bug]
    ' FIXED - #591 - Implement Step 3 [feat]
    ' FIXED - #590 - Step 4 is next - two visibility changes [impr]
    ' FIXED - #589 - Step 1 Implementation
    ' FIXED - #588 - Combo box in ribbon for Chapter and Verse are shorter [bug]
    ' FIXED - #587 - New Ribbon - skeleton implementation of design [feat]
    ' FIXED - #586 - Implement short circuit of TestUpdateCharStyle and also deal with chapter and verse markers at the same time [perf]
    ' FIXED - #585 - Re-capture headings if array is empty - e.g. if IDE is reset for long running process [bug]
    ' FIXED - #584 - Run TestUpdateCharStyle to check long running process is working
    ' FIXED - #583 - Long process plan implementation [feat]
    ' FIXED - #582 - Verify aeLoggerClass works with Run_All_SBL_Tests to create rpt/SBL_Tests.UTF8.txt [test][feat]
    ' FIXED - #581 - Create plan for improvement of long running task process code [impr]
    ' FIXED - #580 - Error - GoTo Book no longer activates after the last fix [regr]
    ' FIXED - #579 - Add GoToH1Deferred as the solution from test of #578
    ' FIXED - #578 - Create TestGoToH1Direct test outside of ribbon to find root cause of double layout spinning block with Revelation [test]
    ' FIXED - #577 - Run specific test to find cause of second 12 sec delay [test]
    ' FIXED - #576 - GoToH1: InvalidateControl After ScreenUpdating=True Causes Second 12-Second Block [bug]
    ' FIXED - #575 - GoToH1: Range.Select Causes Double Layout Pass; Switch to Find [bug]
    ' FIXED - #574 - NextButton Leaves Next Enabled at Revelation; PrevButton Same at Genesis [bug]
    ' FIXED - #573 - GoToH1: Double Layout Pass Caused by InvalidateControl [bug]
    ' FIXED - #572 - Screen flash and 3 second blank screen at Rev -> Jude boundary [bug]
    ' FIXED - #571 - Error in Gen and Rev behavior of Prev/Next buttons [bug]
    ' FIXED - #570 - Implement plan that does not allow wrap-around for Prev/Next book selection [feat]
    ' FIXED - #569 - See #566, #568 -  Pagination still slow for first GoTo Prev when Genesis is the selected book [bug][refac]
    ' FIXED - #568 - Prevent screen updating to reduce pagination cost [perf]
    ' FIXED - #566 - GoTo Prev is very slow on first use once Genesis is entered as the GoTo Book [perf]
    ' FIXED - #565 - #561 finally resolved by #562, #563, #564 bug fixes
    ' FIXED - #564 - No error message but the Prev Next ribbon buttons are enabled instead of disabled [bug]
    ' FIXED - #563 - Type mismatch error twice on ribbon opening [bug]
    ' FIXED - #562 - Wrong number of arguments on ribbon load [bug]
    ' FIXED - #561 - Prev and Next Book buttons should be disabled until GoTo Book is used once
    ' FIXED - #560 - Add ribbon button and code for GoTo Previous Book [impr]
    ' FIXED - #559 - Text selection for citation block captures the end paragraph mark and messes up format on paste [bug]
    ' FIXED - #558 - Whole chapter reference error [bug]
    ' FIXED - #557 - Sentinel fail = -1 not respected [bug]
    ' FIXED - #556 - A  paragraph that contains non-citation text raises an unfriendly error message [bug]
    ' FIXED - #555 - Whole chapter reference is not supported correctly [bug]
    ' FIXED - #554 - The message box window will not allow a fix then retry in citation block repair [bug]
    ' FIXED - #553 - Gen 2:17; 2:25; 3:6-11; Job 31:33; Matt 5:27; 15:11-17; Jas 1:14-15 - Gen output is wrong [bug]
    ' FIXED - #552 - Add Run_Extra_Tests to the citation block tests [impr]
    ' FIXED - #551 - Fix single chapter book issue with citation block verification [bug]
    ' FIXED - #550 - Allow text selection also for citation block verification [impr]
    ' FIXED - #549 - The paragraph mark is deleted on paste of corrected citation block [bug]
    ' FIXED - #548 - Book names should not be repeated for the citation block output [bug]
    ' FIXED - #547 - Add ToSBLShortForm for corrected output of citation block [feat]
    ' FIXED - #546 - Normalize does not deal with Chr(11), manual line feed, so breaks parsing [bug]
    ' FIXED - #545 - Implement the plan for an interactive citation block test and update procedure [feat]
    ' FIXED - #544 - Create a plan for an interactive citation block test and update procedure [feat]
    ' FIXED - #543 - Malformed citation in block is silently skipped [bug]
    ' FIXED - #542 - Implement plan for en-dash in sorted citation block [feat]
    ' FIXED - #541 - Create a plan to ouput en-dash in citation blocks and sorted in canonical order [feat]
    ' FIXED - #540 - Change dash to en dash for verse ranges in doc aeBibleCitationClass.md [impr]
    ' FIXED - #539 - Error 5 in ParseCitationBlock [bug]
    ' FIXED - #538 - Implement code plan for Stage 13a [impr]
    ' FIXED - #537 - Update doc as the en-dash form is pre-normalized and never reaches the parser [doc]
    ' FIXED - #536 - Add a plan for dealing with semicolon use for inherited book name in aeBibleCitationClass [impr]
    ' FIXED - #535 - 1 unexpected FAIL in negative test - [bug]
    ' FIXED - #534 - Fix negative tests of basTEST_aeBibleCitationBlock to use aeAssert framework [impr]
    ' FIXED - #533 - Move python and associated files to py folder and adjust calling scripts [impr]
    ' FIXED - #532 - Move documentation of DSP for SBL Citation to its own md file [doc][impr]
    ' FIXED - #531 - Make code routines explicitly Public or Private in codebase [impr]
    ' FIXED - #530 - Update license headers
    ' FIXED - #529 - Error with 2 unqualified references
    ' FIXED - #528 - Stage 12 error - FAIL: range parsed (expected=John 3:16-18, actual=John 3:16-3:18)
    ' FIXED - #527 - Error in Stage 11 and 12 FAIL [bug]
    ' FIXED - #526 - Update test harness to call aeBibleCitationClass [impr]
    ' FIXED - #525 - There are 2 FAIL errors in Stage 13 [bug]
    ' FIXED - #524 - Normalize to Mid$ [impr]
    ' FIXED - #523 - Create ADODB logger for UTF8 with test data [feat]
    ' FIXED - #522 - Update all tests to use aeAssertClass [refac][impr]
    ' FIXED - #521 - Convert module basSBL_Citation_EBNF to class aeBibleCitationClass. Adjust tests to prevent name clashes [impr]
    ' FIXED - #520 - AddBookNameHeaders adds a second blank line in the header [bug]
    ' FIXED - #509 - AddBookNameHeaders routine
    ' FIXED - #519 - Implement error handler standard in aeBibleClass.cls [impr]
    ' FIXED - #518 - Update definition for standard error handler to be applied in aeBibleClass.cls [doc]
    ' FIXED - #517 - Add test and code for Stage 17 [feat]
    ' FIXED - #516 - Add doc for Stage 17 [doc]
    ' FIXED - #515 - Add test and code for Stage 16 [feat]
    ' FIXED - #514 - Add doc for Stage 16 [doc]
    ' FIXED - #513 - Add test and code for Stage 15 [feat]
    ' FIXED - #512 - Add documentation for Stage 15 [doc]
    ' FIXED - #511 - Add test and code for Stage 14 Canonical Compression [feat]
    ' FIXED - #510 - Fix 3 failures in Test_Harness - corrupted en dash [bug]
    ' FIXED - #509 - AddBookNameHeaders routine
    ' FIXED - #508 - Add Count routines for orphan headers and footers
    ' FIXED - #507 - Fix bug for adde code fixing the footers [bug][feat]
    ' FIXED - #506 - Add AuditDocument module to verify settings when inserted to new docm
    ' FIXED - #505 - Make ribbon code into a class singleton and cleanup [refac]
    ' FIXED - #494 - (done in some earlier commit) Move InterpretStructure from Test harness module to EBNF code module
    ' FIXED - #504 - Comment out ribbon code that is to be replaced by parser code [clean]
    ' FIXED - #503 - See #502 - Fix importer for all files as a generic solution not depending of a list of names
    ' FIXED - #502 - Add Python normalizer + batch file to run after export script completes, no more commit spam from casing drift [feat][bug]
    ' FIXED - #501 - Book Title fix for Export USFM, extraneous use of Heading 1
    ' FIXED - #500 - Replace GoTo SkipLogging with Exit Do [impr]
    ' FIXED - #499 - Correct Defensive Pattern - Capture the error immediately after the risky operation [impr]
    ' FIXED - #498 - Reset error handler immediately after call [bug]
    ' FIXED - #497 - Use assert tests with Stage 8 [impr]
    ' FIXED - #496 - AssertFalse not counting towards Result [bug]
    ' FIXED - #495 - Fix Stage 4 tests (wrong dupcilation of InterpretStructure code) [bug]
    ' FIXED - #493 - Add documentation for Stage 14 Canonical Compression
    ' FIXED - #487 - Add test and code for Stage 13 - Contextual Shorthand Expansion
    ' FIXED - #485 - Update documentation for Stage 13 - Contextual Shorthand Expansion
    ' FIXED - #486 - Rename basChangeLog files to include underscore
    ' FIXED - #483 - Add license LGPL3 and tighten up documentation to include Book-only expansion parsing
    ' FIXED - #482 - Move all Types into Citation_EBNF module
    ' FIXED - #422 - Add per-book chapter and verse bounds (e.g., Jude has max verse 25) [impr]
    ' FIXED - #439 - Implement range validation (e.g., Gen 1:1-2:3) cleanly = all of Genesis 1 plus the first three verses of Genesis 2 [impr]
    ' FIXED - #452 - Temporarily harden parts = Split(normalizedInput, " ") with a quick loop to skip empty tokens
    ' FIXED - #451 - Refactor ParseReferenceStub into proper stage calls
    '                   normalized = NormalizeInput(input), tokens = TokenizeReference(normalized), parsed = InterpretTokens(tokens)
    ' FIXED - #456 - Design extension hooks for future features (ranges, lists, multi-word books) without breaking the contract
    ' FIXED - #416 - Emit fix suggestions (1 JN -> 1 John 1:1)
    ' FIXED - #481 - See #480 - Update code base and test to deal with Book-Only Reference Handling
    ' FIXED - #480 - See #416 - Add to Stage 4 documentation - Book-Only Reference Handling
    ' FIXED - #479 - Add documentation for Stage 11 / 13 - ComposeList / Contextual Shorthand
    ' FIXED - #478 - Fix ScriptureList type to avoid UDT coerce error for late-bound function. Adjust code and tests as needed [bug]
    ' FIXED - #477 - Add test and code for Stages 11, 12
    ' FIXED - #474 - Add new architecture Stage 10, 11, 12, documentation. Test and code for Stage 10
    ' FIXED - #476 - Update extension architecture and explain invariant Stages 1-7
    ' FIXED - #475 - Add AssertFalse to make test harness cleaner
    ' FIXED - #473 - Introduce IsRangeSegment for further clarity when using hypen and en dash
    ' FIXED - #472 - Update Stage 9 doc to show clear distinction of hyphen and en dash (immediate window output is not clear enough)
    ' FIXED - #471 - Implement Stage 9 Range Detection
    ' FIXED - #470 - Extend Test 4 slightly to guarantee future code changes cannot alter the segmentation silently
    ' FIXED - #468 - Update DSP documentation for Extension Hooks Stage 8 initial skeleton, test, and code - lexical only
    ' FIXED - #469 - Clean up task list, mark tasks [obso] that have been dealt with through 7-Stage design
    ' FIXED - #425 - Code development steps: parser stub, semantic tightening, SBL enforcement rules, document-scale validation [impr] [wip]
    ' [obso] - #438 - Add an audit routine that validates MaxChapter against the data structure [impr]
    ' [obso] - #436 - Check Verse upper bound [feat]
    ' [obso] - #429 - Generate the canonical-name aliases automatically from GetCanonicalBookTable [impr] [opt]
    ' [obso] - #448 - Stub Must Normalize Single-Chapter Books - Belongs in the Parser (Not Validator)
    ' [obso] - #415 - Normalize case for output [impr]
    ' [obso] - #414 - Enforce no verse ranges across chapters [impr]
    ' [obso] - #413 - Enforce SBL punctuation (: vs .) [impr]
    ' [obso] - #221 - Add test that will compare DOCVARIABLEs with Result of PrintHeading1sByLogicalPage for page verification [test]
    ' [obso] - #083 - Update name of Bible to Refined Word Bible (RWB) - Michael [idea]
    ' [obso] - #042 - Add readme to aeWordGit [doc]
    ' [obso] - #035 - Add test for page numbers of h1 on odd or even pages [test]
    ' FIXED - #431 - Freeze parser stub scope, Strengthen semantic validator with tests, Add negative tests, Swap parser stub for real parser
    ' FIXED - #467 - Update LexicalScan (with multi-word book support)
    ' FIXED - #466 - Add End-To-End test
    ' FIXED - #465 - Add RUN_FAILURE_DEMOS guard
    ' FIXED - #464 - Match Test_Stagex... to actual routine names
    ' FIXED - #462 - Make AssertTrue show diagnostic information when a test fails, include a fail demo routine as example [impr]
    ' FIXED - #463 - Update all tests to use new assert framework
    ' FIXED - #461 - Update current test harness into something much closer to a real unit-test framework [refac]
    ' FIXED - #460 - Update Test harness to use these stage name routines
    ' FIXED - #459 - Update State Transition Diagram in ASCII form to match new 7 stage design
    ' FIXED - #458 - Add routine names directly under each "Stage Alignment" section
    ' FIXED - #457 - Update to formally express grammar in EBNF for 7 stage design
    ' FIXED - #455 - Refine contracts for 7 stages and add a 12-line formal "Parser Contract" header at the top of the module
    ' FIXED - #454 - Update documentation pipeline overview of 7 stages
    ' FIXED - #450 - Update doc and architectural structure for Stage 2: Lexical Tokenization
    ' FIXED - #449 - Update documentation to reflect status of Stage 1: Input Normalization [doc]
    ' FIXED - #447 - Normalize at Data construction boundary - the correct architectural layer
    ' FIXED - #446 - Enforce 1-Based array usage with assert statements and update documentation
    ' FIXED - #445 - Fix GetMaxVerse for 0 based maps array [regr][bug]
    ' FIXED - #444 - Fix error in dictionary and GetVerseCounts [bug]
    ' FIXED - #443 - Verify Packed map integrity automatically, abort test harness if corrupted, no silent execution, no manual step required
    ' FIXED - #442 - Verify packed verse map [impr]
    ' FIXED - #441 - Correct the Production Implementation (Use Packed Map Only) - runtime GetMaxVerse should not call GetVerseCounts() [bug]
    ' FIXED - #440 - Add the packed verse map using GeneratePackedVerseStrings_FromDictionary and call it from function GetPackedVerseMap [feat]
    ' FIXED - #437 - Generate the packed fixed-width strings automatically [impr]
    ' FIXED - #435 - Check Chapter upper bound - See commit 'Update basSBL_Citation_EBNF.bas'
    ' FIXED - #434 - Add full boundary enforcement: Chapter lower bound, Verse lower bound, Single-chapter special rule
    ' FIXED - #433 - Track failure reason strings (diagnostic only): Do not change control flow, Do not add new layers, Only record why a stage failed
    ' FIXED - #432 - Add negative tests: ResolveBook should succeed, ValidateSBLReference should fail, If ResolveBook fails, that is a test failure, not a pass.
    ' FIXED - #430 - Add a validator that asserts alias coverage completeness
    ' FIXED - #428 - Any book name that may appear as parser output must exist as a key in the alias map [bug]
    ' FIXED - #427 - Book-resolver table is not wired up for full canonical names yet [bug]
    ' FIXED - #426 - Parser stub feeds the semantic pipeline without changing it
    ' FIXED - #424 - Allow reset of AliasMap when running test harness (no use of Static) [bug]
    ' FIXED - #423 - Update initial test harness module [bug]
    ' FIXED - #421 - Add single-chapter book rewriting (Jude 5 > Jude 1:5) [impr]
    ' FIXED - #420 - Add function GetSBLCanonicalBookTable
    ' FIXED - #412 - Add resolver checks for EBNF module
    ' FIXED - #410 - Define canonical Bible book list for Excel/Access-style normalization [refac]
    ' FIXED - #409 - Create initial SBL Unified EBNF design for Bible citation parser
    ' FIXED - #392 - All Bible text paragraphs should be justified. Make test to Count (from P18-931) left justified. Expect 0 Result [test]
    ' FIXED - #398 - Fix RunRepairWrappedVerseMarkers_Across_Pages_From so it DOES NOT put one verse per para for v59 [bug]
    ' FIXED - #397 - Add global OneVersePerPara to separate v59 branch from Main [impr]
    ' FIXED - #408 - Add code from Module1 for #407 wip commit to aeBibleClass and then clean up
    ' FIXED - #407 - When searching for U+0020,U+201D the Result is 0. Using Ctrl+H returns 1 in the footer [bug]
    ' FIXED - #406 - CountUnicodeSeq is not used. CountContraction is the correct function, typo [bug]
    ' FIXED - #405 - Add test for space followed by right double closing quote: U+0020, U+201D [test]
    ' FIXED - #404 - Move test 66 " '" outside of CreateContractionArray so that it is with the UniCode character tests [refac]
    ' FIXED - #401 - Add test to Count - Double Opening Quote, Single Opening Quote, Double Opening Quote [test]
    ' FIXED - #399 - Add test to Count - Double Opening Quote, Single Opening Quote, Double Opening Quote [test]
    ' FIXED - #390 - Create routine to output U+ text from MakeUnicodeSeq and include in test 67, 68 debug and file output
    ' [obso] - #387 - Latest report timings indicate tests 42 and 51 are run when they should be skipped [bug] [regr] [audit]
    ' [obsp] - #206 - See #247 - Add test for all H1 pages to verify no paragraphs have indent setting [test]
    ' [obso] - #048 - Use https://www.bibleprotector.com/editions.htm for comparison of KJV with Pure Cambridge Edition [idea]
    ' [obso] - #268 - Timings of all TestReports to go in csv log file with session ID for each run [impr]
    ' [obso] - #291 - See #300 - Add md doc that shows clearly the workflow for GitHub integration [doc] [flow]
    ' [obso] - #151 - Add test for PrintCompactSectionLayoutInfo, number of one and two col sections, and print layout report file
    ' [obso ] #267 - Add code for CompleteAuditPageLayout [code]
    ' [obso] - #266 - Create design for new routine CompleteAuditPageLayout in md format - Pre, Scan, Post [doc]
    ' [obso] - #226 - Update CompareHeading1sWithShowHideToggle to use CheckShowHideStatus [impr]
    ' FIXED - #287 - Update labels for Tasks and retroactively link to historic issues [doc] [impr]
    ' [obso] - #292 - Add md doc describing use of Copilot for documentation creation [doc]
    ' [obso] - #293 - Add md doc 'Bias Guard' to reduce hallucination (h13n) [doc]
    ' [obso] - #271 - Add routine headers for targeting github.io docs in future [doc] [wip]
    ' [obso] - #170 - See #389 - Check doc and use line feed instead of paragraph mark throughout where verses are divided
    ' FIXED - #044 [wip] - See # 338, #337, #326, #258 - Add extract to text file routine with book chapter reference -
    '           see web.txt from openbible.com as reference [feat]
    ' [obso] - #281 - Explain methodology of Test Driven Development [doc]
    ' [obso] - #259 - Remove old code that regressed [clean]
    ' [obso] - #043 - See #365 - Add extract to USFM routine [feat]
    ' [obso] - #031 - Consider SILAS recommendation for adding pictures in text boxes to support USFM output [idea]
    ' [obso] - #029 - Add versions of usfm_sb.sty to the SILAS folder to be able to track progress [idea]
    ' FIXED - #388 - Quotes missing in contraction array debug and file output [bug]
    ' FIXED - #386 - Add code for DebugAndReportHeader, DRY [refac]
    ' FIXED - #385 - Simplify SKIP test process should always return Result -1, fix for Test 42 and 51, DRY [bug] [refac]
    ' FIXED - #070 - Word automatically adjusts smart quotes to match the context of the text
    '                   Add test for NNBSP followed by any right single closing quote (U+202F followed by U+2019) [test]
    '                   Add test for NNBSP followed by any right double closing quote (U+202F followed by U+201D) [test]
    ' FIXED - #384 - Add a function MakeUnicodeSeq that will make a string from 1~3 U code points
    ' FIXED - #383 - Add test for space followed by U+2019
    ' FIXED - #381 - Add test for Count of "spirit's", expected 1
    ' FIXED - #378 - Simplify use of contraction code [refac]
    ' FIXED - #382 - Add function to replace `'` with  Apostrophe, =ChrW$(AposCP), when calling GetPassFail routine for ResultArray 52+
    ' FIXED - #380 - Create Contraction Array and verify in RunTest 52 and 55
    ' FIXED - #379 - Separate initialization of actual and expected Result arrays from conversion to 1-base array
    ' FIXED - #377 - Add contractions code to test suite [impr]
    ' FIXED - #376 - Add routine to Count use of English contractions e.g. can't, for inclusion in test suite [feat]
    ' FIXED - #375 - Add routine to Show Unicode Of Single Character Selection and account for surrogate pairs as needed [feat]
    ' FIXED - #373 - Add style "Brief" for 'Brief background summary' as USFM \ip
    ' FIXED - #372 - Remove blank lines after verses - from \v in code
    ' FIXED - #371 - Add IsEffectivelyEmpty function to remove extraneous empty lines in exported USFM
    ' FIXED - #370 - Update USFM export for two layer detection model: paragraph and character-level semantics
    ' FIXED - #369 - Update USFM exporter for style definition "Chapter Verse marker" (orange), "Verse marker" (green)
    ' FIXED - #368 - Add a strict USFM validator that logs structural issues to a separate log file in UTF-8 format that uses current code.
    ' FIXED - #367 - UTF-8 output for USFM and Log files are including manual hyphenation characters [bug]
    ' FIXED - #366 - Write USFM file and Log file as UTF-8 with no BOM marker to be safe for Paratext
    ' FIXED - #364 - Add initial scaffolding for USFM export [feat]
    ' FIXED - #362 - Update LogHeadingData.txt for csv output per line of H1 with sessionID and paraIndex updated on change [feat] [impr]
    ' FIXED - #361 - LogHeadingData.txt is empty on first run of CaptureHeading1s [bug]
    ' FIXED - #360 - Add Judge to GetFullBookName [impr]
    ' FIXED - #359 - Add routine to print headingData to Immediate Window and write updated entries to session log file [code]
    ' FIXED - #358 - Add routine to capture Heading 1 text and paragraph index into a static array (1 to 66, 0 to 1) [code]
    ' FIXED - #356 - Speed up FindBookH1 [impr] [refac]
    ' FIXED - #352 - Add routine to find verse based on a given H2 paragraph index [impr] [refac]
    ' FIXED - #355 - Set bookMap to static so it is initialized only once per session and avoid overhead, simplify use of UCase [impr]
    ' FIXED - #354 - Error setting values of chapNum and verseNum for IsOneChapterBook [bug]
    ' FIXED - #353 - Error for '1 Joh' Book not found in FindBookH1 [bug]
    ' FIXED - #350 - Add function ExtractTrailingDigits that extracts the last 1~3 digits from a string and LeftUntilLastSpace
    ' FIXED - #349 - GoTo Book Next finds the next para - positioning error from select work around [bug]
    ' FIXED - #348 - Update search of GoToH1 so pattern does not need to use * or ? - matching the style of abbr in GetFullBookName [impr]
    ' FIXED - #344 - Search for '3 Joh' fails silently in FindBookH1 [bug]
    ' FIXED - #343 - Add [refac] as type task for Refactor [impr]
    ' FIXED - #342 - Add routine FindBookH1 [refac]
    ' FIXED - #341 - Add ParseParts routine to check user input [impr] [refac]
    ' FIXED - #347 - If book is not found the verse search jumps to Genesis but cursor should not move [bug]
    ' FIXED - #346 - Add function IsOneChapterBook [refac]
    ' FIXED - #338 - Use tab separator for console output of routine RunRepairWrappedVerseMarkers_Across_Pages_From [bug]
    ' FIXED - #337 - Josh 12:24 prints "CHAPTER 13" with console text for RunRepairWrappedVerseMarkers_Across_Pages_From. Same for all H2 [bug]
    ' [obso] #024 - ExtractNumbersFromParagraph2 using DoEvents. Still unresponsive after Genesis 50, fifth para [bug]
    ' FIXED - #335 - Add routine SaveAsPDF_NoOpen to avoid auto-open of the PDF with Edge [code]
    ' FIXED - #331 - Add function GetVerseText for console output [feat]
    ' FIXED - #334 - Normalize page to one verse per para and add Count of CRs added [feat]
    ' [regr] #328 - See #333 - Add code in ThisDocument to show/hide the word interface but keep the custom ribbon
    ' FIXED - #333 - Comment out ThisDocument.cls code as it interferes with clean export to docx [bug]
    ' FIXED - #332 - Add function TitleCase to convert header output [code]
    ' FIXED - #330 - Add function GetPageHeaderText and return it in debug output for Chapter/Verse [feat]
    ' FIXED - #329 - Chapter/Verse output missing when marker is at start of page [bug]
    ' FIXED - #327 - Re-run BuildHeadingIndexToCSV to review changes after fixes up to page 225 [audit]
    ' FIXED - #326 - Update RunRepairWrappedVerseMarkers_Across_Pages_From to have SessionID and log [impr]
    ' FIXED - #323 - See #322 - Create index file for H1 and H2 as csv text for speedy lookup [feat] [perf]
    ' FIXED - #325 - Add md for Efficient Book-Chapter Navigation with Pre-Indexed Lookup Table [doc]
    ' FIXED - #195 - Improve verse find - Ps 119:176 is most verses, search is 14 secs, Psalm has most chapters (150), search is 2 secs
    ' FIXED - #321 - Update GoToVerseSBL to use GetParaIndexSafe and speed up verse search [impr]
    ' [obso] #320 - Add code to FindVerseFromLogicalPage [impr]
    ' [obso] #319 - Return logical page number when searching for a chapter [impr]
    ' FIXED - #318 - Add code to skip test 51. It is slow. Run again near book completion.
    ' FIXED - #317 - Use SSOT so GetPassFail is called only once per test, and results are stored in GetPassFailArray [impr]
    ' FIXED - #316 - Uses SSOT in GetPassFail to remove code duplication [impr]
    ' FIXED - #289 See #318 - Add test for Count of H2 with style [test]
    ' FIXED - #290 - Add test for Count of H1 with style [test]
    ' FIXED - #315 - Add code to make CountAndCreateDefinitionForH2 responsive
    ' FIXED - #280 - Add test to Count H2, "How many Chapters are in the Bible", Copilot -> 1,189
    ' FIXED - #313 See #280 - Update routine name and definition for H2 to include Count
    ' FIXED - #311 See #312 - Use SSOT for TestReportFlag check in RunTest [impr]
    ' FIXED - #312 - Total time of TestReport to go in csv file in rpt with Session ID
    ' FIXED - #295 - See #309 - Verify use of late binding in all code base so there is no need to set references [code]
    ' FIXED - #310 - Add code to locally auto tag a version release and push it to GitHub
    ' FIXED - #294 - Cut a 0.1.1 release and tag it on GitHub [doc] [cp]
    ' FIXED - #309 - Add code to scan modules in .docm to flag early-bound object declarations [code]
    ' FIXED - #303 - Fix single RUN_THE_TESTS(x) so it does not run AppendToFile and kill the full report [bug]
    ' FIXED - #305 - Check header writing standard, vbnet vs. vba, use style ' ============== [cp] [clean]
    ' [obso] #100 - Continue check multipage view from 300 for orphans of H2
    ' FIXED - #308 - Update all use of TestReportFlag to -> If TestReportFlag And OneTest = 0 [bug]
    ' FIXED - #307 - Remove bGoTo16, not needed with use of run single test [obso] [clean]
    ' FIXED - #306 - Add audit log from squash #274 [doc] [audit]
    ' FIXED - #279 - Add routine to define H2 style and reapply it in the project, add code header [impr] [code]
    ' FIXED - #300 - Add md doc to outline a Compact Strategy for Squashed Audit Commits and reduce GitHub commit log spam
    ' FIXED - #304 - Add task type [wip] - it will prepend the task commits until replaced by FIXED
    ' FIXED - #302 - Update PrintCompactSectionLayoutInfo to output in rpt folder, move to basTESTaeBibleTools and add doc header [doc]
    ' FIXED - #301 - 999 AppendToFile should be "SKIPPED" [bug]
    ' FIXED - #274 - Fix output path so 'Style Usage Distribution.txt' goes to rpt folder, add code header [bug] [doc]
    ' FIXED - #299 - Add initial README and Bias Guard md files [doc] [cp]
    ' FIXED - #298 - Use SSOT with Select Case statements for values such as num and verify with RUN_THE_TESTS [impr]
    ' FIXED - #297 - Create file to hold Audits for Commit Log [feat]
    ' FIXED - #296 - Add code for ValidateTaskInChangelogModule [code]
    ' FIXED - #286 - Update Heading 2 with DisableKeepLinesTogetherForHeading2 [doc]
    ' FIXED - #285 - Update Heading 2 with EnforceHeading2WidowOrphanControl [doc]
    ' FIXED - #284 - Update Heading 2 KeepWithNext [audit]
    ' FIXED - #283 - Add code GetHeadingDefinitionsWithDescriptions to tools [audit]
    ' FIXED - #282 - Update guide with 'Example of Tags use for Audit Clarity' [doc]
    ' FIXED - #278 - Use Single Source of Truth (SSOT) to fix multiple locations of array definition via MaxTests - see #273 [impr]
    ' FIXED - #277 - Define standard for types of "Tasks" to use with git commit messages [doc]
    ' FIXED - #273 - New error: Erl = 0 Error = 9 (Subscript out of range) in procedure RunBibleClassTests of Class BibleClass [bug]
    ' FIXED - #276 - git mv TestReport to rpt/ and delete old version [impr]
    ' FIXED - #275 - Create md folder for docs - md format, target github.io in future, git mv "Editorial Design and Style Guide.md" [doc]
    ' FIXED - #272 - Add section on Architecture Overview: DOCM-Coupled Macro System [doc]
    ' FIXED - #269 - All reports to be output to rpt folder [feat]
    ' FIXED - #265 - Add SKIP option to RUN_THE_TESTS for slow tests. Return -1 in report log, and GetPassFail return SKIP!!!! [feat]
    ' FIXED - #270 - Add test for SummarizeHeaderFooterAuditToFile [test]
    ' FIXED - #264 - Add test for Style Usage Distribution [test]
    ' FIXED - #263 - Add CountAuditStyles_ToFile [test]
    ' FIXED - #262 - Update code module names to match EDSG manifest [doc] [impr]
    ' FIXED - #261 - Add initial Editorial Design and Style Guide [doc]
    ' FIXED - #257 - Update SmartPrefixRepairOnPage to give a Count of Ascii 160 chars and any other e.g. hair space [impr]
    ' FIXED - #260 - Update RepairWrappedVerseMarkers_MergedPrefix_ByColumnContext_SinglePage to give a Count of Ascii 12 chars [impr]
    ' FIXED - #258 - Add RunRepairWrappedVerseMarkers_Across_Pages_From to allow per page testing [impr]
    ' [obso] [regr] - #256 - Update SmartPrefixRepairOnPage to give a Count of Ascii 12 chars
    ' FIXED - #255 - Update SmartPrefixRepairOnPage for details on unprintable characters [impr]
    ' FIXED - #254 - Add code for FindInvisibleFormFeeds_InPages [code]
    ' FIXED - #253 - Add code for LogExpandedMarkerContext [code]
    ' FIXED - #252 - Add code SmartPrefixRepairOnPage with Diagnostic Counter [code]
    ' FIXED - #251 - Add header to csv forecast output file [feat]
    ' FIXED - #250 - Wire up dummy repair test with stats collection logic [impr]
    ' FIXED - #249 - Add skeleton for StartRepairTimingSession [impr]
    ' FIXED - #248 - Update repair tool for 10 pages [impr]
    ' FIXED - #174 - Add tests for Count tab para in headers and footers [test]
    ' FIXED - #088 - Add tests for Footnote Reference (in doc and footnote) to Count those that are not bold with correct style [test]
    ' FIXED - #246 - Add test for styles using Liberation Sans Narrow [test]
    ' FIXED - #245 - Add code Identify_ArialUnicodeMS_Paragraphs [code]
    ' FIXED - #244 - Unlink heading numbering, should not display Article... or Section... for H1 or H2 [bug]
    ' FIXED - #243 - Add code RedefinePictureCaptionStyle_NotoSans, step 3 of removing Lieration Sans Narrow reference [impr]
    ' FIXED - #242 - Add code RedefineFootnoteNormalStyle_NotoSans, step 2 of removing Lieration Sans Narrow reference [impr]
    ' FIXED - #241 - Add code RedefineFootnoteStyle_NotoSans, step 1 of removing Liberation Sans Narrow reference [impr]
    ' FIXED - #240 - Update all repair code and add runner for checking 5 pages at a time [impr]
    ' FIXED - #239 - Add routine ReportDigitAtCursor_Diagnostics_Expanded [code]
    ' FIXED - #238 - Update Chapter Verse marker repair tool with latest RepairWrappedVerseMarkerPrefixes_AdjacencyWithContext_Navigate
    ' FIXED - #237 - Add diagnostic code to get character information around the cursor position [code]
    ' FIXED - #236 - Add routine to report Report Page Layout Metrics for a particular page [feat]
    ' FIXED - #235 - Add code to repair "Chapter Verse marker" per page, add vbCr if on column edge with space before, defrag as needed [impr]
    ' FIXED - #234 - Add test to Count footers that have only a tab character [test]
    ' FIXED - #212 - Add test for CountFindNotEmphasisBlack = 0 and CountFindNotEmphasisRed = 0 when all have been set [test]
    ' FIXED - #233 - Add test for CountParagraphMarks_CalibriDarkRed [test]
    ' FIXED - #232 - Add word version into to output and test report [doc]
' 20250719 - v010
    ' FIXED - #148 - Add version info to TestReport output
    ' FIXED - #231 - Reapply explicit formatting (Segoe UI 8, Bold, Blue, Superscript) for Footnote Reference, Fix for #230
    ' [obso] #230 - Add code to fix Footnote Reference by reapplying style
    ' FIXED - #225 - Add code to verify Show/Hide is True when tests are run else stop with error message
    ' FIXED - #229 - Add code to verify all necessary settings of Word are enabled - basWordSettingsDiagnostic
    ' FIXED - #228 - Abort tests if Show/Hide is not set
    ' FIXED - #227 - Update CheckShowHideStatus as ActiveWindow.View.ShowAll is the only reliable source of truth
    ' FIXED - #224 - Fix error in CheckShowHideStatus to make it reliable
    ' FIXED - #223 - Add routine with two different ways to check Show/Hide status
    ' FIXED - #222 - Add routine to compare Heading 1s with Show/Hide toggled
    ' FIXED - #220 - Update DOCVARIABLEs based on results of PrintHeading1sByLogicalPage
    ' FIXED - #219 - Add routine to Count search hits with match case true
    ' FIXED - #218 - Add routine to print logical page numbers with Heading 1, in a list, for Bible book page check
    ' FIXED - #217 - Update "I am The lord" to "I am the Lord" x42
    ' FIXED - #210 - See #213 - WoJ emphasised is 9pt, use that in search then set to 8pt As Word.Style EmphsasisRed
    ' FIXED - #184 - See #211 - Add test for Footnote Text to Count those that have any bold text [test]
    ' FIXED - #215 - Add test for paragraph mark styled - Calibri 9 Dark Red - should be color Automatic [test]
    ' FIXED - #216 - Error with H1 Count of 66 vs 59 for show/hide true false
    '    Problem list = "Joshua", "2 Kings", "Nehemiah", "Habakkuk", "Haggai", "Philemon", "1 Peter"
    '        The issue wasn�t with the styles or outline levels themselves, but with invisible or corrupted inline content
    '        (probably non-printing characters or hidden formatting) hiding in those paragraphs. When one cleaned one (Joshua),
    '        it likely triggered a reflow/re-rendering in Word that corrected the others.
    '    Solution - Click at the end of the word "Joshua" and press Delete Then press Enter once.
    '        This clears any hidden/invisible content after the heading text that may prevent proper recognition.
    '        Reselect the paragraph and reapply Heading 1 style
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
    ' [impr] FIXED - #204 - Add next- book button to ribbon : Refer also to customUI14backupRWB.xml
    ' [obso] #157 - Add Word OT DOCVARIABLEs, Ctrl + F9 field brackets { }, right-click the field, select Update Field - verify
    ' FIXED - #045 - Test call to SILAS from ribbon using customUI14.xml OnHelloWorldButtonClick routine
    ' [obso] #041 - Add auto-generated TOC for maps : auto gen too slow
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
    ' FIXED - #123 - Add file TestReport.txt output additional to console Result for GitHub tracking
    ' FIXED - #046 - Update style of cv marker to be smaller than Verse marker
    ' FIXED - #082 - Fix Word paragraph style so minimal empty paragraphs are needed
    ' [obso] #039 - Replace manual TOC with auto-generated version (this is too slow)
    ' FIXED - #141 - Update UTF8bom-Template.txt with multiple language sample of "Hello, World!" ala C style, plus phonetics
' 20250420 - v008
    ' FIXED - #140 - Set version info as global variables and assign in Class_Initialize
    ' FIXED - #139 - Add UTF8bom-Template.txt with BunnyEgg emoji for Easter using :emojisense in VS Code
    ' FIXED - #133 - Store actual Result is 1 based results array for comparison without recalc
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
    ' FIXED - #125 - Add test to Count number of footers with style "Footer" [test]
    ' FIXED - #124 - Add test for Count linefeed and space linefeed in footnotes [test]
    ' FIXED - #122 - Add test for Count linefeed and space linefeed in doc [test]
    ' FIXED - #115 - Add style "TheFooters" based on "TheHeaders" and update all footer sections, use Noto Sans font
    ' FIXED - #121 - Update debug output of Expected1BasedArray for Test(x) to be 15 per line
    ' FIXED - #120 - Add test for "TheHeaders" style as there should be only one paragraph mark per section [test]
    ' FIXED - #118 - Add test for use of "Header" style, should be 0 as "TheHeaders" has to be used instead [test]
    ' FIXED - #112 - Clear all tab stops from para headers, default is 0.1", add one tab to empty headers
    ' FIXED - #117 - See #113 - Add test to Count tab followed by paragraph mark in headers [test]
    ' FIXED - #119 - See #113 - Add test to Count paragraph mark in headers that does not have a tab [test]
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
    ' FIXED - #099 - Add test to Count number of colors in Footnote Reference [test]
    ' FIXED - #098 - Add test to Count number of Footnote References [test]
    ' FIXED - #096 - Add test for Count/delete empty para before H2, related #084 [test]
    ' FIXED - #084 - Update Heading 2 style paragraph to before 12 pt and delete the previous empty para
    ' [obso] #017 - Add optional variant to aeBibleClass for indicating Copy (x) under testing
    ' FIXED - #094 - Add test to List And Count Font Colors, and print the name from a conversion function
    ' FIXED - #090 - Work through Count Spaces After Footnotes debug output and fix as appropriate, split from ch/v numbers
    ' FIXED - #016 - Add function to print pass/fail based on comparing Result with Expected
    ' FIXED - #066 - Add tests to Count paragraphs, empty, [test]
    ' FIXED - See #073 - #067 - Add test to Count Red Footnote References
    ' FIXED - See #091 - #053 - Add test for Footnote Reference followed by a space
    ' FIXED - #089 - Continue find of footnote followed by space ("^f ") from 500 on, and fix as appropriate
    ' FIXED - #093 - Add initial PassFail test for Result and expected
    ' FIXED - #080 - Review all footnote references so that, as much as possible, they are at the end of a paragraph
' 20250402 - v006
    ' FIXED - #091 - Add test for CountSpacesAfterFootnotes - also shows Footnote References and Following Chars (ASCII Val) [test]
    ' FIXED - #092 - Add test for CountFootnotesFollowedByDigit [test]
    ' FIXED - #073 - Run test to verify Count of red footnote reference is zero [test]
    ' FIXED - #072 - Check red footnote reference from Genesis to end of Study Bible
    ' FIXED - #071 - Finish check of red footnote reference from Ezek 39 to end of Bible
    ' FIXED - #038 - Add test for no empty para after h2 headings in doc - total Count should be 0 [test]
    ' FIXED - #079 - Resolve issue around name of REV Bible - see #083
    ' FIXED - #078 - Add test to Count number of h1 heading, should be 66 for Bible books [test]
    ' FIXED - #074 - Set Heading 1 to 144 points before, follows section break so each book is on a new page with
    '           no empty first para, and delete existing 144 pt empty para
    ' FIXED - #081 - Check Books with only one chapter and verify references only use verse number per SBL abbreviations
    '           Obad, Phlm, 2 John, 3 John, Jude
    ' FIXED - #077 - Check Ezek for three in a row closing quotes
    ' FIXED - #068 - Check Ezek 1 to 26 for proper use of "'" and Ezek 39 to end of book for "'"
    '     Double quotes to indicate the start and end of the direct speech.
    '     Single quotes within the double quotes to emphasize the words spoken by God.
    '     Closing double quotes to complete the direct speech.
    '        Opening double quote:   (ASCII code: 147 or Unicode: U+201C)
    '        Closing double quote:   (ASCII code: 148 or Unicode: U+201D)
    '        Opening single quote:   (ASCII code: 145 or Unicode: U+2018)
    '        Closing single quote:   (ASCII code: 146 or Unicode: U+2019)
    '     These smart quotes are different from the straight quotes (" and ') which have ASCII codes 34 and 39, respectively.
    '     To insert these characters manually, you can use the following key combinations in Word:
    '        Opening double quote: Alt + 0147
    '        Closing double quote: Alt + 0148
    '        Opening single quote: Alt + 0145
    '        Closing single quote: Alt + 0146    ' FIXED - #064 - When bTimeAllTests is True it does not show total time
    ' FIXED - #063 - Update RunTest so it will allow a range of tests to be run (15 tests range)
    ' FIXED - #065 - Add module basTESTaeBibleTools for routines that are useful to tests outside of the class
    ' [obso] - Replaced with #062 - #036 - Add test for h1 pages that have heading
    ' FIXED - #062 - Add test for Sections With Different FirstPage selected [test]
    ' FIXED - #055 - Update RunTest so expected gets values from Expected string array
    ' FIXED - #061 - Add variant get array function of Expected to aeBibleClass and initialize with RunTest expected values
    ' FIXED - #059 - Add boolean flag to class to turn off timing for all tests
    ' FIXED - #058 - Add timer to each test and output total runtime of all tests
    ' [obso] #054 - Add string array Expected to aeBibleClass and initialize with RunTest expected values
    ' FIXED - #056 - Add test for white paragraph marks [test]
' 20250323 - v005
    ' FIXED - #052 - Add vba message box with yes/no choice to continue or stop for RunTest error
    ' FIXED - #051 - Add Yes/No continuation message to RunTest error
    ' FIXED - #050 - Error Test num = 11 Function RunTest - need to fix it [test]
    ' FIXED - #049 - Add test for Count of empty paragraphs with no theme color, wdColorAutomatic = -1 [test]
    ' FIXED - #025 (Ref #034) - Check if para is continuous break or section break next page then read the next para
    ' [obso] #027 - Create SILAS dir and add Normal.dot then extract the code to GitHub - code provided by Jim
    ' FIXED - #034 - Add routine to Count of all paragraphs types
    ' FIXED - #033 - Add Hello World custom menu tab as example for ribbon integration
    ' FIXED - #032 - Revert RunTest (12) as form feeds are needed in page and section breaks
    ' FIXED - #030 - Add routine to Count and review Form feed char positions. Needed in docx as part of page and section breaks
' 20250317 - v004
    ' FIXED - #028 - Add test to Count Hex 12 i.e. Form feed - it can cause Word not responding [test]
    ' [obso] FIXED - #026 - Add debugging code to deal with empty paragrahs in ExtractNumbersFromParagraph2
    ' FIXED - #022 - Add routine to print book h1, chapter h2, verse number - based on #021
    ' FIXED - #023 - PrintBibleHeading1Info outputs the CR of Heading 1. Remove it so output is all on one line
    ' FIXED - #021 - Add routine to print Bible book headings
    ' FIXED - #020 - Add routine to print Bible book heading details - could be used as manual page number verification
    ' FIXED - #019 - Add module for interactive slow tests not in aeBibleClass
    ' FIXED - #015 - Add test for Count number dash number in footnotes only [test]
    ' FIXED - #018 - Update Copy(???) in test runner to default Copy () as current version
    ' FIXED - #014 - Add test for Count number dash number [test]
    ' FIXED - #013 - Add test to Count number of nonbreaking spaces [test]
    ' FIXED - #012 - Add test to Count number of period space left parenthesis [test]
    ' FIXED - #011 - Add test to Count style with number and space [test]
    ' FIXED - #010 - Add copy(???) to output as placeholder for revision under test
    ' FIXED - #009 - Add test to Count style with space and number [test]
    ' FIXED - #008 - Add test to Count quadruple paragraph marks [test]
' 20250221 - v003
    ' FIXED - #007 - Add test to Count space followed by carriage return with white font color [test]
    ' FIXED - #006 - Add test to Count number of double tabs [test]
    ' FIXED - #005 - Add test to Count space followed by carriage [test]
    ' FIXED - #004 - Add tests to Count double spaces in doc, and in shapes including groups [test]
    ' FIXED - #003 - Change module name to basTESTaeBibleClass
' 20250219 - v002
    ' FIXED - #002 - Update class name to aeBibleClass
' 20250217 - v001
    ' FIXED - #001 - Create Bible Class base template, initial test module, and change log

