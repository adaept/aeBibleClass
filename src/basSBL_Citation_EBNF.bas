Attribute VB_Name = "basSBL_Citation_EBNF"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

'=======================================
' SBL Scripture Citation - Unified EBNF
'=======================================
' Citation
'    ::= WS? Reference (WS? RefSep WS? Reference)* WS?
' Reference
'    ::= BookRef (WS ChapterSpec)?
' BookRef
'    ::= Prefix? WS? BookName
' Prefix
'    ::= ArabicPrefix | RomanPrefix
' ArabicPrefix
'    ::= "1" | "2" | "3"
' RomanPrefix
'    ::= "I" | "II" | "III"
' NOTE: Prefix may be adjacent to BookWord (e.g., "1John", "IJohn")
' BookName
'    ::= BookWord (WS BookWord)*
' BookWord
'    ::= Letter+ ("." )?
' NOTE: BOOK_WORD may include a trailing . but never internal punctuation
' ChapterSpec
'    ::= Chapter
'     | Chapter ":" VerseSpec
'     | ChapterRange
'     | Chapter ":" VerseRangeSpec
' ChapterRange
'    ::= Chapter "-" Chapter
' VerseSpec
'    ::= VerseItem ("," VerseItem)*
' VerseRangeSpec
'    ::= VerseRange ("," VerseRange)*
' VerseItem
'    ::= Verse | VerseRange
' VerseRange
'    ::= Verse "-" Verse
' Verse
'    ::= Digit+ VerseSuffix?
' NOTE: VerseSuffix letters (e.g., "a", "b") are captured
'       during tokenization and validated in post-processing
' VerseSuffix
'    ::= Letter
' Chapter
'    ::= Digit+
' RefSep
'    ::= ";" | ","
' WS
'    ::= " " { " " }
' Letter
'    ::= "A"..."Z" | "a"..."z"
' Digit
'    ::= "0"..."9"
' NOTE: This DFA validates structural syntax only.
'       Semantic constraints are enforced post-parse.
'=====================================================
' Embedded Extension Hooks (Implicit but Intentional)
'=====================================================
' The grammar is designed to allow future expansion without structural change:
' Single-chapter books > semantic rewrite (Jude 5 ? Jude 1:5)
' Abbreviations / aliases > BookWord resolution table
' Verse lists & ranges > already supported
' Multiple references > ; and ,
' Roman numeral normalization > Prefix
' Language variants > alternate BookName lexemes
' Pericope titles / version tags > append after Reference
'=====================================================
' Canonical Normal Form (Post-Parse Contract)
'=====================================================
' <BookName> <Chapter>:<VerseSpec>   (lists and ranges preserved)

'=====================================================
' 1. Token Stream Definition
' 1.1 Token Types
' | Token           | Description                            | Examples          |
' | --------------- | -------------------------------------- | ----------------- |
' | BOOK_WORD       | Alphabetic word, optional trailing `.` | Genesis, Gen.     |
' | PREFIX_ARABIC   | Arabic numeric prefix                  | 1, 2, 3           |
' | PREFIX_ROMAN    | Roman numeral prefix                   | I, II, III        |
' | DIGITS          | One or more digits                     | 1, 23, 150        |
' | COLON           | Chapter-verse separator                | :                 |
' | DASH            | Range separator                        | -                 |
' | COMMA           | List separator                         | ,                 |
' | SEMICOLON       | Reference separator                    | ;                 |
' | WS              | One or more spaces (collapsed)         | " "               |
' | EOF             | displayed as <END> in debug output     |                   |

'=====================================================
' 1.2 Tokenization Rules (Critical)
' Collapse whitespace ? emit a single WS
' Case-insensitive for BOOK_WORD, PREFIX_ROMAN
' BOOK_WORD may include a trailing . but never internal punctuation
' DIGITS is greedy
' : - , ; are single-character tokens
' Whitespace is significant only between book and chapter

'=====================================================
' 1.3 Example Token Streams
' Input:
' I Cor. 13:1-3,5; Rom 8:1
' Tokens:
' PREFIX_ROMAN ("I")
' WS
' BOOK_WORD ("Cor.")
' WS
' DIGITS ("13")
' COLON
' DIGITS ("1")
' DASH
' DIGITS ("3")
' COMMA
' DIGITS ("5")
' SEMICOLON
' WS
' BOOK_WORD ("Rom")
' WS
' DIGITS ("8")
' COLON
' DIGITS ("1")
' EOF

'=====================================================
' 2. Deterministic State Machine
' This is a single-pass, left-to-right DFA.
' 2.1 State Definitions
' | State | Meaning                | Accepting |
' | ----- | ---------------------- | --------- |
' | S0    | Start                  | X         |
' | S1    | Reading numeric prefix | X         |
' | S2    | Reading book name      | X         |
' | S3    | Expecting chapter      | X         |
' | S4    | Reading chapter        | ^         |
' | S6    | Reading verse          | ^         |
' | S7    | After dash (range)     | X         |
' | S8    | After comma (list)     | X         |
' | SX    | Error                  | X         |

'=====================================================
' 2.2 State Transition Table
' Legend:
' >Sx   transition to state Sx
' X     non-accepting
' ^     conditionally accepting (see Acceptance Rules)
' SX    error state (terminal)
'----------------------------------------------
' Acceptance Rules(EXPLICIT)
' A state marked ^ is accepting only if the next token is:
' <END> => ACCEPT (end of citation)
' SEMICOLON => >S0 (start next reference)
' Any other token from an accepting state => SX.
' NOTE: Transitions to >ACCEPT and >S0 are shown explicitly
' for readability; acceptance is governed by the rules above.
'----------------------------------------------
' S0 - Start:
' | Token           | Action |
' | --------------- | ------ |
' | WS              | >S0    |
' | PREFIX_ARABIC   | >S1    |
' | PREFIX_ROMAN    | >S1    |
' | BOOK_WORD       | >S2    |
' | otherwise       | >SX    |
' S1 - Prefix:
' | Token       | Action |
' | ----------- | ------ |
' | WS          | >S2    |
' | BOOK_WORD   | >S2    |
' | otherwise   | >SX    |
' S2 - Book Name:
' | Token       | Action |
' | ----------- | ------ |
' | BOOK_WORD   | >S2    |
' | WS          | >S3    |
' | otherwise   | >SX    |
' S3 - Expect Chapter:
' | Token     | Action |
' | --------- | ------ |
' | DIGITS    | >S4    |
' | otherwise | >SX    |
' S4 - Chapter(^):
' | Token       | Action   |
' | ----------- | -------- |
' | DIGITS      | >S4      |
' | COLON       | >S6      |
' | DASH        | >S7      |
' | <END>       | >ACCEPT  |
' | SEMICOLON   | >S0      |
' | otherwise   | >SX      |
' S6 - Verse(^):
' | Token       | Action   |
' | ----------- | -------- |
' | DIGITS      | >S6      |
' | DASH        | >S7      |
' | COMMA       | >S8      |
' | <END>       | >ACCEPT  |
' | SEMICOLON   | >S0      |
' | otherwise   | >SX      |
' S7 - After Dash (Range)
' | Token     | Action |
' | --------- | ------ |
' | DIGITS    | >S6    |
' | otherwise | >SX    |
' S8 - After Comma (List)
' | Token     | Action |
' | --------- | ------ |
' | DIGITS    | >S6    |
' | otherwise | >SX    |
' SX - Error
' | Token | Action |
' | ----- | ------ |
' | any   | >SX    |
' NOTE: <END> represents the EOF token in debug output

'=====================================================
' 3. Semantic Post-Processing
' (Outside Deterministic Finite Automaton (DFA))
' Handled after a successful parse:
' Normalize prefixes > I > 1
' Collapse whitespace > single space
' Validate book name via SBL alias table
' Resolve single-chapter books
' Enforce chapter/verse bounds
' Normalize output (Book Chapter:VerseSpec)

