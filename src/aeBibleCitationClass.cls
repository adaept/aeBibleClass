VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aeBibleCitationClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'Copyright (c) 2025-2026 Peter F. Ennis
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation;
'version 3.0, or (at your option) any later version.
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'Lesser General Public License for more details.
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, visit
'https://www.gnu.org/licenses/lgpl-3.0.txt

' ================================================================
' Author:   Peter F. Ennis
' Date:     March 10, 2025
' Comment:  Bible Citation Parser for Word 365 VBA
'           - This is a Deterministic Structural Parser (DSP)
' History:  See comment details, basChangeLog_aeBibleClass, commit messages on GitHub
' GitHub:   https://github.com/adaept/aeBibleClass
' ================================================================
' The parser correctly handles three expansion classes:
'   Book-only expansion
'   Single-chapter book normalization
'   Verse shorthand expansion
' Those three transformations together form the complete
' canonicalization layer of the reference parser.
'
'========================================================
' PARSER PIPELINE
'========================================================
' Stage 1  LexicalScan
' Stage 2  ResolveAlias
' Stage 3  InterpretStructure
' Stage 4  ValidateCanonical
' Stage 5  RewriteSingleChapterRef
' Stage 6  ComposeCanonicalReference
' Stage 7  FinalParser
' Stages 8-12 implement list and range extensions.
'
' ================================================================
'  SCRIPTURE REFERENCE PARSER - FORMAL CONTRACT
' ================================================================
' 1. ParseReference(rawInput As String) is a total function:
'       It returns exactly one ScriptureRef for every input string.
' 2. The parser never raises user-facing runtime errors.
' 3. All user input errors are reported via ScriptureRef:
'       IsValid=False, ErrorCode<>0, ErrorText<>"".
' 4. If IsValid=True then:
'       BookID => [1..66],
'       Chapter >= 1,
'       Canonical output always includes a verse number.
'           Verse >= 0
'           Verse = 0 means the user did not specify a verse.
'           The canonicalization layer expands this to verse 1.
'           Examples:
'           Input           Internal State         Canonical Output
'           -------------------------------------------------------
'           Romans 8        Chapter=8  Verse=0     Romans 8:1
'           John            Chapter=1  Verse=1     John 1:1
'           Jude 5          Chapter=0  Verse=5     Jude 1:5
'       ErrorCode=0,
'       ErrorText="",
'       NormalizedRef<>"".
' 5. If IsValid=False then:
'       NormalizedRef="", ErrorCode<>0.
' 6. Raw input is not stored in the returned ScriptureRef.
'    Only normalized canonical output is preserved.
' 7. Each stage (1-7) performs a single responsibility
'    and does not mutate external state.
' 8. No stage performs responsibilities assigned to another stage.
' 9. Canonical metadata is authoritative and immutable.
'    - Canonical Book ID Mapping
'       1  Genesis
'       2  Exodus
'       3  Leviticus
'       4  Numbers
'       5  Deuteronomy
'       6  Joshua
'       7  Judges
'       8  Ruth
'       9  1 Samuel
'      10  2 Samuel
'      11  1 Kings
'      12  2 Kings
'      13  1 Chronicles
'      14  2 Chronicles
'      15  Ezra
'      16  Nehemiah
'      17  Esther
'      18  Job
'      19  Psalms
'      20  Proverbs
'      21  Ecclesiastes
'      22  Solomon
'      23  Isaiah
'      24  Jeremiah
'      25  Lamentations
'      26  Ezekiel
'      27  Daniel
'      28  Hosea
'      29  Joel
'      30  Amos
'      31  Obadiah
'      32  Jonah
'      33  Micah
'      34  Nahum
'      35  Habakkuk
'      36  Zephaniah
'      37  Haggai
'      38  Zechariah
'      39  Malachi
'      40  Matthew
'      41  Mark
'      42  Luke
'      43  John
'      44  Acts
'      45  Romans
'      46  1 Corinthians
'      47  2 Corinthians
'      48  Galatians
'      49  Ephesians
'      50  Philippians
'      51  Colossians
'      52  1 Thessalonians
'      53  2 Thessalonians
'      54  1 Timothy
'      55  2 Timothy
'      56  Titus
'      57  Philemon
'      58  Hebrews
'      59  James
'      60  1 Peter
'      61  2 Peter
'      62  1 John
'      63  2 John
'      64  3 John
'      65  Jude
'      66  Revelation
' 10. Validation is deterministic: first failure wins.
' 11. Formatting is canonical SBL-style and alias-independent.
' 12. Only internal invariant violations may raise the Err.Raise error.
' 13. Scripture structure is defined entirely by metadata.
'     The parser contains no hard-coded knowledge of Bible structure.
'     All information about books, chapters, and verse counts is stored
'     in metadata tables.
'     Examples of metadata:
'        - Canonical book list (BookID ? BookName)
'        - Number of chapters per book
'        - Maximum verses per chapter
'     The parser reads this metadata when validating references.
'     Therefore validation rules such as:
'        "Genesis has 50 chapters"
'        "Romans 16 has 27 verses"
'     are not implemented in code. They are enforced by metadata lookups.
'     Example validation flow:
'        Input:  Romans 16:30
'        Parser checks metadata:
'            GetMaxVerse(BookID=45, Chapter=16) ? 27
'        Because 30 > 27, the reference is rejected.
'     Benefits of metadata-driven structure:
'        - Parser code remains generic and simple.
'        - Structural data can be corrected without changing code.
'        - The same parser can support alternate canons or
'          translations by swapping metadata tables.
' *** Later stages are are described in the ***
' ***     Extension Layer Classification    ***
' Extension Layer Invariant
'   Stages 1-7 parse atomic references only.
'   Stages 8-12 must not modify the behavior of
'   the atomic parser.
' ================================================================
'
' Design Goals:
'
' 1. VBA coding often causes Type Mismatch and Bounds errors due to 0 vs. 1 based arrays.
'    Bible books in the Protestant Canon are numbered 1-66.
'    This design will enforce 1-Based arrays throughout the system and use Assert statements
'    to raise errors immediately for any 0-Based array usage.
'    It makes sure that the code and documentation are in synch, so strict task separation is used.
'    The test harness should follow the same principle, and documentation be updated to include
'    code routine names as part of the development strategy.
' 1a. - We are building:
'       - A disciplined, testable VBA module(s)
'       - With documentation synchronized to code
'       - With stage-specific harness testing
'       - With architectural clarity as a primary goal
'       - For a validation-heavy parser, stage isolation is gold.
'     - That creates:
'       - A formal trace from documentation => procedure
'       - A test harness entry point per stage
'       - A maintenance roadmap.
'
' 2. Parser Workflow (Pipeline Overview)
'    The reference engine follows a strict multi-stage pipeline.
'    Each stage has a single responsibility.
'    No stage may perform work assigned to another stage.
'   -------------------------------------
'   DETERMINISTIC STRUCTURAL PARSER (DSP)
'   -------------------------------------
'   Stage 1   Normalize
'   Stage 2   LexicalScan
'   Stage 3   ResolveAlias
'   Stage 4   InterpretStructure
'   Stage 5   ValidateCanonical
'   Stage 6   FormatCanonical
'   Stage 7   EmitResult
'
'   EXTENSION PARSING LAYER
'   -----------------------
'   Stage 8   ListDetection
'   Stage 9   RangeDetection
'   Stage 10  RangeComposition
'   Stage 11  ListComposition
'   Stage 12  ExtendedParse
'
'   CONTEXT RESOLUTION LAYER
'   ------------------------
'   Stage 13  ContextualShorthand
'
'   CANONICAL OUTPUT LAYER
'   ----------------------
'   Stage 14  CanonicalCompression
'
' The atomic parser is a mathematically deterministic function:
'   ParseReference : String -> ScriptureRef
' The extended parser becomes:
'   ParseReferenceExtended : String -> ScriptureRef | ScriptureRange | ScriptureList
'
'================================================================
' Design Principle
'================================================================
'       Later stages may depend on earlier stages.
'       Earlier stages must never depend on later stages.
'    Invariant:
'       No user-facing errors are raised during parsing.
'       Only internal invariant violations may Err.Raise.
'
'    ------------------------------------------------------------
'    Stage 1: Input Normalization (Normalize)
'    ------------------------------------------------------------
'    PURPOSE:
'       Normalize raw input string into a canonical internal
'       working form suitable for tokenization.
'    INPUT:
'       rawInput As String
'    OUTPUT:
'       normalizedInput As String
'    RESPONSIBILITIES:
'       - Trim leading/trailing whitespace
'       - Collapse multiple internal spaces to single space
'       - Normalize Unicode punctuation (e.g., smart quotes, en-dash)
'    NON-RESPONSIBILITIES:
'       - No alias resolution
'       - No numeric parsing
'       - No structural interpretation
'       - No validation
'       - No logging or diagnostic preservation
'    DESIGN DECISION:
'       Raw input is NOT preserved.
'       The parser is intentionally stateless and does not
'       retain original user input.
'    FAILURE MODEL:
'       - Never raises user-facing errors.
'       - Returns normalized string (possibly empty).
'    CURRENT IMPLEMENTATION STATUS:
'       - Trim whitespace Y
'       - Preserve original input N (intentionally omitted)
'       - Collapse multiple spaces N (pending)
'       - Unicode normalization N (pending)
'    NOTE:
'       Currently partially implemented inline in ParseReferenceStub.
'       Will be refactored to:
'           NormalizeInput(rawInput As String) As String
'
'    ------------------------------------------------------------
'    Stage 2: Lexical Tokenization (Tokenize)
'    ------------------------------------------------------------
'    PURPOSE:
'       Convert normalized input string into primitive lexical tokens.
'       This stage performs ZERO semantic interpretation.
'    INPUT:
'       normalizedInput As String
'           (Output of Stage 1: NormalizeInput)
'    CURRENT IMPLEMENTATION:
'       Public Function TokenizeReference(ByVal normalizedInput As String) As LexTokens
'    CURRENT BEHAVIOR:
'       - Splits on single space
'       - First token => RawAlias
'       - Second token (if present) => numeric reference block
'       - Detects colon separator
'       - Extracts up to two numeric substrings
'    OUTPUT STRUCTURE:
'       Type LexTokens
'           RawAlias    As String
'           Num1        As Long
'           Num2        As Long
'           HasColon    As Boolean
'           Num1IsValid As Boolean
'           Num2IsValid As Boolean
'       End Type
'    NUMERIC HANDLING RULE:
'       - Numeric conversion must NEVER raise runtime errors.
'       - Use safe parsing (e.g., IsNumeric check or guarded CLng).
'       - If conversion fails:
'             NumXIsValid = False
'             NumX = 0
'       Stage 2 does NOT determine whether values are
'       canonically valid - only whether they are numeric.
'    LIMITATIONS (INTENTIONAL - STAGE 3 WILL HANDLE):
'       - Multi-word book names (e.g., "1 John", "Song of Solomon")
'       - Range detection (1-3)
'       - List detection (1,3,5)
'       - Canonical validation of chapter/verse
'       - Alias resolution
'    DESIGN RULE:
'       Stage 2 extracts lexical structure only.
'       Stage 2 NEVER raises runtime errors for malformed input.
'       Stage 3 interprets meaning.
'    ASSUMPTION:
'         Stage 1 guarantees single-space separation.
'    FAILURE MODEL:
'       - Invalid numeric text does NOT raise error.
'       - Invalid numeric text is structurally flagged.
'       - Structural errors are interpreted in Stage 3.
'    TEST HARNESS:
'       Public Sub Test_TokenizeReference()
'    STATUS:
'       Y Book token extraction
'       Y Numeric extraction (single or colon pair)
'       Y Colon detection
'       Y Safe numeric parsing (required)
'       N Range/list parsing (future)
'       N Multi-token book support (future)
'
'    ------------------------------------------------------------
'    Stage 3: Alias Resolution (Resolve Book Identity)
'    ------------------------------------------------------------
'    PURPOSE:
'       Resolve RawAlias into canonical BookID.
'       Preserve lexical numeric tokens without interpreting them.
'    INPUT:
'       LexTokens
'           RawAlias    As String
'           Num1        As Long
'           Num2        As Long
'           HasColon    As Boolean
'           Num1IsValid As Boolean
'           Num2IsValid As Boolean
'    OUTPUT:
'       Type ParsedRef
'           BookID      As Long   ' 0 if alias unresolved
'           AliasFound  As Boolean
'           ' Forwarded lexical state (UNINTERPRETED)
'           Num1        As Long
'           Num2        As Long
'           HasColon    As Boolean
'           Num1IsValid As Boolean
'           Num2IsValid As Boolean
'           ' Structural fields NOT assigned in Stage 3
'           Chapter     As Long   ' Uninitialized here
'           Verse       As Long   ' Uninitialized here
'       End Type
'    RESPONSIBILITIES:
'       1. Normalize RawAlias (case-insensitive lookup).
'       2. Lookup alias in canonical alias dictionary.
'       3. If found:
'              BookID = canonical ID (1-66)
'              AliasFound = True
'          If NOT found:
'              BookID = 0
'              AliasFound = False
'       4. Forward Num1/Num2/HasColon and validity flags unchanged.
'    NON-RESPONSIBILITIES:
'       - Does NOT assign Chapter or Verse.
'       - No structural interpretation.
'       - No chapter/verse validation.
'       - No canonical bounds checking.
'       - No formatting.
'    FAILURE PROPAGATION RULE:
'       Stage 3 NEVER raises runtime errors.
'       Alias resolution failure does NOT terminate parsing.
'       Instead:
'           AliasFound = False
'           BookID     = 0
'       Lexical numeric failures are preserved and forwarded.
'    DESIGN INVARIANT:
'       BookID = 0 is the ONLY legal representation
'       of unresolved book identity.
'       No negative BookID values.
'    DESIGN RULE:
'       Stage 3 establishes canonical identity ONLY.
'       Structural meaning begins in Stage 4.
'
'    -----------------------------------------------------------------
'    Stage 4: Structural Interpretation (Determine structural meaning)
'    -----------------------------------------------------------------
'    PURPOSE:
'       Convert lexical numeric tokens into structural Chapter/Verse
'       values using canonical metadata (BookID).
'    INPUT:
'       ParsedRef from Stage 3:
'           BookID
'           AliasFound
'           Num1
'           Num2
'           Num1IsValid
'           Num2IsValid
'           HasColon
'    OUTPUT (updates ParsedRef):
'           Chapter As Long
'           Verse   As Long
'    RESPONSIBILITY:
'       Stage 4 is the SOLE owner of structural interpretation.
'       It assigns Chapter and Verse exactly once.
'    PRECONDITION:
'       None. Must tolerate:
'           - BookID = 0
'           - Invalid numeric tokens
'    LOGIC:
'       ' Default structural assignment (no metadata reliance)
'           Chapter = Num1
'           Verse   = 0
'       If Num1IsValid = False And HasColon = False Then
'       ' Book-only reference
'           Chapter = 1
'           Verse = 1
'       ' If colon present, explicit Chapter:Verse structure
'       If HasColon = True Then
'           Verse = Num2
'           Exit Stage
'       End If
'       ' No colon present
'       ' Metadata may be consulted ONLY if alias resolved
'       If AliasFound = True And Num1IsValid = True Then
'           If GetChapterCount(BookID) = 1 Then
'               ' Single-chapter book (e.g., Jude)
'               Chapter = 1
'               Verse   = Num1
'           End If
'       End If
'    STRUCTURAL RULES:
'       - Stage 4 NEVER changes Num1 or Num2.
'       - Stage 4 NEVER converts invalid numeric tokens into valid ones.
'       - Stage 4 NEVER performs range checking.
'       - Stage 4 assigns Chapter/Verse exactly once.
'    STRUCTURAL ASSERTIONS (not semantic validation):
'       Chapter >= 0
'       Verse   >= 0
'    IMPORTANT:
'       Stage 4 determines structure only.
'       Stage 5 validates:
'           - AliasFound = True
'           - Num1IsValid / Num2IsValid
'           - Chapter >= 1
'           - Chapter <= MaxChapter(BookID)
'           - Verse bounds
'    DESIGN RULE:
'       Structural interpretation must not depend on
'       chapter/verse bounds.
'       Metadata (GetChapterCount) may be consulted ONLY
'       when BookID is valid.
'    =====================================================
'    Book-Only Reference Handling
'    =====================================================
'    If a reference contains only a book alias and no
'    numeric component, the parser defaults to:
'       Chapter = 1
'       Verse   = 1
'    Example:
'       Input:  "John"
'       Output: "John 1:1"
'       Input:  "1 Jn"
'       Output: "1 John 1:1"
'    Rationale:
'       This behavior matches common Bible software
'       navigation conventions and provides a safe
'       canonical anchor for book-level references.
'    Implementation:
'       When refPart is empty, it is internally
'       replaced with "1:1" before numeric validation.
'    =====================================================
'
'    ------------------------------------------------------------
'    Stage 5: Canonical Validation (Validate reference)
'    ------------------------------------------------------------
'    PURPOSE:
'       Determine whether the structurally interpreted reference
'       is canonically valid.
'    INPUT:
'       ParsedRef from Stage 4:
'           BookID      As Long
'           AliasFound  As Boolean
'           Chapter     As Long
'           Verse       As Long
'           HasColon    As Boolean
'           Num1IsValid As Boolean
'           Num2IsValid As Boolean
'    OUTPUT:
'       IsValid   As Boolean
'       ErrorCode As Long
'       ErrorText As String
'    RESPONSIBILITY:
'       - Validate canonical correctness only.
'       - Do NOT modify Chapter or Verse.
'       - Do NOT format output.
'       - Do NOT raise runtime errors.
'    VALIDATION ORDER (STRICT PRECEDENCE):
'       1. Alias resolution
'       2. Lexical numeric validity
'       3. Structural minimums (>=1)
'       4. Chapter upper bound
'       5. Verse upper bound
'    FIRST FAILURE WINS.
'
'    ------------------------------------------------------------
'    Stage 6: Canonical Normalization (Output Formatting)
'    ------------------------------------------------------------
'    PURPOSE:
'       Produce canonical SBL-style reference string including
'       canonical Book Name.
'    INPUT:
'       BookID  As Long
'       Chapter As Long
'       Verse   As Long
'    PRECONDITION:
'       Stage 5 validation has succeeded.
'       Therefore the following invariants hold:
'           BookID  => [1..66]
'           Chapter >= 1
'           Chapter <= MaxChapter(BookID)
'           If Verse = 0:
'               Reference is chapter-only
'           If Verse > 0:
'               Verse = MaxVerse(BookID, Chapter)
'    OUTPUT:
'       NormalizedRef As String
'    LOGIC:
'       CanonicalBookName = GetCanonicalBookName(BookID)
'       If Verse = 0 Then
'           ' Chapter-only reference
'           NormalizedRef = CanonicalBookName & " " & _
'                           CStr(Chapter)
'       Else
'           ' Chapter + Verse reference
'           NormalizedRef = CanonicalBookName & " " & _
'                           CStr(Chapter) & ":" & CStr(Verse)
'       End If
'    FORMAT RULES:
'       - Exactly one space between BookName and Chapter.
'       - No trailing spaces.
'       - No zero-padding.
'       - No leading zeros.
'       - No alias text allowed.
'    RESTRICTIONS:
'       - Do NOT perform validation here.
'       - Do NOT modify structural values.
'       - Do NOT consult alias dictionary.
'       - Do NOT emit output if Stage 5 failed.
'    DESIGN RULE:
'       Stage 6 is a pure formatting function.
'       Given valid canonical input,
'       output is uniquely determined.
'
'    ------------------------------------------------------------
'    Stage 7: Structured Result Object (Emit Immutable Result)
'    ------------------------------------------------------------
'    PURPOSE:
'       Construct and return the final immutable ScriptureRef
'           result object.
'       ParseReference() must always return a ScriptureRef,
'           never Nothing, never uninitialized.
'    INPUT:
'       Final state from Stage 5 and (if valid) Stage 6.
'    OUTPUT:
'       Public Type ScriptureRef
'           BookID        As Long
'           Chapter       As Long
'           Verse         As Long
'           NormalizedRef As String
'           IsValid       As Boolean
'           ErrorCode     As Long
'           ErrorText     As String
'       End Type
'    CONSTRUCTION RULE:
'       ScriptureRef is constructed exactly once
'       inside ParseReference().
'       No other procedure may partially construct it.
'    STATE INVARIANTS:
'       If IsValid = True Then
'           BookID        => [1..66]
'           Chapter       >= 1
'           Verse         >= 0
'           ErrorCode     = 0
'           ErrorText     = ""
'           NormalizedRef <> ""
'       If IsValid = False Then
'           NormalizedRef = ""
'           ErrorCode     <> 0
'           ErrorText     <> ""
'           BookID may be 0
'           Chapter and Verse may be 0
'    ILLEGAL STATES (MUST NEVER OCCUR):
'       IsValid = True  AND ErrorCode <> 0
'       IsValid = True  AND NormalizedRef = ""
'       IsValid = False AND ErrorCode = 0
'    FAILURE MODEL:
'       - Parser NEVER raises user-facing runtime errors.
'       - All user input errors are reported via:
'             IsValid   = False
'             ErrorCode <> 0
'             ErrorText <> ""
'       - Only internal invariant violations
'         (e.g., corrupted canonical metadata)
'         may raise Err.Raise.
'       - Because RawInput is discarded:
'           ErrorText must never embed raw user input.
'    DESIGN RULE:
'       Stage 7 performs:
'           - No parsing
'           - No validation
'           - No metadata lookup
'       It only packages the final canonical state.
'       ScriptureRef is immutable after return.
'

'=====================================================
' SBL Scripture Citation - Structural EBNF
' Aligned to 7-Stage Deterministic Parser
'=====================================================
' PURPOSE:
'   Defines structural syntax only.
'   No semantic validation is expressed here.
'   Canonical bounds enforcement occurs in Stage 5.
' NOTE:
'   This grammar describes lexical and structural form.
'   It does NOT:
'       - Validate book identity
'       - Validate chapter/verse bounds
'       - Normalize aliases
'       - Enforce canonical metadata
'=====================================================
' Top-Level
'=====================================================
' Citation
'    ::= WS? Reference (WS? RefSep WS? Reference)* WS?
' RefSep
'    ::= ";" | ","
'=====================================================
' Core Reference
'=====================================================
' Reference
'    ::= BookRef
'     |  BookRef WS ChapterSpec
' BookRef
'    ::= Prefix? WS? BookName
'-----------------------------------------------------
' Book-only references are permitted.
'    When no ChapterSpec is present, the reference is
'    normalized to the first verse of the book.
' Examples
'    John      -> John 1:1
'    Romans    -> Romans 1:1
'    Jude      -> Jude 1:1
'    This normalization occurs in Stage 6
'    (canonical formatting / rewrite)
'-----------------------------------------------------
' Prefix
'    ::= ArabicPrefix | RomanPrefix
' ArabicPrefix
'    ::= "1" | "2" | "3"
' RomanPrefix
'    ::= "I" | "II" | "III"
' NOTE:
'   Prefix may be adjacent to BookName
'   (e.g., "1John", "IJohn")
' BookName
'    ::= BookWord (WS BookWord)*
' BookWord
'    ::= Letter+ "."?
' CONSTRAINT (Structural Only):
'   - No internal punctuation.
'   - Trailing period permitted.
'=====================================================
' Chapter / Verse Structure
'=====================================================
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
'=====================================================
' Atomic Numeric Units
'=====================================================
' Chapter
'    ::= Digit+
' Verse
'    ::= Digit+ VerseSuffix?
' VerseSuffix
'    ::= Letter
' VERSE_SUFFIX | Alphabetic suffix after verse digit
' NOTE:
'   VerseSuffix (e.g., "a", "b") is lexically captured.
'   Canonical validity is enforced post-parse.
'=====================================================
' Lexical Primitives
'=====================================================
' WS
'    ::= " " { " " }
' Letter
'    ::= "A"..."Z" | "a"..."z"
' Digit
'    ::= "0"..."9"
'=====================================================
' Parser Alignment Notes
'=====================================================
' Stage 2:
'   Performs tokenization according to this grammar.
' Stage 3:
'   Resolves BookRef ? BookID via alias dictionary.
' Stage 4:
'   Interprets structural meaning:
'       - Chapter-only
'       - Chapter:Verse
'       - Single-chapter inference (semantic layer)
' Stage 5:
'   Enforces semantic constraints:
'       - AliasFound = True
'       - Chapter >= 1
'       - Chapter <= MaxChapter(BookID)
'       - Verse bounds
' Stage 6:
'   Produces canonical SBL formatting.
' Stage 7:
'   Emits immutable ScriptureRef result object.
'=====================================================
' Canonical Normal Form (Post-Validation Output)
'=====================================================
' Single reference:
'     <CanonicalBookName> <Chapter>
'     <CanonicalBookName> <Chapter>:<Verse>
' Lists and ranges preserve structural form
' after semantic validation.
'=====================================================
'NOTE: This DSP validates structural syntax only.
'      Semantic correctness is enforced in Stage 5.
'=====================================================

'=====================================================
' Parser Alignment to Implementation (7-Stage Model)
'=====================================================
' Stage 1 - Preprocessing / Normalization
' -----------------------------------------------------
' Routine:
'     Private Function Stage1_NormalizeInput( _
'         ByVal rawInput As String) As String
' Responsibility:
'     - Trim leading/trailing whitespace
'     - Normalize internal whitespace
'     - Standardize dash characters if required
'     - Discard raw input after normalization
' -----------------------------------------------------
' Stage 2 - Lexical Tokenization
' -----------------------------------------------------
' Routine:
'     Private Function Stage2_LexicalScan( _
'         ByVal normalizedInput As String) As LexTokens
' Responsibility:
'     - Extract RawAlias
'     - Extract Num1, Num2
'     - Detect HasColon
'     - Detect lexical numeric validity
'     - Capture VerseSuffix (if applicable)
' -----------------------------------------------------
' Stage 3 - Alias Resolution (Canonical Identity)
' -----------------------------------------------------
' Routine:
'     Private Function Stage3_ResolveAlias( _
'         ByVal tokens As LexTokens) As ParsedRef
' Responsibility:
'     - Resolve RawAlias ? BookID
'     - Set AliasFound flag
'     - Forward lexical numeric tokens unchanged
'     - Do NOT assign Chapter/Verse
' -----------------------------------------------------
' Stage 4 - Structural Interpretation
' -----------------------------------------------------
' Routine:
'     Private Sub Stage4_InterpretStructure( _
'         ByRef state As ParsedRef)
' Responsibility:
'     - Assign Chapter and Verse exactly once
'     - Apply colon structure logic
'     - Apply single-chapter inference
'     - Do NOT validate bounds
' -----------------------------------------------------
' Stage 5 - Canonical Semantic Validation
' -----------------------------------------------------
' Routine:
'     Private Sub Stage5_ValidateCanonical( _
'         ByRef state As ParsedRef)
' Responsibility:
'     - Enforce validation matrix
'     - First-failure-wins ordering
'     - Assign ErrorCode / ErrorText
'     - Set IsValid flag
'     - Do NOT modify structural values
' -----------------------------------------------------
' Stage 6 - Canonical Normalization (Formatting)
' -----------------------------------------------------
' Routine:
'     Private Sub Stage6_FormatCanonical( _
'         ByRef state As ParsedRef)
' Responsibility:
'     - Produce canonical SBL-style string
'     - Use canonical BookName from metadata
'     - Execute only if IsValid = True
'     - Do NOT perform validation
' -----------------------------------------------------
' Stage 7 - Immutable Result Emission
' -----------------------------------------------------
' Routine:
'     Private Function Stage7_EmitResult( _
'         ByVal state As ParsedRef) As ScriptureRef
' Responsibility:
'     - Enforce object state invariants
'     - Construct final ScriptureRef
'     - Guarantee total-function return
'     - Perform no parsing or validation
' -----------------------------------------------------
' Public Entry Point
' -----------------------------------------------------
' Routine:
'     Public Function ParseReference( _
'         ByVal rawInput As String) As ScriptureRef
'
' Execution Order:
'     1. normalized  = Stage1_NormalizeInput(rawInput)
'     2. tokens      = Stage2_LexicalScan(normalized)
'     3. state       = Stage3_ResolveAlias(tokens)
'     4. Stage4_InterpretStructure state
'     5. Stage5_ValidateCanonical state
'     6. Stage6_FormatCanonical state
'     7. ParseReference = Stage7_EmitResult(state)
'=====================================================

'=====================================================
' Extension Layer Classification
'=====================================================
' The core parser (Stages 1-7) is modeled as a Deterministic
' Structural Parser (DSP) that produces a single atomic ScriptureRef.
'
' *** Extension stages operate OUTSIDE the DSP and must never ***
' ***      influence the DSP state machine.                   ***
'
' Invariant
'   Stages 1-7 must only parse atomic references.
'   Extension stages must not modify or bypass
'   core validation logic.
' These stages are strictly lexical segmentation layers.
'   Stage 8  List Detection
'   Stage 9  Range Detection
'   Stage 10 RangeComposition     -> builds ranges
'   Stage 11 ListComposition      -> builds lists
'   Stage 12 ExtendedParse        -> orchestrate pipeline
'   Stage 13 ContextualShorthand
'Recursion
'   Stage-12 may call ParseReferenceExtended() recursively
'   when processing list segments. This allows nested list
'   and range structures to be parsed without additional
'   grammar rules.
' Responsibility boundaries:
' Stage 8
'   Detect top-level list separators
'   Output ListTokens
' Stage 9
'   Detect reference ranges
'   Output RangeTokens
' Stage 10
'   Compose atomic ScriptureRef results into:
'       ScriptureList
'       ScriptureRange
'
' The DSP always processes a single atomic reference.
'=====================================================

'=====================================================
' Stage 8 - List Detection (Extension Layer)
'=====================================================
' Purpose
'   Detect multiple references separated by list delimiters.
' Supported separators
'   ,  comma
'   ;  semicolon
' Examples
'   John 3:16,18,20
'   John 3,4,5
'   John 3:16; 4:1
' Output
'   Type ListTokens
'       IsList As Boolean
'       Segments() As String
'   End Type
' Rules
'   1. Detection is lexical only.
'   2. No interpretation of structure occurs.
'   3. Segments are returned exactly as written.
'   4. If no separator is found, IsList=False.
' Determinism
'   Stage 8 MUST NOT perform:
'       - alias resolution
'       - verse validation
'       - canonical formatting
'   These remain Stage 3-6 responsibilities.
' List Detection Rule
'   Separators are recognized only at the top lexical level.
'   Stage 8 must not split inside a detected range.
'Example Input:
'   John 3:16-18,20
' Output segments:
'   John 3:16-18
'   20
' NOT:
'   John 3:16
'   18
'   20
' Stage Ordering
'   Stage 8 executes before Stage 9.
'   Lists are segmented before range interpretation occurs.
' Range interpretation is handled in Stage 9.
' Stage 9 evaluates each segment independently.
'=====================================================

'=====================================================
' Stage 9 - Range Detection (Extension Layer)
'=====================================================
' Purpose
'   Detect reference ranges using hyphen or en dash.
' Supported separators
'   -  = ASCII hyphen-minus  (ChrW(45))
'   ?  = Unicode en dash     (ChrW(&H2013))
'-----------------------------------------------------
' A Note on Range Delimiter Characters
'-----------------------------------------------------
' The parser supports two range delimiters:
' ASCII hyphen-minus
'   Character: "-"
'   Unicode:   U+002D
'   VBA:       ChrW(45)
' Unicode en dash
'   Character: "?"
'   Unicode:   U+2013
'   VBA:       ChrW(&H2013)
' NOTE
'   The VBA Immediate Window may not display the en dash
'   correctly. In Git it may appear as a placeholder.
'   Therefore code comparisons should use ChrW values.
' Example
'   If ch = "-" Or ch = ChrW(&H2013) Then
'       ' range delimiter comparison code
'   End If
'-----------------------------------------------------
' Example Input
'   John 3:16-18
'   John 3-5
'   John 3:16-4:2
' Output
'   Type RangeTokens
'       IsRange As Boolean
'       LeftRaw  As String
'       RightRaw As String
'   End Type
' Rules
'   1. Detection is lexical only.
'   2. No interpretation of structure occurs.
'   3. Left and Right expressions are returned exactly as written.
'   4. If no range delimiter exists, IsRange=False.
' Determinism
'   Stage 9 MUST NOT perform:
'       - alias resolution
'       - verse validation
'       - canonical formatting
'   These remain Stage 3-6 responsibilities.
' Stage Ordering
'   Stage 8 executes before Stage 9.
' Example
'   Input:
'       John 3:16-18,20
'   Stage 8 output:
'       John 3:16-18
'       20
'   Stage 9 then evaluates each segment independently.
' Range Detection Rule
'   A range delimiter must appear after the first numeric token.
' This prevents false detection in book prefixes such as:
'   1-2 Samuel
' Stage-9 must remain lexical only and must not interpret structure.
' Deterministic Structural Parser (DSP) intact because Stage-9
' remains outside the core parser.
' Stage 9 Evaluation:
'    RangeDetection("John 3:16-18") -> Range
'    RangeDetection("20") -> Not a range
'=====================================================

'=====================================================
' Stage 10 - RangeComposition (Extension Layer)
'=====================================================
' Purpose
'   Construct a ScriptureRange from the tokens produced
'   by Stage 9 (RangeDetection).
' Stage-10 performs no lexical parsing.
'   It only composes structured results using the atomic parser.
' Composition Type
'   ScriptureRange
' Rules
'   1. Stage-10 must call ParseReference() for both
'      sides of the detected range.
'   2. LeftRaw and RightRaw must be parsed independently.
'   3. Stage-10 must not modify ScriptureRef.
' Atomic Parser Guarantee
'   ParseReference() remains the only function that
'   produces ScriptureRef.
' Example Input
'   John 3:16-18
' Stage 9
'   RangeTokens
'       LeftRaw  = "John 3:16"
'       RightRaw = "18"
' Stage 10
'   ScriptureRange
'       StartRef -> ScriptureRef(John 3:16)
'       EndRef   -> ScriptureRef(John 3:18)
'=====================================================

'=====================================================
' Stage 11 - ListComposition (Extension Layer)
'=====================================================
' Purpose
'   Compose segmented references into a ScriptureList.
' Stage-11 operates on the segments produced by
'   Stage 8 (ListDetection).
' Composition Type
'   ScriptureList
' Rules
'   1. Each segment must be processed independently.
'   2. Stage-11 must determine whether a segment is
'      a range using Stage 9.
'   3. Range segments must be composed using Stage 10.
'   4. Non-range segments must be parsed using
'      ParseReference().
' Result
'   ScriptureList.Items() may contain
'       ScriptureRef
'       ScriptureRange
' Example Input
'   John 3:16-18,20
' Stage 8
'   Segments
'       John 3:16-18
'       20
' Stage 9
'   Segment 1 -> Range
'   Segment 2 -> Single
' Stage 11
'   ScriptureList
'       Item 1 -> ScriptureRange
'       Item 2 -> ScriptureRef
'=====================================================

'=====================================================
' Post-Parse Canonical Processing (Stages 12-14)
'=====================================================
' After atomic parsing (Stages 1-7) and structural
' composition (Stages 8-11), the extension pipeline
' performs higher-level reference interpretation.
' Stage Responsibilities:
'   Stage 12  ExtendedParse
'       Orchestrates list and range parsing
'   Stage 13  ContextualShorthand
'       Resolves omitted book/chapter context
'   Stage 14  CanonicalCompression
'       Produces minimal canonical reference form
' Only after Stage 13 are references guaranteed to
' represent fully-resolved canonical references.
' Some implementations may internally expand
' references to verse-level triples:
'       (BookID, Chapter, Verse)
' Such expansion is an internal optimization and is
' not required by the specification.
' Stage 14 operates on the resolved canonical
' reference set and performs deterministic
' structural compression.
'=====================================================

'=====================================================
' Stage 12 - ExtendedParse (Extension Entry Point)
'=====================================================
' Purpose
'   Provide a high-level parser capable of handling
'   lists and ranges.
' Stage-12 orchestrates the extension pipeline while
'   preserving the atomic parser contract.
' Pipeline
'   Stage 8  ListDetection
'   Stage 9  RangeDetection
'   Stage 10 RangeComposition
'   Stage 11 ListComposition
' Return Types
'   ScriptureRef
'   ScriptureRange
'   ScriptureList
' Atomic Parser Guarantee
'   Stages 1-7 remain responsible for parsing
'   atomic references only.
' Example Input
'   John 3:16-18,20; 4:1-3
' Result
'   ScriptureList
'       Item 1 -> ScriptureRange (3:16-3:18)
'       Item 2 -> ScriptureRef   (3:20)
'       Item 3 -> ScriptureRange (4:1-4:3)
'=====================================================

'=====================================================
' Stage 13 - Contextual Shorthand Expansion
'            (Post-Parser Context Layer)
'=====================================================
' Public API:
'     ComposeList(ByVal raw As String) As Collection
' Purpose:
'     Parses a list of scripture references from a string and returns
'     a collection of canonical reference strings.
' Supported Forms:
'   Single references
'       "John 3:16"
'   Ranges
'       Same chapter:      "John 3:16-18"
'       Cross chapter:     "John 3:16-4:2"
'   Lists
'       "John 3:16, 18; 4:1"
'   Contextual shorthand
'       Missing book    -> inherited from previous segment
'       Missing chapter -> inherited from current context
' Example:
'   "3:16-4:2, 5"
'       -> "John 3:16-4:2"
'       -> "John 4:5"
'-----------------------------------------------------
' Context State
'-----------------------------------------------------
' During processing the engine maintains:
'       CurrentBook
'       CurrentChapter
' Context is updated after each resolved segment.
'-----------------------------------------------------
' Processing Order
'-----------------------------------------------------
' Segments are processed left-to-right.
' Context updates occur after each resolved segment.
'-----------------------------------------------------
' Range Context Rules
'-----------------------------------------------------
' After a range resolves, the context chapter becomes
' the ending chapter of the range.
' Example:
'   "John 3:16-4:2, 5"
' Range resolved:
'   John 3:16-4:2
' Context becomes:
'   Book = John
'   Chapter = 4
' Next segment:
'   5 -> John 4:5
'-----------------------------------------------------
' Canonical Output Rules
'-----------------------------------------------------
' Same-chapter ranges collapse chapter repetition:
'       John 3:16-18
' Cross-chapter ranges preserve both chapters:
'       John 3:16-4:2
' Single verses always include book and chapter:
'       John 3:16
'-----------------------------------------------------
' Examples
'-----------------------------------------------------
' Input:
'       "John 3:16, 18, 20-22"
' Output:
'       "John 3:16"
'       "John 3:18"
'       "John 3:20-22"
' Input:
'       "3:16-4:2, 5"
' Output:
'       "John 3:16-4:2"
'       "John 4:5"
' Input:
'       "Romans 8; 9"
' Output:
'       "Romans 8"
'       "Romans 9"
'-----------------------------------------------------
' Implementation
'-----------------------------------------------------
'   Stage 11  -> ListTokens (segment structure)
'   Stage 12  -> ScriptureRef / ScriptureRange objects
'   Stage 13  -> Contextual shorthand resolution
' Output:
'   Collection of canonical reference strings
'=====================================================

'=====================================================
' Stage 14 - Canonical Compression
'=====================================================
' Input:
'   A resolved set of canonical references produced
'   after contextual shorthand expansion (Stage 13).
' Output:
'   Minimal canonical citation form.
' Stage 14 performs deterministic compression of the
' expanded canonical verse stream.
' Adjacent verses are collapsed into ranges while
' preserving canonical ordering and semantic meaning.
' This stage is purely structural and does not alter
' interpretation of references.
'-----------------------------------------------------
' Canonical Output Grammar
'-----------------------------------------------------
' CanonicalCitation
'    ::= CanonicalBookRef
'     | CanonicalBookRef (";" WS? CanonicalBookRef)*
' CanonicalBookRef
'    ::= BookName WS CanonicalChapterSpec
' CanonicalChapterSpec
'    ::= CanonicalChapterUnit
'     | CanonicalChapterUnit (";" WS? CanonicalChapterUnit)*
' CanonicalChapterUnit
'    ::= Chapter ":" CanonicalVerseSpec
' CanonicalVerseSpec
'    ::= CanonicalVerseItem
'     | CanonicalVerseItem ("," CanonicalVerseItem)*
' CanonicalVerseItem
'    ::= Verse
'     | Verse "-" Verse
'-----------------------------------------------------
' Canonical Compression Rules
'-----------------------------------------------------
' 1. Sequential verses collapse into ranges.
'      John 3:16
'      John 3:17
'      John 3:18
'  ->  John 3:16-18
' 2. Non-sequential verses remain comma separated.
'      John 3:16
'      John 3:18
'  ->  John 3:16,18
' 3. Chapter boundaries are never merged.
'      Genesis 1:31
'      Genesis 2:1
'  ->  Genesis 1:31; 2:1
' 4. Compression must preserve canonical ordering:
'      BookID
'      Chapter
'      Verse
' 5. Compression is deterministic and lossless.
'=====================================================

'=====================================================
' Stage 15 - Canonical Validation
'=====================================================
' Purpose:
'     Ensure the final canonical reference set contains only valid,
'     in-range scripture references after all expansion, ordering,
'     and compression operations have completed.
' Position in Pipeline:
'     Stage 15 executes AFTER Stage 14 Canonical Compression.
'     At this point, the reference set is:
'         - Fully expanded
'         - Deduplicated
'         - Ordered
'         - Canonically compressed
'     Stage 15 performs the final integrity check before output.
' Responsibilities:
'     1. Remove Invalid Chapters
'         Eliminate references to chapters that do not exist
'         in the specified book.
'         Example:
'             Gen 51      -> removed
'             Matt 29     -> removed
'     2. Remove Invalid Verses
'         Eliminate verse numbers exceeding chapter limits.
'         Example:
'             Gen 1:999   -> removed
'             Jude 1:50   -> removed
'     3. Clamp Range Boundaries
'         If a range extends past valid scripture bounds,
'         trim the range to the last valid verse.
'         Example:
'             Gen 1:1-999
'         becomes:
'             Gen 1:1-31
'     4. Remove Empty Ranges
'         If validation removes all verses in a range,
'         discard the range entirely.
'         Example:
'             Matt 29:1-10 -> removed
'     5. Preserve Canonical Order
'         Validation must NOT reorder references.
'         The Stage 12 ordering must remain intact.
'     6. Preserve Compression
'         Stage 15 must not expand ranges.
'         It may only:
'             - Trim
'             - Remove
'             - Keep
' Input:
'     Packed canonical verse map (compressed, ordered)
' Output:
'     Fully validated packed canonical verse map
' Guarantees After Stage 15:
'     - All books valid
'     - All chapters valid
'     - All verses valid
'     - No empty ranges
'     - Canonically ordered
'     - Canonically compressed
'     - Engine-safe output
' Design Rules:
'     - No allocation of new structures
'     - Operate directly on packed verse map
'     - O(n) scan across canonical set
'     - Validation tables must be constant lookup
' Required Data:
'     ChaptersPerBook(book)
'     VersesPerChapter(book, chapter)
' Example Input:
'     Gen 1:1-999, 51:1-10, Exod 1:1-5
' After Stage 15:
'     Gen 1:1-31, Exod 1:1-5
' Summary:
'     Stage 15 is the final correctness gate.
'     After this stage, the parser output is guaranteed
'     to represent only valid scripture references.
'=====================================================

'=====================================================
' Stage 16 - Canonical Range Builder
'=====================================================
' Purpose:
'     Convert the validated verse-level canonical reference set into
'     contiguous canonical ranges prior to string formatting.
' Position in Pipeline:
'     Stage 16 executes AFTER Stage 15 Canonical Validation
'     and BEFORE Stage 17 Canonical String Formatter.
'     At this point the reference set is:
'         - Fully expanded
'         - Deduplicated
'         - Ordered
'         - Canonically valid
'     Stage 16 groups adjacent verses into canonical ranges.
' Responsibilities:
'     1. Detect Contiguous Verses
'         Identify adjacent verses within same book and chapter.
'         Example:
'             John 3:16
'             John 3:17
'             John 3:18
'         becomes:
'             John 3:16-3:18
'     2. Preserve Non-Adjacent Verses
'         Example:
'             John 3:16
'             John 3:18
'         remains:
'             John 3:16
'             John 3:18
'     3. Stop Range at Chapter Boundary
'         Example:
'             John 3:36
'             John 4:1
'         remains:
'             John 3:36
'             John 4:1
'     4. Stop Range at Book Boundary
'         Example:
'             John 3:16
'             Romans 8:1
'         remains separate
'     5. Produce Minimal Canonical Ranges
'         Example:
'             16,17,18,19
'         becomes:
'             16-19
' Input:
'     Ordered validated verse list
' Output:
'     Canonical range collection
' Design Rules:
'     - Single forward pass
'     - O(n)
'     - No reordering
'     - No expansion
'     - Only grouping
' Guarantees After Stage 16:
'     - Contiguous verses grouped
'     - Minimal ranges created
'     - Canonical order preserved
'     - No cross-chapter merging
'     - No cross-book merging
' Summary:
'     Stage 16 builds canonical ranges from validated verse-level
'     references so Stage 17 can perform final string formatting.
'=====================================================

'=====================================================
' Stage 17 - Canonical String Formatter
'=====================================================
' Purpose:
'     Convert the validated canonical reference ranges into a single
'     properly formatted SBL-style reference string for display,
'     logging, export, and round-trip parsing.
' Position in Pipeline:
'     Stage 17 executes AFTER:
'         Stage 15 - Canonical Validation
'     At this point the reference set is:
'         - Fully expanded
'         - Deduplicated
'         - Canonically ordered
'         - Canonically compressed into ranges
'         - Fully validated
'     Stage 17 performs formatting only.
'     It MUST NOT change the reference structure.
' Responsibilities:
'     1. Render Canonical Ranges
'         Convert each canonical range into text.
'         Examples:
'             John 3:16
'             John 3:16-3:18
'             Romans 8:1-8:2
'     2. Suppress Repeated Book Names
'         The book name is printed once per contiguous group.
'         Example:
'             John 3:16
'             John 3:18
'         becomes:
'             John 3:16, 18
'     3. Suppress Repeated Chapter Numbers
'         Chapter number appears once when ranges remain
'         within the same chapter.
'         Example:
'             John 3:16
'             John 3:18
'         becomes:
'             John 3:16, 18
'     4. Use Comma for Same-Chapter Separation
'         Example:
'             John 3:16
'             John 3:18
'         Output:
'             John 3:16, 18
'     5. Use Semicolon for Chapter Breaks
'         Example:
'             John 3:16-18
'             John 4:1-3
'         Output:
'             John 3:16-18; 4:1-3
'     6. Use Semicolon for Book Breaks
'         Example:
'             John 3:16
'             Romans 8:1
'         Output:
'             John 3:16; Rom 8:1
'     7. Preserve Canonical Order
'         Formatter must not reorder references.
'     8. Preserve Canonical Compression
'         Formatter must not recompute ranges.
'         It only renders existing canonical ranges.
' Input:
'     Validated canonical range collection
' Output:
'     Single formatted SBL reference string
' Formatting Rules:
'     Same chapter        -> comma
'     Chapter change      -> semicolon
'     Book change         -> semicolon
'     Range separator     -> hyphen or en dash
' Example:
' Input:
'     John 3:16-3:18
'     John 4:1-4:2
'     Romans 8:1-8:2
' Output:
'     John 3:16-18; 4:1-2; Rom 8:1-2
' Additional Examples:
'     Input:
'         Gen 1:1-1:3
'     Output:
'         Gen 1:1-3
'     Input:
'         Gen 1:1
'         Gen 1:3
'     Output:
'         Gen 1:1, 3
'     Input:
'         Gen 1:1
'         Gen 2:1
'     Output:
'         Gen 1:1; 2:1
'     Input:
'         Gen 1:1
'         Exod 1:1
'     Output:
'         Gen 1:1; Exod 1:1
' Design Rules:
'     - Pure formatting stage
'     - No mutation of canonical references
'     - Single pass O(n)
'     - Minimal string allocations
'     - Deterministic output
' Guarantees After Stage 17:
'     - Proper SBL formatting
'     - Minimal repetition
'     - Canonical ordering preserved
'     - Canonical compression preserved
'     - Human-readable output
'     - Round-trip parser safe
' Summary:
'     Stage 17 converts the validated canonical range set into a
'     compact, human-readable SBL formatted reference string suitable
'     for display, storage, and round-trip parsing.
'=====================================================

'=====================================================
' Deterministic Structural Parser (DSP)
' Aligned to 7-Stage Parser Architecture
'=====================================================
' PURPOSE:
'   Defines lexical token stream and structural syntax only.
' SCOPE:
'   This DSP:
'       - Validates structural ordering of tokens
'       - Detects lists and ranges syntactically
'   This DSP does NOT:
'       - Validate canonical book identity
'       - Enforce chapter/verse bounds
'       - Perform single-chapter inference
'       - Rewrite structure
'       - Normalize aliases
'   Semantic processing occurs in:
'       Stage 3 - Alias Resolution
'       Stage 4 - Structural Interpretation
'       Stage 5 - Canonical Validation
'=====================================================
' 1. Token Stream Definition (Stage 2 Output)
'=====================================================
' Token Types
' ------------------------------------------------------------
' | Token           | Description                            |
' | --------------- | -------------------------------------- |
' | BOOK_WORD       | Alphabetic word, optional trailing .   |
' | PREFIX_ARABIC   | 1 | 2 | 3                              |
' | PREFIX_ROMAN    | I | II | III                           |
' | DIGITS          | One or more digits                     |
' | COLON           | :                                      |
' | DASH            | -                                      |
' | COMMA           | ,                                      |
' | SEMICOLON       | ;                                      |
' | WS              | One or more spaces (collapsed)         |
' | EOF             | End-of-input                           |
' ------------------------------------------------------------
' Tokenization Rules:
' ------------------------------------------------------------
'   - Whitespace collapsed to single WS
'   - DIGITS is greedy
'   - Case-insensitive BOOK_WORD and PREFIX_ROMAN
'   - No internal punctuation in BOOK_WORD
'=====================================================
' 2. Deterministic State Machine (Structural Only)
'=====================================================
' NOTE:
'   This is a single-pass left-to-right DSP.
'   State numbering is symbolic.
'=====================================================
' 2.1 State Definitions
' -------------------------------------------------
' | State | Meaning                   | Accepting |
' | ----- | ------------------------- | --------- |
' | S0    | Start                     | No        |
' | S1    | Reading numeric prefix    | No        |
' | S2    | Reading book name         | Yes*      |
' | S3    | Expecting chapter digits  | No        |
' | S4    | Reading chapter digits    | Yes*      |
' | S6    | Reading verse digits      | Yes*      |
' | S7    | After dash                | No        |
' | S8    | After comma               | No        |
' | SX    | Error                     | No        |
' -------------------------------------------------
' *Accepting states require EOF or SEMICOLON next.
'  Book-only references (state S2) are normalized to
'  Chapter 1, Verse 1 during canonical formatting.
'==================================================
' 2.2 Transition Table
' -------------------------------------------------
' S0 - Start
'   WS              -> S0
'   PREFIX_*        -> S1
'   BOOK_WORD       -> S2
'   otherwise       -> SX
' S1 - Prefix
'   WS              -> S2
'   BOOK_WORD       -> S2
'   otherwise       -> SX
'S2 - Book Name
'   BOOK_WORD       -> S2
'   WS              -> S3
'   SEMICOLON       -> S0
'   EOF             -> ACCEPT
'   otherwise       -> SX
' S3 - Expect Chapter
'   DIGITS          -> S4
'   otherwise       -> SX
' S4 - Chapter (Accepting*)
'   DIGITS          -> S4
'   COLON           -> S6
'   DASH            -> S7
'   SEMICOLON       -> S0
'   EOF             -> ACCEPT
'   otherwise       -> SX
' S6 - Verse (Accepting*)
'   DIGITS          -> S6
'   DASH            -> S7
'   COMMA           -> S8
'   SEMICOLON       -> S0
'   EOF             -> ACCEPT
'   otherwise       -> SX
' S7 - After Dash
'   DIGITS          -> S6
'   otherwise       -> SX
' S8 - After Comma
'   DIGITS          -> S6
'   otherwise       -> SX
' SX - Error
'   any             -> SX
'=====================================================
' 3. Structural Acceptance Conditions
'=====================================================
' A reference is structurally valid if:
'   - Final state is S4 or S6
'   - AND next token is EOF or SEMICOLON
' All other terminal states are structural errors.
'=====================================================
' 4. Post-DSP Processing (Outside DSP Scope)
'=====================================================
' After structural acceptance:
' Stage 3:
'   Resolve BookRef -> BookID
' Stage 4:
'   Assign Chapter and Verse exactly once
'   Apply single-chapter inference if applicable
' Stage 5:
'   Enforce canonical validation matrix:
'       - AliasFound = True
'       - Chapter >= 1
'       - Chapter <= MaxChapter(BookID)
'       - Verse bounds
' Stage 6:
'   Format canonical SBL reference
' Stage 7:
'   Emit immutable ScriptureRef result
'=====================================================
' 5. Canonical Output Form (After Stage 6)
'=====================================================
' <CanonicalBookName> <Chapter>
' <CanonicalBookName> <Chapter>:<Verse>
' Lists and ranges are preserved structurally
' but validated semantically in Stage 5.
'=====================================================

' NOTES:
' What makes the compressed verse-map approach strong in your architecture is:
'1. Data Is Data
'   The validator performs only two checks:
'       - Is chapter within the book's chapter count?
'       - Is verse within the chapter's verse count?
'       It does not contain any hard-coded knowledge of
'       Bible structure.
'   Architectural Principle: Metadata Authority
'       Canonical Bible structure (chapter counts and
'       verse counts) is stored entirely in metadata tables.
'       Parser logic never encodes Bible structure.
'   Benefits:
'       - deterministic validation
'       - zero structural branching
'       - translation-specific metadata substitution
'       - Proper separation of concerns
'2. Deterministic O(1) Lookup
'   With fixed-width packed strings:
'       maxV = CLng(Mid$(map, (Chapter - 1) * 3 + 1, 3))
'   The lookup is constant time:
'       - No loops
'       - No Select Case
'       - No Split
'       - No dynamic lookup structures
'       - Direct index addressing only
' 3. It Scales Without Growing Code
'   Adding all 66 books:
'     - Adds data only
'     - Adds zero logic
'     - Validator does not encode Bible structure.
'       It only verifies numeric bounds using metadata lookups
' 4. It Matches the Design Philosophy
'    Parser stages remain isolated:
'       - Tokenizer
'       - Alias Resolver
'       - Structural Interpreter
'       - Semantic Validator
'       - Canonical Formatter
'5. The Key Benefit
'   Because structural rules are stored in metadata,
'   different validation policies can be implemented
'   without modifying parser logic.
'   Examples:
'       - Strict SBL validation
'       - Relaxed validation (bounds skipped)
'       - Alternate verse maps for
'         LXX, Vulgate, or other traditions
'   The validator remains unchanged.
'   Only the metadata source is replaced.
'
' The design has moved from string parsing to formal citation semantics - a big architectural leap.

' Special case:
' Psalm 119 contains 176 verses, which exceeds two-digit storage limits.
' Therefore all verse counts are stored using fixed-width three-digit encoding:
'    Right$("000" & verseCount, 3)
' This ensures constant-width indexing for direct addressing.

'===========================================================
' Type Declarations
' NOTE: VBA class modules require all UDTs to be Private.
' Functions using these types in their signatures are
' therefore also Private (internal to this class).
'===========================================================

Private Type ParsedReference
    ' Only structure needed for test harness
    RawInput   As String
    BookAlias  As String   ' e.g. "JUDE", "ROM"
    Chapter    As Long     ' 0 if omitted
    VerseSpec  As String   ' always string ("5", "1-3", "3,5")
End Type

Private Type LexTokens
    RawAlias As String
    Num1     As Long
    Num2     As Long
    HasColon As Boolean
    VerseSpec As String
End Type

Private Type ListTokens
    IsList As Boolean
    Segments() As String
End Type

'===========================================================
' RangeTokens
' Stage 9 output structure
'===========================================================
Private Type RangeTokens
    IsRange As Boolean
    LeftRaw As String
    RightRaw As String
End Type

'===========================================================
' ScriptureRef
'   Represents a single canonical Bible reference.
' Examples
'   John 3:16
'   Jude 1:5
'===========================================================
Private Type ScriptureRef
    BookID As Integer
    Chapter As Integer
    Verse As Integer
End Type

'===========================================================
' ScriptureRange
'   Represents a continuous reference range.
' Examples
'   John 3:16-18
'   John 3-5
'   John 3:16-4:2
'===========================================================
Private Type ScriptureRange
    StartRef  As ScriptureRef
    EndRef    As ScriptureRef
    ErrorCode As Long
    ErrorText As String
End Type

'===========================================================
' ScriptureList
'   Represents a list of references or ranges.
' Example
'   John 3:16-18,20; 4:1
'===========================================================
Private Type ScriptureList
    IsValid As Boolean
    ItemType() As Integer
        ' 1 = ScriptureRef
        ' 2 = ScriptureRange
    Refs() As ScriptureRef
    Ranges() As ScriptureRange
    ErrorCode As Long
    ErrorText As String
End Type

Public Enum CitationMode
    ModeGeneric = 0   ' Accept common abbreviations
    ModeSBL = 1       ' Enforce SBL Study Bible rules
End Enum

'===========================================================
' ContextState (Stage 13 internal only)
'   - Tracks the inherited book/chapter while resolving
'     contextual shorthand.
' This type is NOT returned outside Stage 13.
'===========================================================
Private Type ContextState
    BookID As Integer
    Chapter As Integer
End Type

Private aliasMap As Object

Private Sub Class_Initialize()
    On Error GoTo PROC_ERR
    Debug.Print "aeSBL_Citation_Class.Class_Initialize"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    Debug.Print "ERROR in aeSBL_Citation_Class.Class_Initialize | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    Resume PROC_EXIT
End Sub

Private Sub Class_Terminate()
    On Error GoTo PROC_ERR
    Debug.Print
    Debug.Print "aeSBL_Citation_Class.Class_Terminate"
    Set aliasMap = Nothing
PROC_EXIT:
    Exit Sub
PROC_ERR:
    Debug.Print "ERROR in aeSBL_Citation_Class.Class_Terminate | Erl: " & Erl _
        & " | Err: " & Err.Number & " | " & Err.Description
    Resume PROC_EXIT
End Sub


Public Sub ResetBookAliasMap()
    Set aliasMap = Nothing
End Sub

Private Function CanonicalFromRef(r As ScriptureRef) As String
    Dim books As Object
    Dim b As Variant
    Dim name As String

    Set books = GetCanonicalBookTable
    b = books.Item(r.BookID)
    name = b(1)   ' canonical book name

    If r.Verse = 0 Then
        CanonicalFromRef = name & " " & r.Chapter
    Else
        CanonicalFromRef = name & " " & r.Chapter & ":" & r.Verse
    End If
End Function

Private Function PackRef(r As ScriptureRef) As Variant
    PackRef = Array("REF", r.BookID, r.Chapter, r.Verse)
End Function

Private Function PackRange(r As ScriptureRange) As Variant
    Dim s As ScriptureRef
    Dim e As ScriptureRef
    s = r.StartRef
    e = r.EndRef
    PackRange = Array("RANGE", _
                      s.BookID, s.Chapter, s.Verse, _
                      e.BookID, e.Chapter, e.Verse)
End Function

Private Function CanonicalFromRange(rg As ScriptureRange) As String
    Dim startText As String, endText As String
    startText = CanonicalFromRef(rg.StartRef)
    If rg.StartRef.BookID = rg.EndRef.BookID Then
        endText = CStr(rg.EndRef.Chapter)
        If rg.EndRef.Verse <> 0 Then
            endText = endText & ":" & rg.EndRef.Verse
        End If
    Else
        endText = CanonicalFromRef(rg.EndRef)
    End If
    CanonicalFromRange = startText & "-" & endText
End Function

Public Function ComposeList(raw As String) As Collection
    Dim list As ScriptureList
    Dim Result As New Collection
    Dim i As Long

    ComposeList_Internal raw, list
    '------------------------------------------
    ' Stage 13 - Contextual Shorthand Expansion
    '------------------------------------------
    ApplyContextShorthand list
    For i = LBound(list.ItemType) To UBound(list.ItemType)
        Select Case list.ItemType(i)
        Case 1
            Result.Add CanonicalFromRef(list.Refs(i))
        Case 2
            Result.Add CanonicalFromRange(list.Ranges(i))
        End Select
    Next i

    Set ComposeList = Result
End Function

'=====================================================
' ComposeList_Internal
' Stage 11
'   Parses a list of scripture references.
'
' Items in the resulting ScriptureList.Items() array
' may contain either:
'   ScriptureRef
'   ScriptureRange
' A Variant array is used because VBA Collections
' cannot store UDT structures. Variant arrays *can*
' store copies of UDT values safely.
' Examples
'   "John 3:16, 18"
'   "John 3:16-18, 20"
' Contextual shorthand supported:
'   "John 3:16, 18, 20-22"
'       inherits book/chapter from previous item
' Pipeline stages used
'   Stage 8  - ListDetection
'   Stage 9  - IsRangeSegment
'   Stage 10 - ParseReferenceRef
'   Stage 11 - ComposeRange
'=====================================================
Private Sub ComposeList_Internal(raw As String, ByRef Result As ScriptureList)
    Dim t As ListTokens
    Dim i As Long
    Dim count As Long
    Dim capacity As Long
    Dim prevRef As ScriptureRef
    Dim havePrev As Boolean
    Dim seg As String
    Dim p As Long

    capacity = 8
    ReDim Result.ItemType(0 To capacity - 1)
    ReDim Result.Refs(0 To capacity - 1)
    ReDim Result.Ranges(0 To capacity - 1)
    '------------------------------------------
    ' Stage 8 - list detection
    '------------------------------------------
    t = ListDetection(raw)
    For i = LBound(t.Segments) To UBound(t.Segments)
        '------------------------------------------
        ' grow arrays if needed
        '------------------------------------------
        If count >= capacity Then
            capacity = capacity * 2
            ReDim Preserve Result.ItemType(0 To capacity - 1)
            ReDim Preserve Result.Refs(0 To capacity - 1)
            ReDim Preserve Result.Ranges(0 To capacity - 1)
        End If
        seg = Trim$(t.Segments(i))
        '------------------------------------------
        ' Stage 9 - range detection
        '------------------------------------------
        If IsRangeSegment(seg) Then
            '------------------------------------------
            ' Stage 13 shorthand numeric range
            '------------------------------------------
            If havePrev And IsNumericRange(seg) Then
                p = InStr(seg, "-")
                Result.ItemType(count) = 2
                Result.Ranges(count).StartRef.BookID = prevRef.BookID
                Result.Ranges(count).StartRef.Chapter = prevRef.Chapter
                Result.Ranges(count).StartRef.Verse = CLng(Left$(seg, p - 1))
                Result.Ranges(count).EndRef.BookID = prevRef.BookID
                Result.Ranges(count).EndRef.Chapter = prevRef.Chapter
                Result.Ranges(count).EndRef.Verse = CLng(Mid$(seg, p + 1))
            Else
                Result.ItemType(count) = 2
                Result.Ranges(count) = ComposeRange(seg)
            End If
            prevRef = Result.Ranges(count).EndRef
            havePrev = True
        Else
            '------------------------------------------
            ' Stage 13 shorthand single verse
            '------------------------------------------
            If havePrev And IsNumeric(seg) And InStr(seg, ":") = 0 Then
                Result.ItemType(count) = 1
                Result.Refs(count).BookID = prevRef.BookID
                Result.Refs(count).Chapter = prevRef.Chapter
                Result.Refs(count).Verse = CLng(seg)
            Else
                Result.ItemType(count) = 1
                Result.Refs(count) = ParseReferenceRef(seg)
            End If
            prevRef = Result.Refs(count)
            havePrev = True
        End If
        count = count + 1
    Next i
    '------------------------------------------
    ' trim arrays to actual size
    '------------------------------------------
    If count > 0 Then
        ReDim Preserve Result.ItemType(0 To count - 1)
        ReDim Preserve Result.Refs(0 To count - 1)
        ReDim Preserve Result.Ranges(0 To count - 1)
    End If
    Result.IsValid = True
End Sub

'=====================================================
' ApplyContextShorthand
' Stage 13
'   Expands contextual shorthand by inheriting
'   BookID and Chapter from the previous reference.
' Examples:
'   John 3:16, 18
'       -> John 3:16
'       -> John 3:18
'   John 3:16-4:2, 5
'       -> John 3:16-4:2
'       -> John 4:5
'=====================================================
Private Sub ApplyContextShorthand(ByRef list As ScriptureList)
    Dim i As Long
    Dim currentBook As Integer
    Dim currentChapter As Integer

    For i = LBound(list.ItemType) To UBound(list.ItemType)
        Select Case list.ItemType(i)
        '------------------------------------------
        ' Single reference
        '------------------------------------------
        Case 1
            If list.Refs(i).BookID = 0 Then
                list.Refs(i).BookID = currentBook
            End If
            If list.Refs(i).Chapter = 0 Then
                list.Refs(i).Chapter = currentChapter
            End If
            currentBook = list.Refs(i).BookID
            currentChapter = list.Refs(i).Chapter
        '------------------------------------------
        ' Range
        '------------------------------------------
        Case 2
            With list.Ranges(i)
                If .StartRef.BookID = 0 Then
                    .StartRef.BookID = currentBook
                End If
                If .StartRef.Chapter = 0 Then
                    .StartRef.Chapter = currentChapter
                End If
                If .EndRef.BookID = 0 Then
                    .EndRef.BookID = .StartRef.BookID
                End If
                If .EndRef.Chapter = 0 Then
                    .EndRef.Chapter = .StartRef.Chapter
                End If
                currentBook = .EndRef.BookID
                currentChapter = .EndRef.Chapter
            End With
        End Select
    Next i
End Sub

Private Function ListDetection(ByVal RawInput As String) As ListTokens
    Dim Result As ListTokens
    Dim parts() As String

    If InStr(RawInput, ",") > 0 Then
        parts = Split(RawInput, ",")
        Result.IsList = True
        Result.Segments = parts
        ListDetection = Result
        Exit Function
    End If

    If InStr(RawInput, ";") > 0 Then
        parts = Split(RawInput, ";")
        Result.IsList = True
        Result.Segments = parts
        ListDetection = Result
        Exit Function
    End If

    Result.IsList = False
    ListDetection = Result
End Function

'===========================================================
' RangeDetection Stage 9
' Detects and tokenizes reference ranges
' Example
'   John 3:16-18
'       LeftRaw  = "John 3:16"
'       RightRaw = "18"
'===========================================================
Private Function RangeDetection(ByVal segment As String) As RangeTokens
    Dim r As RangeTokens

    If Not IsRangeSegment(segment) Then
        r.IsRange = False
        RangeDetection = r
        Exit Function
    End If

    Dim dashPos As Long
    Dim enDash As String
    enDash = ChrW(8211)

    dashPos = InStr(segment, "-")
    If dashPos = 0 Then
        dashPos = InStr(segment, enDash)
    End If

    r.IsRange = True
    r.LeftRaw = Trim(Left(segment, dashPos - 1))
    r.RightRaw = Trim(Mid$(segment, dashPos + 1))

    RangeDetection = r
End Function

'===========================================================
' IsRangeSegment
'   Determines if a segment contains a range separator.
' Supported separators:
'   Hyphen  (-) ASCII 45
'   En dash (?) Unicode U+2013
' NOTE
' Use ChrW(8211) so the source code does not depend
' on editors correctly displaying the en dash.
'===========================================================
Public Function IsRangeSegment(ByVal segment As String) As Boolean
    Dim enDash As String
    enDash = ChrW(8211)   ' Unicode U+2013
    If InStr(segment, "-") > 0 Then
        IsRangeSegment = True
        Exit Function
    End If
    If InStr(segment, enDash) > 0 Then
        IsRangeSegment = True
        Exit Function
    End If
End Function

Private Function ComposeRange(ByVal raw As String) As ScriptureRange
    Dim tokens As RangeTokens
    Dim r As ScriptureRange
    Dim StartRef As ScriptureRef
    Dim EndRef As ScriptureRef
    Dim normalizedRight As String

    tokens = RangeDetection(raw)
    StartRef = ParseReferenceRef(tokens.LeftRaw)
    normalizedRight = ExpandRightSide(tokens.LeftRaw, tokens.RightRaw)
    EndRef = ParseReferenceRef(normalizedRight)
    r.StartRef = StartRef
    r.EndRef = EndRef

    ComposeRange = r   ' returns Variant instead of UDT
End Function

Public Function ParseReference( _
        ByVal inputRef As String, _
        Optional ByVal mode As CitationMode = ModeGeneric _
    ) As String
' ParseReference()  -> canonical string
' ParseReferenceRef() -> structured reference

    Dim t As LexTokens
    Dim BookID As Long
    Dim canonical As String
    Dim Chapter As Long
    Dim VerseSpec As String
    Dim Verse As Long
    Dim formatted As String

    '------------------------------------------
    ' Stage 2 - Lexical scan
    '------------------------------------------
    t = LexicalScan(inputRef)
    '------------------------------------------
    ' Stage 3 - Resolve alias
    '------------------------------------------
    canonical = ResolveBookStrict(t.RawAlias, BookID, mode)
    '------------------------------------------
    ' Stage 4 - Interpret structure
    '------------------------------------------
    Dim maxCh As Long
    maxCh = GetMaxChapter(BookID)
    If t.HasColon Then
        ' Book Chapter:Verse
        Chapter = t.Num1
        VerseSpec = CStr(t.Num2)
    Else
        If maxCh = 1 Then
            ' Single chapter book (Jude 5 -> 1:5)
            Chapter = 0
            VerseSpec = CStr(t.Num1)
        Else
            ' Multi chapter book (Romans 8)
            Chapter = t.Num1
            VerseSpec = ""
        End If
    End If
    '------------------------------------------
    ' Stage 5 - Canonical validation
    '------------------------------------------
    If Not ValidateSBLReference(BookID, canonical, Chapter, VerseSpec, mode) Then
        Err.Raise vbObjectError + 600, , "Invalid scripture reference"
    End If
    '------------------------------------------
    ' Stage 6 - Canonical formatting
    '------------------------------------------
    If VerseSpec = "" Then
        If Chapter = 0 Then
            formatted = RewriteSingleChapterRef(BookID, 0, 0)
        Else
            formatted = CStr(Chapter)
        End If
    Else
        formatted = RewriteSingleChapterRef(BookID, Chapter, CLng(VerseSpec))
    End If
    ParseReference = canonical & " " & formatted
End Function

'=====================================================
' ExpandRightSide
'   Expands shorthand range tokens so they become valid
'   canonical references for the atomic parser.
' Examples
'   John 3:16-18   -> John 3:18
'   Matt 5:3-12    -> Matt 5:12
'   John 3-5       -> John 5
'   Gen 1:31-2:3   -> Gen 2:3
'=====================================================
Private Function ExpandRightSide( _
        ByVal LeftRaw As String, _
        ByVal RightRaw As String _
    ) As String

    Dim leftParts() As String
    leftParts = Split(LeftRaw, " ")

    Dim book As String
    book = leftParts(0)
    '------------------------------------------
    ' Case 1: Right side already includes book
    '------------------------------------------
    If InStr(RightRaw, " ") > 0 Then
        ExpandRightSide = RightRaw
        Exit Function
    End If
    '------------------------------------------
    ' Case 2: Cross chapter range (2:3)
    '------------------------------------------
    If InStr(RightRaw, ":") > 0 Then
        ExpandRightSide = book & " " & RightRaw
        Exit Function
    End If
    '------------------------------------------
    ' Case 3: Verse shorthand
    '------------------------------------------
    If InStr(leftParts(1), ":") > 0 Then
        Dim cv() As String
        cv = Split(leftParts(1), ":")

        ExpandRightSide = book & " " & cv(0) & ":" & RightRaw
        Exit Function
    End If
    '------------------------------------------
    ' Case 4: Chapter shorthand
    '------------------------------------------
    ExpandRightSide = book & " " & RightRaw
End Function

'=====================================================
' ParseReferenceRef
' Adapter for extension stages (Stage 10+).
'   The atomic parser (ParseReference) returns canonical
'   reference text. This adapter converts that canonical
'   string into a ScriptureRef structure.
' Example
'   ParseReference("John 3:16")
'       -> "John 3:16"
'   ParseReferenceRef("John 3:16")
'       -> ScriptureRef(BookID=43, Chapter=3, Verse=16)
'=====================================================
Private Function ParseReferenceRef( _
        ByVal inputRef As String, _
        Optional ByVal mode As CitationMode = ModeGeneric _
    ) As ScriptureRef
' ParseReference()  -> canonical string
' ParseReferenceRef() -> structured reference

    Dim canon As String
    canon = ParseReference(inputRef, mode)

    Dim r As ScriptureRef
    Dim parts() As String
    Dim cv() As String
    Dim BookID As Long

    parts = Split(canon, " ")
    '------------------------------------------
    ' Resolve book ID again from canonical name
    '------------------------------------------
    Call ResolveBookStrict(parts(0), BookID, mode)
    r.BookID = BookID
    '------------------------------------------
    ' Parse chapter / verse
    '------------------------------------------
    If InStr(parts(1), ":") > 0 Then
        cv = Split(parts(1), ":")
        r.Chapter = CLng(cv(0))
        r.Verse = CLng(cv(1))
    Else
        r.Chapter = CLng(parts(1))
        r.Verse = 0
    End If

    ParseReferenceRef = r
End Function

Public Function IsListSegment(ByVal raw As String) As Boolean
    ' A list is indicated by comma or semicolon
    If InStr(raw, ",") > 0 Then
        IsListSegment = True
        Exit Function
    End If

    If InStr(raw, ";") > 0 Then
        IsListSegment = True
        Exit Function
    End If

    IsListSegment = False
End Function

Public Function ParseScripture(ByVal raw As String) As Variant
    raw = Trim$(raw)

    '------------------------------------------
    ' List
    '------------------------------------------
    If IsListSegment(raw) Then
        Set ParseScripture = ComposeList(raw)
        Exit Function
    End If
    '------------------------------------------
    ' Range
    '------------------------------------------
    If IsRangeSegment(raw) Then
        Dim r As ScriptureRange
        r = ComposeRange(raw)

        Dim startText As String
        Dim endText As String
        startText = CanonicalFromRef(r.StartRef)

        If r.StartRef.BookID = r.EndRef.BookID Then
            endText = CStr(r.EndRef.Chapter)
            If r.EndRef.Verse <> 0 Then
                endText = endText & ":" & r.EndRef.Verse
            End If
        Else
            endText = CanonicalFromRef(r.EndRef)
        End If
        ParseScripture = startText & "-" & endText
        Exit Function
    End If
    '------------------------------------------
    ' Single reference
    '------------------------------------------
    ParseScripture = ParseReference(raw)
End Function

'=====================================================
' LexicalScan
'    Stage 2 of the SBL parser pipeline.
' Purpose:
'   Splits a normalized scripture reference into:
'       RawAlias  -> book alias
'       Num1      -> chapter
'       Num2      -> verse
'       HasColon  -> indicates verse present
' Special Handling:
'   Book aliases beginning with a number (e.g. "1 John",
'   "2 Tim", "3 Jn") may appear without a numeric
'   reference (e.g. "1 Jn").
'   In this case the entire input is treated as the
'   RawAlias and numeric parsing is skipped. Later
'   stages will expand this to Chapter 1 Verse 1.
' Examples:
'   "Genesis"            -> RawAlias="Genesis"
'   "Romans 8"           -> RawAlias="Romans", Num1=8
'   "John 3:16"          -> RawAlias="John", Num1=3, Num2=16
'   "1 Jn"               -> RawAlias="1 Jn"
'   "1 Corinthians 13:4" -> RawAlias="1 Corinthians"
'=====================================================
Private Function LexicalScan(ByVal normalizedInput As String) As LexTokens
    Dim t As LexTokens
    Dim parts() As String
    Dim refPart As String

    parts = Split(Trim$(normalizedInput), " ")
    Select Case UBound(parts)
        Case 0
            ' Example: "Genesis"
            t.RawAlias = parts(0)
            LexicalScan = t
            Exit Function
        Case 1
            ' Example: "Romans 8"
            ' Example: "1 Jn"
            If IsNumeric(parts(0)) Then
                t.RawAlias = parts(0) & " " & parts(1)
                LexicalScan = t
                Exit Function
            Else
                t.RawAlias = parts(0)
                refPart = parts(1)
            End If
        Case Else
            ' Example: "3 John 4"
            ' Example: "1 Corinthians 13:4"

            t.RawAlias = parts(0) & " " & parts(1)
            refPart = parts(2)
    End Select
    '----------------------------------
    ' Parse numeric portion
    '----------------------------------
    If InStr(refPart, ":") > 0 Then
        Dim subParts() As String
        subParts = Split(refPart, ":")
        ' Chapter must be numeric
        If Not IsNumeric(subParts(0)) Then
            Err.Raise vbObjectError + 1000, , _
                "Invalid chapter reference: " & refPart
        End If
        t.Num1 = CLng(subParts(0))
        ' Preserve verse specification exactly
        t.VerseSpec = subParts(1)
        ' Populate Num2 only for simple verse
        If IsNumeric(subParts(1)) Then
            t.Num2 = CLng(subParts(1))
        Else
            t.Num2 = 0
        End If
        t.HasColon = True
    Else
        ' Single numeric reference
        If Not IsNumeric(refPart) Then
            Err.Raise vbObjectError + 1001, , _
                "Invalid numeric reference: " & refPart
        End If
        t.Num1 = CLng(refPart)
        t.Num2 = 0
        t.HasColon = False
    End If

    LexicalScan = t
End Function


Public Function IsValidSBLAlias(BookID As Long, aliasText As String) As Boolean
    Dim canonical As String
    Dim books As Object
    Dim expected As String

    Set books = GetCanonicalBookTable
    canonical = books(BookID)(1)    ' e.g. "1 John"

    ' Normalize both sides
    expected = UCase$(canonical)
    aliasText = UCase$(Trim$(aliasText))

    IsValidSBLAlias = (aliasText = expected)
End Function

Public Function ResolveBookStrict( _
        abbr As String, _
        Optional BookID As Long, _
        Optional mode As CitationMode = ModeGeneric _
    ) As String

    Dim canonical As String

    ' Step 1: Resolve (existing logic)
    canonical = ResolveAlias(abbr, BookID)

    ' Step 2: Validate (NEW)
    If mode = ModeSBL Then
        If Not IsValidSBLAlias(BookID, abbr) Then
            Err.Raise vbObjectError + 20, , _
                "Non-SBL book form: '" & abbr & _
                "'. Expected '" & canonical & "'"
        End If
    End If

    ResolveBookStrict = canonical
End Function

'========================================================
' Important distinctions - Canonical vs SBL tables
'========================================================
' | Aspect   | Canonical Table | SBL Table              |
' | -------- | --------------- | ---------------------- |
' | Purpose  | Identity        | Style enforcement      |
' | Case     | Mixed           | **Uppercase required** |
' | Variants | Allowed         | **Exactly one**        |
' | Usage    | Output          | Validation             |
'========================================================
Public Function GetCanonicalBookTable() As Object
    Static books As Object

    If books Is Nothing Then
        Set books = CreateObject("Scripting.Dictionary")

        books.Add 1, Array(1, "Genesis", 50)
        books.Add 2, Array(2, "Exodus", 40)
        books.Add 3, Array(3, "Leviticus", 27)
        books.Add 4, Array(4, "Numbers", 36)
        books.Add 5, Array(5, "Deuteronomy", 34)
        books.Add 6, Array(6, "Joshua", 24)
        books.Add 7, Array(7, "Judges", 21)
        books.Add 8, Array(8, "Ruth", 4)
        books.Add 9, Array(9, "1 Samuel", 31)
        books.Add 10, Array(10, "2 Samuel", 24)
        books.Add 11, Array(11, "1 Kings", 22)
        books.Add 12, Array(12, "2 Kings", 25)
        books.Add 13, Array(13, "1 Chronicles", 29)
        books.Add 14, Array(14, "2 Chronicles", 36)
        books.Add 15, Array(15, "Ezra", 10)
        books.Add 16, Array(16, "Nehemiah", 13)
        books.Add 17, Array(17, "Esther", 10)
        books.Add 18, Array(18, "Job", 42)
        books.Add 19, Array(19, "Psalms", 150)
        books.Add 20, Array(20, "Proverbs", 31)
        books.Add 21, Array(21, "Ecclesiastes", 12)
        books.Add 22, Array(22, "Solomon", 8)
        books.Add 23, Array(23, "Isaiah", 66)
        books.Add 24, Array(24, "Jeremiah", 52)
        books.Add 25, Array(25, "Lamentations", 5)
        books.Add 26, Array(26, "Ezekiel", 48)
        books.Add 27, Array(27, "Daniel", 12)
        books.Add 28, Array(28, "Hosea", 14)
        books.Add 29, Array(29, "Joel", 3)
        books.Add 30, Array(30, "Amos", 9)
        books.Add 31, Array(31, "Obadiah", 1)
        books.Add 32, Array(32, "Jonah", 4)
        books.Add 33, Array(33, "Micah", 7)
        books.Add 34, Array(34, "Nahum", 7)
        books.Add 35, Array(35, "Habakkuk", 3)
        books.Add 36, Array(36, "Zephaniah", 3)
        books.Add 37, Array(37, "Haggai", 2)
        books.Add 38, Array(38, "Zechariah", 14)
        books.Add 39, Array(39, "Malachi", 4)
        books.Add 40, Array(40, "Matthew", 28)
        books.Add 41, Array(41, "Mark", 16)
        books.Add 42, Array(42, "Luke", 24)
        books.Add 43, Array(43, "John", 21)
        books.Add 44, Array(44, "Acts", 28)
        books.Add 45, Array(45, "Romans", 16)
        books.Add 46, Array(46, "1 Corinthians", 16)
        books.Add 47, Array(47, "2 Corinthians", 13)
        books.Add 48, Array(48, "Galatians", 6)
        books.Add 49, Array(49, "Ephesians", 6)
        books.Add 50, Array(50, "Philippians", 4)
        books.Add 51, Array(51, "Colossians", 4)
        books.Add 52, Array(52, "1 Thessalonians", 5)
        books.Add 53, Array(53, "2 Thessalonians", 3)
        books.Add 54, Array(54, "1 Timothy", 6)
        books.Add 55, Array(55, "2 Timothy", 4)
        books.Add 56, Array(56, "Titus", 3)
        books.Add 57, Array(57, "Philemon", 1)
        books.Add 58, Array(58, "Hebrews", 13)
        books.Add 59, Array(59, "James", 5)
        books.Add 60, Array(60, "1 Peter", 5)
        books.Add 61, Array(61, "2 Peter", 3)
        books.Add 62, Array(62, "1 John", 5)
        books.Add 63, Array(63, "2 John", 1)
        books.Add 64, Array(64, "3 John", 1)
        books.Add 65, Array(65, "Jude", 1)
        books.Add 66, Array(66, "Revelation", 22)
    End If

    Set GetCanonicalBookTable = books
End Function

Public Function GetSBLCanonicalBookTable() As Object
    Static sbl As Object

    If sbl Is Nothing Then
        Set sbl = CreateObject("Scripting.Dictionary")

        sbl.Add 1, "GENESIS"
        sbl.Add 2, "EXODUS"
        sbl.Add 4, "Numbers"
        sbl.Add 5, "DEUTERONOMY"
        sbl.Add 6, "JOSHUA"
        sbl.Add 7, "JUDGES"
        sbl.Add 8, "RUTH"
        sbl.Add 9, "1 SAMUEL"
        sbl.Add 10, "2 SAMUEL"
        sbl.Add 11, "1 KINGS"
        sbl.Add 12, "2 KINGS"
        sbl.Add 13, "1 CHRONICLES"
        sbl.Add 14, "2 CHRONICLES"
        sbl.Add 15, "EZRA"
        sbl.Add 16, "NEHEMIAH"
        sbl.Add 17, "ESTHER"
        sbl.Add 18, "JOB"
        sbl.Add 19, "PSALMS"
        sbl.Add 20, "PROVERBS"
        sbl.Add 21, "ECCLESIASTES"
        sbl.Add 22, "SOLOMON"
        sbl.Add 23, "ISAIAH"
        sbl.Add 24, "JEREMIAH"
        sbl.Add 25, "LAMENTATIONS"
        sbl.Add 26, "EZEKIEL"
        sbl.Add 27, "DANIEL"
        sbl.Add 28, "HOSEA"
        sbl.Add 29, "JOEL"
        sbl.Add 30, "AMOS"
        sbl.Add 31, "OBADIAH"
        sbl.Add 32, "JONAH"
        sbl.Add 33, "MICAH"
        sbl.Add 34, "NAHUM"
        sbl.Add 35, "HABAKKUK"
        sbl.Add 36, "ZEPHANIAH"
        sbl.Add 37, "HAGGAI"
        sbl.Add 38, "ZECHARIAH"
        sbl.Add 39, "MALACHI"
        sbl.Add 40, "MATTHEW"
        sbl.Add 41, "MARK"
        sbl.Add 42, "LUKE"
        sbl.Add 43, "JOHN"
        sbl.Add 44, "ACTS"
        sbl.Add 45, "ROMANS"
        sbl.Add 46, "1 CORINTHIANS"
        sbl.Add 47, "2 CORINTHIANS"
        sbl.Add 48, "GALATIONS"
        sbl.Add 49, "EPHESIANS"
        sbl.Add 50, "PHILIPPIANS"
        sbl.Add 51, "COLOSSIANS"
        sbl.Add 52, "1 THESSALONIANS"
        sbl.Add 53, "2 THESSALONIANS"
        sbl.Add 54, "1 TIMOTHY"
        sbl.Add 55, "2 TIMOTHY"
        sbl.Add 56, "TITUS"
        sbl.Add 57, "PHILEMON"
        sbl.Add 58, "HEBREWS"
        sbl.Add 59, "JAMES"
        sbl.Add 60, "1 PETER"
        sbl.Add 62, "1 JOHN"
        sbl.Add 63, "2 JOHN"
        sbl.Add 64, "3 JOHN"
        sbl.Add 64, "JUDE"
        sbl.Add 66, "REVELATION"
    End If

    Set GetSBLCanonicalBookTable = sbl
End Function

Public Sub ValidateBookSBL( _
        ByVal BookID As Long, _
        ByVal InputBook As String)

    Dim sbl As Object
    Dim expected As String

    Set sbl = GetSBLCanonicalBookTable

    If Not sbl.Exists(BookID) Then
        Err.Raise vbObjectError + 400, , _
            "Book not defined in SBL canon: " & BookID
    End If

    expected = sbl(BookID)

    If UCase(Trim(InputBook)) <> expected Then
        Err.Raise vbObjectError + 401, , _
            "Non-SBL book form. Expected '" & expected & _
            "', got '" & InputBook & "'"
    End If
End Sub

Public Function GetMaxChapter(BookID As Long) As Long
    Dim books As Object
    Dim b As Variant

    Set books = GetCanonicalBookTable

    If Not books.Exists(BookID) Then
        Err.Raise vbObjectError + 20, , "Unknown BookID in GetMaxChapter"
    End If

    b = books(BookID)
    GetMaxChapter = CLng(b(2))
End Function

Public Sub AssertOneBased(arr As Variant, context As String)
    If LBound(arr) <> 1 Then
        Err.Raise vbObjectError + 900, , _
            "0-Based array detected in " & context
    End If
End Sub

Public Function GetMaxVerse(BookID As Long, Chapter As Long) As Long
    Dim maps As Variant
    maps = GetChapterVerseMap()   ' actually raw map now

    AssertOneBased maps(BookID), "GetMaxVerse chapters"

    If BookID < 1 Or BookID > 66 Then
        Err.Raise vbObjectError + 41, , "Invalid BookID"
    End If

    If Not IsArray(maps(BookID)) Then
        Err.Raise vbObjectError + 42, , "Chapter map missing"
    End If

    If Chapter < 1 Or Chapter > UBound(maps(BookID)) + 1 Then
        Err.Raise vbObjectError + 40, , "Invalid chapter for book"
    End If

    GetMaxVerse = maps(BookID)(Chapter)
End Function

Public Function GetChapterVerseMap() As Variant

    Static cached As Variant

    If IsEmpty(cached) Then

        Dim d As Object
        Set d = GetVerseCounts()

        Dim maps(1 To 66) As Variant
        Dim BookID As Long

        For BookID = 1 To 66
            maps(BookID) = d(BookID)
            AssertOneBased maps(BookID), "GetChapterVerseMap chapters"
        Next BookID

        AssertOneBased maps, "GetChapterVerseMap books"

        cached = maps
    End If

    GetChapterVerseMap = cached

End Function

Public Function ValidateSBLReference( _
        BookID As Long, _
        canonicalName As String, _
        Chapter As Long, _
        VerseSpec As String, _
        mode As CitationMode, _
        Optional silent As Boolean = False) As Boolean

    ' Generic mode: always valid at this layer
    If mode = ModeGeneric Then
        ValidateSBLReference = True
        Exit Function
    End If

    '------------------------------------------
    ' Normalize single-chapter shorthand
    '------------------------------------------
    If Chapter = 0 And VerseSpec <> "" Then
        If GetMaxChapter(BookID) = 1 Then
            ' Single-chapter book
            ' Example: Jude 5 -> Jude 1:5
            Chapter = 1
        Else
            ' Multi-chapter book
            ' Example: Romans 8 -> Romans 8
            Chapter = CLng(VerseSpec)
            VerseSpec = ""
        End If
    End If

    ' ---------- SBL MODE BELOW ----------
    ' 1. Book must exist in canonical table
    If Not GetCanonicalBookTable.Exists(BookID) Then
        Debug.Print "SBL FAIL: Unknown BookID " & BookID
        Exit Function
    End If

    ' 2. Canonical name must match SBL form EXACTLY
    ' (SBL is case-insensitive in print, but normalized internally)
    Dim canon As Variant
    canon = GetCanonicalBookTable(BookID)

    If UCase(canon(1)) <> UCase(canonicalName) Then
        Debug.Print "SBL FAIL: Non-canonical book name"
        Exit Function
    End If

    ' 3. Chapter rules
    If Chapter < 1 Then
        If Not silent Then Debug.Print "SBL FAIL: Chapter must be >= 1"
        ValidateSBLReference = False
        Exit Function
    End If

    Dim maxCh As Long
    maxCh = GetMaxChapter(BookID)
    If Chapter > maxCh Then
        If Not silent Then Debug.Print "SBL FAIL: Chapter exceeds max (" & maxCh & ")"
        ValidateSBLReference = False
        Exit Function
    End If

    ' 4. Single-chapter book rules
    If GetSingleChapterBookSet.Exists(BookID) Then
        If Chapter > 1 Then
            Debug.Print "SBL FAIL: Chapter > 1 for single-chapter book"
            Exit Function
        End If
    End If

    ' 5. Verse spec must exist
    If Len(VerseSpec) = 0 Then
        Debug.Print "SBL FAIL: Missing verse specification"
        Exit Function
    End If

    Dim v As Long

    If Not IsNumeric(VerseSpec) Then
        Debug.Print "SBL FAIL: Non-numeric verse"
        Exit Function
    End If

    v = CLng(VerseSpec)

    If v < 1 Then
        If Not silent Then Debug.Print "SBL FAIL: Verse must be >= 1"
        ValidateSBLReference = False
        Exit Function
    End If

    Dim maxV As Long
    maxV = GetMaxVerse(BookID, Chapter)

    If v > maxV Then
        Debug.Print "SBL FAIL: Verse exceeds max (" & maxV & ")"
        Exit Function
    End If

    ' NOTE:
    ' Verse range bounds (max verse per chapter)
    ' are intentionally NOT enforced here yet.
    ' That is a later enhancement.

    ValidateSBLReference = True
End Function

Public Function GetSingleChapterBookSet() As Object
    Static sc As Object

    If sc Is Nothing Then
        Set sc = CreateObject("Scripting.Dictionary")
        ' Old Testament
        sc.Add 31, True   ' Obadiah
        ' New Testament
        sc.Add 57, True   ' Philemon
        sc.Add 63, True   ' 2 John
        sc.Add 64, True   ' 3 John
        sc.Add 65, True   ' Jude
    End If

    Set GetSingleChapterBookSet = sc
End Function

Private Function InterpretStructure(ByRef tokens As LexTokens) As ParsedReference
    Dim ref As ParsedReference
    '------------------------------------------
    ' Chapter : Verse structure
    '------------------------------------------
    If tokens.HasColon Then
        ref.Chapter = tokens.Num1
        ' Preserve verse specification
        If Len(tokens.VerseSpec) > 0 Then
            ref.VerseSpec = tokens.VerseSpec
        ElseIf tokens.Num2 > 0 Then
            ref.VerseSpec = CStr(tokens.Num2)
        Else
            ref.VerseSpec = ""
        End If
    Else
        '------------------------------------------
        ' Single numeric reference (ambiguous)
        '------------------------------------------
        If tokens.Num1 > 0 Then
            ref.Chapter = 0
            ref.VerseSpec = CStr(tokens.Num1)
        Else
            ref.Chapter = 0
            ref.VerseSpec = ""
        End If
    End If

    InterpretStructure = ref
End Function

Public Function RewriteSingleChapterRef( _
        ByVal BookID As Long, _
        ByVal Chapter As Long, _
        ByVal Verse As Long) As String

    Dim sc As Object
    Set sc = GetSingleChapterBookSet
    '-----------------------------------------
    ' BOOK ONLY
    ' Example: John -> 1:1
    '-----------------------------------------
    If Chapter = 0 And Verse = 0 Then
        RewriteSingleChapterRef = "1:1"
        Exit Function
    End If
    '-----------------------------------------
    ' Single-chapter books
    ' Example: Jude 5 -> 1:5
    '-----------------------------------------
    If sc.Exists(BookID) And Chapter = 0 Then
        RewriteSingleChapterRef = "1:" & Verse
        Exit Function
    End If
    '-----------------------------------------
    ' Normal reference
    '-----------------------------------------
    If Verse > 0 Then
        RewriteSingleChapterRef = Chapter & ":" & Verse
    Else
        ' Chapter-only reference OR book-only reference
        If Chapter = 0 Then
            RewriteSingleChapterRef = "1:1"
        Else
            RewriteSingleChapterRef = CStr(Chapter)
        End If
    End If
End Function

Public Function AliasCoverage( _
        Optional ByRef report As String = "" _
    ) As Boolean
' Validate the Alias Coverage

    Dim books As Object
    Dim aMap As Object
    Dim missing As Collection
    Dim k As Variant
    Dim canon As String

    Set books = GetCanonicalBookTable
    Set aMap = GetBookAliasMap
    Set missing = New Collection

    For Each k In books.Keys
        canon = UCase$(books(k)(1))   ' Canonical name

        If Not aMap.Exists(canon) Then
            missing.Add canon
        End If
    Next k

    If missing.count > 0 Then
        Dim i As Long
        report = "Missing canonical aliases:" & vbCrLf
        For i = 1 To missing.count
            report = report & "  - " & missing(i) & vbCrLf
        Next i

        AliasCoverage = False
    Else
        report = "Alias coverage complete (canonical names present)."
        AliasCoverage = True
    End If
End Function

Public Function GetBookAliasMap() As Object
    ' Single-letter aliases are not allowed due to potential false positives
    ' Sort form allowed, common in Europe, (International / Critical Apparatus Style)

    If aliasMap Is Nothing Then
        Set aliasMap = CreateObject("Scripting.Dictionary")

        ' Genesis
        aliasMap.Add "GENESIS", 1
        aliasMap.Add "GEN", 1
        aliasMap.Add "GE", 1
        aliasMap.Add "GN", 1
        ' Exodus
        aliasMap.Add "EXODUS", 2
        aliasMap.Add "EXOD", 2
        aliasMap.Add "EXO", 2
        aliasMap.Add "EX", 2
        ' Leviticus
        aliasMap.Add "LEVITICUS", 3
        aliasMap.Add "LEV", 3
        aliasMap.Add "LE", 3
        aliasMap.Add "LV", 3
        ' Numbers
        aliasMap.Add "NUMBERS", 4
        aliasMap.Add "NUM", 4
        aliasMap.Add "NU", 4
        aliasMap.Add "NM", 4
        ' Deuteronomy
        aliasMap.Add "DEUTERONOMY", 5
        aliasMap.Add "DEUT", 5
        aliasMap.Add "DEU", 5
        aliasMap.Add "DE", 5
        aliasMap.Add "DT", 5
        ' Joshua
        aliasMap.Add "JOSHUA", 6
        aliasMap.Add "JOSH", 6
        aliasMap.Add "JOS", 6
        ' Judges
        aliasMap.Add "JUDGES", 7
        aliasMap.Add "JUDGE", 7
        aliasMap.Add "JUDG", 7
        aliasMap.Add "JGS", 7
        ' Ruth
        aliasMap.Add "RUTH", 8
        aliasMap.Add "RUT", 8
        aliasMap.Add "RU", 8
        ' 1 Samuel
        aliasMap.Add "1 SAMUEL", 9
        aliasMap.Add "1 SAM", 9
        aliasMap.Add "1 SA", 9
        aliasMap.Add "1 SM", 9
        ' 2 Samuel
        aliasMap.Add "2 SAMUEL", 10
        aliasMap.Add "2 SAM", 10
        aliasMap.Add "2 SA", 10
        aliasMap.Add "2 SM", 10
        ' 1 Kings
        aliasMap.Add "1 KINGS", 11
        aliasMap.Add "1 KGS", 11
        aliasMap.Add "1 KING", 11
        aliasMap.Add "1 KIN", 11
        aliasMap.Add "1 KI", 11
        ' 2 Kings
        aliasMap.Add "2 KINGS", 12
        aliasMap.Add "2 KGS", 12
        aliasMap.Add "2 KING", 12
        aliasMap.Add "2 KIN", 12
        aliasMap.Add "2 KI", 12
        ' 1 Chronicles
        aliasMap.Add "1 CHRONICLES", 13
        aliasMap.Add "1 CHRON", 13
        aliasMap.Add "1 CHRO", 13
        aliasMap.Add "1 CHR", 13
        aliasMap.Add "1 CH", 13
        ' 2 Chronicles
        aliasMap.Add "2 CHRONICLES", 14
        aliasMap.Add "2 CHRON", 14
        aliasMap.Add "2 CHRO", 14
        aliasMap.Add "2 CHR", 14
        aliasMap.Add "2 CH", 14
        ' Ezra
        aliasMap.Add "EZRA", 15
        aliasMap.Add "EZR", 15
        ' Nehemiah
        aliasMap.Add "NEHEMIAH", 16
        aliasMap.Add "NEH", 16
        aliasMap.Add "NE", 16
        ' Esther
        aliasMap.Add "ESTHER", 17
        aliasMap.Add "ESTH", 17
        aliasMap.Add "EST", 17
        aliasMap.Add "ES", 17
        ' Job
        aliasMap.Add "JOB", 18
        aliasMap.Add "JB", 18
        ' Psalms
        aliasMap.Add "PSALMS", 19
        aliasMap.Add "PSALM", 19
        aliasMap.Add "PSA", 19
        aliasMap.Add "PS", 19
        ' Proverbs
        aliasMap.Add "PROVERBS", 20
        aliasMap.Add "PROV", 20
        aliasMap.Add "PRO", 20
        aliasMap.Add "PR", 20
        aliasMap.Add "PRV", 20
        ' Ecclesiastes
        aliasMap.Add "ECCLESIASTES", 21
        aliasMap.Add "ECCL", 21
        aliasMap.Add "ECC", 21
        aliasMap.Add "EC", 21
        ' Solomon
        aliasMap.Add "SOLOMON", 22
        aliasMap.Add "SOLO", 22
        aliasMap.Add "SOL", 22
        aliasMap.Add "SO", 22
        aliasMap.Add "SONG", 22
        aliasMap.Add "SG", 22
        ' Isaiah
        aliasMap.Add "ISAIAH", 23
        aliasMap.Add "ISA", 23
        aliasMap.Add "IS", 23
        ' Jeremiah
        aliasMap.Add "JEREMIAH", 24
        aliasMap.Add "JER", 24
        aliasMap.Add "JE", 24
        ' Lamentations
        aliasMap.Add "LAMENTATIONS", 25
        aliasMap.Add "LAM", 25
        aliasMap.Add "LA", 25
        ' Ezekiel
        aliasMap.Add "EZEKIEL", 26
        aliasMap.Add "EZEK", 26
        aliasMap.Add "EZE", 26
        aliasMap.Add "EZ", 26
        ' Daniel
        aliasMap.Add "DANIEL", 27
        aliasMap.Add "DAN", 27
        aliasMap.Add "DA", 27
        aliasMap.Add "DN", 27
        ' Hosea
        aliasMap.Add "HOSEA", 28
        aliasMap.Add "HOS", 28
        aliasMap.Add "HO", 28
        ' Joel
        aliasMap.Add "JOEL", 29
        aliasMap.Add "JOE", 29
        aliasMap.Add "JL", 29
        ' Amos
        aliasMap.Add "AMOS", 30
        aliasMap.Add "AMO", 30
        aliasMap.Add "AM", 30
        ' Obadiah
        aliasMap.Add "OBADIAH", 31
        aliasMap.Add "OBAD", 31
        aliasMap.Add "OBA", 31
        aliasMap.Add "OB", 31
        ' Jonah
        aliasMap.Add "JONAH", 32
        aliasMap.Add "JONA", 32
        aliasMap.Add "JON", 32
        ' Micah
        aliasMap.Add "MICAH", 33
        aliasMap.Add "MIC", 33
        aliasMap.Add "MI", 33
        ' Nahum
        aliasMap.Add "NAHUM", 34
        aliasMap.Add "NAH", 34
        aliasMap.Add "NA", 34
        ' Habakkuk
        aliasMap.Add "HABAKKUK", 35
        aliasMap.Add "HAB", 35
        aliasMap.Add "HA", 35
        aliasMap.Add "HB", 35
        ' Zephaniah
        aliasMap.Add "ZEPHANIAH", 36
        aliasMap.Add "ZEPH", 36
        aliasMap.Add "ZEP", 36
        ' Haggai
        aliasMap.Add "HAGGAI", 37
        aliasMap.Add "HAG", 37
        aliasMap.Add "HG", 37
        ' Zechariah
        aliasMap.Add "ZECHARIAH", 38
        aliasMap.Add "ZECH", 38
        aliasMap.Add "ZEC", 38
        ' Malachi
        aliasMap.Add "MALACHI", 39
        aliasMap.Add "MAL", 39
        ' Matthew
        aliasMap.Add "MATTHEW", 40
        aliasMap.Add "MATT", 40
        aliasMap.Add "MAT", 40
        aliasMap.Add "MT", 40
        ' Mark
        aliasMap.Add "MARK", 41
        aliasMap.Add "MAR", 41
        aliasMap.Add "MK", 41
        ' Luke
        aliasMap.Add "LUKE", 42
        aliasMap.Add "LUK", 42
        aliasMap.Add "LU", 42
        aliasMap.Add "LK", 42
        ' John
        aliasMap.Add "JOHN", 43
        aliasMap.Add "JOH", 43
        aliasMap.Add "JN", 43
        ' Acts
        aliasMap.Add "ACTS", 44
        aliasMap.Add "ACT", 44
        aliasMap.Add "AC", 44
        ' Romans
        aliasMap.Add "ROMANS", 45
        aliasMap.Add "ROM", 45
        aliasMap.Add "RO", 45
        ' 1 Corinthians
        aliasMap.Add "1 CORINTHIANS", 46
        aliasMap.Add "1 COR", 46
        aliasMap.Add "1 CO", 46
        ' 2 Corinthians
        aliasMap.Add "2 CORINTHIANS", 47
        aliasMap.Add "2 COR", 47
        aliasMap.Add "2 CO", 47
        ' Galatians
        aliasMap.Add "GALATIANS", 48
        aliasMap.Add "GAL", 48
        aliasMap.Add "GA", 48
        ' Ephesians
        aliasMap.Add "EPHESIANS", 49
        aliasMap.Add "EPH", 49
        aliasMap.Add "EP", 49
        ' Philippians
        aliasMap.Add "PHILIPPIANS", 50
        aliasMap.Add "PHILI", 50
        aliasMap.Add "PHIL", 50
        ' Colossians
        aliasMap.Add "COLOSSIANS", 51
        aliasMap.Add "COL", 51
        aliasMap.Add "CO", 51
        ' 1 Thessalonians
        aliasMap.Add "1 THESSALONIANS", 52
        aliasMap.Add "1 THESS", 52
        aliasMap.Add "1 THES", 52
        aliasMap.Add "1 THE", 52
        aliasMap.Add "1 THS", 52
        ' 2 Thessalonians
        aliasMap.Add "2 THESSALONIANS", 53
        aliasMap.Add "2 THESS", 53
        aliasMap.Add "2 THES", 53
        aliasMap.Add "2 THE", 53
        aliasMap.Add "2 THS", 53
        ' 1 Timothy
        aliasMap.Add "1 TIMOTHY", 54
        aliasMap.Add "1 TIM", 54
        aliasMap.Add "1 TI", 54
        aliasMap.Add "1 TM", 54
        ' 2 Timothy
        aliasMap.Add "2 TIMOTHY", 55
        aliasMap.Add "2 TIM", 55
        aliasMap.Add "2 TI", 55
        aliasMap.Add "2 TM", 55
        ' Titus
        aliasMap.Add "TITUS", 56
        aliasMap.Add "TIT", 56
        aliasMap.Add "TI", 56
        ' Philemon
        aliasMap.Add "PHILEMON", 57
        aliasMap.Add "PHILE", 57
        aliasMap.Add "PHLM", 57
        ' Hebrews
        aliasMap.Add "HEBREWS", 58
        aliasMap.Add "HEB", 58
        aliasMap.Add "HE", 58
        ' James
        aliasMap.Add "JAMES", 59
        aliasMap.Add "JAM", 59
        aliasMap.Add "JA", 59
        aliasMap.Add "JAS", 59
        ' 1 Peter
        aliasMap.Add "1 PETER", 60
        aliasMap.Add "1 PET", 60
        aliasMap.Add "1 PE", 60
        aliasMap.Add "1 PT", 60
        ' 2 Peter
        aliasMap.Add "2 PETER", 61
        aliasMap.Add "2 PET", 61
        aliasMap.Add "2 PE", 61
        aliasMap.Add "2 PT", 61
        ' 1 John
        aliasMap.Add "1 JOHN", 62
        aliasMap.Add "1 JOH", 62
        aliasMap.Add "1 JN", 62
        ' 2 John
        aliasMap.Add "2 JOHN", 63
        aliasMap.Add "2 JOH", 63
        aliasMap.Add "2 JN", 63
        ' 3 John
        aliasMap.Add "3 JOHN", 64
        aliasMap.Add "3 JOH", 64
        aliasMap.Add "3 JN", 64
        ' Jude
        aliasMap.Add "JUDE", 65
        aliasMap.Add "JUD", 65
        ' Revelation
        aliasMap.Add "REVELATION", 66
        aliasMap.Add "REV", 66
        aliasMap.Add "RE", 66
        aliasMap.Add "RV", 66
    End If

    Set GetBookAliasMap = aliasMap
End Function

Private Function IsNumericRange(ByVal seg As String) As Boolean
    Dim p As Long
    p = InStr(seg, "-")
    If p = 0 Then Exit Function
    IsNumericRange = _
        IsNumeric(Left$(seg, p - 1)) And _
        IsNumeric(Mid$(seg, p + 1))
End Function

'=====================================================
' Stage 14 - Canonical Compression
'=====================================================
' Purpose:
'   Collapse adjacent consecutive same-chapter verses
'   into ranges in a resolved canonical reference list.
' Input:
'   A Collection of canonical reference strings,
'   such as the output of ComposeList / CanonicalFromRef.
'   Example: "John 3:16", "John 3:17", "John 3:18", "Romans 8:1"
' Output:
'   A Collection of compressed canonical strings.
'   Example: "John 3:16-3:18", "Romans 8:1"
' Compression rules:
'   1. Two refs are adjacent when same book, same chapter,
'      and second verse = first verse + 1.
'   2. A run of 2+ adjacent verses becomes a range:
'         BookName startChapter:startVerse-endChapter:endVerse
'   3. Cross-chapter adjacency is never collapsed.
'   4. Cross-book adjacency is never collapsed.
'   5. Order is preserved.
'   6. Non-canonical strings (chapter-only refs) are passed
'      through unchanged.
'=====================================================
Public Function CompressCanonical(ByVal Refs As Collection) As Collection
    On Error GoTo PROC_ERR
    Dim Result As New Collection
    Dim i As Long
    Dim n As Long

    n = Refs.count
    If n = 0 Then
        Set CompressCanonical = Result
        GoTo PROC_EXIT
    End If
    '------------------------------------------
    ' Parse all refs into parallel arrays
    '------------------------------------------
    Dim bookNames()  As String
    Dim chapters()   As Long
    Dim verses()     As Long
    Dim parsedOk()   As Boolean

    ReDim bookNames(1 To n)
    ReDim chapters(1 To n)
    ReDim verses(1 To n)
    ReDim parsedOk(1 To n)

    Dim refStr As String
    For i = 1 To n
        refStr = Refs(i)
        parsedOk(i) = ParseCanonicalRef( _
                          refStr, _
                          bookNames(i), _
                          chapters(i), _
                          verses(i))
    Next i
    '------------------------------------------
    ' Walk through refs, accumulating runs
    '------------------------------------------
    i = 1
    Do While i <= n
        '------------------------------------------
        ' If this ref could not be parsed as a
        ' chapter:verse ref, emit it unchanged
        '------------------------------------------
        If Not parsedOk(i) Then
            Result.Add Refs(i)
            i = i + 1
        Else
            '------------------------------------------
            ' Start a potential run
            '------------------------------------------
            Dim runStart As Long
            runStart = i
            '------------------------------------------
            ' Extend the run while adjacent.
            ' NOTE: VBA And is not short-circuit - guard
            ' the array accesses with a Do While i < n
            ' outer check and inner If/Exit Do tests.
            '------------------------------------------
            Do While i < n
                If Not parsedOk(i + 1) Then Exit Do
                If bookNames(i + 1) <> bookNames(i) Then Exit Do
                If chapters(i + 1) <> chapters(i) Then Exit Do
                If verses(i + 1) <> verses(i) + 1 Then Exit Do
                i = i + 1
            Loop
            '------------------------------------------
            ' Emit the run as a range or single ref
            '------------------------------------------
            If i > runStart Then
                ' Collapsed range: BookName startCh:startV-endCh:endV
                Result.Add bookNames(runStart) & " " & _
                           chapters(runStart) & ":" & verses(runStart) & _
                           "-" & _
                           chapters(i) & ":" & verses(i)
            Else
                ' Single ref: emit as-is (canonical form preserved)
                Result.Add Refs(i)
            End If

            i = i + 1
        End If
    Loop

    Set CompressCanonical = Result

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure CompressCanonical of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

'=====================================================
' ParseCanonicalRef
'   Parses a canonical reference string into its
'   component parts: book name, chapter, verse.
' Accepts strings of the form:
'   "BookName chapter:verse"   e.g. "John 3:16"
'   "BookName chapter"         e.g. "Romans 8"
'                              (verse returns 0; parsedOk=False)
' Multi-word book names ("1 John", "Song of Solomon")
'   are handled by treating everything up to the last
'   space-delimited token that contains ":" (or is purely
'   numeric) as the book name.
' Returns True when a chapter:verse pair was successfully
'   extracted. Returns False for chapter-only refs or
'   strings that cannot be parsed.
'=====================================================
Private Function ParseCanonicalRef( _
        ByVal canonical As String, _
        ByRef bookName As String, _
        ByRef Chapter As Long, _
        ByRef Verse As Long _
    ) As Boolean
    On Error GoTo PROC_ERR

    bookName = ""
    Chapter = 0
    Verse = 0

    canonical = Trim$(canonical)
    If Len(canonical) = 0 Then GoTo PROC_EXIT

    '------------------------------------------
    ' Find the last space ? everything after it
    ' is the chapter[:verse] token.
    '------------------------------------------
    Dim lastSpace As Long
    lastSpace = 0
    Dim k As Long
    For k = Len(canonical) To 1 Step -1
        If Mid$(canonical, k, 1) = " " Then
            lastSpace = k
            Exit For
        End If
    Next k

    If lastSpace = 0 Then GoTo PROC_EXIT   ' no space found ? unparseable

    Dim refToken As String
    refToken = Mid$(canonical, lastSpace + 1)
    bookName = Left$(canonical, lastSpace - 1)

    '------------------------------------------
    ' Parse chapter[:verse] token
    '------------------------------------------
    Dim colonPos As Long
    colonPos = InStr(refToken, ":")

    If colonPos > 0 Then
        Dim chStr As String
        Dim vStr As String
        chStr = Left$(refToken, colonPos - 1)
        vStr = Mid$(refToken, colonPos + 1)
        If IsNumeric(chStr) And IsNumeric(vStr) Then
            Chapter = CLng(chStr)
            Verse = CLng(vStr)
            ParseCanonicalRef = True   ' chapter:verse successfully parsed
        End If
    End If
    ' If no colon (chapter-only ref), we return False
    ' (chapter-only refs are emitted unchanged by CompressCanonical)

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ParseCanonicalRef of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

Public Function ResolveAlias(abbr As String, _
                            Optional BookID As Long) As String
    Dim key As String
    Dim aMap As Object
    Dim books As Object
    Dim b As Variant

    key = UCase$(Trim$(abbr))
    Set aMap = GetBookAliasMap

    If Not aMap.Exists(key) Then
        Debug.Print "BAD  > ResolveBook(" & key & ")"
        Err.Raise vbObjectError + 10, , "Unknown book alias: " & abbr
    End If

    BookID = aMap(key)
    Set books = GetCanonicalBookTable

    b = books.Item(BookID)     ' SAFE
    ResolveAlias = b(1)         ' Canonical name
End Function

'=====================================================
' Stage 15 - Canonical Validation
'=====================================================
' Purpose:
'     Remove or trim canonical reference strings that fall
'     outside the valid scripture bounds of their book.
' Input:
'     A Collection of canonical strings ? output of Stage 14.
'     Strings may be:
'         Single ref:   "John 3:16"
'         Range:        "John 3:16-3:18"
'         Chapter-only: "Romans 8"   (passed through unchanged)
' Output:
'     A new Collection with invalid refs removed or trimmed.
' Processing rules per item:
'     Single ref "Book ch:v"
'         - ch > ChaptersInBook  -> remove
'         - v  > VersesInChapter -> remove
'         - otherwise            -> keep
'     Range "Book startCh:startV-endCh:endV"
'         - startCh > ChaptersInBook        -> remove
'         - startV  > VersesInChapter(start) -> remove
'         - endCh: clamp to ChaptersInBook
'         - endV:  clamp to VersesInChapter(endCh)
'         - after clamping, end < start      -> remove
'         - after clamping, end = start      -> emit single ref
'         - otherwise                        -> emit clamped range
'     Chapter-only: pass through unchanged
'=====================================================

'-----------------------------------------------------
' ChaptersInBook
'   Returns the number of chapters in the named book.
'   Returns 0 if the book name cannot be resolved.
'-----------------------------------------------------
Private Function ChaptersInBook(ByVal bookName As String) As Long
    On Error GoTo PROC_ERR
    Dim aMap As Object
    Dim key As String
    Dim BookID As Long

    key = UCase$(Trim$(bookName))
    Set aMap = GetBookAliasMap

    If Not aMap.Exists(key) Then
        ChaptersInBook = 0
        GoTo PROC_EXIT
    End If
    BookID = aMap(key)
    ChaptersInBook = GetMaxChapter(BookID)

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ChaptersInBook of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

'-----------------------------------------------------
' VersesInChapter
'   Returns the number of verses in the specified
'   chapter of the named book.
'   Returns 0 if the book name or chapter is invalid.
'-----------------------------------------------------
Private Function VersesInChapter(ByVal bookName As String, ByVal Chapter As Long) As Long
    On Error GoTo PROC_ERR

    Dim aMap As Object
    Dim key As String
    Dim BookID As Long
    Dim maxCh As Long

    key = UCase$(Trim$(bookName))
    Set aMap = GetBookAliasMap

    If Not aMap.Exists(key) Then
        VersesInChapter = 0
        GoTo PROC_EXIT
    End If
    BookID = aMap(key)
    maxCh = GetMaxChapter(BookID)

    If Chapter < 1 Or Chapter > maxCh Then
        VersesInChapter = 0
        GoTo PROC_EXIT
    End If
    VersesInChapter = GetMaxVerse(BookID, Chapter)

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure VersesInChapter of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

'-----------------------------------------------------
' ParseCanonicalRange
'   Parses a range string of the form:
'       "BookName startCh:startV-endCh:endV"
'   into its component parts.
'   Returns True if successfully parsed as a range.
'   Returns False for single refs, chapter-only refs,
'   or strings that cannot be parsed.
'-----------------------------------------------------
Private Function ParseCanonicalRange( _
        ByVal s As String, _
        ByRef bookName As String, _
        ByRef startCh As Long, ByRef startV As Long, _
        ByRef endCh As Long, ByRef endV As Long _
    ) As Boolean
    On Error GoTo PROC_ERR

    bookName = ""
    startCh = 0
    startV = 0
    endCh = 0
    endV = 0

    s = Trim$(s)
    If Len(s) = 0 Then GoTo PROC_EXIT

    '------------------------------------------
    ' Find the last space to isolate the
    ' reference token "startCh:startV-endCh:endV"
    '------------------------------------------
    Dim lastSpace As Long
    Dim k As Long
    lastSpace = 0
    For k = Len(s) To 1 Step -1
        If Mid$(s, k, 1) = " " Then
            lastSpace = k
            Exit For
        End If
    Next k

    If lastSpace = 0 Then GoTo PROC_EXIT

    Dim refToken As String
    refToken = Mid$(s, lastSpace + 1)
    bookName = Left$(s, lastSpace - 1)
    '------------------------------------------
    ' The reference token must contain a dash
    ' separating start ch:v from end ch:v.
    ' Find the dash that follows a digit.
    '------------------------------------------
    Dim dashPos As Long
    dashPos = 0
    Dim c As String
    For k = 2 To Len(refToken)
        c = Mid$(refToken, k, 1)
        If c = "-" Then
            ' Verify the preceding character is a digit
            If Mid$(refToken, k - 1, 1) >= "0" And Mid$(refToken, k - 1, 1) <= "9" Then
                dashPos = k
                Exit For
            End If
        End If
    Next k

    If dashPos = 0 Then GoTo PROC_EXIT

    Dim leftPart As String
    Dim rightPart As String
    leftPart = Left$(refToken, dashPos - 1)
    rightPart = Mid$(refToken, dashPos + 1)
    '------------------------------------------
    ' Parse left side: startCh:startV
    '------------------------------------------
    Dim colonLeft As Long
    colonLeft = InStr(leftPart, ":")
    If colonLeft = 0 Then GoTo PROC_EXIT

    Dim chStrL As String
    Dim vStrL As String
    chStrL = Left$(leftPart, colonLeft - 1)
    vStrL = Mid$(leftPart, colonLeft + 1)
    If Not IsNumeric(chStrL) Or Not IsNumeric(vStrL) Then GoTo PROC_EXIT
    startCh = CLng(chStrL)
    startV = CLng(vStrL)
    '------------------------------------------
    ' Parse right side: endCh:endV
    '------------------------------------------
    Dim colonRight As Long
    colonRight = InStr(rightPart, ":")
    If colonRight = 0 Then GoTo PROC_EXIT

    Dim chStrR As String
    Dim vStrR As String
    chStrR = Left$(rightPart, colonRight - 1)
    vStrR = Mid$(rightPart, colonRight + 1)
    If Not IsNumeric(chStrR) Or Not IsNumeric(vStrR) Then GoTo PROC_EXIT
    endCh = CLng(chStrR)
    endV = CLng(vStrR)

    ParseCanonicalRange = True
PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ParseCanonicalRange of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

'=====================================================
' ValidateCanonical
' Stage 15
'   Validates a Collection of canonical reference
'   strings against scripture metadata.
'   Invalid references are removed.
'   Range boundaries are clamped to valid limits.
' Input:
'   Refs -  Collection of canonical strings from Stage 14
' Output:
'   New Collection with only valid references.
'=====================================================
Public Function ValidateCanonical(ByVal Refs As Collection) As Collection
    On Error GoTo PROC_ERR

    Dim Result As New Collection
    Dim item As Variant
    Dim s As String

    For Each item In Refs
        s = CStr(item)

        '------------------------------------------
        ' Attempt to parse as a range first
        '------------------------------------------
        Dim bkName As String
        Dim sCh As Long, sV As Long
        Dim eCh As Long, eV As Long

        If ParseCanonicalRange(s, bkName, sCh, sV, eCh, eV) Then
            '------------------------------------------
            ' It is a range: validate and clamp
            '------------------------------------------
            Dim maxChBook As Long
            maxChBook = ChaptersInBook(bkName)

            ' Start chapter invalid -> remove
            If sCh < 1 Or sCh > maxChBook Then GoTo NextItem

            ' Start verse invalid -> remove
            Dim maxVStart As Long
            maxVStart = VersesInChapter(bkName, sCh)
            If sV < 1 Or sV > maxVStart Then GoTo NextItem

            ' Clamp end chapter
            If eCh > maxChBook Then
                eCh = maxChBook
            End If

            ' Clamp end verse
            Dim maxVEnd As Long
            maxVEnd = VersesInChapter(bkName, eCh)
            If eV > maxVEnd Then
                eV = maxVEnd
            End If

            ' After clamping, check end >= start canonically
            ' (same chapter: verse comparison; cross-chapter: chapter wins)
            Dim endBeforeStart As Boolean
            If eCh < sCh Then
                endBeforeStart = True
            ElseIf eCh = sCh And eV < sV Then
                endBeforeStart = True
            Else
                endBeforeStart = False
            End If

            If endBeforeStart Then GoTo NextItem

            ' Check if clamped range collapses to single ref
            If eCh = sCh And eV = sV Then
                ' Emit as single ref
                Result.Add bkName & " " & sCh & ":" & sV
            Else
                ' Emit clamped range
                Result.Add bkName & " " & sCh & ":" & sV & "-" & eCh & ":" & eV
            End If

        Else
            '------------------------------------------
            ' Not a range - try to parse as single ref
            '------------------------------------------
            Dim singleBook As String
            Dim singleCh As Long
            Dim singleV As Long
            Dim isParsed As Boolean

            isParsed = ParseCanonicalRef(s, singleBook, singleCh, singleV)

            If Not isParsed Then
                ' Chapter-only ref or unparseable - pass through unchanged
                Result.Add s
            Else
                ' Single ch:v ref ? validate
                Dim maxChSingle As Long
                maxChSingle = ChaptersInBook(singleBook)

                If maxChSingle = 0 Then GoTo NextItem
                If singleCh < 1 Or singleCh > maxChSingle Then GoTo NextItem

                Dim maxVSingle As Long
                maxVSingle = VersesInChapter(singleBook, singleCh)
                If singleV < 1 Or singleV > maxVSingle Then GoTo NextItem

                ' Valid - keep as-is
                Result.Add s
            End If
        End If

NextItem:
    Next item

    Set ValidateCanonical = Result
PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure ValidateCanonical of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

'=====================================================
' BuildCanonicalRanges
' Stage 16
'   Groups adjacent consecutive same-chapter verses
'   from a validated verse-level canonical reference
'   list into contiguous canonical ranges.
' Input:
'   Refs - Ordered Collection of validated verse-level
'          canonical strings from Stage 15.
'          Each item is of the form "BookName ch:v".
' Output:
'   New Collection of canonical range strings.
'   Adjacent verses within same book and chapter are
'   grouped; non-adjacent, cross-chapter, and
'   cross-book verses remain as separate items.
'=====================================================
Public Function BuildCanonicalRanges(ByVal Refs As Collection) As Collection
    On Error GoTo PROC_ERR

    Dim Result As New Collection
    Dim i As Long
    Dim n As Long

    n = Refs.count
    If n = 0 Then
        Set BuildCanonicalRanges = Result
        GoTo PROC_EXIT
    End If
    '------------------------------------------
    ' Parse all refs into parallel arrays
    '------------------------------------------
    Dim bookNames()  As String
    Dim chapters()   As Long
    Dim verses()     As Long
    Dim parsedOk()   As Boolean

    ReDim bookNames(1 To n)
    ReDim chapters(1 To n)
    ReDim verses(1 To n)
    ReDim parsedOk(1 To n)

    Dim refStr As String
    For i = 1 To n
        refStr = Refs(i)
        parsedOk(i) = ParseCanonicalRef( _
                          refStr, _
                          bookNames(i), _
                          chapters(i), _
                          verses(i))
    Next i
    '------------------------------------------
    ' Walk through refs, accumulating runs
    '------------------------------------------
    i = 1
    Do While i <= n
        '------------------------------------------
        ' If this ref could not be parsed as a
        ' chapter:verse ref, emit it unchanged
        '------------------------------------------
        If Not parsedOk(i) Then
            Result.Add Refs(i)
            i = i + 1
        Else
            '------------------------------------------
            ' Start a potential run
            '------------------------------------------
            Dim runStart As Long
            runStart = i
            '------------------------------------------
            ' Extend the run while adjacent.
            ' NOTE: VBA And is not short-circuit - guard
            ' the array accesses with a Do While i < n
            ' outer check and inner If/Exit Do tests.
            '------------------------------------------
            Do While i < n
                If Not parsedOk(i + 1) Then Exit Do
                If bookNames(i + 1) <> bookNames(i) Then Exit Do
                If chapters(i + 1) <> chapters(i) Then Exit Do
                If verses(i + 1) <> verses(i) + 1 Then Exit Do
                i = i + 1
            Loop
            '------------------------------------------
            ' Emit the run as a range or single ref
            '------------------------------------------
            If i > runStart Then
                Result.Add bookNames(runStart) & " " & _
                           chapters(runStart) & ":" & verses(runStart) & _
                           "-" & _
                           chapters(i) & ":" & verses(i)
            Else
                Result.Add Refs(i)
            End If

            i = i + 1
        End If
    Loop

    Set BuildCanonicalRanges = Result

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure BuildCanonicalRanges of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

'=====================================================
' FormatCanonicalString
' Stage 17
'   Converts a validated canonical range collection
'   into a single formatted SBL-style reference string.
' Input:
'   Refs - Collection of canonical range strings
'          from Stage 16 (or Stage 15).
'          Items are "BookName ch:v" or
'          "BookName ch:v-ch:v".
' Output:
'   Single formatted SBL reference string.
'   Book names suppressed after first occurrence per
'   contiguous group.  Chapter numbers suppressed when
'   same as previous item.  Commas separate same-chapter
'   items; semicolons separate chapter and book changes.
'=====================================================
Public Function FormatCanonicalString(ByVal Refs As Collection) As String
    On Error GoTo PROC_ERR

    If Refs.count = 0 Then
        FormatCanonicalString = ""
        GoTo PROC_EXIT
    End If

    Dim Result    As String
    Dim prevBook  As String
    Dim prevEndCh As Long
    Dim item      As Variant
    Dim s         As String
    Dim bk        As String
    Dim sCh       As Long
    Dim sV        As Long
    Dim eCh       As Long
    Dim eV        As Long
    Dim ch        As Long
    Dim v         As Long
    Dim IsRange   As Boolean
    Dim sep       As String
    Dim showBook  As Boolean
    Dim showCh    As Boolean
    Dim fragment  As String

    Result = ""
    prevBook = ""
    prevEndCh = 0

    For Each item In Refs
        s = CStr(item)
        bk = ""
        sCh = 0: sV = 0: eCh = 0: eV = 0

        IsRange = ParseCanonicalRange(s, bk, sCh, sV, eCh, eV)

        If Not IsRange Then
            ch = 0: v = 0
            If ParseCanonicalRef(s, bk, ch, v) Then
                sCh = ch: sV = v
                eCh = ch: eV = v
            Else
                ' Unparseable (chapter-only ref) - pass through
                If Len(Result) > 0 Then Result = Result & "; "
                Result = Result & s
                GoTo NextItem
            End If
        End If

        ' Determine separator and what context to suppress
        If Len(prevBook) = 0 Then
            sep = ""
            showBook = True
            showCh = True
        ElseIf bk <> prevBook Then
            sep = "; "
            showBook = True
            showCh = True
        ElseIf sCh <> prevEndCh Then
            sep = "; "
            showBook = False
            showCh = True
        Else
            sep = ", "
            showBook = False
            showCh = False
        End If

        ' Build the reference fragment
        fragment = ""
        If showBook Then fragment = bk & " "
        If showCh Then
            fragment = fragment & sCh & ":" & sV
        Else
            fragment = fragment & sV
        End If

        ' Append range end when this item spans more than one verse
        If eCh <> sCh Or eV <> sV Then
            If eCh = sCh Then
                ' Same-chapter range: suppress repeated chapter
                fragment = fragment & "-" & eV
            Else
                ' Cross-chapter range: show full end ch:v
                fragment = fragment & "-" & eCh & ":" & eV
            End If
        End If

        Result = Result & sep & fragment

        prevBook = bk
        prevEndCh = eCh

NextItem:
    Next item

    FormatCanonicalString = Result

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure FormatCanonicalString of Class aeSBL_Citation_Class"
    Resume PROC_EXIT
End Function

Public Sub xxxTest_AllBookAliases_STRICT()
    Dim aMap As Object
    Dim books As Object
    Dim k As Variant
    Dim BookID As Long
    Dim canonicalActual As String
    Dim failures As Long

    Set aMap = GetBookAliasMap
    Set books = GetCanonicalBookTable

    Debug.Print "=== STRICT Alias Validation ==="

    For Each k In aMap.Keys
        On Error GoTo AliasFail

        canonicalActual = ResolveAlias(CStr(k), BookID)

        ' 1. BookID must be in canonical range
        If BookID < 1 Or BookID > 66 Then
            Debug.Print "FAIL (INVALID BOOK ID): "; k; " ? "; BookID
            failures = failures + 1
            GoTo NextAlias
        End If

        ' 2. BookID must exist in canonical table
        If Not books.Exists(BookID) Then
            Debug.Print "FAIL (MISSING CANONICAL): "; k; " ? "; BookID
            failures = failures + 1
            GoTo NextAlias
        End If

        ' 3. Canonical name must match
        If canonicalActual <> books(BookID)(1) Then
            Debug.Print "FAIL (NAME MISMATCH): "; k; _
                        " ? "; canonicalActual; _
                        " (expected "; books(BookID)(1); ")"
            failures = failures + 1
        End If

NextAlias:
        On Error GoTo 0
    Next k

    If failures = 0 Then
        Debug.Print "PASS: All aliases are canonical and valid."
    Else
        Debug.Print "FAILURES: "; failures
    End If
    Exit Sub

AliasFail:
    Debug.Print "FAIL (ERROR): "; k; " > "; Err.Description
    Err.Clear
    failures = failures + 1
    Resume NextAlias
End Sub



