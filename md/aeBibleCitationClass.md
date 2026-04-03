# Deterministic Structural Parser (DSP) for SBL Bible Citation

The parser correctly handles three expansion classes:
- Book-only expansion
- Single-chapter book normalization
- Verse shorthand expansion

Those three transformations together form the complete canonicalization layer of the reference parser.

---

## parser PIPELINE

- Stage 1  LexicalScan
- Stage 2  ResolveAlias
- Stage 3  InterpretStructure
- Stage 4  ValidateCanonical
- Stage 5  RewriteSingleChapterRef
- Stage 6  ComposeCanonicalReference
- Stage 7  FinalParser

Stages 8-12 implement list and range extensions.

---

## SCRIPTURE REFERENCE PARSER - FORMAL CONTRACT

1. `ParseReference(rawInput As String)` is a total function: It returns exactly one ScriptureRef for every input string.
2. The parser never raises user-facing runtime errors.
3. All user input errors are reported via ScriptureRef: `IsValid=False`, `ErrorCode<>0`, `ErrorText<>""`.
4. If `IsValid=True` then:
   - `BookID => [1..66]`
   - `Chapter >= 1`
   - Canonical output always includes a verse number.
     - `Verse >= 0`
     - `Verse = 0` means the user did not specify a verse.
     - The canonicalization layer expands this to verse 1.

   Examples:

   | Input | Internal State | Canonical Output |
   |-------|---------------|-----------------|
   | Romans 8 | Chapter=8  Verse=0 | Romans 8:1 |
   | John | Chapter=1  Verse=1 | John 1:1 |
   | Jude 5 | Chapter=0  Verse=5 | Jude 1:5 |

   - `ErrorCode=0`
   - `ErrorText=""`
   - `NormalizedRef<>""`

5. If `IsValid=False` then: `NormalizedRef=""`, `ErrorCode<>0`.
6. Raw input is not stored in the returned ScriptureRef. Only normalized canonical output is preserved.
7. Each stage (1-7) performs a single responsibility and does not mutate external state.
8. No stage performs responsibilities assigned to another stage.
9. Canonical metadata is authoritative and immutable.

   **Canonical Book ID Mapping**

   | ID | Book | | ID | Book |
   |----|------|-|----|------|
   | 1 | Genesis | | 34 | Nahum |
   | 2 | Exodus | | 35 | Habakkuk |
   | 3 | Leviticus | | 36 | Zephaniah |
   | 4 | Numbers | | 37 | Haggai |
   | 5 | Deuteronomy | | 38 | Zechariah |
   | 6 | Joshua | | 39 | Malachi |
   | 7 | Judges | | 40 | Matthew |
   | 8 | Ruth | | 41 | Mark |
   | 9 | 1 Samuel | | 42 | Luke |
   | 10 | 2 Samuel | | 43 | John |
   | 11 | 1 Kings | | 44 | Acts |
   | 12 | 2 Kings | | 45 | Romans |
   | 13 | 1 Chronicles | | 46 | 1 Corinthians |
   | 14 | 2 Chronicles | | 47 | 2 Corinthians |
   | 15 | Ezra | | 48 | Galatians |
   | 16 | Nehemiah | | 49 | Ephesians |
   | 17 | Esther | | 50 | Philippians |
   | 18 | Job | | 51 | Colossians |
   | 19 | Psalms | | 52 | 1 Thessalonians |
   | 20 | Proverbs | | 53 | 2 Thessalonians |
   | 21 | Ecclesiastes | | 54 | 1 Timothy |
   | 22 | Solomon | | 55 | 2 Timothy |
   | 23 | Isaiah | | 56 | Titus |
   | 24 | Jeremiah | | 57 | Philemon |
   | 25 | Lamentations | | 58 | Hebrews |
   | 26 | Ezekiel | | 59 | James |
   | 27 | Daniel | | 60 | 1 Peter |
   | 28 | Hosea | | 61 | 2 Peter |
   | 29 | Joel | | 62 | 1 John |
   | 30 | Amos | | 63 | 2 John |
   | 31 | Obadiah | | 64 | 3 John |
   | 32 | Jonah | | 65 | Jude |
   | 33 | Micah | | 66 | Revelation |

10. Validation is deterministic: first failure wins.
11. Formatting is canonical SBL-style and alias-independent.
12. Only internal invariant violations may raise the Err.Raise error.
13. Scripture structure is defined entirely by metadata. The parser contains no hard-coded knowledge of Bible structure. All information about books, chapters, and verse counts is stored in metadata tables.

    Examples of metadata:
    - Canonical book list (BookID → BookName)
    - Number of chapters per book
    - Maximum verses per chapter

    The parser reads this metadata when validating references. Therefore validation rules such as:
    - "Genesis has 50 chapters"
    - "Romans 16 has 27 verses"

    are not implemented in code. They are enforced by metadata lookups.

    Example validation flow:
    ```
    Input:  Romans 16:30
    Parser checks metadata:
        GetMaxVerse(BookID=45, Chapter=16) ? 27
    Because 30 > 27, the reference is rejected.
    ```

    Benefits of metadata-driven structure:
    - Parser code remains generic and simple.
    - Structural data can be corrected without changing code.
    - The same parser can support alternate canons or translations by swapping metadata tables.

### Extension Layer Invariant

 *** Later stages are are described in the ***
 ***     Extension Layer Classification    ***

- Stages 1-7 parse atomic references only.
- Stages 8-12 must not modify the behavior of the atomic parser.

---

## design Goals

1. VBA coding often causes Type Mismatch and Bounds errors due to 0 vs. 1 based arrays. Bible books in the Protestant Canon are numbered 1-66. This design will enforce 1-Based arrays throughout the system and use Assert statements to raise errors immediately for any 0-Based array usage. It makes sure that the code and documentation are in synch, so strict task separation is used. The test harness should follow the same principle, and documentation be updated to include code routine names as part of the development strategy.

1a. We are building:
   - A disciplined, testable VBA module(s)
   - With documentation synchronized to code
   - With stage-specific harness testing
   - With architectural clarity as a primary goal
   - For a validation-heavy parser, stage isolation is gold.

   That creates:
   - A formal trace from documentation => procedure
   - A test harness entry point per stage
   - A maintenance roadmap.

### Parser Workflow (Pipeline Overview)

The reference engine follows a strict multi-stage pipeline. Each stage has a single responsibility. No stage may perform work assigned to another stage.

```
DETERMINISTIC STRUCTURAL PARSER (DSP)
--------------------------------------
Stage 1   Normalize
Stage 2   LexicalScan
Stage 3   ResolveAlias
Stage 4   InterpretStructure
Stage 5   ValidateCanonical
Stage 6   FormatCanonical
Stage 7   EmitResult

EXTENSION PARSING LAYER
-----------------------
Stage 8   ListDetection
Stage 9   RangeDetection
Stage 10  RangeComposition
Stage 11  ListComposition
Stage 12  ExtendedParse

CONTEXT RESOLUTION LAYER
------------------------
Stage 13  ContextualShorthand

CANONICAL OUTPUT LAYER
----------------------
Stage 14  CanonicalCompression
```

The atomic parser is a mathematically deterministic function:
```
ParseReference : String -> ScriptureRef
```
The extended parser becomes:
```
ParseReferenceExtended : String -> ScriptureRef | ScriptureRange | ScriptureList
```

---

## design Principle

Later stages may depend on earlier stages. Earlier stages must never depend on later stages.

**Invariant:**
- No user-facing errors are raised during parsing.
- Only internal invariant violations may Err.Raise.

### Stage 1: Input Normalization (Normalize)

**PURPOSE:**
Normalize raw input string into a canonical internal working form suitable for tokenization.

**INPUT:**
```
rawInput As String
```

**Output:**
```
normalizedInput As String
```

**Responsibilities:**
- Trim leading/trailing whitespace
- Collapse multiple internal spaces to single space
- Normalize Unicode punctuation (e.g., smart quotes, en-dash)

**NON-Responsibilities:**
- No alias resolution
- No numeric parsing
- No structural interpretation
- No validation
- No logging or diagnostic preservation

**design DECISION:**
Raw input is NOT preserved. The parser is intentionally stateless and does not retain original user input.

**FAILURE MODEL:**
- Never raises user-facing errors.
- Returns normalized string (possibly empty).

**CURRENT IMPLEMENTATION STATUS:**
- Trim whitespace Y
- Preserve original input N (intentionally omitted)
- Collapse multiple spaces N (pending)
- Unicode normalization N (pending)

**NOTE:**
Currently partially implemented inline in ParseReferenceStub. Will be refactored to:
```vb
NormalizeInput(rawInput As String) As String
```

### Stage 2: Lexical Tokenization (Tokenize)

**PURPOSE:**
Convert normalized input string into primitive lexical tokens. This stage performs ZERO semantic interpretation.

**INPUT:**
```
normalizedInput As String
    (Output of Stage 1: NormalizeInput)
```

**current Implementation:**
```vb
Public Function TokenizeReference(ByVal normalizedInput As String) As LexTokens
```

**current behavior:**
- Splits on single space
- First token => RawAlias
- Second token (if present) => numeric reference block
- Detects colon separator
- Extracts up to two numeric substrings

**Output structure:**
```vb
Type LexTokens
    RawAlias    As String
    Num1        As Long
    Num2        As Long
    HasColon    As Boolean
    Num1IsValid As Boolean
    Num2IsValid As Boolean
End Type
```

**NUMERIC HANDLING RULE:**
- Numeric conversion must NEVER raise runtime errors.
- Use safe parsing (e.g., IsNumeric check or guarded CLng).
- If conversion fails:
  - NumXIsValid = False
  - NumX = 0

Stage 2 does NOT determine whether values are canonically valid — only whether they are numeric.

**LIMITATIONS (INTENTIONAL - STAGE 3 WILL HANDLE):**
- Multi-word book names (e.g., "1 John", "Song of Solomon")
- Range detection (1-3)
- List detection (1,3,5)
- Canonical validation of chapter/verse
- Alias resolution

**design RULE:**
Stage 2 extracts lexical structure only. Stage 2 NEVER raises runtime errors for malformed input. Stage 3 interprets meaning.

**ASSUMPTION:**
Stage 1 guarantees single-space separation.

**FAILURE MODEL:**
- Invalid numeric text does NOT raise error.
- Invalid numeric text is structurally flagged.
- Structural errors are interpreted in Stage 3.

**Test harness:**
```vb
Public Sub Test_TokenizeReference()
```

**STATUS:**
- Y Book token extraction
- Y Numeric extraction (single or colon pair)
- Y Colon detection
- Y Safe numeric parsing (required)
- N Range/list parsing (future)
- N Multi-token book support (future)

### Stage 3: Alias Resolution (Resolve Book Identity)

**PURPOSE:**
Resolve RawAlias into canonical BookID. Preserve lexical numeric tokens without interpreting them.

**INPUT:**
```
LexTokens
    RawAlias    As String
    Num1        As Long
    Num2        As Long
    HasColon    As Boolean
    Num1IsValid As Boolean
    Num2IsValid As Boolean
```

**Output:**
```vb
Type ParsedRef
    BookID      As Long   ' 0 if alias unresolved
    AliasFound  As Boolean
    ' Forwarded lexical state (UNINTERPRETED)
    Num1        As Long
    Num2        As Long
    HasColon    As Boolean
    Num1IsValid As Boolean
    Num2IsValid As Boolean
    ' Structural fields NOT assigned in Stage 3
    Chapter     As Long   ' Uninitialized here
    Verse       As Long   ' Uninitialized here
End Type
```

**Responsibilities:**
1. Normalize RawAlias (case-insensitive lookup).
2. Lookup alias in canonical alias dictionary.
3. If found:
   - BookID = canonical ID (1-66)
   - AliasFound = True

   If NOT found:
   - BookID = 0
   - AliasFound = False
4. Forward Num1/Num2/HasColon and validity flags unchanged.

**NON-Responsibilities:**
- Does NOT assign Chapter or Verse.
- No structural interpretation.
- No chapter/verse validation.
- No canonical bounds checking.
- No formatting.

**FAILURE PROPAGATION RULE:**
Stage 3 NEVER raises runtime errors. Alias resolution failure does NOT terminate parsing.

Instead:
- AliasFound = False
- BookID = 0

Lexical numeric failures are preserved and forwarded.

**design Invariant:**
BookID = 0 is the ONLY legal representation of unresolved book identity. No negative BookID values.

**design RULE:**
Stage 3 establishes canonical identity ONLY. Structural meaning begins in Stage 4.

### Stage 4: Structural Interpretation (Determine structural meaning)

**PURPOSE:**
Convert lexical numeric tokens into structural Chapter/Verse values using canonical metadata (BookID).

**INPUT:**
```
ParsedRef from Stage 3:
    BookID
    AliasFound
    Num1
    Num2
    Num1IsValid
    Num2IsValid
    HasColon
```

**OUTPUT** (updates ParsedRef):
```
Chapter As Long
Verse   As Long
```

**Responsibility:**
Stage 4 is the SOLE owner of structural interpretation. It assigns Chapter and Verse exactly once.

**PRECONDITION:**
None. Must tolerate:
- BookID = 0
- Invalid numeric tokens

**LOGIC:**
```vb
' Default structural assignment (no metadata reliance)
    Chapter = Num1
    Verse = 0
If Num1IsValid = False And HasColon = False Then
' Book-only reference
    Chapter = 1
    Verse = 1
' If colon present, explicit Chapter:Verse structure
If HasColon = True Then
    Verse = Num2
    Exit Stage
End If
' No colon present
' Metadata may be consulted ONLY if alias resolved
If AliasFound = True And Num1IsValid = True Then
    If GetChapterCount(BookID) = 1 Then
        ' Single-chapter book (e.g., Jude)
        Chapter = 1
        Verse = Num1
    End If
End If
```

**STRUCTURAL rules:**
- Stage 4 NEVER changes Num1 or Num2.
- Stage 4 NEVER converts invalid numeric tokens into valid ones.
- Stage 4 NEVER performs range checking.
- Stage 4 assigns Chapter/Verse exactly once.

**STRUCTURAL ASSERTIONS** (not semantic validation):
- Chapter >= 0
- Verse >= 0

**IMPORTANT:**
Stage 4 determines structure only. Stage 5 validates:
- AliasFound = True
- Num1IsValid / Num2IsValid
- Chapter >= 1
- Chapter <= MaxChapter(BookID)
- Verse bounds

**design RULE:**
Structural interpretation must not depend on chapter/verse bounds. Metadata (GetChapterCount) may be consulted ONLY when BookID is valid.

#### Book-Only Reference Handling

If a reference contains only a book alias and no numeric component, the parser defaults to:
- Chapter = 1
- Verse = 1

**Example:**
```
Input:  "John"
Output: "John 1:1"
Input:  "1 Jn"
Output: "1 John 1:1"
```

**Rationale:**
This behavior matches common Bible software navigation conventions and provides a safe canonical anchor for book-level references.

**Implementation:**
When refPart Is Empty, It Is internally replaced with "1:1" before numeric validation.

### Stage 5: Canonical Validation (Validate reference)

**PURPOSE:**
Determine whether the structurally interpreted reference is canonically valid.

**INPUT:**
```
ParsedRef from Stage 4:
    BookID      As Long
    AliasFound  As Boolean
    Chapter     As Long
    Verse       As Long
    HasColon    As Boolean
    Num1IsValid As Boolean
    Num2IsValid As Boolean
```

**Output:**
```
IsValid   As Boolean
ErrorCode As Long
ErrorText As String
```

**Responsibility:**
- Validate canonical correctness only.
- Do NOT modify Chapter or Verse.
- Do NOT format output.
- Do NOT raise runtime errors.

**VALIDATION ORDER (STRICT PRECEDENCE):**
1. Alias resolution
2. Lexical numeric validity
3. Structural minimums (>=1)
4. Chapter upper bound
5. Verse upper bound

FIRST FAILURE WINS.

### Stage 6: Canonical Normalization (Output Formatting)

**PURPOSE:**
Produce canonical SBL-style reference string including canonical Book Name.

**INPUT:**
```
BookID  As Long
Chapter As Long
Verse   As Long
```

**PRECONDITION:**
Stage 5 validation has succeeded. Therefore the following invariants hold:
- `BookID  => [1..66]`
- `Chapter >= 1`
- `Chapter <= MaxChapter(BookID)`
- If Verse = 0: Reference is chapter-only
- If Verse > 0: `Verse = MaxVerse(BookID, Chapter)`

**Output:**
```
NormalizedRef As String
```

**LOGIC:**
```vb
CanonicalBookName = GetCanonicalBookName(BookID)
If Verse = 0 Then
    ' Chapter-only reference
    NormalizedRef = CanonicalBookName & " " & _
                    CStr(Chapter)
Else
    ' Chapter + Verse reference
    NormalizedRef = CanonicalBookName & " " & _
                    CStr(Chapter) & ":" & CStr(Verse)
End If
```

**Format rules:**
- Exactly one space between BookName and Chapter.
- No trailing spaces.
- No zero-padding.
- No leading zeros.
- No alias text allowed.

**RESTRICTIONS:**
- Do NOT perform validation here.
- Do NOT modify structural values.
- Do NOT consult alias dictionary.
- Do NOT emit output if Stage 5 failed.

**design RULE:**
Stage 6 is a pure formatting function. Given valid canonical input, output is uniquely determined.

### Stage 7: Structured Result Object (Emit Immutable Result)

**PURPOSE:**
Construct and return the final immutable ScriptureRef result object. `ParseReference()` must always return a ScriptureRef, never Nothing, never uninitialized.

**INPUT:**
Final state from Stage 5 and (if valid) Stage 6.

**Output:**
```vb
Public Type ScriptureRef
    BookID        As Long
    Chapter       As Long
    Verse         As Long
    NormalizedRef As String
    IsValid       As Boolean
    ErrorCode     As Long
    ErrorText     As String
End Type
```

**CONSTRUCTION RULE:**
ScriptureRef is constructed exactly once inside `ParseReference()`. No other procedure may partially construct it.

**state INVARIANTS:**
```
If IsValid = True Then
    BookID        => [1..66]
    Chapter       >= 1
    Verse         >= 0
    ErrorCode = 0
    ErrorText = ""
    NormalizedRef <> ""
If IsValid = False Then
    NormalizedRef = ""
    ErrorCode     <> 0
    ErrorText     <> ""
    BookID may be 0
    Chapter and Verse may be 0
```

**ILLEGAL STATES (MUST NEVER OCCUR):**
- `IsValid = True And ErrorCode <> 0`
- `IsValid = True And NormalizedRef = ""`
- `IsValid = False And ErrorCode = 0`

**FAILURE MODEL:**
- Parser NEVER raises user-facing runtime errors.
- All user input errors are reported via:
  - `IsValid = False`
  - `ErrorCode <> 0`
  - `ErrorText <> ""`
- Only internal invariant violations (e.g., corrupted canonical metadata) may raise Err.Raise.
- Because RawInput is discarded: ErrorText must never embed raw user input.

**design RULE:**
Stage 7 performs:
- No parsing
- No validation
- No metadata lookup

It only packages the final canonical state. ScriptureRef is immutable after return.

---

## SBL Scripture Citation - Structural EBNF — Aligned to 7-Stage Deterministic Parser

**PURPOSE:**
Defines structural syntax only. No semantic validation is expressed here. Canonical bounds enforcement occurs in Stage 5.

**NOTE:**
This grammar describes lexical and structural form. It does NOT:
- Validate book identity
- Validate chapter/verse bounds
- Normalize aliases
- Enforce canonical metadata

### Top-level

```ebnf
Citation
   ::= WS? Reference (WS? RefSep WS? Reference)* WS?
RefSep
   ::= ";" | ","
```

### core Reference

```ebnf
Reference
   ::= BookRef
    |  BookRef WS ChapterSpec
BookRef
   ::= Prefix? WS? BookName
```

Book-only references are permitted. When no ChapterSpec is present, the reference is normalized to the first verse of the book.

**Examples:**
```
John      -> John 1:1
Romans    -> Romans 1:1
Jude      -> Jude 1:1
```
This normalization occurs in Stage 6 (canonical formatting / rewrite)

```ebnf
prefix
   ::= ArabicPrefix | RomanPrefix
ArabicPrefix
   ::= "1" | "2" | "3"
RomanPrefix
   ::= "I" | "II" | "III"
```

**NOTE:**
Prefix may be adjacent to BookName (e.g., "1John", "IJohn")

```ebnf
bookName
   ::= BookWord (WS BookWord)*
BookWord
   ::= Letter+ "."?
```

**CONSTRAINT** (Structural Only):
- No internal punctuation.
- Trailing period permitted.

### Chapter / Verse Structure

```ebnf
ChapterSpec
   ::= Chapter
    | Chapter ":" VerseSpec
    | ChapterRange
    | Chapter ":" VerseRangeSpec
ChapterRange
   ::= Chapter "-" Chapter
VerseSpec
   ::= VerseItem ("," VerseItem)*
VerseRangeSpec
   ::= VerseRange ("," VerseRange)*
VerseItem
   ::= Verse | VerseRange
VerseRange
   ::= Verse "-" Verse
```

### Atomic Numeric Units

```ebnf
Chapter
   ::= Digit+
Verse
   ::= Digit+ VerseSuffix?
VerseSuffix
   ::= Letter
```

VERSE_SUFFIX | Alphabetic suffix after verse digit

**NOTE:**
VerseSuffix (e.g., "a", "b") is lexically captured. Canonical validity is enforced post-parse.

### lexical Primitives

```ebnf
WS
   ::= " " { " " }
Letter
   ::= "A"..."Z" | "a"..."z"
Digit
   ::= "0"..."9"
```

### Parser Alignment Notes

**stage 2:**
Performs tokenization according to this grammar.

**stage 3:**
Resolves BookRef → BookID via alias dictionary.

**stage 4:**
Interprets structural meaning:
- Chapter-only
- Chapter:Verse
- Single-chapter inference (semantic layer)

**stage 5:**
Enforces semantic constraints:
- AliasFound = True
- Chapter >= 1
- Chapter <= MaxChapter(BookID)
- Verse bounds

**stage 6:**
Produces canonical SBL formatting.

**stage 7:**
Emits immutable ScriptureRef result object.

### Canonical Normal Form (Post-Validation Output)

```
Single reference:
    <CanonicalBookName> <Chapter>
    <CanonicalBookName> <Chapter>:<Verse>
```

Lists and ranges preserve structural form after semantic validation.

**NOTE:** This DSP validates structural syntax only. Semantic correctness is enforced in Stage 5.

---

## Parser Alignment to Implementation (7-Stage Model)

### Stage 1 - Preprocessing / Normalization

**Routine:**
```vb
Private Function Stage1_NormalizeInput( _
    ByVal rawInput As String) As String
```
**Responsibility:**
- Trim leading/trailing whitespace
- Normalize internal whitespace
- Standardize dash characters if required
- Discard raw input after normalization

### Stage 2 - Lexical Tokenization

**Routine:**
```vb
Private Function Stage2_LexicalScan( _
    ByVal normalizedInput As String) As LexTokens
```
**Responsibility:**
- Extract RawAlias
- Extract Num1, Num2
- Detect HasColon
- Detect lexical numeric validity
- Capture VerseSuffix (if applicable)

### Stage 3 - Alias Resolution (Canonical Identity)

**Routine:**
```vb
Private Function Stage3_ResolveAlias( _
    ByVal tokens As LexTokens) As ParsedRef
```
**Responsibility:**
- Resolve RawAlias → BookID
- Set AliasFound flag
- Forward lexical numeric tokens unchanged
- Do NOT assign Chapter/Verse

### Stage 4 - Structural Interpretation

**Routine:**
```vb
Private Sub Stage4_InterpretStructure( _
    ByRef state As ParsedRef)
```
**Responsibility:**
- Assign Chapter and Verse exactly once
- Apply colon structure logic
- Apply single-chapter inference
- Do NOT validate bounds

### Stage 5 - Canonical Semantic Validation

**Routine:**
```vb
Private Sub Stage5_ValidateCanonical( _
    ByRef state As ParsedRef)
```
**Responsibility:**
- Enforce validation matrix
- First-failure-wins ordering
- Assign ErrorCode / ErrorText
- Set IsValid flag
- Do NOT modify structural values

### Stage 6 - Canonical Normalization (Formatting)

**Routine:**
```vb
Private Sub Stage6_FormatCanonical( _
    ByRef state As ParsedRef)
```
**Responsibility:**
- Produce canonical SBL-style string
- Use canonical BookName from metadata
- Execute only if IsValid = True
- Do NOT perform validation

### Stage 7 - Immutable Result Emission

**Routine:**
```vb
Private Function Stage7_EmitResult( _
    ByVal state As ParsedRef) As ScriptureRef
```
**Responsibility:**
- Enforce object state invariants
- Construct final ScriptureRef
- Guarantee total-function return
- Perform no parsing or validation

### Public Entry Point

**Routine:**
```vb
Public Function ParseReference( _
    ByVal rawInput As String) As ScriptureRef
```

**Execution Order:**
```
1. normalized  = Stage1_NormalizeInput(rawInput)
2. tokens      = Stage2_LexicalScan(normalized)
3. state       = Stage3_ResolveAlias(tokens)
4. Stage4_InterpretStructure state
5. Stage5_ValidateCanonical state
6. Stage6_FormatCanonical state
7. ParseReference = Stage7_EmitResult(state)
```

---

## Extension Layer Classification

The core parser (Stages 1-7) is modeled as a Deterministic Structural Parser (DSP) that produces a single atomic ScriptureRef.

 *** Extension stages operate OUTSIDE the DSP and must never ***
 ***      influence the DSP state machine.                   ***

**Invariant:**
- Stages 1-7 must only parse atomic references.
- Extension stages must not modify or bypass core validation logic.

These stages are strictly lexical segmentation layers.
- Stage 8  List Detection
- Stage 9  Range Detection
- Stage 10 RangeComposition     -> builds ranges
- Stage 11 ListComposition      -> builds lists
- Stage 12 ExtendedParse        -> orchestrate pipeline
- Stage 13 ContextualShorthand

**Recursion:**
Stage-12 may call `ParseReferenceExtended()` recursively when processing list segments. This allows nested list and range structures to be parsed without additional grammar rules.

**Responsibility boundaries:**

**stage 8:**
- Detect top-level list separators
- Output ListTokens

**stage 9:**
- Detect reference ranges
- Output RangeTokens

**stage 10:**
Compose atomic ScriptureRef results into:
- ScriptureList
- ScriptureRange

The DSP always processes a single atomic reference.

---

## Stage 8 - List Detection (Extension Layer)

**PURPOSE:**
Detect multiple references separated by list delimiters.

**Supported separators:**
- `,`  comma
- `;`  semicolon

**Examples:**
```
John 3:16,18,20
John 3, 4, 5
John 3:16; 4:1
```

**Output:**
```vb
Type ListTokens
    IsList As Boolean
    Segments() As String
End Type
```

**rules:**
1. Detection is lexical only.
2. No interpretation of structure occurs.
3. Segments are returned exactly as written.
4. If no separator is found, IsList=False.

**Determinism:**
Stage 8 MUST NOT perform:
- alias resolution
- verse validation
- canonical formatting

These remain Stage 3-6 responsibilities.

**List Detection Rule:**
Separators are recognized only at the top lexical level. Stage 8 must not split inside a detected range.

**Example Input:**
```
John 3:16-18,20
```
**Output Segments:**
```
John 3:16-18
20
```
**NOT:**
```
John 3:16
18
20
```

**stage ordering:**
Stage 8 executes before Stage 9. Lists are segmented before range interpretation occurs. Range interpretation is handled in Stage 9. Stage 9 evaluates each segment independently.

---

## Stage 9 - Range Detection (Extension Layer)

**PURPOSE:**
Detect reference ranges using hyphen or en dash.

**Supported separators:**
- `-`  = ASCII hyphen-minus  (ChrW(45))
- `–`  = Unicode en dash     (ChrW(&H2013))

### A Note on Range Delimiter Characters

The parser supports two range delimiters:

**ASCII hyphen-minus:**
- Character: `-`
- Unicode:   U+002D
- VBA:       `ChrW(45)`

**Unicode en dash:**
- Character: `–`
- Unicode:   U+2013
- VBA:       `ChrW(&H2013)`

**NOTE:**
The VBA Immediate Window may not display the en dash correctly. In Git it may appear as a placeholder. Therefore code comparisons should use ChrW values.

**Example:**
```vb
If ch = "-" Or ch = ChrW(&H2013) Then
    ' range delimiter comparison code
End If
```

**Example Input:**
```
John 3:16-18
John 3 - 5
John 3:16-4:2
```

**Output:**
```vb
Type RangeTokens
    IsRange As Boolean
    LeftRaw  As String
    RightRaw As String
End Type
```

**rules:**
1. Detection is lexical only.
2. No interpretation of structure occurs.
3. Left and Right expressions are returned exactly as written.
4. If no range delimiter exists, IsRange=False.

**Determinism:**
Stage 9 MUST NOT perform:
- alias resolution
- verse validation
- canonical formatting

These remain Stage 3-6 responsibilities.

**stage ordering:**
Stage 8 executes before Stage 9.

**Example:**
```
Input:
    John 3:16-18,20
Stage 8 output:
    John 3:16-18
    20
Stage 9 then evaluates each segment independently.
```

**Range Detection Rule:**
A range delimiter must appear after the first numeric token. This prevents false detection in book prefixes such as: `1-2 Samuel`

Stage-9 must remain lexical only and must not interpret structure. Deterministic Structural Parser (DSP) intact because Stage-9 remains outside the core parser.

**Stage 9 Evaluation:**
```
RangeDetection("John 3:16-18") -> Range
RangeDetection("20") -> Not a range
```

---

## Stage 10 - RangeComposition (Extension Layer)

**PURPOSE:**
Construct a ScriptureRange from the tokens produced by Stage 9 (RangeDetection).

Stage-10 performs no lexical parsing. It only composes structured results using the atomic parser.

**Composition Type:**
ScriptureRange

**rules:**
1. Stage-10 must call `ParseReference()` for both sides of the detected range.
2. LeftRaw and RightRaw must be parsed independently.
3. Stage-10 must not modify ScriptureRef.

**Atomic Parser Guarantee:**
`ParseReference()` remains the only function that produces ScriptureRef.

**Example Input:**
```
John 3:16-18
```

**stage 9:**
```
RangeTokens
    LeftRaw = "John 3:16"
    RightRaw = "18"
```

**stage 10:**
```
ScriptureRange
    StartRef -> ScriptureRef(John 3:16)
    EndRef   -> ScriptureRef(John 3:18)
```

---

## Stage 11 - ListComposition (Extension Layer)

**PURPOSE:**
Compose segmented references into a ScriptureList.

Stage-11 operates on the segments produced by Stage 8 (ListDetection).

**Composition Type:**
ScriptureList

**rules:**
1. Each segment must be processed independently.
2. Stage-11 must determine whether a segment is a range using Stage 9.
3. Range segments must be composed using Stage 10.
4. Non-range segments must be parsed using `ParseReference()`.

**Result:**
ScriptureList.Items() may contain:
- ScriptureRef
- ScriptureRange

**Example Input:**
```
John 3:16-18,20
```

**stage 8:**
```
Segments
    John 3:16-18
    20
```

**stage 9:**
```
Segment 1 -> Range
Segment 2 -> Single
```

**stage 11:**
```
ScriptureList
    Item 1 -> ScriptureRange
    Item 2 -> ScriptureRef
```

---

## Post-Parse Canonical Processing (Stages 12-14)

After atomic parsing (Stages 1-7) and structural composition (Stages 8-11), the extension pipeline performs higher-level reference interpretation.

**stage Responsibilities:**
- Stage 12  ExtendedParse — Orchestrates list and range parsing
- Stage 13  ContextualShorthand — Resolves omitted book/chapter context
- Stage 14  CanonicalCompression — Produces minimal canonical reference form

Only after Stage 13 are references guaranteed to represent fully-resolved canonical references.

Some implementations may internally expand references to verse-level triples:
```
(BookID, Chapter, Verse)
```
Such expansion is an internal optimization and is not required by the specification.

Stage 14 operates on the resolved canonical reference set and performs deterministic structural compression.

---

## Stage 12 - ExtendedParse (Extension Entry Point)

**PURPOSE:**
Provide a high-level parser capable of handling lists and ranges. Stage-12 orchestrates the extension pipeline while preserving the atomic parser contract.

**PIPELINE:**
```
Stage 8  ListDetection
Stage 9  RangeDetection
Stage 10 RangeComposition
Stage 11 ListComposition
```

**Return Types:**
- ScriptureRef
- ScriptureRange
- ScriptureList

**Atomic Parser Guarantee:**
Stages 1-7 remain responsible for parsing atomic references only.

**Example Input:**
```
John 3:16-18,20; 4:1-3
```

**Result:**
```
ScriptureList
    Item 1 -> ScriptureRange (3:16-3:18)
    Item 2 -> ScriptureRef   (3:20)
    Item 3 -> ScriptureRange (4:1-4:3)
```

---

## Stage 13 - Contextual Shorthand Expansion (Post-Parser Context Layer)

**Public API:**
```vb
ComposeList(ByVal raw As String) As Collection
```

**PURPOSE:**
Parses a list of scripture references from a string and returns a collection of canonical reference strings.

**Supported Forms:**
- Single references: `"John 3:16"`
- Ranges:
  - Same chapter: `"John 3:16-18"`
  - Cross chapter: `"John 3:16-4:2"`
- Lists: `"John 3:16, 18; 4:1"`
- Contextual shorthand:
  - Missing book    -> inherited from previous segment
  - Missing chapter -> inherited from current context

**Example:**
```
"3:16-4:2, 5"
    -> "John 3:16-4:2"
    -> "John 4:5"
```

### context state

During processing the engine maintains:
- currentBook
- currentChapter

Context is updated after each resolved segment.

### processing Order

Segments are processed left-to-right. Context updates occur after each resolved segment.

### Range Context Rules

After a range resolves, the context chapter becomes the ending chapter of the range.

**Example:**
```
"John 3:16-4:2, 5"
Range resolved:
    John 3:16-4:2
context becomes:
    book = John
    Chapter = 4
Next segment:
    5 -> John 4:5
```

### Canonical Output Rules

- Same-chapter ranges collapse chapter repetition: `John 3:16-18`
- Cross-chapter ranges preserve both chapters: `John 3:16-4:2`
- Single verses always include book and chapter: `John 3:16`

### Examples

**Input:** `"John 3:16, 18, 20-22"`

**Output:**
```
"John 3:16"
"John 3:18"
"John 3:20-22"
```

**Input:** `"3:16-4:2, 5"`

**Output:**
```
"John 3:16-4:2"
"John 4:5"
```

**Input:** `"Romans 8; 9"`

**Output:**
```
"Romans 8"
"Romans 9"
```

### Implementation

```
Stage 11  -> ListTokens (segment structure)
Stage 12  -> ScriptureRef / ScriptureRange objects
Stage 13  -> Contextual shorthand resolution
```

**Output:**
Collection of canonical reference strings

---

## Stage 14 - Canonical Compression

**Input:**
A resolved set of canonical references produced after contextual shorthand expansion (Stage 13).

**Output:**
Minimal canonical citation form.

Stage 14 performs deterministic compression of the expanded canonical verse stream. Adjacent verses are collapsed into ranges while preserving canonical ordering and semantic meaning. This stage is purely structural and does not alter interpretation of references.

### Canonical Output Grammar

```ebnf
CanonicalCitation
   ::= CanonicalBookRef
    | CanonicalBookRef (";" WS? CanonicalBookRef)*
CanonicalBookRef
   ::= BookName WS CanonicalChapterSpec
CanonicalChapterSpec
   ::= CanonicalChapterUnit
    | CanonicalChapterUnit (";" WS? CanonicalChapterUnit)*
CanonicalChapterUnit
   ::= Chapter ":" CanonicalVerseSpec
CanonicalVerseSpec
   ::= CanonicalVerseItem
    | CanonicalVerseItem ("," CanonicalVerseItem)*
CanonicalVerseItem
   ::= Verse
    | Verse "-" Verse
```

### Canonical Compression Rules

1. Sequential verses collapse into ranges.
   ```
   John 3:16
   John 3:17
   John 3:18
   ->  John 3:16-18
   ```
2. Non-sequential verses remain comma separated.
   ```
   John 3:16
   John 3:18
   ->  John 3:16,18
   ```
3. Chapter boundaries are never merged.
   ```
   Genesis 1:31
   Genesis 2:1
   ->  Genesis 1:31; 2:1
   ```
4. Compression must preserve canonical ordering: BookID, Chapter, Verse.
5. Compression is deterministic and lossless.

---

## Stage 15 - Canonical Validation

**PURPOSE:**
Ensure the final canonical reference set contains only valid, in-range scripture references after all expansion, ordering, and compression operations have completed.

**Position in Pipeline:**
Stage 15 executes AFTER Stage 14 Canonical Compression. At this point, the reference set is:
- Fully expanded
- Deduplicated
- Ordered
- Canonically compressed

Stage 15 performs the final integrity check before output.

**Responsibilities:**

1. **Remove Invalid Chapters** — Eliminate references to chapters that do not exist in the specified book.
   ```
   Gen 51      -> removed
   Matt 29     -> removed
   ```

2. **Remove Invalid Verses** — Eliminate verse numbers exceeding chapter limits.
   ```
   Gen 1:999   -> removed
   Jude 1:50   -> removed
   ```

3. **Clamp Range Boundaries** — If a range extends past valid scripture bounds, trim the range to the last valid verse.
   ```
   Gen 1:1-999
   becomes:
   Gen 1:1-31
   ```

4. **Remove Empty Ranges** — If validation removes all verses in a range, discard the range entirely.
   ```
   Matt 29:1-10 -> removed
   ```

5. **Preserve Canonical Order** — Validation must NOT reorder references. The Stage 12 ordering must remain intact.

6. **Preserve Compression** — Stage 15 must not expand ranges. It may only: Trim, Remove, or Keep.

**Input:**
Packed canonical verse map (compressed, ordered)

**Output:**
Fully validated packed canonical verse map

**Guarantees After Stage 15:**
- All books valid
- All chapters valid
- All verses valid
- No empty ranges
- Canonically ordered
- Canonically compressed
- Engine-safe output

**design rules:**
- No allocation of new structures
- Operate directly on packed verse map
- O(n) scan across canonical set
- Validation tables must be constant lookup

**required Data:**
- ChaptersPerBook (book)
- VersesPerChapter(book, chapter)

**Example Input:**
```
Gen 1:1-999, 51:1-10, Exod 1:1-5
```

**After Stage 15:**
```
Gen 1:1-31, Exod 1:1-5
```

**Summary:**
Stage 15 is the final correctness gate. After this stage, the parser output is guaranteed to represent only valid scripture references.

---

## Stage 16 - Canonical Range Builder

**PURPOSE:**
Convert the validated verse-level canonical reference set into contiguous canonical ranges prior to string formatting.

**Position in Pipeline:**
Stage 16 executes AFTER Stage 15 Canonical Validation and BEFORE Stage 17 Canonical String Formatter. At this point the reference set is:
- Fully expanded
- Deduplicated
- Ordered
- Canonically valid

Stage 16 groups adjacent verses into canonical ranges.

**Responsibilities:**

1. **Detect Contiguous Verses** — Identify adjacent verses within same book and chapter.
   ```
   John 3:16
   John 3:17
   John 3:18
   becomes:
   John 3:16-3:18
   ```

2. **Preserve Non-Adjacent Verses**
   ```
   John 3:16
   John 3:18
   remains:
   John 3:16
   John 3:18
   ```

3. **Stop Range at Chapter Boundary**
   ```
   John 3:36
   John 4:1
   remains:
   John 3:36
   John 4:1
   ```

4. **Stop Range at Book Boundary**
   ```
   John 3:16
   Romans 8:1
   remains separate
   ```

5. **Produce Minimal Canonical Ranges**
   ```
   16,17,18,19
   becomes:
   16-19
   ```

**Input:**
Ordered validated verse list

**Output:**
Canonical range collection

**design rules:**
- Single forward pass
- O(n)
- No reordering
- No expansion
- Only grouping

**Guarantees After Stage 16:**
- Contiguous verses grouped
- Minimal ranges created
- Canonical order preserved
- No cross-chapter merging
- No cross-book merging

**Summary:**
Stage 16 builds canonical ranges from validated verse-level references so Stage 17 can perform final string formatting.

---

## Stage 17 - Canonical String Formatter

**PURPOSE:**
Convert the validated canonical reference ranges into a single properly formatted SBL-style reference string for display, logging, export, and round-trip parsing.

**Position in Pipeline:**
Stage 17 executes AFTER:
- Stage 15 - Canonical Validation

At this point the reference set is:
- Fully expanded
- Deduplicated
- Canonically ordered
- Canonically compressed into ranges
- Fully validated

Stage 17 performs formatting only. It MUST NOT change the reference structure.

**Responsibilities:**

1. **Render Canonical Ranges** — Convert each canonical range into text.
   ```
   John 3:16
   John 3:16-3:18
   Romans 8:1-8:2
   ```

2. **Suppress Repeated Book Names** — The book name is printed once per contiguous group.
   ```
   John 3:16
   John 3:18
   becomes:
   John 3:16, 18
   ```

3. **Suppress Repeated Chapter Numbers** — Chapter number appears once when ranges remain within the same chapter.
   ```
   John 3:16
   John 3:18
   becomes:
   John 3:16, 18
   ```

4. **Use Comma for Same-Chapter Separation**
   ```
   John 3:16
   John 3:18
   Output:
   John 3:16, 18
   ```

5. **Use Semicolon for Chapter Breaks**
   ```
   John 3:16-18
   John 4:1-3
   Output:
   John 3:16-18; 4:1-3
   ```

6. **Use Semicolon for Book Breaks**
   ```
   John 3:16
   Romans 8:1
   Output:
   John 3:16; Rom 8:1
   ```

7. **Preserve Canonical Order** — Formatter must not reorder references.

8. **Preserve Canonical Compression** — Formatter must not recompute ranges. It only renders existing canonical ranges.

**Input:**
Validated canonical range collection

**Output:**
Single formatted SBL reference string

**Formatting rules:**
- Same chapter        -> comma
- Chapter change      -> semicolon
- Book change         -> semicolon
- Range separator     -> hyphen or en dash

**Example:**

Input:
```
John 3:16-3:18
John 4:1-4:2
Romans 8:1-8:2
```
Output:
```
John 3:16-18; 4:1-2; Rom 8:1-2
```

**Additional Examples:**

```
Input:   Gen 1:1-1:3
Output:  Gen 1:1-3

Input:   Gen 1:1  /  Gen 1:3
Output:  Gen 1:1, 3

Input:   Gen 1:1  /  Gen 2:1
Output:  Gen 1:1; 2:1

Input:   Gen 1:1  /  Exod 1:1
Output:  Gen 1:1; Exod 1:1
```

**design rules:**
- Pure formatting stage
- No mutation of canonical references
- Single pass O(n)
- Minimal string allocations
- Deterministic output

**Guarantees After Stage 17:**
- Proper SBL formatting
- Minimal repetition
- Canonical ordering preserved
- Canonical compression preserved
- Human-readable output
- Round-trip parser safe

**Summary:**
Stage 17 converts the validated canonical range set into a compact, human-readable SBL formatted reference string suitable for display, storage, and round-trip parsing.

---

## Deterministic Structural Parser (DSP) — Aligned to 7-Stage Parser Architecture

**PURPOSE:**
Defines lexical token stream and structural syntax only.

**SCOPE:**
This DSP:
- Validates structural ordering of tokens
- Detects lists and ranges syntactically

This DSP does NOT:
- Validate canonical book identity
- Enforce chapter/verse bounds
- Perform single-chapter inference
- Rewrite structure
- Normalize aliases

Semantic processing occurs in:
- Stage 3 - Alias Resolution
- Stage 4 - Structural Interpretation
- Stage 5 - Canonical Validation

### 1. Token Stream Definition (Stage 2 Output)

**token Types:**

| Token | Description |
|-------|-------------|
| BOOK_WORD | Alphabetic word, optional trailing `.` |
| PREFIX_ARABIC | `1` \| `2` \| `3` |
| PREFIX_ROMAN | `I` \| `II` \| `III` |
| DIGITS | One or more digits |
| COLON | `:` |
| DASH | `-` |
| COMMA | `,` |
| SEMICOLON | `;` |
| WS | One or more spaces (collapsed) |
| EOF | End-of-input |

**Tokenization rules:**
- Whitespace collapsed to single WS
- DIGITS is greedy
- Case-insensitive BOOK_WORD and PREFIX_ROMAN
- No internal punctuation in BOOK_WORD

### 2. Deterministic State Machine (Structural Only)

**NOTE:**
This is a single-pass left-to-right DSP. State numbering is symbolic.

#### 2.1 State Definitions

| State | Meaning | Accepting |
|-------|---------|-----------|
| S0 | Start | No |
| S1 | Reading numeric prefix | No |
| S2 | Reading book name | Yes* |
| S3 | Expecting chapter digits | No |
| S4 | Reading chapter digits | Yes* |
| S6 | Reading verse digits | Yes* |
| S7 | After dash | No |
| S8 | After comma | No |
| SX | Error | No |

\*Accepting states require EOF or SEMICOLON next. Book-only references (state S2) are normalized to Chapter 1, Verse 1 during canonical formatting.

#### 2.2 Transition Table

```
S0 - Start
  WS              -> S0
  PREFIX_*        -> S1
  BOOK_WORD       -> S2
  otherwise       -> SX

S1 - prefix
  WS              -> S2
  BOOK_WORD       -> S2
  otherwise       -> SX

S2 - Book Name
  BOOK_WORD       -> S2
  WS              -> S3
  SEMICOLON       -> S0
  EOF             -> ACCEPT
  otherwise       -> SX

S3 - Expect Chapter
  DIGITS          -> S4
  otherwise       -> SX

S4 - Chapter (Accepting*)
  DIGITS          -> S4
  COLON           -> S6
  DASH            -> S7
  SEMICOLON       -> S0
  EOF             -> ACCEPT
  otherwise       -> SX

S6 - Verse (Accepting*)
  DIGITS          -> S6
  DASH            -> S7
  COMMA           -> S8
  SEMICOLON       -> S0
  EOF             -> ACCEPT
  otherwise       -> SX

S7 - After Dash
  DIGITS          -> S6
  otherwise       -> SX

S8 - After Comma
  DIGITS          -> S6
  otherwise       -> SX

SX - Error
  any             -> SX
```

### 3. Structural Acceptance Conditions

A reference is structurally valid if:
- Final state is S4 or S6
- AND next token is EOF or SEMICOLON

All other terminal states are structural errors.

### 4. Post-DSP Processing (Outside DSP Scope)

After structural acceptance:

**stage 3:**
Resolve BookRef -> BookID

**stage 4:**
- Assign Chapter and Verse exactly once
- Apply single-chapter inference if applicable

**stage 5:**
Enforce canonical validation matrix:
- AliasFound = True
- Chapter >= 1
- Chapter <= MaxChapter(BookID)
- Verse bounds

**stage 6:**
Format canonical SBL reference

**stage 7:**
Emit immutable ScriptureRef result

### 5. Canonical Output Form (After Stage 6)

```
<CanonicalBookName> <Chapter>
<CanonicalBookName> <Chapter>:<Verse>
```

Lists and ranges are preserved structurally but validated semantically in Stage 5.

---

## NOTES

What makes the compressed verse-map approach strong in your architecture is:

### 1. Data Is Data

The validator performs only two checks:
- Is chapter within the book's chapter count?
- Is verse within the chapter's verse count?

It does not contain any hard-coded knowledge of Bible structure.

**Architectural Principle: metadata Authority**

Canonical Bible structure (chapter counts and verse counts) is stored entirely in metadata tables. Parser logic never encodes Bible structure.

**Benefits:**
- deterministic validation
- zero structural branching
- translation-specific metadata substitution
- Proper separation of concerns

### 2. Deterministic O(1) Lookup

With fixed-width packed strings:
```vb
maxV = CLng(Mid$(map, (Chapter - 1) * 3 + 1, 3))
```
The lookup is constant time:
- No loops
- No Select Case
- No Split
- No dynamic lookup structures
- Direct index addressing only

### 3. It Scales Without Growing Code

Adding all 66 books:
- Adds data only
- Adds zero logic
- Validator does not encode Bible structure. It only verifies numeric bounds using metadata lookups

### 4. It Matches the Design Philosophy

Parser stages remain isolated:
- Tokenizer
- Alias Resolver
- Structural Interpreter
- Semantic Validator
- Canonical Formatter

### 5. The Key Benefit

Because structural rules are stored in metadata, different validation policies can be implemented without modifying parser logic.

**Examples:**
- Strict SBL validation
- Relaxed validation (bounds skipped)
- Alternate verse maps for LXX, Vulgate, or other traditions

The validator remains unchanged. Only the metadata source is replaced.

---

The design has moved from string parsing to formal citation semantics - a big architectural leap.

**Special case:**
Psalm 119 contains 176 verses, which exceeds two-digit storage limits. Therefore all verse counts are stored using fixed-width three-digit encoding:
```vb
Right$("000" & verseCount, 3)
```
This ensures constant-width indexing for direct addressing.
