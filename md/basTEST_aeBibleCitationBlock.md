# Plan: SBL Citation Block Verifier
**Date:** 2026-04-01
**Target module:** `src/basTEST_aeBibleCitationBlock.bas` (new standard module)

---

## Problem Statement

The existing parser pipeline (`ComposeList`, `ParseScripture`) requires a book alias on **every** semicolon-separated segment. Study Bible citation blocks use **book-context propagation** — after `Ps`, references like `23:1; 28:7; 68:5` carry no book name and inherit `Ps` from the preceding segment. This is standard SBL citation format and is the central gap.

### Sample citation block to verify

```
Gen 1:27; Num 14:18; Deut 32:6; Josh 1:9; 1 Sam 2:2; 1 Chr 29:10–13;
Ps 19:1–2; 23:1; 28:7; 68:5; 103:8–11; 111:3–5; 145:8–9,17; Isa 40:28; 63:16; 64:8;
Jer 33:11; Nah 1:3; Mal 2:10–15; Matt 6:9; 7:11; 23:9; John 3:16; 4:24;
Rom 1:20; 8:15; 1 Cor 8:6; 14:33; Gal 3:20; Eph 4:6; Heb 13:6; 1 Pet 1:17;
2 Pet 3:9; 1 John 4:16
```

### Three additional issues the pipeline cannot currently handle

| Issue | Example | Impact |
|---|---|---|
| Multi-line input | Line breaks between `1 Chr 29:10–13;` and `Ps 19:1–2` | Must be stripped before any parsing |
| Range verse specs | `103:8–11`, `29:10–13` | `ValidateSBLReference` rejects non-numeric `VerseSpec`; ranges must be decomposed to atomic start/end verses before validation |
| Comma-within-range | `145:8–9,17` | Two tokens: range `8–9` and bare verse `17`, both inheriting `Ps 145` context |

---

## Architecture

All new code goes in a new standard module `basTEST_aeBibleCitationBlock.bas`. The class `aeBibleCitationClass.cls` is **not modified**. Existing class methods are called via the predeclared `aeBibleCitationClass` singleton instance.

---

## Data Structure: `BlockToken`

```vb
Private Type BlockToken
    InputAlias  As String   ' e.g. "Ps", "1 Cor" — empty string if inherited from context
    BookID      As Long     ' 0 if unresolved
    CanonName   As String   ' canonical name from ResolveAlias
    Chapter     As Long
    StartVerse  As Long     ' after DecomposeVerseSpec
    EndVerse    As Long     ' = StartVerse if not a range
    IsRange     As Boolean
    SegText     As String   ' original segment text for error messages
    ErrorCode   As Long     ' 0 = ok; see constants
    ErrorText   As String
End Type
```

### Error constants

```vb
Public Const E_ALIAS_UNRESOLVED As Long = 1001  ' ResolveAlias raised error
Public Const E_CHAPTER_MISSING  As Long = 1002  ' No chapter could be inferred
Public Const E_VERSE_MALFORMED  As Long = 1003  ' VerseSpec not numeric and not range
Public Const E_SBL_FAIL         As Long = 1006  ' ValidateSBLReference returned False
```

---

## Functions to Build

All functions are `Private` unless noted. Implementation order matches dependency order.

| # | Modifier | Name | Signature | Purpose |
|---|---|---|---|---|
| 1 | Private | `NormalizeBlockInput` | `(raw As String) As String` | Replace `vbCr`/`vbLf`/`vbCrLf` with space; collapse multiple spaces |
| 2 | Private | `TryResolveAlias` | `(alias As String, ByRef BookID As Long, ByRef canonName As String) As Boolean` | Error-safe wrapper around `aeBibleCitationClass.ResolveAlias`; returns False on any error |
| 3 | Private | `DetectBookAliasInSegment` | `(seg As String, contextBookID As Long, ByRef alias As String, ByRef refPart As String) As Boolean` | Determines whether a segment begins with a new book alias; returns True if new book found |
| 4 | Private | `DecomposeVerseSpec` | `(spec As String, ByRef startV As Long, ByRef endV As Long) As Boolean` | Splits `"8–9"` or `"8-9"` → `startV=8, endV=9, True`; returns False for single verse |
| 5 | Private | `TokenizeCitationBlock` | `(raw As String) As BlockToken()` | Stage 0 — produces flat array of `BlockToken`; propagates book and chapter context |
| 6 | Private | `FormatTokenRef` | `(t As BlockToken) As String` | Formats a `BlockToken` as a canonical reference string for display |
| 7 | Public  | `VerifyCitationBlock` | `(rawBlock As String)` | Iterates tokens; validates each atomic endpoint via `ValidateSBLReference(ModeSBL)`; prints PASS/FAIL |
| 8 | Public  | `Test_VerifyCitationBlock` | `()` | Positive test — full citation block; expected: all 35 tokens pass |
| 9 | Public  | `Test_VerifyCitationBlock_Negative` | `()` | 3-case negative test: bad alias, verse out of range, chapter out of range |

---

## Function Details

### `DetectBookAliasInSegment` — alias detection rule

Split `seg` on spaces. Classify `parts(0)`:

1. **Single digit** (`"1"`, `"2"`, `"3"`) → try `parts(0) & " " & parts(1)` as a two-token alias via `TryResolveAlias`. If it resolves: new book found, `refPart = parts(2..n)`. Handles `1 Sam`, `1 Chr`, `1 Cor`, `1 Pet`, `2 Pet`, `1 John`.
2. **Starts with a letter** → try `parts(0)` as a one-token alias. If it resolves: new book found, `refPart = parts(1..n)`. Handles `Gen`, `Ps`, `Isa`, `Matt`, `John`, etc.
3. **Multi-digit number** (e.g. `"23"`) → bare `chapter:verse` continuation; inherit context. Return `False`.

If alias detection fails for case 1 or 2, fall back to treating the full segment as a bare reference inheriting the current context.

### `DecomposeVerseSpec` — dash handling

Scans `spec` character by character. Detects either ASCII hyphen (`-`, Chr(45)) or en dash (`–`, ChrW(8211)). If found, splits at that position into `startV` and `endV`. Returns `True` (is range). If no dash found, sets `startV = endV = CLng(spec)`, returns `False`.

### `TokenizeCitationBlock` — Stage 0 algorithm

```
1. NormalizeBlockInput(raw)
2. Split on ";" → major segments
3. contextBookID = 0 : contextCanon = "" : contextChapter = 0
4. For each segment (trimmed):
   a. DetectBookAliasInSegment → updates contextBookID/contextCanon if new book
   b. Parse refPart for ":"
        If ":" found: left = chapter, right = verseSpec string
        If no ":":    treat as chapter-only (verseSpec = "")
        Update contextChapter
   c. Split verseSpec on "," → sub-segments (handles "8–9,17")
   d. For each sub-segment:
        DecomposeVerseSpec → startV, endV, isRange
        Populate BlockToken; append to result array
5. Return array
```

### `VerifyCitationBlock` — validation loop

```
1. Dim tokens() = TokenizeCitationBlock(rawBlock)
2. For each token:
   a. If token.ErrorCode <> 0: print FAIL + ErrorText; failCount++; continue
   b. Validate start verse:
        ok = aeBibleCitationClass.ValidateSBLReference(
                 token.BookID, token.CanonName,
                 token.Chapter, CStr(token.StartVerse), ModeSBL, True)
        If Not ok: print FAIL; failCount++
   c. If token.IsRange: validate EndVerse the same way
   d. If all ok: print "PASS: " & FormatTokenRef(token); passCount++
3. Print summary: passCount & " passed, " & failCount & " failed."
```

### `Test_VerifyCitationBlock` — positive test

Constructs the raw string using `ChrW(8211)` for en dashes (exercises the Unicode path in `DecomposeVerseSpec`):

```vb
rawBlock = "Gen 1:27; Num 14:18; Deut 32:6; Josh 1:9; 1 Sam 2:2; " & _
           "1 Chr 29:10" & ChrW(8211) & "13; " & _
           "Ps 19:1" & ChrW(8211) & "2; 23:1; 28:7; 68:5; " & _
           "103:8" & ChrW(8211) & "11; 111:3" & ChrW(8211) & "5; " & _
           "145:8" & ChrW(8211) & "9,17; Isa 40:28; 63:16; 64:8; " & _
           "Jer 33:11; Nah 1:3; Mal 2:10" & ChrW(8211) & "15; " & _
           "Matt 6:9; 7:11; 23:9; John 3:16; 4:24; " & _
           "Rom 1:20; 8:15; 1 Cor 8:6; 14:33; Gal 3:20; Eph 4:6; " & _
           "Heb 13:6; 1 Pet 1:17; 2 Pet 3:9; 1 John 4:16"
VerifyCitationBlock rawBlock
```

### `Test_VerifyCitationBlock_Negative` — 3 deliberate failures

**Note on alias test design:** The alias used must be provably absent from the alias map — not just a non-SBL form of a real book. `Jeremiah` is the canonical name and IS in the alias map, so `TryResolveAlias("Jeremiah", ...)` succeeds and the test passes, defeating its purpose. The correct approach is a string that cannot resolve: a misspelling, a punctuated form, or a fabricated abbreviation. `Jerimiah` (common misspelling of Jeremiah) is confirmed absent from the alias map and produces a clean, readable failure.

| Case | Change | Expected error |
|---|---|---|
| Bad alias | Replace `Jer 33:11` with `Jerimiah 33:11` | `E_ALIAS_UNRESOLVED` — misspelling absent from alias map |
| Verse out of range | Replace `103:8–11` with `103:8–200` | `E_SBL_FAIL` — verse 200 > max for Ps 103 |
| Chapter out of range | Replace `Jer 33:11` with `Jer 99:1` | `E_SBL_FAIL` — Jeremiah has 52 chapters |

---

## Token Trace — Full Citation Block (36 tokens)

| Segment (trimmed) | New book? | BookID | Ch | VerseSpec | Token count |
|---|---|---|---|---|---|
| `Gen 1:27` | Yes — Gen | 1 | 1 | 27 | 1 |
| `Num 14:18` | Yes — Num | 4 | 14 | 18 | 1 |
| `Deut 32:6` | Yes — Deut | 5 | 32 | 6 | 1 |
| `Josh 1:9` | Yes — Josh | 6 | 1 | 9 | 1 |
| `1 Sam 2:2` | Yes — 1 Sam | 9 | 2 | 2 | 1 |
| `1 Chr 29:10–13` | Yes — 1 Chr | 13 | 29 | 10–13 | 1 (range) |
| `Ps 19:1–2` | Yes — Ps | 19 | 19 | 1–2 | 1 (range) |
| `23:1` | inherit Ps | 19 | 23 | 1 | 1 |
| `28:7` | inherit Ps | 19 | 28 | 7 | 1 |
| `68:5` | inherit Ps | 19 | 68 | 5 | 1 |
| `103:8–11` | inherit Ps | 19 | 103 | 8–11 | 1 (range) |
| `111:3–5` | inherit Ps | 19 | 111 | 3–5 | 1 (range) |
| `145:8–9,17` | inherit Ps | 19 | 145 | 8–9 + 17 | **2 tokens** |
| `Isa 40:28` | Yes — Isa | 23 | 40 | 28 | 1 |
| `63:16` | inherit Isa | 23 | 63 | 16 | 1 |
| `64:8` | inherit Isa | 23 | 64 | 8 | 1 |
| `Jer 33:11` | Yes — Jer | 24 | 33 | 11 | 1 |
| `Nah 1:3` | Yes — Nah | 34 | 1 | 3 | 1 |
| `Mal 2:10–15` | Yes — Mal | 39 | 2 | 10–15 | 1 (range) |
| `Matt 6:9` | Yes — Matt | 40 | 6 | 9 | 1 |
| `7:11` | inherit Matt | 40 | 7 | 11 | 1 |
| `23:9` | inherit Matt | 40 | 23 | 9 | 1 |
| `John 3:16` | Yes — John | 43 | 3 | 16 | 1 |
| `4:24` | inherit John | 43 | 4 | 24 | 1 |
| `Rom 1:20` | Yes — Rom | 45 | 1 | 20 | 1 |
| `8:15` | inherit Rom | 45 | 8 | 15 | 1 |
| `1 Cor 8:6` | Yes — 1 Cor | 46 | 8 | 6 | 1 |
| `14:33` | inherit 1 Cor | 46 | 14 | 33 | 1 |
| `Gal 3:20` | Yes — Gal | 48 | 3 | 20 | 1 |
| `Eph 4:6` | Yes — Eph | 49 | 4 | 6 | 1 |
| `Heb 13:6` | Yes — Heb | 58 | 13 | 6 | 1 |
| `1 Pet 1:17` | Yes — 1 Pet | 60 | 1 | 17 | 1 |
| `2 Pet 3:9` | Yes — 2 Pet | 61 | 3 | 9 | 1 |
| `1 John 4:16` | Yes — 1 John | 62 | 4 | 16 | 1 |

**35 atomic tokens total. All expected to PASS.**

---

## Reuse vs. Build Summary

### Existing class functions — reused as-is

| Function | How used |
|---|---|
| `aeBibleCitationClass.ResolveAlias` | Book alias → BookID + canonical name; wrapped in `TryResolveAlias` |
| `aeBibleCitationClass.ValidateSBLReference(ModeSBL)` | Validates each atomic verse endpoint |
| `aeBibleCitationClass.GetMaxChapter` | Chapter bounds (used internally by `ValidateSBLReference`) |
| `aeBibleCitationClass.GetMaxVerse` | Verse bounds (used internally by `ValidateSBLReference`) |

### Existing class functions — NOT used

| Function | Reason |
|---|---|
| `ComposeList` | Requires book alias on every segment |
| `ParseScripture` | Same limitation |
| `ListDetection` | Replaced by the new tokenizer which handles semicolons and commas together |
| `RangeDetection` | Operates on full reference strings; `DecomposeVerseSpec` handles bare verse ranges |
| `ParseReference` | Requires book alias; crashes on context-only segments |

---

## Key Risk: `ValidateSBLReference` and non-numeric `VerseSpec`

`ValidateSBLReference` checks `IsNumeric(VerseSpec)` and returns False for range strings. Every range token (`1 Chr 29:10–13`, `Ps 19:1–2`, etc.) **must** be decomposed by `DecomposeVerseSpec` before `ValidateSBLReference` is called. `StartVerse` and `EndVerse` are validated as separate `CStr(Long)` calls. Without this, every range token produces a false failure.

---

## Implementation Sequence

1. Create `src/basTEST_aeBibleCitationBlock.bas` — `BlockToken` type, error constants, module header ✓
2. Implement `NormalizeBlockInput` and `TryResolveAlias` ✓
3. Implement `DetectBookAliasInSegment` and `DecomposeVerseSpec` ✓
4. Implement `TokenizeCitationBlock` ✓
5. Implement `FormatTokenRef` and `VerifyCitationBlock` ✓
6. Implement `Test_VerifyCitationBlock` (positive — 36 tokens, all pass) ✓
7. Implement `Test_VerifyCitationBlock_Negative` (3 deliberate failures) ✓

---

## Implementation Log

### 2026-04-01 — Initial implementation

`src/basTEST_aeBibleCitationBlock.bas` created with all 9 functions from the plan. CRLF normalized via WSL Python.

**Helper added beyond plan:** `SliceArray` — a private helper required by `DetectBookAliasInSegment` to slice a `String()` array from a given index to `UBound`. VBA has no built-in array slice; `Join(Array(...))` cannot take a sub-range directly. One private sub added: `AppendToken` — grows the `BlockToken()` array by one and stores the new token. Kept private; called only from `TokenizeCitationBlock`.

**Bug fix — `TryResolveAlias` signature mismatch:**

The plan described `ResolveAlias` as taking three parameters `(alias, ByRef BookID, ByRef canonName)`. The actual class signature is:

```vb
Public Function ResolveAlias(abbr As String, Optional BookID As Long) As String
```

The canonical name is the **return value**, not a third `ByRef` parameter. Fix applied:

```vb
' Before (wrong — wrong number of arguments):
aeBibleCitationClass.ResolveAlias alias, BookID, canonName

' After (correct):
canonName = aeBibleCitationClass.ResolveAlias(alias, BookID)
If canonName = "" Then GoTo RESOLVE_FAIL
```

The empty-string guard handles the case where `ResolveAlias` returns `""` for an unknown alias without raising an error, ensuring `TryResolveAlias` still returns `False`.

### 2026-04-01 — Fix `SliceArray` syntax error

**Error:** `ReDim empty(0 To -1)` — VBA does not permit a negative upper bound; this is a syntax error at compile time.

**Root cause:** The plan assumed an empty dynamic array could be returned by declaring a zero-length array with a negative upper bound. VBA forbids this.

**Fix:** Replace the invalid `ReDim` with `Split(vbNullString)`, which returns `Array("")` — a one-element array containing an empty string. `Join` on this result yields `""`, which is the correct value for `refPart` when no tokens follow the book alias.

```vb
' Before (syntax error):
Dim empty() As String
ReDim empty(0 To -1)
SliceArray = empty

' After (correct):
SliceArray = Split(vbNullString)  ' returns Array(""); Join gives ""
```

### 2026-04-01 — Fix `Chr(8211)` → `ChrW(8211)` in test data

**Error:** Runtime Error 5 "Invalid procedure call or argument" on the first line of `Test_VerifyCitationBlock` that calls `Chr(8211)`.

**Root cause:** VBA's `Chr()` function only accepts values 0–255. En dash (U+2013 = 8211) is outside that range. Any call to `Chr(8211)` raises Runtime Error 5 at runtime.

**Fix:** Replace every `Chr(8211)` with `ChrW(8211)`. `ChrW` accepts the full Unicode range (0–65535).

```vb
' Before (runtime error):
"1 Chr 29:10" & Chr(8211) & "13"

' After (correct):
"1 Chr 29:10" & ChrW(8211) & "13"
```

Same replacement applied to all six en-dash occurrences in `Test_VerifyCitationBlock` and the one occurrence in `Test_VerifyCitationBlock_Negative`. The `DecomposeVerseSpec` detection (`AscW(ch) = 8211`) was already correct — `AscW` handles Unicode input without issue.
