# Code Review: Stage 13a — Book Context Propagation

**Date:** 2026-04-03
**Parser doc updated:** `md/aeBibleCitationClass.md`

---

## Problem

`ComposeList` / `ParseScripture` require a book alias on every segment. Study Bible
citation blocks use book-context propagation — after `Ps`, segments like `23:1; 28:7`
carry no book name and must inherit `Ps` from the preceding segment.

---

## Where the Fix Goes

**Stage 13a — not between Stage 7 and Stage 8.**

`ListDetection` (Stage 8) already splits on `;`, producing clean segments. The failure
happens in `ComposeList_Internal` when it calls `ParseReferenceRef("23:1")`:
`LexicalScan` takes `"23:1"` as a whole-string alias → `ResolveBookStrict` fails.

Stage 13 already has two inline shorthand cases in `ComposeList_Internal`:
- bare verse number (`"18"` after `"John 3:16"`)
- numeric range (`"16-18"` after `"John 3:16"`)

Stage 13a adds a third: **if a segment has a colon and the left side is numeric, it is a
`chapter:verse` with no book alias — inherit `BookID` from the previous token** and bypass
`ParseReferenceRef` entirely.

**Failure trace for `"Ps 19:1; 23:1"`:**

| Stage | Function | What happens with `"23:1"` |
|---|---|---|
| 8 | `ListDetection` | Splits on `;` → produces segment `"23:1"` |
| 11 | `ComposeList_Internal` | Calls `ParseReferenceRef("23:1")` |
| 2 | `LexicalScan("23:1")` | `Split("23:1"," ")` → one element; whole string taken as alias |
| 3 | `ResolveBookStrict("23:1",...)` | No alias matches → **raises error** |

**Qualifying condition (all three must hold):**
1. `havePrev = True`
2. Segment contains `:`
3. Left of `:` is numeric

For range segments (`"103:8-11"` after `"Ps"`), the same rule applies: if the segment
matches `\d+:\d+-\d+` it is a chapter:verse-range with inherited book.

---

## Semicolon Handling

Semicolons are already consumed by `ListDetection` before Stage 13a sees any segment.
No change to semicolon handling is needed for the common case.

**However — `ListDetection` is an either/or splitter:**

```vb
If InStr(rawInput, ",") > 0 Then
    Split(rawInput, ",")   ' exits; semicolons ignored
    Exit Function
End If
If InStr(rawInput, ";") > 0 Then
    Split(rawInput, ";")
End If
```

For inputs with **both** delimiters (e.g. `"Ps 145:8-9,17; Isa 40:28"`), it splits on
comma first — semicolons survive inside segments and `ParseReferenceRef` then fails on
`"17; Isa 40:28"`.

**Conclusion:** Stage 13a in `ComposeList_Internal` is correct and sufficient for
**pure-semicolon** inputs. Mixed `;` and `,` inputs (the full citation block) require a
new entry point.

---

## New Entry Point: `ParseCitationBlock`

Two-level split is not possible in `ListDetection` without breaking existing callers.
A new public method on the class handles the citation block structure:

```
1. NormalizeRawInput  — strip line breaks; replace en-dash with ASCII hyphen
2. Split on ";"       — major segments (book/chapter context propagates across these)
3. For each segment:
   a. Detect book alias (Stage 13a qualifying condition)
   b. Split on ","    — verse sub-items within the segment
   c. DecomposeVerseSpec -> StartVerse, EndVerse
   d. ValidateSBLReference(ModeSBL) per atomic endpoint
4. Return Collection of canonical reference strings
```

This is the class-level replacement for `TokenizeCitationBlock` in
`basTEST_aeBibleCitationBlock.bas`.

---

## En-Dash Normalization

`IsRangeSegment` and `RangeDetection` only recognize ASCII hyphen. Citation blocks use
en-dash (`–`, ChrW(8211)). `NormalizeRawInput` must replace en-dash with ASCII hyphen
before any parsing. Called at the top of `ComposeList` and `ParseCitationBlock`.

---

## Procedures Moving from `basTEST_aeBibleCitationBlock.bas`

| Procedure | Disposition |
|---|---|
| `NormalizeBlockInput` | → class as `Public Function NormalizeRawInput`; add en-dash normalization |
| `DecomposeVerseSpec` | → absorbed into extended `IsRangeSegment` / `RangeDetection` |
| `DetectBookAliasInSegment` | → absorbed into Stage 13a inline case in `ComposeList_Internal` |
| `TokenizeCitationBlock` | → replaced by `ParseCitationBlock` in the class |
| `TryResolveAlias` | → removed; `ResolveAlias` with error handling exists in the class |
| `AppendToken`, `SliceArray`, `FormatTokenRef`, `BlockToken` | → removed |
| `VerifyCitationBlock` | **stays** — simplified to call `ParseCitationBlock` |
| `Test_VerifyCitationBlock` | **stays** — positive full-block integration test |
| `Test_VerifyCitationBlock_Negative` | → removed; all cases covered by `Test_Stage13a_BookContextPropagation` |

---

## Updated `Run_All_SBL_Tests` Sequence

```vb
Test_Stage13_ContextShorthand           ' existing
Test_Stage13a_BookContextPropagation    ' NEW — insert here
Test_Stage14_CanonicalCompression
```

---

## New Test: `Test_Stage13a_BookContextPropagation`

Added to `basTEST_aeBibleCitationClass.bas`. Called from `Run_All_SBL_Tests` immediately
after `Test_Stage13_ContextShorthand`.

**Positive cases:**

```vb
' Single-book propagation — Psalms context
Set c = aeBibleCitationClass.ComposeList("Ps 19:1; 23:1; 28:7")
aeAssert.AssertEqual 3, c.Count, "Stage13a: 3 Psalm refs"
aeAssert.AssertEqual "Psalms 19:1", c(1), "Stage13a: Ps 19:1"
aeAssert.AssertEqual "Psalms 23:1", c(2), "Stage13a: Ps 23:1 inherited"
aeAssert.AssertEqual "Psalms 28:7", c(3), "Stage13a: Ps 28:7 inherited"

' Cross-book transition — Psalms then Isaiah
Set c = aeBibleCitationClass.ComposeList("Ps 103:8; Isa 40:28; 63:16")
aeAssert.AssertEqual 3, c.Count, "Stage13a: cross-book count"
aeAssert.AssertEqual "Psalms 103:8", c(1), "Stage13a: Ps 103:8"
aeAssert.AssertEqual "Isaiah 40:28", c(2), "Stage13a: Isa 40:28"
aeAssert.AssertEqual "Isaiah 63:16", c(3), "Stage13a: Isa 63:16 inherited"

' Range with inherited book
Set c = aeBibleCitationClass.ComposeList("Ps 19:1-2; 103:8-11")
aeAssert.AssertEqual 2, c.Count, "Stage13a: Psalm range count"
aeAssert.AssertEqual "Psalms 19:1-2",   c(1), "Stage13a: Ps 19:1-2"
aeAssert.AssertEqual "Psalms 103:8-11", c(2), "Stage13a: Ps 103:8-11 inherited"
```

**Negative cases (correctly rejected = test passes):**

```vb
' Bad alias — "Jerimiah" is a misspelling; absent from alias map
Dim ok As Boolean
On Error Resume Next
ok = False
aeBibleCitationClass.ComposeList "Gen 1:1; Jerimiah 33:11; Mal 1:1"
ok = (Err.Number <> 0)
On Error GoTo 0
aeAssert.AssertTrue ok, "Stage13a neg: bad alias (Jerimiah) rejected"

' Verse out of range — Ps 103 has 22 verses; verse 200 invalid
Dim valid As Boolean
valid = aeBibleCitationClass.ValidateSBLReference(19, "Psalms", 103, "200", ModeSBL, True)
aeAssert.AssertTrue Not valid, "Stage13a neg: Ps 103:200 rejected"

' Chapter out of range — Jeremiah has 52 chapters; chapter 99 invalid
valid = aeBibleCitationClass.ValidateSBLReference(24, "Jeremiah", 99, "1", ModeSBL, True)
aeAssert.AssertTrue Not valid, "Stage13a neg: Jer 99:1 rejected"

' Jude 99 — single-chapter book; implicit Chapter 1; verse 99 > max (25)
valid = aeBibleCitationClass.ValidateSBLReference(65, "Jude", 0, "99", ModeSBL, True)
aeAssert.AssertTrue Not valid, "Stage13a neg: Jude 99 rejected (max verse 25)"
```

**Note on `Jude 99`:** `ValidateSBLReference` normalizes `Chapter=0, VerseSpec="99"` to
`Chapter=1, VerseSpec="99"` for single-chapter books. `GetMaxVerse(65,1)=25`, so
`99 > 25` → False. Tests both the implicit-chapter rule and verse bounds in one case.

---

## Goal State After Implementation

### `src/aeBibleCitationClass.cls` additions

| New item | Purpose |
|---|---|
| `Public Function NormalizeRawInput` | Strip CR/LF, collapse spaces, replace en-dash with hyphen |
| Extended `IsRangeSegment` / `RangeDetection` | Accept en-dash as range separator |
| Stage 13a inline case in `ComposeList_Internal` | Handle `chapter:verse` segments with inherited book |
| `Public Function ParseCitationBlock` | Two-level split for mixed `;` and `,` citation blocks |

### `src/basTEST_aeBibleCitationBlock.bas` after cleanup

| Remains | Purpose |
|---|---|
| `VerifyCitationBlock` (simplified) | Calls `ParseCitationBlock` + `ValidateSBLReference` per item; no `BlockToken` type |
| `Test_VerifyCitationBlock` | Full 35-token positive integration test |

Everything else removed. The module deals only with citation blocks. All atomic and
list-level logic lives in the class.

---

## Implementation Order

1. Add `NormalizeRawInput` to class; call from `ComposeList` and `ParseScripture`
2. Extend `IsRangeSegment` / `RangeDetection` to accept en-dash
3. Add Stage 13a inline case to `ComposeList_Internal`
4. Add `ParseCitationBlock` to class (two-level split)
5. Add `Test_Stage13a_BookContextPropagation`; wire into `Run_All_SBL_Tests`
6. Simplify `VerifyCitationBlock` to call `ParseCitationBlock`
7. Remove `BlockToken`, `TokenizeCitationBlock`, `DetectBookAliasInSegment`,
   `DecomposeVerseSpec`, `TryResolveAlias`, `AppendToken`, `SliceArray`,
   `FormatTokenRef`, `Test_VerifyCitationBlock_Negative` from
   `basTEST_aeBibleCitationBlock.bas`
8. Run `Run_All_SBL_Tests` and `Test_VerifyCitationBlock`; verify all pass

---

## EBNF Updates — `md/aeBibleCitationClass.md`

Four changes made to the documentation on 2026-04-04.

**1. Section heading**
Updated from "Aligned to 7-Stage Deterministic Parser" to "Aligned to DSP Pipeline
(Stages 1–13a)" to reflect that the grammar now covers the full pipeline including the
context resolution layer.

**2. `Reference` rule — third production added**

```ebnf
Reference
   ::= BookRef
    |  BookRef WS ChapterSpec
    |  ChapterSpec            (* Stage 13a: book inherited from context *)
```

Added semantic constraint note: a bare `ChapterSpec` reference is valid only when a
preceding `BookRef` exists in the same `Citation`. This constraint is context-sensitive
and cannot be expressed in context-free EBNF; it is enforced by `ComposeList_Internal`
via the `havePrev` guard. Added parse examples for single-book propagation, range
propagation, cross-book transition, and the ill-formed case (bare `ChapterSpec` with no
preceding `BookRef`).

**3. `VerseRange` rule — en-dash normalization note**

Added note below the `VerseRange ::= Verse "-" Verse` production: study Bible citation
blocks use en-dash (`–`, U+2013) as the range separator. `NormalizeRawInput` replaces
en-dash with ASCII hyphen before any parsing stage is reached. The grammar uses only `-`;
the en-dash form never reaches the parser.

**4. Canonical EBNF — Stage 13a resolution note**

No structural change to the canonical grammar — every canonical output item is already
`CanonicalBookRef ::= BookName WS CanonicalChapterSpec`, which requires a book name.
Added note confirming that Stage 13a resolves all inherited-book inputs to fully-qualified
`CanonicalBookRef` form before output. The bare `ChapterSpec` form is input-only and
never appears in canonical output.

---

## Implementation Complete — 2026-04-04

### Files changed

**`src/aeBibleCitationClass.cls`**
- `NormalizeRawInput` (new Public) — strips CR/LF, collapses spaces, replaces en-dash with ASCII hyphen
- `IsBooklessChapRef` (new Private) — returns True when left of `:` is numeric; used by Stage 13a
- `ComposeList_Internal` — two new `ElseIf` branches for Stage 13a: `chapter:verse` and `chapter:verse-range` with inherited book; four new variables declared (`cp13`, `ch13`, `vr13`, `dp13`)
- `ComposeList` — added `raw = NormalizeRawInput(raw)` before `ComposeList_Internal`
- `ParseScripture` — replaced `raw = Trim$(raw)` with `raw = NormalizeRawInput(raw)`
- `ParseCitationBlock` (new Public) — two-level split (`";"` then `","`) with book-context propagation; replaces `TokenizeCitationBlock` from the block module

**`src/basTEST_aeBibleCitationClass.bas`**
- `Test_Stage13a_BookContextPropagation` (new) — 3 positive assertions + 4 negative assertions
- `Run_All_SBL_Tests` — added `Test_Stage13a_BookContextPropagation` after `Test_Stage13_ContextShorthand`

**`src/basTEST_aeBibleCitationBlock.bas`**
- Removed: `BlockToken` type, error constants, `NormalizeBlockInput`, `TryResolveAlias`, `DetectBookAliasInSegment`, `SliceArray`, `DecomposeVerseSpec`, `TokenizeCitationBlock`, `AppendToken`, `FormatTokenRef`, `Test_VerifyCitationBlock_Negative`
- `VerifyCitationBlock` rewritten — calls `aeBibleCitationClass.ParseCitationBlock`; parses each returned canonical string; validates via `ValidateSBLReference(ModeSBL)`; returns failCount
- `Test_VerifyCitationBlock` retained unchanged (positive 35-token integration test)

### Pre-existing — no change required
- `IsRangeSegment` — already handled en-dash (ChrW(8211)) ✓
- `RangeDetection` — already handled en-dash ✓

---

## Post-Implementation Fixes — 2026-04-04

### Standard Error Handlers Added

Added `On Error GoTo PROC_ERR` / `PROC_EXIT:` / `PROC_ERR:` / `MsgBox` / `Resume PROC_EXIT`
to all Stage 13a procedures that were missing them.

**`src/aeBibleCitationClass.cls`**
- `NormalizeRawInput` — added standard handler
- `ParseCitationBlock` — added standard handler; internal `On Error Resume Next` blocks
  for alias resolution now restore to `GoTo PROC_ERR` instead of `GoTo 0`

**`src/basTEST_aeBibleCitationClass.bas`**
- `Test_Stage13a_BookContextPropagation` — added standard handler; `On Error Resume Next`
  block for bad-alias negative test restores to `GoTo PROC_ERR`

**`src/basTEST_aeBibleCitationBlock.bas`**
- `VerifyCitationBlock` — replaced non-standard `On Error GoTo PARSE_ERR` / `PARSE_ERR:`
  handler with standard `PROC_ERR` / `PROC_EXIT` pattern

`IsBooklessChapRef` (Private, trivial) — no handler, consistent with other private helpers
such as `IsNumericRange`.

### Bug Fix — `ParseCitationBlock` Error 5

**Symptom:** `Test_VerifyCitationBlock` raised Error 5 (Invalid procedure call or argument)
at the verse-range detection line in `ParseCitationBlock`.

**Cause:** VBA does not short-circuit `And`. The compound condition
`If dp > 0 And IsNumeric(Left$(vsRaw, dp - 1)) ...` evaluated `Left$(vsRaw, -1)` when
`dp = 0`, producing Error 5.

**Fix:** Split the condition — `Left$` and `Mid$` calls are now nested inside
`If dp > 0 Then`, so they execute only when `dp` is valid.

### Documentation — En Dash in Verse Ranges

Updated `md/aeBibleCitationClass.md` to use en dash (`–`, U+2013) for all verse range
examples throughout the document (e.g. `John 3:16–18`, `Ps 103:8–11`). Changes apply to
prose, inline code, and `text` blocks in Stages 8–17 and the EBNF parse examples.

Not changed: EBNF grammar terminal strings (`"-"` in `VerseRange ::= Verse "-" Verse`),
VBA code blocks, and the en-dash normalization NOTE which explicitly contrasts the two
separator characters.
