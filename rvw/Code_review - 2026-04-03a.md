# Code Review: Book-Context Propagation — Architecture Plan

**Date:** 2026-04-03
**Target module (changes):** `src/aeBibleCitationClass.cls`
**Target module (tests):** `src/basTEST_aeBibleCitationClass.bas`
**Target module (simplified):** `src/basTEST_aeBibleCitationBlock.bas`

---

## Problem Statement

The existing parser pipeline (`ComposeList` / `ParseScripture`) requires a book alias on
**every** semicolon-separated segment. Study Bible citation blocks use **book-context
propagation** — after `Ps`, references like `23:1; 28:7; 68:5` carry no book name and
inherit `Ps` from the preceding segment. This is standard SBL citation format and is the
central gap.

---

## Where Does the Fix Go?

**Not between Stage 7 and Stage 8. The fix belongs in Stage 13.**

Here is why, tracing the failure path for `"Ps 19:1; 23:1"`:

| Stage | Function | What happens with `"23:1"` |
|---|---|---|
| 8 | `ListDetection` | Splits on `;` → produces segment `"23:1"` |
| 11 | `ComposeList_Internal` | Iterates segments; calls `ParseReferenceRef("23:1")` |
| 2 | `LexicalScan("23:1")` | `Split("23:1", " ")` → `parts(0)="23:1"`, UBound=0, Case 0 → `RawAlias = "23:1"` |
| 3 | `ResolveBookStrict("23:1", ...)` | No alias matches → **raises error** |

Stage 7 (`ParseReference`) is the **atomic** single-reference parser — it operates on one
complete reference and has no memory of prior references. Stage 8 (`ListDetection`) is a
dumb splitter by design. The semantic context already lives in Stage 13.

Stage 13 already handles **two** shorthand cases inline in `ComposeList_Internal`:

```vb
' Bare verse number: "18" after "John 3:16" → John 3:18
If havePrev And IsNumeric(seg) And InStr(seg, ":") = 0 Then ...

' Numeric range: "16-18" after "John 3:16" → John 3:16-18
If havePrev And IsNumericRange(seg) Then ...
```

The missing case is the **book-context** shorthand: `"23:1"` after `"Ps 19:1"` — a
`chapter:verse` segment where the left side of the colon is numeric and there is no book
alias. This is Stage 13 extended, labelled **Stage 13a**.

---

## Stage 13a Contract

A segment qualifies for book-context propagation when ALL of:

1. `havePrev` is True (a prior reference exists to inherit from)
2. The segment contains `:`
3. Left of `:` is numeric (i.e., it is a chapter number, not a book alias)

When these conditions hold, the segment is parsed as `chapter:verse` inheriting
`BookID` from the previous token — bypassing `ParseReferenceRef` entirely.

For range segments (`"103:8-11"` after `"Ps"`), the same rule applies: if the segment
matches `\d+:\d+-\d+` it is a chapter:verse-range with inherited book.

---

## Additional Prerequisite: En-Dash Normalization

`IsRangeSegment` and `RangeDetection` in the class currently detect only ASCII hyphen
(`-`, Chr(45)). Study Bible citation blocks use en-dash (`–`, ChrW(8211)):

```
103:8–11   Ps 19:1–2   1 Chr 29:10–13
```

`NormalizeBlockInput` in `basTEST_aeBibleCitationBlock.bas` strips line breaks but does
**not** normalize en-dashes. A `NormalizeRawInput` method must be added to the class that
replaces en-dashes with ASCII hyphens, and called at the top of `ComposeList` and
`ParseScripture` before any parsing begins.

---

## Procedures to Move from `basTEST_aeBibleCitationBlock.bas`

| Procedure | Disposition |
|---|---|
| `NormalizeBlockInput` | Move to `aeBibleCitationClass.cls` as `Public Function NormalizeRawInput`; add en-dash → hyphen normalization; call from `ComposeList` and `ParseScripture` |
| `DecomposeVerseSpec` | Logic absorbed into extended `IsRangeSegment` / `RangeDetection` (en-dash detection already needed there) |
| `DetectBookAliasInSegment` | Logic absorbed into Stage 13a inline check in `ComposeList_Internal` |
| `TokenizeCitationBlock` | Replaced by enhanced `ComposeList` + `ValidateSBLReference` loop |
| `TryResolveAlias` | Removed — `ResolveAlias` with error handling exists in the class |
| `AppendToken`, `SliceArray` | Removed — no longer needed |
| `BlockToken` type | Removed — no longer needed |
| `FormatTokenRef` | Removed — `CanonicalFromRef` in the class covers this |
| `VerifyCitationBlock` | **Remains** but simplified: calls `ComposeList`, then `ValidateSBLReference` per item |
| `Test_VerifyCitationBlock` | **Remains** — positive full-block integration test |
| `Test_VerifyCitationBlock_Negative` | **Removed** — all three cases covered by `Test_Stage13a_BookContextPropagation` |

---

## New Test: `Test_Stage13a_BookContextPropagation`

Added to `basTEST_aeBibleCitationClass.bas`. Inserted into `Run_All_SBL_Tests`
immediately after `Test_Stage13_ContextShorthand`.

### Positive cases

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
aeAssert.AssertEqual "Psalms 103:8",  c(1), "Stage13a: Ps 103:8"
aeAssert.AssertEqual "Isaiah 40:28",  c(2), "Stage13a: Isa 40:28"
aeAssert.AssertEqual "Isaiah 63:16",  c(3), "Stage13a: Isa 63:16 inherited"

' Range with inherited book — "103:8-11" after "Ps"
Set c = aeBibleCitationClass.ComposeList("Ps 19:1-2; 103:8-11")
aeAssert.AssertEqual 2, c.Count, "Stage13a: Psalm range count"
aeAssert.AssertEqual "Psalms 19:1-2",   c(1), "Stage13a: Ps 19:1-2"
aeAssert.AssertEqual "Psalms 103:8-11", c(2), "Stage13a: Ps 103:8-11 inherited"
```

### Negative cases (expected rejection → assertion passes when failCount >= 1)

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
' ParseReference("Jude 99") expands to Jude 1:99 via single-chapter rule,
' then ValidateSBLReference rejects verse 99.
valid = aeBibleCitationClass.ValidateSBLReference(65, "Jude", 0, "99", ModeSBL, True)
aeAssert.AssertTrue Not valid, "Stage13a neg: Jude 99 rejected (max verse 25)"
```

**Note on `Jude 99`:** Jude is a single-chapter book. `ValidateSBLReference` normalizes
`Chapter=0, VerseSpec="99"` to `Chapter=1, VerseSpec="99"` (lines 1116–1121 of the
class). `GetMaxVerse(65, 1) = 25`, so `99 > 25` → False. This tests both the
single-chapter implicit-chapter rule and the verse bounds check in one case.

---

## Updated `Run_All_SBL_Tests` Sequence

```vb
Test_Stage13_ContextShorthand       ' existing
Test_Stage13a_BookContextPropagation' NEW — book-context propagation + Jude 99
Test_Stage14_CanonicalCompression
...
```

---

## Goal State After Implementation

### `src/aeBibleCitationClass.cls` additions

| New item | Purpose |
|---|---|
| `Public Function NormalizeRawInput(raw As String) As String` | Strip CR/LF, collapse spaces, replace en-dash with hyphen |
| Extended `IsRangeSegment` / `RangeDetection` | Accept en-dash as range separator |
| Stage 13a inline case in `ComposeList_Internal` | Handle `chapter:verse` segments with inherited book |
| `Public Sub Test_Stage13a_BookContextPropagation` (in basTEST) | 4 positive + 4 negative assertions |

### `src/basTEST_aeBibleCitationBlock.bas` after cleanup

| Remains | Purpose |
|---|---|
| `VerifyCitationBlock` (simplified) | Calls `ComposeList` + `ValidateSBLReference` per item; no `BlockToken` type |
| `Test_VerifyCitationBlock` | Full 35-token positive integration test |

Everything else in the module is removed. The module deals only with citation blocks
(multi-book, multi-segment, full-paragraph input). All atomic and list-level logic lives
in the class.

---

## Implementation Order

1. Add `NormalizeRawInput` to class; call from `ComposeList` and `ParseScripture`; extend `IsRangeSegment` / `RangeDetection` for en-dash
2. Add Stage 13a inline case to `ComposeList_Internal` (book-context propagation)
3. Add `Test_Stage13a_BookContextPropagation` to `basTEST_aeBibleCitationClass.bas`; add call to `Run_All_SBL_Tests`
4. Simplify `VerifyCitationBlock` in `basTEST_aeBibleCitationBlock.bas` to use `ComposeList` directly
5. Remove `BlockToken`, `TokenizeCitationBlock`, `DetectBookAliasInSegment`, `DecomposeVerseSpec`, `TryResolveAlias`, `AppendToken`, `SliceArray`, `FormatTokenRef` from `basTEST_aeBibleCitationBlock.bas`
6. Remove `Test_VerifyCitationBlock_Negative` from `basTEST_aeBibleCitationBlock.bas` (covered by Stage 13a)
7. Run `Run_All_SBL_Tests` and `Test_VerifyCitationBlock`; verify all pass

---

## Q: How Does Stage 13a Deal with the Semicolon?

### Part 1 — Semicolons are already consumed before Stage 13a sees anything

`ListDetection` (Stage 8) runs before any shorthand logic. It splits the raw input into
segments and stores them in `t.Segments`. `ComposeList_Internal` then iterates over those
segments. By the time Stage 13a checks a segment like `"23:1"`, the semicolons are gone —
they were the split delimiters. Stage 13a requires no change to handle semicolons; it
operates on already-split segments.

Trace for `"Ps 19:1; 23:1; 28:7"`:

```
ListDetection input:  "Ps 19:1; 23:1; 28:7"
  no comma found
  semicolon found -> Split(";") -> ["Ps 19:1", " 23:1", " 28:7"]

ComposeList_Internal loop:
  seg = "Ps 19:1"  -> ParseReferenceRef -> Psalms 19:1   (prevRef set)
  seg = "23:1"     -> Stage 13a: colon present, left of colon is "23" (numeric)
                      -> inherit BookID from prevRef (Psalms)
                      -> Chapter=23, Verse=1  -> Psalms 23:1
  seg = "28:7"     -> Stage 13a: same rule -> Psalms 28:7
```

### Part 2 — `ListDetection` is an either/or splitter: this breaks mixed input

`ListDetection` checks for comma **first**. If ANY comma is present in the entire input
string, it splits on comma and exits — semicolons are never checked.

```vb
If InStr(rawInput, ",") > 0 Then
    parts = Split(rawInput, ",")   ' exits here; semicolons ignored
    ...
    Exit Function
End If
If InStr(rawInput, ";") > 0 Then
    parts = Split(rawInput, ";")
    ...
End If
```

The full citation block contains BOTH delimiters: semicolons between references and commas
within verse sub-lists (`145:8-9,17`). For that input `ListDetection` splits on commas
first:

```
ListDetection input:  "Ps 145:8-9,17; Isa 40:28"
  comma found -> Split(",") -> ["Ps 145:8-9", "17; Isa 40:28"]

ComposeList_Internal loop:
  seg = "Ps 145:8-9"   -> range -> OK
  seg = "17; Isa 40:28" -> ParseReferenceRef("17; Isa 40:28") -> ERROR (semicolon inside)
```

**Conclusion:** Stage 13a (book-context propagation) is correct and sufficient for
pure-semicolon inputs — the common case for cross-book reference lists with no
comma-separated verse sub-lists. But the full citation block (which has `145:8-9,17`)
requires a two-level split that `ListDetection` does not currently perform.

### The two-level split required for citation blocks

Study Bible citation blocks use a two-level delimiter structure:

| Level | Delimiter | Separates |
|---|---|---|
| Outer | `;` | Major segments (each has its own book/chapter context) |
| Inner | `,` | Verse sub-items within a single chapter (`145:8-9,17`) |

The correct algorithm:
1. Split on `;` → major segments
2. For each major segment, split on `,` → verse sub-items
3. Apply book-context and chapter-context propagation across both levels

This is exactly what `TokenizeCitationBlock` in `basTEST_aeBibleCitationBlock.bas`
implements. The migration plan must therefore treat `TokenizeCitationBlock` not as logic
to discard but as the specification for a new entry point in the class.

### Revised plan for `ListDetection`

`ListDetection` cannot be changed to "split on semicolon, then comma within segments"
without breaking existing callers that pass purely comma-delimited input
(`"John 3:16, 18, 20-22"`). The existing behaviour must be preserved for `ComposeList`.

The correct approach is a new public method on the class:

```vb
Public Function ParseCitationBlock(rawBlock As String) As Collection
```

This method:
1. Calls `NormalizeRawInput` (strip line breaks, normalize en-dash)
2. Splits on `;` → major segments
3. For each major segment, detects book alias (Stage 13a rule) and splits on `,` for
   verse sub-items
4. Validates each atomic endpoint via `ValidateSBLReference(ModeSBL)`
5. Returns a `Collection` of canonical reference strings

`VerifyCitationBlock` in `basTEST_aeBibleCitationBlock.bas` is then a thin wrapper that
calls `ParseCitationBlock` and reports PASS/FAIL per item — no `BlockToken` type, no
custom tokenizer. `Test_VerifyCitationBlock` calls `VerifyCitationBlock` unchanged.

### Updated implementation order

The step "Add Stage 13a inline case to `ComposeList_Internal`" from the earlier plan
still stands — it fixes the common case for pure-semicolon inputs and is the right place
for that logic. But `ComposeList_Internal` / `ListDetection` are NOT sufficient for the
full citation block with mixed delimiters. That requires the additional step:

**Step 2b:** Add `Public Function ParseCitationBlock` to `aeBibleCitationClass.cls`,
implementing the two-level split. This is the class-level home for the logic currently
in `TokenizeCitationBlock`.
