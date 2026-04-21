# Adding a New Test to the Bible Test Class

## Overview

The Bible QA test suite lives in `aeBibleClass.cls` and is driven by
`basTEST_aeBibleClass.bas`. Tests are numbered sequentially. Adding a new test
requires coordinated changes in four locations within the class, plus updating the
test count constant. This document records the process used to add Test 73
(`CountInvisibleCharacters`) and serves as the template for all future additions.

---

## Architecture

```
basTEST_aeBibleClass.bas
  └─ RUN_THE_TESTS([n | "varDebug"])   ← public entry point (Immediate Window)
       └─ aeBibleClassTest(varDebug)
            └─ aeBibleClass.TheBibleClassTests(varDebug)  ← Property Get on class instance
                 └─ RunBibleClassTests(varDebug)           ← Private Function
                      ├─ InitializeGlobalResultArrayToMinusOne
                      ├─ MakeSkipTestArray
                      ├─ Expected1BasedArray                ← expected values per test
                      └─ RunTest(n) for n = 1 to MaxTests
                           ├─ GetPassFail(n)                ← calls the test function
                           ├─ Debug.Print result line
                           └─ OutputTestReport(n)           ← writes rpt/TestReport.txt
```

Tests store their result in `ResultArray(n)` and compare against
`oneBasedExpectedArray(n)`. Pass = result equals expected; Fail = anything else.

---

## Checklist: Adding Test N

| Step | Location | What to change |
|------|----------|----------------|
| 1 | `aeBibleClass.cls` line with `MaxTests` | Increment constant by 1 |
| 2 | `Expected1BasedArray` — `values` array | Append expected result for test N |
| 3 | `Expected1BasedArray` — comment line above values | Append `N` to the RunTest list |
| 4 | `GetPassFail` Select Case | Add `Case N: ResultArray(TestNum) = YourFunction()` |
| 5 | `RunBibleClassTests` — RunTest call sequence | Add `RunTest (N)` after `RunTest (N-1)` |
| 6 | `RunTest` Select Case | Add `Case N: Debug.Print ... "YourFunctionName"` |
| 7 | `OutputTestReport` Select Case | Add `Case N: AppendToFile ...` with same label |
| 8 | `aeBibleClass.cls` body | Add `Private Function YourFunction() As Long` |

If the function already exists in another module as `Private`, decide whether to:
- **Copy** the logic into the class as a new `Private Function` returning `Long`
  (cleanest — self-contained, follows existing pattern)
- **Make it `Public`** in the source module and call it from the class
  (avoids duplication but creates a cross-module dependency)

For Test 73, the logic was **copied** into the class as a Long-returning variant.
The source function in `basTEST_aeBibleConfig.bas` was also changed from
`Private` to `Public` so it remains independently callable via `TestInvisible`.

---

## Test 73 — CountInvisibleCharacters

### What it tests

Invisible Unicode characters that are visually silent but can corrupt:
- Word's Find/Replace results
- Style normalization runs
- USFM export output

| Code point | Name |
|------------|------|
| U+200B | ZERO WIDTH SPACE |
| U+200C | ZERO WIDTH NON-JOINER |
| U+200D | ZERO WIDTH JOINER |
| U+FEFF | ZERO WIDTH NO-BREAK SPACE (byte-order mark) |
| U+2060 | WORD JOINER |

### Expected result

`0` — the Bible document must contain none of these characters.

### Function added to `aeBibleClass.cls`

```vba
Private Function CountInvisibleCharacters() As Long
    Dim r As Word.Range
    Dim targets As Variant
    Dim counts() As Long
    Dim i As Long, total As Long
    targets = Array(ChrW(&H200B), ChrW(&H200C), ChrW(&H200D), ChrW(&HFEFF), ChrW(&H2060))
    ReDim counts(UBound(targets))
    For Each r In ActiveDocument.StoryRanges
        For i = 0 To UBound(targets)
            counts(i) = counts(i) + UBound(Split(r.Text, targets(i)))
        Next i
    Next r
    For i = 0 To UBound(counts) : total = total + counts(i) : Next i
    CountInvisibleCharacters = total
End Function
```

Scans all story ranges (body, headers, footers, footnotes, text boxes). Uses `Split`
on each target character: `UBound(Split(text, char))` equals the count of occurrences
because splitting on a character that appears N times produces N+1 parts.

### Source function in `basTEST_aeBibleConfig.bas`

`CountInvisibleCharacters` in that module was `Private` and returned `String`
("0" = clean; multi-line report = violations). Changed to `Public` so it is still
callable by `TestInvisible` and accessible directly from the Immediate Window.
The class method is a separate `Long`-returning variant — no naming conflict because
the class method is resolved first within the class scope.

---

## Running the test

```
' Run test 73 only:
RUN_THE_TESTS(73)

' Run all tests (includes 73):
RUN_THE_TESTS

' Check detailed report for test 73:
' Open rpt/TestReport.txt after a full run
```

Pass output in Immediate Window:
```
PASS        Copy ()     Test = 73       0               0               CountInvisibleCharacters
```

Fail output (example — 3 invisible chars found):
```
FAIL!!!!    Copy ()     Test = 73       3               0               CountInvisibleCharacters
```

If the test fails, use `TestInvisible` in `basTEST_aeBibleConfig.bas` for a detailed
per-character breakdown with Unicode labels.

---

## Files changed for Test 73

| File | Change |
|------|--------|
| `src/aeBibleClass.cls` | `MaxTests` 72 → 73; values array +1 entry; `Case 73` in `GetPassFail`, `RunTest`, `OutputTestReport`; `RunTest(73)` in sequence; `CountInvisibleCharacters` function added |
| `src/basTEST_aeBibleConfig.bas` | `CountInvisibleCharacters` visibility `Private` → `Public` |

Import both files into the VBA editor after making changes.
