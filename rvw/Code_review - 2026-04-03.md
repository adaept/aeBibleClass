# Code Review: Negative Test Design in `basTEST_aeBibleCitationBlock`

**Date:** 2026-04-03
**Module:** `src/basTEST_aeBibleCitationBlock.bas`
**Function:** `Test_VerifyCitationBlock_Negative`

---

## Problem

`Test_VerifyCitationBlock_Negative` was written as a demonstration, not a test. It called
`VerifyCitationBlock` three times with invalid input and expected to see three `FAIL` lines
printed. But it made no assertions — no call to `aeAssertClass.AssertTrue` — so the test
harness (`Run_All_SBL_Tests` / `aeAssertClass`) saw nothing. No PASS was recorded, no FAIL
was recorded. The three negative cases were invisible to the test summary.

The deeper conceptual error: a negative test **passes** when the code correctly rejects
invalid input. "Three failures expected" in the banner described the behavior of the
*validator under test*, not the outcome of the *test itself*. That inversion caused the
confusion. Seeing `FAIL: Cannot resolve alias: "Jerimiah"` in the Immediate Window is a
**PASS** for the test — the validator did its job.

---

## Standard Used in This Project

`basTEST_aeBibleCitationClass.bas` / `Run_All_SBL_Tests` uses `aeAssertClass`:

```vb
aeAssert.AssertTrue condition, label
```

- `condition = True` → prints `PASS: label`, increments `mTestsRun`
- `condition = False` → prints `FAIL: label`, increments `mTestsFailed`

Every test outcome — positive or negative — must go through `AssertTrue` so the harness
can tally results. A negative test simply uses the inverted condition:

```vb
' The test PASSES when the validator correctly rejects the bad input
aeAssert.AssertTrue failCount >= 1, "Bad alias correctly rejected"
```

---

## Solution

### Change 1 — `VerifyCitationBlock`: `Sub` → `Function` returning `Long`

`VerifyCitationBlock` was a `Sub` with no return value. The caller had no way to know how
many failures were detected. Promoted to a `Function` returning `failCount As Long`.

```vb
' Before
Public Sub VerifyCitationBlock(rawBlock As String)
    ...
    Debug.Print "--- " & passCount & " passed, " & failCount & " failed. ---"
End Sub

' After
Public Function VerifyCitationBlock(rawBlock As String) As Long
    ...
    Debug.Print "--- " & passCount & " passed, " & failCount & " failed. ---"
    VerifyCitationBlock = failCount
End Function
```

Existing callers that ignore the return value (`Test_VerifyCitationBlock`) require no
change — calling a `Function` as a `Sub` is legal in VBA.

### Change 2 — `Test_VerifyCitationBlock_Negative`: assertions added

Each case now captures the return value and asserts that at least one failure was detected.
A local `aeAssertClass` instance is used so the test works both standalone and when called
from `Run_All_SBL_Tests`.

```vb
' Before — no assertions; test outcomes invisible to harness
Public Sub Test_VerifyCitationBlock_Negative()
    Debug.Print "=== Test_VerifyCitationBlock_Negative (3 failures expected) ==="
    VerifyCitationBlock "Gen 1:1; Jerimiah 33:11; Mal 1:1"
    VerifyCitationBlock "Ps 103:8" & ChrW(8211) & "200"
    VerifyCitationBlock "Jer 99:1"
End Sub

' After — each case recorded as PASS/FAIL by the harness
Public Sub Test_VerifyCitationBlock_Negative()
    Debug.Print "=== Test_VerifyCitationBlock_Negative ==="
    Dim localAssert As aeAssertClass
    Set localAssert = New aeAssertClass
    localAssert.Initialize

    Dim fc1 As Long
    fc1 = VerifyCitationBlock("Gen 1:1; Jerimiah 33:11; Mal 1:1")
    localAssert.AssertTrue fc1 >= 1, "Case 1: Bad alias (Jerimiah) correctly rejected"

    Dim fc2 As Long
    fc2 = VerifyCitationBlock("Ps 103:8" & ChrW(8211) & "200")
    localAssert.AssertTrue fc2 >= 1, "Case 2: Verse out of range (Ps 103:8-200) correctly rejected"

    Dim fc3 As Long
    fc3 = VerifyCitationBlock("Jer 99:1")
    localAssert.AssertTrue fc3 >= 1, "Case 3: Chapter out of range (Jer 99:1) correctly rejected"

    localAssert.Terminate
    Set localAssert = Nothing
End Sub
```

---

## Expected Output After Fix

```
=== Test_VerifyCitationBlock_Negative ===
--- Case 1: Bad alias (Jerimiah) ---
FAIL [1001]: Cannot resolve alias: "Jerimiah" (segment: "Jerimiah 33:11")
PASS: Genesis 1:1
PASS: Malachi 1:1
--- 2 passed, 1 failed. ---
PASS: Case 1: Bad alias (Jerimiah) correctly rejected
--- Case 2: Verse out of range (Ps 103:8-200) ---
FAIL [1006]: Psalms 103:8-200 (end verse 200 failed ValidateSBLReference)
--- 0 passed, 1 failed. ---
PASS: Case 2: Verse out of range (Ps 103:8-200) correctly rejected
--- Case 3: Chapter out of range (Jer 99:1) ---
FAIL [1006]: Jeremiah 99:1 (start verse failed ValidateSBLReference)
--- 0 passed, 1 failed. ---
PASS: Case 3: Chapter out of range (Jer 99:1) correctly rejected
------------------------------------------
 TEST SUMMARY
------------------------------------------
Tests Run:  3
Failures:   0
RESULT: PASS
```

The validator's `FAIL` lines (detecting bad input) are **evidence** that each test case
passed. The harness `RESULT: PASS` at the end confirms all three negative tests succeeded.

---

## Second run result — fix overwritten; new bug found

The fix to `Test_VerifyCitationBlock_Negative` (Issue 1 above) was applied but then
overwritten, restoring the original buggy fallthrough in `DetectBookAliasInSegment`.
The second test run produced the same failure:

```
--- Case 1: Bad alias (Jerimiah) ---
BAD  > ResolveBook(JERIMIAH)
PASS: Genesis 1:1
PASS: Genesis 1:11          <- spurious token
PASS: Malachi 1:1
--- 3 passed, 0 failed. ---
FAIL: Case 1: Bad alias (Jerimiah) correctly rejected
Tests Run: 3  Failures: 1  RESULT: FAIL
```

### Root cause — `DetectBookAliasInSegment` Case 2 silent fallthrough

When `TryResolveAlias("Jerimiah")` failed, Case 2 executed:

```vb
alias = ""
refPart = seg        ' "Jerimiah 33:11"
DetectBookAliasInSegment = False
```

`TokenizeCitationBlock` received `newBook = False` and skipped the error-handling block
entirely. It then parsed `refPart = "Jerimiah 33:11"` as a bare chapter:verse:

- colon found → `chStr = "Jerimiah 33"` → `IsNumeric` = False → chapter stays at 1 (Gen 1:1 context)
- `verseSpecStr = "11"` → token emitted as **Genesis 1:11**, no error

Three tokens passed, `fc1 = 0`, assertion failed.

### Fix re-applied — `DetectBookAliasInSegment` Case 2

Return `True` with `alias = p0` when alias resolution fails for a letter-prefix. This
routes the bad alias into `TokenizeCitationBlock`'s existing error-token path
(`TryResolveAlias` → False → `E_ALIAS_UNRESOLVED` → `GoTo NEXT_SEG`).

```vb
' Before (buggy — silent fallthrough)
alias = ""
refPart = seg
DetectBookAliasInSegment = False
Exit Function

' After (correct — signal as unresolved alias)
alias = p0
If UBound(parts) >= 1 Then
    refPart = Join(SliceArray(parts, 1), " ")
Else
    refPart = ""
End If
DetectBookAliasInSegment = True
Exit Function
```

Cases 2 and 3 were correct in both runs. After this fix all three cases should produce
`fc >= 1` and the summary should show `Tests Run: 3 / Failures: 0 / RESULT: PASS`.
