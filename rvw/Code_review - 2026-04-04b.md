# Code Review: In-Depth Analysis — All Source Files

**Date:** 2026-04-04

---

## Scope

Full review of `src/aeBibleCitationClass.cls` and `src/basTEST_aeBibleCitationBlock.bas`
covering: On Error standard compliance, security and injection risk, input bounds,
multi-language handling, test gaps, and general code quality.

---

## 1 — On Error Standard Compliance

**Standard pattern required:**

```vb
Public Sub/Function ProcName()
    On Error GoTo PROC_ERR
    ...
PROC_EXIT:
    Exit Sub/Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure X of Module Y"
    Resume PROC_EXIT
End Sub/Function
```

**Verdict: COMPLIANT**

All 45+ Public Sub/Function procedures follow the standard pattern. Error messages
include `Erl`, error number, description, procedure name, and module name consistently
across all modules.

**Intentional `On Error Resume Next` blocks (acceptable):**

Several functions temporarily suppress errors for optional alias lookups:

```vb
On Error Resume Next
ttCan = ResolveAlias(tt, ttID)
On Error GoTo PROC_ERR   ' always restored immediately after
```

Used in `ParseCitationBlock` Cases 1 and 2 and in `SortCitationBlock`. Pattern is
correct: suppress → test result → restore. No path leaves `On Error Resume Next`
active at exit.

**`Class_Initialize` / `Class_Terminate`:**

Use `Debug.Print` for error logging rather than `MsgBox`. Correct — initialization
hooks must not block the UI.

---

## 2 — Security and Injection Risk

**Verdict: LOW RISK**

### 2.1 Shell / Run / Filesystem Calls

None found. No command injection surface exists.

### 2.2 Clipboard Write

```vb
Dim dataObj As Object
Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
dataObj.SetText Result
dataObj.PutInClipboard
```

- CLSID is hardcoded — not user-supplied.
- `Result` is composed entirely of parser output: canonical book names, digits,
  colons, hyphens, semicolons. No user-supplied raw string reaches the clipboard
  without being parsed and reconstructed.
- `SetText` has a 2 GB internal limit; a realistic 100-token block is under 2 KB.

**Risk: NONE**

### 2.3 MsgBox Content from User Input

Two MsgBox calls include content derived from user input:

- Validation report (`report`) — built from parsed canonical references, not raw input.
- Replacement prompt (`Result`) — parser output only.

Worst case for a 35-token block with all failures: ~600 characters. MsgBox limit is
~32 KB. **Risk: NONE**

### 2.4 `Err.Raise` with User-Supplied Content

```vb
Err.Raise vbObjectError + 1001, , _
    "Cannot resolve alias: """ & p0 & """"
```

`p0` is a single space-delimited token that has passed `Like "[A-Za-z]*"`, so it
contains only ASCII letters. Maximum practical length: ~10 characters.
**Risk: NONE**

### 2.5 `CreateObject` Calls

All ProgIDs and CLSIDs are hardcoded constants (`Scripting.Dictionary`,
`MSForms.DataObject`). None are derived from user input. **Risk: NONE**

---

## 3 — Input Length and Bounds

**Verdict: ACCEPTABLE with one recommended guard**

### 3.1 No Upper Bound on `rawBlock`

`ParseCitationBlock` accepts any string length with no guard:

```vb
Public Function ParseCitationBlock(ByVal rawBlock As String) As Collection
    Dim normalized As String
    normalized = NormalizeRawInput(rawBlock)
    segs = Split(normalized, ";")   ' no limit
```

In the Word document context users type citation blocks manually; pathological inputs
are implausible. However, a defensive upper-bound check is good practice.

**Recommendation:** add to `ParseCitationBlock` after `NormalizeRawInput`:

```vb
If Len(normalized) > 10000 Then
    Err.Raise vbObjectError + 1002, , "Citation block too long (max 10000 chars)"
End If
```

### 3.2 Character-by-Character Backward Loops

Pattern used in multiple locations to find the last space:

```vb
For k = Len(canon) To 1 Step -1
    If Mid$(canon, k, 1) = " " Then lastSp = k: Exit For
Next k
```

All instances operate on canonical reference strings (max ~25 characters) or
`rawBlock` segments (typically <50 characters). No performance risk.

**Recommendation (low priority):** replace with `InStrRev(s, " ")` — cleaner and
avoids the loop entirely. Not a safety issue.

### 3.3 Collection Growth

`ParseCitationBlock` adds one item per resolved verse reference with no cap. A
100-reference block produces a 100-item Collection (~10 KB) — acceptable. The
`VerifyCitationBlockReport` MsgBox would become long before memory is a concern.

---

## 4 — Multi-Language and Non-ASCII Input

**Verdict: BY DESIGN (English SBL only) — graceful failure on non-ASCII**

### 4.1 `Like "[A-Za-z]*"` Character Class

```vb
If Not newBook And p0 Like "[A-Za-z]*" Then
```

Accepts only ASCII letters. Any token beginning with a non-ASCII character (accented,
Cyrillic, Hebrew, Arabic, etc.) silently falls through to Case 3 (no new book
detected), leaving the previous book context active.

This is by design — the alias map contains English SBL abbreviations only. However
the failure mode is silent (wrong output rather than a clear error).

**Recommendation:** if non-ASCII is detected as `p0`, raise a descriptive error
rather than silently misattributing the reference:

```vb
If Not newBook And p0 Like "[A-Za-z]*" Then
    ' ... existing Case 2 ...
ElseIf Not newBook Then
    Err.Raise vbObjectError + 1003, , _
        "Unrecognised token (non-ASCII or unsupported script): """ & p0 & """"
End If
```

### 4.2 `UCase$` on Non-ASCII

`ResolveAlias` does `key = UCase$(Trim$(abbr))` before dictionary lookup. VBA's
`UCase$` is locale-dependent for characters outside A–Z. For the English alias map
this is irrelevant — `UCase$("Ps") = "PS"` is always correct. If non-English aliases
were ever added, locale-specific upcasing could cause lookup misses.

**No action required** for current scope.

### 4.3 `Scripting.Dictionary` Case Sensitivity

The dictionary is created with default `vbBinaryCompare`. All keys are stored as
`UCase$` strings and all lookups use `UCase$`, so case sensitivity is neutralised
correctly regardless of dictionary mode.

### 4.4 RTL Text

Right-to-left characters (Hebrew, Arabic) in the input would produce a token that
fails `Like "[A-Za-z]*"` and Case 2 would raise `vbObjectError + 1001`. Graceful
failure — no data corruption.

### 4.5 Future Internationalisation Path

To support non-English abbreviation sets (e.g. German Luther Bible abbreviations),
the minimum changes would be:

1. Extend `GetBookAliasMap` with locale-specific entries (or load a secondary map).
2. Change `Like "[A-Za-z]*"` to `Like "[A-Za-z\u00C0-\u024F]*"` or use `Asc(p0) > 0`
   to allow extended Latin characters.
3. Keep `UCase$` for ASCII ranges; use `StrConv(key, vbUpperCase)` for better
   locale handling.

---

## 5 — Test Gaps

### 5.1 `ToSBLShortForm` — No Unit Test

New function added in this session. Covers 66 books including single-chapter shorthand
(`Jude 1:6` → `Jude 6`). No dedicated test exists.

**Recommended test:**

```vb
Public Sub Test_ToSBLShortForm()
    On Error GoTo PROC_ERR
    ' Multi-word book
    aeAssert.AssertEqual "1 Chr 29:10-13", _
        aeBibleCitationClass.ToSBLShortForm("1 Chronicles 29:10-13"), _
        "ToSBLShortForm: 1 Chronicles"
    ' Single-chapter book — chapter number omitted
    aeAssert.AssertEqual "Jude 6", _
        aeBibleCitationClass.ToSBLShortForm("Jude 1:6"), _
        "ToSBLShortForm: Jude single-chapter"
    ' Single-chapter range
    aeAssert.AssertEqual "Obad 3-5", _
        aeBibleCitationClass.ToSBLShortForm("Obadiah 1:3-5"), _
        "ToSBLShortForm: Obadiah range"
    ' Standard multi-chapter
    aeAssert.AssertEqual "Ps 23:1", _
        aeBibleCitationClass.ToSBLShortForm("Psalms 23:1"), _
        "ToSBLShortForm: Psalms"
PROC_EXIT:
    Exit Sub
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure Test_ToSBLShortForm of Module basTEST_aeBibleCitationBlock"
    Resume PROC_EXIT
End Sub
```

### 5.2 `VerifyCitationBlockReport` — No Unit Test

Used by `RepairCitationBlockInParagraph` but has no standalone test. A basic test
should verify that it returns the correct pass/fail counts and a non-empty report
string for a known input.

### 5.3 `ctxChapter` Reset on Book Switch — No Explicit Test

The fix `If newBook Then ctxChapter = 0` is covered implicitly by the 35-token
integration test but has no targeted test. A minimal explicit test:

```vb
' "2 Pet 2:4; Jude 6" — Jude must not inherit chapter 2 from 2 Peter
Dim items As Collection
Set items = aeBibleCitationClass.ParseCitationBlock("2 Pet 2:4; Jude 6")
aeAssert.AssertEqual "Jude 1:6", items(2), "ctxChapter reset: Jude after 2 Pet"
```

### 5.4 `Chr(11)` Normalization — No Unit Test

The `NormalizeRawInput` fix for Word forced line breaks has no test. Suggested:

```vb
' Block with Chr(11) between book switch
Dim raw As String
raw = "1 Chr 29:10-13;" & Chr(11) & "Ps 19:1-2"
Dim items As Collection
Set items = aeBibleCitationClass.ParseCitationBlock(raw)
aeAssert.AssertEqual "Psalms 19:1-2", items(2), "Chr(11) normalization"
```

### 5.5 Single-Chapter Books — Partial Coverage

| Book | BookID | Tested |
|---|---|---|
| Obadiah | 31 | No |
| Philemon | 57 | No |
| 2 John | 63 | No |
| 3 John | 64 | Partial |
| Jude | 65 | Yes (new) |

### 5.6 Edge Cases — Missing

| Case | Status |
|---|---|
| Empty string input to `ParseCitationBlock` | Not tested |
| Single reference (`"John 3:16"`) | Not tested |
| Block with only single-chapter books | Not tested |
| Very long block (100+ tokens) | Not tested |
| Block ending with trailing semicolon | Not tested |
| All-whitespace input | Not tested |

---

## 6 — General Code Quality

### 6.1 `InStrRev` Opportunity

Seven locations use a backward character loop to find the last space:

```vb
For k = Len(canon) To 1 Step -1
    If Mid$(canon, k, 1) = " " Then lastSp = k: Exit For
Next k
```

`InStrRev(canon, " ")` is a direct replacement — same result, no loop.
Low priority; no correctness or safety impact.

### 6.2 Debug Output Left in Production Code

`RepairCitationBlockInParagraph` contains:

```vb
Debug.Print "workRng = " & workRng.Text
Debug.Print "report = " & report
```

These were added for testing this session. Remove before release.

### 6.3 `xxxTest_AllBookAliases_STRICT` — Incomplete

`aeBibleCitationClass.cls` contains a procedure prefixed `xxx` (indicating disabled
or incomplete). Should be completed or removed before release.

### 6.4 `GetSBLCanonicalBookTable` — Data Errors

`GetSBLCanonicalBookTable` (line 964) has several issues:
- BookID 3 (Leviticus) is absent.
- BookID 61 (2 Peter) is absent.
- BookID 64 is assigned twice (3 John and Jude conflict).
- BookID 65 (Jude) is missing — it appears on line 1032 as `sbl.Add 64, "JUDE"`.

This function is not used by the current citation pipeline (`ToSBLShortForm` uses
`GetMaxChapter` + `GetCanonicalBookTable`, not `GetSBLCanonicalBookTable`), so there
is no current runtime impact. The function should be corrected or removed to avoid
future confusion.

---

## 7 — Summary

| Area | Status | Action |
|---|---|---|
| On Error standard | COMPLIANT | None |
| Security — injection | LOW RISK | None |
| Security — MsgBox length | SAFE | None |
| Input length bound | UNGUARDED | Add `Len > 10000` guard to `ParseCitationBlock` |
| Non-ASCII input | BY DESIGN | Optionally add explicit error for non-ASCII tokens |
| Multi-language path | NOT SUPPORTED | Document as out-of-scope; note extension path |
| `ToSBLShortForm` test | MISSING | Add `Test_ToSBLShortForm` |
| `VerifyCitationBlockReport` test | MISSING | Add standalone test |
| `ctxChapter` reset test | MISSING | Add targeted assertion |
| `Chr(11)` normalization test | MISSING | Add `NormalizeRawInput` test with `Chr(11)` |
| Single-chapter books (Obad, Phlm, 2 John) | PARTIAL | Add integration tests |
| Edge cases (empty, single, long) | MISSING | Add `Test_ParseCitationBlock_EdgeCases` |
| `InStrRev` refactor | LOW PRIORITY | Optional cleanup |
| Debug prints in production code | PRESENT | Remove before release |
| `xxxTest_AllBookAliases_STRICT` | INCOMPLETE | Complete or remove |
| `GetSBLCanonicalBookTable` data | ERRORS | Fix or remove |

---

## Questions for Clarification

1. **`GetSBLCanonicalBookTable`** — Is this function used anywhere outside the
   reviewed files, or is it safe to remove/replace?

2. **Multi-language scope** — Is there any plan to support non-English abbreviation
   sets (e.g. German, French Bible abbreviations)? If so, the `Like "[A-Za-z]*"`
   gate and alias map design need revisiting now rather than later.

3. **Input length guard** — Is 10,000 characters a reasonable upper bound for a
   citation block in this project, or should it be higher/lower?

4. **Debug prints** — Should `Debug.Print "workRng"` and `Debug.Print "report"` be
   removed now, or left for continued testing?

5. **`ToSBLShortForm` abbreviation choices** — The function uses standard SBL
   abbreviations (e.g. `Exod`, `1 Kgs`, `Phlm`). Are these the exact forms expected
   in the Study Bible document output, or does the document use a different set?
