# Code Review: Stage 13b — Canonical Sort; Stage 17 En-Dash Option

**Date:** 2026-04-04

---

## Overview

Two capabilities needed for Study Bible document output:

1. **Canonical sort (`SortCitationBlock`) — Stage 13b.** Citation blocks may list
   references in arbitrary order. Stages 14–17 all assume canonical order already exists:
   Stage 14 compresses *adjacent* verses (adjacency requires ordering), Stages 15–17
   explicitly forbid reordering. No existing stage establishes canonical order — this is
   a gap in the pipeline. `SortCitationBlock` fills it, between Stage 13a and Stage 14.

2. **En-dash rendering (`RenderEnDash`) — Stage 17 option, not a new stage.** Stage 17
   (Canonical String Formatter) already specifies "Range separator → hyphen or en dash".
   `ParseCitationBlock` outputs ASCII hyphen internally (en-dash is normalized away at
   input by `NormalizeRawInput`). `RenderEnDash` is a per-item utility that implements
   the en-dash variant of the Stage 17 range separator, used by callers that produce
   Study Bible document output.

---

## Task 1 — En-Dash Output (Stage 17 Option)

### Problem

`NormalizeRawInput` converts en-dash to ASCII hyphen before any stage runs. The parser
always works with ASCII hyphen internally. `ParseCitationBlock` therefore returns canonical
strings with ASCII hyphen (e.g. `"Psalms 103:8-11"`). The Study Bible document requires
`"Psalms 103:8–11"`.

### Why a Separate Rendering Step

En-dash normalization at input is correct and must remain — `IsRangeSegment` and
`RangeDetection` rely on ASCII hyphen. The output form is a display concern distinct from
parsing. Keeping them separate preserves the parser invariant and makes the rendering step
optional for callers that do not produce Study Bible output.

### Solution

Add a new public function to the class:

```vb
Public Function RenderEnDash(ByVal canon As String) As String
```

Replaces every ASCII hyphen (`-`, Chr(45)) with en-dash (`–`, ChrW(8211)) in a canonical
reference string. Safe as a simple `Replace` because canonical book names (SBL style)
never contain hyphens; the only hyphens in a canonical string are verse range separators.

**Examples:**

| Input (canonical)         | Output (rendered)          |
|---|---|
| `"Psalms 103:8-11"`       | `"Psalms 103:8–11"`        |
| `"1 Chronicles 29:10-13"` | `"1 Chronicles 29:10–13"` |
| `"Psalms 23:1"`           | `"Psalms 23:1"` (unchanged) |

### Typical Call Pattern

```vb
Dim items As Collection
Set items = aeBibleCitationClass.ParseCitationBlock(rawBlock)
' optional sort (Task 2)
Set items = aeBibleCitationClass.SortCitationBlock(items)
' render for Study Bible output
Dim item As Variant
For Each item In items
    Debug.Print aeBibleCitationClass.RenderEnDash(CStr(item))
Next item
```

### Implementation Notes

- Add `On Error GoTo PROC_ERR` / `PROC_EXIT` / `PROC_ERR` / `MsgBox` / `Resume PROC_EXIT`
  standard handler.
- `Replace(canon, "-", ChrW(8211))` is the entire body. No loop, no regex.
- Placed after `ParseCitationBlock` in the class (same region, ~line 2645).

### Test

Add `Test_RenderEnDash` to `basTEST_aeBibleCitationBlock.bas`:

```vb
' Verify range entry gets en-dash
Dim rendered As String
rendered = aeBibleCitationClass.RenderEnDash("Psalms 103:8-11")
aeAssert.AssertEqual "Psalms 103:8" & ChrW(8211) & "11", rendered, _
    "RenderEnDash: range gets en-dash"

' Verify non-range entry is unchanged
rendered = aeBibleCitationClass.RenderEnDash("Psalms 23:1")
aeAssert.AssertEqual "Psalms 23:1", rendered, _
    "RenderEnDash: non-range unchanged"
```

---

## Task 2 — Canonical Sort (Stage 13b)

### Problem

A citation block may present references in sermon or thematic order (e.g.
`"John 3:16; Gen 1:1; Ps 23:1"`). For Study Bible document output, references must follow
canonical book order (Gen=1, Exod=2, ..., Rev=66), then chapter, then verse.

### Why Sort the Output Collection, Not the Input Segments

Sorting input segments would require re-establishing book-context propagation after
reordering — a segment like `"23:1"` inherits its book from the preceding segment, so
reordering segments breaks that chain. Sorting the output Collection avoids this: every
item is already a fully-qualified canonical string with an explicit book name.

### Solution

Add a new public function to the class:

```vb
Public Function SortCitationBlock(ByVal items As Collection) As Collection
```

**Algorithm:**

1. For each item in `items`, extract `(BookID, Chapter, StartVerse)`:
   - Find the last space → left part is book name, right part is `Chapter:Verse[-End]`
   - Call `ResolveAlias(bookName, BookID)` to get the numeric BookID (1–66)
   - Parse `Chapter` from left of `:`
   - Parse `StartVerse` from right of `:`, stopping at `-` if a range
2. Compute sort key: `BookID * 100000000 + Chapter * 100000 + StartVerse`
   (supports up to 999 chapters and 99999 verses — well above any biblical maximum)
3. Load items and keys into parallel arrays
4. Sort using insertion sort (O(n²), acceptable for citation block sizes ≤ ~50 items)
5. Return a new Collection in sorted order

**Note on `ParseCanonicalRef`:** The private function `ParseCanonicalRef` (line 1827)
already extracts `(bookName, Chapter, Verse)` from a canonical string. However, since it
is Private, `SortCitationBlock` will inline the same "find last space" + `ResolveAlias`
logic rather than promoting `ParseCanonicalRef` to Public. Promoting it is an alternative
but widens the public API for a single caller.

### Qualifying Condition for Sort Key

The canonical string from `ParseCitationBlock` always has the form:

```
CanonicalBookName Chapter:StartVerse[-EndVerse]
```

The last space separates book name from the numeric part. `ResolveAlias` accepts the full
canonical book name (e.g. `"Psalms"`, `"1 Chronicles"`) — it is already in the alias map.

### Cross-Book Transition After Sort

After sorting, book context is irrelevant — every output item is fully qualified. The
sorted Collection can be passed directly to `RenderEnDash` or to a Study Bible formatter.

### Implementation Notes

- Add standard `On Error GoTo PROC_ERR` handler.
- Declare: `Dim keys() As Long`, `Dim vals() As String`, sized to `items.Count`.
- Use `On Error Resume Next` around `ResolveAlias` call; restore `On Error GoTo PROC_ERR`
  immediately after. If `ResolveAlias` fails (should not happen for output of
  `ParseCitationBlock`), skip the item or raise.
- Placed after `RenderEnDash` in the class (~line 2660).

### Test

Add `Test_SortCitationBlock` to `basTEST_aeBibleCitationBlock.bas`:

```vb
' Out-of-order input: John before Genesis
Dim raw As String
raw = "John 3:16; Gen 1:1; Ps 23:1"
Dim sorted As Collection
Set sorted = aeBibleCitationClass.SortCitationBlock( _
    aeBibleCitationClass.ParseCitationBlock(raw))
aeAssert.AssertEqual 3, sorted.Count, "Sort: count preserved"
aeAssert.AssertEqual "Genesis 1:1",  sorted(1), "Sort: Gen first"
aeAssert.AssertEqual "Psalms 23:1",  sorted(2), "Sort: Ps second"
aeAssert.AssertEqual "John 3:16",    sorted(3), "Sort: John third"

' Same-book, multi-chapter: chapter order within book
raw = "Ps 103:8; Ps 19:1; Ps 68:5"
Set sorted = aeBibleCitationClass.SortCitationBlock( _
    aeBibleCitationClass.ParseCitationBlock(raw))
aeAssert.AssertEqual "Psalms 19:1",  sorted(1), "Sort: Ps 19 before 68"
aeAssert.AssertEqual "Psalms 68:5",  sorted(2), "Sort: Ps 68 before 103"
aeAssert.AssertEqual "Psalms 103:8", sorted(3), "Sort: Ps 103 last"
```

---

## Updated `Run_All_SBL_Tests` Sequence

No changes to `Run_All_SBL_Tests` — `Test_RenderEnDash` and `Test_SortCitationBlock`
are integration tests that live in `basTEST_aeBibleCitationBlock.bas` and are called from
`Test_VerifyCitationBlock` or independently.

---

## Implementation Order

1. Add `RenderEnDash` to `aeBibleCitationClass.cls` with standard handler
2. Add `Test_RenderEnDash` to `basTEST_aeBibleCitationBlock.bas`
3. Add `SortCitationBlock` to `aeBibleCitationClass.cls` with standard handler
4. Add `Test_SortCitationBlock` to `basTEST_aeBibleCitationBlock.bas`
5. Run `Test_VerifyCitationBlock` and new tests; verify all pass

---

## Goal State After Implementation

### `src/aeBibleCitationClass.cls` additions

| New item | Purpose |
|---|---|
| `Public Function RenderEnDash` | Replace ASCII hyphen with en-dash in a canonical string |
| `Public Function SortCitationBlock` | Sort a Collection of canonical strings by BookID, Chapter, StartVerse |

### `src/basTEST_aeBibleCitationBlock.bas` additions

| New item | Purpose |
|---|---|
| `Test_RenderEnDash` | Verify en-dash replacement on range and non-range strings |
| `Test_SortCitationBlock` | Verify canonical book order and within-book chapter order |
