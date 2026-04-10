# Code Review - 2026-04-09
## Long-Running Process Framework — Design and Implementation Plan

---

## § 1 — Background

`rvw/Code_review - 2026-04-08.md` documented the investigation and resolution of a
37-second double-block on GoToH1 navigation in a full-Bible `.docm`. The final
solution (Option B, §30–31 of that review) is:

- Defer GoToH1 execution via `Application.OnTime` (1-second delay) to escape the
  ribbon callback post-processing pass.
- Pre-warm the layout cache at document open: `WarmLayoutCacheDeferred` is scheduled
  5 seconds after ribbon load; it navigates to the last heading (Revelation) and back,
  forcing Word's layout engine to build the page layout table for the full document.

**The remaining problem:** the warm-up itself causes a ~20-second application block
at startup. The Immediate Window shows:

```
WarmLayoutCache: warming to pos=4119182 at 10:53:20
WarmLayoutCache: complete at 10:54:26
```

This is currently unavoidable with the synchronous `ActiveDocument.Range(lastPos,
lastPos).Select` call. The user wants to explore whether
`src/XLongRunningProcessCode.bas` provides a viable path to eliminate or mitigate
this startup block, and to build a general-purpose framework for other long-running
tasks in the project.

---

## § 2 — Description of XLongRunningProcessCode.bas (current state)

The file is a work-in-progress skeleton with the X prefix (excluded from normal test
runs; run manually from Immediate Window only). It provides:

| Routine                           | Purpose                                                                     |
|-----------------------------------|-----------------------------------------------------------------------------|
| `PauseWithDoEvents(ms)`           | Sleep loop with `DoEvents` to keep Word responsive                          |
| `StartOrResumeUpdate()`           | Entry point: load saved progress, start main loop                           |
| `StopUpdate()`                    | Set flag to stop, save progress                                             |
| `ResetProgress()`                 | Clear saved progress                                                        |
| `SaveProgress()` / `LoadProgress()` | Persist `lastProcessedParagraph` + `progressPercentage` to CustomDocumentProperties |
| `CustomPropertyExists()`          | Helper for safe property access                                             |
| `SetWordHighPriority()`           | WMI call to raise WINWORD.EXE process priority                              |
| `UpdateCharacterStyle()`          | Example task: per-character style updates with progress                     |
| `LongProcessSkeletonWithConsoleProgress()` | Main skeleton: batch loop over paragraphs with DoEvents + pause between batches |

The skeleton structure is:
1. Divide document paragraphs into batches of N.
2. Disable `ScreenUpdating` and `Options.Pagination` within each batch.
3. Call `DoEvents` after each paragraph to yield to Word.
4. Re-enable screen updating and pagination between batches.
5. Pause 60 seconds between batches (`PauseWithDoEvents 60000`).
6. Save progress after each batch so work can be resumed after a stop.

---

## § 3 — Proposed Changes

### 3a. Parameterisable task framework

The skeleton's `LongProcessSkeletonWithConsoleProgress` will be refactored so the
task logic is passed in rather than hard-coded. Each long-running task becomes a
self-contained callable unit.

### 3b. Extract the existing task

The current placeholder (`' CODE GOES HERE`) and the example task
(`UpdateCharacterStyle`) will be extracted to their own routines so the skeleton
remains task-agnostic.

### 3c. Logging to rpt/ via basLogger.bas

The framework will emit structured UTF-8 log output to `rpt/` rather than to the
Immediate Window only. `basLogger.bas` will be updated to support this.

### 3d. Rename with git mv

When complete, `src/XLongRunningProcessCode.bas` will be renamed using `git mv` to
preserve history. Candidate name: `src/basLongProcess.bas`.

---

## § 4 — Issues and Analysis

### Issue 1 — BLOCKER: DoEvents cannot interrupt Word's layout engine

**This is the most critical issue and a potential blocker for the primary goal
(eliminating the 20-second startup block).**

The startup block is caused by `ActiveDocument.Range(lastPos, lastPos).Select`.
When called on a cold document, Word's layout engine must synchronously build the
page layout table from the beginning of the document to position 4119182 before the
VBA line completes. This is Word native C++ code running inside `Range.Select`.

`DoEvents` only processes Windows messages **between** VBA statements. It has no
effect on blocking work happening **inside** a single Word API call. No amount of
`DoEvents` before or after `Range.Select` can interrupt the layout pass that happens
within it.

**Consequence:** The batch+DoEvents pattern in `LongProcessSkeletonWithConsoleProgress`
will NOT eliminate the startup block caused by `WarmLayoutCache`. The block will still
occur — it will simply happen inside one of the batched iterations.

**Possible mitigations (to be evaluated):**
- Navigate to intermediate heading positions in steps (Genesis → mid-Bible → Revelation)
  rather than one jump. Each step warms a portion of the layout cache; the block is
  distributed across smaller jumps with `DoEvents` between them. Whether this actually
  shortens individual blocks needs testing.
- Accept the 20-second block and improve user communication (status bar message is
  already in place).
- Investigate whether `Options.Pagination = False` before `Range.Select` reduces
  the block (it suppresses background repagination but may not affect the synchronous
  layout triggered by a navigation).

### Issue 2 — VBA has no first-class function pointers

The plan calls for passing a "task as a parameter." VBA does not support function
references or delegates. Two viable approaches exist:

**Option A — String name + Application.Run:**
```vba
RunLongTask "MyTaskModule.MyTaskSub"
```
Inside the framework, call `Application.Run taskName, i` for each iteration.
Simple but fragile: task names are stringly typed, refactoring can break silently.
Also, `Application.Run` requires the called sub to be in a non-Option Private module
(same restriction as `Application.OnTime`) or requires a fully qualified
`ProjectName.ModuleName.SubName` string.

**Option B — Interface (IaeLongProcessClass class):**
Define a class `IaeLongProcessClass` with a single method `ExecuteItem(itemIndex As Long)`.
Each task implements this interface. The framework holds a reference `As IaeLongProcessClass`.

This is the more robust pattern. It also allows tasks to carry their own state and
parameters. Recommended.

### Issue 3 — WarmLayoutCache does not fit the paragraph-batch model

The skeleton is designed for tasks that iterate over paragraphs (N per batch).
`WarmLayoutCache` does exactly one navigation — it has no items to batch. It could
be wrapped as a "task with 1 item" but that is awkward and adds ceremony for no
benefit.

Two task shapes exist in the project:
- **Batched/resumable:** RUN_THE_TESTS, USFM Export, UpdateCharacterStyle
  → iterate paragraphs in chunks, save progress, can be stopped and resumed.
- **Single-shot deferred:** WarmLayoutCache
  → one operation, already deferred via `Application.OnTime`.

The framework should distinguish these or provide a trivial path for single-shot tasks.

### Issue 4 — basLogger.bas performance: open/load/save on every write

The current `Log_WriteRaw` uses `ADODB.Stream` with `LoadFromFile` + `saveToFile` on
every call. For a process that logs hundreds of lines, this will be slow:
each write opens the file, loads all existing content into memory, appends one line,
and rewrites the entire file.

For long-running processes, the logger needs a sequential-write mode: open the
stream once at `Log_Init`, keep it open, write incrementally, close at `Log_Close`.

### Issue 5 — basLogger.bas uses Option Private Module

`basLogger.bas` declares `Option Private Module`, which means its Public subs are
not accessible from `basRibbonDeferred.bas` (which intentionally omits that
declaration). The framework module that calls the logger needs to be in a module
where `basLogger` is visible, or the logger's `Option Private Module` must be
removed. This needs a decision before implementation.

### Issue 6 — Progress storage in CustomDocumentProperties modifies the document

Saving `lastProcessedParagraph` to CustomDocumentProperties marks the document as
modified (`ActiveDocument.Saved = False`). For a read-only Bible navigation scenario,
this is undesirable — it will prompt the user to save on close. Consider an
alternative:
- Store progress in a sidecar file in `rpt/` (simpler, no document mutation).
- Store progress in memory only (acceptable if tasks are not expected to survive a
  Word crash).

### Issue 7 — SetWordHighPriority uses WMI (slow, privileged)

The WMI call to set process priority adds latency at task startup and requires WMI
service availability. For most tasks, this is not worth the cost. Should be optional
or removed from the default startup path.

### Issue 8 — Pause of 60 seconds between batches is hardcoded and very long

`PauseWithDoEvents (60000)` at line 259 pauses 60 seconds between every batch of
50 paragraphs. For a full Bible (~31,000 paragraphs = 620 batches), this would take
over 10 hours of pause time alone. This parameter needs to be configurable and
appropriate to the task.

---

## § 5 — Blockers Summary

| # | Blocker | Severity | Notes |
|---|---------|----------|-------|
| 1 | DoEvents cannot interrupt layout engine inside Range.Select | HIGH | May not achieve the primary goal of eliminating startup block; needs a different strategy |
| 2 | VBA has no function pointers | MEDIUM | Solvable via interface pattern (Option B above) |
| 5 | basLogger Option Private Module | LOW | Easy to fix; needs a decision |

---

## § 6 — Assessment: Is the plan beneficial and practical?

**Beneficial — yes, with modified scope.**

The framework is valuable for RUN_THE_TESTS and USFM Export, which genuinely iterate
over document content and can use the batch+DoEvents pattern effectively. These tasks
are good candidates and the investment will pay off.

For WarmLayoutCache, the framework provides a better home for the code (logging,
structured progress) but will not eliminate the 20-second block. That block is
inherent to Word's layout engine and requires a different mitigation strategy
(step-navigation, user communication, or simply accepting it as a one-time startup cost).

**Practical — yes, with realistic expectations.**

The skeleton already exists. The main work is:
- Defining the interface pattern for task injection.
- Fixing basLogger for sequential writes.
- Resolving the Option Private Module question.
- Extracting tasks to separate routines.
- Renaming the file.

None of these are technically risky. The risk is in overestimating what DoEvents can
do against Word's layout engine.

---

## § 7 — Implementation Plan

The plan assumes the interface pattern (Issue 2, Option B) is approved. No code
will be written until this plan is confirmed.

---

### Step 1 — Define IaeLongProcessClass interface class

Create `src/IaeLongProcessClass.cls`:
- One method: `ExecuteItem(itemIndex As Long) As Boolean`
  - Returns `True` to continue, `False` to stop.
- One property: `ItemCount() As Long` — total number of items (paragraphs, or 1 for
  single-shot tasks).
- One property: `TaskName() As String` — used in log output and progress sidecar filename.

**Architecture note:** `basLongProcess.bas` is a thin public skeleton only — entry
points (`StartOrResume`, `Stop`, `Reset`) that delegate all logic to `IaeLongProcessClass`
and its concrete implementations. No business logic lives in the bas file.

---

### Step 2 — Convert basLogger to aeLoggerClass

`src/basLogger.bas` becomes `src/aeLoggerClass.cls` (class module). This change:
- Makes logger instances explicit — each long-running task holds its own `aeLoggerClass`
  instance with its own file path, rather than sharing module-level state.
- Eliminates the `Option Private Module` visibility problem — a class is always
  instantiable from any module that can `New` it.

Redesign for sequential writes — hold the ADODB.Stream open between `Log_Init` and
`Log_Close`:
- `Log_Init` opens the stream, writes the session header, positions to end (append mode).
- `Log_Write` appends to the open stream without reloading the file.
- `Log_Close` flushes and closes the stream.

Add a comment at the top of the class explaining that `Option Private Module` is
intentionally absent (classes never have it) and why that matters for callers.

**First test of aeLoggerClass:** Wire `Run_All_SBL_Tests`
(`src/basTEST_aeBibleCitationClass.bas`) to write its test output to
`rpt/SBL_Tests.UTF8.txt`. This is a self-contained, low-risk integration test that
confirms the logger works correctly before it is used inside the long-process framework.

---

### Step 3 — Rename and refactor the skeleton module

```
git mv src/XLongRunningProcessCode.bas src/basLongProcess.bas
```

`basLongProcess.bas` becomes a thin public skeleton only. Its public subs are the
entry points callable from the Immediate Window or ribbon:

```vba
Public Sub StartOrResume(task As IaeLongProcessClass)
Public Sub StopTask()
Public Sub ResetTask(task As IaeLongProcessClass)
```

All batch loop logic, progress persistence, DoEvents/pause behaviour, and logging
move into the class layer. `basLongProcess.bas` holds no business logic.

Specific changes from the current skeleton:
- Remove hardcoded paragraph iteration — delegate to `task.ExecuteItem(i)`.
- Remove `SetWordHighPriority` from the default path (keep as a standalone opt-in sub).
- Make batch size and pause duration parameters of the class, not hardcoded constants.
- Remove the 60,000 ms hardcoded pause.

---

### Step 4 — Migrate progress storage to rpt/ sidecar file

Replace CustomDocumentProperties with a sidecar file:
`rpt/LongRunningProgress_{TaskName}.txt`

Format: simple key=value pairs. This avoids marking the document as modified.

---

### Step 5 — Extract UpdateCharacterStyle as an IaeLongProcessClass implementation

Create `src/aeUpdateCharStyleClass.cls` implementing `IaeLongProcessClass`.
The task logic from `UpdateCharacterStyle` in `XLongRunningProcessCode.bas` moves here.
The bas file retains only the entry-point stub delegating to this class.

---

### Step 6 — Create WarmLayoutCacheTask as IaeLongProcessClass implementation

DEFERRED — per § 8, item 2. Return to this after the framework is verified working
with `aeUpdateCharStyleClass` and the SBL Tests logger test.

---

### Step 7 — Wire RUN_THE_TESTS as a candidate task (design only at this step)

Document the interface that `RUN_THE_TESTS` would need to implement.
Do not implement yet — this is a design milestone to confirm the interface is
flexible enough before committing to the architecture.

---

### Step 8 — Integration testing

Test in order:
1. WarmLayoutCacheTask: confirm logging to rpt/, confirm stop/resume works.
2. UpdateCharacterStyleTask: confirm batch+DoEvents, progress persistence.
3. RUN_THE_TESTS (if approved after Step 7 review).

---

### Step 9 — Update normalize_vba.py if new identifiers are introduced

Add any new public identifiers from the framework to the casing normalizer rules.

---

## § 8 — Open Questions for User Decision

1. **Interface pattern (§4, Issue 2):** Confirm Option B (IaeLongProcessClass interface)
   is acceptable. This adds one new class module.

2. **WarmLayoutCache strategy (§4, Issue 1):** DEFERRED — return to this after the
   long-running task framework is correctly implemented.

3. **basLogger Option Private Module (§4, Issue 5):** RESOLVED — remove
   `Option Private Module` from `basLogger.bas` and add a comment at the top of the
   module explaining why it is absent (so the logger is callable from any module,
   including those that also omit `Option Private Module` such as `basRibbonDeferred`).

4. **Progress storage (§4, Issue 6):** RESOLVED — use a `rpt/` sidecar file to
   replace `CustomDocumentProperties`.

   **Why CustomDocumentProperties is problematic:** Writing to a document's custom
   properties modifies the document's internal state. Word then considers the document
   unsaved (`ActiveDocument.Saved = False`) even though no content has changed. For a
   Bible study document that users may have open for extended sessions, this causes a
   spurious "Do you want to save changes?" prompt on close every time a long-running
   task saves its progress — even if the user made no edits.

   **What the sidecar file does instead:** Progress is stored in a plain text file at
   `rpt/LongRunningProgress_{TaskName}.txt` alongside the other report files. The
   document itself is never touched. The file uses simple `key=value` pairs
   (e.g. `LastProcessedParagraph=1450`, `ProgressPercentage=4.68`) that are easy to
   read, edit, and delete manually if a task needs to be reset. The file is also
   visible in git status, making it easy to confirm whether a task has run.

5. **Rename target:** RESOLVED — `src/basLongProcess.bas` (from `XLongRunningProcessCode.bas`
   via `git mv`). Interface class: `src/IaeLongProcessClass.cls`. Logger class:
   `src/aeLoggerClass.cls` (from `basLogger.bas`). The bas file is a thin public
   skeleton only; all logic lives in the class layer.

---

## § 9 — The `I` Prefix Convention for Interfaces

**Question:** Is the `I` prefix a standard way to indicate "Interface"?

**Answer:** Yes — with the clarification that the class will be named `IaeLongProcessClass`
to be consistent with the `ae` prefix and `Class` suffix used on all other classes in
this project.

---

### Where the `I` prefix comes from

In statically-typed languages such as Java, C#, and TypeScript, an **interface** is a
formal language construct: a type that declares method signatures with no implementation.
Any class that declares it implements the interface is checked by the compiler — missing
or mismatched methods are compile errors.

The `I` prefix (e.g., `IDisposable`, `IEnumerable`, `IComparable`) is so universal in
those languages that it is part of official naming guidelines. It signals to readers:
*this type is a contract, not an implementation.*

---

### What VBA has instead

VBA has no `Interface` keyword. It achieves the same pattern through **class modules
with `Implements`**:

```vba
' IaeLongProcessClass.cls — acts as the interface (empty method bodies only)
Public Function ExecuteItem(itemIndex As Long) As Boolean
End Function

Public Property Get ItemCount() As Long
End Property

Public Property Get TaskName() As String
End Property
```

```vba
' aeUpdateCharStyleClass.cls — concrete implementation
Implements IaeLongProcessClass

Private Function IaeLongProcessClass_ExecuteItem(itemIndex As Long) As Boolean
    ' actual task logic here
End Function

Private Property Get IaeLongProcessClass_ItemCount() As Long
    IaeLongProcessClass_ItemCount = ActiveDocument.Paragraphs.Count
End Property

Private Property Get IaeLongProcessClass_TaskName() As String
    IaeLongProcessClass_TaskName = "UpdateCharacterStyle"
End Property
```

The calling code in `basLongProcess.bas` holds a variable `As IaeLongProcessClass`
and calls `task.ExecuteItem(i)`. VBA resolves which class is actually running at
runtime — this is polymorphism achieved through COM's `Implements` mechanism rather
than compiler-enforced contracts.

The `I` prefix is the correct convention here because:
- It tells readers the class is not meant to be instantiated directly — it is a contract.
- It is recognized by any developer familiar with COM/ActiveX patterns (which VBA is built on).
- Word's own object model uses it: `IRibbonUI`, `IRibbonControl` — both already appear
  in this project.

---

### The practical difference between VBA and C#

| | C# interface | VBA `Implements` |
|---|---|---|
| Keyword | `interface` | class module with empty methods |
| Enforcement | compiler | compile error if method missing |
| `I` prefix | official guideline | convention only |
| Multiple interfaces per class | yes | yes |
| Abstract methods | yes (no body allowed) | no (empty body required) |

The VBA version is "interface by convention." If a class declares
`Implements IaeLongProcessClass` but does not implement all methods, VBA raises a
compile error when the project is compiled — so there is enforcement, just not as
early or as explicit as a true interface language.

---

### NOTE: Naming decision

The interface class is named **`IaeLongProcessClass`** — the `I` prefix marks it as
an interface contract; the `ae` prefix and `Class` suffix are consistent with all
other classes in this project (`aeRibbonClass`, `aeBibleClass`, `aeLoggerClass`, etc.).

---

## § 10 — Implementation Session Summary (2026-04-09)

### Completed

**File renames (git mv, committed as d4e9278):**
- `src/basLogger.bas` → `src/aeLoggerClass.cls`
- `src/XLongRunningProcessCode.bas` → `src/basLongProcess.bas`

**Content updates to `basLongProcess.bas` (unstaged):**
- `Attribute VB_Name` updated to `"basLongProcess"`
- Module header rewritten: X-prefix convention description replaced with skeleton
  architecture note (bas file is thin public skeleton; logic lives in class layer)
- `StartOrResumeUpdate` → `StartOrResume`
- `StopUpdate` → `StopTask`
- `ResetProgress` → `ResetTask`
- All MsgBox error strings updated to reference `Module basLongProcess`
- Corrected misleading comment on `PauseWithDoEvents` call (was labelled
  `1000 milliseconds = 1 second`; corrected to `60000 milliseconds = 60 seconds`)

**Content updates to `aeLoggerClass.cls` (unstaged):**
- Full VBA class module header added (`VERSION 1.0 CLASS` ... `Attribute VB_Exposed`)
- `Attribute VB_Name` updated to `"aeLoggerClass"`
- `Option Private Module` removed
- Explanatory comment added: class modules never use `Option Private Module`; the
  class is accessible from any module including those that also omit it
- Header comment updated: `LOGGING MODULE` → `LOGGING CLASS`; usage example updated
  to `Dim log As New aeLoggerClass`

**Normalizer update (`py/normalize_vba.py`):**
- Rule added: `.Name` property (fixes `.name` → `.Name` for `VBProject.Name`,
  `Style.Name`, `Document.Name`, and other `.Name` properties)

### Next steps (plan order)

1. Commit the unstaged content changes to `basLongProcess.bas` and `aeLoggerClass.cls` — DONE
2. Step 1 of plan: create `src/IaeLongProcessClass.cls` (interface definition) — DONE
3. Step 2 of plan: redesign `aeLoggerClass` for sequential writes; wire
   `Run_All_SBL_Tests` to write output to `rpt/SBL_Tests.UTF8.txt`
4. Step 3 of plan: refactor `basLongProcess.bas` skeleton entry points to accept
   `IaeLongProcessClass`; move batch loop logic to class layer
5. Step 4 of plan: replace `CustomDocumentProperties` progress storage with
   `rpt/` sidecar file

---

## § 12 — Step 1 Complete: IaeLongProcessClass.cls Created (2026-04-09)

`src/IaeLongProcessClass.cls` created and imported successfully as a class module.

**Contract defined:**
- `ExecuteItem(itemIndex As Long) As Boolean` — process one item; return True to
  continue, False to stop
- `ItemCount() As Long` — total items (used for progress percentage)
- `TaskName() As String` — identifier for log output and sidecar filename

**Reminder:** the `Write` tool produces LF line endings. All new `.cls` and `.bas`
files must be converted to CRLF with the Python one-liner from § 11 before importing
into the VBA IDE.

---

## § 13 — Step 2 Implementation: aeLoggerClass Sequential Writes + SBL Test Wiring (2026-04-09)

### Changes made

**`src/aeLoggerClass.cls` — sequential write redesign:**
- Added `Private m_stm As Object` to hold the ADODB.Stream open between `Log_Init`
  and `Log_Close`
- `Log_Init`: opens the stream once, loads existing file if present and seeks to end
  (append mode), otherwise writes BOM; marks `logIsOpen = True` before writing header
- `Log_WriteRaw`: writes to the held-open stream and calls `SaveToFile` after each
  write (crash-safe: file is always current)
- `Log_Close`: writes session-end line, closes stream, sets `m_stm = Nothing`
- Added `Class_Terminate`: calls `Log_Close` if still open (guards against object
  going out of scope without explicit close)
- `TestLogging` moved to bottom of file as a self-test utility

**`src/aeAssertClass.cls` — logger mirroring:**
- Added `Private m_log As aeLoggerClass`
- Added `Public Sub SetLogger(log As aeLoggerClass)` — caller passes the logger
  instance before `Initialize`; optional, backward compatible
- `Initialize`: mirrors the test harness header to the log
- `Terminate`: mirrors the summary block (Tests Run, Failures, RESULT) to the log;
  sets `m_log = Nothing` after writing
- `AssertTrue`: refactored to build `assertLine` string first, then
  `Debug.Print assertLine` + `m_log.Log_Write assertLine`
- `AssertEqual`: same refactor as `AssertTrue`

**`src/basTEST_aeBibleCitationClass.bas` — wiring:**
- `Run_All_SBL_Tests`: instantiates `Dim log As New aeLoggerClass`, calls
  `log.Log_Init ActiveDocument.Path & "\rpt\SBL_Tests.UTF8.txt"`
- Calls `aeAssert.SetLogger log` before `aeAssert.Initialize`
- Calls `log.Log_Close` / `Set log = Nothing` after `Run_Extra_Tests`
- PROC_ERR: calls `log.Log_Close` before MsgBox to ensure log is flushed on error

### What the log captures

All `PASS`/`FAIL` lines from every assertion, the test harness header, and the
final summary (Tests Run, Failures, RESULT). Stage header lines (e.g.
`------ Test_Stage5 ------`) remain `Debug.Print` only — the individual test
methods are unchanged.

### Next step

Run `Run_All_SBL_Tests` and verify `rpt/SBL_Tests.UTF8.txt` is created containing
all assertion results and the PASS/FAIL summary.

---

## § 14 — Logger Bug: ADODB.Stream Text-Mode Position Seeking Unreliable (2026-04-09)

### Symptom

First test of `aeLoggerClass` produced a corrupted log file:

```
﻿18:26:40 | LOG SESSION END
-------------------------
e 25)
==========
=
```

Only `LOG SESSION END` appeared intact. The session header (LOG SESSION START, User,
Machine) was absent. Earlier content was truncated to partial line fragments.

### Root Cause

ADODB.Stream in text mode does not support reliable position-based append. The
`SaveToFile` method saves from the current stream `Position` (in bytes), not from
position 0. After each `WriteText` call, `Position` advances to the end of the newly
written content. Each subsequent `SaveToFile` therefore writes only a small slice —
the content written in that single call — overwriting the file rather than saving
the full accumulated content.

The session header lines were each overwritten by the next `SaveToFile` call.
Only `LOG SESSION END` survived because it was the last write before `Log_Close`
called `m_stm.Close` without another `SaveToFile`.

The `LoadFromFile` + `Position = .Size` append strategy is also unreliable in
UTF-8 text mode because `Position` is byte-based while the stream internally tracks
a text offset; the BOM (3 bytes) and multi-byte characters create misalignment.

### Fix

Replaced the held-open stream approach with a **string buffer**:

- `m_buffer As String` accumulates all content including the BOM
- `Log_Init` initialises `m_buffer = ChrW(&HFEFF)` (BOM) then writes the header
- `Log_WriteRaw` appends to `m_buffer`, then opens a fresh ADODB.Stream, writes
  the full buffer, saves, and closes
- `Log_Close` writes the session-end line and clears the buffer

This avoids all stream-position issues. The file is always rewritten from the full
buffer, so it is complete and correct after every write (crash-safe). The ADODB.Stream
is opened and closed per write, which was the original pattern — the difference is
that no `LoadFromFile` is needed and the write is always the complete content.

### Note on "held-open stream" goal

The original plan called for holding the stream open for performance. In practice,
ADODB.Stream text mode does not support reliable incremental append via position
seeking. The buffer approach achieves the same performance benefit (no `LoadFromFile`
on each write) without the positioning hazard. The stream object is lightweight to
open/close; the cost is negligible compared to `LoadFromFile` on a large file.

---

## § 15 — Steps 3 and 4 Complete: aeLongProcessClass + basLongProcess Refactor (2026-04-09)

### New file: `src/aeLongProcessClass.cls`

The runner class that executes any `IaeLongProcessClass` task. Contains all logic
previously in the `basLongProcess.bas` skeleton:

| Member | Description |
|--------|-------------|
| `BatchSize As Long` | Items per batch (default 50) |
| `PauseMs As Long` | Pause between batches in ms (default 1000) |
| `Run(task As IaeLongProcessClass)` | Main batch loop with DoEvents, stop/resume, logging |
| `StopTask()` | Sets `m_continueProcessing = False`; loop exits at next item |
| `ResetTask(task)` | Clears progress; next `Run` starts from item 1 |
| `SaveProgress(taskName)` | Writes `LastProcessedItem` + `ProgressPercentage` to `rpt/LongRunningProgress_{TaskName}.txt` |
| `LoadProgress(taskName)` | Reads sidecar file on `Run` start; sets `m_lastProcessedItem` |
| `PauseWithDoEvents(ms)` | Sleep loop with DoEvents between batches |

**Progress storage (Step 4):** `CustomDocumentProperties` replaced by a plain-text
sidecar file at `rpt/LongRunningProgress_{TaskName}.txt` (key=value format). The
document is never marked as modified by progress saves.

**Logging:** `aeLongProcessClass` instantiates `aeLoggerClass` at `Run` start and
writes to `rpt/LongProcess_{TaskName}.txt`. If `ActiveDocument.Path` is empty the
logger is skipped gracefully.

**Task stop signals:** two paths:
- `StopTask()` — external stop (user or ribbon); sets flag, loop exits at next item
- `task.ExecuteItem(i)` returns `False` — task requests its own stop

Both paths save progress and close the log before exiting.

### Rewritten: `src/basLongProcess.bas`

Reduced to a thin public skeleton. Removed: `LongProcessSkeletonWithConsoleProgress`,
`SaveProgress`, `LoadProgress`, `CustomPropertyExists`, `PauseWithDoEvents`, the
`Sleep` declare, and all module-level state variables. These all live in
`aeLongProcessClass` now.

**Kept:**
- `Private s_runner As aeLongProcessClass` — module-level runner instance
- `StartOrResume(task As IaeLongProcessClass)` — creates runner if needed, calls `Run`
- `StopTask()` — delegates to `s_runner.StopTask`
- `ResetTask(task As IaeLongProcessClass)` — delegates to `s_runner.ResetTask`
- `SetWordHighPriority()` — opt-in WMI utility, unchanged
- `UpdateCharacterStyle()` — legacy stub, marked for extraction to `aeUpdateCharStyleClass` in Step 5

### Next step

Step 5: create `src/aeUpdateCharStyleClass.cls` implementing `IaeLongProcessClass`.
Move task logic from `UpdateCharacterStyle` in `basLongProcess.bas` into it.

---

## § 16 — Step 5 Complete: aeUpdateCharStyleClass (2026-04-09)

### What aeUpdateCharStyleClass does

`UpdateCharacterStyle` in `basLongProcess.bas` iterates over every character in the
document looking for characters that carry the `"Chapter Verse marker"` style, then
re-applies that same style to each one. This forces Word to rebuild its internal style
data for those characters — it is a **style refresh/repair operation**, not a content
change.

As `aeUpdateCharStyleClass` implementing `IaeLongProcessClass`:

| Member | Value |
|--------|-------|
| `TaskName` | `"UpdateCharacterStyle"` |
| `ItemCount` | `ActiveDocument.Paragraphs.Count` |
| `ExecuteItem(i)` | Process paragraph `i`: re-apply `StyleName` to matching characters; return `True` to continue, `False` to stop |
| `StyleName` | Public property, default `"Chapter Verse marker"`; can be set before `Run` |

### Design decisions

**One paragraph per ExecuteItem call.** The runner's batch loop handles grouping,
DoEvents, pausing, and progress persistence. The task class has no awareness of
batching.

**pageNumber skip logic removed.** The original `UpdateCharacterStyle` accepted a
`pageNumber` parameter to skip earlier pages. This is replaced by the sidecar-file
resume mechanism: on restart, `aeLongProcessClass` resumes from `LastProcessedItem`
(the paragraph index), so no manual page offset is needed.

**5,000-update hard stop removed.** The original stopped after 5,000 updates as a
safety limit. `StopTask()` on the runner replaces this with a user-controlled stop
at any item boundary.

**Error handling in ExecuteItem:** returns `False` on error (signals runner to stop)
and shows a MsgBox identifying the failing paragraph index.

### `basLongProcess.bas` UpdateCharacterStyle stub

The legacy `UpdateCharacterStyle` sub remains in `basLongProcess.bas` marked with a
comment noting it is superseded by `aeUpdateCharStyleClass`. It can be removed once
the new implementation is validated.

### Usage

```vba
' Immediate Window
Dim t As New aeUpdateCharStyleClass
Dim r As New aeLongProcessClass
StartOrResume t          ' via basLongProcess entry point

' Or directly:
r.Run t

' Custom style:
t.StyleName = "My Character Style"
r.Run t
```

---

## § 17 — Step 9 Complete: normalize_vba.py Updated for Framework Identifiers (2026-04-09)

### Problem fixed

The VBA IDE normalizes identifier casing based on the first `Dim` declaration it
encounters in any loaded module. The local variable `Dim styleName As String` in
`UpdateCharacterStyle` caused the IDE to lowercase the public property `StyleName`
on `aeUpdateCharStyleClass` to `styleName` on export.

### Fix

Added `StyleName` to the normalizer and extended it with all new public identifiers
introduced by the long-process framework this session:

| Rule | Canonical casing | Source |
|------|-----------------|--------|
| `StyleName` | `StyleName` | `aeUpdateCharStyleClass` public property |
| `BatchSize` | `BatchSize` | `aeLongProcessClass` public property |
| `PauseMs` | `PauseMs` | `aeLongProcessClass` public property |
| `TaskName` | `TaskName` | `IaeLongProcessClass` interface property |
| `ItemCount` | `ItemCount` | `IaeLongProcessClass` interface property |
| `ExecuteItem` | `ExecuteItem` | `IaeLongProcessClass` interface method |
| `StartOrResume` | `StartOrResume` | `basLongProcess` entry point |
| `StopTask` | `StopTask` | `basLongProcess` / `aeLongProcessClass` |
| `ResetTask` | `ResetTask` | `basLongProcess` / `aeLongProcessClass` |
| `Log_Init` | `Log_Init` | `aeLoggerClass` method |
| `Log_Write` | `Log_Write` | `aeLoggerClass` method |
| `Log_Close` | `Log_Close` | `aeLoggerClass` method |
| `Log_UnicodeDetail` | `Log_UnicodeDetail` | `aeLoggerClass` method |
| `SetLogger` | `SetLogger` | `aeAssertClass` method |

Rules are grouped with a date comment in `NORMALIZATIONS` for traceability.

### Next step

Step 8: integration testing — import `aeLongProcessClass.cls` and
`aeUpdateCharStyleClass.cls`, run `StartOrResume` with a test task, confirm
progress sidecar file, stop/resume, and log output.

---

## § 11 — Line Ending Fix for VBA Class Import (2026-04-09)

### Problem

After `aeLoggerClass.cls` was written and imported into the VBA IDE, it appeared as
a standard module rather than a class module. The `VERSION 1.0 CLASS ... END` header
was present and correctly cased, but the IDE did not recognise it.

### Root Cause

The `Write` tool produces LF-only line endings (`\n`). All existing VBA source files
in this project use CRLF (`\r\n`), as required by the Windows VBA IDE. When the
class file header has LF-only line endings, the IDE's header parser fails silently
and treats the file as a plain module.

Confirmed by comparing with `aeRibbonClass.cls`:
- `aeRibbonClass.cls`: `VERSION 1.0 CLASS^M$` (CRLF — correct)
- `aeLoggerClass.cls` as written: `VERSION 1.0 CLASS$` (LF only — broken)

### Fix

Both affected files were converted to CRLF using Python:

```python
content = content.replace('\r\n', '\n').replace('\n', '\r\n')
with open(path, 'wb') as f:
    f.write(content.encode('utf-8'))
```

`basLongProcess.bas` was also converted at the same time, and an em dash (`—`) in
the module header comment was replaced with ` - ` to keep the file plain ASCII.

### Rule for Future Work

**Any `.cls` or `.bas` file created with the `Write` tool must be converted to CRLF
before importing into the VBA IDE.** This applies to all new files:
`IaeLongProcessClass.cls`, concrete task classes, and any other new modules.

The conversion command to run after every `Write` to a VBA source file:

```python
python3 -c "
p = 'C:/adaept/aeBibleClass/src/FILENAME.cls'
with open(p, 'r', encoding='utf-8', newline='') as f:
    content = f.read()
content = content.replace('\r\n', '\n').replace('\n', '\r\n')
with open(p, 'wb') as f:
    f.write(content.encode('utf-8'))
"
```
