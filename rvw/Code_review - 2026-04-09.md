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
