# Code Review - 2026-04-13

## Carry-Forward: KeyTip Implementation + Open Items

Continues from `rvw/Code_review - 2026-04-10a.md`.

---

## § 1 — Status of Previous Session (2026-04-10a)

### Completed in previous session

| Item | Detail |
|------|--------|
| Tab trap fix — all three rows | `GetPrevBkEnabled`, `GetNextBkEnabled`, `GetPrevChapterEnabled`, `GetNextChapterEnabled`, `GetPrevVerseEnabled`, `GetNextVerseEnabled` all return `> 0` (always-enable at boundary). Tab no longer traps at Revelation or last chapter/verse. |
| Book row refactored to function-method pattern | `m_btnPrevEnabled` / `m_btnNextEnabled` flags and property getters removed. Book row now consistent with Chapter and Verse rows. |
| `basRibbonStrings.bas` created | 11 keytip constants (`KT_BOOK`, `KT_CHAPTER`, etc.). First module in the P1 string resource infrastructure. |
| `getKeytip` callbacks added | 11 stubs in `basBibleRibbonSetup.bas` referencing `basRibbonStrings.bas` constants. |
| `customUI14backupRWB.xml` updated | All controls use `getKeytip=` callbacks. Tab has static `keytip="RW"`. |
| XML injected into `Blank Bible Copy.docm` | Via `py/inject_ribbon.py`. RibbonX Editor not used. |
| `py/inject_ribbon.py` created | Scriptable replacement for RibbonX Editor save workflow. |
| § 18 added to Code_review - 2026-04-10a.md | Fluent keyboard nav design, Tab trap analysis, keytip i18n cost/benefit. |

### Still pending (carry-forward)

| Item | Detail |
|------|--------|
| Step 5 | GoToVerse — timing test (Psalm 119:176) pending |
| Step 7 | OLD_CODE cleanup — dead stubs (`ExecutePendingChapter`, `m_pendingChapter`, `GoToVerseSBL`) |
| Keytip end-to-end testing | Not yet confirmed working — blocked by Bug 15 (tab unreachable) |
| Keytip testing | Not yet confirmed working end-to-end |

---

## § 2 — Bug 14: Alt+R activates Review tab and triggers Word Count

### Symptom

Pressing **Alt → R → W** opens the Word built-in **Review** tab and launches
**Word Count** — a long-running operation on a 33,857-paragraph document.
Pressing **Escape** during Word Count causes an occasional crash.

### Root cause

The Office Fluent ribbon keytip system resolves characters **greedily**: pressing
**R** after **Alt** immediately activates the built-in **Review** tab because "R"
is its single-character keytip. The "W" keypress then fires within the Review tab
context, where it activates **Word Count**.

Our custom tab keytip `keytip="RW"` is never reached. In Office 365, custom tab
keytips specified via `customUI14` XML are overridden or ignored when any prefix
character conflicts with an existing single-character built-in keytip. "R" is
fully consumed by Review before the two-character "RW" sequence can be evaluated.

**Built-in Word tab keytips (Office 365 English — relevant conflicts):**

| Letter | Built-in tab |
|--------|-------------|
| H | Home |
| N | Insert |
| P | Page Layout |
| S | References |
| M | Mailings |
| **R** | **Review** ← conflict |
| **W** | **View** ← also conflicts as a standalone assignment |
| X | Developer |

Both "R" and "W" are taken. "RW" as a two-character sequence starting with "R"
is unreachable by design.

### Crash note

Word Count on a document of this size (33,857 paragraphs) runs a paragraph scan.
Pressing **Escape** during this scan can leave Word's internal count state
inconsistent, which occasionally causes a crash. This is a pre-existing Word
behaviour triggered by the accidental keytip collision — not a defect in the
project code.

### Fix

Remove `keytip="RW"` from the `<tab>` element in `customUI14backupRWB.xml` and
let Office auto-assign a conflict-free keytip at load time.

```xml
<!-- Before -->
<tab id="RWB" label="Radiant Word Bible" keytip="RW">

<!-- After -->
<tab id="RWB" label="Radiant Word Bible">
```

Office will assign the next available letter — typically the first letter of the
tab label not already in use. For "Radiant Word Bible", "R" is taken, so Office
will likely assign a two-character sequence or an available single letter.

**After applying the fix:**

1. Inject the updated XML via `py/inject_ribbon.py`
2. Open `Blank Bible Copy.docm`
3. Press **Alt** and observe the keytip badge that Office places on the RWB tab
4. Record the assigned keytip
5. Update the keyboard shortcut table in `md/Ribbon Design.md`

The auto-assigned keytip is stable for a given Office installation and language.
It does not change between sessions unless the set of loaded add-ins or tabs changes.

---

## § 3 — Bug 15: RWB tab not reachable by keyboard

### Symptom

After pressing **Alt**, the **Radiant Word Bible** tab does not show a keytip
badge, or the badge shown does not respond to keypresses as expected.

### Root cause

**Same root cause as Bug 14.** The `keytip="RW"` attribute is ignored or
overridden by Office because "R" is already allocated to the built-in Review tab.
Without a working keytip, the RWB tab can only be reached by mouse click or by
pressing **Alt → F10** to activate the ribbon and then using arrow keys to navigate
to the RWB tab.

### Fix

Identical to Bug 14: remove `keytip="RW"` from the tab element. The auto-assigned
keytip will make the tab directly accessible via Alt.

### Note: Alt → F10 + arrow keys (interim workaround)

Until the fix is applied and tested, the RWB tab is reachable without a mouse:

1. Press **F6** or **Alt** to activate the ribbon
2. Press **arrow keys** to navigate between tabs until RWB is selected
3. Press **Enter** or **Down** to enter the tab
4. Use **Tab** and **arrow keys** to move between controls

This is slower than a keytip but fully functional.

---

## § 4 — Bug 16: No keytip badges visible within the RWB tab

### Symptom

After navigating to the RWB tab (by mouse click), pressing **Alt** shows no letter
badges on any ribbon controls — Book, Chapter, Verse selectors and Prev/Next
buttons all appear without keytip overlays.

### Verification

`basRibbonStrings`, `GetBookKeytip`, and `KT_BOOK` are all confirmed present in
`word/vbaProject.bin` — both modules are imported. The callbacks are not the cause.

### Root cause — two-level keytip navigation

Office keytip navigation operates in **two sequential levels**:

```
Level 1 — press Alt from document:   badges appear on each ribbon tab
Level 2 — tab activated via keytip:  badges appear on controls within that tab
```

Level 2 badges only appear after the tab is activated by its Level 1 keytip.
Because the RWB tab keytip `"RW"` is broken (Bug 15 — "R" fires the Review tab
first), keyboard users are blocked at Level 1 and never reach Level 2.

**Mouse users are also affected.** Clicking the RWB tab activates it visually,
but pressing Alt re-enters Level 1 (tab badges), not Level 2 (control badges
within the already-active tab). To see Level 2 badges on an already-active tab,
focus must be on the ribbon — not the document — when Alt is pressed.

The workaround to reach control keytips today:

1. Click the RWB tab to activate it
2. Press **F6** to shift focus from the document to the ribbon
3. Press **Alt** — Level 2 control badges now appear
4. Press the badge letter to activate the control

This confirms the callbacks and XML are correct. The only broken path is the
pure-keyboard route, which is blocked by Bug 15.

### Secondary cause — `getKeytip` on `<comboBox>` unverified

`getKeytip` is used on `cmbBook`, `cmbChapter`, and `cmbVerse`. The customUI14
schema lists it as supported for `<comboBox>`, but this has not been confirmed in
this Office 365 installation. If button keytips appear after Bug 15 is fixed but
combo keytips do not, the fallback is to add static `keytip=` alongside each
`getKeytip=` attribute:

```xml
<comboBox id="cmbBook"    ... keytip="B" getKeytip="GetBookKeytip" .../>
<comboBox id="cmbChapter" ... keytip="C" getKeytip="GetChapterKeytip" .../>
<comboBox id="cmbVerse"   ... keytip="V" getKeytip="GetVerseKeytip" .../>
```

Record the observed behaviour after Bug 15 is fixed and update this section.

### Fix dependency

Fix Bug 14 / Bug 15 first. Once the tab is reachable via keyboard, this bug is
either resolved automatically (Level 2 becomes reachable) or reduced to the
`<comboBox>` fallback question above.

---

## § 5 — Bug 17: Book selection does not scroll document to the selected book

### Symptom

Workflow: **Alt, Y2, B** → type `rev` → **Enter**

The Chapter comboBox enables (correct — `m_currentBookIndex` is set) but the
document does not scroll to Revelation. The viewport stays wherever it was.

### Root cause

`OnBookChanged` deliberately skips document navigation. The Bug 9 fix removed the
`ActiveDocument.Range(foundPos, foundPos).Select` call because `.Select`
unconditionally moves focus to the document, which broke the Tab flow
(Tab after Book was intended to move to the Chapter comboBox, but `.Select` stole
focus first).

The side effect of that fix is that selecting a book provides no visual feedback
in the document — the user has no confirmation that the correct book was matched.

### Why `ScrollIntoView` is safe where `.Select` was not

`Window.ScrollIntoView(Range, Start)` scrolls the viewport to show the given range
but does **not** change the selection and does **not** move keyboard focus. Focus
remains in the ribbon. The Tab flow is unaffected.

`.Select` changes both the selection and focus — that is what triggered Bug 9.
`ScrollIntoView` changes neither.

### Fix

In `OnBookChanged`, after setting `m_currentBookPos`, call `ScrollIntoView`:

```vba
If m_currentBookPos > 0 Then
    ActiveWindow.ScrollIntoView ActiveDocument.Range(m_currentBookPos, m_currentBookPos), True
End If
```

The `m_currentBookPos > 0` guard prevents a call to `ActiveDocument.Range(0, 0)`
if `headingData` does not have a position recorded for the matched book.

### Behaviour after fix

- Typing `rev` → document scrolls to show the Revelation heading; Chapter row enables
- Typing `gen` → document scrolls to Genesis; Chapter row enables
- Typing `r` (matches Ruth first) → scrolls to Ruth; refines as more characters typed
- Tab from Book to Chapter → no focus change to document (ScrollIntoView, not Select)
- Enter from Book → same as Tab; scroll already happened on character match

---

## § 6 — Fix Sequence and Test Checklist

Apply in this order to isolate each bug cleanly:

| # | Action | Verifies |
|---|--------|----------|
| 1 | Remove `keytip="RW"` from `<tab>` in `customUI14backupRWB.xml` | Bug 14 / Bug 15 root cause |
| 2 | Run `python py/inject_ribbon.py` | XML injected into docm |
| 3 | Open `Blank Bible Copy.docm`, press **Alt** | Record auto-assigned tab keytip badge |
| 4 | Confirm Alt+[badge] focuses the RWB tab | Bug 15 resolved |
| 5 | Confirm Alt+R no longer launches Word Count | Bug 14 resolved |
| 6 | Press **F6** then **Alt** while RWB tab is active | Confirm all 11 keytip badges visible (workaround) |
| 7 | Press each keytip letter | Confirm correct control activated |
| 8 | Press each keytip letter | Confirm correct control activated |
| 9 | Record actual tab keytip assigned by Office | Update `md/Ribbon Design.md` |

---

## § 7 — Observations: Alt keytip re-entry and Enter vs Tab

### Alt when RWB tab is already active — expected Office behaviour

Observation: with the RWB tab active, pressing Alt does not show control keytip badges.
Y2 must be pressed again to reach the control badges.

This is the two-level keytip system documented in § 4. Pressing Alt from the document
always re-enters Level 1 (tab badges), regardless of which tab is currently active.
Pressing Y2 at Level 1 activates the RWB tab and simultaneously enters Level 2 (control
badges). There is no way to skip Level 1 from the document.

**Classification: expected Office behaviour — not a project bug.**

The documented keyboard path `Alt → Y2 → [control key]` is correct and intentional.

---

### Enter in a comboBox returns focus to the document — expected Office behaviour

Observation: typing `gen` then pressing Enter enables the Chapter comboBox, but
subsequent keypresses land in the document rather than the Chapter comboBox.

Office ribbon comboBox behaviour:
- **Tab** — confirms the current value, moves focus to the next ribbon control.
- **Enter** — confirms the current value, returns focus to the document.

This cannot be changed from VBA. The `onChange` callback fires identically for both
keys and has no mechanism to control where focus goes afterward.

**Classification: expected Office behaviour — not a project bug.**

The correct workflow is Tab-only navigation between selectors:

```
Book → Tab → Chapter → Tab → Verse → Tab → New Search
```

`md/Ribbon Design.md` updated with an explicit "Tab, not Enter" note.

---

## § 8 — Bug 18: Chapter Enter does not navigate; Enter in Chapter gives 2 Tabs to Verse

### Symptom

Workflow A: `gen Tab Tab 3 Enter` — Chapter comboBox enables, `m_currentChapter` is set,
but the document does not navigate to chapter 3.

Workflow B: `gen Tab Tab 3 Tab Tab 2 Tab` — navigates correctly to Genesis 3:2.

### Root cause — `OnChapterChanged` skips navigation

`GoToChapter` (line 576) calls `ActiveDocument.Range(chPos, chPos).Select`. This steals
focus from the ribbon to the document — the same defect as Bug 9 and Bug 17. The
original `OnChapterChanged` comment records the decision to skip navigation entirely:

> "scheduling deferred document navigation here causes Tab to steal focus to the
> document before the user reaches the Verse row (same root cause as Bug 9)"

So `OnChapterChanged` only updates state (`m_currentChapter`) and invalidates controls.
No scroll, no navigation.

### Why Workflow B works

`OnVerseChanged` uses `Application.OnTime Now` to defer `GoToVerse` until after the
key event clears. `GoToVerse` calls `FindChapterPos(m_currentChapter)` and navigates
to the chapter *as part of finding the verse*. Chapter navigation is embedded inside
verse navigation — `OnChapterChanged` itself never navigates.

### Why `ScrollIntoView` solves this

`Window.ScrollIntoView` scrolls the viewport without changing the selection or
moving keyboard focus — confirmed safe for ribbon callbacks (applied in Bug 17 for
`OnBookChanged`). The `.Select` concern that blocked `GoToChapter` from being called
in `OnChapterChanged` does not apply to `ScrollIntoView`.

### Fix

**Change 1 — `GoToChapter`**: replace `ActiveDocument.Range(chPos, chPos).Select`
with `ActiveWindow.ScrollIntoView ActiveDocument.Range(chPos, chPos), True`.
This also fixes the Prev/Next Chapter buttons, which call `GoToChapter` directly
and currently steal focus on click.

**Change 2 — `OnChapterChanged`**: update comment to reflect the new safe-call
reasoning; add `ScrollIntoView` call after the `InvalidateControl` block using
`FindChapterPos(chNum)`. Fires on every keystroke as the user types — consistent
with `OnBookChanged` behaviour (search-as-you-type scroll).

### Behaviour after fix

| Workflow | Before | After |
|----------|--------|-------|
| `gen Tab Tab 3 Enter` | no scroll, no nav | scrolls to Genesis 3 heading |
| `gen Tab Tab 3 Tab Tab 2 Tab` | navigates to Gen 3:2 | unchanged — GoToVerse handles final position |
| Prev/Next Chapter click | steals focus to document | scrolls to chapter, focus stays in ribbon |

---

## § 9 — Bug 19: Next/Prev Book navigates from stale cursor, not current book

### Symptom

Workflow: `rev Tab → gen Tab → Next (→ Exodus) → rev Tab → Next → EXODUS`

After re-navigating to Revelation via `rev Tab`, pressing Next does not stay at
Revelation (last book). Instead it navigates relative to wherever the document
cursor was left by the previous Next/Prev call.

### Root cause

`NextButton()` and `PrevButton()` used `Selection` (document cursor) to find their
current position:

```vba
curParaEnd = Selection.Paragraphs(1).Range.End
Selection.SetRange curParaEnd, curParaEnd
Selection.Find.style = "Heading 1"
Selection.Find.Forward = True
Selection.Find.Execute
```

`ScrollIntoView` (used in `OnBookChanged` since Bug 17 fix) scrolls the viewport
but does **not** move the document cursor. The cursor lagged behind at whatever
book the previous `Next`/`Prev` call left it. Subsequent Next/Prev navigated from
the stale cursor, not from `m_currentBookIndex`.

### Fix

Rewrote `NextButton` and `PrevButton` to use `m_currentBookIndex` and `headingData`
directly — no `Selection.Find`, no `Application.ScreenUpdating`. Navigate to
`headingData(m_currentBookIndex ± 1, 1)`, update state, call `ScrollIntoView`.

Boundary guards: `<= 1` (Genesis, no prev), `>= 66` (Revelation, no next). Both are
silent no-ops, consistent with the always-enable Prev/Next design.

---

## § 10 — Bug 20: Tab from Chapter comboBox lands in document (inline ScrollIntoView)

### Symptom

After the Bug 18 fix (which added inline `ScrollIntoView` to `OnChapterChanged`):
`rev Tab Tab Tab 3 Tab` — the Tab after typing `3` inserts a Tab character into the
document instead of moving focus to the next ribbon control.

### Root cause

The Bug 18 fix called `ActiveWindow.ScrollIntoView` synchronously inside
`OnChapterChanged`. The `onChange` callback fires while the Tab key event is still
processing. The synchronous scroll caused Word to move focus to the document window
before Tab's ribbon focus movement could complete.

### Fix (first attempt — later superseded by Bug 21 fix)

Removed the inline `ScrollIntoView` from `OnChapterChanged`. Wired up the existing
deferred infrastructure:

```vba
m_pendingChapter = chNum
Application.OnTime Now, projName & ".basRibbonDeferred.GoToChapterDeferred"
```

`GoToChapterDeferred` → `ExecutePendingChapter` → `GoToChapter` → `ScrollIntoView`
fires after the key event clears. This matched the `OnVerseChanged` pattern.

**Result:** Tab no longer stole focus mid-event. However a new symptom appeared
(Bug 21): the deferred `ScrollIntoView` fired after Tab had completed and moved
focus to the next ribbon control, then stole focus from the ribbon to the document.
See § 11.

---

## § 11 — Bug 21: Deferred GoToChapter steals ribbon focus via ScrollIntoView

### Symptom

After the Bug 20 fix (deferred chapter scroll):
`rev Tab Tab Tab 3` → document scrolls to Revelation 3 ("lands at Rev 3") → next
Tab inserts a Tab character in the document instead of moving to the Verse comboBox.

### Root cause

`ScrollIntoView` behaves differently depending on its call context:

| Context | Behaviour |
|---------|-----------|
| Called inside `onChange` (ribbon event) | Safe — ribbon retains focus |
| Called from `Application.OnTime` (Word event loop) | Steals focus to document |

In the deferred path, `Application.OnTime Now` fires after the key event clears.
At that point the ribbon has focus (user is at NextChapterButton after Tab). When
`GoToChapter` calls `ScrollIntoView`, Word moves focus to the document window to
complete the scroll. The user's next Tab then fires in the document, inserting a
Tab character.

`OnVerseChanged` uses the same deferred pattern with no focus issue because
`GoToVerseByCount`/`GoToVerseByScan` use `Selection.SetRange`/`.Select` — they
**intentionally** move focus to the document. That is the expected end of
navigation. Chapter navigation is not the end; the user needs to continue to Verse.

### Fix

`ExecutePendingChapter` is now a **no-op** (clears `m_pendingChapter`, does not call
`GoToChapter`):

```vba
Public Sub ExecutePendingChapter()
    m_pendingChapter = 0
End Sub
```

Chapter-level document scrolling occurs only through paths where focus going to the
document is appropriate:

| Path | Scrolls? | Focus after |
|------|----------|-------------|
| `OnChapterChanged` (Tab/Enter from Chapter comboBox) | No — state only | Ribbon (Tab) or document (Enter) |
| `OnPrevChapterClick` / `OnNextChapterClick` | Yes — `GoToChapter` → `ScrollIntoView` | Document (button click) |
| `GoToVerse` deferred (verse entry) | Yes — navigates to chapter + verse | Document |

### Known limitation

`chapter Enter` (chapter-only navigation without selecting a verse) does not scroll
the document. The book heading remains visible from `OnBookChanged`. To navigate to
a specific chapter without selecting a verse, use the Prev/Next Chapter buttons.

---

## § 12 — Bug 22: First navigation to a distant book is slow (~10s for Revelation)

### Symptom

`rev Tab Tab` — the comboBox shows REVELATION and then the UI freezes at the Next
button for approximately 10 seconds before responding.

### Root cause — Word page layout calculation

Word calculates page layout **lazily**: it computes only the pages that have already
been rendered on screen. The first time a scroll request reaches a page that has
not yet been laid out, Word calculates all page breaks from the last known page to
the target. For a 33,857-paragraph document, reaching Revelation from the start
requires computing all preceding pages — approximately 10 seconds on typical
hardware.

This is a Word architecture cost, not a project code defect. The delay is
proportional to distance from the last rendered page to the target:

| First navigation to | Approximate delay |
|---------------------|-------------------|
| Genesis | < 1 s (document always opens here) |
| Psalms (~midpoint) | ~5 s |
| Revelation (end) | ~10 s |

**After the first navigation the layout is cached for the session.** Subsequent
navigations to the same region or any earlier region are instant. The cost is paid
once per region per session.

### Is it avoidable? Can it be pre-warmed?

`WarmLayoutCache` (line 393 of `aeRibbonClass.cls`) already implements the warm:
it selects the last heading position (Revelation), forcing layout calculation, then
restores the saved position. It was disabled because it caused a **~50s freeze at
document open** and brought other windows to the foreground.

Why is warm-on-demand ~10s but warm-at-open ~50s?

- `ScrollIntoView` (used in `OnBookChanged`) triggers a partial layout — enough to
  scroll the viewport to the target. Word stops calculating once the target is on
  screen.
- `Range.Select` (used in `WarmLayoutCache`) triggers a full layout pass because
  Word must know the precise cursor position for caret placement, selection handles,
  and screen reader APIs. This is a deeper, slower calculation.

A targeted `ScrollIntoView`-based warm (not `Range.Select`) would be faster and
would avoid the 50s freeze.

### Can a background process start it?

VBA has no true background threading. `Application.OnTime` fires on the main UI
thread — it **blocks the UI** while running. The warm cannot proceed in background.

Options:

| Option | Cost | Benefit | Risk |
|--------|------|---------|------|
| StatusBar message during scroll | Trivial | User knows to wait | None |
| On-demand warm after first GoToVerse | One-time ~10s | Warms for session | Adds ~10s to first verse nav |
| Deferred ScrollIntoView warm at open | ~10s at open (vs 50s for Range.Select) | First nav instant | Delays document readiness |
| Accept and document | None | No disruption | User surprised by first freeze |

### Implemented mitigation

Add `Application.StatusBar` message before `ScrollIntoView` in `OnBookChanged`.
`DoEvents` is called after the status bar update to force it to render before the
freeze begins.

```vba
Application.StatusBar = "Navigating to " & CStr(headingData(m_currentBookIndex, 0)) & "..."
DoEvents   ' render status bar before ScrollIntoView blocks the UI thread
ActiveWindow.ScrollIntoView ...
Application.StatusBar = False
```

**Risk with `DoEvents`**: in an `onChange` callback, `DoEvents` processes any
pending Windows messages. The key event that triggered `onChange` has already been
consumed, so no duplicate processing. New keystrokes queued during a fast typing
burst could be dispatched. Acceptable in this context; the user is in a ~10s freeze
regardless.

### Deferred pre-warm (future option)

Replace `WarmLayoutCache`'s `Range.Select` with `ScrollIntoView`. Schedule the
warm via `Application.OnTime Now + TimeValue("00:00:10")` after document open.
Expected cost: ~10s (vs current 50s), with no foreground steal. Re-enable the
commented call in `EnableButtonsRoutine`.

---

## § 13 — Bug 22b: Document snaps back to previous verse when focus returns to document

### Symptom

Workflow: navigate to Rev 3:16 (full Tab chain) → click Book comboBox → type `gen`
(viewport scrolls to Genesis) → Shift-Tab → Enter → document snaps back to Rev 3:16.

### Root cause — Selection cursor not updated by ScrollIntoView

`GoToVerse` uses `Selection.SetRange` + `Selection.MoveDown` which **moves the
document cursor** to the verse. All subsequent navigation uses `ScrollIntoView`
which scrolls the viewport without moving the cursor. When any ribbon action
returns focus to the document, Word scrolls to show the **cursor**, not the
viewport — hence the snap-back.

The cursor diverges from the viewport whenever the user navigates by typing in the
Book comboBox (`OnBookChanged` → `ScrollIntoView`).

### Can Shift-Tab be disabled?

**No.** The Office ribbon handles `Tab` / `Shift-Tab` natively at the Win32 message
level. There is no VBA or `customUI14` API to intercept or disable them. A Win32
keyboard hook (`SetWindowsHookEx`) could intercept all keystrokes but requires a
compiled COM extension — not appropriate in this context.

The `Shift-Tab → Enter` path is an edge case: the user navigated backward to the
Prev Book button and activated it. With Genesis already selected, Prev is a no-op —
no cursor update occurs, snap-back follows.

### Forward-only navigation design

The ribbon's progressive unlock (Book → Chapter → Verse, left-to-right) is designed
for forward navigation. The Shift-Tab path bypasses this intent. Documenting the
ribbon as **forward-only** in `md/Ribbon Design.md` is an appropriate design
boundary. The Tab trap fix (always-enabled Prev/Next buttons) leaves the controls
in the Tab sequence for accessibility, but the documented workflow is Tab-forward.

### Search tracking reset — concept

After `GoToVerse` fires (navigation complete), the cursor is at the correct
location. When the user begins a **new search** (types in the Book comboBox again),
`OnBookChanged` fires. At that transition point, we know:

1. Previous search: cursor at Rev 3:16 (correct)
2. New search starting: user intends to go somewhere new

A "search complete" flag set by `GoToVerse` and cleared by `OnBookChanged` could
trigger a deferred cursor update to the new book position. But:

- The deferred cursor update would require `Selection.SetRange` or `Range.Select`
- Both steal focus when called from `Application.OnTime` context (Bug 21 pattern)
- `Selection.SetRange` from `OnTime` steals focus — not confirmed independently but
  expected (same root cause)

This approach is viable **only if** `Selection.SetRange` called from `OnTime` does
**not** steal focus. This needs a dedicated test. If confirmed safe, the search
tracking reset can be implemented cleanly.

### Pros/Cons: ScrollIntoView vs Range.Select for navigation

| Aspect | ScrollIntoView | Range.Select |
|--------|---------------|--------------|
| Moves cursor | No | Yes |
| Steals ribbon focus | No (in onChange context) | Yes |
| Steals ribbon focus (from OnTime) | Yes (Bug 21) | Yes |
| Safe in onChange | Yes | No (Bug 9) |
| Safe in onAction (button click) | Yes — but snap-back | Yes — cursor moves |
| After button click, Enter returns focus to | Document at old cursor | Document at new position |

**Conclusion**: `Range.Select` is correct for button `onAction` callbacks because
focus goes to the document regardless. `ScrollIntoView` is correct for `onChange`
callbacks because it preserves ribbon focus.

### Fix — revert button handlers to Range.Select

`NextButton`, `PrevButton`, `GoToChapter` are called only from `onAction` button
callbacks (never from `onChange`). Reverting them to `Range.Select` ensures the
cursor moves with the viewport when a button is activated. Snap-back is eliminated
for all button-driven navigation.

Snap-back from comboBox-driven book entry (`OnBookChanged`) remains a known
limitation — addressed by the forward-only navigation design boundary and future
search tracking reset work.

---

## § 14 — Step and Bug Status (as of 2026-04-14)

| Item | Description | Status |
|------|-------------|--------|
| Bug 12 | Tab trap at last Book/Chapter/Verse | **COMPLETE** |
| Bug 13 | Tab after Chapter steals focus to document | **COMPLETE** |
| Bug 14 | Alt+R triggers Review / Word Count | **COMPLETE — keytip="RW" removed** |
| Bug 15 | RWB tab unreachable from keyboard | **COMPLETE — Y2 confirmed** |
| Bug 16 | No keytip badges in RWB tab | **PENDING — test after re-import** |
| Bug 17 | Book selection scrolls document | **COMPLETE — ScrollIntoView in OnBookChanged** |
| Bug 18 | GoToChapter uses ScrollIntoView (not .Select) | **COMPLETE — Prev/Next Chapter buttons fixed** |
| Bug 19 | Next/Prev Book navigates from stale cursor | **COMPLETE — use m_currentBookIndex** |
| Bug 20 | Tab from Chapter (inline ScrollIntoView) | **COMPLETE — switched to deferred** |
| Bug 21 | Deferred GoToChapter steals ribbon focus | **COMPLETE — ExecutePendingChapter is no-op** |
| Bug 22 | First nav to distant book is slow (~10s) | **KNOWN LIMITATION — DoEvents reverted (made worse: 22s); accepted as one-time session cost** |
| Bug 22b | Snap-back to previous verse | **PARTIAL — Range.Select in button handlers; comboBox nav is known limitation** |
| Shift-Tab disable | Cannot intercept ribbon keyboard events | **BY DESIGN — forward-only nav documented** |
| Search tracking reset | Cursor update on new search start | **FUTURE — pending Selection.SetRange focus test** |
| Alt re-entry | Alt requires Y2 when RWB tab already active | **BY DESIGN** |
| Enter vs Tab | Enter drops focus to document | **BY DESIGN — documented** |
| chapter Enter | Chapter-only Enter does not scroll document | **KNOWN LIMITATION** |
| Layout pre-warm | Deferred ScrollIntoView warm at open | **FUTURE — re-enable after replacing Range.Select with ScrollIntoView in WarmLayoutCache** |
| Bug 23a | Layout delay for Psalms (~6s first nav) | **KNOWN LIMITATION — same class as Bug 22** |
| Bug 23b | Tab after multi-digit chapter → document | **FIXED — removed self-invalidation of cmbChapter in OnChapterChanged** |
| Step 5 | GoToVerse — timing test | **BLOCKED — re-test after Bug 23b fix imported** |
| Step 7 | OLD_CODE cleanup | **PENDING** |

---

## § 15 — Pre-test Review: GoToVerse Path (Step 5 — Psalm 119:176)

### Purpose

Before running the Step 5 timing test (Psalm 119:176), a full expert review of the
`GoToVerse` code path was conducted to eliminate known defects that could confound
results or produce misleading failures.

---

### Path under test

```
OnVerseChanged → m_pendingVerse → Application.OnTime → GoToVerseDeferred
  → ExecutePendingVerse → GoToVerse(vsNum)
    → FindChapterPos(m_currentChapter)   [O(N) H2 Find loop]
    → IsStudyVersion()                   [Paragraphs.Count branch selector]
    → GoToVerseByCount(chPos, vsNum)     [Selection.SetRange + MoveDown]
       OR GoToVerseByScan(chPos, vsNum)  [Range.Find loop on "Verse marker" style]
```

Also exercised via Prev/Next Verse buttons:
```
OnPrevVerseClick / OnNextVerseClick → GoToVerse(m_currentVerse ± 1)
```

---

### Issues found and fixed

#### Issue A — `GetPrevVerseEnabled` off-by-one (fixed)

**Before:** `GetPrevVerseEnabled = (m_currentVerse > 0)`
**After:**  `GetPrevVerseEnabled = (m_currentVerse > 1)`

`OnPrevVerseClick` guards on `m_currentVerse > 1`. The enabled callback used `> 0`,
so the Prev Verse button appeared active at verse 1 but clicking it was a silent
no-op. Fixed to `> 1` for consistency.

#### Issue B — `IsStudyVersion()` uncached (fixed)

**Before:** Calls `ActiveDocument.Paragraphs.Count` on every `GoToVerse` invocation.
**After:**  Result cached in `m_studyVersionSet` / `m_studyVersionVal` after first call.

On a 33,857-paragraph document, `Paragraphs.Count` forces Word to enumerate the
paragraph collection. This is non-trivial overhead on every verse navigation.
The document type does not change mid-session; one evaluation is sufficient.

Two new private fields added to class state:
```vba
Private m_studyVersionSet  As Boolean
Private m_studyVersionVal  As Boolean
```
Both initialised to `False` in `Class_Initialize`.

---

### Issues noted — not fixed

#### Issue C — `GetNextVerseEnabled` does not bound-check

`GetNextVerseEnabled = (m_currentVerse > 0 And m_currentChapter > 0 And m_currentBookIndex > 0)`

The Next Verse button stays enabled at the last verse of a chapter. `OnNextVerseClick`
correctly calls `VersesInChapter` and guards against overflow, so the click is safe.
Fixing the enabled callback would require a `VersesInChapter` call in a frequently-
fired GetEnabled callback. Deferred — acceptable UX trade-off.

#### Issue D — `FindChapterPos` is O(N) per call

`FindChapterPos(119)` iterates 119 sequential `Range.Find` passes over H2 headings
from the Psalms book position. This is the primary performance question for the
Step 5 test and is left as-is; the test result will determine whether caching is needed.

#### Issue E — `OnVerseChanged` fires on every keystroke

Typing "1", "1", "9" queues three `Application.OnTime` calls. Each fires `GoToVerse`,
producing visible intermediate scrolls to verses 1, 11, 119 before settling. This is
consistent with the `OnChapterChanged` pattern and accepted as existing behaviour.

#### Issue F — `OnPrevVerseClick` lacks error handler

Unlike all other `onAction` subs, `OnPrevVerseClick` has no `On Error GoTo PROC_ERR`.
Low risk (single guarded call to `GoToVerse` which has its own handler). Deferred.

---

### Bug 22 — DoEvents revert confirmed

`DoEvents` added before `ScrollIntoView` in `OnBookChanged` made the first navigation
to Revelation worse (22s vs 10s) and triggered a "Word not responding" spinner.

**Root cause**: DoEvents processes pending Windows messages before `ScrollIntoView`
starts. This causes Word to perform additional layout pre-calculation work before the
blocking call, adding overhead. DoEvents works in VBA-controlled loops because the
code yields between its own iterations; it cannot insert yield points inside a single
atomic Word API call.

**Decision**: Reverted. StatusBar message also removed — it cannot render before the
UI thread blocks on `ScrollIntoView`, so it would only appear after the freeze ends
(no user benefit). The one-time ~10s layout delay for first navigation to Revelation
is accepted as a known limitation. Future mitigation: `WarmLayoutCache` using
`ScrollIntoView` (not `Range.Select`).

---

### Revised task order (approved 2026-04-14)

| Priority | Task | Rationale |
|----------|------|-----------|
| 1 | **Step 5 — GoToVerse timing test** | Core functionality; determines whether FindChapterPos caching is needed |
| 2 | **Bug 16 — Keytip badges end-to-end test** | Deferred several sessions; low risk, quick to verify |
| 3 | **Step 7 — OLD_CODE cleanup** | Dead stubs (`ExecutePendingChapter`, `m_pendingChapter`, `GoToVerseSBL`); do after Step 5 confirms no regressions |
| 4 | **WarmLayoutCache rewrite** | Replace `Range.Select` with `ScrollIntoView`; re-enable deferred warm-on-open |
| 5 | **Search tracking reset** | Test `Selection.SetRange` from `OnTime` context; implement if focus-safe |

---

## § 16 — Step 5 Test Run: Bugs Found (2026-04-14)

### Test sequence attempted

```
Y2 B  →  ps  Tab  →  Tab Tab  119  →  Tab
```

Expected: Psalms selected → chapter 119 confirmed → verse field active.

---

### Bug 23a — Layout delay on first navigation to Psalms (~6s)

**Symptom**: `ps Tab` paused ~6 seconds before Psalms appeared in the Book comboBox
and the document viewport scrolled to the Psalms heading.

**Root cause**: Same as Bug 22 (Revelation ~10s). Word's page layout engine is lazy;
it calculates page positions only as far as the current viewport has rendered. On a
fresh session the document opens at page 1. Any navigation beyond the last rendered
page triggers layout computation proportional to the distance from the last rendered
point.

| Book | Distance from start | Observed delay |
|------|---------------------|----------------|
| Genesis | ~0 | <1s |
| Psalms (book 19) | ~⅓ of document | ~6s |
| Revelation (book 66) | end of document | ~10s |

**Decision**: Accepted as known limitation. Same class as Bug 22. No fix in this
session. Future mitigation remains `WarmLayoutCache` via `ScrollIntoView`.

---

### Bug 23b — Tab after multi-digit chapter → document, Tab character inserted (re-emergence of Bug 20 class)

**Symptom**: After typing `119` in the Chapter comboBox, pressing Tab caused:

- Focus to jump from the ribbon to the Word document body
- A literal Tab character to be inserted at the document cursor position (beginning
  of document — cursor had not been moved, as ScrollIntoView does not move the cursor)
- No verse field activation; verse navigation never reached

**Root cause — self-invalidation of focused control during Tab commit**:

`OnChapterChanged` calls `m_ribbon.InvalidateControl "cmbChapter"` (line 649 before
the fix). This invalidates the **currently focused control** inside its own `onChange`
callback, at the exact moment the ribbon framework is processing the Tab commit event.

The ribbon framework fires `onChange`, then expects to move Tab focus to the next
control. When `InvalidateControl "cmbChapter"` is called mid-callback, the framework
re-renders the comboBox. This disrupts the Tab event's focus-transition state. The
Tab falls through the ribbon to the document.

**Why more pronounced for "119" than for "3"**:

Each keystroke fires `onChange`, which calls `InvalidateControl "cmbChapter"`.

| Input | onChange events | Self-invalidations |
|-------|----------------|-------------------|
| "3" (1 digit) | 2 (keystroke + Tab commit) | 2 |
| "119" (3 digits) | 4 ("1" + "11" + "119" + Tab commit) | 4 |

The accumulated re-renders from 4 self-invalidations create a larger window for
the Tab commit disruption to occur. With 2 events ("3") it happens to clear; with
4 ("119") it reliably misfires.

**Why the Tab character appears at document start**:

`OnBookChanged` calls `ScrollIntoView` (scrolls viewport, does NOT move document
cursor). The cursor remains at its prior position — position 0 at the start of a
fresh session. When Tab falls through to the document, the Word editing cursor is at
position 0, and the Tab key inserts a Tab character there.

**Fix applied — remove self-invalidation**:

Removed `m_ribbon.InvalidateControl "cmbChapter"` from `OnChapterChanged`.

```vba
' REMOVED:
'   m_ribbon.InvalidateControl "cmbChapter"
' REASON: self-invalidating the focused control during its own onChange (Tab commit)
' disrupts Tab focus transition → Tab sent to document instead of next ribbon control.
' GetChapterText already returns the user-typed value; this call was redundant.
```

The remaining five `InvalidateControl` calls are retained:
- `PrevChapterButton` / `NextChapterButton` — update enabled state after chapter set
- `cmbVerse` — enables the verse row so Tab can reach it
- `PrevVerseButton` / `NextVerseButton` — update enabled state

**Expected behaviour after fix**:

`119 Tab` → Tab commits chapter 119 → focus moves to `NextChapterButton` →
second Tab → `cmbVerse` (now enabled) → type verse → Tab → `GoToVerseDeferred` fires.

---

### Status update

| Item | Status |
|------|--------|
| Bug 23a — Layout delay for Psalms (~6s) | **KNOWN LIMITATION — same class as Bug 22; no fix** |
| Bug 23b — Tab after multi-digit chapter → document | **FIXED — removed self-invalidation of cmbChapter in OnChapterChanged** |
| Step 5 timing test | **BLOCKED — pending re-test after Bug 23b fix** |
