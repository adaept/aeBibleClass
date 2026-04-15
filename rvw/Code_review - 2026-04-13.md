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
| Bug 23b | Tab after multi-digit chapter → document | **FIXED — all InvalidateControl calls moved to ExecutePendingChapter (OnTime)** |
| Bug 23c | cmbVerse disabled after chapter confirm | **FIXED — same root cause as 23b; deferred InvalidateControl ensures cache updated before Tab routing** |
| Step 5 | GoToVerse — timing test | **BLOCKED — re-test after Bug 23b/23c fix imported** |
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

---

## § 17 — Step 5 Test Run: Verse Combo Disabled After Chapter Confirm

### Symptom

After importing the Bug 23b fix (`ps Tab Tab Tab 119 Tab Tab 176 Tab`):
- Bug 23b resolved: Tab from cmbChapter no longer falls through to the document
- New symptom: cmbVerse appeared **disabled** after `119 Tab Tab` — two Tabs after
  chapter confirmation reached the verse row but cmbVerse was grayed out / inactive

---

### Root cause — ribbon Tab-routing cache not updated in time

The ribbon framework maintains an internal **enabled-state cache** for each control,
used when routing Tab focus. This cache is populated when:

1. A **full `m_ribbon.Invalidate`** is called — re-queries all controls
2. **`m_ribbon.InvalidateControl`** is called — re-queries the named control

`OnBookChanged` calls `m_ribbon.Invalidate` (full) while `m_currentChapter = 0`.
This caches `GetVerseEnabled = False` → `cmbVerse = DISABLED`.

The five `InvalidateControl` calls in `OnChapterChanged` (after the Bug 23b partial fix)
fired **synchronously during `onChange`**, at the same moment the Tab-commit event was
being processed. The ribbon had not propagated these updates to its Tab-routing cache
before Tab routing began. Tab saw `cmbVerse = DISABLED` (stale cache) and either
skipped it or left it grayed when focus arrived.

### Why this was not visible for Rev chapter "3" in a previous session

For single-digit "3", only **two** `onChange` events fire (one keystroke + Tab commit).
The timing window is narrower and, depending on Word's internal event scheduling,
the cache update may have coincided with Tab routing. For "119" (four `onChange`
events), the accumulated processing made the timing failure consistent.

---

### Fix — defer all `InvalidateControl` calls to `ExecutePendingChapter`

`OnChapterChanged` now sets state and schedules `GoToChapterDeferred` only — no
`InvalidateControl` calls.

`ExecutePendingChapter` (called via `Application.OnTime`) performs all five
`InvalidateControl` calls after the current event returns.

**Timing guarantee**: `Application.OnTime Now` fires as soon as the current VBA
procedure returns and the event queue clears. For each keystroke ("1", "11", "119"),
`ExecutePendingChapter` fires **between keystrokes**, before the next key event.
By the time the Tab-commit `onChange` fires (and Tab routing begins), `cmbVerse`
has already been reliably enabled by the previous keystroke's deferred call.

```
Keystroke "1"  → onChange → schedules OnTime → returns
                                 ↓ OnTime fires
                          ExecutePendingChapter:
                            InvalidateControl "cmbVerse"  ← cmbVerse NOW ENABLED
Keystroke "11" → onChange → schedules OnTime → returns
                                 ↓ OnTime fires (same pattern)
Keystroke "119"→ onChange → schedules OnTime → returns
                                 ↓ OnTime fires → cmbVerse confirmed ENABLED
Tab commit     → onChange → schedules OnTime → returns
               Tab routing: cmbVerse = ENABLED (from "119" OnTime) ✓
                                 ↓ OnTime fires → redundant re-enable
```

**Edge case**: user types a single digit and presses Tab immediately with no pause,
before the first `OnTime` fires. In this case the cache may still be stale. This is
an unlikely interaction pattern and is accepted as a known edge case.

---

### Summary of `OnChapterChanged` evolution

| Version | InvalidateControl location | Tab result |
|---------|---------------------------|------------|
| Original (Bug 20) | onChange — inline ScrollIntoView | Tab → document |
| After Bug 20 fix | onChange — 6 calls including self-invalidation | Tab → document for multi-digit (Bug 23b) |
| After Bug 23b fix | onChange — 5 calls, no self-invalidation | Tab → NextChapterButton; cmbVerse disabled |
| **Current** | **ExecutePendingChapter (OnTime)** | **Tab → NextChapterButton; cmbVerse enabled** |

---

### Status update

| Item | Status |
|------|--------|
| cmbVerse disabled after chapter confirm | **SUPERSEDED — see § 18** |
| Step 5 timing test | **BLOCKED — re-test after Fix 3 imported** |

---

## § 18 — Step 5 Test Run: Fix 2 Failure and Fix 3 (Final)

### Symptom after Fix 2 import

After importing the Fix 2 change (defer `InvalidateControl` calls to `ExecutePendingChapter`
via `Application.OnTime`):

- `119 Tab` still sent Tab to the document (Bug 23b re-appeared)
- cmbVerse still appeared disabled

---

### Why Fix 2 failed — Application.OnTime fires AFTER Tab routing

The § 17 analysis contained a flawed timing assumption:

> "By the time the Tab-commit `onChange` fires, `cmbVerse` has already been reliably
> enabled by the previous keystroke's deferred call."

This is incorrect. `Application.OnTime Now` does **not** fire between keystrokes while
the user is actively typing. It fires when Word is next **idle** — after the current
event queue drains, including Tab routing. The actual sequence is:

```
Keystroke "1"  → onChange → schedules OnTime → returns
Keystroke "11" → onChange → schedules OnTime → returns
Keystroke "119"→ onChange → schedules OnTime → returns
Tab commit     → onChange → schedules OnTime → returns
Tab routing    → reads enabled-state cache ← cmbVerse = DISABLED (stale)
               ↓
               Tab falls to document (focus lost from ribbon)
                    ↓ Word is now idle
               OnTime fires (too late — Tab already routed)
```

The enabled-state cache for `cmbVerse` was never updated synchronously. The stale
`DISABLED` state from `OnBookChanged`'s full `m_ribbon.Invalidate` (fired when
`m_currentChapter = 0`) was never overwritten before Tab routing read it.

Additionally, removing all synchronous `InvalidateControl` calls from `OnChapterChanged`
meant `NextChapterButton` also stayed in its stale state. Since `m_currentChapter = 0`
when `OnBookChanged` called `m_ribbon.Invalidate`, `NextChapterButton` was cached as
`DISABLED`. Tab skipped it (disabled Tab stop) and fell directly to the document.

**Key principle confirmed**: `InvalidateControl` must remain **synchronous** inside
`onChange` for any control whose enabled state is needed by the **current** Tab event's
routing. `Application.OnTime` is only safe for deferred navigation (scrolling, cursor
movement) — not for enabled-state updates consumed by the same keystroke.

---

### Fix 3 — restore synchronous InvalidateControl; extend "always-enable" to verse buttons

Three changes applied to `aeRibbonClass.cls`:

#### Change 1 — Restore 5 synchronous `InvalidateControl` calls in `OnChapterChanged`

Reverts the Fix 2 deferral. Self-invalidation of `cmbChapter` remains absent (Fix 1
from Bug 23b still applies).

```vba
    If Not m_ribbon Is Nothing Then
        ' Do NOT invalidate "cmbChapter" — self-invalidation during Tab commit → Tab
        ' to document (Bug 23b). All other controls are safe to invalidate here.
        m_ribbon.InvalidateControl "PrevChapterButton"
        m_ribbon.InvalidateControl "NextChapterButton"
        m_ribbon.InvalidateControl "cmbVerse"
        m_ribbon.InvalidateControl "PrevVerseButton"
        m_ribbon.InvalidateControl "NextVerseButton"
    End If
```

These calls fire synchronously **before** `onChange` returns, so Tab routing reads
fresh enabled states.

#### Change 2 — Revert `ExecutePendingChapter` to no-op

`ExecutePendingChapter` no longer needs to perform `InvalidateControl`. It only
clears `m_pendingChapter` (clean-up).

```vba
Public Sub ExecutePendingChapter()
    m_pendingChapter = 0
End Sub
```

#### Change 3 — Extend "always-enable at boundary" invariant to verse row buttons

The fix for Bug 23c (Tab stopping at disabled `PrevVerseButton`) applies the same
invariant already used for `PrevChapterButton` / `NextChapterButton`:

> A Prev/Next button at the boundary of an active row is **always enabled**.
> The click handler guards the actual boundary — the button's enabled state does not.

**Before Fix 3:**
```vba
Public Function GetPrevVerseEnabled(control As IRibbonControl) As Boolean
    GetPrevVerseEnabled = (m_currentVerse > 1)   ' disabled at verse 1
End Function
Public Function GetNextVerseEnabled(control As IRibbonControl) As Boolean
    GetNextVerseEnabled = (m_currentVerse > 0 And m_currentVerse < ...)
End Function
```

`PrevVerseButton` was disabled when `m_currentVerse = 0` (no verse selected yet) or
`m_currentVerse = 1`. Tab stopped at this disabled button between `NextChapterButton`
and `cmbVerse`, blocking the Tab path to the verse combo.

**After Fix 3:**
```vba
Public Function GetPrevVerseEnabled(control As IRibbonControl) As Boolean
    ' Always enabled when chapter is selected — same invariant as GetPrevChapterEnabled.
    ' PrevVerseButton is a Tab stop on the path to cmbVerse; disabling it blocks Tab flow.
    ' OnPrevVerseClick guards the actual boundary (m_currentVerse > 1).
    GetPrevVerseEnabled = (m_currentChapter > 0)
End Function

Public Function GetNextVerseEnabled(control As IRibbonControl) As Boolean
    ' Always enabled when chapter is selected (same invariant as GetPrevVerseEnabled).
    ' OnNextVerseClick guards the click boundary.
    GetNextVerseEnabled = (m_currentChapter > 0)
End Function
```

---

### Updated `OnChapterChanged` evolution table

| Version | `InvalidateControl` location | Tab result |
|---------|------------------------------|------------|
| Original (Bug 20) | onChange — inline `ScrollIntoView` | Tab → document |
| After Bug 20 fix | onChange — 6 calls including self-invalidation | Tab → document for multi-digit (Bug 23b) |
| After Bug 23b fix (Fix 1) | onChange — 5 calls, no self-invalidation | Tab → `NextChapterButton`; `cmbVerse` disabled |
| Fix 2 (failed) | `ExecutePendingChapter` (OnTime) — 5 calls deferred | Tab → document (OnTime fires too late) |
| **Fix 3 (current)** | **onChange — 5 calls, no self-invalidation** + **verse buttons always-enabled** | **Tab → `NextChapterButton` → `PrevVerseButton` → `cmbVerse`** |

---

### Status update

| Item | Status |
|------|--------|
| Bug 23a — Layout delay for Psalms (~6s) | **KNOWN LIMITATION** |
| Bug 23b — Tab after multi-digit chapter → document | **FIXED (Fix 1 — no self-invalidation retained in Fix 3)** |
| Bug 23c — PrevVerseButton blocks Tab path to cmbVerse | **FIXED (Fix 3 — always-enable at boundary)** |
| Step 5 timing test | **BLOCKED — pending import of Fix 3 + Fix 4 and re-test** |

---

## § 19 — Bug 24: First-Load Tab Falls to Document After Book Selection

### Symptom

After a fresh Word open: `ps Tab` → 4-second layout delay for Psalms → continue the
chapter/verse sequence → Tab falls to the document instead of reaching the chapter combo.
After **New Search** and repeating the same sequence, it works correctly.

### Root cause — Tab routing fires during the blocking `ScrollIntoView` call

The ribbon framework maintains a cached enabled-state for each control. This cache is
only updated when `InvalidateControl` or `Invalidate` is called.

**Initial state on fresh load**: `OnRibbonLoad` only calls `InvalidateControl` for
`NextBookButton` and `PrevBookButton`. All other controls retain their initial-render
state: `m_currentBookIndex = 0` → `GetChapterEnabled = False` → **`cmbChapter` cached
as DISABLED**.

When the user selects a book (`ps Tab`), `OnBookChanged` fires:

```
m_currentBookIndex = Psalms     ← set
m_currentChapter = 0            ← set
ScrollIntoView(..., True)       ← BLOCKS for ~4s (first-load page layout)
    │
    └── Word message pump runs during blocking call
        Tab routing fires: reads cached state
        cmbChapter = DISABLED (stale — never been re-queried)
        → Tab skips entire chapter/verse row
        → Tab falls to document at position 0
        → Tab character inserted at cursor

m_ribbon.Invalidate             ← fires after ScrollIntoView returns (TOO LATE)
```

On the second attempt (after New Search), `ScrollIntoView` is fast (layout already done
from the first attempt). `OnBookChanged` returns in milliseconds, `m_ribbon.Invalidate`
fires, and Tab routing sees the fresh cache where `cmbChapter = ENABLED`.

**Key principle**: `m_ribbon.Invalidate` must be called **before** any blocking
operation that could allow Tab routing (or any message-pump processing) to fire with
stale enabled states.

### Fix — move `m_ribbon.Invalidate` before `ScrollIntoView`

```vba
    m_currentBookIndex = i
    m_currentBookPos = CLng(headingData(i, 1))
    m_currentChapter = 0
    m_currentVerse = 0

    ' Invalidate BEFORE ScrollIntoView — Tab routing fires during the blocking layout
    ' calculation on first load (~4-10s). cmbChapter must be ENABLED in the ribbon
    ' cache before ScrollIntoView blocks, otherwise Tab skips the chapter/verse row
    ' and falls to the document (Bug 24). GetChapterEnabled = (m_currentBookIndex > 0)
    ' now returns True because m_currentBookIndex was just set above.
    If Not m_ribbon Is Nothing Then m_ribbon.Invalidate

    If m_currentBookPos > 0 Then
        ActiveWindow.ScrollIntoView ActiveDocument.Range(m_currentBookPos, m_currentBookPos), True
    End If
```

**Why no regression**: `GetBookText` returns the full book name from `headingData` at
this point, so cmbBook updates to show the resolved name — the same result as before,
just occurring before the scroll rather than after. Chapter/verse text fields clear via
`GetChapterText = ""` and `GetVerseText = ""` (m_currentChapter/Verse = 0).
Programmatic Get* updates do not re-trigger `onChange`, so no double-fire of
`OnBookChanged`.

### Status update

| Item | Status |
|------|--------|
| Bug 24 — First-load Tab to document after book selection | **SUPERSEDED — see § 24. Fix 4 (Invalidate before scroll) was correct in principle but the blocking scroll is now removed from OnBookChanged entirely.** |
| Step 5 timing test | **BLOCKED — pending import of fixes and re-test** |

---

## § 20 — Improvement: Pre-built Chapter Position Index

### Background

`FindChapterPos` locates the Nth chapter heading by calling `Range.Find.Execute` in a
loop — one Find call per chapter between the book heading and the target. To navigate to
Psalm 119, this requires **119 consecutive Find calls**. For late chapters of long books,
this is O(n) in chapter number.

`basTEST_aeBibleTools.bas` contains `LoadHeadingIndexFromCSV`, which builds an index
of H1 and H2 heading positions to `rpt\HeadingIndex.txt`. This is larger than just H1
data — it includes all chapter headings.

### Speed benefit for navigation

Replace `FindChapterPos`'s loop with a direct array lookup:

| Approach | Psalm 119 | Rev 22 | Cost |
|----------|-----------|--------|------|
| Current `FindChapterPos` | 119 `Find` calls | 22 `Find` calls | Per navigation |
| Pre-built index (H2 array) | 1 array read | 1 array read | Once at load time |

The index structure would extend the existing `headingData` array (currently H1 positions
only) to include H2 positions, keyed by book index and chapter number:

```vba
' Current: headingData(bookIdx, 0) = bookName, headingData(bookIdx, 1) = H1 charPos
' Extended: chapterData(bookIdx, chapterIdx) = H2 charPos
```

`CaptureHeading1s` already performs a full H1 scan at load time (66 entries).
Extending to H2 would capture all 1,189 chapter positions once per session.
The `LoadHeadingIndexFromCSV` / `HeadingIndex.txt` approach persists the scan result
to disk so the scan does not repeat on every document open.

### Effect on pagination delay

**None** — the pagination delay is caused by `ScrollIntoView` triggering Word's lazy
page layout engine to calculate page breaks for every paragraph between the current
scroll position and the target. This occurs regardless of how quickly the character
position is resolved. Even an instant O(1) lookup still requires the same layout work.

The pre-built index does not interact with page layout. It eliminates the `Find` loop;
it does not change what `ScrollIntoView` must do.

### Path to pagination improvement (future)

A persistent position index could enable an alternative navigation strategy that avoids
`ScrollIntoView` altogether — for example, using `Application.GoTo` with a bookmark
pre-placed at each chapter heading. Bookmarks navigate without triggering a full layout
recalculation. This approach has not been evaluated.

### Proposed task

| Task | Detail |
|------|--------|
| Extend `CaptureHeading1s` to capture H2 positions | Populate `chapterData(bookIdx, chNum)` at load time |
| Rewrite `FindChapterPos` | Direct array lookup instead of Find loop |
| Optional: persist index via `LoadHeadingIndexFromCSV` | Avoid re-scan on every document open |
| Evaluate bookmark-based navigation | Determine whether it avoids layout delay |

---

## § 21 — Bug 25a: First-load verse Tab still goes to document (Fix 5a — RETRACTED)

> **Note**: Fix 5a (double-Invalidate) was applied then retracted. The double-Invalidate
> helped the slow-tab case but failed when the user tabbed quickly (before ScrollIntoView
> returned). The root cause turned out to be the presence of `ScrollIntoView` in
> `OnBookChanged` at all, not the placement of `Invalidate` around it. **See § 24** for
> the final resolution (ScrollIntoView removed from OnBookChanged entirely).

### Symptom (original observation)

After Fix 4 (single `m_ribbon.Invalidate` before `ScrollIntoView`), the chapter Tab
path worked on first load but the verse row still failed: Tab after chapter confirm went
to the document instead of `cmbVerse`. On rapid tab entry the chapter Tab also failed.

### Root cause

`OnBookChanged` places one `m_ribbon.Invalidate` call **before** `ScrollIntoView`. This
ensures `cmbChapter` is ENABLED in the ribbon cache when Tab routing fires **during** the
blocking scroll. However, during the 4–10 second layout delay, Word's internal message
pump continues to fire ribbon `Get*` callbacks. These re-query enabled state using current
class fields. At that point `m_currentChapter = 0`, so `GetPrevVerseEnabled`,
`GetNextVerseEnabled`, and `GetVerseEnabled` all return `False`. The ribbon cache for the
verse row is **overwritten with DISABLED** during the scroll, before the user has a chance
to interact with the chapter row.

When the user later confirms a chapter and presses Tab, the verse row buttons in the cache
are still DISABLED (from the stale re-query during scroll). Tab skips them and falls to the
document.

### Fix 5a (applied 2026-04-14)

Add a second `m_ribbon.Invalidate` immediately **after** `ScrollIntoView` in
`OnBookChanged`. This wipes the stale enabled state written during the blocking call and
re-queries all controls from the current field values (still `m_currentChapter = 0` at
this point — verse row correctly DISABLED until chapter is confirmed).

```vba
' Before:
If Not m_ribbon Is Nothing Then m_ribbon.Invalidate    ' Fix 4 — before scroll
If m_currentBookPos > 0 Then
    ActiveWindow.ScrollIntoView ...
End If
' PROC_EXIT

' After (Fix 5a):
If Not m_ribbon Is Nothing Then m_ribbon.Invalidate    ' before scroll
If m_currentBookPos > 0 Then
    ActiveWindow.ScrollIntoView ...
End If
If Not m_ribbon Is Nothing Then m_ribbon.Invalidate    ' after scroll — clears stale state
```

**Double-Invalidate pattern**: `Invalidate` before a blocking call ensures Tab routing
during the call sees the correct state. `Invalidate` after the call erases any state
re-queried during the call (which used potentially stale field values). Both calls are
needed together.

### No-regression notes

The second `Invalidate` fires after `ScrollIntoView` returns (layout complete). No further
blocking operation follows, so no new stale state can accumulate before the next user
interaction. `m_currentChapter = 0` at this point is correct — verse row DISABLED until
chapter is confirmed.

| Scenario | Before Fix 5a | After Fix 5a |
|----------|--------------|-------------|
| First load, Psalm | verse row Tab fails after chapter confirm | verse row Tab works |
| Second+ load | already worked (layout instant) | unchanged |
| Prev/Next book | unaffected | unchanged |

---

## § 22 — Bug 25b: GoToVerse navigates to wrong verse in Psalm 119 (Fix 5b)

### Symptom

On second load (where verse navigation works), navigating to Psalm 119:176 scrolled to
verse 155 instead of 176. Off by exactly 21 verses.

### Root cause

`GoToVerseByCount` implemented study-version verse navigation as:

```vba
Selection.SetRange chPos, chPos
Selection.MoveDown Unit:=wdParagraph, Count:=vsNum
Selection.Collapse Direction:=wdCollapseStart
```

This moves down `vsNum` paragraph marks from the chapter heading. The assumption was
"one paragraph per verse" — true for most chapters. However, Psalm 119 has **22
Hebrew-letter section headings** (Aleph, Beth, Gimel … Taw) as separate paragraph-level
elements interspersed among the 176 verse paragraphs. Moving 176 paragraphs from H2 passes
approximately 21 of these headings before reaching verse 176, landing instead at verse 155
(176 − 21 = 155).

The error is cumulative: a chapter with N section headings before verse V will be off by
the count of headings that appear before V. Psalm 119 with 22 evenly distributed headings
produces ~21 extra paragraphs before verse 176.

### Document structure (confirmed by user)

Every verse in **both** study and print versions begins with:

1. An inline run styled **"Chapter Verse marker"** — contains the chapter number.
2. Immediately followed by an inline run styled **"Verse marker"** — contains the verse number.

Section headings, H1, H2, and all other non-verse elements do NOT begin with a
"Verse marker" run. Searching for the Nth "Verse marker" occurrence from the chapter
position therefore counts only verse starts, regardless of section headings or paragraph
structure.

### Fix 5b (applied 2026-04-14)

**Unified verse navigation path**: Remove the `IsStudyVersion` branch in `GoToVerse`.
Always call `GoToVerseByScan`. `GoToVerseByCount` is retained as a dead stub.

`GoToVerseByScan` already used `Range.Find` on the "Verse marker" character style —
correct for the print version. This approach is equally correct for the study version.

`ScrollIntoView` added after cursor placement in `GoToVerseByScan` to ensure the viewport
updates on first load (where the layout engine may not auto-scroll to a programmatic
selection).

```vba
' GoToVerse — before:
If IsStudyVersion Then
    GoToVerseByCount chPos, vsNum
Else
    GoToVerseByScan chPos, vsNum
End If

' GoToVerse — after:
GoToVerseByScan chPos, vsNum   ' correct for both versions; IsStudyVersion obsolete here
```

```vba
' GoToVerseByScan — added after Select:
ActiveDocument.Range(r.Start, r.Start).Select
ActiveWindow.ScrollIntoView ActiveDocument.Range(r.Start, r.Start), True
```

### Effect on `IsStudyVersion`

The `IsStudyVersion` function is retained — other code may rely on it. Only the verse
navigation branch in `GoToVerse` has been unified. `m_studyVersionSet` / `m_studyVersionVal`
caching remains in place. `GoToVerseByCount` is marked as a dead stub with an explanatory
comment.

| Path | Before Fix 5b | After Fix 5b |
|------|--------------|-------------|
| Study version, Psalm 119:176 | landed at verse 155 | lands at verse 176 |
| Study version, chapters without section headings | correct | unchanged |
| Print version | correct | unchanged |

---

## § 23 — Bug 25c: Spinner icon on Prev/Next verse navigation (discussion)

### Symptom

A blinking/spinning cursor icon appears when clicking Prev or Next verse. The delay is
perceptible and occurs on every verse navigation, not just the first.

### Probable root cause

`GoToVerse` calls `FindChapterPos(m_currentChapter)` on **every invocation**:

```vba
Dim chPos As Long
chPos = FindChapterPos(m_currentChapter)   ' called on every Prev/Next verse click
```

`FindChapterPos` performs a sequential `Range.Find` loop — one call per chapter from the
book heading to the target chapter. For Psalm 119, this is **119 Find calls on a
33,857-paragraph document** every time a verse navigation button is clicked. This loop is
the likely source of the visible delay.

`GoToChapter` (called by Prev/Next Chapter buttons) also calls `FindChapterPos`, but
chapter navigation is expected to be slower and is less frequent than verse navigation.

### Relationship to § 20 (chapter position index)

§ 20 proposes a full pre-built chapter position array as a longer-term improvement
(extend `CaptureHeading1s` to capture all 1,189 chapter H2 positions at load time). That
fully eliminates `FindChapterPos` for all callers. Bug 25c is the immediate, user-visible
symptom of the same underlying O(n) cost.

### Proposed fix (for discussion)

Add a private field `m_currentChapterPos As Long` that caches the chapter position
after it is first resolved for the current chapter. `GoToVerse` checks the cache before
calling `FindChapterPos`:

```vba
' In GoToVerse — replace:
chPos = FindChapterPos(m_currentChapter)

' With:
If m_currentChapterPos > 0 Then
    chPos = m_currentChapterPos
Else
    chPos = FindChapterPos(m_currentChapter)
    m_currentChapterPos = chPos
End If
```

Clear `m_currentChapterPos = 0` in:
- `OnBookChanged` (book changes → chapter resets)
- `OnChapterChanged` (new chapter entered via cmbChapter)

Set `m_currentChapterPos = chPos` also in `GoToChapter` after its `FindChapterPos` call,
so Prev/Next Chapter navigation pre-populates the cache for the immediately following verse
navigation.

**Cost model after fix**:

| Action | FindChapterPos calls |
|--------|----------------------|
| First verse in a chapter (cmbChapter path) | 1 (result cached) |
| Subsequent Prev/Next verse in same chapter | 0 (cache hit) |
| Prev/Next Chapter button | 1 (result cached in GoToChapter) |
| Book change | 0 (cache cleared; new chapter entry triggers the 1 call) |

### Open question for discussion

The cmbChapter entry path still pays one `FindChapterPos` call on the first verse
navigation. The `ExecutePendingChapter` no-op fires via OnTime after chapter confirmation
but before any verse click — this is the earliest safe point to pre-populate the cache
(paying the cost eagerly when the chapter fires, not when the first verse fires). Is eager
pre-population preferred, or is lazy (first verse click) acceptable?

Eager pre-population requires adding `FindChapterPos` work to `ExecutePendingChapter`
(currently intentionally a no-op per Fix 3). Lazy caching is simpler and still eliminates
the cost on Prev/Next verse 2 through N in the same chapter.

**Decision**: Lazy caching selected. Applied 2026-04-14. See updated cost model above.

**Status**: **APPLIED (Fix 6b)**

---

## § 24 — Bug 25a / Bug 24 Final Resolution: Remove ScrollIntoView from OnBookChanged (Fix 6a)

### Problem with previous approach (Fixes 4 and 5a)

Fix 4 (Invalidate before scroll) and Fix 5a (double-Invalidate bracketing the scroll) were
both working around the same root cause: `ScrollIntoView` in `OnBookChanged` blocked VBA
for 4-10s on first load, and Tab key presses during that block were routed using stale
ribbon state.

Fix 5a fixed the case where the user waited after selecting the book before pressing Tab
(slow-tab scenario). But on rapid tab entry — `ps Tab Tab Tab 119 Tab Tab` — multiple Tab
presses occurred during the blocking call. Some fired before the first Invalidate took
effect; others fired and re-queried stale state anyway. The double-Invalidate pattern only
cleaned up state at the endpoints of the block, not during it.

The correct diagnosis: **the blocking call itself is the problem**. No amount of
Invalidate placement eliminates stale Tab routing caused by a 4-10s block inside a ribbon
callback.

### Fix 6a (applied 2026-04-14)

Remove `ScrollIntoView` entirely from `OnBookChanged`. Replace the double-Invalidate with
a single `m_ribbon.Invalidate`. `OnBookChanged` now returns in microseconds.

```vba
' Before (Fixes 4 + 5a):
If Not m_ribbon Is Nothing Then m_ribbon.Invalidate      ' before scroll
If m_currentBookPos > 0 Then
    ActiveWindow.ScrollIntoView ...                       ' BLOCKS 4-10s on first load
End If
If Not m_ribbon Is Nothing Then m_ribbon.Invalidate      ' after scroll

' After (Fix 6a):
If Not m_ribbon Is Nothing Then m_ribbon.Invalidate      ' single call; no blocking follows
' No scroll here — view scrolls when GoToVerseByScan executes (ScrollIntoView is called
' there after cursor placement, at which point all Tab routing is complete).
```

### Why no regression

`GoToVerseByScan` (Fix 5b) already calls `ScrollIntoView` after placing the cursor at the
verse. This fires only after the user has committed their verse number — all Tab routing is
complete at that point. The 4-10s first-load pagination block moves from "during book
selection" to "during verse confirmation", where it causes no Tab routing issues.

For Prev/Next Chapter buttons: `GoToChapter` calls `Range.Select` which triggers a scroll
anyway. No regression.

For Prev/Next Book buttons: these already called `Range.Select` (no ScrollIntoView). No
change.

**UX tradeoff**: After selecting a book from `cmbBook`, the viewport no longer jumps to
the book heading immediately. The view stays at its current position until the verse is
confirmed. This is acceptable for the rapid-entry workflow (`Book Tab Chapter Tab Verse`).

### Relationship to Bug 24 / § 19 / § 21

| Fix | What it changed | Result |
|-----|-----------------|--------|
| Fix 4 (§ 19) | Invalidate moved before ScrollIntoView | Fixed slow-tab scenario |
| Fix 5a (§ 21) | Double-Invalidate bracketing ScrollIntoView | Partial improvement; fast-tab still failed |
| **Fix 6a (§ 24)** | **ScrollIntoView removed from OnBookChanged** | **Fixed for all tab speeds** |

### Status

| Item | Status |
|------|--------|
| Bug 24 / Bug 25a — First-load Tab to document after book selection | **SUPERSEDED — see § 25** |
| Chapter spinner (Prev/Next Chapter) | **FIXED (Fix 8 / § 26) — confirmed 2026-04-15** — pre-built `chapterData` array; O(1) lookup |
| Bug 25c — Verse navigation spinner | **FIXED (Fix 8 / § 26)** — `chapterData` eliminates the `FindChapterPos` call on first verse |

---

## § 25 — Bug 25a Root Cause Identified: Initial Render Cache + Deferred onChange Invalidate (Fix 7)

### Root cause

The Fluent ribbon builds an internal enabled-state cache for every control by calling all
`Get*` callbacks at **initial render** — this fires **before** `OnRibbonLoad`. At that
moment `m_currentBookIndex = 0` and `m_currentChapter = 0`, so the cache is set to:

| Control | Initial cache state |
|---------|---------------------|
| `cmbBook` | ENABLED (`GetBookEnabled = True`) |
| `cmbChapter` | **DISABLED** (`GetChapterEnabled = m_currentBookIndex > 0 = False`) |
| `cmbVerse` | **DISABLED** (`GetVerseEnabled = m_currentChapter > 0 = False`) |
| All Prev/Next buttons | DISABLED |

`OnRibbonLoad` (Fixes 4/5a/6a) only called `InvalidateControl` for the two book buttons,
then called `m_ribbon.Invalidate`. Neither established a synchronous update of `cmbChapter`
or `cmbVerse` before the user's first Tab interaction.

When the user types "gen" and presses Tab:
1. `OnBookChanged` fires → sets `m_currentBookIndex = 1` → calls `m_ribbon.Invalidate`
2. `m_ribbon.Invalidate` called from within an **`onChange`** callback is **deferred** —
   it is queued to fire after the current event cycle completes.
3. **Tab routing fires before the deferred Invalidate executes**, using the stale initial
   render cache where `cmbChapter = DISABLED`.
4. Tab skips all disabled controls and falls to the document.
5. The deferred Invalidate fires later — too late for this Tab event.

**Why it worked on second load (after New Search)**: `OnNewSearchClick` is an **`onAction`**
callback (button click). `m_ribbon.Invalidate` called from `onAction` fires synchronously,
fully updating the cache. When the user then selects a book and OnBookChanged's deferred
Invalidate queues, the ribbon framework processes it differently — likely because a prior
synchronous Invalidate cycle has already been completed. The cached state at Tab routing
time is correct on the second load.

**Why Fixes 4, 5a, 6a all failed**: They addressed `ScrollIntoView` timing (the blocking
call) without recognising that the deferred `Invalidate` from `onChange` was the fundamental
mechanism. Even with no blocking call, Tab routing on first load always fires before the
deferred Invalidate updates `cmbChapter`.

### Fix 7 (applied 2026-04-14)

**Change 1 — Make `GetChapterEnabled` and `GetVerseEnabled` unconditionally `True`.**

The controls are always tab stops from initial render onward. They are never disabled in
the cache, so Tab routing always reaches them regardless of `Invalidate` timing. The
`onChange` handlers already guard against invalid input:

```vba
' OnChapterChanged:
If m_currentBookIndex = 0 Then GoTo PROC_EXIT   ' silently ignores if no book

' OnVerseChanged:
If m_currentChapter = 0 Then GoTo PROC_EXIT     ' silently ignores if no chapter
```

**Change 2 — Add `m_ribbon.Invalidate` to `OnRibbonLoad` after `EnableButtonsRoutine`.**

This fires synchronously from the `onLoad` callback (the same mechanism as `onAction`),
establishing a correct initial cache state after class setup is complete. Removes the two
targeted `InvalidateControl` calls that only updated the book buttons.

```vba
' Before:
m_ribbon.InvalidateControl "NextBookButton"
m_ribbon.InvalidateControl "PrevBookButton"
Call EnableButtonsRoutine

' After:
Call EnableButtonsRoutine
m_ribbon.Invalidate   ' synchronous from onLoad; sets correct initial cache for all controls
```

### Impact on remaining Invalidate calls

The `m_ribbon.Invalidate` calls in `OnBookChanged` and `OnNewSearchClick`, and the
`InvalidateControl` calls in `OnChapterChanged`, remain. They update the visual state
of buttons and combo text displays (GetText, GetEnabled for visual cues). They are no
longer load-bearing for Tab routing, since the combo controls are always-enabled.

### Status

| Item | Status |
|------|--------|
| Bug 25a — First-load Tab to document after book selection | **FIXED (Fix 7)** |
| First-load Tab to document — root cause | **IDENTIFIED: onChange Invalidate is deferred; initial render cache was stale** |
| Bug 26 — Tab after chapter entry goes to document | **OPEN — Fix 9 failed; see § 28 for new approach** |

---

## § 26 — Improvement: Pre-built Chapter Position Index (Fix 8, applied 2026-04-15)

### Background

`FindChapterPos` located the Nth chapter heading by calling `Range.Find.Execute` in a
loop — one Find call per chapter between the book heading and the target. To navigate to
Psalm 119 this required **119 consecutive Find calls**. The cost was proportional to the
chapter number, making late chapters of long books noticeably slow.

### Fix 8

Extended `CaptureHeading1s` to capture H2 (chapter) positions in the same single-pass
paragraph scan that already captures H1 (book) positions. A new private array
`chapterData(1 To 66, 1 To 150) As Long` stores the character position of every chapter
heading, keyed by book index and chapter number.

**Changes to `aeRibbonClass.cls`:**

1. Added `Private chapterData(1 To 66, 1 To 150) As Long` to private state.

2. Extended `CaptureHeading1s` — same paragraph loop, now branches on `Heading 2` as
   well as `Heading 1`. When H1 is seen, `j` resets to 0. When H2 is seen (and at least
   one H1 has been captured), `j` increments and `chapterData(i - 1, j)` is set.

3. Rewrote `FindChapterPos` — replaced the Find loop with a single array read:
   `FindChapterPos = chapterData(m_currentBookIndex, chNum)`. Returns 0 for out-of-range
   inputs (same contract as the old Find path).

### Performance

| Scenario | Old path | New path |
|----------|----------|----------|
| Psalm 119 | 119 `Range.Find.Execute` calls | 1 array read |
| Revelation 22 | 22 `Range.Find.Execute` calls | 1 array read |
| Any chapter | O(n) in chapter number | O(1) |

The pagination delay from `ScrollIntoView` is unchanged — that delay is Word's layout
engine calculating page breaks, not the position lookup.

### Load-time cost

None measurable. The H2 assignments (`chapterData(i-1, j) = para.Range.Start`) add
~1,189 array writes to the existing paragraph scan. The scan already iterates every
paragraph for H1 detection; the H2 branch adds no additional passes.

### Status

| Item | Status |
|------|--------|
| Chapter spinner (Prev/Next Chapter) O(n) lag | **FIXED (Fix 8) — confirmed 2026-04-15** |
| Bug 25c — first verse click calls `FindChapterPos` once (lazy cache) | **FIXED (Fix 8)** — `chapterData` populated at load; `m_currentChapterPos` cache set on first use eliminates repeat lookups |

---

## § 27 — Bug 26: Tab After Chapter Entry Goes to Document (Fix 9, applied 2026-04-15)

### Symptom

Typing a chapter number in `cmbChapter` and pressing Tab (e.g., Psalms `119 Tab`) sent
focus to the document instead of `cmbVerse`. The bug was present for any chapter entry
and had been active for a long time without a specific tracking number.

### Root cause — Bug 23b analog

`onChange` fires for every keystroke **and** on Tab-commit. When Tab was pressed from
`cmbChapter`, `OnChapterChanged` fired one final time with the committed text. Inside
that handler, `m_ribbon.InvalidateControl "cmbVerse"` was called synchronously — while
Word's Tab routing was still in progress. Invalidating the Tab destination during Tab
routing disrupted the focus transition and sent Tab to the document.

This is the exact same mechanism as Bug 23b, but targeting the next control rather than
the current one:

| Bug | Synchronous invalidation during Tab commit | Effect |
|-----|--------------------------------------------|--------|
| Bug 23b | `InvalidateControl "cmbChapter"` (self) | Tab → document |
| **Bug 26** | `InvalidateControl "cmbVerse"` (Tab destination) | Tab → document |

Bug 23b identified self-invalidation as problematic; the destination-invalidation analog
was not recognised at the time.

### Why reported for Psalms 119 specifically

Psalms 119 (3 digits) fires `OnChapterChanged` four times: once per keystroke (`1`,
`19`, `119`) plus once for the Tab commit. Each firing called `InvalidateControl "cmbVerse"`.
The bug is present for all chapter entries but may have been more consistently observed
at 3-digit chapters because those produce more accumulated invalidation calls before Tab.

### Fix 9

**`OnChapterChanged`** — removed `InvalidateControl "cmbVerse"`, `"PrevVerseButton"`,
`"NextVerseButton"`. Added comment citing Bug 26. Only chapter-row controls
(`PrevChapterButton`, `NextChapterButton`) are invalidated synchronously.

**`ExecutePendingChapter`** — moved the three verse control invalidations here. This
fires from `GoToChapterDeferred` via `Application.OnTime Now`, which executes after the
keyboard event (and Tab routing) has fully cleared. Verse controls are therefore updated
after focus has already reached `cmbVerse`, not during the routing itself.

```vba
' Before (OnChapterChanged):
m_ribbon.InvalidateControl "PrevChapterButton"
m_ribbon.InvalidateControl "NextChapterButton"
m_ribbon.InvalidateControl "cmbVerse"        ' ← caused Bug 26
m_ribbon.InvalidateControl "PrevVerseButton" ' ← moved to ExecutePendingChapter
m_ribbon.InvalidateControl "NextVerseButton" ' ← moved to ExecutePendingChapter

' After (OnChapterChanged):
m_ribbon.InvalidateControl "PrevChapterButton"
m_ribbon.InvalidateControl "NextChapterButton"

' After (ExecutePendingChapter — fires after Tab routing completes):
m_ribbon.InvalidateControl "cmbVerse"
m_ribbon.InvalidateControl "PrevVerseButton"
m_ribbon.InvalidateControl "NextVerseButton"
```

### Status

| Item | Status |
|------|--------|
| Bug 26 — Tab after chapter entry goes to document | **FAILED (Fix 9 did not resolve) — see § 28** |

---

## § 28 — New Navigation Architecture: Default-Fill + Action-Gate (Proposed 2026-04-15)

### Background

Fixes 7, 8, and 9 all attempted to resolve Tab-routing failures by adjusting
when and which controls are invalidated during `onChange` callbacks. Fix 9
(deferring `InvalidateControl "cmbVerse"` to `ExecutePendingChapter`) failed —
focus still goes to the document after chapter entry. The recurring pattern
across Bugs 20, 23b, 25a, and 26 is the same: the ribbon framework Tab routing
interacts unpredictably with `InvalidateControl` calls made from within `onChange`.
Each fix has treated a symptom; none has removed the root dependency.

### Proposed rules

| Rule | Description |
|------|-------------|
| 1 | Navigation requires all three fields (Book, Chapter, Verse) to be filled |
| 2 | Book is always required — no default |
| 2a | When Book is confirmed, Chapter and Verse are immediately set to 1 — no wait for nav |
| 3 | Tab past Chapter accepts the displayed value (1 if set by Book, or user-entered) |
| 4 | Tab past Verse accepts the displayed value (1 if set by Book, or user-entered) |
| 5 | Navigation fires only after B/C/V are all filled (action-gated) |
| 6 | Prev/Next buttons are already guarded against out-of-range values |
| 7 | Prev/Next button presses update all three B/C/V fields appropriately |

**Rule 2a detail:** `OnBookChanged` currently sets `m_currentChapter = 0` and
`m_currentVerse = 0`. Changing both to `1` is the entire implementation. The
existing deferred `m_ribbon.Invalidate` already in `OnBookChanged` then fires
and displays "1" in both `cmbChapter` and `cmbVerse`. No new `InvalidateControl`
calls are needed. When focus arrives at `cmbChapter` after Tab, the field already
shows "1" — the user can Tab past to accept it or type a different number.

### Why this eliminates the Tab-routing bug class

The root cause of Bugs 20, 23b, 25a, and 26 is that `onChange` callbacks call
`InvalidateControl` synchronously while Tab routing is in progress. Every fix has
tried to find the safe window to call `InvalidateControl`; none has succeeded
reliably. The proposed approach removes those calls from `onChange` entirely:

- `GetChapterEnabled` and `GetVerseEnabled` remain unconditionally `True` (Fix 7).
  Tab always reaches both controls — no cache change needed.
- Navigation fires only from verse confirmation. Chapter-row invalidation is no
  longer load-bearing for navigation timing.
- The deferred patterns (`GoToChapterDeferred`, `ExecutePendingChapter`) for
  chapter navigation become unnecessary.

### Pros

**1. Eliminates the Tab-routing bug class permanently**
Removing `InvalidateControl` from `onChange` handlers removes the race condition
that caused Bugs 20, 23b, 25a, and 26. No future variant of this bug can recur.

**2. Single navigation trigger**
Navigation fires in exactly one place: verse confirmation. The current system has
navigation paths from book scroll, chapter deferred, and verse deferred. A single
trigger is easier to reason about, test, and debug.

**3. Simpler onChange handlers**
`OnChapterChanged` reduces to: validate or default to 1, set `m_currentChapter`.
No deferred scheduling, no `InvalidateControl` for downstream controls.

**4. Predictable keyboard workflow**
`Book Tab Chapter Tab Verse Enter` always works. No silent failures or
Tab-to-document edge cases. The user knows exactly what triggers navigation.

**5. Default-to-1 is natural**
"Go to Psalms" followed by Tab → Psalm 1:1. "Go to Psalms 119" followed by Tab
→ Psalm 119:1. These match the most common navigation intentions when a
sub-field is left blank.

**6. Prev/Next button consistency**
Pressing Next Chapter sets Chapter and defaults Verse to 1, keeps all three fields
in sync. B/C/V display becomes a live position indicator, not just an input form.

**7. Fewer Application.OnTime calls**
Under the current model, one deferred call is scheduled per onChange per keystroke.
With navigation gated to verse confirmation only, deferred calls drop to one per
navigation event.

### Cons

**1. Defaulting behaviour change**
After Book is confirmed, Chapter and Verse immediately show "1". The user who tabs
through without changing the defaults will navigate to Book 1:1. This is intentional
and visible — the "1" is displayed before the user reaches `cmbChapter`, so there
is no silent default. Users who want a different chapter simply type over "1".

**2. ~~Display lag for defaulted fields~~ — RESOLVED**
Previously a concern: setting `m_currentChapter = 1` inside `OnChapterChanged`
could not safely call `InvalidateControl "cmbChapter"` (Bug 23b). Under Rule 2a,
the default is set in `OnBookChanged` instead. The existing deferred
`m_ribbon.Invalidate` already in that handler picks up the new values and
displays "1" in both fields. No display lag, no new invalidation calls.

**3. Chapter-level navigation removed from Chapter entry**
Currently, entering a chapter and pressing Tab is intended to scroll the document to
that chapter (via deferred path, when working). Under the new model the document
does not move until verse is confirmed. Users expecting intermediate chapter
navigation would need to use Prev/Next Chapter buttons or accept the default path.

**4. Prev/Next Book and Chapter must update all fields**
NextBook must set Chapter = 1 and Verse = 1, update all three displays, and navigate.
This is a behaviour change: Prev/Next Book currently moves to the book heading only.
The fields must stay in sync, requiring `m_ribbon.Invalidate` from `onAction`
(synchronous and safe — no Tab routing in progress).

**5. Empty Chapter and Verse after New Search**
New Search resets `m_currentBookIndex`, `m_currentChapter`, and `m_currentVerse`
to 0 and calls `m_ribbon.Invalidate`. Chapter and Verse fields return to "" until
the user selects a new book. This is correct behaviour — Book is required, and Rule
2a populates Chapter and Verse to "1" the moment a book is confirmed. No special
handling needed for the New Search → re-entry path.

### Code complexity comparison

| Area | Current model | Proposed model |
|------|---------------|----------------|
| `OnBookChanged` | Set chapter/verse = 0, deferred Invalidate | Set chapter/verse = 1, deferred Invalidate (Rule 2a) |
| `OnChapterChanged` | Validate, 5x `InvalidateControl`, schedule deferred | Validate, accept current value if empty, no invalidation |
| `ExecutePendingChapter` | Clear flag, 3x `InvalidateControl` | Remove |
| `GoToChapterDeferred` | Dispatch to `ExecutePendingChapter` | Remove |
| `OnVerseChanged` | Validate, schedule deferred | Validate or default to 1, schedule deferred if B/C/V set |
| `GoToChapter` | Find pos, navigate, invalidate | Remove (navigation only via GoToVerse) |
| Prev/Next Chapter | Navigate to chapter | Navigate to chapter:1, update all 3 fields |
| Navigation trigger points | Book + Chapter deferred + Verse deferred | Verse confirmation only |
| Deferred OnTime calls | 1 per keystroke in chapter or verse | 1 per navigation event |

Overall complexity **decreases**. The chapter deferred path disappears. The
`InvalidateControl` calls in `onChange` disappear. Fewer moving parts, fewer
interactions between the ribbon framework and navigation logic.

### Reliability assessment

**More reliable by construction:**
- Tab routing no longer depends on `InvalidateControl` timing
- Navigation has one trigger with a well-defined precondition (B/C/V all set)
- Defaults are deterministic (always 1) rather than silent-ignore
- Prev/Next buttons keep B/C/V display in sync with document position

**Display lag risk: eliminated.** Rule 2a sets defaults in `OnBookChanged`, where
the existing deferred `m_ribbon.Invalidate` already handles display updates. When
focus arrives at `cmbChapter` after Tab, the field shows "1" — internal state and
display are always in sync.

### Recommended next steps

Implement in a new `Ribbon test15.docm` iteration:

1. **`OnBookChanged` (Rule 2a — minimal change, high impact):** change
   `m_currentChapter = 0` → `m_currentChapter = 1` and `m_currentVerse = 0` →
   `m_currentVerse = 1`. The existing deferred `m_ribbon.Invalidate` displays "1"
   in both fields. No other changes to this handler.
2. **Simplify `OnChapterChanged`:** validate input; if text is empty or invalid,
   accept the current value (already 1 from Rule 2a). No `InvalidateControl`, no
   deferred scheduling.
3. **Simplify `OnVerseChanged`:** validate or default to 1; schedule
   `GoToVerseDeferred` only when `m_currentBookIndex > 0` and `m_currentChapter > 0`.
4. **Remove `ExecutePendingChapter` and `GoToChapterDeferred`** — chapter entry no
   longer drives navigation or display updates.
5. **Update Prev/Next Book and Chapter buttons** to set all three fields and call
   `m_ribbon.Invalidate` (synchronous from `onAction`, safe).
   *Partially applied 2026-04-15: `NextButton`, `PrevButton`, `GoToChapter` all set
   `m_currentChapter = 1` / `m_currentVerse = 1` (Rule 2a). Display update relies
   on the existing `m_ribbon.Invalidate` already present in each handler.*
6. **Confirm:** `Book Tab Chapter Tab Verse Enter` navigates correctly; Chapter and
   Verse fields show "1" immediately after Book is confirmed.

### Status

| Item | Status |
|------|--------|
| Bug 26 — Tab after chapter entry goes to document | **FIXED (§ 28 architecture)** |
| Architecture — default-fill + action-gate model | **IN PROGRESS** |
| Bug 27 — Enter in Chapter does not navigate (inconsistency with Verse) | **KNOWN LIMITATION** |

**Bug 27 note:** Adding `GoToVerseDeferred` to `OnChapterChanged` caused a regression
(Tab after chapter went to document again). Root cause: `onChange` fires on every
keystroke, so deferreds for intermediate values (`1`, `19`) fire and execute
`Range.Select` + `ScrollIntoView` before Tab routing completes — the same
mechanism as Bug 20. `onChange` provides no way to distinguish Enter from a
keystroke, making per-commit-only navigation impossible without per-keystroke
side effects. Verse confirmation remains the sole navigation trigger by design.
| Step 1 — Rule 2a: `OnBookChanged` sets chapter/verse = 1 | **CONFIRMED 2026-04-15** |
| Step 2 — Simplify `OnChapterChanged`: validate only, no invalidation or deferred | **CONFIRMED 2026-04-15** |
| Bug 27 — Enter in Chapter does not navigate (inconsistency with Verse) | **KNOWN LIMITATION — see note below** |
| Bug 28 — Invalid entry in Chapter/Verse leaves stale display | **FIXED 2026-04-15 — deferred `InvalidateControl` on rejection** |
| Bug 28 — Invalid entry in Book leaves stale display | **KNOWN LIMITATION — mid-typing reset would break abbreviation entry** |
| Bug 29 — First load: Tab after Book goes to document (regression from Rule 2a Step 1) | **FIXED 2026-04-15 — display/state separation** |

### Bug 29 — First-load Tab regression (display/state separation fix)

**Root cause:** Setting `m_currentChapter = 1` in `OnBookChanged` (Rule 2a Step 1) caused
the deferred `m_ribbon.Invalidate` to enable `NextChapterButton` (because `m_currentChapter > 0`).
Tab navigates through ALL enabled controls in XML order, so Tab from `cmbChapter` hit
`NextChapterButton` instead of `cmbVerse`. On first load, `CaptureHeading1s` takes several
seconds, ensuring the deferred Invalidate fires before the user presses Tab — making the
regression consistent on first load but not on subsequent searches.

**Fix — display/state separation:**
- `OnBookChanged`: `m_currentChapter = 0`, `m_currentVerse = 0` (buttons stay disabled for Tab routing)
- `GetChapterText`: returns `"1"` when `m_currentBookIndex > 0` AND `m_currentChapter = 0` (visual default only)
- `GetVerseText`: returns `"1"` when `m_currentBookIndex > 0` AND `m_currentVerse = 0` (visual default only)
- `GoToVerse` invalidation block: adds `InvalidateControl "PrevChapterButton"` and `"NextChapterButton"` — chapter buttons only enable after first navigation fires

**Result:** The comboBoxes display "1" immediately (via `GetChapterText`/`GetVerseText`), but
`NextChapterButton`/`PrevChapterButton` remain disabled until `GoToVerse` executes. Tab from
`cmbChapter` reaches `cmbVerse` correctly on both first load and subsequent searches.
| Step 5 (partial) — Rule 2a in `NextButton`, `PrevButton`, `GoToChapter` | **APPROVED — applied 2026-04-15** |
