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

## § 8 — Step and Bug Status (as of 2026-04-13)

| Item | Description | Status |
|------|-------------|--------|
| Bug 12 | Tab trap at last Book/Chapter/Verse | **COMPLETE** |
| Bug 13 | Tab after Chapter steals focus to document | **COMPLETE** |
| Bug 14 | Alt+R triggers Review / Word Count | **COMPLETE — keytip="RW" removed** |
| Bug 15 | RWB tab unreachable from keyboard | **COMPLETE — Y2 confirmed** |
| Bug 16 | No keytip badges in RWB tab | **PENDING — test after re-import** |
| Bug 17 | Book selection does not scroll document | **PENDING — fix in aeRibbonClass.cls, needs re-import** |
| Alt re-entry | Alt requires Y2 when RWB tab already active | **BY DESIGN** |
| Enter vs Tab | Enter drops focus to document | **BY DESIGN — documented** |
| Step 5 | GoToVerse — timing test pending | **PENDING** |
| Step 7 | OLD_CODE cleanup | **PENDING** |
