# How `BUILD_RECORD.txt` relates to the current release process

(Added 2026-09-13, alongside `C:\adaept\aeBibleClass\aeRibbon\releases\1.0.0+bc71416\BUILD_RECORD.txt`
-- a succinct companion to the fuller history in
`C:\adaept\adaept5tudio\docs\aeBibleClass-word-addin-conversion-plan.md`.)

This one file sits at the intersection of two separate release tracks.

## 1. It's the direct source evidence for the 5 VBA items still open

Tracked in `C:\adaept\adaept5tudio\docs\aeBibleClass-word-addin-conversion-plan.md` §10.6:

| §10.6 item | This file's finding (§G8) |
|---|---|
| 1 -- About dialog shows dev stub, not `RIBBON_VERSION` | "About (Alt,Y2,A)... FAIL -- shows a placeholder/debug message... 'Hello, adaept World!...'" |
| 2 -- last-verse boundary shows nothing | "Last-verse boundary, Revelation 22:21... FAIL -- no last-verse message shown at all" |
| 3 -- first-verse KeyTip conflict | "First-verse boundary... FAIL -- ...Word's Help ribbon tab was activated" |
| 4 -- Next-Chapter-after-Next-Book | "Next Chapter... FLAGGED -- landed on Chapter 2... Root cause not yet diagnosed" |
| 5 -- re-test First-book boundary | "NOT EXPLICITLY RE-TESTED this run" |

Fixing those 5 items means re-running G8 and updating this exact file (or its next-version
successor) with clean results.

## 2. Its G8 section is the same run credited in the conversion plan's §10.5 step 6

("Guard: VBA Gate G8"), from the 2026-09-12 session that produced `Radiant-Word-Bible.docx` and
let `aeBibleAddin v0.1.0` get tagged. The file says so explicitly: *"NOT a full aeRibbon v1.0.0
sign-off run (G1-G6 not re-run here)... in service of the aeBibleAddin docm/docx/JS-taskpane
release-alignment procedure."* One file, double duty -- a VBA-track artifact that also served as
the guard-check for the JS-track release that actually shipped.

## 3. Everything else in the file (G1-G7) is aeRibbon's own, separate, still-unfinished v1.0.0 ship-gate sequence

That track is nowhere near done: G1's one recorded run has 17 failures (including a new
regression); G2-G4 and G6-G7 are still blank placeholders (`____`), never executed; the
**Sign-off checklist at the bottom is entirely unchecked** -- no tag, no `RELEASES.md` row. This
is the pre-existing, explicitly-decoupled-from-the-JS-work thread (per the 2026-09-11 decision
recorded in the conversion plan's §5 Phase 1 update: the VBA state is treated as a frozen porting
baseline regardless of its own outstanding G1 failures).

## Bottom line

This file is not part of the release process that shipped `aeBibleAddin v0.1.0` and its Firebase
Hosting deploy, except for its G8 subsection, which fed into it once. The rest of it belongs to
`aeRibbon`'s own v1.0.0 sign-off -- a separate, currently-stalled track, tracked in
`C:\adaept\aeBibleClass\aeRibbon\RELEASES.md` (still empty) and
`C:\adaept\aeBibleClass\aeRibbon\QA_CHECKLIST.md`, not something the aeBibleAddin work closed out.
