# aeRibbon — QA Release Checklist

> Operator: tick each gate as it passes. Record results in
> `aeRibbon/releases/<version>/BUILD_RECORD.txt`.

Version under test: ________________________________
Dev SHA:           ________________________________
Operator:          ________________________________
Date:              ________________________________

## G1 — Unit tests (dev `.docm`)

Run from VBA editor in the **development** `.docm`, not the template:

- [ ] `basTEST_aeBibleClass` — all `RUN_*` entry points: 0 failures
- [ ] `basTEST_aeBibleCitationClass` — all `RUN_*`: 0 failures

## G2 — Citation block

- [ ] `basTEST_aeBibleCitationBlock` — all entry points: 0 failures

## G3 — Config / styles

- [ ] `basTEST_aeBibleConfig.RUN_TAXONOMY_STYLES` — all checks green

## G4 — Tools

- [ ] `basTEST_aeBibleTools` — all `RUN_*`: 0 failures
- [ ] `basTEST_aeBibleFonts` — all `RUN_*`: 0 failures

## G5 — Export trim integrity

- [ ] `wsl python3 py/ribbon_export_trim.py` runs cleanly
- [ ] `git diff aeRibbon/src/` reviewed; every change is explainable from a
      dev-source change since the previous release
- [ ] `aeRibbon/RoutineLog.md` summary matches `aeRibbon/src/` content
- [ ] No KEPT routine is missing a body; no REMOVED routine has dangling
      callers in the kept set (spot-check on top 5 calls into citation class)

## G6 — Template build

- [ ] `aeRibbon.dotm` builds via `BUILD.md` with no manual edits beyond the
      documented steps
- [ ] `inject_ribbon.py` exits 0
- [ ] VBA editor Debug → Compile VBAProject: 0 errors
- [ ] `RIBBON_VERSION` constant matches `aeRibbon/VERSION`
- [ ] Custom property `aeRibbonVersion` matches `aeRibbon/VERSION`

## G7 — Smoke (empty host docx)

Open `aeRibbon/docx/aeRibbon-host.docx` with template attached:

- [ ] **Radiant Word Bible** tab appears
- [ ] No error dialog on document open
- [ ] Immediate window shows `RibbonOnLoad` debug line
- [ ] Only **Go** and **New Search** render greyed-out. Prev/Next Book,
      Chapter, Verse buttons and the three selectors are enabled by
      design (always-enable per issue #599 + selectors must stay
      enabled tab stops because `m_ribbon.Invalidate` from `onChange`
      is deferred). Click handlers guard the bounds. See
      `aeRibbonClass.GetNextBkEnabled` / `GetChapterEnabled` /
      `GetVerseEnabled` for the in-code rationale.

## G8 — Smoke (production Bible docx)

Open `aeRibbon/docx/Radiant-Word-Bible.docx` (the **code-free** docx
produced per `BUILD.md` "Producing the production Bible `.docx`") with
the template attached:

- [ ] **No macro-security warning appears on docx open** (proves the
      content document is truly code-free)
- [ ] Tab appears; Book selector is enabled
- [ ] Mouse path: click Book, type `Jn`, click Chapter, type `3`, click
      Verse, type `16`, click **Go** → cursor lands at John 3:16
- [ ] Keyboard path: `Alt, Y2, B`, `Jn`, Tab, `3`, Tab, `16`, Tab, Enter
      → cursor lands at John 3:16
- [ ] `Alt, Y2, ]` (Next Book) advances one book and navigates
- [ ] `Alt, Y2, .` (Next Chapter) advances within current book
- [ ] At first verse, `Alt, Y2, <` shows boundary message in status bar
- [ ] At Genesis, `Alt, Y2, [` shows first-book message
- [ ] At Revelation 22:21, `Alt, Y2, >` shows last-verse message
- [ ] **New Search** (`Alt, Y2, S`) resets to NoSelection; Chapter/Verse
      rows disable
- [ ] **About** (`Alt, Y2, A`) opens with `RIBBON_VERSION` shown

## Sign-off

- [ ] All gates passed; release artifacts written to
      `aeRibbon/releases/<version>/`
- [ ] `aeRibbon/RELEASES.md` row appended
- [ ] Commit + tag: `git tag v<version>`
