# RWB 1.1.1 release train, VBA side (aeBibleClass) - plan of 2026-11-08

Hub plan (Node, add-in, release side): `C:\adaept\aeBibleAddin\rvw\RWB 1.1.1 Docm Cleanup and Export 2026-11-07.md`. That file holds the owner's confirm/deny answers with pros, cons and risks, the measured evidence, the decisions S-1 to S-9 and the security and DRY review plan. **This file holds the VBA detail.** The release-train table below is the shared task list and is identical in both files (RT-52 will check it).

Status legend: ✅ Done | 🟡 In progress | 🔴 Blocked | ⚪ Not started.

**Rules for this work**
- The docm is the SSOT: every fix, every test and every export change below is VBA in this repository, run by the owner in a live Word session (a USER ACTION each time: import the module, run the routine). The Node tools in `C:\adaept\aeRWB` and the release gate in `C:\adaept\aeBibleAddin` only read, check and release.
- New code goes in small `.bas` modules, not in `src\aeBibleClass.cls`, except the wiring of a test (nine steps in `md\Adding_To_Bible_Test_Class.md`). After every edit of a VBA file verify the CR count equals the LF count (an LF-only `aeBibleClass.cls` once failed to import, `C:\adaept\aeBibleAddin\rvw\Code_review 2026-10-09.md` section 29).
- Repair routines come as a pair: a **read-only list** first, a **fix** second, both run on a copy of the docm before the real one. No Immediate-window one-liners; full Subs.
- Nothing from this plan is committed yet. Commit only when the owner asks.
- The text fingerprint `c44b0feca5d3dd2ec8496ebe831c8be96aab4e60c39d467665a9f00c9fe099b8` (SHA-256 of the canonical verse stream, computed by `npm run rwb.export` in `C:\adaept\aeRWB`) must not change through any clean-up; a formatting fix that moves a letter is a bug.

## 0. The v1.1 pilot and the Aptos exception (summary; full text in the hub plan sections 0 and 3d)

- **The current v1.1 files are a PILOT TEST of the publishing process, never a final release.** The target is the real repository `C:\adaept\RadiantWordBible`, but **nothing of v1.1 is committed or pushed there until the dry run (RT-06) has passed and the owner says go (RT-07)**. The VBA side of the pilot is small: the two PDFs (print-ready and screen) of the exact v1.1 docx made in Word (RT-05); the v1.1 docm and docx are not changed.
- **Aptos is not needed.** It is in 18 verse runs of the docm (a direct theme-body-font setting; the theme body font is Aptos), 7 `DatAuthRef` runs and 1 footnote run. It is a Microsoft font under a proprietary licence tied to Microsoft 365 (the licence text read from the font file is quoted in the hub plan 3d; on this machine it is a cloud font). The PDF of v1.1 therefore embeds Aptos. `--allow-font Aptos` on `rwb.export` records that as a reviewed exception on that one PDF, **for the pilot only**; the release gate will refuse any manifest with it (RT-32). The fix is VBA: RT-13 (the 18 runs and the others) and RT-15 (the theme fonts), in this repository.

## 1. Release train (shared with the hub plan)

<!-- RELEASE-TRAIN:BEGIN -->
| ID | Task | Where | Linked tasks | Depends on | Status |
|---|---|---|---|---|---|
| **Phase 0: the v1.1 PILOT TEST in `C:\adaept\RadiantWordBible` (never a final release; nothing of v1.1 is committed or pushed there until the dry run has passed)** | | | | | |
| RT-01 | Make a PDF of the v1.1 docx from Word (measured 2026-11-08: 901 pages, 7.6 MB, 82 s) | Word (VBA export later: RT-33) | PB-9, PB-10 | - | ✅ experiment done; the rehearsal copy is made in RT-04 |
| RT-02 | Manifest and verifier learn a file that cannot be rebuilt byte for byte (PDF): `reproducible: false`, recorded SHA-256, identity checks (page count, fonts embedded, metadata), skipped by `--check` | aeRWB | RX-5, RX-10, PB-10 | - | ✅ 2026-11-08 (aeRWB, uncommitted): `--pdf` and `--allow-font` on `rwb.export`, `fonts.mjs` (the allowlist, equal to the VBA constant by test), `pdf-identity.mjs`, a PDF branch in `rwb.verify`; 208 tests. On the real Word PDF: 901 pages, 7 fonts all embedded, Aptos flagged unless excepted; trial set `C:\adaept\aeTest\rwb-export-pdf-trial` verifies, `--check` passes, `verify-rwb.ps1` OK on 72 files. Still open (PB-10): a text sample compared with the canonical text |
| RT-03 | Owner decision 2026-11-08: the real `adaept/RadiantWordBible` repository is the target of the pilot; nothing of v1.1 is committed or pushed into it until the dry run (RT-06) has passed; the pilot is marked PILOT in its tag, release title, README, manifest and ledger | owner | S-6 | - | ✅ decided 2026-11-08 |
| RT-04 | Build what the pilot needs, every outward step with `--dry-run`: edition ledger with a `status` of `pilot` or `release` (RX-24), the docx in `source/` with its SHA test (RX-17, RX-18), `rwb.release` (RX-20), the `gh` release script (RX-23), a publish step that assembles the set in a staging copy of RadiantWordBible, the `status` field in the manifest | aeRWB | RX-17, RX-18, RX-20, RX-23, RX-24 | RT-02, S-10, S-11 | ⚪ |
| RT-05 | The pilot PDFs from the exact v1.1 docx in Word (print-ready and screen, S-7), taken in with `--allow-font Aptos` (the reviewed exception, pilot only) | owner or me via Word | PB-9 | RT-04 (file names) | ⚪ |
| RT-06 | **DRY RUN, nothing committed or pushed:** assemble every format in the staging copy `C:\adaept\aeTest\pilot-stage\`; run every check (`rwb.verify --docx`, `--check`, `sha256sum -c`, `verify-rwb.ps1`, the PDF identity, the gate guard `rwbExports` which must report the exception, `rwb.release --dry-run`, `gh release create --dry-run` listing assets and SHA-256); the reader's view on a copy | both | RX-21, PB-10 | RT-04, RT-05 | ⚪ |
| RT-07 | **Owner go**, then commit and push the v1.1 PILOT into `C:\adaept\RadiantWordBible` (tag `pilot-v1.1`, README banner, binaries as assets of a GitHub pre-release titled PILOT, not committed) | owner + me | RX-26, RX-23 | RT-06 passed | ⚪ |
| RT-08 | Reader's view on the real pilot: download the assets into an empty folder, `sha256sum -c`, `verify-rwb.ps1` (also on a copy re-saved in Word), PDF checks | owner + me | RX-21, PB-10 | RT-07 | ⚪ |
| RT-09 | Lessons and fixes; then retire the pilot (delete the pilot release and tag; decide whether to keep or reset the repository history before it is made public) | both | - | RT-08 | ⚪ |
| **Phase 1: clean the docm, the SSOT (VBA, `aeBibleClass`; every fix gets a test where a test can express it)** | | | | | |
| RT-10 | Read-only font-run audit `basFontRunAudit.bas` (written, tested in Word on copies: docx 18 runs, docm 29) | aeBibleClass | R111-1 | - | ✅ 2026-11-07 (uncommitted) |
| RT-11 | **USER ACTION:** import the module into the dev docm (fresh Word), run `FontRunAudit`, send the result (expect 29) | owner | R111-2 | RT-10 | ✅ 2026-11-08: the real docm reports exactly the 29 runs of the copy (+Body 26, Cambria Math 1, Arial 2) in 208 s; report `C:\adaept\aeBibleClass\rpt\FontRunAudit.txt` (not committed: reports wait for the release) |
| RT-12 | Owner decisions S-1 to S-9 | owner | R111-3 | - | ⚪ |
| RT-13 | Fix routine for font runs: 18 verse runs, 7 `DatAuthRef`, 1 footnote, 3 `BodyText` non-Latin (list first, then fix, on a copy first) | aeBibleClass | R111-4, PB-20, F5 | RT-11, RT-12 | ⚪ |
| RT-14 | Language: English (US) for VerseText, remove 51 run and 17 paragraph-mark `de-DE` tags (list, fix, test) | aeBibleClass | F1 | RT-12 | ⚪ |
| RT-15 | Theme fonts Aptos to Carlito and Noto Sans (list, fix, test) | aeBibleClass | F2 | RT-12 | ⚪ |
| RT-16 | Complex-script and East Asian font slots of the styles in use to free fonts | aeBibleClass | F3 | RT-12 | ⚪ |
| RT-17 | Clear the 12 stray direct formats on verse runs (list, fix, test) | aeBibleClass | F4 | RT-12 | ⚪ |
| RT-18 | New VBA tests (Test 95 fonts; further numbers for RT-14, RT-15, RT-17) with the 9-step checklist of `md/Adding_To_Bible_Test_Class.md`; each must be shown to FAIL on the uncleaned copy | aeBibleClass | R111-5 | RT-13 to RT-17 | ⚪ |
| RT-19 | Port every new test to the add-in's audit engine (rule "port every Case") or register a documented gap; `py/js_port_parity.py` must show 0 NOT PORTED before the baseline re-pin | aeBibleAddin | parity ledger, plan 15.11 | RT-18 | ⚪ |
| **Phase 2: Foreword and Introduction (scripture-only publication, separate from the Study Bible)** | | | | | |
| RT-20 | Decide the route D-13 (A: a kept style in the export; B: print profile; C: outside the docx) and whether it ships in 1.1.1 or 1.1.2 | owner | PB-7, D-13, S-3 | - | ⚪ |
| RT-21 | Owner writes the Foreword and Introduction | owner | PB-7 | - | ⚪ |
| RT-22 | Styles in the docm, approved-styles SSOT, the add-in's expected-style set and profile detection (style count, sections 133) with their tests | aeBibleClass + aeBibleAddin | PB-7, W-2 | RT-20 | ⚪ |
| RT-23 | Export keeps it; Node reads it: USFM front-matter book (`FRT`/`INT`), workbook sheet "Front matter", a second fingerprint for the front matter (the verse-stream fingerprint must not change) | aeRWB | RX-3, RX-6, RX-7, PB-8 | RT-22 | ⚪ |
| RT-24 | Separation test: nothing of the Study Bible in the scripture-only front matter, nothing of the Foreword in the Study Bible export | aeBibleClass | PB-7 | RT-22 | ⚪ |
| **Phase 3: fonts embedded and the PDF (so no release relies on Microsoft fonts as the fallback)** | | | | | |
| RT-30 | Font facts: embedding permission (fsType), copyright lines and OFL text of Carlito, Noto Sans, Liberation Serif; which weights are used | me | S-5 | - | ⚪ |
| RT-31 | Export embeds the fonts (subset or full, S-5) **after** RT-15, RT-16 and RT-13: Word embeds every referenced font, so with the Microsoft names still present it embeds Calibri and Aptos too (measured) | aeBibleClass | F2, F3 | RT-13, RT-15, RT-16 | ⚪ |
| RT-32 | Guards: embedded fonts in the docx are a subset of the three; the same rule read independently in Node (`rwb.fonts`, cross-check of VBA Test 95); allowlist conformance test between the VBA constant and the Node and add-in lists (Node side done in RT-02); **the gate refuses `status: pilot` and any PDF recorded with `allowedExtraFonts`** | aeRWB + aeBibleAddin | RX-11 | RT-31 | ⚪ |
| RT-33 | PDF export from the final docx (VBA), print-ready and screen as decided (D-14), fonts embedded, recorded as non-reproducible | aeBibleClass + aeRWB | PB-9, PB-10, RT-02 | RT-31 | ⚪ |
| RT-34 | NOTICE: fonts are embedded under SIL OFL 1.1 with each font's copyright line (today's text says none is embedded) | RadiantWordBible | PB-6 | RT-30, RT-31 | ⚪ |
| RT-35 | Fonts-installed preflight on every machine that builds (VBA check of `Application.FontNames`, bootstrap step, drill PB-14) | aeBibleClass + aeRWB | PB-13, PB-14 | - | ⚪ |
| **Phase 4: export, gate, release of 1.1.1** | | | | | |
| RT-40 | Export changes: `Options.StoreRSIDOnSave = False` around the save (proven), `ExportVersion` 1.1.1, a halting font pass, optional document properties | aeBibleClass | R111-6, F6, F8 | RT-18 | ⚪ |
| RT-41 | Full VBA suite, fresh Word, AC power | owner | R111-7 | RT-40 | ⚪ |
| RT-42 | `rwb.export` (1.1.1), `rwb.verify --docx`, review in `C:\adaept\aeTest\rwb-export\`, promote | aeRWB | RX-26 | RT-41 | ⚪ |
| RT-43 | Full release gate (guards `rwbExports` and the fonts guard) | aeBibleAddin | RX-11 | RT-42 | ⚪ |
| RT-44 | Tags `rwb-v1.1.1` and `rwb-text-v1.1.1`, ledger, RELEASES rows, baseline re-pin LAST, delete the NOTICE exception | both | R111-8, RX-24 | RT-43 | ⚪ |
| **Cross-cutting (during the process)** | | | | | |
| RT-50 | Security review of both code bases (scope and first-pass notes in the plans) | both | - | before RT-43 | ⚪ |
| RT-51 | DRY review and conformance tests (book names, verse counts, USFM codes, allowlist, style sets, OOXML readers) | both | - | before RT-43 | ⚪ |
| RT-52 | Plan-sync check: this table is identical in both repositories (a script compares the marked block) | aeBibleAddin | - | - | ⚪ |
| RT-53 | After the release: the earlier lists continue (RX-13, RX-25, PB-11, PB-12, PB-14, PB-15, PB-17, PB-19, RX-8, RX-9, RX-15) | both | - | RT-44 | ⚪ |
<!-- RELEASE-TRAIN:END -->

## 2. The VBA tasks in detail

| Train ID | VBA work | Module and routines | Test | Owner action | Acceptance |
|---|---|---|---|---|---|
| RT-10 ✅ | Read-only font-run audit | `src\basFontRunAudit.bas`: `FontRunAudit`, `FontRunAuditToFile`, `CountFontRunsOutsideAllowlist`, `FindFontRunsOutsideAllowlist`, `FontRunAuditLastError` (all take an optional open-document name) | to become Test 95 (RT-18) | RT-11: import, run `FontRunAudit` in the Immediate window (about 2 minutes) | docm copy lists 29 runs (+Body 26: 18 verses and 7 `DatAuthRef` and 1 footnote; Cambria Math 1; Arial 2), docx copy 18; the target stays untouched (confirmed on the real docm 2026-11-08: the same 29 runs, 208 s) |
| RT-13 | Fix the font runs | new `src\basDocmHygiene.bas`: `ListDirectFontRuns` (reuses the audit) and `FixDirectFontRuns` clearing the direct theme-font or named-font setting on exactly the listed runs, text untouched; the 3 `BodyText` non-Latin runs are looked at before any change | Test 95 | run list, then fix on a copy, then on the docm | the audit reports 0 in scope; export fingerprint unchanged |
| RT-14 | Language | `basDocmHygiene.bas`: `ListVerseRunsNotEnglish` / `FixVerseTextLanguage` (VerseText style language English US; remove the 51 run and 17 paragraph-mark `de-DE` tags) | Test 96 `CountVerseRunsNotEnglish` expected 0 | same pattern | test 0; fingerprint unchanged |
| RT-15 | Theme fonts | `basDocmHygiene.bas`: `ListThemeFonts` / `FixThemeFonts` through `ActiveDocument.DocumentTheme.ThemeFontScheme` (minor Latin Carlito, major Latin Noto Sans) | Test 97 `CountThemeFontsOutsideAllowlist` expected 0 | same | test 0; page count and fingerprint unchanged (compare the PDF page count, 901 today) |
| RT-16 | Style font slots | `basDocmHygiene.bas`: `ListStyleSlotFonts` / `FixStyleSlotFonts` (`NameBi` and `NameFarEast` of the styles in use to Carlito or Noto Sans) | folded into Test 97 | same | no Microsoft font name left in a style in use |
| RT-17 | Stray direct formatting | `basDocmHygiene.bas`: `ListStrayVerseFormatting` / `FixStrayVerseFormatting` for the listed runs (bold Job 17:16, complex bold Exodus 38:25, size 10 Ecclesiastes 12:2, colour auto Matthew 19:1 and Luke 21:27, complex size 9 on 8 runs); targeted resets only, character styles kept | Test 98 `CountStrayDirectFormattingRuns` expected 0 | same | test 0 |
| RT-18 | Wire the tests | `src\aeBibleClass.cls`, the nine steps of `md\Adding_To_Bible_Test_Class.md`: `MaxTests` 94 to 98, expected values, `RunTest (95)` to `(98)`, `Case 95` to `Case 98` in the four Select Case blocks, the private wrapper functions, then `python3 py\js_port_parity.py` (new tests show NOT PORTED until RT-19) | Tests 95 to 98 | import the edited `.cls` after the CRLF check | each new test FAILS on the uncleaned copy and PASSES on the cleaned one; the suite is 98/98; `FINAL_TEST = 94` in the gate stays (it names the verse-structure test, the gate reads `MaxTests` itself) |
| RT-22 | Foreword styles (VBA part) | the approved-styles SSOT (`basBiblePalette`), `GetPurgeStyles` and the kept-style list in `basBibleOnlyExport.bas`, new paragraph styles for the Foreword and Introduction | Test 45 and the style-usage tests updated | author the text (RT-21) | the export keeps exactly the approved set; the add-in's expected set is updated in step with it |
| RT-24 | Foreword and Study Bible separation | a test that no study-only style or text is in the scripture-only front matter | new test number, as above | - | 0 hits |
| RT-31 | Embed the fonts at export | `basBibleOnlyExport.bas`: `ActiveDocument.EmbedTrueTypeFonts = True`, `SaveSubsetFonts = True` (S-5) on the work copy before `SaveAs2`, only after RT-13, RT-15, RT-16 | export log line | - | the docx's `word/fontTable.xml` embeds only Carlito, Noto Sans, Liberation Serif |
| RT-33 | PDF from the final docx | `basBibleOnlyExport.bas` (or a new `basBibleOnlyPdf.bas`): `ExportScripturePdf` using `SaveAs2` format 17 or `ExportAsFixedFormat` (print-ready and screen, S-7), fonts embedded by Word | export log: pages, fonts | - | PDF opens; embedded fonts are the three; page count recorded |
| RT-35 | Fonts-installed preflight | `basFontRunAudit.bas`: a cheap check of `Application.FontNames` for the three families, printed at the start of the audit and of the export | part of Test 95's hint | - | a missing font stops the export with a clear message |
| RT-40 | Export changes | `basBibleOnlyExport.bas`: (a) `Options.StoreRSIDOnSave = False` just before `SaveAs2`, the previous value restored in **every** exit path including HALT (proven on a copy: rsid attributes 236,210 to 0, runs 123,251 to 96,138, adjacent same-format runs 27,017 to 0, fingerprint identical); (b) `d("ExportVersion") = "1.1.1"` (line 1324); (c) a halting pass that calls `CountFontRunsOutsideAllowlist` on the work copy; (d) optional document properties | export log | - | 0 `w:rsid` attributes in the docx; fingerprint unchanged; HALT demonstrated once on a copy with a seeded Aptos run |
| RT-41 | Full suite | `RUN_THE_TESTS` (98 tests, about 27 minutes plus the new tests, AC power only) | all | fresh Word | 98/98 PASS, no SKIP |

The intended shape of the RT-40 (a) change, so it can be reviewed before it is written:

```vb
Dim bRsidWas As Boolean
bRsidWas = Options.StoreRSIDOnSave
Options.StoreRSIDOnSave = False          ' Word then writes no rsid attributes and merges identical neighbouring runs
On Error GoTo PROC_ERR
oDoc.SaveAs2 fileName:=destPath, FileFormat:=wdFormatXMLDocument
Options.StoreRSIDOnSave = bRsidWas
' ... and the same restore in PROC_ERR and in HaltExport, logged with both values
```

## 3. How the font audit was tested (repeatable)

A scratch macro-enabled document carried the module; the target was a **copy** opened read-only with `Application.AutomationSecurity = 3` so the docm's `Document_Open` never ran; `Application.Run "basFontRunAudit.FontRunAuditToFile", report, docName` with the scratch document active. Details and Word's behaviour (`Find.Font.Name = "+Body"`, resolved names, why `AuditFontsUsage` missed the runs) are in the hub plan appendix C. Results on 2026-11-07: docx copy 18 runs (40 s); docm copy 29 runs (123 s); docx copy with the theme body font set to Carlito in memory: 0. Reports: `C:\adaept\aeTest\vba-test\FontRunAudit-docx.txt` and `FontRunAudit-docm.txt`.

## 4. VBA-side security and DRY review scope (RT-50, RT-51; owner decision S-8 for timing)

- Security: every `Shell`, `WScript.Shell`, `Environ` and file-write call in `src\*.bas` and `aeRibbon\src\*.bas`; the export's `cmd.exe` and WSL call at `basBibleOnlyExport.bas` line 1554 (a `%` in a folder name is expanded by `cmd` even inside quotes; the output file `%TEMP%\rwb_export_step.txt` has a fixed name); the git calls in `basTEST_aeBibleClass.bas`; macro signing of the docm; what `Document_Open` and `AutoExec` do; the public `rpt\` files tracked in this repository and their licence notice (`rpt\NOTICE-RWB-TEXT.md`, `bc2abba`).
- DRY: the 66-book table appears in `aeBibleCitationClass.cls` (twice: `src` and `aeRibbon\src`), `basSBL_VerseCountsGenerator.bas` (twice), `basBibleOnlyExport.bas`, `basRWBTextExport.bas`, `basVerseStructureAudit.bas` and `basTEST_aeBibleCitationClass.bas`; decide the one source and let the others call it or be generated; the verse counts are already generated for the add-in (the model for the allowlist and the book table).

## 5. Where the other side is tracked

Node, release and add-in tasks: the hub plan above and `C:\adaept\aeBibleAddin\rvw\RWB Export Set and Text Identity 2026-11-06.md` (RX tasks) and `C:\adaept\aeBibleAddin\rvw\RWB Publication and Dev Environment 2026-11-06.md` (PB tasks). The session list is `sync\session_manifest.txt` in this repository.

Release-train block SHA-256 (text between the markers, LF line ends): `1d46954135ac4d836969abe264d02d8256b4304bc6bd253c6ce2b77ddea29023`
