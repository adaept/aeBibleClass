# Word COM test tools for the RWB docm (VBScript) and the analysis tools next to them

Born during the RWB 1.1.1 clean-up of 2026-11-10 (release train RT-13 to RT-19, follow-up register `C:\adaept\aeBibleAddin\rvw\RWB 1.1.1 Follow-up Issues.md`). They test a VBA module in a real Word without touching the real docm, and read a saved docm without Word at all.

## 1. The COM scripts in this folder (a hidden Word of their own)

| Script | What it does |
|---|---|
| `run-counts.vbs <copy.docm> <modules> <Module.Function;...> <result.txt>` | Imports the modules into a scratch document, opens the COPY read-only with macros off, runs each function with the copy's name as argument, writes the results (the five clean-up counts in one run). |
| `fix-and-verify.vbs <copy.docm> <modules> <Module.Count> <Module.Fix> <result.txt>` | count, fix, count, fix again (expect the second fix to return 0); saves the document text before and after (`<result>.before.txt`, `.after.txt`) so `sha256sum` shows the fix did not change the text. Closes the copy WITHOUT saving. |
| `compile-check.vbs <src folder> <ClassName> <numbers 94;95> <result.txt>` | Imports every module of `src` into a scratch document, adds a probe that instantiates the class and reads `GetTestDescription(n)`: proves an edited `aeBibleClass.cls` compiles before it is imported into the real docm. |
| `run-export.vbs <src folder> <source.docm> <result.txt>` | Runs the scripture-only export (`ExportScriptureOnlyDocx`) on a COPY in a hidden Word. The export takes every path from the folder of the source: put the copy in its own new folder under `C:\adaept\aeTest\` with a `py\` folder (`strip_ribbon.py`, `verify_char_style_change.py`) and an empty `rpt\`. About 5 minutes; start it in the background and poll `<result.txt>` and `rpt\RadiantWordBibleExport.txt` (the verdict is the last line, `Result: CLEAN (0 issues)`). Does NOT force macros off: with `AutomationSecurity = 3` the export's own `Documents.Open` fails with error 40040 after 6 s (probe, 2026-11-11). Runtime errors inside the export land in `<result.txt>.err.txt`. |
| `insert-front.vbs <src folder> <copy.docm> <front-matter.txt> <result.txt>` | RT-22a: creates the RWB front-matter styles and inserts the front-matter text into a COPY (`CreateFrontMatterStyles`, `FixStyleSlotFonts`, `InsertFrontMatter`, a second insert that must refuse), writes the section page setup and the first 20 paragraphs, and saves the copy. The proof that followed: `py\docm_front_verify.py` on the saved copy, `py\docm_front_compare.py` against another docm, then `run-export.vbs`. A hidden Word repaginates slowly, so the vertical positions are indicative. |
| `dump-wordopenxml.vbs <copy.docm> <out.xml>` | Writes `Content.WordOpenXML` (about 31 MB, about 10 s) of the copy, the same package Office.js `getOoxml()` returns. |
| `examples\hyg_rt13_compare.vbs` | A template, not a tool (paths inside): records the properties of every touched run before a fix and compares them after. It found that `Font.Reset` strips character styles. |

Every script refuses to run (exit 2) while any `WINWORD.EXE` is running: `CreateObject("Word.Application")` may hand the script the owner's open Word and `Quit` would close it. Close Word first.

### Rules learned the hard way

1. **Copies only**, under `C:\adaept\aeTest\`. Never OneDrive, never the real docm. `AutomationSecurity = 3` (macros off) is set before the copy is opened so `Document_Open` never runs.
2. **`Application.Run` needs the scratch document active** and the name `Module.Procedure`; pass the target document's NAME (the modules take an optional document name instead of `ActiveDocument`).
3. **A hang with idle CPU is a hidden compile-error dialog** (`DisplayAlerts` does not hide it), for example a module-level `Const` placed after a procedure. Check before running: no `Const` after the first procedure (`grep`), and run `compile-check.vbs`. If a run hangs, find the hidden Word by start time and empty window title and kill only that PID, never by image name.
4. **A hidden Word does not run Word's background work.** Automatic language detection re-tagged the name-list verses German within hours in a visible session, and no copy test could show it. A clean count on a copy is a snapshot: verify in a visible session AND on the saved file (section 2).
5. **Long runs go in the background** (the full docm takes minutes): `( cscript //nologo script.vbs args > run.log 2>&1 & )`, then poll the result file. The Immediate window is not visible to a script.
6. **Walking `doc.Paragraphs` through COM on the 33,824-paragraph docm is far too slow** (over ten minutes); bisect on `Range.LanguageID` or read `Content.WordOpenXML` once instead.
7. Git Bash has no `python`, `strings` or `wmic`; run Python in WSL (`wsl.exe --exec python3 -I`, with `MSYS_NO_PATHCONV=1` so Git Bash does not rewrite the paths) and PowerShell from a `.ps1` file.
8. **A style's Next style is resolved in the ACTIVE document.** Assigning `NextParagraphStyle` while another document is active (the scratch document that holds the modules) fails with 5171; activate the target first. Setting it on a 12 pt style also resets the size to Normal's 9, so set the next style before the size (RT-22a, 2026-11-11).
9. **Insert and export from VBA need macros enabled in the hidden Word.** `run-export.vbs` and `insert-front.vbs` open copies of the docm; with `AutomationSecurity = 3` a `Documents.Open` made inside the export fails with 40040 after about 6 s. Kill only the hidden Word by its process ID (empty window title, start time) if a probe hangs, never by image name; the copy was not saved, so nothing is lost.
10. **A section break inserted at a paragraph mark leaves the mark behind** (an empty paragraph at the start of the next section); insert the last paragraph of a section without its own mark and put the break right after its text. `PageSetup.TextColumns.Spacing` cannot be set on a single column (error 5993). A section break copies the page setup of the section it splits, so copy the wanted page setup explicitly (FU-34).

## 2. Reading a SAVED docm without Word

```
mkdir diag && cd diag
unzip -p ..\copy.docm word/document.xml > document.xml
unzip -p ..\copy.docm word/styles.xml > styles.xml
unzip -p ..\copy.docm word/theme/theme1.xml > theme1.xml
unzip -p ..\copy.docm word/footnotes.xml > footnotes.xml
```

- **The add-in's own port** (the strictest check): in `C:\adaept\aeBibleAddin\taskpane`, `DIAG_DIR=C:/path/to/diag npx vitest run src/audit-hygiene.saved-docm.test.ts --reporter=verbose` (the flag is needed: vitest hides the output of a passing test) prints the four numbers, every finding (paragraph, style, font or language and its source, text) and every slot violation.
- **The front-matter checks** (RT-22a): `py\docm_front_verify.py <docm or docx> [text file]` (every front paragraph against `frontmatter\rwb-front-matter.txt`: style, wording, italic spans, justification, section, no stray fonts, the first sections' page setup and footers) and `py\docm_front_compare.py <A.docm> <B.docm>` (the first 18 paragraphs of two files, 0 differences expected between the owner's docm and a scripted insert).
- **The Python tools** in `..\..\py` (run each with no argument to see its usage line): `docm_text_compare.py` (text identity: character count, SHA-256, de-DE count), `docm_text_diff.py`, `docm_style_compare.py` (character properties of every style), `docm_lang_compare.py` (the German tags between two saves), `docm_rpr_signatures.py` (run-property signatures of the verse paragraphs), `docm_stray_runs.py` (Test 98 rule), `docm_paragraph_containers.py` (table cells, text boxes stored twice), `docm_show_paragraphs.py`, `docm_inuse_rule.py`.

## 3. Facts that these tools established (so they are not rediscovered)

- The cleaned docm's identity: `w:t` text SHA-256 `027b4dfd6313470e...` (4,200,306 characters) from 2026-11-10 on; a clean-up must not change it.
- `w:noProof` on every run is intended (owner decision A, FU-32); "Detect language automatically" stays off (FU-30).
- Text boxes are stored twice (`mc:Choice` and `mc:Fallback`) and are not body paragraphs for Office.js (FU-28).
- Word's Find and `Font.Name` do not see a direct Microsoft font on a character that Word draws with the East Asian font (the en dash, FU-26, FU-31).
