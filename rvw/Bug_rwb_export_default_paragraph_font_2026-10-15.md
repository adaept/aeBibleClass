# Bug: RWB export carries a flipped `Default Paragraph Font` (found 2026-10-15)

Found by the aeBibleAddin JS port's live audit (`aeAuditSpike.auditApprovedStyles()` in English Word on
`RadiantWordBible.docx`): **1 unhide-when-used violation and 1 wrong-priority violation, both on
`Default Paragraph Font`** (UnhideWhenUsed on; Priority 2, expected 34 = its 1-based position in
`GetApprovedStyles()`). Tracking: `C:\adaept\aeBibleAddin\rvw\Code_review 2026-10-09.md`, task R-15, sections 17-18.

## Finding

- This is the known Word behaviour recorded in `rvw/Bug_docm_verse_export_skip_duplicate_2026-09-16.md` (Tests 77/78
  failed on the same style): Word silently re-surfaces `Default Paragraph Font` on almost any edit. A fix applied to the
  docm (`WordEditingConfig`, save) therefore is **not idempotent** - the next edit undoes it.
- Evidence that the flipped state is in the source and travels through the export: `word/styles.xml` of BOTH
  `Blank Bible Copy.docm` (saved 2026-10-13) and `RadiantWordBible.docx` has, for `DefaultParagraphFont`,
  `<w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/>`.
- **Root cause in the export:** `ExportScriptureOnlyDocx` (`src/basBibleOnlyExport.bas`, Pass 1 .. Pass 7) copies the
  docm to a working copy, runs many editing passes and goes straight to `oDoc.SaveAs2` (Pass 7). It never calls
  `WordEditingConfig`, `PromoteApprovedStyles` or anything touching `Priority` / `UnhideWhenUsed`. So the RWB output
  depends on whatever state the source docm happened to be in.

## Solution (owner decision 2026-10-15: fix at the source, no RWB exception; VBA stays the one oracle)

Re-apply the style flags **as the last pass before the save**, on the export's own document `oDoc`, then verify them.

- Re-run only the style-FLAG parts of `WordEditingConfig`: `PromoteApprovedStyles` and `HideUnapprovedBuiltInStyles`
  (both are property-only, no layout effect).
- Do NOT re-run `EnforceHeading2WidowControl` in the export: it changes paragraph layout (could shift pages after
  Pass 6 regenerated the Bible Index page numbers), and its result is already stored in the docm's paragraphs.
- Do NOT run `DumpPrioritiesSorted` (diagnostic only).
- Place the new pass AFTER Pass 6 (Pass 6 edits index cells and can flip the style again) and immediately BEFORE Pass 7.
- Test first (project rule): Step 1 adds the verify check alone, so its first export run is the audit (expect exactly the
  `Default Paragraph Font` finding); Step 2 adds the fix and turns the check into a halt.

## Draft VBA

### Step 1 - the check only (basBibleOnlyExport.bas)

Add this function (next to `VerifyScriptureOnlyStrip`):

```vba
' ==========================================================================
' VerifyApprovedStyleFlags
' ==========================================================================
' Read-only. For every style in GetApprovedStyles() that exists in oDoc,
' counts a violation for each of: UnhideWhenUsed = True, QuickStyle = True,
' Priority <> its 1-based position in the list - the same three properties as
' Cases 5 / 77 / 78. Styles absent from oDoc are skipped (the RWB purge removes
' many approved styles). Logs every violation. Returns the violation count
' (expected 0).
' Why: Word silently flips these (Default Paragraph Font) on almost any edit,
' so the export output must not depend on the source docm's state
' (rvw/Bug_rwb_export_default_paragraph_font_2026-10-15.md).
' ==========================================================================
Public Function VerifyApprovedStyleFlags(ByVal oDoc As Object) As Long
    On Error GoTo PROC_ERR

    Dim vApproved As Variant
    vApproved = GetApprovedStyles()
    Dim i As Long
    Dim expectedPriority As Long
    Dim oSty As Object
    Dim violations As Long

    For i = LBound(vApproved) To UBound(vApproved)
        If StyleExists(oDoc, CStr(vApproved(i))) Then
            Set oSty = oDoc.Styles(CStr(vApproved(i)))
            expectedPriority = (i - LBound(vApproved)) + 1
            If oSty.UnhideWhenUsed Then
                violations = violations + 1
                ExportLog "VerifyApprovedStyleFlags: " & vApproved(i) & " : UnhideWhenUsed=True"
            End If
            If oSty.QuickStyle Then
                violations = violations + 1
                ExportLog "VerifyApprovedStyleFlags: " & vApproved(i) & " : QuickStyle=True"
            End If
            If oSty.Priority <> expectedPriority Then
                violations = violations + 1
                ExportLog "VerifyApprovedStyleFlags: " & vApproved(i) & " : Priority " & _
                    oSty.Priority & " (expected " & expectedPriority & ")"
            End If
            Set oSty = Nothing
        End If
    Next i

    VerifyApprovedStyleFlags = violations

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & _
           ") in procedure VerifyApprovedStyleFlags of Module basBibleOnlyExport"
    VerifyApprovedStyleFlags = -1
    Resume PROC_EXIT
End Function
```

Call site in `ExportScriptureOnlyDocx`, between Pass 6 and Pass 7 (log only; a nonzero count is an "issue" so the
working copy is kept, but the export still saves):

```vba
    ' Pass 6c - approved style flags (Step 1: log only)
    Dim flagViolations As Long
    flagViolations = VerifyApprovedStyleFlags(oDoc)
    LapTimer "Pass6c_VerifyApprovedStyleFlags"
    ExportLog "Pass6c_VerifyApprovedStyleFlags: " & flagViolations & " violation(s)"
    If flagViolations <> 0 Then m_issues = m_issues + 1
```

**Expected first run:** `Pass6c_VerifyApprovedStyleFlags: 2 violation(s)` listing `Default Paragraph Font` (UnhideWhenUsed=True,
Priority 2 expected 34). If more styles are listed, that is new information - report them, do not fix yet.

### Step 2 - the fix (only after Step 1 has been run and reported)

1. `basTEST_aeBibleConfig.bas`: split `PromoteApprovedStyles` so the export can target `oDoc`:
   - rename the existing body to `Public Sub PromoteApprovedStylesIn(ByVal oDoc As Object)`;
   - change `Dim s As Word.Style` to `Dim s As Object` and every `ActiveDocument` in it to `oDoc`;
   - add the wrapper:

```vba
Private Sub PromoteApprovedStyles()
    PromoteApprovedStylesIn ActiveDocument
End Sub
```

   `WordEditingConfig` is unchanged and keeps calling `PromoteApprovedStyles`.

2. `basStyleInspector.bas`: same split for `HideUnapprovedBuiltInStyles`:
   - rename the existing body to `Public Sub HideUnapprovedBuiltInStylesIn(ByVal oDoc As Object)`;
   - change the two `ActiveDocument` references (`For Each s In ActiveDocument.Styles` and
     `Set s = ActiveDocument.Styles(CStr(nameKey))`) to `oDoc`, and `Dim s As Word.Style` to `Dim s As Object`;
   - add the wrapper:

```vba
Public Sub HideUnapprovedBuiltInStyles()
    HideUnapprovedBuiltInStylesIn ActiveDocument
End Sub
```

3. `basBibleOnlyExport.bas`: replace the Step 1 call site with the fix plus a halting check:

```vba
    ' Pass 6b - re-apply the approved style flags LAST. Word flips
    ' Default Paragraph Font's UnhideWhenUsed / Priority on almost any edit, so the output must not
    ' depend on the source docm's state. Property-only (no layout effect), hence safe after Pass 6's
    ' page-number regeneration. Nothing may edit oDoc between here and the save.
    PromoteApprovedStylesIn oDoc
    HideUnapprovedBuiltInStylesIn oDoc
    LapTimer "Pass6b_NormalizeStyleFlags"

    ' Pass 6c - verify (halts on any violation)
    Dim flagViolations As Long
    flagViolations = VerifyApprovedStyleFlags(oDoc)
    LapTimer "Pass6c_VerifyApprovedStyleFlags"
    If flagViolations <> 0 Then
        HaltExport "Pass6c_VerifyApprovedStyleFlags", flagViolations & " violation(s) - see the report"
        GoTo PROC_HALT
    End If
```

## Risks to check (do not skip)

- `PromoteApprovedStyles` first sets Priority = 99 on every paragraph/character style. The 2026-10-11 RWB check found
  linked "... Char" styles and system styles NOT at 99 in the export; running it in the export may change those and
  could differ from the v0.0 baseline compare (Pass 9) or the style inventory (`rpt\RadiantWordBibleStyleInventory.txt`).
  Re-run both after Step 2 and diff.
- Word could re-flip the style during `SaveAs2` itself. Check the saved `.docx` (not only `oDoc`): the
  `DefaultParagraphFont` entry in `word/styles.xml` must have no `w:unhideWhenUsed` and a priority that follows the list.
- `HideUnapprovedBuiltInStyles` has a `MsgBox` error handler; an error would block an unattended export.
- Cases 77/78 on the docm itself still need `WordEditingConfig` before they run, as today.

## Verification after Step 2

1. Claude: inspect `styles.xml` of the new `RadiantWordBible.docx`.
2. Owner, English Word: `await aeAuditSpike.auditApprovedStyles()` -> 0/0/0/0; W-4 regression: `countSections` 133,
   `countSectionsWithDifferentFirstPage` 0, `countVerseTextNotAligned` 0.
3. Owner, German Word: `auditApprovedStyles` still lists the unresolved approved built-ins.

## Status log

- **2026-10-15 14:47 - Step 1 run (export log):** `VerifyApprovedStyleFlags: Default Paragraph Font : UnhideWhenUsed=True` and
  `Default Paragraph Font : Priority 2 (expected 34)`; `Pass6c_VerifyApprovedStyleFlags: 2 violation(s)`; export finished with 1 issue
  (as designed for Step 1). Exactly the predicted finding, no other style listed. Test-before-fix satisfied.
- **2026-10-15 - Step 2 APPLIED to `src/basTEST_aeBibleConfig.bas` (`PromoteApprovedStylesIn`), `src/basStyleInspector.bas`
  (`HideUnapprovedBuiltInStylesIn`) and `src/basBibleOnlyExport.bas` (Pass 6b + halting Pass 6c). Not yet imported, compiled or run.**
- **2026-10-15 15:15 - Step 2 imported, compiled, `WordEditingConfig` and `ExportScriptureOnlyDocx` run (owner):** export log
  `ExportScriptureOnlyDocx: COMPLETE ... all checks passed`, working copy deleted, Pass 6c silent (0 violations), Pass 8 0 mismatches,
  Pass 9 PASS, 133 sections, 66 index rows. Pass 6b printed the expected "styles NOT found" list (the 27 styles the RWB purge removed).
  Total 328.9 s (Step 1 run 416.2 s; Pass 6b adds about 1.5 s).
- **Saved-file check (Claude, `word/styles.xml` of the new `RadiantWordBible.docx`):** `DefaultParagraphFont` is now
  `uiPriority 33` (= Priority 34) with NO `unhideWhenUsed` (before: `uiPriority 1` + `unhideWhenUsed`). Style-by-style diff against the
  2026-10-13 export (`aeTest\smoke\RadiantWordBible.docx`): same 435 styles; only 3 differ: `DefaultParagraphFont` (the fix), and
  `NoList` / `TableNormal` (`uiPriority 99` + `unhideWhenUsed` -> `98`, no `unhideWhenUsed`, i.e. now Priority 99 like the 402 other
  hidden styles). No other style changed, so the Pass 6b side effects are limited to making the two other default system styles
  consistent. Not yet done: English-Word `auditApprovedStyles` = 0/0/0/0 and the W-4 regression (owner, live).
- **2026-10-15 ~15:40 - LIVE RESULT, premise falsified (owner run, desktop Word, English, RWB doc confirmed by `VerseText` p33 / `Heading 1` p26):**
  a freshly opened `RadiantWordBible.docx` (saved file: `DefaultParagraphFont` uiPriority 33, no unhide; `NoList` / `TableNormal` uiPriority 98, no unhide)
  reads back in Word as `Default Paragraph Font` Priority 2, `No List` 100, `Table Normal` 100, ALL with `UnhideWhenUsed = true`
  = Word's own defaults for its three default system styles. **Word re-applies these defaults on every load**, so no file content can
  persist them; the in-session value set by `PromoteApprovedStyles` (34) survives only until the document is closed. This also explains the
  2026-09-14 / 2026-09-16 "flip" history (it is a load effect, not an edit effect) and why Tests 77/78 pass only in the session in which
  `WordEditingConfig` ran. Consequences: (1) the export-side Pass 6b/6c fix cannot change what the add-in reads (the file is fixed, Word
  ignores it); (2) the JS audit, which always reads a freshly loaded document, will always report `Default Paragraph Font` 1/1 (unhide + priority);
  `No List` / `Table Normal` are not in the approved list and are Table/List types, so Case 45 skips them. The earlier "earlier open was a stale
  window" hypothesis is superseded. Decision pending (owner): see the review, section 18.
- **2026-10-15 - Step 1 and Step 2 REVERTED in `src` (owner decision 2), `git revert` commits `e17df4e` and `7973620`** (originals `88c383b`, `29e0d6f` stay in history). `src/basBibleOnlyExport.bas`, `src/basTEST_aeBibleConfig.bas` and `src/basStyleInspector.bas` are back to their pre-fix state. The Draft VBA sections above are kept as the record of what was tried; do not re-apply. USER ACTION: re-import those three modules from `src` into the working docm so the docm matches `src` again (it currently has the Step 2 versions).
