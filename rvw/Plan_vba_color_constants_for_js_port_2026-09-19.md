# VBA color-constant grounding for the aeBibleAddin JS port - 2026-09-19

## Origin

`aeBibleAddin`'s Phase 3 audit-engine port (`adaept5tudio/docs/aeBibleClass-word-addin-conversion-plan.md`
§15.11/§15.12) is blocked on four cases -
`CountParagraphMarks_ArialBlack` (39), `CountParagraphMarks_ArialBlackDarkRed` (40),
`Count_ArialBlack8pt_Normal_DarkRed_NotEmphasisRed` (41),
`CountParagraphMarksWithDarkRedFormatting` (43) - because they compare
`Font.Color` against Word's built-in `wdColorAutomatic`/`wdColorBlack`/
`wdColorDarkRed` constants, and Office.js has no equivalent named
constants. The port needs the actual numeric values to hard-code.

Operator decision (2026-09-19): answer this with a checked-in,
re-runnable VBA test (`RUN_THE_TESTS(88)` / `AuditColorConstants`) instead
of a one-off `?Hex(wdColorAutomatic)` query in the Immediate window, so
the values are established *actively* going forward - versioned,
diffable, and re-runnable if Word/the constants ever change - rather than
captured once and left to rot in a memory file or doc.

## a) Test 88 - AuditColorConstants (built, awaiting operator run)

Added to `src/basBiblePalette.bas` (`Public Function AuditColorConstants()
As Long`, plus private helper `LiveCheckAutomaticColorReadback`) and wired
into `src/aeBibleClass.cls` as `RUN_THE_TESTS(88)`, following the standard
four-location pattern in `md/Adding_To_Bible_Test_Class.md`:

- `MaxTests` 87 -> 88
- `Expected1BasedArray` values +1 entry (expected `6`)
- `RunTest (88)` added after `RunTest (87)` in the run sequence
- `GetTestDescription`, `GetPassFail`, `RunTest`, `OutputTestReport` all
  have a `Case 88`
- Class-side `Private Function AuditColorConstants() As Long` is a thin
  wrapper delegating to `basBiblePalette.AuditColorConstants()`, which
  already owns the named-color SSOT for this project - no duplicated
  color logic.

**What it prints:** Long/Hex/RGB for every built-in `WdColor*` constant
actually referenced anywhere in this codebase's runtime logic (confirmed
by `grep -orhE "wdColor[A-Za-z]+"` across `src/*.bas src/*.cls` - exactly
six, not the full ~50-member enum):

- `wdColorAutomatic`, `wdColorBlack`, `wdColorRed`, `wdColorDarkRed`,
  `wdColorBlue`, `wdColorDarkBlue`

**What it also does:** a live check. It opens a throwaway invisible
document (`Documents.Add(Visible:=False)`, never `ActiveDocument` - this
diagnostic cannot touch production content), sets a paragraph's
`Font.Color` explicitly to `wdColorAutomatic`, reads it back, and reports
whether the sentinel survived or was resolved to a concrete color. This
was never confirmed before - Shape 2's original live-check (§15.6 in the
addin conversion plan) only tested inequality against white, not the
literal value "Automatic" reads back as.

**Return value:** the count of constants audited (`6`) - fixed and
content-independent, so this participates in the normal PASS/FAIL suite
going forward instead of being a one-off diagnostic that has to be run
and read by hand every time.

**Note on reading the Hex/RGB columns once run:** `WdColor` encodes
"special" values like `wdColorAutomatic` with a flag in the high byte;
the low three bytes of that encoding are not a real RGB triple. The Hex
column will very likely show `#000000` for `wdColorAutomatic` - that is
*not* a claim that Automatic equals Black. Only the raw Long value and
the live-check line answer what Automatic actually resolves to.

## c) Operator action needed - DONE 2026-09-19

`RUN_THE_TESTS(88)` run against the production document. PASS, result 6,
expected 6. Full output:

```
AuditColorConstants: built-in WdColor constants used in this codebase
  Name                Long          Hex        RGB
  wdColorAutomatic    -16777216     #000000    (0,0,0)
  wdColorBlack        0             #000000    (0,0,0)
  wdColorRed          255           #FF0000    (255,0,0)
  wdColorDarkRed      128           #800000    (128,0,0)
  wdColorBlue         16711680      #0000FF    (0,0,255)
  wdColorDarkBlue     8388608       #000080    (0,0,128)
Live check: paragraph explicitly set to wdColorAutomatic, then read back:
  Font.Color = -16777216 #000000  (sentinel preserved on readback)
```

## d) RGB/Hex reference table (confirmed 2026-09-19)

| Constant | Long | Hex* | RGB | Matches basBiblePalette entry? |
|---|---|---|---|---|
| `wdColorAutomatic` | -16777216 | #000000* | (0,0,0)* | N/A - sentinel, intentionally excluded from the palette; *do not read as "black" - see the note in (a) above* |
| `wdColorBlack` | 0 | #000000 | (0,0,0) | `Black` = RGB(0,0,0) - **confirmed match** |
| `wdColorRed` | 255 | #FF0000 | (255,0,0) | `Red` = RGB(255,0,0) - **confirmed match** |
| `wdColorDarkRed` | 128 | #800000 | (128,0,0) | `DarkRed` = RGB(128,0,0) - **confirmed match.** This was the real open question Cases 40/41/43 depended on - resolved clean, no drift |
| `wdColorBlue` | 16711680 | #0000FF | (0,0,255) | `Blue` = RGB(0,0,255) - **confirmed match** |
| `wdColorDarkBlue` | 8388608 | #000080 | (0,0,128) | `DarkBlue` = RGB(0,0,128) - **confirmed match**, closing the "unverified until now" flag on `basBiblePalette.bas` line 99's comment |

**Live-check result:** `Font.Color` read back as `-16777216` after being
explicitly set to `wdColorAutomatic` - the sentinel is **preserved on
readback**, not resolved to a concrete color (not `0`/Black). This
matters for the JS port: `Case 39`'s VBA logic already checks for *two*
distinct values (`fontColor = wdColorAutomatic Or fontColor =
wdColorBlack`), and this result confirms that dual check is necessary,
not defensive over-coding - a paragraph mark that inherited "Automatic"
will never equal Black's numeric value, so both branches are live.

**Bottom line for the JS port (Cases 39/40/41/43): unblocked.** No VBA
built-in constant used by these four cases diverges from
`basBiblePalette.bas`'s existing named-color values. The port can hard-code
`#000000`/`0` for Black and `#800000`/`128,0,0` for DarkRed with full
confidence; Automatic still needs its own sentinel handling on the
Office.js side (a separate, still-open question - see (i) below), since
`-16777216` is a Word-internal encoding, not a portable RGB value.

## e) Standard going forward

**RGB values keyed to human-readable named colors, not raw built-in enum
constants or hand-typed hex literals.** `src/basBiblePalette.bas` already
implements this exactly: a `Name -> {R, G, B, RgbLong, HexCode, Usage}`
dictionary with `ColorFromName`/`NameFromColor`/`LongToHex` conversions.
This plan adopts that module as the project's SSOT for color work going
forward rather than introducing a second scheme.

Why this, and not the built-in `WdColor` enum: RGB/hex is a universal
representation (Office.js, CSS, SVG, any renderer) - Word's `WdColor`
named constants are not. Every call site that compares against a named
`wdColor*` constant is a call site that cannot port without first
reverse-engineering a magic number, which is exactly the friction Cases
39-41/43 hit.

## f) Codebase examination - what uses built-in wdColor* today

All six confirmed by grep (`wdColor[A-Za-z]+` across `src/*.bas
src/*.cls`):

| Constant | File:Line | In scope for the JS port? | Recommendation |
|---|---|---|---|
| `wdColorAutomatic`, `wdColorBlack` | `aeBibleClass.cls` Case 39 (`CountParagraphMarks_ArialBlack`) | Yes | Once Test 88 confirms Black's value, the JS port can hard-code it; Automatic has no RGB equivalent and needs its own sentinel handling on the JS side (see §i) |
| `wdColorDarkRed` | `aeBibleClass.cls` Cases 40, 41, 43 | Yes | Blocked on Test 88 confirming whether this equals `basBiblePalette`'s `DarkRed` (128,0,0) or is a distinct value - see the open question in the table above |
| `wdColorRed` | `basFixDocxRoutines.bas:1121` | No - VBA-only doc-repair tooling, not part of `TheBibleClassTests` | No action; included in Test 88 anyway since it was cheap and closes out the full grep-confirmed list |
| `wdColorBlue` | `basTEST_aeBibleTools.bas:701,754`, `Module1.bas:1154` | No - VBA-only diagnostic/legacy tooling, not part of `TheBibleClassTests` | No action |

Per this project's own "no refactor beyond what's needed" discipline
([[feedback_late_binding]]-adjacent principle, same spirit), this plan
does **not** propose rewriting the four aeBibleClass.cls call sites to use
`basBiblePalette.ColorFromName(...)` right now - that is a separate,
explicit follow-up decision once Test 88's values are in hand and it's
clear whether the built-in and palette values actually agree. If they
disagree for `DarkRed`, that is a substantive finding (two different
"dark reds" already coexisting silently) deserving its own decision, not
a drive-by fix bundled into this task.

## g) Pointer from aeBibleAddin

`adaept5tudio/docs/aeBibleClass-word-addin-conversion-plan.md` §15.11's
Cases 39/40/41/43 row and §15.12 point 1 now point back to this document
as the mechanism for closing the gap (edited 2026-09-19).

## h) Pros / cons / benefits / risks

**Pros / benefits:**
- Turns a one-off Immediate-window fact into a versioned, re-runnable,
  git-diffable regression check - if Word ever changes these constants
  (unlikely, but so is manually re-deriving them years from now), Test 88
  catches it automatically instead of relying on someone's memory of a
  pasted value.
- Answers the Automatic-readback question with real evidence instead of
  documentation-only speculation - the exact gap §15.6 in the addin plan
  already flagged as never confirmed.
- Reuses `basBiblePalette.bas` as-is rather than inventing a second color
  mechanism - zero new duplicated logic.
- The RGB/named-color standard is inherently portable: once adopted, any
  future renderer target (web, mobile - see [[project_i18n_architecture_vision]])
  gets colors "for free" the same way Office.js will.

**Cons / risks:**
- Test 88 opens and closes a throwaway Word document on every full
  `RUN_THE_TESTS` pass - minor COM overhead, mitigated by placing it in
  the same `DoEvents`-throttled tail already used for Tests 74-87.
- The `expected = 6` check only guards the *count* of constants audited,
  not whether their *values* still mean what this plan assumes - a wrong
  value could still report PASS. This is a documentation-generation
  mechanism dressed as a regression test, not a full correctness guard.
- If `wdColorDarkRed` turns out not to equal the palette's `DarkRed`, that
  surfaces a pre-existing inconsistency in the production document/styles
  that this plan cannot fix today - it can only report it accurately.
- Migrating the in-scope call sites to `basBiblePalette.ColorFromName()`
  (beyond just documenting the values) is real, separate work that must
  be scoped and approved on its own, not silently bundled in.

## i) Notes for the global i18n / cross-platform strategy

- This is a small, concrete first step toward the renderer-agnostic color
  model implied by [[project_i18n_architecture_vision]]: RGB/hex is the
  one color representation shared by Word, Office.js, CSS, and SVG alike.
  Treat `basBiblePalette.bas`'s dictionary as the color SSOT any future
  renderer (web, mobile) should read from, not re-derive.
- Recommend the eventual JS-side color table
  (`aeBibleAddin/taskpane/src/data/*.ts`, following the existing
  `approved-styles.ts` pattern) be kept in lockstep with
  `basBiblePalette.bas` by convention (manually mirrored, count-checked,
  the same discipline already used for `approved-styles.ts`'s 49/49
  check) rather than hand-duplicated with no cross-check - this is how
  `wdColorDarkRed` vs. the palette's `DarkRed` were able to drift apart
  unnoticed in the first place.
- Worth checking whether "Automatic" needs its own first-class sentinel
  in the JS-side color model (distinct from "no color property set") -
  the live-check here answers Word's in-memory object-model behavior, but
  the OOXML-level representation (`w:color val="auto"` vs. an explicit hex
  value) is a related, still-open question that ties directly into the
  already-documented Cases 82/83/87 OOXML-parsing spike
  (`range.getOoxml()`/`DOMParser`, §15.10 in the addin conversion plan) -
  the same technique would answer both questions in one pass if that
  spike is ever built.
