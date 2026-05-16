# 12 - Module vs Class (contributor-facing architecture rule)

This page exists for one reason: to tell a future contributor -
especially someone preparing a localization (i18n) - **when they
are allowed to edit freely and when they should stop and ask
first**.

## TL;DR

| You opened a... | What that means | What to do |
|---|---|---|
| `.bas` module | Data, helpers, lookup tables, ribbon XML callbacks, normalization rules | Edit freely; review by diff |
| `.cls` class | A stateful actor enforcing an invariant (document workflow, assertion sink, logger, ribbon controller, citation parser, long-process driver) | **Stop. Ask first.** |

The file extension is the permission gate. You do not need to
know what a class *is* to follow this rule; you only need to
notice the file extension.

## The rule

- Class-related code stays inside the class itself.
- If a class needs behaviour from elsewhere, it calls into
  another class (e.g. `aeAssertClass`, `aeLoggerClass`,
  `aeBibleCitationClass`, `aeRibbonClass`, `aeLongProcessClass`,
  `aeUpdateCharStyleClass`) - not into a module.
- Stateless specs, lookup tables, and config SSOTs live in
  modules.

## Why this exists

There are two reasons, in increasing order of importance:

**1. Architectural.** Test code and coherent stateful workflow
read better in one file. The `RUN_THE_TESTS` slot dispatcher in
`aeBibleClass.cls` is easier to audit when every slot body lives
in the same file - no hunting through `bas` modules to find what
a slot is actually doing.

**2. Social (the bigger reason).** Most VBA contributors edit
modules and treat classes as advanced territory. By concentrating
stateful behaviour inside classes, the file boundary becomes a
permission signal:

- Modules say: "data and helpers, no invariants. Safe to edit."
- Classes say: "an invariant is being maintained here. Read the
  class header, understand what it enforces, and only then
  change it - or ask."

This means a future contributor preparing an i18n release can
edit the palette table, the abbreviation list, the approved
styles SSOT, the ribbon XML, and the normalization rules
**without ever opening a class file**. The moment they need to
open a `.cls`, that is the signal to stop.

## What lives where

**Modules (`.bas`) - safe-to-edit data and helpers:**

- `basBiblePalette` - palette colour names + helpers
  (`ColorFromName`, `LongToHex`, `NameFromColor`)
- `basTEST_aeBibleConfig` - `GetApprovedStyles()` SSOT
- Ribbon XML (`xml/customUI14backupRWB.xml`) + ribbon callback
  shims in `aeRibbon*` modules
- Abbreviation maps (`md/BibleAbbreviationList.md` mirrored in
  `GetBookAliasMap`)
- Normalization rules (`py/normalize_vba.py`)
- Document fixer routines (`basFixDocxRoutines`)

**Classes (`.cls`) - ask-first stateful actors:**

- `aeBibleClass` - top-level document workflow + RUN_THE_TESTS
  dispatch. Test slot bodies live here.
- `aeAssertClass` - assertion sink
- `aeLoggerClass` - structured logging
- `aeBibleCitationClass` - citation parser
- `aeLongProcessClass` - progress / cancellation for long
  operations
- `aeRibbonClass` - ribbon state coordinator
- `aeUpdateCharStyleClass` - character-style update workflow

## Why `basBiblePalette` is a module, not a class

Pure stateless lookup table. Converting it to a class would
dissolve the safety signal described above - every palette-row
edit would require opening a class file. There is no behavioural
gain to offset that cost. The discipline only works if the split
means something, and the right meaning is "stateless data vs
stateful actor," not "everything important is a class."

The same logic applies to `GetApprovedStyles()`, the
abbreviation map, and the normalization rules. These are *specs*
edited like configuration; they should read like configuration.

## i18n-specific guidance

Preparing a localization typically touches:

- Palette additions / re-mappings - **module** edit.
- Bible-book abbreviation table for the target language -
  **module** edit.
- Approved-styles list (if locale needs different styles) -
  **module** edit.
- Ribbon strings / labels - **XML** edit + run
  `py/inject_ribbon.py`.
- Document-fixer routines that depend on locale-specific
  punctuation, soft-hyphen behaviour, casing rules - **module**
  edit.

If your locale work requires *changing the assertion harness*,
the *citation parser grammar*, the *ribbon state machine*, or
the *long-process driver*, that's a class edit - stop and ask.

## When this rule is broken

Anytime a `RUN_THE_TESTS` slot body or a stateful workflow lives
in a module, that's a divergence. The fix is "body in class,
delegate stub in module" - the class owns the behaviour, the
module exposes a one-line wrapper for Immediate-window
convenience. See `aeBibleClass.AuditBookHyperlinkStyling` (in
class) paired with `basStyleInspector.AuditBookHyperlinkStyling`
(stub) for the canonical template, established 2026-05-15.

## Related

- Memory rule: `feedback-class-encapsulation` in user memory.
- Originating discussion: `rvw/Code_review 2026-05-15.md` § 9.
- Casual-coder rationale rephrased from the
  "ask first / failure blast radius" framing used during item
  13 / 2.4.
