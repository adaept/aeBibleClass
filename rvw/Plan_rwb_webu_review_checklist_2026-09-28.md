# R15 plan: a reviewable rwb/webu (and later rwb/webbe) checklist file - 2026-09-28

Follow-on to the R14 categorizer work (`aeRWB/tools/web-diff/README.md`'s
"docm-vs-WEBU/WEBBE full change categorization" section;
`project_docm_categorization_and_nt_variant_policy` memory). That work
found the multi-hunk `other` bucket has hit diminishing returns for
classifier-rule batch-mining - 1,116 of 1,211 distinct leftover hunk
shapes occur exactly once (checked against `aeRWB` commit `7645e9e1`,
`aeBibleClass` commit `6c95c71`: WEBU `other` 1,454/12,201 = 11.9%, WEBBE
`other` 1,952/14,243 = 13.7%). What remains is genuine editorial
divergence needing individual human judgment, not more pattern-finding.

**This is a plan document only - nothing here is built yet.** Per the
operator's 10-point brief (2026-09-28), captured verbatim as the ten
numbered requirements below, then worked through with a concrete design,
a status table, and an explicit pros/cons/risks/suggestions section.

## Status

⚪ Not started - awaiting operator review of this plan before any code is written.

## The ten requirements, as given

1. Update the memory/README with current status, but do not consider the
   categorization work closed yet. **✅ Done** as a prerequisite to this
   plan - see `aeRWB/tools/web-diff/README.md` commit `5fd4f4d` and the
   `project_docm_categorization_and_nt_variant_policy` memory update.
2. Prepare a plan that will list the remaining "leftover-hunk shapes" to
   create one file with both comparisons per verse.
3. The file will be the shape of `rwb.txt` itself and also UTF-8.
4. It will add a new tab-separated starting field of `rwb` and `webu` for
   each verse record.
5. The script generating the file will also allow the same to be run
   against WEBBE so that comparison can be made for that separate
   publication too.
6. This file is intended to be reviewed in GitHub Desktop so the list of
   changes can be readily reviewed.
7. A mechanism is desired such that a checkmark can be added to the `rwb`
   record to signal its acceptance, connected to the classifier such that
   it is idempotent and continues to work across a number of editing and
   approval cycles.
8. The list is expected to reduce significantly; any remainder will be
   corrected such that a new export from the docm run through the
   classifier will allow them to also be checked off.
9. This process should conclude the review of the rwb/webu (and later
   rwb/webbe) work.
10. Pros/cons/benefits/risks and suggestions for consideration to be
    reviewed in this planning doc.

## 1. Scope and definition of done

**In scope for the first pass: WEBU only** (point 9's own phrasing, "and
later rwb/webbe", plus the general-divergence risk noted in §8 below -
reviewing the same content twice in parallel wastes effort). WEBBE's
checklist is the same tool, same file shape, run a second time once WEBU's
list is worked down - not a separate design.

**"Done" is objectively defined**: the checklist file's unchecked-entry
count reaches zero. At that point every docm-vs-WEBU divergence has either
been (a) fixed in the docm, (b) explained by a classifier rule, or (c)
explicitly reviewed and accepted by the operator as deliberate RWB
editorial content. This is a stronger, more auditable completion criterion
than "the `other` percentage looks low enough."

## 2. File shape

Matches `rwb.txt`'s own conventions (UTF-8, tab-separated, `Book C:V` ref
key) per requirement 3, with one addition per requirement 4 - a leading
`rwb`/`webu` field, so **every verse becomes two adjacent lines** instead
of one:

```
RWB-REVIEW
docm-vs-WEBU, generated <date> from aeBibleClass <docm commit> / aeRWB <engwebu.txt commit> - see aeRWB/tools/web-diff/README.md R15
rwb[ ]	Genesis 2:5	No plant of the field was yet in the earth, and no herb of the field had yet sprung up...
webu	Genesis 2:5	No plant of the field was yet on the earth, and no herb of the field had yet sprung up...
rwb[ ]	Genesis 4:7	If you do well, will it not be lifted up?...
webu	Genesis 4:7	If you do well, won't it be lifted up?...
```

- Two header lines, matching `rwb.txt`'s own `translation`/`source` header
  shape (so a reader immediately recognizes the family of file this is).
- One `rwb`/`webu` pair per unresolved verse, in canonical book order
  (same order `diffBibles` already walks) - not grouped by category, since
  the whole point is these are the *un*explained ones.
- The `rwb` field carries the acceptance mark (requirement 7) - proposed
  as `rwb[ ]` (unchecked) / `rwb[x]` (checked), a plain-text checkbox the
  operator edits directly in a text editor or even GitHub Desktop's own
  diff-adjacent editor. `webu` never carries a mark - it's read-only
  reference context sitting next to the record being judged.
- **"rwb" here means the docm's current text** (the same content that
  becomes `rwb.txt` after a sync), not `rwb.txt` itself - `rwb.txt` can
  lag behind an unsynced docm fix by a cycle, and reviewing stale text
  would defeat the purpose.
- **Explicitly NOT interchangeable with `rwb.txt`** despite the shape
  resemblance - four columns after splitting on the `rwb[ ]`/`webu`
  marker vs. `rwb.txt`'s plain two, and every other line is WEBU reference
  text, not RWB content. No existing tool (`docm.rwb.diff`,
  `apply-docm-rwb-*-sync.mjs`, etc.) should ever be pointed at this file.

## 3. Idempotency and the checkmark/classifier connection

Requirement 7 asks for the checkmark to be "connected to the classifier"
and to survive "a number of editing and approval cycles." This is the
part of the design with the most real risk (see §5) and the part most
worth the operator's attention before building anything.

**Proposed mechanism, mirroring a pattern already proven in this
codebase** (`divine-names-census.mjs`'s `RATIFIED_EXCEPTIONS`/
`KNOWN_INSCRIPTION_REFS`/`KNOWN_TITLE_REFS` - ref-keyed manual overrides
that sit *alongside* the mechanical classifier, not inside it):

1. **Source of truth for acceptance state is the checklist file itself**,
   git-tracked (see §4) - not a separate ledger. On each regeneration, the
   script reads the *previous* committed version of the file first.
2. For every verse currently in the categorizer's `other` bucket:
   - If it existed in the previous file **with `rwb[x]`** *and* both the
     `rwb` and `webu` text are byte-identical to what was there before ->
     carry the checkmark forward (`rwb[x]`), still listed.
   - If it existed before but the text changed (either side) -> **reset to
     `rwb[ ]`** - an approval is tied to the exact text pair it was given
     for, never silently carried onto different content. This is the
     single most important correctness rule in the whole design.
   - If it's new (never seen before) -> `rwb[ ]`.
3. For every verse that was in the previous file but is **no longer** in
   the current `other` bucket at all (whether because the docm was fixed,
   or because a classifier rule now explains it) -> **drop it from the
   file entirely.** Its disappearance from the diff *is* the resolution
   signal (requirement 8) - no separate "resolved" list needed.
4. **The "connected to the classifier" part**: once a verse carries
   `rwb[x]`, feed that ref set back into `docm.categorize` itself as a new
   ref-keyed override (same shape as `nt-variant-addition`'s
   `KNOWN_VARIANT_ADDITION_REFS`) - e.g. an `operator-accepted` category -
   so the categorizer's own `other` percentage in its console/summary
   output reflects human sign-off too, not just mechanical pattern
   matches. Read from the checklist file at classify time (or a small
   generated sidecar derived from it - see the open question in §7).

## 4. Script and tracking

- New script, next in the existing `R`-numbering: `aeRWB/tools/web-diff/
  generate-review-checklist.mjs` (R15), `npm run docm.review-checklist --
  webu` / `-- webbe` (source parameterized per requirement 5, reusing
  `classifyVerse`/`splitHunks`/`diffBibles` exactly like `docm.categorize`
  does - no new diffing logic, just new output shaping).
- Output: `aeRWB/rwb-webu-review.txt` (and later `rwb-webbe-review.txt`),
  at repo root alongside `rwb.txt`/`engwebu.txt`/`webbe.txt` - **git-tracked,
  not gitignored**, unlike `categorize/`. This is a deliberate departure
  from this repo's usual "regenerable output stays gitignored" convention
  (see `.gitignore`'s existing `diff/`/`census/`/`webbe-diff/`/
  `categorize/` entries) - required because requirement 6's GitHub Desktop
  review workflow only works on a tracked file (untracked files show as a
  single opaque "new file" add, not a line-by-line diff), and because the
  checkmark state itself needs to persist and be reviewable via git
  history across sessions.

## 5. Pros / Cons / Benefits / Risks (requirement 10)

**Benefits**
- Turns an otherwise-unreviewable 1,454-verse list into a concrete,
  git-diffable, incrementally-workable artifact - GitHub Desktop's native
  diff view needs no new tooling to consume this.
- Durable audit trail: git history on this one file *is* the review log -
  who accepted what, when, visible directly in commit history.
- Architecturally consistent - reuses the ref-keyed-override pattern
  already proven for `nt-variant-addition` and in
  `divine-names-census.mjs`, rather than inventing a new mechanism.
- Generalizes to WEBBE with the same script, no new design.
- Gives the categorization effort a genuine, objective "done" (count -> 0)
  instead of an open-ended percentage to eyeball.

**Risks / open design questions**
- **Scale.** ~1,454 verses x 2 lines is a ~2,900-line file. A single
  GitHub Desktop diff of that size across many small edits per session
  may be slow or unwieldy to review in one sitting. (Splitting per-book is
  a possible future refinement - not in the requirements as given, noted
  as a suggestion below rather than assumed.)
- **Round-trip parsing fragility - the biggest risk in this design.** The
  generator must re-parse its own *previous* output faithfully. If the
  operator hand-edits the file in ways beyond toggling `rwb[ ]`/`rwb[x]`
  (adds a note, reorders lines, fixes a typo in the `webu` reference text
  by mistake), the next regeneration needs a well-defined, tested
  contract for what happens - silently guessing wrong here could lose
  real review work. Needs explicit test coverage before this is trusted
  with real approval cycles, the same rigor `lib.test.mjs` already applies
  to the parsing primitives this would build on.
- **Checkbox semantics are ambiguous as specified.** Does `rwb[x]` mean
  "permanently correct as deliberate RWB content" (never revisit) or "I
  looked, fine for now" (still open to reconsideration)? The mechanism as
  designed treats it as the former (a checked verse silently drops out of
  future review once resolved-or-accepted) - worth the operator
  confirming that's the intended meaning before building it.
- **File-format collision risk.** Structurally resembling `rwb.txt`
  invites an accidental "wire it into an existing sync tool" mistake
  later, precisely because it looks similar. Mitigated by naming
  (`rwb-webu-review.txt`, not anything containing bare `rwb.txt`) and by
  the README documenting the distinction explicitly (§2 above) - still
  worth flagging as a real risk class, not just a naming nicety.
- **Departure from this repo's gitignore convention.** Every other
  regenerable categorization artifact (`categorize/`, `diff/`, `census/`)
  is gitignored; this one deliberately isn't, for good reason (§4) - but
  it means every regeneration is a commit-worthy event, and the operator
  should expect more commit traffic on this one file than the rest of the
  gitignored-output family.
- **WEBBE overlap.** Per requirement 9's own "(and later rwb/webbe)"
  phrasing, running both checklists in parallel would mean reviewing
  substantially the same underlying content twice (WEBU and WEBBE only
  differ in spelling, already handled separately by the
  `spelling-variant` category) - sequencing WEBU-first, WEBBE-second
  avoids that duplication.

**Suggestions for consideration (not requested, offered as options)**
- Sequence strictly WEBU-first, WEBBE-second (matches requirement 9's own
  ordering) rather than generating both checklists at once.
- A free-text notes column (a 4th tab-separated field on the `rwb` line)
  for the operator's own bookkeeping - e.g. "deliberate paraphrase" vs.
  "fixed in Word, pending re-export" - genuinely optional, easy to add
  later without a redesign if wanted.
- Have the generator **warn**, not just silently reset, when a
  previously-`[x]`-checked verse's text changed and its mark was dropped -
  makes the "why did this come back?" case legible without digging into
  git history.
- Per-book splitting of the output file as a scale mitigation, only if the
  single-file version proves unwieldy in practice - deliberately not
  designed in up front, since it adds real complexity (N files instead of
  1, cross-file idempotency) for a problem that might not materialize.

## 6. Workflow once built

1. `npm run docm.review-checklist -- webu` generates/updates
   `rwb-webu-review.txt` against the current docm export.
2. Operator reviews in GitHub Desktop - each verse pair is one small,
   readable diff hunk if anything changed since last run.
3. Operator either (a) edits the docm in Word for a real defect, re-runs
   `ExportDocmVersesToRWBFormat`, then regenerates the checklist (the
   fixed verse drops off automatically - §3 step 3), or (b) marks `rwb[x]`
   directly in the file for a verse judged to be deliberate RWB content,
   commits that change themselves.
4. Repeat until the file's unchecked-entry count reaches zero.
5. Only then, generate and begin the WEBBE checklist the same way.

## 7. Open questions for the operator before implementation begins

- Confirm the `rwb[ ]`/`rwb[x]` inline-checkbox design (§2/§3) over the
  alternative of a separate small sidecar "accepted-refs" ledger file
  (lower round-trip risk, but the checkmark wouldn't live directly next to
  what's being judged, and a second file to keep in sync is its own
  complexity).
- Confirm checkbox semantics: permanent acceptance vs. "revisit later" (§5).
- Confirm this file should be git-tracked (departure from this repo's
  usual convention) rather than gitignored with approvals stored
  elsewhere.
- Confirm scope (WEBU-only first pass, per §1) before WEBBE work starts.
