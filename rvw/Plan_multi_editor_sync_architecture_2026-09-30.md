# Plan: Multi-Editor / Multi-Locale Sync Architecture (Provisional)

**Status: ⚪ Not started - provisional reference document only, not a committed plan.**

## 1. Origin of this document

Raised 2026-09-30 by the operator after hitting a real VBA bug during the R15
review-checklist session (see
[[project_rwb_review_checklist_tool]]/`Plan_rwb_webu_review_checklist_2026-09-28.md`):
after an `ImportAllVBAFiles` full reimport, the ribbon's Go-To navigation
resolved a typed book name but silently stuck at Chapter 1 for chapter/verse,
even after re-entering the values. Root cause (confirmed by reading
`src/aeRibbonClass.cls`): `CaptureHeading1s` caches book/chapter heading
character positions in `Static`/`Private` VBA state, valid only within one
VBA-project session; a full reimport resets that state, and the next
navigation attempt needs a fresh rescan (close/reopen the `.docm`, or call
`CaptureHeading1s bForce:=True`) that wasn't obvious from the symptom alone.
This was already documented once, in passing, in
`rvw/Code_review 2026-04-20.md` §6 - not surfaced prominently enough. Fixed
2026-09-30 by stating the reminder directly in `ImportAllVBAFiles`'s own
"Import Complete" `MsgBox` (`src/basImportWordGitFiles.bas`) and in
[[feedback_importallvbafiles_error17]].

That incident prompted a broader question, since the underlying cause -
**a single-user desktop app's in-memory index/cache has no concept of "this
needs to be kept in sync with something else"** - is a narrow instance of a
much bigger question this project will eventually have to answer: **today,
exactly one person edits exactly one `.docm` file, serially, with git as the
audit trail.** What happens when the i18n vision
([[project_i18n_architecture_vision]]) is no longer non-blocking, and there
are multiple editors, multiple locales, and a `.docx`/JS editing surface
instead of `.docm`/VBA?

This document is a provisional placeholder to think through that question
*before* it's urgent, not a design being built now.

## 2. The five sub-questions asked

a. Is the VBA ribbon-navigation-cache-reset-after-import behavior a known,
   documented limitation? **Yes - confirmed and now documented**, see §1
   above and [[feedback_importallvbafiles_error17]] for the full mechanism.

b. **Task for the JS port**: check whether an analogous "stale cached
   index/state after reload" issue exists in the `aeBibleAddin` taskpane.
   See §3.

c. How would multiple (future, i18n) editors work with the `.docx`/JS
   pathway during an editing cycle? See §4.

d. Per-verse real-time sync (e.g. via GitHub) vs. a batch/daily push-pull
   model - which is better? See §5.

e. This document itself, per the operator's request for "a provisional plan
   document that can be kept as a reference with a background summary of
   the origin for the idea."

## 3. Task for the JS port: check for analogous stale-cache bugs

Not yet investigated - recording as a task, not a finding.

The VBA bug's shape: a per-session, in-memory index built lazily and cached
across calls, invalidated only by a specific signal (`ActiveDocument.Saved`
state) that turned out not to cover every case that actually invalidates it
(a VBA project reset isn't a "the document changed" event, but it does wipe
the cache's backing state).

`aeBibleAddin`'s taskpane (per
[[project_phase3_audit_engine_port_plan]]) has at least one analogous shape
already: `ooxml-run-styles.ts`'s OOXML-parsing approach and any per-run
"Shape" collection caching built up during a session. Office.js taskpanes
also reload differently from VBA - a taskpane reload (F5, or Word
re-sideloading the add-in) is a full JS-runtime restart (closer to VBA's
project-reset case), but the *document* itself persists across that reload
in a way the old in-memory JS state doesn't - structurally the same failure
shape as the VBA bug (index says "stale" or "empty," document says
"unchanged," UI shows something inconsistent with either).

**Concrete check to run** (not done yet): does any `aeBibleAddin`
taskpane code cache a document-derived index (style lookups, heading
positions, run-style scans) across `Office.onReady`/reload boundaries
without an explicit invalidation check? If so, does reloading the taskpane
mid-session (without closing/reopening the Word document) produce a
VBA-style "stuck" navigation state? This is a live-check item, in the same
family as the other pending Shape live-checks already tracked in
`word-audit-scanner.LIVE_CHECK.md` - not urgent (no known symptom yet, this
is a preventive check prompted by the VBA analog), but worth a line item
there next time that file is picked up.

## 4. Multi-editor workflow for `.docx`/JS

Not designed - the i18n vision itself is still non-blocking/long-term per
[[project_i18n_architecture_vision]]. Two shapes worth distinguishing when
this becomes real, framed as a question rather than a decision:

- **One editor per locale (mirrors today's model)**: each language gets its
  own `.docx`, one primary translator/reviewer at a time, serial edits,
  no real concurrent-editing problem to solve - just the export/verify/
  commit/push cycle this exact session has been running all day, generalized
  to N language repos instead of one. Lowest new infrastructure cost.
- **Multiple simultaneous editors on the same locale's `.docx`**: a real
  concurrent-editing problem (two people editing the same verse, or adjacent
  verses, at the same time) - this is the case that would actually need new
  tooling (locking, conflict detection, or a real collaborative-editing
  backend). Not clearly needed yet; most Bible-translation review workflows
  (per the SIL/Paratext ecosystem this project explicitly doesn't want to
  depend on, but also hasn't disproven the working patterns of) tend toward
  one translator + one or more reviewers in sequence, not simultaneous
  same-verse editing - worth confirming this assumption before ever building
  for the harder case.

## 5. Sync model: per-verse real-time vs. batch/daily push-pull

**Observation, not yet a final recommendation**: the batch/git model is
already in active, successful use *today* - this exact session ran it
manually and effectively (many rounds of: Word edit -> `Export...ToRWBFormat`
-> diff-verify against WEBU -> `git commit` -> next batch), resolving 200+
verses in one day with a full audit trail (commit messages naming exactly
which verses changed and why) and zero data loss. That's real, observed
evidence for one side of this tradeoff, not just theory.

**Per-verse real-time sync (e.g., "push straight to GitHub on every verse
edit"):**
- Pro: editors always see the latest state; no batch lag.
- Con: GitHub/git is not designed as a low-latency single-record datastore -
  either every verse edit becomes its own commit (extremely noisy history,
  defeats the purpose of a reviewable audit trail) or it needs a *different*
  backend entirely (a real API/database) with git relegated to a periodic
  export target instead of the live system of record.
- Con: requires the editor to be online continuously; translators in
  low-connectivity regions (a real consideration for a global-reach i18n
  goal) are worse served by this than by an offline-first model.
- Con: introduces a genuinely new failure class - simultaneous edits to the
  *same* verse racing each other - that the current model structurally
  avoids (one person, one file, one export cycle at a time).

**Batch/daily push-pull (git, used as it normally is):**
- Pro: already proven this session at real scale (200+ verses, one day).
- Pro: git's diff/blame/revert/PR tooling is mature, and - critically for a
  Bible text - gives a *reviewable, attributable, revertable* record of
  every wording change, matching this project's own established practice of
  verifying every single change against WEBU/KJV before committing. A
  real-time system would need to reinvent this review gate, not just make
  edits appear faster.
- Pro: offline-friendly - an editor can work all day locally and sync once,
  which matters if this project ever has translators with poor connectivity.
- Con: eventual consistency (a translator's changes aren't visible to
  others until the next sync) - assessed as a low cost for a slow-cadence,
  carefully-reviewed text, not a live-collaboration document.

**Provisional recommendation**: keep the batch/git model as the system of
record when this becomes real, for the reasons above - it's already
validated by this session's own workflow, not just a theoretical preference.
If a future need for lower latency or a friendlier non-technical-editor
experience emerges, the natural evolution is a **hybrid**: a lightweight
web-based review UI (in spirit, like this session's own R15
review-checklist tool, or a purpose-built "translator's workbench") that
non-technical editors use directly, which writes to a per-locale git repo
on the backend, invisible to the editor - approachability on the front end,
git's audit trail and offline-friendliness preserved on the back end. Not a
concrete design - a direction to keep in mind if/when this stops being
non-blocking.

## 6. Explicitly out of scope for this document

- No engineering scheduled by this document. Per
  [[project_i18n_architecture_vision]]'s existing "don't propose actually
  starting the architecture prep checklist items unprompted" guidance, this
  extends the same discipline to sync architecture specifically.
- No decision between "one editor per locale" vs. "concurrent editors per
  locale" in §4 - flagged as an open question to confirm with real
  translator workflow data if/when this becomes active, not decided here.
- The JS-port task in §3 is unstarted - a check to run, not a finding to
  act on yet.
