# Log - WEB (2013) → WEBU structural correction findings

A running, append-only register of specific verses where the 2013 `web.txt`
baseline (and anything that inherited it unchanged - old `rwb.txt`, the
docm before editing) has been found **structurally inferior** to WEBU's
revision - not just a different editorial choice, but a demonstrable
correctness issue (an unclosed quote, a nesting inconsistency, etc.) that
WEBU's update fixed. Distinct from ordinary WEB→WEBU wording differences
(expected, not tracked here) and from RWB's own intentional editorial
choices (tracked via R3/R4's diff register, also not here).

**Why this log exists (operator, 2026-09-15):** found while doing Phase 4
Pass 2 verification (`Plan_rwb_phase4_content_sync_2026-09-15.md`) - given
`web.txt`'s age (2013, over a decade old relative to WEBU's ongoing
updates), more of these are expected to surface as Phase 4 and later work
continues. This doc is the dedicated place to record them with full
reasoning, so each one has a reviewable provenance trail instead of being a
one-off aside buried in a session transcript.

**Format per entry:** reference, the old/current states compared, the
structural reasoning for why WEBU is more correct (not just "policy says
match WEBU"), the decision, and commit references once applied.

---

## 1. ✅ 2 Kings 19:13 (opening/closing quote-mark structure across v9-13)

**Found:** while verifying Phase 4 Pass 2 (`aeRWB` census showed `rwb.txt`
at 64 hits post-sync instead of the expected 63; traced to this verse,
which sits outside both Test 70 and Test 71's pattern coverage entirely -
neither pattern's exact 3-character sequence matches any of the three
states below, so this was invisible to all prior census-based verification).

**States compared (v9-13, 2 Kings 19):**

| Source | v9 (Tirhakah's exclamation) | v10 nesting order | v13 ending |
|---|---|---|---|
| `web.txt` (2013) / pre-fix docm | **unclosed** - no `"` after "against you," | outer `'`, inner `"` (reversed) | `"'"` (3 marks) |
| WEBU (current) | closed: `..."Behold...against you," he sent...` | outer `"`, inner `'` | `'"` (2 marks) |

**Why WEBU is more correct here (not just "the policy default"):** Tirhakah's
exclamation (v9, a report about an unrelated military threat) and the
Assyrian king's taunting letter (v10-13, addressed to Hezekiah) are two
unrelated statements from unrelated speakers, narrated back-to-back - it is
not semantically coherent for one to nest inside the other. WEBU closes
Tirhakah's quote exactly where it semantically ends (v9) and uses a clean,
fully-accounted-for 2-level structure for the letter (L1 `"..."` wraps the
whole letter, L2 `'...'` wraps "the specific words" the letter contains) -
every mark WEBU opens, it closes exactly once, exactly where the content
ends. The old WEB/docm version leaves v9's quote unclosed and produces a
trailing 3rd mark at v13 that doesn't cleanly resolve to anything - either a
genuine editorial omission (the close was simply dropped at v9) or an
unusual deferred-close choice that reads as inconsistent either way. This
is a real correctness improvement in WEBU's revision, not an arbitrary
convention difference.

**Options considered:**
- **A (chosen):** match WEBU exactly - close v9's quote, swap v10's nesting
  polarity (outer `'`→`"`, inner `"`→`'`), close v13 with `'"` (2 marks,
  WEBU's order). Two of the three edits change an *existing* mark's
  identity, not just add one - more invasive than any single-mark-insertion
  done elsewhere in this project, so this went through explicit operator
  sign-off before being applied (unlike simple append-only fixes).
- **B (rejected):** leave v9/v10 as-is, only make v13 self-consistent with
  the *old* structure (`"'"` closing in old order) - internally consistent
  but doesn't match WEBU, and reproduces the same underlying inconsistency
  rather than fixing it.
- **C (rejected):** just restore old WEB's original 3-mark ending at v13 -
  reproduces the less-correct reading; contradicts this project's
  established full-WEBU-fidelity policy for no offsetting benefit.

**Decision:** Option A.

**Status:** ✅ Done - docm fixed and verified byte-for-byte against WEBU,
`aeBibleClass` `cfedd99`. `rwb.txt` synced to match (also picked up two
stale-wording differences at v9/v10, per Phase 4's established full-verse-
replacement approach) - 🟡 `aeRWB` commit pending operator review/push.
`aeRWB` census confirms `rwb.txt` back to 63/63 matching WEBU.
