# LINE ENDINGS: the one rule for every repository (owner decision 2026-11-14)

**This file is identical in `aeBibleAddin`, `aeBibleClass` and `aeRWB`.** The rule is enforced, not advisory: a violation fails the check, the pre-commit hook and the tests.

## The rule

| What | Line ending in the working copy | In git (the index) |
|---|---|---|
| VBA source: `.bas`, `.cls`, `.frm` | **CRLF** | LF |
| Windows batch files: `.bat`, `.cmd` | **CRLF** | LF |
| Everything else (`.mjs`, `.ts`, `.json`, `.md`, `.py`, `.sh`, `.vbs`, `.txt`, `.yml` ...) | **LF** | LF |
| Byte-exact files (`-text` in `.gitattributes`: release artefacts, `exports/**`, `source/**`) and binaries | untouched | untouched |
| Output that a tool rewrites in its own style (marked `eol-exempt`, for example the VBA reports in `aeBibleClass/rpt/`) | not checked | LF |

Git always stores LF. A file with CR in the index is an error whoever wrote it. The rule for a path is the `eol=` in the repository's `.gitattributes`; that file is the single source of truth, and a file type with no explicit rule is also an error.

## Why (the origin of the decision)

On 2026-11-14 `git status` in `aeRWB` listed seven files (`tools/rwb-export/canonical.mjs`, `config`, `docx-reader`, `manifest`, `plan`, `usfm`, `xlsx`) as modified with no change in them. The cause: git remembered each file with CRLF-sized stat data, while the file on disk now had LF. Looking wider showed there was no rule at all, only whatever the last writer left:

- `aeBibleAddin`: `.ts` files 86 CRLF and 80 LF, `.json` 32 CRLF and 10 LF; no `.gitattributes`.
- `aeBibleClass`: VBA source CRLF (the VBA editor writes CRLF), most `.txt` reports CRLF (VBA writes them), `.md` and `.py` mostly LF.
- `aeRWB`: everything LF.

Three writers with no coordination: git itself (`core.autocrlf=true` on this Windows machine gives CRLF on checkout), the VBA editor and VBA export code (CRLF), and Claude's Write tool and the Node scripts (LF). The owner asked whether all files should be CRLF. **No**: shell scripts and `.githooks/*` break on CR; release artefacts and `SHA256SUMS` are byte-exact; Node, Vite, TypeScript and Python in WSL read LF; and every file Claude writes is LF. VBA is CRLF because that is what the VBA editor exports and imports. So: one rule per file type, written in `.gitattributes`, checked by a program.

## How it is enforced, and what is not

- **`.gitattributes`** in each repository holds the rule (`* text=auto eol=lf`, then the exceptions). An explicit `eol=` overrides `core.autocrlf`, so a checkout is the same on every machine.
- **`check-eol.mjs`** (`aeBibleAddin/scripts/`, `aeBibleClass/tools/`, `aeRWB/tools/`: identical copies, change all three together) reads `git ls-files --eol` and exits 1 with the list of violations. `node <folder>/check-eol.mjs --fix` rewrites the working files to the rule (only line endings change; nothing is staged).
- **Pre-commit hook** `.githooks/pre-commit` runs it on every commit. Install once per clone: `git config core.hooksPath .githooks` (`aeBibleAddin`: `npm run hooks.install`).
- **Tests:** `scripts/check-eol.test.mjs` / `tools/check-eol.test.mjs` (10 tests: throw-away repositories prove every error path, and the last one checks the real repository). `aeBibleAddin`: `npm run check.eol`, `npm run test.eol`, both inside `npm run check` (`test.eol` is in `test.all`). `aeRWB`: `npm test`. `aeBibleClass`: `node --test tools/check-eol.test.mjs`.
- **What is NOT enforced:** `git commit --no-verify` skips the hook (git offers no way to forbid it); the next `check`, test run or hook run then fails until the files are fixed. There is no CI step yet (a workflow edit that cannot be tested here). Untracked files are not seen until they are staged. A tool that rewrites a tracked file in another style after a check is caught only at the next check.

## What to do

- **A violation:** `node <folder>/check-eol.mjs --fix`, then re-run the check. Never convert by hand or with an editor setting.
- **A new kind of file with a different rule:** add the `eol=` line to `.gitattributes` first (that is the only place the rule lives), then run the check.
- **Output that a tool writes in its own style** and that is tracked: mark the path `eol-exempt` in `.gitattributes`, with a comment saying which tool writes it.
- **Claude (and any script that writes files):** the Write tool and Node scripts write LF, which is right for everything except VBA and batch files. After writing or editing a `.bas`, `.cls`, `.frm`, `.bat` or `.cmd`, run `--fix`. Before ending any task that wrote files, run the check. Edit keeps the endings the file has.
- **Do not** change `core.autocrlf` to "fix" this: the `.gitattributes` rule already decides.
