#!/usr/bin/env node
// check-eol.mjs - the line-ending policy check (owner decision 2026-11-14; read LINE_ENDINGS.md at the root of the repository).
//
// THE RULE (one source: the .gitattributes of the repository, which every repository carries):
//   - git STORES every text file with LF (the index): a file with CR in the index is an ERROR, whoever wrote it.
//   - the WORKING copy of a file must have the line ending its .gitattributes rule names: `eol=crlf` for VBA source (.bas .cls .frm) and the
//     Windows scripts (.bat .cmd), `eol=lf` for everything else. A file that has no explicit rule is an ERROR (the rule must be written down).
//   - `-text` (byte-exact release files, binaries) and files marked `eol-exempt` (output that a tool rewrites in its own style, for example the
//     reports VBA writes) are not checked in the working tree; the index check still applies to them unless they are `-text`.
// It reads `git ls-files --eol`, so it checks tracked files (staged ones included). Exit 0 = clean, 1 = violations (listed), 2 = usage.
//   node check-eol.mjs            check the repository that holds the current folder
//   node check-eol.mjs --fix      rewrite the WORKING files to the rule (only line endings change; no content is staged: it only refreshes git's
//                                 size cache for files without real changes, so `git status` does not list them as modified)
//   node check-eol.mjs --repo <dir>
// IDENTICAL COPIES of this file live in aeBibleAddin/scripts, aeBibleClass/tools and aeRWB/tools: change all three together.
import { execFileSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

/** One entry of `git ls-files --eol -z`: "i/lf    w/crlf  attr/text eol=lf \tpath". */
export function parseEolEntry(entry) {
  const tab = entry.indexOf('\t');
  if (tab < 0) return null;
  const m = entry.slice(0, tab).trim().match(/^i\/(\S*)\s+w\/(\S*)\s+attr\/(.*)$/);
  return m ? { index: m[1], work: m[2], attr: m[3].trim(), path: entry.slice(tab + 1) } : null;
}

/** The violations among parsed entries. `exempt(path)` says whether the working-tree check is skipped for a path. */
export function findProblems(entries, exempt = () => false) {
  const out = [];
  for (const e of entries) {
    const bytesExact = e.index === '-text' || e.work === '-text' || /(^|\s)-text(\s|$)|(^|\s)binary(\s|$)/.test(e.attr);
    if (bytesExact) continue;
    if (e.index === 'crlf' || e.index === 'mixed' || e.index === 'lone-cr') {
      out.push({ path: e.path, kind: 'index', why: `committed with ${e.index} line endings (git must store LF)` });
      continue;
    }
    const want = (e.attr.match(/eol=(lf|crlf)/) || [])[1];
    if (!want) { out.push({ path: e.path, kind: 'rule', why: 'no explicit eol rule in .gitattributes' }); continue; }
    if (exempt(e.path)) continue;
    if (e.work === '' || e.work === 'none') continue; // empty file
    if (e.work !== want) out.push({ path: e.path, kind: 'working', want, why: `working file is ${e.work}, the rule says ${want}` });
  }
  return out;
}

const git = (root, args, input) => execFileSync('git', ['-C', root, ...args], { input, maxBuffer: 1 << 28, encoding: 'utf8' });

export function listEntries(root) {
  return git(root, ['ls-files', '--eol', '-z']).split('\0').filter(Boolean).map(parseEolEntry).filter(Boolean);
}

/** Paths that carry the attribute `eol-exempt` (set in .gitattributes for output that a tool writes in its own style). */
export function exemptPaths(root, paths) {
  if (!paths.length) return new Set();
  const res = git(root, ['check-attr', '-z', '--stdin', 'eol-exempt'], paths.join('\0') + '\0').split('\0');
  const set = new Set();
  for (let i = 0; i + 2 < res.length + 1; i += 3) if (res[i + 2] === 'set') set.add(res[i]);
  return set;
}

export function toLineEnding(text, want) {
  const lf = text.replace(/\r\n/g, '\n');
  return want === 'crlf' ? lf.replace(/\n/g, '\r\n') : lf;
}

export function check(root) {
  const entries = listEntries(root);
  const ex = exemptPaths(root, entries.map((e) => e.path));
  return { entries, problems: findProblems(entries, (p) => ex.has(p)) };
}

function main(argv) {
  const fix = argv.includes('--fix');
  const ri = argv.indexOf('--repo');
  if (argv.some((a) => a !== '--fix' && a !== '--repo' && !(ri >= 0 && argv[ri + 1] === a))) { console.error('usage: node check-eol.mjs [--fix] [--repo <dir>]'); return 2; }
  const root = ri >= 0 ? argv[ri + 1] : git(process.cwd(), ['rev-parse', '--show-toplevel']).trim();
  let { problems } = check(root);
  if (fix) {
    const rewritten = [];
    for (const p of problems.filter((x) => x.kind === 'working')) {
      const file = join(root, p.path);
      const before = readFileSync(file, 'latin1');
      const after = toLineEnding(before, p.want);
      if (after !== before) { writeFileSync(file, after, 'latin1'); rewritten.push(p.path); }
    }
    // git remembers the file size of each tracked file; after a rewrite the old size makes `git status` list the file as modified although git
    // diff finds nothing. Re-adding such a file with --renormalize only refreshes that cache (the content is the same blob). Files that also
    // have real changes are NOT touched, so nothing unrelated is ever staged here.
    let refreshed = 0;
    for (const f of rewritten) {
      let sameContent = true;
      try { git(root, ['diff', '--quiet', '--', f]); } catch { sameContent = false; }
      if (sameContent) { git(root, ['add', '--renormalize', '--', f]); refreshed++; }
    }
    console.log(`check-eol --fix: rewrote the line endings of ${rewritten.length} working file(s) to their rule; refreshed git's cache for ${refreshed} of them (the others have real changes and stay as they are).`);
    problems = check(root).problems;
  }
  if (!problems.length) { console.log('check-eol: OK (every tracked text file follows its .gitattributes eol rule; the index holds LF only).'); return 0; }
  console.error(`check-eol: ${problems.length} LINE-ENDING POLICY VIOLATION(S) in ${root} (see LINE_ENDINGS.md):`);
  for (const p of problems.slice(0, 40)) console.error(`  ${p.kind.padEnd(7)} ${p.path}: ${p.why}`);
  if (problems.length > 40) console.error(`  ... and ${problems.length - 40} more`);
  const onlyWorking = problems.every((p) => p.kind === 'working');
  console.error(onlyWorking ? 'Fix: node check-eol.mjs --fix (rewrites working files only).' : 'A file in the index has CR, or has no rule: fix .gitattributes or re-add the file after fixing it (git add --renormalize <path>). Do not bypass with --no-verify.');
  return 1;
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href) process.exit(main(process.argv.slice(2)));
