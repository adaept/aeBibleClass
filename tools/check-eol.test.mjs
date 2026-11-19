// Tests of check-eol.mjs (the line-ending policy check, LINE_ENDINGS.md). Run: node --test scripts/check-eol.test.mjs  (aeRWB: npm test)
// Every case builds a throw-away git repository, so the real repositories are never touched.
// IDENTICAL COPIES of this file live next to check-eol.mjs in every repository that follows the policy (list: adaept5tudio/CLAUDE.md, "Line endings").
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync, spawnSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, rmSync, writeFileSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { check, findControlChar, findProblems, parseEolEntry, toLineEnding } from './check-eol.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const SCRIPT = join(HERE, 'check-eol.mjs');
const ATTRS = [
  '* text=auto eol=lf', '*.bas text eol=crlf', '*.bin binary', 'keep/** -text', 'out/** eol-exempt', '',
].join('\n');

function repo(files, { attrs = ATTRS, attrsFirst = true } = {}) {
  const dir = mkdtempSync(join(tmpdir(), 'eol-test-'));
  const run = (...a) => execFileSync('git', ['-C', dir, '-c', 'core.autocrlf=false', '-c', 'user.name=t', '-c', 'user.email=t@t', ...a], { encoding: 'utf8' });
  run('init', '-q');
  const put = (p, text) => { mkdirSync(dirname(join(dir, p)), { recursive: true }); writeFileSync(join(dir, p), text, 'latin1'); };
  if (attrsFirst) put('.gitattributes', attrs);
  for (const [p, text] of Object.entries(files)) put(p, text);
  run('add', '-A');
  run('commit', '-q', '-m', 'x', '--allow-empty');
  return { dir, run, put, done: () => rmSync(dir, { recursive: true, force: true }) };
}
const kinds = (dir) => check(dir).problems.map((p) => `${p.kind}:${p.path}`).sort();

test('parseEolEntry reads the index, working and attribute columns and the path', () => {
  const e = parseEolEntry('i/lf    w/crlf  attr/text eol=crlf  \tsrc/a b.bas');
  assert.deepEqual(e, { index: 'lf', work: 'crlf', attr: 'text eol=crlf', path: 'src/a b.bas' });
  assert.equal(parseEolEntry('garbage'), null);
});

test('findProblems: byte-exact files, empty files and exempt paths are never reported in the working tree', () => {
  const E = (index, work, attr, path) => ({ index, work, attr, path });
  assert.deepEqual(findProblems([E('-text', '-text', 'binary', 'a.png'), E('lf', 'none', 'text eol=lf', 'empty.txt')]), []);
  assert.deepEqual(findProblems([E('lf', 'crlf', 'text eol=lf', 'out/r.txt')], (p) => p.startsWith('out/')), []);
  assert.equal(findProblems([E('lf', 'crlf', 'text eol=lf', 'a.ts')]).length, 1);
});

test('a repository that follows its rules is clean (LF files LF, VBA CRLF)', () => {
  const r = repo({ 'a.ts': 'x\ny\n', 'm.bas': 'Sub A()\r\nEnd Sub\r\n', 'b.bin': 'x\r\ny', 'keep/f.txt': 'k\r\n', 'out/r.txt': 'r\r\n' });
  try { assert.deepEqual(kinds(r.dir), []); } finally { r.done(); }
});

test('a working file with the wrong line ending is reported, in both directions', () => {
  const r = repo({ 'a.ts': 'x\ny\n', 'm.bas': 'Sub A()\r\nEnd Sub\r\n' });
  try {
    r.put('a.ts', 'x\r\ny\r\n');
    r.put('m.bas', 'Sub A()\nEnd Sub\n');
    assert.deepEqual(kinds(r.dir), ['working:a.ts', 'working:m.bas']);
    r.put('a.ts', 'x\r\ny\n'); // mixed
    assert.ok(kinds(r.dir).includes('working:a.ts'));
  } finally { r.done(); }
});

test('CR in the index is an error that cannot be exempted: a file committed before the rule existed with CRLF', () => {
  const r = repo({ 'old.ts': 'x\r\ny\r\n', 'out/r.txt': 'r\r\n' }, { attrsFirst: false });
  try {
    r.put('.gitattributes', ATTRS);
    r.run('add', '.gitattributes');
    assert.deepEqual(kinds(r.dir), ['index:old.ts', 'index:out/r.txt']);
  } finally { r.done(); }
});

test('a type without an explicit eol rule is an error (the rule must be written in .gitattributes)', () => {
  const r = repo({ 'a.ts': 'x\n' }, { attrs: '* text=auto\n' });
  try { assert.deepEqual(kinds(r.dir), ['rule:.gitattributes', 'rule:a.ts']); } finally { r.done(); }
});

test('--fix rewrites only the working endings, content stays, nothing is staged, and the repo is clean afterwards', () => {
  const r = repo({ 'a.ts': 'x\ny\n', 'm.bas': 'Sub A()\r\nEnd Sub\r\n' });
  try {
    r.put('a.ts', 'x\r\ny\r\n');
    r.put('m.bas', 'Sub A()\nEnd Sub\n');
    const res = spawnSync('node', [SCRIPT, '--fix', '--repo', r.dir], { encoding: 'utf8' });
    assert.equal(res.status, 0, res.stderr);
    assert.equal(readFileSync(join(r.dir, 'a.ts'), 'latin1'), 'x\ny\n');
    assert.equal(readFileSync(join(r.dir, 'm.bas'), 'latin1'), 'Sub A()\r\nEnd Sub\r\n');
    assert.equal(r.run('diff', '--cached', '--name-only').trim(), '');
    assert.deepEqual(kinds(r.dir), []);
  } finally { r.done(); }
});

test('the command line exits 1 with the list on a violation and 0 when clean', () => {
  const r = repo({ 'a.ts': 'x\n' });
  try {
    assert.equal(spawnSync('node', [SCRIPT, '--repo', r.dir], { encoding: 'utf8' }).status, 0);
    r.put('a.ts', 'x\r\n');
    const bad = spawnSync('node', [SCRIPT, '--repo', r.dir], { encoding: 'utf8' });
    assert.equal(bad.status, 1);
    assert.match(bad.stderr, /LINE-ENDING POLICY VIOLATION/);
    assert.match(bad.stderr, /a\.ts/);
    assert.equal(spawnSync('node', [SCRIPT, '--nonsense'], { encoding: 'utf8' }).status, 2);
  } finally { r.done(); }
});

test('toLineEnding converts both ways and leaves a file that is already right alone', () => {
  assert.equal(toLineEnding('a\r\nb\n', 'lf'), 'a\nb\n');
  assert.equal(toLineEnding('a\r\nb\n', 'crlf'), 'a\r\nb\r\n');
  assert.equal(toLineEnding('a\nb\n', 'lf'), 'a\nb\n');
});

test('THE REAL REPOSITORY follows its own policy (this is the test that fails if the policy is bypassed)', () => {
  const root = execFileSync('git', ['-C', HERE, 'rev-parse', '--show-toplevel'], { encoding: 'utf8' }).trim();
  const { problems } = check(root);
  assert.deepEqual(problems.map((p) => `${p.kind}:${p.path}: ${p.why}`), [], 'line-ending policy violated: run node ' + SCRIPT + ' --fix, read LINE_ENDINGS.md');
});

test('Q-T1: a control character in a tracked text file is a problem that names the file, line and character', () => {
  const r = repo({ 'a.md': 'fine\n', 'dir/b.md': 'one\nC:\x07daept\n', 'c.txt': 'tab\there\r\n', 'd.bin': '\x07\x00', 'keep/e.txt': '\x07', 'out/f.txt': '\x0b' });
  try {
    assert.deepEqual(kinds(r.dir).filter((k) => k.startsWith('control:')), ['control:dir/b.md']);
    const found = check(r.dir).problems.find((p) => p.kind === 'control');
    assert.match(found.why, /BEL \(code 7\) at line 2, column 3/);
  } finally { r.done(); }
});

test('Q-T1: findControlChar knows BEL VT BS NUL ESC DEL, allows TAB LF CR, and ignores UTF-8 bytes above 127', () => {
  for (const [ch, name] of [['\x07', 'BEL'], ['\x0b', 'VT'], ['\x08', 'BS'], ['\x00', 'NUL'], ['\x1b', 'ESC'], ['\x7f', 'DEL'], ['\x01', 'U+0001']]) {
    assert.equal(findControlChar(`x${ch}y`).name, name);
  }
  assert.equal(findControlChar('a\tb\r\nc\n'), null);
  assert.equal(findControlChar(Buffer.from('caf\u00e9 \u2014 \u201cquote\u201d \u0085', 'utf8').toString('latin1')), null, 'UTF-8 continuation bytes are not controls');
});

test('Q-T1: the command exits 1 and says why, and --fix does not pretend to repair it', () => {
  const r = repo({ 'doc.md': 'path C:\x07daept\n' });
  try {
    const res = spawnSync(process.execPath, [SCRIPT, '--repo', r.dir], { encoding: 'utf8' });
    assert.equal(res.status, 1);
    assert.match(res.stderr, /control\s+doc\.md: control character BEL/);
    assert.match(res.stderr, /mangled Windows path/);
    assert.equal(spawnSync(process.execPath, [SCRIPT, '--repo', r.dir, '--fix'], { encoding: 'utf8' }).status, 1);
  } finally { r.done(); }
});
