// verify-docm-edits.mjs: prove that a saved docm differs from a reference docm ONLY by the edits of an apply-edits.vbs edit list.
//   node verify-docm-edits.mjs <before.docm> <after.docm> <edits.txt>
// Reads word/document.xml of both files (no Word). Checks: same paragraph count; exactly one paragraph per edit differs and its text equals the old
// text with the replacement applied; every other paragraph is identical in text; for the changed paragraphs the paragraph style and the sequence of
// character styles (rStyle) of the runs are the same before and after. Exit 0 = all true, 1 = something differs.
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';

const [before, after, editsFile] = process.argv.slice(2);
if (!editsFile) { console.error('usage: node verify-docm-edits.mjs <before.docm> <after.docm> <edits.txt>'); process.exit(2); }
const unesc = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
function paragraphs(file) {
  const xml = execFileSync('unzip', ['-p', file, 'word/document.xml'], { maxBuffer: 1 << 29 }).toString('utf8');
  return xml.split('</w:p>').map((p) => {
    const text = unesc([...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join(''));
    const pStyle = (p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || '';
    const runStyles = [];
    for (const r of p.split(/<w:r[ >]/).slice(1)) {
      if (!/<w:t[ >]/.test(r)) continue;
      const st = (r.match(/<w:rStyle w:val="([^"]+)"/) || [])[1] || '-';
      if (runStyles[runStyles.length - 1] !== st) runStyles.push(st);
    }
    return { text, pStyle, runStyles: runStyles.join(',') };
  });
}
const edits = readFileSync(editsFile, 'latin1').split(/\r?\n/).filter((l) => l && !l.startsWith('#')).map((l) => l.split('\t'));
const A = paragraphs(before), B = paragraphs(after);
let bad = 0;
const fail = (m) => { console.log('FAIL', m); bad++; };
if (A.length !== B.length) fail(`paragraph count ${A.length} -> ${B.length}`);
const diff = [];
for (let i = 0; i < Math.min(A.length, B.length); i++) if (A[i].text !== B[i].text || A[i].pStyle !== B[i].pStyle) diff.push(i);
console.log(`paragraphs ${A.length} -> ${B.length} | paragraphs that differ: ${diff.length} | edits in the list: ${edits.length}`);
if (diff.length !== edits.length) fail(`expected ${edits.length} changed paragraphs, found ${diff.length}`);
for (const [label, find, rep] of edits) {
  const idx = diff.filter((i) => A[i].text.includes(find));
  if (idx.length !== 1) { fail(`${label}: the old text is in ${idx.length} changed paragraph(s)`); continue; }
  const i = idx[0];
  const want = A[i].text.replace(find, rep);
  const okText = B[i].text === want;
  const okStyle = A[i].pStyle === B[i].pStyle && A[i].runStyles === B[i].runStyles;
  console.log(`${okText && okStyle ? 'ok  ' : 'FAIL'} ${label.padEnd(11)} paragraph ${i} | text as expected: ${okText} | paragraph style ${A[i].pStyle} same: ${A[i].pStyle === B[i].pStyle} | run character styles [${A[i].runStyles}] -> [${B[i].runStyles}]`);
  if (!okText || !okStyle) bad++;
}
console.log(bad ? `RESULT: ${bad} problem(s)` : 'RESULT: only the listed edits differ; text, paragraph styles and run character styles are as expected');
process.exit(bad ? 1 : 0);
