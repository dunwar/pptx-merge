/**
 * Self-contained structural test for the PPTX merge engine.
 * Fixtures are generated in-memory with pptxgenjs (no private files),
 * then merged in both modes and the resulting OPC package is verified:
 * slide counts, every internal rel target resolves inside the merged
 * zip, [Content_Types] covers every part, sldIdLst matches the
 * presentation rels. Error branches are covered too.
 *
 * Run: npm test  (tsx test/merge.test.mjs)
 */
import JSZip from 'jszip';
import PptxGenJS from 'pptxgenjs';

import { MergeError, mergePptxFiles } from '../src/pptx-merge.ts';

// 1x1 red PNG — gives decks a media part so merge exercises image copying.
const RED_PX =
  'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==';

async function makeDeck(theme, slides) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  pptx.theme = { headFontFace: 'Calibri', bodyFontFace: 'Calibri' };
  for (const s of slides) {
    const slide = pptx.addSlide();
    if (s.bg) slide.background = { color: s.bg };
    if (s.title) slide.addText(s.title, { x: 0.5, y: 0.4, w: 9, h: 0.8, fontSize: 28, bold: true });
    if (s.bullets) slide.addText(s.bullets.map((t) => ({ text: t, options: { bullet: true, breakLine: true } })), { x: 0.7, y: 1.4, w: 8.6, h: 3 });
    if (s.image) slide.addImage({ data: RED_PX, x: 6, y: 4, w: 2, h: 1 });
  }
  return Buffer.from(await pptx.write({ outputType: 'nodebuffer' }));
}

function attr(tag, name) {
  const m = tag.match(new RegExp(`\\b${name}="([^"]*)"`));
  return m ? m[1] : null;
}
function dirOf(p) {
  const i = p.lastIndexOf('/');
  return i < 0 ? '' : p.slice(0, i);
}
function resolveRel(fromDir, target) {
  if (target.startsWith('/')) return target.slice(1);
  const segs = (fromDir ? fromDir.split('/') : []).concat(target.split('/'));
  const out = [];
  for (const s of segs) {
    if (s === '..') out.pop();
    else if (s === '' || s === '.') continue;
    else out.push(s);
  }
  return out.join('/');
}

async function countSlides(buf) {
  const zip = await JSZip.loadAsync(buf);
  const xml = await zip.file('ppt/presentation.xml').async('string');
  return (xml.match(/<p:sldId\b/g) || []).length;
}

async function verify(mergedBuf, expectedSlides, label) {
  const zip = await JSZip.loadAsync(mergedBuf);
  const paths = Object.keys(zip.files).filter((p) => !zip.files[p].dir);
  const lower = new Set(paths.map((p) => p.toLowerCase()));
  const problems = [];

  // 1. presentation sldIdLst vs rels
  const presXml = await zip.file('ppt/presentation.xml').async('string');
  const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
  const relIds = new Map();
  for (const m of presRelsXml.matchAll(/<Relationship\b[^>]*?\/?>/g)) {
    const id = attr(m[0], 'Id');
    const t = attr(m[0], 'Target');
    if (id) relIds.set(id, t);
  }
  const sldTags = [...presXml.matchAll(/<p:sldId\b[^>]*?\/?>/g)].map((m) => m[0]);
  if (sldTags.length !== expectedSlides)
    problems.push(`slide count ${sldTags.length} != expected ${expectedSlides}`);
  const ids = new Set();
  for (const tag of sldTags) {
    const id = Number(attr(tag, 'id'));
    const rid = attr(tag, 'r:id');
    if (ids.has(id)) problems.push(`duplicate sldId ${id}`);
    ids.add(id);
    if (!rid || !relIds.has(rid)) problems.push(`sldId ${id} has unresolved ${rid}`);
    else {
      const resolved = resolveRel('ppt', relIds.get(rid));
      if (!lower.has(resolved.toLowerCase()))
        problems.push(`slide target missing: ${resolved}`);
    }
  }

  // 2. every rels file: internal targets resolve
  let relFiles = 0;
  for (const p of paths) {
    if (!p.includes('/_rels/') || !p.endsWith('.rels')) continue;
    relFiles++;
    const ownerDir = dirOf(dirOf(p)); // strip trailing _rels
    const xml = await zip.file(p).async('string');
    for (const m of xml.matchAll(/<Relationship\b[^>]*?\/?>/g)) {
      if (attr(m[0], 'TargetMode') === 'External') continue;
      const resolved = resolveRel(ownerDir, attr(m[0], 'Target'));
      if (!lower.has(resolved.toLowerCase()))
        problems.push(`dangling rel in ${p}: ${attr(m[0], 'Target')} -> ${resolved}`);
    }
  }

  // 3. content types cover every part
  const ctXml = await zip.file('[Content_Types].xml').async('string');
  const defaults = new Set();
  for (const m of ctXml.matchAll(/<Default\b[^>]*?\/?>/g))
    defaults.add((attr(m[0], 'Extension') || '').toLowerCase());
  const overrides = new Set();
  for (const m of ctXml.matchAll(/<Override\b[^>]*?\/?>/g))
    overrides.add((attr(m[0], 'PartName') || '').replace(/^\//, '').toLowerCase());
  for (const p of paths) {
    if (p.endsWith('.rels')) continue;
    const ext = (p.slice(p.lastIndexOf('.') + 1) || '').toLowerCase();
    if (!defaults.has(ext) && !overrides.has(p.toLowerCase()))
      problems.push(`no content type for ${p}`);
  }

  // 3b. known part families must be typed by an Override with the exact
  // content type — pptxgenjs ships Default Extension="xml", so a prefixed
  // slide without an Override is retyped application/xml and PowerPoint
  // repairs it away (merged deck opens with only the first file's slides)
  const CT_BY_DIR = {
    'ppt/slides/': 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
    'ppt/slideLayouts/': 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
    'ppt/slideMasters/': 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
    'ppt/notesSlides/': 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
    'ppt/notesMasters/': 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
    'ppt/theme/': 'application/vnd.openxmlformats-officedocument.theme+xml',
  };
  const ctByPart = new Map();
  for (const m of ctXml.matchAll(/<Override\b[^>]*?\/?>/g))
    ctByPart.set(
      (attr(m[0], 'PartName') || '').replace(/^\//, '').toLowerCase(),
      attr(m[0], 'ContentType')
    );
  for (const p of paths) {
    if (p.endsWith('.rels')) continue;
    for (const [dir, ct] of Object.entries(CT_BY_DIR)) {
      if (!p.toLowerCase().startsWith(dir)) continue;
      const got = ctByPart.get(p.toLowerCase());
      if (got !== ct)
        problems.push(`${p} must carry an Override typed ${ct} (got ${got ?? 'only a Default'})`);
    }
  }

  const bytes = mergedBuf.size ?? mergedBuf.length ?? 0;
  console.log(`[${label}] parts=${paths.length} relsFiles=${relFiles} slides=${sldTags.length} size=${(bytes / 1048576).toFixed(2)}MB`);
  if (problems.length) {
    console.log(`  FAIL (${problems.length}):`);
    for (const pr of problems.slice(0, 12)) console.log('   -', pr);
    process.exitCode = 1;
  } else {
    console.log('  PASS: all rels resolve, content types complete, ids unique');
  }
}

// ---- generate fixtures (3 decks, 6 slides total, 2 media images) ----
const deckA = await makeDeck('business', [
  { title: 'Deck A — Cover', bullets: null, image: true },
  { title: 'Deck A — Points', bullets: ['first point', 'second point', 'third point'] },
  { title: 'Deck A — Closing', bullets: ['thanks'] },
]);
const deckB = await makeDeck('dark', [
  { title: 'Deck B — Section', bg: '1E293B', image: true },
  { title: 'Deck B — Data', bullets: ['42 percent', 'up from 31'] },
]);
const deckC = await makeDeck('light', [{ title: 'Deck C — Single', bullets: ['only slide'] }]);

const inputs = [
  { name: 'deck-a.pptx', size: deckA.length, buffer: deckA },
  { name: 'deck-b.pptx', size: deckB.length, buffer: deckB },
  { name: 'deck-c.pptx', size: deckC.length, buffer: deckC },
];
const perFile = [];
for (const f of inputs) perFile.push(await countSlides(f.buffer));
console.log('inputs:', inputs.map((f, i) => `${f.name}(${perFile[i]}slides, ${(f.size / 1024).toFixed(1)}KB)`).join(' + '));
const expected = perFile.reduce((a, b) => a + b, 0);

// ---- preserve mode ----
const p1 = await mergePptxFiles(inputs, { mode: 'preserve' });
await verify(p1.blob, expected, 'preserve');

// ---- unify mode ----
const p2 = await mergePptxFiles(inputs, { mode: 'unify', outputName: 'test.pptx' });
await verify(p2.blob, expected, 'unify');
console.log(`unify result: mismatch=${p2.sizeMismatch} name=${p2.fileName}`);

// ---- error branches ----
const fakeOle = Buffer.concat([Buffer.from([0xd0, 0xcf, 0x11, 0xe0]), Buffer.alloc(64)]);
const cases = [
  ['encrypted', 'encrypted', [{ name: 'locked.pptx', size: 68, buffer: fakeOle }]],
  ['ppt', 'ppt', [{ name: 'old.ppt', size: 68, buffer: fakeOle }]],
  [
    'corrupt',
    'corrupt',
    [{ name: 'junk.pptx', size: 16, buffer: Buffer.from('not a zip at all') }],
  ],
];
for (const [label, code, extra] of cases) {
  try {
    await mergePptxFiles([inputs[0], ...extra]);
    console.log(`[${label}] FAIL: no error`);
    process.exitCode = 1;
  } catch (e) {
    const ok = e instanceof MergeError && e.code === code;
    console.log(`[${label}] ${ok ? 'PASS' : 'FAIL ' + e.message}`);
    if (!ok) process.exitCode = 1;
  }
}
expectNeedTwo: {
  try {
    await mergePptxFiles([inputs[0]]);
    console.log('[need_two] FAIL: no error');
    process.exitCode = 1;
  } catch (e) {
    const ok = e instanceof MergeError && e.code === 'need_two';
    console.log(`[need_two] ${ok ? 'PASS' : 'FAIL ' + e.message}`);
    if (!ok) process.exitCode = 1;
  }
}

console.log(process.exitCode ? '\nRESULT: FAIL' : '\nRESULT: ALL PASS');
