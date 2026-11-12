/**
 * Browser-local PPTX merger.
 *
 * A .pptx is an OPC package: parts + rels + [Content_Types].xml. To merge
 * decks we use deck #1 as the base package and, for every other deck,
 * copy its slides plus everything they transitively reference (layouts,
 * masters, themes, media, notes…) under collision-free names, rewrite
 * the copied .rels targets, append to the base sldIdLst / presentation
 * rels, and merge [Content_Types].xml.
 *
 * Safety rules:
 * - Slide XML is copied byte-identical — we never touch slide content,
 *   so text, animations and placeholders arrive exactly as authored.
 * - r:id values inside a part stay valid because each part keeps its
 *   own rels (only Target paths inside those rels are rewritten).
 * - The file never leaves the browser; there is no upload path here.
 *
 * DOM-free by design (regex surgery on XML strings) so the exact same
 * code runs in node for structural tests (test/merge.test.mjs).
 */

import JSZip from 'jszip';

export type MergeMode = 'preserve' | 'unify';

export type MergeProgress = (info: {
  step: 'read' | 'parse' | 'merge' | 'package';
  current: number;
  total: number;
  detail?: string;
}) => void;

export type MergeInput = {
  name: string;
  size: number;
  buffer: ArrayBuffer | Uint8Array | Buffer;
};

export type MergeResult = {
  blob: Blob | Buffer;
  fileName: string;
  slideCount: number;
  originalBytes: number;
  mergedBytes: number;
  sizeMismatch: boolean;
};

export type MergeErrorCode =
  | 'need_two'
  | 'too_big'
  | 'total_too_big'
  | 'ppt'
  | 'encrypted'
  | 'corrupt'
  | 'no_slides';

export class MergeError extends Error {
  code: MergeErrorCode;
  file?: string;
  constructor(code: MergeErrorCode, file?: string) {
    super(`merge:${code}${file ? `:${file}` : ''}`);
    this.code = code;
    this.file = file;
  }
}

const REL_SLIDE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const REL_SLIDE_LAYOUT = `${REL_SLIDE}Layout`;
const REL_SLIDE_MASTER = `${REL_SLIDE}Master`;

export const MERGE_MAX_FILES = 20;
export const MERGE_MAX_FILE_BYTES = 100 * 1024 * 1024;
export const MERGE_MAX_TOTAL_BYTES = 300 * 1024 * 1024;

// ---------- path helpers ----------

function dirOf(p: string): string {
  const i = p.lastIndexOf('/');
  return i < 0 ? '' : p.slice(0, i);
}
function baseOf(p: string): string {
  return p.slice(p.lastIndexOf('/') + 1);
}
function extOf(p: string): string {
  const b = baseOf(p);
  const i = b.lastIndexOf('.');
  return i < 0 ? '' : b.slice(i + 1).toLowerCase();
}
function relsPathOf(p: string): string {
  return `${dirOf(p)}/_rels/${baseOf(p)}.rels`;
}

/** Resolve a rel Target (may be relative or lead with /) against the part dir. */
function resolveRel(fromDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const segs = (fromDir ? fromDir.split('/') : []).concat(target.split('/'));
  const out: string[] = [];
  for (const s of segs) {
    if (s === '..') out.pop();
    else if (s === '' || s === '.') continue;
    else out.push(s);
  }
  return out.join('/');
}

/** Relative path from a part's directory to another part. */
function relTarget(fromDir: string, toPath: string): string {
  const from = fromDir ? fromDir.split('/') : [];
  const to = toPath.split('/');
  let i = 0;
  const lim = Math.min(from.length, to.length - 1);
  while (i < lim && from[i] === to[i]) i++;
  const ups = from.length - i;
  return (
    (ups ? Array<string>(ups).fill('..') : []).concat(to.slice(i)).join('/') ||
    baseOf(toPath)
  );
}

// ---------- XML helpers (regex surgery, order-independent attrs) ----------

function attrOf(tag: string, name: string): string | null {
  const m = tag.match(new RegExp(`\\b${name}="([^"]*)"`));
  return m ? m[1] : null;
}

type Rel = {
  id: string;
  type: string;
  target: string;
  external: boolean;
};

const REL_TAG = /<Relationship\b[^>]*?\/?>/g;

function parseRels(xml: string): Rel[] {
  const rels: Rel[] = [];
  for (const m of xml.matchAll(REL_TAG)) {
    const tag = m[0];
    const id = attrOf(tag, 'Id');
    const type = attrOf(tag, 'Type');
    const target = attrOf(tag, 'Target');
    if (id && type && target) {
      rels.push({
        id,
        type,
        target,
        external: attrOf(tag, 'TargetMode') === 'External',
      });
    }
  }
  return rels;
}

function cSldName(xml: string | null): string {
  if (!xml) return '';
  const m = xml.match(/<p:cSld\b[^>]*\bname="([^"]*)"/);
  return m ? m[1] : '';
}

// ---------- package model ----------

type Package = {
  name: string;
  zip: JSZip;
  index: Map<string, string>; // lowerPath → actual path
  presXml: string;
  presRelsXml: string;
  presRels: Rel[];
  slideTargets: string[]; // actual paths, sldIdLst order
  sldIds: number[];
  cx: number;
  cy: number;
  notesMasterTargets: string[];
  ctXml: string;
  ctDefaults: Map<string, string>; // ext → type
  ctOverrides: Map<string, string>; // lowerPath (no slash) → type
};

function fail(code: MergeErrorCode, file?: string): never {
  throw new MergeError(code, file);
}

async function readTextAsync(
  zip: JSZip,
  index: Map<string, string>,
  path: string
): Promise<string | null> {
  const actual = index.get(path.toLowerCase()) ?? path;
  const entry = zip.file(actual);
  if (!entry) return null;
  return entry.async('string');
}

async function parsePackage(name: string, zip: JSZip): Promise<Package> {
  const index = new Map<string, string>();
  for (const p of Object.keys(zip.files)) {
    if (!zip.files[p].dir) index.set(p.toLowerCase(), p);
  }

  const presXml = await readTextAsync(zip, index, 'ppt/presentation.xml');
  const presRelsXml = await readTextAsync(
    zip,
    index,
    'ppt/_rels/presentation.xml.rels'
  );
  if (!presXml || !presRelsXml) fail('corrupt', name);

  const presRels = parseRels(presRelsXml);
  const relById = new Map(presRels.map((r) => [r.id, r]));

  // Slides in presentation order.
  const slideTargets: string[] = [];
  const sldIds: number[] = [];
  for (const m of presXml.matchAll(/<p:sldId\b[^>]*?\/?>/g)) {
    const tag = m[0];
    const rid = attrOf(tag, 'r:id');
    const id = Number(attrOf(tag, 'id') ?? '0');
    if (!rid) continue;
    const rel = relById.get(rid);
    if (!rel || rel.external) fail('corrupt', name);
    const resolved = resolveRel('ppt', rel.target);
    const actual = index.get(resolved.toLowerCase());
    if (!actual) fail('corrupt', name);
    slideTargets.push(actual);
    sldIds.push(id);
  }

  // Slide size.
  const szm = presXml.match(/<p:sldSz\b[^>]*>/);
  const cx = Number(szm ? (attrOf(szm[0], 'cx') ?? '0') : '0');
  const cy = Number(szm ? (attrOf(szm[0], 'cy') ?? '0') : '0');

  const notesMasterTargets = presRels
    .filter((r) => !r.external && r.type.endsWith('/notesMaster'))
    .map((r) => index.get(resolveRel('ppt', r.target).toLowerCase()))
    .filter(Boolean) as string[];

  const ctXml = await readTextAsync(zip, index, '[Content_Types].xml');
  if (!ctXml) fail('corrupt', name);

  const ctDefaults = new Map<string, string>();
  for (const m of ctXml.matchAll(/<Default\b[^>]*?\/?>/g)) {
    const ext = (attrOf(m[0], 'Extension') ?? '').toLowerCase();
    const ct = attrOf(m[0], 'ContentType');
    if (ext && ct) ctDefaults.set(ext, ct);
  }
  const ctOverrides = new Map<string, string>();
  for (const m of ctXml.matchAll(/<Override\b[^>]*?\/?>/g)) {
    const pn = (attrOf(m[0], 'PartName') ?? '').replace(/^\//, '');
    const ct = attrOf(m[0], 'ContentType');
    if (pn && ct) ctOverrides.set(pn.toLowerCase(), ct);
  }

  return {
    name,
    zip,
    index,
    presXml,
    presRelsXml,
    presRels,
    slideTargets,
    sldIds,
    cx,
    cy,
    notesMasterTargets,
    ctXml,
    ctDefaults,
    ctOverrides,
  };
}

// ---------- merge ----------

export async function mergePptxFiles(
  files: MergeInput[],
  opts: { mode?: MergeMode; outputName?: string } = {},
  onProgress?: MergeProgress
): Promise<MergeResult> {
  if (files.length < 2) fail('need_two');

  let total = 0;
  for (const f of files) {
    if (f.size > MERGE_MAX_FILE_BYTES) fail('too_big', f.name);
    total += f.size;
  }
  if (total > MERGE_MAX_TOTAL_BYTES) fail('total_too_big');

  // ---- read + magic-byte checks ----
  const zips: JSZip[] = [];
  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    onProgress?.({
      step: 'read',
      current: i + 1,
      total: files.length,
      detail: f.name,
    });
    const u8 =
      f.buffer instanceof Uint8Array
        ? f.buffer
        : new Uint8Array(f.buffer as ArrayBuffer);
    if (
      u8.length > 3 &&
      u8[0] === 0xd0 &&
      u8[1] === 0xcf &&
      u8[2] === 0x11 &&
      u8[3] === 0xe0
    ) {
      // OLE container: encrypted pptx or legacy binary .ppt.
      if (/\.pptx$/i.test(f.name)) fail('encrypted', f.name);
      else fail('ppt', f.name);
    }
    try {
      zips.push(await JSZip.loadAsync(u8));
    } catch {
      fail('corrupt', f.name);
    }
  }

  // ---- parse ----
  const packs: Package[] = [];
  for (let i = 0; i < zips.length; i++) {
    onProgress?.({
      step: 'parse',
      current: i + 1,
      total: zips.length,
      detail: files[i].name,
    });
    const pack = await parsePackage(files[i].name, zips[i]);
    if (pack.slideTargets.length === 0) fail('no_slides', files[i].name);
    packs.push(pack);
  }

  const base = packs[0];
  const baseZip = zips[0];
  const mode = opts.mode ?? 'preserve';
  const sizeMismatch = packs.some((p) => p.cx !== base.cx || p.cy !== base.cy);

  // ---- content-type accumulators (seeded from base) ----
  const ctDefaults = new Map(base.ctDefaults);
  const ctOverrides = new Map(base.ctOverrides);
  const ctAdditions: string[] = [];

  // ---- rId / sldId counters ----
  let nextRid = 1;
  for (const r of base.presRels) {
    const n = Number(r.id.replace(/\D/g, ''));
    if (Number.isFinite(n) && n >= nextRid) nextRid = n + 1;
  }
  let nextSldId = 256;
  for (const id of base.sldIds) if (id >= nextSldId) nextSldId = id + 1;

  // ---- unify mode: pick base layouts to retarget to ----
  let baseLayouts: { path: string; name: string }[] = [];
  if (mode === 'unify') {
    const masterRel = base.presRels.find(
      (r) => !r.external && r.type === REL_SLIDE_MASTER
    );
    if (masterRel) {
      const mPath = base.index.get(
        resolveRel('ppt', masterRel.target).toLowerCase()
      );
      if (mPath) {
        const mRelsXml = await readTextAsync(
          baseZip,
          base.index,
          relsPathOf(mPath)
        );
        const layouts = parseRels(mRelsXml ?? '')
          .filter((r) => !r.external && r.type === REL_SLIDE_LAYOUT)
          .map((r) =>
            base.index.get(resolveRel(dirOf(mPath), r.target).toLowerCase())
          )
          .filter(Boolean) as string[];
        for (const lp of layouts) {
          baseLayouts.push({
            path: lp,
            name: cSldName(await readTextAsync(baseZip, base.index, lp)),
          });
        }
      }
    }
  }
  const fallbackLayout = baseLayouts[0]?.path ?? null;

  const usedPaths = new Set(
    Object.keys(baseZip.files).filter((p) => !baseZip.files[p].dir)
  );
  const appendRels: string[] = [];
  const appendSldIds: string[] = [];

  // ---- import every non-base deck ----
  for (let fi = 1; fi < packs.length; fi++) {
    const pack = packs[fi];
    onProgress?.({
      step: 'merge',
      current: fi,
      total: packs.length - 1,
      detail: pack.name,
    });
    const prefix = `f${fi}_`;

    // BFS over rels from slides (+ notes masters). Keyed by lower path.
    const rename = new Map<string, string>(); // lowerOld → new
    const queue: string[] = [...pack.slideTargets, ...pack.notesMasterTargets];
    const seen = new Set(queue.map((p) => p.toLowerCase()));

    while (queue.length) {
      const p = queue.shift()!;
      const lp = p.toLowerCase();
      if (rename.has(lp)) continue;
      rename.set(lp, `${dirOf(p)}/${prefix}${baseOf(p)}`);

      const relsXml = await readTextAsync(pack.zip, pack.index, relsPathOf(p));
      if (!relsXml) continue;
      for (const rel of parseRels(relsXml)) {
        if (rel.external) continue;
        if (mode === 'unify' && rel.type === REL_SLIDE_LAYOUT) continue;
        const tp = resolveRel(dirOf(p), rel.target);
        const actual = pack.index.get(tp.toLowerCase());
        if (!actual) continue; // dangling rel — leave it behind
        if (!seen.has(actual.toLowerCase())) {
          seen.add(actual.toLowerCase());
          queue.push(actual);
        }
      }
    }

    // Unify mode: per-slide layout choice (match by layout name).
    const unifyChoice = new Map<string, string>(); // slide lower → base layout path
    if (mode === 'unify' && fallbackLayout) {
      for (const sp of pack.slideTargets) {
        const sRelsXml = await readTextAsync(
          pack.zip,
          pack.index,
          relsPathOf(sp)
        );
        const layoutRel = parseRels(sRelsXml ?? '').find(
          (r) => !r.external && r.type === REL_SLIDE_LAYOUT
        );
        let chosen = fallbackLayout;
        if (layoutRel) {
          const origLayout = pack.index.get(
            resolveRel(dirOf(sp), layoutRel.target).toLowerCase()
          );
          const origName = origLayout
            ? cSldName(await readTextAsync(pack.zip, pack.index, origLayout))
            : '';
          chosen =
            baseLayouts.find((l) => l.name && l.name === origName)?.path ??
            fallbackLayout;
        }
        unifyChoice.set(sp.toLowerCase(), chosen);
      }
    }

    // Copy parts + rewritten rels.
    for (const [lp, npRaw] of rename) {
      const actual = pack.index.get(lp)!;
      let np = npRaw;
      let k = 2;
      while (usedPaths.has(np)) {
        np = npRaw.replace(/(\.[^./]+)$/, `_${k++}$1`);
      }
      usedPaths.add(np);

      const entry = pack.zip.file(actual);
      if (!entry) continue;
      baseZip.file(np, await entry.async('uint8array'));

      const relsXml = await readTextAsync(
        pack.zip,
        pack.index,
        relsPathOf(actual)
      );
      if (relsXml) {
        const rewritten = relsXml.replace(REL_TAG, (tag) => {
          const target = attrOf(tag, 'Target');
          const type = attrOf(tag, 'Type');
          if (!target || attrOf(tag, 'TargetMode') === 'External') return tag;
          if (type && relTypeShort(type) === 'slideLayout') {
            if (mode === 'unify') {
              const chosen = unifyChoice.get(lp) ?? fallbackLayout;
              if (chosen) {
                return tag.replace(
                  /Target="[^"]*"/,
                  `Target="${relTarget(dirOf(np), chosen)}"`
                );
              }
            }
          }
          const resolved = resolveRel(dirOf(actual), target);
          const mapped = rename.get(resolved.toLowerCase());
          if (!mapped) return tag; // external-ish or dangling — keep
          return tag.replace(
            /Target="[^"]*"/,
            `Target="${relTarget(dirOf(np), mapped)}"`
          );
        });
        baseZip.file(relsPathOf(np), rewritten);
      }

      // Content types. An Override in the source package must be carried to
      // the new path even when a Default covers the extension — packages
      // that ship Default Extension="xml" (pptxgenjs, some WPS decks) would
      // otherwise retype prefixed slides as application/xml and PowerPoint
      // repairs them away.
      const fromOverride = pack.ctOverrides.get(actual.toLowerCase());
      if (fromOverride) {
        if (!ctOverrides.has(np.toLowerCase())) {
          ctOverrides.set(np.toLowerCase(), fromOverride);
          ctAdditions.push(
            `<Override PartName="/${np}" ContentType="${fromOverride}"/>`
          );
        }
      } else {
        const ext = extOf(np);
        if (!ctDefaults.has(ext)) {
          const fromDefault = pack.ctDefaults.get(ext);
          const ct = fromDefault ?? 'application/octet-stream';
          ctDefaults.set(ext, ct);
          ctAdditions.push(`<Default Extension="${ext}" ContentType="${ct}"/>`);
        }
      }
    }

    // Append slides in order.
    for (const sp of pack.slideTargets) {
      const np = rename.get(sp.toLowerCase());
      if (!np) continue;
      appendRels.push(
        `<Relationship Id="rId${nextRid}" Type="${REL_SLIDE}" Target="${relTarget('ppt', np)}"/>`
      );
      appendSldIds.push(`<p:sldId id="${nextSldId}" r:id="rId${nextRid}"/>`);
      nextRid++;
      nextSldId++;
    }
  }

  // ---- commit presentation edits ----
  if (!base.presXml.includes('</p:sldIdLst>')) fail('corrupt', base.name);
  baseZip.file(
    'ppt/presentation.xml',
    base.presXml.replace(
      '</p:sldIdLst>',
      `${appendSldIds.join('')}</p:sldIdLst>`
    )
  );
  baseZip.file(
    'ppt/_rels/presentation.xml.rels',
    base.presRelsXml.replace(
      '</Relationships>',
      `${appendRels.join('')}</Relationships>`
    )
  );
  if (ctAdditions.length) {
    baseZip.file(
      '[Content_Types].xml',
      base.ctXml.replace('</Types>', `${ctAdditions.join('')}</Types>`)
    );
  }

  const slideCount =
    base.slideTargets.length +
    packs.slice(1).reduce((n, p) => n + p.slideTargets.length, 0);

  // ---- package ----
  onProgress?.({ step: 'package', current: 0, total: 100 });
  // Node 18+ has a global Blob, so detect Node properly (tests run there).
  const inNode = typeof process !== 'undefined' && !!process.versions?.node;
  const out: Blob | Buffer = await baseZip.generateAsync(
    {
      type: (inNode ? 'nodebuffer' : 'blob') as any,
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    },
    (meta) =>
      onProgress?.({
        step: 'package',
        current: Math.round(meta.percent),
        total: 100,
      })
  );

  const mergedBytes = inNode ? (out as Buffer).length : (out as Blob).size;

  return {
    blob: out,
    fileName: opts.outputName || 'merged.pptx',
    slideCount,
    originalBytes: total,
    mergedBytes,
    sizeMismatch,
  };
}

function relTypeShort(type: string): string {
  const i = type.lastIndexOf('/');
  return i < 0 ? type : type.slice(i + 1);
}

export function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(0)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}
