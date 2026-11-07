/**
 * Browser-local PPTX compressor.
 *
 * A .pptx is a ZIP; the bulk of a heavy deck is almost always
 * ppt/media/* raster images. We unzip with JSZip, re-encode oversized
 * images via canvas (downscale > MAX_EDGE px, JPEG re-quantized), and
 * re-zip. The file never leaves the browser — there is no upload path
 * in this code at all.
 *
 * Safety rules:
 * - Output format always matches the entry extension (.jpeg→JPEG,
 *   .png→PNG) so [Content_Types].xml stays truthful — no renames, no
 *   rels surgery.
 * - A re-encoded image only replaces the original when it is actually
 *   smaller; otherwise the original bytes are kept.
 * - XML, fonts and everything else are copied byte-identical.
 */

import JSZip from 'jszip';

export type CompressProgress = (info: {
  current: number;
  total: number;
  name: string;
}) => void;

export type CompressResult = {
  blob: Blob;
  fileName: string;
  originalSize: number;
  compressedSize: number;
  imagesOptimized: number;
  imagesScanned: number;
};

/**
 * Compression levels — user-facing presets. Each maps to a
 * long-edge ceiling and JPEG quality; PNG stays format-preserving and
 * only gets the downscale.
 */
export type CompressLevel = 'light' | 'balanced' | 'max';

export const COMPRESS_LEVELS: Record<
  CompressLevel,
  { maxEdge: number; quality: number }
> = {
  // Print-grade fidelity — near-lossless on paper.
  light: { maxEdge: 2900, quality: 0.85 },
  // Screen presentations — invisible loss at 1080p (default).
  balanced: { maxEdge: 1920, quality: 0.72 },
  // Smallest sensible file — email caps and slow uploads.
  max: { maxEdge: 1400, quality: 0.6 },
};

const MIN_SAVING_BYTES = 1024; // ignore sub-KB wins
const IMAGE_EXT = /\.(jpe?g|png)$/i;

function canvasToBlob(
  canvas: HTMLCanvasElement,
  mime: string,
  quality?: number
): Promise<Blob | null> {
  return new Promise((resolve) => canvas.toBlob(resolve, mime, quality));
}

async function decodeImage(blob: Blob): Promise<HTMLImageElement | null> {
  const url = URL.createObjectURL(blob);
  try {
    const img = new Image();
    img.decoding = 'async';
    await new Promise<void>((resolve, reject) => {
      img.onload = () => resolve();
      img.onerror = () => reject(new Error('decode failed'));
      img.src = url;
    });
    return img;
  } finally {
    // Revoke async — Safari revokes synchronously before draw sometimes.
    setTimeout(() => URL.revokeObjectURL(url), 5_000);
  }
}

/** Re-encode one image: downscale beyond the level's edge, same format. */
async function reencodeImage(
  blob: Blob,
  ext: string,
  maxEdge: number,
  quality: number
): Promise<Blob | null> {
  const img = await decodeImage(blob);
  if (!img || !img.width || !img.height) return null;

  const scale = Math.min(1, maxEdge / Math.max(img.width, img.height));
  const w = Math.max(1, Math.round(img.width * scale));
  const h = Math.max(1, Math.round(img.height * scale));

  const canvas = document.createElement('canvas');
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(img, 0, 0, w, h);

  const isJpg = /\.jpe?g$/i.test(ext);
  // PNG stays PNG (format-preserving rule); JPEG gets re-quantized.
  return canvasToBlob(
    canvas,
    isJpg ? 'image/jpeg' : 'image/png',
    isJpg ? quality : undefined
  );
}

export async function compressPptx(
  file: File,
  opts: { level?: CompressLevel } = {},
  onProgress?: CompressProgress
): Promise<CompressResult> {
  const { maxEdge, quality } = COMPRESS_LEVELS[opts.level ?? 'balanced'];
  const zip = await JSZip.loadAsync(file);

  // Collect media entries (order stable for progress reporting).
  const entries = Object.keys(zip.files)
    .filter((p) => p.startsWith('ppt/media/') && IMAGE_EXT.test(p))
    .sort();

  let imagesOptimized = 0;
  for (let i = 0; i < entries.length; i++) {
    const path = entries[i];
    onProgress?.({ current: i + 1, total: entries.length, name: path });

    const entry = zip.files[path];
    if (entry.dir) continue;

    const original = await entry.async('blob');
    // Tiny images can't pay for a re-encode.
    if (original.size < 50 * 1024) continue;

    try {
      const reencoded = await reencodeImage(original, path, maxEdge, quality);
      if (reencoded && reencoded.size + MIN_SAVING_BYTES < original.size) {
        zip.file(path, reencoded);
        imagesOptimized++;
      }
    } catch {
      // A bad image must never break the whole deck — keep original.
    }
  }

  const blob = await zip.generateAsync({
    type: 'blob',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  // Never hand back something bigger than the user sent.
  const finalBlob = blob.size < file.size ? blob : file;

  return {
    blob: finalBlob,
    fileName: file.name,
    originalSize: file.size,
    compressedSize: finalBlob.size,
    imagesOptimized,
    imagesScanned: entries.length,
  };
}

export function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(0)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}
