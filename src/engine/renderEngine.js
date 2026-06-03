'use strict';

// renderEngine — rasterizes PDF pages with pdfjs-dist (offline) and produces
// PNG/JPG via Electron's nativeImage (no native canvas binary needed, so it
// builds cleanly on the CI Windows runner without node-gyp). All on-device.

const { nativeImage } = require('electron');

let pdfjsPromise = null;
async function getPdfjs() {
  if (!pdfjsPromise) {
    // pdfjs-dist v4 ships ESM; load the legacy build for Node/Electron-main.
    pdfjsPromise = import('pdfjs-dist/legacy/build/pdf.mjs');
  }
  return pdfjsPromise;
}

/**
 * A minimal canvas factory backed by @napi-rs/canvas is overkill and pulls a
 * native binary. Instead we render to an RGBA pixel buffer using pdfjs' ability
 * to draw into a provided canvas-like object. pdfjs requires a real 2D context
 * for full fidelity, so we use the "OffscreenCanvas" path when available, else
 * fall back to producing an image of the page via the canvas the renderer owns.
 *
 * In Electron's main process there is no DOM canvas. To keep zero native deps
 * we therefore do rasterization in a hidden renderer (see main.js
 * rasterizeInRenderer). renderEngine only handles text + image embedding that
 * does NOT need a canvas.
 */

/** Extract all text from a PDF, page by page. Returns {pages:[string], text}. */
async function extractText(bytes) {
  const pdfjs = await getPdfjs();
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const doc = await pdfjs.getDocument({
    data,
    isEvalSupported: false,
    useSystemFonts: true,
  }).promise;
  const pages = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const content = await page.getTextContent();
    let last = null;
    let line = '';
    const lines = [];
    for (const item of content.items) {
      if (!('str' in item)) continue;
      const y = item.transform[5];
      if (last != null && Math.abs(y - last) > 2) {
        lines.push(line.trimEnd());
        line = '';
      }
      line += item.str + (item.hasEOL ? '\n' : ' ');
      last = y;
    }
    if (line.trim()) lines.push(line.trimEnd());
    pages.push(lines.join('\n'));
  }
  await doc.destroy();
  return { pages, text: pages.join('\n\n') };
}

/** Page count + per-page viewport sizes at scale 1. */
async function pageGeometry(bytes) {
  const pdfjs = await getPdfjs();
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const doc = await pdfjs.getDocument({ data, isEvalSupported: false }).promise;
  const sizes = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const vp = page.getViewport({ scale: 1 });
    sizes.push({ width: vp.width, height: vp.height });
  }
  const n = doc.numPages;
  await doc.destroy();
  return { pageCount: n, sizes };
}

// Convert an RGBA pixel buffer (from the renderer) into encoded PNG/JPG bytes.
function encodeRGBA(rgba, width, height, format = 'png', quality = 92) {
  const img = nativeImage.createFromBitmap(Buffer.from(rgba), { width, height });
  if (format === 'jpg' || format === 'jpeg') {
    return img.toJPEG(quality);
  }
  return img.toPNG();
}

module.exports = { getPdfjs, extractText, pageGeometry, encodeRGBA };
