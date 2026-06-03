'use strict';

// ocrEngine — Tesseract OCR, 100% offline. Language data is BUNDLED locally
// (assets/lang/eng.traineddata.gz) and tesseract.js is pointed at that path so
// it NEVER fetches anything at runtime. Works air-gapped.

const path = require('path');
const fs = require('fs');

let createWorkerFn = null;

function resolveLangPath() {
  // In dev: <project>/assets/lang. In a packaged asar build the lang folder is
  // unpacked (see asarUnpack), so resolve relative to app resources.
  const candidates = [
    path.join(__dirname, '..', '..', 'assets', 'lang'),
    process.resourcesPath ? path.join(process.resourcesPath, 'assets', 'lang') : null,
    process.resourcesPath
      ? path.join(process.resourcesPath, 'app.asar.unpacked', 'assets', 'lang')
      : null,
  ].filter(Boolean);
  for (const c of candidates) {
    if (fs.existsSync(path.join(c, 'eng.traineddata.gz')) ||
        fs.existsSync(path.join(c, 'eng.traineddata'))) {
      return c;
    }
  }
  return candidates[0];
}

async function getCreateWorker() {
  if (!createWorkerFn) {
    const tess = require('tesseract.js');
    createWorkerFn = tess.createWorker;
  }
  return createWorkerFn;
}

/**
 * Run OCR on a single image (PNG/JPG bytes). Returns recognized text + word
 * boxes. lang defaults to 'eng' (the bundled model). onProgress(0..1).
 */
async function ocrImage(imageBytes, { lang = 'eng', onProgress } = {}) {
  const createWorker = await getCreateWorker();
  const langPath = resolveLangPath();
  const worker = await createWorker(lang, 1, {
    langPath,                 // local dir holding eng.traineddata.gz
    gzip: true,               // our bundled file is gzipped
    cacheMethod: 'none',      // never write a cache that could imply a fetch
    logger: (m) => {
      if (onProgress && m.status === 'recognizing text') onProgress(m.progress);
    },
  });
  try {
    const { data } = await worker.recognize(Buffer.from(imageBytes));
    return { text: data.text, words: data.words || [], confidence: data.confidence };
  } finally {
    await worker.terminate();
  }
}

/**
 * OCR a sequence of page images (already rasterized by the renderer) and build
 * a single searchable text result, plus per-page word boxes so the caller can
 * stamp an invisible text layer with pdf-lib.
 */
async function ocrPages(pageImages, { lang = 'eng', onProgress } = {}) {
  const createWorker = await getCreateWorker();
  const langPath = resolveLangPath();
  const worker = await createWorker(lang, 1, {
    langPath,
    gzip: true,
    cacheMethod: 'none',
  });
  const results = [];
  try {
    for (let i = 0; i < pageImages.length; i++) {
      const { data } = await worker.recognize(Buffer.from(pageImages[i].bytes));
      results.push({
        text: data.text,
        words: data.words || [],
        width: pageImages[i].width,
        height: pageImages[i].height,
      });
      if (onProgress) onProgress((i + 1) / pageImages.length);
    }
  } finally {
    await worker.terminate();
  }
  return results;
}

module.exports = { ocrImage, ocrPages, resolveLangPath };
