'use strict';

// raster-preload — runs inside the hidden raster window. It loads pdf.js (the
// real browser build) against the local file, renders pages onto an offscreen
// canvas, and hands back RGBA pixels. Pure on-device: pdf.js worker is the
// local bundled .mjs; no network. Exposes window.__rasterize.

const { contextBridge } = require('electron');
const path = require('path');

// Resolve the local pdf.js ESM build + worker. We load them as file:// URLs
// (allowed by the network lockdown) via dynamic import in the page context.
const pdfBuildDir = path.join(
  __dirname, '..', '..', 'node_modules', 'pdfjs-dist', 'build'
);

function toFileUrl(p) {
  let u = p.replace(/\\/g, '/');
  if (!u.startsWith('/')) u = '/' + u;
  return 'file://' + encodeURI(u);
}

contextBridge.exposeInMainWorld('__folioRaster', {
  pdfUrl: toFileUrl(path.join(pdfBuildDir, 'pdf.mjs')),
  workerUrl: toFileUrl(path.join(pdfBuildDir, 'pdf.worker.mjs')),
});
