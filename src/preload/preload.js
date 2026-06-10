'use strict';

// preload — the ONLY bridge between the editorial renderer UI and the privileged
// main process. contextIsolation is on, nodeIntegration is off; the renderer
// sees nothing but this typed `window.folio` surface.

const { contextBridge, ipcRenderer } = require('electron');

const folio = {
  // ── import / export ────────────────────────────────────────────────────
  openPdf: (multi = false) => ipcRenderer.invoke('dlg:openPdf', multi),
  openImages: () => ipcRenderer.invoke('dlg:openImages'),
  openOffice: (kind) => ipcRenderer.invoke('dlg:openOffice', kind),
  openXml: () => ipcRenderer.invoke('dlg:openXml'),
  openXps: () => ipcRenderer.invoke('dlg:openXps'),
  openAny: () => ipcRenderer.invoke('dlg:openAny'),
  saveBytes: (bytes, defaultName, kind) =>
    ipcRenderer.invoke('save:bytes', { bytes, defaultName, kind }),
  saveMulti: (files) => ipcRenderer.invoke('save:multi', { files }),

  // ── organize ───────────────────────────────────────────────────────────
  merge: (buffers) => ipcRenderer.invoke('pdf:merge', buffers),
  split: (bytes, baseName) => ipcRenderer.invoke('pdf:split', { bytes, baseName }),
  extract: (bytes, range) => ipcRenderer.invoke('pdf:extract', { bytes, range }),
  deletePages: (bytes, range) => ipcRenderer.invoke('pdf:delete', { bytes, range }),
  rotate: (bytes, angle, range) => ipcRenderer.invoke('pdf:rotate', { bytes, angle, range }),
  compress: (bytes, opts = {}) => ipcRenderer.invoke('pdf:compress', { bytes, ...opts }),
  info: (bytes) => ipcRenderer.invoke('pdf:info', bytes),

  // ── images ─────────────────────────────────────────────────────────────
  pdfToImages: (bytes, format = 'png', scale = 2) =>
    ipcRenderer.invoke('pdf:toImages', { bytes, format, scale }),
  imagesToPdf: (images, pageSize = 'fit') =>
    ipcRenderer.invoke('pdf:fromImages', { images, pageSize }),

  // ── viewer ─────────────────────────────────────────────────────────────
  renderPage: (bytes, pageIndex, scale = 1.5) =>
    ipcRenderer.invoke('pdf:renderPage', { bytes, pageIndex, scale }),
  geometry: (bytes) => ipcRenderer.invoke('pdf:geometry', bytes),

  // ── secure & sign ────────────────────────────────────────────────────────
  watermark: (bytes, text, opacity = 0.18) =>
    ipcRenderer.invoke('pdf:watermark', { bytes, text, opacity }),
  pageNumbers: (bytes, format, position) =>
    ipcRenderer.invoke('pdf:pageNumbers', { bytes, format, position }),
  flatten: (bytes) => ipcRenderer.invoke('pdf:flatten', bytes),
  sign: (bytes, signaturePng, opts = {}) =>
    ipcRenderer.invoke('pdf:sign', { bytes, signaturePng, ...opts }),

  // ── scan & read ──────────────────────────────────────────────────────────
  extractText: (bytes) => ipcRenderer.invoke('pdf:extractText', bytes),
  ocrImage: (bytes) => ipcRenderer.invoke('ocr:image', { bytes }),
  ocrPdf: (bytes, scale = 2) => ipcRenderer.invoke('ocr:pdf', { bytes, scale }),
  onOcrProgress: (cb) => {
    const listener = (_e, p) => cb(p);
    ipcRenderer.on('ocr:progress', listener);
    return () => ipcRenderer.removeListener('ocr:progress', listener);
  },

  // ── convert ────────────────────────────────────────────────────────────
  pdfToWord: (bytes) => ipcRenderer.invoke('conv:pdfToWord', bytes),
  pdfToExcel: (bytes) => ipcRenderer.invoke('conv:pdfToExcel', bytes),
  pdfToPpt: (bytes) => ipcRenderer.invoke('conv:pdfToPpt', bytes),
  wordToPdf: (bytes) => ipcRenderer.invoke('conv:wordToPdf', bytes),
  excelToPdf: (bytes) => ipcRenderer.invoke('conv:excelToPdf', bytes),

  // ── convert — XML / XPS ──────────────────────────────────────────────────
  xmlToPdf: (bytes) => ipcRenderer.invoke('conv:xmlToPdf', bytes),
  pdfToXml: (bytes) => ipcRenderer.invoke('conv:pdfToXml', bytes),
  viewXml: (bytes) => ipcRenderer.invoke('conv:viewXml', bytes),
  pdfToXps: (bytes, scale = 2) => ipcRenderer.invoke('conv:pdfToXps', { bytes, scale }),
  xpsToPdf: (bytes) => ipcRenderer.invoke('conv:xpsToPdf', bytes),
  viewXps: (bytes) => ipcRenderer.invoke('conv:viewXps', bytes),

  // ── meta ───────────────────────────────────────────────────────────────
  meta: () => ipcRenderer.invoke('app:meta'),
};

contextBridge.exposeInMainWorld('folio', folio);
