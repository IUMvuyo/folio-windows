'use strict';

const { app, BrowserWindow, ipcMain, dialog, session, nativeImage } = require('electron');
const path = require('path');
const fs = require('fs/promises');

const pdfEngine = require('../engine/pdfEngine');
const renderEngine = require('../engine/renderEngine');
const ocrEngine = require('../engine/ocrEngine');
const officeEngine = require('../engine/officeEngine');
const xmlEngine = require('../engine/xmlEngine');
const xpsEngine = require('../engine/xpsEngine');

// ── Privacy enforcement ──────────────────────────────────────────────────────
// Folio's whole promise is "nothing leaves this device." We HARD-BLOCK every
// network request at the session level. file:// (local app assets) is allowed;
// anything else is cancelled. This is belt-and-suspenders on top of "we wrote
// no fetch/network code."
function lockdownNetwork() {
  const ses = session.defaultSession;
  ses.webRequest.onBeforeRequest((details, cb) => {
    const url = details.url || '';
    const ok =
      url.startsWith('file://') ||
      url.startsWith('devtools://') ||
      url.startsWith('chrome-extension://') ||
      url.startsWith('blob:') ||
      url.startsWith('data:');
    if (!ok) {
      console.warn('[FOLIO] BLOCKED network request:', url);
      return cb({ cancel: true });
    }
    cb({ cancel: false });
  });
  // Deny all permission prompts (geolocation, media, etc.) — none are needed.
  ses.setPermissionRequestHandler((_wc, _perm, cb) => cb(false));
}

let mainWindow = null;
let rasterWindow = null; // hidden renderer used only to rasterize PDF pages

function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 1180,
    height: 820,
    minWidth: 880,
    minHeight: 620,
    backgroundColor: '#F4EFE6',
    title: 'Folio',
    show: false,
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, '..', 'preload', 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false, // preload needs require() for the IPC bridge wiring
      spellcheck: false,
    },
  });
  mainWindow.loadFile(path.join(__dirname, '..', 'renderer', 'index.html'));
  mainWindow.once('ready-to-show', () => mainWindow.show());
  mainWindow.on('closed', () => (mainWindow = null));
}

// Hidden window that loads pdf.js in a real DOM, so we can rasterize pages to
// canvas (PDF→images, raster compress) without any native canvas binary.
function ensureRasterWindow() {
  if (rasterWindow && !rasterWindow.isDestroyed()) return rasterWindow;
  rasterWindow = new BrowserWindow({
    show: false,
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, '..', 'preload', 'raster-preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      offscreen: false,
    },
  });
  rasterWindow.loadFile(path.join(__dirname, '..', 'renderer', 'raster.html'));
  return rasterWindow;
}

function rasterReady() {
  const win = ensureRasterWindow();
  return new Promise((resolve) => {
    if (win.webContents.isLoadingMainFrame && !win.webContents.isLoading()) {
      return resolve(win);
    }
    win.webContents.once('did-finish-load', () => resolve(win));
    // If already loaded, the once may not fire — guard with a probe.
    win.webContents
      .executeJavaScript('window.__folioRasterReady === true')
      .then((r) => { if (r) resolve(win); })
      .catch(() => {});
  });
}

/**
 * Rasterize PDF pages to PNG/JPG bytes via the hidden pdf.js renderer.
 * Returns [{ index, width, height, bytes }].
 */
async function rasterizePages(pdfBytes, { scale = 2, format = 'png', quality = 92, pages } = {}) {
  const win = await rasterReady();
  // Pass the PDF as a base64 string to avoid structured-clone issues, then the
  // raster renderer returns RGBA buffers we encode here in main.
  const b64 = Buffer.from(pdfBytes).toString('base64');
  const result = await win.webContents.executeJavaScript(
    `window.__rasterize(${JSON.stringify(b64)}, ${JSON.stringify({ scale, pages })})`
  );
  // result: [{ index, width, height, rgbaBase64 }]
  return result.map((p) => {
    const rgba = Buffer.from(p.rgbaBase64, 'base64');
    const bytes = renderEngine.encodeRGBA(rgba, p.width, p.height, format, quality);
    return { index: p.index, width: p.width, height: p.height, bytes };
  });
}

// ── File dialog helpers ──────────────────────────────────────────────────────

async function pickOpen({ multi = false, filters } = {}) {
  const props = multi ? ['openFile', 'multiSelections'] : ['openFile'];
  const { canceled, filePaths } = await dialog.showOpenDialog(mainWindow, {
    properties: props,
    filters: filters || [{ name: 'All Files', extensions: ['*'] }],
  });
  if (canceled || !filePaths.length) return null;
  const out = [];
  for (const fp of filePaths) {
    out.push({ path: fp, name: path.basename(fp), bytes: await fs.readFile(fp) });
  }
  return out;
}

async function pickSave(defaultName, filters) {
  const { canceled, filePath } = await dialog.showSaveDialog(mainWindow, {
    defaultPath: defaultName,
    filters,
  });
  if (canceled || !filePath) return null;
  return filePath;
}

async function pickDir() {
  const { canceled, filePaths } = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory', 'createDirectory'],
  });
  if (canceled || !filePaths.length) return null;
  return filePaths[0];
}

const PDF_FILTER = [{ name: 'PDF', extensions: ['pdf'] }];
const IMG_FILTER = [{ name: 'Images', extensions: ['png', 'jpg', 'jpeg'] }];
const XML_FILTER = [{ name: 'XML', extensions: ['xml'] }];
const XPS_FILTER = [{ name: 'XPS', extensions: ['xps', 'oxps'] }];

function toBuf(x) {
  if (Buffer.isBuffer(x)) return x;
  if (x instanceof Uint8Array) return Buffer.from(x);
  if (x instanceof ArrayBuffer) return Buffer.from(x);
  if (x && x.buffer) return Buffer.from(x.buffer, x.byteOffset, x.byteLength);
  return Buffer.from(x);
}

// ── IPC: import / export ─────────────────────────────────────────────────────

ipcMain.handle('dlg:openPdf', (_e, multi) => pickOpen({ multi, filters: PDF_FILTER }));
ipcMain.handle('dlg:openImages', () => pickOpen({ multi: true, filters: IMG_FILTER }));
ipcMain.handle('dlg:openOffice', (_e, kind) => {
  const map = {
    word: [{ name: 'Word', extensions: ['docx'] }],
    excel: [{ name: 'Excel', extensions: ['xlsx'] }],
    ppt: [{ name: 'PowerPoint', extensions: ['pptx'] }],
  };
  return pickOpen({ multi: false, filters: map[kind] || [{ name: 'All', extensions: ['*'] }] });
});
ipcMain.handle('dlg:openXml', () => pickOpen({ multi: false, filters: XML_FILTER }));
ipcMain.handle('dlg:openXps', () => pickOpen({ multi: false, filters: XPS_FILTER }));
ipcMain.handle('dlg:openAny', () => pickOpen({ multi: false }));

ipcMain.handle('save:bytes', async (_e, { bytes, defaultName, kind }) => {
  const filters = {
    pdf: PDF_FILTER,
    png: [{ name: 'PNG', extensions: ['png'] }],
    jpg: [{ name: 'JPEG', extensions: ['jpg', 'jpeg'] }],
    txt: [{ name: 'Text', extensions: ['txt'] }],
    docx: [{ name: 'Word', extensions: ['docx'] }],
    xlsx: [{ name: 'Excel', extensions: ['xlsx'] }],
    pptx: [{ name: 'PowerPoint', extensions: ['pptx'] }],
    xml: XML_FILTER,
    xps: XPS_FILTER,
  }[kind] || [{ name: 'File', extensions: ['*'] }];
  const fp = await pickSave(defaultName, filters);
  if (!fp) return { saved: false };
  await fs.writeFile(fp, toBuf(bytes));
  return { saved: true, path: fp };
});

ipcMain.handle('save:multi', async (_e, { files }) => {
  // files: [{ name, bytes }]
  const dir = await pickDir();
  if (!dir) return { saved: false };
  const written = [];
  for (const f of files) {
    const fp = path.join(dir, f.name);
    await fs.writeFile(fp, toBuf(f.bytes));
    written.push(fp);
  }
  return { saved: true, dir, count: written.length };
});

// ── IPC: Organize ────────────────────────────────────────────────────────────

ipcMain.handle('pdf:merge', async (_e, buffers) =>
  pdfEngine.merge(buffers.map(toBuf)));
ipcMain.handle('pdf:split', async (_e, { bytes, baseName }) =>
  pdfEngine.splitToPages(toBuf(bytes), baseName));
ipcMain.handle('pdf:extract', async (_e, { bytes, range }) =>
  pdfEngine.extractPages(toBuf(bytes), range));
ipcMain.handle('pdf:delete', async (_e, { bytes, range }) =>
  pdfEngine.deletePages(toBuf(bytes), range));
ipcMain.handle('pdf:rotate', async (_e, { bytes, angle, range }) =>
  pdfEngine.rotate(toBuf(bytes), angle, range));
ipcMain.handle('pdf:info', async (_e, bytes) => pdfEngine.info(toBuf(bytes)));

ipcMain.handle('pdf:compress', async (_e, { bytes, raster, scale, quality }) => {
  if (!raster) return pdfEngine.compressStructural(toBuf(bytes));
  // Raster compress: render every page to JPG then rebuild PDF from images.
  const imgs = await rasterizePages(toBuf(bytes), {
    scale: scale || 1.5, format: 'jpg', quality: quality || 60,
  });
  return pdfEngine.imagesToPdf(
    imgs.map((p) => ({ bytes: p.bytes, type: 'jpg' })),
    { pageSize: 'fit' }
  );
});

// ── IPC: Images ──────────────────────────────────────────────────────────────

ipcMain.handle('pdf:toImages', async (_e, { bytes, format, scale }) => {
  const imgs = await rasterizePages(toBuf(bytes), {
    scale: scale || 2, format: format || 'png',
  });
  return imgs.map((p) => ({
    name: `page-${String(p.index + 1).padStart(3, '0')}.${format === 'jpg' ? 'jpg' : 'png'}`,
    bytes: p.bytes,
    width: p.width,
    height: p.height,
  }));
});

ipcMain.handle('pdf:fromImages', async (_e, { images, pageSize }) => {
  const norm = images.map((im) => ({
    bytes: toBuf(im.bytes),
    type: /png/i.test(im.name) ? 'png' : 'jpg',
  }));
  return pdfEngine.imagesToPdf(norm, { pageSize: pageSize || 'fit' });
});

// ── IPC: Viewer (rasterize one page for display) ─────────────────────────────

ipcMain.handle('pdf:renderPage', async (_e, { bytes, pageIndex, scale }) => {
  const imgs = await rasterizePages(toBuf(bytes), {
    scale: scale || 1.5, format: 'png', pages: [pageIndex],
  });
  if (!imgs.length) return null;
  const p = imgs[0];
  return { dataUrl: 'data:image/png;base64,' + Buffer.from(p.bytes).toString('base64'),
           width: p.width, height: p.height };
});

ipcMain.handle('pdf:geometry', async (_e, bytes) =>
  renderEngine.pageGeometry(toBuf(bytes)));

// ── IPC: Secure & Sign ───────────────────────────────────────────────────────

ipcMain.handle('pdf:watermark', async (_e, { bytes, text, opacity }) =>
  pdfEngine.watermark(toBuf(bytes), text, { opacity }));
ipcMain.handle('pdf:pageNumbers', async (_e, { bytes, format, position }) =>
  pdfEngine.pageNumbers(toBuf(bytes), { format, position }));
ipcMain.handle('pdf:flatten', async (_e, bytes) => pdfEngine.flatten(toBuf(bytes)));
ipcMain.handle('pdf:sign', async (_e, { bytes, signaturePng, pageIndex, x, y, width, height }) =>
  pdfEngine.stampSignature(toBuf(bytes), toBuf(signaturePng), { pageIndex, x, y, width, height }));

// ── IPC: Scan & Read (OCR + text extract) ────────────────────────────────────

ipcMain.handle('pdf:extractText', async (_e, bytes) =>
  renderEngine.extractText(toBuf(bytes)));

ipcMain.handle('ocr:image', async (e, { bytes }) => {
  return ocrEngine.ocrImage(toBuf(bytes), {
    onProgress: (p) => e.sender.send('ocr:progress', p),
  });
});

ipcMain.handle('ocr:pdf', async (e, { bytes, scale }) => {
  // Rasterize each page, OCR each image, return combined text.
  e.sender.send('ocr:progress', 0.02);
  const imgs = await rasterizePages(toBuf(bytes), { scale: scale || 2, format: 'png' });
  const pageImages = imgs.map((p) => ({ bytes: p.bytes, width: p.width, height: p.height }));
  const results = await ocrEngine.ocrPages(pageImages, {
    onProgress: (p) => e.sender.send('ocr:progress', 0.05 + p * 0.95),
  });
  return { pages: results.map((r) => r.text), text: results.map((r) => r.text).join('\n\n') };
});

// ── IPC: Convert (PDF → Office) ──────────────────────────────────────────────

ipcMain.handle('conv:pdfToWord', async (_e, bytes) => officeEngine.pdfToWord(toBuf(bytes)));
ipcMain.handle('conv:pdfToExcel', async (_e, bytes) => officeEngine.pdfToExcel(toBuf(bytes)));
ipcMain.handle('conv:pdfToPpt', async (_e, bytes) => officeEngine.pdfToPpt(toBuf(bytes)));

// ── IPC: Convert (Office → PDF) via offline printToPDF ───────────────────────

async function htmlToPdf(html) {
  // A throwaway hidden window renders the HTML, then we print it to PDF fully
  // offline (network is already locked down at the session level).
  const win = new BrowserWindow({
    show: false,
    webPreferences: { contextIsolation: true, nodeIntegration: false, javascript: false },
  });
  try {
    await win.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(html));
    const pdf = await win.webContents.printToPDF({
      printBackground: true,
      margins: { marginType: 'default' },
    });
    return pdf;
  } finally {
    win.destroy();
  }
}

ipcMain.handle('conv:wordToPdf', async (_e, bytes) => {
  const html = await officeEngine.wordToHtml(toBuf(bytes));
  return htmlToPdf(html);
});
ipcMain.handle('conv:excelToPdf', async (_e, bytes) => {
  const html = await officeEngine.excelToHtml(toBuf(bytes));
  return htmlToPdf(html);
});
// PowerPoint → PDF: pptx parsing to layout is out of scope for pure-JS fidelity;
// the renderer marks this "coming soon" rather than faking it.

// ── IPC: Convert (XML / XPS) — shared xmlEngine + xpsEngine ──────────────────

ipcMain.handle('conv:xmlToPdf', async (_e, bytes) => xmlEngine.xmlToPdf(toBuf(bytes)));
ipcMain.handle('conv:pdfToXml', async (_e, bytes) => xmlEngine.pdfToXml(toBuf(bytes)));
ipcMain.handle('conv:viewXml', async (_e, bytes) => xmlEngine.viewXml(toBuf(bytes)));

ipcMain.handle('conv:pdfToXps', async (e, { bytes, scale }) =>
  xpsEngine.pdfToXps(
    toBuf(bytes),
    (b, o) => rasterizePages(b, { scale: (o && o.scale) || scale || 2, format: 'png' }),
    { scale: scale || 2 }
  ));
ipcMain.handle('conv:xpsToPdf', async (_e, bytes) => xpsEngine.xpsToPdf(toBuf(bytes)));
ipcMain.handle('conv:viewXps', async (_e, bytes) => xpsEngine.xpsToPdf(toBuf(bytes)));

ipcMain.handle('app:meta', () => ({
  version: app.getVersion(),
  electron: process.versions.electron,
  node: process.versions.node,
  langPath: ocrEngine.resolveLangPath(),
}));

// ── lifecycle ────────────────────────────────────────────────────────────────

app.whenReady().then(() => {
  lockdownNetwork();
  createMainWindow();
  ensureRasterWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createMainWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
