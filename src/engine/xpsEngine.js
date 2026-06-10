'use strict';

// xpsEngine — PDF ↔ XPS + view, best-effort, all on-device, no network.
//
// SHARED ENGINE (same logic copied into Folio-Android / Folio-Business as ESM;
// this is the CommonJS variant for the Electron main process).
//
// There is no clean pure-JS XPS renderer, so — like the existing PDF↔Office
// path — we go image/text-based:
//
//   • PDF → XPS : rasterize each PDF page to a PNG (caller injects a `rasterize`
//                 fn that returns [{index,width,height,bytes}], same shape as
//                 main.js rasterizePages), then build a VALID XPS OPC package
//                 (a ZIP). ← works well.
//   • XPS → PDF : unzip, walk FixedDocumentSequence → FixedDocument → each
//                 FixedPage; positioned <Glyphs> render + embedded page images,
//                 with a text-extraction fallback. ← best-effort / "beta".
//   • viewXps  : render via XPS → PDF, then show in the PDF viewer.
//
// No file ever leaves the machine.

const JSZip = require('jszip');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');

// XPS units are 1/96 inch (DIPs); PDF user units are 1/72 inch.
const DIP_TO_PT = 72 / 96;

// ── small XML helpers (no DOM) ───────────────────────────────────────────────

function xmlEsc(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function xmlUnesc(s) {
  return String(s)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
    .replace(/&amp;/g, '&');
}

function attr(tag, name) {
  const m = tag.match(new RegExp(`\\b${name}\\s*=\\s*"([^"]*)"`)) ||
            tag.match(new RegExp(`\\b${name}\\s*=\\s*'([^']*)'`));
  return m ? m[1] : null;
}

function toU8(x) {
  if (x instanceof Uint8Array) return x;
  if (x instanceof ArrayBuffer) return new Uint8Array(x);
  if (typeof Buffer !== 'undefined' && Buffer.isBuffer && Buffer.isBuffer(x)) return new Uint8Array(x);
  if (x && x.buffer) return new Uint8Array(x.buffer, x.byteOffset, x.byteLength);
  return new Uint8Array(x);
}

// ═══ PDF → XPS ═══════════════════════════════════════════════════════════════

async function pdfToXps(pdfBytes, rasterize, { scale = 2 } = {}) {
  if (typeof rasterize !== 'function') {
    throw new Error('pdfToXps needs a rasterize() function.');
  }
  const imgs = await rasterize(toU8(pdfBytes), { scale, format: 'png' });
  if (!imgs || !imgs.length) throw new Error('No pages to convert.');

  const zip = new JSZip();

  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="fdseq" ContentType="application/vnd.ms-package.xps-fixeddocumentsequence+xml"/>' +
      '<Default Extension="fdoc" ContentType="application/vnd.ms-package.xps-fixeddocument+xml"/>' +
      '<Default Extension="fpage" ContentType="application/vnd.ms-package.xps-fixedpage+xml"/>' +
      '<Default Extension="png" ContentType="image/png"/>' +
      '</Types>'
  );

  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" ' +
      'Type="http://schemas.microsoft.com/xps/2005/06/fixedrepresentation" ' +
      'Target="/FixedDocumentSequence.fdseq"/>' +
      '</Relationships>'
  );

  zip.file(
    'FixedDocumentSequence.fdseq',
    '<?xml version="1.0" encoding="UTF-8"?>' +
      '<FixedDocumentSequence xmlns="http://schemas.microsoft.com/xps/2005/06">' +
      '<DocumentReference Source="/Documents/1/FixedDocument.fdoc"/>' +
      '</FixedDocumentSequence>'
  );

  let fdoc =
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<FixedDocument xmlns="http://schemas.microsoft.com/xps/2005/06">';
  imgs.forEach((_, idx) => {
    fdoc += `<PageContent Source="/Documents/1/Pages/${idx + 1}.fpage"/>`;
  });
  fdoc += '</FixedDocument>';
  zip.file('Documents/1/FixedDocument.fdoc', fdoc);

  imgs.forEach((img, idx) => {
    const i = idx + 1;
    const wDip = Math.max(1, Math.round(img.width / scale));
    const hDip = Math.max(1, Math.round(img.height / scale));
    const imgName = `${i}.png`;
    const imgPath = `/Documents/1/Resources/Images/${imgName}`;

    const fpage =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      `<FixedPage xmlns="http://schemas.microsoft.com/xps/2005/06" ` +
      `Width="${wDip}" Height="${hDip}" xml:lang="en-US">` +
      `<Path Data="M 0,0 L ${wDip},0 ${wDip},${hDip} 0,${hDip} Z">` +
      '<Path.Fill>' +
      `<ImageBrush ImageSource="${xmlEsc(imgPath)}" ` +
      `Viewbox="0,0 ${img.width},${img.height}" ViewboxUnits="Absolute" ` +
      `Viewport="0,0 ${wDip},${hDip}" ViewportUnits="Absolute" TileMode="None"/>` +
      '</Path.Fill>' +
      '</Path>' +
      '</FixedPage>';
    zip.file(`Documents/1/Pages/${i}.fpage`, fpage);

    zip.file(
      `Documents/1/Pages/_rels/${i}.fpage.rels`,
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId1" ` +
        `Type="http://schemas.microsoft.com/xps/2005/06/required-resource" ` +
        `Target="/Documents/1/Resources/Images/${imgName}"/>` +
        '</Relationships>'
    );

    zip.file(`Documents/1/Resources/Images/${imgName}`, toU8(img.bytes));
  });

  // In Node/Electron, return a Buffer.
  return zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}

// ═══ XPS → PDF ═══════════════════════════════════════════════════════════════

function resolvePath(base, target) {
  if (!target) return null;
  let t = target.replace(/^\.\//, '');
  if (t.startsWith('/')) return t.slice(1);
  const dir = base.slice(0, base.lastIndexOf('/') + 1);
  const parts = (dir + t).split('/');
  const stack = [];
  for (const p of parts) {
    if (p === '..') stack.pop();
    else if (p !== '.' && p !== '') stack.push(p);
  }
  return stack.join('/');
}

function parseGlyphs(tag) {
  const unicode = attr(tag, 'UnicodeString');
  if (unicode == null) return null;
  const text = xmlUnesc(unicode);
  if (!text) return null;
  const ox = parseFloat(attr(tag, 'OriginX') || '0') || 0;
  const oy = parseFloat(attr(tag, 'OriginY') || '0') || 0;
  const em = parseFloat(attr(tag, 'FontRenderingEmSize') || '12') || 12;
  const fill = attr(tag, 'Fill') || '#FF000000';
  const bold = /Bold/i.test(attr(tag, 'FontUri') || attr(tag, 'StyleSimulations') || '');
  return { text, ox, oy, em, fill, bold };
}

function parseColor(s) {
  if (!s) return rgb(0.1, 0.09, 0.08);
  let h = s.trim().replace(/^#/, '');
  if (h.length === 8) h = h.slice(2);
  if (h.length === 6) {
    const r = parseInt(h.slice(0, 2), 16) / 255;
    const g = parseInt(h.slice(2, 4), 16) / 255;
    const b = parseInt(h.slice(4, 6), 16) / 255;
    if ([r, g, b].every((v) => !Number.isNaN(v))) return rgb(r, g, b);
  }
  return rgb(0.1, 0.09, 0.08);
}

function matchTags(xml, name) {
  const re = new RegExp(`<${name}\\b[^>]*?/?>`, 'g');
  return xml.match(re) || [];
}

function parseFixedPage(xml) {
  const openTag = (xml.match(/<FixedPage\b[^>]*>/) || [''])[0];
  const width = parseFloat(attr(openTag, 'Width') || '816') || 816;
  const height = parseFloat(attr(openTag, 'Height') || '1056') || 1056;

  const glyphs = [];
  for (const g of matchTags(xml, 'Glyphs')) {
    const parsed = parseGlyphs(g);
    if (parsed) glyphs.push(parsed);
  }

  const images = [];
  const seen = new Set();
  const srcRe = /ImageSource\s*=\s*"([^"]+)"/g;
  let m;
  while ((m = srcRe.exec(xml)) !== null) {
    const ref = xmlUnesc(m[1]).replace(/^\{[^}]*\}/, '');
    const clean = ref.split(' ')[0];
    if (clean && !seen.has(clean)) { seen.add(clean); images.push(clean); }
  }
  return { width, height, glyphs, images };
}

async function pageOrder(zip) {
  const fileKey = (name) =>
    Object.keys(zip.files).find((k) => k.toLowerCase() === name.toLowerCase());

  const order = [];
  const seqKey = fileKey('FixedDocumentSequence.fdseq') ||
    Object.keys(zip.files).find((k) => /\.fdseq$/i.test(k));
  if (seqKey) {
    const seq = await zip.file(seqKey).async('string');
    const docRefs = (seq.match(/<DocumentReference\b[^>]*>/g) || [])
      .map((t) => attr(t, 'Source'))
      .filter(Boolean);
    for (const dref of docRefs) {
      const docKey = resolvePath(seqKey, dref);
      const k = fileKey(docKey) || docKey;
      if (!zip.files[k]) continue;
      const fdoc = await zip.file(k).async('string');
      const pageRefs = (fdoc.match(/<PageContent\b[^>]*>/g) || [])
        .map((t) => attr(t, 'Source'))
        .filter(Boolean);
      for (const pref of pageRefs) {
        const pk = resolvePath(k, pref);
        order.push(fileKey(pk) || pk);
      }
    }
  }
  if (order.length) return order.filter((k) => zip.files[k]);
  return Object.keys(zip.files)
    .filter((k) => /\.fpage$/i.test(k))
    .sort();
}

async function xpsToPdf(xpsBytes) {
  const zip = await JSZip.loadAsync(toU8(xpsBytes));
  const pageKeys = await pageOrder(zip);
  if (!pageKeys.length) throw new Error('No FixedPage parts found — not a valid XPS.');

  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);

  for (const key of pageKeys) {
    let xml;
    try { xml = await zip.file(key).async('string'); } catch (_) { continue; }
    const { width, height, glyphs, images } = parseFixedPage(xml);

    const wPt = Math.max(1, width * DIP_TO_PT);
    const hPt = Math.max(1, height * DIP_TO_PT);
    const page = pdf.addPage([wPt, hPt]);

    let drewImage = false;
    for (const ref of images) {
      const imgKey = resolvePath(key, ref);
      const k = Object.keys(zip.files).find((f) => f.toLowerCase() === (imgKey || '').toLowerCase());
      if (!k) continue;
      try {
        const bytes = await zip.file(k).async('uint8array');
        const isPng = /\.png$/i.test(k) || (bytes[0] === 0x89 && bytes[1] === 0x50);
        const embedded = isPng ? await pdf.embedPng(bytes) : await pdf.embedJpg(bytes);
        page.drawImage(embedded, { x: 0, y: 0, width: wPt, height: hPt });
        drewImage = true;
      } catch (_) { /* skip unembeddable image */ }
    }

    let drewGlyph = false;
    for (const g of glyphs) {
      const size = Math.max(1, g.em * DIP_TO_PT);
      const x = g.ox * DIP_TO_PT;
      const y = hPt - g.oy * DIP_TO_PT;
      try {
        page.drawText(g.text, {
          x,
          y: y - size,
          size,
          font: g.bold ? bold : font,
          color: parseColor(g.fill),
        });
        drewGlyph = true;
      } catch (_) { /* unsupported glyph chars — skip */ }
    }

    if (!drewImage && !drewGlyph) {
      const all = glyphs.map((g) => g.text).join('\n') ||
        (xml.match(/UnicodeString\s*=\s*"([^"]*)"/g) || [])
          .map((t) => xmlUnesc(attr(t, 'UnicodeString') || '')).join('\n');
      const size = 11;
      const lead = 15;
      let yy = hPt - 48;
      for (const lineRaw of (all || '(no extractable content)').split('\n')) {
        for (const line of wrapText(lineRaw, font, size, wPt - 96)) {
          if (yy < 40) break;
          try { page.drawText(line, { x: 48, y: yy, size, font, color: rgb(0.16, 0.14, 0.12) }); } catch (_) {}
          yy -= lead;
        }
      }
    }
  }

  if (pdf.getPageCount() === 0) pdf.addPage();
  return Buffer.from(await pdf.save());
}

function wrapText(text, font, size, maxWidth) {
  if (!text) return [''];
  const words = String(text).split(/\s+/);
  const lines = [];
  let cur = '';
  for (const w of words) {
    const trial = cur ? cur + ' ' + w : w;
    let width;
    try { width = font.widthOfTextAtSize(trial, size); } catch (_) { width = trial.length * size * 0.5; }
    if (width <= maxWidth) cur = trial;
    else { if (cur) lines.push(cur); cur = w; }
  }
  if (cur) lines.push(cur);
  return lines.length ? lines : [''];
}

async function viewXps(xpsBytes) {
  return xpsToPdf(xpsBytes);
}

module.exports = { pdfToXps, xpsToPdf, viewXps };
