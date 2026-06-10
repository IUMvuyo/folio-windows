'use strict';

// xmlEngine — XML ↔ PDF + view/format, all on-device, no network.
//
// SHARED ENGINE (same logic copied into Folio-Android / Folio-Business as ESM;
// this is the CommonJS variant for the Electron main process). All three
// operations are pure JS:
//
//   • XML → PDF        : pretty-print/indent the XML, paginate as monospace text.
//   • PDF → XML        : extract text per page (renderEngine.extractText) and
//                        emit  <document><page number="1"><text>…</text></page>…
//                        </document>  with every value XML-escaped. Best-effort.
//   • formatXml(view)  : parse + pretty-print the XML for the in-app viewer.
//
// No file ever leaves the machine.

const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const { extractText } = require('./renderEngine');

// ── XML pretty-printer (dependency-free) ─────────────────────────────────────

function xmlEscape(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function decodeBytesToString(input) {
  if (typeof input === 'string') return input;
  const u8 = input instanceof Uint8Array ? input : new Uint8Array(input);
  let start = 0;
  if (u8.length >= 3 && u8[0] === 0xef && u8[1] === 0xbb && u8[2] === 0xbf) start = 3;
  if (typeof TextDecoder !== 'undefined') {
    return new TextDecoder('utf-8').decode(u8.subarray(start));
  }
  return Buffer.from(u8.subarray(start)).toString('utf-8');
}

function formatXml(input, { indent = '  ' } = {}) {
  const xml = decodeBytesToString(input).trim();
  if (!xml) throw new Error('Empty XML.');

  const out = [];
  let depth = 0;
  let i = 0;
  const n = xml.length;
  const pad = (d) => indent.repeat(Math.max(0, d));

  while (i < n) {
    if (xml[i] === '<') {
      let close;
      if (xml.startsWith('<![CDATA[', i)) {
        close = xml.indexOf(']]>', i);
        close = close === -1 ? n : close + 3;
      } else if (xml.startsWith('<!--', i)) {
        close = xml.indexOf('-->', i);
        close = close === -1 ? n : close + 3;
      } else {
        close = xml.indexOf('>', i);
        close = close === -1 ? n : close + 1;
      }
      const token = xml.slice(i, close);
      i = close;

      const isDecl = token.startsWith('<?');
      const isComment = token.startsWith('<!--');
      const isCData = token.startsWith('<![CDATA[');
      const isDoctype = /^<!DOCTYPE/i.test(token);
      const isClose = token.startsWith('</');
      const isSelfClose = token.endsWith('/>') || isDecl || isComment || isCData || isDoctype;
      const isOpen = !isClose && !isSelfClose;

      if (isClose) depth = Math.max(0, depth - 1);

      if (isOpen) {
        const nextLt = xml.indexOf('<', i);
        const between = nextLt === -1 ? xml.slice(i) : xml.slice(i, nextLt);
        const trimmed = between.trim();
        if (trimmed && nextLt !== -1 && xml.startsWith('</', nextLt)) {
          const closeEnd = xml.indexOf('>', nextLt);
          const closeTok = xml.slice(nextLt, closeEnd === -1 ? n : closeEnd + 1);
          out.push(pad(depth) + token + trimmed + closeTok);
          i = closeEnd === -1 ? n : closeEnd + 1;
          continue;
        }
      }

      out.push(pad(depth) + token);
      if (isOpen) depth += 1;
    } else {
      const nextLt = xml.indexOf('<', i);
      const text = (nextLt === -1 ? xml.slice(i) : xml.slice(i, nextLt));
      const trimmed = text.trim();
      if (trimmed) out.push(pad(depth) + trimmed);
      i = nextLt === -1 ? n : nextLt;
    }
  }
  return out.join('\n');
}

// ── XML → PDF (paginated monospace render) ───────────────────────────────────

const A4 = { w: 595.28, h: 841.89 };
const MARGIN = 48;

function wrapMono(line, font, size, maxWidth) {
  if (line === '') return [''];
  const lines = [];
  let cur = '';
  for (const ch of line) {
    const trial = cur + ch;
    if (font.widthOfTextAtSize(trial, size) > maxWidth && cur) {
      lines.push(cur);
      cur = ch;
    } else {
      cur = trial;
    }
  }
  if (cur) lines.push(cur);
  return lines.length ? lines : [''];
}

async function xmlToPdf(input, { title = 'XML' } = {}) {
  const pretty = formatXml(input);
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Courier);
  const bold = await doc.embedFont(StandardFonts.CourierBold);
  const size = 9;
  const lead = 12;
  const colW = A4.w - MARGIN * 2;

  let page = doc.addPage([A4.w, A4.h]);
  let y = A4.h - MARGIN;
  const newPage = () => { page = doc.addPage([A4.w, A4.h]); y = A4.h - MARGIN; };

  if (title) {
    page.drawText(title, { x: MARGIN, y: y - 12, size: 13, font: bold, color: rgb(0.75, 0.14, 0.13) });
    y -= 13 + lead;
  }

  for (const raw of pretty.split('\n')) {
    for (const line of wrapMono(raw, font, size, colW)) {
      if (y - lead < MARGIN) newPage();
      page.drawText(line, { x: MARGIN, y: y - size, size, font, color: rgb(0.16, 0.14, 0.12) });
      y -= lead;
    }
  }
  return doc.save();
}

// ── PDF → XML (structured, best-effort) ──────────────────────────────────────

async function pdfToXml(pdfBytes) {
  const { pages } = await extractText(pdfBytes);
  const parts = ['<?xml version="1.0" encoding="UTF-8"?>', '<document>'];
  const src = pages.length ? pages : [''];
  src.forEach((pageText, idx) => {
    parts.push(`  <page number="${idx + 1}">`);
    parts.push(`    <text>${xmlEscape(pageText || '')}</text>`);
    parts.push('  </page>');
  });
  parts.push('</document>');
  const xml = parts.join('\n');
  return Buffer.from(xml, 'utf-8');
}

function viewXml(input) {
  return formatXml(input);
}

module.exports = { xmlEscape, formatXml, xmlToPdf, pdfToXml, viewXml };
