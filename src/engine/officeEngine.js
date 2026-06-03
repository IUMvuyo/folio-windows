'use strict';

// officeEngine — PDF ↔ Office, all on-device.
//   PDF → Word/Excel/PPT : write OOXML with docx / exceljs / pptxgenjs.
//   Office → PDF         : parse with mammoth (docx) / exceljs (xlsx), render
//                          to HTML, then the main process turns that HTML into
//                          a PDF with Electron's webContents.printToPDF (offline).
// No file ever leaves the machine.

const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
} = require('docx');
const ExcelJS = require('exceljs');
const PptxGenJS = require('pptxgenjs');
const mammoth = require('mammoth');

const { extractText } = require('./renderEngine');

// ── PDF → Office (write OOXML) ───────────────────────────────────────────────

/** PDF → Word (.docx): extract text per page, write paragraphs. Best-effort. */
async function pdfToWord(pdfBytes) {
  const { pages } = await extractText(pdfBytes);
  const children = [];
  pages.forEach((pageText, idx) => {
    if (idx > 0) {
      children.push(
        new Paragraph({
          children: [new TextRun({ text: `Page ${idx + 1}`, bold: true, color: 'C02423' })],
          heading: HeadingLevel.HEADING_3,
          spacing: { before: 240, after: 120 },
        })
      );
    }
    for (const line of pageText.split('\n')) {
      children.push(new Paragraph({ children: [new TextRun(line)] }));
    }
  });
  const doc = new Document({
    creator: 'Folio — Private Edition',
    title: 'Converted from PDF',
    sections: [{ children: children.length ? children : [new Paragraph('')] }],
  });
  return Packer.toBuffer(doc);
}

/** PDF → Excel (.xlsx): one row per text line; naive column split on 2+ spaces. */
async function pdfToExcel(pdfBytes) {
  const { pages } = await extractText(pdfBytes);
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Folio — Private Edition';
  pages.forEach((pageText, idx) => {
    const ws = wb.addWorksheet(`Page ${idx + 1}`);
    for (const line of pageText.split('\n')) {
      if (!line.trim()) continue;
      const cols = line.split(/\s{2,}|\t/).map((c) => c.trim());
      ws.addRow(cols);
    }
  });
  if (wb.worksheets.length === 0) wb.addWorksheet('Sheet1');
  // exceljs returns ArrayBuffer in Node; normalize to Buffer.
  const ab = await wb.xlsx.writeBuffer();
  return Buffer.from(ab);
}

/** PDF → PowerPoint (.pptx): one slide per page with that page's text. */
async function pdfToPpt(pdfBytes) {
  const { pages } = await extractText(pdfBytes);
  const pptx = new PptxGenJS();
  pptx.author = 'Folio — Private Edition';
  pptx.defineLayout({ name: 'A4', width: 10, height: 7.5 });
  pptx.layout = 'A4';
  if (pages.length === 0) pages.push('');
  pages.forEach((pageText, idx) => {
    const slide = pptx.addSlide();
    slide.addText(`Page ${idx + 1}`, {
      x: 0.4, y: 0.2, w: 9.2, h: 0.5, fontFace: 'Georgia', fontSize: 18, bold: true, color: 'C02423',
    });
    slide.addText(pageText || '(no extractable text)', {
      x: 0.4, y: 0.8, w: 9.2, h: 6.4, fontFace: 'Georgia', fontSize: 12,
      color: '28231E', valign: 'top',
    });
  });
  // pptxgenjs write returns a Node Buffer when outputType is 'nodebuffer'.
  return pptx.write({ outputType: 'nodebuffer' });
}

// ── Office → HTML (the main process prints the HTML to PDF, fully offline) ───

const HTML_SHELL = (title, body) => `<!doctype html><html><head><meta charset="utf-8">
<title>${title}</title>
<style>
  @page { margin: 18mm; }
  body { font-family: Georgia, 'Times New Roman', serif; color:#28231E; font-size:12pt; line-height:1.5; }
  h1,h2,h3 { font-family: Georgia, serif; }
  table { border-collapse: collapse; width:100%; margin:8pt 0; }
  td,th { border:1px solid #2A2621; padding:4pt 6pt; font-size:10pt; text-align:left; }
  th { background:#EDE6D8; }
  .folio-slide { page-break-after: always; border-bottom:2px solid #C02423; padding-bottom:12pt; margin-bottom:12pt; }
  img { max-width:100%; }
</style></head><body>${body}</body></html>`;

/** Word (.docx) → HTML (caller prints to PDF). */
async function wordToHtml(docxBytes) {
  const { value } = await mammoth.convertToHtml({ buffer: Buffer.from(docxBytes) });
  return HTML_SHELL('Document', value || '<p>(empty document)</p>');
}

/** Excel (.xlsx) → HTML tables (caller prints to PDF). */
async function excelToHtml(xlsxBytes) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(xlsxBytes);
  let body = '';
  wb.eachSheet((ws) => {
    body += `<h2>${escapeHtml(ws.name)}</h2><table>`;
    ws.eachRow((row) => {
      body += '<tr>';
      const vals = Array.isArray(row.values) ? row.values.slice(1) : [];
      const maxCol = Math.max(vals.length, ws.columnCount || 0);
      for (let c = 0; c < maxCol; c++) {
        const cell = row.getCell(c + 1);
        const v = cell && cell.value != null ? cellText(cell.value) : '';
        body += `<td>${escapeHtml(v)}</td>`;
      }
      body += '</tr>';
    });
    body += '</table>';
  });
  return HTML_SHELL('Spreadsheet', body || '<p>(empty workbook)</p>');
}

function cellText(v) {
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.text) return String(v.text);
    if (v.result != null) return String(v.result);
    if (v.richText) return v.richText.map((r) => r.text).join('');
    if (v.hyperlink) return String(v.hyperlink);
    return '';
  }
  return String(v);
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

module.exports = {
  pdfToWord,
  pdfToExcel,
  pdfToPpt,
  wordToHtml,
  excelToHtml,
  HTML_SHELL,
  escapeHtml,
};
