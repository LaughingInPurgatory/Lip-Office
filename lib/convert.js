// Format dispatch: every supported file type goes in and out through here.
import { docxToHtml, htmlToDocx, xlsxToSheets, sheetsToXlsx } from './ooxml.js';
import { odtToHtml, htmlToOdt, odsToSheets, sheetsToOds } from './odf.js';
import { rtfToText, textToHtml, htmlToText, htmlToRtf } from './rtf.js';
import { sanitizeHtml } from '../shared/sanitize.js';
import { parseCsv, toCsv } from '../shared/csv.js';
import { normalizeWord, normalizeSheet, usedRange, extensionOf, kindOf } from '../shared/model.js';
import { makeEngine, numToCol } from '../shared/formula.js';
import { docToHtml, xlsToSheets } from './legacy.js';

const MIME = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  odt: 'application/vnd.oasis.opendocument.text',
  rtf: 'application/rtf',
  txt: 'text/plain; charset=utf-8',
  html: 'text/html; charset=utf-8',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  ods: 'application/vnd.oasis.opendocument.spreadsheet',
  csv: 'text/csv; charset=utf-8',
  doc: 'application/msword',
  dot: 'application/msword',
  xls: 'application/vnd.ms-excel',
  xlt: 'application/vnd.ms-excel',
};

export function mimeFor(extension) {
  return MIME[String(extension).toLowerCase()] || 'application/octet-stream';
}

function baseName(name) {
  return String(name).replace(/\.[^.]+$/, '') || 'Untitled';
}

function wordDocument(html, title, layout = {}) {
  return { kind: 'word', doc: normalizeWord({ ...layout, html: sanitizeHtml(html), title }) };
}

function csvToSheet(text, name) {
  const rows = parseCsv(text);
  const cells = {};
  rows.forEach((row, r) => {
    row.forEach((value, c) => {
      if (value !== '') cells[`${numToCol(c + 1)}${r + 1}`] = value;
    });
  });
  return { sheets: [{ name: String(name).slice(0, 31) || 'Sheet1', cells }], active: 0 };
}

function sheetToCsv(doc) {
  const sheet = doc.sheets[0] || { cells: {} };
  const engine = makeEngine(sheet.cells);
  const bounds = usedRange(sheet.cells);
  if (!bounds) return '';
  const rows = [];
  for (let row = 1; row <= bounds.rows; row++) {
    const values = [];
    for (let col = 1; col <= bounds.cols; col++) {
      const ref = `${numToCol(col)}${row}`;
      values.push(sheet.cells[ref] === undefined ? '' : engine.display(ref));
    }
    rows.push(values);
  }
  return toCsv(rows);
}

/**
 * @returns {{kind: 'word'|'sheet', doc: object}}
 * Word docs come back as HTML, sheets as `{sheets, active}`.
 */
export function parseDocument(name, buffer) {
  try {
    return readDocument(name, buffer);
  } catch (err) {
    // Anything that fails in here is the caller's file, not a server fault.
    if (!err.status) err.status = 400;
    throw err;
  }
}

function readDocument(name, buffer) {
  const extension = extensionOf(name);
  const kind = kindOf(name);

  if (!kind) throw new Error(`unsupported file type: ${name}`);
  const base = baseName(name);

  if (kind === 'word') {
    if (extension === 'doc' || extension === 'dot') {
      if (buffer.length < 512) throw new Error(`.${extension} is a legacy binary format and the uploaded file is too short to be valid; save it as .docx and retry.`);
      return wordDocument(docToHtml(buffer), base);
    }
    if (extension === 'docx') {
      const parsed = docxToHtml(buffer);
      return wordDocument(parsed.html, parsed.title || base, parsed);
    }
    if (extension === 'odt') {
      const parsed = odtToHtml(buffer);
      return wordDocument(parsed.html, parsed.title || base, parsed);
    }
    if (extension === 'rtf') return wordDocument(textToHtml(rtfToText(buffer).text), base);
    if (extension === 'txt') return wordDocument(textToHtml(buffer.toString('utf8')), base);
    return wordDocument(buffer.toString('utf8'), base); // html / htm
  }

  if (extension === 'xls' || extension === 'xlt') {
    if (buffer.length < 512) throw new Error(`.${extension} is a legacy binary format and the uploaded file is too short to be valid; save it as .xlsx and retry.`);
    return { kind, doc: normalizeSheet(xlsToSheets(buffer)) };
  }
  if (extension === 'xlsx') return { kind, doc: normalizeSheet(xlsxToSheets(buffer)) };
  if (extension === 'ods') return { kind, doc: normalizeSheet(odsToSheets(buffer)) };
  return { kind, doc: normalizeSheet(csvToSheet(buffer.toString('utf8'), base)) };
}

/** @returns {{buffer: Buffer, mime: string}} */
export function serializeDocument(doc, format, kind) {
  if (kind === 'word') {
    const normalized = normalizeWord({ ...(doc || {}), html: sanitizeHtml(doc?.html || ''), header: sanitizeHtml(doc?.header || ''), footer: sanitizeHtml(doc?.footer || '') });
    if (format === 'docx') return { buffer: htmlToDocx(normalized.html, normalized.title, normalized), mime: MIME.docx };
    if (format === 'odt') return { buffer: htmlToOdt(normalized.html, normalized.title, normalized), mime: MIME.odt };
    if (format === 'rtf') return { buffer: Buffer.from(htmlToRtf(normalized.html, normalized.title), 'latin1'), mime: MIME.rtf };
    if (format === 'txt') return { buffer: Buffer.from(htmlToText(normalized.html), 'utf8'), mime: MIME.txt };
    return { buffer: Buffer.from(normalized.html, 'utf8'), mime: MIME.html };
  }

  const sheet = normalizeSheet(doc);
  if (format === 'xlsx') return { buffer: sheetsToXlsx(sheet), mime: MIME.xlsx };
  if (format === 'ods') return { buffer: sheetsToOds(sheet), mime: MIME.ods };
  return { buffer: Buffer.from(sheetToCsv(sheet), 'utf8'), mime: MIME.csv };
}
