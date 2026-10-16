// Document shapes shared by the server and the browser, plus the validation
// that runs on anything arriving from the network.
import { parseImageDataUrl } from './images.js';

export const MAX_SHEET_ROWS = 5000;
export const MAX_SHEET_COLS = 256;

export function blankWord() {
  return { title: 'Untitled document', html: '<p><br></p>', header: '', footer: '', columns: 1, pageSize: 'letter', orientation: 'portrait', margin: 1 };
}

export function blankSheet() {
  return { sheets: [{ name: 'Sheet1', cells: {}, images: {} }], active: 0 };
}

/** File extension -> component. Anything absent here is not openable. */
export const EXTENSIONS = {
  doc: 'word', dot: 'word', docx: 'word', odt: 'word', rtf: 'word', txt: 'word', html: 'word', htm: 'word',
  xls: 'sheet', xlt: 'sheet', xlsx: 'sheet', ods: 'sheet', csv: 'sheet',
};

/** Formats each component can write. */
export const SAVE_FORMATS = {
  word: ['docx'],
  sheet: ['xlsx'],
};

export function extensionOf(name) {
  const match = /\.([A-Za-z0-9]+)$/.exec(String(name));
  return match ? match[1].toLowerCase() : '';
}

export function kindOf(name) {
  return EXTENSIONS[extensionOf(name)] || null;
}


const REF_RE = /^[A-Z]{1,3}[0-9]{1,7}$/;

export function normalizeWord(doc) {
  const source = doc && typeof doc === 'object' ? doc : {};
  const html = typeof source.html === 'string' ? source.html : '';
  const title = typeof source.title === 'string' && source.title.trim() ? source.title.trim() : 'Untitled document';
  const header = typeof source.header === 'string' ? source.header : '';
  const footer = typeof source.footer === 'string' ? source.footer : '';
  const columns = [1, 2, 3].includes(Number(source.columns)) ? Number(source.columns) : 1;
  const pageSize = source.pageSize === 'a4' ? 'a4' : 'letter';
  const orientation = source.orientation === 'landscape' ? 'landscape' : 'portrait';
  const margin = [0.5, 0.75, 1].includes(Number(source.margin)) ? Number(source.margin) : 1;
  return { title, html, header, footer, columns, pageSize, orientation, margin };
}

export function normalizeSheet(doc) {
  const source = doc && typeof doc === 'object' ? doc : {};
  const input = Array.isArray(source.sheets) ? source.sheets : [];

  const sheets = input.slice(0, 50).map((sheet, index) => {
    const cells = {};
    const images = {};
    const rawCells = sheet && typeof sheet.cells === 'object' && sheet.cells !== null ? sheet.cells : {};
    for (const [ref, value] of Object.entries(rawCells)) {
      const key = ref.toUpperCase();
      if (!REF_RE.test(key)) continue;
      if (typeof value !== 'string' && typeof value !== 'number') continue;
      const text = String(value);
      if (text !== '') cells[key] = text;
    }
    const rawImages = sheet && typeof sheet.images === 'object' && sheet.images !== null ? sheet.images : {};
    for (const [ref, image] of Object.entries(rawImages)) {
      const key = ref.toUpperCase();
      if (!REF_RE.test(key) || !image || typeof image !== 'object') continue;
      const src = typeof image.src === 'string' ? image.src : '';
      if (!parseImageDataUrl(src)) continue;
      images[key] = {
        src,
        width: Math.min(800, Math.max(1, Number(image.width) || 160)),
        height: Math.min(600, Math.max(1, Number(image.height) || 120)),
        alt: typeof image.alt === 'string' ? image.alt.slice(0, 300) : '',
      };
    }
    const name = typeof sheet?.name === 'string' && sheet.name.trim() ? sheet.name.trim() : `Sheet${index + 1}`;
    const columnWidths = {};
    for (const [column, width] of Object.entries(sheet?.columnWidths || {})) {
      if (/^[A-Z]{1,3}$/.test(column.toUpperCase()) && Number.isFinite(Number(width))) columnWidths[column.toUpperCase()] = Math.min(600, Math.max(40, Number(width)));
    }
    const rowHeights = {};
    for (const [row, height] of Object.entries(sheet?.rowHeights || {})) {
      if (/^[1-9]\d{0,3}$/.test(row) && Number.isFinite(Number(height))) rowHeights[row] = Math.min(300, Math.max(18, Number(height)));
    }
    const cellStyles = {};
    const rawStyles = sheet?.cellStyles && typeof sheet.cellStyles === 'object' ? sheet.cellStyles : {};
    for (const [ref, style] of Object.entries(rawStyles)) {
      const key = ref.toUpperCase();
      if (!REF_RE.test(key) || !style || typeof style !== 'object') continue;
      const fontFamily = typeof style.fontFamily === 'string' && /^[\w .,'-]{1,80}$/.test(style.fontFamily) ? style.fontFamily : '';
      const fontSize = Number(style.fontSize);
      if (fontFamily || (Number.isFinite(fontSize) && fontSize >= 6 && fontSize <= 96)) {
        cellStyles[key] = { ...(fontFamily ? { fontFamily } : {}), ...(Number.isFinite(fontSize) && fontSize >= 6 && fontSize <= 96 ? { fontSize } : {}) };
      }
    }
    return { name, cells, images, columnWidths, rowHeights, cellStyles };
  });

  if (!sheets.length) return blankSheet();
  const active = Number.isInteger(source.active) && source.active >= 0 && source.active < sheets.length ? source.active : 0;
  return { sheets, active };
}

/** Bounding box of used cells, or null when the sheet is empty. */
export function usedRange(cells) {
  let maxRow = 0;
  let maxCol = 0;
  let any = false;
  for (const ref of Object.keys(cells)) {
    const m = /^([A-Z]+)([0-9]+)$/.exec(ref);
    if (!m) continue;
    any = true;
    let col = 0;
    for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
    maxRow = Math.max(maxRow, Number(m[2]));
    maxCol = Math.max(maxCol, col);
  }
  return any ? { rows: maxRow, cols: maxCol } : null;
}
