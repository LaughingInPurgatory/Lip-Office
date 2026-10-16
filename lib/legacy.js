// Read the text and cell data that older Word and Excel files keep inside OLE.
// Formatting and embedded objects are intentionally left to the newer formats.
import { readOle } from './ole.js';
import { numToCol } from '../shared/formula.js';
import { escapeHtml } from '../shared/html.js';

const CP1252 = new TextDecoder('windows-1252');

function safeSlice(buffer, offset, length) {
  if (!Number.isInteger(offset) || !Number.isInteger(length) || offset < 0 || length < 0 || offset + length > buffer.length) {
    throw new Error('malformed legacy Office document');
  }
  return buffer.subarray(offset, offset + length);
}

function decodeFcText(word, fc, count) {
  const compressed = Boolean(fc & 0x40000000);
  const offset = (fc & 0x3fffffff) / (compressed ? 2 : 1);
  if (compressed) return CP1252.decode(safeSlice(word, offset, count));
  return safeSlice(word, offset, count * 2).toString('utf16le');
}

/** Extract the main text story using the Word 97+ FIB and CLX piece table. */
export function docToText(buffer) {
  const streams = readOle(buffer);
  const word = streams.get('WordDocument');
  if (!word || word.length < 34 || ![0xa5dc, 0xa5ec].includes(word.readUInt16LE(0))) throw new Error('not a supported Word binary document');
  const version = word.readUInt16LE(2);
  const flags = word.readUInt16LE(10);
  if (flags & 0x0100) throw new Error('encrypted Word documents are not supported');

  // Word 6/95 stored the main text contiguously between fcMin and fcMac.
  if (version < 0x00c1) {
    const fcMin = word.readUInt32LE(0x18);
    const fcMac = word.readUInt32LE(0x1c);
    if (fcMac > fcMin && fcMac <= word.length) {
      const text = CP1252.decode(word.subarray(fcMin, fcMac)).replace(/\u0000/g, '');
      if (text) return text;
    }
  }

  let offset = 32;
  const csw = safeSlice(word, offset, 2).readUInt16LE(0);
  offset += 2 + csw * 2;
  const cslw = safeSlice(word, offset, 2).readUInt16LE(0);
  offset += 2;
  const rgLw = safeSlice(word, offset, cslw * 4);
  offset += cslw * 4;
  const ccpText = cslw > 3 ? rgLw.readUInt32LE(12) : 0;
  const cbRgFcLcb = safeSlice(word, offset, 2).readUInt16LE(0);
  offset += 2;
  const pairs = safeSlice(word, offset, cbRgFcLcb * 8);
  // fcClx/lcbClx are pair 33 in FibRgFcLcb97.
  if (cbRgFcLcb <= 33) throw new Error('Word document has no supported text piece table');
  const fcClx = pairs.readUInt32LE(33 * 8);
  const lcbClx = pairs.readUInt32LE(33 * 8 + 4);
  const table = streams.get(flags & 0x0200 ? '1Table' : '0Table');
  if (!table || !lcbClx) throw new Error('Word document has no supported text piece table');

  let clx = safeSlice(table, fcClx, lcbClx);
  let clxOffset = 0;
  while (clxOffset < clx.length && clx[clxOffset] === 0x01) {
    const size = safeSlice(clx, clxOffset + 1, 2).readUInt16LE(0);
    clxOffset += 3 + size;
  }
  if (clx[clxOffset] !== 0x02) throw new Error('malformed Word text piece table');
  const plcLength = safeSlice(clx, clxOffset + 1, 4).readUInt32LE(0);
  const plc = safeSlice(clx, clxOffset + 5, plcLength);
  if (plc.length < 16 || (plc.length - 4) % 12 !== 0) throw new Error('malformed Word text piece table');

  const pieces = (plc.length - 4) / 12;
  const cpBytes = (pieces + 1) * 4;
  const text = [];
  for (let i = 0; i < pieces; i++) {
    const cpStart = plc.readUInt32LE(i * 4);
    const cpEnd = plc.readUInt32LE((i + 1) * 4);
    if (cpEnd < cpStart) throw new Error('malformed Word text piece table');
    const pcd = cpBytes + i * 8;
    const fc = safeSlice(plc, pcd + 2, 4).readUInt32LE(0);
    text.push(decodeFcText(word, fc, cpEnd - cpStart));
  }
  const result = text.join('').slice(0, ccpText || undefined);
  if (!result) throw new Error('Word document contains no readable text');
  return result;
}

export function docToHtml(buffer) {
  const text = docToText(buffer).replace(/\u0000/g, '');
  const paragraphs = text.split(/\r|\n/).map((part) => {
    const content = part.replace(/\u0007/g, '\t').replace(/\u000b/g, '\n');
    return `<p>${content ? escapeHtml(content).replace(/\n/g, '<br>') : '<br>'}</p>`;
  });
  return paragraphs.join('');
}

function records(stream, start = 0) {
  const out = [];
  for (let offset = start; offset + 4 <= stream.length;) {
    const id = stream.readUInt16LE(offset);
    const length = stream.readUInt16LE(offset + 2);
    const dataStart = offset + 4;
    if (dataStart + length > stream.length) throw new Error('malformed Excel BIFF record');
    out.push({ id, data: stream.subarray(dataStart, dataStart + length), offset });
    offset = dataStart + length;
    if (id === 0x000a) break;
  }
  return out;
}

function parseBiffString(data, offset = 0) {
  if (offset + 3 > data.length) return null;
  const chars = data.readUInt16LE(offset);
  const flags = data[offset + 2];
  offset += 3;
  const richRuns = flags & 0x08 ? data.readUInt16LE(offset) : 0;
  if (flags & 0x08) offset += 2;
  const extensionSize = flags & 0x04 ? data.readUInt32LE(offset) : 0;
  if (flags & 0x04) offset += 4;
  const byteLength = chars * (flags & 0x01 ? 2 : 1);
  if (offset + byteLength + richRuns * 4 + extensionSize > data.length) return null;
  const value = flags & 0x01
    ? data.toString('utf16le', offset, offset + byteLength)
    : CP1252.decode(data.subarray(offset, offset + byteLength));
  return { value, next: offset + byteLength + richRuns * 4 + extensionSize };
}

function parseSst(records, start) {
  const chunks = [records[start].data];
  for (let i = start + 1; i < records.length && records[i].id === 0x003c; i++) chunks.push(records[i].data);
  let chunk = 0;
  let offset = 0;
  const advance = () => {
    while (chunk < chunks.length && offset >= chunks[chunk].length) { chunk += 1; offset = 0; }
    return chunk < chunks.length;
  };
  const byte = () => {
    if (!advance()) return null;
    return chunks[chunk][offset++];
  };
  const integer = (size) => {
    let value = 0;
    for (let i = 0; i < size; i++) {
      const next = byte();
      if (next === null) return null;
      value += next * (2 ** (8 * i));
    }
    return value;
  };
  const skip = (size) => { while (size-- > 0 && byte() !== null) { /* consume */ } };

  const total = integer(4);
  const unique = integer(4);
  if (total === null || unique === null) return [];
  const count = Math.min(unique, 1_000_000);
  const strings = [];
  for (let i = 0; i < count; i++) {
    const chars = integer(2);
    let flags = byte();
    if (chars === null || flags === null) break;
    const richRuns = flags & 0x08 ? integer(2) : 0;
    const extensionSize = flags & 0x04 ? integer(4) : 0;
    if (richRuns === null || extensionSize === null) break;
    let remaining = Math.min(chars, 1_000_000);
    let value = '';
    while (remaining > 0) {
      if (chunk >= chunks.length) break;
      if (offset >= chunks[chunk].length) {
        chunk += 1;
        offset = 0;
        if (!advance()) break;
        flags = byte(); // continuation character data begins with its encoding flag
        if (flags === null) break;
      }
      const width = flags & 1 ? 2 : 1;
      if (offset + width > chunks[chunk].length) break;
      value += width === 2
        ? chunks[chunk].toString('utf16le', offset, offset + 2)
        : CP1252.decode(chunks[chunk].subarray(offset, offset + 1));
      offset += width;
      remaining -= 1;
    }
    strings.push(value);
    skip((richRuns || 0) * 4 + (extensionSize || 0));
  }
  return strings;
}

function decodeRk(data, offset) {
  const rk = data.readUInt32LE(offset);
  let value;
  if (rk & 0x02) value = rk >> 2;
  else {
    const bits = Buffer.alloc(8);
    bits.writeUInt32LE(rk & 0xfffffffc, 4);
    value = bits.readDoubleLE(0);
  }
  return rk & 0x01 ? value / 100 : value;
}

function addCell(cells, row, col, value) {
  if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0 || col >= 256) return;
  const text = String(value);
  if (text !== '') cells[`${numToCol(col + 1)}${row + 1}`] = text;
}

function sheetCells(stream, start, shared, biff8) {
  const cells = {};
  let pendingStringCell = null;
  for (const { id, data } of records(stream, start)) {
    if (id === 0x0207 && pendingStringCell) {
      const parsed = parseBiffString(data);
      if (parsed) addCell(cells, pendingStringCell.row, pendingStringCell.col, parsed.value);
      pendingStringCell = null;
      continue;
    }
    if (id === 0x00fd && data.length >= 10) {
      addCell(cells, data.readUInt16LE(0), data.readUInt16LE(2), shared[data.readUInt32LE(6)] ?? '');
    } else if (id === 0x0203 && data.length >= 14) {
      addCell(cells, data.readUInt16LE(0), data.readUInt16LE(2), data.readDoubleLE(6));
    } else if (id === 0x027e && data.length >= 10) {
      addCell(cells, data.readUInt16LE(0), data.readUInt16LE(2), decodeRk(data, 6));
    } else if (id === 0x00bd && data.length >= 12) {
      const row = data.readUInt16LE(0);
      const first = data.readUInt16LE(2);
      const last = data.readUInt16LE(data.length - 2);
      for (let col = first, offset = 4; col <= last && offset + 6 <= data.length - 2; col++, offset += 6) {
        addCell(cells, row, col, decodeRk(data, offset + 2));
      }
    } else if (id === 0x0204 && data.length >= 9) {
      const chars = data.readUInt16LE(6);
      const flags = biff8 ? data[8] : 0;
      const textStart = biff8 ? 9 : 8;
      const size = chars * (flags & 1 ? 2 : 1);
      if (textStart + size <= data.length) {
        const value = flags & 1 ? data.toString('utf16le', textStart, textStart + size) : CP1252.decode(data.subarray(textStart, textStart + size));
        addCell(cells, data.readUInt16LE(0), data.readUInt16LE(2), value);
      }
    } else if (id === 0x0205 && data.length >= 8) {
      const value = data[7] ? `#ERR${data[6]}` : data[6] ? 'TRUE' : 'FALSE';
      addCell(cells, data.readUInt16LE(0), data.readUInt16LE(2), value);
    } else if (id === 0x0006 && data.length >= 14) {
      const row = data.readUInt16LE(0);
      const col = data.readUInt16LE(2);
      if (data[12] === 0xff && data[13] === 0xff) {
        if (data[6] === 1) addCell(cells, row, col, data[8] ? 'TRUE' : 'FALSE');
        else if (data[6] === 2) addCell(cells, row, col, `#ERR${data[8]}`);
        else if (data[6] === 0) pendingStringCell = { row, col };
      } else addCell(cells, row, col, data.readDoubleLE(6));
    }
  }
  return cells;
}

/** Read common BIFF5/BIFF8 value records from an OLE Excel workbook stream. */
export function xlsToSheets(buffer) {
  const streams = readOle(buffer);
  const workbook = streams.get('Workbook') || streams.get('Book');
  if (!workbook) throw new Error('not a supported Excel binary workbook');
  const globalRecords = records(workbook);
  const bof = globalRecords.find((record) => record.id === 0x0809 || record.id === 0x0409);
  const biff8 = Boolean(bof && bof.data.length >= 2 && bof.data.readUInt16LE(0) >= 0x0600);
  const sheets = [];
  let shared = [];

  for (let i = 0; i < globalRecords.length; i++) {
    const { id, data } = globalRecords[i];
    if (id === 0x00fc) shared = parseSst(globalRecords, i);
    else if (id === 0x0085 && data.length >= 8) {
      const flags = biff8 ? data[7] : 0;
      const chars = data[6];
      const size = chars * (flags & 1 ? 2 : 1);
      const nameStart = biff8 ? 8 : 7;
      if (nameStart + size <= data.length) {
        const name = flags & 1 ? data.toString('utf16le', nameStart, nameStart + size) : CP1252.decode(data.subarray(nameStart, nameStart + size));
        sheets.push({ offset: data.readUInt32LE(0), name, kind: data[5] });
      }
    }
  }

  const result = sheets.filter((sheet) => sheet.kind === 0).map((sheet) => ({
    name: sheet.name,
    cells: sheetCells(workbook, sheet.offset, shared, biff8),
  }));
  if (!result.length) throw new Error('Excel workbook contains no readable worksheets');
  return { sheets: result, active: 0 };
}
