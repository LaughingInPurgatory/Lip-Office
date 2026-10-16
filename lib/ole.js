// OLE2 / Compound File Binary container. Word 97-2003 .doc and legacy .xls
// are both CFB files: one header, a FAT, a directory tree, and streams that
// may live in the main file or in a "mini" stream for small ones.

const SIGNATURE = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;

function sectorAt(buffer, index, sectorSize) {
  return buffer.subarray((index + 1) * sectorSize, (index + 2) * sectorSize);
}

// Follow a FAT/miniFAT chain, stopping on ENDOFCHAIN or on a loop so a corrupt
// file cannot hang the reader.
function chain(next, start) {
  const sectors = [];
  const seen = new Set();
  while (start !== ENDOFCHAIN && start < next.length && !seen.has(start)) {
    seen.add(start);
    sectors.push(start);
    start = next[start];
  }
  return sectors;
}

function fatEntries(whole, sectors, sectorSize) {
  const fat = [];
  for (const index of sectors) {
    const data = sectorAt(whole, index, sectorSize);
    for (let i = 0; i + 4 <= data.length; i += 4) fat.push(data.readUInt32LE(i));
  }
  return fat;
}

export function readOle(buffer) {
  if (buffer.length < 512 || !buffer.subarray(0, 8).equals(SIGNATURE)) {
    throw new Error('not an OLE2 compound file');
  }
  const sectorSize = 1 << buffer.readUInt16LE(0x1e);
  const miniSize = 1 << buffer.readUInt16LE(0x20);
  const cutoff = buffer.readUInt32LE(0x38);

  // The header lists only 109 FAT sectors; after those, DIFAT sectors each
  // hold 127 more, then the number of the next DIFAT sector.
  const fatSectors = [];
  for (let i = 0; i < 109; i++) {
    const entry = buffer.readUInt32LE(0x4c + i * 4);
    if (entry === FREESECT) break;
    fatSectors.push(entry);
  }
  let difat = buffer.readUInt32LE(0x44);
  const difatCount = buffer.readUInt32LE(0x48);
  const seenDifat = new Set();
  for (let n = 0; n < difatCount && difat !== ENDOFCHAIN && difat !== FREESECT && !seenDifat.has(difat); n++) {
    seenDifat.add(difat);
    const entries = sectorAt(buffer, difat, sectorSize);
    if (entries.length < 4) break;
    for (let i = 0; i + 8 <= entries.length; i += 4) {
      const entry = entries.readUInt32LE(i);
      if (entry === FREESECT) break;
      fatSectors.push(entry);
    }
    difat = entries.readUInt32LE(entries.length - 4);
  }
  const fat = fatEntries(buffer, fatSectors, sectorSize);

  const directory = readChain(buffer, fat, buffer.readUInt32LE(0x30), sectorSize, Infinity, sectorSize);
  const entries = [];
  for (let offset = 0; offset + 128 <= directory.length; offset += 128) {
    const type = directory.readUInt8(offset + 0x42);
    if (type !== 1 && type !== 2 && type !== 5) continue;
    const nameLength = directory.readUInt16LE(offset + 0x40);
    entries.push({
      type,
      name: nameLength >= 2 ? directory.subarray(offset, offset + nameLength - 2).toString('utf16le') : '',
      start: directory.readUInt32LE(offset + 0x74),
      size: directory.readUInt32LE(offset + 0x78),
    });
  }

  // Everything below `cutoff` bytes is stored in the root entry's mini stream,
  // whose own sectors are chained through the miniFAT. Real .doc files keep
  // the piece table and formatting FKPs there.
  const root = entries.find((entry) => entry.type === 5);
  const miniFat = fatEntries(buffer, chain(fat, buffer.readUInt32LE(0x3c)), sectorSize);
  const miniStream = root && root.start !== ENDOFCHAIN
    ? readChain(buffer, fat, root.start, sectorSize, root.size, sectorSize)
    : Buffer.alloc(0);

  const streams = new Map();
  for (const entry of entries) {
    if (entry.type !== 2 || entry.size === 0) continue;
    streams.set(entry.name, entry.size < cutoff
      ? readChain(miniStream, miniFat, entry.start, miniSize, entry.size)
      : readChain(buffer, fat, entry.start, sectorSize, entry.size, sectorSize));
  }
  return streams;
}

function readChain(source, next, start, unit, size, headerSize = 0) {
  const out = Buffer.concat(chain(next, start).map((index) => source.subarray(headerSize + index * unit, headerSize + (index + 1) * unit)));
  return size < out.length ? out.subarray(0, size) : out;
}
