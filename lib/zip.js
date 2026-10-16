// Minimal ZIP reader/writer. Office and OpenDocument files are ZIP archives,
// so this is all we need to get at their XML payloads.
//
// shortcut: no ZIP64 and no encrypted archives, because office documents are
// never that large. Upgrade if a >4GB or password-protected file shows up.
import zlib from 'node:zlib';

const CRC_TABLE = (() => {
  const table = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c;
  }
  return table;
})();

function crc32(buf) {
  let c = -1;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

const EOCD = 0x06054b50;
const CEN = 0x02014b50;
const LOC = 0x04034b50;
const MAX_ENTRY_SIZE = 128 * 1024 * 1024;
const MAX_TOTAL_SIZE = 256 * 1024 * 1024;

function findEocd(buf) {
  if (buf.length < 22) throw new Error('not a ZIP archive');
  const floor = Math.max(0, buf.length - 22 - 65535);
  for (let i = buf.length - 22; i >= floor; i--) {
    if (buf.readUInt32LE(i) === EOCD) return i;
  }
  throw new Error('not a ZIP archive');
}

/** @returns {Map<string, Buffer>} entry name -> decompressed bytes */
export function unzip(buf) {
  const eocd = findEocd(buf);
  const count = buf.readUInt16LE(eocd + 10);
  const entries = new Map();
  let totalSize = 0;
  let p = buf.readUInt32LE(eocd + 16);

  for (let i = 0; i < count; i++) {
    if (p + 46 > buf.length || buf.readUInt32LE(p) !== CEN) throw new Error('malformed ZIP directory');
    const method = buf.readUInt16LE(p + 10);
    const checksum = buf.readUInt32LE(p + 16);
    const compSize = buf.readUInt32LE(p + 20);
    const rawSize = buf.readUInt32LE(p + 24);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commentLen = buf.readUInt16LE(p + 32);
    const localOff = buf.readUInt32LE(p + 42);
    const name = buf.toString('utf8', p + 46, p + 46 + nameLen);

    if (rawSize > MAX_ENTRY_SIZE || totalSize + rawSize > MAX_TOTAL_SIZE) throw new Error('ZIP archive expands beyond the size limit');
    if (method !== 0 && method !== 8) throw new Error(`unsupported ZIP compression method: ${method}`);
    if (localOff + 30 > buf.length || buf.readUInt32LE(localOff) !== LOC) throw new Error('malformed ZIP entry');
    const start = localOff + 30 + buf.readUInt16LE(localOff + 26) + buf.readUInt16LE(localOff + 28);
    if (start > buf.length || compSize > buf.length - start) throw new Error('truncated ZIP entry');
    const data = buf.subarray(start, start + compSize);
    const output = method === 0
      ? Buffer.from(data)
      : zlib.inflateRawSync(data, { maxOutputLength: Math.max(1, rawSize) });
    if (output.length !== rawSize || crc32(output) !== checksum) throw new Error('corrupt ZIP entry');
    totalSize += output.length;
    entries.set(name, output);
    p += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

/**
 * @param {Array<{name: string, data: Buffer|string, store?: boolean}>} files
 * `store: true` writes uncompressed, which the ODF spec demands for `mimetype`.
 */
export function zip(files) {
  const parts = [];
  const central = [];
  let offset = 0;

  for (const file of files) {
    const raw = Buffer.isBuffer(file.data) ? file.data : Buffer.from(file.data, 'utf8');
    const name = Buffer.from(file.name, 'utf8');
    const crc = crc32(raw);

    let method = 0;
    let body = raw;
    if (!file.store) {
      const deflated = zlib.deflateRawSync(raw);
      if (deflated.length < raw.length) {
        method = 8;
        body = deflated;
      }
    }

    const local = Buffer.alloc(30);
    local.writeUInt32LE(LOC, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(method, 8);
    local.writeUInt16LE(0x21, 12); // 1980-01-01
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(body.length, 18);
    local.writeUInt32LE(raw.length, 22);
    local.writeUInt16LE(name.length, 26);
    parts.push(local, name, body);

    const cd = Buffer.alloc(46);
    cd.writeUInt32LE(CEN, 0);
    cd.writeUInt16LE(20, 4);
    cd.writeUInt16LE(20, 6);
    cd.writeUInt16LE(method, 10);
    cd.writeUInt16LE(0x21, 14);
    cd.writeUInt32LE(crc, 16);
    cd.writeUInt32LE(body.length, 20);
    cd.writeUInt32LE(raw.length, 24);
    cd.writeUInt16LE(name.length, 28);
    cd.writeUInt32LE(offset, 42);
    central.push(cd, name);

    offset += 30 + name.length + body.length;
  }

  const dir = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(EOCD, 0);
  end.writeUInt16LE(files.length, 8);
  end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(dir.length, 12);
  end.writeUInt32LE(offset, 16);

  return Buffer.concat([...parts, dir, end]);
}
