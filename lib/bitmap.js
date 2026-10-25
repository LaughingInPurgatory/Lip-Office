// Small BMP and Amiga IFF ILBM readers/writers used at the format boundary.
// Pixel editing stays in the layered canvas model; these codecs only translate
// flattened RGBA exports.

export function decodeBmp(buffer) {
  const data = Buffer.from(buffer);
  if (data.length < 54 || data.toString('ascii', 0, 2) !== 'BM') throw new Error('invalid BMP image');
  const pixelOffset = data.readUInt32LE(10);
  const headerSize = data.readUInt32LE(14);
  if (headerSize < 40 || data.length < 14 + headerSize) throw new Error('unsupported BMP header');
  const rawWidth = data.readInt32LE(18), rawHeight = data.readInt32LE(22);
  const width = Math.abs(rawWidth), height = Math.abs(rawHeight), topDown = rawHeight < 0;
  const bits = data.readUInt16LE(28), compression = data.readUInt32LE(30);
  if (!width || !height || width * height > 36000000) throw new Error('BMP dimensions are invalid or too large');
  if (data.readUInt16LE(26) !== 1 || ![8, 24, 32].includes(bits) || compression !== 0) throw new Error('BMP must be uncompressed 8-bit, 24-bit, or 32-bit RGB');
  const palette = [];
  if (bits === 8) {
    const colors = data.readUInt32LE(46) || 256;
    const paletteStart = 14 + headerSize;
    if (paletteStart + colors * 4 > pixelOffset) throw new Error('invalid BMP palette');
    for (let i = 0; i < colors; i++) palette.push([data[paletteStart + i * 4 + 2], data[paletteStart + i * 4 + 1], data[paletteStart + i * 4]]);
  }
  const stride = Math.ceil(width * bits / 32) * 4;
  if (pixelOffset + stride * height > data.length) throw new Error('truncated BMP pixel data');
  const rgba = new Uint8Array(width * height * 4);
  for (let y = 0; y < height; y++) {
    const sourceY = topDown ? y : height - 1 - y;
    for (let x = 0; x < width; x++) {
      const target = (y * width + x) * 4, source = pixelOffset + sourceY * stride + x * (bits / 8);
      if (bits === 8) {
        const color = palette[data[source]];
        if (!color) throw new Error('invalid BMP palette index');
        rgba[target] = color[0]; rgba[target + 1] = color[1]; rgba[target + 2] = color[2]; rgba[target + 3] = 255;
      } else {
        rgba[target] = data[source + 2]; rgba[target + 1] = data[source + 1]; rgba[target + 2] = data[source];
        rgba[target + 3] = 255;
      }
    }
  }
  return { width, height, data: rgba };
}

export function encodeBmp(rgba, width, height) {
  const stride = Math.ceil(width * 24 / 32) * 4;
  const pixelsSize = stride * height;
  const out = Buffer.alloc(54 + pixelsSize);
  out.write('BM', 0, 'ascii'); out.writeUInt32LE(out.length, 2); out.writeUInt32LE(54, 10);
  out.writeUInt32LE(40, 14); out.writeInt32LE(width, 18); out.writeInt32LE(height, 22);
  out.writeUInt16LE(1, 26); out.writeUInt16LE(24, 28); out.writeUInt32LE(0, 30);
  out.writeUInt32LE(pixelsSize, 34); out.writeInt32LE(2835, 38); out.writeInt32LE(2835, 42);
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
    const source = ((height - 1 - y) * width + x) * 4, target = 54 + y * stride + x * 3;
    const alpha = rgba[source + 3] / 255;
    out[target] = Math.round(rgba[source + 2] * alpha + 255 * (1 - alpha));
    out[target + 1] = Math.round(rgba[source + 1] * alpha + 255 * (1 - alpha));
    out[target + 2] = Math.round(rgba[source] * alpha + 255 * (1 - alpha));
  }
  return out;
}

function chunk(id, bytes) {
  const header = Buffer.alloc(8); header.write(id, 0, 4, 'ascii'); header.writeUInt32BE(bytes.length, 4);
  return bytes.length % 2 ? Buffer.concat([header, bytes, Buffer.alloc(1)]) : Buffer.concat([header, bytes]);
}

function decodeByteRun1(bytes, expected) {
  const out = new Uint8Array(expected);
  let input = 0, output = 0;
  while (input < bytes.length && output < expected) {
    const control = bytes.readInt8(input++);
    if (control >= 0) {
      const count = control + 1;
      if (input + count > bytes.length || output + count > expected) throw new Error('invalid IFF ByteRun1 data');
      out.set(bytes.subarray(input, input + count), output); input += count; output += count;
    } else if (control !== -128) {
      const count = 1 - control;
      if (input >= bytes.length || output + count > expected) throw new Error('invalid IFF ByteRun1 data');
      out.fill(bytes[input++], output, output + count); output += count;
    }
  }
  if (output !== expected) throw new Error('truncated IFF image data');
  return out;
}

export function decodeIff(buffer) {
  const data = Buffer.from(buffer);
  if (data.length < 12 || data.toString('ascii', 0, 4) !== 'FORM') throw new Error('invalid Amiga IFF image');
  const form = data.toString('ascii', 8, 12);
  if (!['ILBM', 'PBM '].includes(form)) throw new Error('IFF file is not an ILBM or PBM image');
  let bmhd, cmap, body;
  for (let at = 12; at + 8 <= data.length;) {
    const id = data.toString('ascii', at, at + 4), size = data.readUInt32BE(at + 4), start = at + 8, end = start + size;
    if (end > data.length) throw new Error('truncated IFF chunk');
    if (id === 'BMHD') bmhd = data.subarray(start, end);
    else if (id === 'CMAP') cmap = data.subarray(start, end);
    else if (id === 'BODY') body = data.subarray(start, end);
    at = end + (size & 1);
  }
  if (!bmhd || bmhd.length < 20 || !body) throw new Error('IFF is missing its bitmap header or body');
  const width = bmhd.readUInt16BE(0), height = bmhd.readUInt16BE(2), planes = bmhd[8], masking = bmhd[9], compression = bmhd[10], transparent = bmhd.readUInt16BE(12);
  if (!width || !height || width * height > 36000000 || planes < 1 || planes > 8 && !(form === 'PBM ' && planes === 24)) throw new Error('IFF image dimensions or colour depth are not supported');
  if (compression > 1 || masking > 2) throw new Error('IFF compression or transparency mode is not supported');
  const palette = [];
  if (cmap) for (let i = 0; i + 2 < cmap.length; i += 3) palette.push([cmap[i], cmap[i + 1], cmap[i + 2]]);
  if (!palette.length) for (let i = 0; i < 1 << planes; i++) { const g = Math.round(i * 255 / ((1 << planes) - 1)); palette.push([g, g, g]); }
  const rowBytes = Math.ceil(width / 16) * 2;
  let rgba = new Uint8Array(width * height * 4);
  if (form === 'PBM ') {
    if (planes === 24) {
      const stride = (width * 3 + 1) & ~1;
      const raw = compression ? decodeByteRun1(body, stride * height) : body;
      if (raw.length < stride * height) throw new Error('truncated IFF 24-bit PBM body');
      for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
        const source = y * stride + x * 3, target = (y * width + x) * 4;
        rgba[target] = raw[source]; rgba[target + 1] = raw[source + 1]; rgba[target + 2] = raw[source + 2]; rgba[target + 3] = 255;
      }
      return { width, height, data: rgba };
    }
    const stride = (width + 1) & ~1;
    const raw = compression ? decodeByteRun1(body, stride * height) : body;
    if (raw.length < stride * height) throw new Error('truncated IFF PBM body');
    for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
      const index = raw[y * stride + x], color = palette[index] || [0, 0, 0], p = (y * width + x) * 4;
      rgba[p] = color[0]; rgba[p + 1] = color[1]; rgba[p + 2] = color[2]; rgba[p + 3] = masking === 2 && index === transparent ? 0 : 255;
    }
  } else {
    const storedPlanes = planes + (masking === 1 ? 1 : 0), expected = rowBytes * height * storedPlanes;
    const raw = compression ? decodeByteRun1(body, expected) : body;
    if (raw.length < expected) throw new Error('truncated IFF ILBM body');
    for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
      let index = 0, alpha = 255;
      for (let plane = 0; plane < planes; plane++) {
        const value = raw[(y * storedPlanes + plane) * rowBytes + (x >> 3)] >> (7 - (x & 7)) & 1;
        index |= value << plane;
      }
      if (masking === 1) alpha = (raw[(y * storedPlanes + planes) * rowBytes + (x >> 3)] >> (7 - (x & 7)) & 1) ? 255 : 0;
      else if (masking === 2 && index === transparent) alpha = 0;
      const halfBright = planes === 6 && palette.length === 32 && index >= 32;
      let color = palette[halfBright ? index - 32 : index] || [0, 0, 0];
      if (halfBright) color = color.map((channel) => Math.round(channel / 2));
      const p = (y * width + x) * 4; rgba[p] = color[0]; rgba[p + 1] = color[1]; rgba[p + 2] = color[2]; rgba[p + 3] = alpha;
    }
  }
  return { width, height, data: rgba };
}

export function encodeIff(rgba, width, height, quantize, applyPalette) {
  if (width > 65535 || height > 65535) throw new Error('IFF supports images up to 65,535 pixels per side');
  const palette = quantize(rgba, 256, { format: 'rgba4444', oneBitAlpha: true, clearAlpha: true, clearAlphaThreshold: 127 });
  const indices = applyPalette(rgba, palette, 'rgba4444');
  const rowBytes = Math.ceil(width / 16) * 2, planes = 8, masking = 1;
  const header = Buffer.alloc(20);
  header.writeUInt16BE(width, 0); header.writeUInt16BE(height, 2); header.writeInt16BE(0, 4); header.writeInt16BE(0, 6);
  header[8] = planes; header[9] = masking; header[10] = 0; header[11] = 0; header.writeUInt16BE(0, 12);
  header[14] = 1; header[15] = 1; header.writeUInt16BE(width, 16); header.writeUInt16BE(height, 18);
  const colors = Buffer.alloc(256 * 3);
  palette.forEach((color, index) => { colors[index * 3] = color[0]; colors[index * 3 + 1] = color[1]; colors[index * 3 + 2] = color[2]; });
  const body = Buffer.alloc(rowBytes * height * (planes + 1));
  for (let y = 0; y < height; y++) for (let plane = 0; plane < planes + 1; plane++) for (let x = 0; x < width; x++) {
    const alpha = rgba[(y * width + x) * 4 + 3];
    const set = plane === planes ? alpha >= 128 : ((indices[y * width + x] >> plane) & 1) !== 0;
    if (set) body[(y * (planes + 1) + plane) * rowBytes + (x >> 3)] |= 1 << (7 - (x & 7));
  }
  const chunks = Buffer.concat([chunk('BMHD', header), chunk('CMAP', colors), chunk('BODY', body)]);
  const out = Buffer.alloc(12 + chunks.length); out.write('FORM', 0, 'ascii'); out.writeUInt32BE(4 + chunks.length, 4); out.write('ILBM', 8, 'ascii'); chunks.copy(out, 12);
  return out;
}
