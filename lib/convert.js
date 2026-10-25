// Format dispatch: every supported file type goes in and out through here.
import { docxToHtml, htmlToDocx, xlsxToSheets, sheetsToXlsx } from './ooxml.js';
import { odtToHtml, htmlToOdt, odsToSheets, sheetsToOds } from './odf.js';
import { rtfToText, textToHtml, htmlToText, htmlToRtf } from './rtf.js';
import { sanitizeHtml } from '../shared/sanitize.js';
import { parseCsv, toCsv } from '../shared/csv.js';
import { normalizeWord, normalizeSheet, normalizeImage, usedRange, extensionOf, kindOf } from '../shared/model.js';
import { imageDataUrl, parseImageDataUrl } from '../shared/images.js';
import { makeEngine, numToCol } from '../shared/formula.js';
import { docToHtml, xlsToSheets } from './legacy.js';
import { htmlToMarkdown, markdownToHtml } from './markdown.js';
import { PNG } from 'pngjs';
import * as UTIF from 'utif';
import { readPsd, writePsdBuffer } from 'ag-psd';
import { parseGIF, decompressFrames } from 'gifuct-js';
import gifenc from 'gifenc';
import sharp from 'sharp';
import { decodeBmp, encodeBmp, decodeIff, encodeIff } from './bitmap.js';

const { GIFEncoder, quantize, applyPalette } = gifenc;

const MIME = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  odt: 'application/vnd.oasis.opendocument.text',
  rtf: 'application/rtf',
  txt: 'text/plain; charset=utf-8',
  md: 'text/markdown; charset=utf-8',
  html: 'text/html; charset=utf-8',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  ods: 'application/vnd.oasis.opendocument.spreadsheet',
  csv: 'text/csv; charset=utf-8',
  doc: 'application/msword',
  dot: 'application/msword',
  xls: 'application/vnd.ms-excel',
  xlt: 'application/vnd.ms-excel',
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  bmp: 'image/bmp',
  iff: 'image/x-iff',
  lbm: 'image/x-iff',
  psd: 'image/vnd.adobe.photoshop',
  webp: 'image/webp',
  avif: 'image/avif',
  lipimg: 'application/vnd.lip-office.image+json',
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

function rgbaPngDataUrl({ width, height, data }) {
  if (!width || !height || width * height > 36000000 || data.length !== width * height * 4) throw new Error('image dimensions or pixel data are invalid');
  const png = new PNG({ width, height });
  png.data = Buffer.from(data);
  const encoded = PNG.sync.write(png);
  const src = imageDataUrl('image/png', encoded);
  if (!src) throw new Error('decoded image layer is larger than the 8 MB layer limit');
  return src;
}

function rgbaToImageDoc(name, pixels) {
  return normalizeImage({ width: 1200, height: 800, fitToImage: true, layers: [{ name, src: rgbaPngDataUrl(pixels) }] });
}

function readTiff(buffer) {
  const bytes = Buffer.from(buffer);
  const arrayBuffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  const pages = UTIF.decode(arrayBuffer);
  const page = pages?.[0];
  if (!page) throw new Error('TIFF contains no image pages');
  const width = Number(page.t256?.[0]), height = Number(page.t257?.[0]);
  if (!width || !height || width * height > 36000000) throw new Error('TIFF dimensions are invalid or too large');
  UTIF.decodeImage(arrayBuffer, page);
  return { width: page.width, height: page.height, data: UTIF.toRGBA8(page) };
}

function readGif(buffer) {
  const gif = parseGIF(Buffer.from(buffer));
  const width = Number(gif.lsd?.width), height = Number(gif.lsd?.height);
  if (!width || !height || width * height > 36000000) throw new Error('GIF dimensions are invalid or too large');
  const first = decompressFrames(gif, true)[0];
  if (!first?.patch) throw new Error('GIF contains no decodable frames');
  const data = new Uint8Array(width * height * 4);
  const { left = 0, top = 0, width: frameWidth, height: frameHeight } = first.dims;
  for (let y = 0; y < frameHeight; y++) {
    for (let x = 0; x < frameWidth; x++) {
      const sx = (y * frameWidth + x) * 4, dx = ((top + y) * width + left + x) * 4;
      if (top + y >= height || left + x >= width) continue;
      data.set(first.patch.subarray(sx, sx + 4), dx);
    }
  }
  return { width, height, data };
}

async function readSharpImage(buffer) {
  const input = sharp(buffer, { page: 0, pages: 1, animated: false });
  const metadata = await input.metadata();
  if (!metadata.width || !metadata.height || metadata.width * metadata.height > 36000000) throw new Error('image dimensions are invalid or too large');
  const { data, info } = await input.ensureAlpha().raw().toBuffer({ resolveWithObject: true });
  return { width: info.width, height: info.height, data };
}

function psdPixelData(imageData) {
  if (!imageData?.data || !imageData.width || !imageData.height) return null;
  const input = imageData.data;
  const rgba = new Uint8Array(imageData.width * imageData.height * 4);
  if (input.length !== rgba.length) return null;
  if (input instanceof Uint16Array) for (let i = 0; i < input.length; i++) rgba[i] = input[i] >> 8;
  else if (input instanceof Float32Array) for (let i = 0; i < input.length; i++) rgba[i] = Math.max(0, Math.min(255, Math.round(input[i] <= 1 ? input[i] * 255 : input[i])));
  else rgba.set(input);
  return { width: imageData.width, height: imageData.height, data: rgba };
}

function readPsdImage(buffer) {
  const header = readPsd(Buffer.from(buffer), { skipLayerImageData: true, skipCompositeImageData: true, skipThumbnail: true, logMissingFeatures: false });
  if (!header.width || !header.height || header.width * header.height > 36000000) throw new Error('PSD dimensions are invalid or too large');
  const psd = readPsd(Buffer.from(buffer), { useImageData: true, skipThumbnail: true, logMissingFeatures: false, totalMemoryLimit: 512 * 1024 * 1024 });
  const width = Number(psd.width), height = Number(psd.height);
  if (!width || !height || width * height > 36000000) throw new Error('PSD dimensions are invalid or too large');
  const layers = [];
  let totalLayerData = 0;
  const blend = { normal: 'source-over', multiply: 'multiply', screen: 'screen', overlay: 'overlay', darken: 'darken', lighten: 'lighten' };
  const collect = (nodes, parents = []) => {
    for (const node of nodes || []) {
      if (node.children?.length) collect(node.children, [...parents, node.name].filter(Boolean));
      else {
        if (layers.length >= 40) continue;
        const pixels = psdPixelData(node.imageData);
        if (!pixels) continue;
        const src = rgbaPngDataUrl(pixels);
        totalLayerData += src.length;
        if (totalLayerData > 24 * 1024 * 1024) throw new Error('PSD raster layers exceed the 24 MB editable image limit');
        layers.push({
          name: [...parents, node.name || 'Layer'].join(' / '), src,
          x: Number(node.left) || 0, y: Number(node.top) || 0,
          opacity: Number.isFinite(node.opacity) ? node.opacity : 1,
          visible: !node.hidden,
          blendMode: blend[node.blendMode] || 'source-over',
        });
      }
    }
  };
  collect(psd.children);
  layers.reverse(); // PSD stores topmost layers first; the editor paints bottom to top.
  if (!layers.length) {
    const composite = psdPixelData(psd.imageData);
    if (composite) layers.push({ name: 'PSD composite', src: rgbaPngDataUrl(composite), x: 0, y: 0, opacity: 1, visible: true });
  }
  if (!layers.length) throw new Error('PSD has no readable raster layers or composite image');
  return normalizeImage({ width, height, layers });
}

function writePsdImage(doc, composite) {
  const blend = { 'source-over': 'normal', multiply: 'multiply', screen: 'screen', overlay: 'overlay', darken: 'darken', lighten: 'lighten' };
  const children = [...doc.layers].reverse().filter((layer) => layer.src).map((layer) => {
    const parsed = parseImageDataUrl(layer.src);
    if (!parsed) throw new Error(`PSD layer “${layer.name}” has invalid image data`);
    const png = PNG.sync.read(Buffer.from(parsed.base64, 'base64'));
    return {
      name: layer.name, left: Math.round(layer.x), top: Math.round(layer.y), opacity: layer.opacity,
      hidden: !layer.visible, blendMode: blend[layer.blendMode] || 'normal',
      imageData: { width: png.width, height: png.height, data: new Uint8Array(png.data) },
    };
  });
  return writePsdBuffer({
    width: doc.width, height: doc.height,
    imageData: { width: composite.width, height: composite.height, data: new Uint8Array(composite.data) },
    children,
  }, { noBackground: true, generateThumbnail: false });
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
    const result = readDocument(name, buffer);
    return result && typeof result.then === 'function' ? result.catch((err) => {
      if (!err.status) err.status = 400;
      throw err;
    }) : result;
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
    if (extension === 'md') return wordDocument(markdownToHtml(buffer.toString('utf8')), base);
    return wordDocument(buffer.toString('utf8'), base); // html / htm
  }

  if (kind === 'image') {
    if (extension === 'lipimg') {
      try { return { kind: 'image', doc: normalizeImage(JSON.parse(buffer.toString('utf8'))) }; }
      catch (err) { throw new Error(err.message || 'invalid layered image project'); }
    }
    if (extension === 'psd') return { kind: 'image', doc: readPsdImage(buffer) };
    if (extension === 'iff' || extension === 'lbm') return { kind: 'image', doc: rgbaToImageDoc(base, decodeIff(buffer)) };
    if (extension === 'bmp') return { kind: 'image', doc: rgbaToImageDoc(base, decodeBmp(buffer)) };
    if (extension === 'tif' || extension === 'tiff') return { kind: 'image', doc: rgbaToImageDoc(base, readTiff(buffer)) };
    if (extension === 'gif') return { kind: 'image', doc: rgbaToImageDoc(base, readGif(buffer)) };
    if (extension === 'webp' || extension === 'avif') {
      return readSharpImage(buffer).then((pixels) => ({ kind: 'image', doc: rgbaToImageDoc(base, pixels) }));
    }
    const mime = MIME[extension];
    const src = imageDataUrl(mime, buffer);
    if (!src) throw new Error('Image file must be smaller than 8 MB.');
    return { kind: 'image', doc: normalizeImage({ width: 1200, height: 800, fitToImage: true, layers: [{ name: base, src }] }) };
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
  if (kind === 'image') {
    const normalized = normalizeImage(doc);
    if (format === 'lipimg') return { buffer: Buffer.from(JSON.stringify(normalized), 'utf8'), mime: MIME.lipimg };
    const image = parseImageDataUrl(doc?.exportDataUrl);
    if (!image) throw new Error(`cannot create .${format} image from the current export`);
    const encoded = Buffer.from(image.base64, 'base64');
    if (['png', 'jpg', 'jpeg'].includes(format)) {
      const expected = ['jpg', 'jpeg'].includes(format) ? 'image/jpeg' : 'image/png';
      if (image.mime !== expected) throw new Error(`cannot create .${format} image from the current export`);
      return { buffer: encoded, mime: MIME[format] };
    }
    const png = PNG.sync.read(encoded);
    if (format === 'gif') {
      const palette = quantize(png.data, 256, { format: 'rgba4444', oneBitAlpha: true, clearAlpha: true, clearAlphaThreshold: 127 });
      const indices = applyPalette(png.data, palette, 'rgba4444');
      const transparentIndex = palette.findIndex((color) => color.length > 3 && color[3] < 128);
      const gif = GIFEncoder();
      gif.writeFrame(indices, png.width, png.height, { palette, transparent: transparentIndex >= 0, transparentIndex: Math.max(0, transparentIndex) });
      gif.finish();
      return { buffer: Buffer.from(gif.bytes()), mime: MIME.gif };
    }
    if (format === 'tif' || format === 'tiff') return { buffer: Buffer.from(new Uint8Array(UTIF.encodeImage(png.data, png.width, png.height))), mime: MIME[format] };
    if (format === 'bmp') return { buffer: encodeBmp(png.data, png.width, png.height), mime: MIME.bmp };
    if (format === 'iff' || format === 'lbm') return { buffer: encodeIff(png.data, png.width, png.height, quantize, applyPalette), mime: MIME[format] };
    if (format === 'psd') return { buffer: writePsdImage({ ...normalized, layers: doc?.exportPsdLayers || normalized.layers }, png), mime: MIME.psd };
    if (format === 'webp' || format === 'avif') {
      const operation = sharp(encoded);
      return (format === 'webp' ? operation.webp({ quality: 90 }) : operation.avif({ quality: 85 })).toBuffer().then((buffer) => ({ buffer, mime: MIME[format] }));
    }
    throw new Error(`cannot create .${format} images`);
  }
  if (kind === 'word') {
    const normalized = normalizeWord({ ...(doc || {}), html: sanitizeHtml(doc?.html || ''), header: sanitizeHtml(doc?.header || ''), footer: sanitizeHtml(doc?.footer || '') });
    if (format === 'docx') return { buffer: htmlToDocx(normalized.html, normalized.title, normalized), mime: MIME.docx };
    if (format === 'odt') return { buffer: htmlToOdt(normalized.html, normalized.title, normalized), mime: MIME.odt };
    if (format === 'rtf') return { buffer: Buffer.from(htmlToRtf(normalized.html, normalized.title), 'latin1'), mime: MIME.rtf };
    if (format === 'txt') return { buffer: Buffer.from(htmlToText(normalized.html), 'utf8'), mime: MIME.txt };
    if (format === 'md') return { buffer: Buffer.from(htmlToMarkdown(normalized.html), 'utf8'), mime: MIME.md };
    return { buffer: Buffer.from(normalized.html, 'utf8'), mime: MIME.html };
  }

  const sheet = normalizeSheet(doc);
  if (format === 'xlsx') return { buffer: sheetsToXlsx(sheet), mime: MIME.xlsx };
  if (format === 'ods') return { buffer: sheetsToOds(sheet), mime: MIME.ods };
  return { buffer: Buffer.from(sheetToCsv(sheet), 'utf8'), mime: MIME.csv };
}
