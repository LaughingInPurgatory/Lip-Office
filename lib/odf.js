// OpenDocument (.odt / .ods) import and export — what LibreOffice writes.
import { unzip, zip } from './zip.js';
import { parseMarkup, findAll, find, textOf, tableRows, tableCells, escapeXml } from '../shared/html.js';
import { makeEngine, colToNum, numToCol } from '../shared/formula.js';
import { MAX_SHEET_ROWS, MAX_SHEET_COLS } from '../shared/model.js';
import { imageDataUrl, parseImageDataUrl } from '../shared/images.js';

const NS = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
  draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
  number: 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
  meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  dc: 'http://purl.org/dc/elements/1.1/',
  xlink: 'http://www.w3.org/1999/xlink',
  manifest: 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0',
};

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8"?>\n';

function nsAttrs(keys) {
  return keys.map((key) => `xmlns:${key}="${NS[key]}"`).join(' ');
}

function odfParts(kind, mediaFiles = []) {
  const media = kind === 'text'
    ? 'application/vnd.oasis.opendocument.text'
    : 'application/vnd.oasis.opendocument.spreadsheet';
  // `mimetype` must be first and uncompressed, or readers reject the file.
  return [
    { name: 'mimetype', data: media, store: true },
    {
      name: 'META-INF/manifest.xml',
      data: `${XML_HEAD}<manifest:manifest xmlns:manifest="${NS.manifest}" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="${media}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>${mediaFiles.map((item) => `<manifest:file-entry manifest:full-path="${item.path}" manifest:media-type="${item.mime}"/>`).join('')}</manifest:manifest>`,
    },
  ];
}

function metaXml(title) {
  const stamp = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  return `${XML_HEAD}<office:document-meta ${nsAttrs(['office', 'meta', 'dc'])} office:version="1.2"><office:meta><dc:title>${escapeXml(title)}</dc:title><meta:creation-date>${stamp}</meta:creation-date></office:meta></office:document-meta>`;
}

function readOdfTitle(files) {
  const meta = files.get('meta.xml');
  if (!meta) return '';
  const title = find(parseMarkup(meta.toString('utf8')), 'dc:title');
  return title ? textOf(title).trim() : '';
}

function odfInches(value, fallback) {
  const match = /^\s*([\d.]+)\s*(in|cm|mm|pt)?\s*$/i.exec(String(value || ''));
  if (!match) return fallback;
  const number = Number(match[1]);
  const unit = (match[2] || 'in').toLowerCase();
  return number * ({ in: 1, cm: 1 / 2.54, mm: 1 / 25.4, pt: 1 / 72 })[unit];
}

// -------------------------------------------------------------- .odt reading

/** Automatic text/paragraph styles as plain formatting records. */
function readTextStyles(section) {
  const styles = new Map();
  if (!section) return styles;
  for (const style of section.children) {
    if (typeof style === 'string' || style.tag !== 'style:style') continue;
    const name = style.attrs['style:name'];
    if (!name) continue;
    const props = {};

    const text = find(style, 'style:text-properties');
    if (text) {
      if (/^(bold|[6-9]00)$/i.test(text.attrs['fo:font-weight'] || '')) props.b = true;
      if (/^italic$/i.test(text.attrs['fo:font-style'] || '')) props.i = true;
      if ((text.attrs['style:text-underline-style'] || 'none') !== 'none') props.u = true;
      if (/line-through/.test(text.attrs['style:text-line-through-style'] || '')) props.strike = true;
      const color = text.attrs['fo:color'];
      if (color && /^#[0-9a-fA-F]{6}$/.test(color)) props.color = color.toLowerCase();
      const size = parseFloat(text.attrs['fo:font-size']);
      if (Number.isFinite(size)) props.size = size;
      const position = text.attrs['style:text-position'] || '';
      if (position.startsWith('super')) props.vert = 'superscript';
      else if (position.startsWith('sub')) props.vert = 'subscript';
    }

    const para = find(style, 'style:paragraph-properties');
    if (para && para.attrs['fo:text-align']) {
      const align = para.attrs['fo:text-align'];
      if (['center', 'right', 'justify'].includes(align)) props.align = align === 'justify' ? 'justify' : align;
    }
    styles.set(name, props);
  }
  return styles;
}

function odfInline(nodes, styles, files) {
  let html = '';
  for (const node of nodes) {
    if (typeof node === 'string') { html += node; continue; }
    switch (node.tag) {
      case 'text:span': html += wrap(String(node.attrs['text:style-name']), styles, odfInline(node.children, styles, files)); break;
      case 'text:a': html += `<a href="${escapeXml(node.attrs['xlink:href'] || '')}">${odfInline(node.children, styles, files)}</a>`; break;
      case 'text:s': html += ' '.repeat(Math.max(1, Number(node.attrs['text:c'] || 1))); break;
      case 'text:tab': html += '\t'; break;
      case 'text:line-break': html += '<br>'; break;
      case 'text:p': html += odfInline(node.children, styles, files); break;
      case 'draw:frame': {
        const image = find(node, 'draw:image');
        const href = image?.attrs['xlink:href'];
        const bytes = href && files.get(href);
        if (bytes) {
          const ext = href.split('.').pop().toLowerCase();
          const mime = ({ png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif' })[ext];
          const src = imageDataUrl(mime, bytes);
          if (src) html += `<img src="${src}" width="${Math.max(1, Math.round(parseFloat(node.attrs['svg:width']) || 160))}" height="${Math.max(1, Math.round(parseFloat(node.attrs['svg:height']) || 100))}">`;
        }
        break;
      }
      default: html += odfInline(node.children, styles, files); break;
    }
  }
  return html;
}

function wrap(styleName, styles, inner) {
  const props = styles.get(styleName);
  if (!props) return inner;
  let open = '';
  let close = '';
  if (props.b) { open += '<strong>'; close = `</strong>${close}`; }
  if (props.i) { open += '<em>'; close = `</em>${close}`; }
  if (props.u) { open += '<u>'; close = `</u>${close}`; }
  if (props.strike) { open += '<s>'; close = `</s>${close}`; }
  if (props.vert === 'superscript') { open += '<sup>'; close = `</sup>${close}`; }
  if (props.vert === 'subscript') { open += '<sub>'; close = `</sub>${close}`; }
  const rules = [];
  if (props.color) rules.push(`color:${props.color}`);
  if (props.size) rules.push(`font-size:${props.size}pt`);
  const inner2 = rules.length ? `<span style="${rules.join(';')}">${inner}</span>` : inner;
  return open + inner2 + close;
}

function odfBlocks(nodes, styles, files) {
  let html = '';
  for (const node of nodes) {
    if (typeof node === 'string') continue;
    switch (node.tag) {
      case 'text:h': {
        const level = Math.min(6, Math.max(1, Number(node.attrs['text:outline-level'] || 1)));
        html += `<h${level}>${odfInline(node.children, styles, files) || '<br>'}</h${level}>`;
        break;
      }
      case 'text:p': {
        const props = styles.get(node.attrs['text:style-name']) || {};
        const style = props.align && props.align !== 'left' ? ` style="text-align:${props.align}"` : '';
        html += `<p${style}>${odfInline(node.children, styles, files) || '<br>'}</p>`;
        break;
      }
      case 'text:list': {
        const items = node.children.filter((child) => typeof child !== 'string' && child.tag === 'text:list-item');
        html += `<ul>${items.map((item) => `<li>${odfBlocks(item.children, styles, files) || '<br>'}</li>`).join('')}</ul>`;
        break;
      }
      case 'table:table': html += odfTable(node, styles, files); break;
      case 'text:section': case 'office:annotation': html += odfBlocks(node.children, styles, files); break;
      default: break;
    }
  }
  return html;
}

function odfTable(table, styles, files) {
  const rows = [];
  for (const row of findAll(table, 'table:table-row')) {
    const cells = [];
    for (const cell of row.children) {
      if (typeof cell === 'string') continue;
      if (cell.tag === 'table:table-cell' || cell.tag === 'table:covered-table-cell') {
        const span = Math.min(64, Math.max(1, Number(cell.attrs['table:number-columns-spanned'] || 1)));
        cells.push(`<td${span > 1 ? ` colspan="${span}"` : ''}>${odfBlocks(cell.children, styles, files) || '<br>'}</td>`);
      }
    }
    if (cells.length) rows.push(`<tr>${cells.join('')}</tr>`);
  }
  return `<table>${rows.join('')}</table>`;
}

export function odtToHtml(buffer) {
  const files = unzip(buffer);
  const content = files.get('content.xml');
  if (!content) throw new Error('not an OpenDocument text file: content.xml is missing');
  const tree = parseMarkup(content.toString('utf8'));
  const styles = readTextStyles(find(tree, 'office:automatic-styles'));
  const text = find(tree, 'office:text');
  if (!text) throw new Error('malformed content.xml');
  const stylesXml = files.get('styles.xml');
  const stylesTree = stylesXml ? parseMarkup(stylesXml.toString('utf8')) : null;
  const master = stylesTree ? find(stylesTree, 'style:master-page') : null;
  for (const [name, props] of readTextStyles(stylesTree ? find(stylesTree, 'office:automatic-styles') : null)) styles.set(name, props);
  const header = master ? find(master, 'style:header') : null;
  const footer = master ? find(master, 'style:footer') : null;
  const layout = (stylesTree ? findAll(stylesTree, 'style:page-layout-properties') : [])[0];
  const columns = Math.min(3, Math.max(1, Math.trunc(Number(layout && find(layout, 'style:columns')?.attrs['fo:column-count'] || 1))));
  const width = odfInches(layout?.attrs['fo:page-width'], 8.5);
  const height = odfInches(layout?.attrs['fo:page-height'], 11);
  const pageSize = Math.min(width, height) < 8.4 ? 'a4' : 'letter';
  const orientation = width > height ? 'landscape' : 'portrait';
  const rawMargin = odfInches(layout?.attrs['fo:margin'], odfInches(layout?.attrs['fo:margin-left'], 1));
  const margin = [0.5, 0.75, 1].reduce((best, value) => Math.abs(value - rawMargin) < Math.abs(best - rawMargin) ? value : best, 1);
  return { html: odfBlocks(text.children, styles, files), title: readOdfTitle(files), header: header ? odfBlocks(header.children, styles, files) : '', footer: footer ? odfBlocks(footer.children, styles, files) : '', columns, pageSize, orientation, margin };
}

// -------------------------------------------------------------- .odt writing

function styleKey(props) {
  return [props.b ? 'b' : '', props.i ? 'i' : '', props.u ? 'u' : '', props.strike ? 's' : '',
    props.vert || '', props.color || '', props.size || ''].join('|');
}

function odfInlineOut(nodes, registry, inherited) {
  let xml = '';
  for (const node of nodes) {
    if (typeof node === 'string') { xml += escapeXml(node); continue; }
    switch (node.tag) {
      case 'br': xml += '<text:line-break/>'; continue;
      case 'img': {
        const parsed = parseImageDataUrl(node.attrs.src);
        if (parsed) {
          const index = registry.images.length + 1;
          const path = `Pictures/image${index}.${parsed.ext}`;
          registry.images.push({ path, mime: parsed.mime, bytes: Buffer.from(parsed.base64, 'base64') });
          const width = Math.min(800, Math.max(1, Number(node.attrs.width) || 320));
          const height = Math.min(600, Math.max(1, Number(node.attrs.height) || 200));
          xml += `<draw:frame draw:name="Image${index}" text:anchor-type="as-char" svg:width="${width}px" svg:height="${height}px"><draw:image xlink:href="${path}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`;
        }
        continue;
      }
      case 'a': xml += `<text:a xlink:type="simple" xlink:href="${escapeXml(node.attrs.href || '')}">${odfInlineOut(node.children, registry, inherited)}</text:a>`; continue;
      default: break;
    }
    const props = { ...inherited };
    switch (node.tag) {
      case 'strong': case 'b': props.b = true; break;
      case 'em': case 'i': props.i = true; break;
      case 'u': props.u = true; break;
      case 's': case 'strike': case 'del': props.strike = true; break;
      case 'sup': props.vert = 'superscript'; break;
      case 'sub': props.vert = 'subscript'; break;
      case 'span': case 'font': {
        const style = node.attrs.style || '';
        const color = /(?:^|;)\s*color\s*:\s*(#[0-9a-fA-F]{3,6})/.exec(style);
        if (color) props.color = color[1].toLowerCase();
        const size = /(?:^|;)\s*font-size\s*:\s*([\d.]+)pt/.exec(style);
        if (size) props.size = Number(size[1]);
        if (/(?:^|;)\s*font-weight\s*:\s*(bold|[6-9]00)/.test(style)) props.b = true;
        if (/(?:^|;)\s*font-style\s*:\s*italic/.test(style)) props.i = true;
        break;
      }
      default: break;
    }
    const inner = odfInlineOut(node.children, registry, props);
    const key = styleKey(props);
    if (key === '||||||') { xml += inner; continue; }
    let name = registry.text.get(key);
    if (!name) {
      name = `T${registry.text.size + 1}`;
      registry.text.set(key, name);
    }
    xml += `<text:span text:style-name="${name}">${inner}</text:span>`;
  }
  return xml;
}

const ODF_ALIGN = { center: 'center', right: 'right', justify: 'justify' };

function paragraphAlign(node) {
  const value = node.attrs?.align;
  if (ODF_ALIGN[value]) return ODF_ALIGN[value];
  return ODF_ALIGN[/(?:^|;)\s*text-align\s*:\s*(\w+)/.exec(node.attrs?.style || '')?.[1]];
}

function odfBlocksOut(nodes, registry) {
  let xml = '';
  for (const node of nodes) {
    if (typeof node === 'string') {
      if (node.trim() !== '') xml += `<text:p>${escapeXml(node)}</text:p>`;
      continue;
    }
    if (node.tag === 'ul' || node.tag === 'ol') {
      const style = node.tag === 'ol' ? 'LNumber' : 'LBullet';
      const items = node.children.filter((child) => typeof child !== 'string' && child.tag === 'li')
        .map((item) => `<text:list-item><text:p>${odfInlineOut(item.children, registry, {}) || '<text:s/>'}</text:p></text:list-item>`)
        .join('');
      xml += `<text:list text:style-name="${style}">${items}</text:list>`;
      continue;
    }
    if (node.tag === 'table') { xml += odfTableOut(node, registry); continue; }
    if (node.tag === 'hr') { xml += '<text:p>\u2014\u2014\u2014</text:p>'; continue; }

    const align = paragraphAlign(node);
    const styleName = align ? `P${align.charAt(0).toUpperCase()}${align.slice(1)}` : '';
    const attr = styleName ? ` text:style-name="${styleName}"` : '';
    const outline = /^h([1-6])$/.exec(node.tag);
    const inner = odfInlineOut(node.children, registry, {}) || '<text:s/>';

    if (outline) xml += `<text:h${attr} text:outline-level="${outline[1]}">${inner}</text:h>`;
    else if (node.tag === 'blockquote') xml += `<text:p${attr}>${inner}</text:p>`;
    else xml += `<text:p${attr}>${inner}</text:p>`;
    if (align) registry.paragraphs.add(align);
  }
  return xml;
}

function odfTableOut(table, registry) {
  const rows = [];
  let index = 0;
  for (const row of tableRows(table)) {
    const cells = [];
    for (const cell of tableCells(row)) {
      index += 1;
      const content = odfBlocksOut(cell.children, registry) || '<text:p><text:s/></text:p>';
      const span = Math.max(1, Number(cell.attrs.colspan || 1));
      const spanAttr = span > 1 ? ` table:number-columns-spanned="${span}"` : '';
      cells.push(`<table:table-cell${spanAttr} office:value-type="string">${content}</table:table-cell>`);
      for (let extra = 1; extra < span; extra++) cells.push('<table:covered-table-cell/>');
    }
    if (cells.length) rows.push(`<table:table-row>${cells.join('')}</table:table-row>`);
  }
  registry.hasTable = rows.length > 0;
  return `<table:table table:name="Table${index}">${rows.join('')}</table:table>`;
}

function automaticStyles(registry) {
  const text = [...registry.text.entries()].map(([key, name]) => {
    const [b, i, u, strike, vert, color, size] = key.split('|');
    const rules = [];
    if (b) rules.push('fo:font-weight="bold"');
    if (i) rules.push('fo:font-style="italic"');
    if (u) rules.push('style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"');
    if (strike) rules.push('style:text-line-through-style="solid"');
    if (vert === 'superscript') rules.push('style:text-position="super 58%"');
    if (vert === 'subscript') rules.push('style:text-position="sub 58%"');
    if (color) rules.push(`fo:color="${color}"`);
    if (size) rules.push(`fo:font-size="${size}pt"`);
    return `<style:style style:name="${name}" style:family="text"><style:text-properties ${rules.join(' ')}/></style:style>`;
  }).join('');

  const paragraphs = [...registry.paragraphs].map((align) => {
    const name = `P${align.charAt(0).toUpperCase()}${align.slice(1)}`;
    return `<style:style style:name="${name}" style:family="paragraph"><style:paragraph-properties fo:text-align="${align}"/></style:style>`;
  }).join('');

  const tables = registry.hasTable
    ? '<style:style style:name="TblBorder" style:family="table"><style:table-properties table:border-model="collapsing"/></style:style><style:style style:name="CellBorder" style:family="table-cell"><style:table-cell-properties fo:border="0.5pt solid #bfbfbf" style:padding="0.05in"/></style:style>'
    : '';

  return text + paragraphs + tables;
}

function odtStylesXml(registry, layout = {}, header = '', footer = '') {
  const listStyles = `<text:list-style style:name="LBullet"><text:list-level-style-bullet text:level="1" text:bullet-char="\u2022"/></text:list-style><text:list-style style:name="LNumber"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/></text:list-style>`;
  const columns = Math.min(3, Math.max(1, Number(layout.columns) || 1));
  const pageSize = layout.pageSize === 'a4' ? 'a4' : 'letter';
  const orientation = layout.orientation === 'landscape' ? 'landscape' : 'portrait';
  const margin = [0.5, 0.75, 1].includes(Number(layout.margin)) ? Number(layout.margin) : 1;
  const dimensions = pageSize === 'a4' ? [8.27, 11.69] : [8.5, 11];
  if (orientation === 'landscape') dimensions.reverse();
  const headerXml = header ? `<style:header>${header}</style:header>` : '';
  const footerXml = footer ? `<style:footer>${footer}</style:footer>` : '';
  return `${XML_HEAD}<office:document-styles ${nsAttrs(['office', 'style', 'text', 'fo', 'svg', 'table', 'draw', 'xlink'])} office:version="1.2"><office:styles>${listStyles}<style:default-style style:family="paragraph"><style:paragraph-properties fo:text-align="start"/></style:default-style></office:styles><office:automatic-styles>${automaticStyles(registry)}<style:page-layout style:name="PageLayout"><style:page-layout-properties fo:page-width="${dimensions[0]}in" fo:page-height="${dimensions[1]}in" fo:margin="${margin}in"><style:columns fo:column-count="${columns}" fo:column-gap="0.3in"/></style:page-layout-properties></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="PageLayout">${headerXml}${footerXml}</style:master-page></office:master-styles></office:document-styles>`;
}

export function htmlToOdt(html, title = 'Document', layout = {}) {
  const tree = parseMarkup(html);
  const registry = { text: new Map(), paragraphs: new Set(), hasTable: false, images: [] };
  const body = odfBlocksOut(tree.children, registry);
  const header = odfBlocksOut(parseMarkup(layout.header || '').children, registry);
  const footer = odfBlocksOut(parseMarkup(layout.footer || '').children, registry);

  return zip([
    ...odfParts('text', registry.images),
    {
      name: 'content.xml',
      data: `${XML_HEAD}<office:document-content ${nsAttrs(['office', 'style', 'text', 'table', 'fo', 'svg', 'xlink', 'draw'])} office:version="1.2"><office:automatic-styles>${automaticStyles(registry)}</office:automatic-styles><office:body><office:text>${body}</office:text></office:body></office:document-content>`,
    },
    { name: 'styles.xml', data: odtStylesXml(registry, layout, header, footer) },
    { name: 'meta.xml', data: metaXml(title) },
    ...registry.images.map((item) => ({ name: item.path, data: item.bytes })),
  ]);
}

// -------------------------------------------------------------- .ods reading

function odfFormula(raw) {
  let formula = String(raw).replace(/^(of|oooc|ooxml):=/i, '');
  formula = formula.replace(/\[\.?/g, '').replace(/\]/g, '').replace(/:\s*\./g, ':');
  return formula.replace(/\$/g, '').replace(/;/g, ',');
}

function readOdfSheetTable(table, files) {
  const cells = {};
  const images = {};
  let row = 0;

  for (const rowNode of table.children) {
    if (typeof rowNode === 'string' || rowNode.tag !== 'table:table-row') continue;
    const rowRepeat = Math.min(MAX_SHEET_ROWS, Math.max(1, Number(rowNode.attrs['table:number-rows-repeated'] || 1)));
    const rowCells = [];
    for (const cell of rowNode.children) {
      if (typeof cell === 'string') continue;
      if (cell.tag === 'table:covered-table-cell') { rowCells.push(null); continue; }
      if (cell.tag !== 'table:table-cell') continue;
      const repeat = Math.min(MAX_SHEET_COLS, Math.max(1, Number(cell.attrs['table:number-columns-repeated'] || 1)));
      for (let n = 0; n < repeat; n++) rowCells.push(n === 0 ? cell : null);
    }

    for (let n = 0; n < rowRepeat; n++) {
      row += 1;
      if (row > MAX_SHEET_ROWS) break;
      let col = 0;
      for (const cell of rowCells) {
        col += 1;
        if (col > MAX_SHEET_COLS) break;
        if (!cell) continue;
        const value = odfCellValue(cell);
        const ref = `${numToCol(col)}${row}`;
        if (value !== '') cells[ref] = value;
        const frame = find(cell, 'draw:frame');
        const drawImage = frame && find(frame, 'draw:image');
        const href = drawImage?.attrs['xlink:href'];
        const bytes = href && files.get(href);
        if (bytes) {
          const ext = href.split('.').pop().toLowerCase();
          const mime = ({ png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif' })[ext];
          const src = imageDataUrl(mime, bytes);
          if (src) images[ref] = { src, width: 160, height: 100, alt: '' };
        }
      }
    }
    if (row > MAX_SHEET_ROWS) break;
  }
  return { cells, images };
}

function odfCellValue(cell) {
  const formula = cell.attrs['table:formula'];
  if (formula) return `=${odfFormula(formula)}`;

  const type = cell.attrs['office:value-type'];
  if (type === 'float' || type === 'percentage' || type === 'currency') return String(cell.attrs['office:value'] ?? '');
  if (type === 'date') return String(cell.attrs['office:date-value'] ?? '').slice(0, 10);
  if (type === 'time') return String(cell.attrs['office:time-value'] ?? '');
  if (type === 'boolean') return String(cell.attrs['office:boolean-value']) === 'true' ? 'TRUE' : 'FALSE';

  const paragraph = find(cell, 'text:p');
  return paragraph ? textOf(paragraph) : '';
}

export function odsToSheets(buffer) {
  const files = unzip(buffer);
  const content = files.get('content.xml');
  if (!content) throw new Error('not an OpenDocument spreadsheet: content.xml is missing');
  const spreadsheet = find(parseMarkup(content.toString('utf8')), 'office:spreadsheet');
  if (!spreadsheet) throw new Error('malformed content.xml');

  const sheets = [];
  for (const table of spreadsheet.children) {
    if (typeof table === 'string' || table.tag !== 'table:table') continue;
    const data = readOdfSheetTable(table, files);
    sheets.push({
      name: String(table.attrs['table:name'] || `Sheet${sheets.length + 1}`),
      cells: data.cells,
      images: data.images,
    });
  }
  if (!sheets.length) throw new Error('spreadsheet has no sheets');
  return { sheets, active: 0 };
}

// -------------------------------------------------------------- .ods writing

function odsCellXml(ref, raw, engine, image, registry) {
  let picture = '';
  if (image) {
    const parsed = parseImageDataUrl(image.src);
    if (parsed) {
      const index = registry.length + 1;
      const path = `Pictures/image${index}.${parsed.ext}`;
      registry.push({ path, mime: parsed.mime, bytes: Buffer.from(parsed.base64, 'base64') });
      picture = `<draw:frame draw:name="Image${index}" text:anchor-type="as-char" svg:width="${Math.min(800, Math.max(1, Number(image.width) || 160))}px" svg:height="${Math.min(600, Math.max(1, Number(image.height) || 100))}px"><draw:image xlink:href="${path}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`;
    }
  }
  if (raw[0] === '=') {
    const shown = engine.display(ref);
    const numeric = shown !== '' && Number.isFinite(Number(shown));
    const typeAttr = numeric
      ? `office:value-type="float" office:value="${Number(shown)}"`
      : 'office:value-type="string"';
    return `<table:table-cell table:formula="of:=${escapeXml(raw.slice(1))}" ${typeAttr}><text:p>${escapeXml(shown)}${picture}</text:p></table:table-cell>`;
  }
  const text = raw.trim();
  if (text !== '' && Number.isFinite(Number(text)) && !/^0\d/.test(text)) {
    return `<table:table-cell office:value-type="float" office:value="${Number(text)}"><text:p>${escapeXml(text)}${picture}</text:p></table:table-cell>`;
  }
  return `<table:table-cell office:value-type="string"><text:p>${escapeXml(raw)}${picture}</text:p></table:table-cell>`;
}

function odsTableXml(sheet, media) {
  const engine = makeEngine(sheet.cells);
  const byRow = new Map();
  let maxCol = 0;
  for (const [ref, raw] of Object.entries(sheet.cells)) {
    const m = /^([A-Z]+)([0-9]+)$/.exec(ref);
    if (!m) continue;
    const row = Number(m[2]);
    if (row > MAX_SHEET_ROWS) continue;
    const col = colToNum(m[1]);
    maxCol = Math.max(maxCol, col);
    if (!byRow.has(row)) byRow.set(row, new Map());
    byRow.get(row).set(col, { ref, raw: String(raw) });
  }
  for (const [ref, image] of Object.entries(sheet.images || {})) {
    const m = /^([A-Z]+)([0-9]+)$/.exec(ref.toUpperCase());
    if (!m || Number(m[2]) > MAX_SHEET_ROWS || colToNum(m[1]) > MAX_SHEET_COLS || !parseImageDataUrl(image?.src)) continue;
    const row = Number(m[2]);
    const col = colToNum(m[1]);
    maxCol = Math.max(maxCol, col);
    if (!byRow.has(row)) byRow.set(row, new Map());
    const existing = byRow.get(row).get(col) || { ref: `${m[1]}${row}`, raw: '' };
    existing.image = image;
    byRow.get(row).set(col, existing);
  }

  const rows = [];
  let emptyRun = 0;
  const flushEmpty = () => {
    if (!emptyRun) return;
    rows.push(emptyRun === 1 ? '<table:table-row/>' : `<table:table-row table:number-rows-repeated="${emptyRun}"/>`);
    emptyRun = 0;
  };

  const lastRow = byRow.size ? Math.max(...byRow.keys()) : 0;
  for (let row = 1; row <= lastRow; row++) {
    const cells = byRow.get(row);
    if (!cells || cells.size === 0) { emptyRun += 1; continue; }
    flushEmpty();
    const parts = [];
    let empty = 0;
    for (let col = 1; col <= maxCol; col++) {
      const cell = cells.get(col);
      if (!cell) { empty += 1; continue; }
      if (empty) {
        parts.push(empty === 1 ? '<table:table-cell/>' : `<table:table-cell table:number-columns-repeated="${empty}"/>`);
        empty = 0;
      }
      parts.push(odsCellXml(cell.ref, cell.raw, engine, cell.image, media));
    }
    rows.push(`<table:table-row>${parts.join('')}</table:table-row>`);
  }
  flushEmpty();

  const columns = `<table:table-column table:number-columns-repeated="${Math.max(1, maxCol)}"/>`;
  return `<table:table table:name="${escapeXml(sheet.name.slice(0, 31))}">${columns}${rows.join('')}</table:table>`;
}

export function sheetsToOds(doc) {
  const media = [];
  const tables = doc.sheets.map((sheet) => odsTableXml(sheet, media)).join('');
  return zip([
    ...odfParts('spreadsheet', media),
    {
      name: 'content.xml',
      data: `${XML_HEAD}<office:document-content ${nsAttrs(['office', 'style', 'text', 'table', 'fo', 'svg', 'number', 'draw', 'xlink'])} office:version="1.2"><office:automatic-styles/><office:body><office:spreadsheet>${tables}</office:spreadsheet></office:body></office:document-content>`,
    },
    {
      name: 'styles.xml',
      data: `${XML_HEAD}<office:document-styles ${nsAttrs(['office', 'style', 'text', 'table', 'fo'])} office:version="1.2"><office:styles><style:style style:name="Default" style:family="table-cell"/></office:styles><office:automatic-styles/></office:document-styles>`,
    },
    { name: 'meta.xml', data: metaXml('Spreadsheet') },
    ...media.map((item) => ({ name: item.path, data: item.bytes })),
  ]);
}
