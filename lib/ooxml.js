// OOXML (.docx / .xlsx) import and export, built straight on the ZIP and
// markup primitives. No XML library: the two sides of each format are small
// enough that direct tree walking is clearer than a generic DOM.
import { unzip, zip } from './zip.js';
import { parseMarkup, findAll, find, textOf, tableRows, tableCells, escapeXml, escapeHtml } from '../shared/html.js';
import { makeEngine, colToNum, numToCol } from '../shared/formula.js';
import { imageDataUrl, parseImageDataUrl } from '../shared/images.js';
import { paperSizeFromDimensions, paperTwips } from '../shared/paper.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const SS_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

// ------------------------------------------------------------------- reading

function flag(parent, tag) {
  const node = parent ? find(parent, tag) : null;
  if (!node) return false;
  const value = node.attrs['w:val'];
  if (value === undefined) return true;
  return !['0', 'false', 'off'].includes(String(value).toLowerCase());
}

function renderRun(run, imageForRel = () => '') {
  let text = '';
  const pieces = [];
  const flushText = () => {
    if (!text) return;
    pieces.push(applyRunStyle(escapeHtml(text).replace(/\n/g, '<br>'), find(run, 'w:rPr')));
    text = '';
  };
  for (const child of run.children) {
    if (typeof child === 'string') continue;
    if (child.tag === 'w:t') text += textOf(child);
    else if (child.tag === 'w:tab') text += '\t';
    else if (child.tag === 'w:br' || child.tag === 'w:cr') text += '\n';
    else if (child.tag === 'w:nobreakhyphen') text += '-';
    else if (child.tag === 'w:drawing') { flushText(); pieces.push(renderDrawing(child, imageForRel)); }
  }
  flushText();
  return pieces.join('');
}

function applyRunStyle(html, rPr) {
  if (!rPr) return html;
  const style = [];
  let open = '';
  let close = '';

  if (flag(rPr, 'w:b')) { open += '<strong>'; close = `</strong>${close}`; }
  if (flag(rPr, 'w:i')) { open += '<em>'; close = `</em>${close}`; }
  if (find(rPr, 'w:u')) { open += '<u>'; close = `</u>${close}`; }
  if (flag(rPr, 'w:strike') || flag(rPr, 'w:dstrike')) { open += '<s>'; close = `</s>${close}`; }

  const vert = find(rPr, 'w:vertAlign')?.attrs['w:val'];
  if (vert === 'superscript') { open += '<sup>'; close = `</sup>${close}`; }
  if (vert === 'subscript') { open += '<sub>'; close = `</sub>${close}`; }

  const color = find(rPr, 'w:color')?.attrs['w:val'];
  if (color && /^[0-9a-fA-F]{6}$/.test(color)) style.push(`color:#${color}`);
  const size = Number(find(rPr, 'w:sz')?.attrs['w:val']);
  if (Number.isFinite(size) && size > 0) style.push(`font-size:${size / 2}pt`);
  const fonts = find(rPr, 'w:rFonts');
  const family = fonts?.attrs['w:ascii'] || fonts?.attrs['w:hAnsi'] || fonts?.attrs['w:eastAsia'];
  if (family && /^[\w .,'-]{1,80}$/.test(family)) style.push(`font-family:${family}`);

  const inner = style.length ? `<span style="${style.join(';')}">${html}</span>` : html;
  return open + inner + close;
}

const INLINE_WRAPPERS = new Set(['w:hyperlink', 'w:ins', 'w:smarttag', 'w:sdtcontent', 'w:sdt']);

function renderInline(nodes, imageForRel) {
  let out = '';
  for (const node of nodes) {
    if (typeof node === 'string') continue;
    if (node.tag === 'w:r') out += renderRun(node, imageForRel);
    else if (INLINE_WRAPPERS.has(node.tag)) out += renderInline(node.children, imageForRel);
    // w:pPr, w:proofErr, w:bookmarkStart and friends carry no user-visible text.
  }
  return out;
}

const ALIGN = { both: 'justify', distribute: 'justify', center: 'center', right: 'right', left: 'left' };

function renderParagraph(p, imageForRel) {
  const pr = find(p, 'w:pPr');
  let tag = 'p';
  const rules = [];

  if (pr) {
    const name = find(pr, 'w:pStyle')?.attrs['w:val'] || '';
    const heading = /^Heading([1-6])$/i.exec(name);
    if (heading) tag = `h${heading[1]}`;
    else if (/^Title$/i.test(name)) tag = 'h1';
    else if (/^Quote$/i.test(name)) tag = 'blockquote';
    const align = ALIGN[find(pr, 'w:jc')?.attrs['w:val']];
    if (align && align !== 'left') rules.push(`text-align:${align}`);
  }

  const body = renderInline(p.children, imageForRel);
  const style = rules.length ? ` style="${rules.join(';')}"` : '';
  return `<${tag}${style}>${body || '<br>'}</${tag}>`;
}

function renderTable(tbl, imageForRel) {
  const rows = [];
  for (const row of tbl.children) {
    if (typeof row === 'string' || row.tag !== 'w:tr') continue;
    const cells = [];
    for (const cell of row.children) {
      if (typeof cell === 'string' || cell.tag !== 'w:tc') continue;
      const span = Number(find(cell, 'w:tcPr') ? find(find(cell, 'w:tcPr'), 'w:gridSpan')?.attrs['w:val'] || 1 : 1);
      const attr = span > 1 ? ` colspan="${span}"` : '';
      cells.push(`<td${attr}>${renderBlocks(cell.children, imageForRel) || '<br>'}</td>`);
    }
    rows.push(`<tr>${cells.join('')}</tr>`);
  }
  return `<table>${rows.join('')}</table>`;
}

function hasNumbering(paragraph) {
  const pr = find(paragraph, 'w:pPr');
  return pr ? Boolean(find(pr, 'w:numPr')) : false;
}

function renderBlocks(nodes, imageForRel = () => '') {
  let out = '';
  let i = 0;
  while (i < nodes.length) {
    const node = nodes[i];
    if (typeof node === 'string') { i++; continue; }

    if (node.tag === 'w:p' && hasNumbering(node)) {
      const items = [];
      while (i < nodes.length && typeof nodes[i] !== 'string' && nodes[i].tag === 'w:p' && hasNumbering(nodes[i])) {
        items.push(`<li>${renderInline(nodes[i].children, imageForRel) || '<br>'}</li>`);
        i++;
      }
      out += `<ul>${items.join('')}</ul>`;
      continue;
    }
    if (node.tag === 'w:p') { out += renderParagraph(node, imageForRel); i++; continue; }
    if (node.tag === 'w:tbl') { out += renderTable(node, imageForRel); i++; continue; }
    if (node.tag === 'w:sdt') {
      const content = find(node, 'w:sdtContent');
      if (content) out += renderBlocks(content.children, imageForRel);
      i++;
      continue;
    }
    i++;
  }
  return out;
}

export function docxToHtml(buffer) {
  const files = unzip(buffer);
  const document = files.get('word/document.xml');
  if (!document) throw new Error('not a Word document: word/document.xml is missing');
  const body = find(parseMarkup(document.toString('utf8')), 'w:body');
  if (!body) throw new Error('malformed word/document.xml');
  const relationships = new Map();
  const relsXml = files.get('word/_rels/document.xml.rels');
  if (relsXml) for (const rel of findAll(parseMarkup(relsXml.toString('utf8')), 'Relationship')) {
    if (String(rel.attrs.Type || '').endsWith('/image')) relationships.set(rel.attrs.Id, rel.attrs.Target || '');
  }
  const imageForRel = (id, width, height, align = '', style = '') => {
    let target = relationships.get(id);
    if (!target) return '';
    target = target.replace(/^\//, '');
    if (!target.startsWith('word/')) target = `word/${target.replace(/^\.\//, '')}`;
    const bytes = files.get(target);
    if (!bytes) return '';
    const ext = target.split('.').pop().toLowerCase();
    const mime = ({ png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif' })[ext];
    const src = imageDataUrl(mime, bytes);
    const imageAlign = ['left', 'center', 'right'].includes(align) ? ` align="${align}"` : '';
    const imageStyle = style ? ` style="${style}"` : '';
    return src ? `<img src="${src}"${width ? ` width="${Math.max(1, Math.round(width / 9525))}"` : ''}${height ? ` height="${Math.max(1, Math.round(height / 9525))}"` : ''}${imageAlign}${imageStyle}>` : '';
  };
  const imageResolverForPart = (partPath) => {
    const fileName = partPath.split('/').pop();
    const relPath = `${partPath.slice(0, partPath.lastIndexOf('/') + 1)}_rels/${fileName}.rels`;
    const partRels = files.get(relPath);
    const images = new Map();
    if (partRels) for (const rel of findAll(parseMarkup(partRels.toString('utf8')), 'Relationship')) {
      if (String(rel.attrs.Type || '').endsWith('/image')) images.set(rel.attrs.Id, rel.attrs.Target || '');
    }
    return (id, width, height, align = '', style = '') => {
      const target = images.get(id);
      if (!target) return '';
      const stack = partPath.slice(0, partPath.lastIndexOf('/') + 1).split('/').filter(Boolean);
      for (const segment of target.split('/')) {
        if (segment === '..') stack.pop();
        else if (segment && segment !== '.') stack.push(segment);
      }
      const mediaPath = stack.join('/');
      const bytes = files.get(mediaPath);
      if (!bytes) return '';
      const ext = mediaPath.split('.').pop().toLowerCase();
      const mime = ({ png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif' })[ext];
      const src = imageDataUrl(mime, bytes);
      const imageAlign = ['left', 'center', 'right'].includes(align) ? ` align="${align}"` : '';
      const imageStyle = style ? ` style="${style}"` : '';
      return src ? `<img src="${src}"${width ? ` width="${Math.max(1, Math.round(width / 9525))}"` : ''}${height ? ` height="${Math.max(1, Math.round(height / 9525))}"` : ''}${imageAlign}${imageStyle}>` : '';
    };
  };
  const section = find(body, 'w:sectPr');
  const columns = Math.min(3, Math.max(1, Math.trunc(Number(section && find(section, 'w:cols')?.attrs['w:num'] || 1))));
  const page = section && find(section, 'w:pgSz');
  const pageWidth = Number(page?.attrs['w:w'] || 12240);
  const pageHeight = Number(page?.attrs['w:h'] || 15840);
  const pageSize = paperSizeFromDimensions(pageWidth / 1440, pageHeight / 1440);
  const orientation = pageWidth > pageHeight ? 'landscape' : 'portrait';
  const margins = section && find(section, 'w:pgMar');
  const rawMargin = Number(margins?.attrs['w:left'] || 1440) / 1440;
  const margin = [0.5, 0.75, 1].reduce((best, value) => Math.abs(value - rawMargin) < Math.abs(best - rawMargin) ? value : best, 1);
  const headerFooter = (tag) => {
    const type = tag === 'w:hdr' ? 'header' : 'footer';
    const rel = findAll(parseMarkup(relsXml?.toString('utf8') || ''), 'Relationship')
      .find((item) => String(item.attrs.Type || '').endsWith(`/${type}`));
    if (!rel) return '';
    let target = rel.attrs.Target || '';
    target = target.replace(/^\//, '');
    if (!target.startsWith('word/')) target = `word/${target}`;
    const part = files.get(target);
    if (!part) return '';
    const node = find(parseMarkup(part.toString('utf8')), tag);
    return node ? renderBlocks(node.children, imageResolverForPart(target)) : '';
  };
  return { html: renderBlocks(body.children, imageForRel), title: readCoreTitle(files), header: headerFooter('w:hdr'), footer: headerFooter('w:ftr'), columns, pageSize, orientation, margin };
}

function renderDrawing(drawing, imageForRel) {
  const blip = find(drawing, 'a:blip');
  if (!blip) return '';
  const extent = find(drawing, 'wp:extent');
  const anchor = find(drawing, 'wp:anchor');
  const position = anchor && find(anchor, 'wp:positionH');
  const align = position ? textOf(position).trim().toLowerCase() : '';
  const pxOffset = (node) => {
    const raw = Number(textOf(node));
    return node && Number.isFinite(raw) ? Math.max(-2000, Math.min(2000, Math.round(raw / 9525))) : 0;
  };
  const x = pxOffset(position && find(position, 'wp:posOffset'));
  const vertical = anchor && find(anchor, 'wp:positionV');
  const y = pxOffset(vertical && find(vertical, 'wp:posOffset'));
  const offsetStyle = x || y ? `position:relative;left:${x}px;top:${y}px` : '';
  return imageForRel(blip.attrs['r:embed'], Number(extent?.attrs.cx), Number(extent?.attrs.cy), align, offsetStyle);
}

function readCoreTitle(files) {
  const core = files.get('docProps/core.xml');
  if (!core) return '';
  const title = find(parseMarkup(core.toString('utf8')), 'dc:title');
  return title ? textOf(title).trim() : '';
}

// ------------------------------------------------------------------- writing

function pushRun(runs, text, props) {
  if (text === '') return;
  const previous = runs[runs.length - 1];
  if (previous && !previous.br && sameProps(previous.props, props)) {
    previous.text += text;
    return;
  }
  runs.push({ text, props });
}

function sameProps(a, b) {
  const keys = new Set([...Object.keys(a || {}), ...Object.keys(b || {})]);
  for (const key of keys) if ((a || {})[key] !== (b || {})[key]) return false;
  return true;
}

function styleProps(node) {
  const style = node.attrs?.style;
  if (!style) return {};
  const props = {};
  const color = /(?:^|;)\s*color\s*:\s*(#[0-9a-fA-F]{3,6})/.exec(style);
  if (color) props.color = color[1].slice(1).toLowerCase();
  const size = /(?:^|;)\s*font-size\s*:\s*([\d.]+)pt/.exec(style);
  if (size) props.size = Number(size[1]);
  const family = /(?:^|;)\s*font-family\s*:\s*([^;]+)/i.exec(style);
  if (family && /^[\w .,'-]{1,80}$/.test(family[1].trim())) props.font = family[1].trim().replace(/^['"]|['"]$/g, '');
  const weight = /(?:^|;)\s*font-weight\s*:\s*(bold|[6-9]00)/.test(style);
  if (weight) props.b = true;
  const italic = /(?:^|;)\s*font-style\s*:\s*italic/.test(style);
  if (italic) props.i = true;
  const decoration = /(?:^|;)\s*text-decoration[^;]*line-through/.test(style);
  if (decoration) props.strike = true;
  return props;
}

function emitInline(nodes, props, runs, images) {
  for (const node of nodes) {
    if (typeof node === 'string') { pushRun(runs, node, props); continue; }
    const next = { ...props };
    switch (node.tag) {
      case 'br': runs.push({ br: true, props: next }); continue;
      case 'img': {
        const image = parseImageDataUrl(node.attrs.src);
        if (image) {
          const id = `rIdImage${images.length + 1}`;
          const bytes = Buffer.from(image.base64, 'base64');
          const align = ['left', 'center', 'right'].includes(node.attrs.align) ? node.attrs.align : '';
          const offset = (property) => {
            const match = new RegExp(`(?:^|;)\\s*${property}\\s*:\\s*(-?(?:\\d+(?:\\.\\d+)?|\\.\\d+))px`, 'i').exec(node.attrs.style || '');
            return match ? Math.round(Math.max(-2000, Math.min(2000, Number(match[1]))) * 9525) : 0;
          };
          images.push({ id, ...image, bytes, width: Math.min(1600, Math.max(1, Number(node.attrs.width) || 320)), height: Math.min(1600, Math.max(1, Number(node.attrs.height) || 200)), align, offsetX: offset('left'), offsetY: offset('top') });
          runs.push({ image: id, props: next });
        }
        continue;
      }
      case 'strong': case 'b': next.b = true; break;
      case 'em': case 'i': next.i = true; break;
      case 'u': next.u = true; break;
      case 's': case 'strike': case 'del': next.strike = true; break;
      case 'sup': next.vert = 'superscript'; break;
      case 'sub': next.vert = 'subscript'; break;
      case 'code': case 'tt': case 'pre': next.mono = true; break;
      case 'a': next.href = node.attrs.href || ''; break;
      case 'span': case 'font': Object.assign(next, styleProps(node)); break;
      default: break;
    }
    emitInline(node.children, next, runs, images);
  }
  return runs;
}

function runXml(run, links) {
  if (run.br) return '<w:r><w:br/></w:r>';
  if (run.image) {
    const image = links.images.find((item) => item.id === run.image);
    const cx = Math.round(image.width * 9525);
    const cy = Math.round(image.height * 9525);
    const id = links.nextPicture++;
    const anchored = Boolean(image.align || image.offsetX || image.offsetY);
    const horizontal = image.align ? `<wp:align>${image.align}</wp:align>` : `<wp:posOffset>${image.offsetX || 0}</wp:posOffset>`;
    const position = anchored
      ? `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin">${horizontal}</wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${image.offsetY || 0}</wp:posOffset></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/><wp:wrapSquare wrapText="bothSides"/>`
      : `<wp:inline><wp:extent cx="${cx}" cy="${cy}"/>`;
    const closePosition = anchored ? '</wp:anchor>' : '</wp:inline>';
    return `<w:r><w:drawing>${position}<wp:docPr id="${id}" name="Picture ${id}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="image.${image.ext}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${image.id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>${closePosition}</w:drawing></w:r>`;
  }

  const props = run.props || {};
  let rPr = '';
  if (props.b) rPr += '<w:b/>';
  if (props.i) rPr += '<w:i/>';
  if (props.u) rPr += '<w:u w:val="single"/>';
  if (props.strike) rPr += '<w:strike/>';
  if (props.vert) rPr += `<w:vertAlign w:val="${props.vert}"/>`;
  if (props.color) rPr += `<w:color w:val="${escapeXml(props.color)}"/>`;
  if (props.font) rPr += `<w:rFonts w:ascii="${escapeXml(props.font)}" w:hAnsi="${escapeXml(props.font)}"/>`;
  if (props.size) {
    const half = Math.round(props.size * 2);
    rPr += `<w:sz w:val="${half}"/><w:szCs w:val="${half}"/>`;
  }
  if (props.mono) rPr += '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>';

  const pieces = String(run.text).split('\n')
    .map((line) => `<w:t xml:space="preserve">${escapeXml(line)}</w:t>`)
    .join('<w:br/>');
  const body = `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${pieces}</w:r>`;

  if (!props.href) return body;
  let id = links.get(props.href);
  if (!id) {
    id = `rIdLink${links.size + 1}`;
    links.set(props.href, id);
  }
  return `<w:hyperlink r:id="${id}">${body}</w:hyperlink>`;
}

function paragraphXml(runs, options, links) {
  let pPr = '';
  if (options.style) pPr += `<w:pStyle w:val="${options.style}"/>`;
  if (options.numId) pPr += `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${options.numId}"/></w:numPr>`;
  if (options.align) pPr += `<w:jc w:val="${options.align}"/>`;

  const body = runs.map((run) => runXml(run, links)).join('');
  const empty = body === '' ? '<w:r><w:t xml:space="preserve"></w:t></w:r>' : '';
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${body}${empty}</w:p>`;
}

const BLOCK_STYLE = { h1: 'Heading1', h2: 'Heading2', h3: 'Heading3', h4: 'Heading4', h5: 'Heading5', h6: 'Heading6', blockquote: 'Quote' };
const ALIGN_IN = { center: 'center', right: 'right', justify: 'both' };

function elementAlign(node) {
  const value = node.attrs?.align;
  if (ALIGN_IN[value]) return ALIGN_IN[value];
  return ALIGN_IN[/(?:^|;)\s*text-align\s*:\s*(\w+)/.exec(node.attrs?.style || '')?.[1]];
}

function emitBlocks(nodes, links) {
  let xml = '';
  for (const node of nodes) {
    if (typeof node === 'string') {
      if (node.trim() !== '') xml += paragraphXml(emitInline([node], {}, [], links.images), {}, links);
      continue;
    }
    if (node.tag === 'ul' || node.tag === 'ol') {
      const numId = node.tag === 'ol' ? 2 : 1;
      for (const item of node.children) {
        if (typeof item === 'string' || item.tag !== 'li') continue;
        xml += paragraphXml(emitInline(item.children, {}, [], links.images), { numId, align: elementAlign(node) }, links);
      }
      continue;
    }
    if (node.tag === 'table') { xml += tableXml(node, links); continue; }
    if (node.tag === 'hr') { xml += paragraphXml([{ text: '\u2014\u2014\u2014', props: {} }], { align: 'center' }, links); continue; }

    const style = BLOCK_STYLE[node.tag];
    const runs = emitInline(node.children, {}, [], links.images);
    if (node.tag === 'pre' && runs[0]) runs[0].props = { ...runs[0].props, mono: true };
    xml += paragraphXml(runs, { style, align: elementAlign(node) }, links);
  }
  return xml;
}

function tableXml(table, links) {
  const rows = [];
  for (const row of tableRows(table)) {
    const cells = [];
    for (const cell of tableCells(row)) {
      const span = Number(cell.attrs.colspan || 1);
      const spanXml = span > 1 ? `<w:gridSpan w:val="${span}"/>` : '';
      const content = emitBlocks(cell.children, links) || paragraphXml([], {}, links);
      cells.push(`<w:tc><w:tcPr><w:tcW w:w="${span > 1 ? Math.floor(9360 / span) : 9360}" w:type="dxa"/>${spanXml}</w:tcPr>${content}</w:tc>`);
    }
    rows.push(`<w:tr>${cells.join('')}</w:tr>`);
  }
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((side) => `<w:${side} w:val="single" w:sz="4" w:color="BFBFBF"/>`).join('')}</w:tblBorders></w:tblPr>${rows.join('')}</w:tbl>`;
}

function stylesXml() {
  const headings = [1, 2, 3, 4, 5, 6].map((level) => `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="${level - 1}"/></w:pPr><w:rPr><w:b/><w:sz w:val="${[32, 28, 24, 22, 20, 20][level - 1]}"/></w:rPr></w:style>`).join('');
  return `${XML_HEAD}<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:rPr><w:i/><w:color w:val="606060"/></w:rPr></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>${headings}</w:styles>`;
}

function numberingXml() {
  const abstract = (id, format, text) => `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>`;
  return `${XML_HEAD}<w:numbering xmlns:w="${W_NS}">${abstract(0, 'bullet', '\uF0B7')}${abstract(1, 'decimal', '%1.')}<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
}

export function htmlToDocx(html, title = 'Document', layout = {}) {
  const tree = parseMarkup(html);
  const links = new Map();
  links.images = [];
  links.nextPicture = 1;
  const body = emitBlocks(tree.children, links);
  const header = emitBlocks(parseMarkup(layout.header || '').children, links);
  const footer = emitBlocks(parseMarkup(layout.footer || '').children, links);
  const columns = Math.min(3, Math.max(1, Number(layout.columns) || 1));
  const pageSize = layout.pageSize || 'letter';
  const orientation = layout.orientation === 'landscape' ? 'landscape' : 'portrait';
  const margin = [0.5, 0.75, 1].includes(Number(layout.margin)) ? Number(layout.margin) : 1;
  const dimensions = paperTwips(pageSize);
  if (orientation === 'landscape') dimensions.reverse();
  const marginTwips = Math.round(margin * 1440);
  const usesLists = `${body}${header}${footer}`.includes('<w:numPr>');

  const rels = [`<Relationship Id="rIdStyles" Type="${R_NS}/styles" Target="styles.xml"/>`];
  if (usesLists) rels.push(`<Relationship Id="rIdNumbering" Type="${R_NS}/numbering" Target="numbering.xml"/>`);
  for (const [href, id] of links) {
    if (href === 'images' || href === 'nextPicture') continue;
    rels.push(`<Relationship Id="${id}" Type="${R_NS}/hyperlink" Target="${escapeXml(href)}" TargetMode="External"/>`);
  }
  for (const image of links.images) rels.push(`<Relationship Id="${image.id}" Type="${R_NS}/image" Target="media/image${image.id.replace('rIdImage', '')}.${image.ext}"/>`);
  if (header) rels.push(`<Relationship Id="rIdHeader" Type="${R_NS}/header" Target="header1.xml"/>`);
  if (footer) rels.push(`<Relationship Id="rIdFooter" Type="${R_NS}/footer" Target="footer1.xml"/>`);
  const imageTypes = [...new Set(links.images.map((image) => image.ext))].map((ext) => `<Default Extension="${ext}" ContentType="image/${ext === 'jpg' ? 'jpeg' : ext}"/>`).join('');

  const files = [
    {
      name: '[Content_Types].xml',
      data: `${XML_HEAD}<Types xmlns="${CT_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${imageTypes}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${header ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : ''}${footer ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : ''}${usesLists ? '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' : ''}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`,
    },
    {
      name: '_rels/.rels',
      data: `${XML_HEAD}<Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="${R_NS}/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`,
    },
    { name: 'docProps/core.xml', data: coreXml(title) },
    { name: 'word/styles.xml', data: stylesXml() },
    {
      name: 'word/document.xml',
      data: `${XML_HEAD}<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${body}<w:sectPr>${header ? '<w:headerReference w:type="default" r:id="rIdHeader"/>' : ''}${footer ? '<w:footerReference w:type="default" r:id="rIdFooter"/>' : ''}<w:pgSz w:w="${dimensions[0]}" w:h="${dimensions[1]}"/><w:pgMar w:top="${marginTwips}" w:right="${marginTwips}" w:bottom="${marginTwips}" w:left="${marginTwips}" w:header="${Math.round(marginTwips / 2)}" w:footer="${Math.round(marginTwips / 2)}" w:gutter="0"/><w:cols w:num="${columns}" w:space="720"/></w:sectPr></w:body></w:document>`,
    },
    {
      name: 'word/_rels/document.xml.rels',
      data: `${XML_HEAD}<Relationships xmlns="${PKG_REL_NS}">${rels.join('')}</Relationships>`,
    },
  ];
  for (const image of links.images) files.push({ name: `word/media/image${image.id.replace('rIdImage', '')}.${image.ext}`, data: image.bytes });
  if (header) files.push({ name: 'word/header1.xml', data: `${XML_HEAD}<w:hdr xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">${header}</w:hdr>` });
  if (footer) files.push({ name: 'word/footer1.xml', data: `${XML_HEAD}<w:ftr xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">${footer}</w:ftr>` });
  const imageRels = links.images.map((image) => `<Relationship Id="${image.id}" Type="${R_NS}/image" Target="media/image${image.id.replace('rIdImage', '')}.${image.ext}"/>`).join('');
  if (header && imageRels) files.push({ name: 'word/_rels/header1.xml.rels', data: `${XML_HEAD}<Relationships xmlns="${PKG_REL_NS}">${imageRels}</Relationships>` });
  if (footer && imageRels) files.push({ name: 'word/_rels/footer1.xml.rels', data: `${XML_HEAD}<Relationships xmlns="${PKG_REL_NS}">${imageRels}</Relationships>` });
  if (usesLists) files.push({ name: 'word/numbering.xml', data: numberingXml() });

  return zip(files);
}

function coreXml(title) {
  const stamp = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  return `${XML_HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${escapeXml(title)}</dc:title><dcterms:created xsi:type="dcterms:W3CDTF">${stamp}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${stamp}</dcterms:modified></cp:coreProperties>`;
}

// ------------------------------------------------------------------ xlsx in

function readSharedStrings(xml) {
  if (!xml) return [];
  return findAll(parseMarkup(xml), 'si').map((si) => textOf(si));
}

const BUILTIN_DATE_FORMATS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);

/** Per-cell style records, just enough to tell a date from a plain number. */
function readStyles(xml) {
  const lookup = new Map();
  if (!xml) return lookup;

  const tree = parseMarkup(xml);
  const custom = new Map();
  for (const fmt of findAll(tree, 'numFmt')) custom.set(Number(fmt.attrs.numFmtId), String(fmt.attrs.formatCode || ''));
  const fonts = find(tree, 'fonts')?.children.filter((item) => typeof item !== 'string' && item.tag === 'font').map((font) => ({
    fontFamily: String(find(font, 'name')?.attrs.val || ''),
    fontSize: Number(find(font, 'sz')?.attrs.val || 0),
  })) || [];

  const cellXfs = find(tree, 'cellXfs');
  for (const xf of cellXfs ? cellXfs.children : []) {
    if (typeof xf === 'string' || xf.tag !== 'xf') continue;
    const id = Number(xf.attrs.numFmtId || 0);
    const code = custom.get(id) || '';
    // A format code is a date format if it has date tokens outside literals.
    const dateish = BUILTIN_DATE_FORMATS.has(id)
      || /[ymdhs]/i.test(code.replace(/\[[^\]]*\]/g, '').replace(/"[^"]*"/g, '').replace(/\\./g, ''));
    lookup.set(lookup.size, { date: dateish, ...(fonts[Number(xf.attrs.fontId || 0)]?.fontFamily ? { fontFamily: fonts[Number(xf.attrs.fontId || 0)].fontFamily } : {}), ...(fonts[Number(xf.attrs.fontId || 0)]?.fontSize ? { fontSize: fonts[Number(xf.attrs.fontId || 0)].fontSize } : {}) });
  }
  return lookup;
}

/** Excel serial day number -> display text. Epoch is 1899-12-30 (with the 1900 leap bug baked in). */
function serialToDate(serial) {
  const ms = Date.UTC(1899, 11, 30) + serial * 86400000;
  const date = new Date(ms);
  const pad = (n) => String(n).padStart(2, '0');
  const day = `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}`;
  const hasTime = Math.abs(serial % 1) > 1e-9;
  if (!hasTime) return day;
  return `${day} ${pad(date.getUTCHours())}:${pad(date.getUTCMinutes())}`;
}

function worksheetCells(xml, shared, styleInfo) {
  const cells = {};
  const cellStyles = {};
  for (const cell of findAll(parseMarkup(xml), 'c')) {
    const ref = String(cell.attrs.r || '').toUpperCase();
    if (!ref) continue;
    const type = cell.attrs.t || 'n';
    const formula = find(cell, 'f');
    const value = find(cell, 'v');
    let text;
    if (type === 'inlineStr') text = textOf(find(cell, 'is') || { children: [] });
    else if (type === 's') text = shared[Number(textOf(value))] ?? '';
    else if (type === 'str') text = textOf(value);
    else if (type === 'b') text = textOf(value) === '1' ? 'TRUE' : 'FALSE';
    else text = value ? textOf(value) : '';

    // Shared formulas (`<f t="shared" si="0"/>`) carry no text in the follower
    // cells, so those fall back to the cached result rather than becoming "=".
    const expression = formula ? textOf(formula).trim() : '';
    if (expression) text = `=${expression}`;
    else if (type === 'n' && text !== '' && styleInfo.get(Number(cell.attrs.s || 0))?.date && Number.isFinite(Number(text))) {
      text = serialToDate(Number(text));
    }
    if (text !== '') cells[ref] = text;
    const style = styleInfo.get(Number(cell.attrs.s || 0));
    if (style?.fontFamily || style?.fontSize) cellStyles[ref] = { ...(style.fontFamily ? { fontFamily: style.fontFamily } : {}), ...(style.fontSize ? { fontSize: style.fontSize } : {}) };
  }
  const tree = parseMarkup(xml);
  const columnWidths = {};
  for (const col of findAll(tree, 'col')) {
    const width = Number(col.attrs.width);
    if (!Number.isFinite(width) || width <= 0) continue;
    for (let index = Number(col.attrs.min || 1); index <= Math.min(256, Number(col.attrs.max || 0)); index++) columnWidths[numToCol(index)] = Math.min(600, Math.max(40, Math.round(width * 7 + 5)));
  }
  const rowHeights = {};
  for (const row of findAll(tree, 'row')) {
    const height = Number(row.attrs.ht);
    if (Number.isFinite(height) && height > 0) rowHeights[row.attrs.r] = Math.min(300, Math.max(18, Math.round(height * 96 / 72)));
  }
  return { cells, cellStyles, columnWidths, rowHeights };
}

export function xlsxToSheets(buffer) {
  const files = unzip(buffer);
  const workbook = files.get('xl/workbook.xml');
  if (!workbook) throw new Error('not an Excel workbook: xl/workbook.xml is missing');

  const rels = new Map();
  const relXml = files.get('xl/_rels/workbook.xml.rels');
  if (relXml) {
    for (const rel of findAll(parseMarkup(relXml.toString('utf8')), 'Relationship')) {
      rels.set(rel.attrs.Id, rel.attrs.Target || '');
    }
  }
  const shared = readSharedStrings(files.get('xl/sharedStrings.xml')?.toString('utf8'));
  const styleIsDate = readStyles(files.get('xl/styles.xml')?.toString('utf8'));
  const sheets = [];

  for (const sheet of findAll(parseMarkup(workbook.toString('utf8')), 'sheet')) {
    let target = rels.get(sheet.attrs['r:id']) || `worksheets/sheet${sheets.length + 1}.xml`;
    target = target.replace(/^\//, '');
    const path = target.startsWith('xl/') ? target : `xl/${target.replace(/^\.\//, '')}`;
    const xml = files.get(path) || files.get(path.replace(/^xl\//, ''));
    const sheetXml = xml?.toString('utf8') || '';
    const parsed = xml ? worksheetCells(sheetXml, shared, styleIsDate) : { cells: {}, cellStyles: {}, columnWidths: {}, rowHeights: {} };
    sheets.push({
      name: String(sheet.attrs.name || `Sheet${sheets.length + 1}`),
      ...parsed,
      images: xml ? readWorksheetImages(files, path, sheetXml) : {},
    });
  }

  if (!sheets.length) throw new Error('workbook has no sheets');
  return { sheets, active: 0 };
}

function readWorksheetImages(files, worksheetPath, worksheetXml) {
  const result = {};
  const relPath = worksheetPath.replace(/([^/]+)$/, '_rels/$1.rels');
  const relXml = files.get(relPath);
  if (!relXml) return result;
  const drawingTarget = findAll(parseMarkup(relXml.toString('utf8')), 'Relationship')
    .find((rel) => String(rel.attrs.Type || '').endsWith('/drawing'))?.attrs.Target;
  if (!drawingTarget) return result;
  const drawingPath = `${worksheetPath.slice(0, worksheetPath.lastIndexOf('/') + 1)}${drawingTarget}`.replace(/\/[^/]+\/\.\./g, '');
  const drawing = files.get(drawingPath);
  if (!drawing) return result;
  const drawingRelPath = drawingPath.replace(/([^/]+)$/, '_rels/$1.rels');
  const drawingRels = new Map();
  const drawingRelXml = files.get(drawingRelPath);
  if (drawingRelXml) for (const rel of findAll(parseMarkup(drawingRelXml.toString('utf8')), 'Relationship')) drawingRels.set(rel.attrs.Id, rel.attrs.Target || '');
  for (const anchor of findAll(parseMarkup(drawing.toString('utf8')), 'xdr:oneCellAnchor')) {
    const from = find(anchor, 'xdr:from');
    const row = Number(find(from, 'xdr:row') && textOf(find(from, 'xdr:row'))) + 1;
    const col = Number(find(from, 'xdr:col') && textOf(find(from, 'xdr:col'))) + 1;
    const blip = find(anchor, 'a:blip');
    const rel = drawingRels.get(blip?.attrs['r:embed']);
    if (!rel || !row || !col || col > 256 || row > 5000) continue;
    const mediaPath = `${drawingPath.slice(0, drawingPath.lastIndexOf('/') + 1)}${rel}`.replace(/\/[^/]+\/\.\./g, '');
    const bytes = files.get(mediaPath);
    if (!bytes) continue;
    const ext = mediaPath.split('.').pop().toLowerCase();
    const mime = ({ png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif' })[ext];
    const src = imageDataUrl(mime, bytes);
    if (!src) continue;
    const size = find(anchor, 'xdr:ext');
    result[`${numToCol(col)}${row}`] = { src, width: Math.max(1, Math.round(Number(size?.attrs.cx || 1524000) / 9525)), height: Math.max(1, Math.round(Number(size?.attrs.cy || 952500) / 9525)), alt: '' };
  }
  return result;
}

// ----------------------------------------------------------------- xlsx out

function sheetName(name, index) {
  // Excel rejects : \ / ? * [ ] and colon-like names, and caps length at 31.
  const clean = String(name || '').replace(/[:\\/?*[\]]/g, '-').trim().slice(0, 31);
  return clean || `Sheet${index + 1}`;
}

function cellXml(ref, raw, engine, styleId = 0) {
  const style = styleId ? ` s="${styleId}"` : '';
  if (raw[0] === '=') {
    const shown = engine.display(ref);
    const numeric = shown !== '' && Number.isFinite(Number(shown));
    const type = numeric ? '' : ' t="str"';
    return `<c r="${ref}"${type}${style}><f>${escapeXml(raw.slice(1))}</f><v>${escapeXml(shown)}</v></c>`;
  }
  const text = raw.trim();
  const numeric = text !== '' && Number.isFinite(Number(text)) && !/^0\d/.test(text);
  if (numeric) return `<c r="${ref}"${style}><v>${Number(text)}</v></c>`;
  if (text.toUpperCase() === 'TRUE' || text.toUpperCase() === 'FALSE') {
    return `<c r="${ref}" t="b"${style}><v>${text.toUpperCase() === 'TRUE' ? 1 : 0}</v></c>`;
  }
  return `<c r="${ref}" t="inlineStr"${style}><is><t xml:space="preserve">${escapeXml(raw)}</t></is></c>`;
}

function worksheetXml(sheet, hasDrawing = false, styleIds = new Map()) {
  const engine = makeEngine(sheet.cells);
  const byRow = new Map();
  for (const [ref, raw] of Object.entries(sheet.cells)) {
    const m = /^([A-Z]+)([0-9]+)$/.exec(ref);
    if (!m) continue;
    const row = Number(m[2]);
    if (!byRow.has(row)) byRow.set(row, []);
    byRow.get(row).push({ ref, col: colToNum(m[1]), raw: String(raw), styleId: styleIds.get(ref) || 0 });
  }
  for (const ref of Object.keys(sheet.cellStyles || {})) {
    const m = /^([A-Z]+)([0-9]+)$/.exec(ref);
    if (!m) continue;
    const row = Number(m[2]);
    if (!byRow.has(row)) byRow.set(row, []);
    if (!byRow.get(row).some((cell) => cell.ref === ref)) byRow.get(row).push({ ref, col: colToNum(m[1]), raw: '', styleId: styleIds.get(ref) || 0 });
  }
  const rows = [...byRow.keys()].sort((a, b) => a - b).map((row) => {
    const cells = byRow.get(row).sort((a, b) => a.col - b.col)
      .map((cell) => cell.raw === '' ? `<c r="${cell.ref}" s="${cell.styleId}"/>` : cellXml(cell.ref, cell.raw, engine, cell.styleId)).join('');
    const height = Number(sheet.rowHeights?.[row]);
    const rowAttrs = height ? ` ht="${height * 72 / 96}" customHeight="1"` : '';
    return `<row r="${row}"${rowAttrs}>${cells}</row>`;
  }).join('');
  const cols = Object.entries(sheet.columnWidths || {}).map(([name, width]) => {
    const index = colToNum(name);
    return `<col min="${index}" max="${index}" width="${Math.max(0, Number(width) - 5) / 7}" customWidth="1"/>`;
  }).join('');
  return `${XML_HEAD}<worksheet xmlns="${SS_NS}" xmlns:r="${R_NS}">${cols ? `<cols>${cols}</cols>` : ''}<sheetData>${rows}</sheetData>${hasDrawing ? '<drawing r:id="rIdDrawing"/>' : ''}</worksheet>`;
}

function sheetStyles(sheets) {
  const fonts = [{ fontFamily: 'Arial', fontSize: 11 }];
  const fontIndex = new Map([['Arial|11', 0]]);
  const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const styleIdsBySheet = sheets.map(() => new Map());
  const xfIndex = new Map();
  sheets.forEach((sheet, sheetIndex) => {
    for (const [ref, style] of Object.entries(sheet.cellStyles || {})) {
      const fontFamily = String(style.fontFamily || 'Arial');
      const fontSize = Number(style.fontSize) || 11;
      const key = `${fontFamily}|${fontSize}`;
      if (!fontIndex.has(key)) { fontIndex.set(key, fonts.length); fonts.push({ fontFamily, fontSize }); }
      if (!xfIndex.has(key)) { xfIndex.set(key, xfs.length); xfs.push(`<xf numFmtId="0" fontId="${fontIndex.get(key)}" fillId="0" borderId="0" xfId="0" applyFont="1"/>`); }
      styleIdsBySheet[sheetIndex].set(ref, xfIndex.get(key));
    }
  });
  const fontXml = fonts.map((font) => `<font><sz val="${font.fontSize}"/><name val="${escapeXml(font.fontFamily)}"/></font>`).join('');
  return { styleIdsBySheet, xml: `${XML_HEAD}<styleSheet xmlns="${SS_NS}"><fonts count="${fonts.length}">${fontXml}</fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs></styleSheet>` };
}

function collectSheetImages(sheet) {
  const result = [];
  for (const [ref, value] of Object.entries(sheet.images || {})) {
    const match = /^([A-Z]{1,3})([1-9]\d*)$/i.exec(ref);
    const parsed = parseImageDataUrl(value?.src);
    if (!match || !parsed) continue;
    const col = colToNum(match[1].toUpperCase());
    const row = Number(match[2]);
    if (col > 256 || row > 5000) continue;
    result.push({ ref: `${match[1].toUpperCase()}${row}`, ...parsed, bytes: Buffer.from(parsed.base64, 'base64'), width: Math.min(800, Math.max(1, Number(value.width) || 160)), height: Math.min(600, Math.max(1, Number(value.height) || 100)) });
  }
  return result;
}

function drawingParts(images, sheetIndex) {
  const anchors = [];
  const rels = [];
  images.forEach((image, i) => {
    const m = /^([A-Z]+)(\d+)$/.exec(image.ref);
    const id = `rIdImage${i + 1}`;
    const cx = image.width * 9525;
    const cy = image.height * 9525;
    anchors.push(`<xdr:oneCellAnchor><xdr:from><xdr:col>${colToNum(m[1]) - 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${Number(m[2]) - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${cx}" cy="${cy}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${i + 1}" name="Image ${i + 1}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${id}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`);
    rels.push(`<Relationship Id="${id}" Type="${R_NS}/image" Target="../media/sheet${sheetIndex}-image${i + 1}.${image.ext}"/>`);
  });
  return {
    xml: `${XML_HEAD}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R_NS}">${anchors.join('')}</xdr:wsDr>`,
    rels: `${XML_HEAD}<Relationships xmlns="${PKG_REL_NS}">${rels.join('')}</Relationships>`,
  };
}

export function sheetsToXlsx(doc) {
  const names = doc.sheets.map((sheet, index) => sheetName(sheet.name, index));
  const images = doc.sheets.map(collectSheetImages);
  const styles = sheetStyles(doc.sheets);
  const drawings = images.map((items, i) => drawingParts(items, i + 1));
  const imageTypes = [...new Set(images.flatMap((items) => items.map((image) => image.ext)))].map((ext) => `<Default Extension="${ext}" ContentType="image/${ext === 'jpg' ? 'jpeg' : ext}"/>`).join('');
  const files = [
    {
      name: '[Content_Types].xml',
      data: `${XML_HEAD}<Types xmlns="${CT_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${imageTypes}${names.map((_, i) => images[i].length ? `<Override PartName="/xl/drawings/drawing${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>` : '').join('')}<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${names.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`,
    },
    {
      name: '_rels/.rels',
      data: `${XML_HEAD}<Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    },
    {
      name: 'xl/workbook.xml',
      data: `${XML_HEAD}<workbook xmlns="${SS_NS}" xmlns:r="${R_NS}"><sheets>${names.map((name, i) => `<sheet name="${escapeXml(name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`,
    },
    {
      name: 'xl/_rels/workbook.xml.rels',
      data: `${XML_HEAD}<Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rIdStyles" Type="${R_NS}/styles" Target="styles.xml"/>${names.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${R_NS}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}</Relationships>`,
    },
    {
      name: 'xl/styles.xml',
      data: styles.xml,
    },
  ];
  doc.sheets.forEach((sheet, i) => {
    const hasImages = images[i].length > 0;
    files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: worksheetXml(sheet, hasImages, styles.styleIdsBySheet[i]) });
    if (hasImages) {
      files.push({ name: `xl/worksheets/_rels/sheet${i + 1}.xml.rels`, data: `${XML_HEAD}<Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rIdDrawing" Type="${R_NS}/drawing" Target="../drawings/drawing${i + 1}.xml"/></Relationships>` });
      files.push({ name: `xl/drawings/drawing${i + 1}.xml`, data: drawings[i].xml });
      files.push({ name: `xl/drawings/_rels/drawing${i + 1}.xml.rels`, data: drawings[i].rels });
      images[i].forEach((image, j) => files.push({ name: `xl/media/sheet${i + 1}-image${j + 1}.${image.ext}`, data: image.bytes }));
    }
  });
  return zip(files);
}
