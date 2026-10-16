// RTF import/export. Word's oldest surviving text format, still common.
//
// shortcut: import produces plain text (no font/size/color table parsing) and
// export carries bold/italic/underline only. Upgrade to a real property table
// if RTF fidelity ever matters more than the other formats.
import { parseMarkup, tableRows, tableCells } from '../shared/html.js';

// RTF \'hh escapes are cp1252 bytes, not latin1, in everything Word writes.
const CP1252 = {
  0x80: '\u20ac', 0x82: '\u201a', 0x83: '\u0192', 0x84: '\u201e', 0x85: '\u2026',
  0x86: '\u2020', 0x87: '\u2021', 0x88: '\u02c6', 0x89: '\u2030', 0x8a: '\u0160',
  0x8b: '\u2039', 0x8c: '\u0152', 0x8e: '\u017d', 0x91: '\u2018', 0x92: '\u2019',
  0x93: '\u201c', 0x94: '\u201d', 0x95: '\u2022', 0x96: '\u2013', 0x97: '\u2014',
  0x98: '\u02dc', 0x99: '\u2122', 0x9a: '\u0161', 0x9b: '\u203a', 0x9c: '\u0153',
  0x9e: '\u017e', 0x9f: '\u0178',
};

const SKIP_DESTINATIONS = new Set([
  'fonttbl', 'colortbl', 'stylesheet', 'info', 'pict', 'xmlnstbl', 'header', 'footer',
  'footnote', 'listtable', 'listoverridetable', 'generator', 'filetbl', 'revtbl',
  'rsidtbl', 'themedata', 'colorschememapping', 'datastore', 'latentstyles', 'docvars',
  'shppict', 'nonshppict', 'fldinst', 'object',
]);

/** @returns {{text: string}} plain text with `\n` paragraph breaks */
export function rtfToText(input) {
  const src = Buffer.isBuffer(input) ? input.toString('latin1') : String(input);
  const stack = [];
  let depth = 0;
  let ignoreDepth = -1;
  let out = '';
  let i = 0;

  while (i < src.length) {
    const ch = src[i];

    if (ch === '{') { stack.push(ignoreDepth); depth += 1; i += 1; continue; }
    if (ch === '}') { ignoreDepth = stack.length ? stack.pop() : -1; depth = Math.max(0, depth - 1); i += 1; continue; }
    if (ch === '\r' || ch === '\n') { i += 1; continue; }

    if (ch !== '\\') {
      if (ignoreDepth < 0) out += ch;
      i += 1;
      continue;
    }

    const next = src[i + 1];

    if (next === "'") {
      const byte = parseInt(src.slice(i + 2, i + 4), 16);
      if (ignoreDepth < 0 && Number.isFinite(byte)) out += CP1252[byte] || String.fromCharCode(byte);
      i += 4;
      continue;
    }
    if (next === '\\' || next === '{' || next === '}') {
      if (ignoreDepth < 0) out += next;
      i += 2;
      continue;
    }
    if (next === '*') { // \* introduces an ignorable destination
      ignoreDepth = depth;
      const name = /^\\\*\\?([a-zA-Z]+)?/.exec(src.slice(i));
      i += name ? name[0].length : 2;
      continue;
    }

    const word = /^\\([a-zA-Z]+)(-?[0-9]+)?[ ]?/.exec(src.slice(i));
    if (word) {
      i += word[0].length;
      const name = word[1].toLowerCase();
      if (SKIP_DESTINATIONS.has(name)) { ignoreDepth = depth; continue; }
      if (ignoreDepth >= 0) continue;
      if (name === 'par' || name === 'line') out += '\n';
      else if (name === 'tab') out += '\t';
      else if (name === 'bullet') out += '\u2022';
      else if (name === 'emdash') out += '\u2014';
      else if (name === 'endash') out += '\u2013';
      else if (name === 'u') {
        let code = Number(word[2] || 0);
        if (code < 0) code += 65536;
        out += String.fromCharCode(code);
        if (src[i] === '?') i += 1; // the ANSI fallback character
      }
      continue;
    }

    i += 2; // unknown control symbol: drop it
  }

  return { text: out.replace(/\n{3,}/g, '\n\n') };
}

/** Text -> HTML so it can go through the normal word document model. */
export function textToHtml(text) {
  const lines = String(text).split(/\n/);
  return lines.map((line) => `<p>${line === '' ? '<br>' : escapeText(line)}</p>`).join('');
}

export function htmlToText(html) {
  const tree = parseMarkup(html);
  const lines = [];
  const flush = (node) => {
    if (typeof node === 'string') { if (node.trim()) lines.push(node.trim()); return; }
    if (node.tag === 'ul' || node.tag === 'ol') {
      let index = 1;
      for (const item of node.children) {
        if (typeof item === 'string' || item.tag !== 'li') continue;
        const text = item.children.map((child) => (typeof child === 'string' ? child : textOnly(child))).join('').trim();
        lines.push(node.tag === 'ol' ? `${index}. ${text}` : `\u2022 ${text}`);
        index += 1;
      }
      return;
    }
    if (node.tag === 'table') {
      for (const row of tableRows(node)) {
        const cells = tableCells(row).map((cell) => cell.children.map((child) => (typeof child === 'string' ? child : textOnly(child))).join('').trim());
        lines.push(cells.join('\t'));
      }
      return;
    }
    const text = node.children.map((child) => (typeof child === 'string' ? child : textOnly(child))).join('');
    lines.push(text);
  };
  for (const node of tree.children) flush(node);
  return lines.join('\n').replace(/\n{3,}/g, '\n\n');
}

function textOnly(node) {
  if (typeof node === 'string') return node;
  if (node.tag === 'br') return '\n';
  return node.children.map(textOnly).join('');
}

function escapeText(text) {
  return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeRtf(text) {
  let out = '';
  for (const ch of String(text)) {
    const code = ch.codePointAt(0);
    if (ch === '\\') out += '\\\\';
    else if (ch === '{') out += '\\{';
    else if (ch === '}') out += '\\}';
    else if (ch === '\t') out += '\\tab ';
    else if (code < 128) out += ch;
    else if (code <= 0xffff) out += `\\u${code > 32767 ? code - 65536 : code}?`;
    else {
      const high = Math.floor((code - 0x10000) / 0x400) + 0xd800;
      const low = ((code - 0x10000) % 0x400) + 0xdc00;
      out += `\\u${high - 65536}?\\u${low - 65536}?`;
    }
  }
  return out;
}

function inlineRtf(nodes, state) {
  let out = '';
  for (const node of nodes) {
    if (typeof node === 'string') { out += escapeRtf(node); continue; }

    const next = { ...state };
    let suffix = '';
    switch (node.tag) {
      case 'br': out += '\\line '; continue;
      case 'img': continue;
      case 'strong': case 'b': next.b = true; break;
      case 'em': case 'i': next.i = true; break;
      case 'u': next.u = true; break;
      case 'sup': out += '\\super '; suffix = '\\nosupersub '; break;
      case 'sub': out += '\\sub '; suffix = '\\nosupersub '; break;
      default: break;
    }

    let prefix = '';
    if (next.b !== state.b) prefix += next.b ? '\\b ' : '\\b0 ';
    if (next.i !== state.i) prefix += next.i ? '\\i ' : '\\i0 ';
    if (next.u !== state.u) prefix += next.u ? '\\ul ' : '\\ulnone ';
    out += prefix + inlineRtf(node.children, next) + suffix;
  }
  return out;
}

export function htmlToRtf(html, title = '') {
  const tree = parseMarkup(html);
  let body = '';

  const paragraph = (nodes, prefix = '') => {
    const runs = inlineRtf(nodes, {});
    // \plain closes any character formatting left open inside the paragraph.
    body += `${prefix}${runs}\\plain\\par\n`;
  };

  for (const node of tree.children) {
    if (typeof node === 'string') { if (node.trim()) paragraph([node]); continue; }
    if (node.tag === 'ul' || node.tag === 'ol') {
      let index = 1;
      for (const item of node.children) {
        if (typeof item === 'string' || item.tag !== 'li') continue;
        paragraph(item.children, node.tag === 'ol' ? `${index}. ` : '\\bullet\\tab ');
        index += 1;
      }
      continue;
    }
    if (node.tag === 'table') {
      for (const row of tableRows(node)) {
        const cells = tableCells(row).map((cell) => inlineRtf(cell.children, {}));
        body += `${cells.join('\\tab ')}\\plain\\par\n`;
      }
      continue;
    }
    if (node.tag === 'hr') { paragraph([{ tag: 'span', attrs: {}, children: ['\u2014\u2014\u2014'] }], ''); continue; }
    paragraph(node.children);
  }

  const info = title ? `{\\info{\\title ${escapeRtf(title)}}}` : '';
  return `{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\fnil\\fcharset0 Arial;}}${info}\n${body}}`;
}
