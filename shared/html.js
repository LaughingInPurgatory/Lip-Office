// Tiny tolerant markup parser plus escaping helpers.
// One parser serves every exporter (docx, odt, rtf, txt), which is why it
// lives on its own instead of using a real DOM library.
//
// Tags are lower-cased, so this reads OOXML (`w:p`) and ODF (`text:p`) too:
// both formats are lower-case anyway. Namespaced names keep their prefix.

const VOID_TAGS = new Set(['br', 'hr', 'img', 'input', 'meta', 'link', 'col', 'source', 'wbr']);
const RAW_TEXT_TAGS = new Set(['script', 'style']);

const ENTITIES = {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: '\u00a0',
  ndash: '\u2013', mdash: '\u2014', hellip: '\u2026', bull: '\u2022',
  lsquo: '\u2018', rsquo: '\u2019', ldquo: '\u201c', rdquo: '\u201d', copy: '\u00a9',
};

export function decodeEntities(text) {
  return text.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (whole, body) => {
    if (body[0] === '#') {
      const code = body[1] === 'x' || body[1] === 'X'
        ? parseInt(body.slice(2), 16)
        : parseInt(body.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : whole;
    }
    const named = ENTITIES[body.toLowerCase()];
    return named === undefined ? whole : named;
  });
}

export function escapeXml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export function escapeHtml(text) {
  return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

const TAG_RE = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<![^>]*>|<\/([a-zA-Z][a-zA-Z0-9:._-]*)\s*>|<([a-zA-Z][a-zA-Z0-9:._-]*)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/g;
const ATTR_RE = /([a-zA-Z_:][a-zA-Z0-9_:.-]*)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+)))?/g;

// Element and attribute names are kept in the case the document used, because
// OOXML and ODF are camelCase-sensitive (`w:pStyle`, `numFmtId`, `Rel Id`).
// HTML lookups in this codebase are all lower-case, so a lower-cased alias is
// stored alongside every name for them to match against.
function parseAttrs(source) {
  const attrs = {};
  ATTR_RE.lastIndex = 0;
  let m;
  while ((m = ATTR_RE.exec(source))) {
    if (!m[1]) continue;
    const name = m[1];
    const value = decodeEntities(m[2] ?? m[3] ?? m[4] ?? '');
    attrs[name] = value;
    const lower = name.toLowerCase();
    if (lower !== name) attrs[lower] = value;
  }
  return attrs;
}

// A block element implicitly ends an open <p>, and list/table items end their
// peers. Without this, unclosed markup (common in hand-written .html imports)
// nests paragraphs inside each other.
const BLOCK_TAGS = new Set(['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'table', 'blockquote', 'pre', 'section', 'article', 'header', 'footer', 'hr']);
const PEER_CLOSERS = { p: ['p'], li: ['li'], tr: ['tr'], td: ['td', 'th'], th: ['td', 'th'] };

function implicitClose(stack, tag) {
  if (stack.length < 2) return;
  const top = stack[stack.length - 1];
  const peers = PEER_CLOSERS[tag];
  if (peers && peers.includes(top.tag)) { stack.pop(); return; }
  if (!peers && BLOCK_TAGS.has(tag) && top.tag === 'p') stack.pop();
}

/** @returns {{tag: string, attrs: object, children: Array}} root node keeping source order */
export function parseMarkup(source) {
  const root = { tag: '#root', attrs: {}, children: [] };
  const stack = [root];
  let cursor = 0;
  let rawUntil = null;
  TAG_RE.lastIndex = 0;

  const pushText = (text) => {
    if (text !== '') stack[stack.length - 1].children.push(decodeEntities(text));
  };

  let m;
  while ((m = TAG_RE.exec(source))) {
    const text = source.slice(cursor, m.index);
    cursor = TAG_RE.lastIndex;

    if (m[1]) { // closing tag
      pushText(text);
      const tag = m[1].toLowerCase();
      if (rawUntil === tag) { rawUntil = null; stack.length = Math.max(1, stack.length - 1); continue; }
      for (let i = stack.length - 1; i > 0; i--) {
        if (stack[i].tag === tag) { stack.length = i; break; }
      }
      continue;
    }

    if (m[0].startsWith('<!')) {
      pushText(text);
      continue;
    }

    pushText(text);
    if (rawUntil) continue; // markup inside <style>/<script> is text, not elements

    const tag = m[2].toLowerCase();
    const node = { tag, attrs: parseAttrs(m[3]), children: [] };
    if (!VOID_TAGS.has(tag)) implicitClose(stack, tag);
    stack[stack.length - 1].children.push(node);
    if (VOID_TAGS.has(tag) || m[4] === '/') continue;
    if (RAW_TEXT_TAGS.has(tag)) rawUntil = tag;
    stack.push(node);
  }
  pushText(source.slice(cursor));
  return root;
}

/** Concatenated text of a subtree, ignoring all markup. */
export function textOf(node) {
  if (typeof node === 'string') return node;
  if (node.tag === 'br') return '\n';
  return node.children.map(textOf).join('');
}

/** Rows of a table, looking through thead/tbody/tfoot (where browsers put them). */
export function tableRows(table) {
  const rows = [];
  const collect = (nodes) => {
    for (const node of nodes) {
      if (typeof node === 'string') continue;
      if (node.tag === 'tr') rows.push(node);
      else if (node.tag === 'thead' || node.tag === 'tbody' || node.tag === 'tfoot') collect(node.children);
    }
  };
  collect(table.children);
  return rows;
}

/** The td/th cells of one table row. */
export function tableCells(row) {
  return row.children.filter((node) => typeof node !== 'string' && (node.tag === 'td' || node.tag === 'th'));
}

/** Depth-first visit. Return `false` from the callback to skip a subtree's children. */
export function walk(node, visit, parent = null) {
  if (typeof node === 'string') {
    visit(node, parent);
    return;
  }
  if (visit(node, parent) === false) return;
  for (const child of node.children) walk(child, visit, node);
}

/** First descendant with the given tag, or null. Matching is case-insensitive. */
export function find(node, tag) {
  if (typeof node === 'string') return null;
  const wanted = String(tag).toLowerCase();
  if (node.tag === wanted) return node;
  for (const child of node.children) {
    const hit = find(child, tag);
    if (hit) return hit;
  }
  return null;
}

export function findAll(node, tag, out = []) {
  if (typeof node === 'string') return out;
  const wanted = String(tag).toLowerCase();
  if (node.tag === wanted) out.push(node);
  else for (const child of node.children) findAll(child, wanted, out);
  return out;
}
