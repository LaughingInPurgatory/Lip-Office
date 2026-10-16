// Whitelist sanitizer for document HTML.
//
// Documents arrive from uploads, get stored on disk, and are later rendered
// with innerHTML in every browser that opens them. That is a stored-XSS path,
// so anything stored or served goes through here first.
import { parseMarkup } from './html.js';
import { parseImageDataUrl } from './images.js';

const ALLOWED_TAGS = new Set([
  'p', 'br', 'div', 'span', 'b', 'strong', 'i', 'em', 'u', 's', 'strike', 'del', 'ins',
  'sub', 'sup', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 'blockquote',
  'pre', 'code', 'table', 'thead', 'tbody', 'tfoot', 'tr', 'td', 'th', 'a', 'hr',
  'mark', 'small', 'big', 'caption', 'img',
]);

// Tags whose text content is code, not prose: drop the element and everything in it.
const DROP_CONTENT = new Set(['script', 'style', 'iframe', 'object', 'embed', 'template', 'noscript', 'form', 'input', 'button', 'select', 'textarea', 'svg', 'math', 'link', 'meta', 'base']);

const VOID_TAGS = new Set(['br', 'hr', 'img']);

const STYLE_PROPS = new Set(['color', 'background-color', 'font-size', 'font-weight', 'font-style', 'text-decoration', 'text-align', 'font-family']);
const SAFE_URL = /^(?:https?:|mailto:|tel:|#|\/|\.\/)/i;

function cleanStyle(value) {
  return String(value)
    .split(';')
    .map((rule) => {
      const colon = rule.indexOf(':');
      if (colon < 0) return '';
      const name = rule.slice(0, colon).trim().toLowerCase();
      const raw = rule.slice(colon + 1).trim();
      if (!STYLE_PROPS.has(name)) return '';
      // url()/expression()/javascript: are the CSS-level injection vectors.
      if (/url\s*\(|expression|javascript:|@import|\\/i.test(raw)) return '';
      return `${name}:${raw}`;
    })
    .filter(Boolean)
    .join(';');
}

function cleanAttrs(tag, attrs) {
  const out = [];
  if (tag === 'img') {
    const image = parseImageDataUrl(attrs.src);
    if (!image) return '';
    out.push(`src="data:${image.mime};base64,${image.base64}"`);
    if (attrs.alt) out.push(`alt="${String(attrs.alt).replace(/[&"<>]/g, (ch) => ({ '&': '&amp;', '"': '&quot;', '<': '&lt;', '>': '&gt;' })[ch])}"`);
    for (const dimension of ['width', 'height']) {
      const size = Math.min(1600, Math.max(1, Number.parseInt(attrs[dimension], 10) || 0));
      if (size) out.push(`${dimension}="${size}"`);
    }
    return out.join(' ');
  }
  for (const [name, value] of Object.entries(attrs)) {
    if (name === 'style') {
      const style = cleanStyle(value);
      if (style) out.push(`style="${style.replace(/"/g, '&quot;')}"`);
      continue;
    }
    if (tag === 'a' && name === 'href') {
      const href = String(value).trim().replace(/[\u0000-\u001f]/g, '');
      if (SAFE_URL.test(href)) out.push(`href="${href.replace(/"/g, '&quot;')}"`);
      continue;
    }
    if ((tag === 'td' || tag === 'th') && (name === 'colspan' || name === 'rowspan')) {
      const span = Math.min(64, Math.max(1, Number(value) || 1));
      if (span > 1) out.push(`${name}="${span}"`);
      continue;
    }
    if (name === 'align' && ['left', 'center', 'right', 'justify'].includes(value)) out.push(`align="${value}"`);
  }
  return out.join(' ');
}

function esc(text) {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function render(node) {
  if (typeof node === 'string') return esc(node);
  if (DROP_CONTENT.has(node.tag)) return '';

  const inner = node.children.map(render).join('');

  // A rejected image is removed entirely rather than leaving an empty tag.
  if (node.tag === 'img') {
    const attrs = cleanAttrs('img', node.attrs);
    return attrs ? `<img ${attrs}>` : '';
  }

  // execCommand('foreColor') emits the legacy <font color>. Normalize it to a
  // span so every exporter only has to understand one inline styling form.
  if (node.tag === 'font') {
    const rules = [];
    if (/^#[0-9a-fA-F]{3,6}$/.test(String(node.attrs.color || '').trim())) {
      rules.push(`color:${String(node.attrs.color).toLowerCase()}`);
    }
    const style = cleanStyle(node.attrs.style || '');
    if (style) rules.push(style);
    return rules.length ? `<span style="${rules.join(';')}">${inner}</span>` : inner;
  }

  if (!ALLOWED_TAGS.has(node.tag)) return inner; // unwrap unknown wrappers, keep their text
  const attrs = cleanAttrs(node.tag, node.attrs);
  const open = `<${node.tag}${attrs ? ` ${attrs}` : ''}>`;
  if (VOID_TAGS.has(node.tag)) return open;
  return `${open}${inner}</${node.tag}>`;
}

export function sanitizeHtml(html) {
  const tree = parseMarkup(String(html || ''));
  return tree.children.map(render).join('');
}
