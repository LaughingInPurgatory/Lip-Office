import { parseMarkup, textOf, tableRows, tableCells } from '../shared/html.js';

const escapeHtml = (value) => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const escapeAttr = (value) => escapeHtml(value).replace(/"/g, '&quot;');
const escapeMarkdown = (value) => String(value).replace(/([\\`*_{}[\]()#+.!|>~-])/g, '\\$1');

function inlineMarkdown(source) {
  let html = escapeHtml(source);
  html = html.replace(/!\[([^\]]*)\]\(([^)\s]+)(?:\s+"([^"]*)")?\)/g,
    (_, alt, src, title = '') => `<img src="${escapeAttr(src)}" alt="${alt}"${title ? ` title="${escapeAttr(title)}"` : ''}>`);
  html = html.replace(/\[([^\]]+)\]\(([^)\s]+)(?:\s+"([^"]*)")?\)/g,
    (_, label, href, title = '') => `<a href="${escapeAttr(href)}"${title ? ` title="${escapeAttr(title)}"` : ''}>${label}</a>`);
  html = html.replace(/`([^`]+)`/g, '<code>$1</code>');
  html = html.replace(/\*\*(.+?)\*\*|__(.+?)__/g, (_, a, b) => `<strong>${a || b}</strong>`);
  html = html.replace(/~~(.+?)~~/g, '<del>$1</del>');
  html = html.replace(/\*(.+?)\*|_(.+?)_/g, (_, a, b) => `<em>${a || b}</em>`);
  return html;
}

function isBlockStart(lines, index) {
  const line = lines[index] || '';
  return /^\s*(?:#{1,6}\s|```|>|[-*_]\s*[-*_]\s*[-*_]|[-+*]\s|\d+[.)]\s)/.test(line)
    || (index + 1 < lines.length && /^\s*\|?\s*:?-{3,}/.test(lines[index + 1]));
}

function markdownBlocks(lines) {
  const html = [];
  for (let i = 0; i < lines.length;) {
    const line = lines[i];
    if (!line.trim()) { i++; continue; }
    const fence = /^\s*```/.exec(line);
    if (fence) {
      const code = [];
      i++;
      while (i < lines.length && !/^\s*```/.test(lines[i])) code.push(lines[i++]);
      if (i < lines.length) i++;
      html.push(`<pre><code>${escapeHtml(code.join('\n'))}</code></pre>`);
      continue;
    }
    const heading = /^(#{1,6})\s+(.+?)\s*#*\s*$/.exec(line);
    if (heading) {
      const level = heading[1].length;
      html.push(`<h${level}>${inlineMarkdown(heading[2])}</h${level}>`);
      i++;
      continue;
    }
    if (/^\s*(?:(?:\*\s*){3,}|(?:-\s*){3,}|(?:_\s*){3,})$/.test(line)) {
      html.push('<hr>'); i++; continue;
    }
    if (/^\s*>/.test(line)) {
      const quote = [];
      while (i < lines.length && /^\s*>/.test(lines[i])) quote.push(lines[i++].replace(/^\s*>\s?/, ''));
      html.push(`<blockquote>${markdownBlocks(quote)}</blockquote>`);
      continue;
    }
    const listStart = /^\s*(?:([-+*])|(\d+)[.)])\s+(.*)$/.exec(line);
    if (listStart) {
      const ordered = Boolean(listStart[2]);
      const items = [];
      while (i < lines.length) {
        const item = /^\s*(?:([-+*])|(\d+)[.)])\s+(.*)$/.exec(lines[i]);
        if (!item || Boolean(item[2]) !== ordered) break;
        items.push(`<li>${inlineMarkdown(item[3])}</li>`);
        i++;
      }
      const tag = ordered ? 'ol' : 'ul';
      html.push(`<${tag}>${items.join('')}</${tag}>`);
      continue;
    }
    if (i + 1 < lines.length && /^\s*\|?\s*:?-{3,}/.test(lines[i + 1])) {
      const cells = (value) => value.trim().replace(/^\||\|$/g, '').split('|').map((cell) => cell.trim());
      const headers = cells(line);
      i += 2;
      const rows = [];
      while (i < lines.length && lines[i].includes('|') && lines[i].trim()) rows.push(cells(lines[i++]));
      html.push(`<table><thead><tr>${headers.map((cell) => `<th>${inlineMarkdown(cell)}</th>`).join('')}</tr></thead><tbody>${rows.map((row) => `<tr>${headers.map((_, col) => `<td>${inlineMarkdown(row[col] || '')}</td>`).join('')}</tr>`).join('')}</tbody></table>`);
      continue;
    }
    const paragraph = [line.trim()];
    i++;
    while (i < lines.length && lines[i].trim() && !isBlockStart(lines, i)) paragraph.push(lines[i++].trim());
    html.push(`<p>${inlineMarkdown(paragraph.join(' '))}</p>`);
  }
  return html.join('');
}

export function markdownToHtml(markdown) {
  return markdownBlocks(String(markdown || '').replace(/\r\n?/g, '\n').split('\n'));
}

function inlineHtml(node) {
  if (typeof node === 'string') return escapeMarkdown(node);
  const content = node.children.map(inlineHtml).join('');
  switch (node.tag) {
    case 'br': return '  \n';
    case 'strong': case 'b': return `**${content}**`;
    case 'em': case 'i': return `*${content}*`;
    case 'del': case 's': case 'strike': return `~~${content}~~`;
    case 'code': return `\`${textOf(node).replace(/`/g, '\\`')}\``;
    case 'a': return `[${content}](${node.attrs.href || ''})`;
    case 'img': return `![${escapeMarkdown(node.attrs.alt || '')}](${node.attrs.src || ''})`;
    default: return content;
  }
}

function blockHtml(node) {
  if (typeof node === 'string') return node.trim() ? escapeMarkdown(node.trim()) : '';
  if (/^h[1-6]$/.test(node.tag)) return `${'#'.repeat(Number(node.tag[1]))} ${node.children.map(inlineHtml).join('').trim()}`;
  if (node.tag === 'p' || node.tag === 'div') return node.children.map(inlineHtml).join('').trim();
  if (node.tag === 'blockquote') return node.children.map(blockHtml).filter(Boolean).join('\n\n').split('\n').map((line) => `> ${line}`).join('\n');
  if (node.tag === 'pre') return `\`\`\`\n${textOf(node).replace(/\n$/, '')}\n\`\`\``;
  if (node.tag === 'hr') return '---';
  if (node.tag === 'ul' || node.tag === 'ol') {
    let number = 1;
    return node.children.filter((child) => typeof child !== 'string' && child.tag === 'li')
      .map((item) => `${node.tag === 'ol' ? `${number++}.` : '-'} ${item.children.map(inlineHtml).join('').trim()}`).join('\n');
  }
  if (node.tag === 'table') {
    const rows = tableRows(node).map((row) => tableCells(row).map((cell) => cell.children.map(inlineHtml).join('').replace(/\|/g, '\\|').replace(/\n/g, ' ').trim()));
    if (!rows.length) return '';
    const header = rows[0];
    return [header, header.map(() => '---'), ...rows.slice(1)].map((row) => `| ${row.join(' | ')} |`).join('\n');
  }
  return node.children.map(blockHtml).filter(Boolean).join('\n\n');
}

export function htmlToMarkdown(html) {
  const tree = parseMarkup(String(html || ''));
  return tree.children.map(blockHtml).filter(Boolean).join('\n\n').replace(/\n{3,}/g, '\n\n') + '\n';
}
