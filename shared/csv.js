// RFC 4180 CSV, shared by the browser grid and the Node exporters.

export function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;
  let hasField = false;
  const src = String(text).replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');

  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (quoted) {
      if (ch === '"') {
        if (src[i + 1] === '"') { field += '"'; i++; }
        else quoted = false;
      } else field += ch;
      continue;
    }
    if (ch === '"' && field === '') { quoted = true; hasField = true; continue; }
    if (ch === ',') { row.push(field); field = ''; hasField = true; continue; }
    if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; hasField = false; continue; }
    field += ch;
  }
  if (hasField || field !== '' || row.length) { row.push(field); rows.push(row); }
  // Trailing newline produces one empty row; drop it.
  if (rows.length && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === '') rows.pop();
  return rows;
}

export function toCsv(rows) {
  return rows
    .map((row) => row
      .map((cell) => {
        const text = cell === undefined || cell === null ? '' : String(cell);
        return /[",\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
      })
      .join(','))
    .join('\r\n');
}
