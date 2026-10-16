// Spreadsheet formula engine.
//
// Shared verbatim by the browser grid and the Node tests, so it must not touch
// the DOM or any Node API. Cells hold their raw typed text ("=SUM(A1:A2)",
// "12", "hello"); evaluation happens here on demand.

export class FormulaError extends Error {
  constructor(code) {
    super(code);
    this.name = 'FormulaError';
    this.code = code;
  }
}

const LETTERS = /^([A-Z]+)([0-9]+)$/;

export function colToNum(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function numToCol(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = (n - r - 1) / 26;
  }
  return s;
}

export function parseRef(ref) {
  const m = LETTERS.exec(String(ref).toUpperCase());
  return m ? { col: colToNum(m[1]), row: Number(m[2]) } : null;
}

export function toRef(col, row) {
  return numToCol(col) + row;
}

// ---------------------------------------------------------------- tokenizer

const TOKEN_RE = /\s*(?:(\d+\.?\d*(?:[eE][+-]?\d+)?|\.\d+(?:[eE][+-]?\d+)?)|"((?:[^"]|"")*)"|(\$?[A-Za-z]{1,3}\$?[0-9]{1,7})|([A-Za-z_][A-Za-z0-9_.]*)|(<=|>=|<>|[+\-*/^%&(),:;<>=]))/g;

function tokenize(src) {
  const tokens = [];
  TOKEN_RE.lastIndex = 0;
  let cursor = 0;
  let m;
  while ((m = TOKEN_RE.exec(src))) {
    if (m.index !== cursor && src.slice(cursor, m.index).trim() !== '') throw new FormulaError('#ERROR!');
    cursor = TOKEN_RE.lastIndex;
    if (m[1] !== undefined) tokens.push({ t: 'num', v: Number(m[1]) });
    else if (m[2] !== undefined) tokens.push({ t: 'str', v: m[2].replace(/""/g, '"') });
    else if (m[3] !== undefined) tokens.push({ t: 'ref', v: m[3].replace(/\$/g, '').toUpperCase() });
    else if (m[4] !== undefined) tokens.push({ t: 'name', v: m[4] });
    else tokens.push({ t: 'op', v: m[5] });
  }
  if (src.slice(cursor).trim() !== '') throw new FormulaError('#ERROR!');
  tokens.push({ t: 'eof' });
  return tokens;
}

// ------------------------------------------------------------------- parser

function parse(src) {
  const tokens = tokenize(src);
  let i = 0;
  const peek = () => tokens[i];
  const eat = (value) => {
    if (tokens[i].t === 'op' && tokens[i].v === value) { i++; return true; }
    return false;
  };

  function comparison() {
    let node = concat();
    while (['=', '<>', '<', '>', '<=', '>='].includes(peek().v) && peek().t === 'op') {
      const op = tokens[i++].v;
      node = { t: 'bin', op, a: node, b: concat() };
    }
    return node;
  }

  function concat() {
    let node = add();
    while (eat('&')) node = { t: 'bin', op: '&', a: node, b: add() };
    return node;
  }

  function add() {
    let node = mul();
    while (peek().t === 'op' && (peek().v === '+' || peek().v === '-')) {
      const op = tokens[i++].v;
      node = { t: 'bin', op, a: node, b: mul() };
    }
    return node;
  }

  function mul() {
    let node = power();
    while (peek().t === 'op' && (peek().v === '*' || peek().v === '/')) {
      const op = tokens[i++].v;
      node = { t: 'bin', op, a: node, b: power() };
    }
    return node;
  }

  // Excel binds unary minus tighter than ^, so -2^2 is 4, not -4.
  function power() {
    const base = signed();
    if (eat('^')) return { t: 'bin', op: '^', a: base, b: power() };
    return base;
  }

  function signed() {
    if (peek().t === 'op' && (peek().v === '-' || peek().v === '+')) {
      const op = tokens[i++].v;
      return op === '-' ? { t: 'unary', op, a: signed() } : signed();
    }
    return percent();
  }

  function percent() {
    let node = primary();
    while (eat('%')) node = { t: 'post', op: '%', a: node };
    return node;
  }

  function primary() {
    const token = peek();
    if (token.t === 'num') { i++; return { t: 'num', v: token.v }; }
    if (token.t === 'str') { i++; return { t: 'str', v: token.v }; }
    if (token.t === 'ref') {
      i++;
      if (eat(':')) {
        const to = peek();
        if (to.t !== 'ref') throw new FormulaError('#ERROR!');
        i++;
        return { t: 'range', from: token.v, to: to.v };
      }
      return { t: 'ref', v: token.v };
    }
    if (token.t === 'name') {
      i++;
      const upper = token.v.toUpperCase().replace(/^_(?:XLFN|XLWS)\./, '');
      if (upper === 'TRUE') return { t: 'bool', v: true };
      if (upper === 'FALSE') return { t: 'bool', v: false };
      if (!eat('(')) throw new FormulaError('#NAME?');
      const args = [];
      if (!eat(')')) {
        do {
          args.push(comparison());
        } while (eat(','));
        if (!eat(')')) throw new FormulaError('#ERROR!');
      }
      return { t: 'call', name: upper, args };
    }
    if (eat('(')) {
      const node = comparison();
      if (!eat(')')) throw new FormulaError('#ERROR!');
      return node;
    }
    throw new FormulaError('#ERROR!');
  }

  const ast = comparison();
  if (peek().t !== 'eof') throw new FormulaError('#ERROR!');
  return ast;
}

// ------------------------------------------------------------------ values

function isBlank(v) {
  return v === '' || v === null || v === undefined;
}

function toNumber(v) {
  if (Array.isArray(v)) return v.length ? toNumber(v[0]) : 0;
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isBlank(v)) return 0;
  const text = String(v).trim();
  if (text === '') return 0;
  const n = Number(text);
  if (!Number.isFinite(n)) throw new FormulaError('#VALUE!');
  return n;
}

function toText(v) {
  if (Array.isArray(v)) return toText(v.length ? v[0] : '');
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isBlank(v)) return '';
  return String(v);
}

function toBool(v) {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (isBlank(v)) return false;
  const text = String(v).trim().toUpperCase();
  if (text === 'TRUE') return true;
  if (text === 'FALSE') return false;
  const n = Number(text);
  return Number.isFinite(n) ? n !== 0 : true;
}

function flatten(values) {
  const out = [];
  for (const v of values) {
    if (Array.isArray(v)) out.push(...flatten(v));
    else out.push(v);
  }
  return out;
}

function numbers(values) {
  const out = [];
  for (const v of flatten(values)) {
    if (typeof v === 'number') out.push(v);
    else if (typeof v === 'boolean') out.push(v ? 1 : 0);
    else if (!isBlank(v)) {
      const n = Number(String(v).trim());
      if (String(v).trim() !== '' && Number.isFinite(n)) out.push(n);
    }
  }
  return out;
}

function compare(op, a, b) {
  const eq = (x, y) => {
    if (typeof x === 'number' && typeof y === 'number') return x === y;
    if (typeof x === 'boolean' || typeof y === 'boolean') return toBool(x) === toBool(y);
    return toText(x).toLowerCase() === toText(y).toLowerCase();
  };
  if (op === '=') return eq(a, b);
  if (op === '<>') return !eq(a, b);

  let x = a;
  let y = b;
  if (typeof x === 'number' && typeof y === 'string') {
    const n = Number(y.trim());
    if (Number.isFinite(n)) y = n;
  } else if (typeof y === 'number' && typeof x === 'string') {
    const n = Number(x.trim());
    if (Number.isFinite(n)) x = n;
  }
  if (typeof x === 'number' && typeof y === 'number') {
    if (op === '<') return x < y;
    if (op === '>') return x > y;
    if (op === '<=') return x <= y;
    return x >= y;
  }
  const left = toText(x).toLowerCase();
  const right = toText(y).toLowerCase();
  if (op === '<') return left < right;
  if (op === '>') return left > right;
  if (op === '<=') return left <= right;
  return left >= right;
}

function list(value) {
  return flatten([value]);
}

function matchesCriteria(value, criteria) {
  const source = toText(criteria);
  const match = /^(<=|>=|<>|=|<|>)(.*)$/.exec(source);
  if (match) return compare(match[1], value, match[2]);
  if (/[?*]/.test(source)) {
    if (isBlank(value)) return false;
    let pattern = '^';
    for (let i = 0; i < source.length; i++) {
      if (source[i] === '~' && i + 1 < source.length) pattern += source[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      else if (source[i] === '*') pattern += '.*';
      else if (source[i] === '?') pattern += '.';
      else pattern += source[i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    return new RegExp(`${pattern}$`, 'i').test(toText(value));
  }
  return compare('=', value, criteria);
}

function filteredValues(rangeNodes, criteriaNodes, ctx) {
  const ranges = rangeNodes.map((node) => list(ctx.eval(node)));
  if (ranges.some((items) => items.length !== ranges[0]?.length)) throw new FormulaError('#VALUE!');
  const criteria = criteriaNodes.map((node) => ctx.eval(node));
  const count = ranges[0]?.length || 0;
  const indices = [];
  for (let i = 0; i < count; i++) {
    if (ranges.every((items, j) => matchesCriteria(items[i], criteria[j]))) indices.push(i);
  }
  return indices;
}

function tableCell(node, row, col, ctx) {
  if (node.t === 'range') {
    const a = parseRef(node.from);
    const b = parseRef(node.to);
    if (!a || !b) throw new FormulaError('#REF!');
    const height = Math.abs(b.row - a.row) + 1;
    const width = Math.abs(b.col - a.col) + 1;
    if (row < 1 || col < 1 || row > height || col > width) throw new FormulaError('#REF!');
    const r = Math.min(a.row, b.row) + row - 1;
    const c = Math.min(a.col, b.col) + col - 1;
    return ctx.get(toRef(c, r));
  }
  if (row !== 1 || col !== 1) throw new FormulaError('#REF!');
  return ctx.eval(node);
}

function arrayShape(node) {
  if (node.t !== 'range') return { rows: 1, cols: 1 };
  const a = parseRef(node.from);
  const b = parseRef(node.to);
  if (!a || !b) throw new FormulaError('#REF!');
  return { rows: Math.abs(b.row - a.row) + 1, cols: Math.abs(b.col - a.col) + 1 };
}

function excelSerial(date) {
  const days = Math.floor((Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()) - Date.UTC(1899, 11, 30)) / 86400000);
  return days > 60 ? days : days - 1;
}

function asDate(value) {
  if (typeof value === 'number') {
    const serial = Math.floor(value);
    return new Date(Date.UTC(1899, 11, 30) + (serial < 60 ? serial + 1 : serial) * 86400000);
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) throw new FormulaError('#VALUE!');
  return date;
}

// ---------------------------------------------------------------- functions

const FUNCS = {
  SUMIF: (args, ctx) => {
    const values = list(ctx.eval(args[0]));
    const criteria = ctx.eval(args[1]);
    const sums = list(args[2] ? ctx.eval(args[2]) : values);
    return values.reduce((total, value, i) => matchesCriteria(value, criteria) ? total + toNumber(sums[i] ?? 0) : total, 0);
  },
  COUNTIF: (args, ctx) => list(ctx.eval(args[0])).filter((value) => matchesCriteria(value, ctx.eval(args[1]))).length,
  AVERAGEIF: (args, ctx) => {
    const values = list(ctx.eval(args[0]));
    const criteria = ctx.eval(args[1]);
    const sums = list(args[2] ? ctx.eval(args[2]) : values);
    const selected = values.flatMap((value, i) => matchesCriteria(value, criteria) ? [sums[i]] : []).filter((v) => !isBlank(v));
    if (!selected.length) throw new FormulaError('#DIV/0!');
    const numeric = numbers(selected);
    if (!numeric.length) throw new FormulaError('#DIV/0!');
    return numeric.reduce((a, b) => a + b, 0) / numeric.length;
  },
  SUMIFS: (args, ctx) => {
    if (args.length < 3 || args.length % 2 !== 1) throw new FormulaError('#VALUE!');
    const sums = list(ctx.eval(args[0]));
    const ranges = args.slice(1).filter((_, i) => i % 2 === 0);
    if (ranges.some((range) => list(ctx.eval(range)).length !== sums.length)) throw new FormulaError('#VALUE!');
    const indexes = filteredValues(ranges, args.slice(2).filter((_, i) => i % 2 === 0), ctx);
    return indexes.reduce((total, i) => total + toNumber(sums[i] ?? 0), 0);
  },
  COUNTIFS: (args, ctx) => {
    if (args.length < 2 || args.length % 2 !== 0) throw new FormulaError('#VALUE!');
    return filteredValues(args.filter((_, i) => i % 2 === 0), args.filter((_, i) => i % 2 === 1), ctx).length;
  },
  AVERAGEIFS: (args, ctx) => {
    if (args.length < 3 || args.length % 2 !== 1) throw new FormulaError('#VALUE!');
    const sums = list(ctx.eval(args[0]));
    const ranges = args.slice(1).filter((_, i) => i % 2 === 0);
    if (ranges.some((range) => list(ctx.eval(range)).length !== sums.length)) throw new FormulaError('#VALUE!');
    const indexes = filteredValues(ranges, args.slice(2).filter((_, i) => i % 2 === 0), ctx);
    const values = indexes.map((i) => sums[i]).filter((v) => !isBlank(v));
    const numeric = numbers(values);
    if (!numeric.length) throw new FormulaError('#DIV/0!');
    return numeric.reduce((a, b) => a + b, 0) / numeric.length;
  },
  MAXIFS: (args, ctx) => {
    if (args.length < 3 || args.length % 2 !== 1) throw new FormulaError('#VALUE!');
    const values = list(ctx.eval(args[0]));
    const ranges = args.slice(1).filter((_, i) => i % 2 === 0);
    if (ranges.some((range) => list(ctx.eval(range)).length !== values.length)) throw new FormulaError('#VALUE!');
    const indexes = filteredValues(ranges, args.slice(2).filter((_, i) => i % 2 === 0), ctx);
    const selected = numbers(indexes.map((i) => values[i]));
    return selected.length ? Math.max(...selected) : 0;
  },
  MINIFS: (args, ctx) => {
    if (args.length < 3 || args.length % 2 !== 1) throw new FormulaError('#VALUE!');
    const values = list(ctx.eval(args[0]));
    const ranges = args.slice(1).filter((_, i) => i % 2 === 0);
    if (ranges.some((range) => list(ctx.eval(range)).length !== values.length)) throw new FormulaError('#VALUE!');
    const indexes = filteredValues(ranges, args.slice(2).filter((_, i) => i % 2 === 0), ctx);
    const selected = numbers(indexes.map((i) => values[i]));
    return selected.length ? Math.min(...selected) : 0;
  },
  COUNTBLANK: (args, ctx) => list(ctx.eval(args[0])).filter(isBlank).length,
  SUMPRODUCT: (args, ctx) => {
    const arrays = args.map((arg) => list(ctx.eval(arg)));
    if (!arrays.length) return 0;
    if (arrays.some((items) => items.length !== arrays[0].length)) throw new FormulaError('#VALUE!');
    return arrays[0].reduce((total, _, i) => total + arrays.reduce((product, items) => product * toNumber(items[i]), 1), 0);
  },
  IFNA: (args, ctx) => {
    try { return ctx.eval(args[0]); }
    catch (err) { if (err instanceof FormulaError && err.code === '#N/A') return ctx.eval(args[1]); throw err; }
  },
  IFS: (args, ctx) => {
    for (let i = 0; i + 1 < args.length; i += 2) if (toBool(ctx.eval(args[i]))) return ctx.eval(args[i + 1]);
    if (args.length % 2) return ctx.eval(args.at(-1));
    throw new FormulaError('#N/A');
  },
  SWITCH: (args, ctx) => {
    if (args.length < 3) throw new FormulaError('#VALUE!');
    const target = ctx.eval(args[0]);
    const pairEnd = args.length % 2 === 0 ? args.length - 1 : args.length;
    for (let i = 1; i + 1 < pairEnd; i += 2) if (compare('=', target, ctx.eval(args[i]))) return ctx.eval(args[i + 1]);
    if (args.length % 2 === 0) return ctx.eval(args.at(-1));
    throw new FormulaError('#N/A');
  },
  CHOOSE: (args, ctx) => {
    const index = Math.trunc(ctx.num(ctx.eval(args[0])));
    if (index < 1 || index >= args.length) throw new FormulaError('#VALUE!');
    return ctx.eval(args[index]);
  },
  XOR: (args, ctx) => flatten(ctx.values(args)).filter(toBool).length % 2 === 1,
  ISBLANK: (args, ctx) => isBlank(ctx.eval(args[0])),
  ISNUMBER: (args, ctx) => typeof ctx.eval(args[0]) === 'number',
  ISTEXT: (args, ctx) => typeof ctx.eval(args[0]) === 'string' && !isBlank(ctx.eval(args[0])),
  ISLOGICAL: (args, ctx) => typeof ctx.eval(args[0]) === 'boolean',
  ISERROR: (args, ctx) => { try { ctx.eval(args[0]); return false; } catch (err) { return err instanceof FormulaError; } },
  ISNA: (args, ctx) => { try { ctx.eval(args[0]); return false; } catch (err) { return err instanceof FormulaError && err.code === '#N/A'; } },
  ISERR: (args, ctx) => { try { ctx.eval(args[0]); return false; } catch (err) { return err instanceof FormulaError && err.code !== '#N/A'; } },
  ISNONTEXT: (args, ctx) => typeof ctx.eval(args[0]) !== 'string',
  N: (args, ctx) => { const value = ctx.eval(args[0]); return typeof value === 'number' ? value : typeof value === 'boolean' ? Number(value) : 0; },
  T: (args, ctx) => { const value = ctx.eval(args[0]); return typeof value === 'string' ? value : ''; },
  NA: () => { throw new FormulaError('#N/A'); },
  VLOOKUP: (args, ctx) => {
    const shape = arrayShape(args[1]);
    const col = Math.trunc(ctx.num(ctx.eval(args[2])));
    if (col < 1 || col > shape.cols) throw new FormulaError('#REF!');
    const lookup = ctx.eval(args[0]);
    const approximate = args.length > 3 ? toBool(ctx.eval(args[3])) : true;
    let candidate = null;
    for (let row = 1; row <= shape.rows; row++) {
      const first = tableCell(args[1], row, 1, ctx);
      if (compare('=', first, lookup)) return tableCell(args[1], row, col, ctx);
      if (approximate && compare('<=', first, lookup)) candidate = row;
    }
    if (candidate !== null) return tableCell(args[1], candidate, col, ctx);
    throw new FormulaError('#N/A');
  },
  HLOOKUP: (args, ctx) => {
    const shape = arrayShape(args[1]);
    const row = Math.trunc(ctx.num(ctx.eval(args[2])));
    if (row < 1 || row > shape.rows) throw new FormulaError('#REF!');
    const lookup = ctx.eval(args[0]);
    const approximate = args.length > 3 ? toBool(ctx.eval(args[3])) : true;
    let candidate = null;
    for (let col = 1; col <= shape.cols; col++) {
      const first = tableCell(args[1], 1, col, ctx);
      if (compare('=', first, lookup)) return tableCell(args[1], row, col, ctx);
      if (approximate && compare('<=', first, lookup)) candidate = col;
    }
    if (candidate !== null) return tableCell(args[1], row, candidate, ctx);
    throw new FormulaError('#N/A');
  },
  XLOOKUP: (args, ctx) => {
    const lookup = ctx.eval(args[0]);
    const keys = list(ctx.eval(args[1]));
    const values = list(ctx.eval(args[2]));
    if (keys.length !== values.length) throw new FormulaError('#VALUE!');
    const mode = args.length > 4 ? ctx.num(ctx.eval(args[4])) : 0;
    const position = keys.findIndex((item) => compare('=', item, lookup));
    if (position >= 0) return values[position];
    if (mode === -1 || mode === 1) {
      const matches = keys.map((item, i) => ({ item, i })).filter(({ item }) => mode === -1 ? compare('<=', item, lookup) : compare('>=', item, lookup));
      const candidate = mode === -1 ? matches.at(-1) : matches[0];
      if (candidate) return values[candidate.i];
    }
    if (args.length > 3) return ctx.eval(args[3]);
    throw new FormulaError('#N/A');
  },
  INDEX: (args, ctx) => {
    const row = Math.trunc(ctx.num(ctx.eval(args[1])));
    const col = args.length > 2 ? Math.trunc(ctx.num(ctx.eval(args[2]))) : 1;
    if (!row || !col) throw new FormulaError('#VALUE!');
    return tableCell(args[0], row, col, ctx);
  },
  MATCH: (args, ctx) => {
    const lookup = ctx.eval(args[0]);
    const values = list(ctx.eval(args[1]));
    const mode = args.length > 2 ? ctx.num(ctx.eval(args[2])) : 1;
    const exact = values.findIndex((value) => compare('=', value, lookup));
    if (exact >= 0) return exact + 1;
    if (mode === 0) throw new FormulaError('#N/A');
    let candidate = -1;
    values.forEach((value, index) => {
      if (mode > 0 && compare('<=', value, lookup)) candidate = index;
      if (mode < 0 && candidate < 0 && compare('>=', value, lookup)) candidate = index;
    });
    if (candidate < 0) throw new FormulaError('#N/A');
    return candidate + 1;
  },
  ROWS: (args) => arrayShape(args[0]).rows,
  COLUMNS: (args) => arrayShape(args[0]).cols,
  LARGE: (args, ctx) => {
    const values = numbers(ctx.values([args[0]])).sort((a, b) => b - a);
    const k = Math.trunc(ctx.num(ctx.eval(args[1])));
    if (k < 1 || k > values.length) throw new FormulaError('#NUM!');
    return values[k - 1];
  },
  SMALL: (args, ctx) => {
    const values = numbers(ctx.values([args[0]])).sort((a, b) => a - b);
    const k = Math.trunc(ctx.num(ctx.eval(args[1])));
    if (k < 1 || k > values.length) throw new FormulaError('#NUM!');
    return values[k - 1];
  },
  MEDIAN: (args, ctx) => {
    const values = numbers(ctx.values(args)).sort((a, b) => a - b);
    if (!values.length) throw new FormulaError('#NUM!');
    const mid = Math.floor(values.length / 2);
    return values.length % 2 ? values[mid] : (values[mid - 1] + values[mid]) / 2;
  },
  'MODE.SNGL': (args, ctx) => {
    const values = numbers(ctx.values(args));
    const counts = new Map();
    for (const value of values) counts.set(value, (counts.get(value) || 0) + 1);
    const best = [...counts.entries()].sort((a, b) => b[1] - a[1])[0];
    if (!best || best[1] < 2) throw new FormulaError('#N/A');
    return best[0];
  },
  MODE: (args, ctx) => FUNCS['MODE.SNGL'](args, ctx),
  'RANK.EQ': (args, ctx) => {
    const value = ctx.num(ctx.eval(args[0]));
    const values = numbers(ctx.values([args[1]]));
    const ascending = args.length > 2 && ctx.num(ctx.eval(args[2])) !== 0;
    return 1 + values.filter((n) => ascending ? n < value : n > value).length;
  },
  RANK: (args, ctx) => FUNCS['RANK.EQ'](args, ctx),
  'PERCENTILE.INC': (args, ctx) => {
    const values = numbers(ctx.values([args[0]])).sort((a, b) => a - b);
    const k = ctx.num(ctx.eval(args[1]));
    if (!values.length || k < 0 || k > 1) throw new FormulaError('#NUM!');
    const rank = k * (values.length - 1);
    const low = Math.floor(rank); const fraction = rank - low;
    return values[low] + (values[Math.min(low + 1, values.length - 1)] - values[low]) * fraction;
  },
  'QUARTILE.INC': (args, ctx) => {
    const quart = ctx.num(ctx.eval(args[1]));
    if (quart < 0 || quart > 4) throw new FormulaError('#NUM!');
    return FUNCS['PERCENTILE.INC']([args[0], { t: 'num', v: quart / 4 }], ctx);
  },
  SIGN: (args, ctx) => Math.sign(ctx.num(ctx.eval(args[0]))),
  QUOTIENT: (args, ctx) => {
    const divisor = ctx.num(ctx.eval(args[1]));
    if (divisor === 0) throw new FormulaError('#DIV/0!');
    return Math.trunc(ctx.num(ctx.eval(args[0])) / divisor);
  },
  TRUNC: (args, ctx) => {
    const n = ctx.num(ctx.eval(args[0]));
    const digits = args.length > 1 ? ctx.num(ctx.eval(args[1])) : 0;
    const factor = 10 ** digits;
    return Math.trunc(n * factor) / factor;
  },
  CEILING: (args, ctx) => {
    const n = ctx.num(ctx.eval(args[0])); const s = ctx.num(ctx.eval(args[1]));
    if (!s) throw new FormulaError('#DIV/0!');
    return Math.ceil(n / s) * s;
  },
  FLOOR: (args, ctx) => {
    const n = ctx.num(ctx.eval(args[0])); const s = ctx.num(ctx.eval(args[1]));
    if (!s) throw new FormulaError('#DIV/0!');
    return Math.floor(n / s) * s;
  },
  MROUND: (args, ctx) => {
    const n = ctx.num(ctx.eval(args[0])); const s = ctx.num(ctx.eval(args[1]));
    if (!s || Math.sign(n) !== Math.sign(s) && n !== 0) throw new FormulaError('#NUM!');
    return Math.round(n / s) * s;
  },
  LOG: (args, ctx) => {
    const n = ctx.num(ctx.eval(args[0])); const base = args.length > 1 ? ctx.num(ctx.eval(args[1])) : 10;
    if (n <= 0 || base <= 0 || base === 1) throw new FormulaError('#NUM!');
    return Math.log(n) / Math.log(base);
  },
  LN: (args, ctx) => { const n = ctx.num(ctx.eval(args[0])); if (n <= 0) throw new FormulaError('#NUM!'); return Math.log(n); },
  LOG10: (args, ctx) => { const n = ctx.num(ctx.eval(args[0])); if (n <= 0) throw new FormulaError('#NUM!'); return Math.log10(n); },
  EXP: (args, ctx) => Math.exp(ctx.num(ctx.eval(args[0]))),
  PI: () => Math.PI,
  FACT: (args, ctx) => { const n = Math.trunc(ctx.num(ctx.eval(args[0]))); if (n < 0) throw new FormulaError('#NUM!'); let value = 1; for (let i = 2; i <= n; i++) value *= i; return value; },
  EVEN: (args, ctx) => { const n = ctx.num(ctx.eval(args[0])); return Math.ceil(Math.abs(n) / 2) * 2 * Math.sign(n || 1); },
  ODD: (args, ctx) => { const n = ctx.num(ctx.eval(args[0])); const value = Math.ceil((Math.abs(n) - 1) / 2) * 2 + 1; return value * Math.sign(n || 1); },
  SUMSQ: (args, ctx) => numbers(ctx.values(args)).reduce((total, n) => total + n * n, 0),
  RAND: () => Math.random(),
  RANDBETWEEN: (args, ctx) => {
    const low = Math.ceil(ctx.num(ctx.eval(args[0]))); const high = Math.floor(ctx.num(ctx.eval(args[1])));
    if (low > high) throw new FormulaError('#NUM!');
    return Math.floor(Math.random() * (high - low + 1)) + low;
  },
  SIN: (args, ctx) => Math.sin(ctx.num(ctx.eval(args[0]))),
  COS: (args, ctx) => Math.cos(ctx.num(ctx.eval(args[0]))),
  TAN: (args, ctx) => Math.tan(ctx.num(ctx.eval(args[0]))),
  ASIN: (args, ctx) => Math.asin(ctx.num(ctx.eval(args[0]))),
  ACOS: (args, ctx) => Math.acos(ctx.num(ctx.eval(args[0]))),
  ATAN: (args, ctx) => Math.atan(ctx.num(ctx.eval(args[0]))),
  ATAN2: (args, ctx) => Math.atan2(ctx.num(ctx.eval(args[1])), ctx.num(ctx.eval(args[0]))),
  RADIANS: (args, ctx) => ctx.num(ctx.eval(args[0])) * Math.PI / 180,
  DEGREES: (args, ctx) => ctx.num(ctx.eval(args[0])) * 180 / Math.PI,
  'STDEV.S': (args, ctx) => {
    const values = numbers(ctx.values(args));
    if (values.length < 2) throw new FormulaError('#DIV/0!');
    const mean = values.reduce((a, b) => a + b, 0) / values.length;
    return Math.sqrt(values.reduce((sum, n) => sum + (n - mean) ** 2, 0) / (values.length - 1));
  },
  STDEV: (args, ctx) => FUNCS['STDEV.S'](args, ctx),
  'STDEV.P': (args, ctx) => {
    const values = numbers(ctx.values(args));
    if (!values.length) throw new FormulaError('#DIV/0!');
    const mean = values.reduce((a, b) => a + b, 0) / values.length;
    return Math.sqrt(values.reduce((sum, n) => sum + (n - mean) ** 2, 0) / values.length);
  },
  'VAR.S': (args, ctx) => {
    const values = numbers(ctx.values(args));
    if (values.length < 2) throw new FormulaError('#DIV/0!');
    const mean = values.reduce((a, b) => a + b, 0) / values.length;
    return values.reduce((sum, n) => sum + (n - mean) ** 2, 0) / (values.length - 1);
  },
  VAR: (args, ctx) => FUNCS['VAR.S'](args, ctx),
  'VAR.P': (args, ctx) => {
    const values = numbers(ctx.values(args));
    if (!values.length) throw new FormulaError('#DIV/0!');
    const mean = values.reduce((a, b) => a + b, 0) / values.length;
    return values.reduce((sum, n) => sum + (n - mean) ** 2, 0) / values.length;
  },
  YEAR: (args, ctx) => asDate(ctx.eval(args[0])).getUTCFullYear(),
  MONTH: (args, ctx) => asDate(ctx.eval(args[0])).getUTCMonth() + 1,
  DAY: (args, ctx) => asDate(ctx.eval(args[0])).getUTCDate(),
  HOUR: (args, ctx) => Math.floor((((toNumber(ctx.eval(args[0])) % 1) + 1) % 1) * 24),
  MINUTE: (args, ctx) => Math.floor(((((toNumber(ctx.eval(args[0])) % 1) + 1) % 1) * 1440) % 60),
  SECOND: (args, ctx) => Math.floor(((((toNumber(ctx.eval(args[0])) % 1) + 1) % 1) * 86400) % 60),
  WEEKDAY: (args, ctx) => {
    const day = asDate(ctx.eval(args[0])).getUTCDay();
    const type = args.length > 1 ? Math.trunc(ctx.num(ctx.eval(args[1]))) : 1;
    if (type === 1) return day + 1;
    if (type === 2 || type === 11) return (day + 6) % 7 + 1;
    if (type === 3) return (day + 6) % 7;
    if (type >= 12 && type <= 17) return (day + type - 12) % 7 + 1;
    throw new FormulaError('#NUM!');
  },
  DAYS: (args, ctx) => excelSerial(asDate(ctx.eval(args[0]))) - excelSerial(asDate(ctx.eval(args[1]))),
  EDATE: (args, ctx) => {
    const date = asDate(ctx.eval(args[0]));
    const months = Math.trunc(ctx.num(ctx.eval(args[1])));
    const day = date.getUTCDate();
    const result = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + months, 1));
    result.setUTCDate(Math.min(day, new Date(Date.UTC(result.getUTCFullYear(), result.getUTCMonth() + 1, 0)).getUTCDate()));
    return excelSerial(result);
  },
  EOMONTH: (args, ctx) => {
    const date = asDate(ctx.eval(args[0]));
    const months = Math.trunc(ctx.num(ctx.eval(args[1])));
    return excelSerial(new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + months + 1, 0)));
  },
  TIME: (args, ctx) => {
    const hour = Math.trunc(ctx.num(ctx.eval(args[0]))); const minute = Math.trunc(ctx.num(ctx.eval(args[1]))); const second = Math.trunc(ctx.num(ctx.eval(args[2])));
    if (hour < 0 || minute < 0 || second < 0) throw new FormulaError('#NUM!');
    return (((hour * 3600 + minute * 60 + second) % 86400) + 86400) % 86400 / 86400;
  },
  TIMEVALUE: (args, ctx) => {
    const match = /^\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?\s*$/i.exec(toText(ctx.eval(args[0])));
    if (!match) throw new FormulaError('#VALUE!');
    let hour = Number(match[1]);
    const minute = Number(match[2]); const second = Number(match[3] || 0);
    if (match[4]) { if (hour < 1 || hour > 12) throw new FormulaError('#VALUE!'); hour = (hour % 12) + (match[4].toUpperCase() === 'PM' ? 12 : 0); }
    if (hour > 23 || minute > 59 || second > 59) throw new FormulaError('#VALUE!');
    return (hour * 3600 + minute * 60 + second) / 86400;
  },
  DATE: (args, ctx) => {
    const year = Math.trunc(ctx.num(ctx.eval(args[0])));
    const month = Math.trunc(ctx.num(ctx.eval(args[1])));
    const day = Math.trunc(ctx.num(ctx.eval(args[2])));
    const date = new Date(Date.UTC(year < 1900 ? year + 1900 : year, month - 1, day));
    return excelSerial(date);
  },
  TODAY: () => excelSerial(new Date()),
  NOW: () => (Date.now() - Date.UTC(1899, 11, 30)) / 86400000,
  VALUE: (args, ctx) => { const text = toText(ctx.eval(args[0])).trim(); const n = Number(text); if (!Number.isFinite(n)) throw new FormulaError('#VALUE!'); return n; },
  EXACT: (args, ctx) => toText(ctx.eval(args[0])) === toText(ctx.eval(args[1])),
  FIND: (args, ctx) => {
    const needle = toText(ctx.eval(args[0])); const text = toText(ctx.eval(args[1]));
    const start = args.length > 2 ? Math.max(1, Math.trunc(ctx.num(ctx.eval(args[2])))) : 1;
    const found = text.indexOf(needle, start - 1);
    if (found < 0) throw new FormulaError('#VALUE!');
    return found + 1;
  },
  SEARCH: (args, ctx) => {
    const needle = toText(ctx.eval(args[0])).toLowerCase(); const text = toText(ctx.eval(args[1])).toLowerCase();
    const start = args.length > 2 ? Math.max(1, Math.trunc(ctx.num(ctx.eval(args[2])))) : 1;
    const found = text.indexOf(needle, start - 1);
    if (found < 0) throw new FormulaError('#VALUE!');
    return found + 1;
  },
  SUBSTITUTE: (args, ctx) => {
    const text = toText(ctx.eval(args[0])); const oldText = toText(ctx.eval(args[1])); const replacement = toText(ctx.eval(args[2]));
    if (!oldText) return text;
    const instance = args.length > 3 ? Math.trunc(ctx.num(ctx.eval(args[3]))) : 0;
    if (instance < 1) return text.split(oldText).join(replacement);
    let count = 0;
    return text.replaceAll(oldText, (found) => (++count === instance ? replacement : found));
  },
  REPLACE: (args, ctx) => {
    const text = toText(ctx.eval(args[0])); const start = Math.trunc(ctx.num(ctx.eval(args[1]))) - 1; const count = Math.max(0, Math.trunc(ctx.num(ctx.eval(args[2]))));
    return text.slice(0, Math.max(0, start)) + toText(ctx.eval(args[3])) + text.slice(Math.max(0, start) + count);
  },
  PROPER: (args, ctx) => toText(ctx.eval(args[0])).toLowerCase().replace(/(^|[^A-Za-z0-9])([A-Za-z])/g, (_, before, letter) => before + letter.toUpperCase()),
  REPT: (args, ctx) => toText(ctx.eval(args[0])).repeat(Math.max(0, Math.min(32767, Math.trunc(ctx.num(ctx.eval(args[1])))))),
  CLEAN: (args, ctx) => toText(ctx.eval(args[0])).replace(/[\x00-\x1f]/g, ''),
  CHAR: (args, ctx) => String.fromCodePoint(Math.max(1, Math.min(255, Math.trunc(ctx.num(ctx.eval(args[0])))))),
  CODE: (args, ctx) => { const first = toText(ctx.eval(args[0]))[0]; if (!first) throw new FormulaError('#VALUE!'); return first.codePointAt(0); },
  TEXTJOIN: (args, ctx) => {
    const delimiter = toText(ctx.eval(args[0])); const ignoreEmpty = toBool(ctx.eval(args[1]));
    const values = flatten(args.slice(2).map((arg) => ctx.eval(arg))).map(toText);
    return (ignoreEmpty ? values.filter((value) => value !== '') : values).join(delimiter);
  },
  TEXTBEFORE: (args, ctx) => {
    const text = toText(ctx.eval(args[0])); const delimiter = toText(ctx.eval(args[1]));
    const instance = args.length > 2 ? Math.trunc(ctx.num(ctx.eval(args[2]))) : 1;
    if (!delimiter || !instance) throw new FormulaError('#VALUE!');
    const parts = text.split(delimiter);
    if (parts.length === 1) throw new FormulaError('#N/A');
    const index = instance > 0 ? instance : parts.length + instance;
    if (index < 1 || index >= parts.length) throw new FormulaError('#VALUE!');
    return parts.slice(0, index).join(delimiter);
  },
  TEXTAFTER: (args, ctx) => {
    const text = toText(ctx.eval(args[0])); const delimiter = toText(ctx.eval(args[1]));
    const instance = args.length > 2 ? Math.trunc(ctx.num(ctx.eval(args[2]))) : 1;
    if (!delimiter || !instance) throw new FormulaError('#VALUE!');
    const parts = text.split(delimiter);
    if (parts.length === 1) throw new FormulaError('#N/A');
    const index = instance > 0 ? instance : parts.length + instance;
    if (index < 1 || index >= parts.length) throw new FormulaError('#VALUE!');
    return parts.slice(index).join(delimiter);
  },
  SUM: (args, ctx) => numbers(ctx.values(args)).reduce((a, b) => a + b, 0),
  PRODUCT: (args, ctx) => numbers(ctx.values(args)).reduce((a, b) => a * b, 1),
  AVERAGE: (args, ctx) => {
    const list = numbers(ctx.values(args));
    if (!list.length) throw new FormulaError('#DIV/0!');
    return list.reduce((a, b) => a + b, 0) / list.length;
  },
  MIN: (args, ctx) => {
    const list = numbers(ctx.values(args));
    return list.length ? Math.min(...list) : 0;
  },
  MAX: (args, ctx) => {
    const list = numbers(ctx.values(args));
    return list.length ? Math.max(...list) : 0;
  },
  COUNT: (args, ctx) => numbers(ctx.values(args)).length,
  COUNTA: (args, ctx) => flatten(ctx.values(args)).filter((v) => !isBlank(v)).length,
  ABS: (args, ctx) => Math.abs(ctx.num(ctx.eval(args[0]))),
  SQRT: (args, ctx) => {
    const n = ctx.num(ctx.eval(args[0]));
    if (n < 0) throw new FormulaError('#NUM!');
    return Math.sqrt(n);
  },
  POWER: (args, ctx) => Math.pow(ctx.num(ctx.eval(args[0])), ctx.num(ctx.eval(args[1]))),
  INT: (args, ctx) => Math.floor(ctx.num(ctx.eval(args[0]))),
  MOD: (args, ctx) => {
    const d = ctx.num(ctx.eval(args[1]));
    if (d === 0) throw new FormulaError('#DIV/0!');
    const n = ctx.num(ctx.eval(args[0]));
    return n - d * Math.floor(n / d);
  },
  ROUND: (args, ctx) => {
    const digits = args.length > 1 ? ctx.num(ctx.eval(args[1])) : 0;
    const factor = Math.pow(10, digits);
    const scaled = ctx.num(ctx.eval(args[0])) * factor;
    // Nudge by one ULP of the scaled value so 2.345 (stored as 2.34499...) rounds to 2.35.
    return Math.round(scaled + Number.EPSILON * Math.abs(scaled)) / factor;
  },
  IF: (args, ctx) => {
    if (args.length < 2) throw new FormulaError('#VALUE!');
    if (toBool(ctx.eval(args[0]))) return ctx.eval(args[1]);
    return args.length > 2 ? ctx.eval(args[2]) : false;
  },
  IFERROR: (args, ctx) => {
    try {
      return ctx.eval(args[0]);
    } catch (err) {
      if (err instanceof FormulaError) return ctx.eval(args[1]);
      throw err;
    }
  },
  AND: (args, ctx) => flatten(ctx.values(args)).every(toBool),
  OR: (args, ctx) => flatten(ctx.values(args)).some(toBool),
  NOT: (args, ctx) => !toBool(ctx.eval(args[0])),
  LEN: (args, ctx) => toText(ctx.eval(args[0])).length,
  LEFT: (args, ctx) => {
    const n = args.length > 1 ? ctx.num(ctx.eval(args[1])) : 1;
    return toText(ctx.eval(args[0])).slice(0, Math.max(0, n));
  },
  RIGHT: (args, ctx) => {
    const n = args.length > 1 ? ctx.num(ctx.eval(args[1])) : 1;
    const text = toText(ctx.eval(args[0]));
    return n <= 0 ? '' : text.slice(-n);
  },
  MID: (args, ctx) => {
    const start = ctx.num(ctx.eval(args[1]));
    const count = ctx.num(ctx.eval(args[2]));
    return toText(ctx.eval(args[0])).slice(Math.max(0, start - 1), Math.max(0, start - 1) + Math.max(0, count));
  },
  UPPER: (args, ctx) => toText(ctx.eval(args[0])).toUpperCase(),
  LOWER: (args, ctx) => toText(ctx.eval(args[0])).toLowerCase(),
  TRIM: (args, ctx) => toText(ctx.eval(args[0])).trim().replace(/\s+/g, ' '),
  CONCAT: (args, ctx) => flatten(ctx.values(args)).map(toText).join(''),
  CONCATENATE: (args, ctx) => flatten(ctx.values(args)).map(toText).join(''),
  ROUNDUP: (args, ctx) => {
    const digits = args.length > 1 ? ctx.num(ctx.eval(args[1])) : 0;
    const factor = Math.pow(10, digits);
    const n = ctx.num(ctx.eval(args[0]));
    return Math.ceil(Math.abs(n) * factor) / factor * Math.sign(n);
  },
  ROUNDDOWN: (args, ctx) => {
    const digits = args.length > 1 ? ctx.num(ctx.eval(args[1])) : 0;
    const factor = Math.pow(10, digits);
    const n = ctx.num(ctx.eval(args[0]));
    return Math.floor(Math.abs(n) * factor) / factor * Math.sign(n);
  },
};

export function functionNames() {
  return Object.keys(FUNCS).sort();
}

// ------------------------------------------------------------------- engine

function evalNode(node, ctx) {
  switch (node.t) {
    case 'num': return node.v;
    case 'str': return node.v;
    case 'bool': return node.v;
    case 'ref': return ctx.get(node.v);
    case 'range': return ctx.range(node.from, node.to);
    case 'unary': return -ctx.num(evalNode(node.a, ctx));
    case 'post': return ctx.num(evalNode(node.a, ctx)) / 100;
    case 'call': {
      const fn = FUNCS[node.name];
      if (!fn) throw new FormulaError('#NAME?');
      const result = fn(node.args, ctx);
      if (typeof result === 'number' && !Number.isFinite(result)) throw new FormulaError('#NUM!');
      return result;
    }
    case 'bin': {
      const { op } = node;
      if (op === '&') return toText(evalNode(node.a, ctx)) + toText(evalNode(node.b, ctx));
      if (['=', '<>', '<', '>', '<=', '>='].includes(op)) {
        return compare(op, evalNode(node.a, ctx), evalNode(node.b, ctx));
      }
      const a = ctx.num(evalNode(node.a, ctx));
      const b = ctx.num(evalNode(node.b, ctx));
      if (op === '+') return a + b;
      if (op === '-') return a - b;
      if (op === '*') return a * b;
      if (op === '/') {
        if (b === 0) throw new FormulaError('#DIV/0!');
        return a / b;
      }
      if (op === '^') {
        const result = Math.pow(a, b);
        if (!Number.isFinite(result)) throw new FormulaError('#NUM!');
        return result;
      }
      throw new FormulaError('#ERROR!');
    }
    default: throw new FormulaError('#ERROR!');
  }
}

/**
 * Build an evaluator over a sheet's cells.
 * @param {Record<string, string>} cells raw typed text per A1 reference
 */
export function makeEngine(cells) {
  const source = cells || {};
  const memo = new Map();
  const active = new Set();

  function get(ref) {
    if (memo.has(ref)) return memo.get(ref);
    if (active.has(ref)) throw new FormulaError('#CYCLE!');

    const cell = source[ref];
    const raw = cell === undefined || cell === null ? '' : String(cell);
    if (raw === '') return (memo.set(ref, ''), '');

    if (raw[0] !== '=') {
      const trimmed = raw.trim();
      const n = Number(trimmed);
      return (memo.set(ref, trimmed !== '' && Number.isFinite(n) ? n : raw), memo.get(ref));
    }

    active.add(ref);
    try {
      const value = evalNode(parse(raw.slice(1)), ctx);
      memo.set(ref, value);
      return value;
    } finally {
      active.delete(ref);
    }
  }

  function range(from, to) {
    const a = parseRef(from);
    const b = parseRef(to);
    if (!a || !b) throw new FormulaError('#REF!');
    const values = [];
    for (let r = Math.min(a.row, b.row); r <= Math.max(a.row, b.row); r++) {
      for (let c = Math.min(a.col, b.col); c <= Math.max(a.col, b.col); c++) {
        values.push(get(toRef(c, r)));
      }
    }
    return values;
  }

  const ctx = {
    get,
    range,
    num: toNumber,
    eval: (node) => evalNode(node, ctx),
    values: (args) => args.map((node) => evalNode(node, ctx)),
  };

  return {
    get,
    /** Formatted display text for a cell, or the `#ERROR!` code on failure. */
    display(ref) {
      try {
        const value = get(ref);
        if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
        if (typeof value === 'number') return String(value);
        return value;
      } catch (err) {
        if (err instanceof FormulaError) return err.code;
        return '#ERROR!';
      }
    },
    isFormula(ref) {
      const cell = source[ref];
      return typeof cell === 'string' && cell[0] === '=';
    },
  };
}
