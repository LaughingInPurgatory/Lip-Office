// Lip Office — HTTP server. Node's own http + fs, no framework, no packages.
import http from 'node:http';
import { isIP } from 'node:net';
import fs from 'node:fs/promises';
import path from 'node:path';
import { randomBytes } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { parseDocument, serializeDocument, mimeFor } from './lib/convert.js';
import { createStore, safeRelativePath } from './lib/store.js';
import { kindOf, extensionOf, SAVE_FORMATS } from './shared/model.js';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT || 3000);
const IN_CONTAINER = process.env.LIP_OFFICE_CONTAINER === '1';
const HOST = process.env.HOST || (IN_CONTAINER ? '0.0.0.0' : '127.0.0.1');
const normalizedHost = HOST.toLowerCase().replace(/^\[|\]$/g, '');
const REMOTE_BIND = normalizedHost !== 'localhost' && normalizedHost !== '::1' && !normalizedHost.startsWith('127.');
if (REMOTE_BIND && !IN_CONTAINER) {
  throw new Error('Lip Office runs without login or HTTPS and must stay local; set HOST=127.0.0.1');
}

const MAX_UPLOAD = 32 * 1024 * 1024;
const SOURCE_DOCX_TTL = 4 * 60 * 60 * 1000;
const SOURCE_DOCX_CACHE_LIMIT = 128 * 1024 * 1024;
const sourceDocxCache = new Map();

const store = createStore(path.join(ROOT, 'documents'));
await Promise.all(['Docs', 'Spreadsheets'].map((folder) => store.ensureDirectory(folder)));

const STATIC_ROOTS = [
  { prefix: '/shared/', directory: path.join(ROOT, 'shared') },
  { prefix: '/', directory: path.join(ROOT, 'public') },
];

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
};

function sendJson(res, status, payload) {
  const body = Buffer.from(JSON.stringify(payload), 'utf8');
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': body.length,
    'Cache-Control': 'no-store',
  });
  res.end(body);
}

function sendError(res, status, message) {
  sendJson(res, status, { error: message });
}

function isLoopbackHostHeader(value) {
  try {
    const parsed = new URL(`http://${value}`);
    const hostname = parsed.hostname.toLowerCase().replace(/^\[|\]$/g, '');
    return !parsed.username && !parsed.password && !parsed.pathname.slice(1) &&
      (hostname === 'localhost' || hostname === '::1' || /^127(?:\.\d{1,3}){3}$/.test(hostname));
  } catch {
    return false;
  }
}

const configuredAllowedHosts = new Set(
  String(process.env.LIP_OFFICE_ALLOWED_HOSTS || '')
    .split(',')
    .map((host) => host.trim().toLowerCase())
    .filter(Boolean),
);
const allowLanHosts = IN_CONTAINER && process.env.LIP_OFFICE_ALLOW_LAN_HOSTS !== '0';

function isPrivateLanHostHeader(value) {
  try {
    const parsed = new URL(`http://${value}`);
    if (parsed.username || parsed.password || parsed.pathname !== '/' || parsed.search || parsed.hash) return false;
    const hostname = parsed.hostname.toLowerCase().replace(/^\[|\]$/g, '').replace(/\.$/, '');
    const ipVersion = isIP(hostname);
    if (ipVersion === 4) {
      const octets = hostname.split('.').map(Number);
      return octets[0] === 10 ||
        (octets[0] === 172 && octets[1] >= 16 && octets[1] <= 31) ||
        (octets[0] === 192 && octets[1] === 168) ||
        (octets[0] === 169 && octets[1] === 254);
    }
    if (ipVersion === 6) {
      const firstBlock = Number.parseInt(hostname.split(':')[0] || '0', 16);
      return (firstBlock & 0xfe00) === 0xfc00 || (firstBlock & 0xffc0) === 0xfe80;
    }
    return hostname.endsWith('.local') || hostname.endsWith('.lan');
  } catch {
    return false;
  }
}

function isAllowedHostHeader(value) {
  const host = String(value || '').toLowerCase();
  return configuredAllowedHosts.has(host) || isLoopbackHostHeader(host) ||
    (allowLanHosts && isPrivateLanHostHeader(host));
}

function sameOriginMutation(req) {
  if (!['POST', 'PUT', 'DELETE'].includes(req.method)) return true;
  if (req.headers['sec-fetch-site'] === 'cross-site') return false;
  const origin = req.headers.origin;
  if (!origin || origin === 'null') return false;
  try {
    const parsed = new URL(origin);
    const protocol = 'http:';
    return parsed.protocol === protocol && parsed.host.toLowerCase() === String(req.headers.host || '').toLowerCase();
  } catch {
    return false;
  }
}

function setSecurityHeaders(res) {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('X-Frame-Options', 'DENY');
  res.setHeader('Referrer-Policy', 'no-referrer');
  res.setHeader('Permissions-Policy', 'camera=(), microphone=(), geolocation=()');
  res.setHeader('Cross-Origin-Resource-Policy', 'same-origin');
  res.setHeader('Content-Security-Policy', "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; font-src 'self' data:; connect-src 'self'; object-src 'none'; base-uri 'none'; form-action 'self'; frame-ancestors 'none'");
  res.setHeader('Cache-Control', 'no-store');
}

async function readBody(req) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > MAX_UPLOAD) throw Object.assign(new Error('file is too large (32 MB limit)'), { status: 413 });
    chunks.push(chunk);
  }
  if (!chunks.length) return {};
  try {
    return JSON.parse(Buffer.concat(chunks).toString('utf8'));
  } catch {
    throw Object.assign(new Error('request body is not valid JSON'), { status: 400 });
  }
}

function decodeBase64(value) {
  if (typeof value !== 'string' || value === '') throw Object.assign(new Error('missing file data'), { status: 400 });
  const buffer = Buffer.from(value, 'base64');
  if (buffer.length === 0) throw Object.assign(new Error('file data is empty or not valid base64'), { status: 400 });
  if (buffer.length > MAX_UPLOAD) throw Object.assign(new Error('file is too large (32 MB limit)'), { status: 413 });
  return buffer;
}

function attachSourceDocx(parsed, name, buffer) {
  if (parsed.kind !== 'word' || !/^docx$/i.test(path.extname(name))) return parsed;
  const now = Date.now();
  for (const [token, entry] of sourceDocxCache) if (entry.expires <= now) sourceDocxCache.delete(token);
  let bytes = [...sourceDocxCache.values()].reduce((sum, entry) => sum + entry.buffer.length, 0);
  while (bytes + buffer.length > SOURCE_DOCX_CACHE_LIMIT && sourceDocxCache.size) {
    const [oldestToken, oldest] = sourceDocxCache.entries().next().value;
    sourceDocxCache.delete(oldestToken);
    bytes -= oldest.buffer.length;
  }
  const token = randomBytes(24).toString('hex');
  sourceDocxCache.set(token, { buffer: Buffer.from(buffer), expires: now + SOURCE_DOCX_TTL });
  parsed.doc.sourceDocxToken = token;
  return parsed;
}

function originalDocxForSave(doc) {
  if (doc?.sourceDocxUnchanged !== true) return null;
  const token = String(doc.sourceDocxToken || '');
  const entry = /^[a-f0-9]{48}$/.test(token) ? sourceDocxCache.get(token) : null;
  if (!entry || entry.expires <= Date.now()) {
    if (entry) sourceDocxCache.delete(token);
    throw Object.assign(new Error('The original Word file preservation session expired. Reopen the document before saving to retain its full formatting.'), { status: 409 });
  }
  return entry.buffer;
}

async function serveStatic(res, pathname, headOnly = false) {
  const entry = STATIC_ROOTS.find((candidate) => pathname.startsWith(candidate.prefix));
  const relative = pathname === '/' ? 'index.html' : pathname.slice(entry.prefix.length);
  const full = path.resolve(entry.directory, relative);

  if (!full.startsWith(entry.directory + path.sep)) return sendError(res, 403, 'forbidden');

  let data;
  try {
    data = await fs.readFile(full);
  } catch {
    return sendError(res, 404, 'not found');
  }

  res.writeHead(200, {
    'Content-Type': MIME_TYPES[path.extname(full).toLowerCase()] || 'application/octet-stream',
    'Content-Length': data.length,
  });
  res.end(headOnly ? undefined : data);
}

/** Names on the import path are never stored, so they skip the storage whitelist
 *  and let parseDocument explain what is wrong with the file itself. */
function displayName(value) {
  const base = String(value ?? '').split(/[\\/]/).pop().trim().slice(0, 120);
  if (!base) throw Object.assign(new Error('missing file name'), { status: 400 });
  return base;
}

async function handleApi(req, res, pathname, query) {
  if (pathname === '/api/formats') return sendJson(res, 200, SAVE_FORMATS);

  if (pathname === '/api/documents') {
    if (req.method === 'GET') return sendJson(res, 200, await store.list(query.get('path') || ''));
    return sendError(res, 405, 'method not allowed');
  }

  if (pathname === '/api/documents/folders' && req.method === 'POST') {
    const body = await readBody(req);
    const path = await store.createFolder(body.parent || '', body.name);
    return sendJson(res, 201, { path });
  }

  // Import a local file straight from the browser: parse it, do not store it.
  if (pathname === '/api/parse' && req.method === 'POST') {
    const body = await readBody(req);
    const name = displayName(body.name);
    const buffer = decodeBase64(body.data);
    const parsed = attachSourceDocx(parseDocument(name, buffer), name, buffer);
    return sendJson(res, 200, { name, kind: parsed.kind, doc: parsed.doc });
  }

  if (pathname === '/api/documents/move' && req.method === 'POST') {
    const body = await readBody(req);
    const source = safeRelativePath(body.source, { file: true });
    const destination = safeRelativePath(body.destination || '');
    return sendJson(res, 200, await store.move(source, destination));
  }

  const match = /^\/api\/documents\/(.+)$/.exec(pathname);
  if (!match) return sendError(res, 404, 'not found');

  let decodedName;
  try {
    decodedName = decodeURIComponent(match[1]);
  } catch {
    return sendError(res, 400, 'invalid document name');
  }
  const name = safeRelativePath(decodedName, { file: true });

  if (req.method === 'GET') {
    const buffer = await store.read(name);

    // Hand back the real file bytes so the browser gets a genuine .docx/.xlsx.
    if (query.has('download')) {
      const downloadName = name.split('/').at(-1);
      res.writeHead(200, {
        'Content-Type': mimeFor(extensionOf(name)),
        'Content-Length': buffer.length,
        'Content-Disposition': `attachment; filename="${downloadName}"`,
      });
      return res.end(buffer);
    }

    const parsed = attachSourceDocx(parseDocument(name, buffer), name, buffer);
    return sendJson(res, 200, { name, kind: parsed.kind, doc: parsed.doc });
  }

  if (req.method === 'PUT') {
    const body = await readBody(req);
    const kind = kindOf(name);
    if (body.kind && body.kind !== kind) {
      return sendError(res, 400, `"${name}" is a ${kind} file, not a ${body.kind} file`);
    }
    const format = String(body.format || name.split('.').pop()).toLowerCase();
    if (!SAVE_FORMATS[kind].includes(format)) {
      return sendError(res, 400, `cannot save ${kind} documents as .${format}`);
    }
    const originalDocx = kind === 'word' && format === 'docx' ? originalDocxForSave(body.doc) : null;
    const written = originalDocx ? { buffer: originalDocx, mime: mimeFor('docx') } : serializeDocument(body.doc, format, kind);
    await store.write(name, written.buffer);
    return sendJson(res, 200, { name, kind, format, bytes: written.buffer.length });
  }

  if (req.method === 'DELETE') {
    await store.remove(name);
    return sendJson(res, 200, { name, deleted: true });
  }

  return sendError(res, 405, 'method not allowed');
}

const handleRequest = async (req, res) => {
  setSecurityHeaders(res);
  try {
    const { pathname, searchParams } = new URL(req.url, 'http://localhost');
    const host = String(req.headers.host || '').toLowerCase();
    if (!isAllowedHostHeader(host)) return sendError(res, 421, 'unrecognized host');
    if (!sameOriginMutation(req)) return sendError(res, 403, 'cross-origin request blocked');

    if (pathname === '/favicon.ico') { res.writeHead(204).end(); return; }
    if (pathname.startsWith('/api/')) await handleApi(req, res, pathname, searchParams);
    else if (req.method === 'GET' || req.method === 'HEAD') await serveStatic(res, pathname, req.method === 'HEAD');
    else sendError(res, 405, 'method not allowed');
  } catch (err) {
    if (res.headersSent) { res.destroy(); return; }
    const status = err.status || (err.code === 'ENOENT' ? 404 : 500);
    if (status === 500) {
      console.error(`${req.method} ${pathname} failed:`, err);
      sendError(res, 500, 'something went wrong on the server');
    } else {
      sendError(res, status, err.message);
    }
  }
};

const server = http.createServer({ maxHeaderSize: 16 * 1024 }, handleRequest);

server.headersTimeout = 10_000;
server.requestTimeout = 30_000;
server.keepAliveTimeout = 5_000;
server.maxRequestsPerSocket = 1_000;

server.listen(PORT, HOST, () => {
  console.log(`Lip Office running at http://${HOST}:${PORT}`);
  console.log(`Documents are stored in ${store.root}`);
});
