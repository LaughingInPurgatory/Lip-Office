// Document storage: a flat directory of files, addressed by name.
//
// Every name here comes from the network, so this is the path-traversal
// boundary. Names are reduced to a basename and matched against a strict
// whitelist before they are ever joined to the storage directory.
import fs from 'node:fs/promises';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { kindOf } from '../shared/model.js';

const SAFE_NAME = /^[\w][\w .()[\]-]*$/;
const RESERVED = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i;

// Every rejection here is caused by the name the caller sent, so it is a 400
// rather than a server fault.
const reject = (message) => Object.assign(new Error(message), { status: 400 });

export function safeFileName(name) {
  const base = path.basename(String(name ?? '')).trim();
  if (!base || base.length > 120) throw reject('invalid file name');
  if (base.includes('..') || !SAFE_NAME.test(base)) {
    throw reject('file names may use letters, digits, spaces, dots, dashes, underscores and parentheses only');
  }
  if (!kindOf(base)) throw reject(`unsupported file type: ${base.split('.').pop()}`);
  if (RESERVED.test(base.replace(/\.[^.]+$/, ''))) throw reject('that file name is reserved by Windows');
  return base;
}

export function safeRelativePath(value, { file = false } = {}) {
  const input = String(value ?? '');
  if (!input && !file) return '';
  if (!input || input.startsWith('/') || input.includes('\\')) throw reject('invalid path');

  const segments = input.split('/');
  for (const segment of segments) {
    if (!segment || segment.length > 120 || segment.includes('..') || !SAFE_NAME.test(segment)) {
      throw reject('invalid path');
    }
  }
  if (file && !kindOf(segments.at(-1))) throw reject(`unsupported file type: ${segments.at(-1).split('.').pop()}`);
  return segments.join('/');
}

export function createStore(directory) {
  const root = path.resolve(directory);

  function locate(name) {
    const relative = safeRelativePath(name, { file: true });
    const full = path.resolve(root, ...relative.split('/'));
    if (!full.startsWith(root + path.sep)) throw reject('invalid file name');
    return { full, relative };
  }

  async function ensureNoSymlinks(relative, { createDirectories = false, includeFinal = true, finalFile = false } = {}) {
    let current = root;
    const segments = relative ? relative.split('/') : [];
    const limit = includeFinal ? segments.length : Math.max(0, segments.length - 1);
    for (let index = 0; index < limit; index++) {
      current = path.join(current, segments[index]);
      try {
        const stat = await fs.lstat(current);
        const isFileTarget = finalFile && index === segments.length - 1;
        if (stat.isSymbolicLink() || (!isFileTarget && !stat.isDirectory())) throw reject('invalid path');
      } catch (err) {
        if (err.code !== 'ENOENT' || !createDirectories) throw err;
        await fs.mkdir(current);
      }
    }
  }

  return {
    root,

    async ensureDirectory(folder) {
      const relative = safeRelativePath(folder);
      await fs.mkdir(root, { recursive: true });
      await ensureNoSymlinks(relative, { createDirectories: true });
    },

    async createFolder(parent, name) {
      const safeParent = safeRelativePath(parent || '');
      const segment = String(name ?? '').trim();
      if (!segment || segment.includes('/') || segment.includes('\\') || segment.length > 120) throw reject('invalid folder name');
      const safeName = safeRelativePath(segment);
      if (safeName !== segment) throw reject('invalid folder name');
      const relative = [safeParent, safeName].filter(Boolean).join('/');
      await ensureNoSymlinks(relative, { includeFinal: false });
      const target = path.join(root, ...relative.split('/'));
      try {
        await fs.mkdir(target);
      } catch (err) {
        if (err.code === 'EEXIST') throw Object.assign(new Error('a file or folder with that name already exists'), { status: 409 });
        throw err;
      }
      return relative;
    },

    async list(folder = '') {
      const relative = safeRelativePath(folder);
      await ensureNoSymlinks(relative);
      const directory = relative ? path.join(root, ...relative.split('/')) : root;
      let entries;
      try {
        entries = await fs.readdir(directory, { withFileTypes: true });
      } catch (err) {
        if (err.code === 'ENOENT') return [];
        throw err;
      }

      const folders = [];
      const files = [];
      for (const entry of entries) {
        if (entry.name.startsWith('.') || entry.isSymbolicLink()) continue;
        if (entry.isDirectory()) {
          folders.push({ name: entry.name, path: [relative, entry.name].filter(Boolean).join('/') });
          continue;
        }
        if (!entry.isFile()) continue;
        const kind = kindOf(entry.name);
        if (!kind) continue;
        const stats = await fs.stat(path.join(directory, entry.name));
        files.push({
          name: entry.name,
          path: [relative, entry.name].filter(Boolean).join('/'),
          kind,
          size: stats.size,
          modified: stats.mtimeMs,
        });
      }
      folders.sort((a, b) => a.name.localeCompare(b.name));
      files.sort((a, b) => b.modified - a.modified);
      return { folders, documents: files };
    },

    async read(name) {
      const target = locate(name);
      await ensureNoSymlinks(target.relative, { finalFile: true });
      const stat = await fs.lstat(target.full);
      if (!stat.isFile()) throw Object.assign(new Error('not a file'), { status: 404 });
      return fs.readFile(target.full);
    },

    async write(name, buffer) {
      await fs.mkdir(root, { recursive: true });
      const target = locate(name);
      await ensureNoSymlinks(target.relative, { createDirectories: true, includeFinal: false });
      const { full } = target;
      // Write beside the target then rename, so an interrupted save can never
      // leave a half-written document in place of a good one.
      const temp = `${full}.tmp-${randomUUID()}`;
      try {
        await fs.writeFile(temp, buffer, { flag: 'wx' });
        await fs.rename(temp, full);
      } finally {
        await fs.rm(temp, { force: true });
      }
    },

    async remove(name) {
      const target = locate(name);
      await ensureNoSymlinks(target.relative, { finalFile: true });
      const stat = await fs.lstat(target.full);
      if (!stat.isFile()) throw Object.assign(new Error('not a file'), { status: 404 });
      return fs.unlink(target.full);
    },

    async move(sourceName, destinationFolder) {
      const source = locate(sourceName);
      const folder = safeRelativePath(destinationFolder);
      const destination = [folder, path.basename(source.relative)].filter(Boolean).join('/');
      if (destination === source.relative) return { source: source.relative, path: destination };

      const target = path.resolve(root, ...(destination ? destination.split('/') : []));
      if (target !== root && !target.startsWith(root + path.sep)) throw reject('invalid path');
      await ensureNoSymlinks(source.relative, { finalFile: true });
      await ensureNoSymlinks(folder);

      const sourceStat = await fs.lstat(source.full);
      if (!sourceStat.isFile()) throw Object.assign(new Error('not a file'), { status: 404 });
      try {
        await fs.lstat(target);
        throw Object.assign(new Error('a file with that name already exists in the destination folder'), { status: 409 });
      } catch (err) {
        if (err.code !== 'ENOENT') throw err;
      }

      await fs.rename(source.full, target);
      return { source: source.relative, path: destination };
    },
  };
}
