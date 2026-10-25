// Suite shell: view switching, file browsing, and save/open/import.
import { createWordEditor } from './word.js?v=image-drop-move-resize';
import { createSheetEditor } from './sheets.js';
import { createImageEditor } from './images.js';
import { SAVE_FORMATS, extensionOf, kindOf, usedRange } from '/shared/model.js';
import { paperDimensions } from '/shared/paper.js';

const ACCEPT = {
  word: '.doc,.dot,.docx,.odt,.rtf,.txt,.md,.html,.htm',
  sheet: '.xls,.xlt,.xlsx,.ods,.csv',
  image: '.lipimg,.png,.jpg,.jpeg,.gif,.tif,.tiff,.bmp,.iff,.lbm,.psd,.webp,.avif',
};
const DEFAULT_FOLDER = { word: 'Docs', sheet: 'Spreadsheets', image: 'Images' };

const elements = {
  tabs: [...document.querySelectorAll('.viewtab')],
  openButton: document.querySelector('#open-button'),
  manageButton: document.querySelector('#manage-button'),
  saveAsButton: document.querySelector('#save-as-button'),
  openDialog: document.querySelector('#open-dialog'),
  closeOpenDialog: document.querySelector('#close-open-dialog'),
  breadcrumbs: document.querySelector('#file-breadcrumbs'),
  fileList: document.querySelector('#file-list'),
  dialogFooter: document.querySelector('#dialog-footer'),
  dialogFooterLabel: document.querySelector('#dialog-footer-label'),
  cancelDialogAction: document.querySelector('#cancel-dialog-action'),
  confirmDialogAction: document.querySelector('#confirm-dialog-action'),
  nameDisplay: document.querySelector('#doc-name-display'),
  name: document.querySelector('#doc-name'),
  saveNameField: document.querySelector('#save-name-field'),
  format: document.querySelector('#doc-format'),
  save: document.querySelector('#save-button'),
  download: document.querySelector('#download-button'),
  print: document.querySelector('#print-button'),
  newButton: document.querySelector('#new-button'),
  importButton: document.querySelector('#import-button'),
  importInput: document.querySelector('#import-input'),
  status: document.querySelector('#status'),
  theme: document.querySelector('#theme-button'),
};

const editors = {
  word: createWordEditor(document.querySelector('#editor-word')),
  sheet: createSheetEditor(document.querySelector('#editor-sheet')),
  image: createImageEditor(document.querySelector('#editor-image')),
};

let view = 'word';
let formatKind = '';
let openFolder = '';
let dialogMode = 'open';
let draggedPath = '';
let dragNavigateTimer = 0;
let dragTargetFolder = '';
let pendingSaveName = '';

const current = () => editors[view];

function applyTheme(theme) {
  const dark = theme === 'dark';
  document.documentElement.dataset.theme = dark ? 'dark' : 'light';
  // Show the icon for the mode the button will switch to.
  elements.theme.querySelector('.theme-icon-sun').toggleAttribute('hidden', !dark);
  elements.theme.querySelector('.theme-icon-moon').toggleAttribute('hidden', dark);
  elements.theme.title = `Switch to ${dark ? 'light' : 'dark'} mode`;
  elements.theme.setAttribute('aria-label', `Switch to ${dark ? 'light' : 'dark'} mode`);
  elements.theme.setAttribute('aria-pressed', String(dark));
}

const savedTheme = localStorage.getItem('lip-office-theme');
applyTheme(savedTheme || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'));
elements.theme.addEventListener('click', () => {
  const next = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
  localStorage.setItem('lip-office-theme', next);
  applyTheme(next);
});

async function request(url, options) {
  const response = await fetch(url, options);
  let payload = {};
  try {
    payload = await response.json();
  } catch {
    /* 204 and error pages have no JSON body */
  }
  if (!response.ok) throw new Error(payload.error || `${response.status} ${response.statusText}`);
  return payload;
}

let statusTimer = 0;

function setStatus(message, isError = false) {
  elements.status.textContent = message;
  elements.status.classList.toggle('error', isError);
  clearTimeout(statusTimer);
  if (!isError && message) {
    statusTimer = setTimeout(() => { elements.status.textContent = ''; }, 5000);
  }
}

function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

const parentFolder = (name) => name.includes('/') ? name.slice(0, name.lastIndexOf('/')) : '';

const base64 = (bytes) => {
  let binary = '';
  const chunk = 8192; // small enough for Safari's argument limit on fromCharCode
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
};

// ------------------------------------------------------------------ the shell

function updateSaveBar() {
  const editor = current();
  const displayName = editor.name ? editor.name.split('/').pop() : pendingSaveName || 'Untitled document';
  elements.nameDisplay.textContent = displayName;
  elements.nameDisplay.title = editor.name || pendingSaveName || 'Untitled document';
  elements.download.disabled = !editor.name;
  elements.importInput.accept = ACCEPT[editor.kind];

  if (formatKind !== editor.kind) {
    elements.format.replaceChildren(...SAVE_FORMATS[editor.kind].map((extension) => {
      const option = document.createElement('option');
      option.value = extension;
      option.textContent = `.${extension}`;
      return option;
    }));
    formatKind = editor.kind;
  }

  if (editor.name) {
    const extension = editor.name.split('.').pop().toLowerCase();
    if (SAVE_FORMATS[editor.kind].includes(extension)) elements.format.value = extension;
  }
  elements.format.disabled = Boolean(editor.name) && !(elements.openDialog.open && dialogMode === 'save');
}

function setView(next) {
  view = next;
  for (const tab of elements.tabs) tab.classList.toggle('active', tab.dataset.view === next);
  document.querySelector('#editor-word').classList.toggle('hidden', next !== 'word');
  document.querySelector('#editor-sheet').classList.toggle('hidden', next !== 'sheet');
  document.querySelector('#editor-image').classList.toggle('hidden', next !== 'image');
  updateSaveBar();
}

// -------------------------------------------------------- the open dialog

function fileRow(entry) {
  const item = document.createElement('li');

  const open = document.createElement('button');
  open.type = 'button';
  open.className = 'file-entry document-entry';
  open.title = `Open ${entry.name}. Drag onto a folder to move it.`;
  open.draggable = true;
  open.addEventListener('dragstart', (event) => {
    draggedPath = entry.path;
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', entry.path);
    open.classList.add('dragging');
  });
  open.addEventListener('dragend', () => {
    draggedPath = '';
    clearTimeout(dragNavigateTimer);
    dragNavigateTimer = 0;
    dragTargetFolder = '';
    open.classList.remove('dragging');
    document.querySelectorAll('.drag-target').forEach((target) => target.classList.remove('drag-target'));
  });
  open.addEventListener('click', () => {
    if (dialogMode === 'save') {
      elements.name.value = entry.name;
      openFolder = parentFolder(entry.path);
      const extension = entry.name.split('.').pop().toLowerCase();
      if (SAVE_FORMATS[current().kind].includes(extension)) elements.format.value = extension;
      refreshDialog();
    } else {
      openDocument(entry.path);
    }
  });

  const icon = document.createElement('span');
  icon.className = 'file-icon';
  icon.setAttribute('aria-hidden', 'true');
  icon.textContent = entry.kind === 'word' ? '▤' : entry.kind === 'sheet' ? '▦' : '▧';

  const label = document.createElement('span');
  label.className = 'file-label';
  label.textContent = entry.name;
  label.title = entry.name;

  const meta = document.createElement('span');
  meta.className = 'file-meta';
  meta.textContent = formatSize(entry.size);
  open.append(icon, label, meta);

  const remove = document.createElement('button');
  remove.type = 'button';
  remove.className = 'file-remove';
  remove.textContent = '\u00d7';
  remove.title = `Delete ${entry.path}`;
  remove.addEventListener('click', async (event) => {
    event.stopPropagation();
    if (!window.confirm(`Delete ${entry.path}? This cannot be undone.`)) return;
    try {
      await request(`/api/documents/${encodeURIComponent(entry.path)}`, { method: 'DELETE' });
      if (current().name === entry.path) {
        current().name = '';
        current().newDocument();
        updateSaveBar();
      }
      setStatus(`Deleted ${entry.path}`);
      await refreshDialog();
    } catch (err) {
      setStatus(err.message, true);
    }
  });

  item.append(open, remove);
  return item;
}

function markDropTarget(element, folder, { navigateOnHover = false } = {}) {
  element.addEventListener('dragover', (event) => {
    if (!draggedPath) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = 'move';
    element.classList.add('drag-target');
    if (navigateOnHover && activeFolder() !== folder && dragTargetFolder !== folder) {
      clearTimeout(dragNavigateTimer);
      dragTargetFolder = folder;
      dragNavigateTimer = setTimeout(() => {
        dragNavigateTimer = 0;
        dragTargetFolder = '';
        navigateFolder(folder);
      }, 1200);
    }
  });
  element.addEventListener('dragleave', (event) => {
    if (event.relatedTarget && element.contains(event.relatedTarget)) return;
    element.classList.remove('drag-target');
    if (dragTargetFolder === folder) {
      clearTimeout(dragNavigateTimer);
      dragNavigateTimer = 0;
      dragTargetFolder = '';
    }
  });
  element.addEventListener('drop', async (event) => {
    if (!draggedPath) return;
    event.preventDefault();
    event.stopPropagation();
    element.classList.remove('drag-target');
    clearTimeout(dragNavigateTimer);
    dragNavigateTimer = 0;
    dragTargetFolder = '';
    await moveDocument(event.dataTransfer.getData('text/plain') || draggedPath, folder);
    draggedPath = '';
  });
}

function folderRow(entry) {
  const item = document.createElement('li');
  const open = document.createElement('button');
  open.type = 'button';
  open.className = 'file-entry folder-entry';
  open.addEventListener('click', () => navigateFolder(entry.path));
  markDropTarget(open, entry.path, { navigateOnHover: true });

  const icon = document.createElement('span');
  icon.className = 'file-icon';
  icon.setAttribute('aria-hidden', 'true');
  icon.textContent = '▰';

  const label = document.createElement('span');
  label.className = 'file-label';
  label.textContent = entry.name;

  const arrow = document.createElement('span');
  arrow.className = 'folder-arrow';
  arrow.setAttribute('aria-hidden', 'true');
  arrow.textContent = '›';
  open.append(icon, label, arrow);
  item.append(open);
  return item;
}

function activeFolder() {
  return openFolder;
}

function setActiveFolder(folder) {
  openFolder = folder;
}

function updateDialogChrome() {
  const title = document.querySelector('#open-dialog-title');
  const subtitle = document.querySelector('.dialog-heading p');
  elements.saveNameField.classList.toggle('hidden', dialogMode !== 'save');
  if (dialogMode === 'save') {
    title.textContent = 'Save a file';
    subtitle.textContent = 'Choose a folder or create one. Drag files onto folders or breadcrumbs to move.';
    elements.dialogFooter.classList.remove('hidden');
    elements.dialogFooterLabel.textContent = `Save in ${openFolder || 'Documents'}`;
    elements.cancelDialogAction.textContent = 'Cancel';
    elements.cancelDialogAction.classList.remove('hidden');
    elements.confirmDialogAction.textContent = 'Save here';
  } else {
    title.textContent = dialogMode === 'manage' ? 'Manage files' : 'Open a file';
    subtitle.textContent = dialogMode === 'manage'
      ? 'Create folders and drag files onto folders or breadcrumbs to move them.'
      : 'Drag files onto folders or breadcrumbs to move; hover to browse.';
    elements.dialogFooter.classList.add('hidden');
  }
}

function renderBreadcrumbs() {
  elements.breadcrumbs.replaceChildren();
  const folder = activeFolder();
  const root = document.createElement('button');
  root.type = 'button';
  root.textContent = 'Documents';
  root.addEventListener('click', () => navigateFolder(''));
  markDropTarget(root, '', { navigateOnHover: true });
  elements.breadcrumbs.append(root);

  let path = '';
  for (const segment of folder.split('/').filter(Boolean)) {
    const separator = document.createElement('span');
    separator.setAttribute('aria-hidden', 'true');
    separator.textContent = '/';
    elements.breadcrumbs.append(separator);

    path = [path, segment].filter(Boolean).join('/');
    const target = path;
    const crumb = document.createElement('button');
    crumb.type = 'button';
    crumb.textContent = segment;
    crumb.addEventListener('click', () => navigateFolder(target));
    markDropTarget(crumb, target, { navigateOnHover: true });
    elements.breadcrumbs.append(crumb);
  }

  if (dialogMode === 'manage' || dialogMode === 'save') {
    const create = document.createElement('button');
    create.type = 'button';
    create.className = 'new-folder-button';
    create.textContent = '+ New folder';
    create.addEventListener('click', createFolderInDialog);
    elements.breadcrumbs.append(create);
  }
}

async function refreshDialog(folder = activeFolder()) {
  setActiveFolder(folder);
  updateDialogChrome();
  renderBreadcrumbs();
  elements.fileList.replaceChildren();
  try {
    const { folders, documents } = await request(`/api/documents?path=${encodeURIComponent(activeFolder())}`);
    const rows = [...folders.map(folderRow), ...documents.map(fileRow)];
    if (!rows.length) {
      const empty = document.createElement('li');
      empty.className = 'file-empty';
      empty.textContent = 'This folder has no supported documents.';
      elements.fileList.append(empty);
    } else {
      elements.fileList.replaceChildren(...rows);
    }
  } catch (err) {
    const error = document.createElement('li');
    error.className = 'file-empty error';
    error.textContent = err.message;
    elements.fileList.append(error);
  }
}

function navigateFolder(folder) {
  refreshDialog(folder);
}

async function createFolderInDialog() {
  const name = window.prompt('New folder name:');
  if (name === null || !name.trim()) return;
  try {
    const result = await request('/api/documents/folders', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ parent: activeFolder(), name: name.trim() }),
    });
    setStatus(`Created folder ${result.path}`);
    await refreshDialog(dialogMode === 'save' ? result.path : activeFolder());
  } catch (err) {
    setStatus(err.message, true);
    window.alert(err.message);
  }
}

async function showFileDialog(mode, folder = '') {
  dialogMode = mode;
  openFolder = folder;
  elements.openDialog.showModal();
  updateSaveBar();
  await refreshDialog(folder);
  if (mode === 'save') {
    elements.name.focus();
    elements.name.select();
  }
}

function showOpenDialog() {
  showFileDialog('open');
}

function showManageDialog() {
  showFileDialog('manage');
}

function showSaveDialog() {
  const editor = current();
  elements.name.value = editor.name?.split('/').pop() || pendingSaveName || `untitled.${elements.format.value}`;
  const folder = editor.name ? parentFolder(editor.name) : DEFAULT_FOLDER[editor.kind];
  showFileDialog('save', folder);
}

function cancelDialogAction() {
  elements.openDialog.close();
}

async function confirmDialogAction() {
  if (dialogMode === 'save') {
    if (await doSave(openFolder, true)) elements.openDialog.close();
  }
}

async function moveDocument(source, destination) {
  if (!source) return;
  if (parentFolder(source) === destination) {
    setStatus(`${source} is already in this folder.`);
    return;
  }
  try {
    const result = await request('/api/documents/move', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ source, destination }),
    });
    for (const editor of Object.values(editors)) {
      if (editor.name === result.source) editor.name = result.path;
    }
    updateSaveBar();
    setStatus(`Moved ${result.source} to ${result.path}`);
    await refreshDialog();
  } catch (err) {
    setStatus(err.message, true);
  }
}

elements.fileList.addEventListener('dragover', (event) => {
  if (!draggedPath) return;
  event.preventDefault();
  event.dataTransfer.dropEffect = 'move';
  elements.fileList.classList.add('drag-target');
});
elements.fileList.addEventListener('dragleave', (event) => {
  if (event.relatedTarget && elements.fileList.contains(event.relatedTarget)) return;
  elements.fileList.classList.remove('drag-target');
});
elements.fileList.addEventListener('drop', async (event) => {
  if (!draggedPath) return;
  event.preventDefault();
  elements.fileList.classList.remove('drag-target');
  await moveDocument(event.dataTransfer.getData('text/plain') || draggedPath, activeFolder());
  draggedPath = '';
});

// -------------------------------------------------------------- the commands

function fileNameFor(value, format, folder, forceFolder = false) {
  const base = String(value || '').trim() || 'untitled';
  let name = extensionOf(base) ? `${base.replace(/\.[^.]+$/, '')}.${format}` : `${base}.${format}`;
  if (forceFolder) name = name.split('/').pop();
  return name.includes('/') || !folder ? name : `${folder}/${name}`;
}

async function openDocument(name) {
  setStatus(`Opening ${name}\u2026`);
  try {
    const payload = await request(`/api/documents/${encodeURIComponent(name)}`);
    const editor = editors[payload.kind];
    if (!editor) throw new Error(`cannot open ${name}`);
    setView(payload.kind);
    await editor.setDocument(payload.doc);
    editor.name = name;
    updateSaveBar();
    if (elements.openDialog.open) elements.openDialog.close();
    editor.focus();
    setStatus(`Opened ${name}`);
  } catch (err) {
    setStatus(err.message, true);
  }
}

async function doSave(targetFolder, forceFolder = false) {
  const editor = current();
  let doc;
  try {
    doc = editor.contents(elements.format.value || SAVE_FORMATS[editor.kind][0]);
  } catch (err) {
    setStatus(err.message, true);
    return false;
  }
  if (editor.kind === 'word' && doc.sourceDocxToken && doc.sourceDocxUnchanged === false && !doc.sourceDocxLossAccepted) {
    const accepted = window.confirm('This imported Word file contains formatting the editor may not preserve after changes. Saving will rebuild it using the editor’s supported formatting. Continue?');
    if (!accepted) return false;
    editor.acceptSourceDocxLoss?.();
    doc = editor.contents();
  }
  const defaultFormat = SAVE_FORMATS[editor.kind][0];
  if (!forceFolder && !editor.name) {
    showSaveDialog();
    return false;
  }
  if (!forceFolder && !SAVE_FORMATS[editor.kind].includes(editor.name.split('.').pop().toLowerCase())) {
    // Imported legacy/open formats remain untouched; Save As writes the modern format.
    elements.name.value = editor.name.split('/').pop();
    showFileDialog('save', parentFolder(editor.name));
    elements.name.value = elements.name.value.replace(/\.[^.]+$/, `.${defaultFormat}`);
    return false;
  }
  const currentFormat = extensionOf(editor.name || '');
  const format = forceFolder
    ? (elements.format.value || defaultFormat)
    : (SAVE_FORMATS[editor.kind].includes(currentFormat) ? currentFormat : defaultFormat);
  const saveFolder = forceFolder ? targetFolder : parentFolder(editor.name);
  const name = forceFolder
    ? fileNameFor(elements.name.value, elements.format.value || format, saveFolder, true)
    : editor.name;
  setStatus('Saving\u2026');
  try {
    await request(`/api/documents/${encodeURIComponent(name)}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ kind: editor.kind, format, doc }),
    });
    editor.name = name;
    pendingSaveName = '';
    updateSaveBar();
    setStatus(`Saved ${name}`);
    return true;
  } catch (err) {
    setStatus(err.message, true);
    return false;
  }
}

function doDownload() {
  if (!current().name) return;
  window.location.href = `/api/documents/${encodeURIComponent(current().name)}?download=1`;
}

function doPrint() {
  const editor = current();
  const printStyle = document.createElement('style');
  printStyle.id = 'print-settings';

  if (editor.kind === 'word') {
    const doc = editor.contents();
    const orientation = doc.orientation === 'landscape' ? 'landscape' : 'portrait';
    const dimensions = paperDimensions(doc.pageSize);
    const [width, height] = orientation === 'landscape' ? [...dimensions].reverse() : dimensions;
    printStyle.textContent = `@page { size: ${width}in ${height}in; margin: 0; }`;
  } else if (editor.kind === 'sheet') {
    const doc = editor.contents();
    const sheet = doc.sheets[doc.active];
    const refs = { ...sheet.cells, ...(sheet.images || {}) };
    const bounds = usedRange(refs) || { rows: 1, cols: 1 };
    const firstUnusedColumn = bounds.cols + 2; // the first grid column is the row number
    printStyle.textContent = `
      @page { size: letter landscape; margin: 0.35in; }
      #editor-sheet .grid tbody tr:nth-child(n + ${bounds.rows + 1}) { display: none; }
      #editor-sheet .grid tr > :nth-child(n + ${firstUnusedColumn}) { display: none; }
    `;
  } else {
    printStyle.textContent = '#editor-image .image-stage-scroll { display:block; overflow:visible; padding:0; background:#fff; } #image-stage { width:auto !important; max-width:100% !important; height:auto !important; box-shadow:none !important; }';
  }

  document.head.append(printStyle);
  const cleanup = () => {
    printStyle.remove();
    window.removeEventListener('afterprint', cleanup);
  };
  window.addEventListener('afterprint', cleanup, { once: true });
  window.print();
}

function doNew() {
  const editor = current();
  editor.newDocument();
  editor.name = '';
  pendingSaveName = '';
  updateSaveBar();
  editor.focus();
  setStatus('New file. Give it a name, then press Save.');
}

async function doImport(file) {
  if (!file) return;
  setStatus(`Importing ${file.name}\u2026`);
  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const payload = await request('/api/parse', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: file.name, data: base64(bytes) }),
    });
    const editor = editors[payload.kind];
    if (!editor) throw new Error(`cannot import ${file.name}`);
    setView(payload.kind);
    await editor.setDocument(payload.doc);
    editor.name = '';
    pendingSaveName = file.name;
    const importedFormat = file.name.split('.').pop().toLowerCase();
    if (SAVE_FORMATS[payload.kind].includes(importedFormat)) elements.format.value = importedFormat;
    updateSaveBar();
    editor.focus();
    setStatus(`Imported ${file.name}. Press Save to choose a name and folder.`);
  } catch (err) {
    setStatus(err.message, true);
  }
}

// ------------------------------------------------------------------- wiring

for (const tab of elements.tabs) {
  tab.addEventListener('click', () => {
    setView(tab.dataset.view);
  });
}

elements.newButton.addEventListener('click', doNew);
elements.openButton.addEventListener('click', showOpenDialog);
elements.manageButton.addEventListener('click', showManageDialog);
elements.saveAsButton.addEventListener('click', showSaveDialog);
elements.closeOpenDialog.addEventListener('click', () => elements.openDialog.close());
elements.openDialog.addEventListener('close', updateSaveBar);
elements.cancelDialogAction.addEventListener('click', cancelDialogAction);
elements.confirmDialogAction.addEventListener('click', confirmDialogAction);
elements.save.addEventListener('click', doSave);
elements.download.addEventListener('click', doDownload);
elements.print.addEventListener('click', doPrint);
document.querySelector('#editor-image').addEventListener('images:newdocument', () => {
  editors.image.name = '';
  pendingSaveName = '';
  updateSaveBar();
  setStatus('New image. Add layers, then press Save.');
});

elements.importButton.addEventListener('click', () => elements.importInput.click());
elements.importInput.addEventListener('change', () => {
  const file = elements.importInput.files[0];
  elements.importInput.value = ''; // so importing the same file twice still fires
  doImport(file);
});

const documentDropZone = document.querySelector('main');
let fileDragDepth = 0;
const hasDroppedFiles = (event) => event.dataTransfer?.types?.includes('Files')
  && !event.dataTransfer.types.includes('application/x-lip-office-image');
const clearFileDropState = () => {
  fileDragDepth = 0;
  documentDropZone.classList.remove('file-drop-active');
};
// Child editors may stop a handled drop from bubbling to the page listener.
documentDropZone.addEventListener('drop', clearFileDropState, true);
documentDropZone.addEventListener('dragenter', (event) => {
  if (!hasDroppedFiles(event)) return;
  event.preventDefault();
  fileDragDepth++;
  documentDropZone.classList.add('file-drop-active');
});
documentDropZone.addEventListener('dragover', (event) => {
  if (!hasDroppedFiles(event)) return;
  event.preventDefault();
  event.dataTransfer.dropEffect = 'copy';
});
documentDropZone.addEventListener('dragleave', (event) => {
  if (!hasDroppedFiles(event) || event.relatedTarget && documentDropZone.contains(event.relatedTarget)) return;
  fileDragDepth = Math.max(0, fileDragDepth - 1);
  if (!fileDragDepth) documentDropZone.classList.remove('file-drop-active');
});
documentDropZone.addEventListener('drop', async (event) => {
  if (!hasDroppedFiles(event)) return;
  event.preventDefault();
  clearFileDropState();
  const files = [...(event.dataTransfer.files || [])];
  const documents = files.filter((file) => kindOf(file.name));
  if (documents.length > 1) {
    setStatus('Drop one document at a time.', true);
    return;
  }
  if (documents.length === 1) {
    await doImport(documents[0]);
    return;
  }
  if (files.length && !files.every((file) => file.type.startsWith('image/'))) {
    setStatus(`Unsupported document type: ${files[0].name}`, true);
  }
});

elements.name.addEventListener('keydown', (event) => {
  if (event.key === 'Enter') {
    event.preventDefault();
    if (elements.openDialog.open && dialogMode === 'save') confirmDialogAction();
    else doSave();
  }
});

document.addEventListener('keydown', (event) => {
  if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 's') {
    if (event.target === elements.name) return;
    event.preventDefault();
    if (elements.openDialog.open && dialogMode === 'save') confirmDialogAction();
    else doSave();
  }
});

editors.word.newDocument();
setView('word');
