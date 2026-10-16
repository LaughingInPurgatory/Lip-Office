// The word processor. The document model is HTML: the browser is the editor,
// and the server converts HTML to and from .docx/.odt/.rtf/.txt on save/open.
import { sanitizeHtml } from '/shared/sanitize.js';
import { blankWord, normalizeWord } from '/shared/model.js';
import { MAX_IMAGE_BYTES } from '/shared/images.js';

const TABLE_HTML = `<table><tbody>${'<tr><td><br></td><td><br></td><td><br></td></tr>'.repeat(3)}</tbody></table><p><br></p>`;

function escapeText(text) {
  return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
}

export function createWordEditor(root) {
  const page = root.querySelector('#word-page');
  const header = root.querySelector('#word-header');
  const footer = root.querySelector('#word-footer');
  const columnsSelect = root.querySelector('#word-columns');
  const paperSelect = root.querySelector('#word-paper');
  const orientationSelect = root.querySelector('#word-orientation');
  const marginSelect = root.querySelector('#word-margin');
  const editables = [header, footer];
  const toolbar = root.querySelector('#word-toolbar');
  const colour = root.querySelector('#word-colour');
  const formatSelect = toolbar.querySelector('[data-command="formatBlock"]');
  const imageInput = root.querySelector('#word-image-input');

  let docTitle = '';
  let columns = 1;
  let pageSize = 'letter';
  let orientation = 'portrait';
  let margin = 1;
  let activeRegion = page;
  let repaginatePending = false;
  let selectedImage = null;
  let draggedImage = null;

  // Clicking a toolbar control blurs the page, so keep the last caret range and
  // put it back before running the command.
  let savedRange = null;

  function reveal(button) {
    activeRegion.focus();
    if (!savedRange) return;
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(savedRange);
  }

  // Toolbar buttons keep the caret thanks to the mousedown handler above, so
  // they run against the live selection. Only controls that steal focus (the
  // colour picker, the style dropdown, the link prompt) need the saved range
  // put back — and `selectionchange` is async, so restoring it when it is not
  // needed would clobber the real selection.
  function runCommand(command, value = null, options = {}) {
    if (options.restore) reveal();
    document.execCommand(command, false, value);
    syncToolbar();
  }

  function updateImageControls() {
    const enabled = Boolean(selectedImage?.isConnected);
    toolbar.querySelectorAll('[data-action="image-smaller"], [data-action="image-larger"]')
      .forEach((button) => { button.disabled = !enabled; });
  }

  function selectImage(image) {
    selectedImage?.classList.remove('image-selected');
    selectedImage = image?.isConnected ? image : null;
    selectedImage?.classList.add('image-selected');
    updateImageControls();
  }

  function resizeSelectedImage(direction) {
    if (!selectedImage?.isConnected) return;
    const width = selectedImage.getBoundingClientRect().width || Number(selectedImage.width) || selectedImage.naturalWidth;
    const ratio = selectedImage.naturalWidth && selectedImage.naturalHeight
      ? selectedImage.naturalHeight / selectedImage.naturalWidth
      : (Number(selectedImage.height) || width) / width;
    const nextWidth = Math.max(24, Math.min(1600, Math.round(width * (direction > 0 ? 1.2 : 0.8))));
    selectedImage.setAttribute('width', String(nextWidth));
    selectedImage.setAttribute('height', String(Math.max(1, Math.round(nextWidth * ratio))));
    scheduleRepaginate();
  }

  function insertImage(dataUrl, alt = '') {
    const probe = new Image();
    probe.addEventListener('load', () => {
      const scale = Math.min(1, 640 / probe.naturalWidth, 480 / probe.naturalHeight);
      const width = Math.max(1, Math.round(probe.naturalWidth * scale));
      const height = Math.max(1, Math.round(probe.naturalHeight * scale));
      reveal();
      const imagesBefore = new Set(page.querySelectorAll('img'));
      const safeAlt = String(alt).replace(/[&"<>]/g, (ch) => ({ '&': '&amp;', '"': '&quot;', '<': '&lt;', '>': '&gt;' })[ch]);
      document.execCommand('insertHTML', false, `<img src="${dataUrl}" alt="${safeAlt}" width="${width}" height="${height}" draggable="true">`);
      const inserted = [...page.querySelectorAll('img')].find((image) => !imagesBefore.has(image));
      if (inserted) selectImage(inserted);
      savedRange = selectionInPage();
      scheduleRepaginate();
    }, { once: true });
    probe.src = dataUrl;
  }

  function readImage(file) {
    if (!file || !['image/png', 'image/jpeg', 'image/gif'].includes(file.type) || file.size > MAX_IMAGE_BYTES) {
      window.alert('Choose a PNG, JPEG, or GIF image smaller than 8 MB.');
      return;
    }
    const reader = new FileReader();
    reader.addEventListener('load', () => {
      if (typeof reader.result === 'string') insertImage(reader.result, file.name || '');
    }, { once: true });
    reader.readAsDataURL(file);
  }

  function caretRangeAtPoint(x, y, targetPage) {
    let range = document.caretRangeFromPoint?.(x, y) || null;
    if (!range && document.caretPositionFromPoint) {
      const position = document.caretPositionFromPoint(x, y);
      if (position) {
        range = document.createRange();
        range.setStart(position.offsetNode, position.offset);
        range.collapse(true);
      }
    }
    if (range && targetPage.contains(range.commonAncestorContainer)) return range;
    range = document.createRange();
    range.selectNodeContents(targetPage);
    range.collapse(false);
    return range;
  }

  function focusAtDrop(event, targetPage) {
    const range = caretRangeAtPoint(event.clientX, event.clientY, targetPage);
    targetPage.focus({ preventScroll: true });
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
    activeRegion = targetPage;
    savedRange = range.cloneRange();
    return range;
  }

  function selectionInPage() {
    const selection = window.getSelection();
    if (!selection.rangeCount) return null;
    const range = selection.getRangeAt(0);
    const region = [...page.querySelectorAll('.page'), ...editables].find((item) => item.contains(range.commonAncestorContainer));
    if (region) activeRegion = region;
    return region ? range.cloneRange() : null;
  }

  function createPage(html = '') {
    const sheet = document.createElement('div');
    sheet.className = 'page';
    sheet.contentEditable = 'true';
    sheet.spellcheck = true;
    sheet.innerHTML = html;
    sheet.addEventListener('focus', () => { activeRegion = sheet; });
    return sheet;
  }

  function pages() { return [...page.querySelectorAll(':scope > .page')]; }

  function capturePageSelection() {
    const selection = window.getSelection();
    if (!selection?.rangeCount) return null;
    const range = selection.getRangeAt(0);
    if (!pages().some((sheet) => sheet.contains(range.commonAncestorContainer))) return null;
    return {
      anchorNode: selection.anchorNode,
      anchorOffset: selection.anchorOffset,
      focusNode: selection.focusNode,
      focusOffset: selection.focusOffset,
    };
  }

  function restorePageSelection(saved) {
    if (!saved || !saved.anchorNode?.isConnected || !saved.focusNode?.isConnected) return false;
    const sheet = pages().find((item) => item.contains(saved.anchorNode) && item.contains(saved.focusNode));
    if (!sheet) return false;
    const clamp = (node, offset) => Math.min(offset, node.nodeType === Node.TEXT_NODE ? node.length : node.childNodes.length);
    sheet.focus({ preventScroll: true });
    const selection = window.getSelection();
    selection.setBaseAndExtent(
      saved.anchorNode, clamp(saved.anchorNode, saved.anchorOffset),
      saved.focusNode, clamp(saved.focusNode, saved.focusOffset),
    );
    activeRegion = sheet;
    return true;
  }

  function overflowsPage(sheet) {
    const last = sheet.lastElementChild;
    if (!last) return false;
    const sheetStyle = getComputedStyle(sheet);
    const limit = sheet.getBoundingClientRect().bottom
      - parseFloat(sheetStyle.paddingBottom)
      - parseFloat(sheetStyle.borderBottomWidth);
    return last.getBoundingClientRect().bottom > limit + 1;
  }

  function logicalText(node) {
    if (node.nodeType === Node.TEXT_NODE) return node.nodeValue || '';
    if (node.nodeType === Node.ELEMENT_NODE && node.tagName === 'BR') return '\n';
    return [...node.childNodes].map(logicalText).join('');
  }

  function textPointAt(rootNode, offset) {
    let remaining = offset;
    let last = null;
    let point = null;
    function visit(parent) {
      for (let index = 0; index < parent.childNodes.length && !point; index++) {
        const node = parent.childNodes[index];
        if (node.nodeType === Node.TEXT_NODE) {
          last = node;
          if (remaining <= node.length) { point = { node, offset: remaining }; return; }
          remaining -= node.length;
        } else if (node.nodeType === Node.ELEMENT_NODE && node.tagName === 'BR') {
          if (remaining === 0) { point = { node: parent, offset: index }; return; }
          remaining--;
          if (remaining === 0) { point = { node: parent, offset: index + 1 }; return; }
        } else {
          visit(node);
        }
      }
    }
    visit(rootNode);
    if (point) return point;
    return last ? { node: last, offset: last.length } : { node: rootNode, offset: rootNode.childNodes.length };
  }

  function textOffsetOf(rootNode, target, targetOffset) {
    const range = document.createRange();
    range.selectNodeContents(rootNode);
    range.setEnd(target, targetOffset);
    return logicalText(range.cloneContents()).length;
  }

  function splitOverflowingBlock(sheet) {
    const block = sheet.firstElementChild;
    if (!block || sheet.childElementCount !== 1 || !['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'BLOCKQUOTE'].includes(block.tagName)) return false;

    const text = logicalText(block);
    if (text.length < 2) return false;
    const style = getComputedStyle(sheet);
    const bottom = sheet.getBoundingClientRect().bottom - parseFloat(style.paddingBottom) - parseFloat(style.borderBottomWidth);
    const range = document.createRange();
    let low = 0;
    let high = text.length;
    while (low < high) {
      const middle = Math.ceil((low + high) / 2);
      const point = textPointAt(block, middle);
      range.selectNodeContents(block);
      range.setEnd(point.node, point.offset);
      const rect = range.getBoundingClientRect();
      if (rect.height && rect.bottom <= bottom) low = middle;
      else high = middle - 1;
    }

    let splitAt = low;
    const boundary = text.slice(0, low).search(/\s+[^\s]*$/);
    if (boundary >= 0) splitAt = boundary + text.slice(boundary).match(/^\s+/)[0].length;
    if (splitAt <= 0 || splitAt >= text.length || !text.slice(splitAt).trim()) return false;

    const selection = window.getSelection();
    const caretInBlock = selection?.isCollapsed && block.contains(selection.anchorNode);
    const oldCaretOffset = caretInBlock ? textOffsetOf(block, selection.anchorNode, selection.anchorOffset) : -1;
    const point = textPointAt(block, splitAt);
    const tailRange = document.createRange();
    tailRange.setStart(point.node, point.offset);
    tailRange.setEnd(block, block.childNodes.length);
    const tail = block.cloneNode(false);
    tail.append(tailRange.extractContents());

    const nextPage = pages()[pages().indexOf(sheet) + 1] || page.appendChild(createPage());
    nextPage.insertBefore(tail, nextPage.firstChild);

    if (caretInBlock) {
      const caretBlock = oldCaretOffset >= splitAt ? tail : block;
      const caretOffset = oldCaretOffset >= splitAt ? oldCaretOffset - splitAt : oldCaretOffset;
      const caret = textPointAt(caretBlock, caretOffset);
      const caretRange = document.createRange();
      caretRange.setStart(caret.node, caret.offset);
      caretRange.collapse(true);
      selection.removeAllRanges();
      selection.addRange(caretRange);
      activeRegion = caretBlock.closest('.page');
      savedRange = caretRange.cloneRange();
    }
    return true;
  }

  function scheduleRepaginate() {
    if (repaginatePending) return;
    repaginatePending = true;
    requestAnimationFrame(() => {
      repaginatePending = false;
      repaginate();
    });
  }

  function repaginate() {
    // Reflow moves blocks between independent contenteditable pages. Browsers
    // can otherwise keep keyboard focus on the old page while the caret's node
    // has moved, so preserve and explicitly restore the caret after the moves.
    const savedSelection = capturePageSelection();
    const paper = pageSize === 'a4' ? [8.27, 11.69] : [8.5, 11];
    const height = (orientation === 'landscape' ? paper[0] : paper[1]) * 96;
    let sheets = pages();
    if (!sheets.length) page.append(createPage('<p><br></p>'));

    // Keep blocks whole at page edges. If a single block is taller than a page,
    // let that sheet grow instead of painting a gap over its text.
    for (let index = 0; index < pages().length; index++) {
      const sheet = pages()[index];
      sheet.removeAttribute('data-overflowing');
      sheet.style.height = `${height / 96}in`;
      if (!overflowsPage(sheet)) continue;
      if (sheet.childNodes.length > 1) {
        const next = pages()[index + 1] || page.appendChild(createPage());
        const moving = sheet.lastChild;
        next.insertBefore(moving, next.firstChild);
        index--;
      } else if (splitOverflowingBlock(sheet)) {
        index--;
      } else {
        sheet.dataset.overflowing = 'true';
        sheet.style.height = 'auto';
      }
    }

    // Reclaim empty space after edits by pulling whole blocks back when they fit.
    for (let index = 0; index < pages().length - 1; index++) {
      const current = pages()[index];
      const next = pages()[index + 1];
      if (current.dataset.overflowing === 'true') continue;
      while (next.firstChild) {
        const block = next.firstChild;
        current.append(block);
        if (overflowsPage(current)) {
          next.insertBefore(block, next.firstChild);
          break;
        }
      }
    }

    sheets = pages();
    while (sheets.length > 1 && !sheets.at(-1).hasChildNodes()) {
      sheets.pop().remove();
    }
    if (restorePageSelection(savedSelection)) {
      savedRange = window.getSelection().getRangeAt(0).cloneRange();
    } else if (!activeRegion?.isConnected || activeRegion === page) {
      activeRegion = sheets[0];
    }
    applyPageLayout(false);
  }

  function syncToolbar() {
    for (const button of toolbar.querySelectorAll('button[data-command]')) {
      let pressed = false;
      try {
        pressed = document.queryCommandState(button.dataset.command);
      } catch {
        pressed = false;
      }
      button.setAttribute('aria-pressed', pressed ? 'true' : 'false');
    }

    try {
      const block = String(document.queryCommandValue('formatBlock')).toUpperCase();
      const match = [...formatSelect.options].find((option) => option.value === block);
      if (match) formatSelect.value = match.value;
    } catch {
      /* queryCommandValue is best-effort */
    }
  }

  function applyPageLayout(schedule = true) {
    const paper = pageSize === 'a4' ? [8.27, 11.69] : [8.5, 11];
    const [width, height] = orientation === 'landscape' ? [paper[1], paper[0]] : paper;
    page.style.setProperty('--paper-height', `${height}in`);
    for (const sheet of pages()) {
      sheet.style.width = `${width}in`;
      sheet.style.height = sheet.dataset.overflowing === 'true' ? 'auto' : `${height}in`;
      sheet.style.minHeight = `${height}in`;
      sheet.style.padding = `${margin}in`;
      sheet.dataset.columns = String(columns);
    }
    for (const region of [header, footer]) {
      region.style.width = `${width}in`;
      region.style.padding = `8px ${margin}in`;
    }
    if (schedule) scheduleRepaginate();
  }

  document.addEventListener('selectionchange', () => {
    const range = selectionInPage();
    if (range) {
      savedRange = range;
      syncToolbar();
    }
  });

  page.addEventListener('focusin', (event) => {
    const sheet = event.target.closest('.page');
    if (sheet) activeRegion = sheet;
  });
  for (const editable of editables) editable.addEventListener('focus', () => { activeRegion = editable; });
  page.addEventListener('input', () => {
    if (!selectedImage?.isConnected) selectImage(null);
    scheduleRepaginate();
  });
  page.addEventListener('load', (event) => { if (event.target.matches('img')) scheduleRepaginate(); }, true);
  page.addEventListener('click', (event) => {
    const image = event.target.closest?.('img');
    selectImage(image && page.contains(image) ? image : null);
  });
  page.addEventListener('dragstart', (event) => {
    const image = event.target.closest?.('img');
    if (!image || !page.contains(image)) return;
    draggedImage = image;
    image.draggable = true;
    if (event.dataTransfer) {
      event.dataTransfer.setData('application/x-lip-office-image', 'move');
      event.dataTransfer.effectAllowed = 'move';
    }
  });
  page.addEventListener('dragend', () => { draggedImage = null; });
  page.addEventListener('dragover', (event) => {
    const imageFile = [...(event.dataTransfer?.items || [])]
      .some((item) => item.kind === 'file' && /^image\/(png|jpeg|gif)$/.test(item.type));
    if (draggedImage || imageFile) {
      event.preventDefault();
      if (event.dataTransfer) event.dataTransfer.dropEffect = draggedImage ? 'move' : 'copy';
    }
  });
  page.addEventListener('drop', (event) => {
    const targetPage = event.target.closest?.('.page');
    if (!targetPage || !page.contains(targetPage)) return;
    const isImageMove = draggedImage || event.dataTransfer?.types?.includes('application/x-lip-office-image');
    if (isImageMove && draggedImage) {
      event.preventDefault();
      const range = focusAtDrop(event, targetPage);
      const moving = draggedImage;
      range.insertNode(moving);
      range.setStartAfter(moving);
      range.collapse(true);
      window.getSelection().removeAllRanges();
      window.getSelection().addRange(range);
      selectImage(moving);
      savedRange = range.cloneRange();
      scheduleRepaginate();
      return;
    }
    const file = [...(event.dataTransfer?.files || [])]
      .find((item) => ['image/png', 'image/jpeg', 'image/gif'].includes(item.type));
    if (!file) return;
    event.preventDefault();
    focusAtDrop(event, targetPage);
    readImage(file);
  });
  window.addEventListener('resize', scheduleRepaginate);

  toolbar.addEventListener('mousedown', (event) => {
    // Keep the caret in the page when a toolbar button is pressed.
    if (event.target.closest('button')) event.preventDefault();
  });

  toolbar.addEventListener('click', (event) => {
    const button = event.target.closest('button');
    if (!button) return;
    const { command, action } = button.dataset;

    if (action === 'link') {
      const url = window.prompt('Link address:', 'https://');
      if (url) runCommand('createLink', url, { restore: true });
      return;
    }
    if (action === 'unlink') return runCommand('unlink');
    if (action === 'table') return runCommand('insertHTML', TABLE_HTML);
    if (action === 'image') {
      imageInput.click();
      return;
    }
    if (action === 'image-smaller') return resizeSelectedImage(-1);
    if (action === 'image-larger') return resizeSelectedImage(1);
    if (command) runCommand(command);
    return undefined;
  });

  toolbar.addEventListener('change', (event) => {
    if (event.target === formatSelect) runCommand('formatBlock', formatSelect.value, { restore: true });
    if (event.target === columnsSelect) {
      columns = Number(columnsSelect.value);
      applyPageLayout();
    }
    if (event.target === paperSelect) { pageSize = paperSelect.value; applyPageLayout(); }
    if (event.target === orientationSelect) { orientation = orientationSelect.value; applyPageLayout(); }
    if (event.target === marginSelect) { margin = Number(marginSelect.value); applyPageLayout(); }
  });

  // `change` rather than `input`: the colour dialog holds focus while it is open.
  colour.addEventListener('change', () => runCommand('foreColor', colour.value, { restore: true }));

  imageInput.addEventListener('change', () => {
    readImage(imageInput.files?.[0]);
    imageInput.value = '';
  });

  for (const editable of [page, ...editables]) editable.addEventListener('paste', (event) => {
    const clipboard = event.clipboardData;
    if (!clipboard) return;
    const html = clipboard.getData('text/html');
    const text = clipboard.getData('text/plain');
    const image = [...(clipboard.items || [])].find((item) => item.kind === 'file' && /^image\/(png|jpeg|gif)$/.test(item.type));
    if (!html && !text && !image) return;
    event.preventDefault();
    if (!html && image) {
      readImage(image.getAsFile());
      return;
    }
    // Same whitelist the server applies on save, so pasted junk never enters the document.
    const clean = html ? sanitizeHtml(html) : '';
    document.execCommand('insertHTML', false, clean || escapeText(text));
  });

  page.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') scheduleRepaginate();
    if (event.key !== 'Enter' || !event.shiftKey) return;
    // Shift+Enter inside a heading should continue the body text, not the heading.
    if (/^(H[1-6]|BLOCKQUOTE|PRE)$/.test(String(document.queryCommandValue('formatBlock')).toUpperCase())) {
      event.preventDefault();
      document.execCommand('insertHTML', false, '<p><br></p>');
    }
  });

  return {
    kind: 'word',
    name: '',

    newDocument() {
      docTitle = '';
      savedRange = null;
      selectImage(null);
      activeRegion = null;
      const blank = blankWord();
      page.replaceChildren(createPage(sanitizeHtml(blank.html)));
      activeRegion = pages()[0];
      header.innerHTML = blank.header;
      footer.innerHTML = blank.footer;
      columns = blank.columns;
      pageSize = blank.pageSize;
      orientation = blank.orientation;
      margin = blank.margin;
      columnsSelect.value = String(columns);
      paperSelect.value = pageSize;
      orientationSelect.value = orientation;
      marginSelect.value = String(margin);
      applyPageLayout();
    },

    setDocument(doc) {
      const safe = normalizeWord(doc);
      savedRange = null;
      selectImage(null);
      activeRegion = null;
      docTitle = safe.title;
      page.replaceChildren(createPage(sanitizeHtml(safe.html) || '<p><br></p>'));
      activeRegion = pages()[0];
      header.innerHTML = sanitizeHtml(safe.header);
      footer.innerHTML = sanitizeHtml(safe.footer);
      columns = safe.columns;
      pageSize = safe.pageSize;
      orientation = safe.orientation;
      margin = safe.margin;
      columnsSelect.value = String(columns);
      paperSelect.value = pageSize;
      orientationSelect.value = orientation;
      marginSelect.value = String(margin);
      applyPageLayout();
    },

    contents() {
      return { title: docTitle, html: pages().map((sheet) => sheet.innerHTML).join(''), header: header.innerHTML, footer: footer.innerHTML, columns, pageSize, orientation, margin };
    },

    focus() { (activeRegion?.matches?.('.page') ? activeRegion : pages()[0])?.focus(); },
  };
}
