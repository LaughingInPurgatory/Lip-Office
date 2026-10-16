// The spreadsheet. A plain HTML table is the grid; formulas are evaluated by
// the shared engine in /shared/formula.js, the same one the server uses when
// writing .xlsx/.ods.
import { makeEngine, numToCol, toRef, parseRef } from '/shared/formula.js';
import { blankSheet, normalizeSheet, usedRange } from '/shared/model.js';
import { MAX_IMAGE_BYTES } from '/shared/images.js';

const START_ROWS = 100;
const START_COLS = 26;
const MAX_ROWS = 5000;
const MAX_COLS = 256;

const clamp = (value, low, high) => Math.min(Math.max(value, low), high);

export function createSheetEditor(root) {
  const host = root.querySelector('#sheet-grid');
  const formulaInput = root.querySelector('#sheet-formula');
  const refBox = root.querySelector('#sheet-ref');
  const tabsHost = root.querySelector('#sheet-tabs');
  const toolbar = root.querySelector('#sheet-toolbar');
  const imageInput = root.querySelector('#sheet-image-input');
  const fontFamilySelect = root.querySelector('#sheet-font-family');
  const fontSizeSelect = root.querySelector('#sheet-font-size');

  let doc = blankSheet();
  let rows = START_ROWS;
  let cols = START_COLS;
  let cursor = { col: 1, row: 1 };
  let nodes = new Map();   // A1 -> <td>
  let editing = null;      // { ref, settle }
  let selectedTds = new Set();
  let selectionAnchor = { col: 1, row: 1 };
  let selectingCells = false;
  let formulaRef = 'A1';

  const sheet = () => doc.sheets[doc.active];

  function header(className, text) {
    const th = document.createElement('th');
    if (className) th.className = className;
    th.textContent = text;
    return th;
  }

  function columnWidth(col) {
    return Number(sheet().columnWidths?.[numToCol(col)] || 96);
  }

  function gridWidth() {
    let width = 46;
    for (let col = 1; col <= cols; col++) width += columnWidth(col);
    return width;
  }

  function build() {
    editing = null;
    selectedTds.clear();
    nodes = new Map();

    const table = document.createElement('table');
    const colgroup = document.createElement('colgroup');
    const rowHeaderCol = document.createElement('col');
    rowHeaderCol.style.width = '46px';
    colgroup.append(rowHeaderCol);
    for (let col = 1; col <= cols; col++) {
      const column = document.createElement('col');
      column.style.width = `${columnWidth(col)}px`;
      colgroup.append(column);
    }
    table.append(colgroup);
    table.style.width = `${gridWidth()}px`;
    const headRow = document.createElement('tr');
    headRow.append(header('corner', ''));
    for (let col = 1; col <= cols; col++) {
      const th = header('col-header', numToCol(col));
      th.dataset.resizeCol = String(col);
      headRow.append(th);
    }
    const thead = document.createElement('thead');
    thead.append(headRow);
    table.append(thead);

    const tbody = document.createElement('tbody');
    for (let row = 1; row <= rows; row++) {
      const tr = document.createElement('tr');
      tr.style.height = `${sheet().rowHeights?.[row] || 24}px`;
      const rowHeading = header('rownum', String(row));
      rowHeading.dataset.resizeRow = String(row);
      tr.append(rowHeading);
      for (let col = 1; col <= cols; col++) {
        const ref = toRef(col, row);
        const td = document.createElement('td');
        td.dataset.ref = ref;
        nodes.set(ref, td);
        tr.append(td);
      }
      tbody.append(tr);
    }
    table.append(tbody);

    host.replaceChildren(table);
    refresh();
    markSelected();
    syncFormula();
  }

  // shortcut: every commit recomputes and repaints the whole visible grid.
  // Fine to a few thousand cells; switch to dirty-range recalculation if the
  // default grid ever gets much bigger.
  function refresh() {
    const cells = sheet().cells;
    const engine = makeEngine(cells);
    for (const [ref, td] of nodes) {
      const raw = cells[ref];
      const image = sheet().images?.[ref];
      const shown = raw === undefined ? '' : engine.display(ref);
      td.replaceChildren();
      if (image) {
        const img = document.createElement('img');
        img.className = 'cell-image';
        img.src = image.src;
        img.alt = image.alt || '';
        img.width = image.width;
        img.height = image.height;
        td.append(img);
      }
      if (shown) {
        const value = document.createElement('span');
        value.textContent = shown;
        td.append(value);
      }
      const style = sheet().cellStyles?.[ref];
      td.style.fontFamily = style?.fontFamily || '';
      td.style.fontSize = style?.fontSize ? `${style.fontSize}pt` : '';
      if (raw === undefined && !image) { td.className = ''; continue; }
      const numeric = shown !== '' && Number.isFinite(Number(shown));
      td.className = [
        numeric ? 'numeric' : '',
        shown.startsWith('#') ? 'error' : '',
        raw?.[0] === '=' ? 'formula' : '',
      ].filter(Boolean).join(' ');
    }
    markSelected();
  }

  function markSelected() {
    for (const td of selectedTds) td.classList.remove('selected', 'active-cell');
    selectedTds.clear();
    const firstCol = Math.min(selectionAnchor.col, cursor.col);
    const lastCol = Math.max(selectionAnchor.col, cursor.col);
    const firstRow = Math.min(selectionAnchor.row, cursor.row);
    const lastRow = Math.max(selectionAnchor.row, cursor.row);
    for (let row = firstRow; row <= lastRow; row++) {
      for (let col = firstCol; col <= lastCol; col++) {
        const td = nodes.get(toRef(col, row));
        if (!td) continue;
        td.classList.add('selected');
        if (col === cursor.col && row === cursor.row) td.classList.add('active-cell');
        selectedTds.add(td);
      }
    }
  }

  function forSelectedCells(callback) {
    const firstCol = Math.min(selectionAnchor.col, cursor.col);
    const lastCol = Math.max(selectionAnchor.col, cursor.col);
    const firstRow = Math.min(selectionAnchor.row, cursor.row);
    const lastRow = Math.max(selectionAnchor.row, cursor.row);
    for (let row = firstRow; row <= lastRow; row++) {
      for (let col = firstCol; col <= lastCol; col++) callback(toRef(col, row));
    }
  }

  function syncFormula() {
    const ref = toRef(cursor.col, cursor.row);
    formulaRef = ref;
    refBox.textContent = ref;
    formulaInput.value = sheet().cells[ref] ?? '';
    const style = sheet().cellStyles?.[ref] || {};
    fontFamilySelect.value = style.fontFamily || '';
    fontSizeSelect.value = style.fontSize ? String(style.fontSize) : '';
  }

  function moveCursor(col, row, options = {}) {
    commitEdit();
    cursor = { col: clamp(col, 1, cols), row: clamp(row, 1, rows) };
    if (!options.extend) selectionAnchor = { ...cursor };
    markSelected();
    syncFormula();
    nodes.get(toRef(cursor.col, cursor.row))?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    if (options.edit) beginEdit(options.initial ?? null);
    else host.focus();
  }

  function commitEdit() {
    if (editing) editing.settle(true);
  }

  function beginEdit(initial) {
    const ref = toRef(cursor.col, cursor.row);
    const td = nodes.get(ref);
    if (!td) return;

    const input = document.createElement('input');
    input.className = 'cell-input';
    input.value = initial === null ? (sheet().cells[ref] ?? '') : initial;

    let settled = false;
    const settle = (commit) => {
      if (settled) return;
      settled = true;
      editing = null;
      if (commit) {
        const value = input.value;
        if (value === '') delete sheet().cells[ref];
        else sheet().cells[ref] = value;
      }
      refresh();
      syncFormula();
    };

    editing = { ref, settle };
    td.replaceChildren(input);
    input.focus();
    input.setSelectionRange(input.value.length, input.value.length);

    input.addEventListener('keydown', (event) => {
      event.stopPropagation();
      if (event.key === 'Enter') {
        event.preventDefault();
        settle(true);
        moveCursor(cursor.col, cursor.row + 1);
      } else if (event.key === 'Tab') {
        event.preventDefault();
        settle(true);
        moveCursor(cursor.col + (event.shiftKey ? -1 : 1), cursor.row);
      } else if (event.key === 'Escape') {
        event.preventDefault();
        settle(false);
      }
    });
    input.addEventListener('blur', () => settle(true));
  }

  function commitFormula() {
    const value = formulaInput.value;
    const prior = sheet().cells[formulaRef] ?? '';
    if (value === prior) return;
    if (value === '') delete sheet().cells[formulaRef];
    else sheet().cells[formulaRef] = value;
    refresh();
  }

  function switchSheet(index) {
    commitEdit();
    doc.active = index;
    cursor = { col: 1, row: 1 };
    selectionAnchor = { ...cursor };
    renderTabs();
    build();
  }

  function renderTabs() {
    tabsHost.replaceChildren();
    doc.sheets.forEach((item, index) => {
      const tab = document.createElement('button');
      tab.type = 'button';
      tab.className = `sheet-tab${index === doc.active ? ' active' : ''}`;
      tab.textContent = item.name;
      tab.addEventListener('click', () => { if (index !== doc.active) switchSheet(index); });
      tab.addEventListener('dblclick', () => {
        const next = window.prompt('Sheet name:', item.name);
        if (next && next.trim()) {
          item.name = next.trim().slice(0, 31);
          renderTabs();
        }
      });
      tabsHost.append(tab);
    });

    const add = document.createElement('button');
    add.type = 'button';
    add.className = 'sheet-tab add';
    add.textContent = '+ Sheet';
    add.addEventListener('click', () => {
      doc.sheets.push({ name: `Sheet${doc.sheets.length + 1}`, cells: {}, images: {}, columnWidths: {}, rowHeights: {}, cellStyles: {} });
      switchSheet(doc.sheets.length - 1);
    });
    tabsHost.append(add);
  }

  host.addEventListener('mousedown', (event) => {
    const td = event.target.closest('td[data-ref]');
    if (!td) return;
    if (editing && editing.ref === td.dataset.ref) return; // click landed inside the open editor
    if (event.button !== 0) return;
    const position = parseRef(td.dataset.ref);
    if (!position) return;
    event.preventDefault();
    commitEdit();
    if (!event.shiftKey) selectionAnchor = position;
    cursor = position;
    selectingCells = true;
    markSelected();
    syncFormula();
    host.focus();
  });

  host.addEventListener('mousemove', (event) => {
    if (!selectingCells) return;
    const td = event.target.closest('td[data-ref]');
    const position = td && parseRef(td.dataset.ref);
    if (!position || position.col === cursor.col && position.row === cursor.row) return;
    cursor = position;
    markSelected();
    syncFormula();
  });
  document.addEventListener('mouseup', () => { selectingCells = false; });

  host.addEventListener('dblclick', (event) => {
    const td = event.target.closest('td[data-ref]');
    if (!td) return;
    if (editing && editing.ref === td.dataset.ref) return;
    const position = parseRef(td.dataset.ref);
    if (!position) return;
    moveCursor(position.col, position.row);
    beginEdit(null);
  });

  host.addEventListener('keydown', (event) => {
    if (editing) return;
    const move = (dCol, dRow, extend = false) => {
      event.preventDefault();
      moveCursor(cursor.col + dCol, cursor.row + dRow, { extend });
    };

    switch (event.key) {
      case 'ArrowUp': return move(0, -1, event.shiftKey);
      case 'ArrowDown': return move(0, 1, event.shiftKey);
      case 'ArrowLeft': return move(-1, 0, event.shiftKey);
      case 'ArrowRight': return move(1, 0, event.shiftKey);
      case 'PageUp': return move(0, -20, event.shiftKey);
      case 'PageDown': return move(0, 20, event.shiftKey);
      case 'Home': return move(-cols, 0, event.shiftKey);
      case 'End': return move(cols, 0, event.shiftKey);
      case 'Tab': return move(event.shiftKey ? -1 : 1, 0);
      case 'Enter': return move(0, event.shiftKey ? -1 : 1);
      case 'F2':
        event.preventDefault();
        beginEdit(null);
        return;
      case 'Delete':
      case 'Backspace':
        event.preventDefault();
        forSelectedCells((ref) => { delete sheet().cells[ref]; });
        refresh();
        syncFormula();
        return;
      default: break;
    }

    if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
      event.preventDefault();
      beginEdit(event.key);
    }
  });

  formulaInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      commitFormula();
      moveCursor(cursor.col, cursor.row + 1);
    } else if (event.key === 'Tab') {
      event.preventDefault();
      commitFormula();
      moveCursor(cursor.col + (event.shiftKey ? -1 : 1), cursor.row);
    } else if (event.key === 'Escape') {
      event.preventDefault();
      syncFormula();
      host.focus();
    }
  });
  formulaInput.addEventListener('blur', commitFormula);

  function applySelectedFont() {
    sheet().cellStyles ||= {};
    forSelectedCells((ref) => {
      const style = { ...(sheet().cellStyles[ref] || {}) };
      if (fontFamilySelect.value) style.fontFamily = fontFamilySelect.value;
      else delete style.fontFamily;
      if (fontSizeSelect.value) style.fontSize = Number(fontSizeSelect.value);
      else delete style.fontSize;
      if (Object.keys(style).length) sheet().cellStyles[ref] = style;
      else delete sheet().cellStyles[ref];
    });
    refresh();
  }
  fontFamilySelect.addEventListener('change', applySelectedFont);
  fontSizeSelect.addEventListener('change', applySelectedFont);

  let resize = null;
  host.addEventListener('pointerdown', (event) => {
    const colHeader = event.target.closest('[data-resize-col]');
    const rowHeader = event.target.closest('[data-resize-row]');
    const target = colHeader || rowHeader;
    if (!target || event.button !== 0) return;
    event.preventDefault();
    const isColumn = Boolean(colHeader);
    const index = Number(target.dataset[isColumn ? 'resizeCol' : 'resizeRow']);
    resize = {
      isColumn, index, start: isColumn ? event.clientX : event.clientY,
      size: isColumn ? Number(sheet().columnWidths?.[numToCol(index)] || 96) : Number(sheet().rowHeights?.[index] || 24),
    };
    target.setPointerCapture(event.pointerId);
  });
  host.addEventListener('pointermove', (event) => {
    if (!resize) return;
    const delta = (resize.isColumn ? event.clientX : event.clientY) - resize.start;
    const value = clamp(Math.round(resize.size + delta), resize.isColumn ? 40 : 18, resize.isColumn ? 600 : 300);
    const table = host.querySelector('table');
    if (resize.isColumn) {
      table.querySelectorAll('col')[resize.index].style.width = `${value}px`;
      table.style.width = `${gridWidth() + value - columnWidth(resize.index)}px`;
    }
    else table.querySelectorAll('tbody tr')[resize.index - 1].style.height = `${value}px`;
  });
  const finishResize = (event) => {
    if (!resize) return;
    const state = resize;
    const delta = (state.isColumn ? event.clientX : event.clientY) - state.start;
    const value = clamp(Math.round(state.size + delta), state.isColumn ? 40 : 18, state.isColumn ? 600 : 300);
    if (state.isColumn) {
      sheet().columnWidths ||= {};
      sheet().columnWidths[numToCol(state.index)] = value;
      const table = host.querySelector('table');
      table.querySelectorAll('col')[state.index].style.width = `${value}px`;
      table.style.width = `${gridWidth()}px`;
    } else {
      sheet().rowHeights ||= {};
      sheet().rowHeights[state.index] = value;
    }
    resize = null;
  };
  host.addEventListener('pointerup', finishResize);
  host.addEventListener('pointercancel', finishResize);

  toolbar.addEventListener('click', (event) => {
    const button = event.target.closest('button');
    if (!button) return;
    const { action } = button.dataset;
    if (action === 'add-rows') { rows = Math.min(MAX_ROWS, rows + 100); build(); }
    else if (action === 'add-columns') { cols = Math.min(MAX_COLS, cols + 10); build(); }
    else if (action === 'clear-cell') {
      forSelectedCells((ref) => {
        delete sheet().cells[ref];
        if (sheet().images) delete sheet().images[ref];
      });
      refresh();
      syncFormula();
    } else if (action === 'image') imageInput.click();
  });

  imageInput.addEventListener('change', () => {
    const file = imageInput.files?.[0];
    imageInput.value = '';
    if (!file) return;
    if (!['image/png', 'image/jpeg', 'image/gif'].includes(file.type) || file.size > MAX_IMAGE_BYTES) {
      window.alert('Choose a PNG, JPEG, or GIF image smaller than 8 MB.');
      return;
    }
    const reader = new FileReader();
    reader.addEventListener('load', () => {
      if (typeof reader.result !== 'string') return;
      const probe = new Image();
      probe.addEventListener('load', () => {
        const scale = Math.min(1, 320 / probe.naturalWidth, 240 / probe.naturalHeight);
        const ref = toRef(cursor.col, cursor.row);
        sheet().images ||= {};
        sheet().images[ref] = { src: reader.result, width: Math.max(1, Math.round(probe.naturalWidth * scale)), height: Math.max(1, Math.round(probe.naturalHeight * scale)), alt: file.name };
        refresh();
      }, { once: true });
      probe.src = reader.result;
    }, { once: true });
    reader.readAsDataURL(file);
  });

  renderTabs();
  build();

  return {
    kind: 'sheet',
    name: '',

    newDocument() {
      doc = blankSheet();
      rows = START_ROWS;
      cols = START_COLS;
      cursor = { col: 1, row: 1 };
      selectionAnchor = { ...cursor };
      renderTabs();
      build();
    },

    setDocument(next) {
      doc = normalizeSheet(next);
      rows = START_ROWS;
      cols = START_COLS;
      for (const item of doc.sheets) {
        const bounds = usedRange({ ...item.cells, ...item.images });
        if (!bounds) continue;
        rows = Math.min(MAX_ROWS, Math.max(rows, bounds.rows + 20));
        cols = Math.min(MAX_COLS, Math.max(cols, bounds.cols + 5));
      }
      cursor = { col: 1, row: 1 };
      selectionAnchor = { ...cursor };
      renderTabs();
      build();
    },

    contents() {
      commitEdit();
      return {
        sheets: doc.sheets.map((item) => ({ name: item.name, cells: { ...item.cells }, images: { ...(item.images || {}) }, columnWidths: { ...(item.columnWidths || {}) }, rowHeights: { ...(item.rowHeights || {}) }, cellStyles: Object.fromEntries(Object.entries(item.cellStyles || {}).map(([ref, style]) => [ref, { ...style }])) })),
        active: doc.active,
      };
    },

    focus() { host.focus(); },
  };
}
