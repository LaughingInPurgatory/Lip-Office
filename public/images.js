// Layered raster editor for Images. Pixels stay on the browser canvas; only
// saved projects and exported images are sent to the document store.
export function createImageEditor(root) {
  const canvas = root.querySelector('#image-stage');
  const stageScroll = root.querySelector('.image-stage-scroll');
  const ctx = canvas.getContext('2d', { willReadFrequently: true });
  const list = root.querySelector('#image-layer-list');
  const input = root.querySelector('#image-input');
  const widthInput = root.querySelector('#image-width');
  const heightInput = root.querySelector('#image-height');
  const zoomInput = root.querySelector('#image-zoom');
  const opacityInput = root.querySelector('#image-layer-opacity');
  const opacityValue = root.querySelector('#image-layer-opacity-value');
  const blendInput = root.querySelector('#image-layer-blend');
  const shadowToggle = root.querySelector('#image-layer-shadow');
  const shadowOptions = root.querySelector('#image-shadow-options');
  const shadowColor = root.querySelector('#image-shadow-color');
  const shadowBlur = root.querySelector('#image-shadow-blur');
  const shadowDistance = root.querySelector('#image-shadow-distance');
  const layerMenu = root.querySelector('#image-layer-menu');
  const effectPicker = root.querySelector('#image-effect-picker');
  const effectToggle = root.querySelector('#image-effect-toggle');
  const effectMenu = root.querySelector('#image-effect-menu');
  const effectLabel = root.querySelector('#image-effect-label');
  const effectChoices = [...root.querySelectorAll('[data-effect]')];
  const layerMenuTitle = root.querySelector('#image-layer-menu-title');
  const editTextButton = root.querySelector('#image-edit-text');
  const backgroundColorInput = root.querySelector('#image-background-color');
  const transparentBackgroundInput = root.querySelector('#image-background-transparent');
  const textSizeInput = root.querySelector('#image-text-size');
  const textFontInput = root.querySelector('#image-text-font');
  const textColorInput = root.querySelector('#image-text-color');
  const textBoldInput = root.querySelector('#image-text-bold');
  const textItalicInput = root.querySelector('#image-text-italic');
  const textUnderlineInput = root.querySelector('#image-text-underline');
  const transformInputs = {
    scaleX: root.querySelector('#image-transform-x'),
    scaleY: root.querySelector('#image-transform-y'),
    rotation: root.querySelector('#image-transform-rotation'),
    skewX: root.querySelector('#image-transform-skew-x'),
    skewY: root.querySelector('#image-transform-skew-y'),
  };
  const brushTypeInput = root.querySelector('#image-brush-type');
  const textDialog = root.querySelector('#image-text-dialog');
  const textInput = root.querySelector('#image-text-value');
  let width = 1200;
  let height = 800;
  let layers = [];
  let active = 0;
  let tool = 'move';
  let selection = null;
  let gesture = null;
  let undoStack = [];
  let redoStack = [];
  let name = '';
  let backgroundColor = '';
  let textEditIndex = null;
  let cloneSource = null;

  const blankCanvas = () => {
    const layer = document.createElement('canvas');
    layer.width = width;
    layer.height = height;
    return layer;
  };

  function snapshot() {
    return { width, height, active, backgroundColor, layers: layers.map((layer) => ({
      name: layer.name, src: layer.canvas.toDataURL('image/png'), x: layer.x, y: layer.y,
      opacity: layer.opacity, visible: layer.visible, blendMode: layer.blendMode, shadow: { ...layer.shadow },
      transform: { ...layer.transform }, text: layer.text ? { ...layer.text } : null,
    })) };
  }

  function checkpoint() {
    undoStack.push(snapshot());
    if (undoStack.length > 15) undoStack.shift();
    redoStack = [];
  }

  function drawLayer(target, layer) {
    if (!layer.visible) return;
    target.save();
    target.globalAlpha = layer.opacity;
    target.globalCompositeOperation = layer.blendMode;
    target.shadowColor = layer.shadow.enabled ? layer.shadow.color : 'transparent';
    target.shadowBlur = layer.shadow.enabled ? layer.shadow.blur : 0;
    target.shadowOffsetX = layer.shadow.enabled ? layer.shadow.x : 0;
    target.shadowOffsetY = layer.shadow.enabled ? layer.shadow.y : 0;
    const transform = layer.transform;
    target.translate(layer.x + width / 2, layer.y + height / 2);
    target.rotate(transform.rotation * Math.PI / 180);
    target.transform(1, Math.tan(transform.skewY * Math.PI / 180), Math.tan(transform.skewX * Math.PI / 180), 1, 0, 0);
    target.scale(transform.scaleX, transform.scaleY);
    target.drawImage(layer.canvas, -width / 2, -height / 2);
    target.restore();
  }

  function paint() {
    canvas.width = width;
    canvas.height = height;
    canvas.style.width = `${Math.round(width * Number(zoomInput.value))}px`;
    canvas.style.height = `${Math.round(height * Number(zoomInput.value))}px`;
    ctx.clearRect(0, 0, width, height);
    if (backgroundColor) {
      ctx.fillStyle = backgroundColor;
      ctx.fillRect(0, 0, width, height);
    }
    for (const layer of layers) drawLayer(ctx, layer);
    ctx.globalAlpha = 1;
    drawSelection();
    if (tool === 'free-transform') drawTransformControls();
    renderLayers();
  }

  function transformHandles(layer) {
    const corners = [[0, 0], [width, 0], [width, height], [0, height]].map(([x, y]) => forwardLayerPoint({ x, y }, layer));
    const midpoint = (a, b) => ({ x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 });
    const top = midpoint(corners[0], corners[1]), left = midpoint(corners[0], corners[3]);
    const zoom = Number(zoomInput.value);
    return [
      ...corners.map((p, index) => ({ kind: `corner-${index}`, p })),
      { kind: 'scale-x', p: midpoint(corners[1], corners[2]) },
      { kind: 'scale-y', p: midpoint(corners[2], corners[3]) },
      { kind: 'skew-x', p: { x: top.x, y: top.y + 12 / zoom } },
      { kind: 'skew-y', p: { x: left.x + 12 / zoom, y: left.y } },
      { kind: 'rotate', p: { x: top.x, y: top.y + 34 / zoom } },
    ];
  }

  function drawTransformControls() {
    const layer = layers[active];
    if (!layer) return;
    const corners = [[0, 0], [width, 0], [width, height], [0, height]].map(([x, y]) => forwardLayerPoint({ x, y }, layer));
    const handles = transformHandles(layer), zoom = Number(zoomInput.value), size = 8 / zoom;
    ctx.save();
    ctx.lineWidth = 1.5 / zoom;
    ctx.strokeStyle = '#1a73e8';
    ctx.setLineDash([5 / zoom, 3 / zoom]);
    ctx.beginPath();
    corners.forEach((p, index) => index ? ctx.lineTo(p.x, p.y) : ctx.moveTo(p.x, p.y));
    ctx.closePath(); ctx.stroke(); ctx.setLineDash([]);
    const rotate = handles.find((handle) => handle.kind === 'rotate').p;
    const top = { x: (corners[0].x + corners[1].x) / 2, y: (corners[0].y + corners[1].y) / 2 };
    ctx.beginPath(); ctx.moveTo(top.x, top.y); ctx.lineTo(rotate.x, rotate.y); ctx.stroke();
    for (const { kind, p } of handles) {
      ctx.fillStyle = kind.startsWith('skew') ? '#f9ab00' : '#fff';
      ctx.strokeStyle = '#1a73e8';
      if (kind === 'rotate') {
        ctx.beginPath(); ctx.arc(p.x, p.y, size * .65, 0, Math.PI * 2); ctx.fill(); ctx.stroke();
      } else if (kind.startsWith('skew')) {
        ctx.beginPath(); ctx.moveTo(p.x, p.y - size / 2); ctx.lineTo(p.x + size / 2, p.y); ctx.lineTo(p.x, p.y + size / 2); ctx.lineTo(p.x - size / 2, p.y); ctx.closePath(); ctx.fill(); ctx.stroke();
      } else {
        ctx.fillRect(p.x - size / 2, p.y - size / 2, size, size); ctx.strokeRect(p.x - size / 2, p.y - size / 2, size, size);
      }
    }
    ctx.restore();
  }

  function drawSelection() {
    if (!selection) return;
    if (selection.type === 'mask') {
      const bounds = selectionBounds();
      if (!bounds) return;
      const outlineCanvas = document.createElement('canvas');
      outlineCanvas.width = bounds.w;
      outlineCanvas.height = bounds.h;
      const outlineContext = outlineCanvas.getContext('2d');
      const outline = outlineContext.createImageData(bounds.w, bounds.h);
      const mask = selection.data;
      const edge = (x, y) => x < 0 || y < 0 || x >= width || y >= height || !mask[y * width + x];
      for (let y = bounds.y; y < bounds.y + bounds.h; y++) for (let x = bounds.x; x < bounds.x + bounds.w; x++) {
        if (!mask[y * width + x] || ![edge(x - 1, y), edge(x + 1, y), edge(x, y - 1), edge(x, y + 1)].some(Boolean)) continue;
        const offset = ((y - bounds.y) * bounds.w + x - bounds.x) * 4;
        const color = (x + y) % 6 < 3 ? 255 : 32;
        outline.data[offset] = outline.data[offset + 1] = outline.data[offset + 2] = color;
        outline.data[offset + 3] = 255;
      }
      outlineContext.putImageData(outline, 0, 0);
      ctx.drawImage(outlineCanvas, bounds.x, bounds.y);
      return;
    }
    ctx.save();
    ctx.strokeStyle = '#fff';
    ctx.lineWidth = 2 / Number(zoomInput.value);
    ctx.setLineDash([5 / Number(zoomInput.value), 4 / Number(zoomInput.value)]);
    ctx.lineDashOffset = 0;
    ctx.beginPath();
    if (selection.type === 'lasso') {
      selection.points.forEach(([x, y], index) => index ? ctx.lineTo(x, y) : ctx.moveTo(x, y));
      ctx.closePath();
    } else {
      const { x, y, w, h } = selection;
      if (selection.type === 'ellipse') ctx.ellipse(x + w / 2, y + h / 2, Math.abs(w / 2), Math.abs(h / 2), 0, 0, Math.PI * 2);
      else ctx.rect(x, y, w, h);
    }
    ctx.stroke();
    ctx.strokeStyle = '#222';
    ctx.lineDashOffset = 5 / Number(zoomInput.value);
    ctx.stroke();
    ctx.restore();
  }

  function renderLayers() {
    list.replaceChildren();
    [...layers.keys()].reverse().forEach((index) => {
      const layer = layers[index];
      const row = document.createElement('li');
      row.className = `image-layer-row${index === active ? ' active' : ''}`;
      row.dataset.index = String(index);
      const visible = document.createElement('input');
      visible.type = 'checkbox';
      visible.checked = layer.visible;
      visible.setAttribute('aria-label', `Show ${layer.name}`);
      visible.addEventListener('change', () => { checkpoint(); layer.visible = visible.checked; paint(); });
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'image-layer-name';
      button.textContent = layer.name;
      button.title = 'Click to open layer controls';
      button.addEventListener('click', (event) => showLayerMenu(event, index));
      row.append(visible, button);
      list.append(row);
    });
    if (!layers.length) {
      const empty = document.createElement('li');
      empty.className = 'image-layer-empty';
      empty.textContent = 'Add a layer or open an image to begin.';
      list.append(empty);
    }
  }

  function syncLayerMenu() {
    const layer = layers[active];
    const value = layer ? Math.round(layer.opacity * 100) : 100;
    opacityInput.value = String(value);
    opacityValue.textContent = `${value}%`;
    layerMenuTitle.textContent = layer?.name || 'Layer';
    blendInput.value = layer?.blendMode || 'source-over';
    shadowToggle.checked = Boolean(layer?.shadow.enabled);
    shadowColor.value = layer?.shadow.color || '#000000';
    shadowBlur.value = String(layer?.shadow.blur ?? 12);
    shadowDistance.value = String(Math.round(Math.hypot(layer?.shadow.x || 0, layer?.shadow.y || 0)) || 8);
    shadowOptions.hidden = !shadowToggle.checked;
    editTextButton.hidden = !layer?.text;
    for (const [key, input] of Object.entries(transformInputs)) {
      input.value = String(key.startsWith('scale') ? Math.round((layer?.transform[key] || 1) * 100) : Math.round(layer?.transform[key] || 0));
    }
  }

  function showLayerMenu(event, index) {
    active = index;
    renderLayers();
    syncLayerMenu();
    layerMenu.hidden = false;
    const rect = layerMenu.getBoundingClientRect();
    layerMenu.style.left = `${Math.max(8, Math.min(event.clientX, window.innerWidth - rect.width - 8))}px`;
    layerMenu.style.top = `${Math.max(8, Math.min(event.clientY, window.innerHeight - rect.height - 8))}px`;
  }

  function addLayer(layer = {}) {
    if (layers.length >= 40) { window.alert('A project can have up to 40 layers.'); return null; }
    const layerCanvas = blankCanvas();
    const created = { name: layer.name || `Layer ${layers.length + 1}`, canvas: layerCanvas,
      x: Number(layer.x) || 0, y: Number(layer.y) || 0, opacity: layer.opacity ?? 1,
      visible: layer.visible !== false, blendMode: layer.blendMode || 'source-over',
      shadow: { enabled: false, color: '#000000', blur: 12, x: 8, y: 8, ...(layer.shadow || {}) },
      transform: { scaleX: 1, scaleY: 1, rotation: 0, skewX: 0, skewY: 0, ...(layer.transform || {}) },
      text: layer.text && typeof layer.text === 'object' ? { ...layer.text } : null };
    layers.push(created);
    active = layers.length - 1;
    return created;
  }

  function setTool(next) {
    tool = next;
    root.querySelectorAll('[data-tool]').forEach((button) => button.setAttribute('aria-pressed', String(button.dataset.tool === tool)));
    canvas.style.cursor = tool === 'move' ? 'move' : tool === 'free-transform' ? 'default' : 'crosshair';
    paint();
  }

  function point(event) {
    const rect = canvas.getBoundingClientRect();
    return { x: Math.max(0, Math.min(width, (event.clientX - rect.left) * width / rect.width)),
      y: Math.max(0, Math.min(height, (event.clientY - rect.top) * height / rect.height)) };
  }

  function clipToSelection(layerContext, layer = null) {
    if (!selection) return;
    if (layer) {
      const inverse = inverseLayerTransform(layer);
      layerContext.transform(inverse.a, inverse.b, inverse.c, inverse.d, inverse.e, inverse.f);
    }
    layerContext.beginPath();
    if (selection.type === 'mask') {
      const bounds = selectionBounds();
      if (!bounds) { layerContext.clip(); if (layer) layerContext.setTransform(1, 0, 0, 1, 0, 0); return; }
      for (let y = bounds.y; y < bounds.y + bounds.h; y++) {
        let start = -1;
        for (let x = bounds.x; x <= bounds.x + bounds.w; x++) {
          const inside = x < width && selection.data[y * width + x] > 0;
          if (inside && start < 0) start = x;
          if (!inside && start >= 0) { layerContext.rect(start, y, x - start, 1); start = -1; }
        }
      }
    } else if (selection.type === 'lasso') {
      selection.points.forEach(([x, y], index) => index ? layerContext.lineTo(x, y) : layerContext.moveTo(x, y));
      layerContext.closePath();
    } else {
      const { x, y, w, h } = selection;
      if (selection.type === 'ellipse') layerContext.ellipse(x + w / 2, y + h / 2, Math.max(0.5, Math.abs(w / 2)), Math.max(0.5, Math.abs(h / 2)), 0, 0, Math.PI * 2);
      else layerContext.rect(Math.min(x, x + w), Math.min(y, y + h), Math.abs(w), Math.abs(h));
    }
    layerContext.clip();
    if (layer) layerContext.setTransform(1, 0, 0, 1, 0, 0);
  }

  function inverseLayerTransform(layer, transform = layer.transform, x = layer.x, y = layer.y) {
    const { scaleX, scaleY, rotation, skewX, skewY } = transform;
    const angle = rotation * Math.PI / 180;
    const sx = Math.tan(skewX * Math.PI / 180), sy = Math.tan(skewY * Math.PI / 180);
    const cos = Math.cos(angle), sin = Math.sin(angle);
    const a = (cos - sin * sy) * scaleX;
    const b = (sin + cos * sy) * scaleX;
    const c = (cos * sx - sin) * scaleY;
    const d = (sin * sx + cos) * scaleY;
    const e = x + width / 2 - a * width / 2 - c * height / 2;
    const f = y + height / 2 - b * width / 2 - d * height / 2;
    const determinant = a * d - b * c;
    if (!determinant) return { a: 1, b: 0, c: 0, d: 1, e: -x, f: -y };
    return { a: d / determinant, b: -b / determinant, c: -c / determinant, d: a / determinant,
      e: (c * f - d * e) / determinant, f: (b * e - a * f) / determinant };
  }

  function inverseTransformPoint(point, layer, transform = layer.transform, x = layer.x, y = layer.y) {
    const inverse = inverseLayerTransform(layer, transform, x, y);
    return { x: inverse.a * point.x + inverse.c * point.y + inverse.e,
      y: inverse.b * point.x + inverse.d * point.y + inverse.f };
  }

  function pointInQuad(point, corners) {
    let inside = false;
    for (let i = 0, j = corners.length - 1; i < corners.length; j = i++) {
      const a = corners[i], b = corners[j];
      if ((a.y > point.y) !== (b.y > point.y) && point.x < (b.x - a.x) * (point.y - a.y) / (b.y - a.y) + a.x) inside = !inside;
    }
    return inside;
  }

  function forwardLayerPoint(pointInLayer, layer) {
    const { scaleX, scaleY, rotation, skewX, skewY } = layer.transform;
    const angle = rotation * Math.PI / 180;
    const sx = Math.tan(skewX * Math.PI / 180), sy = Math.tan(skewY * Math.PI / 180);
    const cos = Math.cos(angle), sin = Math.sin(angle);
    const a = (cos - sin * sy) * scaleX;
    const b = (sin + cos * sy) * scaleX;
    const c = (cos * sx - sin) * scaleY;
    const d = (sin * sx + cos) * scaleY;
    const e = layer.x + width / 2 - a * width / 2 - c * height / 2;
    const f = layer.y + height / 2 - b * width / 2 - d * height / 2;
    return { x: a * pointInLayer.x + c * pointInLayer.y + e, y: b * pointInLayer.x + d * pointInLayer.y + f };
  }

  function layerPoint(pointOnCanvas, layer) {
    const inverse = inverseLayerTransform(layer);
    return { x: inverse.a * pointOnCanvas.x + inverse.c * pointOnCanvas.y + inverse.e,
      y: inverse.b * pointOnCanvas.x + inverse.d * pointOnCanvas.y + inverse.f };
  }

  function applyBrushStyle(target, activeTool) {
    const size = Math.max(1, Math.min(300, Number(root.querySelector('#image-brush-size').value) || 1));
    const brush = brushTypeInput.value;
    target.lineJoin = brush === 'square' ? 'bevel' : 'round';
    target.lineCap = brush === 'square' ? 'butt' : 'round';
    target.lineWidth = brush === 'pencil' ? Math.max(1, size * .35) : brush === 'marker' ? size * 1.8 : size;
    target.strokeStyle = root.querySelector('#image-brush-color').value;
    target.globalAlpha = brush === 'marker' ? .48 : brush === 'airbrush' ? .22 : 1;
    target.shadowColor = brush === 'airbrush' && activeTool !== 'eraser' ? target.strokeStyle : 'transparent';
    target.shadowBlur = brush === 'airbrush' ? size * .7 : 0;
    target.globalCompositeOperation = activeTool === 'eraser' ? 'destination-out' : 'source-over';
  }

  function pointerDown(event) {
    if (!layers[active]) return;
    const p = point(event);
    if (tool === 'free-transform') {
      const layer = layers[active], zoom = Number(zoomInput.value);
      const hit = transformHandles(layer).find(({ p: handle }) => Math.hypot(handle.x - p.x, handle.y - p.y) <= 10 / zoom);
      const original = { ...layer.transform };
      const corners = [[0, 0], [width, 0], [width, height], [0, height]].map(([x, y]) => forwardLayerPoint({ x, y }, layer));
      const centerPoint = { x: layer.x + width / 2, y: layer.y + height / 2 };
      const inside = pointInQuad(p, corners);
      checkpoint(); canvas.setPointerCapture(event.pointerId);
      gesture = { tool, kind: hit?.kind || (inside ? 'move' : 'rotate'), start: p, original, x: layer.x, y: layer.y,
        startLocal: inverseTransformPoint(p, layer, original), center: { x: width / 2, y: height / 2 }, centerPoint,
        angle: Math.atan2(p.y - centerPoint.y, p.x - centerPoint.x) };
      return;
    }
    if (tool === 'fill') { checkpoint(); fillAt(p); return; }
    if (tool === 'picker') {
      const previousSelection = selection;
      selection = null;
      paint();
      const pixel = ctx.getImageData(Math.min(width - 1, Math.floor(p.x)), Math.min(height - 1, Math.floor(p.y)), 1, 1).data;
      if (pixel[3]) root.querySelector('#image-brush-color').value = `#${[pixel[0], pixel[1], pixel[2]].map((channel) => channel.toString(16).padStart(2, '0')).join('')}`;
      selection = previousSelection;
      paint();
      return;
    }
    if (tool === 'wand') {
      selectSimilarPixels(p);
      return;
    }
    if (tool === 'clone') {
      if (event.altKey) { cloneSource = layerPoint(p, layers[active]); return; }
      if (!cloneSource) { window.alert('Alt-click the canvas to set a clone source first.'); return; }
      canvas.setPointerCapture(event.pointerId);
      checkpoint();
      gesture = { tool, start: p, last: p, source: cloneSource };
      return;
    }
    canvas.setPointerCapture(event.pointerId);
    if (['rect', 'ellipse', 'lasso'].includes(tool)) {
      selection = null;
      gesture = { tool, start: p, current: p, points: [[p.x, p.y]] };
    } else if (tool.startsWith('shape-')) {
      checkpoint();
      gesture = { tool, start: p, current: p };
    } else if (tool === 'move') {
      checkpoint();
      gesture = { tool, start: p, x: layers[active].x, y: layers[active].y };
    } else {
      checkpoint();
      gesture = { tool, last: p };
      const layerContext = layers[active].canvas.getContext('2d');
      layerContext.save();
      clipToSelection(layerContext, layers[active]);
      applyBrushStyle(layerContext, tool);
      layerContext.beginPath();
      const local = layerPoint(p, layers[active]);
      layerContext.moveTo(local.x, local.y);
      layerContext.lineTo(local.x + 0.01, local.y + 0.01);
      layerContext.stroke();
      layerContext.restore();
    }
    paint();
  }

  function pointerMove(event) {
    if (!gesture) return;
    const p = point(event);
    if (gesture.tool === 'free-transform') {
      const layer = layers[active], original = gesture.original;
      if (gesture.kind === 'move') {
        layer.x = gesture.x + p.x - gesture.start.x;
        layer.y = gesture.y + p.y - gesture.start.y;
      } else if (gesture.kind === 'rotate') {
        const angle = Math.atan2(p.y - gesture.centerPoint.y, p.x - gesture.centerPoint.x);
        layer.transform.rotation = original.rotation + (angle - gesture.angle) * 180 / Math.PI;
      } else if (gesture.kind.startsWith('skew')) {
        const key = gesture.kind === 'skew-x' ? 'skewX' : 'skewY';
        const delta = key === 'skewX' ? p.x - gesture.start.x : p.y - gesture.start.y;
        layer.transform[key] = Math.max(-80, Math.min(80, original[key] + delta * .12));
      } else {
        const local = inverseTransformPoint(p, layer, original, gesture.x, gesture.y);
        const sx = (local.x - gesture.center.x) / (gesture.startLocal.x - gesture.center.x || 1);
        const sy = (local.y - gesture.center.y) / (gesture.startLocal.y - gesture.center.y || 1);
        if (gesture.kind.startsWith('corner-') || gesture.kind === 'scale-x') layer.transform.scaleX = Math.max(.01, Math.min(10, original.scaleX * sx));
        if (gesture.kind.startsWith('corner-') || gesture.kind === 'scale-y') layer.transform.scaleY = Math.max(.01, Math.min(10, original.scaleY * sy));
      }
    } else if (gesture.tool === 'move') {
      layers[active].x = gesture.x + p.x - gesture.start.x;
      layers[active].y = gesture.y + p.y - gesture.start.y;
    } else if (gesture.tool === 'clone') {
      const layer = layers[active], current = layerPoint(p, layer), start = layerPoint(gesture.start, layer);
      const sourceX = gesture.source.x + current.x - start.x, sourceY = gesture.source.y + current.y - start.y;
      const size = Math.max(2, Number(root.querySelector('#image-brush-size').value) || 18);
      const target = layer.canvas.getContext('2d');
      target.save();
      clipToSelection(target, layer);
      target.beginPath();
      target.arc(current.x, current.y, size / 2, 0, Math.PI * 2);
      target.clip();
      target.drawImage(layer.canvas, sourceX - size / 2, sourceY - size / 2, size, size, current.x - size / 2, current.y - size / 2, size, size);
      target.restore();
    } else if (['rect', 'ellipse', 'lasso', 'shape-rect', 'shape-ellipse'].includes(gesture.tool)) {
      gesture.current = p;
      if (gesture.tool === 'lasso') gesture.points.push([p.x, p.y]);
    } else {
      const layer = layers[active];
      const layerContext = layer.canvas.getContext('2d');
      layerContext.save();
      clipToSelection(layerContext, layer);
      applyBrushStyle(layerContext, gesture.tool);
      layerContext.beginPath();
      const previous = layerPoint(gesture.last, layer), local = layerPoint(p, layer);
      layerContext.moveTo(previous.x, previous.y);
      layerContext.lineTo(local.x, local.y);
      layerContext.stroke();
      layerContext.restore();
      gesture.last = p;
    }
    paint();
    if (gesture.tool.startsWith('shape-')) drawDraftShape(gesture);
    else if (['rect', 'ellipse', 'lasso'].includes(gesture.tool)) {
      drawDraftSelection(gesture);
    }
  }

  function drawDraftShape(draft) {
    const x = Math.min(draft.start.x, draft.current.x), y = Math.min(draft.start.y, draft.current.y);
    const w = Math.abs(draft.current.x - draft.start.x), h = Math.abs(draft.current.y - draft.start.y);
    ctx.save();
    applyBrushStyle(ctx, 'brush');
    ctx.beginPath();
    if (draft.tool === 'shape-ellipse') ctx.ellipse(x + w / 2, y + h / 2, w / 2, h / 2, 0, 0, Math.PI * 2);
    else ctx.rect(x, y, w, h);
    ctx.stroke();
    ctx.restore();
  }

  function drawDraftSelection(draft) {
    ctx.save();
    ctx.strokeStyle = '#fff';
    ctx.lineWidth = 2 / Number(zoomInput.value);
    ctx.setLineDash([5 / Number(zoomInput.value), 4 / Number(zoomInput.value)]);
    ctx.beginPath();
    if (draft.tool === 'lasso') {
      draft.points.forEach(([x, y], index) => index ? ctx.lineTo(x, y) : ctx.moveTo(x, y));
    } else {
      const x = Math.min(draft.start.x, draft.current.x), y = Math.min(draft.start.y, draft.current.y);
      const w = Math.abs(draft.current.x - draft.start.x), h = Math.abs(draft.current.y - draft.start.y);
      if (draft.tool === 'ellipse') ctx.ellipse(x + w / 2, y + h / 2, w / 2, h / 2, 0, 0, Math.PI * 2);
      else ctx.rect(x, y, w, h);
    }
    ctx.stroke();
    ctx.strokeStyle = '#222';
    ctx.lineDashOffset = 5 / Number(zoomInput.value);
    ctx.stroke();
    ctx.restore();
  }

  function pointerUp() {
    if (!gesture) return;
    if (gesture.tool === 'free-transform') syncLayerMenu();
    else if (['rect', 'ellipse'].includes(gesture.tool)) {
      selection = { type: gesture.tool, x: Math.min(gesture.start.x, gesture.current.x), y: Math.min(gesture.start.y, gesture.current.y),
        w: Math.abs(gesture.current.x - gesture.start.x), h: Math.abs(gesture.current.y - gesture.start.y) };
    } else if (gesture.tool === 'lasso') {
      selection = { type: 'lasso', points: gesture.points };
    } else if (gesture.tool.startsWith('shape-')) {
      const layer = layers[active];
      const target = layer.canvas.getContext('2d');
      const start = layerPoint(gesture.start, layer), end = layerPoint(gesture.current, layer);
      const x = Math.min(start.x, end.x), y = Math.min(start.y, end.y);
      const w = Math.abs(end.x - start.x), h = Math.abs(end.y - start.y);
      target.save();
      clipToSelection(target, layer);
      applyBrushStyle(target, 'brush');
      target.beginPath();
      if (gesture.tool === 'shape-ellipse') target.ellipse(x + w / 2, y + h / 2, w / 2, h / 2, 0, 0, Math.PI * 2);
      else target.rect(x, y, w, h);
      target.stroke();
      target.restore();
    }
    gesture = null;
    paint();
  }

  async function loadDocument(doc = {}, resetHistory = true) {
    width = Math.min(6000, Math.max(1, Math.trunc(Number(doc.width) || 1200)));
    height = Math.min(6000, Math.max(1, Math.trunc(Number(doc.height) || 800)));
    if (width * height > 36000000) throw new Error('image dimensions are too large');
    widthInput.value = String(width);
    heightInput.value = String(height);
    backgroundColor = typeof doc.backgroundColor === 'string' ? doc.backgroundColor : '';
    transparentBackgroundInput.checked = !backgroundColor;
    backgroundColorInput.value = backgroundColor || '#ffffff';
    backgroundColorInput.disabled = false;
    layers = [];
    for (const item of (doc.layers || [])) {
      const layer = addLayer({ ...item, name: item.name });
      if (!item.src) continue;
      const image = new Image();
      image.src = item.src;
      await image.decode();
      if (doc.fitToImage && doc.layers.length === 1) {
        width = Math.min(6000, image.naturalWidth);
        height = Math.min(6000, image.naturalHeight);
        if (width * height > 36000000) { height = Math.max(1, Math.floor(36000000 / width)); }
        widthInput.value = String(width);
        heightInput.value = String(height);
        layer.canvas.width = width;
        layer.canvas.height = height;
        const scale = Math.min(width / image.naturalWidth, height / image.naturalHeight, 1);
        layer.canvas.getContext('2d').drawImage(image, 0, 0, Math.round(image.naturalWidth * scale), Math.round(image.naturalHeight * scale));
      } else {
        layer.canvas.getContext('2d').drawImage(image, 0, 0, Math.min(width, image.naturalWidth), Math.min(height, image.naturalHeight));
      }
    }
    if (!layers.length) addLayer({ name: 'Layer 1' });
    active = Math.min(layers.length - 1, Math.max(0, Number(doc.active) || 0));
    selection = null;
    if (resetHistory) { undoStack = []; redoStack = []; }
    paint();
  }

  function newDocument() {
    const nextWidth = Math.min(6000, Math.max(1, Math.trunc(Number(widthInput.value) || 1200)));
    const nextHeight = Math.min(6000, Math.max(1, Math.trunc(Number(heightInput.value) || 800)));
    if (nextWidth * nextHeight > 36000000) { window.alert('Canvas is too large. Maximum is 36 megapixels.'); return; }
    width = nextWidth;
    height = nextHeight;
    backgroundColor = '';
    transparentBackgroundInput.checked = true;
    backgroundColorInput.disabled = false;
    layers = [];
    addLayer({ name: 'Layer 1' });
    selection = null;
    setTool('move');
    undoStack = [];
    redoStack = [];
    name = '';
    paint();
    root.dispatchEvent(new CustomEvent('images:newdocument', { bubbles: true }));
  }

  async function addImage(file) {
    if (!file || file.size > 22 * 1024 * 1024) {
      window.alert('Choose an image smaller than 22 MB.');
      return;
    }
    const bytes = new Uint8Array(await file.arrayBuffer());
    let binary = '';
    for (let offset = 0; offset < bytes.length; offset += 0x8000) binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
    const response = await fetch('/api/parse', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: file.name, data: btoa(binary) }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || `Could not load ${file.name}.`);
    if (result.kind !== 'image' || !result.doc?.layers?.length) throw new Error('That file does not contain a supported image.');
    const imported = result.doc;
    const documentScale = imported.fitToImage ? 1 : Math.min(1, width / imported.width, height / imported.height);
    checkpoint();
    for (let index = 0; index < imported.layers.length; index++) {
      const sourceLayer = imported.layers[index];
      if (!sourceLayer.src) continue;
      const image = new Image(); image.src = sourceLayer.src; await image.decode();
      const scale = imported.fitToImage
        ? Math.min(1, width / image.naturalWidth, height / image.naturalHeight)
        : documentScale;
      const layer = addLayer({
        name: imported.layers.length > 1 ? sourceLayer.name : file.name.replace(/\.[^.]+$/, ''),
        x: imported.fitToImage ? 0 : (Number(sourceLayer.x) || 0) * scale,
        y: imported.fitToImage ? 0 : (Number(sourceLayer.y) || 0) * scale,
        opacity: sourceLayer.opacity, visible: sourceLayer.visible, blendMode: sourceLayer.blendMode,
        shadow: sourceLayer.shadow, transform: sourceLayer.transform,
      });
      if (!layer) continue;
      const w = Math.round(image.naturalWidth * scale), h = Math.round(image.naturalHeight * scale);
      layer.canvas.getContext('2d').drawImage(image, 0, 0, w, h);
    }
    paint();
  }

  async function undo() {
    if (!undoStack.length) return;
    redoStack.push(snapshot());
    await loadDocument(undoStack.pop(), false);
  }

  async function redo() {
    if (!redoStack.length) return;
    undoStack.push(snapshot());
    await loadDocument(redoStack.pop(), false);
  }

  function deleteSelection() {
    if (!selection || !layers[active]) return;
    checkpoint();
    const target = layers[active].canvas.getContext('2d');
    target.save();
    clipToSelection(target, layers[active]);
    target.clearRect(0, 0, width, height);
    target.restore();
    selection = null;
    paint();
  }

  function fillAt(pointOnCanvas) {
    const layer = layers[active];
    if (!layer) return;
    const pointInLayer = layerPoint(pointOnCanvas, layer);
    const x0 = Math.floor(pointInLayer.x), y0 = Math.floor(pointInLayer.y);
    const target = layer.canvas.getContext('2d', { willReadFrequently: true });
    const image = target.getImageData(0, 0, width, height);
    const { data } = image;
    const pixelCount = width * height;
    if (x0 < 0 || y0 < 0 || x0 >= width || y0 >= height) return;
    const previousSelection = selection;
    selection = null;
    paint();
    const composite = ctx.getImageData(0, 0, width, height).data;
    selection = previousSelection;
    paint();

    let mask = null;
    if (selection) {
      const maskCanvas = document.createElement('canvas');
      maskCanvas.width = width;
      maskCanvas.height = height;
      const maskContext = maskCanvas.getContext('2d');
      maskContext.save();
      clipToSelection(maskContext, layer);
      maskContext.fillStyle = '#fff';
      maskContext.fillRect(0, 0, width, height);
      maskContext.restore();
      mask = maskContext.getImageData(0, 0, width, height).data;
    }

    const sampleX = Math.max(0, Math.min(width - 1, Math.floor(pointOnCanvas.x)));
    const sampleY = Math.max(0, Math.min(height - 1, Math.floor(pointOnCanvas.y)));
    const start = (sampleY * width + sampleX) * 4;
    const targetStart = (y0 * width + x0) * 4;
    if (mask && !mask[targetStart + 3]) return;
    const seed = [composite[start], composite[start + 1], composite[start + 2], composite[start + 3]];
    const fill = root.querySelector('#image-brush-color').value.match(/[0-9a-f]{2}/gi).map((part) => parseInt(part, 16)).concat(255);
    const tolerance = Math.max(0, Math.min(255, Number(root.querySelector('#image-fill-tolerance').value) || 0));
    if (seed.every((channel, index) => Math.abs(channel - fill[index]) <= tolerance)) return;
    const visited = new Uint8Array(pixelCount);
    const matches = (x, y) => {
      if (x < 0 || y < 0 || x >= width || y >= height) return false;
      const pixel = y * width + x;
      if (visited[pixel] || mask && !mask[pixel * 4 + 3]) return false;
      const canvasPoint = forwardLayerPoint({ x, y }, layer);
      const canvasX = Math.max(0, Math.min(width - 1, Math.floor(canvasPoint.x)));
      const canvasY = Math.max(0, Math.min(height - 1, Math.floor(canvasPoint.y)));
      const offset = (canvasY * width + canvasX) * 4;
      return seed.every((channel, index) => Math.abs(composite[offset + index] - channel) <= tolerance);
    };
    const spans = [[x0, y0]];
    while (spans.length) {
      const [seedX, y] = spans.pop();
      if (!matches(seedX, y)) continue;
      let left = seedX;
      while (matches(left - 1, y)) left--;
      let inAbove = false, inBelow = false;
      for (let x = left; matches(x, y); x++) {
        const pixel = y * width + x, offset = pixel * 4;
        visited[pixel] = 1;
        data[offset] = fill[0]; data[offset + 1] = fill[1]; data[offset + 2] = fill[2]; data[offset + 3] = 255;
        if (y > 0) {
          if (matches(x, y - 1)) { if (!inAbove) spans.push([x, y - 1]); inAbove = true; } else inAbove = false;
        }
        if (y + 1 < height) {
          if (matches(x, y + 1)) { if (!inBelow) spans.push([x, y + 1]); inBelow = true; } else inBelow = false;
        }
      }
    }
    target.putImageData(image, 0, 0);
    paint();
  }

  function selectSimilarPixels(pointOnCanvas) {
    selection = null;
    paint();
    const image = ctx.getImageData(0, 0, width, height);
    const sx = Math.max(0, Math.min(width - 1, Math.floor(pointOnCanvas.x)));
    const sy = Math.max(0, Math.min(height - 1, Math.floor(pointOnCanvas.y)));
    const start = (sy * width + sx) * 4;
    const seed = [image.data[start], image.data[start + 1], image.data[start + 2], image.data[start + 3]];
    const tolerance = Math.max(0, Math.min(255, Number(root.querySelector('#image-fill-tolerance').value) || 0));
    const visited = new Uint8Array(width * height);
    const mask = new Uint8ClampedArray(width * height);
    let minX = width, minY = height, maxX = -1, maxY = -1;
    const matches = (x, y) => {
      if (x < 0 || y < 0 || x >= width || y >= height) return false;
      const pixel = y * width + x, offset = pixel * 4;
      return !visited[pixel] && seed.every((channel, index) => Math.abs(image.data[offset + index] - channel) <= tolerance);
    };
    const spans = [[sx, sy]];
    while (spans.length) {
      const [seedX, y] = spans.pop();
      if (!matches(seedX, y)) continue;
      let left = seedX;
      while (matches(left - 1, y)) left--;
      let inAbove = false, inBelow = false;
      for (let x = left; matches(x, y); x++) {
        const pixel = y * width + x;
        visited[pixel] = 1;
        mask[pixel] = 255;
        minX = Math.min(minX, x); minY = Math.min(minY, y); maxX = Math.max(maxX, x); maxY = Math.max(maxY, y);
        if (y > 0) {
          if (matches(x, y - 1)) { if (!inAbove) spans.push([x, y - 1]); inAbove = true; } else inAbove = false;
        }
        if (y + 1 < height) {
          if (matches(x, y + 1)) { if (!inBelow) spans.push([x, y + 1]); inBelow = true; } else inBelow = false;
        }
      }
    }
    selection = { type: 'mask', data: mask, bounds: maxX < 0 ? null : { x: minX, y: minY, w: maxX - minX + 1, h: maxY - minY + 1 } };
    paint();
  }

  function selectionBounds() {
    if (!selection) return null;
    if (selection.type === 'mask') {
      if (selection.bounds) return selection.bounds;
      let minX = width, minY = height, maxX = -1, maxY = -1;
      for (let i = 0; i < selection.data.length; i++) if (selection.data[i]) {
        const x = i % width, y = Math.floor(i / width);
        minX = Math.min(minX, x); minY = Math.min(minY, y); maxX = Math.max(maxX, x); maxY = Math.max(maxY, y);
      }
      return maxX < 0 ? null : { x: minX, y: minY, w: maxX - minX + 1, h: maxY - minY + 1 };
    }
    if (selection.type !== 'lasso') return { x: selection.x, y: selection.y, w: selection.w, h: selection.h };
    const xs = selection.points.map(([x]) => x), ys = selection.points.map(([, y]) => y);
    const x = Math.min(...xs), y = Math.min(...ys);
    return { x, y, w: Math.max(...xs) - x, h: Math.max(...ys) - y };
  }

  function cropToSelection() {
    const bounds = selectionBounds();
    if (!bounds || bounds.w < 1 || bounds.h < 1) return;
    const x = Math.floor(bounds.x), y = Math.floor(bounds.y);
    const nextWidth = Math.max(1, Math.ceil(bounds.x + bounds.w) - x);
    const nextHeight = Math.max(1, Math.ceil(bounds.y + bounds.h) - y);
    checkpoint();
    for (const layer of layers) {
      const cropped = document.createElement('canvas');
      cropped.width = nextWidth;
      cropped.height = nextHeight;
      cropped.getContext('2d').drawImage(layer.canvas, layer.x - x, layer.y - y);
      layer.canvas = cropped;
      layer.x = 0;
      layer.y = 0;
    }
    width = nextWidth;
    height = nextHeight;
    widthInput.value = String(width);
    heightInput.value = String(height);
    selection = null;
    paint();
  }

  function resizeImage() {
    const nextWidth = Math.min(6000, Math.max(1, Math.trunc(Number(widthInput.value) || width)));
    const nextHeight = Math.min(6000, Math.max(1, Math.trunc(Number(heightInput.value) || height)));
    if (nextWidth * nextHeight > 36000000) { window.alert('Image is too large. Maximum is 36 megapixels.'); return; }
    if (nextWidth === width && nextHeight === height) return;
    checkpoint();
    const scaleX = nextWidth / width, scaleY = nextHeight / height;
    for (const layer of layers) {
      const resized = document.createElement('canvas');
      resized.width = nextWidth;
      resized.height = nextHeight;
      resized.getContext('2d').drawImage(layer.canvas, layer.x * scaleX, layer.y * scaleY, width * scaleX, height * scaleY);
      layer.canvas = resized;
      layer.x = 0;
      layer.y = 0;
    }
    width = nextWidth;
    height = nextHeight;
    selection = null;
    paint();
  }

  function textSpecFromDialog() {
    return {
      content: textInput.value.slice(0, 2000),
      font: textFontInput.value,
      size: Math.max(6, Math.min(300, Number(textSizeInput.value) || 48)),
      color: textColorInput.value,
      bold: textBoldInput.checked,
      italic: textItalicInput.checked,
      underline: textUnderlineInput.checked,
    };
  }

  function drawTextLayer(layer) {
    const spec = layer.text;
    if (!spec) return;
    const target = layer.canvas.getContext('2d');
    target.clearRect(0, 0, width, height);
    target.fillStyle = spec.color;
    target.font = `${spec.italic ? 'italic ' : ''}${spec.bold ? 'bold ' : ''}${spec.size}px "${spec.font}", sans-serif`;
    target.textBaseline = 'top';
    spec.content.split('\n').slice(0, 100).forEach((line, index) => {
      const y = 40 + index * spec.size * 1.2;
      target.fillText(line, 40, y, width - 80);
      if (spec.underline) {
        const metrics = target.measureText(line);
        target.fillRect(40, y + spec.size * 1.05, Math.min(metrics.width, width - 80), Math.max(1, spec.size / 18));
      }
    });
  }

  function openTextDialog(editIndex = null) {
    textEditIndex = editIndex;
    const spec = editIndex === null ? null : layers[editIndex]?.text;
    root.querySelector('#image-text-title').textContent = spec ? 'Edit text layer' : 'Add text layer';
    textInput.value = spec?.content || '';
    textFontInput.value = spec?.font || 'Arial';
    textSizeInput.value = String(spec?.size || 48);
    textColorInput.value = spec?.color || root.querySelector('#image-brush-color').value;
    textBoldInput.checked = Boolean(spec?.bold);
    textItalicInput.checked = Boolean(spec?.italic);
    textUnderlineInput.checked = Boolean(spec?.underline);
    textDialog.showModal();
    textInput.focus();
  }

  function addTextLayer() {
    const spec = textSpecFromDialog();
    if (!spec.content.trim()) return;
    if (textEditIndex !== null && layers[textEditIndex]) {
      checkpoint();
      const layer = layers[textEditIndex];
      layer.text = spec;
      layer.name = spec.content.trim().split('\n')[0].slice(0, 40) || 'Text';
      active = textEditIndex;
      drawTextLayer(layer);
    } else {
      checkpoint();
      const layer = addLayer({ name: spec.content.trim().split('\n')[0].slice(0, 40) || 'Text', text: spec });
      if (!layer) return;
      layer.text = spec;
      drawTextLayer(layer);
    }
    textEditIndex = null;
    paint();
  }

  function adjustLayer(adjustment) {
    const layer = layers[active];
    if (!layer) return;
    checkpoint();
    try {
      const target = layer.canvas.getContext('2d', { willReadFrequently: true });
      const pixels = target.getImageData(0, 0, width, height);
      const brightness = Math.max(-100, Math.min(100, Number(root.querySelector('#image-brightness').value) || 0)) * 2.55;
      const contrast = Math.max(-100, Math.min(100, Number(root.querySelector('#image-contrast').value) || 0));
      const factor = (259 * (contrast + 255)) / (255 * (259 - contrast));
      for (let i = 0; i < pixels.data.length; i += 4) {
        if (adjustment === 'invert') {
          pixels.data[i] = 255 - pixels.data[i];
          pixels.data[i + 1] = 255 - pixels.data[i + 1];
          pixels.data[i + 2] = 255 - pixels.data[i + 2];
        } else if (adjustment === 'grayscale') {
          const gray = Math.round(.299 * pixels.data[i] + .587 * pixels.data[i + 1] + .114 * pixels.data[i + 2]);
          pixels.data[i] = pixels.data[i + 1] = pixels.data[i + 2] = gray;
        } else {
          for (let channel = 0; channel < 3; channel++) pixels.data[i + channel] = Math.max(0, Math.min(255, factor * (pixels.data[i + channel] - 128) + 128 + brightness));
        }
      }
      target.putImageData(pixels, 0, 0);
      layer.text = null;
      paint();
    } catch {
      window.alert('This image is too large for an adjustment in this browser.');
    }
  }

  async function applyFilter() {
    const layer = layers[active];
    if (!layer) return;
    const kind = root.querySelector('#image-filter-kind').value;
    const amount = Math.max(1, Math.min(100, Number(root.querySelector('#image-filter-amount').value) || 1));
    const filtered = document.createElement('canvas');
    filtered.width = width;
    filtered.height = height;
    const filterContext = filtered.getContext('2d', { willReadFrequently: true });
    const source = layer.canvas.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, width, height);
    const output = filterContext.createImageData(width, height);
    const colorEffects = new Set(['black-white', 'desaturate', 'hue', 'saturate', 'invert', 'noise', 'gamma', 'vibrance', 'heatmap', 'solarize', 'vintage', 'blueprint', 'night-vision', 'pencil', '1977', 'aden', 'clarendon', 'gingham', 'inkwell', 'lo-fi', 'toaster', 'valencia', 'x-pro-ii', 'dither', 'dot-screen', 'enrich', 'posterize', 'duotone', 'cross-process']);
    if (['sharpen', 'emboss', 'edge'].includes(kind)) {
      const strength = Math.min(2.5, amount / 20);
      const kernels = {
        sharpen: { weights: [0, -strength, 0, -strength, 1 + 4 * strength, -strength, 0, -strength, 0], bias: 0 },
        emboss: { weights: [-2, -1, 0, -1, 1, 1, 0, 1, 2], bias: 128 },
        edge: { weights: [0, -1, 0, -1, 4, -1, 0, -1, 0], bias: 128 },
      };
      const { weights, bias } = kernels[kind];
      for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
        const offset = (y * width + x) * 4;
        for (let channel = 0; channel < 3; channel++) {
          let sum = bias;
          let k = 0;
          for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++, k++) {
            const nx = Math.max(0, Math.min(width - 1, x + dx));
            const ny = Math.max(0, Math.min(height - 1, y + dy));
            sum += source.data[(ny * width + nx) * 4 + channel] * weights[k];
          }
          output.data[offset + channel] = Math.max(0, Math.min(255, sum));
        }
        output.data[offset + 3] = source.data[offset + 3];
      }
      filterContext.putImageData(output, 0, 0);
    } else if (kind === 'pixelate') {
      const block = Math.max(2, Math.round(amount / 2));
      const small = document.createElement('canvas');
      small.width = Math.max(1, Math.ceil(width / block)); small.height = Math.max(1, Math.ceil(height / block));
      const smallContext = small.getContext('2d'); smallContext.imageSmoothingEnabled = false; smallContext.drawImage(layer.canvas, 0, 0, small.width, small.height);
      filterContext.imageSmoothingEnabled = false; filterContext.drawImage(small, 0, 0, width, height);
    } else if (kind === 'chromatic') {
      const shift = Math.max(1, Math.round(amount / 8));
      for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
        const i = (y * width + x) * 4, left = (y * width + Math.max(0, x - shift)) * 4, right = (y * width + Math.min(width - 1, x + shift)) * 4;
        output.data[i] = source.data[left]; output.data[i + 1] = source.data[i + 1]; output.data[i + 2] = source.data[right + 2]; output.data[i + 3] = source.data[i + 3];
      }
      filterContext.putImageData(output, 0, 0);
    } else if (kind === 'glow') {
      const blurred = document.createElement('canvas'); blurred.width = width; blurred.height = height;
      const blurContext = blurred.getContext('2d'); blurContext.filter = `blur(${Math.max(2, amount / 5)}px)`; blurContext.drawImage(layer.canvas, 0, 0);
      filterContext.drawImage(layer.canvas, 0, 0); filterContext.globalCompositeOperation = 'screen'; filterContext.globalAlpha = amount / 100;
      filterContext.drawImage(blurred, 0, 0); filterContext.globalCompositeOperation = 'source-over'; filterContext.globalAlpha = 1;
    } else if (kind === 'outline') {
      for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
        const i = (y * width + x) * 4, right = (y * width + Math.min(width - 1, x + 1)) * 4, below = (Math.min(height - 1, y + 1) * width + x) * 4;
        const edge = Math.abs(source.data[i] - source.data[right]) + Math.abs(source.data[i + 1] - source.data[right + 1]) + Math.abs(source.data[i] - source.data[below]) + Math.abs(source.data[i + 1] - source.data[below + 1]);
        const ink = edge > Math.max(12, 180 - amount * 1.6) ? 0 : 255;
        output.data[i] = output.data[i + 1] = output.data[i + 2] = ink; output.data[i + 3] = source.data[i + 3];
      }
      filterContext.putImageData(output, 0, 0);
    } else if (kind === 'vhs') {
      const shift = Math.max(1, Math.round(amount / 15));
      for (let y = 0; y < height; y++) {
        const glitch = Math.random() < amount / 450 ? Math.round((Math.random() * 2 - 1) * amount) : 0;
        for (let x = 0; x < width; x++) {
          const i = (y * width + x) * 4, red = (y * width + Math.max(0, Math.min(width - 1, x - shift + glitch))) * 4, blue = (y * width + Math.max(0, Math.min(width - 1, x + shift + glitch))) * 4;
          const scanline = y % 3 === 0 ? .78 : 1, grain = (Math.random() - .5) * amount * .55;
          output.data[i] = Math.max(0, Math.min(255, (source.data[red] + grain) * scanline));
          output.data[i + 1] = Math.max(0, Math.min(255, (source.data[i + 1] + grain) * scanline));
          output.data[i + 2] = Math.max(0, Math.min(255, (source.data[blue + 2] + grain) * scanline)); output.data[i + 3] = source.data[i + 3];
        }
      }
      filterContext.putImageData(output, 0, 0);
    } else if (kind === 'mosaic') {
      const block = Math.max(2, Math.round(amount / 4));
      const small = document.createElement('canvas');
      small.width = Math.max(1, Math.ceil(width / block));
      small.height = Math.max(1, Math.ceil(height / block));
      const smallContext = small.getContext('2d');
      smallContext.imageSmoothingEnabled = false;
      smallContext.drawImage(layer.canvas, 0, 0, small.width, small.height);
      filterContext.imageSmoothingEnabled = false;
      filterContext.drawImage(small, 0, 0, width, height);
    } else if (['denoise', 'oil'].includes(kind)) {
      const radius = kind === 'oil' ? Math.min(3, Math.max(1, Math.round(amount / 25))) : 1;
      const buckets = new Uint32Array(256);
      const red = new Float32Array(256), green = new Float32Array(256), blue = new Float32Array(256);
      for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
        const bins = kind === 'oil' ? buckets.fill(0) : null;
        if (bins) { red.fill(0); green.fill(0); blue.fill(0); }
        let totalR = 0, totalG = 0, totalB = 0, count = 0;
        for (let dy = -radius; dy <= radius; dy++) for (let dx = -radius; dx <= radius; dx++) {
          const sx = Math.max(0, Math.min(width - 1, x + dx)), sy = Math.max(0, Math.min(height - 1, y + dy));
          const si = (sy * width + sx) * 4;
          if (!source.data[si + 3]) continue;
          const r = source.data[si], g = source.data[si + 1], b = source.data[si + 2];
          totalR += r; totalG += g; totalB += b; count++;
          if (bins) { const bucket = Math.min(255, Math.floor((r + g + b) / 3)); bins[bucket]++; red[bucket] += r; green[bucket] += g; blue[bucket] += b; }
        }
        const i = (y * width + x) * 4;
        if (bins) {
          let best = 0;
          for (let k = 1; k < 256; k++) if (bins[k] > bins[best]) best = k;
          output.data[i] = bins[best] ? red[best] / bins[best] : source.data[i];
          output.data[i + 1] = bins[best] ? green[best] / bins[best] : source.data[i + 1];
          output.data[i + 2] = bins[best] ? blue[best] / bins[best] : source.data[i + 2];
        } else {
          output.data[i] = count ? totalR / count : source.data[i];
          output.data[i + 1] = count ? totalG / count : source.data[i + 1];
          output.data[i + 2] = count ? totalB / count : source.data[i + 2];
        }
        output.data[i + 3] = source.data[i + 3];
      }
      filterContext.putImageData(output, 0, 0);
    } else if (kind === 'bulge') {
      const strength = (amount - 50) / 50;
      const cx = width / 2, cy = height / 2, radius = Math.min(width, height) * .48;
      for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
        const dx = x - cx, dy = y - cy, distance = Math.hypot(dx, dy), i = (y * width + x) * 4;
        let sx = x, sy = y;
        if (distance < radius && distance > 0) {
          const normalized = distance / radius;
          const warped = normalized * (1 + strength * (1 - normalized) ** 2);
          sx = Math.max(0, Math.min(width - 1, Math.round(cx + dx * warped / normalized)));
          sy = Math.max(0, Math.min(height - 1, Math.round(cy + dy * warped / normalized)));
        }
        const si = (sy * width + sx) * 4;
        output.data.set(source.data.subarray(si, si + 4), i);
      }
      filterContext.putImageData(output, 0, 0);
    } else if (kind === 'zoom-blur') {
      filterContext.drawImage(layer.canvas, 0, 0);
      const passes = 8, strength = amount / 100 * .12;
      for (let n = 1; n <= passes; n++) {
        const scale = 1 + strength * n / passes;
        const w = width * scale, h = height * scale;
        filterContext.globalAlpha = .12;
        filterContext.drawImage(layer.canvas, (width - w) / 2, (height - h) / 2, w, h);
      }
      filterContext.globalAlpha = 1;
    } else if (kind === 'tilt-shift') {
      const blurred = document.createElement('canvas'); blurred.width = width; blurred.height = height;
      const blurContext = blurred.getContext('2d'); blurContext.filter = `blur(${Math.max(2, amount / 5)}px)`; blurContext.drawImage(layer.canvas, 0, 0);
      filterContext.drawImage(blurred, 0, 0);
      filterContext.globalCompositeOperation = 'destination-in';
      const mask = filterContext.createLinearGradient(0, 0, 0, height);
      mask.addColorStop(0, 'rgba(0,0,0,.9)'); mask.addColorStop(.35, 'rgba(0,0,0,0)'); mask.addColorStop(.65, 'rgba(0,0,0,0)'); mask.addColorStop(1, 'rgba(0,0,0,.9)');
      filterContext.fillStyle = mask; filterContext.fillRect(0, 0, width, height);
      filterContext.globalCompositeOperation = 'source-over'; filterContext.drawImage(layer.canvas, 0, 0);
    } else if (kind === 'jpg-compression') {
      const image = new Image();
      image.src = layer.canvas.toDataURL('image/jpeg', Math.max(.05, 1 - amount / 110));
      try { await image.decode(); filterContext.drawImage(image, 0, 0, width, height); }
      catch { filterContext.drawImage(layer.canvas, 0, 0); }
    } else if (kind === 'box-blur') {
      const radius = Math.min(12, Math.max(1, Math.round(amount / 8)));
      const horizontal = new Uint8ClampedArray(source.data.length);
      for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) for (let c = 0; c < 4; c++) {
        let sum = 0, n = 0;
        for (let dx = -radius; dx <= radius; dx++) { const sx = Math.max(0, Math.min(width - 1, x + dx)); sum += source.data[(y * width + sx) * 4 + c]; n++; }
        horizontal[(y * width + x) * 4 + c] = sum / n;
      }
      for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) for (let c = 0; c < 4; c++) {
        let sum = 0, n = 0;
        for (let dy = -radius; dy <= radius; dy++) { const sy = Math.max(0, Math.min(height - 1, y + dy)); sum += horizontal[(sy * width + x) * 4 + c]; n++; }
        output.data[(y * width + x) * 4 + c] = sum / n;
      }
      filterContext.putImageData(output, 0, 0);
    } else if (kind === 'stack-blur') {
      const radius = Math.max(1, amount / 10);
      let previous = layer.canvas;
      for (let pass = 0; pass < 3; pass++) {
        const stage = pass === 2 ? filtered : document.createElement('canvas');
        stage.width = width; stage.height = height;
        const stageContext = stage.getContext('2d'); stageContext.filter = `blur(${radius}px)`; stageContext.drawImage(previous, 0, 0); stageContext.filter = 'none';
        previous = stage;
      }
    } else if (kind === 'gaussian-blur' || kind === 'blur') {
      const radius = Math.max(1, amount / 4);
      filterContext.filter = `blur(${radius}px)`; filterContext.drawImage(layer.canvas, 0, 0); filterContext.filter = 'none';
    } else if (kind === 'vignette') {
      filterContext.drawImage(layer.canvas, 0, 0);
      const radius = Math.hypot(width, height) * .7;
      const gradient = filterContext.createRadialGradient(width / 2, height / 2, Math.min(width, height) * .1, width / 2, height / 2, radius);
      gradient.addColorStop(0, 'rgba(0,0,0,0)'); gradient.addColorStop(.65, 'rgba(0,0,0,0)'); gradient.addColorStop(1, `rgba(0,0,0,${amount / 100})`);
      filterContext.globalCompositeOperation = 'source-atop'; filterContext.fillStyle = gradient; filterContext.fillRect(0, 0, width, height); filterContext.globalCompositeOperation = 'source-over';
    } else if (colorEffects.has(kind)) {
      const bayer = [0, 8, 2, 10, 12, 4, 14, 6, 3, 11, 1, 9, 15, 7, 13, 5];
      const instagram = {
        '1977': [1.08, .94, 1.05, 12], aden: [1.04, .92, .98, 8], clarendon: [1.18, 1.12, 1.12, 0], gingham: [.92, .92, .94, 10],
        inkwell: [1, 1, 1, 0], 'lo-fi': [1.18, 1.08, 1.15, -4], toaster: [1.1, .94, .82, 16], valencia: [1.08, 1.02, .92, 8], 'x-pro-ii': [1.22, 1.08, 1.16, -8],
      }[kind];
      const duotoneHighlight = root.querySelector('#image-brush-color').value.match(/[0-9a-f]{2}/gi).map((part) => parseInt(part, 16));
      const luminance = (r, g, b) => .2126 * r + .7152 * g + .0722 * b;
      for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
        const i = (y * width + x) * 4, r = source.data[i], g = source.data[i + 1], b = source.data[i + 2], lum = luminance(r, g, b);
        let nr = r, ng = g, nb = b;
        if (kind === 'black-white') nr = ng = nb = lum < amount * 2.55 ? 0 : 255;
        else if (kind === 'desaturate' || kind === 'inkwell') { const mix = kind === 'inkwell' ? 1 : amount / 100; nr += (lum - r) * mix; ng += (lum - g) * mix; nb += (lum - b) * mix; }
        else if (kind === 'hue') { const angle = amount * 3.6 * Math.PI / 180, cos = Math.cos(angle), sin = Math.sin(angle); nr = (.213 + cos * .787 - sin * .213) * r + (.715 - cos * .715 - sin * .715) * g + (.072 - cos * .072 + sin * .928) * b; ng = (.213 - cos * .213 + sin * .143) * r + (.715 + cos * .285 + sin * .140) * g + (.072 - cos * .072 - sin * .283) * b; nb = (.213 - cos * .213 - sin * .787) * r + (.715 - cos * .715 + sin * .715) * g + (.072 + cos * .928 + sin * .072) * b; }
        else if (kind === 'saturate') { const f = Math.max(0, 1 + amount / 25); nr = lum + (r - lum) * f; ng = lum + (g - lum) * f; nb = lum + (b - lum) * f; }
        else if (kind === 'invert') nr = 255 - r, ng = 255 - g, nb = 255 - b;
        else if (kind === 'noise' || kind === 'vintage') { const grain = kind === 'noise' ? (Math.random() * 2 - 1) * amount * .65 : ((Math.random() * 2 - 1) * amount * .12); nr += grain; ng += grain; nb += grain; if (kind === 'vintage') nr += 14; nb -= 10; }
        else if (kind === 'gamma') { const gamma = 2 ** ((amount - 50) / 25); nr = 255 * (r / 255) ** (1 / gamma); ng = 255 * (g / 255) ** (1 / gamma); nb = 255 * (b / 255) ** (1 / gamma); }
        else if (kind === 'vibrance') { const avg = (r + g + b) / 3, max = Math.max(r, g, b), boost = (max - avg) * (-amount / 100) * .02; nr += (max - r) * boost; ng += (max - g) * boost; nb += (max - b) * boost; }
        else if (kind === 'heatmap') { const t = lum / 255 * 4, segment = Math.min(3, Math.floor(t)), f = t - segment, stops = [[0,0,80],[0,80,255],[0,230,180],[255,220,0],[220,0,0]]; nr = stops[segment][0] * (1 - f) + stops[segment + 1][0] * f; ng = stops[segment][1] * (1 - f) + stops[segment + 1][1] * f; nb = stops[segment][2] * (1 - f) + stops[segment + 1][2] * f; }
        else if (kind === 'solarize') { const threshold = 255 - amount * 2.55; nr = r > threshold ? 255 - r : r; ng = g > threshold ? 255 - g : g; nb = b > threshold ? 255 - b : b; }
        else if (kind === 'blueprint') { nr = Math.min(255, 8 + lum * .15); ng = Math.min(255, 35 + lum * .55); nb = Math.min(255, 70 + lum * .72); }
        else if (kind === 'night-vision') { nr = lum * .08; ng = Math.min(255, lum * 1.45); nb = lum * .12; }
        else if (kind === 'pencil') { const edge = Math.abs(lum - luminance(source.data[Math.max(0, i - 4)], source.data[Math.max(0, i - 3)], source.data[Math.max(0, i - 2)])); nr = ng = nb = Math.max(0, 255 - edge * (1 + amount / 20)); }
        else if (kind === 'dither') { const threshold = (bayer[(y % 4) * 4 + x % 4] / 16) * 255; nr = ng = nb = lum > threshold ? 255 : 0; }
        else if (kind === 'dot-screen') { const pattern = (Math.sin(x * .22) * Math.sin(y * .22) + 1) * 127.5; nr = ng = nb = lum > pattern ? 255 : 0; }
        else if (kind === 'enrich') { const contrast = 1.12 + amount / 160; nr = (r - 128) * contrast + 128; ng = (g - 128) * contrast + 128; nb = (b - 128) * contrast + 128; }
        else if (kind === 'posterize') { const levels = Math.max(2, Math.round(2 + amount * .14)), step = 255 / (levels - 1); nr = Math.round(r / step) * step; ng = Math.round(g / step) * step; nb = Math.round(b / step) * step; }
        else if (kind === 'duotone') { const mix = Math.min(1, .35 + amount * .0065), t = lum / 255; const tone = [18 + (duotoneHighlight[0] - 18) * t, 24 + (duotoneHighlight[1] - 24) * t, 48 + (duotoneHighlight[2] - 48) * t]; nr += (tone[0] - nr) * mix; ng += (tone[1] - ng) * mix; nb += (tone[2] - nb) * mix; }
        else if (kind === 'cross-process') { nr = 255 * (r / 255) ** .82 + amount * .18; ng = 255 * (g / 255) ** 1.08; nb = 255 * (b / 255) ** .78 - amount * .12; }
        else if (instagram) { const [contrast, saturation, warm, tint] = instagram; nr = ((r - 128) * contrast + 128) * saturation + tint; ng = ((g - 128) * contrast + 128) * saturation; nb = ((b - 128) * contrast + 128) * saturation / warm - tint * .25; }
        output.data[i] = Math.max(0, Math.min(255, nr)); output.data[i + 1] = Math.max(0, Math.min(255, ng)); output.data[i + 2] = Math.max(0, Math.min(255, nb)); output.data[i + 3] = source.data[i + 3];
      }
      filterContext.putImageData(output, 0, 0);
    } else {
      filterContext.filter = kind === 'sepia' ? `sepia(${amount}%)` : `saturate(${amount * 3}%)`;
      filterContext.drawImage(layer.canvas, 0, 0);
    }
    checkpoint();
    const target = layer.canvas.getContext('2d');
    target.save();
    clipToSelection(target, layer);
    target.clearRect(0, 0, width, height);
    target.drawImage(filtered, 0, 0);
    target.restore();
    layer.text = null;
    paint();
  }

  function layerAction(action) {
    const layer = layers[active];
    if (action === 'image-new-layer') { checkpoint(); addLayer(); }
    if (action === 'image-duplicate-layer' && layer) {
      if (layers.length >= 40) { window.alert('A project can have up to 40 layers.'); return; }
      checkpoint();
      const copy = addLayer({ name: `${layer.name} copy`, x: layer.x, y: layer.y, opacity: layer.opacity, visible: layer.visible,
        blendMode: layer.blendMode, shadow: { ...layer.shadow }, transform: { ...layer.transform }, text: layer.text ? { ...layer.text } : null });
      copy.canvas.getContext('2d').drawImage(layer.canvas, 0, 0);
    }
    if (action === 'image-remove-layer' && layers.length > 1) { checkpoint(); layers.splice(active, 1); active = Math.min(active, layers.length - 1); }
    if (action === 'image-layer-up' && active < layers.length - 1) { checkpoint(); [layers[active], layers[active + 1]] = [layers[active + 1], layers[active]]; active++; }
    if (action === 'image-layer-down' && active > 0) { checkpoint(); [layers[active], layers[active - 1]] = [layers[active - 1], layers[active]]; active--; }
    if (action === 'image-clear-layer' && layer) { checkpoint(); layer.canvas.getContext('2d').clearRect(0, 0, width, height); layer.text = null; }
    if (action === 'image-rename-layer' && layer) {
      const nextName = window.prompt('Layer name:', layer.name);
      if (nextName?.trim()) { checkpoint(); layer.name = nextName.trim().slice(0, 80); }
    }
    if (action === 'image-merge-down') {
      if (active < 1) window.alert('Select a layer above another layer to merge it down.');
      else {
        checkpoint();
        const lower = layers[active - 1], upper = layer, mergedCanvas = blankCanvas();
        drawLayer(mergedCanvas.getContext('2d'), lower);
        drawLayer(mergedCanvas.getContext('2d'), upper);
        layers.splice(active - 1, 2, { ...lower, name: `${lower.name} + ${upper.name}`, canvas: mergedCanvas, x: 0, y: 0,
          opacity: 1, visible: true, blendMode: 'source-over', shadow: { enabled: false, color: '#000000', blur: 0, x: 0, y: 0 },
          transform: { scaleX: 1, scaleY: 1, rotation: 0, skewX: 0, skewY: 0 }, text: null });
        active--;
      }
    }
    if (action === 'image-flatten') {
      checkpoint();
      const flattened = blankCanvas(), target = flattened.getContext('2d');
      if (backgroundColor) { target.fillStyle = backgroundColor; target.fillRect(0, 0, width, height); }
      for (const item of layers) drawLayer(target, item);
      layers = [{ name: 'Flattened image', canvas: flattened, x: 0, y: 0, opacity: 1, visible: true, blendMode: 'source-over',
        shadow: { enabled: false, color: '#000000', blur: 0, x: 0, y: 0 }, transform: { scaleX: 1, scaleY: 1, rotation: 0, skewX: 0, skewY: 0 }, text: null }];
      backgroundColor = ''; active = 0; selection = null;
    }
    layerMenu.hidden = true;
    paint();
  }

  function updateLayerTransform(key, value) {
    const layer = layers[active];
    if (!layer || !Number.isFinite(value)) return;
    checkpoint();
    const bounds = key.startsWith('scale') ? [.01, 10] : key === 'rotation' ? [-180, 180] : [-80, 80];
    layer.transform[key] = Math.max(bounds[0], Math.min(bounds[1], value));
    syncLayerMenu();
    paint();
  }

  function closeEffectMenu() {
    effectMenu.hidden = true;
    effectToggle.setAttribute('aria-expanded', 'false');
  }

  root.addEventListener('click', (event) => {
    if (!effectPicker.contains(event.target)) closeEffectMenu();
    const button = event.target.closest('button');
    if (!button) return;
    if (button === effectToggle) {
      effectMenu.hidden = !effectMenu.hidden;
      if (!effectMenu.hidden) effectMenu.querySelectorAll('details').forEach((category) => { category.open = false; });
      effectToggle.setAttribute('aria-expanded', String(!effectMenu.hidden));
      return;
    }
    if (button.dataset.effect) {
      root.querySelector('#image-filter-kind').value = button.dataset.effect;
      effectLabel.textContent = button.dataset.label;
      effectChoices.forEach((choice) => choice.setAttribute('aria-pressed', String(choice === button)));
      closeEffectMenu();
      return;
    }
    if (button.dataset.tool) setTool(button.dataset.tool);
    if (button.dataset.action === 'image-add') input.click();
    if (button.dataset.action === 'image-create') newDocument();
    if (button.dataset.action === 'image-delete-selection') deleteSelection();
    if (button.dataset.action === 'image-deselect') { selection = null; paint(); }
    if (button.dataset.action === 'image-crop') cropToSelection();
    if (button.dataset.action === 'image-resize') resizeImage();
    if (button.dataset.action === 'image-add-text') openTextDialog();
    if (button.dataset.action === 'image-adjust') adjustLayer('adjust');
    if (button.dataset.action === 'image-filter') applyFilter();
    if (button.dataset.action === 'image-grayscale') adjustLayer('grayscale');
    if (button.dataset.action === 'image-invert') adjustLayer('invert');
    if (button.dataset.action === 'image-undo') undo();
    if (button.dataset.action === 'image-redo') redo();
    if (button.dataset.action?.startsWith('image-') && button.dataset.action.includes('layer')) layerAction(button.dataset.action);
    if (button.dataset.action === 'image-merge-down' || button.dataset.action === 'image-flatten') layerAction(button.dataset.action);
    const menuActions = { duplicate: 'image-duplicate-layer', up: 'image-layer-up', down: 'image-layer-down', delete: 'image-remove-layer', clear: 'image-clear-layer', rename: 'image-rename-layer', 'merge-down': 'image-merge-down' };
    if (button.dataset.menuAction === 'close') layerMenu.hidden = true;
    else if (menuActions[button.dataset.menuAction]) layerAction(menuActions[button.dataset.menuAction]);
    else if (button.dataset.menuAction === 'edit-text') { layerMenu.hidden = true; openTextDialog(active); }
    else if (button.dataset.menuAction === 'reset-transform' && layers[active]) {
      checkpoint();
      layers[active].transform = { scaleX: 1, scaleY: 1, rotation: 0, skewX: 0, skewY: 0 };
      syncLayerMenu();
      paint();
    }
  });
  textDialog.addEventListener('close', () => {
    if (textDialog.returnValue === 'add') addTextLayer();
    else textEditIndex = null;
  });
  input.addEventListener('change', () => { addImage(input.files?.[0]).catch((err) => window.alert(err.message)); input.value = ''; });
  window.addEventListener('pointerdown', (event) => {
    if (!layerMenu.hidden && !layerMenu.contains(event.target) && !list.contains(event.target)) layerMenu.hidden = true;
    if (!effectMenu.hidden && !effectPicker.contains(event.target)) closeEffectMenu();
  });
  opacityInput.addEventListener('pointerdown', checkpoint);
  opacityInput.addEventListener('input', () => {
    if (!layers[active]) return;
    layers[active].opacity = Number(opacityInput.value) / 100;
    opacityValue.textContent = `${opacityInput.value}%`;
    paint();
  });
  blendInput.addEventListener('change', () => {
    if (!layers[active]) return;
    checkpoint();
    layers[active].blendMode = blendInput.value;
    paint();
  });
  shadowToggle.addEventListener('change', () => {
    if (!layers[active]) return;
    checkpoint();
    layers[active].shadow.enabled = shadowToggle.checked;
    shadowOptions.hidden = !shadowToggle.checked;
    paint();
  });
  shadowColor.addEventListener('change', () => {
    if (!layers[active]) return;
    checkpoint();
    layers[active].shadow.color = shadowColor.value;
    paint();
  });
  for (const range of [shadowBlur, shadowDistance]) range.addEventListener('pointerdown', checkpoint);
  shadowBlur.addEventListener('input', () => {
    if (!layers[active]) return;
    layers[active].shadow.blur = Number(shadowBlur.value);
    paint();
  });
  shadowDistance.addEventListener('input', () => {
    if (!layers[active]) return;
    const distance = Number(shadowDistance.value) / Math.SQRT2;
    layers[active].shadow.x = distance;
    layers[active].shadow.y = distance;
    paint();
  });
  transformInputs.scaleX.addEventListener('change', () => updateLayerTransform('scaleX', Number(transformInputs.scaleX.value) / 100));
  transformInputs.scaleY.addEventListener('change', () => updateLayerTransform('scaleY', Number(transformInputs.scaleY.value) / 100));
  for (const key of ['rotation', 'skewX', 'skewY']) transformInputs[key].addEventListener('change', () => updateLayerTransform(key, Number(transformInputs[key].value)));
  backgroundColorInput.addEventListener('input', () => {
    backgroundColor = backgroundColorInput.value;
    transparentBackgroundInput.checked = false;
    paint();
  });
  backgroundColorInput.addEventListener('pointerdown', checkpoint);
  transparentBackgroundInput.addEventListener('change', () => {
    checkpoint();
    backgroundColor = transparentBackgroundInput.checked ? '' : backgroundColorInput.value;
    paint();
  });
  canvas.addEventListener('pointerdown', pointerDown);
  canvas.addEventListener('pointermove', pointerMove);
  canvas.addEventListener('pointerup', pointerUp);
  canvas.addEventListener('pointercancel', pointerUp);
  zoomInput.addEventListener('change', paint);
  let canvasDropDepth = 0;
  stageScroll.addEventListener('dragenter', (event) => {
    if (!event.dataTransfer?.types?.includes('Files')) return;
    event.preventDefault();
    event.stopPropagation();
    canvasDropDepth++;
    stageScroll.classList.add('image-drop-active');
  });
  stageScroll.addEventListener('dragover', (event) => {
    if (!event.dataTransfer?.types?.includes('Files')) return;
    event.preventDefault();
    event.stopPropagation();
    event.dataTransfer.dropEffect = 'copy';
  });
  stageScroll.addEventListener('dragleave', (event) => {
    if (!event.dataTransfer?.types?.includes('Files')) return;
    event.stopPropagation();
    canvasDropDepth = Math.max(0, canvasDropDepth - 1);
    if (!canvasDropDepth || !stageScroll.contains(event.relatedTarget)) stageScroll.classList.remove('image-drop-active');
  });
  stageScroll.addEventListener('drop', async (event) => {
    if (!event.dataTransfer?.types?.includes('Files')) return;
    event.preventDefault();
    event.stopPropagation();
    canvasDropDepth = 0;
    stageScroll.classList.remove('image-drop-active');
    const files = [...(event.dataTransfer.files || [])];
    if (!files.length) return;
    if (files.some((file) => !/\.(png|jpe?g|gif|tiff?|bmp|iff|lbm|psd|webp|avif)$/i.test(file.name))) {
      window.alert('Drop a supported image file onto the canvas.');
      return;
    }
    for (const file of files) {
      try { await addImage(file); }
      catch (error) { window.alert(error.message || `Could not load ${file.name}.`); }
    }
  });
  window.addEventListener('keydown', (event) => {
    if (root.classList.contains('hidden') || /INPUT|TEXTAREA|SELECT/.test(event.target.tagName)) return;
    if (event.key === 'Escape' && !effectMenu.hidden) { closeEffectMenu(); return; }
    if (event.key === 'Escape' && !layerMenu.hidden) { layerMenu.hidden = true; return; }
    if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'z') { event.preventDefault(); event.shiftKey ? redo() : undo(); }
    else if (event.key.toLowerCase() === 'v') setTool('move');
    else if (event.key.toLowerCase() === 'm') setTool('rect');
    else if (event.key.toLowerCase() === 'b') setTool('brush');
    else if (event.key.toLowerCase() === 'e') setTool('eraser');
    else if (event.key.toLowerCase() === 'g') setTool('fill');
    else if (event.key.toLowerCase() === 'c') setTool('clone');
    else if (event.key === 'Delete' || event.key === 'Backspace') { event.preventDefault(); deleteSelection(); }
  });

  setTool('move');
  backgroundColorInput.disabled = false;
  addLayer({ name: 'Layer 1' });
  paint();

  return {
    kind: 'image',
    get name() { return name; },
    set name(value) { name = value || ''; },
    contents(format = 'lipimg') {
      const savedLayers = layers.map((layer) => ({ name: layer.name, src: layer.canvas.toDataURL('image/png'),
        x: layer.x, y: layer.y, opacity: layer.opacity, visible: layer.visible,
        blendMode: layer.blendMode, shadow: { ...layer.shadow }, transform: { ...layer.transform },
        text: layer.text ? { ...layer.text } : null }));
      if (savedLayers.some((layer) => layer.src.length > 11190000) || savedLayers.reduce((total, layer) => total + layer.src.length, 0) > 24 * 1024 * 1024) {
        throw new Error('Image layers exceed the 24 MB save limit. Reduce the canvas size or remove layers.');
      }
      const doc = { width, height, active, backgroundColor, layers: savedLayers };
      if (format === 'psd') {
        const psdLayers = [];
        if (backgroundColor) {
          const background = document.createElement('canvas'); background.width = width; background.height = height;
          const backgroundContext = background.getContext('2d'); backgroundContext.fillStyle = backgroundColor; backgroundContext.fillRect(0, 0, width, height);
          psdLayers.push({ name: 'Canvas background', src: background.toDataURL('image/png'), x: 0, y: 0, opacity: 1, visible: true, blendMode: 'source-over' });
        }
        for (const layer of layers) {
          const raster = document.createElement('canvas'); raster.width = width; raster.height = height;
          const rasterContext = raster.getContext('2d'); const transform = layer.transform;
          rasterContext.translate(layer.x + width / 2, layer.y + height / 2);
          rasterContext.rotate(transform.rotation * Math.PI / 180);
          rasterContext.transform(1, Math.tan(transform.skewY * Math.PI / 180), Math.tan(transform.skewX * Math.PI / 180), 1, 0, 0);
          rasterContext.scale(transform.scaleX, transform.scaleY);
          rasterContext.drawImage(layer.canvas, -width / 2, -height / 2);
          psdLayers.push({ name: layer.name, src: raster.toDataURL('image/png'), x: 0, y: 0, opacity: layer.opacity, visible: layer.visible, blendMode: layer.blendMode });
        }
        doc.exportPsdLayers = psdLayers;
      }
      if (format !== 'lipimg') {
      const mime = ['jpg', 'jpeg'].includes(format) ? 'image/jpeg' : 'image/png';
        const selectionBeforeExport = selection;
        const previousBackground = backgroundColor;
        selection = null;
        try {
          if (mime === 'image/jpeg' && !backgroundColor) backgroundColor = '#ffffff';
          paint();
          doc.exportDataUrl = canvas.toDataURL(mime, .92);
        } finally {
          backgroundColor = previousBackground;
          selection = selectionBeforeExport;
          paint();
        }
        if (doc.exportDataUrl.length > 11190000) throw new Error('The exported image exceeds the 8 MB file limit. Resize the canvas and try again.');
      }
      return doc;
    },
    async setDocument(doc) { await loadDocument(doc); },
    newDocument,
    focus() { canvas.focus(); },
  };
}
