// Only inert raster formats are accepted. SVG can contain script and external
// references, so images are embedded as PNG, JPEG, or GIF data URLs.
export const MAX_IMAGE_BYTES = 8 * 1024 * 1024;
const IMAGE_TYPES = new Map([
  ['image/png', 'png'],
  ['image/jpeg', 'jpg'],
  ['image/gif', 'gif'],
]);

export function parseImageDataUrl(value) {
  const match = /^data:(image\/(?:png|jpeg|gif));base64,([A-Za-z0-9+/]*={0,2})$/.exec(String(value || ''));
  if (!match || !IMAGE_TYPES.has(match[1]) || !match[2]) return null;
  const bytes = Math.floor(match[2].length * 3 / 4) - (match[2].endsWith('==') ? 2 : match[2].endsWith('=') ? 1 : 0);
  if (bytes < 1 || bytes > MAX_IMAGE_BYTES) return null;
  return { mime: match[1], ext: IMAGE_TYPES.get(match[1]), base64: match[2], bytes };
}

export function imageDataUrl(mime, bytes) {
  const ext = [...IMAGE_TYPES.keys()].includes(mime) ? mime : null;
  if (!ext || !bytes?.length || bytes.length > MAX_IMAGE_BYTES) return '';
  return `data:${ext};base64,${Buffer.from(bytes).toString('base64')}`;
}
