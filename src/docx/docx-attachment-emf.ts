import { pngBytesFromDataUrl } from "./convert-windows-metafile";

const MAX_ICON_SIZE = 512;
const MAX_EMF_BYTES = 2 * 1024 * 1024;

function drawAlphaBitmap(
  context: CanvasRenderingContext2D,
  record: DataView,
  originX: number,
  originY: number,
): boolean {
  // MS-EMF EMR_ALPHABLEND with an identity source transform and a 32-bit DIB.
  if (record.byteLength < 108 || record.getUint8(40) !== 0 ||
    record.getInt32(44, true) !== 0 || record.getInt32(48, true) !== 0 ||
    record.getUint32(80, true) !== 0) return false;
  for (const [offset, value] of [[52, 1], [56, 0], [60, 0], [64, 1], [68, 0], [72, 0]] as const) {
    if (record.getFloat32(offset, true) !== value) return false;
  }
  const info = record.getUint32(84, true);
  const infoLength = record.getUint32(88, true);
  const bits = record.getUint32(92, true);
  const bitsLength = record.getUint32(96, true);
  if (info < 108 || infoLength < 40 || info + infoLength > record.byteLength ||
    bits < 108 || bits + bitsLength > record.byteLength) return false;
  const width = record.getInt32(info + 4, true);
  const signedHeight = record.getInt32(info + 8, true);
  const height = Math.abs(signedHeight);
  const compression = record.getUint32(info + 16, true);
  if (width <= 0 || width > MAX_ICON_SIZE || height <= 0 || height > MAX_ICON_SIZE ||
    record.getUint32(info, true) !== 40 || record.getUint16(info + 12, true) !== 1 ||
    record.getUint16(info + 14, true) !== 32 ||
    width !== record.getInt32(100, true) || height !== record.getInt32(104, true) ||
    width * height * 4 > bitsLength) return false;
  if (compression !== 0 && (compression !== 3 || infoLength < 52 ||
    record.getUint32(info + 40, true) !== 0x00ff0000 ||
    record.getUint32(info + 44, true) !== 0x0000ff00 ||
    record.getUint32(info + 48, true) !== 0x000000ff)) return false;

  const destinationWidth = record.getInt32(32, true);
  const destinationHeight = record.getInt32(36, true);
  if (destinationWidth <= 0 || destinationHeight <= 0 ||
    destinationWidth > MAX_ICON_SIZE || destinationHeight > MAX_ICON_SIZE) return false;
  const alphaMode = record.getUint8(43);
  if (alphaMode > 1) return false;
  const bitmap = context.canvas.ownerDocument.createElement("canvas");
  bitmap.width = width;
  bitmap.height = height;
  const bitmapContext = bitmap.getContext("2d");
  if (bitmapContext === null) return false;
  const pixels = bitmapContext.createImageData(width, height);
  const constantAlpha = record.getUint8(42) / 255;
  for (let y = 0; y < height; y++) {
    const row = signedHeight < 0 ? y : height - y - 1;
    for (let x = 0; x < width; x++) {
      const input = bits + (row * width + x) * 4;
      const output = (y * width + x) * 4;
      const alpha = alphaMode === 1 ? record.getUint8(input + 3) : 255;
      if (alpha === 0) continue;
      pixels.data[output] = record.getUint8(input + 2) * 255 / alpha;
      pixels.data[output + 1] = record.getUint8(input + 1) * 255 / alpha;
      pixels.data[output + 2] = record.getUint8(input) * 255 / alpha;
      pixels.data[output + 3] = alpha * constantAlpha;
    }
  }
  bitmapContext.putImageData(pixels, 0, 0);
  context.drawImage(bitmap, record.getInt32(24, true) - originX,
    record.getInt32(28, true) - originY, destinationWidth, destinationHeight);
  return true;
}

/** Rasterize the bitmap and Unicode caption used by simple Office attachment icons. */
export function renderDocxAttachmentEmf(
  bytes: Uint8Array,
  document: Document,
): { png: Uint8Array; caption: string } | null {
  if (bytes.length < 88 || bytes.length > MAX_EMF_BYTES) return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (view.getUint32(0, true) !== 1 || view.getUint32(40, true) !== 0x464d4520) return null;
  const x = view.getInt32(8, true);
  const y = view.getInt32(12, true);
  const width = view.getInt32(16, true) - x + 1;
  const height = view.getInt32(20, true) - y + 1;
  if (width <= 0 || width > MAX_ICON_SIZE || height <= 0 || height > MAX_ICON_SIZE) return null;
  const canvas = document.createElement("canvas");
  // Render at twice the source size so the recovered caption stays crisp when zoomed.
  canvas.width = 2 * width;
  canvas.height = 2 * height;
  const context = canvas.getContext("2d");
  if (context === null) return null;
  context.scale(2, 2);
  let bitmap = false;
  let align = 0;
  let fontSize = 16;
  let fontWeight = 400;
  let color = "#000";
  let caption = "";
  let endOfFile = false;
  for (let offset = 0; offset + 8 <= bytes.length;) {
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (size < 8 || offset + size > bytes.length) return null;
    const record = new DataView(bytes.buffer, bytes.byteOffset + offset, size);
    offset += size;
    if (type === 14) { endOfFile = true; break; }
    if (type === 114) {
      if (!drawAlphaBitmap(context, record, x, y)) return null;
      bitmap = true;
    } else if (type === 22 && size >= 12) {
      align = record.getUint32(8, true);
    } else if (type === 24 && size >= 12) {
      color = `rgb(${record.getUint8(8)}, ${record.getUint8(9)}, ${record.getUint8(10)})`;
    } else if (type === 82 && size >= 32) {
      fontSize = Math.max(8, Math.min(72, Math.abs(record.getInt32(12, true))));
      fontWeight = record.getInt32(28, true) >= 600 ? 700 : 400;
    } else if (type === 84 && size >= 76) {
      const count = record.getUint32(44, true);
      const start = record.getUint32(48, true);
      const options = record.getUint32(52, true);
      if (count > 256 || start < 76 || start + count * 2 > size || (options & 0x10) !== 0) return null;
      const text = new TextDecoder("utf-16le").decode(new Uint8Array(
        bytes.buffer, record.byteOffset + start, count * 2,
      )).replace(/[\p{C}]/gu, "");
      if (caption.length + text.length > 512) return null;
      caption += text;
      context.save();
      if ((options & 4) !== 0) {
        const left = record.getInt32(56, true);
        const top = record.getInt32(60, true);
        context.beginPath();
        context.rect(left - x, top - y, record.getInt32(64, true) - left, record.getInt32(68, true) - top);
        context.clip();
      }
      context.textAlign = (align & 6) === 6 ? "center" : (align & 2) !== 0 ? "right" : "left";
      context.textBaseline = (align & 24) === 24 ? "alphabetic" : (align & 8) !== 0 ? "bottom" : "top";
      context.fillStyle = color;
      context.font = `${fontWeight} ${fontSize}px "PingFang SC", sans-serif`;
      context.fillText(text, record.getInt32(36, true) - x, record.getInt32(40, true) - y);
      context.restore();
    } else if (![1, 9, 10, 18, 25, 37].includes(type)) {
      return null;
    }
  }
  if (!bitmap || !endOfFile) return null;
  const png = pngBytesFromDataUrl(canvas.toDataURL("image/png"));
  return png === null ? null : { png, caption: caption.trim() };
}
