import * as wmfModule from "wmf";

const wmf = (wmfModule as { default?: typeof wmfModule }).default ?? wmfModule;

const PLACEABLE_WMF_MAGIC = 0x9a_c6_cd_d7;

export function isPlaceableOrRawWmf(bytes: Uint8Array): boolean {
  if (bytes.byteLength < 28) return false;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const magic = view.getUint32(0, true);
  if (magic === PLACEABLE_WMF_MAGIC) {
    const type = view.getUint16(22, true);
    return type === 1 || type === 2;
  }
  const type = view.getUint16(0, true);
  return type === 1 || type === 2;
}

function rawWmfPayload(bytes: Uint8Array): Uint8Array {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (
    bytes.byteLength >= 22 &&
    view.getUint32(0, true) === PLACEABLE_WMF_MAGIC
  ) {
    return bytes.subarray(22);
  }
  return bytes;
}

function decodePossiblyGbk(value: string): string {
  if (![...value].some((char) => char.charCodeAt(0) > 127)) return value;
  const bytes = Uint8Array.from(value, (char) => char.charCodeAt(0) & 0xff);
  try {
    const decoded = new TextDecoder("gbk").decode(bytes);
    return /[\u4e00-\u9fff]/.test(decoded) ? decoded : value;
  } catch {
    return value;
  }
}

export function pngBytesFromDataUrl(dataUrl: string): Uint8Array | null {
  const marker = "base64,";
  const index = dataUrl.indexOf(marker);
  if (index < 0) return null;
  const binary = atob(dataUrl.slice(index + marker.length));
  const bytes = new Uint8Array(binary.length);
  for (let offset = 0; offset < binary.length; offset += 1) {
    bytes[offset] = binary.charCodeAt(offset);
  }
  return bytes;
}

/** Decode WMF drawing actions and re-decode high-byte text as GBK when needed. */
export function getWindowsMetafileActions(
  bytes: Uint8Array,
): ReturnType<typeof wmf.get_actions> | null {
  if (!isPlaceableOrRawWmf(bytes)) return null;
  try {
    const payload = rawWmfPayload(bytes);
    const actions = wmf.get_actions(payload);
    for (const action of actions) {
      if (action.t !== "text" || typeof action.v !== "string") continue;
      action.v = decodePossiblyGbk(action.v);
    }
    return actions;
  } catch {
    return null;
  }
}

/**
 * Convert a Windows placeable/raw WMF image to PNG using a DOM canvas.
 * Returns null when the payload is not WMF or conversion is unavailable.
 */
export function convertWindowsMetafileToPng(
  bytes: Uint8Array,
): Uint8Array | null {
  if (typeof document === "undefined" || !isPlaceableOrRawWmf(bytes)) {
    return null;
  }
  try {
    const payload = rawWmfPayload(bytes);
    const actions = getWindowsMetafileActions(bytes);
    if (actions === null) return null;
    const size = wmf.image_size(payload);
    const width = Math.max(1, Math.ceil(Math.abs(size[0]) || 1));
    const height = Math.max(1, Math.ceil(Math.abs(size[1]) || 1));
    const canvas = createEl("canvas");
    canvas.width = width;
    canvas.height = height;
    wmf.render_canvas(actions, canvas);
    return pngBytesFromDataUrl(canvas.toDataURL("image/png"));
  } catch {
    return null;
  }
}
