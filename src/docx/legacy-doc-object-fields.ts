// MS-DOC FibRgFcLcb97 indexes for PlcBteChpx, PlcffldMom, and Clx.
const CHARACTER_FORMAT_TABLE = 12;
const MAIN_FIELD_TABLE = 16;
const PIECE_TABLE = 33;
const SPRM_PICTURE_LOCATION = 0x6a03;
const SPRM_OLE_OBJECT = 0x080a;
const SPRM_SPECIAL_CHARACTER = 0x0855;
const SPRM_FIELD_OBJECT = 0x0856;

export interface LegacyDocObjectField {
  readonly storage: string;
  readonly previewName: string | undefined;
}

function slice(bytes: Buffer, offset: number, size: number): Buffer {
  if (offset < 0 || size < 0 || offset + size > bytes.length) throw new RangeError("Invalid Word record");
  return bytes.subarray(offset, offset + size);
}

function countEntries(bytes: Buffer, recordSize: number): number {
  if (bytes.length === 0) return 0;
  const count = (bytes.length - 4) / (recordSize + 4);
  if (!Number.isInteger(count) || count < 0) throw new RangeError("Invalid Word PLC");
  return count;
}

function interval(bytes: Buffer, count: number, position: number): number {
  let lower = 0;
  let upper = count;
  while (lower < upper) {
    const middle = Math.floor((lower + upper) / 2);
    if (bytes.readUInt32LE((middle + 1) * 4) <= position) lower = middle + 1;
    else upper = middle;
  }
  if (lower === count || bytes.readUInt32LE(lower * 4) > position) {
    throw new RangeError("Word position outside its table");
  }
  return lower;
}

function pictureShapeName(data: Buffer, location: number): string | undefined {
  const picture = slice(data, location, data.readUInt32LE(location));
  const headerSize = picture.readUInt16LE(4);
  // An inline picture begins with an OfficeArtSpContainer.
  if (headerSize < 6 || picture.readUInt16LE(headerSize + 2) !== 0xf004) return;
  const end = headerSize + 8 + picture.readUInt32LE(headerSize + 4);
  if (end > picture.length) return;
  for (let offset = headerSize + 8; offset + 8 <= end;) {
    const record = slice(picture, offset + 8, picture.readUInt32LE(offset + 4));
    const type = picture.readUInt16LE(offset + 2);
    if (offset + 8 + record.length > end) return;
    if (type === 0xf00b || type === 0xf121 || type === 0xf122) {
      const propertyCount = picture.readUInt16LE(offset) >>> 4;
      let complexOffset = propertyCount * 6;
      for (let index = 0; index < propertyCount; index++) {
        const property = record.readUInt16LE(index * 6);
        if ((property & 0x8000) === 0) continue;
        const length = record.readUInt32LE(index * 6 + 2);
        if ((property & 0x3fff) === 0x0380 && length <= 1024 && length % 2 === 0) {
          // MS-ODRAW wzName. LibreOffice preserves it as wp:docPr/@name.
          const name = slice(record, complexOffset, length).toString("utf16le").split("\0", 1)[0];
          if (name) return name;
        }
        complexOffset += length;
      }
    }
    offset += 8 + record.length;
  }
}

export function readLegacyDocObjectFields(
  document: Buffer,
  table: Buffer,
  data: Buffer,
): LegacyDocObjectField[] {
  const longCountOffset = 34 + 2 * document.readUInt16LE(32);
  const pairCountOffset = longCountOffset + 2 + 4 * document.readUInt16LE(longCountOffset);
  const pairCount = document.readUInt16LE(pairCountOffset);
  const fibPart = (index: number) => {
    if (index >= pairCount) throw new RangeError("Missing Word FIB field");
    const offset = pairCountOffset + 2 + 8 * index;
    return slice(table, document.readUInt32LE(offset), document.readUInt32LE(offset + 4));
  };
  const fields = fibPart(MAIN_FIELD_TABLE);
  const fieldCount = countEntries(fields, 2);
  if (fieldCount === 0) return [];

  const clx = fibPart(PIECE_TABLE);
  let pieceOffset = 0;
  while (clx[pieceOffset] === 1) pieceOffset += 3 + clx.readUInt16LE(pieceOffset + 1);
  if (clx[pieceOffset] !== 2) return [];
  const pieces = slice(clx, pieceOffset + 5, clx.readUInt32LE(pieceOffset + 1));
  const pieceCount = countEntries(pieces, 8);
  const pages = fibPart(CHARACTER_FORMAT_TABLE);
  const pageCount = countEntries(pages, 4);

  const character = (cp: number) => {
    const index = interval(pieces, pieceCount, cp);
    const encodedOffset = pieces.readUInt32LE(4 * (pieceCount + 1) + 8 * index + 2);
    const compressed = (encodedOffset & 0x40000000) !== 0;
    const distance = cp - pieces.readUInt32LE(4 * index);
    const base = encodedOffset & 0x3fffffff;
    const offset = compressed ? base / 2 + distance : base + 2 * distance;
    return { offset, code: compressed ? document.readUInt8(offset) : document.readUInt16LE(offset) };
  };

  const properties = (offset: number) => {
    const page = interval(pages, pageCount, offset);
    const pageNumber = pages.readUInt32LE(4 * (pageCount + 1) + 4 * page) & 0x3fffff;
    const fkp = slice(document, pageNumber * 512, 512);
    const runCount = fkp.readUInt8(511);
    if (runCount < 1 || runCount > 101) throw new RangeError("Invalid Word CHPX run count");
    const index = interval(fkp, runCount, offset);
    const start = fkp.readUInt8(4 * (runCount + 1) + index) * 2;
    const result = new Map<number, number>();
    if (start === 0) return result;
    const sprms = slice(fkp, start + 1, fkp.readUInt8(start));
    for (let position = 0; position + 2 <= sprms.length;) {
      const opcode = sprms.readUInt16LE(position);
      position += 2;
      const operandSize = [1, 1, 2, 4, 2, 2, 0, 3][opcode >>> 13]!;
      if (operandSize === 0) {
        position += 1 + sprms.readUInt8(position);
        continue;
      }
      if (opcode === SPRM_PICTURE_LOCATION) result.set(opcode, sprms.readUInt32LE(position));
      else if (operandSize === 1) result.set(opcode, sprms.readUInt8(position));
      position += operandSize;
    }
    return result;
  };

  const stack: Array<{ type: number; separator?: number }> = [];
  const objects: LegacyDocObjectField[] = [];
  const recordOffset = 4 * (fieldCount + 1);
  for (let index = 0; index < fieldCount && objects.length < 128; index++) {
    const kind = fields.readUInt8(recordOffset + 2 * index) & 0x1f;
    const flags = fields.readUInt8(recordOffset + 2 * index + 1);
    const position = fields.readUInt32LE(4 * index);
    if (kind === 0x13) {
      stack.push({ type: flags });
    } else if (kind === 0x14) {
      const field = stack.at(-1);
      if (field !== undefined) field.separator = position;
    } else if (kind === 0x15) {
      const field = stack.pop();
      // EMBED, excluding zombie fields and link/control field types.
      if (field?.type !== 0x3a || field.separator === undefined || (flags & 0x22) !== 0) continue;
      try {
        const separator = character(field.separator);
        if (separator.code !== 0x14) continue;
        const attrs = properties(separator.offset);
        const storageId = attrs.get(SPRM_PICTURE_LOCATION);
        if (storageId === undefined || attrs.get(SPRM_OLE_OBJECT) !== 1 ||
          attrs.get(SPRM_FIELD_OBJECT) !== 1 || attrs.get(SPRM_SPECIAL_CHARACTER) !== 1) continue;
        let previewName: string | undefined;
        for (let cp = field.separator + 1; cp < position && cp <= field.separator + 32; cp++) {
          const preview = character(cp);
          if (preview.code !== 1) continue;
          const location = properties(preview.offset).get(SPRM_PICTURE_LOCATION);
          if (location === undefined) continue;
          previewName = pictureShapeName(data, location);
          if (previewName !== undefined) break;
        }
        objects.push({ storage: `_${storageId}`, previewName });
      } catch {
        // A damaged field cannot identify an attachment.
      }
    }
  }
  return objects;
}
