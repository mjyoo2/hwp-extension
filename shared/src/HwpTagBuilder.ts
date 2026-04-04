/**
 * HWP Tag Builder - Binary writing utilities for HWP 5.0 format
 * Reverse of the reading functions in HwpParser.standalone.ts
 */

// ============================================================
// HWP Tags (mirrored from HwpParser.standalone.ts)
// ============================================================

export const HWP_TAGS = {
  HWPTAG_DOCUMENT_PROPERTIES: 16,
  HWPTAG_ID_MAPPINGS: 17,
  HWPTAG_BIN_DATA: 18,
  HWPTAG_FACE_NAME: 19,
  HWPTAG_BORDER_FILL: 20,
  HWPTAG_CHAR_SHAPE: 21,
  HWPTAG_TAB_DEF: 22,
  HWPTAG_NUMBERING: 23,
  HWPTAG_BULLET: 24,
  HWPTAG_PARA_SHAPE: 25,
  HWPTAG_STYLE: 26,
  HWPTAG_PARA_HEADER: 66,
  HWPTAG_PARA_TEXT: 67,
  HWPTAG_PARA_CHAR_SHAPE: 68,
  HWPTAG_PARA_LINE_SEG: 69,
  HWPTAG_PARA_RANGE_TAG: 70,
  HWPTAG_CTRL_HEADER: 71,
  HWPTAG_LIST_HEADER: 72,
  HWPTAG_PAGE_DEF: 73,
  HWPTAG_FOOTNOTE_SHAPE: 74,
  HWPTAG_PAGE_BORDER_FILL: 75,
  HWPTAG_SHAPE_COMPONENT: 76,
  HWPTAG_TABLE: 77,
  HWPTAG_SHAPE_COMPONENT_LINE: 78,
  HWPTAG_SHAPE_COMPONENT_RECTANGLE: 79,
  HWPTAG_SHAPE_COMPONENT_ELLIPSE: 80,
  HWPTAG_SHAPE_COMPONENT_ARC: 81,
  HWPTAG_SHAPE_COMPONENT_POLYGON: 82,
  HWPTAG_SHAPE_COMPONENT_CURVE: 83,
  HWPTAG_SHAPE_COMPONENT_OLE: 84,
  HWPTAG_SHAPE_COMPONENT_PICTURE: 85,
  HWPTAG_SHAPE_COMPONENT_CONTAINER: 86,
  HWPTAG_SHAPE_COMPONENT_TEXTBOX: 87,
  HWPTAG_SHAPE_COMPONENT_FORM_OBJECT: 88,
  HWPTAG_MEMO_SHAPE: 92,
  HWPTAG_MEMO_LIST: 93,
} as const;

export const CTRL_ID = {
  TABLE: 0x74626C20,       // 'tbl '
  PICTURE: 0x24706963,     // '$pic'
  SECTION: 0x73656364,     // 'secd'
  COLUMN: 0x636F6C64,      // 'cold'
  FORM: 0x666F726D,        // 'form'
  GSO: 0x67736F20,         // 'gso '
  FOOTER: 0x666F6F74,      // 'foot'
  HEADER: 0x68656164,      // 'head'
  FOOTNOTE: 0x666E2020,    // 'fn  '
  ENDNOTE: 0x656E2020,     // 'en  '
  FIELD_MEMO: 0x2466300c,
  AUTO_NUMBER: 0x61746E6F,
  PAGE_NUMBER_POS: 0x70676E70,
  EQUATION: 0x65716564,
  BOOKMARK: 0x626F6F6B,
  TCPS: 0x74637073,
  LINE: 0x246C696E,
  RECTANGLE: 0x24726563,
  ELLIPSE: 0x24656C6C,
  ARC: 0x24617263,
  POLYGON: 0x24706F6C,
  CURVE: 0x24637276,
  CONTAINER: 0x2463746E,
  OLE: 0x246F6C65,
  TEXTBOX: 0x24747874,
} as const;

export const CTRL_CHAR = {
  SPACE: 0x0020,
  TAB: 0x0009,
  PAGE_BREAK: 0x000C,
  LINE_BREAK: 0x000A,
  SOFT_LINE_BREAK: 0x001B,
  TABLE_DRAWING: 0x0007,
  EXTENDED: 0x0009,
  FIELD_START: 0x0013,
  FIELD_END: 0x0014,
  INLINE: 0x0015,
  BOOKMARK_START: 0x0001,
  BOOKMARK_END: 0x0002,
  SHAPE_DRAWING: 0x0006,
  SECTION_COLUMN_DEF: 0x0003,
} as const;

// ============================================================
// Exported interfaces (mirrored from HwpParser.standalone.ts)
// ============================================================

export interface BinDataInfo {
  id: number;
  type: 'LINK' | 'EMBEDDING' | 'STORAGE';
  extension: string;
}

export interface ParsedFaceName {
  name: string;
  hasSubstitute: boolean;
  hasFontTypeInfo: boolean;
  hasDefaultFont: boolean;
  substitute?: { type: 'unknown' | 'truetype' | 'hwp'; name: string };
  defaultFont?: string;
}

export interface ParsedCharShape {
  fontIds: number[];
  widthRatios: number[];
  spacings: number[];
  relativeSizes: number[];
  charPositions: number[];
  baseSize: number;
  italic: boolean;
  bold: boolean;
  underlineType: number;
  underlineShape: number;
  outlineType: number;
  shadowType: number;
  emboss: boolean;
  engrave: boolean;
  superscript: boolean;
  subscript: boolean;
  strikethrough: number;
  emphasisMark: number;
  useFontSpacing: boolean;
  strikethroughShape: number;
  kerning: boolean;
  shadowOffsetX: number;
  shadowOffsetY: number;
  textColor: number;
  underlineColor: number;
  shadeColor: number;
  shadowColor: number;
  borderFillId?: number;
  strikethroughColor?: number;
}

export interface ParsedParaShape {
  alignment: number;
  leftMargin: number;
  rightMargin: number;
  indent: number;
  spacingBefore: number;
  spacingAfter: number;
  lineSpacing: number;
  lineSpacingType: number;
  tabDefId: number;
  numberingId: number;
  borderFillId: number;
  borderSpacing: { left: number; right: number; top: number; bottom: number };
  wordBreakEnglish: number;
  wordBreakKorean: number;
  widowOrphan: boolean;
  keepWithNext: boolean;
  keepTogether: boolean;
  pageBreakBefore: boolean;
  verticalAlign: number;
  headType: number;
  level: number;
  useGrid: boolean;
  minSpace: number;
  autoSpaceKoreanEnglish: boolean;
  autoSpaceKoreanNumber: boolean;
}

export interface ParsedBorderLine {
  type: number;
  width: number;
  color: number;
}

export interface ParsedSolidFill {
  fillType: 'solid';
  backgroundColor: number;
  patternColor: number;
  patternType: number;
}

export interface ParsedGradientFill {
  fillType: 'gradient';
  gradientType: number;
  angle: number;
  centerX: number;
  centerY: number;
  blur: number;
  colors: Array<{ position?: number; color: number }>;
}

export interface ParsedImageFill {
  fillType: 'image';
  imageType: number;
  brightness: number;
  contrast: number;
  effect: number;
  binItemId: number;
}

export interface ParsedBorderFill {
  effect3d: boolean;
  shadow: boolean;
  slashDiagonal: number;
  backslashDiagonal: number;
  borders: {
    left: ParsedBorderLine;
    right: ParsedBorderLine;
    top: ParsedBorderLine;
    bottom: ParsedBorderLine;
  };
  diagonal?: {
    type: number;
    width: number;
    color: number;
  };
  fill?: ParsedSolidFill | ParsedGradientFill | ParsedImageFill;
}

export interface CharShapePosition {
  startPos: number;
  charShapeId: number;
}

export interface ParsedLineSeg {
  textStartPos: number;
  verticalPos: number;
  lineHeight: number;
  textHeight: number;
  baselineDistance: number;
  lineSpacing: number;
  horizontalStart: number;
  segmentWidth: number;
  flags: {
    isPageFirst: boolean;
    isColumnFirst: boolean;
    isEmpty: boolean;
    isLastInPara: boolean;
    isAutoHyphen: boolean;
    isIndent: boolean;
  };
}

export interface ParsedRangeTag {
  start: number;
  end: number;
  type: number;
  data: number;
}

export interface ParsedHeaderFooter {
  type: 'header' | 'footer';
  applyTo: 'both' | 'even' | 'odd';
  textWidth: number;
  textHeight: number;
  paragraphs: import('./types').HwpxParagraph[];
}

export interface ParsedFootnoteEndnote {
  type: 'footnote' | 'endnote';
  number: number;
  paragraphWidth: number;
  paragraphs: import('./types').HwpxParagraph[];
}

export interface ParsedSectionDef {
  hideHeader: boolean;
  hideFooter: boolean;
  hideMasterPage: boolean;
  hideBorder: boolean;
  hideBackground: boolean;
  hidePageNum: boolean;
  borderOnFirstOnly: boolean;
  backgroundOnFirstOnly: boolean;
  textDirection: number;
  columnGap: number;
  pageNumber: number;
  figureNumber: number;
  tableNumber: number;
  equationNumber: number;
}

export interface ParsedColumnDef {
  columnType: 'normal' | 'distribute' | 'parallel';
  columnCount: number;
  direction: 'left' | 'right' | 'facing';
  sameWidth: boolean;
  gap: number;
}

// ============================================================
// Byte Writing Utilities (reverse of read functions)
// ============================================================

export function writeUint8(value: number): Uint8Array {
  return new Uint8Array([value & 0xFF]);
}

export function writeUint16LE(value: number): Uint8Array {
  const buf = new Uint8Array(2);
  buf[0] = value & 0xFF;
  buf[1] = (value >>> 8) & 0xFF;
  return buf;
}

export function writeUint32LE(value: number): Uint8Array {
  const buf = new Uint8Array(4);
  buf[0] = value & 0xFF;
  buf[1] = (value >>> 8) & 0xFF;
  buf[2] = (value >>> 16) & 0xFF;
  buf[3] = (value >>> 24) & 0xFF;
  return buf;
}

export function writeInt16LE(value: number): Uint8Array {
  if (value < 0) value = value + 0x10000;
  return writeUint16LE(value);
}

export function writeInt32LE(value: number): Uint8Array {
  if (value < 0) value = value + 0x100000000;
  return writeUint32LE(value);
}

export function writeUtf16LE(str: string): Uint8Array {
  const buf = new Uint8Array(str.length * 2);
  for (let i = 0; i < str.length; i++) {
    const code = str.charCodeAt(i);
    buf[i * 2] = code & 0xFF;
    buf[i * 2 + 1] = (code >>> 8) & 0xFF;
  }
  return buf;
}

/**
 * Reverse of colorrefToHex: "#RRGGBB" → COLORREF integer (0x00BBGGRR)
 */
export function colorHexToRef(hex: string): number {
  const h = hex.replace('#', '');
  const r = parseInt(h.substring(0, 2), 16);
  const g = parseInt(h.substring(2, 4), 16);
  const b = parseInt(h.substring(4, 6), 16);
  return r | (g << 8) | (b << 16);
}

/**
 * Reverse of hwpunitToPt: pt → hwpunit
 * hwpunitToPt = hwpunit / 7200 * 72 = hwpunit / 100
 * So ptToHwpunit = pt * 100
 */
export function ptToHwpunit(pt: number): number {
  return Math.round(pt * 100);
}

/**
 * Reverse of uint8ArrayToBase64
 */
export function base64ToUint8Array(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

// ============================================================
// Tag Record Creation
// ============================================================

/**
 * Creates a tag record header + body.
 *
 * Tag header bit layout (from parser lines 1788-1790):
 *   tagId  = header & 0x3FF         (bits 0-9)
 *   level  = (header >>> 10) & 0x3FF (bits 10-19)
 *   size   = (header >>> 20) & 0xFFF (bits 20-31)
 *
 * When size >= 0xFFF, use extended format: header with size=0xFFF, then 4-byte actual size.
 */
export function createTagRecord(tagId: number, level: number, body: Uint8Array): Uint8Array {
  const size = body.length;

  if (size < 0xFFF) {
    // Normal format: 4-byte header + body
    const header = (tagId & 0x3FF) | ((level & 0x3FF) << 10) | ((size & 0xFFF) << 20);
    const result = new Uint8Array(4 + size);
    result[0] = header & 0xFF;
    result[1] = (header >>> 8) & 0xFF;
    result[2] = (header >>> 16) & 0xFF;
    result[3] = (header >>> 24) & 0xFF;
    result.set(body, 4);
    return result;
  } else {
    // Extended format: 4-byte header (size=0xFFF) + 4-byte actual size + body
    const header = (tagId & 0x3FF) | ((level & 0x3FF) << 10) | (0xFFF << 20);
    const result = new Uint8Array(8 + size);
    result[0] = header & 0xFF;
    result[1] = (header >>> 8) & 0xFF;
    result[2] = (header >>> 16) & 0xFF;
    result[3] = (header >>> 24) & 0xFF;
    // Actual size as uint32 LE
    result[4] = size & 0xFF;
    result[5] = (size >>> 8) & 0xFF;
    result[6] = (size >>> 16) & 0xFF;
    result[7] = (size >>> 24) & 0xFF;
    result.set(body, 8);
    return result;
  }
}

// ============================================================
// TagStreamBuilder - Accumulates tag records into a single stream
// ============================================================

export class TagStreamBuilder {
  private records: Uint8Array[] = [];
  private totalSize = 0;

  addRecord(tagId: number, level: number, body: Uint8Array): void {
    const record = createTagRecord(tagId, level, body);
    this.records.push(record);
    this.totalSize += record.length;
  }

  build(): Uint8Array {
    const result = new Uint8Array(this.totalSize);
    let offset = 0;
    for (const record of this.records) {
      result.set(record, offset);
      offset += record.length;
    }
    return result;
  }

  get size(): number {
    return this.totalSize;
  }

  get recordCount(): number {
    return this.records.length;
  }
}

// ============================================================
// BodyBuilder - Utility for building tag record bodies
// ============================================================

export class BodyBuilder {
  private parts: Uint8Array[] = [];
  private totalSize = 0;

  addUint8(value: number): this {
    const buf = writeUint8(value);
    this.parts.push(buf);
    this.totalSize += buf.length;
    return this;
  }

  addUint16(value: number): this {
    const buf = writeUint16LE(value);
    this.parts.push(buf);
    this.totalSize += buf.length;
    return this;
  }

  addUint32(value: number): this {
    const buf = writeUint32LE(value);
    this.parts.push(buf);
    this.totalSize += buf.length;
    return this;
  }

  addInt16(value: number): this {
    const buf = writeInt16LE(value);
    this.parts.push(buf);
    this.totalSize += buf.length;
    return this;
  }

  addInt32(value: number): this {
    const buf = writeInt32LE(value);
    this.parts.push(buf);
    this.totalSize += buf.length;
    return this;
  }

  addBytes(data: Uint8Array): this {
    this.parts.push(data);
    this.totalSize += data.length;
    return this;
  }

  addUtf16String(str: string): this {
    const buf = writeUtf16LE(str);
    this.parts.push(buf);
    this.totalSize += buf.length;
    return this;
  }

  /** Write a length-prefixed UTF-16LE string (uint16 length + UTF-16LE chars) */
  addHwpString(str: string): this {
    this.addUint16(str.length);
    this.addUtf16String(str);
    return this;
  }

  addZeros(count: number): this {
    const buf = new Uint8Array(count);
    this.parts.push(buf);
    this.totalSize += count;
    return this;
  }

  build(): Uint8Array {
    const result = new Uint8Array(this.totalSize);
    let offset = 0;
    for (const part of this.parts) {
      result.set(part, offset);
      offset += part.length;
    }
    return result;
  }

  get size(): number {
    return this.totalSize;
  }
}

// ============================================================
// Helper to concatenate multiple Uint8Arrays
// ============================================================

export function concatUint8Arrays(...arrays: Uint8Array[]): Uint8Array {
  let totalLength = 0;
  for (const arr of arrays) {
    totalLength += arr.length;
  }
  const result = new Uint8Array(totalLength);
  let offset = 0;
  for (const arr of arrays) {
    result.set(arr, offset);
    offset += arr.length;
  }
  return result;
}
