/**
 * HWP Standalone Parser - Binary HWP format handler without VSCode dependencies
 * Extracted from HwpDocument.ts for standalone Node.js usage
 * Based on HWP 5.0 specification (한글문서파일형식_5.0_revision1.3)
 */

import * as CFB from 'cfb';
import * as pako from 'pako';
import {
  HwpxContent,
  HwpxSection,
  HwpxParagraph,
  TextRun,
  HwpxImage,
  HwpxEquation,
  HwpxTable,
  TableRow,
  TableCell,
  CharacterStyle,
  ParagraphStyle,
  GradationType,
} from './types';

// ============================================================
// Helper Interfaces
// ============================================================

interface BinDataInfo {
  id: number;
  type: 'LINK' | 'EMBEDDING' | 'STORAGE';
  extension: string;
}

interface ParsedFaceName {
  name: string;
  hasSubstitute: boolean;
  hasFontTypeInfo: boolean;
  hasDefaultFont: boolean;
  substitute?: { type: 'unknown' | 'truetype' | 'hwp'; name: string };
  defaultFont?: string;
}

interface ParsedCharShape {
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

interface ParsedParaShape {
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

interface CharShapePosition {
  startPos: number;
  charShapeId: number;
}

interface ParsedLineSeg {
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

interface ParsedRangeTag {
  start: number;
  end: number;
  type: number;
  data: number;
}

interface ParsedBorderLine {
  type: number;
  width: number;
  color: number;
}

interface ParsedSolidFill {
  fillType: 'solid';
  backgroundColor: number;
  patternColor: number;
  patternType: number;
}

interface ParsedGradientFill {
  fillType: 'gradient';
  gradientType: number;
  angle: number;
  centerX: number;
  centerY: number;
  blur: number;
  colors: Array<{ position?: number; color: number }>;
}

interface ParsedImageFill {
  fillType: 'image';
  imageType: number;
  brightness: number;
  contrast: number;
  effect: number;
  binItemId: number;
}

interface ParsedBorderFill {
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

interface ParsedHeaderFooter {
  type: 'header' | 'footer';
  applyTo: 'both' | 'even' | 'odd';
  textWidth: number;
  textHeight: number;
  paragraphs: HwpxParagraph[];
}

interface ParsedFootnoteEndnote {
  type: 'footnote' | 'endnote';
  number: number;
  paragraphWidth: number;
  paragraphs: HwpxParagraph[];
}

interface ParsedSectionDef {
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

interface ParsedColumnDef {
  columnType: 'normal' | 'distribute' | 'parallel';
  columnCount: number;
  direction: 'left' | 'right' | 'facing';
  sameWidth: boolean;
  gap: number;
}

interface ParsedField {
  type: string;
  numberType?: number;
  numberShape?: number;
  number?: number;
  isSuperscript?: boolean;
  position?: number;
  name?: string;
  keyword1?: string;
  keyword2?: string;
  command?: string;
  properties?: number;
  readOnlyEditable?: boolean;
  hyperlinkUpdateType?: number;
  modified?: boolean;
  etcProperties?: number;
}

interface ParsedMemoShape {
  memoId: number;
  width: number;
  height: number;
  lineType: number;
  lineColor: string;
  fillColor: string;
  activeColor: number;
}

interface PendingTextSegment {
  start: number;
  end: number;
  text: string;
}

interface ParseContext {
  currentParagraph: HwpxParagraph | null;
  currentCharShapeId: number;
  currentParaShapeId: number;
  currentStyleId: number;
  charShapePositions: CharShapePosition[];
  pendingTextSegments: PendingTextSegment[];
  pendingLineSegs: ParsedLineSeg[];
  pendingRangeTags: ParsedRangeTag[];
  currentTable: HwpxTable | null;
  currentTableRow: number;
  currentTableCol: number;
  currentCtrlId: number;
  inTableCell: boolean;
  cellParagraphs: HwpxParagraph[];
  tableRowCount: number;
  tableColCount: number;
  tableCells: TableCell[][];
  pendingImage: { width: number; height: number } | null;
  pendingSectionImages: HwpxImage[];
  faceNames: Map<number, ParsedFaceName>;
  charShapes: Map<number, ParsedCharShape>;
  paraShapes: Map<number, ParsedParaShape>;
  borderFills: Map<number, ParsedBorderFill>;
  inHeaderFooter: boolean;
  currentHeaderFooter: ParsedHeaderFooter | null;
  headerFooterParagraphs: HwpxParagraph[];
  inFootnoteEndnote: boolean;
  currentFootnoteEndnote: ParsedFootnoteEndnote | null;
  footnoteEndnoteParagraphs: HwpxParagraph[];
  currentSectionDef: ParsedSectionDef | null;
  currentColumnDef: ParsedColumnDef | null;
  nestedLevel: number;
  pendingShape: any | null;
  pendingField: ParsedField | null;
  inMemo: boolean;
  currentMemo: ParsedMemoShape | null;
  memoParagraphs: HwpxParagraph[];
  memos: ParsedMemoShape[];
  inShapeText: boolean;
  shapeTextParagraphs: HwpxParagraph[];
  inEquation: boolean;
  equationScript: string;
}

// ============================================================
// HWP Tags
// ============================================================

const HWP_TAGS = {
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
};

const CTRL_ID = {
  TABLE: 0x74626C20,    // 'tbl ' in ASCII
  PICTURE: 0x24706963,  // '$pic' in ASCII
  SECTION: 0x73656364,  // 'secd' in ASCII
  COLUMN: 0x636F6C64,   // 'cold' in ASCII
  FORM: 0x666F726D,     // 'form' in ASCII
  GSO: 0x67736F20,      // 'gso ' in ASCII
  FOOTER: 0x666F6F74,   // 'foot' in ASCII
  HEADER: 0x68656164,   // 'head' in ASCII
  FOOTNOTE: 0x666E2020, // 'fn  ' in ASCII
  ENDNOTE: 0x656E2020,  // 'en  ' in ASCII
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
};

const CTRL_CHAR = {
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
};

// ============================================================
// Helper Functions
// ============================================================

function generateId(): string {
  return Math.random().toString(36).substring(2, 11);
}

function readUint16(data: Uint8Array, offset: number): number {
  return data[offset] | (data[offset + 1] << 8);
}

function readUint32(data: Uint8Array, offset: number): number {
  return data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24);
}

function readInt32(data: Uint8Array, offset: number): number {
  const val = readUint32(data, offset);
  return val > 0x7FFFFFFF ? val - 0x100000000 : val;
}

function readInt8(data: Uint8Array, offset: number): number {
  const val = data[offset];
  return val > 127 ? val - 256 : val;
}

function readInt16(data: Uint8Array, offset: number): number {
  const val = readUint16(data, offset);
  return val > 0x7FFF ? val - 0x10000 : val;
}

function colorrefToHex(colorref: number): string {
  const r = colorref & 0xFF;
  const g = (colorref >> 8) & 0xFF;
  const b = (colorref >> 16) & 0xFF;
  return `#${r.toString(16).padStart(2, '0').toUpperCase()}${g.toString(16).padStart(2, '0').toUpperCase()}${b.toString(16).padStart(2, '0').toUpperCase()}`;
}

function hwpunitToPt(hwpunit: number): number {
  return hwpunit / 7200 * 72;
}

function charShapeToStyleStandalone(charShape: ParsedCharShape, faceNames: Map<number, ParsedFaceName>): CharacterStyle {
  const fontId = charShape.fontIds[0];
  const faceName = faceNames.get(fontId);
  const langKeys = ['hangul', 'latin', 'hanja', 'japanese', 'other', 'symbol', 'user'] as const;

  const charSpacing: Record<string, number> = {};
  const charOffset: Record<string, number> = {};
  for (let i = 0; i < 7; i++) {
    charSpacing[langKeys[i]] = charShape.spacings[i];
    charOffset[langKeys[i]] = charShape.charPositions[i];
  }

  return {
    fontName: faceName?.name,
    fontSize: charShape.baseSize / 100,
    bold: charShape.bold,
    italic: charShape.italic,
    underline: charShape.underlineType === 1 || charShape.underlineType === 3,
    strikethrough: charShape.strikethrough >= 2,
    fontColor: colorrefToHex(charShape.textColor),
    superscript: charShape.superscript,
    subscript: charShape.subscript,
    emboss: charShape.emboss,
    engrave: charShape.engrave,
    useFontSpace: charShape.useFontSpacing,
    useKerning: charShape.kerning,
    charSpacing,
    charOffset,
  };
}

function uint8ArrayToBase64(bytes: Uint8Array): string {
  const chunkSize = 0x8000;
  const chunks: string[] = [];
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    chunks.push(String.fromCharCode.apply(null, chunk as unknown as number[]));
  }
  return btoa(chunks.join(''));
}

function parseFaceNameStandalone(data: Uint8Array): ParsedFaceName | null {
  if (data.length < 3) return null;

  const props = data[0];
  const hasSubstitute = (props & 0x80) !== 0;
  const hasFontTypeInfo = (props & 0x40) !== 0;
  const hasDefaultFont = (props & 0x20) !== 0;

  const nameLen = readUint16(data, 1);
  if (data.length < 3 + nameLen * 2) return null;

  const name = new TextDecoder('utf-16le').decode(data.slice(3, 3 + nameLen * 2));

  const result: ParsedFaceName = {
    name,
    hasSubstitute,
    hasFontTypeInfo,
    hasDefaultFont,
  };

  let pos = 3 + nameLen * 2;

  if (hasSubstitute && pos + 3 <= data.length) {
    const substType = data[pos];
    pos++;
    const substNameLen = readUint16(data, pos);
    pos += 2;
    if (pos + substNameLen * 2 <= data.length) {
      const substName = new TextDecoder('utf-16le').decode(data.slice(pos, pos + substNameLen * 2));
      result.substitute = {
        type: substType === 1 ? 'truetype' : substType === 2 ? 'hwp' : 'unknown',
        name: substName,
      };
      pos += substNameLen * 2;
    }
  }

  if (hasFontTypeInfo && pos + 10 <= data.length) {
    pos += 10;
  }

  if (hasDefaultFont && pos + 2 <= data.length) {
    const defaultNameLen = readUint16(data, pos);
    pos += 2;
    if (pos + defaultNameLen * 2 <= data.length) {
      result.defaultFont = new TextDecoder('utf-16le').decode(data.slice(pos, pos + defaultNameLen * 2));
    }
  }

  return result;
}

function parseCharShapeStandalone(data: Uint8Array): ParsedCharShape | null {
  if (data.length < 72) return null;

  const fontIds: number[] = [];
  for (let i = 0; i < 7; i++) {
    fontIds.push(readUint16(data, i * 2));
  }

  const widthRatios: number[] = [];
  for (let i = 0; i < 7; i++) {
    widthRatios.push(data[14 + i]);
  }

  const spacings: number[] = [];
  for (let i = 0; i < 7; i++) {
    spacings.push(readInt8(data, 21 + i));
  }

  const relativeSizes: number[] = [];
  for (let i = 0; i < 7; i++) {
    relativeSizes.push(data[28 + i]);
  }

  const charPositions: number[] = [];
  for (let i = 0; i < 7; i++) {
    charPositions.push(readInt8(data, 35 + i));
  }

  const baseSize = readInt32(data, 42);
  const props = readUint32(data, 46);

  const italic = (props & 0x01) !== 0;
  const bold = (props & 0x02) !== 0;
  const underlineType = (props >> 2) & 0x03;
  const underlineShape = (props >> 4) & 0x0F;
  const outlineType = (props >> 8) & 0x07;
  const shadowType = (props >> 11) & 0x03;
  const emboss = (props & (1 << 13)) !== 0;
  const engrave = (props & (1 << 14)) !== 0;
  const superscript = (props & (1 << 15)) !== 0;
  const subscript = (props & (1 << 16)) !== 0;
  const strikethrough = (props >> 18) & 0x07;
  const emphasisMark = (props >> 21) & 0x0F;
  const useFontSpacing = (props & (1 << 25)) !== 0;
  const strikethroughShape = (props >> 26) & 0x0F;
  const kerning = (props & (1 << 30)) !== 0;

  const shadowOffsetX = readInt8(data, 50);
  const shadowOffsetY = readInt8(data, 51);
  const textColor = readUint32(data, 52);
  const underlineColor = readUint32(data, 56);
  const shadeColor = readUint32(data, 60);
  const shadowColor = readUint32(data, 64);

  const result: ParsedCharShape = {
    fontIds,
    widthRatios,
    spacings,
    relativeSizes,
    charPositions,
    baseSize,
    italic,
    bold,
    underlineType,
    underlineShape,
    outlineType,
    shadowType,
    emboss,
    engrave,
    superscript,
    subscript,
    strikethrough,
    emphasisMark,
    useFontSpacing,
    strikethroughShape,
    kerning,
    shadowOffsetX,
    shadowOffsetY,
    textColor,
    underlineColor,
    shadeColor,
    shadowColor,
  };

  if (data.length >= 70) {
    result.borderFillId = readUint16(data, 68);
  }
  if (data.length >= 74) {
    result.strikethroughColor = readUint32(data, 70);
  }

  return result;
}

function parseParaShapeStandalone(data: Uint8Array): ParsedParaShape | null {
  if (data.length < 42) return null;

  const props1 = readUint32(data, 0);

  const lineSpacingTypeOld = props1 & 0x03;
  const alignment = (props1 >> 2) & 0x07;
  const wordBreakEnglish = (props1 >> 5) & 0x03;
  const wordBreakKorean = (props1 >> 7) & 0x01;
  const useGrid = (props1 & (1 << 8)) !== 0;
  const minSpace = (props1 >> 9) & 0x7F;
  const widowOrphan = (props1 & (1 << 16)) !== 0;
  const keepWithNext = (props1 & (1 << 17)) !== 0;
  const keepTogether = (props1 & (1 << 18)) !== 0;
  const pageBreakBefore = (props1 & (1 << 19)) !== 0;
  const verticalAlign = (props1 >> 20) & 0x03;
  const headType = (props1 >> 23) & 0x03;
  const level = (props1 >> 25) & 0x07;

  const leftMargin = readInt32(data, 4);
  const rightMargin = readInt32(data, 8);
  const indent = readInt32(data, 12);
  const spacingBefore = readInt32(data, 16);
  const spacingAfter = readInt32(data, 20);
  let lineSpacing = readInt32(data, 24);
  const tabDefId = readUint16(data, 28);
  const numberingId = readUint16(data, 30);
  const borderFillId = readUint16(data, 32);
  const borderLeft = readInt16(data, 34);
  const borderRight = readInt16(data, 36);
  const borderTop = readInt16(data, 38);
  const borderBottom = readInt16(data, 40);

  let lineSpacingType = lineSpacingTypeOld;
  let autoSpaceKoreanEnglish = false;
  let autoSpaceKoreanNumber = false;

  if (data.length >= 46) {
    const props2 = readUint32(data, 42);
    autoSpaceKoreanEnglish = (props2 & (1 << 4)) !== 0;
    autoSpaceKoreanNumber = (props2 & (1 << 5)) !== 0;
  }

  if (data.length >= 54) {
    const props3 = readUint32(data, 46);
    lineSpacingType = props3 & 0x1F;
    lineSpacing = readUint32(data, 50);
  }

  return {
    alignment,
    leftMargin,
    rightMargin,
    indent,
    spacingBefore,
    spacingAfter,
    lineSpacing,
    lineSpacingType,
    tabDefId,
    numberingId,
    borderFillId,
    borderSpacing: { left: borderLeft, right: borderRight, top: borderTop, bottom: borderBottom },
    wordBreakEnglish,
    wordBreakKorean,
    widowOrphan,
    keepWithNext,
    keepTogether,
    pageBreakBefore,
    verticalAlign,
    headType,
    level,
    useGrid,
    minSpace,
    autoSpaceKoreanEnglish,
    autoSpaceKoreanNumber,
  };
}

function parseBorderFillStandalone(data: Uint8Array): ParsedBorderFill | null {
  if (data.length < 32) return null;

  const props = readUint16(data, 0);
  const effect3d = (props & 0x01) !== 0;
  const shadow = (props & 0x02) !== 0;
  const slashDiagonal = (props >> 2) & 0x07;
  const backslashDiagonal = (props >> 5) & 0x07;

  // HWP 5.0 spec: border data layout is NOT grouped per-border.
  // It's: 4 types (1 byte each), then 4 widths (1 byte each), then 4 colors (4 bytes each)
  // Order within each group: left, right, top, bottom
  // Offset 2: types[4], offset 6: widths[4], offset 10: colors[16]
  const borderWidthMm = [0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0];
  const borderTypes = [data[2], data[3], data[4], data[5]];       // left, right, top, bottom
  const borderWidths = [data[6], data[7], data[8], data[9]];      // left, right, top, bottom
  const borderColors = [
    readUint32(data, 10), readUint32(data, 14),                    // left, right
    readUint32(data, 18), readUint32(data, 22),                    // top, bottom
  ];

  const makeBorder = (i: number): ParsedBorderLine => ({
    type: borderTypes[i],
    width: borderWidthMm[borderWidths[i]] ?? 0.1,
    color: borderColors[i],
  });

  const leftBorder = makeBorder(0);
  const rightBorder = makeBorder(1);
  const topBorder = makeBorder(2);
  const bottomBorder = makeBorder(3);

  const diagonalType = data[26];
  const diagonalWidth = borderWidthMm[data[27]] ?? 0.1;
  const diagonalColor = readUint32(data, 28);

  const result: ParsedBorderFill = {
    effect3d,
    shadow,
    slashDiagonal,
    backslashDiagonal,
    borders: {
      left: leftBorder,
      right: rightBorder,
      top: topBorder,
      bottom: bottomBorder,
    },
  };

  if (diagonalType !== 0) {
    result.diagonal = { type: diagonalType, width: diagonalWidth, color: diagonalColor };
  }

  if (data.length >= 36) {
    const fillType = readUint32(data, 32);
    let fillOffset = 36;

    if (fillType & 0x01) {
      if (data.length >= fillOffset + 12) {
        result.fill = {
          fillType: 'solid',
          backgroundColor: readUint32(data, fillOffset),
          patternColor: readUint32(data, fillOffset + 4),
          patternType: readInt32(data, fillOffset + 8),
        };
        fillOffset += 12;
      }
    } else if (fillType & 0x04) {
      if (data.length >= fillOffset + 12) {
        const gradientType = readInt16(data, fillOffset);
        const angle = readInt16(data, fillOffset + 2);
        const centerX = readInt16(data, fillOffset + 4);
        const centerY = readInt16(data, fillOffset + 6);
        const blur = readInt16(data, fillOffset + 8);
        const numColors = readInt16(data, fillOffset + 10);
        fillOffset += 12;

        const colors: Array<{ position?: number; color: number }> = [];
        const positionsCount = Math.max(0, numColors - 2);
        fillOffset += positionsCount * 4;

        for (let i = 0; i < numColors && fillOffset + 4 <= data.length; i++) {
          colors.push({ color: readUint32(data, fillOffset) });
          fillOffset += 4;
        }

        result.fill = {
          fillType: 'gradient',
          gradientType,
          angle,
          centerX,
          centerY,
          blur,
          colors,
        };
      }
    } else if (fillType & 0x02) {
      if (data.length >= fillOffset + 6) {
        result.fill = {
          fillType: 'image',
          imageType: data[fillOffset],
          brightness: readInt8(data, fillOffset + 1),
          contrast: readInt8(data, fillOffset + 2),
          effect: data[fillOffset + 3],
          binItemId: readUint16(data, fillOffset + 4),
        };
      }
    }
  }

  return result;
}

function parseSectionData(
   data: Uint8Array,
   images: Map<string, HwpxImage>,
   faceNames: Map<number, ParsedFaceName> = new Map(),
   charShapes: Map<number, ParsedCharShape> = new Map(),
   paraShapes: Map<number, ParsedParaShape> = new Map(),
   borderFills: Map<number, ParsedBorderFill> = new Map()
): HwpxSection {
   const section: HwpxSection = {
     elements: [],
     pageSettings: { width: 595, height: 842, marginTop: 56.7, marginBottom: 56.7, marginLeft: 56.7, marginRight: 56.7 },
   };

  const ctx: ParseContext = {
    currentParagraph: null,
    currentCharShapeId: 0,
    currentParaShapeId: 0,
    currentStyleId: 0,
    charShapePositions: [],
    pendingTextSegments: [],
    pendingLineSegs: [],
    pendingRangeTags: [],
    currentTable: null,
    currentTableRow: 0,
    currentTableCol: 0,
    currentCtrlId: 0,
    inTableCell: false,
    cellParagraphs: [],
    tableRowCount: 0,
    tableColCount: 0,
    tableCells: [],
    pendingImage: null,
    pendingSectionImages: [],
    faceNames,
    charShapes,
    paraShapes,
    borderFills,
    inHeaderFooter: false,
    currentHeaderFooter: null,
    headerFooterParagraphs: [],
    inFootnoteEndnote: false,
    currentFootnoteEndnote: null,
    footnoteEndnoteParagraphs: [],
    currentSectionDef: null,
    currentColumnDef: null,
    nestedLevel: 0,
    pendingShape: null,
    pendingField: null,
    inMemo: false,
    currentMemo: null,
    memoParagraphs: [],
    memos: [],
    inShapeText: false,
    shapeTextParagraphs: [],
    inEquation: false,
    equationScript: '',
  };

  let offset = 0;
  let prevLevel = 0;
  let currentParagraphLevel = 0;
  let memoListActive = false;
  let memoListLevel = -1;
  let discardCurrentParagraph = false;
  let paraTextHadInlineControls = false;
  let paraTextWasPresent = false;

  const flushPendingTextSegments = () => {
    if (!ctx.currentParagraph || ctx.pendingTextSegments.length === 0) return;
    for (const segment of ctx.pendingTextSegments) {
      let applicableCharShapeId = ctx.currentCharShapeId;
      for (const pos of ctx.charShapePositions) {
        if (pos.startPos <= segment.start) {
          applicableCharShapeId = pos.charShapeId;
        }
      }
      const charShape = ctx.charShapes.get(applicableCharShapeId);
      const charStyle = charShape ? charShapeToStyleStandalone(charShape, ctx.faceNames) : undefined;
      ctx.currentParagraph.runs.push({ text: segment.text, charStyle });
    }
    ctx.pendingTextSegments = [];
  };

  const pushCurrentParagraph = () => {
    if (!ctx.currentParagraph) return;
    if (discardCurrentParagraph) {
      discardCurrentParagraph = false;
      ctx.currentParagraph = null;
      return;
    }
    flushPendingTextSegments();
    if (ctx.currentParagraph.runs.length === 0) {
      const emptyCS = ctx.charShapePositions.length > 0
        ? ctx.charShapes.get(ctx.charShapePositions[0].charShapeId)
        : ctx.charShapes.get(ctx.currentCharShapeId);
      const emptyStyle = emptyCS ? charShapeToStyleStandalone(emptyCS, ctx.faceNames) : undefined;
      ctx.currentParagraph.runs.push({ text: '', charStyle: emptyStyle });
    }
    if (ctx.inTableCell) {
      // Table cell takes highest priority — paragraphs inside table cells
      // (even within headers/footers/footnotes) must go to cellParagraphs,
      // UNLESS they belong to a shape text overlay (caption etc.)
      if (ctx.inShapeText && currentParagraphLevel >= shapeTextLevel) {
        ctx.shapeTextParagraphs.push(ctx.currentParagraph);
      } else {
        ctx.cellParagraphs.push(ctx.currentParagraph);
      }
    } else if (ctx.inHeaderFooter && currentParagraphLevel > headerFooterLevel) {
      ctx.headerFooterParagraphs.push(ctx.currentParagraph);
    } else if (ctx.inFootnoteEndnote && currentParagraphLevel > footnoteEndnoteLevel) {
      ctx.footnoteEndnoteParagraphs.push(ctx.currentParagraph);
    } else if (ctx.inMemo && currentParagraphLevel > memoLevel) {
      ctx.memoParagraphs.push(ctx.currentParagraph);
    } else if (ctx.inShapeText && currentParagraphLevel >= shapeTextLevel) {
      ctx.shapeTextParagraphs.push(ctx.currentParagraph);
    } else if (currentParagraphLevel === 0) {
      section.elements.push({ type: 'paragraph', data: ctx.currentParagraph });
    }
    ctx.currentParagraph = null;
  };

  type TableStackItem = {
    table: typeof ctx.currentTable;
    cells: typeof ctx.tableCells;
    rowCount: number;
    colCount: number;
    currentRow: number;
    currentCol: number;
    inCell: boolean;
    cellParagraphs: typeof ctx.cellParagraphs;
    level: number;
    cellContentLevel: number;
  };
  const tableStack: TableStackItem[] = [];
  type ShapeTextStackItem = {
    level: number;
    paragraphs: typeof ctx.shapeTextParagraphs;
  };
  const shapeTextStack: ShapeTextStackItem[] = [];
  let currentTableLevel = 0;
  let cellContentLevel = -1;
  let headerFooterLevel = -1;
  let footnoteEndnoteLevel = -1;
  let memoLevel = -1;
  let shapeTextLevel = -1;
  let _traceTableCount = 0;

  const finishCurrentTable = () => {
    if (!ctx.currentTable) return;

    if (ctx.inTableCell && ctx.currentParagraph) {
      flushPendingTextSegments();
      if (ctx.currentParagraph.runs.length === 0) {
        ctx.currentParagraph.runs.push({ text: '' });
      }
      if (ctx.inShapeText && currentParagraphLevel >= shapeTextLevel) {
        ctx.shapeTextParagraphs.push(ctx.currentParagraph);
      } else {
        ctx.cellParagraphs.push(ctx.currentParagraph);
      }
      ctx.currentParagraph = null;
    }

    const row = ctx.currentTableRow;
    const col = ctx.currentTableCol;
    if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]) {
      // Normalize cell paragraphs: remove leading/trailing empty paragraphs and
      // collapse consecutive empty paragraphs into one to match HWPX output.
      let normalized = [...ctx.cellParagraphs];
      const isEmptyParagraph = (p: typeof normalized[0]) =>
        p.runs.every(r => !r.text || r.text.trim() === '');
      // Remove leading empty paragraphs
      while (normalized.length > 0 && isEmptyParagraph(normalized[0])) {
        normalized.shift();
      }
      // Remove trailing empty paragraphs
      while (normalized.length > 0 && isEmptyParagraph(normalized[normalized.length - 1])) {
        normalized.pop();
      }
      // Collapse consecutive empty paragraphs into one
      const collapsed: typeof normalized = [];
      let prevWasEmpty = false;
      for (const p of normalized) {
        const empty = isEmptyParagraph(p);
        if (empty && prevWasEmpty) continue;
        collapsed.push(p);
        prevWasEmpty = empty;
      }
      ctx.tableCells[row][col].paragraphs = collapsed;
      // Populate cell.elements from paragraphs to match HWPX structure.
      // The webview uses cell.elements when available, falling back to cell.paragraphs.
      if (!ctx.tableCells[row][col].elements) ctx.tableCells[row][col].elements = [];
      for (const p of collapsed) {
        ctx.tableCells[row][col].elements!.push({ type: 'paragraph', data: p });
      }
    }

    // Calculate table width and columnWidths from first row cells
    if (ctx.tableColCount > 0 && ctx.tableCells[0]) {
      const colWidths: number[] = [];
      let totalWidth = 0;
      for (let c = 0; c < ctx.tableColCount; c++) {
        const cell = ctx.tableCells[0][c];
        const w = cell?.width || 0;
        colWidths.push(w);
        totalWidth += w;
      }
      ctx.currentTable.columnWidths = colWidths;
      ctx.currentTable.width = totalWidth;
    }

    const coveredCells = new Set<string>();
    for (let r = 0; r < ctx.tableRowCount; r++) {
      for (let c = 0; c < ctx.tableColCount; c++) {
        const cell = ctx.tableCells[r]?.[c];
        if (cell && ((cell.colSpan || 1) > 1 || (cell.rowSpan || 1) > 1)) {
          for (let dr = 0; dr < (cell.rowSpan || 1); dr++) {
            for (let dc = 0; dc < (cell.colSpan || 1); dc++) {
              if (dr !== 0 || dc !== 0) {
                coveredCells.add(`${r + dr},${c + dc}`);
              }
            }
          }
        }
      }
    }

    const rows: TableRow[] = [];
    for (let r = 0; r < ctx.tableRowCount; r++) {
      const cells: TableCell[] = [];
      for (let c = 0; c < ctx.tableColCount; c++) {
        if (coveredCells.has(`${r},${c}`)) continue;
        const cell = ctx.tableCells[r]?.[c];
        if (cell) {
          if (cell.paragraphs.length === 0) cell.paragraphs.push({ id: generateId(), runs: [{ text: '' }] });
          // Ensure cell.elements is populated from paragraphs (+ any existing elements like nested tables)
          if (!cell.elements || cell.elements.length === 0) {
            cell.elements = cell.paragraphs.map(p => ({ type: 'paragraph' as const, data: p }));
          }
          cells.push(cell);
        }
      }
      if (cells.length > 0) rows.push({ cells });
    }
    ctx.currentTable.rows = rows;

    const isNested = tableStack.length > 0;

    if (isNested) {
      const parent = tableStack[tableStack.length - 1];
      const parentRow = parent.currentRow;
      const parentCol = parent.currentCol;
      if (parentRow < parent.rowCount && parentCol < parent.colCount && parent.cells[parentRow]) {
        if (!parent.cells[parentRow][parentCol].elements) {
          parent.cells[parentRow][parentCol].elements = [];
        }
        parent.cells[parentRow][parentCol].elements!.push({ type: 'table', data: ctx.currentTable });
      }
     } else if (!ctx.inHeaderFooter && !ctx.inFootnoteEndnote) {
       section.elements.push({ type: 'table', data: ctx.currentTable });
     }

     if ((global as any).__HWP_TRACE_ACTIVE && !isNested) { (global as any).__HWP_TRACE_ACTIVE = false; console.log(`=== TABLE END ===\n`); }

     // After outermost table finishes, flush any deferred section-level images
     // so they appear after their containing table (matching HWPX element ordering).
     if (!isNested && ctx.pendingSectionImages.length > 0) {
       for (const img of ctx.pendingSectionImages) {
         section.elements.push({ type: 'image', data: img });
       }
       ctx.pendingSectionImages = [];
     }

     ctx.currentTable = null;
     ctx.inTableCell = false;
     ctx.cellParagraphs = [];
   };

  const restoreParentTable = () => {
    if (tableStack.length === 0) return;
    const parent = tableStack.pop()!;
    ctx.currentTable = parent.table;
    ctx.tableCells = parent.cells;
    ctx.tableRowCount = parent.rowCount;
    ctx.tableColCount = parent.colCount;
    ctx.currentTableRow = parent.currentRow;
    ctx.currentTableCol = parent.currentCol;
    ctx.inTableCell = parent.inCell;
    ctx.cellParagraphs = parent.cellParagraphs;
    cellContentLevel = parent.cellContentLevel;
  };

  while (offset < data.length) {
    if (offset + 4 > data.length) break;

    const header = readUint32(data, offset);
    const tagId = header & 0x3FF;
    const level = (header >>> 10) & 0x3FF;
    let size = (header >>> 20) & 0xFFF;
    let nextOffset = offset + 4;

    if (size === 0xFFF) {
      if (nextOffset + 4 > data.length) break;
      size = readUint32(data, nextOffset);
      nextOffset += 4;
    }

     prevLevel = level;

      const recordData = data.subarray(nextOffset, nextOffset + size);

      if ((global as any).__HWP_TRACE_TAGS && (global as any).__HWP_TRACE_ACTIVE) {
        const _tn: Record<number, string> = {66:'PARA_HDR',67:'PARA_TXT',68:'PARA_CS',69:'PARA_LS',70:'PARA_RT',71:'CTRL_HDR',72:'LIST_HDR',73:'PAGE_DEF',77:'TABLE',76:'SHAPE_COMP',85:'SHAPE_PIC',87:'SHAPE_TBOX'};
        const _t = _tn[tagId] || `TAG_${tagId}`;
        let extra = '';
        if (tagId === 71 && recordData.length >= 4) {
          const cid = readUint32(recordData, 0);
          const cn: Record<number, string> = {0x6C626174:'TABLE',0x64736F67:'GSO',0x63697024:'PIC',0x6E717565:'EQU',0x64616568:'HEADER',0x746F6F66:'FOOTER'};
          extra = ` (${cn[cid] || '0x' + cid.toString(16)})`;
        }
        if (tagId === 72) extra = ` [inCell=${ctx.inTableCell} tblLvl=${currentTableLevel}]`;
        console.log(`${'  '.repeat(level)}[L${level}] ${_t}${extra} | shp=${ctx.inShapeText}:${shapeTextLevel} stk=${shapeTextStack.length} cellLvl=${cellContentLevel}`);
      }

     while (ctx.currentTable && level <= currentTableLevel) {
       finishCurrentTable();
       if (tableStack.length > 0) {
         const parent = tableStack[tableStack.length - 1];
         restoreParentTable();
         currentTableLevel = parent.level;
       } else {
         break;
       }
     }

      const _isParaSubTag = tagId >= HWP_TAGS.HWPTAG_PARA_HEADER && tagId <= HWP_TAGS.HWPTAG_PARA_RANGE_TAG;
      if (ctx.inHeaderFooter && headerFooterLevel >= 0 && level <= headerFooterLevel && !_isParaSubTag) {
        pushCurrentParagraph();
        ctx.inHeaderFooter = false;
        headerFooterLevel = -1;
      }
      if (ctx.inFootnoteEndnote && footnoteEndnoteLevel >= 0 && level <= footnoteEndnoteLevel && !_isParaSubTag) {
        pushCurrentParagraph();
        ctx.inFootnoteEndnote = false;
        footnoteEndnoteLevel = -1;
      }
      if (ctx.inMemo && memoLevel >= 0 && level <= memoLevel && !_isParaSubTag) {
        pushCurrentParagraph();
        ctx.inMemo = false;
        memoLevel = -1;
      }
      const _isBodySectionTag = tagId >= HWP_TAGS.HWPTAG_PARA_HEADER;
      if (ctx.inShapeText && shapeTextLevel >= 0 && level <= shapeTextLevel && _isBodySectionTag) {
        if (level < shapeTextLevel || !_isParaSubTag) {
          const isShapeInternalPara = _isParaSubTag && level < shapeTextLevel && ctx.inTableCell && level !== cellContentLevel && level > currentTableLevel;
          if (!isShapeInternalPara) {
            pushCurrentParagraph();
            if (shapeTextStack.length > 0) {
              const parent = shapeTextStack.pop()!;
              shapeTextLevel = parent.level;
              ctx.shapeTextParagraphs = parent.paragraphs;
            } else {
              ctx.inShapeText = false;
              shapeTextLevel = -1;
              // If this was an equation context, emit the equation element
              if (ctx.inEquation) {
                const script = ctx.shapeTextParagraphs
                  .flatMap(p => p.runs.map(r => r.text))
                  .join('');
                const equation: HwpxEquation = { id: generateId(), script };
                // Always emit equations at section level to match HWPX behavior
                if (!ctx.inHeaderFooter && !ctx.inFootnoteEndnote) {
                  section.elements.push({ type: 'equation', data: equation });
                }
                ctx.inEquation = false;
              }
              ctx.shapeTextParagraphs = [];
            }
          }
        }
      }
      if (memoListActive && level < memoListLevel) {
        memoListActive = false;
        memoListLevel = -1;
      }

      switch (tagId) {
       case HWP_TAGS.HWPTAG_PARA_HEADER:
         pushCurrentParagraph();
         ctx.currentParagraph = { id: generateId(), runs: [] };
         currentParagraphLevel = level;
          ctx.charShapePositions = [];
          ctx.pendingTextSegments = [];
          paraTextHadInlineControls = false;
          paraTextWasPresent = false;
          if (memoListActive && level >= memoListLevel) {
           discardCurrentParagraph = true;
         }
        if (recordData.length >= 12) {
          const paraShapeId = readUint16(recordData, 8);
          const paraShape = ctx.paraShapes.get(paraShapeId);
          if (paraShape) {
            const alignMap: Record<number, 'justify' | 'left' | 'right' | 'center' | 'distribute'> = {
              0: 'justify', 1: 'left', 2: 'right', 3: 'center', 4: 'distribute', 5: 'distribute'
            };
            ctx.currentParagraph.paraStyle = {
              align: alignMap[paraShape.alignment] || 'justify',
              marginLeft: hwpunitToPt(paraShape.leftMargin),
              marginRight: hwpunitToPt(paraShape.rightMargin),
              firstLineIndent: hwpunitToPt(paraShape.indent),
              marginTop: hwpunitToPt(paraShape.spacingBefore),
              marginBottom: hwpunitToPt(paraShape.spacingAfter),
              lineSpacing: paraShape.lineSpacingType === 0 ? paraShape.lineSpacing : paraShape.lineSpacing / 100,
              lineSpacingType: paraShape.lineSpacingType === 0 ? 'percent' : 'fixed',
              keepWithNext: paraShape.keepWithNext,
              keepLines: paraShape.keepTogether,
            };
          }
        }
        break;

      case HWP_TAGS.HWPTAG_PARA_CHAR_SHAPE:
        ctx.charShapePositions = [];
        for (let j = 0; j + 8 <= recordData.length; j += 8) {
          const startPos = readUint32(recordData, j);
          const charShapeId = readUint32(recordData, j + 4);
          ctx.charShapePositions.push({ startPos, charShapeId });
        }
        if (ctx.charShapePositions.length > 0) {
          ctx.currentCharShapeId = ctx.charShapePositions[0].charShapeId;
        }
        break;

      case HWP_TAGS.HWPTAG_PARA_TEXT:
        paraTextWasPresent = true;
        if (ctx.currentParagraph) {
          let currentStart = 0;
          const charCodes: number[] = [];
          let charIndex = 0;
          let i = 0;
          while (i < recordData.length - 1) {
            const charCode = readUint16(recordData, i);
            i += 2;
            if (charCode === 0) { charIndex++; continue; }
            if (charCode < 32) {
              if (charCodes.length > 0) {
                ctx.pendingTextSegments.push({ start: currentStart, end: charIndex, text: charCodes.map(c => String.fromCharCode(c)).join('') });
                charCodes.length = 0;
              }
              if (charCode === CTRL_CHAR.LINE_BREAK) {
                ctx.pendingTextSegments.push({ start: charIndex, end: charIndex + 1, text: '\n' });
                charIndex++;
              } else if (charCode === 0x0009) {
                // Skip tab text to match HWPX behavior (tabs become empty runs with tab property)
                i += 14; charIndex += 8;
               } else if (charCode >= 0x0002 && charCode <= 0x0008) {
                 paraTextHadInlineControls = true;
                 i += 14; charIndex += 8;
               } else if (charCode === 0x000B || charCode === 0x000C ||
                          charCode === 0x000E || charCode === 0x000F || charCode === 0x0010 ||
                          charCode === 0x0011 || charCode === 0x0012 || charCode === 0x0013 ||
                          charCode === 0x0015 || charCode === 0x0016 || charCode === 0x0017) {
                 paraTextHadInlineControls = true;
                 i += 14; charIndex += 8;
              } else if (charCode === 0x000D) {
                break;
              } else if (charCode === 0x001E) {
                ctx.pendingTextSegments.push({ start: charIndex, end: charIndex + 1, text: '\u00A0' });
                charIndex++;
              } else if (charCode === 0x001F) {
                ctx.pendingTextSegments.push({ start: charIndex, end: charIndex + 1, text: '\u3000' });
                charIndex++;
              } else {
                charIndex++;
              }
              currentStart = charIndex;
              continue;
            }
            charCodes.push(charCode);
            charIndex++;
          }
          if (charCodes.length > 0) {
            ctx.pendingTextSegments.push({ start: currentStart, end: charIndex, text: charCodes.map(c => String.fromCharCode(c)).join('') });
          }
        }
        break;

         case HWP_TAGS.HWPTAG_CTRL_HEADER:
           if (recordData.length >= 4) {
             ctx.currentCtrlId = readUint32(recordData, 0);
              if (ctx.currentCtrlId === CTRL_ID.TABLE) {
             if ((global as any).__HWP_TRACE_TAGS) {
               const _tt = (global as any).__HWP_TRACE_TARGET_TABLE;
               if (_traceTableCount === _tt) { (global as any).__HWP_TRACE_ACTIVE = true; console.log(`\n=== TABLE ${_traceTableCount} START (level=${level}) ===`); }
               _traceTableCount++;
             }
              if (ctx.currentParagraph && !ctx.inHeaderFooter && !ctx.inFootnoteEndnote
                  && !ctx.inMemo && !ctx.inShapeText && !ctx.inTableCell && !discardCurrentParagraph) {
                pushCurrentParagraph();
              }
              if (ctx.currentTable) {
                 tableStack.push({
                   table: ctx.currentTable,
                   cells: ctx.tableCells,
                   rowCount: ctx.tableRowCount,
                   colCount: ctx.tableColCount,
                   currentRow: ctx.currentTableRow,
                   currentCol: ctx.currentTableCol,
                   inCell: ctx.inTableCell,
                   cellParagraphs: ctx.cellParagraphs,
                   level: currentTableLevel,
                   cellContentLevel,
                 });
              }
              ctx.currentTable = { id: generateId(), rows: [], rowCount: 0, colCount: 0 };

              // Parse table-level fields from CTRL_HEADER common object header
              if (recordData.length >= 24) {
                const props = readUint32(recordData, 4);
                const treatAsChar = !!(props & (1 << 4));

                if (treatAsChar) {
                  // Inline table: textWrap=topAndBottom, relative to paragraph
                  ctx.currentTable.textWrap = 'topAndBottom' as any;
                  ctx.currentTable.position = {
                    treatAsChar: true,
                    flowWithText: true,
                    vertRelTo: 'para' as any,
                    horzRelTo: 'para' as any,
                    vertAlign: 'top' as any,
                    horzAlign: 'left' as any,
                    vertOffset: 0,
                    horzOffset: 0,
                  };
                } else {
                  // Floating table: parse wrap/position from properties
                  const textWrapVal = (props >> 1) & 0x07;
                  const textWrapMap: Record<number, string> = {
                    0: 'square', 1: 'topAndBottom', 2: 'behindText', 3: 'inFrontOfText',
                    4: 'tight', 5: 'through'
                  };
                  const vertRelVal = (props >> 5) & 0x03;
                  const vertRelMap: Record<number, string> = { 0: 'paper', 1: 'page', 2: 'para' };
                  const horzRelVal = (props >> 9) & 0x03;
                  const horzRelMap: Record<number, string> = { 0: 'paper', 1: 'page', 2: 'column', 3: 'para' };

                  ctx.currentTable.textWrap = (textWrapMap[textWrapVal] || 'topAndBottom') as any;
                  ctx.currentTable.position = {
                    treatAsChar: false,
                    flowWithText: !!(props & (1 << 12)),
                    vertRelTo: (vertRelMap[vertRelVal] || 'para') as any,
                    horzRelTo: (horzRelMap[horzRelVal] || 'column') as any,
                    vertOffset: readInt32(recordData, 8) / 7200 * 72,
                    horzOffset: readInt32(recordData, 12) / 7200 * 72,
                  };
                }
              }
              // Parse outMargin (4 x uint16 at offset 28: left, right, top, bottom) after zOrder at 24
              if (recordData.length >= 36) {
                ctx.currentTable.outMargin = {
                  left: readUint16(recordData, 28) / 100,
                  right: readUint16(recordData, 30) / 100,
                  top: readUint16(recordData, 32) / 100,
                  bottom: readUint16(recordData, 34) / 100,
                };
              }
              // borderFillId and pageBreak/repeatHeader will be filled from HWPTAG_TABLE record

              ctx.tableCells = [];
              ctx.inTableCell = false;
              ctx.cellParagraphs = [];
              currentTableLevel = level;
           } else if (ctx.currentCtrlId === CTRL_ID.PICTURE || ctx.currentCtrlId === CTRL_ID.GSO) {
              // Host paragraph for PICTURE/GSO: if at section level with no visible text,
              // discard it since HWPX doesn't emit host paragraphs for images/drawings.
              if (ctx.currentParagraph && !ctx.inHeaderFooter && !ctx.inFootnoteEndnote
                  && !ctx.inMemo && !ctx.inShapeText && !ctx.inTableCell
                  && paraTextWasPresent && paraTextHadInlineControls
                  && ctx.pendingTextSegments.length === 0 && ctx.currentParagraph.runs.length === 0) {
                discardCurrentParagraph = true;
              }
              ctx.pendingImage = { width: 200, height: 150 };
             if (recordData.length >= 46) {
               const w = readUint32(recordData, 16);
               const h = readUint32(recordData, 20);
               if (w > 0) ctx.pendingImage.width = w / 7200 * 72;
               if (h > 0) ctx.pendingImage.height = h / 7200 * 72;
             }
              if (ctx.inShapeText && shapeTextLevel >= 0) {
                 shapeTextStack.push({ level: shapeTextLevel, paragraphs: ctx.shapeTextParagraphs });
               }
                 ctx.inShapeText = true;
                 ctx.shapeTextParagraphs = [];
                 shapeTextLevel = level;
           } else if (ctx.currentCtrlId === CTRL_ID.HEADER || ctx.currentCtrlId === CTRL_ID.FOOTER) {
            ctx.inHeaderFooter = true;
            ctx.headerFooterParagraphs = [];
            headerFooterLevel = level;
           } else if (ctx.currentCtrlId === CTRL_ID.FOOTNOTE || ctx.currentCtrlId === CTRL_ID.ENDNOTE) {
             ctx.inFootnoteEndnote = true;
             ctx.footnoteEndnoteParagraphs = [];
             footnoteEndnoteLevel = level;
           } else if (ctx.currentCtrlId === CTRL_ID.EQUATION) {
              ctx.inEquation = true;
              ctx.equationScript = '';
              if (ctx.inShapeText && shapeTextLevel >= 0) {
                shapeTextStack.push({ level: shapeTextLevel, paragraphs: ctx.shapeTextParagraphs });
              }
              ctx.inShapeText = true;
              ctx.shapeTextParagraphs = [];
              shapeTextLevel = level;
           } else if (ctx.currentCtrlId === CTRL_ID.FIELD_MEMO) {
             ctx.inMemo = true;
             ctx.memoParagraphs = [];
             memoLevel = level;
            } else if (ctx.currentCtrlId !== CTRL_ID.SECTION
                         && ctx.currentCtrlId !== CTRL_ID.COLUMN
                         && ctx.currentCtrlId !== CTRL_ID.AUTO_NUMBER
                         && ctx.currentCtrlId !== CTRL_ID.PAGE_NUMBER_POS
                         && ctx.currentCtrlId !== CTRL_ID.BOOKMARK
                         && ctx.currentCtrlId !== CTRL_ID.TCPS
                         && ctx.currentCtrlId !== CTRL_ID.FORM
                         && ((ctx.currentCtrlId >> 24) & 0xFF) !== 0x25) {
                if (ctx.inShapeText && shapeTextLevel >= 0) {
                  shapeTextStack.push({ level: shapeTextLevel, paragraphs: ctx.shapeTextParagraphs });
                }
                ctx.inShapeText = true;
                ctx.shapeTextParagraphs = [];
                shapeTextLevel = level;
            }
         }
        break;

        case HWP_TAGS.HWPTAG_TABLE:
          if (ctx.currentTable && recordData.length >= 8) {
            const rowCount = readUint16(recordData, 4);
            const colCount = readUint16(recordData, 6);
           ctx.currentTable.rowCount = rowCount;
          ctx.currentTable.colCount = colCount;
          ctx.tableRowCount = rowCount;
          ctx.tableColCount = colCount;
           ctx.tableCells = [];
          for (let r = 0; r < rowCount; r++) {
            ctx.tableCells[r] = [];
            for (let c = 0; c < colCount; c++) {
              ctx.tableCells[r][c] = { paragraphs: [], colAddr: c, rowAddr: r, colSpan: 1, rowSpan: 1 };
            }
          }

          const cellSpacing = recordData.length >= 10 ? readUint16(recordData, 8) / 7200 * 72 : 0;
          ctx.currentTable.cellSpacing = cellSpacing;

          // Extract table internal cell padding (inMargin) from border padding fields
          if (recordData.length >= 18) {
            ctx.currentTable.inMargin = {
              left: readUint16(recordData, 10) / 7200 * 72,
              right: readUint16(recordData, 12) / 7200 * 72,
              top: readUint16(recordData, 14) / 7200 * 72,
              bottom: readUint16(recordData, 16) / 7200 * 72,
            };
          }

          // borderFillId at offset 18 (uint16, 1-based -> convert to 0-based)
          if (recordData.length >= 20) {
            const bfId = readUint16(recordData, 18);
            if (bfId > 0) ctx.currentTable.borderFillId = bfId;
          }

          // Parse properties flags at offset 0 (uint32): pageBreak and repeatHeader
          if (recordData.length >= 4) {
            const tableProps = readUint32(recordData, 0);
            // bit 0-1: pageBreak (0=none, 1=cell, 2=row)
            const pageBreakVal = tableProps & 0x03;
            const pageBreakMap: Record<number, string> = { 0: 'none', 1: 'cell', 2: 'row' };
            ctx.currentTable.pageBreak = (pageBreakMap[pageBreakVal] || 'none') as any;
            // bit 2: repeatHeader
            ctx.currentTable.repeatHeader = !!(tableProps & (1 << 2));
          }

          ctx.currentTableRow = 0;
          ctx.currentTableCol = 0;
        }
        break;

       case HWP_TAGS.HWPTAG_LIST_HEADER:
         if (ctx.currentTable && recordData.length >= 34 && level === currentTableLevel + 1) {
              const headerSize = 8;
             const cellCol = readUint16(recordData, headerSize);
            const cellRow = readUint16(recordData, headerSize + 2);
            if (cellRow >= ctx.tableRowCount || cellCol >= ctx.tableColCount) {
              break;
            }

            if (ctx.inTableCell) {
              if (ctx.currentParagraph) {
                flushPendingTextSegments();
                if (ctx.currentParagraph.runs.length === 0) {
                  ctx.currentParagraph.runs.push({ text: '' });
                }
                if (ctx.inShapeText && currentParagraphLevel >= shapeTextLevel) {
                  ctx.shapeTextParagraphs.push(ctx.currentParagraph);
                } else {
                  ctx.cellParagraphs.push(ctx.currentParagraph);
                }
              }
             const prevRow = ctx.currentTableRow;
             const prevCol = ctx.currentTableCol;
              if (prevRow < ctx.tableRowCount && prevCol < ctx.tableColCount && ctx.tableCells[prevRow]) {
                ctx.tableCells[prevRow][prevCol].paragraphs = [...ctx.cellParagraphs];
              }
            }

            const colSpan = readUint16(recordData, headerSize + 4);
            const rowSpan = readUint16(recordData, headerSize + 6);
           const cellWidth = readUint32(recordData, headerSize + 8) / 7200 * 72;
           const cellHeight = readUint32(recordData, headerSize + 12) / 7200 * 72;
           const marginLeft = readUint16(recordData, headerSize + 16) / 7200 * 72;
           const marginRight = readUint16(recordData, headerSize + 18) / 7200 * 72;
           const marginTop = readUint16(recordData, headerSize + 20) / 7200 * 72;
           const marginBottom = readUint16(recordData, headerSize + 22) / 7200 * 72;
           const borderFillId = recordData.length > headerSize + 24 ? readUint16(recordData, headerSize + 24) : 0;

           // Extract verticalAlign from list header flags (offset 2, 4 bytes)
           const listFlags = readUint32(recordData, 2);
           // Bits 20-21: vertical alignment (0=font-based→top, 1=top, 2=center/middle, 3=bottom)
           const vertAlignBits = (listFlags >> 20) & 0x03;
           const verticalAlign: 'top' | 'middle' | 'bottom' = vertAlignBits === 2 ? 'middle' : vertAlignBits === 3 ? 'bottom' : 'top';

           // hasMargin: In HWPX, hasMargin="0" means use table's inMargin, "1" means use cell's own margin.
           // In HWP binary, cell margins are always stored. Compare against table inMargin to determine.
           const tableInMargin = ctx.currentTable.inMargin;
           const hasMargin = tableInMargin ? !(
             Math.abs(marginLeft - tableInMargin.left) < 0.01 &&
             Math.abs(marginRight - tableInMargin.right) < 0.01 &&
             Math.abs(marginTop - tableInMargin.top) < 0.01 &&
             Math.abs(marginBottom - tableInMargin.bottom) < 0.01
           ) : false;

           ctx.currentTableCol = cellCol;
           ctx.currentTableRow = cellRow;

          if (cellRow < ctx.tableRowCount && cellCol < ctx.tableColCount && ctx.tableCells[cellRow]) {
            const cell = ctx.tableCells[cellRow][cellCol];
            cell.colSpan = colSpan || 1;
            cell.rowSpan = rowSpan || 1;
            cell.width = cellWidth;
            cell.height = cellHeight;
            cell.marginLeft = marginLeft;
            cell.marginRight = marginRight;
            cell.marginTop = marginTop;
            cell.marginBottom = marginBottom;
            cell.borderFillId = borderFillId;
            cell.verticalAlign = verticalAlign;
            cell.hasMargin = hasMargin;
            cell.textDirection = 'horizontal';
            cell.lineWrap = 'break';

            const borderFill = ctx.borderFills.get(borderFillId);
            if (borderFill) {
              const fill = borderFill.fill;
              if (fill && fill.fillType === 'solid' && fill.backgroundColor !== undefined) {
                cell.backgroundColor = colorrefToHex((fill as ParsedSolidFill).backgroundColor);
              } else if (fill && fill.fillType === 'gradient' && (fill as ParsedGradientFill).colors && (fill as ParsedGradientFill).colors.length > 0) {
                const gf = fill as ParsedGradientFill;
                const gradTypeMap: Record<number, string> = { 1: 'Linear', 2: 'Radial', 3: 'Conical', 4: 'Square' };
                cell.backgroundGradation = {
                  type: (gradTypeMap[gf.gradientType] || 'Linear') as GradationType,
                  angle: gf.angle || 0,
                  colors: gf.colors.map(c => colorrefToHex(c.color)),
                };
              }
              if (borderFill.borders) {
                // HWP 5.0 border type: 0=none,1=solid,2=dash,3=dot,4=dash-dot,5=dash-dot-dot,6=long-dash,7=circle,8=double,9=thin-thick,10=thick-thin,11=thin-thick-thin,12=wave,13=double-wave,14=thick-3d,15=thin-3d-reverse,16=3d,17=3d-reverse
                const borderStyleMap: Record<number, string> = { 0: 'none', 1: 'solid', 2: 'dashed', 3: 'dotted', 4: 'dashed', 5: 'dashed', 6: 'dashed', 7: 'dotted', 8: 'double', 255: 'none' };
                const mapBorder = (b: ParsedBorderLine) => {
                  const style = borderStyleMap[b.type] ?? 'solid';
                  if (style === 'none') return undefined;
                  return {
                    style,
                    width: b.width * 2.83465,
                    color: colorrefToHex(b.color)
                  };
                };
                cell.borderTop = mapBorder(borderFill.borders.top);
                cell.borderBottom = mapBorder(borderFill.borders.bottom);
                cell.borderLeft = mapBorder(borderFill.borders.left);
                cell.borderRight = mapBorder(borderFill.borders.right);
              }
            }
          }

           ctx.inTableCell = true;
           cellContentLevel = level + 1;
           ctx.cellParagraphs = [];
           ctx.currentParagraph = null;
           ctx.inShapeText = false;
           shapeTextLevel = -1;
           shapeTextStack.length = 0;
           ctx.shapeTextParagraphs = [];
           if ((global as any).__HWP_TRACE_ACTIVE) console.log(`--- Cell [${cellRow},${cellCol}] cellContentLevel=${cellContentLevel} ---`);
        }
        break;

      case HWP_TAGS.HWPTAG_SHAPE_COMPONENT:
        if (ctx.pendingImage && recordData.length >= 36) {
          // SHAPE_COMPONENT layout: ShapeID(4) + ComponentID(4) + ...
          // Width at offset 28, Height at offset 32 (verified empirically)
          const w = readInt32(recordData, 28);
          const h = readInt32(recordData, 32);
          if (w > 0) ctx.pendingImage.width = w / 7200 * 72;
          if (h > 0) ctx.pendingImage.height = h / 7200 * 72;
        }
        break;

       case HWP_TAGS.HWPTAG_SHAPE_COMPONENT_PICTURE:
        if (recordData.length >= 45) {
          // Try multiple known offsets for binItemId (2-byte uint16)
          const _picOffsets = [71, 43, 45, 67, 69];
          let binItemId = 0;
          let idStr = '';
          // Try primary offset first, then fallbacks
          for (const off of _picOffsets) {
            if (recordData.length >= off + 2) {
              const candidate = readUint16(recordData, off);
              if (candidate > 0) {
                const candidateStr = `BIN${String(candidate).padStart(4, '0')}`;
                if (images.has(candidateStr)) {
                  binItemId = candidate;
                  idStr = candidateStr;
                  break;
                }
              }
            }
          }
          // Final fallback: scan backwards for any valid binItemId in the images map
          if (!idStr) {
            for (let off = recordData.length - 2; off >= 0; off -= 2) {
              const candidate = readUint16(recordData, off);
              if (candidate > 0) {
                const candidateStr = `BIN${String(candidate).padStart(4, '0')}`;
                if (images.has(candidateStr)) {
                  binItemId = candidate;
                  idStr = candidateStr;
                  break;
                }
              }
            }
          }
          // If still nothing found, use offset 71 or 43 as best guess
          if (!idStr) {
            const off = recordData.length >= 73 ? 71 : 43;
            if (recordData.length >= off + 2) {
              binItemId = readUint16(recordData, off);
              idStr = `BIN${String(binItemId).padStart(4, '0')}`;
            }
          }
          const existingImage = idStr ? images.get(idStr) : undefined;
          const image: HwpxImage = {
            id: generateId(),
            binaryId: idStr || 'BIN0000',
            width: ctx.pendingImage?.width || 200,
            height: ctx.pendingImage?.height || 150,
            data: existingImage?.data,
            mimeType: existingImage?.mimeType,
          };
          // Emit images at section level to match HWPX behavior.
          // HWPX puts floating images (textWrap=topAndBottom etc.) at section level
          // even when they're anchored to a paragraph inside a table cell.
          // When inside a table, defer emission so images appear AFTER their
          // containing table in the element list (matching HWPX ordering).
          if (!ctx.inHeaderFooter && !ctx.inFootnoteEndnote) {
            if (ctx.currentTable || tableStack.length > 0) {
              ctx.pendingSectionImages.push(image);
            } else {
              section.elements.push({ type: 'image', data: image });
            }
          }
          ctx.pendingImage = null;
        }
        break;

      case HWP_TAGS.HWPTAG_MEMO_LIST:
        memoListActive = true;
        memoListLevel = level;
        break;

      case HWP_TAGS.HWPTAG_PAGE_DEF:
        if (recordData.length >= 40) {
          section.pageSettings = {
            width: readUint32(recordData, 0) / 100,
            height: readUint32(recordData, 4) / 100,
            marginLeft: readUint32(recordData, 8) / 100,
            marginRight: readUint32(recordData, 12) / 100,
            marginTop: readUint32(recordData, 16) / 100,
            marginBottom: readUint32(recordData, 20) / 100,
            orientation: (readUint32(recordData, 36) & 0x01) ? 'landscape' : 'portrait',
          };
        }
        break;
    }

    offset = nextOffset + size;
  }

  // Finish any remaining table that wasn't closed properly
  if (ctx.currentTable) {
    finishCurrentTable();
    while (tableStack.length > 0) {
      restoreParentTable();
      if (ctx.currentTable) finishCurrentTable();
    }
  }

   if (ctx.currentParagraph) {
    pushCurrentParagraph();
  }

   if (section.elements.length === 0) {
     section.elements.push({ type: 'paragraph', data: { id: generateId(), runs: [{ text: '' }] } });
   }

   return section;
 }

// ============================================================
// Main Export - parseHwpContent
// ============================================================

export function parseHwpContent(data: Uint8Array): HwpxContent {
  const cfb = CFB.read(data, { type: 'array' });

  const content: HwpxContent = {
    metadata: {},
    sections: [],
    images: new Map(),
    binItems: new Map(),
    binData: new Map(),
    footnotes: [],
    endnotes: [],
  };

  const headerData = CFB.find(cfb, '/FileHeader')?.content;
  if (!headerData || (headerData as Uint8Array).length < 256) {
    throw new Error('Invalid HWP file');
  }

  const headerBytes = headerData instanceof Uint8Array ? headerData : new Uint8Array(headerData as unknown as ArrayBuffer);
  const props = readUint32(headerBytes, 36);
  const compressed = (props & 0x01) !== 0;

  const decompress = (d: Uint8Array): Uint8Array => {
    if (!compressed) return d;
    try { return pako.inflateRaw(d); } catch { return d; }
  };

  const getEntryData = (path: string): Uint8Array | null => {
    const entry = CFB.find(cfb, path);
    if (!entry?.content) return null;
    const raw = entry.content instanceof Uint8Array ? entry.content : new Uint8Array(entry.content as unknown as ArrayBuffer);
    return decompress(raw);
  };

  const binDataInfos = new Map<number, BinDataInfo>();
  const faceNames = new Map<number, ParsedFaceName>();
  const charShapes = new Map<number, ParsedCharShape>();
  const paraShapes = new Map<number, ParsedParaShape>();
  const borderFills = new Map<number, ParsedBorderFill>();

  let docInfoData = getEntryData('/DocInfo');
  if (docInfoData) {
    let offset = 0;
    let binDataId = 1;
    let faceNameId = 0;
    let charShapeId = 0;
    let paraShapeId = 0;
    let borderFillId = 1;

    while (offset < docInfoData.length) {
      if (offset + 4 > docInfoData.length) break;
      const header = readUint32(docInfoData, offset);
      const tagId = header & 0x3FF;
      let size = (header >>> 20) & 0xFFF;
      let nextOffset = offset + 4;

      if (size === 0xFFF) {
        if (nextOffset + 4 > docInfoData.length) break;
        size = readUint32(docInfoData, nextOffset);
        nextOffset += 4;
      }

      const recordData = docInfoData.slice(nextOffset, nextOffset + size);

      if (tagId === HWP_TAGS.HWPTAG_BIN_DATA && recordData.length >= 2) {
        const p = readUint16(recordData, 0);
        const type = p & 0x0F;
        let ext = '';

        if (type === 1 && recordData.length > 10) {
          let extOffset = 2;
          if (recordData.length > extOffset + 2) extOffset += 2 + readUint16(recordData, extOffset) * 2;
          if (recordData.length > extOffset + 2) extOffset += 2 + readUint16(recordData, extOffset) * 2;
          extOffset += 2;
          if (recordData.length > extOffset + 2) {
            const extLen = readUint16(recordData, extOffset);
            extOffset += 2;
            if (recordData.length >= extOffset + extLen * 2) {
              ext = new TextDecoder('utf-16le').decode(recordData.slice(extOffset, extOffset + extLen * 2));
            }
          }
        }

        binDataInfos.set(binDataId, { id: binDataId, type: type === 0 ? 'LINK' : type === 1 ? 'EMBEDDING' : 'STORAGE', extension: ext.toLowerCase() });
        binDataId++;
      } else if (tagId === HWP_TAGS.HWPTAG_FACE_NAME && recordData.length >= 3) {
        const faceName = parseFaceNameStandalone(recordData);
        if (faceName) {
          faceNames.set(faceNameId, faceName);
        }
        faceNameId++;
      } else if (tagId === HWP_TAGS.HWPTAG_CHAR_SHAPE && recordData.length >= 72) {
        const charShape = parseCharShapeStandalone(recordData);
        if (charShape) {
          charShapes.set(charShapeId, charShape);
        }
        charShapeId++;
      } else if (tagId === HWP_TAGS.HWPTAG_PARA_SHAPE && recordData.length >= 42) {
        const paraShape = parseParaShapeStandalone(recordData);
        if (paraShape) {
          paraShapes.set(paraShapeId, paraShape);
        }
        paraShapeId++;
      } else if (tagId === HWP_TAGS.HWPTAG_BORDER_FILL && recordData.length >= 32) {
        const borderFill = parseBorderFillStandalone(recordData);
        if (borderFill) {
          borderFills.set(borderFillId, borderFill);
        }
        borderFillId++;
      }

      offset = nextOffset + size;
    }

  }

  for (const entry of cfb.FileIndex) {
    const fullPath = entry.name;
    if (!fullPath.startsWith('BIN') || !entry.content) continue;

    const match = fullPath.match(/BIN(\d+)/i);
    if (!match) continue;

    const binId = parseInt(match[1], 10);
    const binInfo = binDataInfos.get(binId);

    let imageData = entry.content instanceof Uint8Array ? entry.content : new Uint8Array(entry.content as unknown as ArrayBuffer);
    if (compressed) {
      try { imageData = pako.inflateRaw(imageData); } catch {}
    }

    let mimeType = 'image/png';
    if (imageData[0] === 0xFF && imageData[1] === 0xD8) mimeType = 'image/jpeg';
    else if (imageData[0] === 0x89 && imageData[1] === 0x50) mimeType = 'image/png';
    else if (imageData[0] === 0x47 && imageData[1] === 0x49) mimeType = 'image/gif';
    else if (imageData[0] === 0x42 && imageData[1] === 0x4D) mimeType = 'image/bmp';
    else if (binInfo?.extension) {
      const extMap: Record<string, string> = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', bmp: 'image/bmp' };
      mimeType = extMap[binInfo.extension] || mimeType;
    }

    const base64 = uint8ArrayToBase64(imageData);
    const dataUrl = `data:${mimeType};base64,${base64}`;
    const idStr = `BIN${String(binId).padStart(4, '0')}`;

    content.images.set(idStr, { id: idStr, binaryId: idStr, width: 200, height: 150, data: dataUrl, mimeType });
    content.binData.set(idStr, { id: idStr, data: dataUrl });
  }

  let sectionIndex = 0;
  while (true) {
    const sectionData = getEntryData(`/BodyText/Section${sectionIndex}`);
    if (!sectionData) break;
    content.sections.push(parseSectionData(sectionData, content.images, faceNames, charShapes, paraShapes, borderFills));
    sectionIndex++;
  }

  if (content.sections.length === 0) {
    content.sections.push({
      elements: [{ type: 'paragraph', data: { id: generateId(), runs: [{ text: '' }] } }],
      pageSettings: { width: 595, height: 842, marginTop: 56.7, marginBottom: 56.7, marginLeft: 56.7, marginRight: 56.7 },
    });
  }

  return content;
}
