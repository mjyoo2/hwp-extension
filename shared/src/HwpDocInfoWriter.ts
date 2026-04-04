/**
 * HWP DocInfo Stream Writer
 * Serializes document definition records (fonts, char shapes, para shapes, etc.)
 * Reverse of the DocInfo parsing logic in HwpParser.standalone.ts
 */

import {
  HWP_TAGS,
  TagStreamBuilder,
  BodyBuilder,
  BinDataInfo,
  ParsedFaceName,
  ParsedCharShape,
  ParsedParaShape,
  ParsedBorderFill,
} from './HwpTagBuilder';

// ============================================================
// Border width mapping: mm value → index (reverse of parser's lookup table)
// Parser: borderWidthMm = [0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0]
// ============================================================

const BORDER_WIDTH_MM = [0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0];

function borderWidthToIndex(widthMm: number): number {
  let closest = 0;
  let minDiff = Math.abs(BORDER_WIDTH_MM[0] - widthMm);
  for (let i = 1; i < BORDER_WIDTH_MM.length; i++) {
    const diff = Math.abs(BORDER_WIDTH_MM[i] - widthMm);
    if (diff < minDiff) {
      minDiff = diff;
      closest = i;
    }
  }
  return closest;
}

// ============================================================
// Individual record writers
// ============================================================

/**
 * Reverse of parseFaceNameStandalone() (HwpParser.standalone.ts:473)
 *
 * Layout:
 *   byte 0: property flags (bit7=hasSubstitute, bit6=hasFontTypeInfo, bit5=hasDefaultFont)
 *   bytes 1-2: name length (uint16)
 *   following: name in UTF-16LE
 *   if hasSubstitute: type byte + name length (uint16) + name (UTF-16LE)
 *   if hasFontTypeInfo: 10 zero bytes
 *   if hasDefaultFont: name length (uint16) + name (UTF-16LE)
 */
export function writeFaceName(face: ParsedFaceName): Uint8Array {
  const bb = new BodyBuilder();

  let props = 0;
  if (face.hasSubstitute) props |= 0x80;
  if (face.hasFontTypeInfo) props |= 0x40;
  if (face.hasDefaultFont) props |= 0x20;
  bb.addUint8(props);

  // Name as length-prefixed UTF-16LE
  bb.addHwpString(face.name);

  // Substitute font
  if (face.hasSubstitute && face.substitute) {
    const substType = face.substitute.type === 'truetype' ? 1 : face.substitute.type === 'hwp' ? 2 : 0;
    bb.addUint8(substType);
    bb.addHwpString(face.substitute.name);
  }

  // Font type info (10 zero bytes)
  if (face.hasFontTypeInfo) {
    bb.addZeros(10);
  }

  // Default font
  if (face.hasDefaultFont && face.defaultFont) {
    bb.addHwpString(face.defaultFont);
  }

  return bb.build();
}

/**
 * Reverse of parseCharShapeStandalone() (HwpParser.standalone.ts:525)
 *
 * Layout:
 *   0-13:  7 font IDs (uint16 each)
 *   14-20: 7 width ratios (uint8 each)
 *   21-27: 7 spacings (int8 each)
 *   28-34: 7 relative sizes (uint8 each)
 *   35-41: 7 char positions (int8 each)
 *   42:    baseSize (int32) - in 1/100 pt
 *   46:    properties (uint32) - bit flags
 *   50:    shadowOffsetX (int8)
 *   51:    shadowOffsetY (int8)
 *   52:    textColor (uint32)
 *   56:    underlineColor (uint32)
 *   60:    shadeColor (uint32)
 *   64:    shadowColor (uint32)
 *   68:    borderFillId (uint16) [optional]
 *   70:    strikethroughColor (uint32) [optional]
 *   Minimum: 72 bytes (includes borderFillId + 2 padding)
 */
export function writeCharShape(cs: ParsedCharShape): Uint8Array {
  const bb = new BodyBuilder();

  // 7 font IDs (uint16 each) - offsets 0-13
  for (let i = 0; i < 7; i++) {
    bb.addUint16(cs.fontIds[i] ?? 0);
  }

  // 7 width ratios (uint8 each) - offsets 14-20
  for (let i = 0; i < 7; i++) {
    bb.addUint8(cs.widthRatios[i] ?? 100);
  }

  // 7 spacings (int8 each) - offsets 21-27
  for (let i = 0; i < 7; i++) {
    const val = cs.spacings[i] ?? 0;
    bb.addUint8(val < 0 ? val + 256 : val);
  }

  // 7 relative sizes (uint8 each) - offsets 28-34
  for (let i = 0; i < 7; i++) {
    bb.addUint8(cs.relativeSizes[i] ?? 100);
  }

  // 7 char positions (int8 each) - offsets 35-41
  for (let i = 0; i < 7; i++) {
    const val = cs.charPositions[i] ?? 0;
    bb.addUint8(val < 0 ? val + 256 : val);
  }

  // baseSize (int32) at offset 42
  bb.addInt32(cs.baseSize);

  // Properties (uint32) at offset 46
  // Reverse of parser bit extraction (lines 556-570)
  let props = 0;
  if (cs.italic) props |= 0x01;
  if (cs.bold) props |= 0x02;
  props |= (cs.underlineType & 0x03) << 2;
  props |= (cs.underlineShape & 0x0F) << 4;
  props |= (cs.outlineType & 0x07) << 8;
  props |= (cs.shadowType & 0x03) << 11;
  if (cs.emboss) props |= (1 << 13);
  if (cs.engrave) props |= (1 << 14);
  if (cs.superscript) props |= (1 << 15);
  if (cs.subscript) props |= (1 << 16);
  props |= (cs.strikethrough & 0x07) << 18;
  props |= (cs.emphasisMark & 0x0F) << 21;
  if (cs.useFontSpacing) props |= (1 << 25);
  props |= (cs.strikethroughShape & 0x0F) << 26;
  if (cs.kerning) props |= (1 << 30);
  bb.addUint32(props >>> 0);

  // shadowOffsetX (int8) at offset 50
  const sox = cs.shadowOffsetX;
  bb.addUint8(sox < 0 ? sox + 256 : sox);

  // shadowOffsetY (int8) at offset 51
  const soy = cs.shadowOffsetY;
  bb.addUint8(soy < 0 ? soy + 256 : soy);

  // Colors at offsets 52, 56, 60, 64
  bb.addUint32(cs.textColor >>> 0);
  bb.addUint32(cs.underlineColor >>> 0);
  bb.addUint32(cs.shadeColor >>> 0);
  bb.addUint32(cs.shadowColor >>> 0);

  // borderFillId (uint16) at offset 68 - parser checks data.length >= 70
  bb.addUint16(cs.borderFillId ?? 0);

  // strikethroughColor (uint32) at offset 70 - parser checks data.length >= 74
  bb.addUint32((cs.strikethroughColor ?? 0) >>> 0);

  return bb.build();
}

/**
 * Reverse of parseParaShapeStandalone() (HwpParser.standalone.ts:619)
 *
 * Layout:
 *   0:  properties1 (uint32)
 *   4:  leftMargin (int32)
 *   8:  rightMargin (int32)
 *   12: indent (int32)
 *   16: spacingBefore (int32)
 *   20: spacingAfter (int32)
 *   24: lineSpacing (int32) - overridden by props3 if present
 *   28: tabDefId (uint16)
 *   30: numberingId (uint16)
 *   32: borderFillId (uint16)
 *   34: borderSpacing left (int16)
 *   36: borderSpacing right (int16)
 *   38: borderSpacing top (int16)
 *   40: borderSpacing bottom (int16)
 *   42: properties2 (uint32) - autoSpace flags
 *   46: properties3 (uint32) - lineSpacingType
 *   50: lineSpacing (uint32) - final lineSpacing value
 *   Total: 54 bytes
 */
export function writeParaShape(ps: ParsedParaShape): Uint8Array {
  const bb = new BodyBuilder();

  // Properties1 (uint32) at offset 0
  // Reverse of parser bit extraction (lines 624-636)
  // Note: lineSpacingType goes in props3 when writing full 54-byte format
  let props1 = 0;
  props1 |= (0 & 0x03);                              // lineSpacingTypeOld (set to 0, actual goes in props3)
  props1 |= (ps.alignment & 0x07) << 2;
  props1 |= (ps.wordBreakEnglish & 0x03) << 5;
  props1 |= (ps.wordBreakKorean & 0x01) << 7;
  if (ps.useGrid) props1 |= (1 << 8);
  props1 |= (ps.minSpace & 0x7F) << 9;
  if (ps.widowOrphan) props1 |= (1 << 16);
  if (ps.keepWithNext) props1 |= (1 << 17);
  if (ps.keepTogether) props1 |= (1 << 18);
  if (ps.pageBreakBefore) props1 |= (1 << 19);
  props1 |= (ps.verticalAlign & 0x03) << 20;
  props1 |= (ps.headType & 0x03) << 23;
  props1 |= (ps.level & 0x07) << 25;
  bb.addUint32(props1 >>> 0);

  // Margins and spacing
  bb.addInt32(ps.leftMargin);
  bb.addInt32(ps.rightMargin);
  bb.addInt32(ps.indent);
  bb.addInt32(ps.spacingBefore);
  bb.addInt32(ps.spacingAfter);
  bb.addInt32(ps.lineSpacing);  // old lineSpacing at offset 24

  // Tab, numbering, border
  bb.addUint16(ps.tabDefId);
  bb.addUint16(ps.numberingId);
  bb.addUint16(ps.borderFillId);

  // Border spacing
  bb.addInt16(ps.borderSpacing.left);
  bb.addInt16(ps.borderSpacing.right);
  bb.addInt16(ps.borderSpacing.top);
  bb.addInt16(ps.borderSpacing.bottom);

  // Properties2 (uint32) at offset 42
  let props2 = 0;
  if (ps.autoSpaceKoreanEnglish) props2 |= (1 << 4);
  if (ps.autoSpaceKoreanNumber) props2 |= (1 << 5);
  bb.addUint32(props2 >>> 0);

  // Properties3 (uint32) at offset 46 - lineSpacingType
  const props3 = ps.lineSpacingType & 0x1F;
  bb.addUint32(props3 >>> 0);

  // lineSpacing (uint32) at offset 50 - final value
  bb.addUint32(ps.lineSpacing >>> 0);

  return bb.build();
}

/**
 * Reverse of parseBorderFillStandalone() (HwpParser.standalone.ts:697)
 *
 * Layout:
 *   0:  properties (uint16): 3D(bit0) | shadow(bit1) | slashDiag<<2(3bit) | backslashDiag<<5(3bit)
 *   2:  4 border types (uint8 each): left, right, top, bottom
 *   6:  4 border widths (uint8 each): left, right, top, bottom
 *   10: 4 border colors (uint32 each): left, right, top, bottom
 *   26: diagonal type (uint8)
 *   27: diagonal width (uint8)
 *   28: diagonal color (uint32)
 *   32: fill type (uint32) - bit flags
 *   36: fill data (variable)
 */
export function writeBorderFill(bf: ParsedBorderFill): Uint8Array {
  const bb = new BodyBuilder();

  // Properties (uint16)
  let props = 0;
  if (bf.effect3d) props |= 0x01;
  if (bf.shadow) props |= 0x02;
  props |= (bf.slashDiagonal & 0x07) << 2;
  props |= (bf.backslashDiagonal & 0x07) << 5;
  bb.addUint16(props);

  // 4 border types (uint8 each)
  const borders = [bf.borders.left, bf.borders.right, bf.borders.top, bf.borders.bottom];
  for (const b of borders) {
    bb.addUint8(b.type);
  }

  // 4 border widths (uint8 each) - reverse of mm lookup
  for (const b of borders) {
    bb.addUint8(borderWidthToIndex(b.width));
  }

  // 4 border colors (uint32 each)
  for (const b of borders) {
    bb.addUint32(b.color >>> 0);
  }

  // Diagonal border
  if (bf.diagonal) {
    bb.addUint8(bf.diagonal.type);
    bb.addUint8(borderWidthToIndex(bf.diagonal.width));
    bb.addUint32(bf.diagonal.color >>> 0);
  } else {
    bb.addUint8(0);    // type = none
    bb.addUint8(0);    // width index = 0
    bb.addUint32(0);   // color = 0
  }

  // Fill
  if (bf.fill) {
    if (bf.fill.fillType === 'solid') {
      bb.addUint32(0x01);  // fill type flag: solid
      bb.addUint32(bf.fill.backgroundColor >>> 0);
      bb.addUint32(bf.fill.patternColor >>> 0);
      bb.addInt32(bf.fill.patternType);
    } else if (bf.fill.fillType === 'gradient') {
      bb.addUint32(0x04);  // fill type flag: gradient
      bb.addInt16(bf.fill.gradientType);
      bb.addInt16(bf.fill.angle);
      bb.addInt16(bf.fill.centerX);
      bb.addInt16(bf.fill.centerY);
      bb.addInt16(bf.fill.blur);
      bb.addInt16(bf.fill.colors.length);
      // Position values for intermediate colors
      const positionsCount = Math.max(0, bf.fill.colors.length - 2);
      for (let i = 0; i < positionsCount; i++) {
        bb.addUint32(bf.fill.colors[i + 1]?.position ?? 0);
      }
      // Color values
      for (const c of bf.fill.colors) {
        bb.addUint32(c.color >>> 0);
      }
    } else if (bf.fill.fillType === 'image') {
      bb.addUint32(0x02);  // fill type flag: image
      bb.addUint8(bf.fill.imageType);
      const brightness = bf.fill.brightness;
      bb.addUint8(brightness < 0 ? brightness + 256 : brightness);
      const contrast = bf.fill.contrast;
      bb.addUint8(contrast < 0 ? contrast + 256 : contrast);
      bb.addUint8(bf.fill.effect);
      bb.addUint16(bf.fill.binItemId);
    }
  } else {
    bb.addUint32(0);  // no fill
  }

  // Extended fill flag
  bb.addUint32(0);

  return bb.build();
}

/**
 * Reverse of BIN_DATA parsing (HwpParser.standalone.ts:1801)
 *
 * Layout:
 *   0: properties (uint16): type in bits 0-3 (0=LINK, 1=EMBEDDING, 2=STORAGE)
 *   For EMBEDDING: ext string as length-prefixed UTF-16LE
 */
export function writeBinData(info: BinDataInfo): Uint8Array {
  const bb = new BodyBuilder();

  const typeVal = info.type === 'LINK' ? 0 : info.type === 'EMBEDDING' ? 1 : 2;
  bb.addUint16(typeVal & 0x0F);

  if (info.type === 'EMBEDDING') {
    // For embedding type, write the extension string
    bb.addHwpString(info.extension);
  }

  return bb.build();
}

/**
 * Write DOCUMENT_PROPERTIES record body
 * Section count and starting numbers for page, footnote, endnote, figure, table, equation
 */
export function writeDocumentProperties(sectionCount: number): Uint8Array {
  const bb = new BodyBuilder();

  bb.addUint16(sectionCount);
  // Starting numbers: page, footnote, endnote, figure, table, equation
  bb.addUint16(1);  // page
  bb.addUint16(1);  // footnote
  bb.addUint16(1);  // endnote
  bb.addUint16(1);  // figure
  bb.addUint16(1);  // table
  bb.addUint16(1);  // equation

  return bb.build();
}

/**
 * Write ID_MAPPINGS record body
 * 9 counts as int32: binData, faceName, borderFill, charShape, tabDef, numbering, bullet, paraShape, style
 */
export function writeIdMappings(counts: {
  binData: number;
  faceName: number;
  borderFill: number;
  charShape: number;
  tabDef: number;
  numbering: number;
  bullet: number;
  paraShape: number;
  style: number;
}): Uint8Array {
  const bb = new BodyBuilder();

  bb.addInt32(counts.binData);
  bb.addInt32(counts.faceName);
  bb.addInt32(counts.borderFill);
  bb.addInt32(counts.charShape);
  bb.addInt32(counts.tabDef);
  bb.addInt32(counts.numbering);
  bb.addInt32(counts.bullet);
  bb.addInt32(counts.paraShape);
  bb.addInt32(counts.style);

  return bb.build();
}

/**
 * Write a default TAB_DEF record body
 * Minimal tab definition with no custom tab stops
 */
function writeDefaultTabDef(): Uint8Array {
  const bb = new BodyBuilder();
  // Properties (uint32): autoTab left-aligned
  bb.addUint32(0);
  // Tab stop count
  bb.addUint32(0);
  return bb.build();
}

/**
 * Write a default STYLE record body
 * Minimal style entry: normal paragraph style
 */
function writeDefaultStyle(): Uint8Array {
  const bb = new BodyBuilder();
  // Style name (length-prefixed UTF-16LE)
  bb.addHwpString('Normal');
  // English style name
  bb.addHwpString('Normal');
  // Properties (uint8): paragraph style = 0
  bb.addUint8(0);
  // Next style ID (uint8)
  bb.addUint8(0);
  // Language ID (int16)
  bb.addInt16(0);
  // Para shape ID (uint16)
  bb.addUint16(0);
  // Char shape ID (uint16)
  bb.addUint16(0);
  return bb.build();
}

// ============================================================
// Main DocInfo stream builder
// ============================================================

export interface DocInfoMaps {
  binData: BinDataInfo[];
  faceNames: ParsedFaceName[];
  charShapes: ParsedCharShape[];
  paraShapes: ParsedParaShape[];
  borderFills: ParsedBorderFill[];
}

/**
 * Build the complete DocInfo stream from document content maps.
 *
 * Record order (all at level 0):
 *   DOCUMENT_PROPERTIES(16) → ID_MAPPINGS(17) →
 *   BIN_DATA(18)×N → FACE_NAME(19)×N → BORDER_FILL(20)×N →
 *   CHAR_SHAPE(21)×N → TAB_DEF(22)×1 → PARA_SHAPE(25)×N → STYLE(26)×1
 */
export function buildDocInfoStream(sectionCount: number, maps: DocInfoMaps): Uint8Array {
  const stream = new TagStreamBuilder();

  // Ensure we have at least one tab def and one style
  const tabDefCount = 1;
  const styleCount = 1;

  // DOCUMENT_PROPERTIES
  stream.addRecord(
    HWP_TAGS.HWPTAG_DOCUMENT_PROPERTIES, 0,
    writeDocumentProperties(sectionCount),
  );

  // ID_MAPPINGS
  stream.addRecord(
    HWP_TAGS.HWPTAG_ID_MAPPINGS, 0,
    writeIdMappings({
      binData: maps.binData.length,
      faceName: maps.faceNames.length,
      borderFill: maps.borderFills.length,
      charShape: maps.charShapes.length,
      tabDef: tabDefCount,
      numbering: 0,
      bullet: 0,
      paraShape: maps.paraShapes.length,
      style: styleCount,
    }),
  );

  // BIN_DATA records
  for (const bin of maps.binData) {
    stream.addRecord(HWP_TAGS.HWPTAG_BIN_DATA, 0, writeBinData(bin));
  }

  // FACE_NAME records
  for (const face of maps.faceNames) {
    stream.addRecord(HWP_TAGS.HWPTAG_FACE_NAME, 0, writeFaceName(face));
  }

  // BORDER_FILL records
  for (const bf of maps.borderFills) {
    stream.addRecord(HWP_TAGS.HWPTAG_BORDER_FILL, 0, writeBorderFill(bf));
  }

  // CHAR_SHAPE records
  for (const cs of maps.charShapes) {
    stream.addRecord(HWP_TAGS.HWPTAG_CHAR_SHAPE, 0, writeCharShape(cs));
  }

  // TAB_DEF (one default)
  stream.addRecord(HWP_TAGS.HWPTAG_TAB_DEF, 0, writeDefaultTabDef());

  // PARA_SHAPE records
  for (const ps of maps.paraShapes) {
    stream.addRecord(HWP_TAGS.HWPTAG_PARA_SHAPE, 0, writeParaShape(ps));
  }

  // STYLE (one default)
  stream.addRecord(HWP_TAGS.HWPTAG_STYLE, 0, writeDefaultStyle());

  return stream.build();
}
