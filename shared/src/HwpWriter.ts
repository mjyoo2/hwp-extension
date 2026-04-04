/**
 * HWP Writer - Main entry point for writing HWP 5.0 binary files
 * Assembles OLE compound document from document content model
 */

import * as CFB from 'cfb';
import * as pako from 'pako';
import {
  HwpxContent,
  HwpxParagraph,
  TextRun,
  TableCell,
  CharacterStyle,
  ParagraphStyle,
  FontRef,
} from './types';
import {
  writeUint32LE,
  colorHexToRef,
  ptToHwpunit,
  base64ToUint8Array,
  concatUint8Arrays,
  ParsedFaceName,
  ParsedCharShape,
  ParsedParaShape,
  ParsedBorderFill,
  ParsedBorderLine,
  BinDataInfo,
} from './HwpTagBuilder';
import { buildDocInfoStream, DocInfoMaps } from './HwpDocInfoWriter';
import { buildSectionStream, SectionWriterMaps } from './HwpSectionWriter';

// ============================================================
// FileHeader (256 bytes)
// ============================================================

function buildFileHeader(compressed: boolean): Uint8Array {
  const header = new Uint8Array(256);

  // Signature: "HWP Document File" in null-terminated bytes (offset 0, 32 bytes)
  const sig = new TextEncoder().encode('HWP Document File');
  header.set(sig, 0);

  // Version: 5.1.0.0 at offset 32 (4 bytes, little-endian)
  // Version format: major.minor.build.revision packed as uint32
  // 5.1.0.0 = (5 << 24) | (1 << 16) | (0 << 8) | 0
  header[32] = 0;   // revision
  header[33] = 0;   // build
  header[34] = 1;   // minor
  header[35] = 5;   // major

  // Properties at offset 36 (4 bytes)
  let props = 0;
  if (compressed) props |= 0x01;
  header[36] = props & 0xFF;
  header[37] = (props >>> 8) & 0xFF;
  header[38] = (props >>> 16) & 0xFF;
  header[39] = (props >>> 24) & 0xFF;

  return header;
}

// ============================================================
// Map extraction: CharacterStyle/ParagraphStyle → Parsed shapes
// ============================================================

interface ExtractedMaps {
  faceNames: ParsedFaceName[];
  faceNameIndex: Map<string, number>;
  charShapes: ParsedCharShape[];
  charShapeIndex: Map<string, number>;
  paraShapes: ParsedParaShape[];
  paraShapeIndex: Map<string, number>;
  borderFills: ParsedBorderFill[];
  borderFillIndex: Map<string, number>;
  binData: BinDataInfo[];
  binDataIndex: Map<string, number>;
}

function charStyleKey(run: TextRun): string {
  const cs = run.charStyle;
  if (!cs) return 'default';
  return JSON.stringify(cs);
}

function paraStyleKey(para: HwpxParagraph): string {
  const ps = para.paraStyle;
  if (!ps) return 'default';
  return JSON.stringify(ps);
}

function cellBorderKey(cell: TableCell): string {
  // Key on actual border data + background color so cells with different styles get different entries
  const borders = {
    top: cell.borderTop ? JSON.stringify(cell.borderTop) : null,
    bottom: cell.borderBottom ? JSON.stringify(cell.borderBottom) : null,
    left: cell.borderLeft ? JSON.stringify(cell.borderLeft) : null,
    right: cell.borderRight ? JSON.stringify(cell.borderRight) : null,
    bg: cell.backgroundColor || null,
  };
  const hasBorders = borders.top || borders.bottom || borders.left || borders.right;
  const hasBg = !!cell.backgroundColor;
  if (!hasBorders && !hasBg) return 'default';
  return JSON.stringify(borders);
}

/**
 * Convert a CharacterStyle to a ParsedCharShape.
 * Reverse of charShapeToStyleStandalone() in HwpParser.standalone.ts
 */
function charStyleToShape(cs: CharacterStyle | undefined, faceNameIndex: Map<string, number>): ParsedCharShape {
  const fontId = cs?.fontName ? (faceNameIndex.get(cs.fontName) ?? 0) : 0;

  // Default font IDs: all 7 language slots point to the same font
  const fontIds = new Array(7).fill(fontId);
  const widthRatios = new Array(7).fill(100);
  const relativeSizes = new Array(7).fill(100);

  // Parse per-language spacings and offsets
  const spacings = new Array(7).fill(0);
  const charPositions = new Array(7).fill(0);
  const langKeys = ['hangul', 'latin', 'hanja', 'japanese', 'other', 'symbol', 'user'] as const;

  if (cs?.charSpacing && typeof cs.charSpacing === 'object') {
    const ref = cs.charSpacing as FontRef;
    for (let i = 0; i < 7; i++) {
      spacings[i] = (ref as any)[langKeys[i]] ?? 0;
    }
  }

  if (cs?.charOffset && typeof cs.charOffset === 'object') {
    const ref = cs.charOffset as FontRef;
    for (let i = 0; i < 7; i++) {
      charPositions[i] = (ref as any)[langKeys[i]] ?? 0;
    }
  }

  if (cs?.relativeSize && typeof cs.relativeSize === 'object') {
    const ref = cs.relativeSize as FontRef;
    for (let i = 0; i < 7; i++) {
      relativeSizes[i] = (ref as any)[langKeys[i]] ?? 100;
    }
  }

  const baseSize = cs?.fontSize ? Math.round(cs.fontSize * 100) : 1000; // default 10pt

  return {
    fontIds,
    widthRatios,
    spacings,
    relativeSizes,
    charPositions,
    baseSize,
    italic: cs?.italic ?? false,
    bold: cs?.bold ?? false,
    underlineType: cs?.underline
      ? (typeof cs.underline === 'object'
        ? ({ 'None': 0, 'none': 0, 'Bottom': 1, 'bottom': 1, 'Center': 2, 'center': 2, 'Top': 3, 'top': 3 } as Record<string, number>)[cs.underline.type] ?? 1
        : 1)
      : 0,
    underlineShape: cs?.underline && typeof cs.underline === 'object'
      ? ({ 'Solid': 0, 'Dash': 1, 'Dot': 2, 'DashDot': 3, 'DashDotDot': 4, 'LongDash': 5, 'CircleDot': 6, 'DoubleSlim': 7, 'SlimThick': 8, 'ThickSlim': 9, 'SlimThickSlim': 10 } as Record<string, number>)[cs.underline.shape] ?? 0
      : 0,
    outlineType: 0,
    shadowType: 0,
    emboss: cs?.emboss ?? false,
    engrave: cs?.engrave ?? false,
    superscript: cs?.superscript ?? false,
    subscript: cs?.subscript ?? false,
    strikethrough: cs?.strikethrough ? 2 : 0,
    emphasisMark: 0,
    useFontSpacing: cs?.useFontSpace ?? false,
    strikethroughShape: 0,
    kerning: cs?.useKerning ?? false,
    shadowOffsetX: cs?.shadowX ?? 0,
    shadowOffsetY: cs?.shadowY ?? 0,
    textColor: cs?.fontColor ? colorHexToRef(cs.fontColor) : 0x00000000,
    underlineColor: cs?.underline && typeof cs.underline === 'object' && cs.underline.color
      ? colorHexToRef(cs.underline.color)
      : 0x00000000,
    shadeColor: 0xFFFFFFFF,
    shadowColor: 0x00B2B2B2,
  };
}

/**
 * Convert a ParagraphStyle to a ParsedParaShape.
 * Reverse of parseParaShapeStandalone() in HwpParser.standalone.ts
 */
function paraStyleToShape(ps: ParagraphStyle | undefined): ParsedParaShape {
  const alignMap: Record<string, number> = {
    'justify': 0, 'Justify': 0,
    'left': 1, 'Left': 1,
    'right': 2, 'Right': 2,
    'center': 3, 'Center': 3,
    'distribute': 4, 'Distribute': 4,
    'DistributeSpace': 5,
  };

  const lineSpacingTypeMap: Record<string, number> = {
    'percent': 0,
    'fixed': 1,
    'betweenLines': 2,
    'atLeast': 3,
  };

  return {
    alignment: ps?.align ? (alignMap[ps.align] ?? 0) : 0,
    leftMargin: ps?.marginLeft ? ptToHwpunit(ps.marginLeft) : 0,
    rightMargin: ps?.marginRight ? ptToHwpunit(ps.marginRight) : 0,
    indent: ps?.firstLineIndent ? ptToHwpunit(ps.firstLineIndent) : 0,
    spacingBefore: ps?.marginTop ? ptToHwpunit(ps.marginTop) : 0,
    spacingAfter: ps?.marginBottom ? ptToHwpunit(ps.marginBottom) : 0,
    lineSpacing: ps?.lineSpacing ? Math.round(ps.lineSpacing * 100) : 16000, // default 160%
    lineSpacingType: ps?.lineSpacingType ? (lineSpacingTypeMap[ps.lineSpacingType] ?? 0) : 0,
    tabDefId: ps?.tabDefId ?? 0,
    numberingId: 0,
    borderFillId: ps?.borderFillId ?? 0,
    borderSpacing: { left: 0, right: 0, top: 0, bottom: 0 },
    wordBreakEnglish: 0,
    wordBreakKorean: 0,
    widowOrphan: ps?.widowControl ?? false,
    keepWithNext: ps?.keepWithNext ?? false,
    keepTogether: ps?.keepLines ?? false,
    pageBreakBefore: ps?.pageBreakBefore ?? false,
    verticalAlign: 0,
    headType: 0,
    level: 0,
    useGrid: ps?.snapToGrid ?? true,
    minSpace: 0,
    autoSpaceKoreanEnglish: ps?.autoSpaceEAsianEng ?? false,
    autoSpaceKoreanNumber: ps?.autoSpaceEAsianNum ?? false,
  };
}

/**
 * Build a default ParsedBorderFill (no borders, no fill).
 */
function defaultBorderFill(): ParsedBorderFill {
  const noBorder: ParsedBorderLine = { type: 0, width: 0.1, color: 0x00000000 };
  return {
    effect3d: false,
    shadow: false,
    slashDiagonal: 0,
    backslashDiagonal: 0,
    borders: {
      left: { ...noBorder },
      right: { ...noBorder },
      top: { ...noBorder },
      bottom: { ...noBorder },
    },
  };
}

/**
 * Walk all content to extract unique face names, char shapes, para shapes,
 * border fills, and bin data entries.
 */
function extractMapsFromContent(content: HwpxContent): ExtractedMaps {
  const faceNameIndex = new Map<string, number>();
  const faceNames: ParsedFaceName[] = [];
  const charShapeIndex = new Map<string, number>();
  const charShapes: ParsedCharShape[] = [];
  const paraShapeIndex = new Map<string, number>();
  const paraShapes: ParsedParaShape[] = [];
  const borderFillIndex = new Map<string, number>();
  const borderFills: ParsedBorderFill[] = [];
  const binDataIndex = new Map<string, number>();
  const binData: BinDataInfo[] = [];

  // Helper: register a face name, return its index
  function ensureFaceName(name: string): number {
    if (faceNameIndex.has(name)) return faceNameIndex.get(name)!;
    const id = faceNames.length;
    faceNames.push({
      name,
      hasSubstitute: false,
      hasFontTypeInfo: false,
      hasDefaultFont: false,
    });
    faceNameIndex.set(name, id);
    return id;
  }

  // Always register a default font
  ensureFaceName('함초롬바탕');

  // Collect all binary data from images
  for (const [key, img] of Array.from(content.images.entries())) {
    if (!binDataIndex.has(key)) {
      const id = binData.length + 1; // 1-based
      let ext = 'png';
      if (img.mimeType?.includes('jpeg') || img.mimeType?.includes('jpg')) ext = 'jpg';
      else if (img.mimeType?.includes('gif')) ext = 'gif';
      else if (img.mimeType?.includes('bmp')) ext = 'bmp';
      binData.push({ id, type: 'EMBEDDING', extension: ext });
      binDataIndex.set(key, id);
    }
  }

  // Walk all sections to collect styles
  function processRun(run: TextRun): void {
    if (run.charStyle?.fontName) {
      ensureFaceName(run.charStyle.fontName);
    }
    const key = charStyleKey(run);
    if (!charShapeIndex.has(key)) {
      const id = charShapes.length;
      charShapes.push(charStyleToShape(run.charStyle, faceNameIndex));
      charShapeIndex.set(key, id);
    }
  }

  function processParagraph(para: HwpxParagraph): void {
    const key = paraStyleKey(para);
    if (!paraShapeIndex.has(key)) {
      const id = paraShapes.length;
      paraShapes.push(paraStyleToShape(para.paraStyle));
      paraShapeIndex.set(key, id);
    }
    for (const run of para.runs) {
      processRun(run);
    }
  }

  function borderStyleToLine(border: TableCell['borderTop']): ParsedBorderLine {
    if (!border) return { type: 0, width: 0.1, color: 0x00000000 };
    // Map LineType1 string to HWP border type number
    const typeMap: Record<string, number> = {
      'none': 0, 'None': 0,
      'solid': 1, 'Solid': 1,
      'dash': 2, 'Dash': 2,
      'dot': 3, 'Dot': 3,
      'dashDot': 4, 'DashDot': 4,
      'dashDotDot': 5, 'DashDotDot': 5,
      'longDash': 6, 'LongDash': 6,
      'circleDot': 7, 'CircleDot': 7,
      'doubleSlim': 8, 'DoubleSlim': 8,
      'slimThick': 9, 'SlimThick': 9,
      'thickSlim': 10, 'ThickSlim': 10,
      'slimThickSlim': 11, 'SlimThickSlim': 11,
    };
    const styleStr = (border as any).style ?? (border as any).type;
    const borderType = typeof styleStr === 'string' ? (typeMap[styleStr] ?? 1) : 1;
    // border.width is in pt (from HWPX/HWP parser), convert to mm for HWP binary
    const widthPt = typeof border.width === 'number' ? border.width : parseFloat(String(border.width)) || 0.1;
    const widthMm = widthPt / 2.83465;
    const color = border.color ? colorHexToRef(border.color) : 0x00000000;
    return { type: borderType, width: widthMm, color };
  }

  function cellBorderFill(cell: TableCell): ParsedBorderFill {
    const hasBorders = cell.borderTop || cell.borderBottom || cell.borderLeft || cell.borderRight;
    const bf: ParsedBorderFill = {
      effect3d: false,
      shadow: false,
      slashDiagonal: 0,
      backslashDiagonal: 0,
      borders: hasBorders ? {
        left: borderStyleToLine(cell.borderLeft),
        right: borderStyleToLine(cell.borderRight),
        top: borderStyleToLine(cell.borderTop),
        bottom: borderStyleToLine(cell.borderBottom),
      } : defaultBorderFill().borders!,
    };
    // Preserve cell background color
    if (cell.backgroundColor) {
      bf.fill = {
        fillType: 'solid',
        backgroundColor: colorHexToRef(cell.backgroundColor),
        patternColor: 0xFFFFFFFF,
        patternType: -1,
      };
    }
    return bf;
  }

  function processTable(table: any): void {
    if (table.rows) {
      for (const row of table.rows) {
        for (const cell of row.cells) {
          processCell(cell);
        }
      }
    }
  }

  function processCell(cell: TableCell): void {
    const key = cellBorderKey(cell);
    if (!borderFillIndex.has(key)) {
      const id = borderFills.length;
      borderFills.push(cellBorderFill(cell));
      borderFillIndex.set(key, id);
    }
    // Process cell elements (includes nested tables)
    if (cell.elements) {
      for (const elem of cell.elements) {
        if (elem.type === 'paragraph') processParagraph(elem.data as HwpxParagraph);
        else if (elem.type === 'table') processTable(elem.data);
      }
    }
    if (cell.paragraphs) {
      for (const para of cell.paragraphs) {
        processParagraph(para);
      }
    }
  }

  for (const section of content.sections) {
    for (const element of section.elements) {
      switch (element.type) {
        case 'paragraph':
          processParagraph(element.data as HwpxParagraph);
          break;
        case 'table': {
          processTable(element.data as any);
          break;
        }
        case 'image':
          // Image elements may contain inline text
          break;
      }
    }

    // Process header/footer paragraphs
    if (section.header?.paragraphs) {
      for (const para of section.header.paragraphs) {
        processParagraph(para);
      }
    }
    if (section.footer?.paragraphs) {
      for (const para of section.footer.paragraphs) {
        processParagraph(para);
      }
    }
  }

  // Ensure at least one char shape and para shape exist
  if (charShapes.length === 0) {
    charShapes.push(charStyleToShape(undefined, faceNameIndex));
    charShapeIndex.set('default', 0);
  }
  if (paraShapes.length === 0) {
    paraShapes.push(paraStyleToShape(undefined));
    paraShapeIndex.set('default', 0);
  }
  // Ensure at least one border fill
  if (borderFills.length === 0) {
    borderFills.push(defaultBorderFill());
    borderFillIndex.set('default', 0);
  }

  return {
    faceNames,
    faceNameIndex,
    charShapes,
    charShapeIndex,
    paraShapes,
    paraShapeIndex,
    borderFills,
    borderFillIndex,
    binData,
    binDataIndex,
  };
}

// ============================================================
// Create SectionWriterMaps from ExtractedMaps
// ============================================================

function createSectionWriterMaps(extracted: ExtractedMaps, content?: HwpxContent): SectionWriterMaps {
  return {
    getCharShapeId: (run: TextRun) => {
      const key = charStyleKey(run);
      return extracted.charShapeIndex.get(key) ?? 0;
    },
    getParaShapeId: (para: HwpxParagraph) => {
      const key = paraStyleKey(para);
      return extracted.paraShapeIndex.get(key) ?? 0;
    },
    getBorderFillId: (cell: TableCell) => {
      const key = cellBorderKey(cell);
      // BorderFill IDs are 1-based in HWP format
      return (extracted.borderFillIndex.get(key) ?? 0) + 1;
    },
    getBinItemId: (binaryId: string) => {
      return extracted.binDataIndex.get(binaryId) ?? 1;
    },
    getFootnote: content ? (refNumber: number) => {
      return content.footnotes.find(fn => fn.number === refNumber);
    } : undefined,
    getEndnote: content ? (refNumber: number) => {
      return content.endnotes.find(en => en.number === refNumber);
    } : undefined,
  };
}

// ============================================================
// Extract binary image data from content
// ============================================================

function extractImageData(content: HwpxContent): Map<number, Uint8Array> {
  const result = new Map<number, Uint8Array>();

  for (const [key, img] of Array.from(content.images.entries())) {
    if (!img.data) continue;

    // Parse data URL: "data:image/png;base64,..."
    const match = img.data.match(/^data:[^;]+;base64,(.+)$/);
    if (!match) continue;

    const binId = parseInt(key.replace(/\D/g, ''), 10);
    if (isNaN(binId)) continue;

    result.set(binId, base64ToUint8Array(match[1]));
  }

  return result;
}

// ============================================================
// Main: writeHwpContent
// ============================================================

/**
 * Convert HwpxContent to a HWP 5.0 binary file (Uint8Array).
 *
 * OLE compound file structure:
 *   /FileHeader          - 256-byte file header
 *   /DocInfo             - Document info stream (compressed)
 *   /BodyText/Section0   - Section 0 stream (compressed)
 *   /BodyText/Section1   - Section 1 stream (compressed)
 *   ...
 *   /BinData/BIN0001.xxx - Binary data (images, compressed)
 *   ...
 */
export function writeHwpContent(content: HwpxContent): Uint8Array {
  const compressed = true;

  // 1. Extract maps from content
  const extracted = extractMapsFromContent(content);

  // 2. Build DocInfo stream
  const docInfoMaps: DocInfoMaps = {
    binData: extracted.binData,
    faceNames: extracted.faceNames,
    charShapes: extracted.charShapes,
    paraShapes: extracted.paraShapes,
    borderFills: extracted.borderFills,
  };
  const docInfoStream = buildDocInfoStream(content.sections.length, docInfoMaps);

  // 3. Build section streams
  const sectionWriterMaps = createSectionWriterMaps(extracted, content);
  const sectionStreams: Uint8Array[] = [];
  for (const section of content.sections) {
    sectionStreams.push(buildSectionStream(section, sectionWriterMaps));
  }

  // 4. Extract image binary data
  const imageData = extractImageData(content);

  // 5. Compress helper
  const compress = (data: Uint8Array): Uint8Array => {
    if (!compressed) return data;
    return pako.deflateRaw(data);
  };

  // 6. Assemble OLE compound file
  const cfb = CFB.utils.cfb_new();

  // FileHeader (NOT compressed)
  CFB.utils.cfb_add(cfb, '/FileHeader', buildFileHeader(compressed));

  // DocInfo (compressed)
  CFB.utils.cfb_add(cfb, '/DocInfo', compress(docInfoStream));

  // BodyText/SectionN (compressed)
  for (let i = 0; i < sectionStreams.length; i++) {
    CFB.utils.cfb_add(cfb, `/BodyText/Section${i}`, compress(sectionStreams[i]));
  }

  // BinData entries at ROOT level with decimal-padded IDs: BIN0001.ext, BIN0010.ext, etc.
  // The parser uses /BIN(\d+)/i regex (decimal only), so we use decimal padding.
  // Entries are at root level (not under /BinData/) — parser checks entry.name.startsWith('BIN').
  for (const [binId, data] of Array.from(imageData.entries())) {
    const info = extracted.binData.find(b => b.id === binId);
    const ext = info?.extension ?? 'png';
    const decId = String(binId).padStart(4, '0');
    const name = `BIN${decId}.${ext}`;
    CFB.utils.cfb_add(cfb, `/${name}`, compress(data));
  }

  // 7. Write to buffer
  const output = CFB.write(cfb, { type: 'array' });

  // CFB.write with type 'array' returns number[] — convert to Uint8Array
  if (output instanceof Uint8Array) {
    return output;
  }
  if (Array.isArray(output)) {
    return new Uint8Array(output);
  }
  if (output instanceof ArrayBuffer) {
    return new Uint8Array(output);
  }
  // Buffer (Node.js)
  return new Uint8Array(output as any);
}
