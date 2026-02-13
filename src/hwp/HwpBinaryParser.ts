import * as pako from 'pako';
import { OleReader } from './OleReader';
import {
  HwpxContent,
  HwpxSection,
  HwpxParagraph,
  TextRun,
  SectionElement,
  PageSettings,
  HwpxTable,
  TableRow,
  TableCell,
  CharacterStyle,
  ParagraphStyle,
  HwpxImage,
  BinData,
} from '../hwpx/types';

const HWPTAG_BEGIN = 0x010;
const HWPTAG_DOCUMENT_PROPERTIES = HWPTAG_BEGIN;
const HWPTAG_ID_MAPPINGS = HWPTAG_BEGIN + 1;
const HWPTAG_BIN_DATA = HWPTAG_BEGIN + 2;
const HWPTAG_FACE_NAME = HWPTAG_BEGIN + 3;
const HWPTAG_BORDER_FILL = HWPTAG_BEGIN + 4;
const HWPTAG_CHAR_SHAPE = HWPTAG_BEGIN + 5;
const HWPTAG_TAB_DEF = HWPTAG_BEGIN + 6;
const HWPTAG_NUMBERING = HWPTAG_BEGIN + 7;
const HWPTAG_BULLET = HWPTAG_BEGIN + 8;
const HWPTAG_PARA_SHAPE = HWPTAG_BEGIN + 9;
const HWPTAG_STYLE = HWPTAG_BEGIN + 10;

const HWPTAG_PARA_HEADER = 0x042;
const HWPTAG_PARA_TEXT = 0x043;
const HWPTAG_PARA_CHAR_SHAPE = 0x044;
const HWPTAG_PARA_LINE_SEG = 0x045;
const HWPTAG_PARA_RANGE_TAG = 0x046;
const HWPTAG_CTRL_HEADER = 0x047;
const HWPTAG_LIST_HEADER = 0x048;
const HWPTAG_PAGE_DEF = 0x049;
const HWPTAG_FOOTNOTE_SHAPE = 0x04A;
const HWPTAG_PAGE_BORDER_FILL = 0x04B;

const HWPTAG_TABLE = 0x04D;
const HWPTAG_SHAPE_COMPONENT = 0x04C;
const HWPTAG_SHAPE_COMPONENT_PICTURE = 0x055;

interface FontFace {
  name: string;
  type: number;
}

interface CharShape {
  fontId: number[];
  fontRatio: number[];
  fontSpacing: number[];
  fontRelSize: number[];
  fontOffset: number[];
  height: number;
  textColor: number;
  shadeColor: number;
  useFontSpace: boolean;
  useKerning: boolean;
  bold: boolean;
  italic: boolean;
  underline: number;
  strikeout: number;
  shadowType: number;
  emboss: boolean;
  engrave: boolean;
  superscript: boolean;
  subscript: boolean;
}

interface ParaShape {
  align: number;
  marginLeft: number;
  marginRight: number;
  indent: number;
  marginTop: number;
  marginBottom: number;
  lineSpacing: number;
  lineSpacingType: number;
  tabDefId: number;
  breakLatinWord: number;
  breakNonLatinWord: number;
  widowOrphan: boolean;
  keepWithNext: boolean;
  keepLines: boolean;
  pageBreakBefore: boolean;
  fontLineHeight: boolean;
  snapToGrid: boolean;
}

interface BorderLine {
  type: number;   // 0=none, 1=solid, etc.
  width: number;  // 0-15
  color: number;  // COLORREF
}

interface BorderFillFill {
  fillType: 'solid' | 'gradient' | 'image';
  backgroundColor?: number;  // COLORREF for solid fills
  patternColor?: number;
  patternType?: number;
  gradientType?: number;
  angle?: number;
  centerX?: number;
  centerY?: number;
  blur?: number;
  colors?: Array<{ position?: number; color: number }>;
  imageType?: number;
  brightness?: number;
  contrast?: number;
  effect?: number;
  binItemId?: number;
}

interface BorderFill {
  effect3d: boolean;
  shadow: boolean;
  slashDiagonal: number;
  backslashDiagonal: number;
  borders: {
    left: BorderLine;
    right: BorderLine;
    top: BorderLine;
    bottom: BorderLine;
  };
  diagonal?: BorderLine;
  fill?: BorderFillFill;
}

interface RecordHeader {
  tagId: number;
  level: number;
  size: number;
}

interface HwpFileHeader {
  signature: string;
  version: string;
  compressed: boolean;
  encrypted: boolean;
  distributed: boolean;
}

export class HwpBinaryParser {
  private ole: OleReader;
  private fileHeader: HwpFileHeader | null = null;
  private fontFaces: FontFace[] = [];
  private charShapes: CharShape[] = [];
  private paraShapes: ParaShape[] = [];
  private borderFills: BorderFill[] = [];
  private binDataItems: Map<string, BinData> = new Map();

  static parse(data: Uint8Array): HwpxContent {
    const parser = new HwpBinaryParser(data);
    return parser.toHwpxContent();
  }

  constructor(buffer: Uint8Array) {
    this.ole = new OleReader(buffer);
    this.parseFileHeader();
    this.parseDocInfo();
    this.parseBinData();
  }

  private parseFileHeader(): void {
    const headerData = this.ole.readStreamByName('FileHeader');
    if (!headerData) {
      throw new Error('FileHeader stream not found');
    }

    const view = new DataView(headerData.buffer, headerData.byteOffset, headerData.byteLength);
    
    let signature = '';
    for (let i = 0; i < 32; i++) {
      const byte = headerData[i];
      if (byte === 0) break;
      signature += String.fromCharCode(byte);
    }

    const versionDword = view.getUint32(32, true);
    const major = (versionDword >> 24) & 0xFF;
    const minor = (versionDword >> 16) & 0xFF;
    const build = (versionDword >> 8) & 0xFF;
    const revision = versionDword & 0xFF;

    const flags = view.getUint32(36, true);

    this.fileHeader = {
      signature,
      version: `${major}.${minor}.${build}.${revision}`,
      compressed: (flags & 0x01) !== 0,
      encrypted: (flags & 0x02) !== 0,
      distributed: (flags & 0x04) !== 0,
    };

    if (this.fileHeader.encrypted) {
      throw new Error('Encrypted HWP files are not supported');
    }
  }

  private decompressStream(data: Uint8Array): Uint8Array {
    if (!this.fileHeader?.compressed) {
      return data;
    }
    try {
      return pako.inflateRaw(data);
    } catch (e) {
      try {
        return pako.inflate(data);
      } catch (e2) {
        return data;
      }
    }
  }

  private parseDocInfo(): void {
    const docInfoData = this.ole.readStreamByName('DocInfo');
    if (!docInfoData) {
      throw new Error('DocInfo stream not found');
    }

    const decompressed = this.decompressStream(docInfoData);
    this.parseRecords(decompressed, this.handleDocInfoRecord.bind(this));
  }

  private parseBinData(): void {
    const streams = this.ole.listStreams();
    
    for (const stream of streams) {
      if (stream.name.startsWith('BIN') && stream.type === 'stream') {
        const match = stream.name.match(/BIN(\d+)/i);
        if (match) {
          const binId = match[1];
          const rawData = this.ole.readStreamByName(stream.name);
          if (rawData) {
            const decompressed = this.decompressStream(rawData);
            const base64Data = this.uint8ArrayToBase64(decompressed);
            this.binDataItems.set(binId, {
              id: binId,
              size: decompressed.length,
              encoding: 'Base64',
              data: base64Data,
            });
          }
        }
      }
    }
  }

  private parseRecords(data: Uint8Array, handler: (header: RecordHeader, data: Uint8Array) => void): void {
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    let offset = 0;

    while (offset < data.length) {
      if (offset + 4 > data.length) break;

      const headerDword = view.getUint32(offset, true);
      const tagId = headerDword & 0x3FF;
      const level = (headerDword >> 10) & 0x3FF;
      let size = (headerDword >> 20) & 0xFFF;

      offset += 4;

      if (size === 0xFFF) {
        if (offset + 4 > data.length) break;
        size = view.getUint32(offset, true);
        offset += 4;
      }

      if (offset + size > data.length) break;

      const recordData = data.slice(offset, offset + size);
      handler({ tagId, level, size }, recordData);

      offset += size;
    }
  }

  private handleDocInfoRecord(header: RecordHeader, data: Uint8Array): void {
    switch (header.tagId) {
      case HWPTAG_FACE_NAME:
        this.parseFaceName(data);
        break;
      case HWPTAG_CHAR_SHAPE:
        this.parseCharShape(data);
        break;
      case HWPTAG_PARA_SHAPE:
        this.parseParaShape(data);
        break;
      case HWPTAG_BORDER_FILL:
        this.parseBorderFill(data);
        break;
    }
  }

  private parseFaceName(data: Uint8Array): void {
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    const flags = view.getUint8(0);
    const nameLength = view.getUint16(1, true);
    
    let name = '';
    for (let i = 0; i < nameLength; i++) {
      const charCode = view.getUint16(3 + i * 2, true);
      if (charCode === 0) break;
      name += String.fromCharCode(charCode);
    }

    this.fontFaces.push({
      name,
      type: flags,
    });
  }

  private parseCharShape(data: Uint8Array): void {
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    
    const fontId: number[] = [];
    const fontRatio: number[] = [];
    const fontSpacing: number[] = [];
    const fontRelSize: number[] = [];
    const fontOffset: number[] = [];

    for (let i = 0; i < 7; i++) {
      fontId.push(view.getUint16(i * 2, true));
    }
    for (let i = 0; i < 7; i++) {
      fontRatio.push(view.getUint8(14 + i));
    }
    for (let i = 0; i < 7; i++) {
      fontSpacing.push(view.getInt8(21 + i));
    }
    for (let i = 0; i < 7; i++) {
      fontRelSize.push(view.getUint8(28 + i));
    }
    for (let i = 0; i < 7; i++) {
      fontOffset.push(view.getInt8(35 + i));
    }

    const height = view.getInt32(42, true);
    const props = view.getUint32(46, true);
    const shadowGap1 = view.getInt8(50);
    const shadowGap2 = view.getInt8(51);
    const textColor = view.getUint32(52, true);
    const underlineColor = view.getUint32(56, true);
    const shadeColor = view.getUint32(60, true);
    const shadowColor = view.getUint32(64, true);

    this.charShapes.push({
      fontId,
      fontRatio,
      fontSpacing,
      fontRelSize,
      fontOffset,
      height,
      textColor,
      shadeColor,
      useFontSpace: (props & 0x40000000) !== 0,
      useKerning: (props & 0x20000000) !== 0,
      bold: (props & 0x02) !== 0,
      italic: (props & 0x01) !== 0,
      underline: (props >> 2) & 0x03,
      strikeout: (props >> 18) & 0x07,
      shadowType: (props >> 9) & 0x07,
      emboss: (props & 0x4000) !== 0,
      engrave: (props & 0x8000) !== 0,
      superscript: (props & 0x10000) !== 0,
      subscript: (props & 0x20000) !== 0,
    });
  }

  private parseParaShape(data: Uint8Array): void {
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    
    const props1 = view.getUint32(0, true);
    const marginLeft = view.getInt32(4, true);
    const marginRight = view.getInt32(8, true);
    const indent = view.getInt32(12, true);
    const marginTop = view.getInt32(16, true);
    const marginBottom = view.getInt32(20, true);
    const lineSpacing = view.getInt32(24, true);
    const tabDefId = view.getUint16(28, true);
    const numParaHeadId = view.getUint16(30, true);
    const borderFillId = view.getUint16(32, true);
    const borderOffsetLeft = view.getInt16(34, true);
    const borderOffsetRight = view.getInt16(36, true);
    const borderOffsetTop = view.getInt16(38, true);
    const borderOffsetBottom = view.getInt16(40, true);
    const props2 = view.getUint32(42, true);
    const props3 = data.length > 46 ? view.getUint32(46, true) : 0;

    this.paraShapes.push({
      align: props1 & 0x07,
      marginLeft,
      marginRight,
      indent,
      marginTop,
      marginBottom,
      lineSpacing,
      lineSpacingType: (props1 >> 4) & 0x0F,
      tabDefId,
      breakLatinWord: (props1 >> 8) & 0x03,
      breakNonLatinWord: (props1 >> 10) & 0x01,
      widowOrphan: (props1 & 0x1000) !== 0,
      keepWithNext: (props1 & 0x2000) !== 0,
      keepLines: (props1 & 0x4000) !== 0,
      pageBreakBefore: (props1 & 0x8000) !== 0,
      fontLineHeight: (props1 & 0x10000) !== 0,
      snapToGrid: (props1 & 0x20000) !== 0,
    });
  }

  private parseBorderFill(data: Uint8Array): void {
    if (data.length < 32) return;
    
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    
    const props = view.getUint16(0, true);
    const effect3d = (props & 0x01) !== 0;
    const shadow = (props & 0x02) !== 0;
    const slashDiagonal = (props >> 2) & 0x07;
    const backslashDiagonal = (props >> 5) & 0x07;
    
    const borderTypes = [data[2], data[3], data[4], data[5]];
    const borderWidths = [data[6], data[7], data[8], data[9]];
    const borderColors = [
      view.getUint32(10, true),
      view.getUint32(14, true),
      view.getUint32(18, true),
      view.getUint32(22, true),
    ];
    
    const diagonalType = data[26];
    const diagonalWidth = data[27];
    const diagonalColor = view.getUint32(28, true);
    
    const borderFill: BorderFill = {
      effect3d,
      shadow,
      slashDiagonal,
      backslashDiagonal,
      borders: {
        left: { type: borderTypes[0], width: borderWidths[0], color: borderColors[0] },
        right: { type: borderTypes[1], width: borderWidths[1], color: borderColors[1] },
        top: { type: borderTypes[2], width: borderWidths[2], color: borderColors[2] },
        bottom: { type: borderTypes[3], width: borderWidths[3], color: borderColors[3] },
      },
    };
    
    if (diagonalType !== 0) {
      borderFill.diagonal = { type: diagonalType, width: diagonalWidth, color: diagonalColor };
    }
    
    if (data.length >= 36) {
      const fillType = view.getUint32(32, true);
      let fillOffset = 36;
      
      if (fillType & 0x01) {
        if (data.length >= fillOffset + 12) {
          borderFill.fill = {
            fillType: 'solid',
            backgroundColor: view.getUint32(fillOffset, true),
            patternColor: view.getUint32(fillOffset + 4, true),
            patternType: view.getInt32(fillOffset + 8, true),
          };
          fillOffset += 12;
        }
      } else if (fillType & 0x04) {
        if (data.length >= fillOffset + 12) {
          const gradientType = view.getInt16(fillOffset, true);
          const angle = view.getInt16(fillOffset + 2, true);
          const centerX = view.getInt16(fillOffset + 4, true);
          const centerY = view.getInt16(fillOffset + 6, true);
          const blur = view.getInt16(fillOffset + 8, true);
          const numColors = view.getInt16(fillOffset + 10, true);
          fillOffset += 12;
          
          const positionsCount = Math.max(0, numColors - 2);
          fillOffset += positionsCount * 4;
          
          const colors: Array<{ position?: number; color: number }> = [];
          for (let i = 0; i < numColors && fillOffset + 4 <= data.length; i++) {
            colors.push({ color: view.getUint32(fillOffset, true) });
            fillOffset += 4;
          }
          
          borderFill.fill = {
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
          borderFill.fill = {
            fillType: 'image',
            imageType: data[fillOffset],
            brightness: view.getInt8(fillOffset + 1),
            contrast: view.getInt8(fillOffset + 2),
            effect: data[fillOffset + 3],
            binItemId: view.getUint16(fillOffset + 4, true),
          };
        }
      }
    }
    
    this.borderFills.push(borderFill);
  }

  private parseSections(): HwpxSection[] {
    const sections: HwpxSection[] = [];
    let sectionIndex = 0;

    while (true) {
      const sectionName = `Section${sectionIndex}`;
      const sectionData = this.ole.readStreamByName(sectionName);
      
      if (!sectionData) {
        const altName = `BodyText/Section${sectionIndex}`;
        const altData = this.ole.readStreamByPath(altName);
        if (!altData) break;
        const decompressed = this.decompressStream(altData);
        sections.push(this.parseSection(decompressed));
      } else {
        const decompressed = this.decompressStream(sectionData);
        sections.push(this.parseSection(decompressed));
      }
      
      sectionIndex++;
    }

    return sections;
  }

  private parseSection(data: Uint8Array): HwpxSection {
    const section: HwpxSection = {
      elements: [],
      pageSettings: this.getDefaultPageSettings(),
    };

    const elements: SectionElement[] = [];
    
    let currentParagraph: Partial<HwpxParagraph> | null = null;
    let currentText = '';
    let currentCharShapePositions: { pos: number; shapeIndex: number }[] = [];
    let currentLevel = 0;
    let currentParaInHeaderFooter = false;
    let currentParaTableDepth = 0;
    
    interface TableContext {
      table: Partial<HwpxTable>;
      level: number;
      rowCount: number;
      colCount: number;
      cellIndex: number;
      cellParagraphs: HwpxParagraph[];
      cellNestedTables: HwpxTable[];
      isTopLevel: boolean;
      tableWidth: number;
      tableHeight: number;
      currentCellBorderFillId: number;
      currentCellColSpan: number;
      currentCellRowSpan: number;
      currentCellWidth: number;
      currentCellHeight: number;
      currentCellMarginLeft: number;
      currentCellMarginRight: number;
      currentCellMarginTop: number;
      currentCellMarginBottom: number;
    }
    const tableStack: TableContext[] = [];
    
    let inHeaderFooter = false;
    let headerFooterLevel = -1;
    
    interface GsoContext {
      level: number;
      binItemId: number | null;
      inHeaderFooterAtCreation: boolean;
    }
    let gsoStack: GsoContext[] = [];

    const getCurrentTable = (): TableContext | null => {
      return tableStack.length > 0 ? tableStack[tableStack.length - 1] : null;
    };

    const finalizePara = () => {
      if (currentParagraph) {
        const para = currentParagraph as HwpxParagraph;
        para.runs = this.createRuns(currentText, currentCharShapePositions);
        
        const currentTableCtx = getCurrentTable();
        if (currentTableCtx) {
          currentTableCtx.cellParagraphs.push(para);
        } else if (currentLevel <= 1 && currentParaTableDepth === 0) {
          elements.push({ type: 'paragraph', data: para });
        }
      }
    };

    const finalizeCell = () => {
      const currentTableCtx = getCurrentTable();
      if (!currentTableCtx) return;
      
      // Only add cell if we have started processing it (colCount is set)
      // This means LIST_HEADER was called at least once
      if (currentTableCtx.colCount > 0 && currentTableCtx.cellIndex >= 0) {
        const rowIdx = Math.floor(currentTableCtx.cellIndex / currentTableCtx.colCount);
        const colIdx = currentTableCtx.cellIndex % currentTableCtx.colCount;
        
        if (!currentTableCtx.table.rows) currentTableCtx.table.rows = [];
        while (currentTableCtx.table.rows.length <= rowIdx) {
          currentTableCtx.table.rows.push({ cells: [] });
        }
        
        let backgroundColor: string | undefined;
        let borderTop: { width: number; style: string; color: string } | undefined;
        let borderBottom: { width: number; style: string; color: string } | undefined;
        let borderLeft: { width: number; style: string; color: string } | undefined;
        let borderRight: { width: number; style: string; color: string } | undefined;
        
        const borderFillId = currentTableCtx.currentCellBorderFillId;
        if (borderFillId > 0 && borderFillId <= this.borderFills.length) {
          const borderFill = this.borderFills[borderFillId - 1];
          if (borderFill?.fill?.backgroundColor !== undefined) {
            backgroundColor = this.colorToHex(borderFill.fill.backgroundColor);
          }
          if (borderFill?.borders) {
            const borderWidthMap = [0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0];
            const borderStyleMap: Record<number, string> = {
              0: 'solid', 1: 'dashed', 2: 'dotted', 3: 'dash-dot', 4: 'dash-dot-dot',
              5: 'dashed', 6: 'dotted', 7: 'double', 8: 'double', 9: 'double', 10: 'double',
              11: 'wavy', 12: 'wavy', 13: 'solid', 14: 'solid', 15: 'solid', 16: 'solid',
            };
            const mapBorder = (b: BorderLine) => {
              const widthMm = borderWidthMap[b.width] ?? (b.width + 1) * 0.1;
              const widthPt = widthMm * 2.8346;
              return {
                width: widthPt,
                style: borderStyleMap[b.type] || 'solid',
                color: this.colorToHex(b.color),
              };
            };
            if (borderFill.borders.top) borderTop = mapBorder(borderFill.borders.top);
            if (borderFill.borders.bottom) borderBottom = mapBorder(borderFill.borders.bottom);
            if (borderFill.borders.left) borderLeft = mapBorder(borderFill.borders.left);
            if (borderFill.borders.right) borderRight = mapBorder(borderFill.borders.right);
          }
        }
        
        const targetRow = currentTableCtx.table.rows?.[rowIdx];
        if (targetRow && targetRow.cells) {
          targetRow.cells.push({
            paragraphs: currentTableCtx.cellParagraphs,
            nestedTables: currentTableCtx.cellNestedTables.length > 0 ? currentTableCtx.cellNestedTables : undefined,
            colAddr: colIdx,
            rowAddr: rowIdx,
            colSpan: currentTableCtx.currentCellColSpan || 1,
            rowSpan: currentTableCtx.currentCellRowSpan || 1,
            width: currentTableCtx.currentCellWidth,
            height: currentTableCtx.currentCellHeight,
            marginLeft: currentTableCtx.currentCellMarginLeft,
            marginRight: currentTableCtx.currentCellMarginRight,
            marginTop: currentTableCtx.currentCellMarginTop,
            marginBottom: currentTableCtx.currentCellMarginBottom,
            backgroundColor,
            borderTop,
            borderBottom,
            borderLeft,
            borderRight,
          });
        }
      }
      
      // Always reset cell state
      currentTableCtx.cellParagraphs = [];
      currentTableCtx.cellNestedTables = [];
      currentTableCtx.currentCellBorderFillId = 0;
      currentTableCtx.currentCellColSpan = 1;
      currentTableCtx.currentCellRowSpan = 1;
      currentTableCtx.currentCellWidth = 0;
      currentTableCtx.currentCellHeight = 0;
      currentTableCtx.currentCellMarginLeft = 0;
      currentTableCtx.currentCellMarginRight = 0;
      currentTableCtx.currentCellMarginTop = 0;
      currentTableCtx.currentCellMarginBottom = 0;
    };

    const finalizeTable = () => {
      const currentTableCtx = tableStack.pop();
      if (currentTableCtx && currentTableCtx.table.rows && currentTableCtx.table.rows.length > 0) {
        const parentTableCtx = getCurrentTable();
        if (parentTableCtx) {
          parentTableCtx.cellNestedTables.push(currentTableCtx.table as HwpxTable);
        } else {
          elements.push({ type: 'table', data: currentTableCtx.table as HwpxTable });
        }
      }
    };

    const handleRecord = (header: RecordHeader, recordData: Uint8Array) => {
      if (inHeaderFooter && header.level <= headerFooterLevel) {
        inHeaderFooter = false;
        headerFooterLevel = -1;
      }
      
      while (gsoStack.length > 0) {
        const gsoCtx = gsoStack[gsoStack.length - 1];
        if (header.level <= gsoCtx.level) {
          if (gsoCtx.binItemId !== null) {
            const binIdKey = gsoCtx.binItemId.toString().padStart(4, '0');
            const binData = this.binDataItems.get(binIdKey);
            if (binData) {
              const img: HwpxImage = {
                id: this.generateId(),
                binaryId: binIdKey,
                width: 100,
                height: 100,
                data: binData.data,
              };
              elements.push({ type: 'image', data: img });
            }
          }
          gsoStack.pop();
        } else {
          break;
        }
      }
      
      while (tableStack.length > 0) {
        const currentTableCtx = getCurrentTable()!;
        if (header.level <= currentTableCtx.level && header.tagId !== HWPTAG_TABLE) {
          finalizePara();
          finalizeCell();
          finalizeTable();
        } else {
          break;
        }
      }

      switch (header.tagId) {
        case HWPTAG_PARA_HEADER:
          finalizePara();
          currentParagraph = this.parseParaHeader(recordData);
          currentText = '';
          currentCharShapePositions = [];
          currentLevel = header.level;
          currentParaInHeaderFooter = inHeaderFooter;
          currentParaTableDepth = tableStack.length;
          break;

        case HWPTAG_PARA_TEXT:
          currentText = this.parseParaText(recordData);
          break;

        case HWPTAG_PARA_CHAR_SHAPE:
          currentCharShapePositions = this.parseParaCharShape(recordData);
          break;

        case HWPTAG_CTRL_HEADER: {
          const view = new DataView(recordData.buffer, recordData.byteOffset, recordData.byteLength);
          const ctrlId = view.getUint32(0, true);
          
          if (ctrlId === 0x68656164 || ctrlId === 0x746f6f66) {
            inHeaderFooter = true;
            headerFooterLevel = header.level;
          }
          
          if (ctrlId === 0x67736f20) {
            gsoStack.push({
              level: header.level,
              binItemId: null,
              inHeaderFooterAtCreation: inHeaderFooter,
            });
          }
          
          if (ctrlId === 0x74626c20) {
            finalizePara();
            currentParagraph = null;
            
            const isTopLevel = header.level <= 1 && !inHeaderFooter;
            
            // Extract table width and height from common object properties
            // Structure: ctrlId(4) + props(4) + vOffset(4) + hOffset(4) + width(4) + height(4)
            let tableWidth = 0;
            let tableHeight = 0;
            if (recordData.length >= 24) {
              tableWidth = view.getUint32(16, true) / 7200 * 72; // HWPUNIT to pt
              tableHeight = view.getUint32(20, true) / 7200 * 72;
            }
            
            tableStack.push({
              table: {
                id: this.generateId(),
                rows: [],
                width: tableWidth,
              },
              level: header.level,
              rowCount: 0,
              colCount: 0,
              cellIndex: -1,
              cellParagraphs: [],
              cellNestedTables: [],
              isTopLevel,
              tableWidth,
              tableHeight,
              currentCellBorderFillId: 0,
              currentCellColSpan: 1,
              currentCellRowSpan: 1,
              currentCellWidth: 0,
              currentCellHeight: 0,
              currentCellMarginLeft: 0,
              currentCellMarginRight: 0,
              currentCellMarginTop: 0,
              currentCellMarginBottom: 0,
            });
          }
          break;
        }
        
        case HWPTAG_SHAPE_COMPONENT_PICTURE: {
          if (gsoStack.length > 0 && recordData.length > 71) {
            const binItemId = recordData[71];
            gsoStack[gsoStack.length - 1].binItemId = binItemId;
          }
          break;
        }

        case HWPTAG_TABLE: {
          const currentTableCtx = getCurrentTable();
          if (currentTableCtx) {
            const view = new DataView(recordData.buffer, recordData.byteOffset, recordData.byteLength);
            currentTableCtx.rowCount = view.getUint16(4, true);
            currentTableCtx.colCount = view.getUint16(6, true);
            currentTableCtx.table.rowCount = currentTableCtx.rowCount;
            currentTableCtx.table.colCount = currentTableCtx.colCount;
          }
          break;
        }

        case HWPTAG_LIST_HEADER: {
          const currentTableCtx = getCurrentTable();
          if (currentTableCtx) {
            finalizePara();
            finalizeCell();
            currentParagraph = null;
            currentTableCtx.cellIndex++;
            
            const view = new DataView(recordData.buffer, recordData.byteOffset, recordData.byteLength);
            const listHeaderSize = recordData.length >= 34 ? 8 : 6;
            if (recordData.length >= listHeaderSize + 26) {
              currentTableCtx.currentCellColSpan = view.getUint16(listHeaderSize + 4, true) || 1;
              currentTableCtx.currentCellRowSpan = view.getUint16(listHeaderSize + 6, true) || 1;
              currentTableCtx.currentCellWidth = view.getUint32(listHeaderSize + 8, true) / 7200 * 72;
              currentTableCtx.currentCellHeight = view.getUint32(listHeaderSize + 12, true) / 7200 * 72;
              currentTableCtx.currentCellMarginLeft = view.getUint16(listHeaderSize + 16, true) / 7200 * 72;
              currentTableCtx.currentCellMarginRight = view.getUint16(listHeaderSize + 18, true) / 7200 * 72;
              currentTableCtx.currentCellMarginTop = view.getUint16(listHeaderSize + 20, true) / 7200 * 72;
              currentTableCtx.currentCellMarginBottom = view.getUint16(listHeaderSize + 22, true) / 7200 * 72;
              currentTableCtx.currentCellBorderFillId = view.getUint16(listHeaderSize + 24, true);
            }
          }
          break;
        }

        case HWPTAG_PAGE_DEF: {
          if (recordData.length >= 40) {
            const view = new DataView(recordData.buffer, recordData.byteOffset, recordData.byteLength);
            section.pageSettings = {
              width: view.getUint32(0, true) / 100,
              height: view.getUint32(4, true) / 100,
              marginLeft: view.getUint32(8, true) / 100,
              marginRight: view.getUint32(12, true) / 100,
              marginTop: view.getUint32(16, true) / 100,
              marginBottom: view.getUint32(20, true) / 100,
              orientation: (view.getUint32(36, true) & 0x01) ? 'landscape' : 'portrait',
            };
          }
          break;
        }
      }
    };

    this.parseRecords(data, handleRecord);
    finalizePara();
    while (tableStack.length > 0) {
      finalizeCell();
      finalizeTable();
    }
    
    while (gsoStack.length > 0) {
      const gsoCtx = gsoStack.pop()!;
      if (gsoCtx.binItemId !== null) {
        const binIdKey = gsoCtx.binItemId.toString().padStart(4, '0');
        const binData = this.binDataItems.get(binIdKey);
        if (binData) {
          const img: HwpxImage = {
            id: this.generateId(),
            binaryId: binIdKey,
            width: 100,
            height: 100,
            data: binData.data,
          };
          elements.push({ type: 'image', data: img });
        }
      }
    }

    section.elements = elements;
    return section;
  }

  private createRuns(text: string, charShapePositions: { pos: number; shapeIndex: number }[]): TextRun[] {
    if (charShapePositions.length === 0 || text.length === 0) {
      return [{ text }];
    }

    const runs: TextRun[] = [];
    let lastPos = 0;

    for (let i = 0; i < charShapePositions.length; i++) {
      const startPos = charShapePositions[i].pos;
      const endPos = i + 1 < charShapePositions.length 
        ? Math.min(charShapePositions[i + 1].pos, text.length)
        : text.length;
      
      if (startPos > lastPos && lastPos < text.length) {
        const gapText = text.substring(lastPos, Math.min(startPos, text.length));
        if (gapText) runs.push({ text: gapText });
      }
      
      if (startPos < text.length) {
        const runText = text.substring(startPos, endPos);
        if (runText) {
          const shapeIndex = charShapePositions[i].shapeIndex;
          const charShape = this.charShapes[shapeIndex];
          runs.push({
            text: runText,
            charStyle: charShape ? this.convertCharShape(charShape) : undefined,
          });
        }
        lastPos = endPos;
      }
    }

    if (lastPos < text.length) {
      runs.push({ text: text.substring(lastPos) });
    }

    return runs.length > 0 ? runs : [{ text }];
  }

  private parseParaHeader(data: Uint8Array): Partial<HwpxParagraph> {
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    
    const textLength = view.getUint32(0, true);
    const controlMask = view.getUint32(4, true);
    const paraShapeId = view.getUint16(8, true);
    const styleId = view.getUint8(10);
    const breakType = view.getUint8(11);
    const charShapeCount = view.getUint16(12, true);
    const rangeTagCount = view.getUint16(14, true);
    const lineAlignCount = view.getUint16(16, true);
    const instanceId = view.getUint32(18, true);

    const paraShape = this.paraShapes[paraShapeId];
    
    return {
      id: this.generateId(),
      runs: [],
      paraStyle: paraShape ? this.convertParaShape(paraShape) : undefined,
    };
  }

  private parseParaText(data: Uint8Array): string {
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    let text = '';
    let i = 0;
    
    while (i < data.length) {
      const charCode = view.getUint16(i, true);
      
      if (charCode === 0) break;
      
      if (charCode < 32) {
        i += 2;
        
        switch (charCode) {
          case 1: i += 14; break;
          case 2: i += 14; break;
          case 3: i += 14; break;
          case 11: i += 14; break;
          case 12: i += 14; break;
          case 14: i += 14; break;
          case 15: i += 14; break;
          case 16: i += 14; break;
          case 17: i += 14; break;
          case 18: i += 14; break;
          case 21: i += 14; break;
          case 22: i += 14; break;
          case 23: i += 14; break;
          case 9: text += '\t'; break;
          case 10: text += '\n'; break;
          case 13: break;
          case 24: i += 2; break;
          default: break;
        }
        continue;
      }
      
      text += String.fromCharCode(charCode);
      i += 2;
    }
    
    return text;
  }

  private parseParaCharShape(data: Uint8Array): { pos: number; shapeIndex: number }[] {
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    const positions: { pos: number; shapeIndex: number }[] = [];
    
    for (let i = 0; i < data.length; i += 8) {
      const pos = view.getUint32(i, true);
      const shapeIndex = view.getUint32(i + 4, true);
      positions.push({ pos, shapeIndex });
    }
    
    return positions;
  }

  private finalizeParagraph(
    para: HwpxParagraph,
    text: string,
    charShapePositions: { pos: number; shapeIndex: number }[],
    elements: SectionElement[]
  ): void {
    if (charShapePositions.length === 0) {
      para.runs = [{ text }];
    } else {
      const runs: TextRun[] = [];
      
      for (let i = 0; i < charShapePositions.length; i++) {
        const startPos = charShapePositions[i].pos;
        const endPos = i + 1 < charShapePositions.length 
          ? charShapePositions[i + 1].pos 
          : text.length;
        
        const runText = text.substring(startPos, endPos);
        if (runText.length === 0) continue;
        
        const shapeIndex = charShapePositions[i].shapeIndex;
        const charShape = this.charShapes[shapeIndex];
        
        runs.push({
          text: runText,
          charStyle: charShape ? this.convertCharShape(charShape) : undefined,
        });
      }
      
      para.runs = runs.length > 0 ? runs : [{ text }];
    }
    
    elements.push({ type: 'paragraph', data: para });
  }

  private convertParaShape(shape: ParaShape): ParagraphStyle {
    const alignMap: Record<number, ParagraphStyle['align']> = {
      0: 'Justify',
      1: 'Left',
      2: 'Right',
      3: 'Center',
      4: 'Distribute',
      5: 'Distribute',
    };

    const lineSpacingTypeMap: Record<number, ParagraphStyle['lineSpacingType']> = {
      0: 'percent',
      1: 'fixed',
      2: 'betweenLines',
      3: 'atLeast',
    };

    return {
      align: alignMap[shape.align] || 'Justify',
      marginLeft: shape.marginLeft / 100,
      marginRight: shape.marginRight / 100,
      marginTop: shape.marginTop / 100,
      marginBottom: shape.marginBottom / 100,
      firstLineIndent: shape.indent / 100,
      lineSpacing: shape.lineSpacing,
      lineSpacingType: lineSpacingTypeMap[shape.lineSpacingType],
      keepWithNext: shape.keepWithNext,
      keepLines: shape.keepLines,
      pageBreakBefore: shape.pageBreakBefore,
      widowControl: shape.widowOrphan,
    };
  }

  private convertCharShape(shape: CharShape): CharacterStyle {
    const fontFace = this.fontFaces[shape.fontId[0]];
    
    return {
      fontName: fontFace?.name,
      fontSize: shape.height / 100,
      bold: shape.bold,
      italic: shape.italic,
      underline: shape.underline === 1 || shape.underline === 3,
      strikethrough: shape.strikeout >= 2,
      superscript: shape.superscript,
      subscript: shape.subscript,
      fontColor: this.colorToHex(shape.textColor),
      emboss: shape.emboss,
      engrave: shape.engrave,
      useFontSpace: shape.useFontSpace,
      useKerning: shape.useKerning,
    };
  }

  private colorToHex(color: number): string {
    const r = color & 0xFF;
    const g = (color >> 8) & 0xFF;
    const b = (color >> 16) & 0xFF;
    return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
  }

  private getDefaultPageSettings(): PageSettings {
    return {
      width: 595,
      height: 842,
      marginTop: 56.7,
      marginBottom: 56.7,
      marginLeft: 56.7,
      marginRight: 56.7,
      orientation: 'portrait',
    };
  }

  private generateId(): string {
    return Math.random().toString(36).substring(2, 11);
  }

  private uint8ArrayToBase64(data: Uint8Array): string {
    let binary = '';
    for (let i = 0; i < data.length; i++) {
      binary += String.fromCharCode(data[i]);
    }
    return btoa(binary);
  }

  public toHwpxContent(): HwpxContent {
    const sections = this.parseSections();

    return {
      metadata: {
        title: undefined,
        creator: undefined,
      },
      sections,
      images: new Map(),
      binItems: new Map(),
      binData: this.binDataItems,
      footnotes: [],
      endnotes: [],
    };
  }
}
