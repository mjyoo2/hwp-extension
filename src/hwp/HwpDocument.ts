/**
 * HWP Document - Binary HWP file format handler
 * Based on HWP 5.0 specification (한글문서파일형식_5.0_revision1.3)
 */

import * as vscode from 'vscode';
import * as CFB from 'cfb';
import * as pako from 'pako';
import {
  HwpxContent,
  HwpxSection,
  HwpxParagraph,
  TextRun,
  PageSettings,
  HwpxImage,
  HwpxTable,
  TableRow,
  TableCell,
  CharShape,
  ParaShape,
  FontInfo,
} from '../hwpx/types';

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
  HWPTAG_SHAPE_COMPONENT_PICTURE: 85,
  HWPTAG_SHAPE_COMPONENT_CONTAINER: 86,
  HWPTAG_EQEDIT: 88,
  HWPTAG_SHAPE_COMPONENT_OLE: 89,
  HWPTAG_SHAPE_COMPONENT_TEXTART: 90,
  HWPTAG_FORM_OBJECT: 91,
  HWPTAG_MEMO_SHAPE: 92,
  HWPTAG_MEMO_LIST: 93,
  HWPTAG_CHART_DATA: 95,
  HWPTAG_VIDEO_DATA: 98,
  HWPTAG_TRACK_CHANGE: 100,
  HWPTAG_TRACK_CHANGE_AUTHOR: 101,
  HWPTAG_DOC_DATA: 27,
  HWPTAG_DISTRIBUTE_DOC_DATA: 28,
  HWPTAG_COMPATIBLE_DOCUMENT: 30,
  HWPTAG_LAYOUT_COMPATIBILITY: 31,
  HWPTAG_FORBIDDEN_CHAR: 94,
};

const CTRL_CHAR = {
  SECTION_COLUMN_DEF: 2,
  FIELD_START: 3,
  INLINE: 5,
  EXTENDED: 6,
  LINE_BREAK: 10,
  TABLE_DRAWING: 11,
  PARAGRAPH_BREAK: 13,
};

const CTRL_ID = {
  TABLE: 0x74626C20,         // 'tbl '
  PICTURE: 0x24706963,       // '$pic'
  SECTION_DEF: 0x73656364,   // 'secd'
  COLUMN_DEF: 0x636F6C64,    // 'cold'
  FORM: 0x666F726D,          // 'form'
  GSO: 0x67736F20,           // 'gso '
  HEADER: 0x68656164,        // 'head'
  FOOTER: 0x666F6F74,        // 'foot'
  FOOTNOTE: 0x666E2020,      // 'fn  '
  ENDNOTE: 0x656E2020,       // 'en  '
  AUTO_NUMBER: 0x2461746E,   // '$atn'
  NEW_NUMBER: 0x246E776E,    // '$nwn'
  PAGE_NUMBER_POS: 0x24706E70, // '$pnp'
  BOOKMARK: 0x626F6F6B,      // 'book'
  INDEX_MARK: 0x69646D6B,    // 'idmk'
  CHAR_OVERLAP: 0x24636F76,  // '$cov'
  ANNOTATION: 0x246E6F74,    // '$not'
  HIDDEN_COMMENT: 0x24686964, // '$hid'
  PAGE_HIDE: 0x24706867,     // '$phg'
  PAGE_ODD_EVEN: 0x24706F65, // '$poe'
  LINE: 0x246C696E,          // '$lin'
  RECTANGLE: 0x24726563,     // '$rec'
  ELLIPSE: 0x24656C6C,       // '$ell'
  ARC: 0x24617263,           // '$arc'
  POLYGON: 0x24706F6C,       // '$pol'
  CURVE: 0x24637276,         // '$crv'
  CONTAINER: 0x2463746E,     // '$ctn'
  FIELD_UNKNOWN: 0x24663000,     // '$f\0' (0x00)
  FIELD_DATE: 0x24663001,        // '$f\0' (0x01)
  FIELD_DOCDATE: 0x24663002,     // '$f\0' (0x02)
  FIELD_PATH: 0x24663003,        // '$f\0' (0x03)
  FIELD_BOOKMARK: 0x24663004,    // '$f\0' (0x04)
  FIELD_MAILMERGE: 0x24663005,   // '$f\0' (0x05)
  FIELD_CROSSREF: 0x24663006,    // '$f\0' (0x06)
  FIELD_FORMULA: 0x24663007,     // '$f\0' (0x07)
  FIELD_CLICKHERE: 0x24663008,   // '$f\0' (0x08)
  FIELD_SUMMARY: 0x24663009,     // '$f\0' (0x09)
  FIELD_USERINFO: 0x2466300a,    // '$f\0' (0x0a)
  FIELD_HYPERLINK: 0x2466300b,   // '$f\0' (0x0b)
  FIELD_MEMO: 0x2466300c,        // '$f\0' (0x0c)
  FIELD_PRIVATE_INFO: 0x2466300d, // '$f\0' (0x0d)
  FIELD_TOC: 0x2466300e,         // '$f\0' (0x0e)
};

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
  fontIds: number[];  // 7 language font IDs (Korean, English, Chinese, Japanese, Other, Symbol, User)
  widthRatios: number[];  // 7 values, 50-200%
  spacings: number[];  // 7 values, -50 to 50
  relativeSizes: number[];  // 7 values, 10-250%
  charPositions: number[];  // 7 values, -100 to 100
  baseSize: number;  // in HWPUNIT (divide by 100 for pt)
  italic: boolean;
  bold: boolean;
  underlineType: number;  // 0=none, 1=below, 3=above
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
  textColor: number;  // COLORREF
  underlineColor: number;
  shadeColor: number;
  shadowColor: number;
  borderFillId?: number;
  strikethroughColor?: number;
}

interface ParsedParaShape {
  alignment: number;  // 0=justify, 1=left, 2=right, 3=center, 4=distribute, 5=divide
  leftMargin: number;  // HWPUNIT
  rightMargin: number;  // HWPUNIT
  indent: number;  // HWPUNIT (positive=first line, negative=hanging)
  spacingBefore: number;  // HWPUNIT
  spacingAfter: number;  // HWPUNIT
  lineSpacing: number;  // value
  lineSpacingType: number;  // 0=relative, 1=fixed, 2=margin_only, 3=minimum
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

interface ParsedBorderLine {
  type: number;   // 0=solid, 1=long dash, 2=dot, etc. (표 25)
  width: number;  // 0-15 (표 26)
  color: number;  // COLORREF
}

interface ParsedSolidFill {
  fillType: 'solid';
  backgroundColor: number;  // COLORREF
  patternColor: number;     // COLORREF
  patternType: number;      // 0=none, 1=horizontal, 2=vertical, etc.
}

interface ParsedGradientFill {
  fillType: 'gradient';
  gradientType: number;  // 1=linear, 2=radial, 3=conical, 4=square
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

// Phase 2.1 - DocInfo Tags Interfaces

interface ParsedDocumentProperties {
  sectionCount: number;
  startNumbers: {
    page: number;
    footnote: number;
    endnote: number;
    figure: number;
    table: number;
    equation: number;
  };
  caretPosition: {
    listId: number;
    paragraphId: number;
    charPosition: number;
  };
}

interface ParsedIdMappings {
  binaryDataCount: number;
  fontCounts: {
    korean: number;
    english: number;
    chinese: number;
    japanese: number;
    other: number;
    symbol: number;
    user: number;
  };
  borderFillCount: number;
  charShapeCount: number;
  tabDefCount: number;
  numberingCount: number;
  bulletCount: number;
  paraShapeCount: number;
  styleCount: number;
  memoShapeCount: number;
  trackChangeCount: number;
  trackChangeAuthorCount: number;
}

interface ParsedTabDef {
  autoTabLeft: boolean;
  autoTabRight: boolean;
  tabs: Array<{
    position: number;
    type: 'left' | 'right' | 'center' | 'decimal';
    fillType: number;
  }>;
}

interface ParsedParagraphHeadInfo {
  alignment: 'left' | 'center' | 'right';
  widthFollowsInstance: boolean;
  autoIndent: boolean;
  distanceType: 'relative' | 'absolute';
  widthCorrection: number;
  distanceFromText: number;
}

interface ParsedNumbering {
  levels: Array<{
    headInfo: ParsedParagraphHeadInfo;
    format: string;
  }>;
}

interface ParsedBullet {
  headInfo: ParsedParagraphHeadInfo;
  bulletChar: string;
  isImageBullet: boolean;
  imageBulletId?: number;
  imageBulletInfo?: {
    contrast: number;
    brightness: number;
    effect: number;
    id: number;
  };
  checkBulletChar: string;
}

interface ParsedStyle {
  localName: string;
  englishName: string;
  type: 'paragraph' | 'character';
  nextStyleId: number;
  languageId: number;
  paraShapeId: number;
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

interface ParsedFootnoteShape {
  numberType: number;
  superscript: boolean;
  numberingType: 'continuous' | 'perSection' | 'perPage';
  placement: 'column' | 'pageBottom' | 'endOfDocument' | 'endOfSection';
  belowText: boolean;
  userChar: string;
  prefixChar: string;
  suffixChar: string;
  startNumber: number;
  separatorLength: number;
  spaceAbove: number;
  spaceBelow: number;
  spaceBetween: number;
  separatorLineType: number;
  separatorLineWidth: number;
  separatorLineColor: number;
}

interface ParsedPageBorderFill {
  applyTo: 'both' | 'even' | 'odd' | 'firstOnly';
  includeHeader: boolean;
  includeFooter: boolean;
  fillArea: 'paper' | 'page' | 'border';
  offsets: {
    left: number;
    right: number;
    top: number;
    bottom: number;
  };
  borderFillId: number;
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

interface CharShapePosition {
  startPos: number;
  charShapeId: number;
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

interface ParsedTrackChange {
  id: number;
  type: number;
}

interface ParsedTrackChangeAuthor {
  name: string;
}

interface ParsedCompatibleDocument {
  targetProgram: number;
}

interface ParsedLayoutCompatibility {
  lineWrap: number;
  charUnit: number;
  paraBottomSpacing: number;
  underline: number;
  strikeout: number;
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
  pendingShape: { type: string; width: number; height: number; x: number; y: number } | null;
  pendingField: ParsedField | null;
  inMemo: boolean;
  currentMemo: { id: string; linkedText: string; content: string[] } | null;
  memoParagraphs: HwpxParagraph[];
  memos: import('../hwpx/types').Memo[];
  inShapeText: boolean;
  shapeTextParagraphs: HwpxParagraph[];
}

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

function readWString(data: Uint8Array, offset: number, length: number): string {
  const chars: number[] = [];
  for (let i = 0; i < length; i++) {
    const charCode = readUint16(data, offset + i * 2);
    if (charCode === 0) break;
    chars.push(charCode);
  }
  return String.fromCharCode(...chars);
}

function colorrefToHex(colorref: number): string {
  const r = colorref & 0xFF;
  const g = (colorref >> 8) & 0xFF;
  const b = (colorref >> 16) & 0xFF;
  return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
}

function hwpunitToPt(hwpunit: number): number {
  return hwpunit / 7200 * 72;
}

export class HwpDocument implements vscode.CustomDocument {
  private _uri: vscode.Uri;
  private _content: HwpxContent;
  private _cfb: CFB.CFB$Container | null = null;
  private _compressed: boolean = false;
  private _binDataInfos: Map<number, BinDataInfo> = new Map();
  private _faceNames: Map<number, ParsedFaceName> = new Map();
  private _charShapes: Map<number, ParsedCharShape> = new Map();
  private _paraShapes: Map<number, ParsedParaShape> = new Map();
  private _borderFills: Map<number, ParsedBorderFill> = new Map();
  private _documentProperties: ParsedDocumentProperties | null = null;
  private _idMappings: ParsedIdMappings | null = null;
  private _tabDefs: Map<number, ParsedTabDef> = new Map();
  private _numberings: Map<number, ParsedNumbering> = new Map();
  private _bullets: Map<number, ParsedBullet> = new Map();
  private _styles: Map<number, ParsedStyle> = new Map();
  private _footnoteShapes: Map<number, ParsedFootnoteShape> = new Map();
  private _memoShapes: Map<number, ParsedMemoShape> = new Map();
  private _memoCount: number = 0;
  private _trackChanges: ParsedTrackChange[] = [];
  private _trackChangeAuthors: ParsedTrackChangeAuthor[] = [];
  private _docData: Uint8Array | null = null;
  private _distributeDocData: Uint8Array | null = null;
  private _compatibleDocument: ParsedCompatibleDocument | null = null;
  private _layoutCompatibility: ParsedLayoutCompatibility | null = null;
  private _forbiddenChar: Uint8Array | null = null;
  private _isReadOnly: boolean = true;
  
  private readonly _onDidChangeContent = new vscode.EventEmitter<void>();
  public readonly onDidChangeContent = this._onDidChangeContent.event;
  
  private readonly _onDidDispose = new vscode.EventEmitter<void>();
  public readonly onDidDispose = this._onDidDispose.event;

  private constructor(uri: vscode.Uri) {
    this._uri = uri;
    this._content = {
      metadata: {},
      sections: [],
      images: new Map(),
      binItems: new Map(),
      binData: new Map(),
      footnotes: [],
      endnotes: [],
    };
  }

  get uri(): vscode.Uri { return this._uri; }
  get isReadOnly(): boolean { return this._isReadOnly; }

  static async create(uri: vscode.Uri): Promise<HwpDocument> {
    const document = new HwpDocument(uri);
    await document.load();
    return document;
  }

  static parseContent(data: Uint8Array): HwpxContent {
    return parseHwpContent(data);
  }

  private async load(): Promise<void> {
    const fileData = await vscode.workspace.fs.readFile(this._uri);
    this._cfb = CFB.read(fileData, { type: 'array' });
    
    const headerData = this.getEntryData('/FileHeader');
    if (!headerData || headerData.length < 256) {
      throw new Error('Invalid HWP file');
    }
    
    const signature = new TextDecoder('utf-8').decode(headerData.slice(0, 32)).replace(/\0/g, '');
    if (!signature.startsWith('HWP Document File')) {
      throw new Error('Invalid HWP signature');
    }
    
    const props = readUint32(headerData, 36);
    this._compressed = (props & 0x01) !== 0;
    const encrypted = (props & 0x02) !== 0;
    
    if (encrypted) {
      throw new Error('Encrypted HWP files are not supported');
    }
    
    this._content.metadata = { title: this._uri.path.split('/').pop()?.replace('.hwp', '') };
    
    this.parseDocInfo();
    this.parseBinData();
    this.parseBodyText();
  }

  private getEntryData(path: string): Uint8Array | null {
    if (!this._cfb) return null;
    const entry = CFB.find(this._cfb, path);
    if (!entry?.content) return null;
    return entry.content instanceof Uint8Array ? entry.content : new Uint8Array(entry.content as unknown as ArrayBuffer);
  }

  private decompress(data: Uint8Array): Uint8Array {
    try {
      return pako.inflateRaw(data);
    } catch {
      return data;
    }
  }

  private parseDocInfo(): void {
    let data = this.getEntryData('/DocInfo');
    if (!data) return;
    if (this._compressed) data = this.decompress(data);
    
    let offset = 0;
    let binDataId = 1;
    let faceNameId = 0;
    let charShapeId = 0;
    let paraShapeId = 0;
    let borderFillId = 1;
    let tabDefId = 0;
    let numberingId = 0;
    let bulletId = 0;
    let styleId = 0;
    
    while (offset < data.length) {
      const result = this.parseRecordHeader(data, offset);
      if (!result) break;
      
      const { tagId, size, nextOffset } = result;
      const recordData = data.slice(nextOffset, nextOffset + size);
      
      if (tagId === HWP_TAGS.HWPTAG_BIN_DATA && recordData.length >= 2) {
        const props = readUint16(recordData, 0);
        const type = props & 0x0F;
        let ext = '';
        
        if (type === 1) {
          let extOffset = 2;
          if (recordData.length > extOffset + 2) {
            extOffset += 2 + readUint16(recordData, extOffset) * 2;
          }
          if (recordData.length > extOffset + 2) {
            extOffset += 2 + readUint16(recordData, extOffset) * 2;
          }
          extOffset += 2;
          if (recordData.length > extOffset + 2) {
            const extLen = readUint16(recordData, extOffset);
            extOffset += 2;
            if (recordData.length >= extOffset + extLen * 2) {
              ext = new TextDecoder('utf-16le').decode(recordData.slice(extOffset, extOffset + extLen * 2));
            }
          }
        }
        
        this._binDataInfos.set(binDataId, {
          id: binDataId,
          type: type === 0 ? 'LINK' : type === 1 ? 'EMBEDDING' : 'STORAGE',
          extension: ext.toLowerCase(),
        });
        binDataId++;
      } else if (tagId === HWP_TAGS.HWPTAG_FACE_NAME && recordData.length >= 3) {
        const faceName = this.parseFaceName(recordData);
        if (faceName) {
          this._faceNames.set(faceNameId, faceName);
        }
        faceNameId++;
      } else if (tagId === HWP_TAGS.HWPTAG_CHAR_SHAPE && recordData.length >= 72) {
        const charShape = this.parseCharShape(recordData);
        if (charShape) {
          this._charShapes.set(charShapeId, charShape);
        }
        charShapeId++;
      } else if (tagId === HWP_TAGS.HWPTAG_PARA_SHAPE && recordData.length >= 42) {
        const paraShape = this.parseParaShape(recordData);
        if (paraShape) {
          this._paraShapes.set(paraShapeId, paraShape);
        }
        paraShapeId++;
      } else if (tagId === HWP_TAGS.HWPTAG_BORDER_FILL && recordData.length >= 32) {
        const borderFill = this.parseBorderFill(recordData);
        if (borderFill) {
          this._borderFills.set(borderFillId, borderFill);
        }
        borderFillId++;
      } else if (tagId === HWP_TAGS.HWPTAG_DOCUMENT_PROPERTIES && recordData.length >= 26) {
        this._documentProperties = this.parseDocumentProperties(recordData);
      } else if (tagId === HWP_TAGS.HWPTAG_ID_MAPPINGS && recordData.length >= 56) {
        this._idMappings = this.parseIdMappings(recordData);
      } else if (tagId === HWP_TAGS.HWPTAG_TAB_DEF && recordData.length >= 6) {
        const tabDef = this.parseTabDef(recordData);
        if (tabDef) {
          this._tabDefs.set(tabDefId, tabDef);
        }
        tabDefId++;
      } else if (tagId === HWP_TAGS.HWPTAG_NUMBERING && recordData.length >= 8) {
        const numbering = this.parseNumbering(recordData);
        if (numbering) {
          this._numberings.set(numberingId, numbering);
        }
        numberingId++;
      } else if (tagId === HWP_TAGS.HWPTAG_BULLET && recordData.length >= 12) {
        const bullet = this.parseBullet(recordData);
        if (bullet) {
          this._bullets.set(bulletId, bullet);
        }
        bulletId++;
      } else if (tagId === HWP_TAGS.HWPTAG_STYLE && recordData.length >= 10) {
        const style = this.parseStyle(recordData);
        if (style) {
          this._styles.set(styleId, style);
        }
        styleId++;
      } else if (tagId === HWP_TAGS.HWPTAG_FOOTNOTE_SHAPE && recordData.length >= 26) {
        const footnoteShape = this.parseFootnoteShape(recordData);
        if (footnoteShape) {
          this._footnoteShapes.set(this._footnoteShapes.size, footnoteShape);
        }
      } else if (tagId === HWP_TAGS.HWPTAG_DOC_DATA) {
        this._docData = recordData;
      } else if (tagId === HWP_TAGS.HWPTAG_DISTRIBUTE_DOC_DATA) {
        this._distributeDocData = recordData;
      } else if (tagId === HWP_TAGS.HWPTAG_COMPATIBLE_DOCUMENT && recordData.length >= 4) {
        this._compatibleDocument = this.parseCompatibleDocument(recordData);
      } else if (tagId === HWP_TAGS.HWPTAG_LAYOUT_COMPATIBILITY && recordData.length >= 20) {
        this._layoutCompatibility = this.parseLayoutCompatibility(recordData);
      } else if (tagId === HWP_TAGS.HWPTAG_FORBIDDEN_CHAR) {
        this._forbiddenChar = recordData;
      }
      
      offset = nextOffset + size;
    }
  }

  private parseBorderFill(data: Uint8Array): ParsedBorderFill | null {
    if (data.length < 32) return null;
    
    const props = readUint16(data, 0);
    const effect3d = (props & 0x01) !== 0;
    const shadow = (props & 0x02) !== 0;
    const slashDiagonal = (props >> 2) & 0x07;
    const backslashDiagonal = (props >> 5) & 0x07;
    
    const borderTypes = [data[2], data[3], data[4], data[5]];
    const borderWidths = [data[6], data[7], data[8], data[9]];
    const borderColors = [
      readUint32(data, 10),
      readUint32(data, 14),
      readUint32(data, 18),
      readUint32(data, 22),
    ];
    
    const diagonalType = data[26];
    const diagonalWidth = data[27];
    const diagonalColor = readUint32(data, 28);
    
    const result: ParsedBorderFill = {
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

  private parseFaceName(data: Uint8Array): ParsedFaceName | null {
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

  private parseCharShape(data: Uint8Array): ParsedCharShape | null {
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

  private parseParaShape(data: Uint8Array): ParsedParaShape | null {
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

  private parseDocumentProperties(data: Uint8Array): ParsedDocumentProperties | null {
    if (data.length < 26) return null;
    
    return {
      sectionCount: readUint16(data, 0),
      startNumbers: {
        page: readUint16(data, 2),
        footnote: readUint16(data, 4),
        endnote: readUint16(data, 6),
        figure: readUint16(data, 8),
        table: readUint16(data, 10),
        equation: readUint16(data, 12),
      },
      caretPosition: {
        listId: readUint32(data, 14),
        paragraphId: readUint32(data, 18),
        charPosition: readUint32(data, 22),
      },
    };
  }

  private parseIdMappings(data: Uint8Array): ParsedIdMappings | null {
    if (data.length < 56) return null;
    
    return {
      binaryDataCount: readInt32(data, 0),
      fontCounts: {
        korean: readInt32(data, 4),
        english: readInt32(data, 8),
        chinese: readInt32(data, 12),
        japanese: readInt32(data, 16),
        other: readInt32(data, 20),
        symbol: readInt32(data, 24),
        user: readInt32(data, 28),
      },
      borderFillCount: readInt32(data, 32),
      charShapeCount: readInt32(data, 36),
      tabDefCount: readInt32(data, 40),
      numberingCount: readInt32(data, 44),
      bulletCount: readInt32(data, 48),
      paraShapeCount: readInt32(data, 52),
      styleCount: data.length >= 60 ? readInt32(data, 56) : 0,
      memoShapeCount: data.length >= 64 ? readInt32(data, 60) : 0,
      trackChangeCount: data.length >= 68 ? readInt32(data, 64) : 0,
      trackChangeAuthorCount: data.length >= 72 ? readInt32(data, 68) : 0,
    };
  }

  private parseTabDef(data: Uint8Array): ParsedTabDef | null {
    if (data.length < 6) return null;
    
    const props = readUint32(data, 0);
    const tabCount = readInt16(data, 4);
    
    const tabs: ParsedTabDef['tabs'] = [];
    let offset = 6;
    
    for (let i = 0; i < tabCount && offset + 8 <= data.length; i++) {
      const position = readUint32(data, offset);
      const typeVal = data[offset + 4];
      const fillType = data[offset + 5];
      
      const typeMap: Record<number, 'left' | 'right' | 'center' | 'decimal'> = {
        0: 'left', 1: 'right', 2: 'center', 3: 'decimal'
      };
      
      tabs.push({
        position,
        type: typeMap[typeVal] || 'left',
        fillType,
      });
      
      offset += 8;
    }
    
    return {
      autoTabLeft: (props & 0x01) !== 0,
      autoTabRight: (props & 0x02) !== 0,
      tabs,
    };
  }

  private parseParagraphHeadInfo(data: Uint8Array, offset: number): ParsedParagraphHeadInfo | null {
    if (offset + 8 > data.length) return null;
    
    const props = readUint32(data, offset);
    const widthCorrection = readInt16(data, offset + 4);
    const distanceFromText = readInt16(data, offset + 6);
    
    const alignVal = props & 0x03;
    const alignMap: Record<number, 'left' | 'center' | 'right'> = {
      0: 'left', 1: 'center', 2: 'right'
    };
    
    return {
      alignment: alignMap[alignVal] || 'left',
      widthFollowsInstance: (props & 0x04) !== 0,
      autoIndent: (props & 0x08) !== 0,
      distanceType: (props & 0x10) !== 0 ? 'absolute' : 'relative',
      widthCorrection,
      distanceFromText,
    };
  }

  private parseNumbering(data: Uint8Array): ParsedNumbering | null {
    if (data.length < 8) return null;
    
    const levels: ParsedNumbering['levels'] = [];
    let offset = 0;
    
    for (let level = 0; level < 7 && offset < data.length; level++) {
      const headInfo = this.parseParagraphHeadInfo(data, offset);
      if (!headInfo) break;
      offset += 8;
      
      let format = '';
      if (offset + 2 <= data.length) {
        const formatLen = readUint16(data, offset);
        offset += 2;
        if (offset + formatLen * 2 <= data.length) {
          format = new TextDecoder('utf-16le').decode(data.slice(offset, offset + formatLen * 2));
          offset += formatLen * 2;
        }
      }
      
      levels.push({ headInfo, format });
    }
    
    return levels.length > 0 ? { levels } : null;
  }

  private parseBullet(data: Uint8Array): ParsedBullet | null {
    if (data.length < 12) return null;
    
    const headInfo = this.parseParagraphHeadInfo(data, 0);
    if (!headInfo) return null;
    
    const bulletChar = data.length >= 10 ? String.fromCharCode(readUint16(data, 8)) : '';
    const imageBulletFlag = data.length >= 14 ? readInt32(data, 10) : 0;
    
    const result: ParsedBullet = {
      headInfo,
      bulletChar,
      isImageBullet: imageBulletFlag !== 0,
      checkBulletChar: '',
    };
    
    if (imageBulletFlag !== 0 && data.length >= 18) {
      result.imageBulletId = imageBulletFlag;
      result.imageBulletInfo = {
        contrast: data[14],
        brightness: data[15],
        effect: data[16],
        id: data[17],
      };
    }
    
    if (data.length >= 20) {
      result.checkBulletChar = String.fromCharCode(readUint16(data, 18));
    }
    
    return result;
  }

  private parseStyle(data: Uint8Array): ParsedStyle | null {
    if (data.length < 10) return null;
    
    let offset = 0;
    
    const localNameLen = readUint16(data, offset);
    offset += 2;
    if (offset + localNameLen * 2 > data.length) return null;
    const localName = new TextDecoder('utf-16le').decode(data.slice(offset, offset + localNameLen * 2));
    offset += localNameLen * 2;
    
    if (offset + 2 > data.length) return null;
    const englishNameLen = readUint16(data, offset);
    offset += 2;
    if (offset + englishNameLen * 2 > data.length) return null;
    const englishName = new TextDecoder('utf-16le').decode(data.slice(offset, offset + englishNameLen * 2));
    offset += englishNameLen * 2;
    
    if (offset + 8 > data.length) return null;
    
    const propsVal = data[offset];
    const styleType = (propsVal & 0x07);
    const nextStyleId = data[offset + 1];
    const languageId = readInt16(data, offset + 2);
    const paraShapeId = readUint16(data, offset + 4);
    const charShapeId = readUint16(data, offset + 6);
    
    return {
      localName,
      englishName,
      type: styleType === 1 ? 'character' : 'paragraph',
      nextStyleId,
      languageId,
      paraShapeId,
      charShapeId,
    };
  }

  private parseFootnoteShape(data: Uint8Array): ParsedFootnoteShape | null {
    if (data.length < 26) return null;
    
    const props = readUint32(data, 0);
    const numberType = props & 0x0F;
    const superscript = (props & 0x10) !== 0;
    const numberingTypeVal = (props >> 5) & 0x03;
    const placementVal = (props >> 7) & 0x03;
    const belowText = (props & (1 << 9)) !== 0;
    
    const numberingTypeMap: Record<number, 'continuous' | 'perSection' | 'perPage'> = {
      0: 'continuous', 1: 'perSection', 2: 'perPage'
    };
    const placementMap: Record<number, 'column' | 'pageBottom' | 'endOfDocument' | 'endOfSection'> = {
      0: 'column', 1: 'pageBottom', 2: 'endOfDocument', 3: 'endOfSection'
    };
    
    const userChar = String.fromCharCode(readUint16(data, 4));
    const prefixChar = String.fromCharCode(readUint16(data, 6));
    const suffixChar = String.fromCharCode(readUint16(data, 8));
    const startNumber = readUint16(data, 10);
    const separatorLength = readInt16(data, 12);
    const spaceAbove = readInt16(data, 14);
    const spaceBelow = readInt16(data, 16);
    const spaceBetween = readInt16(data, 18);
    const separatorLineType = data[20];
    const separatorLineWidth = data[21];
    const separatorLineColor = readUint32(data, 22);
    
    return {
      numberType,
      superscript,
      numberingType: numberingTypeMap[numberingTypeVal] || 'continuous',
      placement: placementMap[placementVal] || 'column',
      belowText,
      userChar,
      prefixChar,
      suffixChar,
      startNumber,
      separatorLength: hwpunitToPt(separatorLength),
      spaceAbove: hwpunitToPt(spaceAbove),
      spaceBelow: hwpunitToPt(spaceBelow),
      spaceBetween: hwpunitToPt(spaceBetween),
      separatorLineType,
      separatorLineWidth,
      separatorLineColor,
    };
  }

  private parseCompatibleDocument(data: Uint8Array): ParsedCompatibleDocument {
    const targetProgram = readUint32(data, 0);
    return { targetProgram };
  }

  private parseLayoutCompatibility(data: Uint8Array): ParsedLayoutCompatibility {
    return {
      lineWrap: readUint32(data, 0),
      charUnit: readUint32(data, 4),
      paraBottomSpacing: readUint32(data, 8),
      underline: readUint32(data, 12),
      strikeout: readUint32(data, 16),
    };
  }

  private parseBinData(): void {
    if (!this._cfb) return;
    
    for (const entry of this._cfb.FileIndex) {
      const fullPath = entry.name;
      if (!fullPath.startsWith('BIN') || !entry.content) continue;
      
      const match = fullPath.match(/BIN(\d+)/i);
      if (!match) continue;
      
      const binId = parseInt(match[1], 10);
      const binInfo = this._binDataInfos.get(binId);
      
      let imageData = entry.content instanceof Uint8Array 
        ? entry.content 
        : new Uint8Array(entry.content as unknown as ArrayBuffer);
      
      if (this._compressed && imageData.length > 0) {
        try {
          imageData = pako.inflateRaw(imageData);
        } catch {}
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
      
      this._content.images.set(idStr, {
        id: idStr,
        binaryId: idStr,
        width: 200,
        height: 150,
        data: dataUrl,
        mimeType,
      });
      
      this._content.binData.set(idStr, { id: idStr, data: dataUrl });
    }
  }

  private parseBodyText(): void {
    let sectionIndex = 0;
    while (true) {
      let data = this.getEntryData(`/BodyText/Section${sectionIndex}`);
      if (!data) break;
      if (this._compressed) data = this.decompress(data);
      this._content.sections.push(this.parseSection(data));
      sectionIndex++;
    }
    
    if (this._content.sections.length === 0) {
      this._content.sections.push({
        elements: [{ type: 'paragraph', data: { id: generateId(), runs: [{ text: '' }] } }],
        pageSettings: this.getDefaultPageSettings(),
      });
    }
  }

  private parseRecordHeader(data: Uint8Array, offset: number): { tagId: number; level: number; size: number; nextOffset: number } | null {
    if (offset + 4 > data.length) return null;
    
    const header = readUint32(data, offset);
    const tagId = header & 0x3FF;
    const level = (header >>> 10) & 0x3FF;
    let size = (header >>> 20) & 0xFFF;
    let nextOffset = offset + 4;
    
    if (size === 0xFFF) {
      if (nextOffset + 4 > data.length) return null;
      size = readUint32(data, nextOffset);
      nextOffset += 4;
    }
    
    return { tagId, level, size, nextOffset };
  }

  private parseSection(data: Uint8Array): HwpxSection {
    const section: HwpxSection = {
      elements: [],
      pageSettings: this.getDefaultPageSettings(),
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
      faceNames: this._faceNames,
      charShapes: this._charShapes,
      paraShapes: this._paraShapes,
      borderFills: this._borderFills,
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
    };
    
    let offset = 0;
    let pageBorderFill: ParsedPageBorderFill | undefined;
    let prevLevel = 0;
    let shapeTextLevel = -1;
    
    while (offset < data.length) {
      const result = this.parseRecordHeader(data, offset);
      if (!result) break;
      
      const { tagId, level, size, nextOffset } = result;
      const recordData = data.slice(nextOffset, nextOffset + size);
      
      if (level < prevLevel) {
        this.finalizeNestedContext(ctx, section, prevLevel - level);
      }
      prevLevel = level;
      
      const _isParaSubTag = tagId >= HWP_TAGS.HWPTAG_PARA_HEADER && tagId <= HWP_TAGS.HWPTAG_PARA_RANGE_TAG;
      if (ctx.inShapeText && shapeTextLevel >= 0 && level <= shapeTextLevel && !_isParaSubTag) {
        if (ctx.currentParagraph) {
          this.flushPendingTextSegments(ctx);
          if (ctx.currentParagraph.runs.length === 0) {
            ctx.currentParagraph.runs.push({ text: '' });
          }
          ctx.shapeTextParagraphs.push(ctx.currentParagraph);
          ctx.currentParagraph = null;
        }
        ctx.inShapeText = false;
        shapeTextLevel = -1;
      }
      
      switch (tagId) {
        case HWP_TAGS.HWPTAG_PARA_HEADER:
          this.handleParaHeader(ctx, section, recordData, level);
          break;
          
        case HWP_TAGS.HWPTAG_PARA_TEXT:
          this.handleParaText(ctx, recordData);
          break;
          
        case HWP_TAGS.HWPTAG_PARA_CHAR_SHAPE:
          this.handleParaCharShape(ctx, recordData);
          break;
          
        case HWP_TAGS.HWPTAG_CTRL_HEADER:
          shapeTextLevel = this.handleCtrlHeader(ctx, recordData, level, shapeTextLevel);
          break;
          
        case HWP_TAGS.HWPTAG_TABLE:
          this.handleTable(ctx, recordData);
          break;
          
        case HWP_TAGS.HWPTAG_LIST_HEADER:
          this.handleListHeader(ctx, section);
          break;
          
        case HWP_TAGS.HWPTAG_SHAPE_COMPONENT:
          this.handleShapeComponent(ctx, recordData);
          break;
          
        case HWP_TAGS.HWPTAG_SHAPE_COMPONENT_PICTURE:
          this.handlePicture(ctx, recordData, section);
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
          
        case HWP_TAGS.HWPTAG_PARA_LINE_SEG:
          this.handleParaLineSeg(ctx, recordData);
          break;
          
        case HWP_TAGS.HWPTAG_PARA_RANGE_TAG:
          this.handleParaRangeTag(ctx, recordData);
          break;
          
        case HWP_TAGS.HWPTAG_PAGE_BORDER_FILL:
          if (recordData.length >= 14) {
            const parsed = this.parsePageBorderFill(recordData);
            if (parsed) pageBorderFill = parsed;
          }
          break;
          
        case HWP_TAGS.HWPTAG_SHAPE_COMPONENT_LINE:
          this.handleShapeLine(ctx, recordData, section);
          break;
          
        case HWP_TAGS.HWPTAG_SHAPE_COMPONENT_RECTANGLE:
          this.handleShapeRectangle(ctx, recordData, section);
          break;
          
        case HWP_TAGS.HWPTAG_SHAPE_COMPONENT_ELLIPSE:
          this.handleShapeEllipse(ctx, recordData, section);
          break;
          
        case HWP_TAGS.HWPTAG_SHAPE_COMPONENT_ARC:
          this.handleShapeArc(ctx, recordData, section);
          break;
          
        case HWP_TAGS.HWPTAG_SHAPE_COMPONENT_POLYGON:
          this.handleShapePolygon(ctx, recordData, section);
          break;
          
        case HWP_TAGS.HWPTAG_SHAPE_COMPONENT_CURVE:
          this.handleShapeCurve(ctx, recordData, section);
          break;
          
        case HWP_TAGS.HWPTAG_SHAPE_COMPONENT_CONTAINER:
          this.handleShapeContainer(ctx, recordData, section);
          break;
          
        case HWP_TAGS.HWPTAG_EQEDIT:
          this.handleEquation(ctx, recordData, section);
          break;
          
        case HWP_TAGS.HWPTAG_SHAPE_COMPONENT_OLE:
          this.handleOle(ctx, recordData, section);
          break;
          
        case HWP_TAGS.HWPTAG_SHAPE_COMPONENT_TEXTART:
          this.handleTextArt(ctx, recordData, section);
          break;
          
        case HWP_TAGS.HWPTAG_MEMO_SHAPE:
          this.handleMemoShape(ctx, recordData);
          break;
          
        case HWP_TAGS.HWPTAG_MEMO_LIST:
          this.handleMemoList(ctx, recordData);
          break;
          
        case HWP_TAGS.HWPTAG_FORM_OBJECT:
          this.handleFormObject(ctx, recordData);
          break;
          
        case HWP_TAGS.HWPTAG_TRACK_CHANGE:
          this.handleTrackChange(ctx, recordData);
          break;
          
        case HWP_TAGS.HWPTAG_TRACK_CHANGE_AUTHOR:
          this.handleTrackChangeAuthor(ctx, recordData);
          break;
          
        case HWP_TAGS.HWPTAG_VIDEO_DATA:
          this.handleVideoData(ctx, recordData, section);
          break;
          
        case HWP_TAGS.HWPTAG_CHART_DATA:
          this.handleChartData(ctx, recordData, section);
          break;
      }
      
      offset = nextOffset + size;
    }
    
    if (ctx.inMemo && ctx.currentMemo) {
      this.finishMemo(ctx);
    }
    
    if (ctx.inTableCell && ctx.currentTable) {
      if (ctx.currentParagraph) {
        this.flushPendingTextSegments(ctx);
        if (ctx.currentParagraph.runs.length > 0) {
          ctx.cellParagraphs.push(ctx.currentParagraph);
        }
      }
      const row = ctx.currentTableRow;
      const col = ctx.currentTableCol;
      if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]) {
        ctx.tableCells[row][col].paragraphs = [...ctx.cellParagraphs];
      }
      this.finishTable(ctx, section);
    } else if (ctx.currentParagraph) {
      this.flushPendingTextSegments(ctx);
      if (ctx.currentParagraph.runs.length > 0) {
        section.elements.push({ type: 'paragraph', data: ctx.currentParagraph });
      }
    }
    
    if (section.elements.length === 0) {
      section.elements.push({ type: 'paragraph', data: { id: generateId(), runs: [{ text: '' }] } });
    }
    
    if (pageBorderFill) {
      (section as any).pageBorderFill = pageBorderFill;
    }
    
    if (ctx.memos.length > 0) {
      section.memos = ctx.memos;
      console.log(`[HWP DEBUG] Section complete: ${ctx.memos.length} memos saved to section`);
    }
    
    return section;
  }

  private flushPendingTextSegments(ctx: ParseContext): void {
    if (!ctx.currentParagraph || ctx.pendingTextSegments.length === 0) return;
    for (const segment of ctx.pendingTextSegments) {
      let applicableCharShapeId = ctx.currentCharShapeId;
      for (const pos of ctx.charShapePositions) {
        if (pos.startPos <= segment.start) {
          applicableCharShapeId = pos.charShapeId;
        }
      }
      const charShape = ctx.charShapes.get(applicableCharShapeId);
      const charStyle = charShape ? this.charShapeToStyle(charShape, ctx.faceNames) : undefined;
      ctx.currentParagraph.runs.push({ text: segment.text, charStyle });
    }
    ctx.pendingTextSegments = [];
  }

  private handleParaHeader(ctx: ParseContext, section: HwpxSection, recordData?: Uint8Array, level?: number): void {
    this.flushPendingTextSegments(ctx);
    if (ctx.inHeaderFooter && ctx.currentHeaderFooter) {
      if (ctx.currentParagraph?.runs.length) {
        ctx.headerFooterParagraphs.push(ctx.currentParagraph);
      }
    } else if (ctx.inFootnoteEndnote && ctx.currentFootnoteEndnote) {
      if (ctx.currentParagraph?.runs.length) {
        ctx.footnoteEndnoteParagraphs.push(ctx.currentParagraph);
      }
    } else if (ctx.inMemo && ctx.currentMemo) {
      if (ctx.currentParagraph?.runs.length) {
        ctx.memoParagraphs.push(ctx.currentParagraph);
      }
    } else if (ctx.inShapeText) {
      if (ctx.currentParagraph) {
        if (ctx.currentParagraph.runs.length === 0) {
          ctx.currentParagraph.runs.push({ text: '' });
        }
        ctx.shapeTextParagraphs.push(ctx.currentParagraph);
      }
    } else if (ctx.inTableCell) {
      if (ctx.currentParagraph) {
        if (ctx.currentParagraph.runs.length === 0) {
          ctx.currentParagraph.runs.push({ text: '' });
        }
        ctx.cellParagraphs.push(ctx.currentParagraph);
      }
    } else if (level === undefined || level === 0) {
      if (ctx.currentParagraph?.runs.length) {
        section.elements.push({ type: 'paragraph', data: ctx.currentParagraph });
      }
    }
    
    ctx.currentParagraph = { id: generateId(), runs: [] };
    ctx.charShapePositions = [];
    ctx.pendingTextSegments = [];
    
    if (recordData && recordData.length >= 12) {
      ctx.currentParaShapeId = readUint16(recordData, 8);
      ctx.currentStyleId = recordData[10];
      
      const paraShape = ctx.paraShapes.get(ctx.currentParaShapeId);
      if (paraShape) {
        const alignMap: Record<number, 'Justify' | 'Left' | 'Right' | 'Center' | 'Distribute'> = {
          0: 'Justify', 1: 'Left', 2: 'Right', 3: 'Center', 4: 'Distribute', 5: 'Distribute'
        };
        ctx.currentParagraph.paraStyle = {
          align: alignMap[paraShape.alignment] || 'Justify',
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
  }

  private handleParaCharShape(ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 8) return;
    
    ctx.charShapePositions = [];
    for (let i = 0; i + 8 <= data.length; i += 8) {
      const startPos = readUint32(data, i);
      const charShapeId = readUint32(data, i + 4);
      ctx.charShapePositions.push({ startPos, charShapeId });
    }
    
    if (ctx.charShapePositions.length > 0) {
      ctx.currentCharShapeId = ctx.charShapePositions[0].charShapeId;
    }
  }

  private handleParaText(ctx: ParseContext, data: Uint8Array): void {
    if (!ctx.currentParagraph) return;
    
    let currentStart = 0;
    let currentText = '';
    let charIndex = 0;
    let i = 0;
    
    while (i < data.length - 1) {
      const charCode = readUint16(data, i);
      i += 2;
      
      if (charCode === 0) {
        charIndex++;
        continue;
      }
      
      if (charCode < 32) {
        if (currentText) {
          ctx.pendingTextSegments.push({ start: currentStart, end: charIndex, text: currentText });
          currentText = '';
        }
        
         if (charCode === CTRL_CHAR.LINE_BREAK) {
           ctx.pendingTextSegments.push({ start: charIndex, end: charIndex + 1, text: '\n' });
           charIndex++;
         } else if (charCode === 0x0009) {
           ctx.pendingTextSegments.push({ start: charIndex, end: charIndex + 8, text: '\t' });
           i += 14;
           charIndex += 8;
         } else if (charCode >= 0x0002 && charCode <= 0x0008) {
           i += 14;
           charIndex += 8;
         } else if (charCode === 0x000B || charCode === 0x000C ||
                    charCode === 0x000E || charCode === 0x000F || charCode === 0x0010 ||
                    charCode === 0x0011 || charCode === 0x0012 || charCode === 0x0013 ||
                    charCode === 0x0015 || charCode === 0x0016 || charCode === 0x0017) {
           i += 14;
           charIndex += 8;
         } else if (charCode === CTRL_CHAR.PARAGRAPH_BREAK) {
           break;
         } else {
           charIndex++;
         }
         currentStart = charIndex;
         continue;
      }
      
      currentText += String.fromCharCode(charCode);
      charIndex++;
    }
    
    if (currentText) {
      ctx.pendingTextSegments.push({ start: currentStart, end: charIndex, text: currentText });
    }
  }

  private charShapeToStyle(charShape: ParsedCharShape, faceNames: Map<number, ParsedFaceName>): import('../hwpx/types').CharacterStyle {
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

  private handleCtrlHeader(ctx: ParseContext, data: Uint8Array, level: number, shapeTextLevel: number): number {
    if (data.length < 4) return shapeTextLevel;
    
    ctx.currentCtrlId = readUint32(data, 0);
    
    if (ctx.currentCtrlId === CTRL_ID.TABLE) {
       ctx.currentTable = { id: generateId(), rows: [], rowCount: 0, colCount: 0 };
       ctx.tableCells = [];
       ctx.inTableCell = false;
     } else if (ctx.currentCtrlId === CTRL_ID.PICTURE || ctx.currentCtrlId === CTRL_ID.GSO) {
      ctx.pendingImage = { width: 200, height: 150 };
      
      if (data.length >= 46) {
        const width = readUint32(data, 16);
        const height = readUint32(data, 20);
        if (width > 0) ctx.pendingImage.width = width / 7200 * 72;
        if (height > 0) ctx.pendingImage.height = height / 7200 * 72;
      }
      ctx.inShapeText = true;
      ctx.shapeTextParagraphs = [];
      shapeTextLevel = level;
    } else if (ctx.currentCtrlId === CTRL_ID.SECTION_DEF) {
      this.handleSectionDef(ctx, data);
    } else if (ctx.currentCtrlId === CTRL_ID.COLUMN_DEF) {
      this.handleColumnDef(ctx, data);
    } else if (ctx.currentCtrlId === CTRL_ID.HEADER || ctx.currentCtrlId === CTRL_ID.FOOTER) {
      this.handleHeaderFooterStart(ctx, data);
    } else if (ctx.currentCtrlId === CTRL_ID.FOOTNOTE || ctx.currentCtrlId === CTRL_ID.ENDNOTE) {
      this.handleFootnoteEndnoteStart(ctx, data);
    } else if (this.isShapeCtrlId(ctx.currentCtrlId)) {
      this.handleShapeStart(ctx, data);
      ctx.inShapeText = true;
      ctx.shapeTextParagraphs = [];
      shapeTextLevel = level;
    } else if (ctx.currentCtrlId === CTRL_ID.AUTO_NUMBER) {
      this.handleAutoNumber(ctx, data);
    } else if (ctx.currentCtrlId === CTRL_ID.NEW_NUMBER) {
      this.handleNewNumber(ctx, data);
    } else if (ctx.currentCtrlId === CTRL_ID.PAGE_NUMBER_POS) {
      this.handlePageNumberPos(ctx, data);
    } else if (ctx.currentCtrlId === CTRL_ID.BOOKMARK) {
      this.handleBookmark(ctx, data);
    } else if (ctx.currentCtrlId === CTRL_ID.INDEX_MARK) {
      this.handleIndexMark(ctx, data);
    } else if (ctx.currentCtrlId === CTRL_ID.CHAR_OVERLAP) {
      this.handleCharOverlap(ctx, data);
    } else if (ctx.currentCtrlId === CTRL_ID.ANNOTATION) {
      this.handleAnnotation(ctx, data);
    } else if (ctx.currentCtrlId === CTRL_ID.HIDDEN_COMMENT) {
      this.handleHiddenComment(ctx, data);
    } else if (ctx.currentCtrlId === CTRL_ID.PAGE_HIDE) {
      this.handlePageHide(ctx, data);
    } else if (ctx.currentCtrlId === CTRL_ID.PAGE_ODD_EVEN) {
      this.handlePageOddEven(ctx, data);
    } else if (this.isFieldCtrlId(ctx.currentCtrlId)) {
      this.handleFieldStart(ctx, data);
    } else if (ctx.currentCtrlId !== CTRL_ID.SECTION_DEF
               && ctx.currentCtrlId !== CTRL_ID.COLUMN_DEF
               && ctx.currentCtrlId !== CTRL_ID.AUTO_NUMBER
               && ctx.currentCtrlId !== CTRL_ID.PAGE_NUMBER_POS
               && ctx.currentCtrlId !== 0x65716564 /* EQUATION */
               && ctx.currentCtrlId !== CTRL_ID.BOOKMARK
               && ctx.currentCtrlId !== 0x74637073 /* TCPS */
               && ctx.currentCtrlId !== CTRL_ID.FORM) {
      ctx.inShapeText = true;
      ctx.shapeTextParagraphs = [];
      shapeTextLevel = level;
    }
    return shapeTextLevel;
  }

  private isShapeCtrlId(ctrlId: number): boolean {
    return ctrlId === CTRL_ID.LINE || ctrlId === CTRL_ID.RECTANGLE ||
           ctrlId === CTRL_ID.ELLIPSE || ctrlId === CTRL_ID.ARC ||
           ctrlId === CTRL_ID.POLYGON || ctrlId === CTRL_ID.CURVE ||
           ctrlId === CTRL_ID.CONTAINER;
  }

  private isFieldCtrlId(ctrlId: number): boolean {
    return ctrlId === CTRL_ID.FIELD_UNKNOWN || ctrlId === CTRL_ID.FIELD_DATE ||
           ctrlId === CTRL_ID.FIELD_DOCDATE || ctrlId === CTRL_ID.FIELD_PATH ||
           ctrlId === CTRL_ID.FIELD_BOOKMARK || ctrlId === CTRL_ID.FIELD_MAILMERGE ||
           ctrlId === CTRL_ID.FIELD_CROSSREF || ctrlId === CTRL_ID.FIELD_FORMULA ||
           ctrlId === CTRL_ID.FIELD_CLICKHERE || ctrlId === CTRL_ID.FIELD_SUMMARY ||
           ctrlId === CTRL_ID.FIELD_USERINFO || ctrlId === CTRL_ID.FIELD_HYPERLINK ||
           ctrlId === CTRL_ID.FIELD_MEMO || ctrlId === CTRL_ID.FIELD_PRIVATE_INFO ||
           ctrlId === CTRL_ID.FIELD_TOC;
  }

  private handleAutoNumber(ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 12) return;
    
    const properties = readUint32(data, 4);
    const number = readUint16(data, 8);
    const numberType = properties & 0x0F;
    const numberShape = (properties >> 4) & 0xFF;
    const isSuperscript = (properties >> 12) & 0x01;
    
    ctx.pendingField = {
      type: 'autoNumber',
      numberType,
      numberShape,
      number,
      isSuperscript: isSuperscript === 1
    };
  }

  private handleNewNumber(ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 8) return;
    
    const properties = readUint32(data, 4);
    const number = readUint16(data, 6);
    const numberType = properties & 0x0F;
    
    ctx.pendingField = {
      type: 'newNumber',
      numberType,
      number
    };
  }

  private handlePageNumberPos(ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 12) return;
    
    const properties = readUint32(data, 4);
    const numberShape = properties & 0xFF;
    const position = (properties >> 8) & 0x0F;
    
    ctx.pendingField = {
      type: 'pageNumberPos',
      numberShape,
      position
    };
  }

  private handleBookmark(ctx: ParseContext, _data: Uint8Array): void {
    ctx.pendingField = {
      type: 'bookmark',
      name: ''
    };
  }

  private handleIndexMark(ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 6) return;
    
    let offset = 4;
    const keyword1Len = readUint16(data, offset);
    offset += 2;
    
    let keyword1 = '';
    if (keyword1Len > 0 && data.length >= offset + keyword1Len * 2) {
      keyword1 = readWString(data, offset, keyword1Len);
      offset += keyword1Len * 2;
    }
    
    let keyword2 = '';
    if (data.length >= offset + 2) {
      const keyword2Len = readUint16(data, offset);
      offset += 2;
      if (keyword2Len > 0 && data.length >= offset + keyword2Len * 2) {
        keyword2 = readWString(data, offset, keyword2Len);
      }
    }
    
    ctx.pendingField = {
      type: 'indexMark',
      keyword1,
      keyword2
    };
  }

  private handleCharOverlap(ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 10) return;
    
    let offset = 4;
    const textLen = readUint16(data, offset);
    offset += 2;
    
    let overlappingText = '';
    if (textLen > 0 && data.length >= offset + textLen * 2) {
      overlappingText = readWString(data, offset, textLen);
      offset += textLen * 2;
    }
    
    const borderType = data.length > offset ? data[offset] : 0;
    const fontSize = data.length > offset + 1 ? readInt8(data, offset + 1) : 0;
    const expand = data.length > offset + 2 ? data[offset + 2] : 0;
    
    ctx.pendingField = {
      type: 'charOverlap',
      name: overlappingText,
      numberType: borderType,
      number: fontSize,
      position: expand
    };
  }

  private handleAnnotation(ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 18) return;
    
    let offset = 0;
    const mainTextLen = readUint16(data, offset);
    offset += 2;
    
    let mainText = '';
    if (mainTextLen > 0 && data.length >= offset + mainTextLen * 2) {
      mainText = readWString(data, offset, mainTextLen);
      offset += mainTextLen * 2;
    }
    
    const subTextLen = readUint16(data, offset);
    offset += 2;
    
    let subText = '';
    if (subTextLen > 0 && data.length >= offset + subTextLen * 2) {
      subText = readWString(data, offset, subTextLen);
      offset += subTextLen * 2;
    }
    
    const position = data.length >= offset + 4 ? readUint32(data, offset) : 0;
    
    ctx.pendingField = {
      type: 'annotation',
      name: mainText,
      keyword1: subText,
      position
    };
  }

  private handleHiddenComment(ctx: ParseContext, _data: Uint8Array): void {
    ctx.pendingField = {
      type: 'hiddenComment'
    };
  }

  private handlePageHide(ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 2) return;
    
    const hideFlags = readUint16(data, 0);
    
    ctx.pendingField = {
      type: 'pageHide',
      properties: hideFlags
    };
  }

  private handlePageOddEven(ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 4) return;
    
    const properties = readUint32(data, 0);
    const oddEvenType = properties & 0x03;
    
    ctx.pendingField = {
      type: 'pageOddEven',
      numberType: oddEvenType
    };
  }

  private handleFieldStart(ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 15) return;
    
    const properties = readUint32(data, 4);
    const etcProperties = data[8];
    const commandLen = readUint16(data, 9);
    
    let command = '';
    if (commandLen > 0 && data.length >= 11 + commandLen * 2) {
      command = readWString(data, 11, commandLen);
    }
    
    let fieldType = 'unknown';
    if (ctx.currentCtrlId === CTRL_ID.FIELD_HYPERLINK) fieldType = 'hyperlink';
    else if (ctx.currentCtrlId === CTRL_ID.FIELD_DATE) fieldType = 'date';
    else if (ctx.currentCtrlId === CTRL_ID.FIELD_DOCDATE) fieldType = 'docDate';
    else if (ctx.currentCtrlId === CTRL_ID.FIELD_PATH) fieldType = 'path';
    else if (ctx.currentCtrlId === CTRL_ID.FIELD_BOOKMARK) fieldType = 'bookmarkRef';
    else if (ctx.currentCtrlId === CTRL_ID.FIELD_MAILMERGE) fieldType = 'mailMerge';
    else if (ctx.currentCtrlId === CTRL_ID.FIELD_CROSSREF) fieldType = 'crossRef';
    else if (ctx.currentCtrlId === CTRL_ID.FIELD_FORMULA) fieldType = 'formula';
    else if (ctx.currentCtrlId === CTRL_ID.FIELD_CLICKHERE) fieldType = 'clickHere';
    else if (ctx.currentCtrlId === CTRL_ID.FIELD_SUMMARY) fieldType = 'summary';
    else if (ctx.currentCtrlId === CTRL_ID.FIELD_USERINFO) fieldType = 'userInfo';
    else if (ctx.currentCtrlId === CTRL_ID.FIELD_MEMO) fieldType = 'memo';
    else if (ctx.currentCtrlId === CTRL_ID.FIELD_PRIVATE_INFO) fieldType = 'privateInfo';
    else if (ctx.currentCtrlId === CTRL_ID.FIELD_TOC) fieldType = 'tableOfContents';
    
    ctx.pendingField = {
      type: fieldType,
      command,
      properties,
      readOnlyEditable: (properties & 0x01) === 1,
      hyperlinkUpdateType: (properties >> 11) & 0x0F,
      modified: ((properties >> 15) & 0x01) === 1,
      etcProperties
    };
    
    if (fieldType === 'memo') {
      if (ctx.inMemo && ctx.currentMemo) {
        this.finishMemo(ctx);
      }
      ctx.inMemo = true;
      ctx.currentMemo = {
        id: generateId(),
        linkedText: '',
        content: []
      };
      ctx.memoParagraphs = [];
    }
  }

  private handleMemoShape(_ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 22) return;
    
    const memoId = readUint32(data, 0);
    const width = hwpunitToPt(readUint32(data, 4));
    const height = hwpunitToPt(readUint32(data, 8));
    const lineType = data[12];
    const lineColor = colorrefToHex(readUint32(data, 13));
    const fillColor = colorrefToHex(readUint32(data, 17));
    const activeColor = data[21];
    
    this._memoShapes.set(memoId, { memoId, width, height, lineType, lineColor, fillColor, activeColor });
  }

  private handleMemoList(ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 4) return;
    
    const memoCount = readUint32(data, 0);
    this._memoCount = memoCount;
    console.log(`[HWP DEBUG] MEMO_LIST: memoCount=${memoCount}, currentMemos=${ctx.memos.length}`);
    
    if (ctx.inMemo && ctx.currentMemo) {
      this.finishMemo(ctx);
    }
    
    ctx.inMemo = true;
    ctx.currentMemo = {
      id: generateId(),
      linkedText: '',
      content: []
    };
    ctx.memoParagraphs = [];
  }
  
  private finishMemo(ctx: ParseContext): void {
    if (!ctx.currentMemo) return;
    
    if (ctx.currentParagraph) {
      this.flushPendingTextSegments(ctx);
      if (ctx.currentParagraph.runs.length > 0) {
        ctx.memoParagraphs.push(ctx.currentParagraph);
        ctx.currentParagraph = null;
      }
    }
    
    const memoContent: string[] = [];
    for (const para of ctx.memoParagraphs) {
      const text = para.runs.map(r => r.text || '').join('');
      if (text) memoContent.push(text);
    }
    
    const memo: import('../hwpx/types').Memo = {
      id: ctx.currentMemo.id,
      author: 'Unknown',
      date: '',
      content: memoContent.length > 0 ? memoContent : ['(빈 메모)'],
      linkedText: '',
    };
    ctx.memos.push(memo);
    console.log(`[HWP DEBUG] finishMemo: paragraphs=${ctx.memoParagraphs.length}, content="${memoContent.join(' | ')}", totalMemos=${ctx.memos.length}`);
    
    ctx.inMemo = false;
    ctx.currentMemo = null;
    ctx.memoParagraphs = [];
  }

  private handleFormObject(_ctx: ParseContext, _data: Uint8Array): void {
  }

  private handleTrackChange(_ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 8) return;
    
    const trackChangeId = readUint32(data, 0);
    const trackChangeType = readUint32(data, 4);
    
    this._trackChanges.push({ id: trackChangeId, type: trackChangeType });
  }

  private handleTrackChangeAuthor(_ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 4) return;
    
    let offset = 0;
    const nameLen = readUint16(data, offset);
    offset += 2;
    
    let name = '';
    if (nameLen > 0 && data.length >= offset + nameLen * 2) {
      name = readWString(data, offset, nameLen);
    }
    
    this._trackChangeAuthors.push({ name });
  }

  private handleVideoData(ctx: ParseContext, data: Uint8Array, section: HwpxSection): void {
    if (data.length < 4) return;
    
    const videoBinDataId = readUint16(data, 0);
    const thumbnailBinDataId = readUint16(data, 2);
    
    const video: import('../hwpx/types').HwpxVideo = {
      id: generateId(),
      binDataId: videoBinDataId,
      thumbnailBinDataId,
    };
    
    if (ctx.inTableCell && ctx.currentTable) {
      const row = ctx.currentTableRow;
      const col = ctx.currentTableCol;
      if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]?.[col]) {
        if (!ctx.tableCells[row][col].elements) ctx.tableCells[row][col].elements = [];
        ctx.tableCells[row][col].elements!.push({ type: 'video', data: video });
      }
    } else if (!ctx.inHeaderFooter && !ctx.inFootnoteEndnote) {
      section.elements.push({ type: 'video', data: video });
    }
  }

  private handleChartData(ctx: ParseContext, data: Uint8Array, section: HwpxSection): void {
    if (data.length < 4) return;
    
    const chart: import('../hwpx/types').HwpxChart = {
      id: generateId(),
      rawData: data,
    };
    
    if (ctx.inTableCell && ctx.currentTable) {
      const row = ctx.currentTableRow;
      const col = ctx.currentTableCol;
      if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]?.[col]) {
        if (!ctx.tableCells[row][col].elements) ctx.tableCells[row][col].elements = [];
        ctx.tableCells[row][col].elements!.push({ type: 'chart', data: chart });
      }
    } else if (!ctx.inHeaderFooter && !ctx.inFootnoteEndnote) {
      section.elements.push({ type: 'chart', data: chart });
    }
  }

  private handleShapeStart(ctx: ParseContext, data: Uint8Array): void {
    let shapeType = 'unknownobject';
    if (ctx.currentCtrlId === CTRL_ID.LINE) shapeType = 'line';
    else if (ctx.currentCtrlId === CTRL_ID.RECTANGLE) shapeType = 'rect';
    else if (ctx.currentCtrlId === CTRL_ID.ELLIPSE) shapeType = 'ellipse';
    else if (ctx.currentCtrlId === CTRL_ID.ARC) shapeType = 'arc';
    else if (ctx.currentCtrlId === CTRL_ID.POLYGON) shapeType = 'polygon';
    else if (ctx.currentCtrlId === CTRL_ID.CURVE) shapeType = 'curve';
    else if (ctx.currentCtrlId === CTRL_ID.CONTAINER) shapeType = 'container';
    
    ctx.pendingShape = { type: shapeType, width: 100, height: 100, x: 0, y: 0 };
    
    if (data.length >= 24) {
      const width = readUint32(data, 16);
      const height = readUint32(data, 20);
      if (width > 0) ctx.pendingShape.width = hwpunitToPt(width);
      if (height > 0) ctx.pendingShape.height = hwpunitToPt(height);
    }
  }

  private handleShapeLine(ctx: ParseContext, data: Uint8Array, section: HwpxSection): void {
    if (data.length < 18) return;
    
    const startX = hwpunitToPt(readInt32(data, 0));
    const startY = hwpunitToPt(readInt32(data, 4));
    const endX = hwpunitToPt(readInt32(data, 8));
    const endY = hwpunitToPt(readInt32(data, 12));
    
    const line: import('../hwpx/types').HwpxLine = {
      id: generateId(),
      startX, startY, endX, endY,
    };
    
    if (ctx.inTableCell && ctx.currentTable) {
      const row = ctx.currentTableRow;
      const col = ctx.currentTableCol;
      if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]?.[col]) {
        if (!ctx.tableCells[row][col].elements) ctx.tableCells[row][col].elements = [];
        ctx.tableCells[row][col].elements!.push({ type: 'line', data: line });
      }
    } else if (!ctx.inHeaderFooter && !ctx.inFootnoteEndnote) {
      section.elements.push({ type: 'line', data: line });
    }
    ctx.pendingShape = null;
  }

  private handleShapeRectangle(ctx: ParseContext, data: Uint8Array, section: HwpxSection): void {
    if (data.length < 33) return;
    
    const cornerRadius = data[0];
    const xCoords = [
      hwpunitToPt(readInt32(data, 1)),
      hwpunitToPt(readInt32(data, 5)),
      hwpunitToPt(readInt32(data, 9)),
      hwpunitToPt(readInt32(data, 13)),
    ];
    const yCoords = [
      hwpunitToPt(readInt32(data, 17)),
      hwpunitToPt(readInt32(data, 21)),
      hwpunitToPt(readInt32(data, 25)),
      hwpunitToPt(readInt32(data, 29)),
    ];
    
    const minX = Math.min(...xCoords);
    const maxX = Math.max(...xCoords);
    const minY = Math.min(...yCoords);
    const maxY = Math.max(...yCoords);
    
     const rect: import('../hwpx/types').HwpxRect = {
       id: generateId(),
       x: minX,
       y: minY,
       width: maxX - minX || ctx.pendingShape?.width || 100,
       height: maxY - minY || ctx.pendingShape?.height || 100,
       cornerRadius,
     };
     
     if (ctx.inTableCell && ctx.currentTable) {
       const row = ctx.currentTableRow;
       const col = ctx.currentTableCol;
       if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]?.[col]) {
         if (!ctx.tableCells[row][col].elements) ctx.tableCells[row][col].elements = [];
         ctx.tableCells[row][col].elements!.push({ type: 'rect', data: rect });
       }
     } else if (!ctx.inHeaderFooter && !ctx.inFootnoteEndnote) {
       section.elements.push({ type: 'rect', data: rect });
     }
     ctx.pendingShape = null;
  }

  private handleShapeEllipse(ctx: ParseContext, data: Uint8Array, section: HwpxSection): void {
    if (data.length < 60) return;
    
    const centerX = hwpunitToPt(readInt32(data, 4));
    const centerY = hwpunitToPt(readInt32(data, 8));
    const axis1X = hwpunitToPt(readInt32(data, 12));
    const axis1Y = hwpunitToPt(readInt32(data, 16));
    const axis2X = hwpunitToPt(readInt32(data, 20));
    const axis2Y = hwpunitToPt(readInt32(data, 24));
    
    const rx = Math.sqrt(axis1X * axis1X + axis1Y * axis1Y) || (ctx.pendingShape?.width || 100) / 2;
    const ry = Math.sqrt(axis2X * axis2X + axis2Y * axis2Y) || (ctx.pendingShape?.height || 100) / 2;
    
     const ellipse: import('../hwpx/types').HwpxEllipse = {
       id: generateId(),
       centerX, centerY,
       axis1X, axis1Y,
       axis2X, axis2Y,
       cx: centerX,
       cy: centerY,
       rx, ry,
     };
     
     if (ctx.inTableCell && ctx.currentTable) {
       const row = ctx.currentTableRow;
       const col = ctx.currentTableCol;
       if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]?.[col]) {
         if (!ctx.tableCells[row][col].elements) ctx.tableCells[row][col].elements = [];
         ctx.tableCells[row][col].elements!.push({ type: 'ellipse', data: ellipse });
       }
     } else if (!ctx.inHeaderFooter && !ctx.inFootnoteEndnote) {
       section.elements.push({ type: 'ellipse', data: ellipse });
     }
     ctx.pendingShape = null;
  }

  private handleShapeArc(ctx: ParseContext, data: Uint8Array, section: HwpxSection): void {
    if (data.length < 28) return;
    
    const centerX = hwpunitToPt(readInt32(data, 4));
    const centerY = hwpunitToPt(readInt32(data, 8));
    const axis1X = hwpunitToPt(readInt32(data, 12));
    const axis1Y = hwpunitToPt(readInt32(data, 16));
    const axis2X = hwpunitToPt(readInt32(data, 20));
    const axis2Y = hwpunitToPt(readInt32(data, 24));
    
     const arc: import('../hwpx/types').HwpxArc = {
       id: generateId(),
       centerX, centerY,
       axis1X, axis1Y,
       axis2X, axis2Y,
     };
     
     if (ctx.inTableCell && ctx.currentTable) {
       const row = ctx.currentTableRow;
       const col = ctx.currentTableCol;
       if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]?.[col]) {
         if (!ctx.tableCells[row][col].elements) ctx.tableCells[row][col].elements = [];
         ctx.tableCells[row][col].elements!.push({ type: 'arc', data: arc });
       }
     } else if (!ctx.inHeaderFooter && !ctx.inFootnoteEndnote) {
       section.elements.push({ type: 'arc', data: arc });
     }
     ctx.pendingShape = null;
  }

  private handleShapePolygon(ctx: ParseContext, data: Uint8Array, section: HwpxSection): void {
    if (data.length < 2) return;
    
    const pointCount = readInt16(data, 0);
    if (data.length < 2 + pointCount * 8) return;
    
    const points: Array<{ x: number; y: number }> = [];
    let offset = 2;
    
    for (let i = 0; i < pointCount; i++) {
      const x = hwpunitToPt(readInt32(data, offset));
      offset += 4;
      points.push({ x, y: 0 });
    }
    
    for (let i = 0; i < pointCount; i++) {
      const y = hwpunitToPt(readInt32(data, offset));
      offset += 4;
      points[i].y = y;
    }
    
     const polygon: import('../hwpx/types').HwpxPolygon = {
       id: generateId(),
       points,
     };
     
     if (ctx.inTableCell && ctx.currentTable) {
       const row = ctx.currentTableRow;
       const col = ctx.currentTableCol;
       if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]?.[col]) {
         if (!ctx.tableCells[row][col].elements) ctx.tableCells[row][col].elements = [];
         ctx.tableCells[row][col].elements!.push({ type: 'polygon', data: polygon });
       }
     } else if (!ctx.inHeaderFooter && !ctx.inFootnoteEndnote) {
       section.elements.push({ type: 'polygon', data: polygon });
     }
     ctx.pendingShape = null;
  }

  private handleShapeCurve(ctx: ParseContext, data: Uint8Array, section: HwpxSection): void {
    if (data.length < 2) return;
    
    const pointCount = readInt16(data, 0);
    if (data.length < 2 + pointCount * 8) return;
    
    const segments: import('../hwpx/types').CurveSegment[] = [];
    let offset = 2;
    
    const xCoords: number[] = [];
    const yCoords: number[] = [];
    
    for (let i = 0; i < pointCount; i++) {
      xCoords.push(hwpunitToPt(readInt32(data, offset)));
      offset += 4;
    }
    
    for (let i = 0; i < pointCount; i++) {
      yCoords.push(hwpunitToPt(readInt32(data, offset)));
      offset += 4;
    }
    
    for (let i = 0; i < pointCount - 1; i++) {
      segments.push({
        type: 'Line',
        x1: xCoords[i],
        y1: yCoords[i],
        x2: xCoords[i + 1],
        y2: yCoords[i + 1],
      });
    }
    
     const curve: import('../hwpx/types').HwpxCurve = {
       id: generateId(),
       segments,
     };
     
     if (ctx.inTableCell && ctx.currentTable) {
       const row = ctx.currentTableRow;
       const col = ctx.currentTableCol;
       if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]?.[col]) {
         if (!ctx.tableCells[row][col].elements) ctx.tableCells[row][col].elements = [];
         ctx.tableCells[row][col].elements!.push({ type: 'curve', data: curve });
       }
     } else if (!ctx.inHeaderFooter && !ctx.inFootnoteEndnote) {
       section.elements.push({ type: 'curve', data: curve });
     }
     ctx.pendingShape = null;
  }

  private handleShapeContainer(ctx: ParseContext, data: Uint8Array, section: HwpxSection): void {
    if (data.length < 2) return;
    
    const objectCount = readUint16(data, 0);
    
     const container: import('../hwpx/types').HwpxContainer = {
       id: generateId(),
       children: [],
     };
     
     if (ctx.inTableCell && ctx.currentTable) {
       const row = ctx.currentTableRow;
       const col = ctx.currentTableCol;
       if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]?.[col]) {
         if (!ctx.tableCells[row][col].elements) ctx.tableCells[row][col].elements = [];
         ctx.tableCells[row][col].elements!.push({ type: 'container', data: container });
       }
     } else if (!ctx.inHeaderFooter && !ctx.inFootnoteEndnote) {
       section.elements.push({ type: 'container', data: container });
     }
     ctx.pendingShape = null;
  }

  private handleEquation(ctx: ParseContext, data: Uint8Array, section: HwpxSection): void {
    if (data.length < 6) return;
    
    const scriptLen = readUint16(data, 4);
    let script = '';
    
    if (data.length >= 6 + scriptLen * 2) {
      for (let i = 0; i < scriptLen; i++) {
        const charCode = readUint16(data, 6 + i * 2);
        if (charCode > 0) script += String.fromCharCode(charCode);
      }
    }
    
    const equation: import('../hwpx/types').HwpxEquation = {
      id: generateId(),
      script,
    };
    
    if (ctx.inTableCell && ctx.currentTable) {
      const row = ctx.currentTableRow;
      const col = ctx.currentTableCol;
      if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]?.[col]) {
        if (!ctx.tableCells[row][col].elements) ctx.tableCells[row][col].elements = [];
        ctx.tableCells[row][col].elements!.push({ type: 'equation', data: equation });
      }
    } else if (!ctx.inHeaderFooter && !ctx.inFootnoteEndnote) {
      section.elements.push({ type: 'equation', data: equation });
    }
    ctx.pendingShape = null;
  }

  private handleOle(ctx: ParseContext, data: Uint8Array, section: HwpxSection): void {
     const ole: import('../hwpx/types').HwpxOle = {
       id: generateId(),
       extentX: ctx.pendingShape?.width,
       extentY: ctx.pendingShape?.height,
     };
     
     if (ctx.inTableCell && ctx.currentTable) {
       const row = ctx.currentTableRow;
       const col = ctx.currentTableCol;
       if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]?.[col]) {
         if (!ctx.tableCells[row][col].elements) ctx.tableCells[row][col].elements = [];
         ctx.tableCells[row][col].elements!.push({ type: 'ole', data: ole });
       }
     } else if (!ctx.inHeaderFooter && !ctx.inFootnoteEndnote) {
       section.elements.push({ type: 'ole', data: ole });
     }
     ctx.pendingShape = null;
  }

   private handleTextArt(ctx: ParseContext, data: Uint8Array, section: HwpxSection): void {
     const textart: import('../hwpx/types').HwpxTextArt = {
       id: generateId(),
       text: '',
     };
     
     if (ctx.inTableCell && ctx.currentTable) {
       const row = ctx.currentTableRow;
       const col = ctx.currentTableCol;
       if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]?.[col]) {
         if (!ctx.tableCells[row][col].elements) ctx.tableCells[row][col].elements = [];
         ctx.tableCells[row][col].elements!.push({ type: 'textart', data: textart });
       }
     } else if (!ctx.inHeaderFooter && !ctx.inFootnoteEndnote) {
       section.elements.push({ type: 'textart', data: textart });
     }
    ctx.pendingShape = null;
  }

  private handleSectionDef(ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 26) return;
    
    const props = readUint32(data, 0);
    ctx.currentSectionDef = {
      hideHeader: (props & 0x01) !== 0,
      hideFooter: (props & 0x02) !== 0,
      hideMasterPage: (props & 0x04) !== 0,
      hideBorder: (props & 0x08) !== 0,
      hideBackground: (props & 0x10) !== 0,
      hidePageNum: (props & 0x20) !== 0,
      borderOnFirstOnly: (props & 0x100) !== 0,
      backgroundOnFirstOnly: (props & 0x200) !== 0,
      textDirection: (props >> 16) & 0x07,
      columnGap: readUint16(data, 4),
      pageNumber: readUint16(data, 18),
      figureNumber: readUint16(data, 20),
      tableNumber: readUint16(data, 22),
      equationNumber: readUint16(data, 24),
    };
  }

  private handleColumnDef(ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 4) return;
    
    const props = readUint16(data, 0);
    const typeVal = props & 0x03;
    const columnCount = (props >> 2) & 0xFF;
    const directionVal = (props >> 10) & 0x03;
    const sameWidth = (props & (1 << 12)) !== 0;
    
    const typeMap: Record<number, 'normal' | 'distribute' | 'parallel'> = {
      0: 'normal', 1: 'distribute', 2: 'parallel'
    };
    const directionMap: Record<number, 'left' | 'right' | 'facing'> = {
      0: 'left', 1: 'right', 2: 'facing'
    };
    
    ctx.currentColumnDef = {
      columnType: typeMap[typeVal] || 'normal',
      columnCount: columnCount || 1,
      direction: directionMap[directionVal] || 'left',
      sameWidth,
      gap: data.length >= 6 ? hwpunitToPt(readUint16(data, 2)) : 0,
    };
  }

  private handleHeaderFooterStart(ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 14) return;
    
    const isHeader = ctx.currentCtrlId === CTRL_ID.HEADER;
    const props = readUint32(data, 0);
    const applyToVal = props & 0x03;
    const applyToMap: Record<number, 'both' | 'even' | 'odd'> = {
      0: 'both', 1: 'even', 2: 'odd'
    };
    
    ctx.currentHeaderFooter = {
      type: isHeader ? 'header' : 'footer',
      applyTo: applyToMap[applyToVal] || 'both',
      textWidth: hwpunitToPt(readUint32(data, 4)),
      textHeight: hwpunitToPt(readUint32(data, 8)),
      paragraphs: [],
    };
    ctx.inHeaderFooter = true;
    ctx.headerFooterParagraphs = [];
    ctx.nestedLevel++;
  }

  private handleFootnoteEndnoteStart(ctx: ParseContext, data: Uint8Array): void {
    const isFootnote = ctx.currentCtrlId === CTRL_ID.FOOTNOTE;
    
    let number = 0;
    let paragraphWidth = 0;
    if (data.length >= 8) {
      number = readUint16(data, 0);
      paragraphWidth = data.length >= 12 ? hwpunitToPt(readUint32(data, 8)) : 0;
    }
    
    ctx.currentFootnoteEndnote = {
      type: isFootnote ? 'footnote' : 'endnote',
      number,
      paragraphWidth,
      paragraphs: [],
    };
    ctx.inFootnoteEndnote = true;
    ctx.footnoteEndnoteParagraphs = [];
    ctx.nestedLevel++;
  }

  private finalizeNestedContext(ctx: ParseContext, section: HwpxSection, levelDrop: number): void {
    for (let i = 0; i < levelDrop && ctx.nestedLevel > 0; i++) {
      if (ctx.inHeaderFooter && ctx.currentHeaderFooter) {
        if (ctx.currentParagraph) {
          this.flushPendingTextSegments(ctx);
          if (ctx.currentParagraph.runs.length > 0) {
            ctx.headerFooterParagraphs.push(ctx.currentParagraph);
            ctx.currentParagraph = null;
          }
        }
        ctx.currentHeaderFooter.paragraphs = [...ctx.headerFooterParagraphs];
        
        const headerFooterData: any = {
          id: generateId(),
          type: ctx.currentHeaderFooter.type,
          applyTo: ctx.currentHeaderFooter.applyTo,
          textWidth: ctx.currentHeaderFooter.textWidth,
          textHeight: ctx.currentHeaderFooter.textHeight,
          paragraphs: ctx.currentHeaderFooter.paragraphs,
        };
        section.elements.push({ type: ctx.currentHeaderFooter.type, data: headerFooterData });
        
        ctx.inHeaderFooter = false;
        ctx.currentHeaderFooter = null;
        ctx.headerFooterParagraphs = [];
        ctx.nestedLevel--;
      } else if (ctx.inFootnoteEndnote && ctx.currentFootnoteEndnote) {
        if (ctx.currentParagraph) {
          this.flushPendingTextSegments(ctx);
          if (ctx.currentParagraph.runs.length > 0) {
            ctx.footnoteEndnoteParagraphs.push(ctx.currentParagraph);
            ctx.currentParagraph = null;
          }
        }
        ctx.currentFootnoteEndnote.paragraphs = [...ctx.footnoteEndnoteParagraphs];
        
        const footnoteData: any = {
          id: generateId(),
          type: ctx.currentFootnoteEndnote.type,
          number: ctx.currentFootnoteEndnote.number,
          paragraphWidth: ctx.currentFootnoteEndnote.paragraphWidth,
          paragraphs: ctx.currentFootnoteEndnote.paragraphs,
        };
        section.elements.push({ type: ctx.currentFootnoteEndnote.type, data: footnoteData });
        
        ctx.inFootnoteEndnote = false;
        ctx.currentFootnoteEndnote = null;
        ctx.footnoteEndnoteParagraphs = [];
        ctx.nestedLevel--;
      } else if (ctx.inMemo && ctx.currentMemo) {
        this.finishMemo(ctx);
        ctx.nestedLevel--;
      } else if (ctx.inTableCell && ctx.currentTable) {
        if (ctx.currentParagraph) {
          this.flushPendingTextSegments(ctx);
          if (ctx.currentParagraph.runs.length === 0) {
            ctx.currentParagraph.runs.push({ text: '' });
          }
          ctx.cellParagraphs.push(ctx.currentParagraph);
          ctx.currentParagraph = null;
        }
        
        const row = ctx.currentTableRow;
        const col = ctx.currentTableCol;
        if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]) {
          ctx.tableCells[row][col].paragraphs = [...ctx.cellParagraphs];
        }
        
        this.finishTable(ctx, section);
        ctx.nestedLevel--;
      } else {
        ctx.nestedLevel = Math.max(0, ctx.nestedLevel - 1);
      }
    }
  }

  private handleTable(ctx: ParseContext, data: Uint8Array): void {
    if (!ctx.currentTable) {
      return;
    }
    if (data.length < 22) {
      return;
    }
    
    const tableProps = readUint32(data, 0);
    const rowCount = readUint16(data, 4);
    const colCount = readUint16(data, 6);
    const cellSpacing = hwpunitToPt(readUint16(data, 8));
    
    const inMarginLeft = hwpunitToPt(readUint16(data, 10));
    const inMarginRight = hwpunitToPt(readUint16(data, 12));
    const inMarginTop = hwpunitToPt(readUint16(data, 14));
    const inMarginBottom = hwpunitToPt(readUint16(data, 16));
    
    ctx.currentTable.rowCount = rowCount;
    ctx.currentTable.colCount = colCount;
    ctx.currentTable.inMargin = { left: inMarginLeft, right: inMarginRight, top: inMarginTop, bottom: inMarginBottom };
    ctx.currentTable.cellSpacing = cellSpacing;
    
    ctx.tableRowCount = rowCount;
    ctx.tableColCount = colCount;
    
    let offset = 18;
    
    const rowHeights: number[] = [];
    for (let r = 0; r < rowCount && offset + 2 <= data.length; r++) {
      rowHeights.push(hwpunitToPt(readUint16(data, offset)));
      offset += 2;
    }
    (ctx.currentTable as any)._rowHeights = rowHeights;
    
    if (offset + 2 <= data.length) {
      ctx.currentTable.borderFillId = readUint16(data, offset);
      offset += 2;
    }
    
    if (offset + 2 <= data.length) {
      offset += 2;
    }
    
    const cellCount = rowCount * colCount;
    const cellListOffset = offset;
    
    ctx.tableCells = [];
    for (let r = 0; r < rowCount; r++) {
      ctx.tableCells[r] = [];
      for (let c = 0; c < colCount; c++) {
        ctx.tableCells[r][c] = {
          paragraphs: [],
          colAddr: c,
          rowAddr: r,
          colSpan: 1,
          rowSpan: 1,
        };
      }
    }
    
    for (let i = 0; i < cellCount && cellListOffset + (i + 1) * 26 <= data.length; i++) {
      const cellOffset = cellListOffset + i * 26;
      
      const colAddr = readUint16(data, cellOffset);
      const rowAddr = readUint16(data, cellOffset + 2);
      const colSpan = readUint16(data, cellOffset + 4);
      const rowSpan = readUint16(data, cellOffset + 6);
      const cellWidth = hwpunitToPt(readUint32(data, cellOffset + 8));
      const cellHeight = hwpunitToPt(readUint32(data, cellOffset + 12));
      const marginLeft = hwpunitToPt(readUint16(data, cellOffset + 16));
      const marginRight = hwpunitToPt(readUint16(data, cellOffset + 18));
      const marginTop = hwpunitToPt(readUint16(data, cellOffset + 20));
      const marginBottom = hwpunitToPt(readUint16(data, cellOffset + 22));
      const borderFillId = readUint16(data, cellOffset + 24);
      
      if (rowAddr < rowCount && colAddr < colCount && ctx.tableCells[rowAddr]) {
        const borderFill = ctx.borderFills.get(borderFillId);
        
        let backgroundColor: string | undefined;
        let borderTop: { width: number; style: string; color: string } | undefined;
        let borderBottom: { width: number; style: string; color: string } | undefined;
        let borderLeft: { width: number; style: string; color: string } | undefined;
        let borderRight: { width: number; style: string; color: string } | undefined;
        
        if (borderFill) {
          const fill = borderFill.fill as { backgroundColor?: number } | undefined;
          if (fill?.backgroundColor !== undefined) {
            backgroundColor = colorrefToHex(fill.backgroundColor);
          }
          if (borderFill.borders) {
            const mapBorder = (b: { type: number; width: number; color: number }) => ({
              width: b.width * 0.1,
              style: b.type === 0 ? 'none' : 'solid',
              color: colorrefToHex(b.color)
            });
            if (borderFill.borders.top) borderTop = mapBorder(borderFill.borders.top);
            if (borderFill.borders.bottom) borderBottom = mapBorder(borderFill.borders.bottom);
            if (borderFill.borders.left) borderLeft = mapBorder(borderFill.borders.left);
            if (borderFill.borders.right) borderRight = mapBorder(borderFill.borders.right);
          }
        }
        
        ctx.tableCells[rowAddr][colAddr] = {
          paragraphs: [],
          colAddr,
          rowAddr,
          colSpan: colSpan || 1,
          rowSpan: rowSpan || 1,
          width: cellWidth,
          height: cellHeight,
          marginLeft,
          marginRight,
          marginTop,
          marginBottom,
          borderFillId,
          backgroundColor,
          borderTop,
          borderBottom,
          borderLeft,
          borderRight,
        };
      }
    }
    
    ctx.currentTableRow = 0;
    ctx.currentTableCol = 0;
  }

  private handleListHeader(ctx: ParseContext, section: HwpxSection): void {
    if (ctx.inHeaderFooter && ctx.currentHeaderFooter) {
      if (ctx.currentParagraph) {
        this.flushPendingTextSegments(ctx);
        if (ctx.currentParagraph.runs.length > 0) {
          ctx.headerFooterParagraphs.push(ctx.currentParagraph);
        }
      }
      ctx.currentParagraph = null;
      return;
    }
    
    if (ctx.inFootnoteEndnote && ctx.currentFootnoteEndnote) {
      if (ctx.currentParagraph) {
        this.flushPendingTextSegments(ctx);
        if (ctx.currentParagraph.runs.length > 0) {
          ctx.footnoteEndnoteParagraphs.push(ctx.currentParagraph);
        }
      }
      ctx.currentParagraph = null;
      return;
    }
    
    if (ctx.inMemo && ctx.currentMemo) {
      if (ctx.currentParagraph) {
        this.flushPendingTextSegments(ctx);
        if (ctx.currentParagraph.runs.length > 0) {
          ctx.memoParagraphs.push(ctx.currentParagraph);
        }
      }
      ctx.currentParagraph = null;
      return;
    }
    
    if (!ctx.currentTable) return;
    
    if (ctx.inTableCell) {
      if (ctx.currentParagraph) {
        this.flushPendingTextSegments(ctx);
        if (ctx.currentParagraph.runs.length === 0) {
          ctx.currentParagraph.runs.push({ text: '' });
        }
        ctx.cellParagraphs.push(ctx.currentParagraph);
      }
      
      const row = ctx.currentTableRow;
      const col = ctx.currentTableCol;
      
      
      if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]) {
        ctx.tableCells[row][col].paragraphs = [...ctx.cellParagraphs];
      }
      
      ctx.currentTableCol++;
      if (ctx.currentTableCol >= ctx.tableColCount) {
        ctx.currentTableCol = 0;
        ctx.currentTableRow++;
      }
      
      if (ctx.currentTableRow >= ctx.tableRowCount) {
        this.finishTable(ctx, section);
      }
    }
    
    ctx.inTableCell = true;
    ctx.cellParagraphs = [];
    ctx.currentParagraph = null;
  }

  private finishTable(ctx: ParseContext, section: HwpxSection): void {
    if (!ctx.currentTable) return;
    
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
    
    const rowHeights = (ctx.currentTable as any)._rowHeights as number[] | undefined;
    const rows: TableRow[] = [];
    for (let r = 0; r < ctx.tableRowCount; r++) {
      const cells: TableCell[] = [];
      for (let c = 0; c < ctx.tableColCount; c++) {
        if (coveredCells.has(`${r},${c}`)) continue;
        const cell = ctx.tableCells[r]?.[c];
        if (cell) {
          if (cell.paragraphs.length === 0) {
            cell.paragraphs.push({ id: generateId(), runs: [{ text: '' }] });
          }
          cells.push(cell);
        }
      }
      if (cells.length > 0) {
        const row: TableRow = { cells };
        if (rowHeights && rowHeights[r]) {
          row.height = rowHeights[r];
        }
        rows.push(row);
      }
    }
    
    ctx.currentTable.rows = rows;
    delete (ctx.currentTable as any)._rowHeights;
    if (!ctx.inHeaderFooter && !ctx.inFootnoteEndnote) {
      section.elements.push({ type: 'table', data: ctx.currentTable });
    }
    
    ctx.currentTable = null;
    ctx.inTableCell = false;
    ctx.cellParagraphs = [];
  }

  private handleShapeComponent(ctx: ParseContext, data: Uint8Array): void {
    if (!ctx.pendingImage || data.length < 24) return;
    
    const width = readInt32(data, 16);
    const height = readInt32(data, 20);
    
    if (width > 0) ctx.pendingImage.width = width / 7200 * 72;
    if (height > 0) ctx.pendingImage.height = height / 7200 * 72;
  }

  private handlePicture(ctx: ParseContext, data: Uint8Array, section: HwpxSection): void {
    if (data.length < 73) return;
    
    const binItemId = readUint16(data, 71);
    const idStr = `BIN${String(binItemId).padStart(4, '0')}`;
    
    const existingImage = this._content.images.get(idStr);
    
    const image: HwpxImage = {
      id: generateId(),
      binaryId: idStr,
      width: ctx.pendingImage?.width || 200,
      height: ctx.pendingImage?.height || 150,
      data: existingImage?.data,
      mimeType: existingImage?.mimeType,
    };
    
    if (ctx.inTableCell && ctx.currentTable) {
      const row = ctx.currentTableRow;
      const col = ctx.currentTableCol;
      if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]?.[col]) {
        if (!ctx.tableCells[row][col].elements) ctx.tableCells[row][col].elements = [];
        ctx.tableCells[row][col].elements!.push({ type: 'image', data: image });
      }
    } else if (!ctx.inHeaderFooter && !ctx.inFootnoteEndnote) {
      section.elements.push({ type: 'image', data: image });
    }
    ctx.pendingImage = null;
  }

  private handleParaLineSeg(ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 36) return;
    
    ctx.pendingLineSegs = [];
    for (let i = 0; i + 36 <= data.length; i += 36) {
      const textStartPos = readUint32(data, i);
      const verticalPos = readInt32(data, i + 4);
      const lineHeight = readInt32(data, i + 8);
      const textHeight = readInt32(data, i + 12);
      const baselineDistance = readInt32(data, i + 16);
      const lineSpacing = readInt32(data, i + 20);
      const horizontalStart = readInt32(data, i + 24);
      const segmentWidth = readInt32(data, i + 28);
      const flagsVal = readUint32(data, i + 32);
      
      ctx.pendingLineSegs.push({
        textStartPos,
        verticalPos: hwpunitToPt(verticalPos),
        lineHeight: hwpunitToPt(lineHeight),
        textHeight: hwpunitToPt(textHeight),
        baselineDistance: hwpunitToPt(baselineDistance),
        lineSpacing: hwpunitToPt(lineSpacing),
        horizontalStart: hwpunitToPt(horizontalStart),
        segmentWidth: hwpunitToPt(segmentWidth),
        flags: {
          isPageFirst: (flagsVal & 0x01) !== 0,
          isColumnFirst: (flagsVal & 0x02) !== 0,
          isEmpty: (flagsVal & 0x04) !== 0,
          isLastInPara: (flagsVal & 0x08) !== 0,
          isAutoHyphen: (flagsVal & 0x10) !== 0,
          isIndent: (flagsVal & 0x20) !== 0,
        },
      });
    }
    
    if (ctx.currentParagraph && ctx.pendingLineSegs.length > 0) {
      (ctx.currentParagraph as any).lineSegs = ctx.pendingLineSegs;
    }
  }

  private handleParaRangeTag(ctx: ParseContext, data: Uint8Array): void {
    if (data.length < 12) return;
    
    ctx.pendingRangeTags = [];
    for (let i = 0; i + 12 <= data.length; i += 12) {
      const start = readUint32(data, i);
      const end = readUint32(data, i + 4);
      const tagVal = readUint32(data, i + 8);
      const type = (tagVal >> 24) & 0xFF;
      const tagData = tagVal & 0x00FFFFFF;
      
      ctx.pendingRangeTags.push({
        start,
        end,
        type,
        data: tagData,
      });
    }
    
    if (ctx.currentParagraph && ctx.pendingRangeTags.length > 0) {
      (ctx.currentParagraph as any).rangeTags = ctx.pendingRangeTags;
    }
  }

  private parsePageBorderFill(data: Uint8Array): ParsedPageBorderFill | null {
    if (data.length < 14) return null;
    
    const props = readUint32(data, 0);
    const applyToVal = props & 0x03;
    const includeHeader = (props & 0x04) !== 0;
    const includeFooter = (props & 0x08) !== 0;
    const fillAreaVal = (props >> 4) & 0x03;
    
    const applyToMap: Record<number, 'both' | 'even' | 'odd' | 'firstOnly'> = {
      0: 'both', 1: 'even', 2: 'odd', 3: 'firstOnly'
    };
    const fillAreaMap: Record<number, 'paper' | 'page' | 'border'> = {
      0: 'paper', 1: 'page', 2: 'border'
    };
    
    return {
      applyTo: applyToMap[applyToVal] || 'both',
      includeHeader,
      includeFooter,
      fillArea: fillAreaMap[fillAreaVal] || 'paper',
      offsets: {
        left: hwpunitToPt(readInt16(data, 4)),
        right: hwpunitToPt(readInt16(data, 6)),
        top: hwpunitToPt(readInt16(data, 8)),
        bottom: hwpunitToPt(readInt16(data, 10)),
      },
      borderFillId: readUint16(data, 12),
    };
  }

  private getDefaultPageSettings(): PageSettings {
    return {
      width: 595,
      height: 842,
      marginLeft: 85,
      marginRight: 85,
      marginTop: 113,
      marginBottom: 85,
      orientation: 'portrait',
    };
  }

  getContent(): HwpxContent {
    return this._content;
  }

  getSerializableContent(): object {
    return {
      metadata: this._content.metadata,
      sections: this._content.sections,
      images: Object.fromEntries(this._content.images),
      binData: Object.fromEntries(this._content.binData),
      footnotes: this._content.footnotes,
      endnotes: this._content.endnotes,
      isReadOnly: this._isReadOnly,
    };
  }

  dispose(): void {
    this._onDidDispose.fire();
    this._onDidChangeContent.dispose();
    this._onDidDispose.dispose();
  }
}

function parseHwpContent(data: Uint8Array): HwpxContent {
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
  
  let docInfoData = getEntryData('/DocInfo');
  if (docInfoData) {
    let offset = 0;
    let binDataId = 1;
    let faceNameId = 0;
    let charShapeId = 0;
    let paraShapeId = 0;
    
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
    content.sections.push(parseSectionData(sectionData, content.images, faceNames, charShapes, paraShapes));
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
  };
  
  let offset = 0;
  let prevLevel = 0;
  let shapeTextLevel = -1;
  let headerFooterLevel = -1;
  
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
  };
  const tableStack: TableStackItem[] = [];
  let currentTableLevel = 0;

  const flushPending = () => {
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

  const finishCurrentTable = () => {
    if (!ctx.currentTable) return;
    
    if (ctx.inTableCell && ctx.currentParagraph) {
      flushPending();
      if (ctx.currentParagraph.runs.length === 0) {
        ctx.currentParagraph.runs.push({ text: '' });
      }
      ctx.cellParagraphs.push(ctx.currentParagraph);
      ctx.currentParagraph = null;
    }
    
    const row = ctx.currentTableRow;
    const col = ctx.currentTableCol;
    if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]) {
                    ctx.tableCells[row][col].paragraphs = [...ctx.cellParagraphs];
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
    } else if (!ctx.inHeaderFooter) {
      section.elements.push({ type: 'table', data: ctx.currentTable });
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
     
      const recordData = data.slice(nextOffset, nextOffset + size);
     
        // Check if we exited header/footer scope based on level
        if (ctx.inHeaderFooter && headerFooterLevel >= 0 && level <= headerFooterLevel) {
          ctx.inHeaderFooter = false;
          ctx.inFootnoteEndnote = false;
          headerFooterLevel = -1;
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
      if (ctx.inShapeText && shapeTextLevel >= 0 && level <= shapeTextLevel && !_isParaSubTag) {
        if (ctx.currentParagraph) {
          flushPending();
          if (ctx.currentParagraph.runs.length === 0) {
            ctx.currentParagraph.runs.push({ text: '' });
          }
          ctx.shapeTextParagraphs.push(ctx.currentParagraph);
          ctx.currentParagraph = null;
        }
        ctx.inShapeText = false;
        shapeTextLevel = -1;
      }
      
      switch (tagId) {
      case HWP_TAGS.HWPTAG_PARA_HEADER:
        flushPending();
        if (ctx.currentParagraph) {
          if (ctx.currentParagraph.runs.length === 0) {
            const emptyCharShape = ctx.charShapePositions.length > 0
              ? ctx.charShapes.get(ctx.charShapePositions[0].charShapeId)
              : ctx.charShapes.get(ctx.currentCharShapeId);
            const emptyCharStyle = emptyCharShape ? charShapeToStyleStandalone(emptyCharShape, ctx.faceNames) : undefined;
            ctx.currentParagraph.runs.push({ text: '', charStyle: emptyCharStyle });
          }
           if (ctx.inShapeText) {
             ctx.shapeTextParagraphs.push(ctx.currentParagraph);
           } else if (ctx.inTableCell) {
             ctx.cellParagraphs.push(ctx.currentParagraph);
           } else if (level === 0) {
             section.elements.push({ type: 'paragraph', data: ctx.currentParagraph });
           }
         }
         ctx.currentParagraph = { id: generateId(), runs: [] };
         ctx.pendingTextSegments = [];
        if (recordData.length >= 12) {
          const paraShapeId = readUint16(recordData, 8);
          const paraShape = ctx.paraShapes.get(paraShapeId);
          if (paraShape) {
            const alignMap: Record<number, 'Justify' | 'Left' | 'Right' | 'Center' | 'Distribute'> = {
              0: 'Justify', 1: 'Left', 2: 'Right', 3: 'Center', 4: 'Distribute', 5: 'Distribute'
            };
            ctx.currentParagraph.paraStyle = {
              align: alignMap[paraShape.alignment] || 'Justify',
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
          if (ctx.currentParagraph) {
            let currentStart = 0;
            let currentText = '';
            let charIndex = 0;
            let i = 0;
            
            while (i < recordData.length - 1) {
              const charCode = readUint16(recordData, i);
              i += 2;
              
              if (charCode === 0) {
                charIndex++;
                continue;
              }
              
              if (charCode < 32) {
                if (currentText) {
                  ctx.pendingTextSegments.push({ start: currentStart, end: charIndex, text: currentText });
                  currentText = '';
                }
                
                 if (charCode === CTRL_CHAR.LINE_BREAK) {
                   ctx.pendingTextSegments.push({ start: charIndex, end: charIndex + 1, text: '\n' });
                   charIndex++;
                 } else if (charCode === 0x0009) {
                   ctx.pendingTextSegments.push({ start: charIndex, end: charIndex + 8, text: '\t' });
                   i += 14;
                   charIndex += 8;
                 } else if (charCode >= 0x0002 && charCode <= 0x0008) {
                   i += 14;
                   charIndex += 8;
                 } else if (charCode === 0x000B || charCode === 0x000C ||
                            charCode === 0x000E || charCode === 0x000F || charCode === 0x0010 ||
                            charCode === 0x0011 || charCode === 0x0012 || charCode === 0x0013 ||
                            charCode === 0x0015 || charCode === 0x0016 || charCode === 0x0017) {
                   i += 14;
                   charIndex += 8;
                 } else if (charCode === CTRL_CHAR.PARAGRAPH_BREAK) {
                   break;
                 } else {
                   charIndex++;
                 }
                currentStart = charIndex;
                continue;
              }
              
              currentText += String.fromCharCode(charCode);
              charIndex++;
            }
            
            if (currentText) {
              ctx.pendingTextSegments.push({ start: currentStart, end: charIndex, text: currentText });
            }
          }
          break;
        
       case HWP_TAGS.HWPTAG_CTRL_HEADER:
         if (recordData.length >= 4) {
           ctx.currentCtrlId = readUint32(recordData, 0);
            if (ctx.currentCtrlId === CTRL_ID.TABLE) {
              if (ctx.currentTable) {
                 if (ctx.inTableCell && ctx.currentParagraph) {
                    flushPending();
                    if (ctx.currentParagraph.runs.length === 0) {
                      ctx.currentParagraph.runs.push({ text: '' });
                    }
                    ctx.cellParagraphs.push(ctx.currentParagraph);
                    ctx.currentParagraph = null;
                  }
                 if (ctx.inTableCell) {
                  const row = ctx.currentTableRow;
                  const col = ctx.currentTableCol;
                  if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]) {
      ctx.tableCells[row][col].paragraphs = [...ctx.cellParagraphs];
                  }
                }
                
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
                });
              } else {
                if (ctx.currentParagraph?.runs.length && level === 0) {
                  section.elements.push({ type: 'paragraph', data: ctx.currentParagraph });
                  ctx.currentParagraph = null;
                }
              }
              ctx.currentTable = { id: generateId(), rows: [], rowCount: 0, colCount: 0 };
             ctx.tableCells = [];
             ctx.inTableCell = false;
             ctx.cellParagraphs = [];
             currentTableLevel = level;
            } else if (ctx.currentCtrlId === CTRL_ID.PICTURE || ctx.currentCtrlId === CTRL_ID.GSO) {
            ctx.pendingImage = { width: 200, height: 150 };
            if (recordData.length >= 46) {
              const w = readUint32(recordData, 16);
              const h = readUint32(recordData, 20);
              if (w > 0) ctx.pendingImage.width = w / 7200 * 72;
              if (h > 0) ctx.pendingImage.height = h / 7200 * 72;
            }
            ctx.inShapeText = true;
            ctx.shapeTextParagraphs = [];
            shapeTextLevel = level;
           } else if (ctx.currentCtrlId === CTRL_ID.HEADER
                     || ctx.currentCtrlId === CTRL_ID.FOOTER
                     || ctx.currentCtrlId === CTRL_ID.FOOTNOTE
                     || ctx.currentCtrlId === CTRL_ID.ENDNOTE) {
            ctx.inHeaderFooter = true;
            ctx.inFootnoteEndnote = ctx.currentCtrlId === CTRL_ID.FOOTNOTE || ctx.currentCtrlId === CTRL_ID.ENDNOTE;
            headerFooterLevel = level;
          } else if (ctx.currentCtrlId !== CTRL_ID.SECTION_DEF
                     && ctx.currentCtrlId !== CTRL_ID.COLUMN_DEF
                     && ctx.currentCtrlId !== CTRL_ID.TABLE
                     && ctx.currentCtrlId !== CTRL_ID.AUTO_NUMBER
                     && ctx.currentCtrlId !== CTRL_ID.PAGE_NUMBER_POS
                     && ctx.currentCtrlId !== 0x65716564 /* EQUATION */
                     && ctx.currentCtrlId !== CTRL_ID.BOOKMARK
                     && ctx.currentCtrlId !== 0x74637073 /* TCPS */
                     && ctx.currentCtrlId !== CTRL_ID.FORM) {
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
               flushPending();
               if (ctx.currentParagraph.runs.length === 0) {
                 ctx.currentParagraph.runs.push({ text: '' });
               }
               ctx.cellParagraphs.push(ctx.currentParagraph);
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
          
          ctx.currentTableCol = cellCol;
          ctx.currentTableRow = cellRow;
          
          if (cellRow < ctx.tableRowCount && cellCol < ctx.tableColCount && ctx.tableCells[cellRow]) {
            ctx.tableCells[cellRow][cellCol].colSpan = colSpan || 1;
            ctx.tableCells[cellRow][cellCol].rowSpan = rowSpan || 1;
            ctx.tableCells[cellRow][cellCol].width = cellWidth;
            ctx.tableCells[cellRow][cellCol].height = cellHeight;
            ctx.tableCells[cellRow][cellCol].marginLeft = marginLeft;
            ctx.tableCells[cellRow][cellCol].marginRight = marginRight;
            ctx.tableCells[cellRow][cellCol].marginTop = marginTop;
            ctx.tableCells[cellRow][cellCol].marginBottom = marginBottom;
            ctx.tableCells[cellRow][cellCol].borderFillId = borderFillId;
          }
          
           ctx.inTableCell = true;
           ctx.cellParagraphs = [];
           ctx.currentParagraph = null;
         }
         break;
        
      case HWP_TAGS.HWPTAG_SHAPE_COMPONENT:
        if (ctx.pendingImage && recordData.length >= 24) {
          const w = readInt32(recordData, 16);
          const h = readInt32(recordData, 20);
          if (w > 0) ctx.pendingImage.width = w / 7200 * 72;
          if (h > 0) ctx.pendingImage.height = h / 7200 * 72;
        }
        break;
        
      case HWP_TAGS.HWPTAG_SHAPE_COMPONENT_PICTURE:
        if (recordData.length >= 73) {
          const binItemId = readUint16(recordData, 71);
          const idStr = `BIN${String(binItemId).padStart(4, '0')}`;
          const existingImage = images.get(idStr);
          const image: HwpxImage = {
            id: generateId(),
            binaryId: idStr,
            width: ctx.pendingImage?.width || 200,
            height: ctx.pendingImage?.height || 150,
            data: existingImage?.data,
            mimeType: existingImage?.mimeType,
          };
          if (ctx.inTableCell) {
            const row = ctx.currentTableRow;
            const col = ctx.currentTableCol;
            if (row < ctx.tableRowCount && col < ctx.tableColCount && ctx.tableCells[row]?.[col]) {
              if (!ctx.tableCells[row][col].elements) ctx.tableCells[row][col].elements = [];
              ctx.tableCells[row][col].elements!.push({ type: 'image', data: image });
            }
          } else if (!ctx.inHeaderFooter) {
            section.elements.push({ type: 'image', data: image });
          }
          ctx.pendingImage = null;
        }
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
  
  if (ctx.currentParagraph && !ctx.inTableCell) {
    flushPending();
    if (ctx.currentParagraph.runs.length === 0) {
      const emptyCharShape = ctx.charShapePositions.length > 0
        ? ctx.charShapes.get(ctx.charShapePositions[0].charShapeId)
        : ctx.charShapes.get(ctx.currentCharShapeId);
      const emptyCharStyle = emptyCharShape ? charShapeToStyleStandalone(emptyCharShape, ctx.faceNames) : undefined;
      ctx.currentParagraph.runs.push({ text: '', charStyle: emptyCharStyle });
    }
    section.elements.push({ type: 'paragraph', data: ctx.currentParagraph });
  }
  
  if (section.elements.length === 0) {
    section.elements.push({ type: 'paragraph', data: { id: generateId(), runs: [{ text: '' }] } });
  }
  
  return section;
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

function charShapeToStyleStandalone(charShape: ParsedCharShape, faceNames: Map<number, ParsedFaceName>): import('../hwpx/types').CharacterStyle {
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
