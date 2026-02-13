import { parse } from 'hwp.js';
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
  FontRef,
} from '../hwpx/types';

const CommonCtrlID = {
  Table: 0x74626C20,
};

const CONTROL_CHARS = new Set([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31]);

function generateId(): string {
  return Math.random().toString(36).substring(2, 11);
}

export class HwpParser {
  static parse(data: Uint8Array): HwpxContent {
    let hwpDocument: any;
    try {
      hwpDocument = parse(data, { type: 'array' });
    } catch (e: any) {
      console.warn('HWP parsing error:', e?.message || e);
      throw new Error(`HWP 파일을 파싱할 수 없습니다: ${e?.message || 'Unknown error'}`);
    }
    
    const content: HwpxContent = {
      metadata: this.extractMetadata(hwpDocument),
      sections: [],
      images: new Map(),
      binItems: new Map(),
      binData: new Map(),
      footnotes: [],
      endnotes: [],
    };

    for (const hwpSection of hwpDocument.sections) {
      const section = this.convertSection(hwpSection, hwpDocument.info);
      content.sections.push(section);
    }

    return content;
  }

  private static extractMetadata(hwpDoc: any): HwpxContent['metadata'] {
    return {
      title: hwpDoc.info?.title || undefined,
      creator: hwpDoc.info?.author || undefined,
    };
  }

  private static convertSection(hwpSection: any, docInfo: any): HwpxSection {
    const section: HwpxSection = {
      elements: [],
      pageSettings: this.convertPageSettings(hwpSection),
    };

    for (const paragraph of hwpSection.content) {
      const tableElements = this.extractTablesFromParagraph(paragraph, docInfo);
      
      if (tableElements.length > 0) {
        for (const element of tableElements) {
          section.elements.push(element);
        }
      } else {
        const converted = this.convertParagraph(paragraph, docInfo);
        section.elements.push({ type: 'paragraph', data: converted });
      }
    }

    return section;
  }

  private static extractTablesFromParagraph(paragraph: any, docInfo: any): SectionElement[] {
    const elements: SectionElement[] = [];
    
    if (!paragraph.controls || paragraph.controls.length === 0) {
      return elements;
    }

    let hasTable = false;
    for (const control of paragraph.controls) {
      if (control.id === CommonCtrlID.Table) {
        hasTable = true;
        const table = this.convertTableControl(control, docInfo);
        if (table) {
          elements.push({ type: 'table', data: table });
        }
      }
    }

    if (hasTable) {
      const textContent = this.extractTextFromParagraph(paragraph);
      if (textContent.trim()) {
        const textPara = this.convertParagraph(paragraph, docInfo);
        elements.unshift({ type: 'paragraph', data: textPara });
      }
    }

    return elements;
  }

  private static extractTextFromParagraph(paragraph: any): string {
    let text = '';
    for (const char of paragraph.content || []) {
      if (char.type === 0 && char.value) {
        text += this.filterControlChars(char.value);
      }
    }
    return text;
  }

  private static filterControlChars(text: string): string {
    let result = '';
    for (let i = 0; i < text.length; i++) {
      const code = text.charCodeAt(i);
      if (!CONTROL_CHARS.has(code)) {
        result += text[i];
      } else if (code === 10 || code === 13) {
        result += '\n';
      }
    }
    return result.replace(/\n+/g, '\n').replace(/^\n|\n$/g, '');
  }

  private static convertTableControl(control: any, docInfo: any): HwpxTable | null {
    if (!control.content || !Array.isArray(control.content)) {
      return null;
    }

    const rowCount = control.rowCount || control.content.length;
    const colCount = control.columnCount || (control.content[0]?.length || 0);
    
    if (rowCount === 0 || colCount === 0) {
      return null;
    }

    const table: HwpxTable = {
      id: generateId(),
      rowCount,
      colCount,
      rows: [],
      width: control.width ? control.width / 100 : undefined,
      height: control.height ? control.height / 100 : undefined,
    };

    for (let rowIdx = 0; rowIdx < control.content.length; rowIdx++) {
      const rowData = control.content[rowIdx];
      if (!Array.isArray(rowData)) continue;

      const tableRow: TableRow = {
        cells: [],
      };

      for (const cellData of rowData) {
        const cell = this.convertCell(cellData, docInfo);
        tableRow.cells.push(cell);
      }

      if (tableRow.cells.length > 0) {
        table.rows.push(tableRow);
      }
    }

    return table.rows.length > 0 ? table : null;
  }

  private static convertCell(cellData: any, docInfo: any): TableCell {
    const attr = cellData.attribute || {};
    const items = cellData.items || [];

    const cell: TableCell = {
      paragraphs: [],
      colAddr: attr.column ?? 0,
      rowAddr: attr.row ?? 0,
      colSpan: attr.colSpan ?? 1,
      rowSpan: attr.rowSpan ?? 1,
      width: attr.width ? attr.width / 100 : undefined,
      height: attr.height ? attr.height / 100 : undefined,
    };

    if (attr.padding && Array.isArray(attr.padding)) {
      cell.marginLeft = attr.padding[0] ? attr.padding[0] / 100 : undefined;
      cell.marginRight = attr.padding[1] ? attr.padding[1] / 100 : undefined;
      cell.marginTop = attr.padding[2] ? attr.padding[2] / 100 : undefined;
      cell.marginBottom = attr.padding[3] ? attr.padding[3] / 100 : undefined;
    }

    for (const item of items) {
      const para = this.convertParagraph(item, docInfo);
      cell.paragraphs.push(para);
    }

    if (cell.paragraphs.length === 0) {
      cell.paragraphs.push({ id: generateId(), runs: [{ text: '' }] });
    }

    return cell;
  }

  private static convertPageSettings(hwpSection: any): PageSettings {
    return {
      width: hwpSection.width / 100 || 595,
      height: hwpSection.height / 100 || 842,
      marginTop: hwpSection.paddingTop / 100 || 56.7,
      marginBottom: hwpSection.paddingBottom / 100 || 56.7,
      marginLeft: hwpSection.paddingLeft / 100 || 56.7,
      marginRight: hwpSection.paddingRight / 100 || 56.7,
      orientation: hwpSection.orientation === 1 ? 'landscape' : 'portrait',
    };
  }

  private static convertParagraph(hwpParagraph: any, docInfo: any): HwpxParagraph {
    const paragraph: HwpxParagraph = {
      id: generateId(),
      runs: [],
    };

    const shapeIndex = hwpParagraph.shapeIndex;
    if (docInfo?.paragraphShapes?.[shapeIndex]) {
      const paraShape = docInfo.paragraphShapes[shapeIndex];
      paragraph.paraStyle = this.convertParaStyle(paraShape);
    }

    let currentText = '';
    let currentCharShapeIndex = -1;
    let shapePointerIndex = 0;

    for (let i = 0; i < hwpParagraph.content.length; i++) {
      const char = hwpParagraph.content[i];
      
      let charShapeIndex = currentCharShapeIndex;
      if (hwpParagraph.shapeBuffer && hwpParagraph.shapeBuffer.length > 0) {
        const nextShape = hwpParagraph.shapeBuffer[shapePointerIndex + 1];
        if (nextShape && i >= nextShape.pos) {
          shapePointerIndex++;
        }
        charShapeIndex = hwpParagraph.shapeBuffer[shapePointerIndex]?.shapeIndex ?? 0;
      }

      if (char.type === 0) {
        const filteredValue = this.filterControlChars(char.value);
        if (filteredValue) {
          if (charShapeIndex !== currentCharShapeIndex && currentText) {
            paragraph.runs.push(this.createRun(currentText, currentCharShapeIndex, docInfo));
            currentText = '';
          }
          currentCharShapeIndex = charShapeIndex;
          currentText += filteredValue;
        }
      }
    }

    if (currentText) {
      paragraph.runs.push(this.createRun(currentText, currentCharShapeIndex, docInfo));
    }

    if (paragraph.runs.length === 0) {
      paragraph.runs.push({ text: '' });
    }

    return paragraph;
  }

  private static convertParaStyle(paraShape: any): ParagraphStyle {
    const style: ParagraphStyle = {
      align: this.convertAlign(paraShape.align ?? 0),
      lineSpacing: paraShape.lineSpacing ?? 160,
      marginTop: (paraShape.marginTop ?? 0) / 100,
      marginBottom: (paraShape.marginBottom ?? 0) / 100,
      marginLeft: (paraShape.marginLeft ?? paraShape.indent ?? 0) / 100,
      marginRight: (paraShape.marginRight ?? 0) / 100,
      firstLineIndent: (paraShape.firstLineIndent ?? paraShape.firstIndent ?? 0) / 100,
      keepWithNext: !!paraShape.keepWithNext,
      keepLines: !!paraShape.keepLines,
    };

    if (paraShape.lineSpacingType !== undefined) {
      const typeMap: Record<number, ParagraphStyle['lineSpacingType']> = {
        0: 'percent',
        1: 'fixed',
        2: 'betweenLines',
        3: 'atLeast',
      };
      style.lineSpacingType = typeMap[paraShape.lineSpacingType] || 'percent';
    }

    if (paraShape.pageBreakBefore !== undefined) {
      style.pageBreakBefore = !!paraShape.pageBreakBefore;
    }
    if (paraShape.widowOrphan !== undefined) {
      style.widowControl = !!paraShape.widowOrphan;
    }

    return style;
  }

  private static createRun(text: string, charShapeIndex: number, docInfo: any): TextRun {
    const run: TextRun = { text };

    if (charShapeIndex >= 0 && docInfo?.charShapes?.[charShapeIndex]) {
      const charShape = docInfo.charShapes[charShapeIndex];
      run.charStyle = this.convertCharStyle(charShape, docInfo);
    }

    return run;
  }

  private static convertCharStyle(charShape: any, docInfo: any): CharacterStyle {
    let fontName: string | undefined;
    if (docInfo?.fontFaces) {
      const fontIds = charShape.fontId;
      if (Array.isArray(fontIds) && fontIds.length > 0) {
        const hangulFontId = fontIds[0];
        const fontFace = docInfo.fontFaces[hangulFontId];
        if (fontFace) {
          fontName = fontFace.name;
        }
      } else if (typeof fontIds === 'number') {
        const fontFace = docInfo.fontFaces[fontIds];
        if (fontFace) {
          fontName = fontFace.name;
        }
      }
    }

    const defaultSpacing = { hangul: 0, latin: 0, hanja: 0, japanese: 0, other: 0, symbol: 0, user: 0 };
    
    const charSpacing = charShape.charSpacing ?? charShape.spacing;
    let charSpacingObj = defaultSpacing;
    if (charSpacing !== undefined) {
      if (Array.isArray(charSpacing)) {
        charSpacingObj = {
          hangul: charSpacing[0] ?? 0,
          latin: charSpacing[1] ?? 0,
          hanja: charSpacing[2] ?? 0,
          japanese: charSpacing[3] ?? 0,
          other: charSpacing[4] ?? 0,
          symbol: charSpacing[5] ?? 0,
          user: charSpacing[6] ?? 0,
        };
      } else {
        charSpacingObj = {
          hangul: charSpacing, latin: charSpacing, hanja: charSpacing,
          japanese: charSpacing, other: charSpacing, symbol: charSpacing, user: charSpacing,
        };
      }
    }

    const charOffset = charShape.charOffset ?? charShape.offset;
    let charOffsetObj = defaultSpacing;
    if (charOffset !== undefined) {
      if (Array.isArray(charOffset)) {
        charOffsetObj = {
          hangul: charOffset[0] ?? 0,
          latin: charOffset[1] ?? 0,
          hanja: charOffset[2] ?? 0,
          japanese: charOffset[3] ?? 0,
          other: charOffset[4] ?? 0,
          symbol: charOffset[5] ?? 0,
          user: charOffset[6] ?? 0,
        };
      } else {
        charOffsetObj = {
          hangul: charOffset, latin: charOffset, hanja: charOffset,
          japanese: charOffset, other: charOffset, symbol: charOffset, user: charOffset,
        };
      }
    }

    const style: CharacterStyle = {
      fontName,
      fontSize: charShape.height ? charShape.height / 100 : 10,
      bold: !!charShape.bold,
      italic: !!charShape.italic,
      charSpacing: charSpacingObj,
      charOffset: charOffsetObj,
      useFontSpace: !!charShape.useFontSpace,
      useKerning: !!charShape.useKerning,
      emboss: !!charShape.emboss,
      engrave: !!charShape.engrave,
    };

    if (charShape.underline) {
      style.underline = true;
    }
    if (charShape.strikeout) {
      style.strikethrough = true;
    }
    if (charShape.superscript) {
      style.superscript = true;
    }
    if (charShape.subscript) {
      style.subscript = true;
    }

    return style;
  }

  private static convertAlign(align: number): ParagraphStyle['align'] {
    switch (align) {
      case 0: return 'Justify' as ParagraphStyle['align'];
      case 1: return 'Left' as ParagraphStyle['align'];
      case 2: return 'Right' as ParagraphStyle['align'];
      case 3: return 'Center' as ParagraphStyle['align'];
      case 4: return 'Distribute' as ParagraphStyle['align'];
      default: return undefined;
    }
  }

  private static convertColor(color: number): string {
    const r = color & 0xFF;
    const g = (color >> 8) & 0xFF;
    const b = (color >> 16) & 0xFF;
    return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
  }
}
