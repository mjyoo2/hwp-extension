import JSZip from 'jszip';
import { HwpxParser } from './HwpxParser';
import {
  HwpxContent,
  HwpxParagraph,
  TextRun,
  CharacterStyle,
  ParagraphStyle,
  HwpxTable,
  TableCell,
  TableRow,
  SectionElement,
  HwpxSection,
  HwpxImage,
  PageSettings,
  Footnote,
  Endnote,
  Memo,
  ColumnDef,
  CharShape,
  ParaShape,
  StyleDef,
  HwpxLine,
  HwpxRect,
  HwpxEllipse,
  HwpxEquation,
  HeaderFooter,
  HwpxTextBox,
} from './types';

type DocumentFormat = 'hwpx' | 'hwp';

const MAX_UNDO_STACK_SIZE = 50;

export class HwpxDocument {
  private _id: string;
  private _path: string;
  private _zip: JSZip | null;
  private _content: HwpxContent;
  private _isDirty = false;
  private _format: DocumentFormat;

  private _undoStack: string[] = [];
  private _redoStack: string[] = [];
  private _pendingTextReplacements: Array<{ oldText: string; newText: string; options: { caseSensitive?: boolean; regex?: boolean; replaceAll?: boolean } }> = [];
  private _pendingDirectTextUpdates: Array<{ oldText: string; newText: string }> = [];
  private _pendingTableRowInserts: Array<{ tableIndex: number; afterRowIndex: number; cellTexts?: string[] }> = [];
  private _pendingTableRowDeletes: Array<{ tableIndex: number; rowIndex: number }> = [];
  private _pendingTableColumnInserts: Array<{ tableIndex: number; afterColIndex: number }> = [];
  private _pendingTableColumnDeletes: Array<{ tableIndex: number; colIndex: number }> = [];
  private _pendingCellMerges: Array<{ tableIndex: number; startRow: number; startCol: number; endRow: number; endCol: number }> = [];
  private _pendingHeaderFooter: Array<{ sectionIndex: number; type: 'header' | 'footer'; text: string; includePageNumber: boolean; align: 'left' | 'center' | 'right' }> = [];
  private _hasStructuralChanges = false;

  private constructor(id: string, path: string, zip: JSZip | null, content: HwpxContent, format: DocumentFormat) {
    this._id = id;
    this._path = path;
    this._zip = zip;
    this._content = content;
    this._format = format;
  }

  public static async createFromBuffer(id: string, path: string, data: Buffer): Promise<HwpxDocument> {
    const extension = path.toLowerCase();

    if (extension.endsWith('.hwp')) {
      // HWP parsing would go here - for now return empty content
      const content: HwpxContent = {
        metadata: {},
        sections: [],
        images: new Map(),
        binItems: new Map(),
        binData: new Map(),
        footnotes: [],
        endnotes: [],
      };
      return new HwpxDocument(id, path, null, content, 'hwp');
    } else {
      const zip = await JSZip.loadAsync(data);
      const content = await HwpxParser.parse(zip);
      return new HwpxDocument(id, path, zip, content, 'hwpx');
    }
  }

  public static createNew(id: string, title?: string, creator?: string): HwpxDocument {
    const now = new Date().toISOString();
    const content: HwpxContent = {
      metadata: {
        title: title || 'Untitled',
        creator: creator || 'Unknown',
        createdDate: now,
        modifiedDate: now,
      },
      sections: [{
        id: Math.random().toString(36).substring(2, 11),
        elements: [{
          type: 'paragraph',
          data: {
            id: Math.random().toString(36).substring(2, 11),
            runs: [{ text: '' }],
          },
        }],
        pageSettings: {
          width: 59528,
          height: 84188,
          marginTop: 4252,
          marginBottom: 4252,
          marginLeft: 4252,
          marginRight: 4252,
        },
      }],
      images: new Map(),
      binItems: new Map(),
      binData: new Map(),
      footnotes: [],
      endnotes: [],
    };

    // Create a new zip with basic HWPX structure
    const zip = new JSZip();

    // Add minimal required files for a valid HWPX document
    zip.file('mimetype', 'application/hwp+zip');
    zip.file('Contents/header.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:title>${title || 'Untitled'}</hh:title>
  <hh:creator>${creator || 'Unknown'}</hh:creator>
  <hh:createdDate>${now}</hh:createdDate>
  <hh:modifiedDate>${now}</hh:modifiedDate>
</hh:head>`);
    zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hp:sec xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p>
    <hp:run>
      <hp:t></hp:t>
    </hp:run>
  </hp:p>
</hp:sec>`);

    return new HwpxDocument(id, 'new-document.hwpx', zip, content, 'hwpx');
  }

  get id(): string { return this._id; }
  get path(): string { return this._path; }
  get format(): DocumentFormat { return this._format; }
  get isDirty(): boolean { return this._isDirty; }
  get zip(): JSZip | null { return this._zip; }
  get content(): HwpxContent { return this._content; }

  // ============================================================
  // Undo/Redo
  // ============================================================

  private saveState(): void {
    const state = this.serializeContent();
    this._undoStack.push(state);
    if (this._undoStack.length > MAX_UNDO_STACK_SIZE) {
      this._undoStack.shift();
    }
    this._redoStack = [];
  }

  private serializeContent(): string {
    return JSON.stringify({
      sections: this._content.sections,
      metadata: this._content.metadata,
    });
  }

  private deserializeContent(state: string): void {
    const parsed = JSON.parse(state);
    this._content.sections = parsed.sections;
    this._content.metadata = parsed.metadata;
  }

  canUndo(): boolean { return this._undoStack.length > 0; }
  canRedo(): boolean { return this._redoStack.length > 0; }

  undo(): boolean {
    if (!this.canUndo()) return false;
    const currentState = this.serializeContent();
    this._redoStack.push(currentState);
    const previousState = this._undoStack.pop()!;
    this.deserializeContent(previousState);
    this._isDirty = true;
    return true;
  }

  redo(): boolean {
    if (!this.canRedo()) return false;
    const currentState = this.serializeContent();
    this._undoStack.push(currentState);
    const nextState = this._redoStack.pop()!;
    this.deserializeContent(nextState);
    this._isDirty = true;
    return true;
  }

  // ============================================================
  // Content Access
  // ============================================================

  getSerializableContent(): object {
    return {
      metadata: this._content.metadata,
      sections: this._content.sections,
      images: Array.from(this._content.images.entries()),
      footnotes: this._content.footnotes,
      endnotes: this._content.endnotes,
    };
  }

  getAllText(): string {
    let text = '';
    for (const section of this._content.sections) {
      for (const element of section.elements) {
        if (element.type === 'paragraph') {
          text += element.data.runs.map(r => r.text).join('') + '\n';
        }
      }
    }
    return text;
  }

  getStructure(): object {
    return {
      format: this._format,
      sections: this._content.sections.map((s, i) => {
        let paragraphs = 0, tables = 0, images = 0;
        for (const el of s.elements) {
          if (el.type === 'paragraph') paragraphs++;
          if (el.type === 'table') tables++;
          if (el.type === 'image') images++;
        }
        return { section: i, paragraphs, tables, images };
      }),
    };
  }

  // ============================================================
  // Paragraph Operations
  // ============================================================

  private findParagraphByPath(sectionIndex: number, elementIndex: number): HwpxParagraph | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;
    const element = section.elements[elementIndex];
    if (!element || element.type !== 'paragraph') return null;
    return element.data;
  }

  getParagraphs(sectionIndex?: number): Array<{ section: number; index: number; text: string; style?: ParagraphStyle }> {
    const paragraphs: Array<{ section: number; index: number; text: string; style?: ParagraphStyle }> = [];
    const sections = sectionIndex !== undefined
      ? [{ section: this._content.sections[sectionIndex], idx: sectionIndex }]
      : this._content.sections.map((s, i) => ({ section: s, idx: i }));

    for (const { section, idx } of sections) {
      if (!section) continue;
      section.elements.forEach((el, ei) => {
        if (el.type === 'paragraph') {
          paragraphs.push({
            section: idx,
            index: ei,
            text: el.data.runs.map(r => r.text).join(''),
            style: el.data.paraStyle,
          });
        }
      });
    }
    return paragraphs;
  }

  getParagraph(sectionIndex: number, paragraphIndex: number): { text: string; runs: TextRun[]; style?: ParagraphStyle } | null {
    const para = this.findParagraphByPath(sectionIndex, paragraphIndex);
    if (!para) return null;
    return {
      text: para.runs.map(r => r.text).join(''),
      runs: para.runs,
      style: para.paraStyle,
    };
  }

  updateParagraphText(sectionIndex: number, elementIndex: number, runIndex: number, text: string): void {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    if (!paragraph || !paragraph.runs[runIndex]) return;

    // Track the old text for XML update
    const oldText = paragraph.runs[runIndex].text;
    if (oldText && oldText !== text && this._zip) {
      this._pendingDirectTextUpdates.push({ oldText, newText: text });
    }

    this.saveState();
    paragraph.runs[runIndex].text = text;
    this._isDirty = true;
  }

  updateParagraphRuns(sectionIndex: number, elementIndex: number, runs: TextRun[]): void {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    if (!paragraph) return;
    this.saveState();
    paragraph.runs = runs;
    this._isDirty = true;
  }

  insertParagraph(sectionIndex: number, afterElementIndex: number, text: string = ''): number {
    const section = this._content.sections[sectionIndex];
    if (!section) return -1;

    this.saveState();
    const newParagraph: HwpxParagraph = {
      id: Math.random().toString(36).substring(2, 11),
      runs: [{ text }],
    };

    const newElement: SectionElement = { type: 'paragraph', data: newParagraph };
    section.elements.splice(afterElementIndex + 1, 0, newElement);
    this._isDirty = true;
    this._hasStructuralChanges = true;
    return afterElementIndex + 1;
  }

  deleteParagraph(sectionIndex: number, elementIndex: number): boolean {
    const section = this._content.sections[sectionIndex];
    if (!section || elementIndex < 0 || elementIndex >= section.elements.length) return false;

    this.saveState();
    section.elements.splice(elementIndex, 1);
    this._isDirty = true;
    this._hasStructuralChanges = true;
    return true;
  }

  createBulletedList(sectionIndex: number, items: string[], afterElementIndex?: number, bulletChar: string = '•'): number[] {
    const section = this._content.sections[sectionIndex];
    if (!section) return [];

    this.saveState();
    const insertIndex = afterElementIndex !== undefined ? afterElementIndex + 1 : section.elements.length;
    const insertedIndices: number[] = [];

    for (let i = 0; i < items.length; i++) {
      const newParagraph: HwpxParagraph = {
        id: Math.random().toString(36).substring(2, 11),
        runs: [{ text: `${bulletChar} ${items[i]}` }],
        listType: 'bullet',
        listLevel: 0,
      };

      const newElement: SectionElement = { type: 'paragraph', data: newParagraph };
      section.elements.splice(insertIndex + i, 0, newElement);
      insertedIndices.push(insertIndex + i);
    }

    this._isDirty = true;
    this._hasStructuralChanges = true;
    return insertedIndices;
  }

  createNumberedList(sectionIndex: number, items: string[], afterElementIndex?: number, startNumber: number = 1, format: 'decimal' | 'roman' | 'alpha' = 'decimal'): number[] {
    const section = this._content.sections[sectionIndex];
    if (!section) return [];

    this.saveState();
    const insertIndex = afterElementIndex !== undefined ? afterElementIndex + 1 : section.elements.length;
    const insertedIndices: number[] = [];

    for (let i = 0; i < items.length; i++) {
      const number = startNumber + i;
      let prefix: string;
      
      switch (format) {
        case 'roman':
          prefix = this.toRoman(number);
          break;
        case 'alpha':
          prefix = String.fromCharCode(96 + number);
          break;
        default:
          prefix = number.toString();
      }

      const newParagraph: HwpxParagraph = {
        id: Math.random().toString(36).substring(2, 11),
        runs: [{ text: `${prefix}. ${items[i]}` }],
        listType: 'number',
        listLevel: 0,
      };

      const newElement: SectionElement = { type: 'paragraph', data: newParagraph };
      section.elements.splice(insertIndex + i, 0, newElement);
      insertedIndices.push(insertIndex + i);
    }

    this._isDirty = true;
    this._hasStructuralChanges = true;
    return insertedIndices;
  }

  private toRoman(num: number): string {
    const romanNumerals: [number, string][] = [
      [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'],
      [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'],
      [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']
    ];
    let result = '';
    for (const [value, symbol] of romanNumerals) {
      while (num >= value) {
        result += symbol;
        num -= value;
      }
    }
    return result.toLowerCase();
  }

  setParagraphNumbering(sectionIndex: number, paragraphIndex: number, type: 'none' | 'bullet' | 'decimal' | 'roman' | 'alpha', level: number = 0): boolean {
    const paragraph = this.findParagraphByPath(sectionIndex, paragraphIndex);
    if (!paragraph) return false;

    this.saveState();
    
    if (type === 'none') {
      paragraph.listType = undefined;
      paragraph.listLevel = undefined;
    } else if (type === 'bullet') {
      paragraph.listType = 'bullet';
      paragraph.listLevel = level;
    } else {
      paragraph.listType = 'number';
      paragraph.listLevel = level;
    }

    this._isDirty = true;
    return true;
  }

  appendTextToParagraph(sectionIndex: number, elementIndex: number, text: string): void {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    if (!paragraph) return;

    this.saveState();
    paragraph.runs.push({ text });
    this._isDirty = true;
  }

  // ============================================================
  // Character Style Operations
  // ============================================================

  applyCharacterStyle(sectionIndex: number, elementIndex: number, runIndex: number, style: Partial<CharacterStyle>): void {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    if (!paragraph || !paragraph.runs[runIndex]) return;

    this.saveState();
    const run = paragraph.runs[runIndex];
    run.charStyle = { ...run.charStyle, ...style };
    this._isDirty = true;
  }

  getCharacterStyle(sectionIndex: number, elementIndex: number, runIndex?: number): CharacterStyle | CharacterStyle[] | null {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    if (!paragraph) return null;

    if (runIndex !== undefined) {
      return paragraph.runs[runIndex]?.charStyle || null;
    }
    return paragraph.runs.map(r => r.charStyle || {});
  }

  // ============================================================
  // Paragraph Style Operations
  // ============================================================

  applyParagraphStyle(sectionIndex: number, elementIndex: number, style: Partial<ParagraphStyle>): void {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    if (!paragraph) return;

    this.saveState();
    paragraph.paraStyle = { ...paragraph.paraStyle, ...style };
    this._isDirty = true;
  }

  getParagraphStyle(sectionIndex: number, elementIndex: number): ParagraphStyle | null {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    return paragraph?.paraStyle || null;
  }

  // ============================================================
  // Table Operations
  // ============================================================

  private findTable(sectionIndex: number, tableIndex: number): HwpxTable | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;
    const tables = section.elements.filter(el => el.type === 'table');
    return tables[tableIndex]?.data as HwpxTable || null;
  }

  getTables(): Array<{ section: number; index: number; rows: number; cols: number }> {
    const tables: Array<{ section: number; index: number; rows: number; cols: number }> = [];
    this._content.sections.forEach((section, si) => {
      let tableIndex = 0;
      section.elements.forEach(el => {
        if (el.type === 'table') {
          const table = el.data as HwpxTable;
          tables.push({
            section: si,
            index: tableIndex++,
            rows: table.rows.length,
            cols: table.rows[0]?.cells.length || 0,
          });
        }
      });
    });
    return tables;
  }

  getTable(sectionIndex: number, tableIndex: number): { rows: number; cols: number; data: any[][] } | null {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table) return null;

    return {
      rows: table.rows.length,
      cols: table.rows[0]?.cells.length || 0,
      data: table.rows.map(row => row.cells.map(cell => ({
        text: cell.paragraphs.map(p => p.runs.map(r => r.text).join('')).join('\n'),
        style: cell,
      }))),
    };
  }

  getTableCell(sectionIndex: number, tableIndex: number, row: number, col: number): { text: string; cell: TableCell } | null {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table) return null;
    const cell = table.rows[row]?.cells[col];
    if (!cell) return null;
    return {
      text: cell.paragraphs.map(p => p.runs.map(r => r.text).join('')).join('\n'),
      cell,
    };
  }

  updateTableCell(sectionIndex: number, tableIndex: number, row: number, col: number, text: string): boolean {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table) return false;
    const cell = table.rows[row]?.cells[col];
    if (!cell) return false;

    // Track old text for XML update
    if (cell.paragraphs.length > 0 && cell.paragraphs[0].runs.length > 0) {
      const oldText = cell.paragraphs[0].runs[0].text;
      if (oldText && oldText !== text && this._zip) {
        this._pendingDirectTextUpdates.push({ oldText, newText: text });
      }
    }

    this.saveState();
    if (cell.paragraphs.length > 0 && cell.paragraphs[0].runs.length > 0) {
      cell.paragraphs[0].runs[0].text = text;
    } else {
      cell.paragraphs = [{ id: Math.random().toString(36).substring(2, 11), runs: [{ text }] }];
    }
    this._isDirty = true;
    return true;
  }

  setCellProperties(sectionIndex: number, tableIndex: number, row: number, col: number, props: Partial<TableCell>): boolean {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table) return false;
    const cell = table.rows[row]?.cells[col];
    if (!cell) return false;

    this.saveState();
    Object.assign(cell, props);
    this._isDirty = true;
    return true;
  }

  mergeCells(sectionIndex: number, tableIndex: number, startRow: number, startCol: number, endRow: number, endCol: number): boolean {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table) return false;

    if (startRow < 0 || startCol < 0 || endRow >= table.rows.length || endCol >= (table.rows[0]?.cells.length || 0)) {
      return false;
    }
    if (startRow > endRow || startCol > endCol) return false;

    this.saveState();
    
    const rowSpan = endRow - startRow + 1;
    const colSpan = endCol - startCol + 1;

    const topLeftCell = table.rows[startRow]?.cells[startCol];
    if (!topLeftCell) return false;

    topLeftCell.rowSpan = rowSpan;
    topLeftCell.colSpan = colSpan;

    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        if (r === startRow && c === startCol) continue;
        
        const cell = table.rows[r]?.cells[c];
        if (cell) {
          cell.rowSpan = 0;
          cell.colSpan = 0;
        }
      }
    }

    this._pendingCellMerges.push({ tableIndex, startRow, startCol, endRow, endCol });
    this._isDirty = true;
    return true;
  }

  insertTableRow(sectionIndex: number, tableIndex: number, afterRowIndex: number, cellTexts?: string[]): boolean {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table || !table.rows[afterRowIndex]) return false;

    this.saveState();
    const templateRow = table.rows[afterRowIndex];
    const colCount = templateRow.cells.length;

    const newRow = {
      cells: Array.from({ length: colCount }, (_, i) => ({
        paragraphs: [{
          id: Math.random().toString(36).substring(2, 11),
          runs: [{ text: cellTexts?.[i] || '' }],
        }],
        colAddr: i,
        rowAddr: afterRowIndex + 1,
        colSpan: 1,
        rowSpan: 1,
      })),
    };

    table.rows.splice(afterRowIndex + 1, 0, newRow as any);
    this._pendingTableRowInserts.push({ tableIndex, afterRowIndex, cellTexts });
    this._isDirty = true;
    return true;
  }

  deleteTableRow(sectionIndex: number, tableIndex: number, rowIndex: number): boolean {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table || table.rows.length <= 1) return false;

    this.saveState();
    table.rows.splice(rowIndex, 1);
    this._pendingTableRowDeletes.push({ tableIndex, rowIndex });
    this._isDirty = true;
    return true;
  }

  insertTableColumn(sectionIndex: number, tableIndex: number, afterColIndex: number): boolean {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table) return false;

    this.saveState();
    for (const row of table.rows) {
      row.cells.splice(afterColIndex + 1, 0, {
        paragraphs: [{
          id: Math.random().toString(36).substring(2, 11),
          runs: [{ text: '' }],
        }],
        colAddr: afterColIndex + 1,
        rowAddr: 0,
        colSpan: 1,
        rowSpan: 1,
      } as any);
    }
    this._pendingTableColumnInserts.push({ tableIndex, afterColIndex });
    this._isDirty = true;
    return true;
  }

  deleteTableColumn(sectionIndex: number, tableIndex: number, colIndex: number): boolean {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table || (table.rows[0]?.cells.length || 0) <= 1) return false;

    this.saveState();
    for (const row of table.rows) {
      row.cells.splice(colIndex, 1);
    }
    this._pendingTableColumnDeletes.push({ tableIndex, colIndex });
    this._isDirty = true;
    return true;
  }

  getTableAsCsv(sectionIndex: number, tableIndex: number, delimiter: string = ','): string | null {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table) return null;

    return table.rows.map(row =>
      row.cells.map(cell => {
        const text = cell.paragraphs.map(p => p.runs.map(r => r.text).join('')).join(' ');
        if (text.includes(delimiter) || text.includes('"') || text.includes('\n')) {
          return `"${text.replace(/"/g, '""')}"`;
        }
        return text;
      }).join(delimiter)
    ).join('\n');
  }

  // ============================================================
  // Search & Replace
  // ============================================================

  searchText(query: string, options: { caseSensitive?: boolean; regex?: boolean } = {}): Array<{ section: number; element: number; text: string; matches: string[]; count: number }> {
    const { caseSensitive = false, regex = false } = options;
    let pattern: RegExp;

    if (regex) {
      pattern = new RegExp(query, caseSensitive ? 'g' : 'gi');
    } else {
      const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      pattern = new RegExp(escaped, caseSensitive ? 'g' : 'gi');
    }

    const results: Array<{ section: number; element: number; text: string; matches: string[]; count: number }> = [];

    this._content.sections.forEach((section, si) => {
      section.elements.forEach((el, ei) => {
        if (el.type === 'paragraph') {
          const text = el.data.runs.map(r => r.text).join('');
          const found = text.match(pattern);
          if (found) {
            results.push({
              section: si,
              element: ei,
              text,
              matches: found,
              count: found.length,
            });
          }
        }
      });
    });

    return results;
  }

  replaceText(oldText: string, newText: string, options: { caseSensitive?: boolean; regex?: boolean; replaceAll?: boolean } = {}): number {
    const { caseSensitive = false, regex = false, replaceAll = true } = options;
    let pattern: RegExp;

    if (regex) {
      pattern = new RegExp(oldText, caseSensitive ? (replaceAll ? 'g' : '') : (replaceAll ? 'gi' : 'i'));
    } else {
      const escaped = oldText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      pattern = new RegExp(escaped, caseSensitive ? (replaceAll ? 'g' : '') : (replaceAll ? 'gi' : 'i'));
    }

    this.saveState();
    let count = 0;

    // Update in-memory content
    for (const section of this._content.sections) {
      for (const element of section.elements) {
        if (element.type === 'paragraph') {
          for (const run of element.data.runs) {
            const matches = run.text.match(pattern);
            if (matches) {
              count += matches.length;
              run.text = run.text.replace(pattern, newText);
            }
          }
        }
        // Also handle table cells
        if (element.type === 'table') {
          const table = element.data as HwpxTable;
          for (const row of table.rows) {
            for (const cell of row.cells) {
              for (const para of cell.paragraphs) {
                for (const run of para.runs) {
                  const matches = run.text.match(pattern);
                  if (matches) {
                    count += matches.length;
                    run.text = run.text.replace(pattern, newText);
                  }
                }
              }
            }
          }
        }
      }
    }

    // Also update directly in the ZIP XML files for safe saving
    if (count > 0 && this._zip) {
      this._pendingTextReplacements = this._pendingTextReplacements || [];
      this._pendingTextReplacements.push({ oldText, newText, options });
      this._isDirty = true;
    }

    return count;
  }

  // ============================================================
  // Metadata
  // ============================================================

  getMetadata(): HwpxContent['metadata'] {
    return this._content.metadata;
  }

  setMetadata(metadata: Partial<HwpxContent['metadata']>): void {
    this.saveState();
    this._content.metadata = { ...this._content.metadata, ...metadata };
    this._isDirty = true;
  }

  // ============================================================
  // Page Settings
  // ============================================================

  getPageSettings(sectionIndex: number = 0): PageSettings | null {
    const section = this._content.sections[sectionIndex];
    return section?.pageSettings || null;
  }

  setPageSettings(sectionIndex: number, settings: Partial<PageSettings>): boolean {
    const section = this._content.sections[sectionIndex];
    if (!section) return false;

    this.saveState();
    section.pageSettings = { ...section.pageSettings, ...settings } as PageSettings;
    this._isDirty = true;
    return true;
  }

  // ============================================================
  // Statistics
  // ============================================================

  getWordCount(): { characters: number; charactersNoSpaces: number; words: number; paragraphs: number } {
    let characters = 0;
    let charactersNoSpaces = 0;
    let words = 0;
    let paragraphs = 0;

    for (const section of this._content.sections) {
      for (const element of section.elements) {
        if (element.type === 'paragraph') {
          paragraphs++;
          const text = element.data.runs.map(r => r.text).join('');
          characters += text.length;
          charactersNoSpaces += text.replace(/\s/g, '').length;
          words += text.trim().split(/\s+/).filter(w => w.length > 0).length;
        }
      }
    }

    return { characters, charactersNoSpaces, words, paragraphs };
  }

  // ============================================================
  // Copy/Move Operations
  // ============================================================

  copyParagraph(sourceSection: number, sourceParagraph: number, targetSection: number, targetAfter: number): boolean {
    const srcSection = this._content.sections[sourceSection];
    const tgtSection = this._content.sections[targetSection];
    if (!srcSection || !tgtSection) return false;

    const srcElement = srcSection.elements[sourceParagraph];
    if (!srcElement || srcElement.type !== 'paragraph') return false;

    this.saveState();
    const copy = JSON.parse(JSON.stringify(srcElement));
    copy.data.id = Math.random().toString(36).substring(2, 11);
    tgtSection.elements.splice(targetAfter + 1, 0, copy);
    this._isDirty = true;
    return true;
  }

  moveParagraph(sourceSection: number, sourceParagraph: number, targetSection: number, targetAfter: number): boolean {
    const srcSection = this._content.sections[sourceSection];
    const tgtSection = this._content.sections[targetSection];
    if (!srcSection || !tgtSection) return false;

    const srcElement = srcSection.elements[sourceParagraph];
    if (!srcElement || srcElement.type !== 'paragraph') return false;

    this.saveState();
    srcSection.elements.splice(sourceParagraph, 1);
    tgtSection.elements.splice(targetAfter + 1, 0, srcElement);
    this._isDirty = true;
    return true;
  }

  // ============================================================
  // Images
  // ============================================================

  getImages(): Array<{ id: string; width: number; height: number }> {
    return Array.from(this._content.images.values()).map(img => ({
      id: img.id,
      width: img.width,
      height: img.height,
    }));
  }

  // ============================================================
  // Table Creation
  // ============================================================

  insertTable(sectionIndex: number, afterElementIndex: number, rows: number, cols: number, options?: { width?: number; cellWidth?: number }): { tableIndex: number } | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;
    if (rows <= 0 || cols <= 0) return null;

    this.saveState();

    const tableId = Math.random().toString(36).substring(2, 11);
    const defaultWidth = options?.width || 42520; // Default table width in hwpunit
    const cellWidth = options?.cellWidth || Math.floor(defaultWidth / cols);

    const tableRows: TableRow[] = [];
    for (let r = 0; r < rows; r++) {
      const cells: TableCell[] = [];
      for (let c = 0; c < cols; c++) {
        cells.push({
          colAddr: c,
          rowAddr: r,
          colSpan: 1,
          rowSpan: 1,
          width: cellWidth,
          paragraphs: [{
            id: Math.random().toString(36).substring(2, 11),
            runs: [{ text: '' }],
          }],
        });
      }
      tableRows.push({ cells });
    }

    const newTable: HwpxTable = {
      id: tableId,
      rowCount: rows,
      colCount: cols,
      rows: tableRows,
      width: defaultWidth,
    };

    const newElement: SectionElement = { type: 'table', data: newTable };
    section.elements.splice(afterElementIndex + 1, 0, newElement);

    // Calculate table index
    let tableIndex = 0;
    for (let i = 0; i <= afterElementIndex + 1; i++) {
      if (section.elements[i]?.type === 'table') {
        if (i === afterElementIndex + 1) break;
        tableIndex++;
      }
    }

    this._isDirty = true;
    this._hasStructuralChanges = true;
    return { tableIndex };
  }

  // ============================================================
  // Header/Footer Operations
  // ============================================================

  getHeader(sectionIndex: number): { paragraphs: any[] } | null {
    const section = this._content.sections[sectionIndex];
    if (!section || !section.header) return null;
    return {
      paragraphs: section.header.paragraphs.map(p => ({
        id: p.id,
        text: p.runs.map(r => r.text).join(''),
        runs: p.runs,
      })),
    };
  }

  setHeader(sectionIndex: number, text: string, align: 'left' | 'center' | 'right' = 'center'): boolean {
    const section = this._content.sections[sectionIndex];
    if (!section) return false;

    this.saveState();

    const headerParagraph: HwpxParagraph = {
      id: Math.random().toString(36).substring(2, 11),
      runs: [{ text }],
      paraStyle: { align },
    };

    if (!section.header) {
      section.header = {
        paragraphs: [headerParagraph],
      };
    } else {
      section.header.paragraphs = [headerParagraph];
    }

    this._pendingHeaderFooter.push({ sectionIndex, type: 'header', text, includePageNumber: false, align });
    this._isDirty = true;
    return true;
  }

  getFooter(sectionIndex: number): { paragraphs: any[] } | null {
    const section = this._content.sections[sectionIndex];
    if (!section || !section.footer) return null;
    return {
      paragraphs: section.footer.paragraphs.map(p => ({
        id: p.id,
        text: p.runs.map(r => r.text).join(''),
        runs: p.runs,
      })),
    };
  }

  setFooter(sectionIndex: number, text: string, includePageNumber: boolean = false, align: 'left' | 'center' | 'right' = 'center'): boolean {
    const section = this._content.sections[sectionIndex];
    if (!section) return false;

    this.saveState();

    const runs: TextRun[] = [];
    if (text) {
      runs.push({ text });
    }
    if (includePageNumber) {
      runs.push({ text: '', pageNumber: true });
    }

    const footerParagraph: HwpxParagraph = {
      id: Math.random().toString(36).substring(2, 11),
      runs,
      paraStyle: { align },
    };

    if (!section.footer) {
      section.footer = {
        paragraphs: [footerParagraph],
      };
    } else {
      section.footer.paragraphs = [footerParagraph];
    }

    this._pendingHeaderFooter.push({ sectionIndex, type: 'footer', text, includePageNumber, align });
    this._isDirty = true;
    return true;
  }

  // ============================================================
  // Footnote/Endnote Operations
  // ============================================================

  getFootnotes(): Footnote[] {
    return this._content.footnotes || [];
  }

  insertFootnote(sectionIndex: number, paragraphIndex: number, text: string): { id: string } | null {
    const paragraph = this.findParagraphByPath(sectionIndex, paragraphIndex);
    if (!paragraph) return null;

    this.saveState();

    const footnoteId = Math.random().toString(36).substring(2, 11);
    const footnoteNumber = (this._content.footnotes?.length || 0) + 1;

    const footnote: Footnote = {
      id: footnoteId,
      number: footnoteNumber,
      type: 'footnote',
      paragraphs: [{
        id: Math.random().toString(36).substring(2, 11),
        runs: [{ text }],
      }],
    };

    if (!this._content.footnotes) {
      this._content.footnotes = [];
    }
    this._content.footnotes.push(footnote);

    // Add footnote reference to the paragraph
    paragraph.runs.push({
      text: '',
      footnoteRef: footnoteNumber,
    });

    this._isDirty = true;
    return { id: footnoteId };
  }

  getEndnotes(): Endnote[] {
    return this._content.endnotes || [];
  }

  insertEndnote(sectionIndex: number, paragraphIndex: number, text: string): { id: string } | null {
    const paragraph = this.findParagraphByPath(sectionIndex, paragraphIndex);
    if (!paragraph) return null;

    this.saveState();

    const endnoteId = Math.random().toString(36).substring(2, 11);
    const endnoteNumber = (this._content.endnotes?.length || 0) + 1;

    const endnote: Endnote = {
      id: endnoteId,
      number: endnoteNumber,
      paragraphs: [{
        id: Math.random().toString(36).substring(2, 11),
        runs: [{ text }],
      }],
    };

    if (!this._content.endnotes) {
      this._content.endnotes = [];
    }
    this._content.endnotes.push(endnote);

    // Add endnote reference to the paragraph
    paragraph.runs.push({
      text: '',
      endnoteRef: endnoteNumber,
    });

    this._isDirty = true;
    return { id: endnoteId };
  }

  // ============================================================
  // Bookmark/Hyperlink Operations
  // ============================================================

  getBookmarks(): { name: string; section: number; paragraph: number }[] {
    const bookmarks: { name: string; section: number; paragraph: number }[] = [];

    this._content.sections.forEach((section, si) => {
      section.elements.forEach((el, ei) => {
        if (el.type === 'paragraph') {
          for (const run of el.data.runs) {
            if (run.field?.fieldType === 'Bookmark' || run.field?.fieldType === 'bookmark') {
              bookmarks.push({
                name: run.field.name || '',
                section: si,
                paragraph: ei,
              });
            }
          }
        }
      });
    });

    return bookmarks;
  }

  insertBookmark(sectionIndex: number, paragraphIndex: number, name: string): boolean {
    const paragraph = this.findParagraphByPath(sectionIndex, paragraphIndex);
    if (!paragraph) return false;

    this.saveState();

    paragraph.runs.push({
      text: '',
      field: {
        fieldType: 'Bookmark',
        name,
      },
    });

    this._isDirty = true;
    return true;
  }

  getHyperlinks(): { url: string; text: string; section: number; paragraph: number }[] {
    const hyperlinks: { url: string; text: string; section: number; paragraph: number }[] = [];

    this._content.sections.forEach((section, si) => {
      section.elements.forEach((el, ei) => {
        if (el.type === 'paragraph') {
          for (const run of el.data.runs) {
            if (run.hyperlink) {
              hyperlinks.push({
                url: run.hyperlink.url,
                text: run.text || run.hyperlink.name || '',
                section: si,
                paragraph: ei,
              });
            }
          }
        }
      });
    });

    return hyperlinks;
  }

  insertHyperlink(sectionIndex: number, paragraphIndex: number, url: string, text: string): boolean {
    const paragraph = this.findParagraphByPath(sectionIndex, paragraphIndex);
    if (!paragraph) return false;

    this.saveState();

    paragraph.runs.push({
      text,
      hyperlink: {
        fieldType: 'Hyperlink',
        url,
        name: text,
      },
    });

    this._isDirty = true;
    return true;
  }

  // ============================================================
  // Image Operations
  // ============================================================

  insertImage(sectionIndex: number, afterElementIndex: number, imageData: { data: string; mimeType: string; width: number; height: number }): { id: string } | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;

    this.saveState();

    const imageId = Math.random().toString(36).substring(2, 11);
    const binaryId = Math.random().toString(36).substring(2, 11);

    const newImage: HwpxImage = {
      id: imageId,
      binaryId,
      width: imageData.width,
      height: imageData.height,
      data: imageData.data,
      mimeType: imageData.mimeType,
    };

    // Store image in the images map
    this._content.images.set(imageId, newImage);

    // Store binary data
    this._content.binData.set(binaryId, {
      id: binaryId,
      encoding: 'Base64',
      data: imageData.data,
    });

    // Add image element to section
    const newElement: SectionElement = { type: 'image', data: newImage };
    section.elements.splice(afterElementIndex + 1, 0, newElement);

    this._isDirty = true;
    return { id: imageId };
  }

  updateImageSize(imageId: string, width: number, height: number): boolean {
    const image = this._content.images.get(imageId);
    if (!image) return false;

    this.saveState();

    image.width = width;
    image.height = height;

    this._isDirty = true;
    return true;
  }

  deleteImage(imageId: string): boolean {
    const image = this._content.images.get(imageId);
    if (!image) return false;

    this.saveState();

    // Remove from images map
    this._content.images.delete(imageId);

    // Remove binary data if exists
    if (image.binaryId) {
      this._content.binData.delete(image.binaryId);
    }

    // Remove from sections
    for (const section of this._content.sections) {
      const index = section.elements.findIndex(el => el.type === 'image' && el.data.id === imageId);
      if (index !== -1) {
        section.elements.splice(index, 1);
        break;
      }
    }

    this._isDirty = true;
    return true;
  }

  // ============================================================
  // Drawing Objects (Line, Rect, Ellipse)
  // ============================================================

  insertLine(sectionIndex: number, x1: number, y1: number, x2: number, y2: number, options?: { color?: string; width?: number }): { id: string } | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;

    this.saveState();

    const lineId = Math.random().toString(36).substring(2, 11);

    const newLine: HwpxLine = {
      id: lineId,
      x1,
      y1,
      x2,
      y2,
      strokeColor: options?.color || '#000000',
      strokeWidth: options?.width || 1,
    };

    const newElement: SectionElement = { type: 'line', data: newLine };
    section.elements.push(newElement);

    this._isDirty = true;
    return { id: lineId };
  }

  insertRect(sectionIndex: number, x: number, y: number, width: number, height: number, options?: { fillColor?: string; strokeColor?: string }): { id: string } | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;

    this.saveState();

    const rectId = Math.random().toString(36).substring(2, 11);

    const newRect: HwpxRect = {
      id: rectId,
      x,
      y,
      width,
      height,
      fillColor: options?.fillColor,
      strokeColor: options?.strokeColor || '#000000',
    };

    const newElement: SectionElement = { type: 'rect', data: newRect };
    section.elements.push(newElement);

    this._isDirty = true;
    return { id: rectId };
  }

  insertEllipse(sectionIndex: number, cx: number, cy: number, rx: number, ry: number, options?: { fillColor?: string; strokeColor?: string }): { id: string } | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;

    this.saveState();

    const ellipseId = Math.random().toString(36).substring(2, 11);

    const newEllipse: HwpxEllipse = {
      id: ellipseId,
      cx,
      cy,
      rx,
      ry,
      fillColor: options?.fillColor,
      strokeColor: options?.strokeColor || '#000000',
    };

    const newElement: SectionElement = { type: 'ellipse', data: newEllipse };
    section.elements.push(newElement);

    this._isDirty = true;
    return { id: ellipseId };
  }

  insertTextBox(sectionIndex: number, x: number, y: number, width: number, height: number, text: string, options?: { fillColor?: string; strokeColor?: string; strokeWidth?: number }): { id: string } | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;

    this.saveState();

    const textBoxId = Math.random().toString(36).substring(2, 11);

    const newTextBox: HwpxTextBox = {
      id: textBoxId,
      x,
      y,
      width,
      height,
      paragraphs: [{
        id: Math.random().toString(36).substring(2, 11),
        runs: [{ text }],
      }],
      fillColor: options?.fillColor,
      strokeColor: options?.strokeColor || '#000000',
      strokeWidth: options?.strokeWidth ?? 1,
    };

    const newElement: SectionElement = { type: 'textbox', data: newTextBox };
    section.elements.push(newElement);

    this._isDirty = true;
    this._hasStructuralChanges = true;
    return { id: textBoxId };
  }

  getTextBoxes(): { id: string; x: number; y: number; width: number; height: number; text: string }[] {
    const textBoxes: { id: string; x: number; y: number; width: number; height: number; text: string }[] = [];

    for (const section of this._content.sections) {
      for (const element of section.elements) {
        if (element.type === 'textbox') {
          const tb = element.data as HwpxTextBox;
          textBoxes.push({
            id: tb.id,
            x: tb.x,
            y: tb.y,
            width: tb.width,
            height: tb.height,
            text: tb.paragraphs.map(p => p.runs.map(r => r.text).join('')).join('\n'),
          });
        }
      }
    }

    return textBoxes;
  }

  updateTextBoxText(textBoxId: string, text: string): boolean {
    for (const section of this._content.sections) {
      for (const element of section.elements) {
        if (element.type === 'textbox' && (element.data as HwpxTextBox).id === textBoxId) {
          this.saveState();
          const tb = element.data as HwpxTextBox;
          if (tb.paragraphs.length > 0 && tb.paragraphs[0].runs.length > 0) {
            tb.paragraphs[0].runs[0].text = text;
          } else {
            tb.paragraphs = [{
              id: Math.random().toString(36).substring(2, 11),
              runs: [{ text }],
            }];
          }
          this._isDirty = true;
          this._hasStructuralChanges = true;
          return true;
        }
      }
    }
    return false;
  }

  deleteTextBox(textBoxId: string): boolean {
    for (const section of this._content.sections) {
      for (let i = 0; i < section.elements.length; i++) {
        const element = section.elements[i];
        if (element.type === 'textbox' && (element.data as HwpxTextBox).id === textBoxId) {
          this.saveState();
          section.elements.splice(i, 1);
          this._isDirty = true;
          this._hasStructuralChanges = true;
          return true;
        }
      }
    }
    return false;
  }

  // ============================================================
  // Equation Operations
  // ============================================================

  insertEquation(sectionIndex: number, afterElementIndex: number, script: string): { id: string } | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;

    this.saveState();

    const equationId = Math.random().toString(36).substring(2, 11);

    const newEquation: HwpxEquation = {
      id: equationId,
      script,
      lineMode: false,
      baseUnit: 1000,
    };

    const newElement: SectionElement = { type: 'equation', data: newEquation };
    section.elements.splice(afterElementIndex + 1, 0, newElement);

    this._isDirty = true;
    return { id: equationId };
  }

  getEquations(): { id: string; script: string }[] {
    const equations: { id: string; script: string }[] = [];

    for (const section of this._content.sections) {
      for (const element of section.elements) {
        if (element.type === 'equation') {
          equations.push({
            id: element.data.id,
            script: element.data.script || '',
          });
        }
      }
    }

    return equations;
  }

  // ============================================================
  // Memo Operations
  // ============================================================

  getMemos(): Memo[] {
    const memos: Memo[] = [];

    for (const section of this._content.sections) {
      if (section.memos) {
        memos.push(...section.memos);
      }
    }

    return memos;
  }

  insertMemo(sectionIndex: number, paragraphIndex: number, content: string, author?: string): { id: string } | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;

    const paragraph = this.findParagraphByPath(sectionIndex, paragraphIndex);
    if (!paragraph) return null;

    this.saveState();

    const memoId = Math.random().toString(36).substring(2, 11);

    const memo: Memo = {
      id: memoId,
      author: author || 'Unknown',
      date: new Date().toISOString(),
      content: [content],
    };

    if (!section.memos) {
      section.memos = [];
    }
    section.memos.push(memo);

    // Mark the paragraph as having a memo
    if (paragraph.runs.length > 0) {
      paragraph.runs[paragraph.runs.length - 1].hasMemo = true;
      paragraph.runs[paragraph.runs.length - 1].memoId = memoId;
    }

    this._isDirty = true;
    return { id: memoId };
  }

  deleteMemo(memoId: string): boolean {
    let found = false;

    for (const section of this._content.sections) {
      if (section.memos) {
        const index = section.memos.findIndex(m => m.id === memoId);
        if (index !== -1) {
          this.saveState();
          section.memos.splice(index, 1);
          found = true;
          break;
        }
      }
    }

    if (found) {
      // Remove memo reference from paragraphs
      for (const section of this._content.sections) {
        for (const element of section.elements) {
          if (element.type === 'paragraph') {
            for (const run of element.data.runs) {
              if (run.memoId === memoId) {
                run.hasMemo = false;
                run.memoId = undefined;
              }
            }
          }
        }
      }

      this._isDirty = true;
    }

    return found;
  }

  // ============================================================
  // Section Operations
  // ============================================================

  getSections(): { index: number; pageSettings: PageSettings }[] {
    return this._content.sections.map((section, index) => ({
      index,
      pageSettings: section.pageSettings || {
        width: 59528,
        height: 84188,
        marginTop: 4252,
        marginBottom: 4252,
        marginLeft: 4252,
        marginRight: 4252,
      },
    }));
  }

  insertSection(afterSectionIndex: number): number {
    this.saveState();

    const newSection: HwpxSection = {
      id: Math.random().toString(36).substring(2, 11),
      elements: [{
        type: 'paragraph',
        data: {
          id: Math.random().toString(36).substring(2, 11),
          runs: [{ text: '' }],
        },
      }],
      pageSettings: {
        width: 59528,
        height: 84188,
        marginTop: 4252,
        marginBottom: 4252,
        marginLeft: 4252,
        marginRight: 4252,
      },
    };

    const insertIndex = afterSectionIndex + 1;
    this._content.sections.splice(insertIndex, 0, newSection);

    this._isDirty = true;
    this._hasStructuralChanges = true;
    return insertIndex;
  }

  deleteSection(sectionIndex: number): boolean {
    if (sectionIndex < 0 || sectionIndex >= this._content.sections.length) return false;
    if (this._content.sections.length <= 1) return false;

    this.saveState();
    this._content.sections.splice(sectionIndex, 1);
    this._isDirty = true;
    this._hasStructuralChanges = true;
    return true;
  }

  // ============================================================
  // Style Operations
  // ============================================================

  getStyles(): { id: number; name: string; type: string }[] {
    if (!this._content.styles?.styles) return [];

    return Array.from(this._content.styles.styles.values()).map(style => ({
      id: style.id,
      name: style.name || '',
      type: style.type || 'Para',
    }));
  }

  getCharShapes(): CharShape[] {
    if (!this._content.styles?.charShapes) return [];
    return Array.from(this._content.styles.charShapes.values());
  }

  getParaShapes(): ParaShape[] {
    if (!this._content.styles?.paraShapes) return [];
    return Array.from(this._content.styles.paraShapes.values());
  }

  applyStyle(sectionIndex: number, paragraphIndex: number, styleId: number): boolean {
    const paragraph = this.findParagraphByPath(sectionIndex, paragraphIndex);
    if (!paragraph) return false;

    if (!this._content.styles?.styles) return false;
    const style = this._content.styles.styles.get(styleId);
    if (!style) return false;

    this.saveState();

    paragraph.style = styleId;

    // Apply paragraph shape if defined
    if (style.paraPrIdRef !== undefined && this._content.styles.paraShapes) {
      const paraShape = this._content.styles.paraShapes.get(style.paraPrIdRef);
      if (paraShape) {
        paragraph.paraStyle = {
          align: paraShape.align?.toLowerCase() as ParagraphStyle['align'],
          lineSpacing: paraShape.lineSpacing,
          marginTop: paraShape.marginTop,
          marginBottom: paraShape.marginBottom,
          marginLeft: paraShape.marginLeft,
          marginRight: paraShape.marginRight,
          firstLineIndent: paraShape.firstLineIndent,
        };
      }
    }

    // Apply character shape if defined
    if (style.charPrIdRef !== undefined && this._content.styles.charShapes) {
      const charShape = this._content.styles.charShapes.get(style.charPrIdRef);
      if (charShape) {
        for (const run of paragraph.runs) {
          run.charStyle = {
            bold: charShape.bold,
            italic: charShape.italic,
            underline: charShape.underline,
            fontSize: charShape.height ? charShape.height / 100 : undefined,
            fontColor: charShape.textColor,
          };
        }
      }
    }

    this._isDirty = true;
    return true;
  }

  // ============================================================
  // Column Definition Operations
  // ============================================================

  getColumnDef(sectionIndex: number): ColumnDef | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;
    return section.columnDef || null;
  }

  setColumnDef(sectionIndex: number, columns: number, gap?: number): boolean {
    const section = this._content.sections[sectionIndex];
    if (!section) return false;
    if (columns < 1) return false;

    this.saveState();

    const defaultGap = gap || 850; // Default gap in hwpunit (approx 8.5mm)

    section.columnDef = {
      type: 'Newspaper',
      count: columns,
      sameSize: true,
      sameGap: defaultGap,
      columns: Array.from({ length: columns }, () => ({
        width: 0, // Will be calculated based on page width
        gap: defaultGap,
      })),
    };

    this._isDirty = true;
    return true;
  }

  // ============================================================
  // Save
  // ============================================================

  async save(): Promise<Buffer> {
    if (!this._zip) throw new Error('Cannot save HWP files');
    await this.syncContentToZip();
    return await this._zip.generateAsync({ type: 'nodebuffer' });
  }

  private async syncContentToZip(): Promise<void> {
    if (!this._zip) return;

    const hasTextReplacements = this._pendingTextReplacements && this._pendingTextReplacements.length > 0;
    const hasDirectTextUpdates = this._pendingDirectTextUpdates && this._pendingDirectTextUpdates.length > 0;
    const hasTableRowInserts = this._pendingTableRowInserts && this._pendingTableRowInserts.length > 0;
    const hasTableRowDeletes = this._pendingTableRowDeletes && this._pendingTableRowDeletes.length > 0;
    const hasTableColumnInserts = this._pendingTableColumnInserts && this._pendingTableColumnInserts.length > 0;
    const hasTableColumnDeletes = this._pendingTableColumnDeletes && this._pendingTableColumnDeletes.length > 0;
    const hasCellMerges = this._pendingCellMerges && this._pendingCellMerges.length > 0;
    const hasHeaderFooter = this._pendingHeaderFooter && this._pendingHeaderFooter.length > 0;
    const hasTableStructuralChanges = hasTableRowInserts || hasTableRowDeletes || hasTableColumnInserts || hasTableColumnDeletes || hasCellMerges;
    
    const hasOnlyTextChanges = (hasTextReplacements || hasDirectTextUpdates) && 
                               !this._hasStructuralChanges && 
                               !hasTableStructuralChanges &&
                               !hasHeaderFooter;

    if (!hasOnlyTextChanges && !hasTableStructuralChanges && !hasHeaderFooter) {
      await this.syncStructuralChangesToZip();
    }

    if (hasTableStructuralChanges) {
      await this.applyTableStructuralChangesToXml();
      this._pendingTableRowInserts = [];
      this._pendingTableRowDeletes = [];
      this._pendingTableColumnInserts = [];
      this._pendingTableColumnDeletes = [];
    }

    if (hasHeaderFooter) {
      await this.applyHeaderFooterToXml();
      this._pendingHeaderFooter = [];
    }

    if (hasDirectTextUpdates) {
      await this.applyDirectTextUpdatesToXml();
      this._pendingDirectTextUpdates = [];
    }

    if (hasTextReplacements) {
      await this.applyTextReplacementsToXml();
      this._pendingTextReplacements = [];
    }

    await this.syncMetadataToZip();

    this._isDirty = false;
    this._hasStructuralChanges = false;
  }

  private async applyHeaderFooterToXml(): Promise<void> {
    if (!this._zip) return;

    for (const item of this._pendingHeaderFooter) {
      const sectionPath = `Contents/section${item.sectionIndex}.xml`;
      const file = this._zip.file(sectionPath);
      if (!file) continue;

      let xml = await file.async('string');
      
      const tagName = item.type === 'header' ? 'hp:header' : 'hp:footer';
      const existingTagRegex = new RegExp(`<${tagName}[^>]*>[\\s\\S]*?<\\/${tagName}>`, 'g');
      xml = xml.replace(existingTagRegex, '');

      let content = '';
      if (item.text) {
        content += `<hp:t>${this.escapeXml(item.text)}</hp:t>`;
      }
      if (item.includePageNumber) {
        content += `<hp:pageNum/>`;
      }

      const alignAttr = item.align !== 'left' ? ` align="${item.align}"` : '';
      const headerFooterXml = `<${tagName}><hp:p${alignAttr}><hp:run>${content}</hp:run></hp:p></${tagName}>`;

      const closingSecTag = '</hs:sec>';
      const closingSecTagAlt = '</hp:sec>';
      
      if (xml.includes(closingSecTag)) {
        xml = xml.replace(closingSecTag, headerFooterXml + closingSecTag);
      } else if (xml.includes(closingSecTagAlt)) {
        xml = xml.replace(closingSecTagAlt, headerFooterXml + closingSecTagAlt);
      }

      this._zip.file(sectionPath, xml);
    }
  }

  private async applyTableStructuralChangesToXml(): Promise<void> {
    if (!this._zip) return;

    let sectionIndex = 0;
    while (true) {
      const sectionPath = `Contents/section${sectionIndex}.xml`;
      const file = this._zip.file(sectionPath);
      if (!file) break;

      let xml = await file.async('string');

      for (const insert of this._pendingTableRowInserts) {
        xml = this.insertTableRowInXml(xml, insert.tableIndex, insert.afterRowIndex, insert.cellTexts);
      }

      for (const del of this._pendingTableRowDeletes) {
        xml = this.deleteTableRowInXml(xml, del.tableIndex, del.rowIndex);
      }

      for (const insert of this._pendingTableColumnInserts) {
        xml = this.insertTableColumnInXml(xml, insert.tableIndex, insert.afterColIndex);
      }

      for (const del of this._pendingTableColumnDeletes) {
        xml = this.deleteTableColumnInXml(xml, del.tableIndex, del.colIndex);
      }

      for (const merge of this._pendingCellMerges) {
        xml = this.mergeCellsInXml(xml, merge.tableIndex, merge.startRow, merge.startCol, merge.endRow, merge.endCol);
      }

      this._zip.file(sectionPath, xml);
      sectionIndex++;
    }
    
    this._pendingCellMerges = [];
  }

  private insertTableRowInXml(xml: string, tableIndex: number, afterRowIndex: number, cellTexts?: string[]): string {
    const tableRegex = /<hp:tbl\b[^>]*>[\s\S]*?<\/hp:tbl>/g;
    let currentTableIndex = 0;
    
    return xml.replace(tableRegex, (tableMatch) => {
      if (currentTableIndex !== tableIndex) {
        currentTableIndex++;
        return tableMatch;
      }
      currentTableIndex++;

      const rowRegex = /<hp:tr[^>]*>[\s\S]*?<\/hp:tr>/g;
      const rows: string[] = [];
      let rowMatch;
      while ((rowMatch = rowRegex.exec(tableMatch)) !== null) {
        rows.push(rowMatch[0]);
      }

      if (afterRowIndex >= rows.length) return tableMatch;

      const templateRow = rows[afterRowIndex];
      const newRow = this.createNewRowFromTemplate(templateRow, afterRowIndex + 1, cellTexts);
      rows.splice(afterRowIndex + 1, 0, newRow);

      const newRowCount = rows.length;
      let updatedTable = tableMatch.replace(/rowCnt="(\d+)"/, `rowCnt="${newRowCount}"`);
      
      const rowsStart = updatedTable.indexOf('<hp:tr');
      const rowsEnd = updatedTable.lastIndexOf('</hp:tr>') + '</hp:tr>'.length;
      
      if (rowsStart !== -1 && rowsEnd > rowsStart) {
        updatedTable = updatedTable.substring(0, rowsStart) + rows.join('') + updatedTable.substring(rowsEnd);
      }

      return updatedTable;
    });
  }

  private createNewRowFromTemplate(templateRow: string, newRowAddr: number, cellTexts?: string[]): string {
    let newRow = templateRow;
    
    newRow = newRow.replace(/rowAddr="(\d+)"/g, `rowAddr="${newRowAddr}"`);
    
    let cellIndex = 0;
    newRow = newRow.replace(/<hp:tc\b([^>]*)>([\s\S]*?)<\/hp:tc>/g, (cellMatch, attrs, content) => {
      const newText = cellTexts?.[cellIndex] || '';
      cellIndex++;
      
      const simplifiedContent = content.replace(
        /<hp:subList[^>]*>[\s\S]*?<\/hp:subList>/,
        `<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0"><hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t>${this.escapeXml(newText)}</hp:t></hp:run></hp:p></hp:subList>`
      );
      
      return `<hp:tc${attrs}>${simplifiedContent}</hp:tc>`;
    });

    return newRow;
  }

  private deleteTableRowInXml(xml: string, tableIndex: number, rowIndex: number): string {
    const tableRegex = /<hp:tbl\b[^>]*>[\s\S]*?<\/hp:tbl>/g;
    let currentTableIndex = 0;
    
    return xml.replace(tableRegex, (tableMatch) => {
      if (currentTableIndex !== tableIndex) {
        currentTableIndex++;
        return tableMatch;
      }
      currentTableIndex++;

      const rowRegex = /<hp:tr[^>]*>[\s\S]*?<\/hp:tr>/g;
      const rows: string[] = [];
      let rowMatch;
      while ((rowMatch = rowRegex.exec(tableMatch)) !== null) {
        rows.push(rowMatch[0]);
      }

      if (rowIndex >= rows.length || rows.length <= 1) return tableMatch;

      rows.splice(rowIndex, 1);

      const newRowCount = rows.length;
      let updatedTable = tableMatch.replace(/rowCnt="(\d+)"/, `rowCnt="${newRowCount}"`);
      
      const rowsStart = updatedTable.indexOf('<hp:tr');
      const rowsEnd = updatedTable.lastIndexOf('</hp:tr>') + '</hp:tr>'.length;
      
      if (rowsStart !== -1 && rowsEnd > rowsStart) {
        updatedTable = updatedTable.substring(0, rowsStart) + rows.join('') + updatedTable.substring(rowsEnd);
      }

      return updatedTable;
    });
  }

  private insertTableColumnInXml(xml: string, tableIndex: number, afterColIndex: number): string {
    const tableRegex = /<hp:tbl\b[^>]*>[\s\S]*?<\/hp:tbl>/g;
    let currentTableIndex = 0;
    
    return xml.replace(tableRegex, (tableMatch) => {
      if (currentTableIndex !== tableIndex) {
        currentTableIndex++;
        return tableMatch;
      }
      currentTableIndex++;

      let updatedTable = tableMatch.replace(/colCnt="(\d+)"/, (_match, oldCount) => {
        return `colCnt="${parseInt(oldCount) + 1}"`;
      });

      updatedTable = updatedTable.replace(/<hp:tr[^>]*>[\s\S]*?<\/hp:tr>/g, (rowMatch) => {
        const cellRegex = /<hp:tc\b[^>]*>[\s\S]*?<\/hp:tc>/g;
        const cells: string[] = [];
        let cellMatch;
        while ((cellMatch = cellRegex.exec(rowMatch)) !== null) {
          cells.push(cellMatch[0]);
        }

        if (afterColIndex >= cells.length) return rowMatch;

        const templateCell = cells[afterColIndex];
        const newCell = this.createNewCellFromTemplate(templateCell, afterColIndex + 1);
        cells.splice(afterColIndex + 1, 0, newCell);

        const cellsStart = rowMatch.indexOf('<hp:tc');
        const cellsEnd = rowMatch.lastIndexOf('</hp:tc>') + '</hp:tc>'.length;
        
        if (cellsStart !== -1 && cellsEnd > cellsStart) {
          return rowMatch.substring(0, cellsStart) + cells.join('') + rowMatch.substring(cellsEnd);
        }
        return rowMatch;
      });

      return updatedTable;
    });
  }

  private createNewCellFromTemplate(templateCell: string, newColAddr: number): string {
    let newCell = templateCell;
    
    newCell = newCell.replace(/colAddr="(\d+)"/, `colAddr="${newColAddr}"`);
    newCell = newCell.replace(/<hp:t>([^<]*)<\/hp:t>/g, '<hp:t></hp:t>');

    return newCell;
  }

  private deleteTableColumnInXml(xml: string, tableIndex: number, colIndex: number): string {
    const tableRegex = /<hp:tbl\b[^>]*>[\s\S]*?<\/hp:tbl>/g;
    let currentTableIndex = 0;
    
    return xml.replace(tableRegex, (tableMatch) => {
      if (currentTableIndex !== tableIndex) {
        currentTableIndex++;
        return tableMatch;
      }
      currentTableIndex++;

      let updatedTable = tableMatch.replace(/colCnt="(\d+)"/, (_match, oldCount) => {
        const newCount = parseInt(oldCount) - 1;
        return newCount > 0 ? `colCnt="${newCount}"` : `colCnt="1"`;
      });

      updatedTable = updatedTable.replace(/<hp:tr[^>]*>[\s\S]*?<\/hp:tr>/g, (rowMatch) => {
        const cellRegex = /<hp:tc\b[^>]*>[\s\S]*?<\/hp:tc>/g;
        const cells: string[] = [];
        let cellMatch;
        while ((cellMatch = cellRegex.exec(rowMatch)) !== null) {
          cells.push(cellMatch[0]);
        }

        if (colIndex >= cells.length || cells.length <= 1) return rowMatch;

        cells.splice(colIndex, 1);

        const cellsStart = rowMatch.indexOf('<hp:tc');
        const cellsEnd = rowMatch.lastIndexOf('</hp:tc>') + '</hp:tc>'.length;
        
        if (cellsStart !== -1 && cellsEnd > cellsStart) {
          return rowMatch.substring(0, cellsStart) + cells.join('') + rowMatch.substring(cellsEnd);
        }
        return rowMatch;
      });

      return updatedTable;
    });
  }

  private mergeCellsInXml(xml: string, tableIndex: number, startRow: number, startCol: number, endRow: number, endCol: number): string {
    const tableRegex = /<hp:tbl\b[^>]*>[\s\S]*?<\/hp:tbl>/g;
    let currentTableIndex = 0;
    
    return xml.replace(tableRegex, (tableMatch) => {
      if (currentTableIndex !== tableIndex) {
        currentTableIndex++;
        return tableMatch;
      }
      currentTableIndex++;

      const rowSpan = endRow - startRow + 1;
      const colSpan = endCol - startCol + 1;

      let rowIndex = 0;
      return tableMatch.replace(/<hp:tr[^>]*>([\s\S]*?)<\/hp:tr>/g, (rowMatch, rowContent) => {
        const currentRow = rowIndex;
        rowIndex++;

        if (currentRow < startRow || currentRow > endRow) return rowMatch;

        let colIndex = 0;
        const updatedRowContent = rowContent.replace(/<hp:tc\b([^>]*)>([\s\S]*?)<\/hp:tc>/g, (cellMatch: string, attrs: string, content: string) => {
          const currentCol = colIndex;
          colIndex++;

          if (currentCol < startCol || currentCol > endCol) return cellMatch;

          if (currentRow === startRow && currentCol === startCol) {
            let updatedAttrs = attrs;
            updatedAttrs = updatedAttrs.replace(/rowSpan="(\d+)"/, `rowSpan="${rowSpan}"`);
            updatedAttrs = updatedAttrs.replace(/colSpan="(\d+)"/, `colSpan="${colSpan}"`);
            
            if (!updatedAttrs.includes('rowSpan=')) {
              updatedAttrs += ` rowSpan="${rowSpan}"`;
            }
            if (!updatedAttrs.includes('colSpan=')) {
              updatedAttrs += ` colSpan="${colSpan}"`;
            }
            
            return `<hp:tc${updatedAttrs}>${content}</hp:tc>`;
          } else {
            return '';
          }
        });

        return `<hp:tr>${updatedRowContent}</hp:tr>`;
      });
    });
  }

  private async applyDirectTextUpdatesToXml(): Promise<void> {
    if (!this._zip) return;

    let sectionIndex = 0;
    while (true) {
      const sectionPath = `Contents/section${sectionIndex}.xml`;
      const file = this._zip.file(sectionPath);
      if (!file) break;

      let xml = await file.async('string');

      for (const update of this._pendingDirectTextUpdates) {
        const escapedOld = this.escapeXml(update.oldText);
        const escapedNew = this.escapeXml(update.newText);

        // Replace text anywhere within <hp:t> tags (may contain other tags like <hp:tab/>)
        // First try exact match at the start of <hp:t> content
        const pattern1 = new RegExp(`(<hp:t[^>]*>)${this.escapeRegex(escapedOld)}`, 'g');
        xml = xml.replace(pattern1, `$1${escapedNew}`);

        // Also try simple text replacement for cases where text is standalone
        xml = xml.replace(new RegExp(`>${this.escapeRegex(escapedOld)}<`, 'g'), `>${escapedNew}<`);
      }

      this._zip.file(sectionPath, xml);
      sectionIndex++;
    }
  }

  private escapeRegex(str: string): string {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  /**
   * Apply text replacements directly to XML files.
   * This is the safest approach as it preserves the original XML structure.
   */
  private async applyTextReplacementsToXml(): Promise<void> {
    if (!this._zip) return;

    // Get all section files
    let sectionIndex = 0;
    while (true) {
      const sectionPath = `Contents/section${sectionIndex}.xml`;
      const file = this._zip.file(sectionPath);
      if (!file) break;

      let xml = await file.async('string');

      // Apply each pending replacement to the XML
      for (const replacement of this._pendingTextReplacements) {
        const { oldText, newText, options } = replacement;
        const { caseSensitive = false, regex = false, replaceAll = true } = options;

        // Create pattern for matching text inside <hp:t> tags
        let searchPattern: RegExp;
        if (regex) {
          searchPattern = new RegExp(oldText, caseSensitive ? (replaceAll ? 'g' : '') : (replaceAll ? 'gi' : 'i'));
        } else {
          const escaped = oldText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          searchPattern = new RegExp(escaped, caseSensitive ? (replaceAll ? 'g' : '') : (replaceAll ? 'gi' : 'i'));
        }

        // Replace text within <hp:t> tags while preserving XML structure
        xml = xml.replace(/<hp:t([^>]*)>([^<]*)<\/hp:t>/g, (match, attrs, textContent) => {
          const newTextContent = textContent.replace(searchPattern, this.escapeXml(newText));
          return `<hp:t${attrs}>${newTextContent}</hp:t>`;
        });
      }

      this._zip.file(sectionPath, xml);
      sectionIndex++;
    }

    // Update metadata in header.xml if needed
    await this.syncMetadataToZip();
  }

  /**
   * Sync structural changes (paragraph text, table cells, etc.)
   * Regenerates section XML from _content to handle new elements.
   */
  private async syncStructuralChangesToZip(): Promise<void> {
    if (!this._zip) return;

    // Regenerate each section XML from content
    for (let sectionIndex = 0; sectionIndex < this._content.sections.length; sectionIndex++) {
      const sectionPath = `Contents/section${sectionIndex}.xml`;
      const section = this._content.sections[sectionIndex];
      const newXml = this.generateSectionXml(section);
      this._zip.file(sectionPath, newXml);
    }

    // Sync metadata
    await this.syncMetadataToZip();
  }

  private generateSectionXml(section: HwpxSection): string {
    let xml = `<?xml version="1.0" encoding="UTF-8"?>\n`;
    xml += `<hp:sec xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">\n`;

    for (const element of section.elements) {
      if (element.type === 'paragraph') {
        xml += this.generateParagraphXml(element.data as HwpxParagraph);
      } else if (element.type === 'table') {
        xml += this.generateTableXml(element.data as HwpxTable);
      } else if (element.type === 'textbox') {
        xml += this.generateTextBoxXml(element.data as HwpxTextBox);
      }
    }

    if (section.header) {
      xml += this.generateHeaderFooterXml(section.header, 'header');
    }
    if (section.footer) {
      xml += this.generateHeaderFooterXml(section.footer, 'footer');
    }

    xml += `</hp:sec>`;
    return xml;
  }

  private generateHeaderFooterXml(hf: import('./types').HeaderFooter, type: 'header' | 'footer'): string {
    let xml = `  <hp:${type}>\n`;
    for (const para of hf.paragraphs) {
      xml += this.generateParagraphXml(para, 4);
    }
    xml += `  </hp:${type}>\n`;
    return xml;
  }

  private generateParagraphXml(paragraph: HwpxParagraph, indentSpaces: number = 2): string {
    const indent = ' '.repeat(indentSpaces);
    const align = paragraph.paraStyle?.align || 'left';
    let xml = `${indent}<hp:p`;
    if (align !== 'left') {
      xml += ` align="${align}"`;
    }
    xml += `>\n`;
    
    for (const run of paragraph.runs) {
      xml += `${indent}  <hp:run>\n`;
      if (run.pageNumber) {
        xml += `${indent}    <hp:pageNum/>\n`;
      } else if (run.text) {
        xml += `${indent}    <hp:t>${this.escapeXml(run.text)}</hp:t>\n`;
      }
      xml += `${indent}  </hp:run>\n`;
    }
    xml += `${indent}</hp:p>\n`;
    return xml;
  }

  /**
   * Generate table XML from HwpxTable.
   */
  private generateTableXml(table: HwpxTable): string {
    let xml = `  <hp:tbl rowCount="${table.rowCount}" colCount="${table.colCount}">\n`;

    for (const row of table.rows) {
      xml += `    <hp:tr>\n`;
      for (const cell of row.cells) {
        xml += `      <hp:tc colAddr="${cell.colAddr}" rowAddr="${cell.rowAddr}" colSpan="${cell.colSpan}" rowSpan="${cell.rowSpan}">\n`;
        for (const para of cell.paragraphs) {
          xml += `        <hp:p>\n`;
          for (const run of para.runs) {
            xml += `          <hp:run>\n`;
            xml += `            <hp:t>${this.escapeXml(run.text)}</hp:t>\n`;
            xml += `          </hp:run>\n`;
          }
          xml += `        </hp:p>\n`;
        }
        xml += `      </hp:tc>\n`;
      }
      xml += `    </hp:tr>\n`;
    }

    xml += `  </hp:tbl>\n`;
    return xml;
  }

  private generateTextBoxXml(textBox: HwpxTextBox): string {
    const xHwpunit = Math.round(textBox.x * 100);
    const yHwpunit = Math.round(textBox.y * 100);
    const widthHwpunit = Math.round(textBox.width * 100);
    const heightHwpunit = Math.round(textBox.height * 100);

    let xml = `  <hp:p>\n`;
    xml += `    <hp:run>\n`;
    xml += `      <hp:rect id="${textBox.id}" zOrder="0">\n`;
    xml += `        <hp:sz width="${widthHwpunit}" height="${heightHwpunit}" widthRelTo="ABSOLUTE" heightRelTo="ABSOLUTE"/>\n`;
    xml += `        <hp:pos vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="${yHwpunit}" horzOffset="${xHwpunit}"/>\n`;
    
    if (textBox.fillColor) {
      xml += `        <hp:fillBrush><hp:winBrush faceColor="${textBox.fillColor}"/></hp:fillBrush>\n`;
    }
    if (textBox.strokeColor && textBox.strokeWidth) {
      xml += `        <hp:lineShape color="${textBox.strokeColor}" width="${textBox.strokeWidth}"/>\n`;
    }
    
    xml += `        <hp:textbox>\n`;
    for (const para of textBox.paragraphs) {
      xml += this.generateParagraphXml(para, 10);
    }
    xml += `        </hp:textbox>\n`;
    xml += `      </hp:rect>\n`;
    xml += `    </hp:run>\n`;
    xml += `  </hp:p>\n`;
    
    return xml;
  }

  private updateSectionXml(xml: string, section: HwpxSection): string {
    let updatedXml = xml;

    // Build a map of element index to paragraph data for quick lookup
    const paragraphMap = new Map<number, HwpxParagraph>();
    const tableMap = new Map<number, HwpxTable>();

    let paragraphCount = 0;
    let tableCount = 0;

    for (const element of section.elements) {
      if (element.type === 'paragraph') {
        paragraphMap.set(paragraphCount, element.data as HwpxParagraph);
        paragraphCount++;
      } else if (element.type === 'table') {
        tableMap.set(tableCount, element.data as HwpxTable);
        tableCount++;
      }
    }

    // Track table positions to skip paragraphs inside tables
    const tablePositions: Array<{ start: number; end: number }> = [];
    const tableRegex = /<hp:tbl\b[^>]*>[\s\S]*?<\/hp:tbl>/g;
    let tableMatch;
    while ((tableMatch = tableRegex.exec(xml)) !== null) {
      tablePositions.push({ start: tableMatch.index, end: tableMatch.index + tableMatch[0].length });
    }

    // Update paragraphs outside of tables
    let paragraphIndex = 0;
    const paragraphRegex = /<hp:p\b[^>]*>([\s\S]*?)<\/hp:p>/g;
    updatedXml = xml.replace(paragraphRegex, (match, _inner, offset) => {
      // Check if this paragraph is inside a table
      const isInTable = tablePositions.some(pos => offset >= pos.start && offset < pos.end);

      if (isInTable) {
        return match; // Don't modify paragraphs inside tables here
      }

      const paragraph = paragraphMap.get(paragraphIndex);
      paragraphIndex++;

      if (paragraph) {
        return this.updateParagraphXml(match, paragraph);
      }
      return match;
    });

    // Update table cells
    let tableIndex = 0;
    updatedXml = updatedXml.replace(/<hp:tbl\b[^>]*>([\s\S]*?)<\/hp:tbl>/g, (tblMatch) => {
      const table = tableMap.get(tableIndex);
      tableIndex++;

      if (!table) {
        return tblMatch;
      }

      let rowIndex = 0;
      return tblMatch.replace(/<hp:tr[^>]*>([\s\S]*?)<\/hp:tr>/g, (rowMatch) => {
        if (rowIndex >= table.rows.length) {
          rowIndex++;
          return rowMatch;
        }

        const row = table.rows[rowIndex];
        rowIndex++;

        let cellIndex = 0;
        return rowMatch.replace(/<hp:tc\b([^>]*)>([\s\S]*?)<\/hp:tc>/g, (cellMatch, cellAttrs, cellContent) => {
          if (cellIndex >= row.cells.length) {
            cellIndex++;
            return cellMatch;
          }

          const cell = row.cells[cellIndex];
          cellIndex++;

          // Update cell content - replace text in paragraphs
          let updatedCellContent = cellContent;
          if (cell.paragraphs && cell.paragraphs.length > 0) {
            let cellParaIndex = 0;
            updatedCellContent = cellContent.replace(/<hp:p\b[^>]*>([\s\S]*?)<\/hp:p>/g, (paraMatch: string) => {
              if (cellParaIndex < cell.paragraphs.length) {
                const para = cell.paragraphs[cellParaIndex];
                cellParaIndex++;
                return this.updateParagraphXml(paraMatch, para);
              }
              cellParaIndex++;
              return paraMatch;
            });
          }

          return `<hp:tc${cellAttrs}>${updatedCellContent}</hp:tc>`;
        });
      });
    });

    return updatedXml;
  }

  /**
   * Update paragraph XML with new text content.
   */
  private updateParagraphXml(xml: string, paragraph: HwpxParagraph): string {
    const fullText = paragraph.runs.map(r => r.text).join('');

    // Update all <hp:t> tags with the combined text
    // For simplicity, put all text in the first <hp:t> tag and empty the rest
    let firstTextTag = true;
    return xml.replace(/<hp:t([^>]*)>([^<]*)<\/hp:t>/g, (_match, attrs, _oldText) => {
      if (firstTextTag) {
        firstTextTag = false;
        return `<hp:t${attrs}>${this.escapeXml(fullText)}</hp:t>`;
      }
      // Empty subsequent text tags
      return `<hp:t${attrs}></hp:t>`;
    });
  }

  /**
   * Sync metadata to header.xml
   */
  private async syncMetadataToZip(): Promise<void> {
    if (!this._zip) return;

    const headerPath = 'Contents/header.xml';
    let headerXml = await this._zip.file(headerPath)?.async('string');
    if (headerXml && this._content.metadata) {
      const meta = this._content.metadata;
      if (meta.title) {
        headerXml = headerXml.replace(/<hh:title[^>]*>[^<]*<\/hh:title>/,
          `<hh:title>${this.escapeXml(meta.title)}</hh:title>`);
      }
      if (meta.creator) {
        headerXml = headerXml.replace(/<hh:creator[^>]*>[^<]*<\/hh:creator>/,
          `<hh:creator>${this.escapeXml(meta.creator)}</hh:creator>`);
      }
      if (meta.subject) {
        headerXml = headerXml.replace(/<hh:subject[^>]*>[^<]*<\/hh:subject>/,
          `<hh:subject>${this.escapeXml(meta.subject)}</hh:subject>`);
      }
      if (meta.description) {
        headerXml = headerXml.replace(/<hh:description[^>]*>[^<]*<\/hh:description>/,
          `<hh:description>${this.escapeXml(meta.description)}</hh:description>`);
      }
      this._zip.file(headerPath, headerXml);
    }
  }

  private escapeXml(text: string): string {
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }
}
