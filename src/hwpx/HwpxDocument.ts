import * as vscode from 'vscode';
import JSZip from 'jszip';
import { HwpxParser } from '../../shared/src/HwpxParser';
import { HwpParser } from '../hwp/HwpParser';
import { parseHwpContent } from '../../shared/src/HwpParser';
import {
  HwpxContent,
  HwpxParagraph,
  TextRun,
  CharacterStyle,
  ParagraphStyle,
  HwpxTable,
  TableCell,
  SectionElement,
} from '../../shared/src/types';

type DocumentFormat = 'hwpx' | 'hwp';

const MAX_UNDO_STACK_SIZE = 30;

export class HwpxDocument implements vscode.CustomDocument {
  private _uri: vscode.Uri;
  private _zip: JSZip | null;
  private _content: HwpxContent;
  private _isDirty = false;
  private _format: DocumentFormat;

  private _undoStack: string[] = [];
  private _redoStack: string[] = [];
  private _onDidChangeContent = new vscode.EventEmitter<void>();
  public readonly onDidChangeContent = this._onDidChangeContent.event;
  private _sentImageKeys: Set<string> = new Set();

  private constructor(uri: vscode.Uri, zip: JSZip | null, content: HwpxContent, format: DocumentFormat) {
    this._uri = uri;
    this._zip = zip;
    this._content = content;
    this._format = format;
  }

  public static async create(uri: vscode.Uri): Promise<HwpxDocument> {
    const fileData = await vscode.workspace.fs.readFile(uri);
    const extension = uri.fsPath.toLowerCase();

    // Detect actual format by magic bytes: ZIP starts with PK (0x504B), OLE with 0xD0CF
    const isZip = fileData.length >= 2 && fileData[0] === 0x50 && fileData[1] === 0x4B;

    if (extension.endsWith('.hwp') && !isZip) {
      const content = parseHwpContent(fileData);
      return new HwpxDocument(uri, null, content, 'hwp');
    } else {
      const zip = await JSZip.loadAsync(fileData);
      const content = await HwpxParser.parse(zip);
      return new HwpxDocument(uri, zip, content, 'hwpx');
    }
  }

  get format(): DocumentFormat {
    return this._format;
  }

  get uri(): vscode.Uri {
    return this._uri;
  }

  async getContent(): Promise<HwpxContent> {
    return this._content;
  }

  getSerializableContent(): object {
    const content = this._content;
    // Only send images whose keys have not been sent before (image caching)
    const newImages = Array.from(content.images.entries()).filter(([key]) => !this._sentImageKeys.has(key));
    newImages.forEach(([key]) => this._sentImageKeys.add(key));
    return {
      metadata: content.metadata,
      sections: content.sections,
      images: newImages,
      footnotes: content.footnotes,
      endnotes: content.endnotes,
      isReadOnly: false,
    };
  }

  resetImageCache(): void {
    this._sentImageKeys.clear();
  }

  private findParagraphByPath(
    sectionIndex: number,
    elementIndex: number
  ): HwpxParagraph | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;

    const element = section.elements[elementIndex];
    if (!element || element.type !== 'paragraph') return null;

    return element.data;
  }

  updateParagraphText(
    sectionIndex: number,
    elementIndex: number,
    runIndex: number,
    text: string
  ): void {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    if (!paragraph) return;

    if (paragraph.runs[runIndex]) {
      this.saveState(sectionIndex);
      paragraph.runs[runIndex].text = text;
      this._isDirty = true;
    }
  }

  updateParagraphRuns(
    sectionIndex: number,
    elementIndex: number,
    runs: TextRun[]
  ): void {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    if (!paragraph) return;

    this.saveState(sectionIndex);
    paragraph.runs = runs;
    this._isDirty = true;
  }

  applyCharacterStyle(
    sectionIndex: number,
    elementIndex: number,
    runIndex: number,
    style: Partial<CharacterStyle>
  ): void {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    if (!paragraph || !paragraph.runs[runIndex]) return;

    this.saveState(sectionIndex);
    const run = paragraph.runs[runIndex];
    run.charStyle = { ...run.charStyle, ...style };
    this._isDirty = true;
  }

  applyParagraphStyle(
    sectionIndex: number,
    elementIndex: number,
    style: Partial<ParagraphStyle>
  ): void {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    if (!paragraph) return;

    this.saveState(sectionIndex);
    paragraph.paraStyle = { ...paragraph.paraStyle, ...style };
    this._isDirty = true;
  }

  setOutlineLevel(
    sectionIndex: number,
    elementIndex: number,
    outlineLevel: number
  ): void {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    if (!paragraph) return;

    this.saveState(sectionIndex);
    if (outlineLevel >= 1 && outlineLevel <= 7) {
      paragraph.outlineLevel = outlineLevel;
    } else {
      delete paragraph.outlineLevel;
    }
    this._isDirty = true;
  }

  setColumnCount(sectionIndex: number, columnCount: number): void {
    const section = this._content.sections[sectionIndex];
    if (!section) return;

    this.saveState(sectionIndex);
    if (!section.columnDef) {
      section.columnDef = {};
    }
    section.columnDef.count = columnCount;
    section.columnDef.sameSize = true;
    this._isDirty = true;
  }

  setCaption(
    sectionIndex: number,
    elementIndex: number,
    caption: string,
    captionPosition: 'above' | 'below' = 'below'
  ): void {
    const section = this._content.sections[sectionIndex];
    if (!section) return;

    const element = section.elements[elementIndex];
    if (!element) return;

    this.saveState(sectionIndex);
    if (element.type === 'table' || element.type === 'image') {
      (element.data as { caption?: string; captionPosition?: string }).caption = caption;
      (element.data as { caption?: string; captionPosition?: string }).captionPosition = captionPosition;
    }
    this._isDirty = true;
  }

  insertParagraph(sectionIndex: number, afterElementIndex: number): void {
    const section = this._content.sections[sectionIndex];
    if (!section) return;

    this.saveState(sectionIndex);
    const newParagraph: HwpxParagraph = {
      id: Math.random().toString(36).substring(2, 11),
      runs: [{ text: '' }],
    };

    const newElement: SectionElement = { type: 'paragraph', data: newParagraph };
    section.elements.splice(afterElementIndex + 1, 0, newElement);
    this._isDirty = true;
  }

  deleteParagraph(sectionIndex: number, elementIndex: number): void {
    const section = this._content.sections[sectionIndex];
    if (!section) return;

    this.saveState(sectionIndex);
    section.elements.splice(elementIndex, 1);
    this._isDirty = true;
  }

  mergeParagraphWithPrevious(sectionIndex: number, elementIndex: number): void {
    const section = this._content.sections[sectionIndex];
    if (!section || elementIndex <= 0) return;

    const currentElement = section.elements[elementIndex];
    const previousElement = section.elements[elementIndex - 1];
    if (!currentElement || !previousElement) return;
    if (currentElement.type !== 'paragraph' || previousElement.type !== 'paragraph') return;

    this.saveState(sectionIndex);
    const currentParagraph = currentElement.data as HwpxParagraph;
    const previousParagraph = previousElement.data as HwpxParagraph;

    previousParagraph.runs.push(...currentParagraph.runs);
    section.elements.splice(elementIndex, 1);
    this._isDirty = true;
  }

  updateTableCell(
    sectionIndex: number,
    elementIndex: number,
    rowIndex: number,
    cellIndex: number,
    paragraphIndex: number,
    text: string
  ): void {
    const section = this._content.sections[sectionIndex];
    if (!section) return;

    const element = section.elements[elementIndex];
    if (!element || element.type !== 'table') return;

    const table = element.data as HwpxTable;
    const cell = table.rows[rowIndex]?.cells[cellIndex];
    if (!cell) return;

    const paragraph = cell.paragraphs[paragraphIndex];
    if (!paragraph) return;

    this.saveState(sectionIndex);
    if (paragraph.runs.length > 0) {
      paragraph.runs[0].text = text;
    } else {
      paragraph.runs = [{ text }];
    }
    this._isDirty = true;
  }

  insertTableRow(sectionIndex: number, elementIndex: number, afterRowIndex: number): void {
    const section = this._content.sections[sectionIndex];
    if (!section) return;

    const element = section.elements[elementIndex];
    if (!element || element.type !== 'table') return;

    const table = element.data as HwpxTable;
    const templateRow = table.rows[afterRowIndex];
    if (!templateRow) return;

    this.saveState(sectionIndex);
    const newRow = {
      cells: templateRow.cells.map(() => ({
        paragraphs: [{
          id: Math.random().toString(36).substring(2, 11),
          runs: [{ text: '' }],
        }],
      })),
    };

    table.rows.splice(afterRowIndex + 1, 0, newRow);
    this._isDirty = true;
  }

  deleteTableRow(sectionIndex: number, elementIndex: number, rowIndex: number): void {
    const section = this._content.sections[sectionIndex];
    if (!section) return;

    const element = section.elements[elementIndex];
    if (!element || element.type !== 'table') return;

    const table = element.data as HwpxTable;
    if (table.rows.length > 1) {
      this.saveState(sectionIndex);
      table.rows.splice(rowIndex, 1);
      this._isDirty = true;
    }
  }

  setTableColumnWidth(sectionIndex: number, elementIndex: number, colIndex: number, width: number): void {
    const section = this._content.sections[sectionIndex];
    if (!section) return;

    const element = section.elements[elementIndex];
    if (!element || element.type !== 'table') return;

    const table = element.data as HwpxTable;
    
    this.saveState(sectionIndex);
    
    if (!table.columnWidths) {
      const colCount = table.colCount || table.colCnt || table.rows[0]?.cells.length || 0;
      table.columnWidths = new Array(colCount).fill(100);
    }
    
    if (colIndex >= 0 && colIndex < table.columnWidths.length) {
      table.columnWidths[colIndex] = width;
    }
    
    for (const row of table.rows) {
      if (row.cells[colIndex]) {
        row.cells[colIndex].width = width;
      }
    }
    
    this._isDirty = true;
  }

  setTableRowHeight(sectionIndex: number, elementIndex: number, rowIndex: number, height: number): void {
    const section = this._content.sections[sectionIndex];
    if (!section) return;

    const element = section.elements[elementIndex];
    if (!element || element.type !== 'table') return;

    const table = element.data as HwpxTable;
    
    if (rowIndex < 0 || rowIndex >= table.rows.length) return;
    
    this.saveState(sectionIndex);
    
    table.rows[rowIndex].height = height;
    
    for (const cell of table.rows[rowIndex].cells) {
      cell.height = height;
    }
    
    this._isDirty = true;
  }

  insertTableColumn(sectionIndex: number, elementIndex: number, colIndex: number, insertLeft: boolean): void {
    const section = this._content.sections[sectionIndex];
    if (!section) return;

    const element = section.elements[elementIndex];
    if (!element || element.type !== 'table') return;

    const table = element.data as HwpxTable;
    
    this.saveState(sectionIndex);
    
    const insertPos = insertLeft ? colIndex : colIndex + 1;
    const defaultWidth = table.columnWidths?.[colIndex] || 100;
    
    if (table.columnWidths) {
      table.columnWidths.splice(insertPos, 0, defaultWidth);
    }
    
    for (const row of table.rows) {
      const newCell: TableCell = {
        paragraphs: [{
          id: Math.random().toString(36).substring(2, 11),
          runs: [{ text: '' }],
        }],
        width: defaultWidth,
      };
      row.cells.splice(insertPos, 0, newCell);
    }
    
    if (table.colCount !== undefined) table.colCount++;
    if (table.colCnt !== undefined) table.colCnt++;
    
    this._isDirty = true;
  }

  deleteTableColumn(sectionIndex: number, elementIndex: number, colIndex: number): void {
    const section = this._content.sections[sectionIndex];
    if (!section) return;

    const element = section.elements[elementIndex];
    if (!element || element.type !== 'table') return;

    const table = element.data as HwpxTable;
    const colCount = table.colCount || table.colCnt || table.rows[0]?.cells.length || 0;
    
    if (colCount <= 1) return;
    
    this.saveState(sectionIndex);
    
    if (table.columnWidths && table.columnWidths.length > colIndex) {
      table.columnWidths.splice(colIndex, 1);
    }
    
    for (const row of table.rows) {
      if (row.cells.length > colIndex) {
        row.cells.splice(colIndex, 1);
      }
    }
    
    if (table.colCount !== undefined) table.colCount--;
    if (table.colCnt !== undefined) table.colCnt--;
    
    this._isDirty = true;
  }

  mergeTableCells(
    sectionIndex: number,
    elementIndex: number,
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number
  ): void {
    const section = this._content.sections[sectionIndex];
    if (!section) return;

    const element = section.elements[elementIndex];
    if (!element || element.type !== 'table') return;

    const table = element.data as HwpxTable;
    
    if (startRow < 0 || endRow >= table.rows.length) return;
    if (startCol < 0 || endCol >= (table.rows[0]?.cells.length || 0)) return;
    
    this.saveState(sectionIndex);
    
    const mainCell = table.rows[startRow].cells[startCol];
    mainCell.rowSpan = endRow - startRow + 1;
    mainCell.colSpan = endCol - startCol + 1;
    
    let combinedText = '';
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        if (r === startRow && c === startCol) continue;
        const cell = table.rows[r]?.cells[c];
        if (cell?.paragraphs) {
          cell.paragraphs.forEach(p => {
            p.runs.forEach(run => {
              if (run.text) combinedText += run.text;
            });
          });
        }
      }
    }
    
    if (combinedText && mainCell.paragraphs[0]?.runs[0]) {
      const existingText = mainCell.paragraphs[0].runs[0].text || '';
      mainCell.paragraphs[0].runs[0].text = existingText + combinedText;
    }
    
    this._isDirty = true;
  }

  makeEdit(_edit: unknown): void {
    this._isDirty = true;
  }

  async save(): Promise<void> {
    if (this._format === 'hwp') {
      await this.saveAsHwp(this._uri);
      return;
    }
    await this.saveAs(this._uri);
  }

  async saveAs(targetUri: vscode.Uri): Promise<void> {
    const targetPath = targetUri.fsPath.toLowerCase();
    if (targetPath.endsWith('.hwp')) {
      await this.saveAsHwp(targetUri);
      return;
    }
    if (this._format === 'hwp' && !targetPath.endsWith('.hwp')) {
      await this.saveAsHwpx(targetUri);
      return;
    }
    
    if (!this._zip) {
      throw new Error('Cannot save: no ZIP archive available');
    }
    
    await HwpxParser.updateZip(this._zip, this._content);
    const data = await this._zip.generateAsync({ type: 'uint8array' });
    await vscode.workspace.fs.writeFile(targetUri, data);
    this._isDirty = false;
  }

  private async saveAsHwp(targetUri: vscode.Uri): Promise<void> {
    const content = this._content;
    let data: Uint8Array;

    // Generate HWP binary data
    try {
      data = HwpParser.write(content);
    } catch (e: any) {
      vscode.window.showErrorMessage(`HWP 저장 실패: ${e.message}`);
      return;
    }

    // Validate: try to re-parse the written data
    try {
      parseHwpContent(data);
    } catch (e: any) {
      vscode.window.showErrorMessage(`HWP 저장 검증 실패 (파일이 깨질 수 있어 저장하지 않았습니다): ${e.message}`);
      return;
    }

    // Backup original if it exists
    try {
      const backupUri = vscode.Uri.file(targetUri.fsPath + '.bak');
      const originalData = await vscode.workspace.fs.readFile(targetUri);
      await vscode.workspace.fs.writeFile(backupUri, originalData);
    } catch {
      // Original file may not exist (new file) - that's OK
    }

    // Write the validated data
    await vscode.workspace.fs.writeFile(targetUri, data);
    this._isDirty = false;
  }

  private async saveAsHwpx(targetUri: vscode.Uri): Promise<void> {
    const newZip = await HwpxParser.createNewHwpxZip(this._content);
    const data = await newZip.generateAsync({ type: 'uint8array' });
    await vscode.workspace.fs.writeFile(targetUri, data);
    
    this._zip = newZip;
    this._format = 'hwpx';
    this._isDirty = false;
  }

  async revert(): Promise<void> {
    const fileData = await vscode.workspace.fs.readFile(this._uri);
    
    if (this._format === 'hwp') {
      this._content = parseHwpContent(fileData);
    } else {
      this._zip = await JSZip.loadAsync(fileData);
      this._content = await HwpxParser.parse(this._zip);
    }
    this._isDirty = false;
  }

  async backup(destination: vscode.Uri): Promise<vscode.CustomDocumentBackup> {
    let backupUri = destination;
    if (this._format === 'hwp') {
      backupUri = vscode.Uri.file(destination.fsPath.replace(/\.hwp$/i, '.hwpx'));
    }
    await this.saveAs(backupUri);
    return {
      id: backupUri.toString(),
      delete: async () => {
        try {
          await vscode.workspace.fs.delete(backupUri);
        } catch {
        }
      },
    };
  }

  private saveState(sectionIndex?: number): void {
    const state = this.serializeContent(sectionIndex);
    this._undoStack.push(state);
    if (this._undoStack.length > MAX_UNDO_STACK_SIZE) {
      this._undoStack.shift();
    }
    this._redoStack = [];
  }

  private serializeContent(sectionIndex?: number): string {
    if (sectionIndex !== undefined && this._content.sections[sectionIndex] !== undefined) {
      return JSON.stringify({
        type: 'partial',
        sectionIndex,
        sectionData: structuredClone(this._content.sections[sectionIndex]),
        metadata: structuredClone(this._content.metadata),
      });
    }
    return JSON.stringify({
      type: 'full',
      sections: structuredClone(this._content.sections),
      metadata: structuredClone(this._content.metadata),
    });
  }

  private deserializeContent(state: string): void {
    const parsed = JSON.parse(state);
    if (parsed.type === 'partial') {
      this._content.sections[parsed.sectionIndex] = parsed.sectionData;
      this._content.metadata = parsed.metadata;
    } else {
      this._content.sections = parsed.sections;
      this._content.metadata = parsed.metadata;
    }
  }

  canUndo(): boolean {
    return this._undoStack.length > 0;
  }

  canRedo(): boolean {
    return this._redoStack.length > 0;
  }

  undo(): boolean {
    if (!this.canUndo()) return false;

    const currentState = this.serializeContent();
    this._redoStack.push(currentState);

    const previousState = this._undoStack.pop()!;
    this.deserializeContent(previousState);
    this._isDirty = true;
    this._onDidChangeContent.fire();
    return true;
  }

  redo(): boolean {
    if (!this.canRedo()) return false;

    const currentState = this.serializeContent();
    this._undoStack.push(currentState);

    const nextState = this._redoStack.pop()!;
    this.deserializeContent(nextState);
    this._isDirty = true;
    this._onDidChangeContent.fire();
    return true;
  }

  dispose(): void {
    this._onDidChangeContent.dispose();
  }
}
