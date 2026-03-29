import * as vscode from 'vscode';
import { HwpxDocument } from '../hwpx/HwpxDocument';
import { getWebviewContent } from './webview';

export class HwpxEditorProvider implements vscode.CustomEditorProvider<HwpxDocument> {
  private static readonly viewType = 'hwpx.editor';

  public static register(context: vscode.ExtensionContext): vscode.Disposable {
    const provider = new HwpxEditorProvider(context);
    return vscode.window.registerCustomEditorProvider(
      HwpxEditorProvider.viewType,
      provider,
      {
        webviewOptions: {
          retainContextWhenHidden: true,
        },
        supportsMultipleEditorsPerDocument: false,
      }
    );
  }

  private readonly _onDidChangeCustomDocument = new vscode.EventEmitter<vscode.CustomDocumentEditEvent<HwpxDocument>>();
  public readonly onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;

  constructor(private readonly context: vscode.ExtensionContext) {}

  async openCustomDocument(
    uri: vscode.Uri,
    _openContext: vscode.CustomDocumentOpenContext,
    _token: vscode.CancellationToken
  ): Promise<HwpxDocument> {
    const document = await HwpxDocument.create(uri);
    return document;
  }

  async resolveCustomEditor(
    document: HwpxDocument,
    webviewPanel: vscode.WebviewPanel,
    _token: vscode.CancellationToken
  ): Promise<void> {
    webviewPanel.webview.options = {
      enableScripts: true,
    };

    webviewPanel.webview.html = getWebviewContent(webviewPanel.webview, this.context.extensionUri);

    const updateWebview = async () => {
      const content = document.getSerializableContent() as any;

      // Always strip inline image data from sections to avoid
      // RangeError: Invalid string length in postMessage's JSON.stringify
      const imageDataMap: Record<string, string> = {};
      let imgCounter = 0;
      const strippedSections = content.sections.map((section: any) => ({
        ...section,
        elements: section.elements.map((el: any) => {
          if (el.type === 'image' && el.data?.data && typeof el.data.data === 'string' && el.data.data.startsWith('data:')) {
            const id = `img_${imgCounter++}`;
            imageDataMap[id] = el.data.data;
            return { ...el, data: { ...el.data, data: `__deferred:${id}` } };
          }
          return el;
        }),
      }));

      // Send structure without inline images
      webviewPanel.webview.postMessage({
        type: 'update',
        content: { ...content, sections: strippedSections },
      });

      // Send images in small batches
      const entries = Object.entries(imageDataMap);
      const BATCH = 3;
      for (let i = 0; i < entries.length; i += BATCH) {
        const batch = Object.fromEntries(entries.slice(i, i + BATCH));
        webviewPanel.webview.postMessage({
          type: 'updateImages',
          imageData: batch,
        });
      }
    };

    const fireDocumentChange = () => {
      this._onDidChangeCustomDocument.fire({
        document,
        undo: async () => {},
        redo: async () => {},
      });
    };

    // Listen for content changes from undo/redo
    const contentChangeDisposable = document.onDidChangeContent(async () => {
      await updateWebview();
    });

    webviewPanel.onDidDispose(() => {
      contentChangeDisposable.dispose();
    });

    webviewPanel.webview.onDidReceiveMessage(async (message) => {
      switch (message.type) {
        case 'ready':
          await updateWebview();
          break;

        case 'undo':
          if (document.undo()) {
            fireDocumentChange();
          }
          break;

        case 'redo':
          if (document.redo()) {
            fireDocumentChange();
          }
          break;

        case 'openUrl':
          vscode.env.openExternal(vscode.Uri.parse(message.url));
          break;

        case 'updateParagraphText':
          document.updateParagraphText(
            message.sectionIndex,
            message.elementIndex,
            message.runIndex,
            message.text
          );
          fireDocumentChange();
          break;

        case 'updateParagraphRuns':
          document.updateParagraphRuns(
            message.sectionIndex,
            message.elementIndex,
            message.runs
          );
          fireDocumentChange();
          break;

        case 'applyCharacterStyle':
          document.applyCharacterStyle(
            message.sectionIndex,
            message.elementIndex,
            message.runIndex,
            message.style
          );
          fireDocumentChange();
          await updateWebview();
          break;

        case 'applyParagraphStyle':
          document.applyParagraphStyle(
            message.sectionIndex,
            message.elementIndex,
            message.style
          );
          fireDocumentChange();
          await updateWebview();
          break;

        case 'setOutlineLevel':
          document.setOutlineLevel(
            message.sectionIndex,
            message.elementIndex,
            message.outlineLevel
          );
          fireDocumentChange();
          await updateWebview();
          break;

        case 'setColumnCount':
          document.setColumnCount(
            message.sectionIndex,
            message.columnCount
          );
          fireDocumentChange();
          await updateWebview();
          break;

        case 'insertParagraph':
          document.insertParagraph(message.sectionIndex, message.afterElementIndex);
          fireDocumentChange();
          await updateWebview();
          break;

        case 'setCaption':
          document.setCaption(
            message.sectionIndex,
            message.elementIndex,
            message.caption,
            message.captionPosition
          );
          fireDocumentChange();
          await updateWebview();
          break;

        case 'deleteParagraph':
          document.deleteParagraph(message.sectionIndex, message.elementIndex);
          fireDocumentChange();
          await updateWebview();
          break;

        case 'mergeParagraphWithPrevious':
          document.mergeParagraphWithPrevious(message.sectionIndex, message.elementIndex);
          fireDocumentChange();
          await updateWebview();
          break;

        case 'updateTableCell':
          document.updateTableCell(
            message.sectionIndex,
            message.elementIndex,
            message.rowIndex,
            message.cellIndex,
            message.paragraphIndex,
            message.text
          );
          fireDocumentChange();
          break;

        case 'insertTableRow':
          document.insertTableRow(
            message.sectionIndex,
            message.elementIndex,
            message.afterRowIndex
          );
          fireDocumentChange();
          await updateWebview();
          break;

        case 'deleteTableRow':
          document.deleteTableRow(
            message.sectionIndex,
            message.elementIndex,
            message.rowIndex
          );
          fireDocumentChange();
          await updateWebview();
          break;

        case 'tableColumnResize':
          document.setTableColumnWidth(
            message.sectionIndex,
            message.tableIndex,
            message.colIndex,
            message.width
          );
          fireDocumentChange();
          break;

        case 'tableRowResize':
          document.setTableRowHeight(
            message.sectionIndex,
            message.tableIndex,
            message.rowIndex,
            message.height
          );
          fireDocumentChange();
          break;

        case 'insertTableColumn':
          document.insertTableColumn(
            message.sectionIndex,
            message.elementIndex,
            message.colIndex,
            message.insertLeft
          );
          fireDocumentChange();
          await updateWebview();
          break;

        case 'deleteTableColumn':
          document.deleteTableColumn(
            message.sectionIndex,
            message.elementIndex,
            message.colIndex
          );
          fireDocumentChange();
          await updateWebview();
          break;

        case 'mergeTableCells':
          document.mergeTableCells(
            message.sectionIndex,
            message.elementIndex,
            message.startRow,
            message.startCol,
            message.endRow,
            message.endCol
          );
          fireDocumentChange();
          await updateWebview();
          break;
      }
    });
  }

  async saveCustomDocument(
    document: HwpxDocument,
    _cancellation: vscode.CancellationToken
  ): Promise<void> {
    await document.save();
  }

  async saveCustomDocumentAs(
    document: HwpxDocument,
    destination: vscode.Uri,
    _cancellation: vscode.CancellationToken
  ): Promise<void> {
    await document.saveAs(destination);
  }

  async revertCustomDocument(
    document: HwpxDocument,
    _cancellation: vscode.CancellationToken
  ): Promise<void> {
    await document.revert();
  }

  async backupCustomDocument(
    document: HwpxDocument,
    context: vscode.CustomDocumentBackupContext,
    _cancellation: vscode.CancellationToken
  ): Promise<vscode.CustomDocumentBackup> {
    return document.backup(context.destination);
  }
}
