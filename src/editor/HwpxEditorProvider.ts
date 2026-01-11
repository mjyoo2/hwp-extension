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
      const content = document.getSerializableContent();
      webviewPanel.webview.postMessage({
        type: 'update',
        content: content,
      });
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

        case 'insertParagraph':
          document.insertParagraph(message.sectionIndex, message.afterElementIndex);
          fireDocumentChange();
          await updateWebview();
          break;

        case 'deleteParagraph':
          document.deleteParagraph(message.sectionIndex, message.elementIndex);
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
