import * as vscode from 'vscode';
import * as path from 'path';
import JSZip from 'jszip';
import { HwpxEditorProvider } from './editor/HwpxEditorProvider';
import { HwpEditorProvider } from './editor/HwpEditorProvider';
import { HwpDocument } from './hwp/HwpDocument';
import { HwpxParser } from './hwpx/HwpxParser';

export function activate(context: vscode.ExtensionContext) {
  // Register the custom editors
  context.subscriptions.push(HwpxEditorProvider.register(context));
  context.subscriptions.push(HwpEditorProvider.register(context));

  // MCP Server path
  const mcpServerPath = path.join(context.extensionPath, 'out', 'mcp-server.js');

  // Command: Show MCP Configuration
  context.subscriptions.push(
    vscode.commands.registerCommand('hwpx.showMcpConfig', async () => {
      const config = getMcpConfig(mcpServerPath);

      const panel = vscode.window.createWebviewPanel(
        'hwpxMcpConfig',
        'HWP MCP Server Configuration',
        vscode.ViewColumn.One,
        {}
      );

      panel.webview.html = `<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: var(--vscode-font-family); padding: 20px; }
    h1 { color: var(--vscode-foreground); }
    pre { background: var(--vscode-textBlockQuote-background); padding: 15px; border-radius: 5px; overflow-x: auto; }
    code { font-family: var(--vscode-editor-font-family); }
    .section { margin: 20px 0; }
    .path { background: var(--vscode-textBlockQuote-background); padding: 8px; border-radius: 3px; word-break: break-all; }
  </style>
</head>
<body>
  <h1>HWP MCP Server Configuration</h1>

  <div class="section">
    <h2>MCP Server Path</h2>
    <p class="path"><code>${mcpServerPath}</code></p>
  </div>

  <div class="section">
    <h2>Claude Desktop Configuration</h2>
    <p>Add this to your Claude Desktop config file:</p>
    <ul>
      <li><strong>macOS:</strong> ~/Library/Application Support/Claude/claude_desktop_config.json</li>
      <li><strong>Windows:</strong> %APPDATA%\\Claude\\claude_desktop_config.json</li>
    </ul>
    <pre><code>${JSON.stringify(config, null, 2)}</code></pre>
  </div>

  <div class="section">
    <h2>Available Tools (76)</h2>
    <p>Document: open, close, save, create, list</p>
    <p>Text: paragraphs, search, replace, styles</p>
    <p>Tables: create, edit, rows, columns, CSV</p>
    <p>Objects: images, lines, rectangles, ellipses, equations</p>
    <p>Structure: headers, footers, footnotes, endnotes, sections</p>
    <p>And more: bookmarks, hyperlinks, memos, undo/redo</p>
  </div>
</body>
</html>`;
    })
  );

  // Command: Copy MCP Path
  context.subscriptions.push(
    vscode.commands.registerCommand('hwpx.copyMcpPath', async () => {
      await vscode.env.clipboard.writeText(mcpServerPath);
      vscode.window.showInformationMessage(`MCP Server path copied: ${mcpServerPath}`);
    })
  );

  // Command: Convert HWP to HWPX
  context.subscriptions.push(
    vscode.commands.registerCommand('hwpx.convertHwpToHwpx', async (uri?: vscode.Uri) => {
      let sourceUri = uri;

      if (!sourceUri) {
        const selected = await vscode.window.showOpenDialog({
          canSelectMany: false,
          filters: { 'HWP Files': ['hwp'] },
          title: 'Select HWP file to convert'
        });
        if (!selected || selected.length === 0) return;
        sourceUri = selected[0];
      }

      if (!sourceUri.fsPath.toLowerCase().endsWith('.hwp')) {
        vscode.window.showErrorMessage('Selected file is not an HWP file.');
        return;
      }

      const suggestedPath = sourceUri.fsPath.replace(/\.hwp$/i, '.hwpx');
      const targetUri = await vscode.window.showSaveDialog({
        defaultUri: vscode.Uri.file(suggestedPath),
        filters: { 'HWPX Files': ['hwpx'] },
        title: 'Save converted HWPX file'
      });
      if (!targetUri) return;

      await vscode.window.withProgress(
        { location: vscode.ProgressLocation.Notification, title: 'Converting HWP to HWPX...' },
        async () => {
          try {
            const fileData = await vscode.workspace.fs.readFile(sourceUri!);
            const content = HwpDocument.parseContent(fileData);
            const zip = await HwpxParser.createNewHwpxZip(content);
            const data = await zip.generateAsync({ type: 'uint8array' });
            await vscode.workspace.fs.writeFile(targetUri!, data);
            vscode.window.showInformationMessage(`Converted: ${path.basename(targetUri!.fsPath)}`);
          } catch (err: unknown) {
            const msg = err instanceof Error ? err.message : String(err);
            vscode.window.showErrorMessage(`Conversion failed: ${msg}`);
          }
        }
      );
    })
  );
}

function getMcpConfig(mcpServerPath: string) {
  return {
    mcpServers: {
      hwpx: {
        command: "node",
        args: [mcpServerPath]
      }
    }
  };
}

export function deactivate() {}
