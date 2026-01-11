import * as vscode from 'vscode';
import * as path from 'path';
import { HwpxEditorProvider } from './editor/HwpxEditorProvider';

export function activate(context: vscode.ExtensionContext) {
  // Register the custom editor
  context.subscriptions.push(HwpxEditorProvider.register(context));

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
