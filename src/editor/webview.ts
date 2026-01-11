import * as vscode from 'vscode';

export function getWebviewContent(_webview: vscode.Webview, _extensionUri: vscode.Uri): string {
  return `<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>HWPX Editor</title>
  <style>
    :root {
      --toolbar-height: 80px;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', -apple-system, sans-serif;
      background: var(--vscode-editor-background);
      color: var(--vscode-editor-foreground);
      overflow: hidden;
    }
    .toolbar {
      position: fixed;
      top: 0; left: 0; right: 0;
      height: var(--toolbar-height);
      background: var(--vscode-editor-background);
      border-bottom: 1px solid var(--vscode-panel-border);
      display: flex;
      flex-direction: column;
      padding: 4px 8px;
      z-index: 1000;
      gap: 4px;
    }
    .toolbar-row {
      display: flex;
      align-items: center;
      gap: 4px;
      flex-wrap: wrap;
    }
    .toolbar-group {
      display: flex;
      align-items: center;
      gap: 2px;
      padding: 0 8px;
      border-right: 1px solid var(--vscode-panel-border);
    }
    .toolbar-group:last-child { border-right: none; }
    .toolbar button {
      background: transparent;
      border: 1px solid transparent;
      color: var(--vscode-editor-foreground);
      padding: 4px 8px;
      cursor: pointer;
      border-radius: 3px;
      font-size: 13px;
      min-width: 28px;
      height: 28px;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    .toolbar button:hover {
      background: var(--vscode-toolbar-hoverBackground);
    }
    .toolbar button.active {
      background: var(--vscode-button-background);
      color: var(--vscode-button-foreground);
    }
    .toolbar select, .toolbar input[type="number"] {
      background: var(--vscode-dropdown-background);
      color: var(--vscode-dropdown-foreground);
      border: 1px solid var(--vscode-dropdown-border);
      padding: 4px;
      border-radius: 3px;
      font-size: 12px;
      height: 28px;
    }
    .toolbar input[type="number"] { width: 50px; }
    .toolbar input[type="color"] {
      width: 28px; height: 28px;
      padding: 0;
      border: 1px solid var(--vscode-input-border);
      border-radius: 3px;
      cursor: pointer;
    }
    .editor-container {
      position: fixed;
      top: var(--toolbar-height);
      left: 0; right: 300px; bottom: 24px;
      overflow: auto;
      padding: 20px;
      background: #e0e0e0;
      transition: right 0.3s;
    }
    .editor-container.no-memos {
      right: 0;
    }
    .document {
      max-width: 800px;
      margin: 0 auto;
      background: white;
      box-shadow: 0 2px 8px rgba(0,0,0,0.15);
      padding: 60px;
      min-height: 1000px;
      color: #000;
    }
    .memo-panel {
      position: fixed;
      top: var(--toolbar-height);
      right: 0;
      width: 300px;
      bottom: 24px;
      background: var(--vscode-editor-background);
      border-left: 1px solid var(--vscode-panel-border);
      overflow-y: auto;
      padding: 12px;
    }
    .memo-panel.hidden {
      display: none;
    }
    .memo-panel-header {
      font-size: 14px;
      font-weight: bold;
      padding: 8px 0;
      border-bottom: 1px solid var(--vscode-panel-border);
      margin-bottom: 12px;
      color: var(--vscode-editor-foreground);
    }
    .memo-item {
      background: #e8f5e9;
      border-left: 4px solid #4caf50;
      padding: 10px;
      margin-bottom: 10px;
      border-radius: 4px;
      font-size: 12px;
    }
    .memo-item-header {
      display: flex;
      justify-content: space-between;
      margin-bottom: 6px;
      color: #2e7d32;
      font-size: 11px;
    }
    .memo-item-author {
      font-weight: bold;
    }
    .memo-item-date {
      color: #666;
    }
    .memo-item-content {
      color: #1b5e20;
      line-height: 1.5;
    }
    .memo-item-content p {
      margin: 4px 0;
    }
    .memo-item-target {
      background: #fff9c4;
      padding: 4px 8px;
      margin-bottom: 6px;
      border-radius: 3px;
      font-style: italic;
      color: #5d4037;
      font-size: 11px;
      border-left: 3px solid #ffc107;
    }
    /* Footnotes styles - displayed inline at bottom of content, not as separate page */
    .footnotes-section {
      margin-top: 20px;
      padding-top: 8px;
      border-top: 1px solid #666;
      border-width: 1px;
      width: 40%;
    }
    .footnotes-content {
      font-size: 8pt;
      line-height: 1.3;
      color: #333;
    }
    .footnote-item {
      margin-bottom: 2px;
      padding-left: 14px;
      text-indent: -14px;
    }
    .footnote-number {
      font-size: 7pt;
      vertical-align: super;
      color: #333;
      margin-right: 2px;
    }
    .footnote-text {
      color: #333;
    }
    .footnote-ref {
      cursor: pointer;
      color: #0066cc;
    }
    .footnote-ref:hover {
      text-decoration: underline;
    }
    .section { margin-bottom: 20px; }
    .element { margin-bottom: 8px; }
    .paragraph {
      line-height: 1.8;
      min-height: 1.5em;
      outline: none;
      padding: 2px 4px;
      border-radius: 2px;
      cursor: text;
    }
    .paragraph:hover { background: rgba(0,0,0,0.03); }
    .paragraph:focus {
      background: rgba(0,120,215,0.05);
      outline: 1px solid rgba(0,120,215,0.3);
    }
    .paragraph[contenteditable="true"] { white-space: pre-wrap; }
    .paragraph.align-left { text-align: left; }
    .paragraph.align-center { text-align: center; }
    .paragraph.align-right { text-align: right; }
    .paragraph.align-justify { text-align: justify; }
    .paragraph.align-distribute { text-align: justify; text-justify: distribute-all-lines; }
    .paragraph.align-distributespace { text-align: justify; text-justify: distribute-all-lines; }
    .text-run { display: inline; }
    .text-run.bold { font-weight: bold; }
    .text-run.italic { font-style: italic; }
    .text-run.underline { text-decoration: underline; }
    .text-run.strikethrough { text-decoration: line-through; }
    .text-run.underline.strikethrough { text-decoration: underline line-through; }
    .text-run.has-memo {
      background-color: #c8e6c9;
      border-bottom: 2px solid #4caf50;
      cursor: pointer;
      position: relative;
    }
    .text-run.has-memo:hover {
      background-color: #a5d6a7;
    }
    .text-run.has-memo::after {
      content: '\\1F4DD';
      font-size: 10px;
      position: absolute;
      top: -8px;
      right: -2px;
    }
    .table-container { overflow-x: auto; }
    table.hwpx-table { border-collapse: collapse; background: white; table-layout: fixed; }
    .hwpx-table td, .hwpx-table th {
      border: 1px solid #000;
      vertical-align: middle;
      word-wrap: break-word;
      overflow: hidden;
    }
    .hwpx-table td:focus-within {
      outline: 2px solid var(--vscode-focusBorder);
      outline-offset: -2px;
    }
    .hwpx-table .cell-content { min-height: 1em; outline: none; line-height: 1.4; }
    .page-break { margin: 30px 0; text-align: center; position: relative; }
    .page-break-line { border: none; border-top: 2px dashed #999; margin: 0; }
    .page-break-label {
      position: absolute; top: -10px; left: 50%; transform: translateX(-50%);
      background: white; padding: 0 10px;
      color: #999; font-size: 11px;
    }
    .section-break { margin: 40px 0; text-align: center; position: relative; }
    .section-break-line { border: none; border-top: 3px dashed #666; margin: 0; }
    .section-break-label {
      position: absolute; top: -10px; left: 50%; transform: translateX(-50%);
      background: white; padding: 0 10px;
      color: #666; font-size: 11px; font-weight: bold;
    }
    .image-container { margin: 16px 0; text-align: center; }
    .image-container img { max-width: 100%; height: auto; border: 1px solid #ddd; }
    .list-item { padding-left: 24px; position: relative; }
    .list-item::before { position: absolute; left: 8px; }
    .list-item.bullet::before { content: "•"; }
    .list-item.number::before { content: attr(data-number) "."; }
    .graphic-line { margin: 8px 0; }
    .graphic-rect { margin: 8px auto; box-sizing: border-box; }
    .graphic-ellipse { margin: 8px auto; text-align: center; }
    .graphic-textbox { margin: 8px 0; box-sizing: border-box; min-height: 30px; }
    .textbox-para { min-height: 1.2em; }
    .hwpx-hr { border: none; margin: 16px 0; }
    .metadata {
      font-size: 11px; color: #666;
      margin-bottom: 20px; padding-bottom: 10px;
      border-bottom: 1px solid #ddd;
    }
    .loading {
      display: flex; align-items: center; justify-content: center;
      height: 200px; color: var(--vscode-descriptionForeground);
    }
    .context-menu {
      position: fixed;
      background: var(--vscode-menu-background);
      border: 1px solid var(--vscode-menu-border);
      box-shadow: 0 2px 8px rgba(0,0,0,0.2);
      border-radius: 4px;
      padding: 4px 0;
      z-index: 2000;
      display: none;
    }
    .context-menu.visible { display: block; }
    .context-menu-item {
      padding: 6px 16px;
      cursor: pointer;
      font-size: 13px;
      color: var(--vscode-menu-foreground);
    }
    .context-menu-item:hover {
      background: var(--vscode-menu-selectionBackground);
      color: var(--vscode-menu-selectionForeground);
    }
    .context-menu-separator {
      height: 1px;
      background: var(--vscode-menu-separatorBackground);
      margin: 4px 0;
    }
    .status-bar {
      position: fixed;
      bottom: 0; left: 0; right: 0;
      height: 24px;
      background: var(--vscode-statusBar-background);
      color: var(--vscode-statusBar-foreground);
      font-size: 12px;
      display: flex;
      align-items: center;
      padding: 0 12px;
      border-top: 1px solid var(--vscode-statusBar-border);
    }
    .status-item { margin-right: 16px; }
  </style>
</head>
<body>
  <div class="toolbar">
    <div class="toolbar-row">
      <div class="toolbar-group">
        <select id="fontFamily" title="Font">
          <option value="Malgun Gothic">맑은 고딕</option>
          <option value="Batang">바탕</option>
          <option value="Dotum">돋움</option>
          <option value="Gulim">굴림</option>
          <option value="Gungsuh">궁서</option>
          <option value="Arial">Arial</option>
          <option value="Times New Roman">Times New Roman</option>
        </select>
        <input type="number" id="fontSize" value="10" min="6" max="72" title="Font Size (pt)">
      </div>
      <div class="toolbar-group">
        <button id="boldBtn" title="Bold (Ctrl+B)"><b>B</b></button>
        <button id="italicBtn" title="Italic (Ctrl+I)"><i>I</i></button>
        <button id="underlineBtn" title="Underline (Ctrl+U)"><u>U</u></button>
        <button id="strikeBtn" title="Strikethrough"><s>S</s></button>
      </div>
      <div class="toolbar-group">
        <input type="color" id="textColor" value="#000000" title="Text Color">
        <input type="color" id="bgColor" value="#ffffff" title="Background Color">
      </div>
      <div class="toolbar-group">
        <button id="alignLeft" title="Align Left"><svg width="16" height="16" viewBox="0 0 16 16"><path fill="currentColor" d="M1 2h14v2H1zM1 6h10v2H1zM1 10h14v2H1zM1 14h10v2H1z"/></svg></button>
        <button id="alignCenter" title="Align Center"><svg width="16" height="16" viewBox="0 0 16 16"><path fill="currentColor" d="M1 2h14v2H1zM3 6h10v2H3zM1 10h14v2H1zM3 14h10v2H3z"/></svg></button>
        <button id="alignRight" title="Align Right"><svg width="16" height="16" viewBox="0 0 16 16"><path fill="currentColor" d="M1 2h14v2H1zM5 6h10v2H5zM1 10h14v2H1zM5 14h10v2H5z"/></svg></button>
        <button id="alignJustify" title="Justify"><svg width="16" height="16" viewBox="0 0 16 16"><path fill="currentColor" d="M1 2h14v2H1zM1 6h14v2H1zM1 10h14v2H1zM1 14h14v2H1z"/></svg></button>
      </div>
      <div class="toolbar-group">
        <button id="bulletList" title="Bullet List">&#8226; &#8212;</button>
        <button id="numberList" title="Numbered List">1. &#8212;</button>
      </div>
      <div class="toolbar-group">
        <button id="indentDecrease" title="Decrease Indent"><svg width="16" height="16" viewBox="0 0 16 16"><path fill="currentColor" d="M1 2h14v2H1zM5 6h10v2H5zM5 10h10v2H5zM1 14h14v2H1zM1 6l3 2-3 2z"/></svg></button>
        <button id="indentIncrease" title="Increase Indent"><svg width="16" height="16" viewBox="0 0 16 16"><path fill="currentColor" d="M1 2h14v2H1zM5 6h10v2H5zM5 10h10v2H5zM1 14h14v2H1zM4 6l-3 2 3 2z"/></svg></button>
      </div>
    </div>
    <div class="toolbar-row">
      <div class="toolbar-group">
        <label style="font-size:11px;margin-right:4px;">Line:</label>
        <select id="lineSpacing" title="Line Spacing">
          <option value="100">1.0</option>
          <option value="115">1.15</option>
          <option value="150">1.5</option>
          <option value="160" selected>1.6</option>
          <option value="200">2.0</option>
          <option value="250">2.5</option>
        </select>
      </div>
      <div class="toolbar-group">
        <button id="insertTable" title="Insert Table">&#8862;</button>
        <button id="insertLink" title="Insert Hyperlink">&#128279;</button>
      </div>
      <div class="toolbar-group">
        <button id="superscript" title="Superscript">x&#178;</button>
        <button id="subscript" title="Subscript">x&#8322;</button>
      </div>
    </div>
  </div>
  <div class="editor-container">
    <div class="document" id="document">
      <div class="loading">Loading document...</div>
    </div>
  </div>
  <div class="context-menu" id="contextMenu">
    <div class="context-menu-item" data-action="cut">Cut</div>
    <div class="context-menu-item" data-action="copy">Copy</div>
    <div class="context-menu-item" data-action="paste">Paste</div>
    <div class="context-menu-separator"></div>
    <div class="context-menu-item" data-action="insertParagraph">Insert Paragraph</div>
    <div class="context-menu-item" data-action="deleteParagraph">Delete Paragraph</div>
    <div class="context-menu-separator"></div>
    <div class="context-menu-item" data-action="insertRowBelow">Insert Row Below</div>
    <div class="context-menu-item" data-action="deleteRow">Delete Row</div>
  </div>
  <div class="memo-panel hidden" id="memoPanel">
    <div class="memo-panel-header">메모</div>
    <div id="memoList"></div>
  </div>
  <div class="status-bar">
    <span class="status-item" id="statusWordCount">Words: 0</span>
  </div>
  <script>
    const vscode = acquireVsCodeApi();
    let documentContent = null;
    let selectedElement = null;

    function renderDocument(content) {
      documentContent = content;
      const container = document.getElementById('document');
      if (!content || !content.sections || content.sections.length === 0) {
        container.innerHTML = '<div class="loading">No content found in document</div>';
        return;
      }
      let html = '';
      if (content.metadata) {
        const meta = content.metadata;
        html += '<div class="metadata">';
        if (meta.title) html += '<div><strong>' + escapeHtml(meta.title) + '</strong></div>';
        if (meta.creator) html += '<div>Author: ' + escapeHtml(meta.creator) + '</div>';
        if (meta.modifiedDate) html += '<div>Modified: ' + escapeHtml(meta.modifiedDate) + '</div>';
        html += '</div>';
      }

      // Collect all memos from all sections
      const allMemos = [];
      content.sections.forEach((section, sectionIndex) => {
        // Add section break before each section except the first
        if (sectionIndex > 0) {
          html += '<div class="section-break"><hr class="section-break-line"><span class="section-break-label">섹션 ' + (sectionIndex + 1) + '</span></div>';
        }
        html += '<div class="section" data-section="' + sectionIndex + '">';
        section.elements.forEach((element, elementIndex) => {
          html += renderElement(element, sectionIndex, elementIndex);
        });

        // Footnotes are now rendered per-page in calculateAutoPageBreaks

        html += '</div>';

        // Collect memos
        if (section.memos && section.memos.length > 0) {
          allMemos.push(...section.memos);
        }
      });

      container.innerHTML = html;
      attachEventListeners();
      updateWordCount();

      // Render memo panel
      renderMemoPanel(allMemos);

      // Calculate and insert automatic page breaks after rendering
      setTimeout(() => calculateAutoPageBreaks(content), 100);
    }

    function renderMemoPanel(memos) {
      const memoPanel = document.getElementById('memoPanel');
      const memoList = document.getElementById('memoList');
      const editorContainer = document.querySelector('.editor-container');

      if (!memos || memos.length === 0) {
        memoPanel.classList.add('hidden');
        editorContainer.classList.add('no-memos');
        return;
      }

      memoPanel.classList.remove('hidden');
      editorContainer.classList.remove('no-memos');

      let html = '';
      memos.forEach((memo, index) => {
        html += '<div class="memo-item" data-memo-index="' + index + '" data-memo-id="' + memo.id + '">';
        if (memo.linkedText) {
          html += '<div class="memo-item-target">"' + escapeHtml(memo.linkedText) + '"</div>';
        }
        html += '<div class="memo-item-header">';
        html += '<span class="memo-item-author">' + escapeHtml(memo.author) + '</span>';
        if (memo.date) {
          const dateStr = memo.date.replace('T', ' ').replace('Z', '').substring(0, 16);
          html += '<span class="memo-item-date">' + dateStr + '</span>';
        }
        html += '</div>';
        html += '<div class="memo-item-content">';
        memo.content.forEach(line => {
          html += '<p>' + escapeHtml(line) + '</p>';
        });
        html += '</div>';
        html += '</div>';
      });
      memoList.innerHTML = html;
    }

    function renderFootnotesSection(footnotes) {
      let html = '<div class="footnotes-section">';
      html += '<div class="footnotes-content">';
      footnotes.forEach((fn, idx) => {
        const fnNumber = fn.number || (idx + 1);
        const fnType = fn.type || 'footnote';
        html += '<div class="footnote-item" data-footnote-id="' + fn.id + '">';
        html += '<span class="footnote-number">' + fnNumber + ')</span> ';
        if (fn.paragraphs && fn.paragraphs.length > 0) {
          fn.paragraphs.forEach(p => {
            html += '<span class="footnote-text">';
            if (p.runs) {
              p.runs.forEach(run => {
                if (run.text) html += escapeHtml(run.text);
              });
            }
            html += '</span>';
          });
        }
        html += '</div>';
      });
      html += '</div></div>';
      return html;
    }

    function renderElement(element, sectionIndex, elementIndex) {
      switch (element.type) {
        case 'paragraph': return renderParagraph(element.data, sectionIndex, elementIndex);
        case 'table': return renderTable(element.data, sectionIndex, elementIndex);
        case 'image': return renderImage(element.data, sectionIndex, elementIndex);
        case 'line': return renderLine(element.data, sectionIndex, elementIndex);
        case 'rect': return renderRect(element.data, sectionIndex, elementIndex);
        case 'ellipse': return renderEllipse(element.data, sectionIndex, elementIndex);
        case 'textbox': return renderTextBox(element.data, sectionIndex, elementIndex);
        case 'hr': return renderHorizontalRule(element.data, sectionIndex, elementIndex);
        case 'arc': return renderArc(element.data, sectionIndex, elementIndex);
        case 'polygon': return renderPolygon(element.data, sectionIndex, elementIndex);
        case 'curve': return renderCurve(element.data, sectionIndex, elementIndex);
        case 'connectline': return renderConnectLine(element.data, sectionIndex, elementIndex);
        case 'equation': return renderEquation(element.data, sectionIndex, elementIndex);
        case 'ole': return renderOle(element.data, sectionIndex, elementIndex);
        case 'container': return renderContainer(element.data, sectionIndex, elementIndex);
        case 'textart': return renderTextArt(element.data, sectionIndex, elementIndex);
        case 'unknownobject': return renderUnknownObject(element.data, sectionIndex, elementIndex);
        case 'button': return renderButton(element.data, sectionIndex, elementIndex);
        case 'radiobutton': return renderRadioButton(element.data, sectionIndex, elementIndex);
        case 'checkbutton': return renderCheckButton(element.data, sectionIndex, elementIndex);
        case 'combobox': return renderComboBox(element.data, sectionIndex, elementIndex);
        case 'edit': return renderEdit(element.data, sectionIndex, elementIndex);
        case 'listbox': return renderListBox(element.data, sectionIndex, elementIndex);
        case 'scrollbar': return renderScrollBar(element.data, sectionIndex, elementIndex);
        default: return '';
      }
    }

    function renderParagraph(paragraph, sectionIndex, elementIndex) {
      let html = '';

      // Render page break if present (not on first element of first section)
      if (paragraph.pageBreak && !(sectionIndex === 0 && elementIndex === 0)) {
        html += '<div class="page-break"><hr class="page-break-line"><span class="page-break-label">페이지 나눔</span></div>';
      }

      const style = paragraph.paraStyle || {};
      let classes = ['paragraph', 'element'];
      if (style.align) classes.push('align-' + style.align.toLowerCase());
      if (paragraph.listType && paragraph.listType !== 'none') {
        classes.push('list-item', paragraph.listType);
      }
      let inlineStyle = '';
      if (style.lineSpacing) inlineStyle += 'line-height: ' + (style.lineSpacing / 100) + ';';
      if (style.marginTop) inlineStyle += 'margin-top: ' + style.marginTop + 'pt;';
      if (style.marginBottom) inlineStyle += 'margin-bottom: ' + style.marginBottom + 'pt;';
      // Handle margin and indent properly for hanging indent
      // If firstLineIndent is negative, we need positive marginLeft to compensate
      const marginLeft = style.marginLeft || 0;
      const firstLineIndent = style.firstLineIndent || 0;

      if (firstLineIndent < 0) {
        // Hanging indent: need to ensure effective position is not negative
        // effective first line position = marginLeft + firstLineIndent
        // If this would be negative, adjust both values
        const effectivePos = marginLeft + firstLineIndent;
        if (effectivePos < 0) {
          // The indent would push text off the left edge
          // Apply positive padding and adjusted text-indent
          const adjustedMargin = -firstLineIndent;
          inlineStyle += 'padding-left: ' + adjustedMargin + 'pt;';
          inlineStyle += 'text-indent: ' + firstLineIndent + 'pt;';
        } else {
          if (marginLeft > 0) inlineStyle += 'margin-left: ' + marginLeft + 'pt;';
          inlineStyle += 'text-indent: ' + firstLineIndent + 'pt;';
        }
      } else {
        if (marginLeft > 0) inlineStyle += 'margin-left: ' + marginLeft + 'pt;';
        if (firstLineIndent > 0) inlineStyle += 'text-indent: ' + firstLineIndent + 'pt;';
      }

      html += '<div class="' + classes.join(' ') + '" contenteditable="true" ';
      html += 'data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="paragraph" ';
      if (paragraph.listType === 'number') html += 'data-number="' + (elementIndex + 1) + '" ';
      if (style.keepWithNext) html += 'data-keep-with-next="1" ';
      if (style.keepLines) html += 'data-keep-lines="1" ';
      // Add lineseg layout info (vertpos and height in pt)
      if (paragraph.linesegs && paragraph.linesegs.length > 0) {
        const firstSeg = paragraph.linesegs[0];
        const lastSeg = paragraph.linesegs[paragraph.linesegs.length - 1];
        html += 'data-vertpos="' + firstSeg.vertpos + '" ';
        html += 'data-vertend="' + (lastSeg.vertpos + lastSeg.vertsize) + '" ';
      }
      if (inlineStyle) html += 'style="' + inlineStyle + '" ';
      html += '>';
      paragraph.runs.forEach((run, runIndex) => {
        html += renderTextRun(run, runIndex);
      });
      if (paragraph.runs.length === 0 || paragraph.runs.every(r => !r.text)) html += '<br>';
      html += '</div>';
      return html;
    }

    function renderTextRun(run, runIndex) {
      if (run.tab) {
        const width = Math.max(run.tab.width, 20);
        let tabHtml = '<span class="tab-leader" data-run="' + runIndex + '" style="';
        tabHtml += 'display:inline-block;width:' + width + 'pt;text-align:left;overflow:hidden;vertical-align:baseline;';

        if (run.tab.leader === 'dot') {
          // Dot leader: use letter-spacing to create evenly spaced dots
          tabHtml += 'letter-spacing:3px;white-space:nowrap;';
          tabHtml += '">';
          // Fill with dots - calculate approximate number needed (1pt ≈ 1.33px, each dot+spacing ≈ 6px)
          const numDots = Math.ceil(width / 4);
          tabHtml += '<span style="color:#000;font-size:10px;">' + '.'.repeat(numDots) + '</span>';
          tabHtml += '</span>';
        } else if (run.tab.leader === 'dash') {
          tabHtml += 'letter-spacing:2px;white-space:nowrap;';
          tabHtml += '">';
          const numDashes = Math.ceil(width / 6);
          tabHtml += '<span style="color:#000;font-size:10px;">' + '-'.repeat(numDashes) + '</span>';
          tabHtml += '</span>';
        } else if (run.tab.leader === 'solid') {
          tabHtml += 'border-bottom:1px solid #000;';
          tabHtml += '">&nbsp;</span>';
        } else if (run.tab.leader === 'dashDot' || run.tab.leader === 'dashDotDot') {
          tabHtml += 'border-bottom:1px dashed #000;';
          tabHtml += '">&nbsp;</span>';
        } else {
          // No leader, just space
          tabHtml += '">&nbsp;</span>';
        }
        return tabHtml;
      }
      
      const style = run.charStyle || {};
      let classes = ['text-run'];
      let inlineStyle = '';
      if (style.bold) classes.push('bold');
      if (style.italic) classes.push('italic');
      if (style.underline) classes.push('underline');
      if (style.strikethrough) classes.push('strikethrough');
      if (run.hasMemo) classes.push('has-memo');
      if (style.fontName) inlineStyle += "font-family: '" + style.fontName + "';";
      if (style.fontSize) inlineStyle += 'font-size: ' + style.fontSize + 'pt;';
      if (style.fontColor) inlineStyle += 'color: ' + style.fontColor + ';';
      if (style.backgroundColor && style.backgroundColor !== '#ffffff') {
        inlineStyle += 'background-color: ' + style.backgroundColor + ';';
      }
      if (style.superscript) inlineStyle += 'vertical-align: super; font-size: 0.8em;';
      if (style.subscript) inlineStyle += 'vertical-align: sub; font-size: 0.8em;';

      // Check if this is a footnote/endnote reference
      if (run.footnoteRef || run.endnoteRef) {
        classes.push('footnote-ref');
        inlineStyle += 'vertical-align: super; font-size: 0.7em; color: #0066cc; cursor: pointer;';
      }

      let html = '<span class="' + classes.join(' ') + '" data-run="' + runIndex + '" ';
      if (run.memoId) html += 'data-memo-id="' + run.memoId + '" ';
      if (run.footnoteRef) html += 'data-footnote="' + run.footnoteRef + '" ';
      if (run.endnoteRef) html += 'data-endnote="' + run.endnoteRef + '" ';
      if (inlineStyle) html += 'style="' + inlineStyle + '" ';
      html += '>' + escapeHtml(run.text) + '</span>';
      return html;
    }

    function formatBorder(property, border) {
      if (!border) return '';
      const width = border.width || 0.5;
      const style = border.style || 'solid';
      const color = border.color || '#000000';
      if (style === 'none') return property + ':none;';
      return property + ':' + width + 'pt ' + style + ' ' + color + ';';
    }

    function renderTable(table, sectionIndex, elementIndex) {
      // Apply outMargin for table spacing
      let containerStyle = '';
      if (table.outMargin) {
        containerStyle = 'margin:' + (table.outMargin.top || 0) + 'pt ' +
          (table.outMargin.right || 0) + 'pt ' +
          (table.outMargin.bottom || 0) + 'pt ' +
          (table.outMargin.left || 0) + 'pt;';
      }

      // Apply table position/alignment
      if (table.position) {
        // For non-inline tables (treatAsChar=false), apply horizontal alignment
        if (!table.position.treatAsChar) {
          if (table.position.horzAlign === 'center') {
            containerStyle += 'margin-left:auto;margin-right:auto;';
          } else if (table.position.horzAlign === 'right') {
            containerStyle += 'margin-left:auto;margin-right:0;';
          }
          // Apply horizontal offset if present
          if (table.position.horzOffset) {
            containerStyle += 'margin-left:' + table.position.horzOffset + 'pt;';
          }
        }
      }

      // Add lineseg position attributes for page break detection
      let linesegAttrs = '';
      if (table.linesegs && table.linesegs.length > 0) {
        const firstSeg = table.linesegs[0];
        const lastSeg = table.linesegs[table.linesegs.length - 1];
        linesegAttrs = ' data-vertpos="' + firstSeg.vertpos + '" data-vertend="' + (lastSeg.vertpos + lastSeg.vertsize) + '"';
      }
      let html = '<div class="element table-container" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="table"' +
        linesegAttrs + (containerStyle ? ' style="' + containerStyle + '"' : '') + '>';

      let tableStyle = 'border-collapse: collapse; table-layout: fixed;';
      if (table.width) {
        tableStyle += 'width:' + table.width + 'pt;';
      }
      
      html += '<table class="hwpx-table" style="' + tableStyle + '">';
      
      table.rows.forEach((row, rowIndex) => {
        let rowStyle = '';
        if (row.height) {
          rowStyle = ' style="height:' + row.height + 'pt;"';
        }
        html += '<tr data-row="' + rowIndex + '"' + rowStyle + '>';
        row.cells.forEach((cell, cellIndex) => {
          let cellAttrs = '';
          if (cell.rowSpan && cell.rowSpan > 1) cellAttrs += ' rowspan="' + cell.rowSpan + '"';
          if (cell.colSpan && cell.colSpan > 1) cellAttrs += ' colspan="' + cell.colSpan + '"';
          
          let cellStyle = '';
          cellStyle += formatBorder('border-top', cell.borderTop);
          cellStyle += formatBorder('border-right', cell.borderRight);
          cellStyle += formatBorder('border-bottom', cell.borderBottom);
          cellStyle += formatBorder('border-left', cell.borderLeft);
          if (!cell.borderTop && !cell.borderRight && !cell.borderBottom && !cell.borderLeft) {
            cellStyle += 'border: 1px solid #000;';
          }
          if (cell.width) cellStyle += 'width:' + cell.width + 'pt;';
          if (cell.height) cellStyle += 'height:' + cell.height + 'pt;';
          if (cell.backgroundGradation && cell.backgroundGradation.colors.length >= 2) {
            const angle = (cell.backgroundGradation.angle || 0) + 90;
            const colors = cell.backgroundGradation.colors.join(', ');
            cellStyle += 'background:linear-gradient(' + angle + 'deg, ' + colors + ');';
          } else if (cell.backgroundColor) {
            cellStyle += 'background-color:' + cell.backgroundColor + ';';
          }
          if (cell.verticalAlign) cellStyle += 'vertical-align:' + cell.verticalAlign + ';';

          // If hasMargin is false, use table's inMargin; otherwise use cell's own margin
          let padTop, padRight, padBottom, padLeft;
          if (cell.hasMargin === false && table.inMargin) {
            padTop = table.inMargin.top || 1.41;
            padRight = table.inMargin.right || 1.41;
            padBottom = table.inMargin.bottom || 1.41;
            padLeft = table.inMargin.left || 1.41;
          } else {
            padTop = cell.marginTop !== undefined ? cell.marginTop : 1.41;
            padRight = cell.marginRight !== undefined ? cell.marginRight : 1.41;
            padBottom = cell.marginBottom !== undefined ? cell.marginBottom : 1.41;
            padLeft = cell.marginLeft !== undefined ? cell.marginLeft : 1.41;
          }
          cellStyle += 'padding:' + padTop + 'pt ' + padRight + 'pt ' + padBottom + 'pt ' + padLeft + 'pt;';
          
          html += '<td' + cellAttrs + ' style="' + cellStyle + '" data-cell="' + cellIndex + '">';
          if (cell.elements && cell.elements.length > 0) {
            cell.elements.forEach((el, elIndex) => {
              if (el.type === 'paragraph') {
                const p = el.data;
                let paraStyle = '';
                if (p.paraStyle) {
                  if (p.paraStyle.align) paraStyle += 'text-align:' + p.paraStyle.align.toLowerCase() + ';';
                  if (p.paraStyle.lineSpacing) paraStyle += 'line-height:' + (p.paraStyle.lineSpacing / 100) + ';';
                  if (p.paraStyle.marginTop) paraStyle += 'margin-top:' + p.paraStyle.marginTop + 'pt;';
                  if (p.paraStyle.marginBottom) paraStyle += 'margin-bottom:' + p.paraStyle.marginBottom + 'pt;';
                  // Handle hanging indent (negative firstLineIndent) properly in table cells
                  const marginLeftVal = p.paraStyle.marginLeft || 0;
                  const firstLineIndentVal = p.paraStyle.firstLineIndent || 0;
                  if (firstLineIndentVal < 0) {
                    // For hanging indent, add padding-left and use text-indent
                    const hangAmount = Math.abs(firstLineIndentVal);
                    const extraPad = Math.max(0, hangAmount - marginLeftVal);
                    if (extraPad > 0) paraStyle += 'padding-left:' + extraPad + 'pt;';
                    if (marginLeftVal > 0) paraStyle += 'margin-left:' + marginLeftVal + 'pt;';
                    paraStyle += 'text-indent:' + firstLineIndentVal + 'pt;';
                  } else {
                    if (marginLeftVal > 0) paraStyle += 'margin-left:' + marginLeftVal + 'pt;';
                    if (firstLineIndentVal > 0) paraStyle += 'text-indent:' + firstLineIndentVal + 'pt;';
                  }
                }
                html += '<div class="cell-content" contenteditable="true" data-para="' + elIndex + '"' + (paraStyle ? ' style="' + paraStyle + '"' : '') + '>';
                p.runs.forEach((run, runIndex) => {
                  html += renderTextRun(run, runIndex);
                });
                if (p.runs.length === 0 || p.runs.every(r => !r.text && !r.tab)) html += '<br>';
                html += '</div>';
              } else if (el.type === 'table') {
                html += renderNestedTable(el.data);
              }
            });
          } else {
            cell.paragraphs.forEach((p, pIndex) => {
              let paraStyle = '';
              if (p.paraStyle) {
                if (p.paraStyle.align) paraStyle += 'text-align:' + p.paraStyle.align.toLowerCase() + ';';
                if (p.paraStyle.lineSpacing) paraStyle += 'line-height:' + (p.paraStyle.lineSpacing / 100) + ';';
                if (p.paraStyle.marginTop) paraStyle += 'margin-top:' + p.paraStyle.marginTop + 'pt;';
                if (p.paraStyle.marginBottom) paraStyle += 'margin-bottom:' + p.paraStyle.marginBottom + 'pt;';
                // Handle hanging indent (negative firstLineIndent) properly in table cells
                const marginLeftVal = p.paraStyle.marginLeft || 0;
                const firstLineIndentVal = p.paraStyle.firstLineIndent || 0;
                if (firstLineIndentVal < 0) {
                  const hangAmount = Math.abs(firstLineIndentVal);
                  const extraPad = Math.max(0, hangAmount - marginLeftVal);
                  if (extraPad > 0) paraStyle += 'padding-left:' + extraPad + 'pt;';
                  if (marginLeftVal > 0) paraStyle += 'margin-left:' + marginLeftVal + 'pt;';
                  paraStyle += 'text-indent:' + firstLineIndentVal + 'pt;';
                } else {
                  if (marginLeftVal > 0) paraStyle += 'margin-left:' + marginLeftVal + 'pt;';
                  if (firstLineIndentVal > 0) paraStyle += 'text-indent:' + firstLineIndentVal + 'pt;';
                }
              }
              html += '<div class="cell-content" contenteditable="true" data-para="' + pIndex + '"' + (paraStyle ? ' style="' + paraStyle + '"' : '') + '>';
              p.runs.forEach((run, runIndex) => {
                html += renderTextRun(run, runIndex);
              });
              if (p.runs.length === 0 || p.runs.every(r => !r.text && !r.tab)) html += '<br>';
              html += '</div>';
            });
          }
          html += '</td>';
        });
        html += '</tr>';
      });
      html += '</table></div>';
      return html;
    }

    function renderNestedTable(table) {
      let tableStyle = 'border-collapse: collapse; table-layout: fixed; margin: 4px 0;';
      if (table.width) tableStyle += 'width:' + table.width + 'pt;';
      
      let html = '<table class="hwpx-table nested-table" style="' + tableStyle + '">';
      
      table.rows.forEach((row, rowIndex) => {
        let rowStyle = '';
        if (row.height) rowStyle = ' style="height:' + row.height + 'pt;"';
        html += '<tr' + rowStyle + '>';
        row.cells.forEach((cell, cellIndex) => {
          let cellAttrs = '';
          if (cell.rowSpan && cell.rowSpan > 1) cellAttrs += ' rowspan="' + cell.rowSpan + '"';
          if (cell.colSpan && cell.colSpan > 1) cellAttrs += ' colspan="' + cell.colSpan + '"';
          
          let cellStyle = '';
          cellStyle += formatBorder('border-top', cell.borderTop);
          cellStyle += formatBorder('border-right', cell.borderRight);
          cellStyle += formatBorder('border-bottom', cell.borderBottom);
          cellStyle += formatBorder('border-left', cell.borderLeft);
          if (!cell.borderTop && !cell.borderRight && !cell.borderBottom && !cell.borderLeft) {
            cellStyle += 'border: 1px solid #000;';
          }
          if (cell.width) cellStyle += 'width:' + cell.width + 'pt;';
          if (cell.height) cellStyle += 'height:' + cell.height + 'pt;';
          if (cell.backgroundGradation && cell.backgroundGradation.colors.length >= 2) {
            const angle = (cell.backgroundGradation.angle || 0) + 90;
            const colors = cell.backgroundGradation.colors.join(', ');
            cellStyle += 'background:linear-gradient(' + angle + 'deg, ' + colors + ');';
          } else if (cell.backgroundColor) {
            cellStyle += 'background-color:' + cell.backgroundColor + ';';
          }
          if (cell.verticalAlign) cellStyle += 'vertical-align:' + cell.verticalAlign + ';';

          // If hasMargin is false, use table's inMargin; otherwise use cell's own margin
          let padTop, padRight, padBottom, padLeft;
          if (cell.hasMargin === false && table.inMargin) {
            padTop = table.inMargin.top || 1.41;
            padRight = table.inMargin.right || 1.41;
            padBottom = table.inMargin.bottom || 1.41;
            padLeft = table.inMargin.left || 1.41;
          } else {
            padTop = cell.marginTop !== undefined ? cell.marginTop : 1.41;
            padRight = cell.marginRight !== undefined ? cell.marginRight : 1.41;
            padBottom = cell.marginBottom !== undefined ? cell.marginBottom : 1.41;
            padLeft = cell.marginLeft !== undefined ? cell.marginLeft : 1.41;
          }
          cellStyle += 'padding:' + padTop + 'pt ' + padRight + 'pt ' + padBottom + 'pt ' + padLeft + 'pt;';
          
          html += '<td' + cellAttrs + ' style="' + cellStyle + '">';
          
          if (cell.nestedTables && cell.nestedTables.length > 0) {
            cell.nestedTables.forEach(nestedTbl => {
              html += renderNestedTable(nestedTbl);
            });
          }
          
          cell.paragraphs.forEach((p, pIndex) => {
            let paraStyle = '';
            if (p.paraStyle) {
              if (p.paraStyle.align) paraStyle += 'text-align:' + p.paraStyle.align.toLowerCase() + ';';
              if (p.paraStyle.lineSpacing) paraStyle += 'line-height:' + (p.paraStyle.lineSpacing / 100) + ';';
              if (p.paraStyle.marginTop) paraStyle += 'margin-top:' + p.paraStyle.marginTop + 'pt;';
              if (p.paraStyle.marginBottom) paraStyle += 'margin-bottom:' + p.paraStyle.marginBottom + 'pt;';
              // Handle hanging indent (negative firstLineIndent) properly in nested table cells
              const marginLeftVal = p.paraStyle.marginLeft || 0;
              const firstLineIndentVal = p.paraStyle.firstLineIndent || 0;
              if (firstLineIndentVal < 0) {
                const hangAmount = Math.abs(firstLineIndentVal);
                const extraPad = Math.max(0, hangAmount - marginLeftVal);
                if (extraPad > 0) paraStyle += 'padding-left:' + extraPad + 'pt;';
                if (marginLeftVal > 0) paraStyle += 'margin-left:' + marginLeftVal + 'pt;';
                paraStyle += 'text-indent:' + firstLineIndentVal + 'pt;';
              } else {
                if (marginLeftVal > 0) paraStyle += 'margin-left:' + marginLeftVal + 'pt;';
                if (firstLineIndentVal > 0) paraStyle += 'text-indent:' + firstLineIndentVal + 'pt;';
              }
            }
            html += '<div class="cell-content"' + (paraStyle ? ' style="' + paraStyle + '"' : '') + '>';
            p.runs.forEach((run, runIndex) => {
              html += renderTextRun(run, runIndex);
            });
            if (p.runs.length === 0 || p.runs.every(r => !r.text && !r.tab)) html += '<br>';
            html += '</div>';
          });
          html += '</td>';
        });
        html += '</tr>';
      });
      html += '</table>';
      return html;
    }

    function renderImage(image, sectionIndex, elementIndex) {
      let html = '<div class="element image-container" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="image">';
      if (image.data) {
        html += '<img src="' + image.data + '" ';
        if (image.width) html += 'width="' + image.width + '" ';
        if (image.height) html += 'height="' + image.height + '" ';
        html += '/>';
      } else {
        html += '<div style="padding:20px;background:#f0f0f0;color:#666;">[Image: ' + escapeHtml(image.binaryId) + ']</div>';
      }
      html += '</div>';
      return html;
    }

    function renderLine(line, sectionIndex, elementIndex) {
      const width = Math.abs(line.x2 - line.x1) || 1;
      const height = Math.abs(line.y2 - line.y1) || 2;
      const isHorizontal = height <= 2;
      
      let html = '<div class="element graphic-line" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="line">';
      html += '<svg width="100%" height="' + Math.max(height, 10) + '" style="display:block;">';
      html += '<line x1="0" y1="' + (height/2) + '" x2="100%" y2="' + (height/2) + '" ';
      html += 'stroke="' + (line.strokeColor || '#000') + '" ';
      html += 'stroke-width="' + (line.strokeWidth || 1) + '" ';
      if (line.strokeStyle === 'dashed') html += 'stroke-dasharray="8,4" ';
      if (line.strokeStyle === 'dotted') html += 'stroke-dasharray="2,2" ';
      html += '/></svg></div>';
      return html;
    }

    function renderRect(rect, sectionIndex, elementIndex) {
      let html = '<div class="element graphic-rect" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="rect" ';
      html += 'style="';
      html += 'width:' + (rect.width || 100) + 'pt;';
      html += 'height:' + (rect.height || 50) + 'pt;';
      if (rect.fillColor) html += 'background-color:' + rect.fillColor + ';';
      html += 'border:' + (rect.strokeWidth || 1) + 'px solid ' + (rect.strokeColor || '#000') + ';';
      if (rect.cornerRadius) html += 'border-radius:' + rect.cornerRadius + 'pt;';
      html += '">';
      html += '</div>';
      return html;
    }

    function renderEllipse(ellipse, sectionIndex, elementIndex) {
      const width = (ellipse.rx || 50) * 2;
      const height = (ellipse.ry || 50) * 2;
      
      let html = '<div class="element graphic-ellipse" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="ellipse">';
      html += '<svg width="' + width + 'pt" height="' + height + 'pt" style="display:block;">';
      html += '<ellipse cx="50%" cy="50%" rx="' + (ellipse.rx || 50) + '" ry="' + (ellipse.ry || 50) + '" ';
      html += 'fill="' + (ellipse.fillColor || 'transparent') + '" ';
      html += 'stroke="' + (ellipse.strokeColor || '#000') + '" ';
      html += 'stroke-width="' + (ellipse.strokeWidth || 1) + '" ';
      html += '/></svg></div>';
      return html;
    }

    function renderTextBox(textbox, sectionIndex, elementIndex) {
      let html = '<div class="element graphic-textbox" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="textbox" ';
      html += 'style="';
      html += 'width:' + (textbox.width || 200) + 'pt;';
      if (textbox.fillColor) html += 'background-color:' + textbox.fillColor + ';';
      if (textbox.strokeColor) html += 'border:' + (textbox.strokeWidth || 1) + 'px solid ' + textbox.strokeColor + ';';
      html += 'padding:8px;';
      html += '">';
      if (textbox.paragraphs) {
        textbox.paragraphs.forEach(function(para, pIndex) {
          html += '<div class="textbox-para">';
          para.runs.forEach(function(run) {
            html += escapeHtml(run.text);
          });
          html += '</div>';
        });
      }
      html += '</div>';
      return html;
    }

    function renderHorizontalRule(hr, sectionIndex, elementIndex) {
      let style = 'border:none;';
      style += 'border-top:' + (hr.height || 1) + 'px ';
      style += (hr.style || 'solid') + ' ';
      style += (hr.color || '#000') + ';';
      style += 'margin:16px 0;';
      if (hr.width !== 'full') {
        style += 'width:' + hr.width + 'pt;';
        if (hr.align === 'center') style += 'margin-left:auto;margin-right:auto;';
        else if (hr.align === 'right') style += 'margin-left:auto;margin-right:0;';
      }
      
      return '<hr class="element hwpx-hr" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="hr" style="' + style + '">';
    }

    function renderArc(arc, sectionIndex, elementIndex) {
      const cx = arc.centerX || 50;
      const cy = arc.centerY || 50;
      const rx = Math.abs(arc.axis1X || 50);
      const ry = Math.abs(arc.axis2Y || 50);
      const width = rx * 2 + 20;
      const height = ry * 2 + 20;
      
      let html = '<div class="element graphic-arc" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="arc">';
      html += '<svg width="' + width + 'pt" height="' + height + 'pt" viewBox="0 0 ' + width + ' ' + height + '">';
      
      const arcType = arc.type || 'Normal';
      if (arcType === 'Pie') {
        html += '<path d="M ' + (width/2) + ' ' + (height/2) + ' L ' + (width/2 + rx) + ' ' + (height/2) + ' A ' + rx + ' ' + ry + ' 0 0 1 ' + (width/2) + ' ' + (height/2 - ry) + ' Z" ';
      } else {
        html += '<ellipse cx="' + (width/2) + '" cy="' + (height/2) + '" rx="' + rx + '" ry="' + ry + '" ';
      }
      html += 'fill="none" stroke="#000" stroke-width="1"/>';
      html += '</svg></div>';
      return html;
    }

    function renderPolygon(polygon, sectionIndex, elementIndex) {
      if (!polygon.points || polygon.points.length < 2) {
        return '<div class="element graphic-polygon" data-section="' + sectionIndex + '" data-element="' + elementIndex + '">[Polygon]</div>';
      }
      
      let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
      polygon.points.forEach(function(p) {
        if (p.x < minX) minX = p.x;
        if (p.y < minY) minY = p.y;
        if (p.x > maxX) maxX = p.x;
        if (p.y > maxY) maxY = p.y;
      });
      
      const width = (maxX - minX) + 20;
      const height = (maxY - minY) + 20;
      const points = polygon.points.map(function(p) { return (p.x - minX + 10) + ',' + (p.y - minY + 10); }).join(' ');
      
      let html = '<div class="element graphic-polygon" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="polygon">';
      html += '<svg width="' + width + 'pt" height="' + height + 'pt">';
      html += '<polygon points="' + points + '" fill="none" stroke="#000" stroke-width="1"/>';
      html += '</svg></div>';
      return html;
    }

    function renderCurve(curve, sectionIndex, elementIndex) {
      if (!curve.segments || curve.segments.length === 0) {
        return '<div class="element graphic-curve" data-section="' + sectionIndex + '" data-element="' + elementIndex + '">[Curve]</div>';
      }
      
      let minX = 0, minY = 0, maxX = 100, maxY = 100;
      curve.segments.forEach(function(s) {
        if (s.x1 > maxX) maxX = s.x1;
        if (s.y1 > maxY) maxY = s.y1;
        if (s.x2 > maxX) maxX = s.x2;
        if (s.y2 > maxY) maxY = s.y2;
      });
      
      const width = maxX + 20;
      const height = maxY + 20;
      
      let pathD = 'M 10 10';
      curve.segments.forEach(function(s) {
        if (s.type === 'Curve') {
          pathD += ' Q ' + (s.x1 + 10) + ' ' + (s.y1 + 10) + ' ' + (s.x2 + 10) + ' ' + (s.y2 + 10);
        } else {
          pathD += ' L ' + (s.x2 + 10) + ' ' + (s.y2 + 10);
        }
      });
      
      let html = '<div class="element graphic-curve" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="curve">';
      html += '<svg width="' + width + 'pt" height="' + height + 'pt">';
      html += '<path d="' + pathD + '" fill="none" stroke="#000" stroke-width="1"/>';
      html += '</svg></div>';
      return html;
    }

    function renderConnectLine(connectLine, sectionIndex, elementIndex) {
      const x1 = connectLine.startX || 0;
      const y1 = connectLine.startY || 0;
      const x2 = connectLine.endX || 100;
      const y2 = connectLine.endY || 0;
      const width = Math.max(Math.abs(x2 - x1), 10) + 20;
      const height = Math.max(Math.abs(y2 - y1), 10) + 20;
      
      let html = '<div class="element graphic-connectline" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="connectline">';
      html += '<svg width="' + width + 'pt" height="' + height + 'pt">';
      html += '<line x1="10" y1="10" x2="' + (width - 10) + '" y2="' + (height - 10) + '" stroke="#000" stroke-width="1" marker-end="url(#arrowhead)"/>';
      html += '<defs><marker id="arrowhead" markerWidth="10" markerHeight="7" refX="9" refY="3.5" orient="auto"><polygon points="0 0, 10 3.5, 0 7" fill="#000"/></marker></defs>';
      html += '</svg></div>';
      return html;
    }

    function renderEquation(equation, sectionIndex, elementIndex) {
      const script = equation.script || '';
      let html = '<div class="element graphic-equation" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="equation" ';
      html += 'style="padding:8px;margin:8px 0;background:#f8f8f8;border:1px solid #ddd;border-radius:4px;font-family:serif;font-style:italic;">';
      if (script) {
        html += '<span style="color:#333;">' + escapeHtml(script) + '</span>';
      } else {
        html += '<span style="color:#999;">[Equation]</span>';
      }
      html += '</div>';
      return html;
    }

    function renderOle(ole, sectionIndex, elementIndex) {
      const objType = ole.objectType || 'Unknown';
      let html = '<div class="element graphic-ole" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="ole" ';
      html += 'style="padding:16px;margin:8px 0;background:#f0f0f0;border:1px solid #ccc;text-align:center;color:#666;">';
      html += '<div style="font-size:24px;margin-bottom:8px;">📎</div>';
      html += '<div>[OLE Object: ' + escapeHtml(objType) + ']</div>';
      if (ole.binItem) {
        html += '<div style="font-size:11px;color:#999;">Ref: ' + escapeHtml(ole.binItem) + '</div>';
      }
      html += '</div>';
      return html;
    }

    function renderContainer(container, sectionIndex, elementIndex) {
      let html = '<div class="element graphic-container" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="container" ';
      html += 'style="border:1px dashed #999;padding:8px;margin:8px 0;">';
      if (container.children && container.children.length > 0) {
        container.children.forEach(function(child, i) {
          html += renderContainerChild(child, i);
        });
      } else {
        html += '<span style="color:#999;">[Container Group]</span>';
      }
      html += '</div>';
      return html;
    }

    function renderContainerChild(child, index) {
      if (child.points) {
        return renderPolygon(child, 0, index);
      } else if (child.segments) {
        return renderCurve(child, 0, index);
      } else if (child.centerX !== undefined) {
        return renderArc(child, 0, index);
      } else if (child.binaryId) {
        return renderImage(child, 0, index);
      } else if (child.children) {
        return renderContainer(child, 0, index);
      }
      return '<div>[Child Object]</div>';
    }

    function renderTextArt(textArt, sectionIndex, elementIndex) {
      const text = textArt.text || '';
      let html = '<div class="element graphic-textart" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="textart" ';
      html += 'style="padding:16px;margin:8px 0;text-align:center;">';
      html += '<span style="font-size:24px;font-weight:bold;';
      if (textArt.shape && textArt.shape.fontName) {
        html += 'font-family:\\'' + textArt.shape.fontName + '\\';';
      }
      html += 'background:linear-gradient(45deg,#667eea,#764ba2);-webkit-background-clip:text;-webkit-text-fill-color:transparent;">';
      html += escapeHtml(text) || '[TextArt]';
      html += '</span></div>';
      return html;
    }

    function renderUnknownObject(obj, sectionIndex, elementIndex) {
      let html = '<div class="element graphic-unknown" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="unknownobject" ';
      html += 'style="padding:8px;margin:8px 0;background:#fffbe6;border:1px solid #ffe58f;color:#ad8b00;text-align:center;">';
      html += '<div>[Unknown Object';
      if (obj.ctrlId) html += ': ' + escapeHtml(obj.ctrlId);
      html += ']</div></div>';
      return html;
    }

    function renderButton(button, sectionIndex, elementIndex) {
      const form = button.formObject || {};
      const btnSet = form.buttonSet || {};
      const caption = btnSet.caption || 'Button';
      
      let html = '<div class="element form-button" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="button">';
      html += '<button style="padding:4px 12px;cursor:pointer;">' + escapeHtml(caption) + '</button>';
      html += '</div>';
      return html;
    }

    function renderRadioButton(radio, sectionIndex, elementIndex) {
      const form = radio.formObject || {};
      const btnSet = form.buttonSet || {};
      const caption = btnSet.caption || 'Option';
      const groupName = btnSet.radioGroupName || form.groupName || 'radio';
      
      let html = '<div class="element form-radio" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="radiobutton">';
      html += '<label style="cursor:pointer;"><input type="radio" name="' + escapeHtml(groupName) + '"> ' + escapeHtml(caption) + '</label>';
      html += '</div>';
      return html;
    }

    function renderCheckButton(check, sectionIndex, elementIndex) {
      const form = check.formObject || {};
      const btnSet = form.buttonSet || {};
      const caption = btnSet.caption || 'Checkbox';
      const checked = btnSet.value === '1' || btnSet.value === 'true';
      
      let html = '<div class="element form-check" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="checkbutton">';
      html += '<label style="cursor:pointer;"><input type="checkbox"' + (checked ? ' checked' : '') + '> ' + escapeHtml(caption) + '</label>';
      html += '</div>';
      return html;
    }

    function renderComboBox(combo, sectionIndex, elementIndex) {
      const text = combo.text || '';
      
      let html = '<div class="element form-combo" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="combobox">';
      html += '<select style="padding:4px;min-width:100px;">';
      html += '<option>' + escapeHtml(text) + '</option>';
      html += '</select>';
      html += '</div>';
      return html;
    }

    function renderEdit(edit, sectionIndex, elementIndex) {
      const text = edit.text || '';
      const multiLine = edit.multiLine;
      const readOnly = edit.readOnly;
      
      let html = '<div class="element form-edit" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="edit">';
      if (multiLine) {
        html += '<textarea style="padding:4px;min-width:200px;min-height:60px;"' + (readOnly ? ' readonly' : '') + '>' + escapeHtml(text) + '</textarea>';
      } else {
        html += '<input type="text" value="' + escapeHtml(text) + '" style="padding:4px;min-width:150px;"' + (readOnly ? ' readonly' : '') + '>';
      }
      html += '</div>';
      return html;
    }

    function renderListBox(listBox, sectionIndex, elementIndex) {
      const text = listBox.text || '';
      
      let html = '<div class="element form-listbox" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="listbox">';
      html += '<select size="3" style="padding:4px;min-width:100px;">';
      html += '<option>' + escapeHtml(text) + '</option>';
      html += '</select>';
      html += '</div>';
      return html;
    }

    function renderScrollBar(scrollBar, sectionIndex, elementIndex) {
      const min = scrollBar.min || 0;
      const max = scrollBar.max || 100;
      const value = scrollBar.value || 0;
      const isVertical = scrollBar.type === 'vertical' || scrollBar.type === 'Vertical';
      
      let html = '<div class="element form-scrollbar" data-section="' + sectionIndex + '" data-element="' + elementIndex + '" data-type="scrollbar">';
      html += '<input type="range" min="' + min + '" max="' + max + '" value="' + value + '"';
      if (isVertical) {
        html += ' style="writing-mode:bt-lr;-webkit-appearance:slider-vertical;height:100px;"';
      } else {
        html += ' style="min-width:100px;"';
      }
      html += '>';
      html += '</div>';
      return html;
    }

    function attachEventListeners() {
      document.querySelectorAll('.paragraph').forEach(el => {
        el.addEventListener('blur', handleParagraphBlur);
        el.addEventListener('focus', handleElementFocus);
        el.addEventListener('keydown', handleKeyDown);
      });
      document.querySelectorAll('.cell-content').forEach(el => {
        el.addEventListener('blur', handleCellBlur);
        el.addEventListener('focus', handleElementFocus);
      });
      document.querySelectorAll('.element').forEach(el => {
        el.addEventListener('contextmenu', handleContextMenu);
      });
    }

    function handleElementFocus(e) {
      selectedElement = e.target;
      updateToolbarState();
    }

    function handleParagraphBlur(e) {
      const el = e.target;
      const sectionIndex = parseInt(el.dataset.section, 10);
      const elementIndex = parseInt(el.dataset.element, 10);
      const runs = extractRunsFromElement(el);
      vscode.postMessage({ type: 'updateParagraphRuns', sectionIndex, elementIndex, runs });
    }

    function handleCellBlur(e) {
      const el = e.target;
      const td = el.closest('td');
      const tr = td.closest('tr');
      const container = el.closest('.table-container');
      vscode.postMessage({
        type: 'updateTableCell',
        sectionIndex: parseInt(container.dataset.section, 10),
        elementIndex: parseInt(container.dataset.element, 10),
        rowIndex: parseInt(tr.dataset.row, 10),
        cellIndex: parseInt(td.dataset.cell, 10),
        paragraphIndex: parseInt(el.dataset.para, 10),
        text: el.textContent
      });
    }

    function extractRunsFromElement(el) {
      const runs = [];
      const spans = el.querySelectorAll('.text-run');
      if (spans.length > 0) {
        spans.forEach(span => {
          runs.push({ text: span.textContent, charStyle: extractStyleFromSpan(span) });
        });
      } else {
        runs.push({ text: el.textContent });
      }
      return runs;
    }

    function extractStyleFromSpan(span) {
      const style = {};
      if (span.classList.contains('bold')) style.bold = true;
      if (span.classList.contains('italic')) style.italic = true;
      if (span.classList.contains('underline')) style.underline = true;
      if (span.classList.contains('strikethrough')) style.strikethrough = true;
      return style;
    }

    function handleKeyDown(e) {
      if (e.ctrlKey || e.metaKey) {
        switch (e.key.toLowerCase()) {
          case 'b': e.preventDefault(); applyCharStyle('bold'); break;
          case 'i': e.preventDefault(); applyCharStyle('italic'); break;
          case 'u': e.preventDefault(); applyCharStyle('underline'); break;
          case 'z':
            e.preventDefault();
            if (e.shiftKey) {
              vscode.postMessage({ type: 'redo' });
            } else {
              vscode.postMessage({ type: 'undo' });
            }
            break;
          case 'y':
            e.preventDefault();
            vscode.postMessage({ type: 'redo' });
            break;
        }
      }
      if (e.key === 'Enter' && !e.shiftKey) {
        const el = e.target;
        if (el.classList.contains('paragraph')) {
          e.preventDefault();
          vscode.postMessage({
            type: 'insertParagraph',
            sectionIndex: parseInt(el.dataset.section, 10),
            afterElementIndex: parseInt(el.dataset.element, 10)
          });
        }
      }
      // Handle Backspace on empty paragraph - delete the paragraph
      if (e.key === 'Backspace') {
        const el = e.target;
        if (el.classList.contains('paragraph')) {
          const selection = window.getSelection();
          const isAtStart = selection.anchorOffset === 0 && selection.focusOffset === 0;
          const isEmpty = el.textContent.trim() === '' || el.textContent === '';
          const elementIndex = parseInt(el.dataset.element, 10);

          // Delete paragraph if it's empty and not the first one
          if (isEmpty && elementIndex > 0) {
            e.preventDefault();
            vscode.postMessage({
              type: 'deleteParagraph',
              sectionIndex: parseInt(el.dataset.section, 10),
              elementIndex: elementIndex
            });
          }
          // Merge with previous paragraph if cursor is at the start
          else if (isAtStart && elementIndex > 0 && !isEmpty) {
            e.preventDefault();
            vscode.postMessage({
              type: 'mergeParagraphWithPrevious',
              sectionIndex: parseInt(el.dataset.section, 10),
              elementIndex: elementIndex
            });
          }
        }
      }
    }

    function handleContextMenu(e) {
      e.preventDefault();
      const menu = document.getElementById('contextMenu');
      menu.style.left = e.clientX + 'px';
      menu.style.top = e.clientY + 'px';
      menu.classList.add('visible');
      selectedElement = e.target.closest('.element');
    }

    document.addEventListener('click', () => {
      document.getElementById('contextMenu').classList.remove('visible');
    });

    document.querySelectorAll('.context-menu-item').forEach(item => {
      item.addEventListener('click', (e) => handleContextAction(e.target.dataset.action));
    });

    function handleContextAction(action) {
      if (!selectedElement) return;
      const sectionIndex = parseInt(selectedElement.dataset.section, 10);
      const elementIndex = parseInt(selectedElement.dataset.element, 10);
      const type = selectedElement.dataset.type;
      switch (action) {
        case 'insertParagraph':
          vscode.postMessage({ type: 'insertParagraph', sectionIndex, afterElementIndex: elementIndex });
          break;
        case 'deleteParagraph':
          vscode.postMessage({ type: 'deleteParagraph', sectionIndex, elementIndex });
          break;
        case 'insertRowBelow':
          if (type === 'table') {
            const row = selectedElement.querySelector('tr:focus-within, tr:last-child');
            vscode.postMessage({ type: 'insertTableRow', sectionIndex, elementIndex, afterRowIndex: row ? parseInt(row.dataset.row, 10) : 0 });
          }
          break;
        case 'deleteRow':
          if (type === 'table') {
            const row = selectedElement.querySelector('tr:focus-within');
            if (row) vscode.postMessage({ type: 'deleteTableRow', sectionIndex, elementIndex, rowIndex: parseInt(row.dataset.row, 10) });
          }
          break;
        case 'cut': document.execCommand('cut'); break;
        case 'copy': document.execCommand('copy'); break;
        case 'paste': document.execCommand('paste'); break;
      }
    }

    function applyCharStyle(styleName) {
      document.execCommand(styleName);
      if (selectedElement && selectedElement.classList.contains('paragraph')) {
        const style = {};
        style[styleName] = !isStyleActive(styleName);
        vscode.postMessage({
          type: 'applyCharacterStyle',
          sectionIndex: parseInt(selectedElement.dataset.section, 10),
          elementIndex: parseInt(selectedElement.dataset.element, 10),
          runIndex: 0,
          style
        });
      }
      updateToolbarState();
    }

    function applyParaStyle(styleName, value) {
      if (selectedElement && selectedElement.classList.contains('paragraph')) {
        const style = {};
        style[styleName] = value;
        vscode.postMessage({
          type: 'applyParagraphStyle',
          sectionIndex: parseInt(selectedElement.dataset.section, 10),
          elementIndex: parseInt(selectedElement.dataset.element, 10),
          style
        });
      }
    }

    function isStyleActive(styleName) { return document.queryCommandState(styleName); }

    function updateToolbarState() {
      document.getElementById('boldBtn').classList.toggle('active', isStyleActive('bold'));
      document.getElementById('italicBtn').classList.toggle('active', isStyleActive('italic'));
      document.getElementById('underlineBtn').classList.toggle('active', isStyleActive('underline'));
      document.getElementById('strikeBtn').classList.toggle('active', isStyleActive('strikethrough'));
    }

    document.getElementById('boldBtn').addEventListener('click', () => applyCharStyle('bold'));
    document.getElementById('italicBtn').addEventListener('click', () => applyCharStyle('italic'));
    document.getElementById('underlineBtn').addEventListener('click', () => applyCharStyle('underline'));
    document.getElementById('strikeBtn').addEventListener('click', () => applyCharStyle('strikethrough'));
    document.getElementById('alignLeft').addEventListener('click', () => applyParaStyle('align', 'left'));
    document.getElementById('alignCenter').addEventListener('click', () => applyParaStyle('align', 'center'));
    document.getElementById('alignRight').addEventListener('click', () => applyParaStyle('align', 'right'));
    document.getElementById('alignJustify').addEventListener('click', () => applyParaStyle('align', 'justify'));
    document.getElementById('indentIncrease').addEventListener('click', () => {
      if (selectedElement && selectedElement.classList.contains('paragraph')) {
        const currentMargin = parseFloat(selectedElement.style.marginLeft) || 0;
        applyParaStyle('marginLeft', currentMargin + 20);
      }
    });
    document.getElementById('indentDecrease').addEventListener('click', () => {
      if (selectedElement && selectedElement.classList.contains('paragraph')) {
        const currentMargin = parseFloat(selectedElement.style.marginLeft) || 0;
        applyParaStyle('marginLeft', Math.max(0, currentMargin - 20));
      }
    });
    document.getElementById('lineSpacing').addEventListener('change', (e) => applyParaStyle('lineSpacing', parseInt(e.target.value, 10)));
    document.getElementById('fontSize').addEventListener('change', (e) => {
      if (selectedElement && selectedElement.classList.contains('paragraph')) {
        vscode.postMessage({
          type: 'applyCharacterStyle',
          sectionIndex: parseInt(selectedElement.dataset.section, 10),
          elementIndex: parseInt(selectedElement.dataset.element, 10),
          runIndex: 0,
          style: { fontSize: parseInt(e.target.value, 10) }
        });
      }
    });
    document.getElementById('fontFamily').addEventListener('change', (e) => {
      if (selectedElement && selectedElement.classList.contains('paragraph')) {
        vscode.postMessage({
          type: 'applyCharacterStyle',
          sectionIndex: parseInt(selectedElement.dataset.section, 10),
          elementIndex: parseInt(selectedElement.dataset.element, 10),
          runIndex: 0,
          style: { fontName: e.target.value }
        });
      }
    });
    document.getElementById('textColor').addEventListener('change', (e) => {
      if (selectedElement && selectedElement.classList.contains('paragraph')) {
        vscode.postMessage({
          type: 'applyCharacterStyle',
          sectionIndex: parseInt(selectedElement.dataset.section, 10),
          elementIndex: parseInt(selectedElement.dataset.element, 10),
          runIndex: 0,
          style: { fontColor: e.target.value }
        });
      }
    });
    document.getElementById('bgColor').addEventListener('change', (e) => {
      if (selectedElement && selectedElement.classList.contains('paragraph')) {
        vscode.postMessage({
          type: 'applyCharacterStyle',
          sectionIndex: parseInt(selectedElement.dataset.section, 10),
          elementIndex: parseInt(selectedElement.dataset.element, 10),
          runIndex: 0,
          style: { backgroundColor: e.target.value }
        });
      }
    });
    document.getElementById('superscript').addEventListener('click', () => {
      if (selectedElement && selectedElement.classList.contains('paragraph')) {
        vscode.postMessage({
          type: 'applyCharacterStyle',
          sectionIndex: parseInt(selectedElement.dataset.section, 10),
          elementIndex: parseInt(selectedElement.dataset.element, 10),
          runIndex: 0,
          style: { superscript: true, subscript: false }
        });
      }
    });
    document.getElementById('subscript').addEventListener('click', () => {
      if (selectedElement && selectedElement.classList.contains('paragraph')) {
        vscode.postMessage({
          type: 'applyCharacterStyle',
          sectionIndex: parseInt(selectedElement.dataset.section, 10),
          elementIndex: parseInt(selectedElement.dataset.element, 10),
          runIndex: 0,
          style: { subscript: true, superscript: false }
        });
      }
    });

    function calculateAutoPageBreaks(content) {
      // Remove existing auto page breaks and page footnotes
      document.querySelectorAll('.auto-page-break, .page-footnotes').forEach(el => el.remove());

      // Get page settings from first section (or use defaults)
      const pageSettings = content.sections[0]?.pageSettings || {
        width: 595, height: 842,
        marginTop: 56.7, marginBottom: 56.7,
        marginLeft: 56.7, marginRight: 56.7,
        headerMargin: 0, footerMargin: 0
      };

      // Content height in pt (page height - margins)
      const contentHeightPt = pageSettings.height - pageSettings.marginTop - pageSettings.marginBottom;

      const docContainer = document.getElementById('document');

      // Build footnotes map from content
      const footnotesMap = {};
      if (content.footnotes) {
        content.footnotes.forEach(fn => {
          const num = fn.number || 0;
          let text = '';
          if (fn.paragraphs) {
            fn.paragraphs.forEach(p => {
              if (p.runs) {
                p.runs.forEach(run => {
                  if (run.text) text += run.text;
                });
              }
            });
          }
          footnotesMap[num] = text;
        });
      }

      // Get all content elements
      const elements = docContainer.querySelectorAll('.section > .element, .section > .section-break, .page-break');
      if (elements.length === 0) return;

      let currentPageBottomPt = contentHeightPt;
      let pageNum = 1;
      let useLinesegMode = false;

      // Check if we have lineseg data (vertpos/vertend)
      const elementsArray = Array.from(elements);
      const hasLinesegData = elementsArray.some(el => el.dataset.vertpos !== undefined);
      useLinesegMode = hasLinesegData;

      // Debug: count elements with lineseg data
      const linesegCount = elementsArray.filter(el => el.dataset.vertpos !== undefined).length;
      console.log('Page calc: useLinesegMode=' + useLinesegMode + ', contentHeightPt=' + contentHeightPt + ', elements=' + elementsArray.length + ', withLineseg=' + linesegCount);

      // Debug: show first few vertend values
      const vertendValues = elementsArray.slice(0, 20).map(el => el.dataset.vertend).filter(v => v);
      console.log('First vertend values:', vertendValues.join(', '));

      // Find max vertend
      let maxVertend = 0;
      elementsArray.forEach(el => {
        if (el.dataset.vertend) {
          const v = parseFloat(el.dataset.vertend);
          if (v > maxVertend) maxVertend = v;
        }
      });
      console.log('Max vertend:', maxVertend, 'pt, pages needed:', maxVertend / contentHeightPt);

      let i = 0;
      let prevVertend = 0;
      let pageStartIdx = 0; // Track start of current page

      // Helper function to collect footnotes from elements in a range
      function collectPageFootnotes(startIdx, endIdx) {
        const fnNumbers = new Set();
        for (let j = startIdx; j < endIdx && j < elementsArray.length; j++) {
          const el = elementsArray[j];
          // Find all footnote references within this element
          const fnRefs = el.querySelectorAll('[data-footnote]');
          fnRefs.forEach(ref => {
            const fnNum = parseInt(ref.dataset.footnote);
            if (fnNum && footnotesMap[fnNum]) {
              fnNumbers.add(fnNum);
            }
          });
        }
        return Array.from(fnNumbers).sort((a, b) => a - b);
      }

      // Helper function to create footnotes HTML
      function createPageFootnotesDiv(fnNumbers) {
        if (fnNumbers.length === 0) return null;
        const div = document.createElement('div');
        div.className = 'page-footnotes';
        // Short divider line (40%) at top, but content uses full width
        div.style.cssText = 'margin-top:15px;padding-top:8px;font-size:8pt;line-height:1.3;color:#333;position:relative;';
        let html = '<div style="position:absolute;top:0;left:0;width:40%;border-top:1px solid #666;"></div>';
        fnNumbers.forEach(num => {
          html += '<div style="margin-bottom:2px;padding-left:14px;text-indent:-14px;"><span style="font-size:7pt;vertical-align:super;">' + num + ')</span> ' + escapeHtml(footnotesMap[num]) + '</div>';
        });
        div.innerHTML = html;
        return div;
      }

      while (i < elementsArray.length) {
        const el = elementsArray[i];

        // Skip if this element already has a manual page break before it
        if (el.previousElementSibling && el.previousElementSibling.classList.contains('page-break')) {
          // Insert footnotes for current page before page break
          const fnNumbers = collectPageFootnotes(pageStartIdx, i);
          if (fnNumbers.length > 0) {
            const fnDiv = createPageFootnotesDiv(fnNumbers);
            if (fnDiv && el.previousElementSibling) {
              el.parentNode.insertBefore(fnDiv, el.previousElementSibling);
            }
          }
          pageNum++;
          prevVertend = 0; // Reset for new page
          pageStartIdx = i;
          i++;
          continue;
        }

        // Find group of elements that should stay together (keepWithNext chain)
        let groupEnd = i;
        while (groupEnd < elementsArray.length - 1 && elementsArray[groupEnd].dataset.keepWithNext === '1') {
          groupEnd++;
        }

        const firstEl = elementsArray[i];
        const lastEl = elementsArray[groupEnd];

        // Use lineseg data if available
        if (useLinesegMode) {
          const currentVertpos = firstEl.dataset.vertpos !== undefined ? parseFloat(firstEl.dataset.vertpos) : null;
          const currentVertend = lastEl.dataset.vertend !== undefined ? parseFloat(lastEl.dataset.vertend) : null;

          // If this element has no lineseg data (e.g., table), keep prevVertend and skip
          // This allows detecting page breaks after tables
          if (currentVertpos === null) {
            i = groupEnd + 1;
            continue;
          }

          // Detect page break: if current vertpos is less than previous vertend, a new page started
          // (HWPX resets vertpos to ~0 at each page boundary)
          // Only trigger if prevVertend is significant (> 100pt) to avoid false positives
          if (prevVertend > 100 && currentVertpos < prevVertend - 100) {
            console.log('Page break detected! prevVertend=' + prevVertend + ', currentVertpos=' + currentVertpos);

            // Insert footnotes for current page BEFORE page break
            const fnNumbers = collectPageFootnotes(pageStartIdx, i);
            if (fnNumbers.length > 0) {
              const fnDiv = createPageFootnotesDiv(fnNumbers);
              if (fnDiv && firstEl.parentNode) {
                firstEl.parentNode.insertBefore(fnDiv, firstEl);
              }
            }

            // Create auto page break marker
            const pageBreakDiv = document.createElement('div');
            pageBreakDiv.className = 'auto-page-break';
            pageBreakDiv.innerHTML = '<hr class="page-break-line" style="border-top-style:dotted;border-color:#ccc;"><span class="page-break-label" style="color:#aaa;">페이지 ' + (pageNum + 1) + '</span>';
            pageBreakDiv.style.cssText = 'margin:20px 0;text-align:center;position:relative;';

            if (firstEl.parentNode) {
              firstEl.parentNode.insertBefore(pageBreakDiv, firstEl);
            }
            pageNum++;
            // Reset prevVertend for new page
            prevVertend = 0;
            pageStartIdx = i;
          }

          // Update prevVertend
          if (currentVertend !== null) {
            prevVertend = currentVertend;
          } else {
            prevVertend = currentVertpos;
          }

          // Note: Very tall elements spanning multiple pages are not handled in lineseg mode
          // since HWPX already has pre-calculated page breaks
        }

        i = groupEnd + 1;
      }

      // Insert footnotes for the last page (after the loop)
      const lastPageFnNumbers = collectPageFootnotes(pageStartIdx, elementsArray.length);
      if (lastPageFnNumbers.length > 0) {
        const fnDiv = createPageFootnotesDiv(lastPageFnNumbers);
        if (fnDiv && elementsArray.length > 0) {
          const lastEl = elementsArray[elementsArray.length - 1];
          if (lastEl.parentNode) {
            lastEl.parentNode.appendChild(fnDiv);
          }
        }
      }

      // Update status bar with page count
      const statusWordCount = document.getElementById('statusWordCount');
      if (statusWordCount) {
        const wordText = statusWordCount.textContent;
        statusWordCount.textContent = wordText + ' | Pages: ~' + pageNum;
      }
    }

    function updateWordCount() {
      let wordCount = 0;
      document.querySelectorAll('.paragraph').forEach(el => {
        const text = el.textContent.trim();
        if (text) wordCount += text.split(/\\s+/).length;
      });
      document.getElementById('statusWordCount').textContent = 'Words: ' + wordCount;
    }

    document.addEventListener('selectionchange', updateToolbarState);

    // Global keyboard handler for undo/redo
    document.addEventListener('keydown', function(e) {
      if (e.ctrlKey || e.metaKey) {
        if (e.key.toLowerCase() === 'z') {
          e.preventDefault();
          if (e.shiftKey) {
            vscode.postMessage({ type: 'redo' });
          } else {
            vscode.postMessage({ type: 'undo' });
          }
        } else if (e.key.toLowerCase() === 'y') {
          e.preventDefault();
          vscode.postMessage({ type: 'redo' });
        }
      }
    });

    function escapeHtml(text) {
      if (!text) return '';
      const div = document.createElement('div');
      div.textContent = text;
      return div.innerHTML;
    }

    window.addEventListener('message', event => {
      const message = event.data;
      if (message.type === 'update') renderDocument(message.content);
    });

    vscode.postMessage({ type: 'ready' });
  </script>
</body>
</html>`;
}
