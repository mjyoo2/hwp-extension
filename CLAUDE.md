# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

HWP/HWPX Editor - A VSCode extension for editing Hangul Word Processor documents. Supports both modern HWPX (XML-based, editable) and legacy HWP (binary, read-only) formats.

## Build Commands

**VSCode Extension (root directory):**
```bash
npm run compile        # Build extension with esbuild
npm run watch          # Watch mode for development
npm run package        # Package as VSIX (requires vsce)
```

**MCP Server (mcp-server directory):**
```bash
cd mcp-server
npm run build          # Compile TypeScript
npm run start          # Run the server
```

**Development:** Press F5 in VSCode to launch the extension in debug mode (uses `.vscode/launch.json`).

## Architecture

The project consists of two main components:

### 1. VSCode Extension (root)

```
src/
├── extension.ts              # Entry point - registers HwpxEditorProvider
├── editor/
│   ├── HwpxEditorProvider.ts # CustomEditorProvider implementation
│   └── webview.ts            # HTML/CSS/JS for the editor UI
├── hwpx/
│   ├── HwpxParser.ts         # HWPX format parser (ZIP-based XML)
│   ├── HwpxDocument.ts       # Document model and operations
│   └── types.ts              # HWPML type definitions
└── hwp/
    └── HwpParser.ts          # Legacy HWP binary format parser (read-only)
```

**Data flow:** Extension opens file → Parser extracts content → HwpxDocument holds state → Webview renders UI → User edits sent via postMessage → Document updated → Changes saved back to file.

### 2. MCP Server (mcp-server/)

Standalone Model Context Protocol server for AI integration. Provides tools for document operations: `open_document`, `get_document_text`, `get_paragraphs`, `get_tables`, `search_text`, `replace_text`, `save_document`.

## Key Dependencies

- **hwp.js**: Parses legacy HWP binary format
- **jszip**: Handles HWPX ZIP archive structure
- **@modelcontextprotocol/sdk**: MCP protocol implementation (mcp-server only)

## File Format Notes

- **HWPX files** are ZIP archives containing XML files following the HWPML specification
- **HWP files** are binary OLE compound documents (read-only support via hwp.js)
- Type definitions in `src/hwpx/types.ts` are based on the official Korean HWPML specification

## Reference Documentation

The `reference/docs/` folder contains official HWP specification documents:
- `한글문서파일형식3.0_HWPML_revision1.2.md` - HWPML (XML format) specification
- `한글문서파일형식_5.0_revision1.3.md` - HWP 5.0 binary format specification

These documents define all supported elements including:
- Document structure (HEAD, BODY, TAIL)
- Text elements (paragraphs, runs, special characters)
- Drawing objects (tables, images, shapes, equations)
- Control elements (headers/footers, footnotes/endnotes, bookmarks, fields)
- Script and template support
