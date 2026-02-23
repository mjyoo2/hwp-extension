import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import * as fs from 'fs';
import * as path from 'path';
import { HwpxDocument } from './HwpxDocument';

// Document storage
const openDocuments = new Map<string, HwpxDocument>();

function generateId(): string {
  return Math.random().toString(36).substring(2, 11);
}

// ============================================================
// Tool Definitions
// ============================================================

const tools = [
  // === Document Management ===
  {
    name: 'open_document',
    description: 'Open an HWPX or HWP document for reading and editing',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: { type: 'string', description: 'Path to the HWPX or HWP file' },
      },
      required: ['file_path'],
    },
  },
  {
    name: 'close_document',
    description: 'Close an open document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID from open_document' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'save_document',
    description: 'Save the document (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        output_path: { type: 'string', description: 'Output path (optional, saves to original if omitted)' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'list_open_documents',
    description: 'List all currently open documents',
    inputSchema: { type: 'object', properties: {} },
  },

  // === Document Info ===
  {
    name: 'get_document_text',
    description: 'Get all text content from the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'get_document_structure',
    description: 'Get document structure (sections, paragraphs, tables, images count)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'get_document_metadata',
    description: 'Get document metadata (title, author, dates, etc.)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'set_document_metadata',
    description: 'Set document metadata (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        title: { type: 'string', description: 'Document title' },
        creator: { type: 'string', description: 'Author name' },
        subject: { type: 'string', description: 'Subject' },
        description: { type: 'string', description: 'Description' },
      },
      required: ['doc_id'],
    },
  },

  // === Paragraph Operations ===
  {
    name: 'get_paragraphs',
    description: 'Get paragraphs from the document with their text and styles',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (optional, all if omitted)' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'get_paragraph',
    description: 'Get a specific paragraph with full details',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index'],
    },
  },
  {
    name: 'insert_paragraph',
    description: 'Insert a new paragraph (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
        after_index: { type: 'number', description: 'Insert after this paragraph index (-1 for beginning, default: append to end)' },
        text: { type: 'string', description: 'Paragraph text' },
      },
      required: ['doc_id', 'text'],
    },
  },
  {
    name: 'delete_paragraph',
    description: 'Delete a paragraph by element index (HWPX only). Use get_paragraphs to find the element index.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
        paragraph_index: { type: 'number', description: 'Element index of the paragraph to delete (from get_paragraphs or insert_paragraph result)' },
      },
      required: ['doc_id', 'paragraph_index'],
    },
  },
  {
    name: 'update_paragraph_text',
    description: 'Update paragraph text content (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        run_index: { type: 'number', description: 'Run index (default 0)' },
        text: { type: 'string', description: 'New text content' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index', 'text'],
    },
  },
  {
    name: 'append_text_to_paragraph',
    description: 'Append text to an existing paragraph (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        text: { type: 'string', description: 'Text to append' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index', 'text'],
    },
  },

  // === List Operations ===
  {
    name: 'create_bulleted_list',
    description: 'Create a bulleted list (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
        items: { type: 'array', items: { type: 'string' }, description: 'List items' },
        after_element_index: { type: 'number', description: 'Insert after this element (default: append to end)' },
        bullet_char: { type: 'string', description: 'Bullet character (default: •)' },
      },
      required: ['doc_id', 'items'],
    },
  },
  {
    name: 'create_numbered_list',
    description: 'Create a numbered list (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
        items: { type: 'array', items: { type: 'string' }, description: 'List items' },
        after_element_index: { type: 'number', description: 'Insert after this element (default: append to end)' },
        start_number: { type: 'number', description: 'Starting number (default: 1)' },
        format: { type: 'string', enum: ['decimal', 'roman', 'alpha'], description: 'Numbering format (default: decimal)' },
      },
      required: ['doc_id', 'items'],
    },
  },
  {
    name: 'set_paragraph_numbering',
    description: 'Set numbering/bullet type for a paragraph (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        type: { type: 'string', enum: ['none', 'bullet', 'decimal', 'roman', 'alpha'], description: 'Numbering type' },
        level: { type: 'number', description: 'Indent level (0-9, default: 0)' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index', 'type'],
    },
  },

  // === Character Styling ===
  {
    name: 'set_text_style',
    description: 'Apply character formatting to a paragraph run (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        run_index: { type: 'number', description: 'Run index (default 0)' },
        bold: { type: 'boolean', description: 'Bold' },
        italic: { type: 'boolean', description: 'Italic' },
        underline: { type: 'boolean', description: 'Underline' },
        strikethrough: { type: 'boolean', description: 'Strikethrough' },
        font_name: { type: 'string', description: 'Font name' },
        font_size: { type: 'number', description: 'Font size in pt' },
        font_color: { type: 'string', description: 'Text color (hex)' },
        background_color: { type: 'string', description: 'Background color (hex)' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index'],
    },
  },
  {
    name: 'get_text_style',
    description: 'Get character formatting of a paragraph',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        run_index: { type: 'number', description: 'Run index (optional)' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index'],
    },
  },

  // === Paragraph Styling ===
  {
    name: 'set_paragraph_style',
    description: 'Apply paragraph formatting (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        align: { type: 'string', enum: ['left', 'center', 'right', 'justify', 'distribute'], description: 'Text alignment' },
        line_spacing: { type: 'number', description: 'Line spacing in %' },
        margin_left: { type: 'number', description: 'Left margin in pt' },
        margin_right: { type: 'number', description: 'Right margin in pt' },
        margin_top: { type: 'number', description: 'Top margin in pt' },
        margin_bottom: { type: 'number', description: 'Bottom margin in pt' },
        first_line_indent: { type: 'number', description: 'First line indent in pt' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index'],
    },
  },
  {
    name: 'get_paragraph_style',
    description: 'Get paragraph formatting',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index'],
    },
  },

  // === Search & Replace ===
  {
    name: 'search_text',
    description: 'Search for text in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        query: { type: 'string', description: 'Text to search for' },
        case_sensitive: { type: 'boolean', description: 'Case sensitive search (default: false)' },
        regex: { type: 'boolean', description: 'Use regular expression (default: false)' },
      },
      required: ['doc_id', 'query'],
    },
  },
  {
    name: 'replace_text',
    description: 'Replace text in the document (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        old_text: { type: 'string', description: 'Text to find' },
        new_text: { type: 'string', description: 'Replacement text' },
        case_sensitive: { type: 'boolean', description: 'Case sensitive (default: false)' },
        regex: { type: 'boolean', description: 'Use regular expression (default: false)' },
        replace_all: { type: 'boolean', description: 'Replace all occurrences (default: true)' },
      },
      required: ['doc_id', 'old_text', 'new_text'],
    },
  },
  {
    name: 'batch_replace',
    description: 'Perform multiple text replacements at once (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        replacements: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              old_text: { type: 'string' },
              new_text: { type: 'string' },
            },
          },
          description: 'Array of {old_text, new_text} pairs',
        },
      },
      required: ['doc_id', 'replacements'],
    },
  },

  // === Table Operations ===
  {
    name: 'get_tables',
    description: 'Get all tables from the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'get_table',
    description: 'Get a specific table with full data',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index within section' },
      },
      required: ['doc_id', 'section_index', 'table_index'],
    },
  },
  {
    name: 'get_table_cell',
    description: 'Get content of a specific table cell',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        row: { type: 'number', description: 'Row index (0-based)' },
        col: { type: 'number', description: 'Column index (0-based)' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'row', 'col'],
    },
  },
  {
    name: 'update_table_cell',
    description: 'Update content of a table cell (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        row: { type: 'number', description: 'Row index' },
        col: { type: 'number', description: 'Column index' },
        text: { type: 'string', description: 'New cell content' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'row', 'col', 'text'],
    },
  },
  {
    name: 'set_cell_properties',
    description: 'Set table cell properties (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        row: { type: 'number', description: 'Row index' },
        col: { type: 'number', description: 'Column index' },
        width: { type: 'number', description: 'Cell width' },
        height: { type: 'number', description: 'Cell height' },
        background_color: { type: 'string', description: 'Background color (hex)' },
        vertical_align: { type: 'string', enum: ['top', 'middle', 'bottom'], description: 'Vertical alignment' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'row', 'col'],
    },
  },
  {
    name: 'insert_table_row',
    description: 'Insert a new row in a table (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        after_row: { type: 'number', description: 'Insert after this row index (-1 for beginning)' },
        cell_texts: { type: 'array', items: { type: 'string' }, description: 'Text for each cell (optional)' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'after_row'],
    },
  },
  {
    name: 'delete_table_row',
    description: 'Delete a row from a table (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        row_index: { type: 'number', description: 'Row index to delete' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'row_index'],
    },
  },
  {
    name: 'insert_table_column',
    description: 'Insert a new column in a table (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        after_col: { type: 'number', description: 'Insert after this column (-1 for beginning)' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'after_col'],
    },
  },
  {
    name: 'delete_table_column',
    description: 'Delete a column from a table (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        col_index: { type: 'number', description: 'Column index to delete' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'col_index'],
    },
  },
  {
    name: 'get_table_as_csv',
    description: 'Export table content as CSV format',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        delimiter: { type: 'string', description: 'Delimiter character (default: comma)' },
      },
      required: ['doc_id', 'section_index', 'table_index'],
    },
  },

  {
    name: 'merge_table_cells',
    description: 'Merge a range of table cells (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        start_row: { type: 'number', description: 'Starting row index' },
        start_col: { type: 'number', description: 'Starting column index' },
        end_row: { type: 'number', description: 'Ending row index (inclusive)' },
        end_col: { type: 'number', description: 'Ending column index (inclusive)' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'start_row', 'start_col', 'end_row', 'end_col'],
    },
  },

  // === Page Settings ===
  {
    name: 'get_page_settings',
    description: 'Get page settings (paper size, margins)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'set_page_settings',
    description: 'Set page settings (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        width: { type: 'number', description: 'Page width in pt' },
        height: { type: 'number', description: 'Page height in pt' },
        margin_top: { type: 'number', description: 'Top margin in pt' },
        margin_bottom: { type: 'number', description: 'Bottom margin in pt' },
        margin_left: { type: 'number', description: 'Left margin in pt' },
        margin_right: { type: 'number', description: 'Right margin in pt' },
        orientation: { type: 'string', enum: ['portrait', 'landscape'], description: 'Page orientation' },
      },
      required: ['doc_id'],
    },
  },

  // === Copy/Move ===
  {
    name: 'copy_paragraph',
    description: 'Copy a paragraph to another location (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        source_section: { type: 'number', description: 'Source section index (default: 0)' },
        source_paragraph: { type: 'number', description: 'Source element index (from get_paragraphs result)' },
        target_section: { type: 'number', description: 'Target section index (default: 0)' },
        target_after: { type: 'number', description: 'Insert after this element index in target (-1 to insert at beginning)' },
      },
      required: ['doc_id', 'source_paragraph', 'target_after'],
    },
  },
  {
    name: 'move_paragraph',
    description: 'Move a paragraph to another location (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        source_section: { type: 'number', description: 'Source section index (default: 0)' },
        source_paragraph: { type: 'number', description: 'Source element index (from get_paragraphs result)' },
        target_section: { type: 'number', description: 'Target section index (default: 0)' },
        target_after: { type: 'number', description: 'Insert after this element index in target (-1 to insert at beginning)' },
      },
      required: ['doc_id', 'source_paragraph', 'target_after'],
    },
  },

  // === Statistics ===
  {
    name: 'get_word_count',
    description: 'Get word and character count statistics',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },

  // === Images ===
  {
    name: 'get_images',
    description: 'Get all images in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },

  // === Export ===
  {
    name: 'export_to_text',
    description: 'Export document to plain text file',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        output_path: { type: 'string', description: 'Output file path' },
      },
      required: ['doc_id', 'output_path'],
    },
  },
  {
    name: 'export_to_html',
    description: 'Export document to HTML file',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        output_path: { type: 'string', description: 'Output file path' },
      },
      required: ['doc_id', 'output_path'],
    },
  },

  // === Undo/Redo ===
  {
    name: 'undo',
    description: 'Undo the last change',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'redo',
    description: 'Redo the last undone change',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },

  // === Table Creation ===
  {
    name: 'insert_table',
    description: 'Insert a new table (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
        after_index: { type: 'number', description: 'Insert after this element index (-1 for beginning, default: append to end)' },
        rows: { type: 'number', description: 'Number of rows' },
        cols: { type: 'number', description: 'Number of columns' },
        width: { type: 'number', description: 'Table width (optional)' },
      },
      required: ['doc_id', 'rows', 'cols'],
    },
  },

  // === Header/Footer ===
  {
    name: 'get_header',
    description: 'Get header content for a section',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'set_header',
    description: 'Set header content for a section (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
        text: { type: 'string', description: 'Header text content' },
        apply_page_type: { type: 'string', enum: ['both', 'even', 'odd'], description: 'Apply to page type (default: both)' },
      },
      required: ['doc_id', 'text'],
    },
  },
  {
    name: 'get_footer',
    description: 'Get footer content for a section',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'set_footer',
    description: 'Set footer content for a section (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
        text: { type: 'string', description: 'Footer text content' },
        apply_page_type: { type: 'string', enum: ['both', 'even', 'odd'], description: 'Apply to page type (default: both)' },
      },
      required: ['doc_id', 'text'],
    },
  },

  // === Footnotes/Endnotes ===
  {
    name: 'get_footnotes',
    description: 'Get all footnotes in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'insert_footnote',
    description: 'Insert a footnote at a specific location (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        text: { type: 'string', description: 'Footnote text content' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index', 'text'],
    },
  },
  {
    name: 'get_endnotes',
    description: 'Get all endnotes in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'insert_endnote',
    description: 'Insert an endnote at a specific location (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        text: { type: 'string', description: 'Endnote text content' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index', 'text'],
    },
  },

  // === Bookmarks/Hyperlinks ===
  {
    name: 'get_bookmarks',
    description: 'Get all bookmarks in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'insert_bookmark',
    description: 'Insert a bookmark at a specific location (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        name: { type: 'string', description: 'Bookmark name' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index', 'name'],
    },
  },
  {
    name: 'get_hyperlinks',
    description: 'Get all hyperlinks in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'insert_hyperlink',
    description: 'Insert a hyperlink in a paragraph (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
        paragraph_index: { type: 'number', description: 'Paragraph index (default 0)' },
        url: { type: 'string', description: 'URL for the hyperlink' },
        text: { type: 'string', description: 'Display text for the hyperlink' },
      },
      required: ['doc_id', 'url', 'text'],
    },
  },

  // === Images ===
  {
    name: 'insert_image',
    description: 'Insert an image into the document (HWPX only). Provide either image_path (file path) or image_data (base64-encoded image data).',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
        after_index: { type: 'number', description: 'Insert after this element index (-1 for beginning, default: append to end)' },
        image_path: { type: 'string', description: 'Path to the image file' },
        image_data: { type: 'string', description: 'Base64-encoded image data (alternative to image_path)' },
        mime_type: { type: 'string', description: 'MIME type when using image_data (default: image/png)' },
        width: { type: 'number', description: 'Image width in hwpunit (default: 10000)' },
        height: { type: 'number', description: 'Image height in hwpunit (default: 10000)' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'update_image_size',
    description: 'Update the size of an existing image (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        image_index: { type: 'number', description: 'Image index within section' },
        width: { type: 'number', description: 'New width' },
        height: { type: 'number', description: 'New height' },
      },
      required: ['doc_id', 'section_index', 'image_index', 'width', 'height'],
    },
  },
  {
    name: 'delete_image',
    description: 'Delete an image from the document (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        image_index: { type: 'number', description: 'Image index within section' },
      },
      required: ['doc_id', 'section_index', 'image_index'],
    },
  },

  // === Drawing Objects ===
  {
    name: 'insert_line',
    description: 'Insert a line drawing object (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        after_index: { type: 'number', description: 'Insert after this element index (-1 for beginning)' },
        x1: { type: 'number', description: 'Start X coordinate' },
        y1: { type: 'number', description: 'Start Y coordinate' },
        x2: { type: 'number', description: 'End X coordinate' },
        y2: { type: 'number', description: 'End Y coordinate' },
        stroke_color: { type: 'string', description: 'Stroke color (hex)' },
        stroke_width: { type: 'number', description: 'Stroke width' },
      },
      required: ['doc_id', 'section_index', 'after_index', 'x1', 'y1', 'x2', 'y2'],
    },
  },
  {
    name: 'insert_rect',
    description: 'Insert a rectangle drawing object (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        after_index: { type: 'number', description: 'Insert after this element index (-1 for beginning)' },
        x: { type: 'number', description: 'X coordinate' },
        y: { type: 'number', description: 'Y coordinate' },
        width: { type: 'number', description: 'Width' },
        height: { type: 'number', description: 'Height' },
        fill_color: { type: 'string', description: 'Fill color (hex)' },
        stroke_color: { type: 'string', description: 'Stroke color (hex)' },
        stroke_width: { type: 'number', description: 'Stroke width' },
      },
      required: ['doc_id', 'section_index', 'after_index', 'x', 'y', 'width', 'height'],
    },
  },
  {
    name: 'insert_ellipse',
    description: 'Insert an ellipse drawing object (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        after_index: { type: 'number', description: 'Insert after this element index (-1 for beginning)' },
        cx: { type: 'number', description: 'Center X coordinate' },
        cy: { type: 'number', description: 'Center Y coordinate' },
        rx: { type: 'number', description: 'Radius X' },
        ry: { type: 'number', description: 'Radius Y' },
        fill_color: { type: 'string', description: 'Fill color (hex)' },
        stroke_color: { type: 'string', description: 'Stroke color (hex)' },
        stroke_width: { type: 'number', description: 'Stroke width' },
      },
      required: ['doc_id', 'section_index', 'after_index', 'cx', 'cy', 'rx', 'ry'],
    },
  },

  // === TextBox ===
  {
    name: 'insert_textbox',
    description: 'Insert a text box (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        x: { type: 'number', description: 'X position in pt' },
        y: { type: 'number', description: 'Y position in pt' },
        width: { type: 'number', description: 'Width in pt' },
        height: { type: 'number', description: 'Height in pt' },
        text: { type: 'string', description: 'Text content' },
        fill_color: { type: 'string', description: 'Fill color (hex, e.g., "#FFFFFF")' },
        stroke_color: { type: 'string', description: 'Stroke color (hex)' },
        stroke_width: { type: 'number', description: 'Stroke width in pt' },
      },
      required: ['doc_id', 'section_index', 'x', 'y', 'width', 'height', 'text'],
    },
  },
  {
    name: 'get_textboxes',
    description: 'Get all text boxes in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'update_textbox_text',
    description: 'Update text in a text box (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        textbox_id: { type: 'string', description: 'TextBox ID' },
        text: { type: 'string', description: 'New text content' },
      },
      required: ['doc_id', 'textbox_id', 'text'],
    },
  },
  {
    name: 'delete_textbox',
    description: 'Delete a text box (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        textbox_id: { type: 'string', description: 'TextBox ID' },
      },
      required: ['doc_id', 'textbox_id'],
    },
  },

  // === Equations ===
  {
    name: 'get_equations',
    description: 'Get all equations in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'insert_equation',
    description: 'Insert an equation (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        after_index: { type: 'number', description: 'Insert after this element index (-1 for beginning)' },
        script: { type: 'string', description: 'Equation script (HWP equation format)' },
      },
      required: ['doc_id', 'section_index', 'after_index', 'script'],
    },
  },

  // === Memos ===
  {
    name: 'get_memos',
    description: 'Get all memos/comments in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'insert_memo',
    description: 'Insert a memo/comment (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
        paragraph_index: { type: 'number', description: 'Paragraph index (default 0)' },
        author: { type: 'string', description: 'Memo author' },
        text: { type: 'string', description: 'Memo text content' },
        content: { type: 'string', description: 'Memo content (alias for text)' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'delete_memo',
    description: 'Delete a memo/comment (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        memo_id: { type: 'string', description: 'Memo ID to delete' },
      },
      required: ['doc_id', 'memo_id'],
    },
  },

  // === Sections ===
  {
    name: 'get_sections',
    description: 'Get all sections in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'insert_section',
    description: 'Insert a new section (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        after_index: { type: 'number', description: 'Insert after this section index (-1 for beginning)' },
      },
      required: ['doc_id', 'after_index'],
    },
  },
  {
    name: 'delete_section',
    description: 'Delete a section (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index to delete' },
      },
      required: ['doc_id', 'section_index'],
    },
  },

  // === Styles ===
  {
    name: 'get_styles',
    description: 'Get all defined styles in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'get_char_shapes',
    description: 'Get all character shape definitions',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'get_para_shapes',
    description: 'Get all paragraph shape definitions',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'apply_style',
    description: 'Apply a named style to a paragraph (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        style_id: { type: 'number', description: 'Style ID to apply' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index', 'style_id'],
    },
  },

  // === Column Definition ===
  {
    name: 'get_column_def',
    description: 'Get column definition for a section',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'set_column_def',
    description: 'Set column definition for a section (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
        count: { type: 'number', description: 'Number of columns' },
        type: { type: 'string', enum: ['newspaper', 'balanced', 'parallel'], description: 'Column type' },
        same_size: { type: 'boolean', description: 'Whether all columns have same width' },
        gap: { type: 'number', description: 'Gap between columns' },
      },
      required: ['doc_id', 'count'],
    },
  },

  // === New Document Creation ===
  {
    name: 'create_document',
    description: 'Create a new empty HWPX document',
    inputSchema: {
      type: 'object',
      properties: {
        title: { type: 'string', description: 'Document title (optional)' },
        creator: { type: 'string', description: 'Document author (optional)' },
      },
    },
  },
];

// ============================================================
// Server Setup
// ============================================================

const server = new Server(
  {
    name: 'hwpx-mcp-server',
    version: '0.3.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools }));

// ============================================================
// Tool Handlers
// ============================================================

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    switch (name) {
      // === Document Management ===
      case 'open_document': {
        const filePath = args?.file_path as string;
        if (!filePath) return error('file_path is required');

        const absolutePath = path.resolve(filePath);
        const data = fs.readFileSync(absolutePath);
        const docId = generateId();

        const doc = await HwpxDocument.createFromBuffer(docId, absolutePath, data);
        openDocuments.set(docId, doc);

        return success({
          doc_id: docId,
          format: doc.format,
          path: absolutePath,
          structure: doc.getStructure(),
          metadata: doc.getMetadata(),
        });
      }

      case 'close_document': {
        const docId = args?.doc_id as string;
        if (openDocuments.delete(docId)) {
          return success({ message: 'Document closed' });
        }
        return error('Document not found');
      }

      case 'save_document': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const data = await doc.save();
        const savePath = (args?.output_path as string) || doc.path;
        fs.writeFileSync(savePath, data);

        return success({ message: `Saved to ${savePath}` });
      }

      case 'list_open_documents': {
        const docs = Array.from(openDocuments.values()).map(d => ({
          id: d.id,
          path: d.path,
          format: d.format,
          isDirty: d.isDirty,
        }));
        return success({ documents: docs });
      }

      // === Document Info ===
      case 'get_document_text': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ text: doc.getAllText() });
      }

      case 'get_document_structure': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success(doc.getStructure());
      }

      case 'get_document_metadata': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ metadata: doc.getMetadata() });
      }

      case 'set_document_metadata': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const metadata: any = {};
        if (args?.title) metadata.title = args.title;
        if (args?.creator) metadata.creator = args.creator;
        if (args?.subject) metadata.subject = args.subject;
        if (args?.description) metadata.description = args.description;

        doc.setMetadata(metadata);
        return success({ metadata: doc.getMetadata() });
      }

      // === Paragraph Operations ===
      case 'get_paragraphs': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const sectionIndex = args?.section_index as number | undefined;
        const paragraphs = doc.getParagraphs(sectionIndex);
        return success({ paragraphs });
      }

      case 'get_paragraph': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const sectionIdx = Number(args?.section_index ?? 0);
        const paragraphIdx = Number(args?.paragraph_index ?? 0);
        const result = doc.getParagraph(sectionIdx, paragraphIdx);
        if (!result) return error('Paragraph not found');
        return success(result);
      }

      case 'insert_paragraph': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const sectionIdx = (args?.section_index as number) ?? 0;
        const section = doc.content.sections[sectionIdx];
        if (!section) return error('Section not found');
        const afterIdx = (args?.after_index as number) ?? section.elements.length - 1;
        const index = doc.insertParagraph(
          sectionIdx,
          afterIdx,
          args?.text as string
        );

        if (index === -1) return error('Failed to insert paragraph');
        return success({ message: 'Paragraph inserted', index });
      }

      case 'delete_paragraph': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.deleteParagraph(Number(args?.section_index ?? 0), Number(args?.paragraph_index ?? 0))) {
          return success({ message: 'Paragraph deleted' });
        }
        return error('Failed to delete paragraph');
      }

      case 'update_paragraph_text': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        doc.updateParagraphText(
          args?.section_index as number,
          args?.paragraph_index as number,
          args?.run_index as number ?? 0,
          args?.text as string
        );
        return success({ message: 'Paragraph updated' });
      }

      case 'append_text_to_paragraph': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        doc.appendTextToParagraph(
          args?.section_index as number,
          args?.paragraph_index as number,
          args?.text as string
        );
        return success({ message: 'Text appended' });
      }

      // === List Operations ===
      case 'create_bulleted_list': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const indices = doc.createBulletedList(
          (args?.section_index as number) ?? 0,
          args?.items as string[],
          args?.after_element_index as number | undefined,
          (args?.bullet_char as string) ?? '•'
        );
        return success({ inserted_indices: indices, count: indices.length });
      }

      case 'create_numbered_list': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const indices = doc.createNumberedList(
          (args?.section_index as number) ?? 0,
          args?.items as string[],
          args?.after_element_index as number | undefined,
          (args?.start_number as number) ?? 1,
          (args?.format as 'decimal' | 'roman' | 'alpha') ?? 'decimal'
        );
        return success({ inserted_indices: indices, count: indices.length });
      }

      case 'set_paragraph_numbering': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.setParagraphNumbering(
          Number(args?.section_index ?? 0),
          Number(args?.paragraph_index ?? 0),
          args?.type as 'none' | 'bullet' | 'decimal' | 'roman' | 'alpha',
          Number(args?.level ?? 0)
        );
        return success({ success: result });
      }

      // === Character Styling ===
      case 'set_text_style': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const style: any = {};
        if (args?.bold !== undefined) style.bold = args.bold;
        if (args?.italic !== undefined) style.italic = args.italic;
        if (args?.underline !== undefined) style.underline = args.underline;
        if (args?.strikethrough !== undefined) style.strikethrough = args.strikethrough;
        if (args?.font_name) style.fontName = args.font_name;
        if (args?.font_size) style.fontSize = args.font_size;
        if (args?.font_color) style.fontColor = args.font_color;
        if (args?.background_color) style.backgroundColor = args.background_color;

        doc.applyCharacterStyle(
          args?.section_index as number,
          args?.paragraph_index as number,
          args?.run_index as number ?? 0,
          style
        );
        return success({ message: 'Text style applied' });
      }

      case 'get_text_style': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const style = doc.getCharacterStyle(
          args?.section_index as number,
          args?.paragraph_index as number,
          args?.run_index as number | undefined
        );
        return success({ style });
      }

      // === Paragraph Styling ===
      case 'set_paragraph_style': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const style: any = {};
        if (args?.align) style.align = args.align;
        if (args?.line_spacing) style.lineSpacing = args.line_spacing;
        if (args?.margin_left) style.marginLeft = args.margin_left;
        if (args?.margin_right) style.marginRight = args.margin_right;
        if (args?.margin_top) style.marginTop = args.margin_top;
        if (args?.margin_bottom) style.marginBottom = args.margin_bottom;
        if (args?.first_line_indent) style.firstLineIndent = args.first_line_indent;

        doc.applyParagraphStyle(
          args?.section_index as number,
          args?.paragraph_index as number,
          style
        );
        return success({ message: 'Paragraph style applied' });
      }

      case 'get_paragraph_style': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const style = doc.getParagraphStyle(
          args?.section_index as number,
          args?.paragraph_index as number
        );
        return success({ style });
      }

      // === Search & Replace ===
      case 'search_text': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const results = doc.searchText(args?.query as string, {
          caseSensitive: args?.case_sensitive as boolean,
          regex: args?.regex as boolean,
        });

        return success({
          query: args?.query,
          total_matches: results.reduce((sum, r) => sum + r.count, 0),
          locations: results,
        });
      }

      case 'replace_text': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const count = doc.replaceText(args?.old_text as string, args?.new_text as string, {
          caseSensitive: args?.case_sensitive as boolean,
          regex: args?.regex as boolean,
          replaceAll: args?.replace_all as boolean ?? true,
        });

        return success({ message: `Replaced ${count} occurrence(s)`, count });
      }

      case 'batch_replace': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const replacements = args?.replacements as Array<{ old_text: string; new_text: string }>;
        if (!replacements) return error('replacements array is required');

        const results: any[] = [];
        for (const { old_text, new_text } of replacements) {
          const count = doc.replaceText(old_text, new_text);
          results.push({ old_text, new_text, count });
        }

        return success({ results });
      }

      // === Table Operations ===
      case 'get_tables': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ tables: doc.getTables() });
      }

      case 'get_table': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const sectionIdx = Number(args?.section_index ?? 0);
        const tableIdx = Number(args?.table_index ?? 0);
        const table = doc.getTable(sectionIdx, tableIdx);
        if (!table) return error('Table not found');
        return success(table);
      }

      case 'get_table_cell': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const cell = doc.getTableCell(
          Number(args?.section_index ?? 0),
          Number(args?.table_index ?? 0),
          Number(args?.row ?? 0),
          Number(args?.col ?? 0)
        );
        if (!cell) return error('Cell not found');
        return success(cell);
      }

      case 'update_table_cell': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.updateTableCell(
          args?.section_index as number,
          args?.table_index as number,
          args?.row as number,
          args?.col as number,
          args?.text as string
        )) {
          return success({ message: 'Cell updated' });
        }
        return error('Failed to update cell');
      }

      case 'set_cell_properties': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const props: any = {};
        if (args?.width) props.width = args.width;
        if (args?.height) props.height = args.height;
        if (args?.background_color) props.backgroundColor = args.background_color;
        if (args?.vertical_align) props.verticalAlign = args.vertical_align;

        if (doc.setCellProperties(
          args?.section_index as number,
          args?.table_index as number,
          args?.row as number,
          args?.col as number,
          props
        )) {
          return success({ message: 'Cell properties updated' });
        }
        return error('Failed to update cell properties');
      }

      case 'insert_table_row': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.insertTableRow(
          args?.section_index as number,
          args?.table_index as number,
          args?.after_row as number,
          args?.cell_texts as string[]
        )) {
          return success({ message: 'Row inserted' });
        }
        return error('Failed to insert row');
      }

      case 'delete_table_row': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.deleteTableRow(
          args?.section_index as number,
          args?.table_index as number,
          args?.row_index as number
        )) {
          return success({ message: 'Row deleted' });
        }
        return error('Failed to delete row');
      }

      case 'insert_table_column': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.insertTableColumn(
          args?.section_index as number,
          args?.table_index as number,
          args?.after_col as number
        )) {
          return success({ message: 'Column inserted' });
        }
        return error('Failed to insert column');
      }

      case 'delete_table_column': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.deleteTableColumn(
          args?.section_index as number,
          args?.table_index as number,
          args?.col_index as number
        )) {
          return success({ message: 'Column deleted' });
        }
        return error('Failed to delete column');
      }

      case 'get_table_as_csv': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const csv = doc.getTableAsCsv(
          Number(args?.section_index ?? 0),
          Number(args?.table_index ?? 0),
          args?.delimiter as string || ','
        );
        if (!csv) return error('Table not found');
        return success({ csv });
      }

      case 'merge_table_cells': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.mergeCells(
          args?.section_index as number,
          args?.table_index as number,
          args?.start_row as number,
          args?.start_col as number,
          args?.end_row as number,
          args?.end_col as number
        )) {
          return success({ message: 'Cells merged' });
        }
        return error('Failed to merge cells');
      }

      // === Page Settings ===
      case 'get_page_settings': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const settings = doc.getPageSettings(args?.section_index as number || 0);
        return success({ settings });
      }

      case 'set_page_settings': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const settings: any = {};
        if (args?.width) settings.width = args.width;
        if (args?.height) settings.height = args.height;
        if (args?.margin_top) settings.marginTop = args.margin_top;
        if (args?.margin_bottom) settings.marginBottom = args.margin_bottom;
        if (args?.margin_left) settings.marginLeft = args.margin_left;
        if (args?.margin_right) settings.marginRight = args.margin_right;
        if (args?.orientation) settings.orientation = args.orientation;

        if (doc.setPageSettings(args?.section_index as number || 0, settings)) {
          return success({ message: 'Page settings updated' });
        }
        return error('Failed to update page settings');
      }

      // === Copy/Move ===
      case 'copy_paragraph': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const srcSection = Number(args?.source_section ?? 0);
        const srcParagraph = Number(args?.source_paragraph ?? 0);
        const tgtSection = Number(args?.target_section ?? 0);
        const tgtAfter = Number(args?.target_after ?? -1);

        if (doc.copyParagraph(srcSection, srcParagraph, tgtSection, tgtAfter)) {
          return success({ message: 'Paragraph copied' });
        }
        return error('Failed to copy paragraph');
      }

      case 'move_paragraph': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const srcSection = Number(args?.source_section ?? 0);
        const srcParagraph = Number(args?.source_paragraph ?? 0);
        const tgtSection = Number(args?.target_section ?? 0);
        const tgtAfter = Number(args?.target_after ?? -1);

        if (doc.moveParagraph(srcSection, srcParagraph, tgtSection, tgtAfter)) {
          return success({ message: 'Paragraph moved' });
        }
        return error('Failed to move paragraph');
      }

      // === Statistics ===
      case 'get_word_count': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success(doc.getWordCount());
      }

      // === Images ===
      case 'get_images': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ images: doc.getImages() });
      }

      // === Export ===
      case 'export_to_text': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const text = doc.getAllText();
        const outputPath = args?.output_path as string;
        fs.writeFileSync(outputPath, text, 'utf-8');
        return success({ message: `Exported to ${outputPath}`, characters: text.length });
      }

      case 'export_to_html': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        let html = '<!DOCTYPE html><html><head><meta charset="UTF-8">';
        html += '<style>body{font-family:sans-serif;max-width:800px;margin:0 auto;padding:20px;}table{border-collapse:collapse;width:100%;}td,th{border:1px solid #ccc;padding:8px;}</style>';
        html += '</head><body>';

        const content = doc.content;
        for (const section of content.sections) {
          for (const element of section.elements) {
            if (element.type === 'paragraph') {
              const text = element.data.runs.map(r => escapeHtml(r.text)).join('');
              html += `<p>${text}</p>`;
            } else if (element.type === 'table') {
              const table = element.data;
              html += '<table>';
              for (const row of table.rows) {
                html += '<tr>';
                for (const cell of row.cells) {
                  const text = cell.paragraphs.map(p => p.runs.map(r => escapeHtml(r.text)).join('')).join('<br>');
                  html += `<td>${text}</td>`;
                }
                html += '</tr>';
              }
              html += '</table>';
            }
          }
        }

        html += '</body></html>';
        const outputPath = args?.output_path as string;
        fs.writeFileSync(outputPath, html, 'utf-8');
        return success({ message: `Exported to ${outputPath}` });
      }

      // === Undo/Redo ===
      case 'undo': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        if (doc.undo()) {
          return success({ message: 'Undo successful', canUndo: doc.canUndo(), canRedo: doc.canRedo() });
        }
        return error('Nothing to undo');
      }

      case 'redo': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        if (doc.redo()) {
          return success({ message: 'Redo successful', canUndo: doc.canUndo(), canRedo: doc.canRedo() });
        }
        return error('Nothing to redo');
      }

      // === Table Creation ===
      case 'insert_table': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const rowCount = args?.rows as number;
        const colCount = args?.cols as number;
        if (!rowCount || rowCount <= 0) return error('rows must be a positive number');
        if (!colCount || colCount <= 0) return error('cols must be a positive number');

        const sectionIdx = (args?.section_index as number) ?? 0;
        const section = doc.content.sections[sectionIdx];
        if (!section) return error('Section not found');
        const afterIdx = (args?.after_index as number) ?? section.elements.length - 1;

        const result = doc.insertTable(
          sectionIdx,
          afterIdx,
          rowCount,
          colCount,
          { width: args?.width as number | undefined }
        );
        if (!result) return error('Failed to insert table');
        return success({ message: 'Table inserted', tableIndex: result.tableIndex });
      }

      // === Header/Footer ===
      case 'get_header': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const result = doc.getHeader(args?.section_index as number || 0);
        return success({ header: result });
      }

      case 'set_header': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.setHeader(args?.section_index as number || 0, args?.text as string)) {
          return success({ message: 'Header set successfully' });
        }
        return error('Failed to set header');
      }

      case 'get_footer': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const result = doc.getFooter(args?.section_index as number || 0);
        return success({ footer: result });
      }

      case 'set_footer': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.setFooter(args?.section_index as number || 0, args?.text as string)) {
          return success({ message: 'Footer set successfully' });
        }
        return error('Failed to set footer');
      }

      // === Footnotes/Endnotes ===
      case 'get_footnotes': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ footnotes: doc.getFootnotes() });
      }

      case 'insert_footnote': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const fnText = args?.text as string;
        if (!fnText) return error('text is required');

        const result = doc.insertFootnote(
          Number(args?.section_index ?? 0),
          Number(args?.paragraph_index ?? 0),
          fnText
        );
        if (!result) return error('Failed to insert footnote');
        return success({ message: 'Footnote inserted', id: result.id });
      }

      case 'get_endnotes': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ endnotes: doc.getEndnotes() });
      }

      case 'insert_endnote': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const enText = args?.text as string;
        if (!enText) return error('text is required');

        const result = doc.insertEndnote(
          Number(args?.section_index ?? 0),
          Number(args?.paragraph_index ?? 0),
          enText
        );
        if (!result) return error('Failed to insert endnote');
        return success({ message: 'Endnote inserted', id: result.id });
      }

      // === Bookmarks/Hyperlinks ===
      case 'get_bookmarks': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ bookmarks: doc.getBookmarks() });
      }

      case 'insert_bookmark': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const bookmarkName = args?.name as string;
        if (!bookmarkName) return error('Bookmark name is required');

        if (doc.insertBookmark(
          Number(args?.section_index ?? 0),
          Number(args?.paragraph_index ?? 0),
          bookmarkName
        )) {
          return success({ message: 'Bookmark inserted' });
        }
        return error('Failed to insert bookmark');
      }

      case 'get_hyperlinks': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ hyperlinks: doc.getHyperlinks() });
      }

      case 'insert_hyperlink': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const hlUrl = args?.url as string;
        const hlText = args?.text as string;
        if (!hlUrl) return error('url is required');
        if (!hlText) return error('text is required');

        if (doc.insertHyperlink(
          Number(args?.section_index ?? 0),
          Number(args?.paragraph_index ?? 0),
          hlUrl,
          hlText
        )) {
          return success({ message: 'Hyperlink inserted' });
        }
        return error('Failed to insert hyperlink');
      }

      // === Image Operations ===
      case 'insert_image': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        let base64Data: string;
        let mimeType: string;

        const imagePath = args?.image_path as string | undefined;
        const imageDataArg = args?.image_data as string | undefined;

        if (imagePath) {
          if (!fs.existsSync(imagePath)) return error('Image file not found');
          const imageBuffer = fs.readFileSync(imagePath);
          base64Data = imageBuffer.toString('base64');
          const ext = path.extname(imagePath).toLowerCase();
          const mimeTypes: Record<string, string> = {
            '.png': 'image/png',
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.gif': 'image/gif',
            '.bmp': 'image/bmp',
          };
          mimeType = mimeTypes[ext] || 'image/png';
        } else if (imageDataArg) {
          // Strip data URI prefix if present (e.g., "data:image/png;base64,...")
          const dataUriMatch = imageDataArg.match(/^data:([^;]+);base64,(.+)$/s);
          if (dataUriMatch) {
            mimeType = (args?.mime_type as string) || dataUriMatch[1];
            base64Data = dataUriMatch[2];
          } else {
            base64Data = imageDataArg;
            mimeType = (args?.mime_type as string) || 'image/png';
          }
        } else {
          return error('Either image_path or image_data is required');
        }

        const sectionIndex = (args?.section_index as number) ?? 0;
        const section = doc.content.sections[sectionIndex];
        const afterIndex = (args?.after_index as number) ?? (section ? section.elements.length - 1 : -1);

        const result = doc.insertImage(
          sectionIndex,
          afterIndex,
          {
            data: base64Data,
            mimeType,
            width: (args?.width as number) || 10000,
            height: (args?.height as number) || 10000,
          }
        );
        if (!result) return error('Failed to insert image');
        return success({ message: 'Image inserted', id: result.id });
      }

      case 'update_image_size': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const sectionIdx = (args?.section_index as number) ?? 0;
        const imgIdx = args?.image_index as number;
        const sectionImages = doc.getImagesBySectionIndex(sectionIdx);
        if (imgIdx < 0 || imgIdx >= sectionImages.length) return error(`Image not found at index ${imgIdx} in section ${sectionIdx}`);

        if (doc.updateImageSize(
          sectionImages[imgIdx].id,
          args?.width as number,
          args?.height as number
        )) {
          return success({ message: 'Image size updated' });
        }
        return error('Failed to update image size');
      }

      case 'delete_image': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const delSectionIdx = (args?.section_index as number) ?? 0;
        const delImgIdx = args?.image_index as number;
        const delSectionImages = doc.getImagesBySectionIndex(delSectionIdx);
        if (delImgIdx < 0 || delImgIdx >= delSectionImages.length) return error(`Image not found at index ${delImgIdx} in section ${delSectionIdx}`);

        if (doc.deleteImage(delSectionImages[delImgIdx].id)) {
          return success({ message: 'Image deleted' });
        }
        return error('Failed to delete image');
      }

      // === Drawing Objects ===
      case 'insert_line': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.insertLine(
          Number(args?.section_index ?? args?.section ?? 0),
          args?.x1 as number,
          args?.y1 as number,
          args?.x2 as number,
          args?.y2 as number,
          {
            color: args?.stroke_color as string,
            width: args?.stroke_width as number,
          },
          args?.after_index as number
        );
        if (!result) return error('Failed to insert line');
        return success({ message: 'Line inserted', id: result.id });
      }

      case 'insert_rect': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.insertRect(
          Number(args?.section_index ?? args?.section ?? 0),
          args?.x as number,
          args?.y as number,
          args?.width as number,
          args?.height as number,
          {
            fillColor: args?.fill_color as string,
            strokeColor: args?.stroke_color as string,
          },
          args?.after_index as number
        );
        if (!result) return error('Failed to insert rectangle');
        return success({ message: 'Rectangle inserted', id: result.id });
      }

      case 'insert_ellipse': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.insertEllipse(
          Number(args?.section_index ?? args?.section ?? 0),
          (args?.cx ?? args?.x) as number,
          (args?.cy ?? args?.y) as number,
          (args?.rx ?? (args?.width ? (args.width as number) / 2 : undefined)) as number,
          (args?.ry ?? (args?.height ? (args.height as number) / 2 : undefined)) as number,
          {
            fillColor: args?.fill_color as string,
            strokeColor: args?.stroke_color as string,
          },
          args?.after_index as number
        );
        if (!result) return error('Failed to insert ellipse');
        return success({ message: 'Ellipse inserted', id: result.id });
      }

      // === TextBox ===
      case 'insert_textbox': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.insertTextBox(
          Number(args?.section_index ?? args?.section ?? 0),
          args?.x as number,
          args?.y as number,
          args?.width as number,
          args?.height as number,
          args?.text as string,
          {
            fillColor: args?.fill_color as string,
            strokeColor: args?.stroke_color as string,
            strokeWidth: args?.stroke_width as number,
          }
        );
        if (!result) return error('Failed to insert text box');
        return success({ message: 'Text box inserted', id: result.id });
      }

      case 'get_textboxes': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ textboxes: doc.getTextBoxes() });
      }

      case 'update_textbox_text': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        let tbIdU = args?.textbox_id as string;
        if (!tbIdU && args?.textbox_index !== undefined) {
          const tbs = doc.getTextBoxes();
          const tbByIdx = tbs[Number(args.textbox_index)];
          if (tbByIdx) tbIdU = tbByIdx.id;
        }
        if (tbIdU && doc.updateTextBoxText(tbIdU, args?.text as string)) {
          return success({ message: 'Text box updated' });
        }
        return error('Text box not found');
      }

      case 'delete_textbox': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        let tbIdD = args?.textbox_id as string;
        if (!tbIdD && args?.textbox_index !== undefined) {
          const tbs = doc.getTextBoxes();
          const tbByIdx = tbs[Number(args.textbox_index)];
          if (tbByIdx) tbIdD = tbByIdx.id;
        }
        if (tbIdD && doc.deleteTextBox(tbIdD)) {
          return success({ message: 'Text box deleted' });
        }
        return error('Text box not found');
      }

      // === Equations ===
      case 'get_equations': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ equations: doc.getEquations() });
      }

      case 'insert_equation': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.insertEquation(
          Number(args?.section_index ?? args?.section ?? 0),
          Number(args?.after_index ?? args?.paragraph ?? 0),
          (args?.script ?? args?.equation) as string
        );
        if (!result) return error('Failed to insert equation');
        return success({ message: 'Equation inserted', id: result.id });
      }

      // === Memos ===
      case 'get_memos': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ memos: doc.getMemos() });
      }

      case 'insert_memo': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const memoContent = (args?.text ?? args?.content) as string;
        if (!memoContent) return error('text is required');

        const result = doc.insertMemo(
          Number(args?.section_index ?? 0),
          Number(args?.paragraph_index ?? 0),
          memoContent,
          args?.author as string
        );
        if (!result) return error('Failed to insert memo');
        return success({ message: 'Memo inserted', id: result.id });
      }

      case 'delete_memo': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        let memoIdD = args?.memo_id as string;
        if (!memoIdD && args?.memo_index !== undefined) {
          const memos = doc.getMemos();
          const memoByIdx = memos[Number(args.memo_index)];
          if (memoByIdx) memoIdD = memoByIdx.id;
        }
        if (memoIdD && doc.deleteMemo(memoIdD)) {
          return success({ message: 'Memo deleted' });
        }
        return error('Failed to delete memo');
      }

      // === Sections ===
      case 'get_sections': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ sections: doc.getSections() });
      }

      case 'insert_section': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const newIndex = doc.insertSection(args?.after_index as number);
        return success({ message: 'Section inserted', index: newIndex });
      }

      case 'delete_section': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.deleteSection(Number(args?.section_index ?? args?.section ?? 0))) {
          return success({ message: 'Section deleted' });
        }
        return error('Failed to delete section');
      }

      // === Styles ===
      case 'get_styles': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ styles: doc.getStyles() });
      }

      case 'get_char_shapes': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ charShapes: doc.getCharShapes() });
      }

      case 'get_para_shapes': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ paraShapes: doc.getParaShapes() });
      }

      case 'apply_style': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const secIdxAS = Number(args?.section_index ?? args?.section ?? 0);
        const paraIdxAS = Number(args?.paragraph_index ?? args?.paragraph ?? 0);
        let styleIdAS = args?.style_id as number;
        if (styleIdAS === undefined && args?.style_name) {
          // Look up style by name
          const styles = doc.getStyles();
          if (styles?.styles) {
            for (const [id, s] of styles.styles) {
              if (s.name === args.style_name) { styleIdAS = id; break; }
            }
          }
          if (styleIdAS === undefined) styleIdAS = 0;
        }
        if (doc.applyStyle(secIdxAS, paraIdxAS, styleIdAS)) {
          return success({ message: 'Style applied' });
        }
        return error('Failed to apply style');
      }

      // === Column Definition ===
      case 'get_column_def': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ columnDef: doc.getColumnDef(args?.section_index as number || 0) });
      }

      case 'set_column_def': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.setColumnDef(
          args?.section_index as number || 0,
          args?.count as number,
          args?.gap as number
        )) {
          return success({ message: 'Column definition set' });
        }
        return error('Failed to set column definition');
      }

      // === Create New Document ===
      case 'create_document': {
        const docId = generateId();
        const doc = HwpxDocument.createNew(docId, args?.title as string, args?.creator as string);
        openDocuments.set(docId, doc);
        return success({
          doc_id: docId,
          format: 'hwpx',
          message: 'New document created',
        });
      }

      default:
        return error(`Unknown tool: ${name}`);
    }
  } catch (err: any) {
    return error(err.message);
  }
});

// ============================================================
// Helper Functions
// ============================================================

function getDoc(docId: string): HwpxDocument | undefined {
  return openDocuments.get(docId);
}

function success(data: any) {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

function error(message: string) {
  return { content: [{ type: 'text', text: JSON.stringify({ error: message }) }], isError: true };
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// ============================================================
// Main
// ============================================================

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch(console.error);
