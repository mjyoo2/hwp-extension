/**
 * HWP Section Writer - Serializes section content into binary tag records
 * Reverse of parseSectionData() in HwpParser.standalone.ts
 */

import {
  HWP_TAGS,
  CTRL_ID,
  TagStreamBuilder,
  BodyBuilder,
  ptToHwpunit,
} from './HwpTagBuilder';

import {
  HwpxSection,
  HwpxParagraph,
  TextRun,
  HwpxTable,
  TableRow,
  TableCell,
  HwpxImage,
  HwpxLine,
  HwpxRect,
  HwpxEllipse,
  HwpxTextBox,
  HwpxEquation,
  PageSettings,
  SectionElement,
  HeaderFooter,
  Footnote,
  Endnote,
} from './types';

// ============================================================
// Maps for looking up shape IDs from the DocInfo stream
// ============================================================

export interface SectionWriterMaps {
  /** Look up a charShape ID for a given TextRun. Returns 0 as default. */
  getCharShapeId: (run: TextRun) => number;
  /** Look up a paraShape ID for a given paragraph. Returns 0 as default. */
  getParaShapeId: (para: HwpxParagraph) => number;
  /** Look up a borderFill ID for a table cell. Returns 1 as default. */
  getBorderFillId: (cell: TableCell) => number;
  /** Look up a binItem ID for an image binaryId. Returns 1 as default. */
  getBinItemId: (binaryId: string) => number;
  /** Look up a Footnote by its reference number. Returns undefined if not found. */
  getFootnote?: (refNumber: number) => Footnote | undefined;
  /** Look up an Endnote by its reference number. Returns undefined if not found. */
  getEndnote?: (refNumber: number) => Endnote | undefined;
}

// ============================================================
// Section Write Context (replaces module-level mutable state)
// ============================================================

interface SectionWriteContext {
  instanceIdCounter: number;
  pageSettings: PageSettings;
}

/** Module-level context, reset per buildSectionStream call. */
let ctx: SectionWriteContext = {
  instanceIdCounter: 0,
  pageSettings: {
    width: 595, height: 842,
    marginTop: 56.7, marginBottom: 56.7,
    marginLeft: 56.7, marginRight: 56.7,
  },
};

function nextInstanceId(): number {
  return ++ctx.instanceIdCounter;
}

// ============================================================
// A. Text Paragraph Records
// ============================================================

/**
 * HWPTAG_PARA_HEADER (66)
 * Based on parser line ~1196-1228
 */
function writeParaHeader(
  textCharCount: number,
  controlMask: number,
  paraShapeId: number,
  styleId: number,
  charShapeCount: number,
  lineSegCount: number,
  level: number,
  stream: TagStreamBuilder
): void {
  const body = new BodyBuilder();
  // text char count (uint32) - number of UTF-16 code units
  body.addUint32(textCharCount);
  // control mask (uint32) - bit flags for which controls appear
  body.addUint32(controlMask);
  // paraShapeId (uint16)
  body.addUint16(paraShapeId);
  // styleId (uint8)
  body.addUint8(styleId);
  // breakType (uint8): 0 typically
  body.addUint8(0);
  // charShapeCount (uint16)
  body.addUint16(charShapeCount);
  // rangeTagCount (uint16): 0
  body.addUint16(0);
  // lineSegCount (uint16)
  body.addUint16(lineSegCount);
  // instanceId (uint32)
  body.addUint32(nextInstanceId());
  // changeTracking (uint16) - 2 extra bytes present in real HWP files
  body.addUint16(0);

  stream.addRecord(HWP_TAGS.HWPTAG_PARA_HEADER, level, body.build());
}

/**
 * Encode paragraph text runs into HWPTAG_PARA_TEXT (67) format.
 * Based on parser line ~1242-1293
 *
 * Returns the encoded bytes, the total char count, and control positions
 * for any extended controls (tables, images) inserted.
 *
 * Inline control block format (8 UTF-16 code units = 16 bytes):
 *   [ctrl_char, ctrlId_lo16, ctrlId_hi16, 0x0000, 0x0000, 0x0000, 0x0000, ctrl_char]
 * The control char used for section/column defs is 0x0002 (BOOKMARK_END range).
 * The control char used for table/image/shape is 0x0002 as well (range 0x0002-0x0008).
 */
interface ParaTextResult {
  bytes: Uint8Array;
  charCount: number;
  controlMask: number;
}

/**
 * Write one 8-unit inline control block (16 bytes):
 *   ctrl_char | ctrlId_lo | ctrlId_hi | 0 | 0 | 0 | 0 | ctrl_char
 */
function writeInlineControl(body: BodyBuilder, ctrlChar: number, ctrlId: number): void {
  body.addUint16(ctrlChar);
  body.addUint16(ctrlId & 0xFFFF);
  body.addUint16((ctrlId >>> 16) & 0xFFFF);
  body.addUint16(0x0000);
  body.addUint16(0x0000);
  body.addUint16(0x0000);
  body.addUint16(0x0000);
  body.addUint16(ctrlChar);
}

function encodeParaText(
  runs: TextRun[],
  hasTable: boolean,
  hasImage: boolean,
  sectionCtrlIds?: number[],  // ctrlIds to embed as inline controls before text (secd, cold, etc.)
): ParaTextResult {
  const body = new BodyBuilder();
  let charCount = 0;
  let controlMask = 0;

  // Embed section/column definition inline controls at the start (char 0x0002)
  if (sectionCtrlIds && sectionCtrlIds.length > 0) {
    for (const ctrlId of sectionCtrlIds) {
      writeInlineControl(body, 0x0002, ctrlId);
      charCount += 8;
    }
    controlMask |= (1 << 2);
  }

  // Insert extended control for table at start.
  // Tables use control char 0x000B (TABLE_DRAWING marker) and set controlMask bit 11 (0x800).
  if (hasTable) {
    writeInlineControl(body, 0x000B, CTRL_ID.TABLE);
    charCount += 8;
    controlMask |= (1 << 11); // bit 11 for table/drawing objects
  }

  // Insert extended control for image at start.
  // Images/GSO use control char 0x000B as well, and set controlMask bit 11.
  if (hasImage) {
    writeInlineControl(body, 0x000B, CTRL_ID.PICTURE);
    charCount += 8;
    controlMask |= (1 << 11);
  }

  // Encode text from runs
  for (const run of runs) {
    const text = run.text;
    for (let i = 0; i < text.length; i++) {
      const ch = text.charCodeAt(i);
      if (ch === 0x0A) {
        // Line break: write as control char 0x000A
        body.addUint16(0x000A);
        charCount++;
      } else {
        body.addUint16(ch);
        charCount++;
      }
    }
  }

  // Append paragraph break (0x000D)
  body.addUint16(0x000D);
  charCount++;

  return { bytes: body.build(), charCount, controlMask };
}

/**
 * HWPTAG_PARA_CHAR_SHAPE (68)
 * Array of (uint32 position, uint32 charShapeId) pairs
 * Based on parser line ~1230-1239
 * Note: written at paraLevel+1 (sub-record of PARA_HEADER)
 */
function writeParaCharShape(
  positions: Array<{ pos: number; id: number }>,
  paraLevel: number,
  stream: TagStreamBuilder
): void {
  const body = new BodyBuilder();
  for (const p of positions) {
    body.addUint32(p.pos);
    body.addUint32(p.id);
  }
  stream.addRecord(HWP_TAGS.HWPTAG_PARA_CHAR_SHAPE, paraLevel + 1, body.build());
}

/**
 * HWPTAG_PARA_LINE_SEG (69)
 * Simple single-line segment with approximate values.
 * Hangul recalculates line segments when opening.
 * Based on parser line ~107-124 for ParsedLineSeg
 * Note: written at paraLevel+1 (sub-record of PARA_HEADER)
 */
function writeParaLineSeg(
  paraLevel: number,
  stream: TagStreamBuilder
): void {
  const body = new BodyBuilder();
  // Single line segment - 36 bytes (9 uint32 values)
  body.addUint32(0);        // textStartPos
  body.addInt32(0);         // verticalPos
  body.addInt32(1000);      // lineHeight
  body.addInt32(800);       // textHeight
  body.addInt32(850);       // baselineDistance
  body.addInt32(600);       // lineSpacing
  body.addInt32(0);         // horizontalStart
  body.addInt32(42520);     // segmentWidth (approx A4 text width)
  body.addUint32(0x00000008); // flags: isLastInPara=true (bit 3)

  stream.addRecord(HWP_TAGS.HWPTAG_PARA_LINE_SEG, paraLevel + 1, body.build());
}

// ============================================================
// B. Section/Page Definition
// ============================================================

/**
 * Write CTRL_HEADER for section def ('secd') and PAGE_DEF.
 * Based on parser line ~1296-1298 (CTRL_HEADER) and ~1695-1707 (PAGE_DEF)
 */
function writeSectionDef(
  pageSettings: PageSettings,
  level: number,
  stream: TagStreamBuilder
): void {
  // CTRL_HEADER for 'secd'
  const ctrlBody = new BodyBuilder();
  ctrlBody.addUint32(CTRL_ID.SECTION); // ctrlId
  // Properties: flags for hide/show (all visible by default)
  ctrlBody.addUint32(0);
  // Column gap (uint16)
  ctrlBody.addUint16(0);
  // verticalAlign (uint16)
  ctrlBody.addUint16(0);
  // horizontal margin (uint16)
  ctrlBody.addUint16(0);
  // vertical margin (uint16)
  ctrlBody.addUint16(0);

  stream.addRecord(HWP_TAGS.HWPTAG_CTRL_HEADER, level, ctrlBody.build());

  // PAGE_DEF (73) at level+1
  // Based on parser line ~1696-1705
  // pageSettings uses pt units; HWP uses hwpunit (pt * 100)
  const pageDef = new BodyBuilder();
  pageDef.addUint32(ptToHwpunit(pageSettings.width));         // paperWidth
  pageDef.addUint32(ptToHwpunit(pageSettings.height));        // paperHeight
  pageDef.addUint32(ptToHwpunit(pageSettings.marginLeft));    // marginLeft
  pageDef.addUint32(ptToHwpunit(pageSettings.marginRight));   // marginRight
  pageDef.addUint32(ptToHwpunit(pageSettings.marginTop));     // marginTop
  pageDef.addUint32(ptToHwpunit(pageSettings.marginBottom));  // marginBottom
  pageDef.addUint32(ptToHwpunit(pageSettings.headerMargin ?? 11.8));  // headerMargin (~4252 hwpunit)
  pageDef.addUint32(ptToHwpunit(pageSettings.footerMargin ?? 11.8));  // footerMargin
  pageDef.addUint32(ptToHwpunit(pageSettings.gutterMargin ?? 0));     // gutter
  // Properties: bit 0 = landscape
  const props = pageSettings.orientation === 'landscape' ? 1 : 0;
  pageDef.addUint32(props);

  stream.addRecord(HWP_TAGS.HWPTAG_PAGE_DEF, level + 1, pageDef.build());
}

/**
 * Write CTRL_HEADER for column def ('cold').
 * Based on parser for CTRL_ID.COLUMN
 */
function writeColumnDef(
  level: number,
  stream: TagStreamBuilder
): void {
  const ctrlBody = new BodyBuilder();
  ctrlBody.addUint32(CTRL_ID.COLUMN); // ctrlId
  // Column type (uint16): 0=normal
  ctrlBody.addUint16(0);
  // Column count (uint16): 1
  ctrlBody.addUint16(1);
  // Direction (uint16): 0=left
  ctrlBody.addUint16(0);
  // Same width (uint16): 1=yes
  ctrlBody.addUint16(1);
  // Gap (uint16): 0
  ctrlBody.addUint16(0);

  stream.addRecord(HWP_TAGS.HWPTAG_CTRL_HEADER, level, ctrlBody.build());
}

// ============================================================
// C. Table Records
// ============================================================

/**
 * Write all records for a complete table.
 * Based on parser lines ~1296-1608 for table/cell parsing
 */
function writeTableControl(
  table: HwpxTable,
  level: number,
  maps: SectionWriterMaps,
  stream: TagStreamBuilder
): void {
  const rowCount = table.rows.length;
  // Compute colCount from the maximum number of logical columns
  let colCount = table.colCount || 0;
  if (!colCount && rowCount > 0) {
    for (const row of table.rows) {
      let rowCols = 0;
      for (const cell of row.cells) {
        rowCols += cell.colSpan || 1;
      }
      if (rowCols > colCount) colCount = rowCols;
    }
  }

  // Calculate table dimensions in hwpunit
  const tableWidthHwp = ptToHwpunit(table.width || 425.2);
  const tableHeightHwp = ptToHwpunit(table.height || 100);

  // a. CTRL_HEADER (71) for 'tbl '
  // Based on parser lines ~1296-1374
  const ctrlBody = new BodyBuilder();
  ctrlBody.addUint32(CTRL_ID.TABLE);  // ctrlId

  // Object properties: bit 4 = treatAsChar (inline table)
  const treatAsChar = table.position?.treatAsChar !== false;
  let objProps = 0;
  if (treatAsChar) {
    objProps |= (1 << 4);
  } else {
    // TextWrap mapping (bits 1-3)
    const wrapMap: Record<string, number> = {
      'square': 0, 'topAndBottom': 1, 'behindText': 2, 'inFrontOfText': 3,
      'tight': 4, 'through': 5
    };
    const wrapVal = wrapMap[table.textWrap || 'topAndBottom'] ?? 1;
    objProps |= (wrapVal << 1);
    // VertRelTo (bits 5-6)
    const vertRelMap: Record<string, number> = { 'paper': 0, 'page': 1, 'para': 2 };
    objProps |= ((vertRelMap[table.position?.vertRelTo || 'para'] ?? 2) << 5);
    // HorzRelTo (bits 9-10)
    const horzRelMap: Record<string, number> = { 'paper': 0, 'page': 1, 'column': 2, 'para': 3 };
    objProps |= ((horzRelMap[table.position?.horzRelTo || 'column'] ?? 2) << 9);
    // FlowWithText (bit 12)
    if (table.position?.flowWithText) objProps |= (1 << 12);
  }
  ctrlBody.addUint32(objProps);

  // Vertical offset (int32, hwpunit) at offset 8
  ctrlBody.addInt32(ptToHwpunit(table.position?.vertOffset || 0));
  // Horizontal offset (int32, hwpunit) at offset 12
  ctrlBody.addInt32(ptToHwpunit(table.position?.horzOffset || 0));
  // Width (uint32, hwpunit)
  ctrlBody.addUint32(tableWidthHwp);
  // Height (uint32, hwpunit)
  ctrlBody.addUint32(tableHeightHwp);
  // zOrder (int32)
  ctrlBody.addInt32(table.zOrder || 0);
  // OutMargin (4 x uint16, hwpunit)
  ctrlBody.addUint16(ptToHwpunit(table.outMargin?.left || 0));
  ctrlBody.addUint16(ptToHwpunit(table.outMargin?.right || 0));
  ctrlBody.addUint16(ptToHwpunit(table.outMargin?.top || 0));
  ctrlBody.addUint16(ptToHwpunit(table.outMargin?.bottom || 0));

  stream.addRecord(HWP_TAGS.HWPTAG_CTRL_HEADER, level, ctrlBody.build());

  // b. TABLE (77) at level+1
  // Based on parser lines ~1442-1490
  const tableBody = new BodyBuilder();
  // Properties (uint32): pageBreak + repeatHeader
  let tableProps = 0;
  const pageBreakMap: Record<string, number> = { 'none': 0, 'cell': 1, 'table': 2 };
  tableProps |= (pageBreakMap[table.pageBreak || 'none'] ?? 0);
  if (table.repeatHeader) tableProps |= (1 << 2);
  tableBody.addUint32(tableProps);
  // rowCount (uint16)
  tableBody.addUint16(rowCount);
  // colCount (uint16)
  tableBody.addUint16(colCount);
  // cellSpacing (uint16, hwpunit)
  tableBody.addUint16(ptToHwpunit(table.cellSpacing || 0));
  // Border padding / inMargin (4 x uint16, hwpunit)
  tableBody.addUint16(ptToHwpunit(table.inMargin?.left || 1.4));
  tableBody.addUint16(ptToHwpunit(table.inMargin?.right || 1.4));
  tableBody.addUint16(ptToHwpunit(table.inMargin?.top || 1.4));
  tableBody.addUint16(ptToHwpunit(table.inMargin?.bottom || 1.4));
  // borderFillId (uint16)
  tableBody.addUint16(table.borderFillId || 1);
  // Valid zone info: rowCount entries of uint16 (column count per row)
  for (let r = 0; r < rowCount; r++) {
    tableBody.addUint16(colCount);
  }

  stream.addRecord(HWP_TAGS.HWPTAG_TABLE, level + 1, tableBody.build());

  // c. For each cell (row-major order): LIST_HEADER + cell paragraphs
  // HWP binary format requires ALL grid positions to have cell records.
  // For merged cells, the HWPX format omits spanned positions entirely and
  // often sets colSpan/rowSpan=1 on master cells. We must:
  //  1. Place master cells at their colAddr positions
  //  2. Infer actual colSpan/rowSpan from grid gaps
  //  3. Fill remaining positions with filler cells

  // Step 1: Build grid and place master cells at their colAddr positions
  // grid[r][c] = cell reference or null (empty)
  const grid: (TableCell | null)[][] = [];
  for (let r = 0; r < rowCount; r++) {
    grid[r] = new Array(colCount).fill(null);
  }

  // Compute default column widths from the last row (which typically has all cells)
  const defaultColWidths: number[] = new Array(colCount).fill(0);
  // Try to find a row with all columns to get widths
  for (let r = rowCount - 1; r >= 0; r--) {
    const row = table.rows[r];
    if (!row) continue;
    if (row.cells.length === colCount) {
      for (let c = 0; c < colCount; c++) {
        defaultColWidths[c] = row.cells[c]?.width || 0;
      }
      break;
    }
  }
  // If still no widths, compute evenly from table width
  const tableWidthPt = table.width || 425.2;
  if (defaultColWidths.every(w => !w)) {
    const colW = tableWidthPt / colCount;
    for (let c = 0; c < colCount; c++) defaultColWidths[c] = colW;
  }

  for (let r = 0; r < rowCount; r++) {
    const row = table.rows[r];
    if (!row) continue;

    // If this row has exactly colCount cells AND no position is already occupied
    // by a sentinel from a prior row's rowSpan, place them sequentially.
    const hasOccupied = grid[r].some(slot => slot !== null);
    const useSequential = row.cells.length === colCount && !hasOccupied;

    let nextCol = 0;
    for (let c = 0; c < row.cells.length; c++) {
      const cell = row.cells[c];
      if (!cell) continue;

      let gridCol: number;
      if (useSequential) {
        gridCol = c;
      } else if (cell.colAddr !== undefined && cell.colAddr < colCount && grid[r][cell.colAddr] === null) {
        // Use colAddr only if it points to an unoccupied slot
        gridCol = cell.colAddr;
      } else if (cell.colAddr !== undefined && cell.colAddr < colCount
                 && grid[r][cell.colAddr]?.colSpan === -1 && grid[r][cell.colAddr]?.rowSpan === -1) {
        // colAddr points to a sentinel from a prior row's rowSpan.
        // This cell overrides the sentinel - the parent's rowSpan was too large.
        gridCol = cell.colAddr;
      } else {
        // Fall back to next available slot
        while (nextCol < colCount && grid[r][nextCol] !== null) nextCol++;
        gridCol = nextCol;
      }
      if (gridCol >= colCount) continue;

      grid[r][gridCol] = cell;

      // If cell has explicit colSpan/rowSpan > 1, mark spanned positions
      const cs = cell.colSpan || 1;
      const rs = cell.rowSpan || 1;
      if (cs > 1 || rs > 1) {
        for (let dr = 0; dr < rs; dr++) {
          for (let dc = 0; dc < cs; dc++) {
            if (dr === 0 && dc === 0) continue;
            const fr = r + dr, fc = gridCol + dc;
            if (fr < rowCount && fc < colCount && grid[fr][fc] === null) {
              // Sentinel: mark as occupied by a span (use a minimal marker cell)
              grid[fr][fc] = { paragraphs: [], colSpan: -1, rowSpan: -1 } as TableCell;
            }
          }
        }
      }
      nextCol = gridCol + cs;
    }
  }

  // Step 1.5: Fix rowSpan for master cells whose sentinels were overridden.
  // When a later row has a cell at a position that was a sentinel, the master's
  // rowSpan should be truncated to not extend past that row.
  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      const cell = grid[r][c];
      if (!cell || cell.colSpan === -1 || (cell.rowSpan || 1) <= 1) continue;
      // This is a master cell with rowSpan > 1; check if all spanned positions are still sentinels
      let actualRowSpan = 1;
      for (let dr = 1; dr < (cell.rowSpan || 1); dr++) {
        const fr = r + dr;
        if (fr >= rowCount) break;
        let allSentinels = true;
        for (let dc = 0; dc < (cell.colSpan || 1); dc++) {
          const fc = c + dc;
          if (fc >= colCount) break;
          const slot = grid[fr][fc];
          if (!(slot?.colSpan === -1 && slot?.rowSpan === -1)) {
            allSentinels = false;
            break;
          }
        }
        if (!allSentinels) break;
        actualRowSpan++;
      }
      if (actualRowSpan < (cell.rowSpan || 1)) {
        cell.rowSpan = actualRowSpan;
      }
    }
  }

  // Step 2: Infer colSpan/rowSpan for master cells from grid gaps.
  // A null position adjacent to a master cell means the master spans into it.
  // We infer colSpan first (consecutive nulls to the right), then rowSpan
  // (consecutive rows where the same column range is null).
  // "Sentinel" cells (colSpan=-1) are cells already marked by explicit spans.

  // Collect inferred spans: inferredSpans[r][c] = { colSpan, rowSpan }
  const inferredSpans: { colSpan: number; rowSpan: number }[][] = [];
  for (let r = 0; r < rowCount; r++) {
    inferredSpans[r] = [];
    for (let c = 0; c < colCount; c++) {
      inferredSpans[r][c] = { colSpan: 1, rowSpan: 1 };
    }
  }

  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      const cell = grid[r][c];
      if (!cell || (cell.colSpan === -1 && cell.rowSpan === -1)) continue;

      // If cell already has explicit spans > 1, use those
      if ((cell.colSpan || 1) > 1 || (cell.rowSpan || 1) > 1) {
        inferredSpans[r][c] = { colSpan: cell.colSpan || 1, rowSpan: cell.rowSpan || 1 };
        continue;
      }

      // Infer colSpan: count consecutive null columns to the right
      let cs = 1;
      while (c + cs < colCount && grid[r][c + cs] === null) cs++;
      // Infer rowSpan: check how many rows below have the same null pattern.
      // For cs=1 cells, only infer rowSpan if the ENTIRE row below is null
      // (no real cells or sentinels anywhere). A partially filled row with a
      // null at this column just means the row has fewer cells, not a merge.
      let rs = 1;
      let shouldInferRowSpan = false;
      if (cs > 1) {
        shouldInferRowSpan = true;
      } else if (grid[r + 1]?.[c] === null) {
        // For cs=1: check that the entire row below has no real cells
        // (only nulls allowed; sentinels from explicit spans are OK since they
        // indicate the row is part of another merge, not a stand-alone row)
        const nextRow = grid[r + 1];
        if (nextRow) {
          const hasRealCell = nextRow.some(slot =>
            slot !== null && !(slot.colSpan === -1 && slot.rowSpan === -1)
          );
          shouldInferRowSpan = !hasRealCell;
        }
      }
      if (shouldInferRowSpan) {
        // Check downward: for each subsequent row, the columns [c..c+cs-1] must all be null
        while (r + rs < rowCount) {
          let allNull = true;
          for (let dc = 0; dc < cs; dc++) {
            if (grid[r + rs][c + dc] !== null) { allNull = false; break; }
          }
          if (!allNull) break;
          rs++;
        }
      }

      inferredSpans[r][c] = { colSpan: cs, rowSpan: rs };

      // Mark the inferred spanned positions so they won't be double-counted
      for (let dr = 0; dr < rs; dr++) {
        for (let dc = 0; dc < cs; dc++) {
          if (dr === 0 && dc === 0) continue;
          const fr = r + dr, fc = c + dc;
          if (fr < rowCount && fc < colCount) {
            // Place a sentinel to indicate this position is covered
            grid[fr][fc] = { paragraphs: [], colSpan: -1, rowSpan: -1 } as TableCell;
          }
        }
      }
    }
  }

  // Step 3: Write all cells in row-major order
  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      const cell = grid[r][c];
      const spans = inferredSpans[r][c];

      if (cell && cell.colSpan !== -1 && cell.rowSpan !== -1) {
        // Master cell: compute width as sum of spanned column widths
        const masterCell = { ...cell };
        masterCell.colSpan = spans.colSpan;
        masterCell.rowSpan = spans.rowSpan;
        if (!masterCell.width || isNaN(masterCell.width)) {
          let w = 0;
          for (let dc = 0; dc < spans.colSpan; dc++) w += defaultColWidths[c + dc] || (tableWidthPt / colCount);
          masterCell.width = w;
        }
        writeCellListHeader(masterCell, r, c, level + 1, maps, stream);
        writeCellContent(masterCell, level + 2, maps, stream);
      } else {
        // Filler cell (spanned position): write empty cell
        const fillerWidth = defaultColWidths[c] || (tableWidthPt / colCount);
        const fillerCell: TableCell = {
          colSpan: 1,
          rowSpan: 1,
          width: fillerWidth,
          height: 30,
          paragraphs: [{ id: 'filler', runs: [{ text: '' }] }],
        };
        writeCellListHeader(fillerCell, r, c, level + 1, maps, stream);
        writeCellContent(fillerCell, level + 2, maps, stream);
      }
    }
  }
}

/**
 * Write LIST_HEADER (72) for a table cell.
 * Based on parser lines ~1493-1608
 */
function writeCellListHeader(
  cell: TableCell,
  rowIndex: number,
  colIndex: number,
  level: number,
  maps: SectionWriterMaps,
  stream: TagStreamBuilder
): void {
  const body = new BodyBuilder();

  // List header common fields (8 bytes)
  const paraCount = getCellParagraphCount(cell);
  // paraCount (uint16)
  body.addUint16(paraCount);
  // properties (uint32) - includes verticalAlign in bits 20-21
  const vertAlignMap: Record<string, number> = { 'top': 1, 'middle': 2, 'bottom': 3 };
  const vertAlignBits = vertAlignMap[cell.verticalAlign || 'top'] ?? 1;
  const listFlags = (vertAlignBits << 20);
  // Offset 0: paraCount(2), then properties(4), then padding(2) = 8 bytes total header
  body.addUint32(listFlags);
  body.addUint16(0); // padding to align to 8 bytes

  // Cell-specific fields after header (offset headerSize=8)
  // colIndex (uint16)
  body.addUint16(colIndex);
  // rowIndex (uint16)
  body.addUint16(rowIndex);
  // colSpan (uint16)
  body.addUint16(cell.colSpan || 1);
  // rowSpan (uint16)
  body.addUint16(cell.rowSpan || 1);
  // cellWidth (uint32, hwpunit)
  body.addUint32(ptToHwpunit(cell.width || 100));
  // cellHeight (uint32, hwpunit)
  body.addUint32(ptToHwpunit(cell.height || 30));
  // margins (uint16 x 4, hwpunit)
  body.addUint16(ptToHwpunit(cell.marginLeft ?? 1.4));
  body.addUint16(ptToHwpunit(cell.marginRight ?? 1.4));
  body.addUint16(ptToHwpunit(cell.marginTop ?? 1.4));
  body.addUint16(ptToHwpunit(cell.marginBottom ?? 1.4));
  // borderFillId (uint16) - always use maps lookup for HWP-local ID
  body.addUint16(maps.getBorderFillId(cell));

  stream.addRecord(HWP_TAGS.HWPTAG_LIST_HEADER, level, body.build());
}

/**
 * Get paragraphs from a cell, using elements or paragraphs field.
 */
/**
 * Count the total number of paragraphs that will be written for a cell,
 * including host paragraphs for nested tables/images.
 */
function getCellParagraphCount(cell: TableCell): number {
  if (cell.elements && cell.elements.length > 0) {
    let count = 0;
    for (const elem of cell.elements) {
      // Each element (paragraph, table, image) becomes one paragraph in HWP
      count++;
    }
    return count > 0 ? count : 1;
  }
  if (cell.paragraphs && cell.paragraphs.length > 0) {
    return cell.paragraphs.length;
  }
  return 1;
}

function getCellParagraphs(cell: TableCell): HwpxParagraph[] {
  if (cell.elements && cell.elements.length > 0) {
    const paras: HwpxParagraph[] = [];
    for (const elem of cell.elements) {
      if (elem.type === 'paragraph') {
        paras.push(elem.data);
      }
    }
    if (paras.length > 0) return paras;
  }
  if (cell.paragraphs && cell.paragraphs.length > 0) {
    return cell.paragraphs;
  }
  return [{ id: 'empty', runs: [{ text: '' }] }];
}

/**
 * Write all content for a cell: paragraphs, nested tables, and images.
 * When a cell has `elements`, iterate them in order to preserve structure.
 */
function writeCellContent(
  cell: TableCell,
  level: number,
  maps: SectionWriterMaps,
  stream: TagStreamBuilder
): void {
  if (cell.elements && cell.elements.length > 0) {
    for (const elem of cell.elements) {
      if (elem.type === 'paragraph') {
        writeParagraph(elem.data, level, maps, stream);
      } else if (elem.type === 'table') {
        writeTableHostParagraph(elem.data, level, maps, stream);
      } else if (elem.type === 'image') {
        writeImageHostParagraph(elem.data, level, maps, stream);
      }
    }
    return;
  }
  // Fallback: use paragraphs array
  const paras = getCellParagraphs(cell);
  for (const para of paras) {
    writeParagraph(para, level, maps, stream);
  }
}

// ============================================================
// D. Image Records
// ============================================================

/**
 * Write all records for an image control.
 * Based on parser lines ~1381-1687
 */
function writeImageControl(
  image: HwpxImage,
  level: number,
  maps: SectionWriterMaps,
  stream: TagStreamBuilder
): void {
  const widthHwp = ptToHwpunit(image.width);
  const heightHwp = ptToHwpunit(image.height);
  const binItemId = maps.getBinItemId(image.binaryId);

  // CTRL_HEADER (71) for '$pic'
  const ctrlBody = new BodyBuilder();
  ctrlBody.addUint32(CTRL_ID.PICTURE);  // ctrlId
  // Properties: treat as char by default
  ctrlBody.addUint32(1 << 4);  // treatAsChar = true
  // Vertical offset
  ctrlBody.addInt32(0);
  // Horizontal offset
  ctrlBody.addInt32(0);
  // Width
  ctrlBody.addUint32(widthHwp);
  // Height
  ctrlBody.addUint32(heightHwp);
  // zOrder
  ctrlBody.addInt32(0);
  // OutMargin (4 x uint16)
  ctrlBody.addUint16(0);
  ctrlBody.addUint16(0);
  ctrlBody.addUint16(0);
  ctrlBody.addUint16(0);

  stream.addRecord(HWP_TAGS.HWPTAG_CTRL_HEADER, level, ctrlBody.build());

  // SHAPE_COMPONENT (76) at level+1
  // Based on parser lines ~1612-1620
  const shapeBody = new BodyBuilder();
  // ShapeID (uint32)
  shapeBody.addUint32(0);
  // ComponentID (uint32) - use 'pic ' as marker
  shapeBody.addUint32(0);
  // X grouping offset
  shapeBody.addInt32(0);
  // Y grouping offset
  shapeBody.addInt32(0);
  // Group level (uint16)
  shapeBody.addUint16(0);
  // Local file version (uint16)
  shapeBody.addUint16(0);
  // Original width
  shapeBody.addUint32(widthHwp);
  // Original height
  shapeBody.addUint32(heightHwp);
  // Current width (at offset 28)
  shapeBody.addUint32(widthHwp);
  // Current height (at offset 32)
  shapeBody.addUint32(heightHwp);
  // Rotation properties (uint32)
  shapeBody.addUint32(0);
  // Rotation center X (int32)
  shapeBody.addInt32(widthHwp / 2);
  // Rotation center Y (int32)
  shapeBody.addInt32(heightHwp / 2);

  stream.addRecord(HWP_TAGS.HWPTAG_SHAPE_COMPONENT, level + 1, shapeBody.build());

  // SHAPE_COMPONENT_PICTURE (85) at level+1
  // Based on parser lines ~1623-1687
  const picBody = new BodyBuilder();
  // Border color (COLORREF)
  picBody.addUint32(0x00000000);
  // Border thickness (int32)
  picBody.addInt32(0);
  // Border properties (uint32)
  picBody.addUint32(0);
  // Image rect: 4 points x (x,y) = 8 int32 values
  // x0, y0 = top-left
  picBody.addInt32(0);
  picBody.addInt32(0);
  // x1, y1 = top-right
  picBody.addInt32(widthHwp);
  picBody.addInt32(0);
  // x2, y2 = bottom-right (note: some docs say bottom-left for x2,y2)
  picBody.addInt32(widthHwp);
  picBody.addInt32(heightHwp);
  // x3, y3 = bottom-left
  picBody.addInt32(0);
  picBody.addInt32(heightHwp);
  // Crop: left, top, right, bottom
  picBody.addInt32(0);
  picBody.addInt32(0);
  picBody.addInt32(0);
  picBody.addInt32(0);
  // Inner spacing (left, right, top, bottom as uint16)
  picBody.addUint16(0);
  picBody.addUint16(0);
  picBody.addUint16(0);
  picBody.addUint16(0);
  // Brightness (uint8)
  picBody.addUint8(0);
  // Contrast (uint8)
  picBody.addUint8(0);
  // Effect (uint8): 0=RealPic
  picBody.addUint8(0);
  // binItemId (uint16) - at offset ~71 (the parser scans for this)
  picBody.addUint16(binItemId);
  // Border transparency (uint8)
  picBody.addUint8(0);
  // Instance ID (uint32)
  picBody.addUint32(nextInstanceId());

  stream.addRecord(HWP_TAGS.HWPTAG_SHAPE_COMPONENT_PICTURE, level + 1, picBody.build());
}

// ============================================================
// E. Paragraph Writing (combines HEADER + TEXT + CHAR_SHAPE + LINE_SEG)
// ============================================================

// Page settings are stored in ctx.pageSettings (SectionWriteContext).

/**
 * Write a complete paragraph (PARA_HEADER + PARA_TEXT + PARA_CHAR_SHAPE + PARA_LINE_SEG)
 * without any trailing table/image children. Simple version for cell paragraphs.
 */
function writeParagraph(
  para: HwpxParagraph,
  level: number,
  maps: SectionWriterMaps,
  stream: TagStreamBuilder,
  sectionCtrlIds?: number[],
): void {
  writeParagraphWithTrailing(para, level, maps, stream, sectionCtrlIds, [], []);
}

/**
 * Write a paragraph with optional trailing table/image children.
 *
 * In real HWP files, tables and images are sub-records (at level+1) of the paragraph
 * that contains their inline 0x000B control char in PARA_TEXT. There is no separate
 * host paragraph for each table/image.
 *
 * @param sectionCtrlIds - ctrlIds to prepend as secd/cold inline controls (first paragraph only)
 * @param trailingTables - tables to embed as children of this paragraph
 * @param trailingImages - images to embed as children of this paragraph
 */
function writeParagraphWithTrailing(
  para: HwpxParagraph,
  level: number,
  maps: SectionWriterMaps,
  stream: TagStreamBuilder,
  sectionCtrlIds: number[] | undefined,
  trailingTables: HwpxTable[],
  trailingImages: HwpxImage[],
): void {
  const runs = para.runs.length > 0 ? para.runs : [{ text: '' }];

  // Encode PARA_TEXT with:
  //   - secd/cold inline controls at the start (using 0x0002)
  //   - one 0x000B control block per trailing table
  //   - one 0x000B control block per trailing image
  //   - text runs
  //   - paragraph break 0x000D
  const body = new BodyBuilder();
  let charCount = 0;
  let controlMask = 0;

  // Prepend secd/cold inline controls (char 0x0002)
  if (sectionCtrlIds && sectionCtrlIds.length > 0) {
    for (const ctrlId of sectionCtrlIds) {
      writeInlineControl(body, 0x0002, ctrlId);
      charCount += 8;
    }
    controlMask |= (1 << 2);
  }

  // Append table inline controls (char 0x000B)
  for (const _tbl of trailingTables) {
    writeInlineControl(body, 0x000B, CTRL_ID.TABLE);
    charCount += 8;
    controlMask |= (1 << 11);
  }

  // Append image inline controls (char 0x000B)
  for (const _img of trailingImages) {
    writeInlineControl(body, 0x000B, CTRL_ID.PICTURE);
    charCount += 8;
    controlMask |= (1 << 11);
  }

  // Collect footnote/endnote references from runs
  const trailingFootnotes: Array<{ type: 'footnote' | 'endnote'; ref: number; data?: Footnote | Endnote }> = [];
  for (const run of runs) {
    if (run.footnoteRef !== undefined && run.footnoteRef > 0) {
      const fnData = maps.getFootnote?.(run.footnoteRef);
      trailingFootnotes.push({ type: 'footnote', ref: run.footnoteRef, data: fnData });
    }
    if (run.endnoteRef !== undefined && run.endnoteRef > 0) {
      const enData = maps.getEndnote?.(run.endnoteRef);
      trailingFootnotes.push({ type: 'endnote', ref: run.endnoteRef, data: enData });
    }
  }

  // Append footnote/endnote inline controls (char 0x000B)
  for (const fn of trailingFootnotes) {
    const fnCtrlId = fn.type === 'footnote' ? CTRL_ID.FOOTNOTE : CTRL_ID.ENDNOTE;
    writeInlineControl(body, 0x000B, fnCtrlId);
    charCount += 8;
    controlMask |= (1 << 11);
  }

  // Encode text runs
  for (const run of runs) {
    const text = run.text;
    for (let i = 0; i < text.length; i++) {
      const ch = text.charCodeAt(i);
      if (ch === 0x0A) {
        body.addUint16(0x000A);
        charCount++;
      } else {
        body.addUint16(ch);
        charCount++;
      }
    }
  }

  // Paragraph break
  body.addUint16(0x000D);
  charCount++;

  // Build PARA_CHAR_SHAPE positions
  const ctrlOffset = (sectionCtrlIds?.length ?? 0) * 8
    + trailingTables.length * 8
    + trailingImages.length * 8
    + trailingFootnotes.length * 8;
  const charShapePositions: Array<{ pos: number; id: number }> = [];
  if (ctrlOffset > 0) {
    // Use the first run's charShapeId for the control region instead of default (0)
    const firstRunCsId = runs.length > 0 ? maps.getCharShapeId(runs[0]) : 0;
    charShapePositions.push({ pos: 0, id: firstRunCsId });
  }
  let pos = ctrlOffset;
  for (const run of runs) {
    // Only add a charShape position entry if this run contributes characters;
    // empty-text runs would create duplicate positions that overwrite the
    // previous run's charShape mapping
    if (run.text.length > 0) {
      const csId = maps.getCharShapeId(run);
      charShapePositions.push({ pos, id: csId });
    }
    pos += run.text.length;
  }
  if (charShapePositions.length === 0) {
    charShapePositions.push({ pos: 0, id: 0 });
  }

  const paraShapeId = maps.getParaShapeId(para);

  // PARA_HEADER at `level`
  writeParaHeader(
    charCount,
    controlMask,
    paraShapeId,
    para.style || 0,
    charShapePositions.length,
    0, // lineSegCount: 0 in real HWP files
    level,
    stream
  );

  // PARA_TEXT at level+1
  if (charCount > 0) {
    stream.addRecord(HWP_TAGS.HWPTAG_PARA_TEXT, level + 1, body.build());
  }

  // PARA_CHAR_SHAPE at level+1
  writeParaCharShape(charShapePositions, level, stream);

  // PARA_LINE_SEG at level+1
  writeParaLineSeg(level, stream);

  // Trailing secd/cold CTRL_HEADERs at level+1
  if (sectionCtrlIds && sectionCtrlIds.length > 0) {
    writeSectionDefControls(ctx.pageSettings, level, stream);
  }

  // Trailing table CTRL_HEADERs at level+1
  for (const tbl of trailingTables) {
    writeTableControl(tbl, level + 1, maps, stream);
  }

  // Trailing image CTRL_HEADERs at level+1
  for (const img of trailingImages) {
    writeImageControl(img, level + 1, maps, stream);
  }

  // Trailing footnote/endnote CTRL_HEADERs at level+1
  for (const fn of trailingFootnotes) {
    if (fn.data) {
      writeFootnoteEndnoteControl(fn.data, fn.type, level + 1, maps, stream);
    }
  }
}

/**
 * Write a section-definition first paragraph.
 *
 * In real HWP files, the section definition controls (secd, cold) are embedded
 * as inline control blocks at the START of the first content paragraph's PARA_TEXT,
 * NOT in a separate dedicated paragraph. This function handles an empty first
 * paragraph (when there are no content elements) and the secd/cold CTRL_HEADERs.
 *
 * The inline controls use char 0x0002, and are prepended to the paragraph text.
 * CTRL_HEADER records for secd/cold follow at level 1 (same as other PARA_TEXT sub-records).
 */
function writeSectionDefControls(
  pageSettings: PageSettings,
  level: number,
  stream: TagStreamBuilder
): void {
  // Section definition CTRL_HEADER + PAGE_DEF at level+1
  writeSectionDef(pageSettings, level + 1, stream);

  // Column definition CTRL_HEADER at level+1
  writeColumnDef(level + 1, stream);
}

/**
 * Write a table-hosting paragraph (paragraph with embedded table control char).
 * PARA_TEXT/CS/LS at level+1; table CTRL_HEADER at level+1 as well.
 */
function writeTableHostParagraph(
  table: HwpxTable,
  level: number,
  maps: SectionWriterMaps,
  stream: TagStreamBuilder
): void {
  const textResult = encodeParaText([], true, false);

  writeParaHeader(
    textResult.charCount,
    textResult.controlMask,
    0, // paraShapeId
    0, // styleId
    1, // charShapeCount
    0, // lineSegCount: 0 in real HWP files
    level,
    stream
  );

  stream.addRecord(HWP_TAGS.HWPTAG_PARA_TEXT, level + 1, textResult.bytes);
  writeParaCharShape([{ pos: 0, id: 0 }], level, stream);
  writeParaLineSeg(level, stream);

  // Table CTRL_HEADER at level+1
  writeTableControl(table, level + 1, maps, stream);
}

/**
 * Write an image-hosting paragraph (paragraph with embedded image control char).
 * PARA_TEXT/CS/LS at level+1; image CTRL_HEADER at level+1 as well.
 */
function writeImageHostParagraph(
  image: HwpxImage,
  level: number,
  maps: SectionWriterMaps,
  stream: TagStreamBuilder
): void {
  const textResult = encodeParaText([], false, true);

  writeParaHeader(
    textResult.charCount,
    textResult.controlMask,
    0, // paraShapeId
    0, // styleId
    1, // charShapeCount
    0, // lineSegCount: 0 in real HWP files
    level,
    stream
  );

  stream.addRecord(HWP_TAGS.HWPTAG_PARA_TEXT, level + 1, textResult.bytes);
  writeParaCharShape([{ pos: 0, id: 0 }], level, stream);
  writeParaLineSeg(level, stream);

  // Image CTRL_HEADER at level+1
  writeImageControl(image, level + 1, maps, stream);
}

/**
 * Write a table-hosting paragraph that also embeds secd/cold inline controls (first element).
 */
function writeTableHostParagraphWithSectionDef(
  table: HwpxTable,
  pageSettings: PageSettings,
  level: number,
  maps: SectionWriterMaps,
  stream: TagStreamBuilder
): void {
  // Encode: secd + cold inline controls + table control char
  const body = new BodyBuilder();
  let charCount = 0;
  let controlMask = (1 << 2) | (1 << 11); // bit 2 for secd/cold, bit 11 for table

  // secd inline control (uses 0x0002)
  writeInlineControl(body, 0x0002, CTRL_ID.SECTION);
  charCount += 8;
  // cold inline control (uses 0x0002)
  writeInlineControl(body, 0x0002, CTRL_ID.COLUMN);
  charCount += 8;
  // table inline control (uses 0x000B)
  writeInlineControl(body, 0x000B, CTRL_ID.TABLE);
  charCount += 8;
  // paragraph break
  body.addUint16(0x000D);
  charCount++;

  writeParaHeader(charCount, controlMask, 0, 0, 1, 0, level, stream);
  stream.addRecord(HWP_TAGS.HWPTAG_PARA_TEXT, level + 1, body.build());
  writeParaCharShape([{ pos: 0, id: 0 }], level, stream);
  writeParaLineSeg(level, stream);

  // secd/cold CTRL_HEADERs at level+1
  writeSectionDefControls(pageSettings, level, stream);
  // table CTRL_HEADER at level+1
  writeTableControl(table, level + 1, maps, stream);
}

/**
 * Write an image-hosting paragraph that also embeds secd/cold inline controls (first element).
 */
function writeImageHostParagraphWithSectionDef(
  image: HwpxImage,
  pageSettings: PageSettings,
  level: number,
  maps: SectionWriterMaps,
  stream: TagStreamBuilder
): void {
  const body = new BodyBuilder();
  let charCount = 0;
  const controlMask = (1 << 2) | (1 << 11); // bit 2 for secd/cold, bit 11 for image

  writeInlineControl(body, 0x0002, CTRL_ID.SECTION);
  charCount += 8;
  writeInlineControl(body, 0x0002, CTRL_ID.COLUMN);
  charCount += 8;
  writeInlineControl(body, 0x000B, CTRL_ID.PICTURE);
  charCount += 8;
  body.addUint16(0x000D);
  charCount++;

  writeParaHeader(charCount, controlMask, 0, 0, 1, 0, level, stream);
  stream.addRecord(HWP_TAGS.HWPTAG_PARA_TEXT, level + 1, body.build());
  writeParaCharShape([{ pos: 0, id: 0 }], level, stream);
  writeParaLineSeg(level, stream);

  writeSectionDefControls(pageSettings, level, stream);
  writeImageControl(image, level + 1, maps, stream);
}

// ============================================================
// F. Header/Footer Controls
// ============================================================

/**
 * Write all records for a header or footer control.
 * CTRL_HEADER (head/foot) + LIST_HEADER + child paragraphs.
 *
 * Based on parser lines ~1403-1406: CTRL_ID.HEADER/FOOTER triggers inHeaderFooter mode,
 * followed by LIST_HEADER with paragraphs at deeper nesting levels.
 */
function writeHeaderFooterControl(
  hf: HeaderFooter,
  type: 'header' | 'footer',
  level: number,
  maps: SectionWriterMaps,
  stream: TagStreamBuilder
): void {
  const ctrlId = type === 'header' ? CTRL_ID.HEADER : CTRL_ID.FOOTER;

  // CTRL_HEADER at `level`
  const ctrlBody = new BodyBuilder();
  ctrlBody.addUint32(ctrlId);
  // applyTo flag: 0=both, 1=even, 2=odd
  const applyToMap: Record<string, number> = { 'Both': 0, 'Even': 1, 'Odd': 2, 'both': 0, 'even': 1, 'odd': 2 };
  ctrlBody.addUint32(applyToMap[hf.applyPageType || 'Both'] ?? 0);

  stream.addRecord(HWP_TAGS.HWPTAG_CTRL_HEADER, level, ctrlBody.build());

  // LIST_HEADER at level+1
  const paragraphs = hf.paragraphs.length > 0 ? hf.paragraphs : [{ id: 'empty-hf', runs: [{ text: '' }] }];
  const ps = ctx.pageSettings;
  const textWidth = ptToHwpunit(ps.width - (ps.marginLeft || 0) - (ps.marginRight || 0));
  const textHeight = ptToHwpunit(type === 'header' ? (ps.headerMargin ?? 11.8) : (ps.footerMargin ?? 11.8));

  const listBody = new BodyBuilder();
  listBody.addUint16(paragraphs.length); // paraCount
  listBody.addUint32(0);                 // properties
  listBody.addUint16(0);                 // padding
  // Header/footer specific: textWidth, textHeight
  listBody.addUint32(textWidth);
  listBody.addUint32(textHeight);

  stream.addRecord(HWP_TAGS.HWPTAG_LIST_HEADER, level + 1, listBody.build());

  // Child paragraphs at level+2
  for (const para of paragraphs) {
    writeParagraph(para, level + 2, maps, stream);
  }
}

// ============================================================
// G. Footnote/Endnote Controls
// ============================================================

/**
 * Write all records for a footnote or endnote control.
 * CTRL_HEADER (fn/en) + LIST_HEADER + child paragraphs.
 *
 * Based on parser lines ~1407-1410: CTRL_ID.FOOTNOTE/ENDNOTE triggers inFootnoteEndnote mode,
 * followed by LIST_HEADER with paragraphs.
 */
function writeFootnoteEndnoteControl(
  fn: Footnote | Endnote,
  type: 'footnote' | 'endnote',
  level: number,
  maps: SectionWriterMaps,
  stream: TagStreamBuilder
): void {
  const ctrlId = type === 'footnote' ? CTRL_ID.FOOTNOTE : CTRL_ID.ENDNOTE;

  // CTRL_HEADER at `level`
  const ctrlBody = new BodyBuilder();
  ctrlBody.addUint32(ctrlId);
  // Properties: number (uint32) and flags
  ctrlBody.addUint32(fn.number || 1);
  // Additional properties: suffix/prefix type etc. - zero for defaults
  ctrlBody.addUint32(0);
  ctrlBody.addUint32(0);

  stream.addRecord(HWP_TAGS.HWPTAG_CTRL_HEADER, level, ctrlBody.build());

  // LIST_HEADER at level+1
  const paragraphs = fn.paragraphs.length > 0 ? fn.paragraphs : [{ id: 'empty-fn', runs: [{ text: '' }] }];
  const ps = ctx.pageSettings;
  const paragraphWidth = ptToHwpunit(ps.width - (ps.marginLeft || 0) - (ps.marginRight || 0));

  const listBody = new BodyBuilder();
  listBody.addUint16(paragraphs.length); // paraCount
  listBody.addUint32(0);                 // properties
  listBody.addUint16(0);                 // padding
  // Footnote/endnote specific: paragraphWidth
  listBody.addUint32(paragraphWidth);

  stream.addRecord(HWP_TAGS.HWPTAG_LIST_HEADER, level + 1, listBody.build());

  // Child paragraphs at level+2
  for (const para of paragraphs) {
    writeParagraph(para, level + 2, maps, stream);
  }
}

// ============================================================
// H. Drawing Object Controls (Line, Rectangle, Ellipse)
// ============================================================

/**
 * Write the common CTRL_HEADER + SHAPE_COMPONENT for a drawing object.
 * Returns the ctrlId used.
 */
function writeDrawingCtrlHeader(
  ctrlId: number,
  widthHwp: number,
  heightHwp: number,
  level: number,
  stream: TagStreamBuilder
): void {
  // CTRL_HEADER
  const ctrlBody = new BodyBuilder();
  ctrlBody.addUint32(ctrlId);
  ctrlBody.addUint32(1 << 4); // treatAsChar = true
  ctrlBody.addInt32(0);       // vertOffset
  ctrlBody.addInt32(0);       // horzOffset
  ctrlBody.addUint32(widthHwp);
  ctrlBody.addUint32(heightHwp);
  ctrlBody.addInt32(0);       // zOrder
  // OutMargin (4 x uint16)
  ctrlBody.addUint16(0);
  ctrlBody.addUint16(0);
  ctrlBody.addUint16(0);
  ctrlBody.addUint16(0);

  stream.addRecord(HWP_TAGS.HWPTAG_CTRL_HEADER, level, ctrlBody.build());

  // SHAPE_COMPONENT (76) at level+1
  const shapeBody = new BodyBuilder();
  shapeBody.addUint32(0);     // ShapeID
  shapeBody.addUint32(0);     // ComponentID
  shapeBody.addInt32(0);      // X grouping offset
  shapeBody.addInt32(0);      // Y grouping offset
  shapeBody.addUint16(0);     // Group level
  shapeBody.addUint16(0);     // Local file version
  shapeBody.addUint32(widthHwp);  // Original width
  shapeBody.addUint32(heightHwp); // Original height
  shapeBody.addUint32(widthHwp);  // Current width
  shapeBody.addUint32(heightHwp); // Current height
  shapeBody.addUint32(0);     // Rotation properties
  shapeBody.addInt32(widthHwp / 2);  // Rotation center X
  shapeBody.addInt32(heightHwp / 2); // Rotation center Y

  stream.addRecord(HWP_TAGS.HWPTAG_SHAPE_COMPONENT, level + 1, shapeBody.build());
}

/**
 * Write a LINE drawing control.
 * CTRL_HEADER + SHAPE_COMPONENT + SHAPE_COMPONENT_LINE (78)
 */
function writeLineControl(
  line: HwpxLine,
  level: number,
  stream: TagStreamBuilder
): void {
  const x1 = ptToHwpunit(line.startX ?? line.x1 ?? 0);
  const y1 = ptToHwpunit(line.startY ?? line.y1 ?? 0);
  const x2 = ptToHwpunit(line.endX ?? line.x2 ?? 100);
  const y2 = ptToHwpunit(line.endY ?? line.y2 ?? 0);
  const widthHwp = Math.abs(x2 - x1) || ptToHwpunit(100);
  const heightHwp = Math.abs(y2 - y1) || ptToHwpunit(1);

  writeDrawingCtrlHeader(CTRL_ID.LINE, widthHwp, heightHwp, level, stream);

  // SHAPE_COMPONENT_LINE (78) at level+1
  const lineBody = new BodyBuilder();
  lineBody.addInt32(x1);
  lineBody.addInt32(y1);
  lineBody.addInt32(x2);
  lineBody.addInt32(y2);
  // isReverseHV (uint16)
  lineBody.addUint16(line.isReverseHV ? 1 : 0);

  stream.addRecord(HWP_TAGS.HWPTAG_SHAPE_COMPONENT_LINE, level + 1, lineBody.build());
}

/**
 * Write a RECTANGLE drawing control.
 * CTRL_HEADER + SHAPE_COMPONENT + SHAPE_COMPONENT_RECTANGLE (79)
 */
function writeRectControl(
  rect: HwpxRect,
  level: number,
  stream: TagStreamBuilder
): void {
  const widthHwp = ptToHwpunit(rect.width ?? 100);
  const heightHwp = ptToHwpunit(rect.height ?? 50);

  writeDrawingCtrlHeader(CTRL_ID.RECTANGLE, widthHwp, heightHwp, level, stream);

  // SHAPE_COMPONENT_RECTANGLE (79) at level+1
  // 4 corner points + ratio
  const rectBody = new BodyBuilder();
  rectBody.addUint8(rect.ratio ?? 0); // ratio
  // 4 corner coordinates (x0,y0), (x1,y1), (x2,y2), (x3,y3)
  const x0 = ptToHwpunit(rect.x0 ?? rect.x ?? 0);
  const y0 = ptToHwpunit(rect.y0 ?? rect.y ?? 0);
  const x1 = ptToHwpunit(rect.x1 ?? (rect.x ?? 0) + (rect.width ?? 100));
  const y1 = ptToHwpunit(rect.y1 ?? rect.y ?? 0);
  const x2 = ptToHwpunit(rect.x2 ?? (rect.x ?? 0) + (rect.width ?? 100));
  const y2 = ptToHwpunit(rect.y2 ?? (rect.y ?? 0) + (rect.height ?? 50));
  const x3 = ptToHwpunit(rect.x3 ?? rect.x ?? 0);
  const y3 = ptToHwpunit(rect.y3 ?? (rect.y ?? 0) + (rect.height ?? 50));
  rectBody.addInt32(x0); rectBody.addInt32(y0);
  rectBody.addInt32(x1); rectBody.addInt32(y1);
  rectBody.addInt32(x2); rectBody.addInt32(y2);
  rectBody.addInt32(x3); rectBody.addInt32(y3);

  stream.addRecord(HWP_TAGS.HWPTAG_SHAPE_COMPONENT_RECTANGLE, level + 1, rectBody.build());
}

/**
 * Write an ELLIPSE drawing control.
 * CTRL_HEADER + SHAPE_COMPONENT + SHAPE_COMPONENT_ELLIPSE (80)
 */
function writeEllipseControl(
  ellipse: HwpxEllipse,
  level: number,
  stream: TagStreamBuilder
): void {
  const rx = ptToHwpunit(ellipse.rx ?? 50);
  const ry = ptToHwpunit(ellipse.ry ?? 50);
  const widthHwp = rx * 2;
  const heightHwp = ry * 2;

  writeDrawingCtrlHeader(CTRL_ID.ELLIPSE, widthHwp, heightHwp, level, stream);

  // SHAPE_COMPONENT_ELLIPSE (80) at level+1
  const ellipseBody = new BodyBuilder();
  // Flags (uint32): intervalDirty, hasArcProperty, arcType
  let flags = 0;
  if (ellipse.intervalDirty) flags |= 1;
  if (ellipse.hasArcProperty) flags |= 2;
  ellipseBody.addUint32(flags);
  // Center
  ellipseBody.addInt32(ptToHwpunit(ellipse.centerX ?? ellipse.cx ?? 0));
  ellipseBody.addInt32(ptToHwpunit(ellipse.centerY ?? ellipse.cy ?? 0));
  // Axis1
  ellipseBody.addInt32(ptToHwpunit(ellipse.axis1X ?? (ellipse.rx ?? 50)));
  ellipseBody.addInt32(ptToHwpunit(ellipse.axis1Y ?? 0));
  // Axis2
  ellipseBody.addInt32(ptToHwpunit(ellipse.axis2X ?? 0));
  ellipseBody.addInt32(ptToHwpunit(ellipse.axis2Y ?? (ellipse.ry ?? 50)));

  stream.addRecord(HWP_TAGS.HWPTAG_SHAPE_COMPONENT_ELLIPSE, level + 1, ellipseBody.build());
}

// ============================================================
// I. Textbox Control
// ============================================================

/**
 * Write a textbox control.
 * CTRL_HEADER + SHAPE_COMPONENT + SHAPE_COMPONENT_TEXTBOX (87) + LIST_HEADER + child paragraphs
 */
function writeTextboxControl(
  textbox: HwpxTextBox,
  level: number,
  maps: SectionWriterMaps,
  stream: TagStreamBuilder
): void {
  const widthHwp = ptToHwpunit(textbox.width);
  const heightHwp = ptToHwpunit(textbox.height);

  writeDrawingCtrlHeader(CTRL_ID.TEXTBOX, widthHwp, heightHwp, level, stream);

  // SHAPE_COMPONENT_TEXTBOX (87) at level+1
  const tbBody = new BodyBuilder();
  // Text margin (left, right, top, bottom as uint16)
  tbBody.addUint16(ptToHwpunit(1.4)); // default margin
  tbBody.addUint16(ptToHwpunit(1.4));
  tbBody.addUint16(ptToHwpunit(1.4));
  tbBody.addUint16(ptToHwpunit(1.4));
  // lastWidth (uint32)
  tbBody.addUint32(widthHwp);

  stream.addRecord(HWP_TAGS.HWPTAG_SHAPE_COMPONENT_TEXTBOX, level + 1, tbBody.build());

  // LIST_HEADER at level+1 for text content
  const paragraphs = textbox.paragraphs.length > 0
    ? textbox.paragraphs
    : [{ id: 'empty-tb', runs: [{ text: '' }] }];

  const listBody = new BodyBuilder();
  listBody.addUint16(paragraphs.length); // paraCount
  listBody.addUint32(0);                 // properties
  listBody.addUint16(0);                 // padding
  // textbox specific: textWidth, textHeight
  listBody.addUint32(widthHwp);
  listBody.addUint32(heightHwp);

  stream.addRecord(HWP_TAGS.HWPTAG_LIST_HEADER, level + 1, listBody.build());

  // Child paragraphs at level+2
  for (const para of paragraphs) {
    writeParagraph(para, level + 2, maps, stream);
  }
}

// ============================================================
// J. Equation Control
// ============================================================

/**
 * Write an equation control.
 * Equations in HWP are written as CTRL_HEADER (eqed) + SHAPE_COMPONENT + textbox-like
 * text content containing the equation script.
 *
 * For safety, we write the equation script as a paragraph inside a shape text context.
 */
function writeEquationControl(
  equation: HwpxEquation,
  level: number,
  maps: SectionWriterMaps,
  stream: TagStreamBuilder
): void {
  const script = equation.script || '';
  // Default equation dimensions
  const widthHwp = ptToHwpunit(200);
  const heightHwp = ptToHwpunit(30);

  // CTRL_HEADER for equation
  const ctrlBody = new BodyBuilder();
  ctrlBody.addUint32(CTRL_ID.EQUATION);
  ctrlBody.addUint32(1 << 4); // treatAsChar = true
  ctrlBody.addInt32(0);       // vertOffset
  ctrlBody.addInt32(0);       // horzOffset
  ctrlBody.addUint32(widthHwp);
  ctrlBody.addUint32(heightHwp);
  ctrlBody.addInt32(0);       // zOrder
  // OutMargin
  ctrlBody.addUint16(0);
  ctrlBody.addUint16(0);
  ctrlBody.addUint16(0);
  ctrlBody.addUint16(0);

  stream.addRecord(HWP_TAGS.HWPTAG_CTRL_HEADER, level, ctrlBody.build());

  // SHAPE_COMPONENT at level+1
  const shapeBody = new BodyBuilder();
  shapeBody.addUint32(0);     // ShapeID
  shapeBody.addUint32(0);     // ComponentID
  shapeBody.addInt32(0);
  shapeBody.addInt32(0);
  shapeBody.addUint16(0);     // Group level
  shapeBody.addUint16(0);     // Local file version
  shapeBody.addUint32(widthHwp);
  shapeBody.addUint32(heightHwp);
  shapeBody.addUint32(widthHwp);
  shapeBody.addUint32(heightHwp);
  shapeBody.addUint32(0);
  shapeBody.addInt32(widthHwp / 2);
  shapeBody.addInt32(heightHwp / 2);

  stream.addRecord(HWP_TAGS.HWPTAG_SHAPE_COMPONENT, level + 1, shapeBody.build());

  // Write the equation script as a simple paragraph at level+1
  // The parser reads equation text from shape text paragraphs
  const eqPara: HwpxParagraph = { id: 'eq-' + equation.id, runs: [{ text: script }] };
  writeParagraph(eqPara, level + 1, maps, stream);
}

// ============================================================
// K. Main Assembly - buildSectionStream
// ============================================================

/**
 * Build a complete section stream from an HwpxSection.
 * This is the reverse of parseSectionData().
 *
 * Structure:
 * 1. First paragraph: section definition (secd + cold controls) + PAGE_DEF
 * 2. For each element in section.elements:
 *    - paragraph: PARA_HEADER + PARA_TEXT + PARA_CHAR_SHAPE + PARA_LINE_SEG
 *    - table: host paragraph + CTRL_HEADER(tbl) + TABLE + cells(LIST_HEADER + cell paragraphs)
 *    - image: host paragraph + CTRL_HEADER($pic) + SHAPE_COMPONENT + SHAPE_COMPONENT_PICTURE
 */
export function buildSectionStream(
  section: HwpxSection,
  maps: SectionWriterMaps
): Uint8Array {
  // Reset context for each section
  ctx = {
    instanceIdCounter: 0,
    pageSettings: section.pageSettings || {
      width: 595, height: 842,
      marginTop: 56.7, marginBottom: 56.7,
      marginLeft: 56.7, marginRight: 56.7,
    },
  };

  const stream = new TagStreamBuilder();
  const pageSettings = ctx.pageSettings;

  // The secd/cold ctrlIds to embed as inline controls in the first paragraph.
  // In real HWP files, secd/cold are inline in the FIRST paragraph (not a separate paragraph).
  const sectionCtrlIds = [CTRL_ID.SECTION, CTRL_ID.COLUMN];

  // In real HWP files, tables and images are NOT written as separate host paragraphs.
  // They are embedded as children of the PRECEDING paragraph: the preceding paragraph's
  // PARA_TEXT gets a 0x000B control char appended for each trailing table/image,
  // and then the CTRL_HEADER records follow at level+1 (sibling of PARA_TXT).
  //
  // Strategy: collect consecutive table/image elements that follow a paragraph and
  // attach them as trailing controls to that paragraph.
  // If a table/image appears without a preceding paragraph (e.g. first element),
  // emit a minimal empty paragraph to host it.

  const elements = section.elements;
  let i = 0;
  let isFirst = true;

  while (i < elements.length) {
    const element = elements[i];
    const sectionCtrls = isFirst ? sectionCtrlIds : undefined;
    isFirst = false;

    if (element.type === 'paragraph') {
      // Look ahead: collect any immediately following tables/images
      const trailingTables: HwpxTable[] = [];
      const trailingImages: HwpxImage[] = [];
      let j = i + 1;
      while (j < elements.length && (elements[j].type === 'table' || elements[j].type === 'image')) {
        if (elements[j].type === 'table') trailingTables.push(elements[j].data as HwpxTable);
        else trailingImages.push(elements[j].data as HwpxImage);
        j++;
      }

      // Write the paragraph with all trailing tables/images as children
      writeParagraphWithTrailing(
        element.data as HwpxParagraph,
        0, maps, stream, sectionCtrls,
        trailingTables, trailingImages,
      );
      i = j;

    } else if (element.type === 'table' || element.type === 'image') {
      // Table/image without a preceding paragraph — collect consecutive table/images
      const trailingTables: HwpxTable[] = [];
      const trailingImages: HwpxImage[] = [];
      let j = i;
      while (j < elements.length && (elements[j].type === 'table' || elements[j].type === 'image')) {
        if (elements[j].type === 'table') trailingTables.push(elements[j].data as HwpxTable);
        else trailingImages.push(elements[j].data as HwpxImage);
        j++;
      }
      // Emit an empty host paragraph with all trailing tables/images
      writeParagraphWithTrailing(
        { id: 'host', runs: [{ text: '' }] },
        0, maps, stream, sectionCtrls,
        trailingTables, trailingImages,
      );
      i = j;

    } else if (element.type === 'line') {
      // Line drawing object — emit as host paragraph with inline control
      if (sectionCtrls) {
        writeParagraph({ id: 'sectiondef', runs: [{ text: '' }] }, 0, maps, stream, sectionCtrls);
      }
      writeDrawingHostParagraph(CTRL_ID.LINE, 0, stream);
      writeLineControl(element.data as HwpxLine, 1, stream);
      i++;

    } else if (element.type === 'rect') {
      if (sectionCtrls) {
        writeParagraph({ id: 'sectiondef', runs: [{ text: '' }] }, 0, maps, stream, sectionCtrls);
      }
      writeDrawingHostParagraph(CTRL_ID.RECTANGLE, 0, stream);
      writeRectControl(element.data as HwpxRect, 1, stream);
      i++;

    } else if (element.type === 'ellipse') {
      if (sectionCtrls) {
        writeParagraph({ id: 'sectiondef', runs: [{ text: '' }] }, 0, maps, stream, sectionCtrls);
      }
      writeDrawingHostParagraph(CTRL_ID.ELLIPSE, 0, stream);
      writeEllipseControl(element.data as HwpxEllipse, 1, stream);
      i++;

    } else if (element.type === 'textbox') {
      if (sectionCtrls) {
        writeParagraph({ id: 'sectiondef', runs: [{ text: '' }] }, 0, maps, stream, sectionCtrls);
      }
      writeDrawingHostParagraph(CTRL_ID.TEXTBOX, 0, stream);
      writeTextboxControl(element.data as HwpxTextBox, 1, maps, stream);
      i++;

    } else if (element.type === 'equation') {
      if (sectionCtrls) {
        writeParagraph({ id: 'sectiondef', runs: [{ text: '' }] }, 0, maps, stream, sectionCtrls);
      }
      writeDrawingHostParagraph(CTRL_ID.EQUATION, 0, stream);
      writeEquationControl(element.data as HwpxEquation, 1, maps, stream);
      i++;

    } else {
      // Unsupported element type — skip but emit section def if needed
      if (sectionCtrls) {
        writeParagraph({ id: 'sectiondef', runs: [{ text: '' }] }, 0, maps, stream, sectionCtrls);
      }
      i++;
    }
  }

  // If section was empty, emit one paragraph with secd/cold
  if (isFirst) {
    writeParagraph({ id: 'empty', runs: [{ text: '' }] }, 0, maps, stream, sectionCtrlIds);
    writeParagraph({ id: 'empty2', runs: [{ text: '' }] }, 0, maps, stream);
  }

  // Emit header/footer controls after all section content.
  // In HWP binary format, headers/footers appear at level 0, the same level
  // as section content paragraphs. The LIST_HEADER goes at level+1=1, and
  // child paragraphs at level+2=2.
  if (section.header && section.header.paragraphs.length > 0) {
    writeHeaderFooterControl(section.header, 'header', 0, maps, stream);
  }
  if (section.footer && section.footer.paragraphs.length > 0) {
    writeHeaderFooterControl(section.footer, 'footer', 0, maps, stream);
  }

  return stream.build();
}

// ============================================================
// Helper: Drawing host paragraph
// ============================================================

/**
 * Write a minimal host paragraph for a drawing object (line, rect, ellipse, textbox, equation).
 * Similar to writeTableHostParagraph/writeImageHostParagraph but for drawing ctrl IDs.
 * Uses 0x000B inline control char (same as tables/images for shape objects).
 */
function writeDrawingHostParagraph(
  ctrlId: number,
  level: number,
  stream: TagStreamBuilder,
): void {
  const body = new BodyBuilder();
  let charCount = 0;
  const controlMask = (1 << 11); // bit 11 for drawing objects

  // Drawing object inline control (char 0x000B)
  writeInlineControl(body, 0x000B, ctrlId);
  charCount += 8;
  // Paragraph break
  body.addUint16(0x000D);
  charCount++;

  writeParaHeader(charCount, controlMask, 0, 0, 1, 0, level, stream);
  stream.addRecord(HWP_TAGS.HWPTAG_PARA_TEXT, level + 1, body.build());
  writeParaCharShape([{ pos: 0, id: 0 }], level, stream);
  writeParaLineSeg(level, stream);
}
