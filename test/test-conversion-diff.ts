/**
 * Comprehensive HWPX → HWP round-trip conversion diff test.
 * Tests every field: text, char styles, para styles, tables, images,
 * footnotes, headers/footers, page settings, and metadata.
 */

import * as fs from 'fs';
import * as path from 'path';
import * as JSZip from 'jszip';
import { HwpxParser } from '../shared/src/HwpxParser';
import { writeHwpContent } from '../shared/src/HwpWriter';
import { parseHwpContent } from '../shared/src/HwpParser';
import {
  HwpxContent,
  HwpxParagraph,
  TextRun,
  HwpxTable,
  TableRow,
  TableCell,
  HwpxImage,
  CharacterStyle,
  ParagraphStyle,
  PageSettings,
  Footnote,
} from '../shared/src/types';

// ============================================================
// Diff utilities
// ============================================================

interface Diff {
  path: string;
  original: unknown;
  roundtrip: unknown;
  note?: string;
}

const allDiffs: Diff[] = [];
let currentFile = '';

function diff(pathStr: string, original: unknown, roundtrip: unknown, note?: string) {
  allDiffs.push({ path: `[${currentFile}] ${pathStr}`, original, roundtrip, note });
}

function cmpNum(pathStr: string, a: number | undefined, b: number | undefined, tol = 0.5) {
  if (a === undefined && b === undefined) return;
  if (a === undefined || b === undefined) {
    diff(pathStr, a, b, 'one side undefined');
    return;
  }
  if (Math.abs(a - b) > tol) diff(pathStr, a, b);
}

function cmpStr(pathStr: string, a: string | undefined, b: string | undefined) {
  if (a !== b) diff(pathStr, a, b);
}

function cmpBool(pathStr: string, a: boolean | undefined, b: boolean | undefined) {
  if (!!a !== !!b) diff(pathStr, a, b);
}

// ============================================================
// Paragraph text extraction
// ============================================================

function paraText(para: HwpxParagraph): string {
  return para.runs.map((r: TextRun) => r.text).join('');
}

function sectionText(content: HwpxContent): string {
  const lines: string[] = [];
  for (const section of content.sections) {
    for (const el of section.elements) {
      if (el.type === 'paragraph') {
        lines.push(paraText(el.data as HwpxParagraph));
      } else if (el.type === 'table') {
        const tbl = el.data as HwpxTable;
        for (const row of tbl.rows) {
          for (const cell of row.cells) {
            if (cell.paragraphs) {
              for (const p of cell.paragraphs) lines.push(paraText(p));
            }
          }
        }
      }
    }
  }
  return lines.join('\n');
}

// ============================================================
// Character style comparison
// ============================================================

function compareCharStyle(prefix: string, a: CharacterStyle | undefined, b: CharacterStyle | undefined) {
  if (!a && !b) return;
  if (!a || !b) { diff(prefix, a ? 'defined' : 'undefined', b ? 'defined' : 'undefined'); return; }

  cmpStr(`${prefix}.fontName`, a.fontName, b.fontName);
  cmpNum(`${prefix}.fontSize`, a.fontSize, b.fontSize, 0.1);
  cmpBool(`${prefix}.bold`, a.bold, b.bold);
  cmpBool(`${prefix}.italic`, a.italic, b.italic);
  cmpBool(`${prefix}.underline`, a.underline, b.underline);
  cmpBool(`${prefix}.strikethrough`, a.strikethrough, b.strikethrough);
  cmpBool(`${prefix}.superscript`, a.superscript, b.superscript);
  cmpBool(`${prefix}.subscript`, a.subscript, b.subscript);
  cmpBool(`${prefix}.emboss`, a.emboss, b.emboss);
  cmpBool(`${prefix}.engrave`, a.engrave, b.engrave);
  // Color: normalize to uppercase for comparison
  const normColor = (c?: string) => c?.toUpperCase().replace('#', '');
  if (normColor(a.fontColor) !== normColor(b.fontColor)) {
    diff(`${prefix}.fontColor`, a.fontColor, b.fontColor);
  }
  cmpBool(`${prefix}.useKerning`, a.useKerning, b.useKerning);
  cmpBool(`${prefix}.useFontSpace`, a.useFontSpace, b.useFontSpace);
}

// ============================================================
// Paragraph style comparison
// ============================================================

function compareParaStyle(prefix: string, a: ParagraphStyle | undefined, b: ParagraphStyle | undefined) {
  if (!a && !b) return;
  if (!a || !b) { diff(prefix, a ? 'defined' : 'undefined', b ? 'defined' : 'undefined'); return; }

  cmpStr(`${prefix}.align`, a.align, b.align);
  cmpNum(`${prefix}.marginLeft`, a.marginLeft, b.marginLeft);
  cmpNum(`${prefix}.marginRight`, a.marginRight, b.marginRight);
  cmpNum(`${prefix}.marginTop`, a.marginTop, b.marginTop);
  cmpNum(`${prefix}.marginBottom`, a.marginBottom, b.marginBottom);
  cmpNum(`${prefix}.firstLineIndent`, a.firstLineIndent, b.firstLineIndent);
  cmpNum(`${prefix}.lineSpacing`, a.lineSpacing, b.lineSpacing, 1);
  cmpStr(`${prefix}.lineSpacingType`, a.lineSpacingType, b.lineSpacingType);
  cmpBool(`${prefix}.widowControl`, a.widowControl, b.widowControl);
  cmpBool(`${prefix}.keepWithNext`, a.keepWithNext, b.keepWithNext);
  cmpBool(`${prefix}.keepLines`, a.keepLines, b.keepLines);
  cmpBool(`${prefix}.pageBreakBefore`, a.pageBreakBefore, b.pageBreakBefore);
  cmpBool(`${prefix}.snapToGrid`, a.snapToGrid, b.snapToGrid);
  cmpBool(`${prefix}.autoSpaceEAsianEng`, a.autoSpaceEAsianEng, b.autoSpaceEAsianEng);
  cmpBool(`${prefix}.autoSpaceEAsianNum`, a.autoSpaceEAsianNum, b.autoSpaceEAsianNum);
}

// ============================================================
// Paragraph comparison (runs)
// ============================================================

function compareParagraph(prefix: string, a: HwpxParagraph, b: HwpxParagraph) {
  const textA = paraText(a);
  const textB = paraText(b);

  if (textA !== textB) {
    // Show first difference character position
    let diffPos = -1;
    for (let i = 0; i < Math.max(textA.length, textB.length); i++) {
      if (textA[i] !== textB[i]) { diffPos = i; break; }
    }
    diff(`${prefix}.text`, textA.substring(0, 80), textB.substring(0, 80),
      `first diff at char ${diffPos}; len orig=${textA.length} rt=${textB.length}`);
  }

  // Compare para style
  compareParaStyle(`${prefix}.paraStyle`, a.paraStyle, b.paraStyle);

  // Compare first run char style (representative)
  if (a.runs.length > 0 || b.runs.length > 0) {
    if (a.runs.length !== b.runs.length) {
      diff(`${prefix}.runCount`, a.runs.length, b.runs.length);
    }
    // Compare first non-empty run
    const firstA = a.runs.find((r: TextRun) => r.text.length > 0);
    const firstB = b.runs.find((r: TextRun) => r.text.length > 0);
    if (firstA || firstB) {
      compareCharStyle(`${prefix}.run[0].charStyle`, firstA?.charStyle, firstB?.charStyle);
    }
  }
}

// ============================================================
// Table comparison
// ============================================================

function compareTable(prefix: string, a: HwpxTable, b: HwpxTable) {
  if (a.rows.length !== b.rows.length) {
    diff(`${prefix}.rowCount`, a.rows.length, b.rows.length);
  }
  cmpNum(`${prefix}.colCount`, a.colCount, b.colCount, 0);
  cmpNum(`${prefix}.width`, a.width, b.width);
  // cmpNum(`${prefix}.height`, a.height, b.height);  // height often recalculated

  const rowCount = Math.min(a.rows.length, b.rows.length);
  for (let r = 0; r < rowCount; r++) {
    const rowA = a.rows[r];
    const rowB = b.rows[r];
    if (rowA.cells.length !== rowB.cells.length) {
      diff(`${prefix}.row[${r}].cellCount`, rowA.cells.length, rowB.cells.length);
    }
    const cellCount = Math.min(rowA.cells.length, rowB.cells.length);
    for (let c = 0; c < cellCount; c++) {
      const cellA = rowA.cells[c];
      const cellB = rowB.cells[c];
      const cp = `${prefix}.row[${r}].cell[${c}]`;

      cmpNum(`${cp}.colSpan`, cellA.colSpan, cellB.colSpan, 0);
      cmpNum(`${cp}.rowSpan`, cellA.rowSpan, cellB.rowSpan, 0);
      cmpNum(`${cp}.width`, cellA.width, cellB.width);
      cmpStr(`${cp}.verticalAlign`, cellA.verticalAlign, cellB.verticalAlign);
      // cmpStr(`${cp}.backgroundColor`, cellA.backgroundColor, cellB.backgroundColor);

      // Cell text content
      const cellTextA = (cellA.paragraphs || []).map(paraText).join('\n');
      const cellTextB = (cellB.paragraphs || []).map(paraText).join('\n');
      if (cellTextA !== cellTextB) {
        diff(`${cp}.text`, cellTextA.substring(0, 80), cellTextB.substring(0, 80));
      }

      // Cell border styles
      const checkBorder = (side: 'borderTop' | 'borderBottom' | 'borderLeft' | 'borderRight') => {
        const ba = (cellA as any)[side];
        const bb = (cellB as any)[side];
        if (JSON.stringify(ba) !== JSON.stringify(bb)) {
          diff(`${cp}.${side}`, ba, bb);
        }
      };
      checkBorder('borderTop');
      checkBorder('borderBottom');
      checkBorder('borderLeft');
      checkBorder('borderRight');
    }
  }
}

// ============================================================
// Image comparison
// ============================================================

function compareImages(prefix: string, a: HwpxContent, b: HwpxContent) {
  // Count images in section elements
  const imgsA: HwpxImage[] = [];
  const imgsB: HwpxImage[] = [];

  for (const section of a.sections) {
    for (const el of section.elements) {
      if (el.type === 'image') imgsA.push(el.data as HwpxImage);
    }
  }
  for (const section of b.sections) {
    for (const el of section.elements) {
      if (el.type === 'image') imgsB.push(el.data as HwpxImage);
    }
  }

  if (imgsA.length !== imgsB.length) {
    diff(`${prefix}.imageCount`, imgsA.length, imgsB.length);
  }

  const count = Math.min(imgsA.length, imgsB.length);
  for (let i = 0; i < count; i++) {
    const ia = imgsA[i];
    const ib = imgsB[i];
    cmpNum(`${prefix}.image[${i}].width`, ia.width, ib.width);
    cmpNum(`${prefix}.image[${i}].height`, ia.height, ib.height);
    const hasDataA = !!ia.data;
    const hasDataB = !!ib.data;
    if (hasDataA !== hasDataB) {
      diff(`${prefix}.image[${i}].hasData`, hasDataA, hasDataB);
    }
  }

  // Binary images map
  const imgMapA = a.images;
  const imgMapB = b.images;
  if (imgMapA.size !== imgMapB.size) {
    diff(`${prefix}.images.mapSize`, imgMapA.size, imgMapB.size);
  }
}

// ============================================================
// Page settings comparison
// ============================================================

function comparePageSettings(prefix: string, a: PageSettings | undefined, b: PageSettings | undefined) {
  if (!a && !b) return;
  if (!a || !b) { diff(prefix, a ? 'defined' : 'undefined', b ? 'defined' : 'undefined'); return; }

  cmpNum(`${prefix}.width`, a.width, b.width);
  cmpNum(`${prefix}.height`, a.height, b.height);
  cmpNum(`${prefix}.marginTop`, a.marginTop, b.marginTop);
  cmpNum(`${prefix}.marginBottom`, a.marginBottom, b.marginBottom);
  cmpNum(`${prefix}.marginLeft`, a.marginLeft, b.marginLeft);
  cmpNum(`${prefix}.marginRight`, a.marginRight, b.marginRight);
  cmpNum(`${prefix}.headerMargin`, a.headerMargin, b.headerMargin);
  cmpNum(`${prefix}.footerMargin`, a.footerMargin, b.footerMargin);
  cmpStr(`${prefix}.orientation`, a.orientation, b.orientation);
}

// ============================================================
// Header/Footer comparison
// ============================================================

function compareHeaderFooter(prefix: string, a: any, b: any) {
  if (!a && !b) return;
  if (!a || !b) { diff(prefix, a ? 'defined' : 'undefined', b ? 'defined' : 'undefined'); return; }

  if (a.paragraphs && b.paragraphs) {
    if (a.paragraphs.length !== b.paragraphs.length) {
      diff(`${prefix}.paragraphCount`, a.paragraphs.length, b.paragraphs.length);
    }
    const count = Math.min(a.paragraphs.length, b.paragraphs.length);
    for (let i = 0; i < count; i++) {
      compareParagraph(`${prefix}.para[${i}]`, a.paragraphs[i], b.paragraphs[i]);
    }
  }
}

// ============================================================
// Footnote/Endnote comparison
// ============================================================

function compareFootnotes(prefix: string, a: Footnote[], b: Footnote[]) {
  if (a.length !== b.length) {
    diff(`${prefix}.count`, a.length, b.length, 'footnote/endnote count');
    return;
  }
  for (let i = 0; i < a.length; i++) {
    const textA = (a[i].paragraphs || []).map(paraText).join('\n');
    const textB = (b[i].paragraphs || []).map(paraText).join('\n');
    if (textA !== textB) {
      diff(`${prefix}[${i}].text`, textA.substring(0, 80), textB.substring(0, 80));
    }
  }
}

// ============================================================
// Main comparison function
// ============================================================

function compareContent(orig: HwpxContent, rt: HwpxContent, label: string) {
  const prefix = label;

  // --- Section count ---
  if (orig.sections.length !== rt.sections.length) {
    diff(`${prefix}.sectionCount`, orig.sections.length, rt.sections.length);
  }

  const sectionCount = Math.min(orig.sections.length, rt.sections.length);

  for (let si = 0; si < sectionCount; si++) {
    const secA = orig.sections[si];
    const secB = rt.sections[si];
    const sp = `${prefix}.section[${si}]`;

    // --- Page settings ---
    comparePageSettings(`${sp}.pageSettings`, secA.pageSettings, secB.pageSettings);

    // --- Header / Footer ---
    compareHeaderFooter(`${sp}.header`, secA.header, secB.header);
    compareHeaderFooter(`${sp}.footer`, secA.footer, secB.footer);

    // --- Element count ---
    const elemsA = secA.elements;
    const elemsB = secB.elements;

    // Count by type
    const countType = (elems: any[], t: string) => elems.filter((e: any) => e.type === t).length;
    const paraCountA = countType(elemsA, 'paragraph');
    const paraCountB = countType(elemsB, 'paragraph');
    const tableCountA = countType(elemsA, 'table');
    const tableCountB = countType(elemsB, 'table');
    const imgCountA = countType(elemsA, 'image');
    const imgCountB = countType(elemsB, 'image');

    if (paraCountA !== paraCountB) {
      diff(`${sp}.paragraphCount`, paraCountA, paraCountB);
    }
    if (tableCountA !== tableCountB) {
      diff(`${sp}.tableCount`, tableCountA, tableCountB);
    }
    if (imgCountA !== imgCountB) {
      diff(`${sp}.imageCount`, imgCountA, imgCountB);
    }

    // --- Paragraph-by-paragraph comparison ---
    // Align paragraphs (skip non-paragraph elements for sequential comparison)
    const parasA = elemsA.filter((e: any) => e.type === 'paragraph').map((e: any) => e.data as HwpxParagraph);
    const parasB = elemsB.filter((e: any) => e.type === 'paragraph').map((e: any) => e.data as HwpxParagraph);

    const paraCount = Math.min(parasA.length, parasB.length);
    // Compare up to 200 paragraphs per section to keep output manageable
    const maxPara = Math.min(paraCount, 200);
    for (let pi = 0; pi < maxPara; pi++) {
      compareParagraph(`${sp}.para[${pi}]`, parasA[pi], parasB[pi]);
    }

    // --- Table comparison ---
    const tablesA = elemsA.filter((e: any) => e.type === 'table').map((e: any) => e.data as HwpxTable);
    const tablesB = elemsB.filter((e: any) => e.type === 'table').map((e: any) => e.data as HwpxTable);
    const tableCount = Math.min(tablesA.length, tablesB.length);
    for (let ti = 0; ti < tableCount; ti++) {
      compareTable(`${sp}.table[${ti}]`, tablesA[ti], tablesB[ti]);
    }
  }

  // --- Images ---
  compareImages(prefix, orig, rt);

  // --- Footnotes/Endnotes ---
  compareFootnotes(`${prefix}.footnotes`, orig.footnotes || [], rt.footnotes || []);
  compareFootnotes(`${prefix}.endnotes`, orig.endnotes || [], rt.endnotes || []);

  // --- Metadata ---
  const metaFields = ['title', 'creator', 'description', 'keywords', 'subject', 'lastModifiedBy'] as const;
  for (const field of metaFields) {
    if ((orig.metadata as any)[field] !== (rt.metadata as any)[field]) {
      diff(`${prefix}.metadata.${field}`, (orig.metadata as any)[field], (rt.metadata as any)[field]);
    }
  }
}

// ============================================================
// Statistics helper
// ============================================================

function getStats(content: HwpxContent) {
  let totalParas = 0, totalRuns = 0, totalChars = 0, totalTables = 0;
  let totalRows = 0, totalCells = 0, totalImages = 0;

  for (const section of content.sections) {
    for (const el of section.elements) {
      if (el.type === 'paragraph') {
        const para = el.data as HwpxParagraph;
        totalParas++;
        totalRuns += para.runs.length;
        totalChars += para.runs.reduce((s: number, r: TextRun) => s + r.text.length, 0);
      } else if (el.type === 'table') {
        const tbl = el.data as HwpxTable;
        totalTables++;
        for (const row of tbl.rows) {
          totalRows++;
          for (const cell of row.cells) {
            totalCells++;
            if (cell.paragraphs) {
              for (const p of cell.paragraphs) {
                totalParas++;
                totalChars += p.runs.reduce((s: number, r: TextRun) => s + r.text.length, 0);
              }
            }
          }
        }
      } else if (el.type === 'image') {
        totalImages++;
      }
    }
  }

  return { sections: content.sections.length, totalParas, totalRuns, totalChars, totalTables, totalRows, totalCells, totalImages, footnotes: (content.footnotes || []).length, endnotes: (content.endnotes || []).length };
}

// ============================================================
// File test runner
// ============================================================

async function testFile(hwpxPath: string): Promise<{ file: string; diffs: Diff[]; stats: any; error?: string }> {
  const fileName = path.basename(hwpxPath);
  const startDiffCount = allDiffs.length;
  currentFile = fileName;

  console.log(`\n${'='.repeat(70)}`);
  console.log(`FILE: ${fileName}`);
  console.log(`${'='.repeat(70)}`);

  // 1. Parse HWPX
  let orig: HwpxContent;
  try {
    const data = fs.readFileSync(hwpxPath);
    const zip = await JSZip.loadAsync(data);
    orig = await HwpxParser.parse(zip);
    console.log(`  [HWPX parsed] sections=${orig.sections.length}`);
  } catch (e: any) {
    console.log(`  [FAIL] HWPX parse: ${e.message}`);
    return { file: fileName, diffs: [], stats: null, error: `HWPX parse: ${e.message}` };
  }

  const origStats = getStats(orig);
  console.log(`  Original: paras=${origStats.totalParas} chars=${origStats.totalChars} tables=${origStats.totalTables} images=${origStats.totalImages} footnotes=${origStats.footnotes}`);

  // 2. Write HWP
  let hwpData: Uint8Array;
  try {
    hwpData = writeHwpContent(orig);
    console.log(`  [HWP written] ${hwpData.length} bytes`);
  } catch (e: any) {
    console.log(`  [FAIL] HWP write: ${e.message}`);
    return { file: fileName, diffs: [], stats: origStats, error: `HWP write: ${e.message}` };
  }

  // 3. Re-parse HWP
  let roundtrip: HwpxContent;
  try {
    roundtrip = parseHwpContent(hwpData);
    console.log(`  [HWP re-parsed] sections=${roundtrip.sections.length}`);
  } catch (e: any) {
    console.log(`  [FAIL] HWP re-parse: ${e.message}`);
    return { file: fileName, diffs: [], stats: origStats, error: `HWP re-parse: ${e.message}` };
  }

  const rtStats = getStats(roundtrip);
  console.log(`  Roundtrip: paras=${rtStats.totalParas} chars=${rtStats.totalChars} tables=${rtStats.totalTables} images=${rtStats.totalImages} footnotes=${rtStats.footnotes}`);

  // 4. Compare
  compareContent(orig, roundtrip, fileName);

  const newDiffs = allDiffs.slice(startDiffCount);
  console.log(`  Diffs found: ${newDiffs.length}`);

  return { file: fileName, diffs: newDiffs, stats: { original: origStats, roundtrip: rtStats } };
}

// ============================================================
// Report writer
// ============================================================

function printReport(results: Array<{ file: string; diffs: Diff[]; stats: any; error?: string }>) {
  console.log('\n' + '='.repeat(70));
  console.log('COMPREHENSIVE DIFF REPORT');
  console.log('='.repeat(70));

  // Group diffs by category
  const categories: Record<string, Diff[]> = {};
  for (const result of results) {
    for (const d of result.diffs) {
      // Extract category from path
      let cat = 'other';
      if (d.path.includes('.text')) cat = 'text_content';
      else if (d.path.includes('.fontSize') || d.path.includes('.fontName') || d.path.includes('.bold') || d.path.includes('.italic') || d.path.includes('.fontColor') || d.path.includes('.underline') || d.path.includes('.charStyle')) cat = 'char_style';
      else if (d.path.includes('.align') || d.path.includes('paraStyle') || d.path.includes('.lineSpacing') || d.path.includes('margin')) cat = 'para_style';
      else if (d.path.includes('.table') || d.path.includes('.row') || d.path.includes('.cell')) cat = 'tables';
      else if (d.path.includes('.image') || d.path.includes('images')) cat = 'images';
      else if (d.path.includes('footnote') || d.path.includes('endnote')) cat = 'footnotes_endnotes';
      else if (d.path.includes('header') || d.path.includes('footer')) cat = 'headers_footers';
      else if (d.path.includes('pageSettings')) cat = 'page_settings';
      else if (d.path.includes('metadata')) cat = 'metadata';
      else if (d.path.includes('Count') || d.path.includes('count')) cat = 'element_counts';
      (categories[cat] = categories[cat] || []).push(d);
    }
  }

  // Print summary by category
  console.log('\n--- DIFF SUMMARY BY CATEGORY ---');
  for (const [cat, diffs] of Object.entries(categories).sort()) {
    console.log(`\n[${cat.toUpperCase()}] (${diffs.length} diffs)`);
    // Print up to 30 diffs per category
    const shown = diffs.slice(0, 30);
    for (const d of shown) {
      const origStr = JSON.stringify(d.original)?.substring(0, 60) ?? 'undefined';
      const rtStr = JSON.stringify(d.roundtrip)?.substring(0, 60) ?? 'undefined';
      console.log(`  ${d.path}`);
      console.log(`    ORIG: ${origStr}`);
      console.log(`    RT:   ${rtStr}`);
      if (d.note) console.log(`    NOTE: ${d.note}`);
    }
    if (diffs.length > 30) {
      console.log(`  ... and ${diffs.length - 30} more`);
    }
  }

  // Per-file summary
  console.log('\n--- PER-FILE SUMMARY ---');
  for (const result of results) {
    console.log(`\n${result.file}:`);
    if (result.error) {
      console.log(`  ERROR: ${result.error}`);
      continue;
    }
    console.log(`  Total diffs: ${result.diffs.length}`);
    if (result.stats) {
      const o = result.stats.original;
      const r = result.stats.roundtrip;
      console.log(`  Original:  sections=${o?.sections} paras=${o?.totalParas} chars=${o?.totalChars} tables=${o?.totalTables} images=${o?.totalImages} footnotes=${o?.footnotes}`);
      console.log(`  Roundtrip: sections=${r?.sections} paras=${r?.totalParas} chars=${r?.totalChars} tables=${r?.totalTables} images=${r?.totalImages} footnotes=${r?.footnotes}`);
    }
  }

  // Global totals
  const totalDiffs = results.reduce((s, r) => s + r.diffs.length, 0);
  const catCounts = Object.entries(categories).map(([c, d]) => `${c}=${d.length}`).join(', ');
  console.log('\n--- GLOBAL TOTALS ---');
  console.log(`Total diffs across all files: ${totalDiffs}`);
  console.log(`By category: ${catCounts}`);

  // Unique diff paths (to see patterns)
  const pathPatterns = new Map<string, number>();
  for (const result of results) {
    for (const d of result.diffs) {
      // Normalize path (remove array indices for pattern counting)
      const pattern = d.path
        .replace(/\[[\w-]+\]\s+/, '')          // file prefix
        .replace(/\[\d+\]/g, '[N]')             // array indices
        .replace(/section\[N\]/, 'section[N]');
      pathPatterns.set(pattern, (pathPatterns.get(pattern) || 0) + 1);
    }
  }

  console.log('\n--- MOST COMMON DIFF PATTERNS ---');
  const sorted = [...pathPatterns.entries()].sort((a, b) => b[1] - a[1]).slice(0, 40);
  for (const [pattern, count] of sorted) {
    console.log(`  ${count}x  ${pattern}`);
  }

  // Save JSON report
  const reportDir = path.resolve(__dirname, '..', '.omc', 'scientist', 'reports');
  fs.mkdirSync(reportDir, { recursive: true });
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const reportPath = path.join(reportDir, `${timestamp}_hwp_conversion_diff.json`);
  fs.writeFileSync(reportPath, JSON.stringify({
    timestamp: new Date().toISOString(),
    totalDiffs,
    categorySummary: Object.fromEntries(Object.entries(categories).map(([c, d]) => [c, d.length])),
    topPatterns: sorted.map(([p, c]) => ({ pattern: p, count: c })),
    results: results.map(r => ({
      file: r.file,
      error: r.error,
      diffCount: r.diffs.length,
      stats: r.stats,
      diffs: r.diffs,
    })),
  }, null, 2));
  console.log(`\nFull JSON report saved to: ${reportPath}`);
}

// ============================================================
// Main
// ============================================================

async function main() {
  const samplesDir = path.resolve(__dirname, '..', 'samples');

  // Primary target + 3 additional diverse files
  const testFiles = [
    path.join(samplesDir, 'RAPIDS_결과보고서.hwpx'),
    path.join(samplesDir, 'test1.hwpx'),
    path.join(samplesDir, 'test2.hwpx'),
    path.join(samplesDir, 'PRISM_결과보고서.hwpx'),
  ].filter(f => {
    if (!fs.existsSync(f)) {
      console.log(`SKIP (not found): ${f}`);
      return false;
    }
    return true;
  });

  console.log(`Testing ${testFiles.length} files...`);

  const results: Array<{ file: string; diffs: Diff[]; stats: any; error?: string }> = [];
  for (const f of testFiles) {
    const result = await testFile(f);
    results.push(result);
  }

  printReport(results);
}

main().catch(e => {
  console.error('Fatal error:', e);
  process.exit(1);
});
