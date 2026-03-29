/**
 * Compare HWP and HWPX parser outputs for the same document.
 * Usage: npx esbuild compare-formats.ts --bundle --outfile=out/compare-formats.js --format=cjs --platform=node && node out/compare-formats.js samples/test1
 */
import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import { HwpxParser } from './mcp-server/src/HwpxParser';
import { HwpParser } from './mcp-server/src/HwpParser';

interface ElementSummary {
  type: string;
  text?: string;
  runs?: { text: string; bold?: boolean; italic?: boolean; fontSize?: number }[];
  rows?: number;
  cols?: number;
  cellTexts?: string[][];
  imageId?: string;
}

function summarizeElements(elements: any[]): ElementSummary[] {
  return elements.map(el => {
    const summary: ElementSummary = { type: el.type };
    if (el.type === 'paragraph' && el.data) {
      summary.runs = (el.data.runs || []).map((r: any) => ({
        text: r.text || '',
        bold: r.charStyle?.bold,
        italic: r.charStyle?.italic,
        fontSize: r.charStyle?.fontSize,
      }));
      summary.text = summary.runs!.map(r => r.text).join('');
    } else if (el.type === 'table' && el.data) {
      summary.rows = el.data.rows?.length || 0;
      summary.cols = el.data.rows?.[0]?.cells?.length || 0;
      summary.cellTexts = (el.data.rows || []).map((row: any) =>
        (row.cells || []).map((cell: any) =>
          (cell.paragraphs || []).map((p: any) =>
            (p.runs || []).map((r: any) => r.text || '').join('')
          ).join('\n')
        )
      );
    } else if (el.type === 'image' && el.data) {
      summary.imageId = el.data.binaryId || el.data.id || 'unknown';
    }
    return summary;
  });
}

async function main() {
  const basePath = process.argv[2]; // e.g., "samples/test1"
  if (!basePath) {
    console.error('Usage: node out/compare-formats.js samples/test1');
    process.exit(1);
  }

  const hwpPath = basePath + '.hwp';
  const hwpxPath = basePath + '.hwpx';

  console.log(`\n=== Comparing ${path.basename(hwpPath)} vs ${path.basename(hwpxPath)} ===\n`);

  // Parse HWP
  console.log('Parsing HWP...');
  const hwpStart = Date.now();
  const hwpData = fs.readFileSync(hwpPath);
  const hwpContent = HwpParser.parse(new Uint8Array(hwpData));
  const hwpTime = Date.now() - hwpStart;
  console.log(`  HWP parsed in ${hwpTime}ms`);

  // Parse HWPX
  console.log('Parsing HWPX...');
  const hwpxStart = Date.now();
  const hwpxData = fs.readFileSync(hwpxPath);
  const zip = await JSZip.loadAsync(hwpxData);
  const hwpxContent = await HwpxParser.parse(zip);
  const hwpxTime = Date.now() - hwpxStart;
  console.log(`  HWPX parsed in ${hwpxTime}ms`);
  console.log(`  Speed ratio: HWPX is ${(hwpxTime / hwpTime).toFixed(1)}x slower than HWP\n`);

  // Compare sections
  console.log(`Sections: HWP=${hwpContent.sections.length} HWPX=${hwpxContent.sections.length}`);

  const maxSections = Math.max(hwpContent.sections.length, hwpxContent.sections.length);
  for (let s = 0; s < maxSections; s++) {
    const hwpSection = hwpContent.sections[s];
    const hwpxSection = hwpxContent.sections[s];

    if (!hwpSection) { console.log(`  Section ${s}: MISSING in HWP`); continue; }
    if (!hwpxSection) { console.log(`  Section ${s}: MISSING in HWPX`); continue; }

    const hwpElems = summarizeElements(hwpSection.elements);
    const hwpxElems = summarizeElements(hwpxSection.elements);

    console.log(`\nSection ${s}: HWP=${hwpElems.length} elements, HWPX=${hwpxElems.length} elements`);

    // Count by type
    const hwpTypes: Record<string, number> = {};
    const hwpxTypes: Record<string, number> = {};
    hwpElems.forEach(e => hwpTypes[e.type] = (hwpTypes[e.type] || 0) + 1);
    hwpxElems.forEach(e => hwpxTypes[e.type] = (hwpxTypes[e.type] || 0) + 1);
    console.log(`  HWP types:  ${JSON.stringify(hwpTypes)}`);
    console.log(`  HWPX types: ${JSON.stringify(hwpxTypes)}`);

    // Compare paragraphs text
    const hwpParas = hwpElems.filter(e => e.type === 'paragraph');
    const hwpxParas = hwpxElems.filter(e => e.type === 'paragraph');
    console.log(`  Paragraphs: HWP=${hwpParas.length} HWPX=${hwpxParas.length}`);

    let textMismatches = 0;
    const maxParas = Math.min(hwpParas.length, hwpxParas.length);
    for (let p = 0; p < maxParas; p++) {
      const hwpText = hwpParas[p].text || '';
      const hwpxText = hwpxParas[p].text || '';
      if (hwpText !== hwpxText) {
        textMismatches++;
        if (textMismatches <= 5) {
          console.log(`  [TEXT DIFF] Para ${p}:`);
          console.log(`    HWP:  "${hwpText.substring(0, 100)}${hwpText.length > 100 ? '...' : ''}"`);
          console.log(`    HWPX: "${hwpxText.substring(0, 100)}${hwpxText.length > 100 ? '...' : ''}"`);
        }
      }
    }
    if (textMismatches > 5) console.log(`  ... and ${textMismatches - 5} more text mismatches`);
    console.log(`  Text mismatches: ${textMismatches}/${maxParas}`);

    // Compare tables
    const hwpTables = hwpElems.filter(e => e.type === 'table');
    const hwpxTables = hwpxElems.filter(e => e.type === 'table');
    console.log(`  Tables: HWP=${hwpTables.length} HWPX=${hwpxTables.length}`);

    let tableMismatches = 0;
    const maxTables = Math.min(hwpTables.length, hwpxTables.length);
    for (let t = 0; t < maxTables; t++) {
      const hwpT = hwpTables[t];
      const hwpxT = hwpxTables[t];
      if (hwpT.rows !== hwpxT.rows || hwpT.cols !== hwpxT.cols) {
        tableMismatches++;
        console.log(`  [TABLE STRUCT DIFF] Table ${t}: HWP=${hwpT.rows}x${hwpT.cols} HWPX=${hwpxT.rows}x${hwpxT.cols}`);
      } else if (hwpT.cellTexts && hwpxT.cellTexts) {
        // Compare cell texts
        let cellDiffs = 0;
        for (let r = 0; r < (hwpT.cellTexts?.length || 0); r++) {
          for (let c = 0; c < (hwpT.cellTexts[r]?.length || 0); c++) {
            const hwpCell = hwpT.cellTexts[r]?.[c] || '';
            const hwpxCell = hwpxT.cellTexts[r]?.[c] || '';
            if (hwpCell !== hwpxCell) {
              cellDiffs++;
              if (cellDiffs <= 3) {
                console.log(`  [CELL DIFF] Table ${t} [${r},${c}]:`);
                console.log(`    HWP:  "${hwpCell.substring(0, 80)}"`);
                console.log(`    HWPX: "${hwpxCell.substring(0, 80)}"`);
              }
            }
          }
        }
        if (cellDiffs > 0) {
          tableMismatches++;
          if (cellDiffs > 3) console.log(`    ... and ${cellDiffs - 3} more cell diffs in table ${t}`);
        }
      }
    }
    console.log(`  Table mismatches: ${tableMismatches}/${maxTables}`);

    // Compare style info for first few paragraphs
    let styleMismatches = 0;
    for (let p = 0; p < Math.min(maxParas, 20); p++) {
      const hwpRuns = hwpParas[p].runs || [];
      const hwpxRuns = hwpxParas[p].runs || [];
      if (hwpRuns.length !== hwpxRuns.length) {
        styleMismatches++;
        if (styleMismatches <= 3) {
          console.log(`  [RUN COUNT DIFF] Para ${p}: HWP=${hwpRuns.length} runs, HWPX=${hwpxRuns.length} runs`);
        }
      } else {
        for (let r = 0; r < hwpRuns.length; r++) {
          if (hwpRuns[r].bold !== hwpxRuns[r].bold || hwpRuns[r].italic !== hwpxRuns[r].italic) {
            styleMismatches++;
            if (styleMismatches <= 3) {
              console.log(`  [STYLE DIFF] Para ${p}, Run ${r}: HWP bold=${hwpRuns[r].bold} italic=${hwpRuns[r].italic}, HWPX bold=${hwpxRuns[r].bold} italic=${hwpxRuns[r].italic}`);
            }
            break;
          }
        }
      }
    }
    console.log(`  Style mismatches (first 20 paras): ${styleMismatches}`);

    // Images
    const hwpImages = hwpElems.filter(e => e.type === 'image');
    const hwpxImages = hwpxElems.filter(e => e.type === 'image');
    console.log(`  Images: HWP=${hwpImages.length} HWPX=${hwpxImages.length}`);
  }

  console.log('\n=== Done ===\n');
}

main().catch(console.error);
