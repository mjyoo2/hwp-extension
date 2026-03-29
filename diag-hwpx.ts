import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';

async function main() {
  const { HwpxParser } = require('./src/hwpx/HwpxParser');
  const hwpxBuffer = fs.readFileSync(path.join(__dirname, 'samples/test1.hwpx'));
  const zip = await JSZip.loadAsync(hwpxBuffer);
  const hwpxContent = await HwpxParser.parse(zip);

  console.log('=== HWPX Tables ===');
  for (let si = 0; si < hwpxContent.sections.length; si++) {
    const section = hwpxContent.sections[si];
    for (let ei = 0; ei < section.elements.length; ei++) {
      const el = section.elements[ei];
      if (el.type === 'table') {
        const table = el.data;
        console.log(`\nSection ${si}, Element ${ei}: Table with ${table.rows.length} rows`);
        for (const row of table.rows) {
          for (const cell of row.cells) {
            const cellInfo: any = {
              pos: `[${cell.rowAddr},${cell.colAddr}]`,
              bg: cell.backgroundColor || 'none',
              paraCount: cell.paragraphs.length,
              elemCount: cell.elements?.length || 0,
            };
            if (cell.elements) {
              cellInfo.elemTypes = cell.elements.map((e: any) => {
                if (e.type === 'image') return `image(${e.data.binaryId || 'no-id'}, data=${!!e.data.data})`;
                if (e.type === 'paragraph') return `para(${e.data.runs?.length || 0} runs)`;
                return e.type;
              });
            }
            console.log('  Cell:', JSON.stringify(cellInfo));
          }
        }
      } else if (el.type === 'image') {
        console.log(`\nSection ${si}, Element ${ei}: STANDALONE Image id=${el.data.id} binaryId=${el.data.binaryId}`);
      }
    }
  }
}

main().catch(console.error);
