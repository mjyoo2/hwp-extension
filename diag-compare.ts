// Diagnostic: compare table cell borders/backgrounds between test1.hwp and test1.hwpx
import * as fs from 'fs';
import * as path from 'path';

// Use the HwpDocument.parseHwpContent (same path as the editor)
// We need to extract it - it's a local function, access via HwpDocument.parseContent
const { HwpDocument } = require('./src/hwp/HwpDocument');

// Use HwpxParser for HWPX
import JSZip from 'jszip';

async function main() {
  // Parse HWP
  const hwpData = fs.readFileSync(path.join(__dirname, 'samples/test1.hwp'));
  const hwpContent = HwpDocument.parseContent(hwpData);

  console.log('=== HWP Tables ===');
  for (const section of hwpContent.sections) {
    for (const el of section.elements) {
      if (el.type === 'table') {
        const table = el.data;
        console.log(`Table: ${table.rows.length} rows`);
        for (const row of table.rows) {
          for (const cell of row.cells) {
            const info: any = {
              pos: `[${cell.rowAddr},${cell.colAddr}]`,
              bg: cell.backgroundColor || 'none',
              grad: cell.backgroundGradation ? JSON.stringify(cell.backgroundGradation) : 'none',
              borderTop: cell.borderTop ? `${cell.borderTop.style} ${cell.borderTop.width?.toFixed(2)}pt ${cell.borderTop.color}` : 'none',
              borderLeft: cell.borderLeft ? `${cell.borderLeft.style} ${cell.borderLeft.width?.toFixed(2)}pt ${cell.borderLeft.color}` : 'none',
              borderFillId: cell.borderFillId,
            };
            console.log('  Cell:', JSON.stringify(info));
          }
        }
      }
    }
  }

  // Parse HWPX
  const { HwpxParser } = require('./src/hwpx/HwpxParser');
  const hwpxData = fs.readFileSync(path.join(__dirname, 'samples/test1.hwpx'));
  const hwpxContent = await HwpxParser.parse(hwpxData);

  console.log('\n=== HWPX Tables ===');
  for (const section of hwpxContent.sections) {
    for (const el of section.elements) {
      if (el.type === 'table') {
        const table = el.data;
        console.log(`Table: ${table.rows.length} rows`);
        for (const row of table.rows) {
          for (const cell of row.cells) {
            const info: any = {
              pos: `[${cell.rowAddr},${cell.colAddr}]`,
              bg: cell.backgroundColor || 'none',
              grad: cell.backgroundGradation ? JSON.stringify(cell.backgroundGradation) : 'none',
              borderTop: cell.borderTop ? `${cell.borderTop.style} ${cell.borderTop.width?.toFixed(2)}pt ${cell.borderTop.color}` : 'none',
              borderLeft: cell.borderLeft ? `${cell.borderLeft.style} ${cell.borderLeft.width?.toFixed(2)}pt ${cell.borderLeft.color}` : 'none',
              borderFillId: cell.borderFillId,
            };
            console.log('  Cell:', JSON.stringify(info));
          }
        }
      }
    }
  }
}

main().catch(console.error);
