import * as fs from 'fs';
import { HwpParser } from './mcp-server/src/HwpParser';
import JSZip from 'jszip';
import { HwpxParser } from './mcp-server/src/HwpxParser';

function renderCellHtml(cell: any, table: any): string {
  let style = '';
  if (cell.backgroundColor) style += 'background-color:' + cell.backgroundColor + ';';
  if (cell.backgroundGradation?.colors?.length >= 2) {
    style = 'background:linear-gradient(...);'; // simplified
  }
  if (cell.width) style += 'width:' + cell.width + 'pt;';
  return style;
}

async function check(name: string) {
  console.log(`\n=== ${name} ===`);
  
  // HWP
  const hwpData = fs.readFileSync(`samples/${name}.hwp`);
  const hwpContent = HwpParser.parse(new Uint8Array(hwpData));
  const hwpTables = hwpContent.sections[0].elements.filter((e: any) => e.type === 'table');
  
  // HWPX
  const hwpxData = fs.readFileSync(`samples/${name}.hwpx`);
  const zip = await JSZip.loadAsync(hwpxData);
  const hwpxContent = await HwpxParser.parse(zip);
  const hwpxTables = hwpxContent.sections[0].elements.filter((e: any) => e.type === 'table');
  
  // Compare first 3 tables' HTML output
  for (let t = 0; t < Math.min(3, hwpTables.length); t++) {
    const ht = hwpTables[t].data;
    const xt = hwpxTables[t]?.data;
    console.log(`\nTable ${t}:`);
    for (let r = 0; r < Math.min(2, ht.rows.length); r++) {
      for (let c = 0; c < ht.rows[r].cells.length; c++) {
        const hc = ht.rows[r].cells[c];
        const xc = xt?.rows[r]?.cells[c];
        const hStyle = renderCellHtml(hc, ht);
        const xStyle = xc ? renderCellHtml(xc, xt) : 'N/A';
        if (hStyle !== xStyle) {
          console.log(`  [${r},${c}] HWP: "${hStyle}"`);
          console.log(`           HWPX: "${xStyle}"`);
        } else {
          console.log(`  [${r},${c}] SAME: "${hStyle}"`);
        }
      }
    }
  }
  
  // Check if any HWP cells have bg but HWPX doesn't (or vice versa)
  let hwpWithBg = 0, hwpxWithBg = 0, hwpNoBg = 0, hwpxNoBg = 0;
  for (let t = 0; t < hwpTables.length; t++) {
    const ht = hwpTables[t].data;
    const xt = hwpxTables[t]?.data;
    if (!xt) continue;
    for (let r = 0; r < ht.rows.length; r++) {
      for (let c = 0; c < ht.rows[r].cells.length; c++) {
        if (ht.rows[r].cells[c].backgroundColor) hwpWithBg++;
        else hwpNoBg++;
        if (xt.rows[r]?.cells[c]?.backgroundColor) hwpxWithBg++;
        else hwpxNoBg++;
      }
    }
  }
  console.log(`\nBG stats: HWP with=${hwpWithBg} without=${hwpNoBg} | HWPX with=${hwpxWithBg} without=${hwpxNoBg}`);
}

(async () => { await check('test1'); })();
