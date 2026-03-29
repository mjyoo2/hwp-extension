import * as fs from 'fs';
import JSZip from 'jszip';
import { HwpxParser } from './mcp-server/src/HwpxParser';
import { HwpParser } from './mcp-server/src/HwpParser';

async function main() {
  const hwpData = fs.readFileSync('samples/test1.hwp');
  const hwpContent = HwpParser.parse(new Uint8Array(hwpData));
  
  const hwpxData = fs.readFileSync('samples/test1.hwpx');
  const zip = await JSZip.loadAsync(hwpxData);
  const hwpxContent = await HwpxParser.parse(zip);

  // Check cell backgrounds
  for (let s = 0; s < Math.min(hwpContent.sections.length, hwpxContent.sections.length); s++) {
    const hwpTables = hwpContent.sections[s].elements.filter((e: any) => e.type === 'table');
    const hwpxTables = hwpxContent.sections[s].elements.filter((e: any) => e.type === 'table');
    
    for (let t = 0; t < Math.min(hwpTables.length, hwpxTables.length); t++) {
      const hwpRows = hwpTables[t].data.rows || [];
      const hwpxRows = hwpxTables[t].data.rows || [];
      for (let r = 0; r < Math.min(hwpRows.length, hwpxRows.length); r++) {
        for (let c = 0; c < Math.min(hwpRows[r].cells.length, hwpxRows[r].cells.length); c++) {
          const hwpBg = hwpRows[r].cells[c].backgroundColor;
          const hwpxBg = hwpxRows[r].cells[c].backgroundColor;
          if (hwpBg !== hwpxBg) {
            console.log(`Table ${t} [${r},${c}] bg: HWP="${hwpBg}" HWPX="${hwpxBg}"`);
          }
        }
      }
    }
  }

  // Check images
  const hwpImgs = hwpContent.sections[0].elements.filter((e: any) => e.type === 'image');
  const hwpxImgs = hwpxContent.sections[0].elements.filter((e: any) => e.type === 'image');
  console.log(`\nImages: HWP=${hwpImgs.length} HWPX=${hwpxImgs.length}`);
  hwpImgs.forEach((img: any, i: number) => {
    console.log(`  HWP img[${i}]: w=${img.data.width} h=${img.data.height} id=${img.data.binaryId}`);
  });
  hwpxImgs.forEach((img: any, i: number) => {
    console.log(`  HWPX img[${i}]: w=${img.data.width} h=${img.data.height} id=${img.data.binaryId}`);
  });
  
  // Check HWP images map
  console.log(`\nHWP images map: ${hwpContent.images.size} entries`);
  hwpContent.images.forEach((v: any, k: string) => {
    console.log(`  ${k}: ${v.width}x${v.height} ${v.mimeType} data=${v.data ? v.data.substring(0, 30) + '...' : 'none'}`);
  });
  console.log(`HWPX images map: ${hwpxContent.images.size} entries`);
  hwpxContent.images.forEach((v: any, k: string) => {
    console.log(`  ${k}: ${v.width}x${v.height} ${v.mimeType} data=${v.data ? v.data.substring(0, 30) + '...' : 'none'}`);
  });
}
main().catch(console.error);
