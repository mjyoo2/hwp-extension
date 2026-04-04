import * as fs from 'fs';
import * as JSZip from 'jszip';
import { HwpxParser } from '../shared/src/HwpxParser';
import { writeHwpContent } from '../shared/src/HwpWriter';
import { parseHwpContent } from '../shared/src/HwpParser';

async function main() {
  const hwpxPath = 'samples/RAPIDS_결과보고서.hwpx';
  const data = fs.readFileSync(hwpxPath);
  console.log(`HWPX file size: ${data.length} bytes`);

  const zip = await JSZip.loadAsync(data);
  const content = await HwpxParser.parse(zip);
  console.log(`Parsed HWPX: ${content.sections.length} sections, ${content.sections.reduce((a: number, s: any) => a + s.elements.length, 0)} elements`);
  console.log(`Images: ${content.images.size}`);
  console.log(`Tables: ${content.sections.reduce((a: number, s: any) => a + s.elements.filter((e: any) => e.type === 'table').length, 0)}`);

  // Try to convert
  try {
    const hwpData = writeHwpContent(content);
    console.log(`HWP written: ${hwpData.length} bytes`);
    console.log(`First 4 bytes: ${Array.from(hwpData.slice(0, 4)).map(b => b.toString(16).padStart(2, '0')).join(' ')}`);

    // Try to re-parse
    const reread = parseHwpContent(hwpData);
    console.log(`HWP re-parsed OK: ${reread.sections.length} sections`);
  } catch (e: any) {
    console.log(`Convert error: ${e.message}`);
    console.log(`Stack:\n${e.stack?.split('\n').slice(0, 10).join('\n')}`);
  }
}

main();
