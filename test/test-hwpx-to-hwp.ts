import * as fs from 'fs';
import * as path from 'path';
import * as JSZip from 'jszip';
import { HwpxParser } from '../shared/src/HwpxParser';
import { writeHwpContent } from '../shared/src/HwpWriter';
import { parseHwpContent } from '../shared/src/HwpParser';

async function testConvert(hwpxPath: string) {
  console.log(`\n=== Testing: ${path.basename(hwpxPath)} ===`);
  
  // 1. Read HWPX
  const data = fs.readFileSync(hwpxPath);
  const zip = await JSZip.loadAsync(data);
  
  // 2. Parse HWPX content
  let content;
  try {
    content = await HwpxParser.parse(zip);
    console.log(`  [OK] HWPX parsed: ${content.sections.length} sections, ${content.sections.reduce((a: number, s: any) => a + s.elements.length, 0)} elements`);
  } catch (e: any) {
    console.log(`  [FAIL] HWPX parse error: ${e.message}`);
    return;
  }
  
  // 3. Write as HWP
  let hwpData: Uint8Array;
  try {
    hwpData = writeHwpContent(content);
    console.log(`  [OK] HWP written: ${hwpData.length} bytes`);
  } catch (e: any) {
    console.log(`  [FAIL] HWP write error: ${e.message}`);
    console.log(`  Stack: ${e.stack?.split('\n').slice(0,5).join('\n')}`);
    return;
  }
  
  // 4. Re-parse HWP
  try {
    const reread = parseHwpContent(hwpData);
    console.log(`  [OK] HWP re-parsed: ${reread.sections.length} sections`);
  } catch (e: any) {
    console.log(`  [FAIL] HWP re-parse error: ${e.message}`);
    console.log(`  Stack: ${e.stack?.split('\n').slice(0,5).join('\n')}`);
  }
}

async function main() {
  const sampleDir = path.resolve(__dirname, '..', 'samples');
  const hwpxFiles = fs.readdirSync(sampleDir).filter(f => f.endsWith('.hwpx')).slice(0, 5);
  
  for (const f of hwpxFiles) {
    await testConvert(path.join(sampleDir, f));
  }
}

main().catch(e => console.error(e));
