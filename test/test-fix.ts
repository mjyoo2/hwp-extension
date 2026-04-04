/**
 * Test HWPX → HWP conversion for all sample HWPX files
 * Verifies: HWPX parse → writeHwpContent → valid OLE output → re-parse
 */
import * as fs from 'fs';
import * as path from 'path';
import * as JSZip from 'jszip';
import { HwpxParser } from '../shared/src/HwpxParser';
import { writeHwpContent } from '../shared/src/HwpWriter';
import { parseHwpContent } from '../shared/src/HwpParser';

async function testConvert(hwpxPath: string): Promise<boolean> {
  const name = path.basename(hwpxPath);

  // 1. Parse HWPX
  const data = fs.readFileSync(hwpxPath);
  const zip = await JSZip.loadAsync(data);
  let content;
  try {
    content = await HwpxParser.parse(zip);
  } catch (e: any) {
    console.log(`  [FAIL] ${name}: HWPX parse error - ${e.message}`);
    return false;
  }

  // 2. Write as HWP
  let hwpData: Uint8Array;
  try {
    hwpData = writeHwpContent(content);
  } catch (e: any) {
    console.log(`  [FAIL] ${name}: HWP write error - ${e.message}`);
    return false;
  }

  // 3. Check magic bytes
  const isOle = hwpData.length >= 4 && hwpData[0] === 0xD0 && hwpData[1] === 0xCF;
  if (!isOle) {
    console.log(`  [FAIL] ${name}: Output is not OLE format (first bytes: ${Array.from(hwpData.slice(0, 4)).map(b => b.toString(16)).join(' ')})`);
    return false;
  }

  // 4. Re-parse
  try {
    const reread = parseHwpContent(hwpData);
    console.log(`  [OK] ${name}: ${content.sections.length} sections, ${content.sections.reduce((a: number, s: any) => a + s.elements.length, 0)} elements → ${hwpData.length} bytes OLE`);
    return true;
  } catch (e: any) {
    console.log(`  [FAIL] ${name}: HWP re-parse error - ${e.message}`);
    return false;
  }
}

async function main() {
  const sampleDir = path.resolve(__dirname, '..', 'samples');
  const hwpxFiles = fs.readdirSync(sampleDir).filter(f => f.endsWith('.hwpx'));

  console.log(`Testing ${hwpxFiles.length} HWPX files:\n`);

  let passed = 0, failed = 0;
  for (const f of hwpxFiles) {
    const ok = await testConvert(path.join(sampleDir, f));
    if (ok) passed++; else failed++;
  }

  // Also check existing .hwp files on disk
  console.log(`\nChecking .hwp files on disk:`);
  const hwpFiles = fs.readdirSync(sampleDir).filter(f => f.endsWith('.hwp'));
  for (const f of hwpFiles) {
    const d = fs.readFileSync(path.join(sampleDir, f));
    const isZip = d.length >= 2 && d[0] === 0x50 && d[1] === 0x4B;
    const isOle = d.length >= 4 && d[0] === 0xD0 && d[1] === 0xCF;
    console.log(`  ${f}: ${isOle ? 'OLE (valid HWP)' : isZip ? 'ZIP (actually HWPX!)' : 'unknown'}`);
  }

  console.log(`\n=== ${passed} passed, ${failed} failed ===`);
  process.exit(failed > 0 ? 1 : 0);
}

main();
