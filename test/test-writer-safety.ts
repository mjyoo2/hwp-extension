/**
 * Round-trip test for HWP Writer safety
 * Reads a sample HWP file, writes it back, re-parses, and compares.
 *
 * Usage: npx tsx test/test-writer-safety.ts
 */

import * as fs from 'fs';
import * as path from 'path';
import { parseHwpContent } from '../shared/src/HwpParser';
import { writeHwpContent } from '../shared/src/HwpWriter';
import { HwpxContent, SectionElement } from '../shared/src/types';

// ============================================================
// Helpers
// ============================================================

function countElementsByType(content: HwpxContent): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const section of content.sections) {
    for (const element of section.elements) {
      counts[element.type] = (counts[element.type] || 0) + 1;
    }
  }
  return counts;
}

function extractAllText(content: HwpxContent): string {
  const parts: string[] = [];
  for (const section of content.sections) {
    for (const element of section.elements) {
      if (element.type === 'paragraph') {
        const para = element.data as any;
        if (para.runs) {
          for (const run of para.runs) {
            if (run.text) parts.push(run.text);
          }
        }
      } else if (element.type === 'table') {
        const table = element.data as any;
        if (table.rows) {
          for (const row of table.rows) {
            for (const cell of row.cells) {
              if (cell.paragraphs) {
                for (const para of cell.paragraphs) {
                  if (para.runs) {
                    for (const run of para.runs) {
                      if (run.text) parts.push(run.text);
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }
  return parts.join('');
}

function countTables(content: HwpxContent): number {
  let count = 0;
  for (const section of content.sections) {
    for (const element of section.elements) {
      if (element.type === 'table') count++;
    }
  }
  return count;
}

function countImages(content: HwpxContent): number {
  return content.images.size;
}

// ============================================================
// Main test
// ============================================================

interface TestResult {
  name: string;
  passed: boolean;
  detail: string;
}

function runRoundTripTest(filePath: string): TestResult[] {
  const results: TestResult[] = [];
  const fileName = path.basename(filePath);

  console.log(`\n=== Round-trip test: ${fileName} ===\n`);

  // Step 1: Read original
  let originalData: Uint8Array;
  try {
    originalData = new Uint8Array(fs.readFileSync(filePath));
  } catch (e: any) {
    results.push({ name: 'Read original', passed: false, detail: e.message });
    return results;
  }
  results.push({ name: 'Read original', passed: true, detail: `${originalData.length} bytes` });

  // Step 2: Parse original
  let originalContent: HwpxContent;
  try {
    originalContent = parseHwpContent(originalData);
  } catch (e: any) {
    results.push({ name: 'Parse original', passed: false, detail: e.message });
    return results;
  }
  results.push({ name: 'Parse original', passed: true, detail: `${originalContent.sections.length} sections` });

  // Step 3: Write back
  let writtenData: Uint8Array;
  try {
    writtenData = writeHwpContent(originalContent);
  } catch (e: any) {
    results.push({ name: 'Write back', passed: false, detail: e.message });
    return results;
  }
  results.push({ name: 'Write back', passed: true, detail: `${writtenData.length} bytes` });

  // Step 4: Re-parse written data
  let rereadContent: HwpxContent;
  try {
    rereadContent = parseHwpContent(writtenData);
  } catch (e: any) {
    results.push({ name: 'Re-parse written data', passed: false, detail: e.message });
    return results;
  }
  results.push({ name: 'Re-parse written data', passed: true, detail: 'OK' });

  // Step 5: Compare section count
  const origSections = originalContent.sections.length;
  const rereadSections = rereadContent.sections.length;
  results.push({
    name: 'Section count',
    passed: origSections === rereadSections,
    detail: `original=${origSections}, roundtrip=${rereadSections}`,
  });

  // Step 6: Compare element types
  const origTypes = countElementsByType(originalContent);
  const rereadTypes = countElementsByType(rereadContent);
  const typeKeys = new Set([...Object.keys(origTypes), ...Object.keys(rereadTypes)]);
  let typesMatch = true;
  const typeDetails: string[] = [];
  for (const key of typeKeys) {
    const orig = origTypes[key] || 0;
    const reread = rereadTypes[key] || 0;
    typeDetails.push(`${key}: ${orig} -> ${reread}`);
    if (orig !== reread) typesMatch = false;
  }
  results.push({
    name: 'Element types',
    passed: typesMatch,
    detail: typeDetails.join(', '),
  });

  // Step 7: Compare text content
  const origText = extractAllText(originalContent);
  const rereadText = extractAllText(rereadContent);
  const textMatch = origText === rereadText;
  results.push({
    name: 'Text content',
    passed: textMatch,
    detail: textMatch
      ? `${origText.length} chars match`
      : `original=${origText.length} chars, roundtrip=${rereadText.length} chars (diff at position ${findFirstDiff(origText, rereadText)})`,
  });

  // Step 8: Compare table count
  const origTables = countTables(originalContent);
  const rereadTables = countTables(rereadContent);
  results.push({
    name: 'Table count',
    passed: origTables === rereadTables,
    detail: `original=${origTables}, roundtrip=${rereadTables}`,
  });

  // Step 9: Compare image count
  const origImages = countImages(originalContent);
  const rereadImages = countImages(rereadContent);
  results.push({
    name: 'Image count',
    passed: origImages === rereadImages,
    detail: `original=${origImages}, roundtrip=${rereadImages}`,
  });

  return results;
}

function findFirstDiff(a: string, b: string): number {
  const len = Math.min(a.length, b.length);
  for (let i = 0; i < len; i++) {
    if (a[i] !== b[i]) return i;
  }
  return len;
}

// ============================================================
// Entry point
// ============================================================

const sampleDir = path.resolve(__dirname, '..', 'samples');
const testFiles = ['test2.hwp', 'test1.hwp'].map(f => path.join(sampleDir, f)).filter(f => fs.existsSync(f));

if (testFiles.length === 0) {
  console.error('No sample HWP files found in samples/');
  process.exit(1);
}

let totalPassed = 0;
let totalFailed = 0;

for (const filePath of testFiles) {
  const results = runRoundTripTest(filePath);

  for (const r of results) {
    const status = r.passed ? 'PASS' : 'FAIL';
    const icon = r.passed ? '[+]' : '[X]';
    console.log(`  ${icon} ${status}: ${r.name} - ${r.detail}`);
    if (r.passed) totalPassed++;
    else totalFailed++;
  }
}

console.log(`\n=== Summary: ${totalPassed} passed, ${totalFailed} failed ===\n`);
process.exit(totalFailed > 0 ? 1 : 0);
