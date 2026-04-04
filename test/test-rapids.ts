import * as fs from 'fs';
import { parseHwpContent } from '../shared/src/HwpParser';

const hwpPath = 'samples/RAPIDS_결과보고서.hwp';
const data = new Uint8Array(fs.readFileSync(hwpPath));
console.log(`File size: ${data.length} bytes`);
console.log(`First 32 bytes: ${Array.from(data.slice(0, 32)).map(b => b.toString(16).padStart(2, '0')).join(' ')}`);

try {
  const content = parseHwpContent(data);
  console.log(`Parsed OK: ${content.sections.length} sections`);
} catch (e: any) {
  console.log(`Parse error: ${e.message}`);
  console.log(`Stack: ${e.stack?.split('\n').slice(0, 8).join('\n')}`);
}
