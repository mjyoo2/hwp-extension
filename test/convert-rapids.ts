import * as fs from 'fs';
import * as JSZip from 'jszip';
import { HwpxParser } from '../shared/src/HwpxParser';
import { writeHwpContent } from '../shared/src/HwpWriter';
import { parseHwpContent } from '../shared/src/HwpParser';

async function main() {
  const hwpxData = fs.readFileSync('samples/RAPIDS_결과보고서.hwpx');
  const zip = await JSZip.loadAsync(hwpxData);
  const content = await HwpxParser.parse(zip);
  const hwpData = writeHwpContent(content);

  if (hwpData[0] !== 0xD0 || hwpData[1] !== 0xCF) {
    console.log('ERROR: not OLE');
    return;
  }
  parseHwpContent(hwpData);

  fs.writeFileSync('samples/RAPIDS_결과보고서.hwp', Buffer.from(hwpData));
  console.log('OK: ' + hwpData.length + ' bytes → samples/RAPIDS_결과보고서.hwp');
}
main().catch(e => console.error('ERROR:', e.message));
