import * as fs from 'fs';
import { HwpParser } from './mcp-server/src/HwpParser';

const data = fs.readFileSync('samples/test1.hwp');
const content = HwpParser.parse(new Uint8Array(data));
const tables = content.sections[0].elements.filter((e: any) => e.type === 'table');

console.log('=== test1.hwp table cell backgrounds ===');
tables.forEach((t: any, ti: number) => {
  t.data.rows.forEach((r: any, ri: number) => {
    r.cells.forEach((c: any, ci: number) => {
      if (c.backgroundColor) {
        console.log(`  T${ti}[${ri},${ci}] bg=${c.backgroundColor} borders=${!!c.borderTop} hasMargin=${c.hasMargin}`);
      }
    });
  });
});

// Check first table structure in detail
const t0 = tables[0].data;
console.log('\nTable 0 details:');
console.log('  rows:', t0.rows.length, 'width:', t0.width);
console.log('  inMargin:', JSON.stringify(t0.inMargin));
t0.rows.forEach((r: any, ri: number) => {
  r.cells.forEach((c: any, ci: number) => {
    console.log(`  [${ri},${ci}] bg=${c.backgroundColor||'none'} w=${Math.round(c.width||0)} h=${Math.round(c.height||0)} borderFillId=${c.borderFillId} borderTop=${JSON.stringify(c.borderTop||'none').substring(0,40)}`);
  });
});
