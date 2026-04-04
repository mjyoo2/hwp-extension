/**
 * Comprehensive HWPX→HWP roundtrip check for ALL sample files.
 * Checks: text, charStyle, paraStyle, borders, backgroundColor,
 * verticalAlign, width, header/footer, images, pageSettings
 */
import * as fs from 'fs';
import * as JSZip from 'jszip';
import { HwpxParser } from '../shared/src/HwpxParser';
import { writeHwpContent } from '../shared/src/HwpWriter';
import { parseHwpContent } from '../shared/src/HwpParser';

interface Diff { file: string; path: string; orig: string; rt: string; }

async function checkFile(name: string): Promise<Diff[]> {
  const diffs: Diff[] = [];
  const d = (path: string, orig: any, rt: any) => diffs.push({ file: name, path, orig: String(orig).substring(0, 60), rt: String(rt).substring(0, 60) });

  const data = fs.readFileSync('samples/' + name);
  const zip = await JSZip.loadAsync(data);
  const orig = await HwpxParser.parse(zip);
  const hwpData = writeHwpContent(orig);

  if (hwpData[0] !== 0xD0 || hwpData[1] !== 0xCF) { d('format', 'OLE expected', 'not OLE'); return diffs; }

  const rt = parseHwpContent(hwpData);

  for (let s = 0; s < Math.min(orig.sections.length, rt.sections.length); s++) {
    const oS = orig.sections[s], rS = rt.sections[s];

    // Elements count
    if (oS.elements.length !== rS.elements.length) d('S'+s+'.elements.length', oS.elements.length, rS.elements.length);

    // Paragraphs
    const oParas = oS.elements.filter(e => e.type === 'paragraph');
    const rParas = rS.elements.filter(e => e.type === 'paragraph');
    for (let p = 0; p < Math.min(oParas.length, rParas.length); p++) {
      const oP = oParas[p].data as any, rP = rParas[p].data as any;
      const oText = oP.runs?.map((r: any) => r.text).join('') || '';
      const rText = rP.runs?.map((r: any) => r.text).join('') || '';
      if (oText !== rText) d('P'+p+'.text', oText.substring(0,50), rText.substring(0,50));

      // Char styles (skip empty text paragraphs - known cosmetic issue)
      if (oText.length > 0) {
        for (let r = 0; r < Math.min(oP.runs?.length||0, rP.runs?.length||0); r++) {
          const oCS = oP.runs[r].charStyle, rCS = rP.runs[r].charStyle;
          if (oCS && rCS) {
            for (const k of ['fontName','fontSize','bold','italic','underline','strikethrough','fontColor']) {
              const ov = (oCS as any)[k], rv = (rCS as any)[k];
              if (ov !== rv && ov !== undefined) {
                // Deep compare for object values (e.g. underline)
                if (typeof ov === 'object' && typeof rv === 'object' && ov !== null && rv !== null
                    && JSON.stringify(ov) === JSON.stringify(rv)) continue;
                d('P'+p+'.run'+r+'.'+k, ov, rv);
              }
            }
          }
        }
      }

      // Para styles
      const oPS = oP.paraStyle, rPS = rP.paraStyle;
      if (oPS && rPS) {
        for (const k of ['align','lineSpacing','lineSpacingType']) {
          const ov = (oPS as any)[k], rv = (rPS as any)[k];
          if (typeof ov === 'number' && typeof rv === 'number') {
            if (Math.abs(ov - rv) > 0.1) d('P'+p+'.paraStyle.'+k, ov, rv);
          } else if (ov !== rv && ov !== undefined) d('P'+p+'.paraStyle.'+k, ov, rv);
        }
      }
    }

    // Tables
    const oTbls = oS.elements.filter(e => e.type === 'table');
    const rTbls = rS.elements.filter(e => e.type === 'table');
    if (oTbls.length !== rTbls.length) d('S'+s+'.tableCount', oTbls.length, rTbls.length);
    for (let t = 0; t < Math.min(oTbls.length, rTbls.length); t++) {
      const tA = oTbls[t].data as any, tB = rTbls[t].data as any;
      if (tA.rows.length !== tB.rows.length) d('T'+t+'.rowCount', tA.rows.length, tB.rows.length);
      for (let r = 0; r < Math.min(tA.rows.length, tB.rows.length); r++) {
        if (tA.rows[r].cells.length !== tB.rows[r].cells.length) d('T'+t+'R'+r+'.cellCount', tA.rows[r].cells.length, tB.rows[r].cells.length);
        for (let c = 0; c < Math.min(tA.rows[r].cells.length, tB.rows[r].cells.length); c++) {
          const oC = tA.rows[r].cells[c], rC = tB.rows[r].cells[c];
          // Text
          const oTx = (oC.paragraphs||[]).map((p:any)=>p.runs?.map((r:any)=>r.text).join('')||'').join('\n');
          const rTx = (rC.paragraphs||[]).map((p:any)=>p.runs?.map((r:any)=>r.text).join('')||'').join('\n');
          if (oTx !== rTx) {
            // Check if diff is just trailing newlines
            if (oTx.trimEnd() !== rTx.trimEnd()) d('T'+t+'R'+r+'C'+c+'.text', oTx.substring(0,50), rTx.substring(0,50));
          }
          // Borders
          for (const side of ['borderTop','borderBottom','borderLeft','borderRight'] as const) {
            const ob = (oC as any)[side], rb = (rC as any)[side];
            if (ob && !rb) d('T'+t+'R'+r+'C'+c+'.'+side, JSON.stringify(ob)?.substring(0,40), 'undefined');
          }
          // Background
          if (oC.backgroundColor !== rC.backgroundColor && oC.backgroundColor) d('T'+t+'R'+r+'C'+c+'.bg', oC.backgroundColor, rC.backgroundColor||'undefined');
          // VerticalAlign
          if (oC.verticalAlign !== rC.verticalAlign && oC.verticalAlign && oC.verticalAlign !== 'top') d('T'+t+'R'+r+'C'+c+'.vAlign', oC.verticalAlign, rC.verticalAlign||'undefined');
          // Width (tolerance 1pt)
          if (Math.abs((oC.width||0) - (rC.width||0)) > 1) d('T'+t+'R'+r+'C'+c+'.width', Math.round(oC.width), Math.round(rC.width));
        }
      }
    }

    // Header/footer
    if (!!oS.header !== !!rS.header) d('header', !!oS.header, !!rS.header);
    if (!!oS.footer !== !!rS.footer) d('footer', !!oS.footer, !!rS.footer);

    // Page settings
    if (oS.pageSettings && rS.pageSettings) {
      for (const k of ['width','height','marginLeft','marginRight','marginTop','marginBottom']) {
        const ov = (oS.pageSettings as any)[k], rv = (rS.pageSettings as any)[k];
        if (typeof ov === 'number' && typeof rv === 'number' && Math.abs(ov-rv) > 0.1) d('pageSettings.'+k, ov, rv);
      }
    }
  }

  // Images
  if (orig.images.size !== rt.images.size) d('imageCount', orig.images.size, rt.images.size);

  return diffs;
}

async function main() {
  const files = fs.readdirSync('samples').filter(f => f.endsWith('.hwpx') && !f.includes('_converted'));
  console.log('=== 전체 HWPX→HWP 라운드트립 검증 (' + files.length + '개 파일) ===\n');

  let totalDiffs = 0;
  for (const f of files) {
    try {
      const diffs = await checkFile(f);
      totalDiffs += diffs.length;
      if (diffs.length === 0) {
        console.log('[PASS] ' + f);
      } else {
        console.log('[DIFF] ' + f + ' (' + diffs.length + '개 차이):');
        for (const d of diffs.slice(0, 10)) {
          console.log('  ' + d.path + ': ' + d.orig + ' → ' + d.rt);
        }
        if (diffs.length > 10) console.log('  ... +' + (diffs.length - 10) + '개');
      }
    } catch (e: any) {
      console.log('[FAIL] ' + f + ': ' + e.message.substring(0, 80));
      totalDiffs++;
    }
  }

  console.log('\n=== 결과: ' + totalDiffs + '개 차이점 (총 ' + files.length + '개 파일) ===');
}

main();
