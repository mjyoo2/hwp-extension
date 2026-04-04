/**
 * HWP Binary Parser - uses the standalone parser (cfb + pako)
 * Replaces the broken hwp.js v0.0.3 wrapper
 */
import { parseHwpContent } from '../../shared/src/HwpParser';
import { writeHwpContent } from '../../shared/src/HwpWriter';
import { HwpxContent } from '../../shared/src/types';

export class HwpParser {
  static parse(data: Uint8Array): HwpxContent {
    return parseHwpContent(data);
  }

  static write(content: HwpxContent): Uint8Array {
    return writeHwpContent(content);
  }
}
