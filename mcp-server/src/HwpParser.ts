/**
 * HWP Binary Parser - uses the standalone parser (cfb + pako)
 * Replaces the broken hwp.js v0.0.3 wrapper
 */
import { parseHwpContent } from './HwpParser.standalone';
import { HwpxContent } from './types';

export class HwpParser {
  static parse(data: Uint8Array): HwpxContent {
    return parseHwpContent(data);
  }
}
