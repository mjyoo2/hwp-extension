import { parse } from 'hwp.js';
import {
  HwpxContent,
  HwpxSection,
  HwpxParagraph,
  TextRun,
  SectionElement,
  PageSettings,
} from '../hwpx/types';

function generateId(): string {
  return Math.random().toString(36).substring(2, 11);
}

export class HwpParser {
  static parse(data: Uint8Array): HwpxContent {
    const hwpDocument = parse(data);
    
    const content: HwpxContent = {
      metadata: this.extractMetadata(hwpDocument),
      sections: [],
      images: new Map(),
      binItems: new Map(),
      binData: new Map(),
      footnotes: [],
      endnotes: [],
    };

    for (const hwpSection of hwpDocument.sections) {
      const section = this.convertSection(hwpSection, hwpDocument.info);
      content.sections.push(section);
    }

    return content;
  }

  private static extractMetadata(hwpDoc: any): HwpxContent['metadata'] {
    return {
      title: hwpDoc.info?.title || undefined,
      creator: hwpDoc.info?.author || undefined,
    };
  }

  private static convertSection(hwpSection: any, docInfo: any): HwpxSection {
    const section: HwpxSection = {
      elements: [],
      pageSettings: this.convertPageSettings(hwpSection),
    };

    for (const paragraph of hwpSection.content) {
      const converted = this.convertParagraph(paragraph, docInfo);
      section.elements.push({ type: 'paragraph', data: converted });
    }

    return section;
  }

  private static convertPageSettings(hwpSection: any): PageSettings {
    return {
      width: hwpSection.width / 100 || 595,
      height: hwpSection.height / 100 || 842,
      marginTop: hwpSection.paddingTop / 100 || 56.7,
      marginBottom: hwpSection.paddingBottom / 100 || 56.7,
      marginLeft: hwpSection.paddingLeft / 100 || 56.7,
      marginRight: hwpSection.paddingRight / 100 || 56.7,
      orientation: hwpSection.orientation === 1 ? 'landscape' : 'portrait',
    };
  }

  private static convertParagraph(hwpParagraph: any, docInfo: any): HwpxParagraph {
    const paragraph: HwpxParagraph = {
      id: generateId(),
      runs: [],
    };

    const shapeIndex = hwpParagraph.shapeIndex;
    if (docInfo?.paragraphShapes?.[shapeIndex]) {
      const paraShape = docInfo.paragraphShapes[shapeIndex];
      paragraph.paraStyle = {
        align: this.convertAlign(paraShape.align),
        lineSpacing: paraShape.lineSpacing || undefined,
      };
    }

    let currentText = '';
    let currentCharShapeIndex = -1;
    let shapePointerIndex = 0;

    for (let i = 0; i < hwpParagraph.content.length; i++) {
      const char = hwpParagraph.content[i];
      
      let charShapeIndex = currentCharShapeIndex;
      if (hwpParagraph.shapeBuffer && hwpParagraph.shapeBuffer.length > 0) {
        const nextShape = hwpParagraph.shapeBuffer[shapePointerIndex + 1];
        if (nextShape && i >= nextShape.pos) {
          shapePointerIndex++;
        }
        charShapeIndex = hwpParagraph.shapeBuffer[shapePointerIndex]?.shapeIndex ?? 0;
      }

      if (char.type === 0) {
        if (charShapeIndex !== currentCharShapeIndex && currentText) {
          paragraph.runs.push(this.createRun(currentText, currentCharShapeIndex, docInfo));
          currentText = '';
        }
        currentCharShapeIndex = charShapeIndex;
        currentText += char.value;
      }
    }

    if (currentText) {
      paragraph.runs.push(this.createRun(currentText, currentCharShapeIndex, docInfo));
    }

    if (paragraph.runs.length === 0) {
      paragraph.runs.push({ text: '' });
    }

    return paragraph;
  }

  private static createRun(text: string, charShapeIndex: number, docInfo: any): TextRun {
    const run: TextRun = { text };

    if (charShapeIndex >= 0 && docInfo?.charShapes?.[charShapeIndex]) {
      const charShape = docInfo.charShapes[charShapeIndex];
      run.charStyle = {
        fontSize: charShape.height ? charShape.height / 100 : undefined,
        bold: charShape.bold || false,
        italic: charShape.italic || false,
        underline: charShape.underline || false,
        strikethrough: charShape.strikeout || false,
        fontColor: charShape.color ? this.convertColor(charShape.color) : undefined,
      };
      
      if (docInfo.fontFaces && charShape.fontId !== undefined) {
        const fontFace = docInfo.fontFaces[charShape.fontId];
        if (fontFace) {
          run.charStyle.fontName = fontFace.name;
        }
      }
    }

    return run;
  }

  private static convertAlign(align: number): 'left' | 'center' | 'right' | 'justify' | undefined {
    switch (align) {
      case 0: return 'justify';
      case 1: return 'left';
      case 2: return 'right';
      case 3: return 'center';
      default: return undefined;
    }
  }

  private static convertColor(color: number): string {
    const r = color & 0xFF;
    const g = (color >> 8) & 0xFF;
    const b = (color >> 16) & 0xFF;
    return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
  }
}
