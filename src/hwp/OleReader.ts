/**
 * OLE Compound File Binary (CFB) Reader
 * Parses Microsoft Compound Document File Format used by HWP files
 * 
 * Reference: [MS-CFB] Compound File Binary File Format
 */

const HEADER_SIGNATURE = new Uint8Array([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
const ENDOFCHAIN = 0xFFFFFFFE;
const FREESECT = 0xFFFFFFFF;
const DIFSECT = 0xFFFFFFFC;
const FATSECT = 0xFFFFFFFD;

export interface DirectoryEntry {
  name: string;
  type: number; // 0=unknown, 1=storage, 2=stream, 5=root
  colorFlag: number;
  leftSiblingId: number;
  rightSiblingId: number;
  childId: number;
  startSectorLocation: number;
  size: number;
  clsid: Uint8Array;
}

export class OleReader {
  private data: DataView;
  private sectorSize: number = 512;
  private miniSectorSize: number = 64;
  private miniStreamCutoffSize: number = 4096;
  private fat: number[] = [];
  private miniFat: number[] = [];
  private directoryEntries: DirectoryEntry[] = [];
  private miniStreamData: Uint8Array = new Uint8Array(0);

  constructor(buffer: Uint8Array) {
    this.data = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);
    this.parse();
  }

  private parse(): void {
    this.validateHeader();
    this.readHeader();
    this.readFAT();
    this.readDirectoryEntries();
    this.readMiniFAT();
    this.readMiniStream();
  }

  private validateHeader(): void {
    for (let i = 0; i < 8; i++) {
      if (this.data.getUint8(i) !== HEADER_SIGNATURE[i]) {
        throw new Error('Invalid OLE file: signature mismatch');
      }
    }
  }

  private readHeader(): void {
    const majorVersion = this.data.getUint16(0x1A, true);
    
    if (majorVersion === 3) {
      this.sectorSize = 512;
    } else if (majorVersion === 4) {
      this.sectorSize = 4096;
    } else {
      throw new Error(`Unsupported OLE version: ${majorVersion}`);
    }

    const sectorShift = this.data.getUint16(0x1E, true);
    this.sectorSize = 1 << sectorShift;
    
    const miniSectorShift = this.data.getUint16(0x20, true);
    this.miniSectorSize = 1 << miniSectorShift;
    
    this.miniStreamCutoffSize = this.data.getUint32(0x38, true);
  }

  private getSectorOffset(sectorIndex: number): number {
    return (sectorIndex + 1) * this.sectorSize;
  }

  private readFAT(): void {
    const numFATSectors = this.data.getUint32(0x2C, true);
    const firstDIFATSector = this.data.getInt32(0x44, true);
    const numDIFATSectors = this.data.getUint32(0x48, true);

    const difat: number[] = [];
    
    for (let i = 0; i < 109; i++) {
      const sector = this.data.getInt32(0x4C + i * 4, true);
      if (sector >= 0) {
        difat.push(sector);
      }
    }

    let difatSector = firstDIFATSector;
    while (difatSector >= 0 && difatSector !== ENDOFCHAIN) {
      const offset = this.getSectorOffset(difatSector);
      const entriesPerSector = (this.sectorSize / 4) - 1;
      
      for (let i = 0; i < entriesPerSector; i++) {
        const sector = this.data.getInt32(offset + i * 4, true);
        if (sector >= 0) {
          difat.push(sector);
        }
      }
      difatSector = this.data.getInt32(offset + entriesPerSector * 4, true);
    }

    this.fat = [];
    for (const fatSector of difat) {
      const offset = this.getSectorOffset(fatSector);
      const entriesPerSector = this.sectorSize / 4;
      
      for (let i = 0; i < entriesPerSector; i++) {
        this.fat.push(this.data.getInt32(offset + i * 4, true));
      }
    }
  }

  private readDirectoryEntries(): void {
    const firstDirSector = this.data.getInt32(0x30, true);
    const entriesPerSector = this.sectorSize / 128;
    
    const dirSectors = this.getSectorChain(firstDirSector);
    
    this.directoryEntries = [];
    for (const sector of dirSectors) {
      const offset = this.getSectorOffset(sector);
      
      for (let i = 0; i < entriesPerSector; i++) {
        const entryOffset = offset + i * 128;
        const entry = this.readDirectoryEntry(entryOffset);
        if (entry.type !== 0) {
          this.directoryEntries.push(entry);
        }
      }
    }
  }

  private readDirectoryEntry(offset: number): DirectoryEntry {
    const nameLength = this.data.getUint16(offset + 64, true);
    let name = '';
    for (let i = 0; i < Math.min(nameLength - 2, 62); i += 2) {
      const charCode = this.data.getUint16(offset + i, true);
      if (charCode === 0) break;
      name += String.fromCharCode(charCode);
    }

    const clsid = new Uint8Array(16);
    for (let i = 0; i < 16; i++) {
      clsid[i] = this.data.getUint8(offset + 80 + i);
    }

    return {
      name,
      type: this.data.getUint8(offset + 66),
      colorFlag: this.data.getUint8(offset + 67),
      leftSiblingId: this.data.getInt32(offset + 68, true),
      rightSiblingId: this.data.getInt32(offset + 72, true),
      childId: this.data.getInt32(offset + 76, true),
      startSectorLocation: this.data.getInt32(offset + 116, true),
      size: this.data.getUint32(offset + 120, true),
      clsid,
    };
  }

  private readMiniFAT(): void {
    const firstMiniFATSector = this.data.getInt32(0x3C, true);
    
    if (firstMiniFATSector < 0) {
      this.miniFat = [];
      return;
    }

    const miniFATSectors = this.getSectorChain(firstMiniFATSector);
    
    this.miniFat = [];
    for (const sector of miniFATSectors) {
      const offset = this.getSectorOffset(sector);
      const entriesPerSector = this.sectorSize / 4;
      
      for (let i = 0; i < entriesPerSector; i++) {
        this.miniFat.push(this.data.getInt32(offset + i * 4, true));
      }
    }
  }

  private readMiniStream(): void {
    const rootEntry = this.directoryEntries[0];
    if (!rootEntry || rootEntry.startSectorLocation < 0) {
      this.miniStreamData = new Uint8Array(0);
      return;
    }

    this.miniStreamData = this.readStream(rootEntry.startSectorLocation, rootEntry.size, false);
  }

  private getSectorChain(startSector: number): number[] {
    const chain: number[] = [];
    let sector = startSector;
    
    while (sector >= 0 && sector < this.fat.length && sector !== ENDOFCHAIN) {
      chain.push(sector);
      sector = this.fat[sector];
      
      if (chain.length > 1000000) {
        throw new Error('FAT chain too long - possible corruption');
      }
    }
    
    return chain;
  }

  private getMiniSectorChain(startSector: number): number[] {
    const chain: number[] = [];
    let sector = startSector;
    
    while (sector >= 0 && sector < this.miniFat.length && sector !== ENDOFCHAIN) {
      chain.push(sector);
      sector = this.miniFat[sector];
      
      if (chain.length > 1000000) {
        throw new Error('Mini FAT chain too long - possible corruption');
      }
    }
    
    return chain;
  }

  private readStream(startSector: number, size: number, useMiniStream: boolean): Uint8Array {
    const result = new Uint8Array(size);
    let offset = 0;

    if (useMiniStream) {
      const chain = this.getMiniSectorChain(startSector);
      for (const sector of chain) {
        const miniOffset = sector * this.miniSectorSize;
        const bytesToCopy = Math.min(this.miniSectorSize, size - offset);
        
        for (let i = 0; i < bytesToCopy; i++) {
          result[offset + i] = this.miniStreamData[miniOffset + i];
        }
        offset += bytesToCopy;
      }
    } else {
      const chain = this.getSectorChain(startSector);
      for (const sector of chain) {
        const sectorOffset = this.getSectorOffset(sector);
        const bytesToCopy = Math.min(this.sectorSize, size - offset);
        
        for (let i = 0; i < bytesToCopy; i++) {
          result[offset + i] = this.data.getUint8(sectorOffset + i);
        }
        offset += bytesToCopy;
      }
    }

    return result;
  }

  public getStreamNames(): string[] {
    return this.directoryEntries.map(e => e.name);
  }

  public findEntry(path: string): DirectoryEntry | null {
    const parts = path.split('/').filter(p => p.length > 0);
    
    if (parts.length === 0) {
      return this.directoryEntries[0] || null;
    }

    let currentEntry = this.directoryEntries[0];
    
    for (const part of parts) {
      if (!currentEntry || currentEntry.childId < 0) {
        return null;
      }

      const found = this.findInTree(currentEntry.childId, part);
      if (!found) {
        return null;
      }
      currentEntry = found;
    }

    return currentEntry;
  }

  private findInTree(entryId: number, name: string): DirectoryEntry | null {
    if (entryId < 0 || entryId >= this.directoryEntries.length) {
      return null;
    }

    const entry = this.directoryEntries.find((e, idx) => {
      return e.name.toLowerCase() === name.toLowerCase();
    });

    return entry || null;
  }

  public readStreamByName(name: string): Uint8Array | null {
    const entry = this.directoryEntries.find(e => e.name === name);
    if (!entry || entry.type !== 2) {
      return null;
    }

    const useMiniStream = entry.size < this.miniStreamCutoffSize;
    return this.readStream(entry.startSectorLocation, entry.size, useMiniStream);
  }

  public readStreamByPath(path: string): Uint8Array | null {
    const parts = path.split('/').filter(p => p.length > 0);
    const streamName = parts[parts.length - 1];
    
    const entry = this.directoryEntries.find(e => e.name === streamName && e.type === 2);
    if (!entry) {
      return null;
    }

    const useMiniStream = entry.size < this.miniStreamCutoffSize;
    return this.readStream(entry.startSectorLocation, entry.size, useMiniStream);
  }

  public listStreams(): { name: string; type: string; size: number }[] {
    return this.directoryEntries.map(e => ({
      name: e.name,
      type: e.type === 1 ? 'storage' : e.type === 2 ? 'stream' : e.type === 5 ? 'root' : 'unknown',
      size: e.size,
    }));
  }
}
