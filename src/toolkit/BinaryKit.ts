/**
 * OLE2 Compound File Binary Format (CFB) parser.
 * Used for legacy HWP 5.0 files.
 */

export const BinaryKit = {
  readU16LE(buf: Uint8Array, offset: number): number {
    return buf[offset] | (buf[offset + 1] << 8);
  },

  readU32LE(buf: Uint8Array, offset: number): number {
    return (
      (buf[offset] | (buf[offset + 1] << 8) | (buf[offset + 2] << 16)) >>> 0
    ) + buf[offset + 3] * 0x1000000;
  },

  isOle2(data: Uint8Array): boolean {
    return (
      data.length >= 8 &&
      data[0] === 0xD0 && data[1] === 0xCF &&
      data[2] === 0x11 && data[3] === 0xE0 &&
      data[4] === 0xA1 && data[5] === 0xB1 &&
      data[6] === 0x1A && data[7] === 0xE1
    );
  },

  parseCfb(data: Uint8Array): Map<string, Uint8Array> {
    const streams = new Map<string, Uint8Array>();

    if (!this.isOle2(data) || data.length < 512) {
      throw new Error('Not a valid OLE2 file');
    }

    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    const sectorSize   = 1 << view.getUint16(30, true);
    const miniSectorSz = 1 << view.getUint16(32, true);
    const majorVersion = view.getUint16(26, true);
    if (!((majorVersion === 3 && sectorSize === 512) || (majorVersion === 4 && sectorSize === 4096)) || miniSectorSz !== 64) {
      throw new Error('Unsupported CFB sector size/version');
    }
    const dirFirstSec  = view.getUint32(48, true);
    const miniStreamCutoff = view.getUint32(56, true);
    const miniFatFirst = view.getUint32(60, true);
    const miniFatCnt   = view.getUint32(64, true);
    const difatFirst   = view.getUint32(68, true);

    const ENDOFCHAIN = 0xFFFFFFFE;
    const FREESECT   = 0xFFFFFFFF;

    const sectorAt = (sec: number): Uint8Array => {
      const offset = (sec + 1) * sectorSize;
      if (!Number.isInteger(sec) || sec < 0 || offset + sectorSize > data.length) throw new Error('CFB sector outside file');
      return data.subarray(offset, offset + sectorSize);
    };

    // Build FAT from DIFAT
    const fatSecNums: number[] = [];
    for (let i = 0; i < 109; i++) {
      const s = view.getUint32(76 + i * 4, true);
      if (s === FREESECT || s === ENDOFCHAIN) break;
      fatSecNums.push(s);
    }
    if (difatFirst !== ENDOFCHAIN && difatFirst !== FREESECT) {
      let difSec = difatFirst;
      const seen = new Set<number>();
      while (difSec !== ENDOFCHAIN && difSec !== FREESECT) {
        if (seen.has(difSec)) throw new Error('Cyclic CFB DIFAT chain');
        seen.add(difSec);
        const sec = sectorAt(difSec);
        const sv = new DataView(sec.buffer, sec.byteOffset, sec.byteLength);
        for (let i = 0; i < (sectorSize / 4) - 1; i++) {
          const s = sv.getUint32(i * 4, true);
          if (s === FREESECT || s === ENDOFCHAIN) break;
          fatSecNums.push(s);
        }
        difSec = sv.getUint32(sectorSize - 4, true);
      }
    }

    const fat: number[] = [];
    for (const sec of fatSecNums) {
      const s = sectorAt(sec);
      const sv = new DataView(s.buffer, s.byteOffset, s.byteLength);
      for (let i = 0; i < sectorSize / 4; i++) {
        fat.push(sv.getUint32(i * 4, true));
      }
    }

    const readChain = (startSec: number): Uint8Array => {
      const chunks: Uint8Array[] = [];
      let sec = startSec;
      const seen = new Set<number>();
      while (sec !== ENDOFCHAIN && sec !== FREESECT && sec < fat.length) {
        if (seen.has(sec)) throw new Error('Cyclic CFB FAT chain');
        seen.add(sec);
        chunks.push(sectorAt(sec));
        sec = fat[sec];
      }
      return concatUint8(chunks);
    };

    // Directory entries
    const dirData = readChain(dirFirstSec);
    const dirView = new DataView(dirData.buffer, dirData.byteOffset, dirData.byteLength);
    const dirCount = dirData.length / 128;

    interface DirEntry {
      name: string;
      type: number;
      startSec: number;
      size: number;
      childId: number;
      siblingLeftId: number;
      siblingRightId: number;
    }

    const dirEntries: DirEntry[] = [];
    for (let i = 0; i < dirCount; i++) {
      const base = i * 128;
      const nameLen = dirView.getUint16(base + 64, true);
      const nameBytes = dirData.subarray(base, base + Math.max(0, nameLen - 2));
      const name = new TextDecoder('utf-16le').decode(nameBytes);
      const type = dirData[base + 66];
      const childId      = dirView.getInt32(base + 76, true);
      const sibLeft       = dirView.getInt32(base + 68, true);
      const sibRight      = dirView.getInt32(base + 72, true);
      const startSec     = dirView.getUint32(base + 116, true);
      const size         = dirView.getUint32(base + 120, true);
      dirEntries.push({ name, type, startSec, size, childId, siblingLeftId: sibLeft, siblingRightId: sibRight });
    }

    // Mini stream
    const rootEntry = dirEntries[0];
    let miniStreamData: Uint8Array | null = null;
    let miniFat: number[] = [];

    if (rootEntry && rootEntry.startSec !== ENDOFCHAIN && rootEntry.startSec !== FREESECT) {
      miniStreamData = readChain(rootEntry.startSec);
    }

    if (miniFatCnt > 0 && miniFatFirst !== ENDOFCHAIN && miniFatFirst !== FREESECT) {
      const mfData = readChain(miniFatFirst);
      const mfv = new DataView(mfData.buffer, mfData.byteOffset, mfData.byteLength);
      for (let i = 0; i < mfData.length / 4; i++) {
        miniFat.push(mfv.getUint32(i * 4, true));
      }
    }

    const readMiniChain = (startSec: number, size: number): Uint8Array => {
      if (!miniStreamData) return new Uint8Array(0);
      const chunks: Uint8Array[] = [];
      let sec = startSec;
      let remaining = size;
      const seen = new Set<number>();
      while (sec !== ENDOFCHAIN && sec !== FREESECT && sec < miniFat.length && remaining > 0) {
        if (seen.has(sec)) throw new Error('Cyclic CFB mini FAT chain');
        seen.add(sec);
        const off = sec * miniSectorSz;
        const chunk = miniStreamData.subarray(off, off + Math.min(miniSectorSz, remaining));
        if (chunk.length === 0) throw new Error('CFB mini sector outside stream');
        chunks.push(chunk);
        remaining -= chunk.length;
        sec = miniFat[sec];
      }
      return concatUint8(chunks).subarray(0, size);
    };

    // DFS traversal
    const visited = new Set<number>();
    const visit = (id: number, path: string): void => {
      if (id < 0 || id >= dirEntries.length) return;
      if (visited.has(id)) throw new Error('Cyclic CFB directory');
      visited.add(id);
      const entry = dirEntries[id];
      const fullPath = path ? `${path}/${entry.name}` : entry.name;

      if (entry.type === 2) {
        let streamData: Uint8Array;
        if (entry.size < miniStreamCutoff && miniStreamData) {
          streamData = readMiniChain(entry.startSec, entry.size);
        } else {
          streamData = readChain(entry.startSec).subarray(0, entry.size);
        }
        streams.set(fullPath, streamData);
        streams.set(entry.name, streamData);
      }

      if (entry.childId >= 0) visit(entry.childId, fullPath);
      if (entry.siblingLeftId >= 0) visit(entry.siblingLeftId, path);
      if (entry.siblingRightId >= 0) visit(entry.siblingRightId, path);
    };

    if (dirEntries.length > 0 && dirEntries[0].childId >= 0) {
      visit(dirEntries[0].childId, '');
    }

    return streams;
  },
};

function concatUint8(arrays: Uint8Array[]): Uint8Array {
  const total = arrays.reduce((s, a) => s + a.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const a of arrays) { out.set(a, off); off += a.length; }
  return out;
}
