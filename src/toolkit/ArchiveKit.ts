import pako from 'pako';

interface ZipEntry {
  name: string;
  data: Uint8Array;
}

export const ArchiveKit = {
  async inflate(compressed: Uint8Array): Promise<Uint8Array> {
    return pako.inflate(compressed);
  },

  async deflate(data: Uint8Array): Promise<Uint8Array> {
    return pako.deflate(data, { level: 6 });
  },

  async unzip(zipData: Uint8Array): Promise<Map<string, Uint8Array>> {
    const files = new Map<string, Uint8Array>();
    const view = new DataView(zipData.buffer, zipData.byteOffset, zipData.byteLength);

    let offset = 0;
    while (offset < zipData.length - 4) {
      const sig = view.getUint32(offset, true);

      if (sig === 0x04034b50) {
        const compressionMethod = view.getUint16(offset + 8, true);
        const compressedSize    = view.getUint32(offset + 18, true);
        const uncompressedSize  = view.getUint32(offset + 22, true);
        const fileNameLength    = view.getUint16(offset + 26, true);
        const extraLength       = view.getUint16(offset + 28, true);

        const nameBytes = zipData.subarray(offset + 30, offset + 30 + fileNameLength);
        const name = new TextDecoder('utf-8').decode(nameBytes);

        const dataOffset = offset + 30 + fileNameLength + extraLength;
        let fileData: Uint8Array;

        if (compressionMethod === 0) {
          fileData = zipData.subarray(dataOffset, dataOffset + uncompressedSize);
        } else if (compressionMethod === 8) {
          const compressed = zipData.subarray(dataOffset, dataOffset + compressedSize);
          fileData = pako.inflateRaw(compressed);
        } else {
          throw new Error(`Unsupported ZIP compression method: ${compressionMethod}`);
        }

        files.set(name, new Uint8Array(fileData));
        offset = dataOffset + compressedSize;
      } else if (sig === 0x02014b50 || sig === 0x06054b50) {
        break;
      } else {
        offset++;
      }
    }

    return files;
  },

  async zip(entries: ZipEntry[]): Promise<Uint8Array> {
    const localHeaders: Uint8Array[] = [];
    const centralHeaders: Uint8Array[] = [];
    let localOffset = 0;

    for (const entry of entries) {
      const nameBytes = new TextEncoder().encode(entry.name);
      const crc = crc32(entry.data);

      // 'mimetype' and 'version.xml' must be stored uncompressed per HWPX spec
      const store = entry.name === 'mimetype' || entry.name === 'version.xml';
      const method = store ? 0 : 8;
      const payload = store ? entry.data : pako.deflateRaw(entry.data, { level: 6 });

      // Local file header (30 bytes + name + data)
      const local = new Uint8Array(30 + nameBytes.length + payload.length);
      const lv = new DataView(local.buffer);
      lv.setUint32(0, 0x04034b50, true);
      lv.setUint16(4, 20, true);
      lv.setUint16(6, 0, true);
      lv.setUint16(8, method, true);
      lv.setUint16(10, 0, true);
      lv.setUint16(12, 0x0021, true); // date (1980-01-01)
      lv.setUint32(14, crc, true);
      lv.setUint32(18, payload.length, true);
      lv.setUint32(22, entry.data.length, true);
      lv.setUint16(26, nameBytes.length, true);
      lv.setUint16(28, 0, true);
      local.set(nameBytes, 30);
      local.set(payload, 30 + nameBytes.length);

      // Central directory header (46 bytes + name)
      const central = new Uint8Array(46 + nameBytes.length);
      const cv = new DataView(central.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(8, 0, true);
      cv.setUint16(10, method, true);
      cv.setUint16(12, 0, true); // mod time
      cv.setUint16(14, 0x0021, true); // mod date (1980-01-01)
      cv.setUint32(16, crc, true);
      cv.setUint32(20, payload.length, true);
      cv.setUint32(24, entry.data.length, true);
      cv.setUint16(28, nameBytes.length, true);
      cv.setUint16(30, 0, true);
      cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true);
      cv.setUint16(36, 0, true);
      cv.setUint32(38, 0, true);
      cv.setUint32(42, localOffset, true);
      central.set(nameBytes, 46);

      localHeaders.push(local);
      centralHeaders.push(central);
      localOffset += local.length;
    }

    const centralDir = concat(centralHeaders);
    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(4, 0, true);
    ev.setUint16(6, 0, true);
    ev.setUint16(8, entries.length, true);
    ev.setUint16(10, entries.length, true);
    ev.setUint32(12, centralDir.length, true);
    ev.setUint32(16, localOffset, true);
    ev.setUint16(20, 0, true);

    return concat([...localHeaders, centralDir, eocd]);
  },
};

function concat(arrays: Uint8Array[]): Uint8Array {
  const total = arrays.reduce((s, a) => s + a.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const a of arrays) { out.set(a, offset); offset += a.length; }
  return out;
}

function crc32(data: Uint8Array): number {
  let crc = 0xFFFFFFFF;
  for (const byte of data) {
    crc ^= byte;
    for (let i = 0; i < 8; i++) {
      crc = (crc & 1) ? (crc >>> 1) ^ 0xEDB88320 : crc >>> 1;
    }
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}
