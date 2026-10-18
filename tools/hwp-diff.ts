#!/usr/bin/env -S pnpm exec vite-node

/**
 * Inspect or compare HWP 5.x Compound File Binary (CFB) files.
 *
 * Usage:
 *   pnpm hwp:diff file.hwp
 *   pnpm hwp:diff left.hwp right.hwp
 */

import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { inflateRawSync } from "node:zlib";

const CFB_MAGIC = Uint8Array.from([
  0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1,
]);
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;

const HWPTAG_BEGIN = 0x10;

/** All DocInfo tags published in Hancom's HWP 5.0 revision 1.3 spec. */
export const DOC_INFO_TAG_NAMES = new Map<number, string>([
  [HWPTAG_BEGIN + 0, "HWPTAG_DOCUMENT_PROPERTIES"],
  [HWPTAG_BEGIN + 1, "HWPTAG_ID_MAPPINGS"],
  [HWPTAG_BEGIN + 2, "HWPTAG_BIN_DATA"],
  [HWPTAG_BEGIN + 3, "HWPTAG_FACE_NAME"],
  [HWPTAG_BEGIN + 4, "HWPTAG_BORDER_FILL"],
  [HWPTAG_BEGIN + 5, "HWPTAG_CHAR_SHAPE"],
  [HWPTAG_BEGIN + 6, "HWPTAG_TAB_DEF"],
  [HWPTAG_BEGIN + 7, "HWPTAG_NUMBERING"],
  [HWPTAG_BEGIN + 8, "HWPTAG_BULLET"],
  [HWPTAG_BEGIN + 9, "HWPTAG_PARA_SHAPE"],
  [HWPTAG_BEGIN + 10, "HWPTAG_STYLE"],
  [HWPTAG_BEGIN + 11, "HWPTAG_DOC_DATA"],
  [HWPTAG_BEGIN + 12, "HWPTAG_DISTRIBUTE_DOC_DATA"],
  [HWPTAG_BEGIN + 13, "RESERVED_DOC_INFO_13"],
  [HWPTAG_BEGIN + 14, "HWPTAG_COMPATIBLE_DOCUMENT"],
  [HWPTAG_BEGIN + 15, "HWPTAG_LAYOUT_COMPATIBILITY"],
  [HWPTAG_BEGIN + 16, "HWPTAG_TRACKCHANGE"],
  [HWPTAG_BEGIN + 76, "HWPTAG_MEMO_SHAPE"],
  [HWPTAG_BEGIN + 78, "HWPTAG_FORBIDDEN_CHAR"],
  [HWPTAG_BEGIN + 80, "HWPTAG_TRACK_CHANGE"],
  [HWPTAG_BEGIN + 81, "HWPTAG_TRACK_CHANGE_AUTHOR"],
]);

/** All BodyText tags published in Hancom's HWP 5.0 revision 1.3 spec. */
export const BODY_TEXT_TAG_NAMES = new Map<number, string>([
  [HWPTAG_BEGIN + 50, "HWPTAG_PARA_HEADER"],
  [HWPTAG_BEGIN + 51, "HWPTAG_PARA_TEXT"],
  [HWPTAG_BEGIN + 52, "HWPTAG_PARA_CHAR_SHAPE"],
  [HWPTAG_BEGIN + 53, "HWPTAG_PARA_LINE_SEG"],
  [HWPTAG_BEGIN + 54, "HWPTAG_PARA_RANGE_TAG"],
  [HWPTAG_BEGIN + 55, "HWPTAG_CTRL_HEADER"],
  [HWPTAG_BEGIN + 56, "HWPTAG_LIST_HEADER"],
  [HWPTAG_BEGIN + 57, "HWPTAG_PAGE_DEF"],
  [HWPTAG_BEGIN + 58, "HWPTAG_FOOTNOTE_SHAPE"],
  [HWPTAG_BEGIN + 59, "HWPTAG_PAGE_BORDER_FILL"],
  [HWPTAG_BEGIN + 60, "HWPTAG_SHAPE_COMPONENT"],
  [HWPTAG_BEGIN + 61, "HWPTAG_TABLE"],
  [HWPTAG_BEGIN + 62, "HWPTAG_SHAPE_COMPONENT_LINE"],
  [HWPTAG_BEGIN + 63, "HWPTAG_SHAPE_COMPONENT_RECTANGLE"],
  [HWPTAG_BEGIN + 64, "HWPTAG_SHAPE_COMPONENT_ELLIPSE"],
  [HWPTAG_BEGIN + 65, "HWPTAG_SHAPE_COMPONENT_ARC"],
  [HWPTAG_BEGIN + 66, "HWPTAG_SHAPE_COMPONENT_POLYGON"],
  [HWPTAG_BEGIN + 67, "HWPTAG_SHAPE_COMPONENT_CURVE"],
  [HWPTAG_BEGIN + 68, "HWPTAG_SHAPE_COMPONENT_OLE"],
  [HWPTAG_BEGIN + 69, "HWPTAG_SHAPE_COMPONENT_PICTURE"],
  [HWPTAG_BEGIN + 70, "HWPTAG_SHAPE_COMPONENT_CONTAINER"],
  [HWPTAG_BEGIN + 71, "HWPTAG_CTRL_DATA"],
  [HWPTAG_BEGIN + 72, "HWPTAG_EQEDIT"],
  [HWPTAG_BEGIN + 73, "RESERVED_BODY_TEXT_73"],
  [HWPTAG_BEGIN + 74, "HWPTAG_SHAPE_COMPONENT_TEXTART"],
  [HWPTAG_BEGIN + 75, "HWPTAG_FORM_OBJECT"],
  [HWPTAG_BEGIN + 76, "HWPTAG_MEMO_SHAPE"],
  [HWPTAG_BEGIN + 77, "HWPTAG_MEMO_LIST"],
  [HWPTAG_BEGIN + 79, "HWPTAG_CHART_DATA"],
  [HWPTAG_BEGIN + 82, "HWPTAG_VIDEO_DATA"],
  [HWPTAG_BEGIN + 99, "HWPTAG_SHAPE_COMPONENT_UNKNOWN"],
]);

const ID_MAPPING_ITEMS = [
  { name: "BIN_DATA", tag: HWPTAG_BEGIN + 2 },
  { name: "FACE_NAME_HANGUL", tag: HWPTAG_BEGIN + 3 },
  { name: "FACE_NAME_LATIN", tag: HWPTAG_BEGIN + 3 },
  { name: "FACE_NAME_HANJA", tag: HWPTAG_BEGIN + 3 },
  { name: "FACE_NAME_JAPANESE", tag: HWPTAG_BEGIN + 3 },
  { name: "FACE_NAME_OTHER", tag: HWPTAG_BEGIN + 3 },
  { name: "FACE_NAME_SYMBOL", tag: HWPTAG_BEGIN + 3 },
  { name: "FACE_NAME_USER", tag: HWPTAG_BEGIN + 3 },
  { name: "BORDER_FILL", tag: HWPTAG_BEGIN + 4 },
  { name: "CHAR_SHAPE", tag: HWPTAG_BEGIN + 5 },
  { name: "TAB_DEF", tag: HWPTAG_BEGIN + 6 },
  { name: "NUMBERING", tag: HWPTAG_BEGIN + 7 },
  { name: "BULLET", tag: HWPTAG_BEGIN + 8 },
  { name: "PARA_SHAPE", tag: HWPTAG_BEGIN + 9 },
  { name: "STYLE", tag: HWPTAG_BEGIN + 10 },
  { name: "MEMO_SHAPE", tag: HWPTAG_BEGIN + 76 },
  { name: "TRACK_CHANGE", tag: HWPTAG_BEGIN + 80 },
  { name: "TRACK_CHANGE_AUTHOR", tag: HWPTAG_BEGIN + 81 },
] as const;

interface CfbStream {
  path: string;
  data: Uint8Array;
}

interface DirectoryEntry {
  name: string;
  type: number;
  left: number;
  right: number;
  child: number;
  startSector: number;
  size: number;
}

export interface HwpRecord {
  tag: number;
  level: number;
  size: number;
  offset: number;
  headerSize: 4 | 8;
  data: Uint8Array;
}

interface RecordStream {
  path: string;
  compressedSize: number;
  raw: Uint8Array;
  records: HwpRecord[];
}

interface Analysis {
  inputPath: string;
  fileSize: number;
  streams: CfbStream[];
  fileHeader?: Uint8Array;
  flags?: number;
  recordStreams: Map<string, RecordStream>;
  errors: string[];
}

function readU32(data: Uint8Array, offset: number): number {
  if (offset < 0 || offset + 4 > data.length) {
    throw new Error(`UINT32 out of bounds at ${offset}`);
  }
  return new DataView(data.buffer, data.byteOffset, data.byteLength).getUint32(
    offset,
    true,
  );
}

function concat(chunks: Uint8Array[]): Uint8Array {
  const size = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const output = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.length;
  }
  return output;
}

/** Parse CFB while retaining canonical storage-qualified stream paths. */
export function parseCfbStreams(file: Uint8Array): CfbStream[] {
  if (
    file.length < 512 ||
    !CFB_MAGIC.every((value, index) => file[index] === value)
  ) {
    throw new Error("not a CFB/OLE2 file (magic mismatch)");
  }

  const header = new DataView(file.buffer, file.byteOffset, file.byteLength);
  const majorVersion = header.getUint16(26, true);
  const sectorSize = 1 << header.getUint16(30, true);
  const miniSectorSize = 1 << header.getUint16(32, true);
  const fatSectorCount = header.getUint32(44, true);
  const directoryStart = header.getUint32(48, true);
  const miniStreamCutoff = header.getUint32(56, true);
  const miniFatStart = header.getUint32(60, true);
  const miniFatSectorCount = header.getUint32(64, true);
  const difatStart = header.getUint32(68, true);
  const difatSectorCount = header.getUint32(72, true);

  if (majorVersion !== 3 && majorVersion !== 4) {
    throw new Error(`unsupported CFB major version ${majorVersion}`);
  }
  if (sectorSize !== (majorVersion === 3 ? 512 : 4096)) {
    throw new Error(`invalid CFB sector size ${sectorSize}`);
  }
  if (miniSectorSize <= 0 || miniSectorSize > sectorSize) {
    throw new Error(`invalid CFB mini-sector size ${miniSectorSize}`);
  }

  const sectorCount = Math.floor((file.length - 512) / sectorSize);
  const sectorAt = (sector: number): Uint8Array => {
    if (!Number.isInteger(sector) || sector < 0 || sector >= sectorCount) {
      throw new Error(`CFB sector ${sector} is outside 0..${sectorCount - 1}`);
    }
    const start = 512 + sector * sectorSize;
    return file.subarray(start, start + sectorSize);
  };

  const fatSectorIds: number[] = [];
  for (let index = 0; index < 109; index++) {
    const sector = header.getUint32(76 + index * 4, true);
    if (sector !== FREESECT && sector !== ENDOFCHAIN) fatSectorIds.push(sector);
  }

  let difatSector = difatStart;
  const seenDifat = new Set<number>();
  for (let index = 0; index < difatSectorCount; index++) {
    if (difatSector === ENDOFCHAIN || difatSector === FREESECT) break;
    if (seenDifat.has(difatSector)) throw new Error("cycle in CFB DIFAT chain");
    seenDifat.add(difatSector);
    const bytes = sectorAt(difatSector);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    for (let slot = 0; slot < sectorSize / 4 - 1; slot++) {
      const sector = view.getUint32(slot * 4, true);
      if (sector !== FREESECT && sector !== ENDOFCHAIN)
        fatSectorIds.push(sector);
    }
    difatSector = view.getUint32(sectorSize - 4, true);
  }
  if (fatSectorIds.length < fatSectorCount) {
    throw new Error(
      `CFB declares ${fatSectorCount} FAT sectors but DIFAT exposes ${fatSectorIds.length}`,
    );
  }

  const fat: number[] = [];
  for (const sector of fatSectorIds.slice(0, fatSectorCount)) {
    const bytes = sectorAt(sector);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    for (let offset = 0; offset < bytes.length; offset += 4)
      fat.push(view.getUint32(offset, true));
  }

  const readRegularChain = (
    startSector: number,
    requestedSize?: number,
  ): Uint8Array => {
    if (requestedSize === 0) return new Uint8Array();
    const chunks: Uint8Array[] = [];
    const seen = new Set<number>();
    let sector = startSector;
    while (sector !== ENDOFCHAIN && sector !== FREESECT) {
      if (sector >= fat.length) throw new Error(`FAT entry ${sector} is absent`);
      if (seen.has(sector)) throw new Error(`cycle in CFB FAT chain at ${sector}`);
      seen.add(sector);
      chunks.push(sectorAt(sector));
      if (
        requestedSize !== undefined &&
        chunks.length * sectorSize >= requestedSize
      )
        break;
      sector = fat[sector];
    }
    const bytes = concat(chunks);
    return requestedSize === undefined ? bytes : bytes.subarray(0, requestedSize);
  };

  const directoryBytes = readRegularChain(directoryStart);
  if (directoryBytes.length % 128 !== 0) {
    throw new Error("CFB directory stream is not 128-byte aligned");
  }
  const directoryView = new DataView(
    directoryBytes.buffer,
    directoryBytes.byteOffset,
    directoryBytes.byteLength,
  );
  const entries: DirectoryEntry[] = [];
  for (let base = 0; base < directoryBytes.length; base += 128) {
    const nameLength = directoryView.getUint16(base + 64, true);
    if (nameLength > 64 || (nameLength & 1) !== 0) {
      throw new Error(`invalid CFB directory name length ${nameLength}`);
    }
    const name =
      nameLength >= 2
        ? new TextDecoder("utf-16le").decode(
            directoryBytes.subarray(base, base + nameLength - 2),
          )
        : "";
    const lowSize = directoryView.getUint32(base + 120, true);
    const highSize = directoryView.getUint32(base + 124, true);
    const size = majorVersion === 3 ? lowSize : highSize * 0x100000000 + lowSize;
    if (!Number.isSafeInteger(size)) throw new Error(`unsafe stream size for ${name}`);
    entries.push({
      name,
      type: directoryBytes[base + 66],
      left: directoryView.getInt32(base + 68, true),
      right: directoryView.getInt32(base + 72, true),
      child: directoryView.getInt32(base + 76, true),
      startSector: directoryView.getUint32(base + 116, true),
      size,
    });
  }
  if (!entries[0] || entries[0].type !== 5) {
    throw new Error("CFB root directory entry is absent");
  }

  const root = entries[0];
  const miniStream = readRegularChain(root.startSector, root.size);
  const miniFatBytes =
    miniFatSectorCount > 0
      ? readRegularChain(miniFatStart, miniFatSectorCount * sectorSize)
      : new Uint8Array();
  const miniFatView = new DataView(
    miniFatBytes.buffer,
    miniFatBytes.byteOffset,
    miniFatBytes.byteLength,
  );
  const miniFat: number[] = [];
  for (let offset = 0; offset + 4 <= miniFatBytes.length; offset += 4)
    miniFat.push(miniFatView.getUint32(offset, true));

  const readMiniChain = (startSector: number, size: number): Uint8Array => {
    if (size === 0) return new Uint8Array();
    const chunks: Uint8Array[] = [];
    const seen = new Set<number>();
    let sector = startSector;
    let remaining = size;
    while (
      remaining > 0 &&
      sector !== ENDOFCHAIN &&
      sector !== FREESECT
    ) {
      if (sector >= miniFat.length)
        throw new Error(`mini FAT entry ${sector} is absent`);
      if (seen.has(sector))
        throw new Error(`cycle in CFB mini FAT chain at ${sector}`);
      seen.add(sector);
      const offset = sector * miniSectorSize;
      if (offset >= miniStream.length)
        throw new Error(`mini-sector ${sector} is outside the mini stream`);
      const chunk = miniStream.subarray(
        offset,
        Math.min(offset + miniSectorSize, offset + remaining),
      );
      chunks.push(chunk);
      remaining -= chunk.length;
      sector = miniFat[sector];
    }
    const bytes = concat(chunks).subarray(0, size);
    if (bytes.length !== size)
      throw new Error(`mini stream truncated: expected ${size}, got ${bytes.length}`);
    return bytes;
  };

  const streams: CfbStream[] = [];
  const visited = new Set<number>();
  const walkSiblingTree = (entryId: number, parentPath: string): void => {
    if (entryId < 0) return;
    if (entryId >= entries.length)
      throw new Error(`directory entry ${entryId} is absent`);
    if (visited.has(entryId))
      throw new Error(`cycle/duplicate in directory tree at entry ${entryId}`);
    visited.add(entryId);
    const entry = entries[entryId];

    walkSiblingTree(entry.left, parentPath);
    const path = parentPath ? `${parentPath}/${entry.name}` : entry.name;
    if (entry.type === 2) {
      const data =
        entry.size < miniStreamCutoff
          ? readMiniChain(entry.startSector, entry.size)
          : readRegularChain(entry.startSector, entry.size);
      if (data.length !== entry.size) {
        throw new Error(
          `stream ${path} truncated: expected ${entry.size}, got ${data.length}`,
        );
      }
      streams.push({ path, data });
    } else if (entry.type === 1 && entry.child >= 0) {
      walkSiblingTree(entry.child, path);
    }
    walkSiblingTree(entry.right, parentPath);
  };

  if (root.child >= 0) walkSiblingTree(root.child, "");
  return streams.sort((left, right) => left.path.localeCompare(right.path));
}

/** Strict HWP record parser: success means every byte belongs to one record. */
export function parseRecords(data: Uint8Array): HwpRecord[] {
  const records: HwpRecord[] = [];
  let offset = 0;
  while (offset < data.length) {
    const recordOffset = offset;
    if (data.length - offset < 4) {
      throw new Error(`${data.length - offset} trailing byte(s) at ${offset}`);
    }
    const header = readU32(data, offset);
    const tag = header & 0x3ff;
    const level = (header >>> 10) & 0x3ff;
    let size = (header >>> 20) & 0xfff;
    offset += 4;
    let headerSize: 4 | 8 = 4;
    if (size === 0xfff) {
      if (data.length - offset < 4) {
        throw new Error(`extended size missing at ${recordOffset}`);
      }
      size = readU32(data, offset);
      offset += 4;
      headerSize = 8;
    }
    if (size > data.length - offset) {
      throw new Error(
        `record at ${recordOffset} declares ${size} byte(s), only ${data.length - offset} remain`,
      );
    }
    records.push({
      tag,
      level,
      size,
      offset: recordOffset,
      headerSize,
      data: data.subarray(offset, offset + size),
    });
    offset += size;
  }
  return records;
}

function isRecordStream(path: string): boolean {
  return path === "DocInfo" || /(?:^|\/)Section\d+$/.test(path);
}

function tagName(path: string, tag: number): string {
  const table = path === "DocInfo" ? DOC_INFO_TAG_NAMES : BODY_TEXT_TAG_NAMES;
  return table.get(tag) ?? `UNKNOWN_TAG_0x${tag.toString(16).toUpperCase()}`;
}

function sha256(data: Uint8Array, length = 12): string {
  return createHash("sha256").update(data).digest("hex").slice(0, length);
}

function analyze(inputPath: string): Analysis {
  const file = new Uint8Array(readFileSync(inputPath));
  const analysis: Analysis = {
    inputPath,
    fileSize: file.length,
    streams: [],
    recordStreams: new Map(),
    errors: [],
  };
  try {
    analysis.streams = parseCfbStreams(file);
  } catch (error) {
    analysis.errors.push(`CFB: ${error instanceof Error ? error.message : error}`);
    return analysis;
  }

  const fileHeader = analysis.streams.find(
    (stream) => stream.path === "FileHeader",
  )?.data;
  analysis.fileHeader = fileHeader;
  if (!fileHeader) {
    analysis.errors.push("FileHeader stream is missing");
  } else if (fileHeader.length < 40) {
    analysis.errors.push(`FileHeader is too short (${fileHeader.length} bytes)`);
  } else {
    analysis.flags = readU32(fileHeader, 36);
  }

  for (const stream of analysis.streams.filter((item) =>
    isRecordStream(item.path),
  )) {
    try {
      // HWP's compressed streams use raw DEFLATE: zlib windowBits = -15.
      const raw =
        (analysis.flags ?? 1) & 1
          ? new Uint8Array(inflateRawSync(stream.data))
          : stream.data;
      const records = parseRecords(raw);
      analysis.recordStreams.set(stream.path, {
        path: stream.path,
        compressedSize: stream.data.length,
        raw,
        records,
      });
    } catch (error) {
      analysis.errors.push(
        `${stream.path}: ${error instanceof Error ? error.message : error}`,
      );
    }
  }
  return analysis;
}

function countTag(records: HwpRecord[], tag: number): number {
  return records.reduce((count, record) => count + (record.tag === tag ? 1 : 0), 0);
}

function formatIdMappings(records: HwpRecord[]): {
  lines: string[];
  errors: string[];
} {
  const lines: string[] = [];
  const errors: string[] = [];
  const mappingRecords = records.filter(
    (record) => record.tag === HWPTAG_BEGIN + 1,
  );
  if (mappingRecords.length !== 1) {
    errors.push(
      `ID_MAPPINGS record count is ${mappingRecords.length}, expected exactly 1`,
    );
  }
  const mapping = mappingRecords[0];
  if (!mapping) return { lines, errors };
  if (mapping.data.length < ID_MAPPING_ITEMS.length * 4) {
    errors.push(
      `ID_MAPPINGS is ${mapping.data.length} bytes, expected at least ${ID_MAPPING_ITEMS.length * 4}`,
    );
    return { lines, errors };
  }

  const view = new DataView(
    mapping.data.buffer,
    mapping.data.byteOffset,
    mapping.data.byteLength,
  );
  const expected = ID_MAPPING_ITEMS.map((_, index) =>
    view.getInt32(index * 4, true),
  );
  for (let index = 0; index < expected.length; index++) {
    if (expected[index] < 0)
      errors.push(`ID_MAPPINGS[${index}] is negative (${expected[index]})`);
  }

  const actual = ID_MAPPING_ITEMS.map((item) => countTag(records, item.tag));
  const faceTotal = countTag(records, HWPTAG_BEGIN + 3);
  let remainingFaces = faceTotal;
  // FACE_NAME records contain no language-group discriminator. HWP stores the
  // seven groups consecutively, so allocate the observed total in that order
  // and additionally perform the independent aggregate check below.
  for (let index = 1; index <= 7; index++) {
    actual[index] = Math.min(Math.max(expected[index], 0), remainingFaces);
    remainingFaces -= actual[index];
  }

  for (let index = 0; index < ID_MAPPING_ITEMS.length; index++) {
    const ok = expected[index] === actual[index];
    lines.push(
      `[${index.toString().padStart(2, "0")}] ${ID_MAPPING_ITEMS[index].name.padEnd(24)} expected=${expected[index]} actual=${actual[index]} ${ok ? "OK" : "ERROR"}`,
    );
    if (!ok) {
      errors.push(
        `ID_MAPPINGS[${index}] ${ID_MAPPING_ITEMS[index].name}: expected ${expected[index]}, actual ${actual[index]}`,
      );
    }
  }
  const expectedFaceTotal = expected
    .slice(1, 8)
    .reduce((sum, count) => sum + count, 0);
  const facesOk = expectedFaceTotal === faceTotal;
  lines.push(
    `FACE_NAME aggregate          expected=${expectedFaceTotal} actual=${faceTotal} ${facesOk ? "OK" : "ERROR"}`,
  );
  lines.push(
    "note: per-language FACE_NAME actuals are order-inferred; only the aggregate is independently countable.",
  );
  if (!facesOk) {
    errors.push(
      `FACE_NAME aggregate: expected ${expectedFaceTotal}, actual ${faceTotal}`,
    );
  }
  return { lines, errors };
}

function recordLine(path: string, record: HwpRecord): string {
  const indent = "  ".repeat(Math.min(record.level, 40));
  const levelSuffix = record.level > 40 ? `[level=${record.level}] ` : "";
  return `${indent}${levelSuffix}@0x${record.offset.toString(16).padStart(8, "0")} ${tagName(path, record.tag)} (tag=0x${record.tag.toString(16).padStart(3, "0")}, level=${record.level}, size=${record.size}, header=${record.headerSize}, sha256=${sha256(record.data)})`;
}

function dump(analysis: Analysis): string {
  const lines: string[] = [];
  lines.push(`# HWP dump: ${analysis.inputPath}`);
  lines.push(`file-size: ${analysis.fileSize} bytes`);
  lines.push("");
  lines.push(`## CFB streams (${analysis.streams.length})`);
  for (const stream of analysis.streams)
    lines.push(`- ${stream.path}: ${stream.data.length} bytes`);

  if (analysis.fileHeader) {
    const signature = new TextDecoder("ascii")
      .decode(analysis.fileHeader.subarray(0, 32))
      .replace(/\0+$/, "");
    const version =
      analysis.fileHeader.length >= 36 ? readU32(analysis.fileHeader, 32) : 0;
    lines.push("");
    lines.push("## FileHeader");
    lines.push(`size: ${analysis.fileHeader.length} bytes`);
    lines.push(`signature: ${JSON.stringify(signature)}`);
    lines.push(`version: 0x${version.toString(16).padStart(8, "0")}`);
    lines.push(
      `flags: 0x${(analysis.flags ?? 0).toString(16).padStart(8, "0")} (compressed=${Boolean((analysis.flags ?? 0) & 1)}, encrypted=${Boolean((analysis.flags ?? 0) & 2)})`,
    );
  }

  for (const [path, stream] of [...analysis.recordStreams].sort(([left], [right]) =>
    left.localeCompare(right),
  )) {
    lines.push("");
    lines.push(`## Records: ${path}`);
    lines.push(
      `raw-deflate(windowBits=-15): ${stream.compressedSize} -> ${stream.raw.length} bytes`,
    );
    lines.push(
      `record-count: ${stream.records.length}; consumed: ${stream.raw.length}; trailing: 0`,
    );
    for (const record of stream.records) lines.push(recordLine(path, record));
    if (path === "DocInfo") {
      lines.push("");
      lines.push("### ID_MAPPINGS (18 entries)");
      const result = formatIdMappings(stream.records);
      lines.push(...result.lines);
      analysis.errors.push(...result.errors.map((error) => `DocInfo: ${error}`));
    }
  }

  lines.push("");
  lines.push("## Validation");
  if (analysis.errors.length === 0) lines.push("OK");
  else for (const error of analysis.errors) lines.push(`ERROR: ${error}`);
  return lines.join("\n");
}

function recordKey(record: HwpRecord): string {
  return `${record.tag}:${record.level}:${record.size}:${sha256(record.data, 64)}`;
}

/** Hirschberg LCS: stable record alignment with linear memory. */
function lcsMatches(left: HwpRecord[], right: HwpRecord[]): Array<[number, number]> {
  const leftKeys = left.map(recordKey);
  const rightKeys = right.map(recordKey);

  const lengths = (
    aStart: number,
    aEnd: number,
    bStart: number,
    bEnd: number,
    reverse: boolean,
  ): Uint32Array => {
    const bLength = bEnd - bStart;
    let previous = new Uint32Array(bLength + 1);
    let current = new Uint32Array(bLength + 1);
    const aLength = aEnd - aStart;
    for (let ai = 0; ai < aLength; ai++) {
      current[0] = 0;
      const aIndex = reverse ? aEnd - 1 - ai : aStart + ai;
      for (let bi = 0; bi < bLength; bi++) {
        const bIndex = reverse ? bEnd - 1 - bi : bStart + bi;
        current[bi + 1] =
          leftKeys[aIndex] === rightKeys[bIndex]
            ? previous[bi] + 1
            : Math.max(current[bi], previous[bi + 1]);
      }
      [previous, current] = [current, previous];
    }
    return previous;
  };

  const matches: Array<[number, number]> = [];
  const solve = (aStart: number, aEnd: number, bStart: number, bEnd: number): void => {
    if (aStart >= aEnd || bStart >= bEnd) return;
    if (aEnd - aStart === 1) {
      for (let index = bStart; index < bEnd; index++) {
        if (leftKeys[aStart] === rightKeys[index]) {
          matches.push([aStart, index]);
          break;
        }
      }
      return;
    }
    const aMiddle = Math.floor((aStart + aEnd) / 2);
    const forward = lengths(aStart, aMiddle, bStart, bEnd, false);
    const backward = lengths(aMiddle, aEnd, bStart, bEnd, true);
    const bLength = bEnd - bStart;
    let split = 0;
    let best = -1;
    for (let index = 0; index <= bLength; index++) {
      const score = forward[index] + backward[bLength - index];
      if (score > best) {
        best = score;
        split = index;
      }
    }
    const bMiddle = bStart + split;
    solve(aStart, aMiddle, bStart, bMiddle);
    solve(aMiddle, aEnd, bMiddle, bEnd);
  };
  solve(0, left.length, 0, right.length);
  return matches.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
}

function diff(left: Analysis, right: Analysis): string {
  const lines: string[] = [];
  lines.push("# HWP diff");
  lines.push(`left: ${left.inputPath}`);
  lines.push(`right: ${right.inputPath}`);
  lines.push("");
  lines.push("## CFB stream diff");
  const leftStreams = new Map(left.streams.map((stream) => [stream.path, stream]));
  const rightStreams = new Map(right.streams.map((stream) => [stream.path, stream]));
  const streamPaths = [...new Set([...leftStreams.keys(), ...rightStreams.keys()])].sort();
  for (const path of streamPaths) {
    const leftStream = leftStreams.get(path);
    const rightStream = rightStreams.get(path);
    if (!leftStream)
      lines.push(`[ONLY RIGHT STREAM] ${path}: ${rightStream!.data.length} bytes`);
    else if (!rightStream)
      lines.push(`[ONLY LEFT STREAM]  ${path}: ${leftStream.data.length} bytes`);
    else
      lines.push(
        `[BOTH STREAMS]      ${path}: left=${leftStream.data.length} right=${rightStream.data.length}${sha256(leftStream.data) === sha256(rightStream.data) ? " identical" : " different"}`,
      );
  }

  const recordPaths = [
    ...new Set([...left.recordStreams.keys(), ...right.recordStreams.keys()]),
  ].sort();
  for (const path of recordPaths) {
    lines.push("");
    lines.push(`## Record diff: ${path}`);
    const leftRecords = left.recordStreams.get(path)?.records ?? [];
    const rightRecords = right.recordStreams.get(path)?.records ?? [];
    const matches = lcsMatches(leftRecords, rightRecords);
    let leftIndex = 0;
    let rightIndex = 0;
    for (const [matchedLeft, matchedRight] of [
      ...matches,
      [leftRecords.length, rightRecords.length] as [number, number],
    ]) {
      while (leftIndex < matchedLeft) {
        lines.push(`[ONLY LEFT]  ${recordLine(path, leftRecords[leftIndex]).trimStart()}`);
        leftIndex++;
      }
      while (rightIndex < matchedRight) {
        lines.push(`[ONLY RIGHT] ${recordLine(path, rightRecords[rightIndex]).trimStart()}`);
        rightIndex++;
      }
      if (matchedLeft < leftRecords.length && matchedRight < rightRecords.length) {
        lines.push(`[BOTH]       ${recordLine(path, leftRecords[matchedLeft]).trimStart()}`);
        leftIndex = matchedLeft + 1;
        rightIndex = matchedRight + 1;
      }
    }
  }
  return lines.join("\n");
}

function usage(): string {
  return [
    "Usage:",
    "  pnpm hwp:diff <file.hwp>",
    "  pnpm hwp:diff <left.hwp> <right.hwp>",
  ].join("\n");
}

function main(args: string[]): void {
  if (args.length < 1 || args.length > 2 || args.includes("--help")) {
    console.error(usage());
    process.exitCode = args.includes("--help") ? 0 : 2;
    return;
  }
  try {
    const left = analyze(args[0]);
    console.log(dump(left));
    if (args.length === 2) {
      const right = analyze(args[1]);
      console.log("\n");
      console.log(dump(right));
      console.log("\n");
      console.log(diff(left, right));
      if (right.errors.length > 0) process.exitCode = 1;
    }
    if (left.errors.length > 0) process.exitCode = 1;
  } catch (error) {
    console.error(`ERROR: ${error instanceof Error ? error.message : error}`);
    process.exitCode = 1;
  }
}

main(process.argv.slice(2));
