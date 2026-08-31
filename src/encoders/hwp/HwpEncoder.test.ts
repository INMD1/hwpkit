import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import pako from "pako";
import { beforeAll, describe, expect, it } from "vitest";
import type { DocRoot } from "../../model/doc-tree";
import { buildPara, buildRoot, buildSheet, buildSpan } from "../../model/builders";
import { HwpScanner } from "../../decoders/hwp/HwpScanner";
import { HwpxDecoder } from "../../decoders/hwpx/HwpxDecoder";
import { HwpxEncoder } from "../hwpx/HwpxEncoder";
import { ArchiveKit } from "../../toolkit/ArchiveKit";
import { BinaryKit } from "../../toolkit/BinaryKit";
import { HwpEncoder } from "./HwpEncoder";

const HWPTAG_BEGIN = 0x10;
const TAG_PAGE_DEF = HWPTAG_BEGIN + 57;

interface ParsedRecord {
  tag: number;
  level: number;
  data: Uint8Array;
}

function readU32(data: Uint8Array, offset: number): number {
  return new DataView(data.buffer, data.byteOffset, data.byteLength).getUint32(
    offset,
    true,
  );
}

/** Strict parser: it throws unless the final record consumes the final byte. */
function parseRecords(data: Uint8Array): ParsedRecord[] {
  const records: ParsedRecord[] = [];
  let offset = 0;
  while (offset < data.length) {
    if (data.length - offset < 4) {
      throw new Error(`${data.length - offset} trailing record byte(s)`);
    }
    const header = readU32(data, offset);
    const tag = header & 0x3ff;
    const level = (header >>> 10) & 0x3ff;
    let size = (header >>> 20) & 0xfff;
    offset += 4;
    if (size === 0xfff) {
      if (data.length - offset < 4) throw new Error("missing extended record size");
      size = readU32(data, offset);
      offset += 4;
    }
    if (size > data.length - offset) {
      throw new Error(`record size ${size} exceeds ${data.length - offset} remaining bytes`);
    }
    records.push({ tag, level, data: data.subarray(offset, offset + size) });
    offset += size;
  }
  expect(offset).toBe(data.length);
  return records;
}

function countTag(records: ParsedRecord[], tag: number): number {
  return records.filter((record) => record.tag === tag).length;
}

function extractPageDef(hwp: Uint8Array): Uint8Array {
  const streams = BinaryKit.parseCfb(hwp);
  const fileHeader = streams.get("FileHeader");
  const section0 = streams.get("BodyText/Section0");
  if (!fileHeader || !section0) throw new Error("required HWP stream is missing");
  const compressed = (readU32(fileHeader, 36) & 1) !== 0;
  const raw = compressed ? new Uint8Array(pako.inflateRaw(section0)) : section0;
  const pageDefs = parseRecords(raw).filter((record) => record.tag === TAG_PAGE_DEF);
  expect(pageDefs).toHaveLength(1);
  expect(pageDefs[0].data).toHaveLength(40);
  return new Uint8Array(pageDefs[0].data);
}

describe("HwpEncoder structural output", () => {
  let output: Uint8Array;
  let streams: Map<string, Uint8Array>;

  beforeAll(async () => {
    const document = buildRoot(
      { title: "HwpEncoder baseline" },
      [
        buildSheet([
          buildPara(
            [buildSpan("reference 없이 검증하는 HWP", { font: "함초롬바탕", pt: 10 })],
            { align: "left", spaceBefore: 0, spaceAfter: 0, lineHeight: 1.6 },
          ),
        ]),
      ],
    );
    const encoded = await new HwpEncoder().encode(document);
    if (!encoded.ok) throw new Error(encoded.error);
    output = encoded.data;
    streams = BinaryKit.parseCfb(output);
  });

  it("is a parseable CFB file with the required HWP streams", () => {
    expect(BinaryKit.isOle2(output)).toBe(true);
    expect(streams.get("FileHeader")).toBeDefined();
    expect(streams.get("DocInfo")).toBeDefined();
    expect(streams.get("BodyText/Section0")).toBeDefined();
  });

  it("writes the 256-byte FileHeader signature, version, and compression flag", () => {
    const header = streams.get("FileHeader")!;
    expect(header).toHaveLength(256);
    expect(new TextDecoder("ascii").decode(header.subarray(0, 17))).toBe(
      "HWP Document File",
    );
    expect(readU32(header, 32)).toBe(0x05010001);
    expect(readU32(header, 36) & 1).toBe(1);
  });

  it("raw-inflates DocInfo and Section0 and parses every record byte", () => {
    const docInfo = new Uint8Array(pako.inflateRaw(streams.get("DocInfo")!));
    const section0 = new Uint8Array(
      pako.inflateRaw(streams.get("BodyText/Section0")!),
    );
    expect(parseRecords(docInfo).length).toBeGreaterThan(0);
    expect(parseRecords(section0).length).toBeGreaterThan(0);
  });

  it("matches all ID_MAPPINGS counts to the corresponding DocInfo records", () => {
    const raw = new Uint8Array(pako.inflateRaw(streams.get("DocInfo")!));
    const records = parseRecords(raw);
    const mappings = records.filter(
      (record) => record.tag === HWPTAG_BEGIN + 1,
    );
    expect(mappings).toHaveLength(1);
    expect(mappings[0].data.length).toBeGreaterThanOrEqual(18 * 4);

    const values = Array.from({ length: 18 }, (_, index) =>
      new DataView(
        mappings[0].data.buffer,
        mappings[0].data.byteOffset,
        mappings[0].data.byteLength,
      ).getInt32(index * 4, true),
    );
    expect(values.every((value) => value >= 0)).toBe(true);

    // All seven language groups share HWPTAG_FACE_NAME, so their sum is the
    // independently verifiable count in the record stream.
    expect(values.slice(1, 8).reduce((sum, value) => sum + value, 0)).toBe(
      countTag(records, HWPTAG_BEGIN + 3),
    );

    const mappingIndexToTag = new Map<number, number>([
      [0, HWPTAG_BEGIN + 2],
      [8, HWPTAG_BEGIN + 4],
      [9, HWPTAG_BEGIN + 5],
      [10, HWPTAG_BEGIN + 6],
      [11, HWPTAG_BEGIN + 7],
      [12, HWPTAG_BEGIN + 8],
      [13, HWPTAG_BEGIN + 9],
      [14, HWPTAG_BEGIN + 10],
      [15, HWPTAG_BEGIN + 76],
      [16, HWPTAG_BEGIN + 80],
      [17, HWPTAG_BEGIN + 81],
    ]);
    for (const [index, tag] of mappingIndexToTag) {
      expect(values[index], `ID_MAPPINGS[${index}] vs tag 0x${tag.toString(16)}`).toBe(
        countTag(records, tag),
      );
    }
  });
});

describe("PAGE_DEF margin regression", () => {
  let reference: Uint8Array;
  let referencePageDef: Uint8Array;
  let referenceDoc: DocRoot;
  let encodedHwpx: Uint8Array;

  beforeAll(async () => {
    reference = new Uint8Array(
      readFileSync(resolve(process.cwd(), "tests/fixtures/reference.hwp")),
    );
    referencePageDef = extractPageDef(reference);

    const decoded = await new HwpScanner().decode(reference);
    if (!decoded.ok) throw new Error(decoded.error);
    referenceDoc = decoded.data;

    const hwpx = await new HwpxEncoder().encode(referenceDoc);
    if (!hwpx.ok) throw new Error(hwpx.error);
    encodedHwpx = hwpx.data;
  });

  it("preserves the exact 40-byte PAGE_DEF through HWP to HWP", async () => {
    const encoded = await new HwpEncoder().encode(referenceDoc);
    if (!encoded.ok) throw new Error(encoded.error);
    expect(extractPageDef(encoded.data)).toEqual(referencePageDef);
  });

  it("writes HWPX hp:margin with independent page and header/footer margins", async () => {
    const files = await ArchiveKit.unzip(encodedHwpx);
    const section0 = files.get("Contents/section0.xml");
    if (!section0) throw new Error("Contents/section0.xml is missing");
    const xml = new TextDecoder().decode(section0);
    expect(xml).toContain(
      '<hp:margin header="4252" footer="4252" gutter="0" left="8504" right="8504" top="5668" bottom="4252"/>',
    );
  });

  it("preserves the exact PAGE_DEF through HWP to HWPX to HWP", async () => {
    const decoded = await new HwpxDecoder().decode(encodedHwpx);
    if (!decoded.ok) throw new Error(decoded.error);
    const encoded = await new HwpEncoder().encode(decoded.data);
    if (!encoded.ok) throw new Error(encoded.error);
    expect(extractPageDef(encoded.data)).toEqual(referencePageDef);
  });
});
