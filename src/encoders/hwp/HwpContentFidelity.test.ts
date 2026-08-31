import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import pako from "pako";
import type { DocRoot, GridNode, ParaNode } from "../../model/doc-tree";
import {
  buildCell,
  buildGrid,
  buildPara,
  buildRoot,
  buildRow,
  buildSheet,
  buildSpan,
} from "../../model/builders";
import { Pipeline } from "../../pipeline/Pipeline";
import { BinaryKit } from "../../toolkit/BinaryKit";
import { MdEncoder } from "../md/MdEncoder";

const HWPTAG_BEGIN = 0x10;
const TAG_ID_MAPPINGS = HWPTAG_BEGIN + 1;
const TAG_BORDER_FILL = HWPTAG_BEGIN + 4;
const TAG_NUMBERING = HWPTAG_BEGIN + 7;
const TAG_BULLET = HWPTAG_BEGIN + 8;
const TAG_PARA_SHAPE = HWPTAG_BEGIN + 9;
const TAG_STYLE = HWPTAG_BEGIN + 10;
const TAG_PARA_HEADER = HWPTAG_BEGIN + 50;
const TAG_LIST_HEADER = HWPTAG_BEGIN + 56;
const TAG_TABLE = HWPTAG_BEGIN + 61;

interface ParsedRecord {
  tag: number;
  level: number;
  data: Uint8Array;
}

function readU16(data: Uint8Array, offset: number): number {
  return new DataView(data.buffer, data.byteOffset, data.byteLength).getUint16(
    offset,
    true,
  );
}

function readU32(data: Uint8Array, offset: number): number {
  return new DataView(data.buffer, data.byteOffset, data.byteLength).getUint32(
    offset,
    true,
  );
}

function parseRecords(data: Uint8Array): ParsedRecord[] {
  const records: ParsedRecord[] = [];
  let offset = 0;
  while (offset < data.length) {
    if (offset + 4 > data.length) throw new Error("trailing record bytes");
    const header = readU32(data, offset);
    const tag = header & 0x3ff;
    const level = (header >>> 10) & 0x3ff;
    let size = (header >>> 20) & 0xfff;
    offset += 4;
    if (size === 0xfff) {
      if (offset + 4 > data.length) throw new Error("missing extended size");
      size = readU32(data, offset);
      offset += 4;
    }
    if (offset + size > data.length) throw new Error("record exceeds stream");
    records.push({ tag, level, data: data.subarray(offset, offset + size) });
    offset += size;
  }
  expect(offset).toBe(data.length);
  return records;
}

function hwpRecords(
  hwp: Uint8Array,
  path: "DocInfo" | "BodyText/Section0",
): ParsedRecord[] {
  const streams = BinaryKit.parseCfb(hwp);
  const fileHeader = streams.get("FileHeader");
  const stream = streams.get(path);
  if (!fileHeader || !stream) throw new Error(`missing HWP stream: ${path}`);
  const compressed = (readU32(fileHeader, 36) & 1) !== 0;
  return parseRecords(
    compressed ? new Uint8Array(pako.inflateRaw(stream)) : stream,
  );
}

function styleNames(data: Uint8Array): [string, string] {
  let offset = 0;
  const readName = (): string => {
    const length = readU16(data, offset);
    offset += 2;
    const end = offset + length * 2;
    const name = new TextDecoder("utf-16le").decode(data.subarray(offset, end));
    offset = end;
    return name;
  };
  return [readName(), readName()];
}

function styleParaShapeId(data: Uint8Array): number {
  let offset = 0;
  for (let field = 0; field < 2; field++) {
    const length = readU16(data, offset);
    offset += 2 + length * 2;
  }
  return readU16(data, offset + 4);
}

function countTag(records: ParsedRecord[], tag: number): number {
  return records.filter(record => record.tag === tag).length;
}

function expectIdMappingsMatch(records: ParsedRecord[]): void {
  const mapping = records.filter(record => record.tag === TAG_ID_MAPPINGS);
  expect(mapping).toHaveLength(1);
  expect(mapping[0].data.length).toBeGreaterThanOrEqual(18 * 4);
  const values = Array.from({ length: 18 }, (_, index) =>
    new DataView(
      mapping[0].data.buffer,
      mapping[0].data.byteOffset,
      mapping[0].data.byteLength,
    ).getInt32(index * 4, true),
  );
  expect(values.slice(1, 8).reduce((sum, value) => sum + value, 0)).toBe(
    countTag(records, HWPTAG_BEGIN + 3),
  );
  const indexToTag = new Map<number, number>([
    [0, HWPTAG_BEGIN + 2],
    [8, HWPTAG_BEGIN + 4],
    [9, HWPTAG_BEGIN + 5],
    [10, HWPTAG_BEGIN + 6],
    [11, TAG_NUMBERING],
    [12, TAG_BULLET],
    [13, TAG_PARA_SHAPE],
    [14, TAG_STYLE],
    [15, HWPTAG_BEGIN + 76],
    [16, HWPTAG_BEGIN + 80],
    [17, HWPTAG_BEGIN + 81],
  ]);
  for (const [index, tag] of indexToTag) {
    expect(values[index], `ID_MAPPINGS[${index}]`).toBe(countTag(records, tag));
  }
}

function fixture(name: string): string {
  return readFileSync(
    resolve(process.cwd(), "tests/fixtures", name),
    "utf8",
  );
}

async function inspectMarkdown(name: string): Promise<DocRoot> {
  const result = await Pipeline.open(fixture(name), "md").inspect();
  if (!result.ok) throw new Error(result.error);
  return result.data;
}

async function roundTripMarkdownThroughHwp(name: string): Promise<DocRoot> {
  const encoded = await Pipeline.open(fixture(name), "md").to("hwp");
  if (!encoded.ok) throw new Error(encoded.error);
  const decoded = await Pipeline.open(encoded.data, "hwp").inspect();
  if (!decoded.ok) throw new Error(decoded.error);
  return decoded.data;
}

async function encodeMarkdownFixtureToHwp(name: string): Promise<Uint8Array> {
  const encoded = await Pipeline.open(fixture(name), "md").to("hwp");
  if (!encoded.ok) throw new Error(encoded.error);
  return encoded.data;
}

function firstGrid(doc: DocRoot): GridNode {
  const grid = doc.kids[0]?.kids.find(
    (node): node is GridNode => node.tag === "grid",
  );
  if (!grid) throw new Error("grid is missing");
  return grid;
}

function expectGridSize(grid: GridNode, rows: number, columns: number): void {
  expect(grid.kids).toHaveLength(rows);
  for (const row of grid.kids) {
    expect(row.kids).toHaveLength(columns);
    expect(row.kids.reduce((sum, cell) => sum + cell.cs, 0)).toBe(columns);
  }
}

describe("HWP content fidelity", () => {
  it.each([
    ["table-2x2.md", 2, 2],
    ["table-3x4.md", 3, 4],
  ])("preserves %s table dimensions through HWP", async (name, rows, columns) => {
    expectGridSize(firstGrid(await inspectMarkdown(name)), rows, columns);
    const hwp = await encodeMarkdownFixtureToHwp(name);
    const decoded = await Pipeline.open(hwp, "hwp").inspect();
    if (!decoded.ok) throw new Error(decoded.error);
    expectGridSize(
      firstGrid(decoded.data),
      rows,
      columns,
    );
    const records = hwpRecords(hwp, "BodyText/Section0");
    const table = records.filter(record => record.tag === TAG_TABLE);
    expect(table).toHaveLength(1);
    expect(readU16(table[0].data, 4)).toBe(rows);
    expect(readU16(table[0].data, 6)).toBe(columns);
    expect(countTag(records, TAG_LIST_HEADER)).toBe(rows * columns);
  });

  it("emits regular HWP tables and inline formatting as plain Markdown", async () => {
    const hwp = await Pipeline.open(fixture("table-2x2.md"), "md").to("hwp");
    if (!hwp.ok) throw new Error(hwp.error);
    const markdown = await Pipeline.open(hwp.data, "hwp").to("md");
    if (!markdown.ok) throw new Error(markdown.error);
    const text = new TextDecoder().decode(markdown.data);
    expect(text).toContain("| **A** | **B** |");
    expect(text).toContain("| C | D |");
    expect(text).not.toContain("<span");
    expect(text).not.toContain("<table");
  });

  it("uses Markdown markers and warns when unsupported text styles are dropped", async () => {
    const doc = buildRoot({}, [
      buildSheet([
        buildPara([
          buildSpan("styled", {
            b: true,
            i: true,
            s: true,
            font: "함초롬바탕",
            pt: 10,
          }),
          buildSpan(" code", { font: "Courier New" }),
        ]),
      ]),
    ]);
    const encoded = await new MdEncoder().encode(doc);
    if (!encoded.ok) throw new Error(encoded.error);
    const text = new TextDecoder().decode(encoded.data);
    expect(text).toBe("~~***styled***~~` code`");
    expect(text).not.toContain("<span");
    expect(encoded.warns.some(warn => warn.includes("글꼴명"))).toBe(true);
    expect(encoded.warns.some(warn => warn.includes("글자 크기"))).toBe(true);
  });

  it("falls back to an HTML table only for unrepresentable cells and warns", async () => {
    const merged = buildGrid([
      buildRow([
        buildCell([buildPara([buildSpan("merged")])], { cs: 2 }),
      ]),
    ]);
    const encoded = await new MdEncoder().encode(
      buildRoot({}, [buildSheet([merged])]),
    );
    if (!encoded.ok) throw new Error(encoded.error);
    expect(new TextDecoder().decode(encoded.data)).toContain("<table");
    expect(encoded.warns.some(warn => warn.includes("HTML 표로 폴백"))).toBe(true);
  });

  it("preserves H1 through H3 as HWP outline styles and heading IR", async () => {
    const hwp = await encodeMarkdownFixtureToHwp("headings.md");
    const decoded = await Pipeline.open(hwp, "hwp").inspect();
    if (!decoded.ok) throw new Error(decoded.error);
    const headings = decoded.data.kids[0].kids
      .filter(node => node.tag === "para")
      .map(para => para.props.heading)
      .filter(level => level !== undefined);
    expect(headings).toEqual([1, 2, 3]);

    const docInfo = hwpRecords(hwp, "DocInfo");
    const styles = docInfo.filter(record => record.tag === TAG_STYLE);
    expect(styles.map(record => styleNames(record.data))).toEqual([
      ["바탕글", "Normal"],
      ["본문", "Body"],
      ["개요 1", "Outline 1"],
      ["개요 2", "Outline 2"],
      ["개요 3", "Outline 3"],
    ]);

    const paraShapes = docInfo.filter(record => record.tag === TAG_PARA_SHAPE);
    const paraHeaders = hwpRecords(hwp, "BodyText/Section0")
      .filter(record => record.tag === TAG_PARA_HEADER)
      .filter(record => record.data.length >= 11 && record.data[10] >= 2);
    expect(paraHeaders.map(record => record.data[10])).toEqual([2, 3, 4]);
    expect(styles.slice(2).map(record => styleParaShapeId(record.data))).toEqual(
      paraHeaders.map(record => readU16(record.data, 8)),
    );
    paraHeaders.forEach((record, index) => {
      const paraShapeId = readU16(record.data, 8);
      const attr = readU32(paraShapes[paraShapeId].data, 0);
      expect((attr >>> 23) & 0x3).toBe(1);
      expect((attr >>> 25) & 0x7).toBe(index);
    });

    const markdown = await Pipeline.open(hwp, "hwp").to("md");
    if (!markdown.ok) throw new Error(markdown.error);
    const headingLines = new TextDecoder().decode(markdown.data)
      .split(/\r?\n/)
      .filter(line => line.startsWith("#"));
    expect(headingLines.map(line => line.match(/^#+/)?.[0].length)).toEqual([
      1,
      2,
      3,
    ]);
  });

  it("preserves bullet and numbered list markers through HWP", async () => {
    const hwp = await encodeMarkdownFixtureToHwp("lists.md");
    const decoded = await Pipeline.open(hwp, "hwp").inspect();
    if (!decoded.ok) throw new Error(decoded.error);
    const lists = decoded.data.kids[0].kids
      .filter(
        (node): node is ParaNode =>
          node.tag === "para" && node.props.listOrd !== undefined,
      );
    expect(lists.map(para => para.props.listOrd)).toEqual([
      false,
      false,
      false,
      true,
      true,
      true,
    ]);

    const markdown = await Pipeline.open(hwp, "hwp").to("md");
    if (!markdown.ok) throw new Error(markdown.error);
    const lines = new TextDecoder().decode(markdown.data).split(/\r?\n/);
    expect(lines.filter(line => line.startsWith("- "))).toHaveLength(3);
    expect(lines.filter(line => line.startsWith("1. "))).toHaveLength(3);

    const records = hwpRecords(hwp, "DocInfo");
    expect(countTag(records, TAG_NUMBERING)).toBe(1);
    expect(countTag(records, TAG_BULLET)).toBe(1);
    expect(records.find(record => record.tag === TAG_NUMBERING)?.data).toHaveLength(230);
    expect(records.find(record => record.tag === TAG_BULLET)?.data).toHaveLength(20);
    expectIdMappingsMatch(records);

    const listShapes = records
      .filter(record => record.tag === TAG_PARA_SHAPE)
      .map(record => ({
        type: (readU32(record.data, 0) >>> 23) & 0x3,
        level: (readU32(record.data, 0) >>> 25) & 0x7,
        id: readU16(record.data, 30),
      }))
      .filter(shape => shape.type === 2 || shape.type === 3);
    expect(listShapes).toEqual([
      { type: 3, level: 0, id: 1 },
      { type: 2, level: 0, id: 1 },
    ]);
  });

  it("preserves three nested list levels through HWP", async () => {
    const decoded = await roundTripMarkdownThroughHwp("lists-nested.md");
    const lists = decoded.kids[0].kids
      .filter(
        (node): node is ParaNode =>
          node.tag === "para" && node.props.listOrd !== undefined,
      );
    expect(lists.map(para => [para.props.listOrd, para.props.listLv])).toEqual([
      [false, 0],
      [false, 1],
      [false, 2],
      [true, 0],
      [true, 1],
      [true, 2],
    ]);
  });

  it("keeps the combined document record streams structurally valid", async () => {
    const hwp = await encodeMarkdownFixtureToHwp("content-fidelity.md");
    const docInfo = hwpRecords(hwp, "DocInfo");
    const section0 = hwpRecords(hwp, "BodyText/Section0");
    expect(docInfo.length).toBeGreaterThan(0);
    expect(section0.length).toBeGreaterThan(0);
    expectIdMappingsMatch(docInfo);

    const paraHeaders = section0.filter(record => record.tag === TAG_PARA_HEADER);
    expect(paraHeaders.length).toBeGreaterThan(0);
    for (const record of paraHeaders) {
      expect(readU32(record.data, 0) >>> 31).toBe(1);
    }

    const defaultBorder = docInfo.find(record => record.tag === TAG_BORDER_FILL);
    if (!defaultBorder) throw new Error("default BORDER_FILL is missing");
    for (let side = 0; side < 4; side++) {
      const offset = 2 + side * 6;
      expect(defaultBorder.data[offset]).toBe(0);
      expect(defaultBorder.data[offset + 1]).toBe(0);
    }

    const decoded = await Pipeline.open(hwp, "hwp").inspect();
    if (!decoded.ok) throw new Error(decoded.error);
    expect(firstGrid(decoded.data).kids).toHaveLength(2);
    expect(
      decoded.data.kids[0].kids.filter(
        node => node.tag === "para" && node.props.heading !== undefined,
      ),
    ).toHaveLength(2);
    expect(
      decoded.data.kids[0].kids.filter(
        node => node.tag === "para" && node.props.listOrd !== undefined,
      ),
    ).toHaveLength(6);
  });
});
