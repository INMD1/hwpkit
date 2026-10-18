import pako from 'pako';
import { describe, expect, it } from 'vitest';
import {
  BinaryKit, registry, buildRoot, buildSheet, buildPara, buildSpan,
  buildGrid, buildRow, buildCell,
} from '../../index';
import { parseHwpRecords, verifyHwpRecordStreams } from './verify';
import type { GridNode } from '../../model/doc-tree';

const cell = (text: string, cs = 1, rs = 1) => buildCell([buildPara([buildSpan(text)])], { cs, rs });
const u16 = (data: Uint8Array, offset: number) =>
  new DataView(data.buffer, data.byteOffset, data.byteLength).getUint16(offset, true);

async function encode(grid: GridNode) {
  const before = structuredClone(grid);
  const result = await registry.getEncoder('hwp')!.encode(buildRoot({}, [buildSheet([grid])]));
  if (!result.ok) throw new Error(result.error);
  expect(grid).toEqual(before);
  const streams = BinaryKit.parseCfb(result.data);
  return {
    docInfo: pako.inflateRaw(streams.get('DocInfo')!),
    body: pako.inflateRaw(streams.get('BodyText/Section0')!),
  };
}

describe('Hancom rectangular table contract', () => {
  it('fills the uncovered column after a short merged row', async () => {
    // The demo.docx calendar has the same shape: row 0 spans 13 of 14 columns.
    const { body, docInfo } = await encode(buildGrid([
      buildRow([cell('merged', 2)]),
      buildRow([cell('A'), cell('B'), cell('C')]),
    ]));
    const records = parseHwpRecords(body);
    const cells = records.filter(r => r.tag === 72 && r.level === 2);
    expect(cells.map(r => [u16(r.data, 10), u16(r.data, 8), u16(r.data, 12)])).toEqual([
      [0, 0, 2], [0, 2, 1], [1, 0, 1], [1, 1, 1], [1, 2, 1],
    ]);
    expect(verifyHwpRecordStreams(docInfo, body).errors).toEqual([]);

    // Move the filler onto the merged cell: a gap and an overlap must fail.
    new DataView(cells[1].data.buffer, cells[1].data.byteOffset).setUint16(8, 0, true);
    expect(verifyHwpRecordStreams(docInfo, body).errors.join(' ')).toMatch(/gap or overlap|uncovered/);
  });

  it('places cells after columns occupied by a row span', async () => {
    const { body, docInfo } = await encode(buildGrid([
      buildRow([cell('A', 1, 2), cell('B')]),
      buildRow([cell('C')]),
    ]));
    const cells = parseHwpRecords(body).filter(r => r.tag === 72 && r.level === 2);
    expect(cells.map(r => [u16(r.data, 10), u16(r.data, 8), u16(r.data, 14)])).toEqual([
      [0, 0, 2], [0, 1, 1], [1, 1, 1],
    ]);
    expect(verifyHwpRecordStreams(docInfo, body).errors).toEqual([]);
  });

  it('records zero cell starts in a row entirely covered by a merge', async () => {
    const { body, docInfo } = await encode(buildGrid([
      buildRow([cell('merged', 2, 2)]), buildRow([]),
    ]));
    const table = parseHwpRecords(body).find(r => r.tag === 77)!;
    expect([u16(table.data, 18), u16(table.data, 20)]).toEqual([1, 0]);
    expect(verifyHwpRecordStreams(docInfo, body).errors).toEqual([]);
  });
});
