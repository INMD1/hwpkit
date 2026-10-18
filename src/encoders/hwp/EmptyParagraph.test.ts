import pako from 'pako';
import { describe, expect, it } from 'vitest';
import {
  BinaryKit, Pipeline, registry, buildRoot, buildSheet, buildPara, buildSpan,
  buildGrid, buildRow, buildCell,
} from '../../index';
import { parseHwpRecords, verifyHwpRecordStreams } from './verify';

const u32 = (data: Uint8Array, offset = 0) =>
  new DataView(data.buffer, data.byteOffset, data.byteLength).getUint32(offset, true);

async function encodeEmptyParagraphs() {
  const text = (s: string) => buildPara([buildSpan(s)]);
  const empty = () => buildPara([buildSpan('', { pt: 24 })]);
  const nested = buildGrid([buildRow([buildCell([text('nested'), empty()])])]);
  const table = buildGrid([buildRow([
    buildCell([text('cell'), empty(), nested, empty()]),
    buildCell([]),
  ])]);
  const result = await registry.getEncoder('hwp')!.encode(buildRoot({}, [
    buildSheet([text('start'), empty(), table, empty(), text('end'), empty()]),
  ]));
  if (!result.ok) throw new Error(result.error);
  const streams = BinaryKit.parseCfb(result.data);
  return {
    bytes: result.data,
    docInfo: pako.inflateRaw(streams.get('DocInfo')!),
    body: pako.inflateRaw(streams.get('BodyText/Section0')!),
  };
}

describe('Hancom empty paragraph record contract', () => {
  it('omits PARA_TEXT for implicit CRs in body, cells, and nested cells', async () => {
    const { body, docInfo, bytes } = await encodeEmptyParagraphs();
    const records = parseHwpRecords(body);
    const emptyLevels: number[] = [];
    for (let i = 0; i < records.length; i++) {
      const header = records[i];
      if (header.tag !== 66 || (u32(header.data) & 0x7fffffff) !== 1) continue;
      emptyLevels.push(header.level);
      // libCoreEngine.so sub_6547a0 skips text when nChars <= 1, then
      // sub_64afa0 calls ReadRecord(68) unconditionally.
      expect(records[i + 1].tag).toBe(68);
      expect(records[i + 1].level).toBe(header.level + 1);
      expect(u32(records[i + 1].data)).toBe(0);
    }
    expect(emptyLevels).toEqual([0, 2, 4, 2, 2, 0, 0]);
    expect(verifyHwpRecordStreams(docInfo, body).errors).toEqual([]);

    const decoded = await Pipeline.open(bytes).inspect();
    if (!decoded.ok) throw new Error(decoded.error);
    const emptyFonts: number[] = [];
    const visit = (node: any) => {
      if (node.tag === 'span' && node.kids.every((k: any) => k.tag === 'txt' && k.content === ''))
        emptyFonts.push(node.props.pt);
      for (const kid of node.kids ?? []) visit(kid);
    };
    visit(decoded.data);
    expect(emptyFonts.filter(pt => pt === 24)).toHaveLength(6);
  });

  it('rejects a redundant text record after an empty paragraph header', async () => {
    const { body, docInfo } = await encodeEmptyParagraphs();
    const records = parseHwpRecords(body);
    const index = records.findIndex(r => r.tag === 66 && (u32(r.data) & 0x7fffffff) === 1);
    const insertAt = records[index + 1].offset;
    const extra = new Uint8Array(6);
    const view = new DataView(extra.buffer);
    view.setUint32(0, (2 << 20) | (1 << 10) | 67, true);
    view.setUint16(4, 13, true);
    const invalid = new Uint8Array(body.length + extra.length);
    invalid.set(body.subarray(0, insertAt));
    invalid.set(extra, insertAt);
    invalid.set(body.subarray(insertAt), insertAt + extra.length);
    expect(verifyHwpRecordStreams(docInfo, invalid).errors.join(' ')).toContain('must omit PARA_TEXT');
  });
});
