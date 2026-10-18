import pako from 'pako';
import { describe, expect, it } from 'vitest';
import { Pipeline, BinaryKit, buildRoot, buildSheet, buildPara, buildSpan, buildGrid, buildRow, buildCell, registry } from '../../index';
import { parseHwpRecords, verifyHwpRecordStreams } from './verify';

function records(bytes: Uint8Array) {
  const streams = BinaryKit.parseCfb(bytes);
  return parseHwpRecords(pako.inflateRaw(streams.get('BodyText/Section0')!));
}
const endBit = (data: Uint8Array) => new DataView(data.buffer, data.byteOffset).getUint32(0, true) >>> 31;

describe('HWP paragraph list termination (Hancom reader contract)', () => {
  it('a section-control-only paragraph must not terminate the body before a table', async () => {
    const result = await Pipeline.open('| 제목 |\n| --- |\n| 본문 |\n\n마지막').to('hwp');
    if (!result.ok) throw new Error(result.error);
    const headers = records(result.data).filter(r => r.tag === 66 && r.level === 0);
    expect(headers.map(r => endBit(r.data))).toEqual([0, 0, 1]);
  });

  it('terminates each cell independently, including multi-paragraph and nested cells', async () => {
    const p = (s: string) => buildPara([buildSpan(s)]);
    const inner = buildGrid([buildRow([buildCell([p('nested one'), p('nested last')])])]);
    const grid = buildGrid([buildRow([
      buildCell([p('one'), p('two'), inner, p('cell last')]),
      buildCell([p('other cell')]),
    ])]);
    const result = await registry.getEncoder('hwp')!.encode(buildRoot({}, [buildSheet([p('start'), grid, p('end')])]));
    if (!result.ok) throw new Error(result.error);
    const headers = records(result.data).filter(r => r.tag === 66);
    expect(headers.filter(r => r.level === 0).map(r => endBit(r.data))).toEqual([0, 0, 1]);
    expect(headers.filter(r => r.level === 2).map(r => endBit(r.data))).toEqual([0, 0, 0, 1, 1]);
    expect(headers.filter(r => r.level === 4).map(r => endBit(r.data))).toEqual([0, 1]);
  });

  it('the verifier rejects an early body terminator', async () => {
    const result = await Pipeline.open('첫 문단\n\n마지막 문단').to('hwp');
    if (!result.ok) throw new Error(result.error);
    const streams = BinaryKit.parseCfb(result.data);
    const body = pako.inflateRaw(streams.get('BodyText/Section0')!);
    const docInfo = pako.inflateRaw(streams.get('DocInfo')!);
    expect(verifyHwpRecordStreams(docInfo, body).ok).toBe(true);
    const first = parseHwpRecords(body).find(r => r.tag === 66)!;
    const view = new DataView(first.data.buffer, first.data.byteOffset);
    view.setUint32(0, view.getUint32(0, true) | 0x80000000, true);
    expect(verifyHwpRecordStreams(docInfo, body).errors.join(' ')).toContain('paragraph-list end bit');
  });
});
