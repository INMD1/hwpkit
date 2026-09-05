import JSZip from 'jszip';
import { HwpEncoder } from '../encoders/hwp/HwpEncoder';
import { HwpxEncoder } from '../encoders/hwpx/HwpxEncoder';
import pako from 'pako';
import { describe, expect, it, vi } from 'vitest';
import { Pipeline, BinaryKit, buildRoot, buildSheet, buildPara, buildSpan } from '../index';

function paraShapeLefts(hwp: Uint8Array): number[] {
  const streams = BinaryKit.parseCfb(hwp);
  const data = pako.inflateRaw(streams.get('DocInfo')!);
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const lefts: number[] = [];
  for (let offset = 0; offset < data.length;) {
    const header = view.getUint32(offset, true);
    offset += 4;
    let size = header >>> 20;
    if (size === 4095) { size = view.getUint32(offset, true); offset += 4; }
    if ((header & 1023) === 25) lefts.push(view.getInt32(offset + 4, true));
    offset += size;
  }
  return lefts;
}

describe('Hancom margin and model body margin boundary', () => {
  it('reads an independent HWPX hanging fixture into the body margin', async () => {
    const zip = new JSZip();
    zip.file('Contents/header.xml', `<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">
      <hh:refList><hh:paraProperties><hh:paraPr id="0"><hh:margin>
        <hc:left value="3660"/><hc:indent value="-4000"/>
      </hh:margin></hh:paraPr></hh:paraProperties></hh:refList></hh:head>`);
    zip.file('Contents/section0.xml', `<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
      <hp:p paraPrIDRef="0"><hp:run><hp:t>내어쓰기</hp:t></hp:run></hp:p></hs:sec>`);
    const decoded = await Pipeline.open(await zip.generateAsync({ type: 'uint8array' }), 'hwpx').inspect();
    if (!decoded.ok) throw new Error(decoded.error);
    const para = decoded.data.kids[0].kids.find(p => p.tag === 'para')!;
    expect(para.props.indentPt).toBe(38.3);
    expect(para.props.firstLineIndentPt).toBe(-20);
  });

  for (const [body, first, rawLeft] of [[38.3, -20, 3660], [0, -20, -4000], [10, 20, 2000], [-5, 0, -1000]]) {
    for (const format of ['hwp', 'hwpx']) {
      it(`${format} encodes body=${body}, first=${first} and preserves it across three cycles`, async () => {
        let document = buildRoot({}, [buildSheet([buildPara([buildSpan('위치 보존')], {
          indentPt: body, firstLineIndentPt: first,
        })])]);
        for (let cycle = 0; cycle < 3; cycle++) {
          const encoded = await (format === 'hwp' ? new HwpEncoder() : new HwpxEncoder()).encode(document);
          if (!encoded.ok) throw new Error(encoded.error);
          if (format === 'hwp') {
            expect(paraShapeLefts(encoded.data)).toContain(rawLeft);
          } else {
            const zip = await JSZip.loadAsync(encoded.data);
            const header = await zip.file('Contents/header.xml')!.async('string');
            expect(header).toContain(`<hc:left value="${rawLeft}"`);
          }
          const docx = await Pipeline.open(encoded.data, format).to('docx');
          if (!docx.ok) throw new Error(docx.error);
          const decoded = await Pipeline.open(docx.data, 'docx').inspect();
          if (!decoded.ok) throw new Error(decoded.error);
          document = decoded.data;
          const para = document.kids[0].kids.find(p => p.tag === 'para')!;
          expect(para.props.indentPt ?? 0).toBeCloseTo(body, 2);
          expect(para.props.firstLineIndentPt ?? 0).toBeCloseTo(first, 2);
        }
      });
    }
  }

  it('tries raw DEFLATE when wrapped inflate returns no result without throwing', async () => {
    const encoded = await Pipeline.open('압축 회귀 본문').to('hwp');
    if (!encoded.ok) throw new Error(encoded.error);
    const spy = vi.spyOn(pako, 'inflate').mockReturnValue(undefined as unknown as ReturnType<typeof pako.inflate>);
    try {
      const decoded = await Pipeline.open(encoded.data, 'hwp').to('md');
      if (!decoded.ok) throw new Error(decoded.error);
      expect(new TextDecoder().decode(decoded.data)).toContain('압축 회귀 본문');
    } finally {
      spy.mockRestore();
    }
  });
});
