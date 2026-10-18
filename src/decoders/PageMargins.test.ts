import JSZip from 'jszip';
import pako from 'pako';
import { describe, expect, it } from 'vitest';
import { BinaryKit, Pipeline } from '../index';

async function fixture(top: number, header: number, bottom: number, footer: number) {
  const zip = new JSZip();
  zip.file('Contents/header.xml', '<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"/>');
  zip.file('Contents/section0.xml', `<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"><hp:p><hp:run><hp:secPr><hp:pagePr width="59528" height="84188"><hp:margin top="${top}" header="${header}" bottom="${bottom}" footer="${footer}" left="5669" right="5669"/></hp:pagePr></hp:secPr><hp:t>Position anchor</hp:t></hp:run></hp:p></hs:sec>`);
  return zip.generateAsync({ type: 'uint8array' });
}

function rawHwpMargins(bytes: Uint8Array) {
  const stream = pako.inflateRaw(BinaryKit.parseCfb(bytes).get('BodyText/Section0')!);
  const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  for (let offset = 0; offset < stream.length;) {
    const header = view.getUint32(offset, true); offset += 4;
    let size = header >>> 20;
    if (size === 4095) { size = view.getUint32(offset, true); offset += 4; }
    if ((header & 1023) === 73) return [16, 24, 20, 28].map(i => view.getUint32(offset + i, true));
    offset += size;
  }
  throw new Error('Missing PAGE_DEF');
}

describe('Physical body and header distances', () => {
  for (const values of [[4252, 2835, 4252, 2835], [0, 2000, 1000, 0], [3000, 0, 4000, 0]]) {
    it(`preserves independent Hancom margins ${values} through DOCX and 3 HWP/HWPX cycles`, async () => {
      const [top, header, bottom, footer] = values;
      let bytes = await fixture(top, header, bottom, footer);
      let format = 'hwpx';
      for (const target of ['docx', 'hwp', 'hwpx', 'docx', 'hwp', 'hwpx', 'docx', 'hwp', 'hwpx']) {
        const decoded = await Pipeline.open(bytes, format).inspect();
        if (!decoded.ok) throw new Error(decoded.error);
        const dims = decoded.data.kids[0].dims;
        expect(dims.mt).toBeCloseTo((top + header) / 100, 1);
        expect(dims.mb).toBeCloseTo((bottom + footer) / 100, 1);
        expect(dims.headerPt).toBeCloseTo(top / 100, 1);
        expect(dims.footerPt).toBeCloseTo(bottom / 100, 1);
        const converted = await Pipeline.open(bytes, format).to(target);
        if (!converted.ok) throw new Error(converted.error);
        bytes = converted.data; format = target;
        if (target === 'docx') {
          const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
          expect(xml).toContain(`w:top="${Math.round((top + header) / 5)}"`);
          expect(xml).toContain(`w:header="${Math.round(top / 5)}"`);
        } else if (target === 'hwp') {
          rawHwpMargins(bytes).forEach((v, i) => expect(Math.abs(v - values[i])).toBeLessThanOrEqual(3));
        }
      }
    });
  }
});
