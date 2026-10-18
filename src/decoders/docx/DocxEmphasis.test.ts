import JSZip from 'jszip';
import pako from 'pako';
import { SaxesParser } from 'saxes';
import { describe, expect, it } from 'vitest';
import { BinaryKit, Pipeline } from '../../index';
import { parseHwpRecords } from '../../encoders/hwp/verify';

// Independent OOXML fixtures reproduce file-sample_1MB.docx's explicit OFF runs.
async function docx(value: string | undefined) {
  const zip = new JSZip();
  zip.file('word/styles.xml', `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="TextBody"><w:rPr><w:b/><w:i/></w:rPr></w:style></w:styles>`);
  const rpr = value === undefined ? '' : `<w:rPr><w:b w:val="${value}"/><w:i w:val="${value}"/><w:sz w:val="21"/></w:rPr>`;
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
    <w:p>${value === undefined ? '' : '<w:pPr><w:pStyle w:val="TextBody"/></w:pPr>'}<w:r>${rpr}<w:t>plain</w:t></w:r></w:p>
    <w:p><w:r><w:rPr><w:b/><w:i w:val="false"/></w:rPr><w:t>bold</w:t></w:r>
      <w:r><w:rPr><w:b w:val="false"/><w:i/></w:rPr><w:t>italic</w:t></w:r></w:p>
  </w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}

describe('DOCX emphasis does not appear on plain text in Hancom outputs', () => {
  for (const format of ['hwp', 'hwpx']) for (const off of [undefined, '0', 'false', 'off']) {
    it(`${format}: ${off ?? 'absent'} stays plain and genuine emphasis survives`, async () => {
      const output = await Pipeline.open(await docx(off), 'docx').to(format);
      if (!output.ok) throw Error(output.error);
      if (format === 'hwp') {
        const streams = BinaryKit.parseCfb(output.data);
        const info = parseHwpRecords(pako.inflateRaw(streams.get('DocInfo')!));
        const shapes = info.filter(r => r.tag === 21).map(r => new DataView(r.data.buffer, r.data.byteOffset, r.data.byteLength).getUint32(46, true) & 3);
        // Neither bold nor italic is set on the first text's raw CHAR_SHAPE.
        const body = parseHwpRecords(pako.inflateRaw(streams.get('BodyText/Section0')!));
        const pairs = body.find(r => r.tag === 68)!.data;
        const shapeId = new DataView(pairs.buffer, pairs.byteOffset, pairs.byteLength).getUint32(pairs.length - 4, true);
        expect(shapes[shapeId]).toBe(0);
        expect(shapes).toContain(2); // only bold, HWP bit 1
        expect(shapes).toContain(1); // only italic, HWP bit 0
        expect(shapes).not.toContain(3);
      } else {
        const zip = await JSZip.loadAsync(output.data);
        const header = await zip.file('Contents/header.xml')!.async('string');
        const section = await zip.file('Contents/section0.xml')!.async('string');
        let ref = '';
        let currentRef = '';
        const parser = new SaxesParser();
        parser.on('opentag', node => {
          if (node.name === 'hp:run') currentRef = String(node.attributes.charPrIDRef);
        });
        parser.on('text', text => { if (text === 'plain') ref = currentRef; });
        parser.write(section).close();
        expect(ref).not.toBe('');
        const shape = header.match(new RegExp(`<hh:charPr id="${ref}"[\\s\\S]*?<\\/hh:charPr>`))![0];
        expect(shape).not.toMatch(/<hh:(bold|italic)\b/);
      }
      const result = await Pipeline.open(output.data, format).inspect();
      if (!result.ok) throw Error(result.error);
      const chars: Record<string, { b: boolean; i: boolean }> = {};
      const walk = (node: any) => {
        if (node.tag === 'span') {
          const text = node.kids.filter((n: any) => n.tag === 'txt').map((n: any) => n.content).join('');
          if (text) chars[text] = { b: !!node.props.b, i: !!node.props.i };
        }
        for (const kid of node.kids ?? []) walk(kid);
      };
      walk(result.data);
      expect(chars.plain).toEqual({ b: false, i: false });
      expect(chars.bold).toEqual({ b: true, i: false });
      expect(chars.italic).toEqual({ b: false, i: true });
    });
  }
});
