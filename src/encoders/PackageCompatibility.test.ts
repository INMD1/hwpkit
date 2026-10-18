import JSZip from 'jszip';
import { SaxesParser } from 'saxes';
import { describe, expect, it } from 'vitest';
import { Pipeline, TreeWalker, BinaryKit } from '../index';
import pako from 'pako';
import { parseHwpRecords } from './hwp/verify';

describe('external document reader contracts', () => {
  it('decodes Hancom mixed text/lineBreak/tab order and preserves it across binary formats', async () => {
    const zip = new JSZip();
    zip.file('Contents/section0.xml',
      '<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">' +
      '<hp:p><hp:run><hp:t>앞<hp:lineBreak/>중간<hp:tab width="4000" leader="0" type="1"/>뒤<hp:lineBreak/>끝</hp:t></hp:run></hp:p></hs:sec>');
    let bytes = await zip.generateAsync({ type: 'uint8array' });
    let from = 'hwpx';
    for (const target of ['docx', 'hwpx', 'hwp', 'docx', 'hwpx']) {
      const result = await Pipeline.open(bytes, from).to(target);
      if (!result.ok) throw new Error(result.error);
      const parsed = await Pipeline.open(result.data, target).inspect();
      if (!parsed.ok) throw new Error(parsed.error);
      const para = parsed.data.kids[0].kids[0];
      const text = para.tag === 'para' ? para.kids.flatMap(k => k.tag === 'span' ? k.kids : [])
        .map(k => k.tag === 'txt' ? k.content : k.tag === 'br' ? '\n' : '').join('') : '';
      expect(text).toBe('앞\n중간\t뒤\n끝');
      if (target === 'hwpx') {
        const xml = await (await JSZip.loadAsync(result.data)).file('Contents/section0.xml')!.async('string');
        expect(xml).not.toContain('<hp:br');
        expect(xml).toContain('<hp:lineBreak/>');
        const ancestors: string[] = [];
        const parser = new SaxesParser({ xmlns: true });
        parser.on('opentag', tag => {
          if (tag.local === 'lineBreak' || tag.local === 'tab') expect(ancestors.at(-1)).toBe('t');
          ancestors.push(tag.local);
        });
        parser.on('closetag', () => { ancestors.pop(); });
        parser.write(xml).close();
      }
      bytes = result.data;
      from = target;
    }
  });
  it.each(['hwpx', 'docx'])('%s XML is well-formed even with pasted control characters', async format => {
    const result = await Pipeline.open('# 한글 🐱\n\nbefore\0after\u000bnext\u0001end\ud800').to(format);
    if (!result.ok) throw new Error(result.error);
    const zip = await JSZip.loadAsync(result.data, { checkCRC32: true });
    for (const entry of Object.values(zip.files)) {
      if (!/\.(xml|hpf|rels)$/.test(entry.name)) continue;
      const xml = await entry.async('string');
      expect(() => new SaxesParser({ xmlns: true }).write(xml).close(), entry.name).not.toThrow();
    }
    const parsed = await Pipeline.open(result.data).inspect();
    if (!parsed.ok) throw new Error(parsed.error);
    expect(new TreeWalker().extractText(parsed.data)).toContain('한글 🐱');
    expect(new TreeWalker().extractText(parsed.data)).toContain('beforeafternextend');
  });

  it('uses the official Hancom margin element, including switch fallback', async () => {
    const result = await Pipeline.open('> 들여쓴 본문').to('hwpx');
    if (!result.ok) throw new Error(result.error);
    const zip = await JSZip.loadAsync(result.data, { checkCRC32: true });
    const xml = await zip.file('Contents/header.xml')!.async('string');
    // hancom-io/hwpx-owpml-model OWPML/Class/Head/margin.cpp maps "intent".
    expect(xml).not.toContain('<hc:indent');
    expect(xml.match(/<hh:margin>/g)?.length).toBe(xml.match(/<hc:intent /g)?.length);
    // mimetype is the first local ZIP entry, uncompressed, with no extra field.
    const view = new DataView(result.data.buffer, result.data.byteOffset);
    expect(new TextDecoder().decode(result.data.subarray(30, 38))).toBe('mimetype');
    expect(view.getUint16(8, true)).toBe(0);
    expect(view.getUint16(28, true)).toBe(0);
  });

  it('HWP tabs occupy eight words and do not swallow following text or style changes', async () => {
    const result = await Pipeline.open('시작\tABCDEFG **굵게**\t마지막').to('hwp');
    if (!result.ok) throw new Error(result.error);
    const streams = BinaryKit.parseCfb(result.data);
    const textRecord = parseHwpRecords(pako.inflateRaw(streams.get('BodyText/Section0')!)).find(r => r.tag === 67)!;
    const words = new Uint16Array(textRecord.data.buffer, textRecord.data.byteOffset, textRecord.data.length / 2);
    const tab = Array.from(words).indexOf(9);
    expect(Array.from(words.slice(tab, tab + 8))).toEqual([9, 0, 0, 0, 0, 0, 0, 9]);
    const parsed = await Pipeline.open(result.data).inspect();
    if (!parsed.ok) throw new Error(parsed.error);
    expect(new TreeWalker().extractText(parsed.data)).toContain('시작\tABCDEFG 굵게\t마지막');
    const converted = await Pipeline.open(result.data).to('docx');
    if (!converted.ok) throw new Error(converted.error);
    const xml = await (await JSZip.loadAsync(converted.data)).file('word/document.xml')!.async('string');
    expect(xml).toContain('ABCDEFG');
    expect(xml).toContain('<w:b/>');
    expect(xml).toContain('마지막');
  });
});
