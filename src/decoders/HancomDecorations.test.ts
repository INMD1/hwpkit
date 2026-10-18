import JSZip from 'jszip';
import pako from 'pako';
import { describe, expect, it, vi } from 'vitest';
import { BinaryKit, Pipeline, registry, XmlKit, buildRoot, buildSheet, buildPara, buildSpan } from '../index';
import { parseHwpRecords } from '../encoders/hwp/verify';
import type { SpanNode } from '../model/doc-tree';

async function nativeRecordFixture(attr: number) {
  const result = await registry.getEncoder('hwp')!.encode(buildRoot({}, [buildSheet([
    buildPara([buildSpan('decoration 한글', { pt: 12 })]),
  ])]));
  if (!result.ok) throw Error(result.error);
  const streams = BinaryKit.parseCfb(result.data);
  const raw = pako.inflateRaw(streams.get('DocInfo')!);
  // Independent CHAR_SHAPE attribute values. Mock only the CFB container read;
  // the real decoder still parses compression, records, spans and their IDs.
  for (const r of parseHwpRecords(raw)) if (r.tag === 21)
    new DataView(r.data.buffer, r.data.byteOffset, r.data.byteLength).setUint32(46, attr >>> 0, true);
  streams.set('DocInfo', pako.deflateRaw(raw));
  return { bytes: result.data, streams };
}
async function docxRunProperties(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  const root: any = await XmlKit.parseStrict(await zip.file('word/document.xml')!.async('string'));
  const runs: any[] = [];
  function visit(node: any) {
    if (!node || typeof node !== 'object') return;
    for (const [key, value] of Object.entries(node)) {
      if (key === 'w:r') runs.push(...value as any[]);
      else if (!key.startsWith('_') && Array.isArray(value)) for (const child of value) visit(child);
    }
  }
  visit(root);
  return runs.filter(r => r['w:t']?.some((t: any) => t._text?.trim())).map(r => r['w:rPr']?.[0] ?? {});
}
function isUnderline(p: any) { const n = p['w:u']?.[0]; return !!n && n._attr?.['w:val'] !== 'none'; }
function isStrikeout(p: any) { const n = p['w:strike']?.[0]; return !!n && !['0', 'false', 'off'].includes(n._attr?.['w:val']); }

const nativeCases: [string, number, boolean, boolean][] = [
  ['plain', 0, false, false],
  ['reported plain sentinel', 0x3c0400f8, false, false],
  ['reported bold sentinel', 0x3c0400fa, false, false],
  ['sentinel plus italic and kerning', 0x7c0400f9, false, false],
  ['bottom underline', 1 << 2, true, false],
  ['overline', 3 << 2, true, false],
  ['legacy line-through, solid', 2 << 2, false, true],
  ['legacy line-through, dotted', (2 << 2) | (2 << 4), false, true],
  ['legacy shape overrides a stale strikeout shape', (2 << 2) | (15 << 26), false, true],
  ['modern strikeout, solid', 1 << 18, false, true],
  ['modern strikeout, dotted', (1 << 18) | (2 << 26), false, true],
  ['modern strikeout disabled shape', (1 << 18) | (15 << 26), false, false],
  ['stored strikeout shape without enable bit', 2 << 26, false, false],
  ['underline and strikeout', (1 << 2) | (1 << 18), true, true],
];
describe('Hancom → DOCX visible text decorations', () => {
  for (const [label, attr, u, s] of nativeCases) it(`HWP ${label}`, async () => {
    const fixture = await nativeRecordFixture(attr);
    const spy = vi.spyOn(BinaryKit, 'parseCfb').mockReturnValue(fixture.streams);
    try {
      const decoded = await Pipeline.open(fixture.bytes, 'hwp').inspect();
      if (!decoded.ok) throw Error(decoded.error);
      const para = decoded.data.kids[0].kids.find(n => n.tag === 'para')!;
      const span = para.kids.find(n => n.tag === 'span') as SpanNode;
      expect(!!span.props.u).toBe(u); expect(!!span.props.s).toBe(s);
      expect(!!span.props.b).toBe(!!(attr & 2)); expect(!!span.props.i).toBe(!!(attr & 1));
      const result = await Pipeline.open(fixture.bytes, 'hwp').to('docx');
      if (!result.ok) throw Error(result.error);
      const prs = await docxRunProperties(result.data);
      expect(prs.length).toBeGreaterThan(0);
      for (const p of prs) { expect(isUnderline(p)).toBe(u); expect(isStrikeout(p)).toBe(s); }
    } finally { spy.mockRestore(); }
  });

  for (const inline of [false, true]) for (const val of ['NONE', 'none', '0', 'false', 'False', 'off']) {
    it(`HWPX ${inline ? 'inline CHARPR' : 'header charPr'} OFF=${val}`, async () => {
      const zip = new JSZip();
      zip.file('Contents/header.xml', `<hh:head xmlns:hh="h"><hh:refList><hh:charProperties><hh:charPr id="0" height="1200"><hh:underline type="${val}" shape="SOLID"/><hh:strikeout shape="${val}"/></hh:charPr></hh:charProperties></hh:refList></hh:head>`);
      zip.file('Contents/section0.xml', `<hs:sec xmlns:hs="s" xmlns:hp="p"><hp:p><hp:run ${inline ? '' : 'charPrIDRef="0"'}>${inline ? `<hp:CHARPR Underline="${val}" Strikeout="${val}"/>` : ''}<hp:t>plain</hp:t></hp:run></hp:p></hs:sec>`);
      const result = await Pipeline.open(await zip.generateAsync({ type: 'uint8array' }), 'hwpx').to('docx');
      if (!result.ok) throw Error(result.error);
      const prs = await docxRunProperties(result.data);
      expect(prs).toHaveLength(1); expect(isUnderline(prs[0])).toBe(false); expect(isStrikeout(prs[0])).toBe(false);
    });
  }

  it('HWPX preserves active underline/strikeout beside plain NONE/3D, including uppercase element names', async () => {
    const zip = new JSZip();
    zip.file('Contents/header.xml', `<hh:head xmlns:hh="h"><hh:refList><hh:charProperties>
      <hh:charPr id="0"><hh:underline type="NONE" shape="SOLID"/><hh:strikeout shape="3D"/></hh:charPr>
      <hh:charPr id="1"><hh:underline type="BOTTOM" shape="SOLID"/><hh:strikeout shape="NONE"/></hh:charPr>
      <hh:charPr id="2"><hh:underline type="NONE"/><hh:strikeout shape="SOLID"/></hh:charPr>
      <hh:charPr id="3"><hh:UNDERLINE Type="TOP"/><hh:STRIKEOUT Shape="DASH"/></hh:charPr>
      </hh:charProperties></hh:refList></hh:head>`);
    zip.file('Contents/section0.xml', `<hs:sec xmlns:hs="s" xmlns:hp="p"><hp:p>${[0, 1, 2, 3, 0].map(id => `<hp:run charPrIDRef="${id}"><hp:t>run${id}</hp:t></hp:run>`).join('')}</hp:p></hs:sec>`);
    const result = await Pipeline.open(await zip.generateAsync({ type: 'uint8array' }), 'hwpx').to('docx');
    if (!result.ok) throw Error(result.error);
    expect((await docxRunProperties(result.data)).map(p => [isUnderline(p), isStrikeout(p)]))
      .toEqual([[false, false], [true, false], [false, true], [true, true], [false, false]]);
  });

  it('DOCX explicitly writes underline/strikeout OFF between active runs', async () => {
    const result = await registry.getEncoder('docx')!.encode(buildRoot({}, [buildSheet([buildPara([
      buildSpan('active', { u: true, s: true }), buildSpan('off', { u: false, s: false }), buildSpan('plain'),
    ])])]));
    if (!result.ok) throw Error(result.error);
    const prs = await docxRunProperties(result.data);
    expect(prs.map(p => [isUnderline(p), isStrikeout(p)])).toEqual([[true, true], [false, false], [false, false]]);
    expect(prs[1]['w:u'][0]._attr['w:val']).toBe('none');
    expect(prs[1]['w:strike'][0]._attr['w:val']).toBe('0');
  });
});
