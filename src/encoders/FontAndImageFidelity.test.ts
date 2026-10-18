import JSZip from 'jszip';
import pako from 'pako';
import { describe, expect, it } from 'vitest';
import { BinaryKit, Pipeline, registry, buildRoot, buildSheet, buildPara, buildSpan, buildImg,
  buildGrid, buildRow, buildCell } from '../index';
import type { DocRoot, ParaNode, SpanNode } from '../model/doc-tree';
import { parseHwpRecords } from './hwp/verify';

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const u32 = (d: Uint8Array, o = 0) => new DataView(d.buffer, d.byteOffset, d.byteLength).getUint32(o, true);
const u16 = (d: Uint8Array, o = 0) => new DataView(d.buffer, d.byteOffset, d.byteLength).getUint16(o, true);
async function encode(doc: DocRoot, format: string) {
  const result = await registry.getEncoder(format)!.encode(doc);
  if (!result.ok) throw new Error(result.error);
  return result.data;
}
async function open(bytes: Uint8Array, format: string) {
  const result = await Pipeline.open(bytes, format).inspect();
  if (!result.ok) throw new Error(result.error);
  return result.data;
}
function records(bytes: Uint8Array, stream: string) {
  return parseHwpRecords(pako.inflateRaw(BinaryKit.parseCfb(bytes).get(stream)!));
}
const text = (p: ParaNode) => p.kids.map(k => k.tag === 'img' ? '[image]' : k.tag === 'span'
  ? k.kids.map(t => t.tag === 'txt' ? t.content : '').join('') : '').join('');

async function independentHwpx() {
  const zip = new JSZip();
  zip.file('Contents/header.xml', `<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"><hh:refList>
    <hh:fontfaces><hh:fontface lang="HANGUL"><hh:font id="0" face="HY견고딕"/><hh:font id="1" face="바탕"/></hh:fontface>
    <hh:fontface lang="LATIN"><hh:font id="0" face="Arial"/><hh:font id="1" face="Courier New"/></hh:fontface></hh:fontfaces>
    <hh:charProperties><hh:charPr id="0" height="1200"><hh:fontRef hangul="1" latin="0"/><hh:bold/></hh:charPr></hh:charProperties>
  </hh:refList></hh:head>`);
  zip.file('Contents/section0.xml', `<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"><hp:p><hp:run charPrIDRef="0"><hp:t>한글 ABC</hp:t></hp:run></hp:p></hs:sec>`);
  return open(await zip.generateAsync({ type: 'uint8array' }), 'hwpx');
}

describe('font faces, style inheritance and picture anchors', () => {
  it('resolves independent HWPX language banks and retains them through all Office encoders', async () => {
    let doc = await independentHwpx();
    for (const format of ['hwp', 'docx', 'hwpx', 'hwp', 'docx', 'hwpx']) {
      doc = await open(await encode(doc, format), format);
      const span = (doc.kids[0].kids.find(n => n.tag === 'para') as ParaNode).kids[0] as SpanNode;
      expect(span.props.fontFaces).toMatchObject({ hangul: '바탕', latin: 'Arial' });
      expect(span.props.b).toBe(true);
      expect(span.props.pt).toBe(12);
    }
  });

  for (const format of ['hwp', 'hwpx']) {
    it(`${format} keeps the requested face in Latin and Korean banks without forced substitution`, async () => {
      const bytes = await encode(buildRoot({}, [buildSheet([
        buildPara([buildSpan('한글 ABC', { font: 'HY견고딕', b: true, i: true })]),
      ])]), format);
      if (format === 'hwp') {
        const info = records(bytes, 'DocInfo');
        const mapping = info.find(r => r.tag === 17)!.data;
        const names = info.filter(r => r.tag === 19).map(r => new TextDecoder('utf-16le').decode(r.data.subarray(3, 3 + u16(r.data, 1) * 2)));
        const shape = info.filter(r => r.tag === 21).find(r => (u32(r.data, 46) & 3) === 3)!.data;
        let base = 0;
        for (let lang = 0; lang < 7; lang++) {
          expect(names[base + u16(shape, lang * 2)]).toBe('HY견고딕');
          base += u32(mapping, 4 + lang * 4);
        }
      }
      const doc = await open(bytes, format);
      const p = doc.kids[0].kids.find(n => n.tag === 'para') as ParaNode;
      expect((p.kids[0] as SpanNode).props).toMatchObject({ b: true, i: true, fontFaces: { hangul: 'HY견고딕', latin: 'HY견고딕' } });
    });
  }

  it('inherits bold/italic through a size-only style and honors every direct OFF spelling', async () => {
    const zip = new JSZip();
    zip.file('word/styles.xml', `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="바탕"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Base"><w:rPr><w:b/><w:i/></w:rPr></w:style>
      <w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/><w:rPr><w:sz w:val="28"/></w:rPr></w:style></w:styles>`);
    zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr>
      <w:r><w:t>inherited</w:t></w:r>${['0', 'false', 'off'].map(val => `<w:r><w:rPr><w:b w:val="${val}"/><w:i w:val="${val}"/></w:rPr><w:t>${val}</w:t></w:r>`).join('')}</w:p></w:body></w:document>`);
    let doc = await open(await zip.generateAsync({ type: 'uint8array' }), 'docx');
    for (let cycle = 0; cycle < 2; cycle++) {
      const spans = (doc.kids[0].kids[0] as ParaNode).kids as SpanNode[];
      expect(spans[0].props).toMatchObject({ b: true, i: true, pt: 14, fontFaces: { hangul: '바탕', latin: 'Arial' } });
      for (const span of spans.slice(1)) expect(span.props).toMatchObject({ b: false, i: false });
      doc = await open(await encode(doc, 'docx'), 'docx');
    }
  });

  it('does not rewrite symbol fonts, boldness or insert spaces in HWP source text', async () => {
    const input = '◦marker →word □title';
    const bytes = await encode(buildRoot({}, [buildSheet([buildPara([
      buildSpan(input, { font: 'Arial', b: true }),
    ])])]), 'hwp');
    const doc = await open(bytes, 'hwp');
    const para = doc.kids[0].kids[0] as ParaNode;
    expect(text(para)).toBe(input);
    for (const span of para.kids as SpanNode[]) expect(span.props).toMatchObject({ b: true, font: 'Arial' });
  });

  it('keeps distinct HWP character styles when only shading differs', async () => {
    const bytes = await encode(buildRoot({}, [buildSheet([buildPara([
      buildSpan('red', { bg: 'FF0000' }), buildSpan('blue', { bg: '0000FF' }),
    ])])]), 'hwp');
    const shapes = records(bytes, 'DocInfo').filter(r => r.tag === 21);
    const shadeColors = shapes.map(r => Array.from(r.data.subarray(60, 63)));
    expect(shadeColors).toContainEqual([255, 0, 0]);
    expect(shadeColors).toContainEqual([0, 0, 255]);
  });

  for (const [kind, expectedBit] of [['sup', 15], ['sub', 16]] as const) {
    it(`writes and reads HWP ${kind} using the independent table 35 bit position`, async () => {
      const bytes = await encode(buildRoot({}, [buildSheet([buildPara([
        buildSpan('index', { [kind]: true }),
      ])])]), 'hwp');
      const attributes = records(bytes, 'DocInfo').filter(r => r.tag === 21).map(r => u32(r.data, 46));
      expect(attributes).toContain(1 << expectedBit);
      const doc = await open(bytes, 'hwp');
      const span = (doc.kids[0].kids[0] as ParaNode).kids[0] as SpanNode;
      expect(span.props[kind]).toBe(true);
      expect(span.props[kind === 'sup' ? 'sub' : 'sup']).toBeUndefined();
    });
  }

  for (const inCell of [false, true]) {
    it(`preserves text/image/text/image order in ${inCell ? 'a cell' : 'the body'} without extra paragraphs`, async () => {
      const para = buildPara([buildSpan('before', { b: true }), buildImg(PNG, 'image/png', 24, 24),
        buildSpan('between', { i: true }), buildImg(PNG, 'image/png', 16, 16), buildSpan('after')], { align: 'center' });
      const doc = buildRoot({}, [buildSheet(inCell ? [buildGrid([buildRow([buildCell([para])])])] : [para])]);
      const original = JSON.stringify(doc);
      const bytes = await encode(doc, 'hwp');
      const body = records(bytes, 'BodyText/Section0');
      const header = body.find(r => r.tag === 66 && (u32(r.data, 4) & (1 << 11)))!;
      expect(header).toBeDefined();
      expect(body.filter(r => r.tag === 66)).toHaveLength(inCell ? 3 : 1);
      if (inCell) expect(u16(body.find(r => r.tag === 72)!.data)).toBe(1);
      const decoded = await open(bytes, 'hwp');
      const grid = decoded.kids[0].kids.find(n => n.tag === 'grid');
      const p = inCell ? grid!.kids[0].kids[0].kids[0] as ParaNode : decoded.kids[0].kids[0] as ParaNode;
      expect(text(p)).toBe('before[image]between[image]after');
      expect(p.props.align).toBe('center');
      expect((p.kids[0] as SpanNode).props.b).toBe(true);
      expect((p.kids[2] as SpanNode).props.i).toBe(true);
      expect(JSON.stringify(doc)).toBe(original);
    });
  }

  for (const format of ['hwp', 'hwpx']) for (const rel of ['page', 'margin'] as const) {
    it(`${format} distinguishes ${rel} reference coordinates, including zero and negative offsets`, async () => {
      const doc = buildRoot({}, [buildSheet([buildPara([buildSpan('anchor'), buildImg(PNG, 'image/png', 24, 24, '', {
        wrap: 'front', horzRelTo: rel, vertRelTo: rel, horzAlign: 'left', vertAlign: 'top', xPt: 0, yPt: -12,
      })])])]);
      const bytes = await encode(doc, format);
      if (format === 'hwp') {
        const flags = u32(records(bytes, 'BodyText/Section0').find(r => r.tag === 71 && u32(r.data) === 0x67736f20)!.data, 4);
        expect((flags >>> 3) & 3).toBe(rel === 'page' ? 0 : 1);
        expect((flags >>> 8) & 3).toBe(rel === 'page' ? 0 : 1);
      }
      const decoded = await open(bytes, format);
      const img = (decoded.kids[0].kids.find(n => n.tag === 'para') as ParaNode).kids.find(n => n.tag === 'img');
      expect(img?.layout).toMatchObject({ horzRelTo: rel, vertRelTo: rel, xPt: 0, yPt: -12 });
    });
  }
});
