import JSZip from 'jszip';
import pako from 'pako';
import { describe, expect, it } from 'vitest';
import { BinaryKit, Pipeline, registry, buildRoot, buildSheet, buildPara, buildSpan, buildImg } from '../index';
import type { DocRoot, ParaNode, ImgNode, SpanNode } from '../model/doc-tree';
import { parseHwpRecords } from './hwp/verify';
import { renderDocxChart } from '../toolkit/ChartRaster';

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const cache = (values: (string | number)[], kind: 'str' | 'num') => `<c:${kind}Lit><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:${kind}Lit>`;
const CHART = `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:tx><c:v>Column 1</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="004586"/></a:solidFill></c:spPr><c:cat>${cache(['Row 1', 'Row 2'], 'str')}</c:cat><c:val>${cache([9.1, 2.4], 'num')}</c:val></c:ser></c:barChart><c:valAx><c:scaling><c:min val="0"/><c:max val="10"/></c:scaling></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/></c:legend></c:chart></c:chartSpace>`;
async function open(bytes: Uint8Array): Promise<DocRoot> {
  const r = await Pipeline.open(bytes, 'docx').inspect();
  if (!r.ok) throw Error(r.error);
  return r.data;
}
async function encode(doc: DocRoot, format: string) {
  const r = await registry.getEncoder(format)!.encode(doc);
  if (!r.ok) throw Error(r.error);
  return r.data;
}
const records = (bytes: Uint8Array) => parseHwpRecords(pako.inflateRaw(BinaryKit.parseCfb(bytes).get('BodyText/Section0')!));
const i32 = (bytes: Uint8Array, offset: number) => new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getInt32(offset, true);
async function section(bytes: Uint8Array) { return (await JSZip.loadAsync(bytes)).file('Contents/section0.xml')!.async('string'); }
function lines(xml: string) {
  return [...xml.matchAll(/<hp:lineseg\s([^>]+)\/>/g)].map(m => Object.fromEntries([...m[1].matchAll(/(\w+)="(-?\d+)"/g)].map(a => [a[1], Number(a[2])])));
}
async function fixture() {
  const zip = new JSZip();
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>
    <w:p><w:r><w:drawing><wp:inline><wp:extent cx="4099560" cy="2059940"/><a:graphic><a:graphicData><c:chart r:id="chart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
    <w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0"><wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:align>top</wp:align></wp:positionV><wp:extent cx="6120130" cy="4081145"/><wp:wrapSquare wrapText="largest"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="photo"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:bottom="1134" w:left="1134" w:right="1134" w:header="0" w:footer="0"/></w:sectPr></w:body></w:document>`);
  zip.file('word/_rels/document.xml.rels', '<Relationships><Relationship Id="chart" Target="charts/chart1.xml"/><Relationship Id="photo" Target="media/image1.png"/></Relationships>');
  zip.file('word/charts/chart1.xml', CHART);
  zip.file('word/media/image1.png', PNG, { base64: true });
  return zip.generateAsync({ type: 'uint8array' });
}

describe('DOCX chart, first-line and body-edge fidelity', () => {
  it('renders cached native chart data into a valid, nonblank PNG and retains the original inline extent', async () => {
    const doc = await open(await fixture());
    const img = (doc.kids[0].kids[0] as ParaNode).kids[0] as ImgNode;
    expect(img).toMatchObject({ w: 322.8, h: 162.2, layout: { wrap: 'inline' } });
    const png = Uint8Array.from(Buffer.from(img.b64, 'base64'));
    expect([...png.subarray(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    expect(new DataView(png.buffer).getUint32(16)).toBe(646);
    expect(new DataView(png.buffer).getUint32(20)).toBe(325);
    const idat: Uint8Array[] = [];
    for (let o = 8; o < png.length;) {
      const n = new DataView(png.buffer).getUint32(o);
      if (String.fromCharCode(...png.subarray(o + 4, o + 8)) === 'IDAT') idat.push(png.subarray(o + 8, o + 8 + n));
      o += n + 12;
    }
    const raw = pako.inflate(Buffer.concat(idat));
    expect(raw.filter(b => b !== 255 && b !== 0).length).toBeGreaterThan(1000);
    expect(raw.some((b, i) => b === 0 && raw[i + 1] === 69 && raw[i + 2] === 134)).toBe(true);
  });

  it('rejects combination charts rather than silently discarding some series', async () => {
    await expect(renderDocxChart(CHART.replace('</c:barChart>', '</c:barChart><c:lineChart/>'), 300, 160)).rejects.toThrow('unsupported chart type');
  });

  it('reports missing portable glyphs rather than changing chart labels to question marks', async () => {
    await expect(renderDocxChart(CHART.replace('Column 1', '한글'), 300, 160)).rejects.toThrow('browser canvas');
  });

  for (const format of ['hwp', 'hwpx']) {
    it(`${format} embeds chart plus photo, preserves inline/floating anchors and 20mm body edges`, async () => {
      const doc = await open(await fixture());
      expect(doc.kids[0].dims).toMatchObject({ wPt: 595.3, hPt: 841.9, mt: 56.7, mb: 56.7, ml: 56.7, mr: 56.7 });
      const photo = (doc.kids[0].kids[1] as ParaNode).kids[0] as ImgNode;
      expect(photo.layout).toMatchObject({ wrap: 'square', horzRelTo: 'column', vertRelTo: 'para', horzAlign: 'center', vertAlign: 'top' });
      const bytes = await encode(doc, format);
      if (format === 'hwp') {
        const body = records(bytes);
        const page = body.find(r => r.tag === 73)!.data;
        expect([i32(page, 8), i32(page, 12), i32(page, 16) + i32(page, 24), i32(page, 20) + i32(page, 28)]).toEqual([5670, 5670, 5670, 5670]);
        const pics = body.filter(r => r.tag === 71 && i32(r.data, 0) === 0x67736f20);
        expect(pics).toHaveLength(2);
        expect(i32(pics[0].data, 4) & 1).toBe(1);
        expect(i32(pics[1].data, 4) & 1).toBe(0);
        expect([...BinaryKit.parseCfb(bytes).keys()].filter(k => k.startsWith('BinData/'))).toHaveLength(2);
      } else {
        const zip = await JSZip.loadAsync(bytes), xml = await section(bytes);
        expect(Object.keys(zip.files).filter(k => k.startsWith('BinData/') && !zip.files[k].dir)).toHaveLength(2);
        expect(xml).toContain('treatAsChar="1"'); expect(xml).toContain('treatAsChar="0"');
        expect(xml).toContain('horzRelTo="COLUMN"'); expect(xml).toContain('horzAlign="CENTER"');
        const m = /<hp:margin ([^>]+)\/>/.exec(xml)![1];
        const a = Object.fromEntries([...m.matchAll(/(\w+)="(\d+)"/g)].map(v => [v[1], Number(v[2])]));
        expect([a.left, a.right, a.top + a.header, a.bottom + a.footer]).toEqual([5670, 5670, 5670, 5670]);
      }
    });

    for (const [indent, first, expected] of [[36, -18, [1800, 3600]], [0, -21.6, [-2160, 0]]] as const) {
      it(`${format} caches signed first/rest line starts ${indent}/${first} and the correct usable width`, async () => {
        const span = buildSpan('first', { pt: 10 }); span.kids.push({ tag: 'br' }, { tag: 'txt', content: 'second' });
        const doc = buildRoot({}, [buildSheet([buildPara([span], { indentPt: indent, firstLineIndentPt: first, indentRightPt: 18 })], { wPt: 600, hPt: 800, ml: 50, mr: 50, mt: 50, mb: 50 })]);
        const bytes = await encode(doc, format);
        if (format === 'hwp') {
          const d = records(bytes).find(r => r.tag === 69)!.data;
          expect([i32(d, 24), i32(d, 60)]).toEqual(expected);
          expect([i32(d, 28), i32(d, 64)]).toEqual(expected.map(x => 50000 - x - 1800));
        } else {
          const l = lines(await section(bytes));
          expect(l.map(s => s.horzpos)).toEqual(expected);
          expect(l.map(s => s.horzsize)).toEqual(expected.map(x => 50000 - x - 1800));
        }
      });
    }

    it(`${format} advances an inline image by its height plus font leading, without scaling the image height`, async () => {
      const doc = buildRoot({}, [buildSheet([
        buildPara([buildSpan('', { pt: 10 }), buildImg(PNG, 'image/png', 100, 100)], { lineHeight: 1.2 }),
        buildPara([buildSpan('after', { pt: 10 })]),
      ])]);
      const bytes = await encode(doc, format);
      if (format === 'hwp') {
        const l = records(bytes).filter(r => r.tag === 69).map(r => r.data);
        expect(i32(l[0], 16)).toBe(10000);
        expect(i32(l[1], 4)).toBe(10200);
      } else {
        const l = lines(await section(bytes));
        expect(l[0].baseline).toBe(10000); expect(l[1].vertpos).toBe(10200);
      }
    });
  }

  it('inherits numbering-level hanging indent, honors zero overrides, and suppresses numFmt=none markers', async () => {
    const zip = new JSZip();
    zip.file('word/numbering.xml', `<w:numbering xmlns:w="w"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="none"/></w:lvl></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
    const p = (level: number, ind = '', num = 5) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${num}"/></w:numPr>${ind}</w:pPr><w:r><w:t>list</w:t></w:r></w:p>`;
    zip.file('word/document.xml', `<w:document xmlns:w="w"><w:body>${p(0)}${p(0, '<w:ind w:left="0" w:hanging="0"/>')}${p(1)}${p(0, '', 0)}</w:body></w:document>`);
    const doc = await open(await zip.generateAsync({ type: 'uint8array' }));
    const props = doc.kids[0].kids.map(p => (p as ParaNode).props);
    expect(props[0]).toMatchObject({ listLv: 0, listOrd: false, indentPt: 36, firstLineIndentPt: -18 });
    expect(props[1]).toMatchObject({ indentPt: 0, firstLineIndentPt: 0 });
    expect(props[2].listLv).toBeUndefined(); expect(props[3].listLv).toBeUndefined();
  });

  it('preserves empty paragraph-mark font sizes instead of expanding a deliberate small spacer to 10pt', async () => {
    const zip = new JSZip();
    zip.file('word/document.xml', '<w:document xmlns:w="w"><w:body><w:p><w:pPr><w:rPr><w:sz w:val="4"/></w:rPr><w:spacing w:line="240" w:before="0" w:after="0"/></w:pPr></w:p><w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:t>after</w:t></w:r></w:p></w:body></w:document>');
    const doc = await open(await zip.generateAsync({ type: 'uint8array' }));
    expect(((doc.kids[0].kids[0] as ParaNode).kids[0] as SpanNode).props.pt).toBe(2);
    const hwp = records(await encode(doc, 'hwp')).filter(r => r.tag === 69);
    expect(i32(hwp[0].data, 8)).toBe(200); expect(i32(hwp[1].data, 4)).toBe(200);
    const hwpx = lines(await section(await encode(doc, 'hwpx')));
    expect(hwpx[0].textheight).toBe(200); expect(hwpx[1].vertpos).toBe(200);
  });
});
