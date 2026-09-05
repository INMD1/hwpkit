import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { Pipeline, buildRoot, buildSheet, buildPara, buildSpan, buildGrid, buildRow, buildCell, buildImg, A4, BinaryKit } from '../index';
import { DocxEncoder } from '../encoders/docx/DocxEncoder';
import { HwpEncoder } from '../encoders/hwp/HwpEncoder';
import { HwpxEncoder } from '../encoders/hwpx/HwpxEncoder';

const flags = { pageBreakBefore: true, keepWithNext: true, keepLines: true, widowControl: false };
async function parts(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  return { xml: await zip.file('word/document.xml')!.async('string'), styles: await zip.file('word/styles.xml')!.async('string') };
}

describe('Pagination and table style interoperability', () => {
  it('defines every native style used in nested tables and emits one pStyle per list paragraph', async () => {
    const nested = buildGrid([buildRow([buildCell([buildPara([buildSpan('nested')], { hwpStyleId: 44, listOrd: true })])])]);
    const doc = buildRoot({}, [buildSheet([buildGrid([buildRow([buildCell([nested])])])])]);
    const encoded = await new DocxEncoder().encode(doc);
    if (!encoded.ok) throw new Error(encoded.error);
    const { xml, styles } = await parts(encoded.data);
    const defined = new Set([...styles.matchAll(/w:styleId="([^"]+)"/g)].map(m => m[1]));
    for (const m of xml.matchAll(/<w:pStyle w:val="([^"]+)"/g)) expect(defined.has(m[1])).toBe(true);
    expect(xml).toContain('<w:pStyle w:val="44"/>');
    for (const p of xml.matchAll(/<w:pPr>([\s\S]*?)<\/w:pPr>/g)) expect((p[1].match(/<w:pStyle /g) ?? []).length).toBeLessThanOrEqual(1);
  });

  for (const [format, Encoder] of [['docx', DocxEncoder], ['hwp', HwpEncoder], ['hwpx', HwpxEncoder]] as const) {
    it(`${format} preserves pagination switches and signed right margins through DOCX`, async () => {
      const encoded = await new Encoder().encode(buildRoot({}, [buildSheet([
        buildPara([buildSpan('new page')], { ...flags, indentRightPt: -3 }),
        buildPara([buildSpan('ordinary')], { pageBreakBefore: false, keepWithNext: false, keepLines: false, widowControl: true }),
      ])]));
      if (!encoded.ok) throw new Error(encoded.error);
      const docx = await Pipeline.open(encoded.data, format).to('docx');
      if (!docx.ok) throw new Error(docx.error);
      const decoded = await Pipeline.open(docx.data, 'docx').inspect();
      if (!decoded.ok) throw new Error(decoded.error);
      const paras = decoded.data.kids.flatMap(s => s.kids).filter(p => p.tag === 'para');
      const first = paras.find(p => JSON.stringify(p.kids).includes('new page'))!;
      expect(first.props).toMatchObject({ ...flags, indentRightPt: -3 });
      const second = paras.find(p => JSON.stringify(p.kids).includes('ordinary'))!;
      expect(second.props).toMatchObject({ pageBreakBefore: false, keepWithNext: false, keepLines: false, widowControl: true });
      expect((await parts(docx.data)).xml).not.toContain('<w:br w:type="page"/>');
    });
  }

  it('honors inherited pagination flags and explicit false overrides in an independent DOCX fixture', async () => {
    const zip = new JSZip();
    zip.file('word/styles.xml', `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Parent"><w:pPr><w:keepNext/><w:keepLines/><w:pageBreakBefore/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Parent"/><w:pPr><w:keepLines w:val="0"/></w:pPr></w:style></w:styles>`);
    zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Child"/><w:pageBreakBefore w:val="false"/></w:pPr><w:r><w:t>override</w:t></w:r></w:p></w:body></w:document>`);
    const decoded = await Pipeline.open(await zip.generateAsync({ type: 'uint8array' }), 'docx').inspect();
    if (!decoded.ok) throw new Error(decoded.error);
    expect(decoded.data.kids[0].kids[0].props).toMatchObject({ keepWithNext: true, keepLines: false, pageBreakBefore: false });
  });

  it('reads HWPX break settings without inserting an empty line before the text', async () => {
    const zip = new JSZip();
    zip.file('Contents/header.xml', `<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"><hh:refList><hh:paraProperties><hh:paraPr id="0"><hh:breakSetting widowOrphan="0" keepWithNext="1" keepLines="1" pageBreakBefore="1"/></hh:paraPr></hh:paraProperties></hh:refList></hh:head>`);
    zip.file('Contents/section0.xml', `<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"><hp:p paraPrIDRef="0" pageBreak="1"><hp:run><hp:t>new page</hp:t></hp:run></hp:p></hs:sec>`);
    const result = await Pipeline.open(await zip.generateAsync({ type: 'uint8array' }), 'hwpx').to('docx');
    if (!result.ok) throw new Error(result.error);
    const { xml } = await parts(result.data);
    expect(xml).toContain('<w:pageBreakBefore/>');
    expect(xml).toContain('<w:keepNext/>');
    expect(xml).not.toContain('<w:br w:type="page"/>');
  });
  it('preserves section sizes and continuous/odd-page starts without synthetic break paragraphs', async () => {
    const zip = new JSZip();
    zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:sectPr><w:pgSz w:w="12000" w:h="16000"/></w:sectPr></w:pPr><w:r><w:t>portrait</w:t></w:r></w:p>
      <w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="16000" w:h="12000" w:orient="landscape"/></w:sectPr></w:pPr><w:r><w:t>continuous</w:t></w:r></w:p>
      <w:p><w:r><w:t>odd page</w:t></w:r></w:p><w:sectPr><w:type w:val="oddPage"/><w:pgSz w:w="14000" w:h="18000"/></w:sectPr>
    </w:body></w:document>`);
    let bytes = await zip.generateAsync({ type: 'uint8array' });
    for (let i = 0; i < 3; i++) {
      const decoded = await Pipeline.open(bytes, 'docx').inspect();
      if (!decoded.ok) throw new Error(decoded.error);
      const sheets = decoded.data.kids;
      expect(sheets).toHaveLength(3);
      expect(sheets.map(s => [s.dims.wPt, s.dims.hPt])).toEqual([[600, 800], [800, 600], [700, 900]]);
      expect(sheets[1].sectionType).toBe('continuous');
      expect(sheets[2].sectionType).toBe('oddPage');
      expect(sheets.flatMap(s => s.kids)).toHaveLength(3);
      const encoded = await new DocxEncoder().encode(decoded.data);
      if (!encoded.ok) throw new Error(encoded.error);
      bytes = encoded.data;
      expect((await parts(bytes)).xml).not.toContain('<w:br w:type="page"/>');
    }
  });

  it('keeps distinct section headers, first/even variants, and empty overrides', async () => {
    const doc = buildRoot({}, [
      buildSheet([buildPara([buildSpan('body1')])], undefined, { headers: { default: [buildPara([buildSpan('header1')])], first: [buildPara([buildSpan('first')])] } }),
      buildSheet([buildPara([buildSpan('body2')])], undefined, { headers: { default: [buildPara([buildSpan('header2')])], even: [buildPara([buildSpan('even')])] }, footers: { default: [] } }),
    ]);
    const encoded = await new DocxEncoder().encode(doc);
    if (!encoded.ok) throw new Error(encoded.error);
    const decoded = await Pipeline.open(encoded.data, 'docx').inspect();
    if (!decoded.ok) throw new Error(decoded.error);
    expect(JSON.stringify(decoded.data.kids[0].headers?.default)).toContain('header1');
    expect(JSON.stringify(decoded.data.kids[1].headers?.default)).toContain('header2');
    expect(JSON.stringify(decoded.data.kids[0].headers?.first)).toContain('first');
    expect(JSON.stringify(decoded.data.kids[1].headers?.even)).toContain('even');
    const zip = await JSZip.loadAsync(encoded.data);
    expect(await zip.file('word/settings.xml')!.async('string')).toContain('<w:evenAndOddHeaders/>');
    expect((await parts(encoded.data)).xml).toContain('<w:titlePg/>');
  });

  it('places the break before a HWPX table rather than after its anchor', async () => {
    const zip = new JSZip();
    zip.file('Contents/section0.xml', `<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"><hp:p pageBreak="1"><hp:run><hp:tbl rowCnt="1" colCnt="1"><hp:tr><hp:tc><hp:subList><hp:p><hp:run><hp:t>table on new page</hp:t></hp:run></hp:p></hp:subList><hp:cellAddr colAddr="0" rowAddr="0"/><hp:cellSpan colSpan="1" rowSpan="1"/><hp:cellSz width="10000" height="2000"/></hp:tc></hp:tr></hp:tbl></hp:run></hp:p></hs:sec>`);
    const decoded = await Pipeline.open(await zip.generateAsync({ type: 'uint8array' }), 'hwpx').inspect();
    if (!decoded.ok) throw new Error(decoded.error);
    const kids = decoded.data.kids[0].kids;
    expect(kids[0].tag).toBe('para');
    expect(kids[0].props).toMatchObject({ pageBreakBefore: true, lineHeightFixed: 0.05 });
    expect(kids[1].tag).toBe('grid');
    expect(kids[2].props).toMatchObject({ pageBreakBefore: false });
  });

  it('preserves an explicit leading HWP page break and its intentional blank first page', async () => {
    const doc = buildRoot({}, [buildSheet([buildPara([
      { tag: 'span', props: {}, kids: [{ tag: 'pb' }] }, buildSpan('after blank page'),
    ])])]);
    const hwp = await new HwpEncoder().encode(doc);
    if (!hwp.ok) throw new Error(hwp.error);
    const result = await Pipeline.open(hwp.data, 'hwp').to('docx');
    if (!result.ok) throw new Error(result.error);
    const { xml } = await parts(result.data);
    expect(xml.indexOf('<w:br w:type="page"/>')).toBeGreaterThan(-1);
    expect(xml.indexOf('<w:br w:type="page"/>')).toBeLessThan(xml.indexOf('after blank page'));
  });

  it('does not activate unused first/even header variants during a DOCX round trip', async () => {
    const first = buildSheet([buildPara([buildSpan('body')])], undefined, { headers: {
      first: [buildPara([buildSpan('unused first')])], even: [buildPara([buildSpan('unused even')])],
    } });
    first.differentFirstPage = false;
    const encoded = await new DocxEncoder().encode(buildRoot({ evenAndOddHeaders: false }, [first]));
    if (!encoded.ok) throw new Error(encoded.error);
    const decoded = await Pipeline.open(encoded.data, 'docx').inspect();
    if (!decoded.ok) throw new Error(decoded.error);
    expect(decoded.data.meta.evenAndOddHeaders).toBe(false);
    expect(decoded.data.kids[0].differentFirstPage).toBe(false);
    const again = await new DocxEncoder().encode(decoded.data);
    if (!again.ok) throw new Error(again.error);
    expect((await parts(again.data)).xml).not.toContain('<w:titlePg/>');
    const zip = await JSZip.loadAsync(again.data);
    expect(await zip.file('word/settings.xml')!.async('string')).not.toContain('<w:evenAndOddHeaders/>');
  });

  for (const [format, Encoder] of [['hwp', HwpEncoder], ['hwpx', HwpxEncoder]] as const) {
    it(`${format} preserves all twelve sections, their sizes, tables and shared image references`, async () => {
      const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Zl1sAAAAASUVORK5CYII=';
      const doc = buildRoot({}, Array.from({ length: 12 }, (_, i) => buildSheet([
        buildPara([buildSpan(`SECTION ${i}`)]),
        buildGrid([buildRow([buildCell([buildPara([buildSpan(`CELL ${i}`)])])])]),
        buildPara([buildImg(png, 'image/png', 5, 5)]),
      ], { ...A4, wPt: 500 + i * 10 })));
      const encoded = await new Encoder().encode(doc);
      if (!encoded.ok) throw new Error(encoded.error);
      if (format === 'hwp') {
        const streams = BinaryKit.parseCfb(encoded.data);
        expect(streams.has('BodyText/Section11')).toBe(true);
      } else {
        const zip = await JSZip.loadAsync(encoded.data);
        expect(zip.file('Contents/section11.xml')).not.toBeNull();
        expect(await zip.file('Contents/header.xml')!.async('string')).toContain('secCnt="12"');
        expect(await zip.file('Contents/content.hpf')!.async('string')).toContain('idref="section11"');
      }
      const decoded = await Pipeline.open(encoded.data, format).inspect();
      if (!decoded.ok) throw new Error(decoded.error);
      expect(decoded.data.kids).toHaveLength(12);
      decoded.data.kids.forEach((sheet, i) => {
        expect(sheet.dims.wPt).toBeCloseTo(500 + i * 10, 1);
        expect(JSON.stringify(sheet.kids)).toContain(`SECTION ${i}`);
        expect(JSON.stringify(sheet.kids)).toContain(`CELL ${i}`);
        expect(JSON.stringify(sheet.kids)).toContain('image/png');
      });
    });
  }

  it('preserves Hancom 160 percent as DOCX 384 line units without compressing pagination', async () => {
    const zip = new JSZip();
    zip.file('Contents/header.xml', `<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"><hh:refList><hh:paraProperties><hh:paraPr id="0"><hh:lineSpacing type="PERCENT" value="160"/></hh:paraPr></hh:paraProperties></hh:refList></hh:head>`);
    zip.file('Contents/section0.xml', `<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"><hp:p paraPrIDRef="0"><hp:run><hp:t>line spacing</hp:t></hp:run></hp:p></hs:sec>`);
    const result = await Pipeline.open(await zip.generateAsync({ type: 'uint8array' }), 'hwpx').to('docx');
    if (!result.ok) throw new Error(result.error);
    expect((await parts(result.data)).xml).toContain('w:line="384" w:lineRule="auto"');
  });

});
