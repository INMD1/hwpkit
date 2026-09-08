import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { Pipeline } from '../index';
import { DocxDecoder } from './docx/DocxDecoder';
import { DocxEncoder } from '../encoders/docx/DocxEncoder';
import { HwpxDecoder } from './hwpx/HwpxDecoder';
import { HwpxEncoder } from '../encoders/hwpx/HwpxEncoder';
import { HwpEncoder } from '../encoders/hwp/HwpEncoder';
import {
  buildRoot,
  buildSheet,
  buildPara,
  buildSpan,
  buildGrid,
  buildRow,
  buildCell,
} from '../model/builders';
import type { SpanNode, ParaNode } from '../model/doc-tree';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** Flatten a paragraph's spans into their inline kids, in order. */
function paraInlineKids(para: ParaNode): any[] {
  const out: any[] = [];
  for (const k of para.kids) {
    if (k.tag === 'span') out.push(...(k as SpanNode).kids);
    else if (k.tag === 'link') for (const s of (k as any).kids) out.push(...s.kids);
  }
  return out;
}

/** A table cell whose single paragraph is: <t>one</t> <br/> <t>two</t> */
function docxCellWithBreak(bodyParaXml: string) {
  const zip = new JSZip();
  zip.file(
    'word/styles.xml',
    `<w:styles xmlns:w="${W_NS}"><w:style w:styleId="Normal" w:type="paragraph"/></w:styles>`,
  );
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="${W_NS}"><w:body>${bodyParaXml}<w:p/></w:body></w:document>`,
  );
  return zip.generateAsync({ type: 'uint8array' });
}

const cellBreakXml = `<w:tbl><w:tr><w:tc><w:p>` +
  `<w:r><w:t xml:space="preserve">one</w:t></w:r>` +
  `<w:r><w:br/></w:r>` +
  `<w:r><w:t xml:space="preserve">two</w:t></w:r>` +
  `</w:p></w:tc></w:tr></w:tbl>`;

describe('Line breaks (txt newline and <br/>)', () => {
  it('DocxEncoder turns a txt newline into <w:br/> and keeps <w:t> newline-free', async () => {
    // HWP decodes in-cell line breaks into TxtNode content with a raw \n.
    const doc = buildRoot({}, [
      buildSheet([
        buildPara([buildSpan('')]),
        buildGrid([
          buildRow([buildCell([buildPara([buildSpan('one\ntwo')])])]),
        ]),
      ]),
    ]);
    const out = await new DocxEncoder().encode(doc);
    if (!out.ok) throw new Error(out.error);
    const xml = await (await JSZip.loadAsync(out.data)).file('word/document.xml')!.async('string');
    expect(xml).toMatch(/<w:t xml:space="preserve">one<\/w:t>/);
    expect(xml).toContain('<w:br/>');
    expect(xml).toMatch(/<w:t xml:space="preserve">two<\/w:t>/);
    // No literal newline may survive inside a w:t element.
    for (const m of xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)) {
      expect(m[1]).not.toMatch(/[\r\n]/);
    }
  });

  it('DocxDecoder reads <w:br/> as a BrNode interleaved with text, and round trips', async () => {
    let bytes = await docxCellWithBreak(cellBreakXml);
    let format = 'docx';
    for (let i = 0; i < 3; i++) {
      const decoded = await Pipeline.open(bytes, format).inspect();
      if (!decoded.ok) throw new Error(decoded.error);
      const cell = ((decoded.data.kids[0].kids[0] as any).kids[0].kids[0]) as any; // grid>row>cell (table is the first body child)
      const para = cell.kids[0] as ParaNode;
      const kids = paraInlineKids(para);
      expect(kids.map((k: any) => k.tag)).toEqual(['txt', 'br', 'txt']);
      expect(kids.filter((k: any) => k.tag === 'txt').map((k: any) => k.content)).toEqual(['one', 'two']);
      const converted = await Pipeline.open(bytes, format).to('docx');
      if (!converted.ok) throw new Error(converted.error);
      bytes = converted.data;
      const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
      expect(xml).toContain('<w:br/>');
      for (const m of xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)) {
        expect(m[1]).not.toMatch(/[\r\n]/);
      }
    }
  });

  it('HWP round trip keeps an in-cell line break as a binary LF', async () => {
    const doc = buildRoot({}, [
      buildSheet([buildGrid([buildRow([buildCell([buildPara([buildSpan('one\ntwo')])])])])]),
    ]);
    const hwp = await new HwpEncoder().encode(doc);
    if (!hwp.ok) throw new Error(hwp.error);
    // 0x000A (LF) must appear in a PARA_TEXT record; re-decode confirms it survives.
    const round = await Pipeline.open(hwp.data, 'hwp').inspect();
    if (!round.ok) throw new Error(round.error);
    const joined = JSON.stringify(round.data).replace(/\\n/g, '\n');
    expect(joined).toMatch(/one\ntwo/);
  });

  it('HwpxEncoder turns a txt newline into a nested hp:lineBreak', async () => {
    const doc = buildRoot({}, [
      buildSheet([buildGrid([buildRow([buildCell([buildPara([buildSpan('one\ntwo')])])])])]),
    ]);
    const out = await new HwpxEncoder().encode(doc);
    if (!out.ok) throw new Error(out.error);
    const zip = await JSZip.loadAsync(out.data);
    const sec = zip.file('Contents/section0.xml');
    expect(sec).toBeTruthy();
    const xml = await sec!.async('string');
    expect(xml).toMatch(/<hp:t xml:space="preserve">one<\/hp:t>/);
    expect(xml).toContain('<hp:t><hp:lineBreak/></hp:t>');
    expect(xml).not.toContain('<hp:br/>');
    expect(xml).toMatch(/<hp:t xml:space="preserve">two<\/hp:t>/);
    for (const m of xml.matchAll(/<hp:t(?:\s[^>]*)?>([\s\S]*?)<\/hp:t>/g)) {
      expect(m[1]).not.toMatch(/[\r\n]/);
    }
  });

  it('HwpxDecoder reads <hp:br/> as a BrNode and round trips', async () => {
    const zip = new JSZip();
    zip.file('Contents/header.xml', '<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"/>');
    const section =
      '<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">' +
      '<hp:tbl id="1" numberingType="TABLE" zOrder="0" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" rowCnt="1" colCnt="1" cellSpacing="0" borderFillIDRef="0">' +
      '<hp:tr><hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="0">' +
      '<hp:p><hp:run><hp:t xml:space="preserve">one</hp:t><hp:br/><hp:t xml:space="preserve">two</hp:t></hp:run></hp:p>' +
      '</hp:tc></hp:tr></hp:tbl>' +
      '</hs:sec>';
    zip.file('Contents/section0.xml', section);
    const bytes = await zip.generateAsync({ type: 'uint8array' });

    const decoded = await new HwpxDecoder().decode(bytes);
    if (!decoded.ok) throw new Error(decoded.error);
    const cell = ((decoded.data.kids[0].kids[0] as any).kids[0].kids[0]) as any;
    const kids = paraInlineKids(cell.kids[0] as ParaNode);
    expect(kids.map((k: any) => k.tag)).toEqual(['txt', 'br', 'txt']);

    const re = await new HwpxEncoder().encode(decoded.data);
    if (!re.ok) throw new Error(re.error);
    const xml = await (await JSZip.loadAsync(re.data)).file('Contents/section0.xml')!.async('string');
    expect(xml).toContain('<hp:t><hp:lineBreak/></hp:t>');
    expect(xml).not.toContain('<hp:br/>');
  });
});
