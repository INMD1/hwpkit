import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { Pipeline } from '../../pipeline/Pipeline';
import { buildRoot, buildSheet, buildPara, buildSpan } from '../../model/builders';
import { DocxEncoder } from '../../encoders/docx/DocxEncoder';
import { DocxDecoder } from './DocxDecoder';
import type { ParaNode } from '../../model/doc-tree';

async function decode(ind: string, styleInd = 'w:left="720" w:right="360" w:firstLine="240"') {
  const zip = new JSZip();
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  zip.file('word/styles.xml', `<w:styles xmlns:w="${ns}"><w:style w:styleId="Base" w:type="paragraph"><w:pPr><w:ind ${styleInd}/></w:pPr></w:style><w:style w:styleId="Child" w:type="paragraph"><w:basedOn w:val="Base"/><w:pPr><w:jc w:val="left"/></w:pPr></w:style></w:styles>`);
  zip.file('word/document.xml', `<w:document xmlns:w="${ns}"><w:body><w:p><w:pPr><w:pStyle w:val="Child"/><w:ind ${ind}/></w:pPr><w:r><w:t>indent</w:t></w:r></w:p></w:body></w:document>`);
  const result = await new DocxDecoder().decode(await zip.generateAsync({ type: 'uint8array' }));
  if (!result.ok) throw new Error(result.error);
  return (result.data.kids[0].kids[0] as ParaNode).props;
}
describe('ECMA-376 §17.3.1.12 indentation', () => {
  it('retains base style values through a child style with no indentation', async () => {
    expect(await decode('')).toMatchObject({ indentPt: 36, indentRightPt: 18, firstLineIndentPt: 12 });
  });
  it('explicit zero overrides inherited values', async () => {
    expect(await decode('w:left="0" w:right="0" w:firstLine="0"')).toMatchObject({ indentPt: 0, indentRightPt: 0, firstLineIndentPt: 0 });
  });
  it('hanging takes precedence over firstLine, including zero', async () => {
    expect((await decode('w:firstLine="240" w:hanging="360"')).firstLineIndentPt).toBe(-18);
    expect((await decode('w:firstLine="240" w:hanging="0"')).firstLineIndentPt).toBe(0);
    expect((await decode('', 'w:firstLine="240" w:hanging="360"')).firstLineIndentPt).toBe(-18);
  });
  it('retains negative left and right indentation', async () => {
    expect(await decode('w:left="-240" w:right="-120"')).toMatchObject({ indentPt: -12, indentRightPt: -6 });
  });
});


it('DOCX round trips do not accumulate hanging indentation or erase negative/zero values', async () => {
  for (const props of [
    { indentPt: 36, indentRightPt: 18, firstLineIndentPt: -18 },
    { indentPt: -12, indentRightPt: -6, firstLineIndentPt: 0 },
    { indentPt: 0, indentRightPt: 0, firstLineIndentPt: 0 },
  ]) {
    const source = buildRoot({}, [buildSheet([buildPara([buildSpan('문단')], props)])]);
    let encoded = await new DocxEncoder().encode(source);
    for (let i = 0; i < 3; i++) {
      if (!encoded.ok) throw new Error(encoded.error);
      const decoded = await Pipeline.open(encoded.data, 'docx').inspect();
      if (!decoded.ok) throw new Error(decoded.error);
      expect((decoded.data.kids[0].kids[0] as ParaNode).props).toMatchObject(props);
      encoded = await Pipeline.open(encoded.data, 'docx').to('docx');
    }
  }
});
