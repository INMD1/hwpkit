import type { Encoder } from '../../contract/encoder';
import type { DocRoot, ParaNode, SpanNode, GridNode, ContentNode } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import type { PageDims } from '../../model/doc-props';
import { A4 } from '../../model/doc-props';
import { succeed, fail } from '../../contract/result';
import { Metric } from '../../safety/StyleBridge';
import { ArchiveKit } from '../../toolkit/ArchiveKit';
import { TextKit } from '../../toolkit/TextKit';
import { registry } from '../../pipeline/registry';

export class DocxEncoder implements Encoder {
  readonly format = 'docx';

  async encode(doc: DocRoot): Promise<Outcome<Uint8Array>> {
    try {
      const sheet = doc.kids[0];
      const dims  = sheet?.dims ?? A4;
      const kids  = sheet?.kids ?? [];

      const entries = [
        { name: '[Content_Types].xml',            data: TextKit.encode(contentTypes()) },
        { name: '_rels/.rels',                    data: TextKit.encode(pkgRels()) },
        { name: 'word/document.xml',              data: TextKit.encode(documentXml(kids, dims)) },
        { name: 'word/_rels/document.xml.rels',   data: TextKit.encode(docRels()) },
        { name: 'docProps/app.xml',               data: TextKit.encode(appXml()) },
        { name: 'docProps/core.xml',              data: TextKit.encode(coreXml(doc.meta)) },
      ];

      return succeed(await ArchiveKit.zip(entries));
    } catch (e: any) {
      return fail(`DOCX encode error: ${e?.message ?? String(e)}`);
    }
  }
}

// ─── OOXML boilerplate ──────────────────────────────────────

function contentTypes(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
}

function pkgRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
}

function docRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
}

function appXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>hwpkit</Application>
</Properties>`;
}

function coreXml(meta: any): string {
  const now = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${esc(meta.title ?? '')}</dc:title>
  <dc:creator>${esc(meta.author ?? '')}</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">${meta.created ?? now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`;
}

// ─── document.xml ───────────────────────────────────────────

function documentXml(kids: ContentNode[], dims: PageDims): string {
  const body = kids.map(encodeContent).join('\n');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
${body}
    <w:sectPr>
      <w:pgSz w:w="${Metric.ptToDxa(dims.wPt)}" w:h="${Metric.ptToDxa(dims.hPt)}" w:orient="${dims.orient ?? 'portrait'}"/>
      <w:pgMar w:top="${Metric.ptToDxa(dims.mt)}" w:right="${Metric.ptToDxa(dims.mr)}" w:bottom="${Metric.ptToDxa(dims.mb)}" w:left="${Metric.ptToDxa(dims.ml)}" w:header="709" w:footer="709" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

function encodeContent(node: ContentNode): string {
  return node.tag === 'grid' ? encodeGrid(node) : encodePara(node);
}

function encodePara(para: ParaNode): string {
  const align = para.props.align ?? 'left';
  const headStyle = para.props.heading ? `<w:pStyle w:val="Heading${para.props.heading}"/>` : '';

  const runs = para.kids.map(k => k.tag === 'span' ? encodeRun(k) : '').join('');
  return `    <w:p>
      <w:pPr>${headStyle}<w:jc w:val="${align === 'justify' ? 'both' : align}"/></w:pPr>
      ${runs}
    </w:p>`;
}

function encodeRun(span: SpanNode): string {
  const p = span.props;
  const rPr: string[] = [];
  if (p.b) rPr.push('<w:b/>');
  if (p.i) rPr.push('<w:i/>');
  if (p.u) rPr.push('<w:u w:val="single"/>');
  if (p.s) rPr.push('<w:strike/>');
  if (p.pt) rPr.push(`<w:sz w:val="${Metric.ptToHalfPt(p.pt)}"/>`);
  if (p.color) rPr.push(`<w:color w:val="${p.color}"/>`);
  if (p.font) rPr.push(`<w:rFonts w:ascii="${p.font}" w:hAnsi="${p.font}" w:eastAsia="${p.font}"/>`);

  const content = span.kids.filter(k => k.tag === 'txt').map(k => k.tag === 'txt' ? esc(k.content) : '').join('');
  return `<w:r><w:rPr>${rPr.join('')}</w:rPr><w:t xml:space="preserve">${content}</w:t></w:r>`;
}

function encodeGrid(grid: GridNode): string {
  const rows = grid.kids.map(row => {
    const cells = row.kids.map(cell => {
      const tcPr = cell.props.bg
        ? `<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${cell.props.bg}"/></w:tcPr>`
        : '<w:tcPr/>';
      return `      <w:tc>${tcPr}${cell.kids.map(p => encodePara(p)).join('')}</w:tc>`;
    }).join('\n');
    return `    <w:tr>\n${cells}\n    </w:tr>`;
  }).join('\n');

  return `    <w:tbl>
      <w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>
${rows}
    </w:tbl>`;
}

function esc(s: string): string { return TextKit.escapeXml(s); }

registry.registerEncoder(new DocxEncoder());
