/**
 * DocxEmitVisitor - IR → DOCX (ZIP)
 *
 * WriteStrategy 구현체이며 NodeVisitor<string, DocxEmitContext>를 구현합니다.
 * OOXML 표준에 따른 DOCX 패키지를 생성합니다.
 *
 * IR 노드 → OOXML XML 문자열 변환:
 *   IrParagraphNode → <w:p>...</w:p>
 *   IrRunNode       → <w:r>...</w:r>
 *   IrTableNode     → <w:tbl>...</w:tbl>
 *   IrImageNode     → <w:drawing>...</w:drawing>
 */

import type { WriteStrategy } from '../core/strategy';
import type { NodeVisitor } from '../core/visitor';
import { dispatchVisit } from '../core/visitor';
import type {
  IrDocumentNode,
  IrSectionNode,
  IrParagraphNode,
  IrRunNode,
  IrTextNode,
  IrImageNode,
  IrHyperlinkNode,
  IrTableNode,
  IrTableRowNode,
  IrTableCellNode,
  IrLineBreakNode,
  IrPageBreakNode,
  RunStyle,
  PageLayout,
  DocumentMeta,
} from '../core/ir';
import { buildZip, ZipWriteEntry, zipToBlob } from '../services/zip';
import { escapeXml } from '../services/xml';
import { fromBase64 } from '../services/encoding';

// ─── 단위 변환 ────────────────────────────────────────────────────────────────

const PT_TO_TWIP = 20;         // 1pt = 20 twips
const PT_TO_EMU = 12700;       // 1pt = 12700 EMU
const PX_TO_EMU = 9525;        // 1px = 9525 EMU (96 DPI 기준)

function pt2Twip(pt: number): number { return Math.round(pt * PT_TO_TWIP); }
function pt2Emu(pt: number): number { return Math.round(pt * PT_TO_EMU); }
function px2Emu(px: number): number { return Math.round(px * PX_TO_EMU); }

// ─── ID 생성기 ────────────────────────────────────────────────────────────────

class IdCounter {
  private n = 0;
  next(): number { return ++this.n; }
}

// ─── Emit Context ─────────────────────────────────────────────────────────────

interface DocxEmitContext {
  relId: IdCounter;
  drawingId: IdCounter;
  images: { relId: string; name: string; data: Uint8Array; ext: string }[];
  rels: { id: string; type: string; target: string }[];
  insideCell: boolean;
}

function makeContext(): DocxEmitContext {
  return {
    relId: new IdCounter(),
    drawingId: new IdCounter(),
    images: [],
    rels: [],
    insideCell: false,
  };
}

// ─── DocxEmitVisitor ──────────────────────────────────────────────────────────

class DocxEmitVisitor implements NodeVisitor<string, DocxEmitContext> {

  visitDocument(_node: IrDocumentNode, _ctx: DocxEmitContext): string {
    return '';
  }

  visitSection(node: IrSectionNode, ctx: DocxEmitContext): string {
    const blocks = node.blocks.map(block => dispatchVisit(this, block, ctx)).join('\n');
    const sectPr = buildSectPr(node.layout);
    return `${blocks}\n${sectPr}`;
  }

  visitParagraph(node: IrParagraphNode, ctx: DocxEmitContext): string {
    const pPr = buildPPr(node);
    const content = node.children.map(child => dispatchVisit(this, child, ctx)).join('');
    return `<w:p>\n${pPr}\n${content}\n</w:p>`;
  }

  visitRun(node: IrRunNode, ctx: DocxEmitContext): string {
    const rPr = buildRPr(node.style);
    const content = node.children.map(child => dispatchVisit(this, child, ctx)).join('');
    return `<w:r>\n${rPr}\n${content}\n</w:r>`;
  }

  visitText(node: IrTextNode, _ctx: DocxEmitContext): string {
    return `<w:t xml:space="preserve">${escapeXml(node.value)}</w:t>`;
  }

  visitImage(node: IrImageNode, ctx: DocxEmitContext): string {
    const rId = `rId${ctx.relId.next()}`;
    const dId = ctx.drawingId.next();
    const ext = node.mimeType.split('/')[1] ?? 'png';
    const name = `media/image${ctx.images.length + 1}.${ext}`;

    ctx.images.push({ relId: rId, name, data: fromBase64(node.dataBase64), ext });
    ctx.rels.push({
      id: rId,
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      target: name,
    });

    const cx = px2Emu(node.widthPx);
    const cy = px2Emu(node.heightPx);
    const alt = escapeXml(node.altText ?? 'image');

    return `<w:r><w:drawing>
  <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
             distT="0" distB="0" distL="0" distR="0">
    <wp:extent cx="${cx}" cy="${cy}"/>
    <wp:docPr id="${dId}" name="${alt}"/>
    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:nvPicPr>
            <pic:cNvPr id="${dId}" name="${alt}"/>
            <pic:cNvPicPr/>
          </pic:nvPicPr>
          <pic:blipFill>
            <a:blip r:embed="${rId}"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
            <a:stretch><a:fillRect/></a:stretch>
          </pic:blipFill>
          <pic:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing></w:r>`;
  }

  visitHyperlink(node: IrHyperlinkNode, ctx: DocxEmitContext): string {
    const rId = `rId${ctx.relId.next()}`;
    ctx.rels.push({
      id: rId,
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
      target: node.url,
    });
    const runs = node.children.map(run => this.visitRun(run, ctx)).join('');
    return `<w:hyperlink r:id="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${runs}</w:hyperlink>`;
  }

  visitTable(node: IrTableNode, ctx: DocxEmitContext): string {
    const tblPr = `<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>
      <w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>
      <w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>
      <w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>
    </w:tblBorders></w:tblPr>`;
    const rows = node.rows.map(row => this.visitTableRow(row, ctx)).join('\n');
    return `<w:tbl>\n${tblPr}\n${rows}\n</w:tbl>`;
  }

  visitTableRow(node: IrTableRowNode, ctx: DocxEmitContext): string {
    const cells = node.cells.map(cell => this.visitTableCell(cell, ctx)).join('\n');
    return `<w:tr>\n${cells}\n</w:tr>`;
  }

  visitTableCell(node: IrTableCellNode, ctx: DocxEmitContext): string {
    const cellCtx: DocxEmitContext = { ...ctx, insideCell: true };
    const tcPr = buildTcPr(node);
    const content = node.children.map(p => this.visitParagraph(p, cellCtx)).join('\n');
    return `<w:tc>\n${tcPr}\n${content}\n</w:tc>`;
  }

  visitLineBreak(_node: IrLineBreakNode, _ctx: DocxEmitContext): string {
    return `<w:br/>`;
  }

  visitPageBreak(_node: IrPageBreakNode, _ctx: DocxEmitContext): string {
    return `</w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p>`;
  }
}

// ─── XML 빌더 함수들 ──────────────────────────────────────────────────────────

function buildPPr(node: IrParagraphNode): string {
  const { align, headingLevel, indentLeftPt, indentRightPt,
          spacingBeforePt, spacingAfterPt, lineSpacing } = node.style;

  const alignMap: Record<string, string> = {
    left: 'left', center: 'center', right: 'right', justify: 'both',
  };

  let xml = '<w:pPr>';
  if (headingLevel) xml += `<w:pStyle w:val="Heading${headingLevel}"/>`;
  if (align && align !== 'left') xml += `<w:jc w:val="${alignMap[align] ?? 'left'}"/>`;
  if (indentLeftPt || indentRightPt) {
    xml += `<w:ind w:left="${pt2Twip(indentLeftPt ?? 0)}" w:right="${pt2Twip(indentRightPt ?? 0)}"/>`;
  }
  if (spacingBeforePt || spacingAfterPt || lineSpacing) {
    const before = pt2Twip(spacingBeforePt ?? 0);
    const after = pt2Twip(spacingAfterPt ?? 0);
    const line = lineSpacing ? Math.round(lineSpacing * 240) : 240;
    xml += `<w:spacing w:before="${before}" w:after="${after}" w:line="${line}" w:lineRule="auto"/>`;
  }
  xml += '</w:pPr>';
  return xml;
}

function buildRPr(style: RunStyle): string {
  let xml = '<w:rPr>';
  if (style.bold)          xml += '<w:b/><w:bCs/>';
  if (style.italic)        xml += '<w:i/><w:iCs/>';
  if (style.underline)     xml += '<w:u w:val="single"/>';
  if (style.strikethrough) xml += '<w:strike/>';
  if (style.superscript)   xml += '<w:vertAlign w:val="superscript"/>';
  if (style.subscript)     xml += '<w:vertAlign w:val="subscript"/>';
  if (style.fontSizePt) {
    const hp = Math.round(style.fontSizePt * 2);
    xml += `<w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/>`;
  }
  if (style.fontFamily) {
    const f = escapeXml(style.fontFamily);
    xml += `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}"/>`;
  }
  if (style.colorHex) xml += `<w:color w:val="${style.colorHex.padStart(6, '0')}"/>`;
  xml += '</w:rPr>';
  return xml;
}

function buildTcPr(cell: IrTableCellNode): string {
  let xml = '<w:tcPr>';
  if (cell.colSpan > 1) xml += `<w:gridSpan w:val="${cell.colSpan}"/>`;
  if (cell.rowSpan === 0) xml += '<w:vMerge/>';
  else if (cell.rowSpan > 1) xml += '<w:vMerge w:val="restart"/>';
  if (cell.style.bgColorHex) xml += `<w:shd w:val="clear" w:fill="${cell.style.bgColorHex}"/>`;
  xml += '</w:tcPr>';
  return xml;
}

function buildSectPr(layout: PageLayout): string {
  return `<w:sectPr>
  <w:pgSz w:w="${pt2Twip(layout.widthPt)}" w:h="${pt2Twip(layout.heightPt)}"${layout.orientation === 'landscape' ? ' w:orient="landscape"' : ''}/>
  <w:pgMar w:top="${pt2Twip(layout.marginTopPt)}"
           w:bottom="${pt2Twip(layout.marginBottomPt)}"
           w:left="${pt2Twip(layout.marginLeftPt)}"
           w:right="${pt2Twip(layout.marginRightPt)}"/>
</w:sectPr>`;
}

// ─── DOCX 패키지 XML 생성 ─────────────────────────────────────────────────────

const DOCX_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml"
    ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
    Target="docProps/core.xml"/>
</Relationships>`;

function buildDocumentRels(rels: DocxEmitContext['rels']): string {
  const items = rels.map(r =>
    `  <Relationship Id="${r.id}" Type="${r.type}" Target="${escapeXml(r.target)}"${r.type.includes('hyperlink') ? ' TargetMode="External"' : ''}/>`,
  ).join('\n');
  return `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${items}
</Relationships>`;
}

function buildDocumentXml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
${body}
  </w:body>
</w:document>`;
}

function buildStylesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr><w:jc w:val="left"/></w:pPr>
  </w:style>
  ${[1,2,3,4,5,6].map(n => `<w:style w:type="paragraph" w:styleId="Heading${n}">
    <w:name w:val="heading ${n}"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr><w:b/><w:sz w:val="${Math.max(24, 40 - n * 4)}"/></w:rPr>
  </w:style>`).join('\n  ')}
</w:styles>`;
}

function buildCoreXml(meta: DocumentMeta): string {
  const now = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:title>${escapeXml(meta.title ?? '')}</dc:title>
  <dc:creator>${escapeXml(meta.author ?? '')}</dc:creator>
  <dc:subject>${escapeXml(meta.subject ?? '')}</dc:subject>
  <dc:description>${escapeXml(meta.description ?? '')}</dc:description>
  <dcterms:created>${meta.createdAt ?? now}</dcterms:created>
  <dcterms:modified>${now}</dcterms:modified>
</cp:coreProperties>`;
}

// ─── DocxWriter ───────────────────────────────────────────────────────────────

export class DocxWriter implements WriteStrategy<Blob> {
  async write(document: IrDocumentNode): Promise<Blob> {
    const visitor = new DocxEmitVisitor();
    const ctx = makeContext();

    const bodyParts: string[] = [];
    for (const section of document.sections) {
      bodyParts.push(visitor.visitSection(section, ctx));
    }

    const docBody = bodyParts.join('\n');
    const documentXml = buildDocumentXml(docBody);

    const entries: ZipWriteEntry[] = [
      { path: '[Content_Types].xml', data: DOCX_CONTENT_TYPES },
      { path: '_rels/.rels', data: ROOT_RELS },
      { path: 'word/document.xml', data: documentXml },
      { path: 'word/styles.xml', data: buildStylesXml() },
      { path: 'word/_rels/document.xml.rels', data: buildDocumentRels(ctx.rels) },
      { path: 'docProps/core.xml', data: buildCoreXml(document.meta) },
    ];

    for (const img of ctx.images) {
      entries.push({ path: `word/${img.name}`, data: img.data });
    }

    const zipData = await buildZip(entries);
    const buf = zipData.buffer.slice(zipData.byteOffset, zipData.byteOffset + zipData.byteLength) as ArrayBuffer;
    return new Blob([buf], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
  }
}
