/**
 * HwpxEmitVisitor - IR → HWPX (ZIP)
 *
 * WriteStrategy 구현체이며 NodeVisitor<string, HwpxEmitContext>를 구현합니다.
 * HWPX는 ZIP 컨테이너이며 내부에 XML 파일들을 포함합니다.
 *
 * IR 노드 → XML 문자열 변환:
 *   IrDocumentNode  → header.xml + section0.xml + manifest
 *   IrSectionNode   → section{n}.xml 내용
 *   IrParagraphNode → <hp:p>...</hp:p>
 *   IrRunNode       → <hp:run>...</hp:run>
 *   IrTableNode     → <hp:tbl>...</hp:tbl>
 *   IrImageNode     → <hp:ctrl id="GENSM">...</hp:ctrl>
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
  ParagraphStyle,
  PageLayout,
  DocumentMeta,
} from '../core/ir';
import { buildZip, ZipWriteEntry, zipToBlob } from '../services/zip';
import { TagBuilder, escapeXml } from '../services/xml';
import { fromBase64 } from '../services/encoding';

// ─── 단위 변환 ────────────────────────────────────────────────────────────────

const PT_TO_HWP = 100;       // 1pt → 100 HWPUNIT (font size용)
const PT_TO_HWPU = 7200 / 72; // 1pt → 100 HWPUNIT (레이아웃용, = 100)

function pt2HwpLayout(pt: number): number {
  return Math.round(pt * PT_TO_HWPU);
}

function pt2HwpFont(pt: number): number {
  return Math.round(pt * PT_TO_HWP);
}

// ─── ID 생성기 ────────────────────────────────────────────────────────────────

class IdCounter {
  private counts: Record<string, number> = {};
  next(key: string): number {
    this.counts[key] = (this.counts[key] ?? -1) + 1;
    return this.counts[key];
  }
  current(key: string): number {
    return this.counts[key] ?? 0;
  }
}

// ─── Emit Context ─────────────────────────────────────────────────────────────

interface HwpxEmitContext {
  ids: IdCounter;
  charShapes: string[];     // CharShape XML 목록
  paraShapes: string[];     // ParaShape XML 목록
  images: { name: string; data: Uint8Array }[];  // 이미지 바이너리
  insideCell: boolean;
}

function makeContext(): HwpxEmitContext {
  return {
    ids: new IdCounter(),
    charShapes: [],
    paraShapes: [],
    images: [],
    insideCell: false,
  };
}

// ─── CharShape 등록 ───────────────────────────────────────────────────────────

function registerCharShape(style: RunStyle, ctx: HwpxEmitContext): number {
  const id = ctx.ids.next('cs');
  const sizePt = style.fontSizePt ?? 10;
  const colorStr = style.colorHex ? style.colorHex.padStart(6, '0') : '000000';
  const xml = `<hh:charShape id="${id}"`
    + ` height="${pt2HwpFont(sizePt)}"`
    + ` bold="${style.bold ? '1' : '0'}"`
    + ` italic="${style.italic ? '1' : '0'}"`
    + ` underline="${style.underline ? 'SINGLE' : 'NONE'}"`
    + ` strikeout="${style.strikethrough ? 'SINGLE' : 'NONE'}"`
    + ` superScript="${style.superscript ? '1' : '0'}"`
    + ` subScript="${style.subscript ? '1' : '0'}"`
    + ` textColor="#${colorStr}"`
    + `/>\n`;
  ctx.charShapes.push(xml);
  return id;
}

function registerParaShape(style: ParagraphStyle, ctx: HwpxEmitContext): number {
  const id = ctx.ids.next('ps');
  const alignMap: Record<string, string> = {
    left: 'LEFT', center: 'CENTER', right: 'RIGHT', justify: 'JUSTIFY',
  };
  const align = alignMap[style.align ?? 'left'] ?? 'LEFT';
  const xml = `<hh:paraShape id="${id}"`
    + ` align="${align}"`
    + ` indentLeft="${pt2HwpLayout(style.indentLeftPt ?? 0)}"`
    + ` indentRight="${pt2HwpLayout(style.indentRightPt ?? 0)}"`
    + ` spacingBefore="${pt2HwpLayout(style.spacingBeforePt ?? 0)}"`
    + ` spacingAfter="${pt2HwpLayout(style.spacingAfterPt ?? 0)}"`
    + ` lineSpacing="${Math.round((style.lineSpacing ?? 1.0) * 100)}"`
    + `/>\n`;
  ctx.paraShapes.push(xml);
  return id;
}

// ─── HwpxEmitVisitor ─────────────────────────────────────────────────────────

class HwpxEmitVisitor implements NodeVisitor<string, HwpxEmitContext> {

  visitDocument(_node: IrDocumentNode, _ctx: HwpxEmitContext): string {
    // document는 buildPackage에서 처리
    return '';
  }

  visitSection(node: IrSectionNode, ctx: HwpxEmitContext): string {
    const layout = node.layout;
    const pagepr = `<hp:pagepr`
      + ` width="${pt2HwpLayout(layout.widthPt)}"`
      + ` height="${pt2HwpLayout(layout.heightPt)}"`
      + ` headerLen="${pt2HwpLayout(layout.marginTopPt)}"`
      + ` footerLen="${pt2HwpLayout(layout.marginBottomPt)}"`
      + ` leftMargin="${pt2HwpLayout(layout.marginLeftPt)}"`
      + ` rightMargin="${pt2HwpLayout(layout.marginRightPt)}"`
      + `/>`;

    const blocks = node.blocks.map(block => dispatchVisit(this, block, ctx)).join('\n');
    return `<hh:body>\n${pagepr}\n${blocks}\n</hh:body>`;
  }

  visitParagraph(node: IrParagraphNode, ctx: HwpxEmitContext): string {
    const psId = registerParaShape(node.style, ctx);

    // 제목 레벨이 있으면 스타일 id 적용
    const styleAttr = node.style.headingLevel
      ? ` styleId="heading${node.style.headingLevel}"`
      : '';

    const pPr = `<hp:pPr paraShapeId="${psId}"${styleAttr}/>`;
    const content = node.children.map(child => dispatchVisit(this, child, ctx)).join('');

    return `<hp:p>\n${pPr}\n${content}\n</hp:p>`;
  }

  visitRun(node: IrRunNode, ctx: HwpxEmitContext): string {
    const csId = registerCharShape(node.style, ctx);
    const runPr = `<hp:runPr charShapeId="${csId}"/>`;
    const content = node.children.map(child => dispatchVisit(this, child, ctx)).join('');
    return `<hp:run>\n${runPr}\n${content}\n</hp:run>`;
  }

  visitText(node: IrTextNode, _ctx: HwpxEmitContext): string {
    return `<hp:t>${escapeXml(node.value)}</hp:t>`;
  }

  visitImage(node: IrImageNode, ctx: HwpxEmitContext): string {
    const imgId = ctx.ids.next('img');
    const ext = node.mimeType.split('/')[1] ?? 'png';
    const fileName = `BIN${String(imgId).padStart(4, '0')}.${ext}`;
    ctx.images.push({ name: fileName, data: fromBase64(node.dataBase64) });

    const w = pt2HwpLayout(node.widthPx * 0.75); // px → pt 근사 변환
    const h = pt2HwpLayout(node.heightPx * 0.75);

    return `<hp:ctrl id="GENSM">`
      + `<hp:pic binaryItemIDRef="${imgId}">`
      + `<hp:sz width="${w}" height="${h}"/>`
      + `</hp:pic>`
      + `</hp:ctrl>`;
  }

  visitHyperlink(node: IrHyperlinkNode, ctx: HwpxEmitContext): string {
    const runs = node.children.map(run => this.visitRun(run, ctx)).join('');
    return `<hp:ctrl id="FIELD">`
      + `<hp:fieldBegin type="HYPERLINK" url="${escapeXml(node.url)}"/>`
      + runs
      + `<hp:fieldEnd/>`
      + `</hp:ctrl>`;
  }

  visitTable(node: IrTableNode, ctx: HwpxEmitContext): string {
    const rows = node.rows.map(row => this.visitTableRow(row, ctx)).join('\n');
    const colCount = node.rows[0]?.cells.length ?? 1;
    return `<hp:tbl numCols="${colCount}" numRows="${node.rows.length}">\n${rows}\n</hp:tbl>`;
  }

  visitTableRow(node: IrTableRowNode, ctx: HwpxEmitContext): string {
    const cells = node.cells.map(cell => this.visitTableCell(cell, ctx)).join('\n');
    return `<hp:tr>\n${cells}\n</hp:tr>`;
  }

  visitTableCell(node: IrTableCellNode, ctx: HwpxEmitContext): string {
    const cellCtx: HwpxEmitContext = { ...ctx, insideCell: true };
    const tcPr = `<hp:tcPr colSpan="${node.colSpan}" rowSpan="${node.rowSpan}"`
      + (node.style.bgColorHex ? ` fillColor="#${node.style.bgColorHex}"` : '')
      + `/>`;
    const content = node.children.map(p => this.visitParagraph(p, cellCtx)).join('\n');
    return `<hp:tc>\n${tcPr}\n${content}\n</hp:tc>`;
  }

  visitLineBreak(_node: IrLineBreakNode, _ctx: HwpxEmitContext): string {
    return `<hp:lf/>`;
  }

  visitPageBreak(_node: IrPageBreakNode, _ctx: HwpxEmitContext): string {
    return `</hp:p><hp:p><hp:pPr pageBreak="1"/></hp:p><hp:p>`;
  }
}

// ─── HWPX 패키지 XML 생성 ─────────────────────────────────────────────────────

function buildHeaderXml(meta: DocumentMeta, ctx: HwpxEmitContext): string {
  const charShapesXml = ctx.charShapes.join('');
  const paraShapesXml = ctx.paraShapes.join('');

  return `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2012/head"
         xmlns:hp="http://www.hancom.co.kr/hwpml/2012/paragraph">
  <hh:docInfo>
    <hh:title>${escapeXml(meta.title ?? '')}</hh:title>
    <hh:creator>${escapeXml(meta.author ?? '')}</hh:creator>
  </hh:docInfo>
  <hh:charShapes>
${charShapesXml}
  </hh:charShapes>
  <hh:paraShapes>
${paraShapesXml}
  </hh:paraShapes>
  <hh:styles>
    <hh:style id="0" type="para" name="Normal" paraShapeId="0" charShapeId="0"/>
    <hh:style id="1" type="para" name="Heading 1" paraShapeId="0" charShapeId="0"/>
    <hh:style id="2" type="para" name="Heading 2" paraShapeId="0" charShapeId="0"/>
    <hh:style id="3" type="para" name="Heading 3" paraShapeId="0" charShapeId="0"/>
  </hh:styles>
</hh:head>`;
}

function buildSectionXml(body: string, index: number): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<hh:sec xmlns:hh="http://www.hancom.co.kr/hwpml/2012/head"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2012/paragraph"
        number="${index}">
${body}
</hh:sec>`;
}

function buildManifestXml(sectionCount: number, imageNames: string[]): string {
  const sectionItems = Array.from({ length: sectionCount }, (_, i) =>
    `<manifest:item full-path="Contents/section${i}.xml" media-type="application/xml"/>`,
  ).join('\n');

  const imageItems = imageNames.map(name =>
    `<manifest:item full-path="Contents/BinData/${name}" media-type="image/png"/>`,
  ).join('\n');

  return `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:item full-path="/" media-type="application/hwp+zip"/>
  <manifest:item full-path="Contents/header.xml" media-type="application/xml"/>
${sectionItems}
${imageItems}
</manifest:manifest>`;
}

function buildVersionXml(): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<hh:version xmlns:hh="http://www.hancom.co.kr/hwpml/2012/head"
  appVersion="9.9.9.9" major="5" minor="0" micro="2" buildNumber="0"/>`;
}

// ─── HwpxWriter ───────────────────────────────────────────────────────────────

export class HwpxWriter implements WriteStrategy<Blob> {
  async write(document: IrDocumentNode): Promise<Blob> {
    const visitor = new HwpxEmitVisitor();
    const ctx = makeContext();

    const sectionBodies: string[] = [];

    for (const section of document.sections) {
      const bodyXml = visitor.visitSection(section, ctx);
      sectionBodies.push(bodyXml);
    }

    const headerXml = buildHeaderXml(document.meta, ctx);

    const entries: ZipWriteEntry[] = [
      { path: 'META-INF/manifest.xml', data: buildManifestXml(sectionBodies.length, ctx.images.map(i => i.name)) },
      { path: 'version.xml', data: buildVersionXml() },
      { path: 'Contents/header.xml', data: headerXml },
    ];

    // 섹션 파일들
    for (let i = 0; i < sectionBodies.length; i++) {
      entries.push({
        path: `Contents/section${i}.xml`,
        data: buildSectionXml(sectionBodies[i], i),
      });
    }

    // 이미지 파일들
    for (const img of ctx.images) {
      entries.push({ path: `Contents/BinData/${img.name}`, data: img.data });
    }

    const zipData = await buildZip(entries);
    return zipToBlob(zipData);
  }
}
