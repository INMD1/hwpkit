/**
 * DocxReader - DOCX → IrDocumentNode
 *
 * Strategy: ReadStrategy 구현체
 * DOCX는 OOXML 표준의 ZIP 컨테이너입니다.
 *
 * 데이터 흐름:
 *   Uint8Array (DOCX ZIP)
 *     → ZipAdapter.extractZip()         (fflate 기반 ZIP 추출)
 *     → RelationshipIndex               (이미지/스타일 관계 인덱싱)
 *     → StyleIndex                      (DOCX 스타일 파싱)
 *     → DocxBlockAssembler              (document.xml → IR Block 조립)
 *     → IrDocumentNode
 */

import type { ReadStrategy } from '../core/strategy';
import type {
  IrDocumentNode,
  IrSectionNode,
  IrBlockNode,
  IrParagraphNode,
  IrRunNode,
  IrTableNode,
  IrTableRowNode,
  IrTableCellNode,
  RunStyle,
  ParagraphStyle,
  CellStyle,
  CellBorder,
  PageLayout,
  DocumentMeta,
  HeadingLevel,
  TextAlign,
} from '../core/ir';
import {
  makeDocument,
  makeSection,
  makeParagraph,
  makeRun,
  makeImage,
  makeTable,
  makeTableRow,
  makeTableCell,
  DEFAULT_PAGE_LAYOUT,
} from '../core/ir';
import {
  extractZip,
  getEntry,
  requireEntry,
  entryToText,
  ZipEntries,
} from '../services/zip';
import {
  parseXml,
  asArray,
  asString,
  attr,
  numAttr,
  XmlObject,
} from '../services/xml';
import { toBase64, detectMimeType } from '../services/encoding';

// ─── OOXML 배열 태그 ──────────────────────────────────────────────────────────

const DOCX_ARRAY_TAGS = [
  'w:p', 'w:r', 'w:tbl', 'w:tr', 'w:tc',
  'w:rPr', 'w:pPr', 'w:trPr', 'w:tcPr',
  'w:style', 'w:font',
  'Relationship', 'Override',
];

// ─── 관계(Relationship) 인덱스 ───────────────────────────────────────────────

type RelType = 'image' | 'hyperlink' | 'other';

interface Relationship {
  id: string;
  type: RelType;
  target: string;
}

function parseRels(relsXml: string): Map<string, Relationship> {
  const obj = parseXml(relsXml, { arrayTags: ['Relationship'] }) as XmlObject;
  const rels = asArray<XmlObject>((obj['Relationships'] as XmlObject | undefined)?.['Relationship']);
  const map = new Map<string, Relationship>();

  for (const rel of rels) {
    const id = attr(rel, 'Id');
    const typeUri = attr(rel, 'Type');
    const target = attr(rel, 'Target');
    let type: RelType = 'other';
    if (typeUri.includes('/image')) type = 'image';
    else if (typeUri.includes('/hyperlink')) type = 'hyperlink';
    map.set(id, { id, type, target });
  }

  return map;
}

// ─── 스타일 인덱스 ────────────────────────────────────────────────────────────

interface DocxStyle {
  headingLevel?: HeadingLevel;
  bold?: boolean;
  italic?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  align?: TextAlign;
}

function parseStyles(stylesXml: string): Map<string, DocxStyle> {
  const obj = parseXml(stylesXml, { arrayTags: ['w:style'] }) as XmlObject;
  const styles = asArray<XmlObject>((obj['w:styles'] as XmlObject | undefined)?.['w:style']);
  const map = new Map<string, DocxStyle>();

  for (const style of styles) {
    const styleId = attr(style, 'w:styleId');
    const name = asString(((style['w:name'] as XmlObject | undefined)?.['@_w:val'])).toLowerCase();

    const docxStyle: DocxStyle = {};

    // Heading 감지
    const headMatch = name.match(/^heading\s*(\d)/i);
    if (headMatch) docxStyle.headingLevel = Number(headMatch[1]) as HeadingLevel;

    // rPr에서 기본 서식 추출
    const rPr = (style['w:rPr'] ?? {}) as XmlObject;
    if (rPr['w:b'] !== undefined) docxStyle.bold = true;
    if (rPr['w:i'] !== undefined) docxStyle.italic = true;
    const szHp = numAttr(rPr['w:sz'] as XmlObject ?? {}, 'w:val');
    if (szHp > 0) docxStyle.fontSizePt = szHp / 2;

    // pPr에서 정렬 추출
    const pPr = (style['w:pPr'] ?? {}) as XmlObject;
    const jc = attr(pPr['w:jc'] as XmlObject ?? {}, 'w:val');
    if (jc) docxStyle.align = normalizeAlign(jc);

    map.set(styleId, docxStyle);
  }

  return map;
}

function normalizeAlign(raw: string): TextAlign {
  switch (raw.toLowerCase()) {
    case 'center': return 'center';
    case 'right': return 'right';
    case 'both':
    case 'distribute':
    case 'justify': return 'justify';
    default: return 'left';
  }
}

// ─── 페이지 레이아웃 ──────────────────────────────────────────────────────────

function parsePageSetup(docXml: XmlObject): PageLayout {
  // w:body > w:sectPr > w:pgSz, w:pgMar
  const body = (docXml['w:document'] as XmlObject | undefined)?.['w:body'] as XmlObject | undefined;
  if (!body) return { ...DEFAULT_PAGE_LAYOUT };

  const sectPr = (body['w:sectPr'] ?? {}) as XmlObject;
  const pgSz = (sectPr['w:pgSz'] ?? {}) as XmlObject;
  const pgMar = (sectPr['w:pgMar'] ?? {}) as XmlObject;

  const EMU_PT = 12700; // 1pt = 12700 EMU
  const twip2Pt = (t: number) => t / 20; // 1pt = 20 twips

  const wTwip = numAttr(pgSz, 'w:w');
  const hTwip = numAttr(pgSz, 'w:h');

  return {
    widthPt: wTwip > 0 ? twip2Pt(wTwip) : DEFAULT_PAGE_LAYOUT.widthPt,
    heightPt: hTwip > 0 ? twip2Pt(hTwip) : DEFAULT_PAGE_LAYOUT.heightPt,
    marginTopPt: twip2Pt(numAttr(pgMar, 'w:top') || 1134),
    marginBottomPt: twip2Pt(numAttr(pgMar, 'w:bottom') || 1134),
    marginLeftPt: twip2Pt(numAttr(pgMar, 'w:left') || 1417),
    marginRightPt: twip2Pt(numAttr(pgMar, 'w:right') || 1417),
    orientation: wTwip > hTwip ? 'landscape' : 'portrait',
  };
}

// ─── Run 파싱 ─────────────────────────────────────────────────────────────────

function parseDocxRun(
  rNode: XmlObject,
  styleMap: Map<string, DocxStyle>,
  relsMap: Map<string, Relationship>,
  entries: ZipEntries,
): IrRunNode {
  const rPr = (rNode['w:rPr'] ?? {}) as XmlObject;

  // 서식 파싱
  const bold = rPr['w:b'] !== undefined || rPr['w:bCs'] !== undefined;
  const italic = rPr['w:i'] !== undefined || rPr['w:iCs'] !== undefined;
  const underline = rPr['w:u'] !== undefined && attr(rPr['w:u'] as XmlObject, 'w:val') !== 'none';
  const strike = rPr['w:strike'] !== undefined || rPr['w:dstrike'] !== undefined;
  const vertAlign = attr(rPr['w:vertAlign'] as XmlObject ?? {}, 'w:val');
  const superscript = vertAlign === 'superscript';
  const subscript = vertAlign === 'subscript';

  const szHp = numAttr(rPr['w:sz'] as XmlObject ?? {}, 'w:val');
  const fontSizePt = szHp > 0 ? szHp / 2 : undefined;

  const fontFamily =
    asString((rPr['w:rFonts'] as XmlObject | undefined)?.['@_w:ascii'])
    || asString((rPr['w:rFonts'] as XmlObject | undefined)?.['@_w:hAnsi'])
    || undefined;

  const colorVal = attr(rPr['w:color'] as XmlObject ?? {}, 'w:val');
  const colorHex = colorVal && colorVal !== 'auto' ? colorVal : undefined;

  const style: RunStyle = {
    bold: bold || undefined,
    italic: italic || undefined,
    underline: underline || undefined,
    strikethrough: strike || undefined,
    superscript: superscript || undefined,
    subscript: subscript || undefined,
    fontSizePt,
    fontFamily,
    colorHex,
  };

  const children: IrRunNode['children'] = [];

  // 텍스트
  const tNodes = asArray<XmlObject | string>(rNode['w:t']);
  for (const t of tNodes) {
    const val = typeof t === 'string' ? t : asString((t as XmlObject)['#text']);
    if (val) children.push({ kind: 'text', value: val });
  }

  // 줄바꿈
  if (rNode['w:br'] !== undefined) {
    const brType = attr(rNode['w:br'] as XmlObject ?? {}, 'w:type');
    if (brType === 'page') children.push({ kind: 'pageBreak' });
    else children.push({ kind: 'lineBreak' });
  }

  return { kind: 'run', style, children };
}

// ─── Paragraph 파싱 ───────────────────────────────────────────────────────────

function parseDocxParagraph(
  pNode: XmlObject,
  styleMap: Map<string, DocxStyle>,
  relsMap: Map<string, Relationship>,
  entries: ZipEntries,
): IrParagraphNode {
  const pPr = (pNode['w:pPr'] ?? {}) as XmlObject;
  const styleId = attr(pPr['w:pStyle'] as XmlObject ?? {}, 'w:val');
  const docxStyle = styleMap.get(styleId) ?? {};

  // 정렬
  const jcVal = attr(pPr['w:jc'] as XmlObject ?? {}, 'w:val');
  const align = jcVal ? normalizeAlign(jcVal) : (docxStyle.align ?? 'left');

  // 들여쓰기
  const ind = (pPr['w:ind'] ?? {}) as XmlObject;
  const indLeftTwip = numAttr(ind, 'w:left') || numAttr(ind, 'w:start');
  const indRightTwip = numAttr(ind, 'w:right') || numAttr(ind, 'w:end');
  const twip2Pt = (t: number) => t / 20;

  // 간격
  const spacing = (pPr['w:spacing'] ?? {}) as XmlObject;
  const beforeTwip = numAttr(spacing, 'w:before');
  const afterTwip = numAttr(spacing, 'w:after');
  const lineRule = attr(spacing, 'w:lineRule');
  const lineVal = numAttr(spacing, 'w:line');
  let lineSpacing: number | undefined;
  if (lineVal > 0) {
    lineSpacing = lineRule === 'auto' ? lineVal / 240 : lineVal / 240;
  }

  // 목록
  const numPr = (pPr['w:numPr'] ?? {}) as XmlObject;
  const numId = numAttr(numPr['w:numId'] as XmlObject ?? {}, 'w:val');
  const ilvl = numAttr(numPr['w:ilvl'] as XmlObject ?? {}, 'w:val');

  const style: ParagraphStyle = {
    align: align !== 'left' ? align : undefined,
    headingLevel: docxStyle.headingLevel,
    indentLeftPt: indLeftTwip > 0 ? twip2Pt(indLeftTwip) : undefined,
    indentRightPt: indRightTwip > 0 ? twip2Pt(indRightTwip) : undefined,
    spacingBeforePt: beforeTwip > 0 ? twip2Pt(beforeTwip) : undefined,
    spacingAfterPt: afterTwip > 0 ? twip2Pt(afterTwip) : undefined,
    lineSpacing,
    listLevel: numId > 0 ? ilvl : undefined,
    listOrdered: numId > 0 ? false : undefined,
  };

  const children: IrParagraphNode['children'] = [];

  // 하이퍼링크
  const hyperlinks = asArray<XmlObject>(pNode['w:hyperlink']);
  for (const hl of hyperlinks) {
    const relId = attr(hl, 'r:id');
    const rel = relsMap.get(relId);
    const url = rel?.target ?? attr(hl, 'w:anchor');
    if (url) {
      const hlRuns = asArray<XmlObject>(hl['w:r']).map(r =>
        parseDocxRun(r, styleMap, relsMap, entries),
      );
      if (hlRuns.length > 0) {
        children.push({ kind: 'hyperlink', url, children: hlRuns });
      }
    }
  }

  // 일반 Run
  for (const rNode of asArray<XmlObject>(pNode['w:r'])) {
    // 이미지 Run (w:drawing 또는 w:pict)
    const drawing = rNode['w:drawing'] as XmlObject | undefined;
    const pic = drawing
      ? ((drawing['wp:inline'] ?? drawing['wp:anchor']) as XmlObject | undefined)
      : undefined;

    if (pic) {
      const blip = ((pic['a:graphic'] as XmlObject | undefined)
        ?.['a:graphicData'] as XmlObject | undefined)
        ?.['pic:pic'] as XmlObject | undefined;
      const relId = attr(
        (blip?.['pic:blipFill'] as XmlObject | undefined)?.['a:blip'] as XmlObject ?? {},
        'r:embed',
      );
      const rel = relsMap.get(relId);
      if (rel?.type === 'image') {
        const imgPath = `word/${rel.target.replace(/^\//, '')}`;
        const imgData = getEntry(entries, imgPath);
        if (imgData) {
          const mime = detectMimeType(imgData);
          const b64 = toBase64(imgData);
          const ext = (pic['wp:extent'] ?? {}) as XmlObject;
          const cx = numAttr(ext, 'cx');
          const cy = numAttr(ext, 'cy');
          const EMU_PX = 9525;
          children.push(makeImage(b64, mime, Math.round(cx / EMU_PX), Math.round(cy / EMU_PX)));
          continue;
        }
      }
    }

    const run = parseDocxRun(rNode, styleMap, relsMap, entries);
    if (run.children.length > 0) children.push(run);
  }

  return makeParagraph(style, children);
}

// ─── Table 파싱 ───────────────────────────────────────────────────────────────

function parseDocxTable(
  tblNode: XmlObject,
  styleMap: Map<string, DocxStyle>,
  relsMap: Map<string, Relationship>,
  entries: ZipEntries,
): IrTableNode {
  const rows: IrTableRowNode[] = [];

  for (const trNode of asArray<XmlObject>(tblNode['w:tr'])) {
    const cells: IrTableCellNode[] = [];

    for (const tcNode of asArray<XmlObject>(trNode['w:tc'])) {
      const tcPr = (tcNode['w:tcPr'] ?? {}) as XmlObject;
      const gridSpan = numAttr(tcPr['w:gridSpan'] as XmlObject ?? {}, 'w:val') || 1;
      const vMerge = tcPr['w:vMerge'];
      const rowSpan = vMerge !== undefined ? 0 : 1; // 0 = 병합된 셀 (추후 후처리)

      const shading = (tcPr['w:shd'] ?? {}) as XmlObject;
      const bgHex = attr(shading, 'w:fill');

      const cellStyle: CellStyle = {
        bgColorHex: bgHex && bgHex !== 'auto' ? bgHex : undefined,
      };

      const cellParas = asArray<XmlObject>(tcNode['w:p']).map(p =>
        parseDocxParagraph(p, styleMap, relsMap, entries),
      );

      cells.push(makeTableCell(cellParas, { colSpan: gridSpan, rowSpan, style: cellStyle }));
    }

    rows.push(makeTableRow(cells));
  }

  return makeTable(rows);
}

// ─── DocxReader ───────────────────────────────────────────────────────────────

export class DocxReader implements ReadStrategy {
  async read(data: Uint8Array): Promise<IrDocumentNode> {
    const entries = await extractZip(data);

    // 관계 파싱
    const relsXmlData = getEntry(entries, 'word/_rels/document.xml.rels');
    const relsMap = relsXmlData ? parseRels(entryToText(relsXmlData)) : new Map<string, Relationship>();

    // 스타일 파싱
    const stylesData = getEntry(entries, 'word/styles.xml');
    const styleMap = stylesData ? parseStyles(entryToText(stylesData)) : new Map<string, DocxStyle>();

    // 메타데이터
    const meta = parseDocxMeta(entries);

    // document.xml 파싱
    const docXmlText = entryToText(requireEntry(entries, 'word/document.xml'));
    const docXmlObj = parseXml(docXmlText, { arrayTags: DOCX_ARRAY_TAGS }) as XmlObject;

    const layout = parsePageSetup(docXmlObj);

    const body = ((docXmlObj['w:document'] as XmlObject | undefined)?.['w:body'] ?? {}) as XmlObject;
    const blocks: IrBlockNode[] = [];

    for (const pNode of asArray<XmlObject>(body['w:p'])) {
      blocks.push(parseDocxParagraph(pNode, styleMap, relsMap, entries));
    }

    for (const tblNode of asArray<XmlObject>(body['w:tbl'])) {
      blocks.push(parseDocxTable(tblNode, styleMap, relsMap, entries));
    }

    const section = makeSection(layout, blocks);
    return makeDocument(meta, [section]);
  }
}

function parseDocxMeta(entries: ZipEntries): DocumentMeta {
  const meta: DocumentMeta = {};
  const coreData = getEntry(entries, 'docProps/core.xml');
  if (!coreData) return meta;

  const obj = parseXml(entryToText(coreData)) as XmlObject;
  const cp = (obj['cp:coreProperties'] ?? obj['coreProperties'] ?? {}) as XmlObject;

  meta.title = asString(cp['dc:title']) || undefined;
  meta.author = asString(cp['dc:creator']) || undefined;
  meta.subject = asString(cp['dc:subject']) || undefined;
  meta.description = asString(cp['dc:description']) || undefined;
  meta.createdAt = asString(cp['dcterms:created']) || undefined;
  meta.modifiedAt = asString(cp['dcterms:modified']) || undefined;

  return meta;
}
