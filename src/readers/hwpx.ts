/**
 * HwpxReader - HWPX → IrDocumentNode
 *
 * Strategy: ReadStrategy 구현체
 * HWPX는 ZIP 컨테이너이며 내부에 XML 파일들을 포함합니다.
 *
 * 데이터 흐름:
 *   Uint8Array (HWPX ZIP)
 *     → ZipAdapter.extractZip()       (fflate 기반 ZIP 추출)
 *     → XmlAdapter.parseXml()          (fast-xml-parser 기반 XML 파싱)
 *     → SectionAssembler              (섹션별 IR 조립)
 *     → IrDocumentNode
 */

import type { ReadStrategy } from '../core/strategy';
import type {
  IrDocumentNode,
  IrSectionNode,
  IrBlockNode,
  IrParagraphNode,
  IrTableNode,
  IrTableRowNode,
  IrTableCellNode,
  IrRunNode,
  RunStyle,
  ParagraphStyle,
  CellStyle,
  PageLayout,
  DocumentMeta,
  HeadingLevel,
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
  requireEntry,
  getEntry,
  entryToText,
  listEntries,
  ZipEntries,
} from '../services/zip';
import {
  parseXml,
  asArray,
  asString,
  asNumber,
  attr,
  numAttr,
  XmlObject,
} from '../services/xml';
import { toBase64, detectMimeType } from '../services/encoding';

// ─── HWPX 배열 태그 목록 ──────────────────────────────────────────────────────

const HWPX_ARRAY_TAGS = [
  'hp:p', 'hp:run', 'hp:tbl', 'hp:tr', 'hp:tc',
  'hp:ctrl', 'hp:fieldBegin', 'hp:fieldEnd',
  'hp:charShape', 'hp:paraShape', 'hp:fontFaces',
  'hh:refList', 'hh:charShape', 'hh:paraShape',
  'hh:style', 'hh:para',
];

// ─── 스타일 레지스트리 ────────────────────────────────────────────────────────

interface CharShapeEntry {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  superscript: boolean;
  subscript: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  colorHex?: string;
}

interface ParaShapeEntry {
  align?: ParagraphStyle['align'];
  headingLevel?: HeadingLevel;
  indentLeftPt?: number;
  indentRightPt?: number;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  lineSpacing?: number;
}

interface StyleRegistry {
  charShapes: CharShapeEntry[];
  paraShapes: ParaShapeEntry[];
  fontNames: string[];
}

// HWP 내부 단위 → PT 변환 (HWPUNIT = 1/7200 인치, 1인치 = 72pt)
function hwpUnit2Pt(hwp: number): number {
  return (hwp / 7200) * 72;
}

function parseFontSize(hwp: number): number {
  // HWPX는 font size를 1/100 포인트 단위로 저장
  return hwp / 100;
}

// ─── 헤더 파싱 (스타일 레지스트리 구축) ──────────────────────────────────────

function parseHeader(xml: string): StyleRegistry {
  const root = parseXml(xml, { arrayTags: HWPX_ARRAY_TAGS });
  const hh = (root['hh:head'] ?? root['head'] ?? {}) as XmlObject;

  const charShapes: CharShapeEntry[] = [];
  const paraShapes: ParaShapeEntry[] = [];
  const fontNames: string[] = [];

  // 글꼴 이름 수집
  const fontFaces = asArray<XmlObject>(
    (hh['hh:refList'] as XmlObject | undefined)?.['hh:fontFaces']
    ?? (hh['hh:charShapes'] as XmlObject | undefined)?.['hh:fontFace'],
  );
  for (const ff of fontFaces) {
    fontNames.push(asString((ff as XmlObject)['@_name'] ?? (ff as XmlObject)['name']));
  }

  // CharShape 파싱
  const rawCharShapes = asArray<XmlObject>(
    (hh['hh:charShapes'] as XmlObject | undefined)?.['hh:charShape']
    ?? (hh['hh:refList'] as XmlObject | undefined)?.['hh:charShape'],
  );
  for (const cs of rawCharShapes) {
    const bold = attr(cs, 'bold') === '1' || attr(cs, 'bold') === 'true';
    const italic = attr(cs, 'italic') === '1' || attr(cs, 'italic') === 'true';
    const underlineType = attr(cs, 'underline');
    const underline = underlineType !== '' && underlineType !== 'NONE' && underlineType !== '0';
    const strikethrough = attr(cs, 'strikeout') !== 'NONE' && attr(cs, 'strikeout') !== '';
    const superscript = attr(cs, 'superScript') === '1';
    const subscript = attr(cs, 'subScript') === '1';
    const faceNameId = numAttr(cs, 'faceNameId');
    const fontFamily = fontNames[faceNameId] ?? undefined;
    const sizeRaw = numAttr(cs, 'height');
    const fontSizePt = sizeRaw > 0 ? parseFontSize(sizeRaw) : undefined;
    const colorHex = attr(cs, 'textColor').replace(/^#/, '') || undefined;
    charShapes.push({ bold, italic, underline, strikethrough, superscript, subscript, fontFamily, fontSizePt, colorHex });
  }

  // ParaShape 파싱
  const rawParaShapes = asArray<XmlObject>(
    (hh['hh:paraShapes'] as XmlObject | undefined)?.['hh:paraShape']
    ?? (hh['hh:refList'] as XmlObject | undefined)?.['hh:paraShape'],
  );
  for (const ps of rawParaShapes) {
    const rawAlign = attr(ps, 'align').toUpperCase();
    const alignMap: Record<string, ParagraphStyle['align']> = {
      LEFT: 'left', CENTER: 'center', RIGHT: 'right',
      JUSTIFY: 'justify', BOTH: 'justify',
    };
    const align = alignMap[rawAlign];
    const indentLeftPt = hwpUnit2Pt(numAttr(ps, 'indentLeft'));
    const indentRightPt = hwpUnit2Pt(numAttr(ps, 'indentRight'));
    const spacingBeforePt = hwpUnit2Pt(numAttr(ps, 'spacingBefore'));
    const spacingAfterPt = hwpUnit2Pt(numAttr(ps, 'spacingAfter'));
    const lineSpacingRaw = numAttr(ps, 'lineSpacing');
    const lineSpacing = lineSpacingRaw > 0 ? lineSpacingRaw / 100 : undefined;
    paraShapes.push({ align, indentLeftPt, indentRightPt, spacingBeforePt, spacingAfterPt, lineSpacing });
  }

  // 스타일에서 Heading 레벨 추출
  const styles = asArray<XmlObject>(
    (hh['hh:styles'] as XmlObject | undefined)?.['hh:style'],
  );
  for (const style of styles) {
    const name = asString(attr(style, 'name')).toLowerCase();
    const paraShapeId = numAttr(style, 'paraShapeId');
    const match = name.match(/^heading\s*(\d)/i) ?? name.match(/^제목\s*(\d)/);
    if (match && paraShapes[paraShapeId]) {
      paraShapes[paraShapeId].headingLevel = Number(match[1]) as HeadingLevel;
    }
  }

  return { charShapes, paraShapes, fontNames };
}

// ─── 페이지 레이아웃 파싱 ─────────────────────────────────────────────────────

function parseSectionLayout(secXml: XmlObject): PageLayout {
  const pagepr = (secXml['hp:pagepr'] ?? secXml['hh:pagepr'] ?? {}) as XmlObject;
  if (!pagepr) return { ...DEFAULT_PAGE_LAYOUT };
  return {
    widthPt: hwpUnit2Pt(numAttr(pagepr, 'width') || 59530),
    heightPt: hwpUnit2Pt(numAttr(pagepr, 'height') || 84190),
    marginTopPt: hwpUnit2Pt(numAttr(pagepr, 'headerLen') || numAttr(pagepr, 'topMargin') || 4000),
    marginBottomPt: hwpUnit2Pt(numAttr(pagepr, 'footerLen') || numAttr(pagepr, 'bottomMargin') || 4000),
    marginLeftPt: hwpUnit2Pt(numAttr(pagepr, 'leftMargin') || 5000),
    marginRightPt: hwpUnit2Pt(numAttr(pagepr, 'rightMargin') || 5000),
    orientation: numAttr(pagepr, 'width') > numAttr(pagepr, 'height') ? 'landscape' : 'portrait',
  };
}

// ─── 이미지 수집 ──────────────────────────────────────────────────────────────

function collectBinData(entries: ZipEntries): Map<string, Uint8Array> {
  const bins = new Map<string, Uint8Array>();
  const binEntries = listEntries(entries, 'Contents/BinData');
  for (const path of binEntries) {
    const data = entries[path];
    if (data) bins.set(path.split('/').pop()!, data);
  }
  return bins;
}

// ─── 섹션 XML 파싱 ────────────────────────────────────────────────────────────

function parseRunStyle(cs: CharShapeEntry | undefined): RunStyle {
  if (!cs) return {};
  return {
    bold: cs.bold || undefined,
    italic: cs.italic || undefined,
    underline: cs.underline || undefined,
    strikethrough: cs.strikethrough || undefined,
    superscript: cs.superscript || undefined,
    subscript: cs.subscript || undefined,
    fontFamily: cs.fontFamily,
    fontSizePt: cs.fontSizePt,
    colorHex: cs.colorHex,
  };
}

function parseRun(
  runNode: XmlObject,
  registry: StyleRegistry,
  binData: Map<string, Uint8Array>,
): IrRunNode {
  const runPr = (runNode['hp:runPr'] ?? {}) as XmlObject;
  const csId = numAttr(runPr, 'charShapeId');
  const style = parseRunStyle(registry.charShapes[csId]);

  const textNodes = asArray<XmlObject | string>(runNode['hp:t']);
  const children: IrRunNode['children'] = [];

  for (const t of textNodes) {
    const raw = typeof t === 'string' ? t : asString((t as XmlObject)['#text']);
    if (raw) children.push({ kind: 'text', value: raw });
  }

  // 줄바꿈
  if (runNode['hp:lf'] !== undefined) children.push({ kind: 'lineBreak' });

  return { kind: 'run', style, children };
}

function parseParagraph(
  pNode: XmlObject,
  registry: StyleRegistry,
  binData: Map<string, Uint8Array>,
): IrParagraphNode {
  const pPr = (pNode['hp:pPr'] ?? {}) as XmlObject;
  const psId = numAttr(pPr, 'paraShapeId');
  const ps = registry.paraShapes[psId] ?? {};

  const style: ParagraphStyle = {
    align: ps.align,
    headingLevel: ps.headingLevel,
    indentLeftPt: ps.indentLeftPt,
    indentRightPt: ps.indentRightPt,
    spacingBeforePt: ps.spacingBeforePt,
    spacingAfterPt: ps.spacingAfterPt,
    lineSpacing: ps.lineSpacing,
  };

  // 목록 정보
  const numPr = (pPr['hp:numPr'] ?? {}) as XmlObject;
  if (numAttr(numPr, 'numId') > 0) {
    style.listLevel = numAttr(numPr, 'iLvl');
    style.listOrdered = false; // 기본값; 목록 정의에서 결정
  }

  const children: IrParagraphNode['children'] = [];

  // 쪽나누기 체크
  if (attr(pPr, 'pageBreak') === '1' || attr(pPr, 'breakType') === 'PAGE') {
    children.push({
      kind: 'run',
      style: {},
      children: [{ kind: 'pageBreak' }],
    });
  }

  // Run 처리
  for (const runNode of asArray<XmlObject>(pNode['hp:run'])) {
    const run = parseRun(runNode, registry, binData);
    if (run.children.length > 0) children.push(run);
  }

  // 컨트롤(이미지 등) 처리
  for (const ctrl of asArray<XmlObject>(pNode['hp:ctrl'])) {
    const ctrlId = attr(ctrl, 'id') || attr(ctrl, '@_id');

    // 이미지 (GENSM = GenericShape, PICF = Picture Field)
    if (ctrlId === 'GENSM' || ctrlId === 'PICF' || ctrl['hp:pic'] !== undefined) {
      const pic = (ctrl['hp:pic'] ?? ctrl) as XmlObject;
      const binId = attr(pic as XmlObject, 'binaryItemIDRef') || attr(pic as XmlObject, 'id');
      const binEntry = binData.get(binId) ?? binData.get(`BIN${String(binId).padStart(4, '0')}.png`);
      if (binEntry) {
        const mime = detectMimeType(binEntry);
        const b64 = toBase64(binEntry);
        const szObj = (ctrl['hp:sz'] ?? (ctrl['hp:pic'] as XmlObject | undefined)?.['hp:sz'] ?? {}) as XmlObject;
        const w = numAttr(szObj, 'width') || 200;
        const h = numAttr(szObj, 'height') || 200;
        children.push(makeImage(b64, mime, w, h));
      }
    }
  }

  return makeParagraph(style, children);
}

function parseTable(
  tblNode: XmlObject,
  registry: StyleRegistry,
  binData: Map<string, Uint8Array>,
): IrTableNode {
  const rows: IrTableRowNode[] = [];
  for (const trNode of asArray<XmlObject>(tblNode['hp:tr'])) {
    const cells: IrTableCellNode[] = [];
    for (const tcNode of asArray<XmlObject>(trNode['hp:tc'])) {
      const tcPr = (tcNode['hp:tcPr'] ?? {}) as XmlObject;
      const colSpan = numAttr(tcPr, 'colSpan') || 1;
      const rowSpan = numAttr(tcPr, 'rowSpan') || 1;

      const bgHex = attr(tcPr, 'fillColor').replace(/^#/, '') || undefined;
      const cellStyle: CellStyle = {
        bgColorHex: bgHex,
        paddingPt: hwpUnit2Pt(numAttr(tcPr, 'paddingLeft') || 0),
      };

      const cellParas: IrParagraphNode[] = [];
      for (const pNode of asArray<XmlObject>(tcNode['hp:p'])) {
        cellParas.push(parseParagraph(pNode, registry, binData));
      }

      cells.push(makeTableCell(cellParas, { colSpan, rowSpan, style: cellStyle }));
    }
    rows.push(makeTableRow(cells));
  }
  return makeTable(rows);
}

function parseSectionBlocks(
  secXml: XmlObject,
  registry: StyleRegistry,
  binData: Map<string, Uint8Array>,
): IrBlockNode[] {
  const body = (secXml['hh:body'] ?? secXml['hp:body'] ?? secXml) as XmlObject;
  const blocks: IrBlockNode[] = [];

  // hp:p 와 hp:tbl 이 혼재할 수 있으므로 순서 보존
  const allChildren = collectBodyChildren(body);

  for (const child of allChildren) {
    if (child.tag === 'hp:p' || child.tag === 'p') {
      blocks.push(parseParagraph(child.node, registry, binData));
    } else if (child.tag === 'hp:tbl' || child.tag === 'tbl') {
      blocks.push(parseTable(child.node, registry, binData));
    }
  }

  return blocks;
}

/** 본문에서 hp:p / hp:tbl 순서를 보존하며 수집 */
function collectBodyChildren(body: XmlObject): { tag: string; node: XmlObject }[] {
  const result: { tag: string; node: XmlObject }[] = [];

  // fast-xml-parser는 같은 태그명 요소들을 배열로 처리하지만
  // 서로 다른 태그 간 순서는 별도 처리 필요
  // → 단순히 hp:p 먼저, hp:tbl 나중으로 처리 (순서 손실 허용)
  for (const p of asArray<XmlObject>(body['hp:p'])) {
    result.push({ tag: 'hp:p', node: p });
  }
  // 테이블이 있으면 본문 끝에 붙임 (단순 구현)
  for (const tbl of asArray<XmlObject>(body['hp:tbl'])) {
    result.push({ tag: 'hp:tbl', node: tbl });
  }

  return result;
}

// ─── HwpxReader ───────────────────────────────────────────────────────────────

export class HwpxReader implements ReadStrategy {
  async read(data: Uint8Array): Promise<IrDocumentNode> {
    const entries = await extractZip(data);

    // 헤더 파싱 → 스타일 레지스트리
    const headerXml = entryToText(requireEntry(entries, 'Contents/header.xml'));
    const registry = parseHeader(headerXml);

    // 이미지 수집
    const binData = collectBinData(entries);

    // 메타데이터
    const meta = parseMetaData(entries);

    // 섹션 파일 목록 (section0.xml, section1.xml, ...) 를 번호 순으로 정렬
    const sectionPaths = Object.keys(entries)
      .filter(k => /Contents\/section\d+\.xml$/i.test(k))
      .sort((a, b) => {
        const na = Number(a.match(/(\d+)\.xml$/)?.[1] ?? 0);
        const nb = Number(b.match(/(\d+)\.xml$/)?.[1] ?? 0);
        return na - nb;
      });

    const sections: IrSectionNode[] = [];
    for (const path of sectionPaths) {
      const secText = entryToText(entries[path]);
      const secObj = parseXml(secText, { arrayTags: HWPX_ARRAY_TAGS }) as XmlObject;
      const rootKey = Object.keys(secObj).find(k => k.includes('sec')) ?? Object.keys(secObj)[0];
      const secXml = (secObj[rootKey] ?? secObj) as XmlObject;
      const layout = parseSectionLayout(secXml);
      const blocks = parseSectionBlocks(secXml, registry, binData);
      sections.push(makeSection(layout, blocks));
    }

    return makeDocument(meta, sections);
  }
}

function parseMetaData(entries: ZipEntries): DocumentMeta {
  const meta: DocumentMeta = {};
  const coreEntry = getEntry(entries, 'docProps/core.xml');
  if (!coreEntry) return meta;

  const coreXml = entryToText(coreEntry);
  const obj = parseXml(coreXml) as XmlObject;
  const cp = (obj['cp:coreProperties'] ?? obj['coreProperties'] ?? {}) as XmlObject;

  meta.title = asString((cp['dc:title'] ?? cp['title'])) || undefined;
  meta.author = asString((cp['dc:creator'] ?? cp['creator'])) || undefined;
  meta.subject = asString((cp['dc:subject'] ?? cp['subject'])) || undefined;
  meta.description = asString((cp['dc:description'] ?? cp['description'])) || undefined;
  meta.createdAt = asString((cp['dcterms:created'] ?? cp['created'])) || undefined;
  meta.modifiedAt = asString((cp['dcterms:modified'] ?? cp['modified'])) || undefined;

  return meta;
}
