/**
 * HWPKit Intermediate Representation (IR)
 *
 * 모든 문서 포맷은 이 IR로 파싱된 뒤 목표 포맷으로 직렬화됩니다.
 * 파싱 단계(ReadStrategy)와 직렬화 단계(WriteStrategy)가 완전히 분리됩니다.
 */

// ─── Node Types ──────────────────────────────────────────────────────────────

export type IrNodeKind =
  | 'document'
  | 'section'
  | 'paragraph'
  | 'run'
  | 'text'
  | 'image'
  | 'hyperlink'
  | 'table'
  | 'tableRow'
  | 'tableCell'
  | 'lineBreak'
  | 'pageBreak';

export interface IrBaseNode {
  kind: IrNodeKind;
}

// ─── Leaf Nodes ──────────────────────────────────────────────────────────────

export interface IrTextNode extends IrBaseNode {
  kind: 'text';
  value: string;
}

export interface IrLineBreakNode extends IrBaseNode {
  kind: 'lineBreak';
}

export interface IrPageBreakNode extends IrBaseNode {
  kind: 'pageBreak';
}

export interface IrImageNode extends IrBaseNode {
  kind: 'image';
  /** Base64 인코딩된 바이너리 데이터 */
  dataBase64: string;
  mimeType: 'image/png' | 'image/jpeg' | 'image/gif' | 'image/webp';
  widthPx: number;
  heightPx: number;
  altText?: string;
}

// ─── Run (인라인 서식 단위) ────────────────────────────────────────────────────

export interface RunStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  colorHex?: string;        // 예: 'FF0000'
  bgColorHex?: string;
}

export interface IrRunNode extends IrBaseNode {
  kind: 'run';
  style: RunStyle;
  children: (IrTextNode | IrLineBreakNode | IrPageBreakNode)[];
}

// ─── Hyperlink ───────────────────────────────────────────────────────────────

export interface IrHyperlinkNode extends IrBaseNode {
  kind: 'hyperlink';
  url: string;
  children: IrRunNode[];
}

// ─── Paragraph ───────────────────────────────────────────────────────────────

export type TextAlign = 'left' | 'center' | 'right' | 'justify';
export type HeadingLevel = 1 | 2 | 3 | 4 | 5 | 6;

export interface ParagraphStyle {
  align?: TextAlign;
  headingLevel?: HeadingLevel;
  indentLeftPt?: number;
  indentRightPt?: number;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  lineSpacing?: number;       // 배수 (예: 1.5)
  listLevel?: number;
  listOrdered?: boolean;
  listMarker?: string;        // 순서 없는 목록의 마커 문자
}

export interface IrParagraphNode extends IrBaseNode {
  kind: 'paragraph';
  style: ParagraphStyle;
  children: (IrRunNode | IrImageNode | IrHyperlinkNode)[];
}

// ─── Table ────────────────────────────────────────────────────────────────────

export interface CellBorder {
  widthPt: number;
  colorHex?: string;
  style?: 'solid' | 'dashed' | 'dotted' | 'none';
}

export interface CellStyle {
  borderTop?: CellBorder;
  borderBottom?: CellBorder;
  borderLeft?: CellBorder;
  borderRight?: CellBorder;
  bgColorHex?: string;
  paddingPt?: number;
  align?: TextAlign;
  vAlign?: 'top' | 'middle' | 'bottom';
}

export interface IrTableCellNode extends IrBaseNode {
  kind: 'tableCell';
  colSpan: number;
  rowSpan: number;
  style: CellStyle;
  children: IrParagraphNode[];
}

export interface IrTableRowNode extends IrBaseNode {
  kind: 'tableRow';
  cells: IrTableCellNode[];
}

export interface TableStyle {
  widthPct?: number;   // 전체 너비 대비 %
  borderAll?: CellBorder;
}

export interface IrTableNode extends IrBaseNode {
  kind: 'table';
  style: TableStyle;
  rows: IrTableRowNode[];
}

// ─── Section (페이지 레이아웃 단위) ────────────────────────────────────────────

export interface PageLayout {
  widthPt: number;
  heightPt: number;
  marginTopPt: number;
  marginBottomPt: number;
  marginLeftPt: number;
  marginRightPt: number;
  orientation?: 'portrait' | 'landscape';
}

export type IrBlockNode = IrParagraphNode | IrTableNode;

export interface IrSectionNode extends IrBaseNode {
  kind: 'section';
  layout: PageLayout;
  blocks: IrBlockNode[];
}

// ─── Document (루트) ──────────────────────────────────────────────────────────

export interface DocumentMeta {
  title?: string;
  author?: string;
  subject?: string;
  description?: string;
  keywords?: string;
  createdAt?: string;
  modifiedAt?: string;
}

export interface IrDocumentNode extends IrBaseNode {
  kind: 'document';
  meta: DocumentMeta;
  sections: IrSectionNode[];
}

// ─── Default Layout ───────────────────────────────────────────────────────────

export const DEFAULT_PAGE_LAYOUT: PageLayout = {
  widthPt: 595.28,     // A4
  heightPt: 841.89,
  marginTopPt: 56.69,
  marginBottomPt: 56.69,
  marginLeftPt: 70.87,
  marginRightPt: 70.87,
  orientation: 'portrait',
};

// ─── Builder Helpers ──────────────────────────────────────────────────────────

export function makeDocument(
  meta: DocumentMeta = {},
  sections: IrSectionNode[] = [],
): IrDocumentNode {
  return { kind: 'document', meta, sections };
}

export function makeSection(
  layout: Partial<PageLayout> = {},
  blocks: IrBlockNode[] = [],
): IrSectionNode {
  return {
    kind: 'section',
    layout: { ...DEFAULT_PAGE_LAYOUT, ...layout },
    blocks,
  };
}

export function makeParagraph(
  style: ParagraphStyle = {},
  children: IrParagraphNode['children'] = [],
): IrParagraphNode {
  return { kind: 'paragraph', style, children };
}

export function makeRun(
  text: string,
  style: RunStyle = {},
): IrRunNode {
  return {
    kind: 'run',
    style,
    children: [{ kind: 'text', value: text }],
  };
}

export function makeImage(
  dataBase64: string,
  mimeType: IrImageNode['mimeType'],
  widthPx: number,
  heightPx: number,
  altText?: string,
): IrImageNode {
  return { kind: 'image', dataBase64, mimeType, widthPx, heightPx, altText };
}

export function makeTable(
  rows: IrTableRowNode[],
  style: TableStyle = {},
): IrTableNode {
  return { kind: 'table', style, rows };
}

export function makeTableRow(cells: IrTableCellNode[]): IrTableRowNode {
  return { kind: 'tableRow', cells };
}

export function makeTableCell(
  children: IrParagraphNode[],
  options: Partial<Omit<IrTableCellNode, 'kind' | 'children'>> = {},
): IrTableCellNode {
  return {
    kind: 'tableCell',
    colSpan: options.colSpan ?? 1,
    rowSpan: options.rowSpan ?? 1,
    style: options.style ?? {},
    children,
  };
}
