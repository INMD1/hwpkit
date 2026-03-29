export type Align = 'left' | 'center' | 'right' | 'justify';

// ─── 이미지 배치 ────────────────────────────────────────────
export type ImgWrap = 'inline' | 'square' | 'tight' | 'through' | 'none' | 'behind' | 'front';
export type ImgHorzAlign = 'left' | 'center' | 'right';
export type ImgVertAlign = 'top' | 'center' | 'bottom';
export type ImgHorzRelTo = 'margin' | 'column' | 'page' | 'para';
export type ImgVertRelTo = 'margin' | 'line' | 'page' | 'para';

export interface ImgLayout {
  wrap: ImgWrap;
  horzAlign?: ImgHorzAlign;   // 정렬 기준 (xPt 없을 때)
  vertAlign?: ImgVertAlign;   // 정렬 기준 (yPt 없을 때)
  horzRelTo?: ImgHorzRelTo;   // 가로 기준점
  vertRelTo?: ImgVertRelTo;   // 세로 기준점
  xPt?: number;               // 명시적 가로 오프셋 (pt)
  yPt?: number;               // 명시적 세로 오프셋 (pt)
  distT?: number;             // 텍스트와의 거리 top (pt)
  distB?: number;
  distL?: number;
  distR?: number;
  behindDoc?: boolean;        // 텍스트 뒤에 배치
  zOrder?: number;
}
export type VAlign = 'top' | 'mid' | 'bot';
export type Heading = 1 | 2 | 3 | 4 | 5 | 6;
export type StrokeKind = 'solid' | 'dash' | 'dot' | 'double' | 'none';

export interface TextProps {
  b?: boolean;
  i?: boolean;
  u?: boolean;
  s?: boolean;
  sup?: boolean;
  sub?: boolean;
  font?: string;
  pt?: number;
  color?: string;
  bg?: string;
}

export interface ParaProps {
  align?: Align;
  heading?: Heading;
  indentPt?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  lineHeight?: number;
  listLv?: number;
  listOrd?: boolean;
  listMark?: string;
}

export interface Stroke {
  kind: StrokeKind;
  pt: number;
  color: string;
}

export interface CellProps {
  top?: Stroke;
  bot?: Stroke;
  left?: Stroke;
  right?: Stroke;
  bg?: string;
  padPt?: number;
  align?: Align;
  va?: VAlign;
  isHeader?: boolean;
}

export interface TableLook {
  firstRow?: boolean;
  lastRow?: boolean;
  firstCol?: boolean;
  lastCol?: boolean;
  bandedRows?: boolean;
  bandedCols?: boolean;
}

export interface GridProps {
  widthPct?: number;
  colWidths?: number[];   // column widths in points
  defaultStroke?: Stroke;
  look?: TableLook;
  headerRow?: boolean;
}

export interface PageDims {
  wPt: number;
  hPt: number;
  mt: number;
  mb: number;
  ml: number;
  mr: number;
  orient?: 'portrait' | 'landscape';
}

export interface DocMeta {
  title?: string;
  author?: string;
  subject?: string;
  desc?: string;
  keywords?: string;
  created?: string;
  modified?: string;
}

export const A4: PageDims = {
  wPt: 595.28, hPt: 841.89,
  mt: 56.69, mb: 56.69, ml: 70.87, mr: 70.87,
  orient: 'portrait',
};

export const A4_LANDSCAPE: PageDims = {
  wPt: 841.89, hPt: 595.28,
  mt: 56.69, mb: 56.69, ml: 70.87, mr: 70.87,
  orient: 'landscape',
};

/**
 * orient === 'landscape'일 때 wPt < hPt이면 swap,
 * orient === 'portrait'일 때 wPt > hPt이면 swap하여
 * 방향과 치수가 항상 일치하도록 정규화합니다.
 */

export function normalizeDims(dims: PageDims): PageDims {
  const orient = dims.orient ?? 'portrait';
  if (orient === 'landscape' && dims.wPt < dims.hPt) {
    return { ...dims, wPt: dims.hPt, hPt: dims.wPt };
  }
  if (orient === 'portrait' && dims.wPt > dims.hPt) {
    return { ...dims, wPt: dims.hPt, hPt: dims.wPt };
  }
  return dims;
}

export const DEFAULT_STROKE: Stroke = { kind: 'solid', pt: 0.5, color: '000000' };
