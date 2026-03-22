export type Align = 'left' | 'center' | 'right' | 'justify';
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

export const DEFAULT_STROKE: Stroke = { kind: 'solid', pt: 0.5, color: '000000' };
