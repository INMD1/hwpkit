import type { TextProps, ParaProps, CellProps, GridProps, PageDims, DocMeta, ImgLayout } from './doc-props';

// ─── 노드 종류 ─────────────────────────────────────────────
export type BlockTag =
  | 'root' | 'sheet' | 'para' | 'span'
  | 'txt' | 'img' | 'link'
  | 'grid' | 'row' | 'cell'
  | 'br' | 'pb' | 'pagenum';

// ─── 리프 노드 ─────────────────────────────────────────────
export interface TxtNode   { tag: 'txt'; content: string }
export interface BrNode    { tag: 'br' }
export interface PbNode    { tag: 'pb' }

export interface PageNumNode {
  tag: 'pagenum';
  format?: 'decimal' | 'roman' | 'romanCaps';
}

export interface ImgNode {
  tag: 'img';
  b64: string;
  mime: 'image/png' | 'image/jpeg' | 'image/gif' | 'image/bmp';
  w: number;
  h: number;
  alt?: string;
  layout?: ImgLayout;  // 배치/위치 정보 (없으면 inline으로 취급)
}

// ─── 인라인 노드 ───────────────────────────────────────────
export interface SpanNode {
  tag: 'span';
  props: TextProps;
  kids: (TxtNode | BrNode | PbNode | PageNumNode)[];
}

export interface LinkNode {
  tag: 'link';
  href: string;
  kids: SpanNode[];
}

// ─── 블록 노드 ─────────────────────────────────────────────
export interface ParaNode {
  tag: 'para';
  props: ParaProps;
  kids: (SpanNode | ImgNode | LinkNode | GridNode)[];
}

// ─── 표(Grid) 노드 ─────────────────────────────────────────
export interface CellNode {
  tag: 'cell';
  cs: number;
  rs: number;
  props: CellProps;
  kids: ParaNode[];
}

export interface RowNode  { tag: 'row'; kids: CellNode[]; heightPt?: number }

export interface GridNode {
  tag: 'grid';
  props: GridProps;
  kids: RowNode[];
}

// ─── 섹션(Sheet) 노드 ──────────────────────────────────────
export type ContentNode = ParaNode | GridNode;

export interface SheetNode {
  tag: 'sheet';
  dims: PageDims;
  kids: ContentNode[];
  headers?: {
    default?: ParaNode[];
    first?: ParaNode[];
    even?: ParaNode[];
  };
  footers?: {
    default?: ParaNode[];
    first?: ParaNode[];
    even?: ParaNode[];
  };
}

// ─── 루트 노드 ─────────────────────────────────────────────
export interface DocRoot {
  tag: 'root';
  meta: DocMeta;
  kids: SheetNode[];
}

// ─── 모든 노드의 유니온 ─────────────────────────────────────
export type AnyNode =
  | DocRoot | SheetNode | ParaNode | SpanNode
  | TxtNode | ImgNode | LinkNode
  | GridNode | RowNode | CellNode
  | BrNode | PbNode | PageNumNode;
