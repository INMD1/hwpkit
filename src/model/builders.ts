import type {
  DocRoot, SheetNode, ParaNode, SpanNode, ImgNode,
  GridNode, RowNode, CellNode, ContentNode, TxtNode, PageNumNode, BrNode, PbNode,
} from './doc-tree';
import type { TextProps, ParaProps, CellProps, GridProps, DocMeta, PageDims, ImgLayout } from './doc-props';
import { A4 } from './doc-props';

export function buildRoot(meta: DocMeta = {}, kids: SheetNode[] = []): DocRoot {
  return { tag: 'root', meta, kids };
}

export function buildSheet(
  kids: ContentNode[] = [],
  dims: PageDims = A4,
  opts?: { headers?: SheetNode["headers"]; footers?: SheetNode["footers"] },
): SheetNode {
  const node: SheetNode = { tag: 'sheet', dims, kids };
  if (opts?.headers) node.headers = opts.headers;
  if (opts?.footers) node.footers = opts.footers;
  return node;
}

export function buildPageNum(format?: PageNumNode['format']): PageNumNode {
  return { tag: 'pagenum', format };
}

export function buildBr(): BrNode { return { tag: 'br' }; }
export function buildPb(): PbNode { return { tag: 'pb' }; }

export function buildPara(kids: ParaNode['kids'] = [], props: ParaProps = {}): ParaNode {
  return { tag: 'para', props, kids };
}

export function buildSpan(content: string, props: TextProps = {}): SpanNode {
  const txt: TxtNode = { tag: 'txt', content };
  return { tag: 'span', props, kids: [txt] };
}

export function buildImg(
  b64: string,
  mime: ImgNode['mime'],
  w: number,
  h: number,
  alt?: string,
  layout?: ImgLayout,
): ImgNode {
  const node: ImgNode = { tag: 'img', b64, mime, w, h };
  if (alt) node.alt = alt;
  if (layout) node.layout = layout;
  return node;
}

export function buildGrid(kids: RowNode[], props: GridProps = {}): GridNode {
  return { tag: 'grid', props, kids };
}

export function buildRow(kids: CellNode[], heightPt?: number): RowNode {
  const node: RowNode = { tag: 'row', kids };
  if (heightPt != null) node.heightPt = heightPt;
  return node;
}

export function buildCell(
  kids: ParaNode[],
  opts: { cs?: number; rs?: number; props?: CellProps } = {},
): CellNode {
  return { tag: 'cell', cs: opts.cs ?? 1, rs: opts.rs ?? 1, props: opts.props ?? {}, kids };
}
