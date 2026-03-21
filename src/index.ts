// ─── 공개 API ───────────────────────────────────────────────

// Pipeline
export { Pipeline } from './pipeline/Pipeline';
export { registry } from './pipeline/registry';

// Side-effect imports: register all decoders/encoders
import './decoders/md/MdDecoder';
import './decoders/hwpx/HwpxDecoder';
import './decoders/docx/DocxDecoder';
import './decoders/hwp/HwpScanner';
import './encoders/md/MdEncoder';
import './encoders/hwpx/HwpxEncoder';
import './encoders/docx/DocxEncoder';

// Model
export type {
  DocRoot, SheetNode, ParaNode, SpanNode, GridNode, RowNode, CellNode,
  ImgNode, LinkNode, TxtNode, BrNode, PbNode, PageNumNode, ContentNode, AnyNode, BlockTag,
} from './model/doc-tree';
export type {
  TextProps, ParaProps, CellProps, GridProps, TableLook, PageDims, DocMeta,
  Align, VAlign, Heading, StrokeKind, Stroke,
} from './model/doc-props';
export { A4, DEFAULT_STROKE } from './model/doc-props';
export { buildRoot, buildSheet, buildPara, buildSpan, buildImg, buildGrid, buildRow, buildCell, buildPageNum } from './model/builders';

// Contract
export type { Decoder } from './contract/decoder';
export type { Encoder } from './contract/encoder';
export type { Outcome, Ok, Fail } from './contract/result';
export { succeed, fail } from './contract/result';

// Safety
export { ShieldedParser } from './safety/ShieldedParser';
export { Metric, safeHex, safeAlign, safeFont, safeStrokeHwpx, safeStrokeDocx } from './safety/StyleBridge';

// Walk
export { TreeWalker, walkNode } from './walk/TreeWalker';
export { countNodes, validateRoot } from './walk/tree-ops';

// Toolkit
export { XmlKit } from './toolkit/XmlKit';
export { ArchiveKit } from './toolkit/ArchiveKit';
export { BinaryKit } from './toolkit/BinaryKit';
export { TextKit } from './toolkit/TextKit';
