/**
 * HWPKit - 공개 API 진입점
 *
 * 변환 흐름:
 *   입력 파일 → ReadStrategy (Reader) → IrDocumentNode → WriteStrategy (Writer) → 출력
 *
 * ┌─────────────────────────────────────────────────────┐
 * │  Transformer (공개 API)                              │
 * │  ├── MdTransformer.fromHwpx(file) → string         │
 * │  ├── MdTransformer.fromHwp(file)  → string         │
 * │  ├── MdTransformer.fromDocx(file) → string         │
 * │  ├── HwpxTransformer.fromDocx(file) → Blob         │
 * │  ├── HwpxTransformer.fromMd(md)   → Blob           │
 * │  ├── DocxTransformer.fromHwpx(file) → Blob         │
 * │  ├── DocxTransformer.fromHwp(file)  → Blob         │
 * │  └── DocxTransformer.fromMd(md)     → Blob         │
 * └─────────────────────────────────────────────────────┘
 */

// ─── 공개 Transformer API ─────────────────────────────────────────────────────
export { MdTransformer } from './transformers/MdTransformer';
export { HwpxTransformer } from './transformers/HwpxTransformer';
export { DocxTransformer } from './transformers/DocxTransformer';

// ─── 직접 Reader/Writer 사용 (고급) ──────────────────────────────────────────
export { HwpxReader } from './readers/hwpx';
export { HwpReader } from './readers/hwp';
export { DocxReader } from './readers/docx';
export { MarkdownReader } from './readers/markdown';
export { MdWriter } from './writers/md';
export { HwpxWriter } from './writers/hwpx';
export { DocxWriter } from './writers/docx';

// ─── IR 타입 ──────────────────────────────────────────────────────────────────
export type {
  IrDocumentNode,
  IrSectionNode,
  IrBlockNode,
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
  TableStyle,
  CellStyle,
  CellBorder,
  TextAlign,
  HeadingLevel,
} from './core/ir';
export {
  makeDocument,
  makeSection,
  makeParagraph,
  makeRun,
  makeImage,
  makeTable,
  makeTableRow,
  makeTableCell,
  DEFAULT_PAGE_LAYOUT,
} from './core/ir';

// ─── Strategy 인터페이스 (커스텀 Reader/Writer 작성 시) ─────────────────────
export type { ReadStrategy, WriteStrategy } from './core/strategy';

// ─── Visitor 인터페이스 (커스텀 IR 순회 시) ──────────────────────────────────
export type { NodeVisitor } from './core/visitor';
export { BaseVisitor, dispatchVisit } from './core/visitor';
