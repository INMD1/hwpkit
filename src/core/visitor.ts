/**
 * Visitor 패턴
 *
 * IR 트리를 순회하는 방식을 외부에서 주입합니다.
 * 각 Writer(MdEmitVisitor, HwpxEmitVisitor 등)는 이 인터페이스를 구현합니다.
 *
 * R: 반환 타입, C: 컨텍스트 타입
 */

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
} from './ir';

// ─── Visitor Interface ────────────────────────────────────────────────────────

export interface NodeVisitor<R = void, C = unknown> {
  visitDocument(node: IrDocumentNode, ctx: C): R;
  visitSection(node: IrSectionNode, ctx: C): R;
  visitParagraph(node: IrParagraphNode, ctx: C): R;
  visitRun(node: IrRunNode, ctx: C): R;
  visitText(node: IrTextNode, ctx: C): R;
  visitImage(node: IrImageNode, ctx: C): R;
  visitHyperlink(node: IrHyperlinkNode, ctx: C): R;
  visitTable(node: IrTableNode, ctx: C): R;
  visitTableRow(node: IrTableRowNode, ctx: C): R;
  visitTableCell(node: IrTableCellNode, ctx: C): R;
  visitLineBreak(node: IrLineBreakNode, ctx: C): R;
  visitPageBreak(node: IrPageBreakNode, ctx: C): R;
}

// ─── Dispatch Helper ──────────────────────────────────────────────────────────

export function dispatchVisit<R, C>(
  visitor: NodeVisitor<R, C>,
  node: { kind: string },
  ctx: C,
): R {
  switch (node.kind) {
    case 'document':   return visitor.visitDocument(node as IrDocumentNode, ctx);
    case 'section':    return visitor.visitSection(node as IrSectionNode, ctx);
    case 'paragraph':  return visitor.visitParagraph(node as IrParagraphNode, ctx);
    case 'run':        return visitor.visitRun(node as IrRunNode, ctx);
    case 'text':       return visitor.visitText(node as IrTextNode, ctx);
    case 'image':      return visitor.visitImage(node as IrImageNode, ctx);
    case 'hyperlink':  return visitor.visitHyperlink(node as IrHyperlinkNode, ctx);
    case 'table':      return visitor.visitTable(node as IrTableNode, ctx);
    case 'tableRow':   return visitor.visitTableRow(node as IrTableRowNode, ctx);
    case 'tableCell':  return visitor.visitTableCell(node as IrTableCellNode, ctx);
    case 'lineBreak':  return visitor.visitLineBreak(node as IrLineBreakNode, ctx);
    case 'pageBreak':  return visitor.visitPageBreak(node as IrPageBreakNode, ctx);
    default:
      throw new Error(`[HWPKit] 알 수 없는 IR 노드: ${(node as { kind: string }).kind}`);
  }
}

// ─── BaseVisitor (기본 구현 - 자식 노드를 재귀 순회) ──────────────────────────

export abstract class BaseVisitor<R, C = unknown> implements NodeVisitor<R, C> {
  abstract visitText(node: IrTextNode, ctx: C): R;
  abstract visitImage(node: IrImageNode, ctx: C): R;
  abstract visitLineBreak(node: IrLineBreakNode, ctx: C): R;
  abstract visitPageBreak(node: IrPageBreakNode, ctx: C): R;

  visitDocument(node: IrDocumentNode, ctx: C): R {
    let last!: R;
    for (const section of node.sections) last = this.visitSection(section, ctx);
    return last;
  }

  visitSection(node: IrSectionNode, ctx: C): R {
    let last!: R;
    for (const block of node.blocks) last = dispatchVisit(this, block, ctx);
    return last;
  }

  visitParagraph(node: IrParagraphNode, ctx: C): R {
    let last!: R;
    for (const child of node.children) last = dispatchVisit(this, child, ctx);
    return last;
  }

  visitRun(node: IrRunNode, ctx: C): R {
    let last!: R;
    for (const child of node.children) last = dispatchVisit(this, child, ctx);
    return last;
  }

  visitHyperlink(node: IrHyperlinkNode, ctx: C): R {
    let last!: R;
    for (const run of node.children) last = this.visitRun(run, ctx);
    return last;
  }

  visitTable(node: IrTableNode, ctx: C): R {
    let last!: R;
    for (const row of node.rows) last = this.visitTableRow(row, ctx);
    return last;
  }

  visitTableRow(node: IrTableRowNode, ctx: C): R {
    let last!: R;
    for (const cell of node.cells) last = this.visitTableCell(cell, ctx);
    return last;
  }

  visitTableCell(node: IrTableCellNode, ctx: C): R {
    let last!: R;
    for (const para of node.children) last = this.visitParagraph(para, ctx);
    return last;
  }
}
