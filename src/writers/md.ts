/**
 * MdEmitVisitor - IR → Markdown 문자열
 *
 * WriteStrategy 구현체이며 NodeVisitor<string, MdEmitContext>를 구현합니다.
 * IR 트리를 방문(visit)하며 Markdown 텍스트를 조립합니다.
 *
 * 기존 코드와의 차이:
 *  - if-else 체인 없음: Visitor 디스패치
 *  - 출력 버퍼 없음: 각 visit 메서드가 문자열 반환
 *  - 섹션 분리 없음: 모든 섹션을 --- 구분자로 연결
 */

import type { WriteStrategy } from '../core/strategy';
import type { NodeVisitor } from '../core/visitor';
import { dispatchVisit } from '../core/visitor';
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
} from '../core/ir';

// ─── Emit Context ─────────────────────────────────────────────────────────────

interface MdEmitContext {
  listDepth: number;
  insideTable: boolean;
  insideCell: boolean;
}

function defaultCtx(): MdEmitContext {
  return { listDepth: 0, insideTable: false, insideCell: false };
}

// ─── MdEmitVisitor ────────────────────────────────────────────────────────────

class MdEmitVisitor implements NodeVisitor<string, MdEmitContext> {

  visitDocument(node: IrDocumentNode, ctx: MdEmitContext): string {
    const parts: string[] = [];

    // YAML frontmatter (메타데이터)
    if (node.meta.title || node.meta.author) {
      const lines = ['---'];
      if (node.meta.title) lines.push(`title: "${node.meta.title}"`);
      if (node.meta.author) lines.push(`author: "${node.meta.author}"`);
      if (node.meta.subject) lines.push(`subject: "${node.meta.subject}"`);
      lines.push('---');
      parts.push(lines.join('\n'));
    }

    for (const section of node.sections) {
      parts.push(this.visitSection(section, ctx));
    }

    return parts.join('\n\n');
  }

  visitSection(node: IrSectionNode, ctx: MdEmitContext): string {
    const blocks: string[] = [];
    for (const block of node.blocks) {
      const text = dispatchVisit(this, block, ctx);
      if (text.trim()) blocks.push(text);
    }
    return blocks.join('\n\n');
  }

  visitParagraph(node: IrParagraphNode, ctx: MdEmitContext): string {
    // 자식 노드 텍스트 수집
    const inline = node.children
      .map(child => dispatchVisit(this, child, ctx))
      .join('');

    if (!inline.trim()) return '';

    const { headingLevel, listLevel, listOrdered, listMarker } = node.style;

    // Heading
    if (headingLevel) {
      return `${'#'.repeat(headingLevel)} ${inline.trim()}`;
    }

    // List item
    if (listLevel !== undefined) {
      const indent = '  '.repeat(listLevel);
      const marker = listOrdered ? `${listLevel + 1}.` : (listMarker ?? '-');
      return `${indent}${marker} ${inline.trim()}`;
    }

    // 셀 내부는 줄바꿈 없이
    if (ctx.insideCell) return inline.replace(/\n/g, ' ').trim();

    return inline;
  }

  visitRun(node: IrRunNode, ctx: MdEmitContext): string {
    const inner = node.children
      .map(child => dispatchVisit(this, child, ctx))
      .join('');

    if (!inner) return '';

    // 서식 적용 (안쪽 → 바깥쪽 순서로 wrap)
    let result = inner;
    const { bold, italic, underline, strikethrough, superscript, subscript } = node.style;

    if (strikethrough) result = `~~${result}~~`;
    if (underline)     result = `<u>${result}</u>`;
    if (bold && italic) result = `***${result}***`;
    else if (bold)     result = `**${result}**`;
    else if (italic)   result = `*${result}*`;
    if (superscript)   result = `<sup>${result}</sup>`;
    if (subscript)     result = `<sub>${result}</sub>`;

    return result;
  }

  visitText(node: IrTextNode, _ctx: MdEmitContext): string {
    return node.value;
  }

  visitImage(node: IrImageNode, ctx: MdEmitContext): string {
    const alt = node.altText ?? 'image';
    const src = `data:${node.mimeType};base64,${node.dataBase64}`;
    if (ctx.insideCell) return `![${alt}](${src})`;
    return `![${alt}](${src})`;
  }

  visitHyperlink(node: IrHyperlinkNode, ctx: MdEmitContext): string {
    const text = node.children
      .map(run => this.visitRun(run, ctx))
      .join('');
    return `[${text}](${node.url})`;
  }

  visitTable(node: IrTableNode, ctx: MdEmitContext): string {
    const tableCtx: MdEmitContext = { ...ctx, insideTable: true };
    const rows = node.rows.map(row => this.visitTableRow(row, tableCtx));

    if (rows.length === 0) return '';

    // GFM 테이블: 첫 번째 행을 헤더로, 두 번째 행에 구분선 삽입
    const [header, ...body] = rows;
    const cols = node.rows[0]?.cells.length ?? 1;
    const separator = `| ${Array(cols).fill('---').join(' | ')} |`;

    return [header, separator, ...body].join('\n');
  }

  visitTableRow(node: IrTableRowNode, ctx: MdEmitContext): string {
    const cellCtx: MdEmitContext = { ...ctx, insideCell: true };
    const cells = node.cells.map(cell => this.visitTableCell(cell, cellCtx));
    return `| ${cells.join(' | ')} |`;
  }

  visitTableCell(node: IrTableCellNode, ctx: MdEmitContext): string {
    if (node.rowSpan === 0) return '';  // 병합된 셀
    const content = node.children
      .map(p => this.visitParagraph(p, ctx))
      .filter(Boolean)
      .join('<br>');
    return content || ' ';
  }

  visitLineBreak(_node: IrLineBreakNode, ctx: MdEmitContext): string {
    return ctx.insideCell ? '<br>' : '\n';
  }

  visitPageBreak(_node: IrPageBreakNode, _ctx: MdEmitContext): string {
    return '\n\n---\n\n';  // Markdown 수평선으로 쪽 나누기 표현
  }
}

// ─── MdWriter (WriteStrategy 구현) ───────────────────────────────────────────

export class MdWriter implements WriteStrategy<string> {
  private visitor = new MdEmitVisitor();

  async write(document: IrDocumentNode): Promise<string> {
    const result = this.visitor.visitDocument(document, defaultCtx());
    // 연속된 빈 줄 정리
    return result.replace(/\n{3,}/g, '\n\n').trim();
  }
}
