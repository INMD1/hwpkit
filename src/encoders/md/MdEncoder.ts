import type { Encoder } from '../../contract/encoder';
import type { DocRoot, ParaNode, SpanNode, GridNode, ContentNode } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import { succeed, fail } from '../../contract/result';
import { TextKit } from '../../toolkit/TextKit';
import { registry } from '../../pipeline/registry';

export class MdEncoder implements Encoder {
  readonly format = 'md';

  async encode(doc: DocRoot): Promise<Outcome<Uint8Array>> {
    try {
      const parts: string[] = [];
      for (const sheet of doc.kids) {
        for (const kid of sheet.kids) parts.push(encodeContent(kid));
      }
      return succeed(TextKit.encode(parts.join('\n\n')));
    } catch (e: any) {
      return fail(`MD encode error: ${e?.message ?? String(e)}`);
    }
  }
}

function encodeContent(node: ContentNode): string {
  return node.tag === 'grid' ? encodeGrid(node) : encodePara(node);
}

function encodePara(para: ParaNode): string {
  const text = para.kids.map(k => k.tag === 'span' ? encodeSpan(k) : '').join('');

  if (para.props.heading) return `${'#'.repeat(para.props.heading)} ${text}`;

  if (para.props.listOrd !== undefined) {
    const indent = '  '.repeat(para.props.listLv ?? 0);
    return `${indent}${para.props.listOrd ? '1.' : '-'} ${text}`;
  }

  return text;
}

function encodeSpan(span: SpanNode): string {
  const text = span.kids.filter(k => k.tag === 'txt').map(k => k.tag === 'txt' ? k.content : '').join('');
  let r = text;
  if (span.props.b && span.props.i) r = `***${r}***`;
  else if (span.props.b) r = `**${r}**`;
  else if (span.props.i) r = `*${r}*`;
  if (span.props.s) r = `~~${r}~~`;
  return r;
}

function encodeGrid(grid: GridNode): string {
  if (grid.kids.length === 0) return '';

  const rows = grid.kids.map(row =>
    `| ${row.kids.map(cell => cell.kids.map(p => encodePara(p)).join(' ')).join(' | ')} |`,
  );

  if (rows.length > 0) {
    const cols = grid.kids[0].kids.length;
    rows.splice(1, 0, `| ${Array(cols).fill('---').join(' | ')} |`);
  }

  return rows.join('\n');
}

registry.registerEncoder(new MdEncoder());
