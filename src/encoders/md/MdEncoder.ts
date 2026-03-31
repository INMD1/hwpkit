import type { Encoder } from '../../contract/encoder';
import type { DocRoot, ParaNode, SpanNode, GridNode, ContentNode, ImgNode } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import { succeed, fail } from '../../contract/result';
import { TextKit } from '../../toolkit/TextKit';
import { registry } from '../../pipeline/registry';

export class MdEncoder implements Encoder {
  readonly format = 'md';

  async encode(doc: DocRoot): Promise<Outcome<Uint8Array>> {
    try {
      const warns: string[] = [];
      const parts: string[] = [];
      for (const sheet of doc.kids) {
        // Warn about header/footer loss
        if (sheet.header && sheet.header.length > 0) warns.push('[SHIELD] MD: 머리글(header) 표현 불가 — 손실됨');
        if (sheet.footer && sheet.footer.length > 0) warns.push('[SHIELD] MD: 바닥글(footer) 표현 불가 — 손실됨');

        for (const kid of sheet.kids) parts.push(encodeContent(kid, warns));
      }
      return succeed(TextKit.encode(parts.join('\n\n')), warns);
    } catch (e: any) {
      return fail(`MD encode error: ${e?.message ?? String(e)}`);
    }
  }
}

function encodeContent(node: ContentNode, warns: string[]): string {
  return node.tag === 'grid' ? encodeGrid(node, warns) : encodePara(node, warns);
}

function encodePara(para: ParaNode, warns: string[]): string {
  const text = para.kids.map(k => {
    if (k.tag === 'span') return encodeSpan(k, warns);
    if (k.tag === 'img') return encodeImage(k);
    return '';
  }).join('');

  if (para.props.heading) return `${'#'.repeat(para.props.heading)} ${text}`;

  if (para.props.listOrd !== undefined) {
    const indent = '  '.repeat(para.props.listLv ?? 0);
    return `${indent}${para.props.listOrd ? '1.' : '-'} ${text}`;
  }

  // Alignment: use HTML fallback for non-left
  if (para.props.align && para.props.align !== 'left' && para.props.align !== 'justify') {
    return `<div align="${para.props.align}">${text}</div>`;
  }

  return text;
}

function encodeSpan(span: SpanNode, warns: string[]): string {
  let hasPageNum = false;
  const textParts: string[] = [];
  for (const kid of span.kids) {
    if (kid.tag === 'txt') textParts.push(kid.content);
    else if (kid.tag === 'pagenum') {
      hasPageNum = true;
      warns.push('[SHIELD] MD: 페이지 번호 표현 불가 — 손실됨');
    }
  }

  let r = textParts.join('');
  if (hasPageNum && r === '') r = '[페이지 번호]';

  // Collect CSS styles for font/color/size/bg — use HTML span so fonts can be
  // loaded externally via the page's stylesheet or @font-face rules.
  const cssStyles: string[] = [];
  if (span.props.font) cssStyles.push(`font-family: ${span.props.font}`);
  if (span.props.pt) cssStyles.push(`font-size: ${span.props.pt}pt`);
  if (span.props.color) cssStyles.push(`color: #${span.props.color}`);
  if (span.props.bg) cssStyles.push(`background-color: #${span.props.bg}`);

  const hasHtmlStyle = cssStyles.length > 0;

  if (hasHtmlStyle) {
    // When style properties are present, use HTML for all formatting so that
    // markdown markers inside an HTML element don't break parsers.
    if (span.props.b) cssStyles.push('font-weight: bold');
    if (span.props.i) cssStyles.push('font-style: italic');
    if (span.props.s) cssStyles.push('text-decoration: line-through');
    if (span.props.u) {
      // combine underline with possible line-through
      const existing = cssStyles.find(s => s.startsWith('text-decoration:'));
      if (existing) {
        const idx = cssStyles.indexOf(existing);
        cssStyles[idx] = existing.replace('line-through', 'underline line-through');
        if (!existing.includes('line-through')) cssStyles[idx] = existing + ' underline';
      } else {
        cssStyles.push('text-decoration: underline');
      }
    }
    const styleAttr = cssStyles.join('; ');
    if (span.props.sup) return `<sup style="${styleAttr}">${r}</sup>`;
    if (span.props.sub) return `<sub style="${styleAttr}">${r}</sub>`;
    return `<span style="${styleAttr}">${r}</span>`;
  }

  // No CSS styles needed — use plain Markdown formatting
  if (span.props.b && span.props.i) r = `***${r}***`;
  else if (span.props.b) r = `**${r}**`;
  else if (span.props.i) r = `*${r}*`;
  if (span.props.s) r = `~~${r}~~`;
  if (span.props.u) r = `<u>${r}</u>`;
  if (span.props.sup) r = `<sup>${r}</sup>`;
  if (span.props.sub) r = `<sub>${r}</sub>`;

  return r;
}

function encodeImage(img: ImgNode): string {
  return `![${img.alt ?? ''}](data:${img.mime};base64,${img.b64})`;
}

function encodeGrid(grid: GridNode, warns: string[]): string {
  if (grid.kids.length === 0) return '';

  // Warn about table style loss
  if (grid.props.look) warns.push('[SHIELD] MD: 표 스타일(색상, 테두리, 머리행 강조) 표현 불가 — 손실됨');

  const rows = grid.kids.map(row =>
    `| ${row.kids.map(cell => cell.kids.map(p => encodePara(p, warns)).join(' ')).join(' | ')} |`,
  );

  if (rows.length > 0) {
    const cols = grid.kids[0].kids.length;
    rows.splice(1, 0, `| ${Array(cols).fill('---').join(' | ')} |`);
  }

  return rows.join('\n');
}

registry.registerEncoder(new MdEncoder());
