import type { DocRoot, ParaNode, SpanNode, GridNode, ContentNode, ImgNode } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import type { Stroke } from '../../model/doc-props';
import { succeed, fail } from '../../contract/result';
import { TextKit } from '../../toolkit/TextKit';
import { registry } from '../../pipeline/registry';
import { BaseEncoder } from '../../core/BaseEncoder';

export class MdEncoder extends BaseEncoder {
  protected getFormat(): string { return 'md'; }

  async encode(doc: DocRoot): Promise<Outcome<Uint8Array>> {
    try {
      const warns: string[] = [];
      const parts: string[] = [];
      for (const sheet of doc.kids) {
        // Warn about header/footer loss
        if (sheet.headers && sheet.headers.default && sheet.headers.default.length > 0) warns.push('[SHIELD] MD: 머리글(header) 표현 불가 — 손실됨');
        if (sheet.footers && sheet.footers.default && sheet.footers.default.length > 0) warns.push('[SHIELD] MD: 바닥글(footer) 표현 불가 — 손실됨');

        for (const kid of sheet.kids) parts.push(encodeContent(kid, warns));
      }
      return succeed(this.stringToBytes(parts.join('\n\n')), warns);
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

/** pt → CSS border shorthand (only if stroke is visible) */
function strokeToCss(s?: Stroke): string | undefined {
  if (!s || s.kind === 'none' || s.pt <= 0) return undefined;
  const kindMap: Record<string, string> = { solid: 'solid', dash: 'dashed', dot: 'dotted', double: 'double', none: 'none' };
  const style = kindMap[s.kind] ?? 'solid';
  const px = Math.max(1, Math.round(s.pt * 96 / 72));
  const color = s.color.startsWith('#') ? s.color : `#${s.color}`;
  return `${px}px ${style} ${color}`;
}

function encodeGrid(grid: GridNode, warns: string[]): string {
  if (grid.kids.length === 0) return '';

  // HTML 테이블로 출력 — 테두리/배경색을 인라인 스타일로 유지
  const rowCount = grid.kids.length;

  // Build occupancy map for rowspan
  const occupancy: Set<number>[] = Array.from({ length: rowCount }, () => new Set());
  let colCount = 0;
  for (let ri = 0; ri < rowCount; ri++) {
    const row = grid.kids[ri];
    let ci = 0;
    for (const cell of row.kids) {
      while (occupancy[ri].has(ci)) ci++;
      if (cell.rs > 1) {
        for (let r = ri + 1; r < ri + cell.rs && r < rowCount; r++) {
          for (let c = ci; c < ci + cell.cs; c++) occupancy[r].add(c);
        }
      }
      ci += cell.cs;
    }
    while (occupancy[ri].has(ci)) ci++;
    if (ci > colCount) colCount = ci;
  }

  let rows = '';
  for (let ri = 0; ri < rowCount; ri++) {
    const row = grid.kids[ri];
    let cells = '';
    let colIdx = 0;

    for (const cell of row.kids) {
      while (occupancy[ri].has(colIdx)) colIdx++;

      const cs = cell.cs > 1 ? ` colspan="${cell.cs}"` : '';
      const rs = cell.rs > 1 ? ` rowspan="${cell.rs}"` : '';

      const styles: string[] = ['padding:4px 6px', 'vertical-align:top'];
      const top    = strokeToCss(cell.props.top);
      const bot    = strokeToCss(cell.props.bot);
      const left   = strokeToCss(cell.props.left);
      const right  = strokeToCss(cell.props.right);
      if (top)   styles.push(`border-top:${top}`);
      if (bot)   styles.push(`border-bottom:${bot}`);
      if (left)  styles.push(`border-left:${left}`);
      if (right) styles.push(`border-right:${right}`);
      if (cell.props.bg) styles.push(`background-color:#${cell.props.bg}`);
      if (cell.props.va === 'mid') styles[1] = 'vertical-align:middle';
      else if (cell.props.va === 'bot') styles[1] = 'vertical-align:bottom';

      const tag = (grid.props.headerRow && ri === 0) || cell.props.isHeader ? 'th' : 'td';
      const content = cell.kids.map(p => encodePara(p, warns)).join('\n');
      cells += `<${tag}${cs}${rs} style="${styles.join(';')}">${content}</${tag}>`;
      colIdx += cell.cs;
    }
    rows += `<tr>${cells}</tr>\n`;
  }

  return `<table style="border-collapse:collapse;width:100%">\n<tbody>\n${rows}</tbody>\n</table>\n`;
}

registry.registerEncoder(new MdEncoder());
