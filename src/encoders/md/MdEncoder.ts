import type { DocRoot, ParaNode, SpanNode, GridNode, ContentNode, ImgNode } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import type { Stroke } from '../../model/doc-props';
import type { EncoderOptions } from '../../contract/encoder';
import { succeed, fail } from '../../contract/result';
import { TextKit } from '../../toolkit/TextKit';
import { registry } from '../../pipeline/registry';
import { BaseEncoder } from '../../core/BaseEncoder';

export class MdEncoder extends BaseEncoder {
  protected getFormat(): string { return 'md'; }

  async encode(doc: DocRoot, options?: EncoderOptions): Promise<Outcome<Uint8Array>> {
    const includeImages = options?.includeImages !== false;  // default: true
    try {
      const warns: string[] = [];
      const parts: string[] = [];
      for (const sheet of doc.kids) {
        // Warn about header/footer loss
        if (sheet.headers && sheet.headers.default && sheet.headers.default.length > 0) warns.push('[SHIELD] MD: 머리글(header) 표현 불가 — 손실됨');
        if (sheet.footers && sheet.footers.default && sheet.footers.default.length > 0) warns.push('[SHIELD] MD: 바닥글(footer) 표현 불가 — 손실됨');

        for (const kid of sheet.kids) parts.push(encodeContent(kid, warns, includeImages));
      }
      return succeed(this.stringToBytes(parts.join('\n\n')), warns);
    } catch (e: any) {
      return fail(`MD encode error: ${e?.message ?? String(e)}`);
    }
  }
}

function encodeContent(node: ContentNode, warns: string[], includeImages: boolean): string {
  return node.tag === 'grid' ? encodeGrid(node, warns, includeImages) : encodePara(node, warns, includeImages);
}

function encodePara(para: ParaNode, warns: string[], includeImages: boolean): string {
  const text = para.kids.map(k => {
    if (k.tag === 'span') return encodeSpan(k, warns);
    if (k.tag === 'img') return encodeImage(k, includeImages);
    if (k.tag === 'link') {
      const label = k.kids.map(span => encodeSpan(span, warns)).join('');
      return `[${label}](${k.href})`;
    }
    if (k.tag === 'pagenum') {
      warnOnce(warns, '[SHIELD] MD: 페이지 번호 표현 불가 — 자리표시자로 대체됨');
      return '[페이지 번호]';
    }
    return '';
  }).join('');

  if (para.props.heading) return `${'#'.repeat(para.props.heading)} ${text}`;

  if (para.props.listOrd !== undefined) {
    const indent = '  '.repeat(para.props.listLv ?? 0);
    const sourceMark = para.props.listMark;
    const marker = para.props.listOrd
      ? (sourceMark && /^\d+\.$/.test(sourceMark) ? sourceMark : '1.')
      : (sourceMark && /^[-*+]$/.test(sourceMark) ? sourceMark : '-');
    return `${indent}${marker} ${text}`;
  }

  // Markdown has no paragraph-alignment syntax. Keep the content, not HTML.
  if (para.props.align && para.props.align !== 'left' && para.props.align !== 'justify') {
    warnOnce(warns, '[SHIELD] MD: 문단 정렬 표현 불가 — 정렬 정보가 손실됨');
  }

  return text;
}

function warnOnce(warns: string[], warning: string): void {
  if (!warns.includes(warning)) warns.push(warning);
}

function isCodeFont(font?: string): boolean {
  return !!font && /courier|consolas|monaco|menlo|monospace/i.test(font);
}

function wrapInlineCode(text: string): string {
  const longestRun = Math.max(
    0,
    ...(text.match(/`+/g) ?? []).map(run => run.length),
  );
  const fence = '`'.repeat(longestRun + 1);
  const pad = text.startsWith('`') || text.endsWith('`') ? ' ' : '';
  return `${fence}${pad}${text}${pad}${fence}`;
}

function encodeSpan(span: SpanNode, warns: string[]): string {
  let hasPageNum = false;
  const textParts: string[] = [];
  for (const kid of span.kids) {
    if (kid.tag === 'txt') textParts.push(kid.content);
    else if (kid.tag === 'br') textParts.push('  \n');
    else if (kid.tag === 'pb') {
      warnOnce(warns, '[SHIELD] MD: 쪽 나누기 표현 불가 — 손실됨');
    }
    else if (kid.tag === 'pagenum') {
      hasPageNum = true;
      warnOnce(warns, '[SHIELD] MD: 페이지 번호 표현 불가 — 자리표시자로 대체됨');
    }
  }

  let r = textParts.join('');
  if (hasPageNum && r === '') r = '[페이지 번호]';

  const code = isCodeFont(span.props.font);
  if (span.props.font && !code) {
    warnOnce(warns, '[SHIELD] MD: 글꼴명 표현 불가 — 글꼴 정보가 손실됨');
  }
  if (span.props.pt !== undefined) {
    warnOnce(warns, '[SHIELD] MD: 글자 크기 표현 불가 — 크기 정보가 손실됨');
  }
  if (span.props.color || span.props.bg) {
    warnOnce(warns, '[SHIELD] MD: 글자색/배경색 표현 불가 — 색상 정보가 손실됨');
  }
  if (span.props.u || span.props.sup || span.props.sub) {
    warnOnce(warns, '[SHIELD] MD: 밑줄/위첨자/아래첨자 표현 불가 — 해당 서식이 손실됨');
  }

  if (code) r = wrapInlineCode(r);
  if (span.props.b && span.props.i) r = `***${r}***`;
  else if (span.props.b) r = `**${r}**`;
  else if (span.props.i) r = `*${r}*`;
  if (span.props.s) r = `~~${r}~~`;

  return r;
}

function encodeImage(img: ImgNode, includeImages: boolean): string {
  if (!includeImages) {
    return `![${img.alt ?? ''}]`;  // alt text only, no data URI
  }
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

function encodeGrid(grid: GridNode, warns: string[], includeImages: boolean): string {
  if (grid.kids.length === 0) return '';

  if (canEncodePipeTable(grid)) {
    const losesLayout =
      Object.keys(grid.props).length > 0 ||
      grid.kids.some(row =>
        row.heightPt !== undefined ||
        row.kids.some(cell => Object.keys(cell.props).length > 0),
      );
    if (losesLayout) {
      warnOnce(
        warns,
        '[SHIELD] MD: 파이프 표가 지원하지 않는 너비/테두리/배경/정렬 정보가 손실됨',
      );
    }
    return encodePipeTable(grid, warns, includeImages);
  }

  warnOnce(
    warns,
    '[SHIELD] MD: 병합 셀 또는 셀 내부 개행/블록 요소 때문에 HTML 표로 폴백함',
  );
  return encodeHtmlTable(grid, warns, includeImages);
}

function paraHasLineBreak(para: ParaNode): boolean {
  return para.kids.some(kid => {
    if (kid.tag === 'grid') return true;
    if (kid.tag === 'span') {
      return kid.kids.some(child =>
        child.tag === 'br' ||
        child.tag === 'pb' ||
        (child.tag === 'txt' && /[\r\n]/.test(child.content)),
      );
    }
    if (kid.tag === 'link') {
      return kid.kids.some(span => span.kids.some(child =>
        child.tag === 'br' ||
        child.tag === 'pb' ||
        (child.tag === 'txt' && /[\r\n]/.test(child.content)),
      ));
    }
    return false;
  });
}

function canEncodePipeTable(grid: GridNode): boolean {
  const columns = grid.kids[0]?.kids.length ?? 0;
  if (columns === 0) return false;
  return grid.kids.every(row =>
    row.kids.length === columns &&
    row.kids.every(cell =>
      cell.cs === 1 &&
      cell.rs === 1 &&
      cell.kids.length === 1 &&
      cell.kids[0].tag === 'para' &&
      cell.kids[0].props.heading === undefined &&
      cell.kids[0].props.listOrd === undefined &&
      !paraHasLineBreak(cell.kids[0]),
    ),
  );
}

function encodePipeTable(
  grid: GridNode,
  warns: string[],
  includeImages: boolean,
): string {
  const rows = grid.kids.map(row => row.kids.map(cell => {
    const para = cell.kids[0] as ParaNode;
    return encodePara(para, warns, includeImages).replace(/\|/g, '\\|');
  }));
  const renderRow = (cells: string[]) => `| ${cells.join(' | ')} |`;
  const separator = renderRow(rows[0].map(() => '---'));
  return [renderRow(rows[0]), separator, ...rows.slice(1).map(renderRow)].join('\n');
}

function encodeHtmlTable(grid: GridNode, warns: string[], includeImages: boolean): string {

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
      const content = cell.kids.map(p => p.tag === 'para' ? encodePara(p, warns, includeImages) : encodeGrid(p, warns, includeImages)).join('\n');
      cells += `<${tag}${cs}${rs} style="${styles.join(';')}">${content}</${tag}>`;
      colIdx += cell.cs;
    }
    rows += `<tr>${cells}</tr>\n`;
  }

  return `<table style="border-collapse:collapse;width:100%">\n<tbody>\n${rows}</tbody>\n</table>\n`;
}

registry.registerEncoder(new MdEncoder());
