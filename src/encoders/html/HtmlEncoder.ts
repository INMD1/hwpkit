import type { Encoder } from '../../contract/encoder';
import type { DocRoot, ParaNode, SpanNode, GridNode, ContentNode, ImgNode, LinkNode } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import { succeed, fail } from '../../contract/result';
import { TextKit } from '../../toolkit/TextKit';
import { registry } from '../../pipeline/registry';

export class HtmlEncoder implements Encoder {
  readonly format = 'html';

  async encode(doc: DocRoot): Promise<Outcome<Uint8Array>> {
    try {
      const warns: string[] = [];
      const bodyParts: string[] = [];

      for (const sheet of doc.kids) {
        // Header/footer as comments
        if (sheet.header && sheet.header.length > 0) {
          const hText = sheet.header.map(p => encodePara(p, warns)).join('');
          bodyParts.push(`<div class="hwp-header">${hText}</div>`);
        }

        for (const kid of sheet.kids) {
          bodyParts.push(encodeContent(kid, warns));
        }

        if (sheet.footer && sheet.footer.length > 0) {
          const fText = sheet.footer.map(p => encodePara(p, warns)).join('');
          bodyParts.push(`<div class="hwp-footer">${fText}</div>`);
        }
      }

      const title = esc(doc.meta?.title ?? '');
      const html = `<!DOCTYPE html>\n<html lang="ko">\n<head>\n<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width, initial-scale=1.0">\n<title>${title}</title>\n<style>\n${BASE_CSS}\n</style>\n</head>\n<body>\n<div class="hwp-doc">\n${bodyParts.join('\n')}\n</div>\n</body>\n</html>`;

      return succeed(TextKit.encode(html), warns);
    } catch (e: any) {
      return fail(`HTML encode error: ${e?.message ?? String(e)}`);
    }
  }
}

const BASE_CSS = `
body { margin: 0; padding: 0; background: #f0f0f0; }
.hwp-doc { max-width: 800px; margin: 0 auto; background: #fff; padding: 40px 60px; box-shadow: 0 0 8px rgba(0,0,0,0.15); }
.hwp-header, .hwp-footer { color: #666; font-size: 0.9em; border-bottom: 1px solid #ddd; margin-bottom: 8px; padding-bottom: 4px; }
.hwp-footer { border-top: 1px solid #ddd; border-bottom: none; margin-top: 8px; padding-top: 4px; }
p { margin: 0; padding: 0; line-height: 1.6; }
table { border-collapse: collapse; width: 100%; margin: 8px 0; }
td, th { border: 1px solid #ccc; padding: 4px 8px; vertical-align: top; }
img { max-width: 100%; height: auto; }
`.trim();

function encodeContent(node: ContentNode, warns: string[]): string {
  return node.tag === 'grid' ? encodeGrid(node, warns) : encodePara(node, warns);
}

function encodePara(para: ParaNode, warns: string[]): string {
  const kids = para.kids.map((k): string => {
    if (k.tag === 'span') return encodeSpan(k, warns);
    if (k.tag === 'img') return encodeImage(k);
    if (k.tag === 'link') {
      const link = k as LinkNode;
      const inner = link.kids.map(s => encodeSpan(s, warns)).join('');
      return `<a href="${esc(link.href)}">${inner}</a>`;
    }
    return '';
  }).join('');

  // Heading
  if (para.props.heading) {
    const tag = `h${para.props.heading}`;
    return `<${tag}>${kids}</${tag}>\n`;
  }

  // List
  if (para.props.listOrd !== undefined) {
    const indent = (para.props.listLv ?? 0) * 20;
    const style = indent > 0 ? ` style="margin-left:${indent}px"` : '';
    const marker = para.props.listOrd ? `<span class="list-marker">1. </span>` : `<span class="list-marker">• </span>`;
    return `<p${style}>${marker}${kids}</p>\n`;
  }

  // Alignment
  const align = para.props.align;
  const styleAttrs: string[] = [];
  if (align && align !== 'left') styleAttrs.push(`text-align:${align}`);
  if (para.props.indentPt) styleAttrs.push(`margin-left:${para.props.indentPt.toFixed(1)}pt`);
  if (para.props.spaceBefore) styleAttrs.push(`margin-top:${para.props.spaceBefore.toFixed(1)}pt`);
  if (para.props.spaceAfter) styleAttrs.push(`margin-bottom:${para.props.spaceAfter.toFixed(1)}pt`);
  if (para.props.lineHeight) styleAttrs.push(`line-height:${para.props.lineHeight}`);

  const styleAttr = styleAttrs.length > 0 ? ` style="${styleAttrs.join(';')}"` : '';
  return `<p${styleAttr}>${kids || '&nbsp;'}</p>\n`;
}

function encodeSpan(span: SpanNode, warns: string[]): string {
  const parts: string[] = [];
  let hasPageNum = false;

  for (const kid of span.kids) {
    if (kid.tag === 'txt') {
      parts.push(esc(kid.content));
    } else if (kid.tag === 'br') {
      parts.push('<br>');
    } else if (kid.tag === 'pb') {
      parts.push('<div style="page-break-after:always"></div>');
    } else if (kid.tag === 'pagenum') {
      hasPageNum = true;
      warns.push('[SHIELD] HTML: 페이지 번호 — 정적 값으로 대체됨');
      parts.push('<span class="page-num">[페이지]</span>');
    }
  }

  let text = parts.join('');
  if (hasPageNum && text.trim() === '<span class="page-num">[페이지]</span>') {
    // keep as-is
  }

  const p = span.props;
  const css: string[] = [];
  if (p.font) css.push(`font-family:${esc(p.font)}`);
  if (p.pt) css.push(`font-size:${p.pt}pt`);
  if (p.color) css.push(`color:#${p.color}`);
  if (p.bg) css.push(`background-color:#${p.bg}`);
  if (p.b) css.push('font-weight:bold');
  if (p.i) css.push('font-style:italic');

  const decorations: string[] = [];
  if (p.u) decorations.push('underline');
  if (p.s) decorations.push('line-through');
  if (decorations.length > 0) css.push(`text-decoration:${decorations.join(' ')}`);

  if (p.sup) return `<sup${css.length ? ` style="${css.join(';')}"` : ''}>${text}</sup>`;
  if (p.sub) return `<sub${css.length ? ` style="${css.join(';')}"` : ''}>${text}</sub>`;
  if (css.length > 0) return `<span style="${css.join(';')}">${text}</span>`;
  return text;
}

function encodeImage(img: ImgNode): string {
  const wStyle = img.w ? ` width="${Math.round(img.w / 72 * 96)}px"` : '';
  const hStyle = img.h ? ` height="${Math.round(img.h / 72 * 96)}px"` : '';
  const alt = esc(img.alt ?? '');
  return `<img src="data:${img.mime};base64,${img.b64}" alt="${alt}"${wStyle}${hStyle}>`;
}

function encodeGrid(grid: GridNode, warns: string[]): string {
  if (grid.kids.length === 0) return '';

  // Build occupancy map for rowspan
  const rowCount = grid.kids.length;
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
    let ci = 0;
    for (const cell of row.kids) {
      while (occupancy[ri].has(ci)) ci++;

      const isHeader = cell.props.isHeader || (grid.props.headerRow && ri === 0);
      const tag = isHeader ? 'th' : 'td';

      const cs = cell.cs > 1 ? ` colspan="${cell.cs}"` : '';
      const rs = cell.rs > 1 ? ` rowspan="${cell.rs}"` : '';

      const styleAttrs: string[] = [];
      if (cell.props.bg) styleAttrs.push(`background-color:#${cell.props.bg}`);
      const va = cell.props.va;
      if (va === 'mid') styleAttrs.push('vertical-align:middle');
      else if (va === 'bot') styleAttrs.push('vertical-align:bottom');
      const styleAttr = styleAttrs.length > 0 ? ` style="${styleAttrs.join(';')}"` : '';

      const content = cell.kids.map(p => encodePara(p, warns)).join('');
      cells += `<${tag}${cs}${rs}${styleAttr}>${content}</${tag}>`;
      ci += cell.cs;
    }
    rows += `<tr>${cells}</tr>\n`;
  }

  return `<table>\n<tbody>\n${rows}</tbody>\n</table>\n`;
}

function esc(s: string): string {
  return TextKit.escapeXml(s);
}

registry.registerEncoder(new HtmlEncoder());
