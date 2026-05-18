import type { DocRoot, ContentNode, ParaNode, SpanNode, ImgNode } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import type { ParaProps, TextProps } from '../../model/doc-props';
import { A4 } from '../../model/doc-props';
import { succeed, fail } from '../../contract/result';
import { buildRoot, buildSheet, buildPara, buildSpan, buildImg, buildGrid, buildRow, buildCell } from '../../model/builders';
import { ShieldedParser } from '../../safety/ShieldedParser';
import { TextKit } from '../../toolkit/TextKit';
import { registry } from '../../pipeline/registry';
import { BaseDecoder } from '../../core/BaseDecoder';

export class MdDecoder extends BaseDecoder {
  protected getFormat(): string { return 'md'; }

  async decode(data: Uint8Array): Promise<Outcome<DocRoot>> {
    const shield = new ShieldedParser();
    const warns: string[] = [];

    try {
      const text = this.bytesToString(data);
      const lines = text.split(/\r?\n/);
      const kids: ContentNode[] = [];

      let i = 0;
      while (i < lines.length) {
        const line = lines[i];

        // Heading
        const headingMatch = line.match(/^(#{1,6})\s+(.+)$/);
        if (headingMatch) {
          const level = headingMatch[1].length as 1 | 2 | 3 | 4 | 5 | 6;
          kids.push(buildPara([buildSpan(headingMatch[2], { b: level <= 2 })], { heading: level }));
          i++; continue;
        }

        // Table (pipe + separator line)
        if (line.includes('|') && i + 1 < lines.length && lines[i + 1].match(/^\s*\|?\s*[-:]+\s*\|/)) {
          const tableResult = shield.guard(() => parseMdTable(lines, i), null, `md:table@${i}`);
          if (tableResult) { kids.push(tableResult.node); i = tableResult.nextLine; continue; }
        }

        // HR
        if (line.match(/^[-*_]{3,}$/)) { kids.push(buildPara([buildSpan('')], {})); i++; continue; }

        // List item
        const listMatch = line.match(/^(\s*)([-*+]|\d+\.)\s+(.+)$/);
        if (listMatch) {
          kids.push(buildPara(parseInline(listMatch[3]), {
            listLv: Math.floor(listMatch[1].length / 2),
            listOrd: /\d+\./.test(listMatch[2]),
          }));
          i++; continue;
        }

        // Blockquote
        const bqMatch = line.match(/^>\s*(.*)$/);
        if (bqMatch) { kids.push(buildPara([buildSpan(bqMatch[1])], { indentPt: 28 })); i++; continue; }

        // Code block
        if (line.startsWith('```')) {
          const codeLines: string[] = [];
          i++;
          while (i < lines.length && !lines[i].startsWith('```')) { codeLines.push(lines[i]); i++; }
          i++;
          kids.push(buildPara([buildSpan(codeLines.join('\n'), { font: 'Courier New' })], {}));
          continue;
        }

        // Empty line
        if (line.trim() === '') { i++; continue; }

        // Regular paragraph — check for alignment div
        const alignMatch = line.match(/^<div\s+align="(center|right|left)">(.*?)<\/div>$/i);
        if (alignMatch) {
          const align = alignMatch[1].toLowerCase() as 'left' | 'center' | 'right';
          kids.push(buildPara(parseInline(alignMatch[2]), { align }));
          i++; continue;
        }

        // Regular paragraph
        kids.push(buildPara(parseInline(line), {}));
        i++;
      }

      warns.push(...shield.flush());
      const sheet = buildSheet(kids.length > 0 ? kids : [buildPara([buildSpan('')])], A4);
      return succeed(buildRoot({}, [sheet]), warns);
    } catch (e: any) {
      warns.push(...shield.flush());
      return fail(`MD decode error: ${e?.message ?? String(e)}`, warns);
    }
  }
}

function parseInline(text: string): (SpanNode | ImgNode)[] {
  const result: (SpanNode | ImgNode)[] = [];
  let rem = text;

  while (rem.length > 0) {
    // Image: ![alt](data:mime;base64,...)
    let m = rem.match(/^(.*?)!\[([^\]]*)\]\((data:([^;]+);base64,([^)]+))\)(.*)/s);
    if (m) {
      if (m[1]) result.push(buildSpan(m[1]));
      const mime = m[4] as ImgNode['mime'];
      const validMimes = ['image/png', 'image/jpeg', 'image/gif', 'image/bmp'];
      result.push(buildImg(m[5], validMimes.includes(mime) ? mime : 'image/png', 100, 100, m[2] || undefined));
      rem = m[6]; continue;
    }

    // Image: ![alt](url) — non-base64
    m = rem.match(/^(.*?)!\[([^\]]*)\]\(([^)]+)\)(.*)/s);
    if (m) {
      if (m[1]) result.push(buildSpan(m[1]));
      // Can't convert URL to base64, just preserve alt text
      result.push(buildSpan(`[이미지: ${m[2] || m[3]}]`));
      rem = m[4]; continue;
    }

    // Bold+italic
    m = rem.match(/^(.*?)\*\*\*(.+?)\*\*\*(.*)/s);
    if (m) { if (m[1]) result.push(buildSpan(m[1])); result.push(buildSpan(m[2], { b: true, i: true })); rem = m[3]; continue; }

    // Bold
    m = rem.match(/^(.*?)\*\*(.+?)\*\*(.*)/s);
    if (m) { if (m[1]) result.push(buildSpan(m[1])); result.push(buildSpan(m[2], { b: true })); rem = m[3]; continue; }

    // Italic
    m = rem.match(/^(.*?)\*(.+?)\*(.*)/s);
    if (m) { if (m[1]) result.push(buildSpan(m[1])); result.push(buildSpan(m[2], { i: true })); rem = m[3]; continue; }

    // Strikethrough ~~text~~
    m = rem.match(/^(.*?)~~(.+?)~~(.*)/s);
    if (m) { if (m[1]) result.push(buildSpan(m[1])); result.push(buildSpan(m[2], { s: true })); rem = m[3]; continue; }

    // Underline <u>text</u>
    m = rem.match(/^(.*?)<u>(.+?)<\/u>(.*)/si);
    if (m) { if (m[1]) result.push(buildSpan(m[1])); result.push(buildSpan(m[2], { u: true })); rem = m[3]; continue; }

    // Superscript <sup>text</sup>
    m = rem.match(/^(.*?)<sup>(.+?)<\/sup>(.*)/si);
    if (m) { if (m[1]) result.push(buildSpan(m[1])); result.push(buildSpan(m[2], { sup: true })); rem = m[3]; continue; }

    // Subscript <sub>text</sub>
    m = rem.match(/^(.*?)<sub>(.+?)<\/sub>(.*)/si);
    if (m) { if (m[1]) result.push(buildSpan(m[1])); result.push(buildSpan(m[2], { sub: true })); rem = m[3]; continue; }

    // Inline code
    m = rem.match(/^(.*?)`(.+?)`(.*)/s);
    if (m) { if (m[1]) result.push(buildSpan(m[1])); result.push(buildSpan(m[2], { font: 'Courier New' })); rem = m[3]; continue; }

    result.push(buildSpan(rem));
    break;
  }

  return result.length > 0 ? result : [buildSpan(text)];
}

function parseMdTable(lines: string[], startLine: number): { node: any; nextLine: number } | null {
  const parse = (line: string) => line.split('|').map(c => c.trim()).filter((c, i, arr) => i > 0 || c !== '');
  const headers = parse(lines[startLine]);

  let cur = startLine + 2;
  const rows: string[][] = [];
  while (cur < lines.length) {
    if (!lines[cur].includes('|')) break;
    const cells = parse(lines[cur]);
    if (cells.length === 0) break;
    rows.push(cells);
    cur++;
  }

  const allRows = [headers, ...rows];
  const gridRows = allRows.map((row, ri) =>
    buildRow(row.map(cell => buildCell([buildPara([buildSpan(cell, ri === 0 ? { b: true } : {})])]))),
  );

  return { node: buildGrid(gridRows), nextLine: cur };
}

registry.registerDecoder(new MdDecoder());
