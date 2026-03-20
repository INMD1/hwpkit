import type { Decoder } from '../../contract/decoder';
import type { DocRoot, ContentNode, ParaNode, SpanNode } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import type { ParaProps, TextProps } from '../../model/doc-props';
import { A4 } from '../../model/doc-props';
import { succeed, fail } from '../../contract/result';
import { buildRoot, buildSheet, buildPara, buildSpan, buildGrid, buildRow, buildCell } from '../../model/builders';
import { ShieldedParser } from '../../safety/ShieldedParser';
import { TextKit } from '../../toolkit/TextKit';
import { registry } from '../../pipeline/registry';

export class MdDecoder implements Decoder {
  readonly format = 'md';

  async decode(data: Uint8Array): Promise<Outcome<DocRoot>> {
    const shield = new ShieldedParser();
    const warns: string[] = [];

    try {
      const text = TextKit.decode(data);
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

function parseInline(text: string): SpanNode[] {
  const spans: SpanNode[] = [];
  let rem = text;

  while (rem.length > 0) {
    // Bold+italic
    let m = rem.match(/^(.*?)\*\*\*(.+?)\*\*\*(.*)/s);
    if (m) { if (m[1]) spans.push(buildSpan(m[1])); spans.push(buildSpan(m[2], { b: true, i: true })); rem = m[3]; continue; }

    // Bold
    m = rem.match(/^(.*?)\*\*(.+?)\*\*(.*)/s);
    if (m) { if (m[1]) spans.push(buildSpan(m[1])); spans.push(buildSpan(m[2], { b: true })); rem = m[3]; continue; }

    // Italic
    m = rem.match(/^(.*?)\*(.+?)\*(.*)/s);
    if (m) { if (m[1]) spans.push(buildSpan(m[1])); spans.push(buildSpan(m[2], { i: true })); rem = m[3]; continue; }

    // Inline code
    m = rem.match(/^(.*?)`(.+?)`(.*)/s);
    if (m) { if (m[1]) spans.push(buildSpan(m[1])); spans.push(buildSpan(m[2], { font: 'Courier New' })); rem = m[3]; continue; }

    spans.push(buildSpan(rem));
    break;
  }

  return spans.length > 0 ? spans : [buildSpan(text)];
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
