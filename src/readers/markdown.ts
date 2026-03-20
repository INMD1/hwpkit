/**
 * MarkdownReader - Markdown 텍스트 → IrDocumentNode
 *
 * ReadStrategy 구현체.
 * 정규표현식 기반 자체 파서 (외부 의존성 없음).
 *
 * 처리 순서:
 *   1. 줄 단위 분류 (HeadingLine, ListLine, TableLine, FencedCodeLine, TextLine)
 *   2. 연속된 같은 타입 줄 → Block 단위로 그룹핑
 *   3. Block → IR Node 변환
 *   4. 인라인 서식 (볼드, 이탤릭 등) → RunNode 분해
 */

import type { ReadStrategy } from '../core/strategy';
import type {
  IrDocumentNode,
  IrBlockNode,
  IrParagraphNode,
  IrRunNode,
  IrTableNode,
  RunStyle,
  ParagraphStyle,
  HeadingLevel,
} from '../core/ir';
import {
  makeDocument,
  makeSection,
  makeParagraph,
  makeRun,
  makeTable,
  makeTableRow,
  makeTableCell,
  DEFAULT_PAGE_LAYOUT,
} from '../core/ir';

// ─── 줄 타입 분류 ─────────────────────────────────────────────────────────────

type LineKind =
  | 'heading'
  | 'listUnordered'
  | 'listOrdered'
  | 'tableRow'
  | 'tableSep'
  | 'fenceOpen'
  | 'fenceClose'
  | 'hrule'
  | 'blank'
  | 'text';

interface ClassifiedLine {
  kind: LineKind;
  raw: string;
  depth?: number;    // heading level or list indent
  marker?: string;   // list marker
}

const RE_HEADING       = /^(#{1,6})\s+(.*)/;
const RE_UNORDERED     = /^(\s*)[-*+]\s+(.*)/;
const RE_ORDERED       = /^(\s*)\d+[.)]\s+(.*)/;
const RE_TABLE_ROW     = /^\s*\|.+\|/;
const RE_TABLE_SEP     = /^\s*\|[\s|:-]+\|/;
const RE_FENCE         = /^```/;
const RE_HRULE         = /^[-*_]{3,}\s*$/;

function classifyLine(raw: string): ClassifiedLine {
  if (!raw.trim()) return { kind: 'blank', raw };

  const headingMatch = raw.match(RE_HEADING);
  if (headingMatch) return { kind: 'heading', raw, depth: headingMatch[1].length };

  if (RE_FENCE.test(raw)) return { kind: 'fenceOpen', raw };
  if (RE_HRULE.test(raw)) return { kind: 'hrule', raw };

  const tableRowMatch = raw.match(RE_TABLE_ROW);
  if (tableRowMatch) {
    if (RE_TABLE_SEP.test(raw)) return { kind: 'tableSep', raw };
    return { kind: 'tableRow', raw };
  }

  const ulMatch = raw.match(RE_UNORDERED);
  if (ulMatch) return { kind: 'listUnordered', raw, depth: Math.floor(ulMatch[1].length / 2), marker: ulMatch[0].trim().charAt(0) };

  const olMatch = raw.match(RE_ORDERED);
  if (olMatch) return { kind: 'listOrdered', raw, depth: Math.floor(olMatch[1].length / 2) };

  return { kind: 'text', raw };
}

// ─── 인라인 파서 ──────────────────────────────────────────────────────────────

/**
 * 인라인 서식 문자열을 IrRunNode 배열로 분해합니다.
 * 처리 패턴: ***bold+italic***, **bold**, *italic*, ~~strike~~, `code`,
 *            <u>underline</u>, <sup>...</sup>, <sub>...</sub>, [text](url)
 */
export function parseInline(text: string): IrRunNode[] {
  const runs: IrRunNode[] = [];

  // 누적 패턴 정규식
  const RE = /(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*|~~(.+?)~~|`(.+?)`|<u>(.+?)<\/u>|<sup>(.+?)<\/sup>|<sub>(.+?)<\/sub>|\[([^\]]+)\]\(([^)]+)\))/g;

  let lastIndex = 0;
  let match: RegExpExecArray | null;

  while ((match = RE.exec(text)) !== null) {
    // 앞의 일반 텍스트
    if (match.index > lastIndex) {
      const plain = text.slice(lastIndex, match.index);
      if (plain) runs.push(makeRun(plain));
    }

    const [full, , boldItalic, bold, italic, strike, code, underline, sup, sub, linkText, linkUrl] = match;

    if (boldItalic)   runs.push(makeRun(boldItalic, { bold: true, italic: true }));
    else if (bold)    runs.push(makeRun(bold, { bold: true }));
    else if (italic)  runs.push(makeRun(italic, { italic: true }));
    else if (strike)  runs.push(makeRun(strike, { strikethrough: true }));
    else if (code)    runs.push(makeRun(code, { fontFamily: 'Courier New' }));
    else if (underline) runs.push(makeRun(underline, { underline: true }));
    else if (sup)     runs.push(makeRun(sup, { superscript: true }));
    else if (sub)     runs.push(makeRun(sub, { subscript: true }));
    else if (linkText && linkUrl) {
      runs.push({
        kind: 'run',
        style: { underline: true, colorHex: '0563C1' },
        children: [{ kind: 'text', value: linkText }],
      });
    }

    lastIndex = RE.lastIndex;
  }

  // 남은 텍스트
  if (lastIndex < text.length) {
    const remaining = text.slice(lastIndex);
    if (remaining) runs.push(makeRun(remaining));
  }

  if (runs.length === 0 && text) runs.push(makeRun(text));

  return runs;
}

// ─── 블록 파서 ────────────────────────────────────────────────────────────────

function parseHeading(line: ClassifiedLine): IrParagraphNode {
  const match = line.raw.match(RE_HEADING)!;
  const level = match[1].length as HeadingLevel;
  const content = match[2];
  const runs = parseInline(content);
  return makeParagraph({ headingLevel: level }, runs);
}

function parseListItem(line: ClassifiedLine): IrParagraphNode {
  const isOrdered = line.kind === 'listOrdered';
  const re = isOrdered ? RE_ORDERED : RE_UNORDERED;
  const match = line.raw.match(re)!;
  const content = isOrdered ? line.raw.replace(/^\s*\d+[.)]\s+/, '') : line.raw.replace(/^\s*[-*+]\s+/, '');
  const runs = parseInline(content.trim());
  return makeParagraph(
    { listLevel: line.depth ?? 0, listOrdered: isOrdered },
    runs,
  );
}

function parseTextLine(raw: string): IrParagraphNode {
  const runs = parseInline(raw);
  return makeParagraph({}, runs);
}

function parseTableLines(lines: ClassifiedLine[]): IrTableNode {
  const rows = lines.filter(l => l.kind === 'tableRow');
  const tableRows = rows.map(row => {
    const cells = row.raw
      .split('|')
      .map(c => c.trim())
      .filter((c, i, arr) => i > 0 && i < arr.length - 1); // 앞뒤 | 제거

    return makeTableRow(
      cells.map(cellText => {
        const runs = parseInline(cellText);
        return makeTableCell([makeParagraph({}, runs)]);
      }),
    );
  });
  return makeTable(tableRows);
}

// ─── 블록 그룹핑 ──────────────────────────────────────────────────────────────

interface LineGroup {
  type: 'heading' | 'list' | 'table' | 'code' | 'hr' | 'paragraph';
  lines: ClassifiedLine[];
}

function groupLines(classified: ClassifiedLine[]): LineGroup[] {
  const groups: LineGroup[] = [];
  let i = 0;

  while (i < classified.length) {
    const line = classified[i];

    if (line.kind === 'blank') { i++; continue; }

    if (line.kind === 'heading') {
      groups.push({ type: 'heading', lines: [line] });
      i++;
    } else if (line.kind === 'hrule') {
      groups.push({ type: 'hr', lines: [line] });
      i++;
    } else if (line.kind === 'fenceOpen') {
      // 코드 블록: 닫는 ``` 까지 수집
      const codeLines: ClassifiedLine[] = [line];
      i++;
      while (i < classified.length && classified[i].kind !== 'fenceClose' && classified[i].kind !== 'fenceOpen') {
        codeLines.push(classified[i]);
        i++;
      }
      if (i < classified.length) { codeLines.push(classified[i]); i++; }
      groups.push({ type: 'code', lines: codeLines });
    } else if (line.kind === 'tableRow' || line.kind === 'tableSep') {
      const tableLines: ClassifiedLine[] = [];
      while (i < classified.length && (classified[i].kind === 'tableRow' || classified[i].kind === 'tableSep')) {
        tableLines.push(classified[i]);
        i++;
      }
      groups.push({ type: 'table', lines: tableLines });
    } else if (line.kind === 'listUnordered' || line.kind === 'listOrdered') {
      const listLines: ClassifiedLine[] = [];
      while (i < classified.length && (classified[i].kind === 'listUnordered' || classified[i].kind === 'listOrdered')) {
        listLines.push(classified[i]);
        i++;
      }
      groups.push({ type: 'list', lines: listLines });
    } else {
      // 일반 텍스트: 빈 줄이 나올 때까지 합침
      const textLines: ClassifiedLine[] = [];
      while (i < classified.length && classified[i].kind === 'text') {
        textLines.push(classified[i]);
        i++;
      }
      groups.push({ type: 'paragraph', lines: textLines });
    }
  }

  return groups;
}

// ─── 블록 → IR 변환 ───────────────────────────────────────────────────────────

function groupToBlocks(group: LineGroup): IrBlockNode[] {
  switch (group.type) {
    case 'heading':
      return [parseHeading(group.lines[0])];

    case 'list':
      return group.lines.map(parseListItem);

    case 'table':
      return [parseTableLines(group.lines)];

    case 'hr': {
      const run: IrRunNode = { kind: 'run', style: {}, children: [{ kind: 'pageBreak' }] };
      return [makeParagraph({}, [run])];
    }

    case 'code': {
      const codeText = group.lines
        .slice(1, -1)
        .map(l => l.raw)
        .join('\n');
      return [makeParagraph({}, [makeRun(codeText, { fontFamily: 'Courier New' })])];
    }

    case 'paragraph': {
      const text = group.lines.map(l => l.raw).join('\n');
      return [parseTextLine(text)];
    }

    default:
      return [];
  }
}

// ─── MarkdownReader ───────────────────────────────────────────────────────────

export class MarkdownReader implements ReadStrategy {
  async read(data: Uint8Array): Promise<IrDocumentNode> {
    const text = new TextDecoder('utf-8').decode(data);
    return this.readText(text);
  }

  async readText(markdown: string): Promise<IrDocumentNode> {
    // YAML frontmatter 추출
    let meta = {};
    let body = markdown;
    const fmMatch = markdown.match(/^---\n([\s\S]*?)\n---\n([\s\S]*)$/);
    if (fmMatch) {
      meta = parseYamlFrontmatter(fmMatch[1]);
      body = fmMatch[2];
    }

    const lines = body.split('\n').map(classifyLine);
    const groups = groupLines(lines);
    const blocks: IrBlockNode[] = groups.flatMap(groupToBlocks);
    const section = makeSection({ ...DEFAULT_PAGE_LAYOUT }, blocks);
    return makeDocument(meta, [section]);
  }
}

function parseYamlFrontmatter(yaml: string): Record<string, string> {
  const result: Record<string, string> = {};
  for (const line of yaml.split('\n')) {
    const match = line.match(/^(\w+):\s*"?([^"]*)"?/);
    if (match) result[match[1]] = match[2].trim();
  }
  return result;
}
