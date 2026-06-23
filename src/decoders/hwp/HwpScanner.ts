import type { Decoder } from '../../contract/decoder';
import type { DocRoot, ContentNode, ParaNode, SpanNode, ImgNode, GridNode, PageNumNode } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import type { Align, Stroke, StrokeKind, PageDims, TextProps, ParaProps, CellProps, GridProps } from '../../model/doc-props';
import { succeed, fail } from '../../contract/result';
import { buildRoot, buildSheet, buildPara, buildSpan, buildGrid, buildRow, buildCell, buildImg, buildPb, buildPageNum } from '../../model/builders';
import { ShieldedParser } from '../../safety/ShieldedParser';
import { BinaryKit } from '../../toolkit/BinaryKit';
import { TextKit } from '../../toolkit/TextKit';
import { Metric, safeHex, safeFont } from '../../safety/StyleBridge';
import { registry } from '../../pipeline/registry';
import { A4 } from '../../model/doc-props';
import pako from 'pako';

/* ═══════════════════════════════════════════════════════════════
   HWP 5.0 Tag Constants
   ═══════════════════════════════════════════════════════════════ */

const HWPTAG_BEGIN = 16;

const TAG_FACE_NAME       = HWPTAG_BEGIN + 3;   // 19
const TAG_BORDER_FILL     = HWPTAG_BEGIN + 4;   // 20
const TAG_CHAR_SHAPE      = HWPTAG_BEGIN + 5;   // 21
const TAG_PARA_SHAPE      = HWPTAG_BEGIN + 9;   // 25
const TAG_PARA_HEADER     = HWPTAG_BEGIN + 50;  // 66
const TAG_PARA_TEXT       = HWPTAG_BEGIN + 51;  // 67
const TAG_PARA_CHAR_SHAPE = HWPTAG_BEGIN + 52;  // 68
const TAG_CTRL_HEADER     = HWPTAG_BEGIN + 55;  // 71
const TAG_PAGE_DEF        = HWPTAG_BEGIN + 57;  // 73

// TABLE / CELL tags vary by HWP version
const TAG_LIST_HEADER = HWPTAG_BEGIN + 56;  // 72
const TAG_TABLE_A = HWPTAG_BEGIN + 61;  // 77
const TAG_CELL_A  = HWPTAG_BEGIN + 62;  // 78
const TAG_TABLE_B = HWPTAG_BEGIN + 64;  // 80
const TAG_CELL_B  = HWPTAG_BEGIN + 65;  // 81

function isTableTag(t: number) { return t === TAG_TABLE_A || t === TAG_TABLE_B; }
function isCellTag(t: number)  { return t === TAG_CELL_A || t === TAG_CELL_B || t === TAG_LIST_HEADER; }

// CTRL_HEADER ctrlId values (UINT32-LE as ASCII)
const CTRL_TABLE = 0x74626C20;  // 'tbl ' = 표(table)
const CTRL_IMAGE = 0x696D6720;  // 'img '
const CTRL_OBJ   = 0x6F626A20;  // 'obj '
const CTRL_FIG   = 0x66696720;  // 'fig '
const CTRL_GSO   = 0x67736F20;  // 'gso ' = 그리기 객체 (drawing object, contains embedded images)
const CTRL_HEAD  = 0x68656164;  // 'head' = 머리말
const CTRL_FOOT  = 0x666F6F74;  // 'foot' = 꼬리말
const CTRL_ATNO  = 0x61746E6F;  // 'atno' = 자동 번호 (쪽번호 등)

/* ═══════════════════════════════════════════════════════════════
   Types
   ═══════════════════════════════════════════════════════════════ */

interface HwpRecord {
  tag: number;
  level: number;
  data: Uint8Array;
}

interface HwpCharShape {
  faceIds: number[];
  height: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikeout: boolean;
  superscript: boolean;
  subscript: boolean;
  textColor: string;
}
interface HwpParaShape {
  align: Align;
  spaceBefore: number;
  spaceAfter: number;
  lineSpacing: number;
  lineSpacingType: 0 | 1 | 2 | 3; // 0=PERCENT, 1=FIXED, 2=BETWEEN_LINES, 3=AT_LEAST
  leftMargin: number;
  rightMargin: number;
  indent: number;
  verAlign?: 'baseline' | 'top' | 'center' | 'bottom';
  lineWrap?: 'break' | 'squeeze' | 'keep';
}
interface HwpBorderFill {
  borders: { type: number; widthPt: number; color: string }[];
  bgColor?: string;
}

interface DocInfo {
  faceNames: string[];
  charShapes: HwpCharShape[];
  paraShapes: HwpParaShape[];
  borderFills: HwpBorderFill[];
}

interface ParsedChar { pos: number; ch: string }
interface ParsedCtrl { pos: number; ctrlId: number; objId: number; matched: boolean }
interface ParaTextResult { chars: ParsedChar[]; controls: ParsedCtrl[] }

interface OleObject {
  id: number;
  data: Uint8Array;
  mimeType: string;
}

/* ═══════════════════════════════════════════════════════════════
   Low-level record parsing
   ═══════════════════════════════════════════════════════════════ */

function parseRecords(data: Uint8Array): HwpRecord[] {
  const out: HwpRecord[] = [];
  let off = 0;
  while (off + 4 <= data.length) {
    const hdr = BinaryKit.readU32LE(data, off);
    const tag   = hdr & 0x3FF;
    const level = (hdr >> 10) & 0x3FF;
    let size    = (hdr >> 20) & 0xFFF;
    off += 4;
    if (size === 0xFFF) {
      if (off + 4 > data.length) break;
      size = BinaryKit.readU32LE(data, off);
      off += 4;
    }
    if (off + size > data.length) break;
    out.push({ tag, level, data: data.subarray(off, off + size) });
    off += size;
  }
  return out;
}

function tryInflate(data: Uint8Array): Uint8Array {
  try { return pako.inflate(data); } catch {
    try { return pako.inflateRaw(data); } catch { return data; }
  }
}

/* ═══════════════════════════════════════════════════════════════
   FileHeader
   ═══════════════════════════════════════════════════════════════ */

function parseFileHeader(buf: Uint8Array) {
  if (buf.length < 40) return { compressed: true, encrypted: false };
  const props = BinaryKit.readU32LE(buf, 36);
  return { compressed: (props & 1) !== 0, encrypted: (props & 2) !== 0 };
}

/* ═══════════════════════════════════════════════════════════════
   DocInfo parsing
   ═══════════════════════════════════════════════════════════════ */

function parseDocInfo(data: Uint8Array, compressed: boolean): DocInfo {
  const raw = compressed ? tryInflate(data) : data;
  const recs = parseRecords(raw);
  const info: DocInfo = { faceNames: [], charShapes: [], paraShapes: [], borderFills: [] };

  for (const r of recs) {
    try {
      if (r.tag === TAG_FACE_NAME)   info.faceNames.push(parseFaceName(r.data));
      if (r.tag === TAG_CHAR_SHAPE)  info.charShapes.push(parseCharShape(r.data));
      if (r.tag === TAG_PARA_SHAPE)  info.paraShapes.push(parseParaShape(r.data));
      if (r.tag === TAG_BORDER_FILL) info.borderFills.push(parseBorderFill(r.data));
    } catch { /* skip malformed record */ }
  }
  return info;
}

/* ── FACE_NAME ──────────────────────────────────────────────── */

function parseFaceName(d: Uint8Array): string {
  if (d.length < 3) return '';
  const len = BinaryKit.readU16LE(d, 1);          // UTF-16 char count
  if (d.length < 3 + len * 2) return '';
  return new TextDecoder('utf-16le').decode(d.subarray(3, 3 + len * 2));
}

/* ── CHAR_SHAPE ─────────────────────────────────────────────── */
/*  offset  size  field
    0       14    faceId[7]   (UINT16 × 7)
    14       7    ratio[7]
    21       7    spacing[7]
    28       7    relSize[7]
    35       7    offset[7]
    42       4    height      (UINT32, HWP-units  100 = 1pt)
    46       4    attr        (UINT32, bit flags)
    50       1    shadowX
    51       1    shadowY
    52       4    textColor   (COLORREF R,G,B,0)                     */

function parseCharShape(d: Uint8Array): HwpCharShape {
  const faceIds: number[] = [];
  for (let i = 0; i < 7; i++) faceIds.push(d.length >= (i + 1) * 2 ? BinaryKit.readU16LE(d, i * 2) : 0);

  const height = d.length >= 46 ? BinaryKit.readU32LE(d, 42) : 1000;
  const attr   = d.length >= 50 ? BinaryKit.readU32LE(d, 46) : 0;

  // attr bit layout (HWP 5.0 spec Table 35):
  //  0: italic, 1: bold, 2-4: underline type(3), 5-8: underline shape(4),
  //  9-11: outline(3), 12-13: shadow(2), 14: emboss, 15: engrave,
  //  16-17: super/sub(2, 0=none,1=super,2=sub), 18-20: strikeout type(3),
  //  21-24: strikeout shape(4), 25: annotLine, 26-28: annotLine type,
  //  29: useFontSpace, 30: kerning
  const ulType  = (attr >> 2)  & 0x7;   // 3 bits at 2-4
  const skType  = (attr >> 18) & 0x7;   // 3 bits at 18-20
  const suType  = (attr >> 16) & 0x3;   // 2 bits at 16-17 (0=none,1=super,2=sub)

  return {
    faceIds,
    height: (height > 0 && height < 100000) ? height : 1000,
    italic:      (attr & 1) !== 0,
    bold:        ((attr >> 1) & 1) !== 0,
    underline:   ulType !== 0,
    strikeout:   skType !== 0,
    superscript: suType === 1,
    subscript:   suType === 2,
    textColor:   d.length >= 56 ? colorRef(d, 52) : '000000',
  };
}

/* ── PARA_SHAPE ─────────────────────────────────────────────── */
/*  offset  size  field
    0       4     attr1   (bits 0-1 = line spacing type, bits 2-4 = alignment)
    4       4     leftMargin   (HWPUNIT)
    8       4     rightMargin
    12      4     indent
    16      4     spaceBefore
    20      4     spaceAfter
    24      4     lineSpacing                                         */

const ALIGN_TBL: Record<number, Align> = { 0: 'justify', 1: 'left', 2: 'right', 3: 'center', 4: 'distribute', 5: 'distribute_space' };

function parseParaShape(d: Uint8Array): HwpParaShape {
  if (d.length < 4) return { align: 'left', spaceBefore: 0, spaceAfter: 0, lineSpacing: 160, lineSpacingType: 0, leftMargin: 0, rightMargin: 0, indent: 0 };
  const attr = BinaryKit.readU32LE(d, 0);

  // bits 0-1: 줄 간격 종류 (0=PERCENT, 1=FIXED, 2=BETWEEN_LINES, 3=AT_LEAST)
  const lineSpacingType = (attr & 0x3) as 0 | 1 | 2 | 3;

  // bits 2-4: 정렬 방식 (0=justify,1=left,2=right,3=center,4=distribute,5=split)
  const align = ALIGN_TBL[(attr >> 2) & 0x7] ?? 'left';

  // 세로 정렬 (Bit 18 ~ Bit 19)
  const vVal = (attr >> 18) & 0x3;
  const verAlign = vVal === 1 ? 'top' : vVal === 2 ? 'center' : vVal === 3 ? 'bottom' : 'baseline';

  // 줄 바꿈 기준: attr1 에는 별도 비트 없음, 기본값 'break'
  const lineWrap: 'break' = 'break';

  return {
    align,
    lineSpacingType,
    leftMargin:  d.length >= 8  ? i32(d, 4)  : 0,  // offset 4: 문단 몸체 왼쪽 여백 (HWPUNIT)
    rightMargin: d.length >= 12 ? i32(d, 8)  : 0,  // offset 8: 문단 몸체 오른쪽 여백 (HWPUNIT)
    indent:      d.length >= 16 ? i32(d, 12) : 0,  // offset 12: 첫 줄 들여쓰기 (HWPUNIT)
    spaceBefore: d.length >= 20 ? i32(d, 16) : 0,
    spaceAfter:  d.length >= 24 ? i32(d, 20) : 0,
    lineSpacing: d.length >= 28 ? i32(d, 24) : 160,
    verAlign,
    lineWrap,
  };
}

/* ── BORDER_FILL ────────────────────────────────────────────── */
/*  [0:2]  attr
    For each of 5 borders (left,right,top,bottom,diagonal): 6 bytes
      +0 type(BYTE)  +1 widthIdx(BYTE)  +2 color(COLORREF)
    [32:4] fillType
    [36:4] faceColor (bgColor for solid fill)                        */

const BORDER_W_PT = [0.28, 0.34, 0.43, 0.57, 0.71, 0.85, 1.13, 1.42, 1.70, 1.98, 2.84, 4.25, 5.67, 8.50, 11.34, 14.17];
const BORDER_KIND: Record<number, StrokeKind> = { 0:'solid',1:'dash',2:'dash',3:'dot',4:'dash',5:'dash',6:'dash',7:'double',8:'double',9:'double',10:'none' };

function parseBorderFill(d: Uint8Array): HwpBorderFill {
  // Spec grouped format (표 23):
  //   [0:2]   attr
  //   [2:4]   4 border types  (left, right, top, bottom) — 1 byte each
  //   [6:4]   4 border widths (left, right, top, bottom) — 1 byte each (index into BORDER_W_PT)
  //   [10:16] 4 border colors (left, right, top, bottom) — 4 bytes each (COLORREF)
  //   [26:3]  diagonal: type(1) + width(1) + color(4) = 6 bytes actually [26:6]
  //   [32:4]  fillType
  //   [36:4]  faceColor (bgColor for solid fill)
  const borders: HwpBorderFill['borders'] = [];
  const BASE_TYPE  = 2;   // 4 type bytes
  const BASE_WIDTH = 6;   // 4 width bytes
  const BASE_COLOR = 10;  // 4 × 4-byte colors
  for (let i = 0; i < 4; i++) {
    const type    = BASE_TYPE  + i     < d.length ? d[BASE_TYPE  + i]              : 0;
    const widthPt = BASE_WIDTH + i     < d.length ? (BORDER_W_PT[d[BASE_WIDTH + i]] ?? 0.5) : 0.5;
    const color   = BASE_COLOR + i * 4 + 4 <= d.length ? colorRef(d, BASE_COLOR + i * 4) : '000000';
    borders.push({ type, widthPt, color });
  }
  let bgColor: string | undefined;
  // after attr(2) + 4 types(4) + 4 widths(4) + 4 colors(16) + diagonal(6) = offset 32
  const fOff = 32;
  if (d.length >= fOff + 8) {
    const ft = BinaryKit.readU32LE(d, fOff);
    if (ft & 1) bgColor = colorRef(d, fOff + 4);
  }
  return { borders, bgColor };
}

/* ═══════════════════════════════════════════════════════════════
   Body section parsing
   ═══════════════════════════════════════════════════════════════ */

// gsoCtx: shared mutable counter for 'gso ' drawing objects.
// Each 'gso ' CTRL_HEADER encountered increments this counter.
// objectMap is keyed by 0-based gso order = sequential BinData insertion order.
interface GsoCtx {
  count: number;
  headers?: ParaNode[];
  footers?: ParaNode[];
}

function parseBody(
  raw: Uint8Array, compressed: boolean, di: DocInfo, shield: ShieldedParser, gsoCtx: GsoCtx,
): { content: ContentNode[]; pageDims?: PageDims } {
  const recs = parseRecords(compressed ? tryInflate(raw) : raw);
  const content: ContentNode[] = [];
  let pageDims: PageDims | undefined;

  // Pre-scan for PAGE_DEF at any nesting level (real HWP stores it at level 2 inside section ctrl)
  for (const r of recs) {
    if (r.tag === TAG_PAGE_DEF) {
      pageDims = shield.guard(() => parsePageDef(r.data), A4, 'hwp:pageDef');
      break;
    }
  }

  let i = 0;
  while (i < recs.length) {
    if (recs[i].tag === TAG_PAGE_DEF) {
      i++; // already handled above; skip at top level
    } else if (recs[i].tag === TAG_PARA_HEADER) {
      const r = shield.guard(
        () => parseParagraphGroup(recs, i, di, shield, gsoCtx),
        { nodes: [] as ContentNode[], next: i + 1 },
        `hwp:para@${i}`,
      );
      content.push(...r.nodes);
      i = r.next;
    } else {
      i++;
    }
  }
  return { content, pageDims };
}

/* ── Paragraph group ────────────────────────────────────────── */

function parseParagraphGroup(
  recs: HwpRecord[], start: number, di: DocInfo, shield: ShieldedParser, gsoCtx: GsoCtx,
): { nodes: ContentNode[]; next: number } {
  const hdr = recs[start];
  const lv  = hdr.level;

  // P1: PARA_HEADER 레이아웃
  //   offset 8-9: paraShapeId (UINT16)
  //   offset 10:  styleId (UINT8)
  //   offset 11:  divideSort (UINT8) — 0x04=쪽나누기
  const psId       = hdr.data.length >= 10 ? BinaryKit.readU16LE(hdr.data, 8) : 0;
  const hwpStyleId = hdr.data.length >= 11 ? hdr.data[10] : 0;
  const divideSort = hdr.data.length >= 12 ? hdr.data[11] : 0;
  const ps         = di.paraShapes[psId];

  let text: ParaTextResult | null = null;
  let csPairs: [number, number][] = [];
  const grids: ContentNode[] = [];
  // imgId: for 'gso' uses sequential gsoCtx.count; for others uses flags-based objId
  const ctrlHeaders: { ctrlId: number; imgId: number; wPt: number; hPt: number; atnoType?: number }[] = [];
  let i = start + 1;

  while (i < recs.length && recs[i].level > lv) {
    const r = recs[i];

    if (r.tag === TAG_PARA_TEXT && r.level === lv + 1) {
      text = decodeParaText(r.data);
      i++;
    } else if (r.tag === TAG_PARA_CHAR_SHAPE && r.level === lv + 1) {
      csPairs = parseCharShapePairs(r.data);
      i++;
    } else if (r.tag === TAG_CTRL_HEADER && r.level === lv + 1) {
      if (r.data.length >= 4) {
        const ctrlId = BinaryKit.readU32LE(r.data, 0);

        if (ctrlId === CTRL_HEAD || ctrlId === CTRL_FOOT) {
          // P8: 머리말/꼬리말 컨트롤 — 자식 문단을 파싱해 gsoCtx에 저장
          const ctrlLv = r.level;
          const hfParas: ParaNode[] = [];
          let j = i + 1;
          while (j < recs.length && recs[j].level > ctrlLv) {
            if (recs[j].tag === TAG_PARA_HEADER) {
              const pr = shield.guard(
                () => parseParagraphGroup(recs, j, di, shield, gsoCtx),
                { nodes: [] as ContentNode[], next: j + 1 },
                `hwp:hf@${j}`,
              );
              hfParas.push(...pr.nodes.filter((n): n is ParaNode => n.tag === 'para'));
              j = pr.next;
            } else {
              j++;
            }
          }
          if (hfParas.length > 0) {
            const key = ctrlId === CTRL_HEAD ? 'headers' : 'footers';
            if (!gsoCtx[key]) gsoCtx[key] = hfParas;
          }
          i = j;
        } else {
          // HWP 5.0 general-object layout:
          //   [0:4] ctrlId  [4:4] flags  [8:4] xOff  [12:4] yOff
          //   [16:4] width(HWPUNIT)  [20:4] height(HWPUNIT)
          const MAX_HWP = 1_000_000;
          const rawW = r.data.length >= 24 ? BinaryKit.readU32LE(r.data, 16) : 0;
          const rawH = r.data.length >= 28 ? BinaryKit.readU32LE(r.data, 20) : 0;
          const wPt = rawW > 0 && rawW < MAX_HWP ? Metric.hwpToPt(rawW) : 0;
          const hPt = rawH > 0 && rawH < MAX_HWP ? Metric.hwpToPt(rawH) : 0;

          // P9: atno — offset 4 u32 하위 4bit = 번호 종류 (0=쪽번호, 6=전체쪽수)
          const atnoType = ctrlId === CTRL_ATNO && r.data.length >= 8
            ? BinaryKit.readU32LE(r.data, 4) & 15
            : undefined;

          // 'gso ' (그리기 객체) uses sequential counter; others use flags-based id
          const imgId = ctrlId === CTRL_GSO ? gsoCtx.count++ : (r.data.length >= 6 ? BinaryKit.readU16LE(r.data, 4) : 0);
          ctrlHeaders.push({ ctrlId, imgId, wPt, hPt, atnoType });

          if (ctrlId === CTRL_TABLE) {
            const tr = shield.guard(
              () => parseTableCtrl(recs, i, di, shield, gsoCtx),
              { grid: null, next: skipKids(recs, i) },
              `hwp:tbl@${i}`,
            );
            if (tr.grid) grids.push(tr.grid);
            i = tr.next;
          } else {
            i = skipKids(recs, i);
          }
        }
      } else {
        i = skipKids(recs, i);
      }
    } else {
      i++;
    }
  }

  const nodes: ContentNode[] = [];

  {
    const paraContent: Array<SpanNode | GridNode | PageNumNode> = [];

    // P9: atno 컨트롤 위치 수집 (pos 기준 정렬)
    const atnoCtrls: { pos: number; type: number }[] = [];
    if (text && text.controls.length > 0) {
      for (let ci = 0; ci < text.controls.length; ci++) {
        const ch = ctrlHeaders[ci];
        if (ch && ch.ctrlId === CTRL_ATNO)
          atnoCtrls.push({ pos: text.controls[ci].pos, type: ch.atnoType ?? 0 });
      }
      atnoCtrls.sort((a, b) => a.pos - b.pos);
    }

    // P9: 텍스트 chars를 atno 위치 기준으로 분할하여 PageNumNode 삽입
    if (text && text.chars.length > 0) {
      if (atnoCtrls.length > 0) {
        let k = 0;
        for (const ac of atnoCtrls) {
          const seg: ParsedChar[] = [];
          while (k < text.chars.length && text.chars[k].pos < ac.pos) seg.push(text.chars[k++]);
          if (seg.length > 0) paraContent.push(...resolveCharShapes(seg, csPairs, di));
          paraContent.push(buildPageNum(ac.type === 0 ? 'decimal' : 'total'));
        }
        const rest = text.chars.slice(k);
        if (rest.length > 0) paraContent.push(...resolveCharShapes(rest, csPairs, di));
      } else {
        paraContent.push(...resolveCharShapes(text.chars, csPairs, di));
      }
    } else if (atnoCtrls.length > 0) {
      for (const ac of atnoCtrls) paraContent.push(buildPageNum(ac.type === 0 ? 'decimal' : 'total'));
    }

    // Image placeholder spans: only for actual image controls.
    // Non-image controls (footnotes, TOC entries, etc.) are silently skipped.
    if (text && text.controls.length > 0) {
      for (let ci = 0; ci < text.controls.length; ci++) {
        const ch = ctrlHeaders[ci];
        if (!ch) continue; // anchor-only ctrl (gso is sibling, not inline)
        const isImg = ch.ctrlId === CTRL_IMAGE || ch.ctrlId === CTRL_FIG || ch.ctrlId === CTRL_OBJ || ch.ctrlId === CTRL_GSO;
        if (!isImg) continue; // skip footnotes, TOC, page num, etc.
        const dimStr = (ch.wPt > 0 && ch.hPt > 0)
          ? `_W${Math.round(ch.wPt)}_H${Math.round(ch.hPt)}`
          : '';
        paraContent.push(buildSpan(`__EXT_${ch.imgId}${dimStr}__`));
      }
    }

    // P5: 쪽나누기(divideSort & 4) → page-break 문단 먼저 출력
    if (divideSort & 4) {
      nodes.push(buildPara([{ tag: 'span', props: {}, kids: [buildPb()] } as SpanNode]));
    }
    // P5: 표 → 앵커 문단 순서 (앵커 문단 드롭 금지)
    nodes.push(...grids);
    nodes.push(buildPara(
      paraContent.length > 0 ? paraContent as any : [buildSpan('')],
      buildParaProps(ps, hwpStyleId),
    ));
  }

  return { nodes, next: i };
}

function skipKids(recs: HwpRecord[], idx: number): number {
  const lv = recs[idx].level;
  let i = idx + 1;
  while (i < recs.length && recs[i].level > lv) i++;
  return i;
}

/* ── PARA_TEXT ───────────────────────────────────────────────── */

// Extended controls: 8 WORDs, associated CTRL_HEADER (16-25 also skip 16 bytes)
const EXT_CTRL = new Set([2, 3, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25]);
// Inline controls: 8 WORDs, no CTRL_HEADER
const INL_CTRL = new Set([4, 5, 6, 7, 8]);

function decodeParaText(d: Uint8Array): ParaTextResult {
  const chars: ParsedChar[] = [];
  const controls: ParsedCtrl[] = [];
  let i = 0, pos = 0;

  while (i + 1 < d.length) {
    const c = d[i] | (d[i + 1] << 8);
    if (c === 0)  { i += 2; pos++; continue; }
    if (c === 13) { break; }                             // paragraph end
    if (c === 10) { chars.push({ pos, ch: '\n' }); i += 2; pos++; continue; }

    if (EXT_CTRL.has(c)) {
      // Extended control: 8 WORDs (16 bytes)
      // WORD 4 contains objId (for images, charts, etc.)
      let objId = 0;
      if (i + 16 <= d.length) {
        objId = BinaryKit.readU16LE(d, i + 8); // 4th WORD (offset 8) contains objId
      }
      controls.push({ pos, ctrlId: 0, objId, matched: false });
      i += 16; pos += 8; continue;
    }
    if (INL_CTRL.has(c)) {
      i += 16; pos += 8; continue;
    }
    if (c === 9) {                                        // tab (inline 8 WORDs)
      chars.push({ pos, ch: '\t' });
      i += 16; pos += 8; continue;
    }
    if (c >= 1 && c <= 31) { i += 2; pos++; continue; }  // other control

    chars.push({ pos, ch: String.fromCharCode(c) });
    i += 2; pos++;
  }
  return { chars, controls };
}

/* ── PARA_CHAR_SHAPE ────────────────────────────────────────── */

function parseCharShapePairs(d: Uint8Array): [number, number][] {
  const out: [number, number][] = [];
  for (let i = 0; i + 7 < d.length; i += 8)
    out.push([BinaryKit.readU32LE(d, i), BinaryKit.readU32LE(d, i + 4)]);
  return out;
}

/* ── Char-shape → SpanNode resolution ───────────────────────── */

function resolveCharShapes(chars: ParsedChar[], pairs: [number, number][], di: DocInfo): SpanNode[] {
  if (chars.length === 0) return [buildSpan('')];

  const defaultId = pairs.length > 0 ? pairs[0][1] : 0;

  function idFor(pos: number): number {
    let id = defaultId;
    for (const [p, sid] of pairs) { if (p <= pos) id = sid; else break; }
    return id;
  }

  const spans: SpanNode[] = [];
  let curId = idFor(chars[0].pos);
  let buf   = chars[0].ch;

  for (let k = 1; k < chars.length; k++) {
    const sid = idFor(chars[k].pos);
    if (sid !== curId) { spans.push(styledSpan(buf, curId, di)); buf = ''; curId = sid; }
    buf += chars[k].ch;
  }
  if (buf) spans.push(styledSpan(buf, curId, di));
  return spans;
}

function styledSpan(text: string, shapeId: number, di: DocInfo): SpanNode {
  const cs = di.charShapes[shapeId];
  if (!cs) return buildSpan(text);

  const props: TextProps = {};
  const fid = cs.faceIds[0] ?? 0;
  if (fid < di.faceNames.length && di.faceNames[fid]) props.font = safeFont(di.faceNames[fid]);
  if (cs.height > 0) props.pt = Metric.hwpToPt(cs.height);
  if (cs.bold)        props.b = true;
  if (cs.italic)      props.i = true;
  if (cs.underline)   props.u = true;
  if (cs.strikeout)   props.s = true;
  if (cs.superscript) props.sup = true;
  if (cs.subscript)   props.sub = true;

  const hex = safeHex(cs.textColor);
  if (hex && hex !== '000000') props.color = hex;

  return buildSpan(text, props);
}

/* ── Table control parsing ──────────────────────────────────── */

function parseTableCtrl(
  recs: HwpRecord[], ctrlIdx: number, di: DocInfo, shield: ShieldedParser, gsoCtx: GsoCtx,
): { grid: ContentNode | null; next: number } {
  const ctrlLv = recs[ctrlIdx].level;
  let i = ctrlIdx + 1;

  let tblData: Uint8Array | null = null;
  const cells: { data: Uint8Array; tag: number; cStart: number; cEnd: number }[] = [];

  // Collect TABLE and cell records within this control's scope
  const tblLevel = ctrlLv + 1;

  while (i < recs.length && recs[i].level > ctrlLv) {
    const r = recs[i];

    if (isTableTag(r.tag) && r.level === tblLevel) {
      tblData = r.data;
      i++;
    } else if (r.tag === TAG_LIST_HEADER && r.level === tblLevel) {
      // LIST_HEADER as cell: paraCount tells how many paragraphs follow
      const cellData = r.data;
      const paraCount = cellData.length >= 2 ? BinaryKit.readU16LE(cellData, 0) : 0;
      i++;
      const cStart = i;
      // Consume exactly paraCount paragraphs (each with its child records)
      let consumed = 0;
      while (i < recs.length && consumed < paraCount) {
        if (recs[i].tag === TAG_PARA_HEADER && recs[i].level === tblLevel) {
          consumed++;
          i++;
          // Skip child records of this paragraph
          while (i < recs.length && recs[i].level > tblLevel) i++;
        } else if (recs[i].level > tblLevel) {
          i++;
        } else {
          break; // hit next sibling at same level
        }
      }
      cells.push({ data: cellData, tag: TAG_LIST_HEADER, cStart, cEnd: i });
    } else if (isCellTag(r.tag) && r.level === tblLevel) {
      // Full CELL record (with cell-specific fields)
      const cellData = r.data;
      const cellTag = r.tag;
      i++;
      const cStart = i;
      while (i < recs.length && recs[i].level > tblLevel) i++;
      cells.push({ data: cellData, tag: cellTag, cStart, cEnd: i });
    } else {
      i++;
    }
  }

  if (!tblData || cells.length === 0) return { grid: null, next: i };

  const rowCnt = tblData.length >= 6 ? BinaryKit.readU16LE(tblData, 4) : 1;
  const colCnt = tblData.length >= 8 ? BinaryKit.readU16LE(tblData, 6) : 1;

  interface PC { row: number; col: number; cs: number; rs: number; widthHwp: number; heightHwp?: number; props: CellProps; cellChildren: (ParaNode | GridNode)[] }
  const parsed: PC[] = [];

  for (let ci = 0; ci < cells.length; ci++) {
    const c = cells[ci];
    const seqIdx = ci;
    const pc = shield.guard(
      () => parseCellRec(c.data, c.tag, recs, c.cStart, c.cEnd, di, shield, seqIdx, colCnt, gsoCtx),
      { row: Math.floor(ci / (colCnt || 1)), col: ci % (colCnt || 1), cs: 1, rs: 1, widthHwp: 0, heightHwp: undefined, props: {}, cellChildren: [buildPara([buildSpan('')])] },
      `hwp:cell@${c.cStart}`,
    );
    parsed.push(pc);
  }

  // Determine actual row count from cell data (may exceed rowCnt for merged cells)
  const maxRow = parsed.reduce((m, c) => Math.max(m, c.row + c.rs), 0);
  const actualRowCnt = Math.max(rowCnt, maxRow);

  // Validate cell positions; fallback to sequential layout if invalid
  const posValid = parsed.every(c => c.row >= 0 && c.col >= 0 && c.col < colCnt);
  if (!posValid) {
    let idx = 0;
    for (const c of parsed) { c.row = Math.floor(idx / colCnt); c.col = idx % colCnt; idx++; }
  }

  // Compute column widths in points from cell widths
  const colWidthsPt: number[] = new Array(colCnt).fill(0);
  // Pass 1: use cells with cs=1 for exact column widths
  for (const c of parsed) {
    if (c.cs === 1 && c.widthHwp > 0) {
      const wPt = Metric.hwpToPt(c.widthHwp);
      if (wPt > colWidthsPt[c.col]) colWidthsPt[c.col] = wPt;
    }
  }
  // Pass 2: for columns still 0, try to derive from multi-span cells
  // Sort by span size ascending so smaller, more precise spans fill widths before larger spans
  const zeroColumns = colWidthsPt.filter(w => w === 0).length;
  if (zeroColumns > 0) {
    const spanCells = parsed.filter(c => c.cs > 1 && c.widthHwp > 0).sort((a, b) => a.cs - b.cs);
    for (const c of spanCells) {
      if (c.cs > 1 && c.widthHwp > 0) {
        // Subtract known column widths from the span
        let known = 0;
        let unknownCols = 0;
        for (let ci = c.col; ci < c.col + c.cs && ci < colCnt; ci++) {
          if (colWidthsPt[ci] > 0) known += colWidthsPt[ci];
          else unknownCols++;
        }
        if (unknownCols > 0) {
          const remaining = Metric.hwpToPt(c.widthHwp) - known;
          const each = remaining > 0 ? remaining / unknownCols : 0;
          for (let ci = c.col; ci < c.col + c.cs && ci < colCnt; ci++) {
            if (colWidthsPt[ci] === 0 && each > 0) colWidthsPt[ci] = each;
          }
        }
      }
    }
  }

  // Post-process: clamp near-zero column widths (< 1pt = floating-point artifact) to minimum 1pt
  for (let i = 0; i < colWidthsPt.length; i++) {
    if (colWidthsPt[i] > 0 && colWidthsPt[i] < 1) colWidthsPt[i] = 1;
  }

  const rows = [];
  for (let r = 0; r < actualRowCnt; r++) {
    const rc = parsed.filter(c => c.row === r).sort((a, b) => a.col - b.col);
    if (rc.length === 0) continue;

    // Calculate row height — prefer rs=1 cells (exact per-row height)
    let rowHeightPt: number | undefined = undefined;
    for (const c of rc) {
      if (c.heightHwp && c.heightHwp > 0 && c.rs === 1) {
        const hPt = Metric.hwpToPt(c.heightHwp);
        if (rowHeightPt == null || hPt > rowHeightPt) rowHeightPt = hPt;
      }
    }
    // Fallback: all cells span multiple rows → approximate height per row
    if (rowHeightPt == null) {
      for (const c of rc) {
        if (c.heightHwp && c.heightHwp > 0) {
          const hPt = Metric.hwpToPt(c.heightHwp) / c.rs;
          if (rowHeightPt == null || hPt > rowHeightPt) rowHeightPt = hPt;
        }
      }
    }

    rows.push(buildRow(rc.map(c => {
      return buildCell(c.cellChildren, { cs: c.cs, rs: c.rs, props: c.props });
    }), rowHeightPt));
  }
  if (rows.length === 0) return { grid: null, next: i };

  // Table-level default stroke
  let defStroke: Stroke | undefined;
  const bfOff = 18 + rowCnt * 2;
  if (tblData.length >= bfOff + 2) {
    const bfId = BinaryKit.readU16LE(tblData, bfOff);
    defStroke = strokeFromBF(bfId, di);
  }

  const gp: GridProps = {};
  if (defStroke) gp.defaultStroke = defStroke;
  const hasWidths = colWidthsPt.some(w => w > 0);
  if (hasWidths) gp.colWidths = colWidthsPt;
  return { grid: buildGrid(rows, gp), next: i };
}

/* ── Cell record ────────────────────────────────────────────── */
/*  LIST_HEADER for cells (HWP 5.0/5.1):
    [0:2]  paraCount   [2:4]  attr (bits 6-7 = vertAlign)
    [6:2]  unknown     [8:2]  rowAddr   [10:2] colAddr
    [12:2] rowSpan     [14:2] colSpan
    [16:4] width(HWPUNIT)  [20:4] height(HWPUNIT)
    [24:8] padding[4]      [32:2] borderFillId                  */

function parseCellRec(
  d: Uint8Array, tag: number, recs: HwpRecord[], cStart: number, cEnd: number,
  di: DocInfo, shield: ShieldedParser, seqIdx: number, colCnt: number, gsoCtx: GsoCtx,
) {
  let col: number, row: number, cs = 1, rs = 1;
  let widthHwp = 0;
  let heightHwp = 0;
  const props: CellProps = {};

  const attr = d.length >= 6 ? BinaryKit.readU32LE(d, 2) : 0;
  const va = (attr >> 6) & 0x3;
  if (va === 1) props.va = 'mid';
  else if (va === 2) props.va = 'bot';

  const HWP_PAD_LR_DEFAULT = 360;
  const HWP_PAD_TB_DEFAULT = 141;

  if (tag === TAG_LIST_HEADER && d.length >= 22) {
    col = BinaryKit.readU16LE(d, 8);
    row = BinaryKit.readU16LE(d, 10);
    cs  = Math.max(1, BinaryKit.readU16LE(d, 12));
    rs  = Math.max(1, BinaryKit.readU16LE(d, 14));
    widthHwp = BinaryKit.readU32LE(d, 16);
    heightHwp = d.length >= 24 ? BinaryKit.readU32LE(d, 20) : 0;
    if (d.length >= 32) {
      const pL = BinaryKit.readU16LE(d, 24); const pR = BinaryKit.readU16LE(d, 26);
      const pT = BinaryKit.readU16LE(d, 28); const pB = BinaryKit.readU16LE(d, 30);
      if (pL !== HWP_PAD_LR_DEFAULT) props.padL = Metric.hwpToPt(pL);
      if (pR !== HWP_PAD_LR_DEFAULT) props.padR = Metric.hwpToPt(pR);
      if (pT !== HWP_PAD_TB_DEFAULT) props.padT = Metric.hwpToPt(pT);
      if (pB !== HWP_PAD_TB_DEFAULT) props.padB = Metric.hwpToPt(pB);
    }
    const bfId = d.length >= 34 ? BinaryKit.readU16LE(d, 32) : 0;
    if (bfId > 0 && bfId <= di.borderFills.length) applyCellBorderFill(di.borderFills[bfId - 1], props);
  } else if (tag !== TAG_LIST_HEADER) {
    col = d.length >= 8  ? BinaryKit.readU16LE(d, 6) : seqIdx % (colCnt || 1);
    row = d.length >= 10 ? BinaryKit.readU16LE(d, 8) : Math.floor(seqIdx / (colCnt || 1));
    cs  = d.length >= 12 ? Math.max(1, BinaryKit.readU16LE(d, 10)) : 1;
    rs  = d.length >= 14 ? Math.max(1, BinaryKit.readU16LE(d, 12)) : 1;
    widthHwp = d.length >= 18 ? BinaryKit.readU32LE(d, 14) : 0;
    heightHwp = d.length >= 22 ? BinaryKit.readU32LE(d, 18) : 0;
    if (d.length >= 30) {
      const pL = BinaryKit.readU16LE(d, 22); const pR = BinaryKit.readU16LE(d, 24);
      const pT = BinaryKit.readU16LE(d, 26); const pB = BinaryKit.readU16LE(d, 28);
      if (pL !== HWP_PAD_LR_DEFAULT) props.padL = Metric.hwpToPt(pL);
      if (pR !== HWP_PAD_LR_DEFAULT) props.padR = Metric.hwpToPt(pR);
      if (pT !== HWP_PAD_TB_DEFAULT) props.padT = Metric.hwpToPt(pT);
      if (pB !== HWP_PAD_TB_DEFAULT) props.padB = Metric.hwpToPt(pB);
    }
    const bfId = d.length >= 32 ? BinaryKit.readU16LE(d, 30) : 0;
    if (bfId > 0 && bfId <= di.borderFills.length) applyCellBorderFill(di.borderFills[bfId - 1], props);
  } else {
    row = Math.floor(seqIdx / (colCnt || 1));
    col = seqIdx % (colCnt || 1);
  }

  const cellChildren: (ParaNode | GridNode)[] = [];
  const MAX_HWP = 1_000_000;
  let k = cStart;

  while (k < cEnd) {
    if (recs[k].tag === TAG_PARA_HEADER) {
      // Parse paragraph inside cell — also extracts nested tables within the paragraph
      const r = shield.guard(
        () => {
          const hdr = recs[k];
          const lv = hdr.level;
          const psId = hdr.data.length >= 10 ? BinaryKit.readU16LE(hdr.data, 8) : 0;
          // P6: 셀 내부 문단의 styleId / divideSort 읽기
          const cellStyleId = hdr.data.length >= 11 ? hdr.data[10] : 0;
          const cellDivide  = hdr.data.length >= 12 ? hdr.data[11] : 0;
          const ps = di.paraShapes[psId];
          let txt: ParaTextResult | null = null;
          let csp: [number, number][] = [];
          const ctrlHdrs: { ctrlId: number; imgId: number; wPt: number; hPt: number }[] = [];
          const innerGrids: GridNode[] = [];
          let j = k + 1;
          while (j < cEnd && recs[j].level > lv) {
            if (recs[j].tag === TAG_PARA_TEXT) { txt = decodeParaText(recs[j].data); j++; }
            else if (recs[j].tag === TAG_PARA_CHAR_SHAPE) { csp = parseCharShapePairs(recs[j].data); j++; }
            else if (recs[j].tag === TAG_CTRL_HEADER && recs[j].level === lv + 1) {
              if (recs[j].data.length >= 4) {
                const ctrlId = BinaryKit.readU32LE(recs[j].data, 0);
                if (ctrlId === CTRL_TABLE) {
                  // Nested table inside a cell paragraph — recurse into parseTableCtrl
                  const nestedTr = shield.guard(
                    () => parseTableCtrl(recs, j, di, shield, gsoCtx),
                    { grid: null, next: skipKids(recs, j) },
                    `hwp:innerNestedTbl@${j}`,
                  );
                  if (nestedTr.grid) innerGrids.push(nestedTr.grid as GridNode);
                  j = nestedTr.next;
                } else {
                  const rawW = recs[j].data.length >= 24 ? BinaryKit.readU32LE(recs[j].data, 16) : 0;
                  const rawH = recs[j].data.length >= 28 ? BinaryKit.readU32LE(recs[j].data, 20) : 0;
                  const wPt = rawW > 0 && rawW < MAX_HWP ? Metric.hwpToPt(rawW) : 0;
                  const hPt = rawH > 0 && rawH < MAX_HWP ? Metric.hwpToPt(rawH) : 0;
                  const imgId = ctrlId === CTRL_GSO ? gsoCtx.count++ : (recs[j].data.length >= 6 ? BinaryKit.readU16LE(recs[j].data, 4) : 0);
                  ctrlHdrs.push({ ctrlId, imgId, wPt, hPt });
                  j = skipKids(recs, j);
                }
              } else {
                j = skipKids(recs, j);
              }
            }
            else j++;
          }
          const paraContent: (SpanNode | ContentNode)[] = [];
          if (txt && txt.chars.length > 0) paraContent.push(...resolveCharShapes(txt.chars, csp, di));
          if (txt && txt.controls.length > 0) {
            for (let ci = 0; ci < txt.controls.length; ci++) {
              const ch = ctrlHdrs[ci];
              if (!ch) continue;
              const isImg = ch.ctrlId === CTRL_IMAGE || ch.ctrlId === CTRL_FIG || ch.ctrlId === CTRL_OBJ || ch.ctrlId === CTRL_GSO;
              if (!isImg) continue;
              const dimStr = (ch.wPt > 0 && ch.hPt > 0) ? `_W${Math.round(ch.wPt)}_H${Math.round(ch.hPt)}` : '';
              paraContent.push(buildSpan(`__EXT_${ch.imgId}${dimStr}__`));
            }
          }
          const kids = paraContent.length > 0 ? paraContent as any : [buildSpan('')];
          // P6: innerGrids 먼저, 앵커 문단 나중 (P5와 동일한 순서)
          const items: (ParaNode | GridNode)[] = [...innerGrids, buildPara(kids, buildParaProps(ps, cellStyleId))];
          if (cellDivide & 4) items.unshift(buildPara([{ tag: 'span', props: {}, kids: [buildPb()] } as SpanNode]));
          return { items, next: j };
        },
        { items: [buildPara([buildSpan('')])] as (ParaNode | GridNode)[], next: k + 1 },
        `hwp:cellP@${k}`,
      );
      cellChildren.push(...r.items);
      k = r.next;
    } else if (recs[k].tag === TAG_CTRL_HEADER && recs[k].data.length >= 4) {
      // CTRL_HEADER at cell level (sibling of PARA_HEADER) — anchored 'gso' images and outer-level nested tables
      const cellCtrlId = BinaryKit.readU32LE(recs[k].data, 0);
      if (cellCtrlId === CTRL_GSO) {
        const gsoId = gsoCtx.count++;
        const rawW = recs[k].data.length >= 24 ? BinaryKit.readU32LE(recs[k].data, 16) : 0;
        const rawH = recs[k].data.length >= 28 ? BinaryKit.readU32LE(recs[k].data, 20) : 0;
        const wPt = rawW > 0 && rawW < MAX_HWP ? Metric.hwpToPt(rawW) : 0;
        const hPt = rawH > 0 && rawH < MAX_HWP ? Metric.hwpToPt(rawH) : 0;
        const dimStr = (wPt > 0 && hPt > 0) ? `_W${Math.round(wPt)}_H${Math.round(hPt)}` : '';
        cellChildren.push(buildPara([buildSpan(`__EXT_${gsoId}${dimStr}__`)]));
        k = skipKids(recs, k);
      } else if (cellCtrlId === CTRL_TABLE) {
        const tr = shield.guard(
          () => parseTableCtrl(recs, k, di, shield, gsoCtx),
          { grid: null, next: skipKids(recs, k) },
          `hwp:nestedTbl@${k}`,
        );
        if (tr.grid) cellChildren.push(tr.grid as GridNode);
        k = tr.next;
      } else {
        k = skipKids(recs, k);
      }
    } else { k++; }
  }

  return {
    row, col, cs, rs, props, widthHwp,
    heightHwp: heightHwp || undefined,
    cellChildren: cellChildren.length ? cellChildren : [buildPara([buildSpan('')])],
  };
}

/* ── PAGE_DEF ───────────────────────────────────────────────── */
/*  [0:4] width  [4:4] height  [8:4] ml  [12:4] mr
    [16:4] mt  [20:4] mb  [24:4] header  [28:4] footer  [36:4] attr (bit0=landscape) */

function parsePageDef(d: Uint8Array): PageDims {
  if (d.length < 24) return A4;
  const w  = BinaryKit.readU32LE(d, 0);
  const h  = BinaryKit.readU32LE(d, 4);
  const ml = BinaryKit.readU32LE(d, 8);
  const mr = BinaryKit.readU32LE(d, 12);
  const mt = BinaryKit.readU32LE(d, 16);
  const mb = BinaryKit.readU32LE(d, 20);
  const header = d.length >= 28 ? BinaryKit.readU32LE(d, 24) : 0;
  const footer = d.length >= 32 ? BinaryKit.readU32LE(d, 28) : 0;
  const at = d.length >= 40 ? BinaryKit.readU32LE(d, 36) : 0;
  return {
    wPt: Metric.hwpToPt(w),  hPt: Metric.hwpToPt(h),
    ml: Metric.hwpToPt(ml),  mr: Metric.hwpToPt(mr),
    mt: Metric.hwpToPt(mt),  mb: Metric.hwpToPt(mb),
    headerPt: header > 0 ? Metric.hwpToPt(header) : undefined,
    footerPt: footer > 0 ? Metric.hwpToPt(footer) : undefined,
    orient: (at & 1) ? 'landscape' : 'portrait',
  };
}

/* ═══════════════════════════════════════════════════════════════
   Helpers
   ═══════════════════════════════════════════════════════════════ */

function i32(d: Uint8Array, o: number): number {
  const u = BinaryKit.readU32LE(d, o);
  return u > 0x7FFFFFFF ? u - 0x100000000 : u;
}

function colorRef(d: Uint8Array, o: number): string {
  if (o + 3 > d.length) return '000000';
  return ((d[o] << 16) | (d[o + 1] << 8) | d[o + 2]).toString(16).padStart(6, '0').toUpperCase();
}

function toStroke(b: { type: number; widthPt: number; color: string }): Stroke {
  return { kind: BORDER_KIND[b.type] ?? 'solid', pt: b.widthPt, color: b.color };
}

// Apply borderFill to CellProps. Preserve explicit NONE so DOCX tcBorders can
// override the table-level tblBorders. Filtering NONE would let tblBorders bleed through.
function applyCellBorderFill(bf: HwpBorderFill, props: CellProps): void {
  if (bf.borders.length >= 4) {
    props.left  = toStroke(bf.borders[0]);
    props.right = toStroke(bf.borders[1]);
    props.top   = toStroke(bf.borders[2]);
    props.bot   = toStroke(bf.borders[3]);
  }
  if (bf.bgColor && bf.bgColor !== 'FFFFFF') props.bg = bf.bgColor;
}

function strokeFromBF(bfId: number, di: DocInfo): Stroke | undefined {
  if (bfId <= 0 || bfId > di.borderFills.length) return undefined;
  const bf = di.borderFills[bfId - 1];
  if (!bf.borders.length) return undefined;
  const b = bf.borders[0];
  return { kind: BORDER_KIND[b.type] ?? 'solid', pt: b.widthPt, color: b.color };
}

function buildParaProps(ps?: HwpParaShape, hwpStyleId?: number): ParaProps {
  // P2: hwpStyleId를 초기값으로 포함 (undefined이면 빈 객체)
  const p: ParaProps = hwpStyleId !== undefined ? { hwpStyleId } : {};
  if (!ps) return p;
  if (ps.align && ps.align !== 'left') p.align = ps.align;
  if (ps.spaceBefore > 0) p.spaceBefore = Metric.hwpToPt(ps.spaceBefore);
  if (ps.spaceAfter > 0)  p.spaceAfter  = Metric.hwpToPt(ps.spaceAfter);
  // 줄 간격: type=0(PERCENT) → lineHeight, type=1(FIXED) → lineHeightFixed
  if (ps.lineSpacingType === 1) {
    if (ps.lineSpacing > 0) p.lineHeightFixed = Metric.hwpToPt(ps.lineSpacing);
  } else {
    // P10: 160%(HWP 기본값) 생략 버그 수정 — 항상 lineHeight 설정
    if (ps.lineSpacing > 0) p.lineHeight = ps.lineSpacing / 100;
  }
  // leftMargin (offset 4) = 문단 몸체 왼쪽 여백 → leftMargin (pt), ensure non-negative
  const leftMarginPt = Math.max(0, Metric.hwpToPt(ps.leftMargin));
  if (leftMarginPt > 0) p.leftMargin = leftMarginPt;
  // rightMargin (offset 8) = 문단 몸체 오른쪽 여백 → indentRightPt (pt)
  const rightMarginPt = Math.max(0, Metric.hwpToPt(ps.rightMargin));
  if (rightMarginPt > 0) p.indentRightPt = rightMarginPt;
  // indent (offset 12) = 첫 줄 들여쓰기(양수) / 내어쓰기(음수) → firstLineIndentPt
  if (ps.indent !== 0) p.firstLineIndentPt = Metric.hwpToPt(ps.indent);
  if (ps.verAlign && ps.verAlign !== 'baseline') p.verAlign = ps.verAlign;
  if (ps.lineWrap && ps.lineWrap !== 'break') p.lineWrap = ps.lineWrap;
  return p;
}

/* ═══════════════════════════════════════════════════════════════
   Decoder class
   ═══════════════════════════════════════════════════════════════ */

export class HwpScanner implements Decoder {
  readonly format = 'hwp';
  readonly aliases = ['application/vnd.hancom.hwp'];

  async decode(data: Uint8Array): Promise<Outcome<DocRoot>> {
    const shield = new ShieldedParser();
    const warns: string[] = [];

    try {
      if (!BinaryKit.isOle2(data)) return fail('HWP: Invalid OLE2 signature');
      const streams = BinaryKit.parseCfb(data);

      // FileHeader
      const fh = streams.get('FileHeader');
      const { compressed, encrypted } = fh ? parseFileHeader(fh) : { compressed: true, encrypted: false };
      if (encrypted) return fail('HWP: 암호화된 파일은 지원하지 않습니다');

      // DocInfo
      const diRaw = streams.get('DocInfo');
      let di: DocInfo = { faceNames: [], charShapes: [], paraShapes: [], borderFills: [] };
      if (diRaw) {
        di = shield.guard(() => parseDocInfo(diRaw, compressed), di, 'hwp:docInfo');
      }

      // Extract images from BinData streams.
      // HWP duplicates each BinData entry: once as "BinData/BIN0001.jpg" and once as "BIN0001.jpg".
      // We keep only the "BinData/" prefixed versions, sort by BIN number, then assign 0-based keys
      // matching the order 'gso' CTRL_HEADER records are encountered during body parsing.
      const binEntries: { binNum: number; data: Uint8Array }[] = [];
      for (const [path, streamData] of streams) {
        // Match "BinData/BIN0001.jpg" style — the canonical form
        const m = path.match(/^BinData[/\\]BIN(\d+)\.\w+$/i);
        if (m) binEntries.push({ binNum: parseInt(m[1], 10), data: streamData });
      }
      // Sort by BIN number (ascending) so BIN0001→idx0, BIN0002→idx1, …
      binEntries.sort((a, b) => a.binNum - b.binNum);

      const objectMap = new Map<number, ImgNode>();
      for (let idx = 0; idx < binEntries.length; idx++) {
        const { data: imgData } = binEntries[idx];

        // Determine MIME type from binary signature first, then fall back to extension
        let mimeType: ImgNode['mime'] = 'image/jpeg';
        if (imgData[0] === 0x89 && imgData[1] === 0x50) mimeType = 'image/png';
        else if (imgData[0] === 0x47 && imgData[1] === 0x49) mimeType = 'image/gif';
        else if (imgData[0] === 0x42 && imgData[1] === 0x4D) mimeType = 'image/bmp';

        const base64 = TextKit.base64Encode(imgData);
        const { wPt, hPt } = getImageDimsPt(imgData, mimeType);
        objectMap.set(idx, buildImg(base64, mimeType, wPt, hPt));
      }

      // gsoCtx tracks sequential 'gso' encounter order — must be shared across all sections
      const gsoCtx: GsoCtx = { count: 0 };

      // Body sections
      const allContent: ContentNode[] = [];
      let pageDims: PageDims = A4;

      for (let s = 0; s < 100; s++) {
        const sec = streams.get(`BodyText/Section${s}`) ?? streams.get(`Section${s}`);
        if (!sec) {
          if (s === 0) {
            const fb = findBodySection(streams);
            if (fb) {
              const r = parseBody(fb, compressed, di, shield, gsoCtx);
              allContent.push(...r.content);
              if (r.pageDims) pageDims = r.pageDims;
            }
          }
          break;
        }
        const r = shield.guard(
          () => parseBody(sec, compressed, di, shield, gsoCtx),
          { content: [], pageDims: undefined },
          `hwp:sec${s}`,
        );
        allContent.push(...r.content);
        if (r.pageDims) pageDims = r.pageDims;
      }

      if (objectMap.size > 0) {
        injectImagesIntoContent(allContent, objectMap);
      }

      warns.push(...shield.flush());
      const content = allContent.length > 0 ? allContent : [buildPara([buildSpan('')])];
      // P8: 머리말/꼬리말을 gsoCtx에서 가져와 buildSheet에 전달
      return succeed(buildRoot({}, [buildSheet(content, pageDims, {
        headers: gsoCtx.headers ? { default: gsoCtx.headers } : undefined,
        footers: gsoCtx.footers ? { default: gsoCtx.footers } : undefined,
      })]), warns);
    } catch (e: any) {
      warns.push(...shield.flush());
      return fail(`HWP decode error: ${e?.message ?? String(e)}`, warns);
    }
  }
}

function findBodySection(streams: Map<string, Uint8Array>): Uint8Array | undefined {
  for (const [k, v] of streams)
    if (k.includes('Section') && !k.includes('Header') && !k.includes('Info')) return v;
  return undefined;
}

/* ═══════════════════════════════════════════════════════════════
   Image dimension extraction from binary headers
   ════════════════════════════════════════════════════════════ */

// Returns { wPt, hPt } by parsing image headers; falls back to { wPt: 72, hPt: 72 } (1-inch)
function getImageDimsPt(data: Uint8Array, mime: string): { wPt: number; hPt: number } {
  const fallback = { wPt: 72, hPt: 72 };
  try {
    if (mime === 'image/png' && data.length >= 24) {
      // PNG IHDR: sig(8) + length(4) + type(4) + width(4) + height(4) — all big-endian
      const w = (data[16] << 24 | data[17] << 16 | data[18] << 8 | data[19]) >>> 0;
      const h = (data[20] << 24 | data[21] << 16 | data[22] << 8 | data[23]) >>> 0;
      if (w > 0 && h > 0) return { wPt: w * 0.75, hPt: h * 0.75 }; // 96 DPI → pt
    }
    if (mime === 'image/jpeg') {
      // Scan for SOF markers: FF C0 / C1 / C2 / C3
      let i = 2;
      while (i + 8 < data.length) {
        if (data[i] !== 0xFF) { i++; continue; }
        const marker = data[i + 1];
        if (marker >= 0xC0 && marker <= 0xC3) {
          // SOF: 2-byte marker + 2-byte length + 1-byte precision + 2-byte height + 2-byte width
          const h = (data[i + 5] << 8 | data[i + 6]) >>> 0;
          const w = (data[i + 7] << 8 | data[i + 8]) >>> 0;
          if (w > 0 && h > 0) return { wPt: w * 0.75, hPt: h * 0.75 };
        }
        const segLen = data[i + 2] << 8 | data[i + 3];
        i += 2 + (segLen > 0 ? segLen : 2);
      }
    }
    if (mime === 'image/bmp' && data.length >= 26) {
      // BMP DIB header: width at 18, height at 22 (signed int32 LE; negative = top-down)
      const w = BinaryKit.readU32LE(data, 18);
      const h = Math.abs(BinaryKit.readU32LE(data, 22) | 0);
      if (w > 0 && h > 0) return { wPt: w * 0.75, hPt: h * 0.75 };
    }
    if (mime === 'image/gif' && data.length >= 10) {
      // GIF: width at 6, height at 8 (uint16 LE)
      const w = data[6] | data[7] << 8;
      const h = data[8] | data[9] << 8;
      if (w > 0 && h > 0) return { wPt: w * 0.75, hPt: h * 0.75 };
    }
  } catch { /* ignore */ }
  return fallback;
}

/* ═══════════════════════════════════════════════════════════════
   OLE Object extraction (images)
   ════════════════════════════════════════════════════════════ */

function extractImagesFromOleObjectLink(data: Uint8Array): OleObject[] {
  const objects: OleObject[] = [];
  let off = 0;

  while (off + 8 <= data.length) {
    const objId = BinaryKit.readU32LE(data, off);
    const dataSize = BinaryKit.readU32LE(data, off + 4);
    const reserved = BinaryKit.readU32LE(data, off + 8);

    if (objId === 0 || dataSize === 0) break;

    const objOff = off + 16;
    if (objOff + dataSize > data.length) break;

    const objData = data.subarray(objOff, objOff + dataSize);

    // Detect MIME type from signature
    let mimeType = 'application/octet-stream';
    if (objData[0] === 0xFF && objData[1] === 0xD8 && objData[2] === 0xFF) {
      mimeType = 'image/jpeg';
    } else if (objData[0] === 0x89 && objData[1] === 0x50 && objData[2] === 0x4E && objData[3] === 0x47) {
      mimeType = 'image/png';
    } else if (objData[0] === 0x47 && objData[1] === 0x49 && objData[2] === 0x46 && objData[3] === 0x3538) {
      mimeType = 'image/gif';
    } else if (objData[0] === 0x42 && objData[1] === 0x4D) {
      mimeType = 'image/bmp';
    }

    objects.push({ id: objId, data: objData, mimeType });
    off = objOff + dataSize;
  }

  return objects;
}

/* ═══════════════════════════════════════════════════════════════
   Helper to inject images into paragraph content
   ════════════════════════════════════════════════════════════ */

function injectImagesIntoContent(
  content: ContentNode[],
  objectMap: Map<number, ImgNode>
): void {
  if (objectMap.size === 0) return;

  // Helper function to process a list of kids (spans, images, etc.)
  const processKids = (kids: any[]) => {
    for (let i = 0; i < kids.length; i++) {
      const kid = kids[i];
      // Span node structure: { tag: 'span', props, kids: [{ tag: 'txt', content }] }
      if (kid.tag === 'span' && kid.kids && kid.kids[0]?.tag === 'txt') {
        const text = kid.kids[0].content;
        // __EXT_N__ or __EXT_N_W<wPt>_H<hPt>__ (with encoded display size)
        // N is the objId that matches the index in objectMap
        const match = text.match?.(/^__(?:IMG|EXT)_(\d+)(?:_W(\d+)_H(\d+))?__$/);
        if (match) {
          const objId = parseInt(match[1], 10);
          const base = objectMap.get(objId);
          if (base) {
            const wPt = match[2] ? parseInt(match[2], 10) : 0;
            const hPt = match[3] ? parseInt(match[3], 10) : 0;
            // Use encoded display size when valid; otherwise keep pixel-based dims
            kids[i] = (wPt > 0 && hPt > 0) ? { ...base, w: wPt, h: hPt } : base;
          }
        }
      }
    }
  };

  // Recursively process a grid (table): resolves image placeholders in all cells,
  // including nested grids inside cells.
  const processGridKids = (grid: any) => {
    if (!grid.kids || !Array.isArray(grid.kids)) return;

    for (const row of grid.kids) {
      if (!row.kids || !Array.isArray(row.kids)) continue;

      for (const cell of row.kids) {
        if (!cell.kids || !Array.isArray(cell.kids)) continue;

        for (const cellKid of cell.kids) {
          if (cellKid.tag === 'grid') {
            // Nested table inside cell — recurse
            processGridKids(cellKid);
          } else if (cellKid.tag === 'para' && cellKid.kids) {
            processKids(cellKid.kids);
          }
        }
      }
    }
  };

  for (const node of content) {
    if (node.tag === 'para' && node.kids) {
      // Process paragraph kids (spans, images, links, grids)
      processKids(node.kids);

      // Also process any nested grids inside the paragraph
      for (const kid of node.kids) {
        if (kid.tag === 'grid') {
          processGridKids(kid);
        }
      }
    } else if (node.tag === 'grid') {
      // Process grid nodes (tables)
      processGridKids(node);
    }
  }
}

registry.registerDecoder(new HwpScanner());
