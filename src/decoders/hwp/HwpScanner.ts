import type { Decoder } from '../../contract/decoder';
import type { DocRoot, ContentNode, ParaNode, SpanNode, ImgNode, GridNode, PageNumNode } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import type { Align, Stroke, StrokeKind, PageDims, TextProps, ParaProps, CellProps, GridProps, ImgLayout, Heading } from '../../model/doc-props';
import { succeed, fail } from '../../contract/result';
import { buildRoot, buildSheet, buildPara, buildSpan, buildGrid, buildRow, buildCell, buildImg, buildPb, buildPageNum } from '../../model/builders';
import { ShieldedParser } from '../../safety/ShieldedParser';
import { BinaryKit } from '../../toolkit/BinaryKit';
import { TextKit } from '../../toolkit/TextKit';
import { Metric, safeHex, safeFont } from '../../safety/StyleBridge';
import { registry } from '../../pipeline/registry';
import { A4 } from '../../model/doc-props';
import { inferColumnWidths } from '../../toolkit/TableGeometry';
import pako from 'pako';

/* ═══════════════════════════════════════════════════════════════
   HWP 5.0 Tag Constants
   ═══════════════════════════════════════════════════════════════ */

const HWPTAG_BEGIN = 16;

const TAG_FACE_NAME       = HWPTAG_BEGIN + 3;   // 19
const TAG_BORDER_FILL     = HWPTAG_BEGIN + 4;   // 20
const TAG_CHAR_SHAPE      = HWPTAG_BEGIN + 5;   // 21
const TAG_NUMBERING       = HWPTAG_BEGIN + 7;   // 23
const TAG_BULLET          = HWPTAG_BEGIN + 8;   // 24
const TAG_PARA_SHAPE      = HWPTAG_BEGIN + 9;   // 25
const TAG_STYLE           = HWPTAG_BEGIN + 10;  // 26
const TAG_PARA_HEADER     = HWPTAG_BEGIN + 50;  // 66
const TAG_PARA_TEXT       = HWPTAG_BEGIN + 51;  // 67
const TAG_PARA_CHAR_SHAPE = HWPTAG_BEGIN + 52;  // 68
const TAG_CTRL_HEADER     = HWPTAG_BEGIN + 55;  // 71
const TAG_PAGE_DEF        = HWPTAG_BEGIN + 57;  // 73
const TAG_SHAPE_COMPONENT_PICTURE = HWPTAG_BEGIN + 69; // 85

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
const CTRL_PIC   = 0x24706963;  // '$pic' = picture object
const CTRL_OBJ   = 0x6F626A20;  // 'obj '
const CTRL_FIG   = 0x66696720;  // 'fig '
const CTRL_GSO   = 0x67736F20;  // 'gso ' = 그리기 객체 (drawing object, contains embedded images)
const CTRL_HEAD  = 0x68656164;  // 'head' = 머리말
const CTRL_FOOT  = 0x666F6F74;  // 'foot' = 꼬리말
const CTRL_ATNO  = 0x61746E6F;  // 'atno' = 자동 번호 (쪽번호 등)
const CTRL_SECD  = 0x73656364;  // 'secd' = 구역 정의

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
  pageBreakBefore?: boolean;
  keepWithNext?: boolean;
  keepLines?: boolean;
  widowControl?: boolean;
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
  heading?: Heading;
  listOrd?: boolean;
  listLevel?: number;
  listId?: number;
}
interface HwpStyle {
  name: string;
  engName: string;
  paraShapeId: number;
  charShapeId: number;
}
interface HwpNumbering {
  formats: string[];
}
interface HwpBullet {
  character: string;
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
  styles: HwpStyle[];
  numberings: HwpNumbering[];
  bullets: HwpBullet[];
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
  // Pako can return undefined for an incomplete stream without throwing.
  // Only accept actual bytes, and still try raw DEFLATE before falling back.
  for (const inflate of [pako.inflate, pako.inflateRaw]) {
    try {
      const result = inflate(data);
      if (result instanceof Uint8Array) return result;
    } catch { /* try the next representation */ }
  }
  return data;
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
  const info: DocInfo = {
    faceNames: [],
    charShapes: [],
    paraShapes: [],
    borderFills: [],
    styles: [],
    numberings: [],
    bullets: [],
  };

  for (const r of recs) {
    try {
      if (r.tag === TAG_FACE_NAME)   info.faceNames.push(parseFaceName(r.data));
      if (r.tag === TAG_CHAR_SHAPE)  info.charShapes.push(parseCharShape(r.data));
      if (r.tag === TAG_PARA_SHAPE)  info.paraShapes.push(parseParaShape(r.data));
      if (r.tag === TAG_BORDER_FILL) info.borderFills.push(parseBorderFill(r.data));
      if (r.tag === TAG_STYLE)       info.styles.push(parseStyle(r.data));
      if (r.tag === TAG_NUMBERING)   info.numberings.push(parseNumbering(r.data));
      if (r.tag === TAG_BULLET)      info.bullets.push(parseBullet(r.data));
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

function parseStyle(d: Uint8Array): HwpStyle {
  let offset = 0;
  const readName = (): string => {
    if (offset + 2 > d.length) throw new Error('truncated STYLE name length');
    const length = BinaryKit.readU16LE(d, offset);
    offset += 2;
    const end = offset + length * 2;
    if (end > d.length) throw new Error('truncated STYLE name');
    const value = new TextDecoder('utf-16le').decode(d.subarray(offset, end));
    offset = end;
    return value;
  };
  const name = readName();
  const engName = readName();
  if (offset + 8 > d.length) throw new Error('truncated STYLE fields');
  offset += 4; // type, nextStyleId, languageId
  const paraShapeId = BinaryKit.readU16LE(d, offset);
  const charShapeId = BinaryKit.readU16LE(d, offset + 2);
  return { name, engName, paraShapeId, charShapeId };
}

function parseNumbering(d: Uint8Array): HwpNumbering {
  const formats: string[] = [];
  let offset = 0;
  for (let level = 0; level < 7; level++) {
    if (offset + 14 > d.length) throw new Error('truncated NUMBERING level');
    offset += 12; // 문단 머리 정보
    const length = BinaryKit.readU16LE(d, offset);
    offset += 2;
    const end = offset + length * 2;
    if (end > d.length) throw new Error('truncated NUMBERING format');
    formats.push(
      new TextDecoder('utf-16le').decode(d.subarray(offset, end)),
    );
    offset = end;
  }
  return { formats };
}

function parseBullet(d: Uint8Array): HwpBullet {
  if (d.length < 10) throw new Error('truncated BULLET record');
  // Conforming HWP 5.x records include the paragraph-head charShapeId at
  // offset 8, followed by the bullet character at offset 12. Keep the old
  // compact hwpkit layout readable for files emitted before that fix.
  const characterOffset = d.length >= 23 ? 12 : 8;
  return { character: String.fromCharCode(BinaryKit.readU16LE(d, characterOffset)) };
}

/* ── CHAR_SHAPE ─────────────────────────────────────────────── */
/*  offset  size  field
    0       14    faceId[7]   (UINT16 × 7)
    14       7    ratio[7]
    21       7    spacing[7]
    28       7    relSize[7]
    35       7    offset[7]
    42       4    height      (UINT32, HWP-units 100 = 1pt)
    46       4    attr/textColor

    Hancom-authored HWP files commonly store text-color bytes at offset 46,
    while hwpkit-authored files store the compact style bitfield there. Keep
    bold/italic compatibility, but only accept underline/strike from unambiguous
    compact flags so color bytes do not become bogus decorations. */

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
  const compactStyleFlags = (attr & 0xFF000000) === 0;
  const suType  = (attr >> 16) & 0x3;   // 2 bits at 16-17 (0=none,1=super,2=sub)

  return {
    faceIds,
    height: (height > 0 && height < 100000) ? height : 1000,
    italic:      (attr & 1) !== 0,
    bold:        ((attr >> 1) & 1) !== 0,
    underline:   compactStyleFlags && (attr & (1 << 2)) !== 0,
    strikeout:   compactStyleFlags && ((attr >> 18) & 0x7) !== 0,
    superscript: suType === 1,
    subscript:   suType === 2,
    textColor:   d.length >= 56 ? colorRef(d, 52) : '000000',
  };
}

/* ── PARA_SHAPE ─────────────────────────────────────────────── */
/*  offset  size  field
    0       4     attr1   (bits 0-1 = line spacing type, bits 2-4 = alignment)
    4       4     leftMargin   (HWPUNIT * 2)
    8       4     rightMargin
    12      4     indent
    16      4     spaceBefore
    20      4     spaceAfter
    24      4     lineSpacing                                         */

const ALIGN_TBL: Record<number, Align> = { 0: 'justify', 1: 'left', 2: 'right', 3: 'center', 4: 'distribute', 5: 'distribute_space' };

function parseParaShape(d: Uint8Array): HwpParaShape {
  if (d.length < 4) return { align: 'justify', spaceBefore: 0, spaceAfter: 0, lineSpacing: 160, lineSpacingType: 0, leftMargin: 0, rightMargin: 0, indent: 0 };
  const attr = BinaryKit.readU32LE(d, 0);

  // HWP 5.0.2.5+ stores the active type/value in attr3 (offset 46) and
  // lineSpacing2 (offset 50).  Older files use attr1 and offset 24.
  const legacyLineSpacingType = (attr & 0x3) as 0 | 1 | 2 | 3;
  const extendedLineSpacingType =
    d.length >= 54 ? BinaryKit.readU32LE(d, 46) & 0x1f : -1;
  const lineSpacingType =
    extendedLineSpacingType >= 0 && extendedLineSpacingType <= 3
      ? extendedLineSpacingType as 0 | 1 | 2 | 3
      : legacyLineSpacingType;
  const lineSpacing =
    d.length >= 54
      ? BinaryKit.readU32LE(d, 50)
      : d.length >= 28
        ? i32(d, 24)
        : 160;

  // bits 2-4: 정렬 방식 (0=justify,1=left,2=right,3=center,4=distribute,5=split)
  const align = ALIGN_TBL[(attr >> 2) & 0x7] ?? 'justify';

  // 세로 정렬 (Bit 20 ~ Bit 21)
  const vVal = (attr >> 20) & 0x3;
  const verAlign = vVal === 1 ? 'top' : vVal === 2 ? 'center' : vVal === 3 ? 'bottom' : 'baseline';

  // 줄 바꿈 기준: attr1 에는 별도 비트 없음, 기본값 'break'
  const lineWrap: 'break' = 'break';
  const headingType = (attr >>> 23) & 0x3;
  const headingLevel = (attr >>> 25) & 0x7;
  const heading = headingType === 1 && headingLevel < 6
    ? (headingLevel + 1) as Heading
    : undefined;
  const listOrd = headingType === 2
    ? true
    : headingType === 3
      ? false
      : undefined;
  const listId = d.length >= 32 ? BinaryKit.readU16LE(d, 30) : 0;

  return {
    align,
    pageBreakBefore: (attr & (1 << 19)) !== 0,
    keepWithNext: (attr & (1 << 17)) !== 0,
    keepLines: (attr & (1 << 18)) !== 0,
    widowControl: (attr & (1 << 16)) !== 0,
    lineSpacingType,
    leftMargin:  d.length >= 8  ? i32(d, 4)  : 0,  // offset 4: 문단 몸체 왼쪽 여백 (HWPUNIT * 2)
    rightMargin: d.length >= 12 ? i32(d, 8)  : 0,  // offset 8: 문단 몸체 오른쪽 여백 (HWPUNIT * 2)
    indent:      d.length >= 16 ? i32(d, 12) : 0,  // offset 12: 첫 줄 들여쓰기 (HWPUNIT * 2)
    spaceBefore: d.length >= 20 ? i32(d, 16) : 0,
    spaceAfter:  d.length >= 24 ? i32(d, 20) : 0,
    lineSpacing,
    verAlign,
    lineWrap,
    heading,
    listOrd,
    listLevel: listOrd === undefined ? undefined : headingLevel,
    listId: listOrd === undefined ? undefined : listId,
  };
}

/* ── BORDER_FILL ────────────────────────────────────────────── */
/*  [0:2]  attr
    For each of 5 borders (left,right,top,bottom,diagonal): 6 bytes
      +0 type(BYTE)  +1 widthIdx(BYTE)  +2 color(COLORREF)
    [32:4] fillType
    [36:4] faceColor (bgColor for solid fill)                        */

const BORDER_W_PT = [0.28, 0.34, 0.43, 0.57, 0.71, 0.85, 1.13, 1.42, 1.70, 1.98, 2.84, 4.25, 5.67, 8.50, 11.34, 14.17];
const BORDER_KIND: Record<number, StrokeKind> = { 0:'none',1:'solid',2:'dash',3:'dot',4:'dash',5:'dash',6:'dash',7:'double',8:'double',9:'double',10:'none' };

function parseBorderFill(d: Uint8Array): HwpBorderFill {
  // HWP 5 BorderFill:
  //   [0:2] attr
  //   [2:8] left   border: type(1), width(1), colorRef(4)
  //   [8:14] right border
  //   [14:20] top border
  //   [20:26] bottom border
  //   [26:32] diagonal border
  //   [32:4] fillType, followed by fill data
  const borders: HwpBorderFill['borders'] = [];
  for (let i = 0; i < 4; i++) {
    const off = 2 + i * 6;
    const type = off < d.length ? d[off] : 0;
    const widthPt = off + 1 < d.length ? (BORDER_W_PT[d[off + 1]] ?? 0.5) : 0.5;
    const color = off + 6 <= d.length ? colorRef(d, off + 2) : '000000';
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
  objects: Map<number, { wPt: number; hPt: number; layout?: ImgLayout; binIndex?: number }>;
  headers?: ParaNode[];
  footers?: ParaNode[];
}

/** Decode HWP 5 common object properties (spec table 69/70). */
function parseObjectLayout(data: Uint8Array): ImgLayout | undefined {
  if (data.length < 28) return undefined;
  const flags = BinaryKit.readU32LE(data, 4);
  if ((flags & 1) !== 0) return { wrap: 'inline' };

  const vertRelCode = (flags >>> 3) & 0x3;
  const horzRelCode = (flags >>> 8) & 0x3;
  const vertAlignCode = (flags >>> 5) & 0x7;
  const horzAlignCode = (flags >>> 10) & 0x7;
  const wrapCode = (flags >>> 21) & 0x7;

  const vertRelTo: ImgLayout['vertRelTo'] =
    vertRelCode === 2 ? 'para' : 'page';
  const horzRelTo: ImgLayout['horzRelTo'] =
    horzRelCode === 2 ? 'column' : horzRelCode === 3 ? 'para' : 'page';
  const vertAlign = (['top', 'center', 'bottom'] as const)[vertAlignCode];
  const horzAlign = (['left', 'center', 'right'] as const)[horzAlignCode];
  const wrap = (
    ['square', 'tight', 'through', 'topAndBottom', 'behind', 'front'] as const
  )[wrapCode] ?? 'square';

  // The binary layout stores vertical offset first and horizontal offset next.
  const rawY = i32(data, 8);
  const rawX = i32(data, 12);
  const layout: ImgLayout = {
    wrap,
    horzRelTo,
    vertRelTo,
    horzAlign,
    vertAlign,
    xPt: rawX !== 0 ? Metric.hwpToPt(rawX) : undefined,
    yPt: rawY !== 0 ? Metric.hwpToPt(rawY) : undefined,
    behindDoc: wrap === 'behind' || undefined,
    zOrder: Math.max(0, i32(data, 24)),
  };

  // HWPUNIT16 outer margins: left, right, top, bottom.
  if (data.length >= 36) {
    layout.distL = Metric.hwpToPt(BinaryKit.readU16LE(data, 28));
    layout.distR = Metric.hwpToPt(BinaryKit.readU16LE(data, 30));
    layout.distT = Metric.hwpToPt(BinaryKit.readU16LE(data, 32));
    layout.distB = Metric.hwpToPt(BinaryKit.readU16LE(data, 34));
  }
  return layout;
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
        () => parseParagraphGroup(recs, i, di, shield, gsoCtx, content.length === 0),
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
  firstBodyParagraph = false,
): { nodes: ContentNode[]; next: number } {
  const hdr = recs[start];
  const lv  = hdr.level;

  // P1: PARA_HEADER 레이아웃
  //   offset 0-3: 글자 수 (최상위 비트는 유효 플래그이므로 제외)
  //   offset 8-9: paraShapeId (UINT16)
  //   offset 10:  styleId (UINT8)
  //   offset 11:  divideSort (UINT8) — 0x04=쪽나누기
  const _nchars    = hdr.data.length >= 4
    ? BinaryKit.readU32LE(hdr.data, 0) & 0x7fffffff
    : 0;
  const psId       = hdr.data.length >= 10 ? BinaryKit.readU16LE(hdr.data, 8) : 0;
  const hwpStyleId = hdr.data.length >= 11 ? hdr.data[10] : undefined;
  const divideSort = hdr.data.length >= 12 ? hdr.data[11] : 0;
  const ps         = di.paraShapes[psId];

  let text: ParaTextResult | null = null;
  let csPairs: [number, number][] = [];
  const grids: ContentNode[] = [];
  // imgId: for 'gso' uses sequential gsoCtx.count; for others uses flags-based objId
  const ctrlHeaders: { ctrlId: number; imgId: number; wPt: number; hPt: number; layout?: ImgLayout; atnoType?: number }[] = [];
  let hasSectionCtrl = false;
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
        if (ctrlId === CTRL_SECD) hasSectionCtrl = true;

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
          const rawH = r.data.length >= 24 ? BinaryKit.readU32LE(r.data, 20) : 0;
          const wPt = rawW > 0 && rawW < MAX_HWP ? Metric.hwpToPt(rawW) : 0;
          const hPt = rawH > 0 && rawH < MAX_HWP ? Metric.hwpToPt(rawH) : 0;
          const layout = parseObjectLayout(r.data);

          // P9: atno — offset 4 u32 하위 4bit = 번호 종류 (0=쪽번호, 6=전체쪽수)
          const atnoType = ctrlId === CTRL_ATNO && r.data.length >= 8
            ? BinaryKit.readU32LE(r.data, 4) & 15
            : undefined;

          const isPicture = ctrlId === CTRL_GSO || ctrlId === CTRL_PIC;
          const imgId = isPicture
            ? gsoCtx.count++
            : (r.data.length >= 6 ? BinaryKit.readU16LE(r.data, 4) : 0);
          const binIndex = isPicture ? pictureBinIndex(recs, i) : undefined;
          ctrlHeaders.push({ ctrlId, imgId, wPt, hPt, layout, atnoType });

          const isImageCtrl =
            ctrlId === CTRL_IMAGE || ctrlId === CTRL_PIC ||
            ctrlId === CTRL_FIG || ctrlId === CTRL_OBJ || ctrlId === CTRL_GSO;
          if (isImageCtrl) gsoCtx.objects.set(imgId, { wPt, hPt, layout, binIndex });

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
        const isImg = ch.ctrlId === CTRL_IMAGE || ch.ctrlId === CTRL_PIC || ch.ctrlId === CTRL_FIG || ch.ctrlId === CTRL_OBJ || ch.ctrlId === CTRL_GSO;
        if (!isImg) continue; // skip footnotes, TOC, page num, etc.
        paraContent.push(buildSpan(`__EXT_${ch.imgId}__`));
      }
    }

    const leadingExplicitBreak = firstBodyParagraph && (divideSort & 4) !== 0 &&
      !ps?.pageBreakBefore && grids.length === 0 && paraContent.length > 0;
    // A literal leading page break creates a blank first page. Word ignores
    // pageBreakBefore on its first paragraph, so keep the explicit break here.
    if (leadingExplicitBreak) {
      paraContent.unshift({ tag: 'span', props: {}, kids: [buildPb()] });
    }
    const hasPageBreakBefore = (divideSort & 4) !== 0 && !leadingExplicitBreak;
    // A grid has no paragraph properties of its own, so retain a standalone
    // break only when the break must precede a grid or has no following text.
    if (
      hasPageBreakBefore &&
      (grids.length > 0 || paraContent.length === 0)
    ) {
      nodes.push(buildPara([{ tag: 'span', props: {}, kids: [buildPb()] } as SpanNode]));
    }
    // P5: 표 → 앵커 문단 순서 (앵커 문단 드롭 금지)
    nodes.push(...grids);
    const isWhitespaceSectionPara =
      hasSectionCtrl &&
      grids.length === 0 &&
      paraContent.length > 0 &&
      paraContent.every((n: any) => {
        if (n?.tag !== 'span') return false;
        const text = (n.kids ?? [])
          .filter((kid: any) => kid?.tag === 'txt')
          .map((kid: any) => kid.content ?? '')
          .join('');
        return text.trim() === '';
      });
    const isSectionOnlyPara =
      hasSectionCtrl &&
      grids.length === 0 &&
      (paraContent.length === 0 || isWhitespaceSectionPara);
    const isPageBreakOnlyPara = (divideSort & 4) && paraContent.length === 0 && grids.length === 0;
    if (!isSectionOnlyPara && !isPageBreakOnlyPara) {
      const paraProps = buildParaProps(ps, hwpStyleId, di);
      if (hasPageBreakBefore && grids.length === 0)
        paraProps.pageBreakBefore = true;
      nodes.push(buildPara(
        paraContent.length > 0
          ? paraContent as any
          : resolveCharShapes([], csPairs, di),
        paraProps,
      ));
    }
  }

  return { nodes, next: i };
}

function skipKids(recs: HwpRecord[], idx: number): number {
  const lv = recs[idx].level;
  let i = idx + 1;
  while (i < recs.length && recs[i].level > lv) i++;
  return i;
}

/** Resolve the referenced BinData item from the standard picture record. */
function pictureBinIndex(recs: HwpRecord[], ctrlIdx: number): number | undefined {
  const end = skipKids(recs, ctrlIdx);
  for (let i = ctrlIdx + 1; i < end; i++) {
    const data = recs[i].data;
    // HWP 5.0 spec table 107: picture-info starts at 68, BinItem id at 71.
    if (recs[i].tag === TAG_SHAPE_COMPONENT_PICTURE && data.length >= 73) {
      const binId = BinaryKit.readU16LE(data, 71);
      if (binId > 0) return binId - 1;
    }
  }
  return undefined;
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
    if (sid !== curId) { spans.push(...styledSpans(buf, curId, di)); buf = ''; curId = sid; }
    buf += chars[k].ch;
  }
  if (buf) spans.push(...styledSpans(buf, curId, di));
  return spans;
}

function styledSpans(text: string, shapeId: number, di: DocInfo): SpanNode[] {
  const cs = di.charShapes[shapeId];
  if (!cs) return [buildSpan(text)];

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

  return splitLeadingSymbolRuns(text, props, di);
}

function splitLeadingSymbolRuns(text: string, props: TextProps, di: DocInfo): SpanNode[] {
  if (!text) return [buildSpan(text, props)];

  const symbolFont = firstAvailableFont(di, ['한양신명조', 'HY신명조']) ?? props.font;
  const leadFont = firstAvailableFont(di, ['HCI Poppy']) ?? symbolFont;
  const out: SpanNode[] = [];
  let rest = text;

  const lead = rest.match(/^(\s+)([◦→])/);
  if (lead?.[1]) {
    out.push(buildSpan(lead[1], { ...props, b: false, font: leadFont }));
    rest = rest.slice(lead[1].length);
  }

  const marker = rest.match(/^([□◦→])(\s*)/);
  if (marker) {
    out.push(buildSpan(marker[1], { ...props, font: symbolFont }));
    if (marker[2] && marker[1] !== '□')
      out.push(buildSpan(marker[2], { ...props, b: false, font: leadFont }));
    rest = rest.slice(marker[0].length);
    if (marker[2] && marker[1] === '□') rest = `${marker[2]}${rest}`;
    if (!marker[2] && rest && (marker[1] === '◦' || marker[1] === '→')) {
      rest = ` ${rest}`;
    }
  }

  if (rest) appendLatinAwareSpans(out, rest, props, leadFont);
  return out.length ? out : [buildSpan(text, props)];
}

function firstAvailableFont(di: DocInfo, names: string[]): string | undefined {
  return names.find(name => di.faceNames.includes(name));
}

function appendLatinAwareSpans(out: SpanNode[], text: string, props: TextProps, _latinFont?: string): void {
  out.push(buildSpan(text, props));
}

/* ── Table control parsing ──────────────────────────────────── */

interface HwpCellPadding {
  left: number;
  right: number;
  top: number;
  bottom: number;
}

const HWP_DEFAULT_CELL_PADDING: HwpCellPadding = {
  left: 510,
  right: 510,
  top: 141,
  bottom: 141,
};

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

  const rowCnt = Math.max(1, tblData.length >= 6 ? BinaryKit.readU16LE(tblData, 4) : 1);
  const colCnt = Math.max(1, tblData.length >= 8 ? BinaryKit.readU16LE(tblData, 6) : 1);
  const tablePadding: HwpCellPadding = tblData.length >= 18
    ? {
        left: inheritedHwpPadding(BinaryKit.readU16LE(tblData, 10), HWP_DEFAULT_CELL_PADDING.left),
        right: inheritedHwpPadding(BinaryKit.readU16LE(tblData, 12), HWP_DEFAULT_CELL_PADDING.right),
        top: inheritedHwpPadding(BinaryKit.readU16LE(tblData, 14), HWP_DEFAULT_CELL_PADDING.top),
        bottom: inheritedHwpPadding(BinaryKit.readU16LE(tblData, 16), HWP_DEFAULT_CELL_PADDING.bottom),
      }
    : HWP_DEFAULT_CELL_PADDING;

  interface PC { row: number; col: number; cs: number; rs: number; widthHwp: number; heightHwp?: number; props: CellProps; cellChildren: (ParaNode | GridNode)[] }
  const parsed: PC[] = [];

  for (let ci = 0; ci < cells.length; ci++) {
    const c = cells[ci];
    const seqIdx = ci;
    const pc = shield.guard(
      () => parseCellRec(c.data, c.tag, recs, c.cStart, c.cEnd, di, shield, seqIdx, colCnt, gsoCtx, tablePadding),
      { row: Math.floor(ci / (colCnt || 1)), col: ci % (colCnt || 1), cs: 1, rs: 1, widthHwp: 0, heightHwp: undefined, props: {}, cellChildren: [buildPara([buildSpan('')])] },
      `hwp:cell@${c.cStart}`,
    );
    parsed.push(pc);
  }

  // Validate geometry before it reaches the DOCX table renderer.
  // Corrupt/misread cell spans can otherwise create thousands of virtual columns.
  const rowLimit = Math.max(rowCnt, Math.ceil(parsed.length / colCnt), 1);
  for (let idx = 0; idx < parsed.length; idx++) {
    const c = parsed[idx];
    const badPosition =
      !Number.isFinite(c.row) ||
      !Number.isFinite(c.col) ||
      c.row < 0 ||
      c.col < 0 ||
      c.col >= colCnt ||
      c.row > rowLimit * 4 + 20;
    if (badPosition) {
      c.row = Math.floor(idx / colCnt);
      c.col = idx % colCnt;
    }

    const maxColSpan = Math.max(1, colCnt - c.col);
    if (!Number.isFinite(c.cs) || c.cs < 1) c.cs = 1;
    if (c.cs > maxColSpan) c.cs = maxColSpan;

    const maxRowSpan = Math.max(1, rowCnt - Math.min(c.row, rowCnt - 1));
    if (!Number.isFinite(c.rs) || c.rs < 1) c.rs = 1;
    if (c.rs > maxRowSpan) c.rs = maxRowSpan;
  }

  // Determine actual row count from normalized cell data.
  const maxRow = parsed.reduce((m, c) => Math.max(m, c.row + c.rs), 0);
  const actualRowCnt = Math.max(rowCnt, maxRow);

  const colWidthsPt = inferColumnWidths(
    colCnt,
    parsed
      .filter(c => c.widthHwp > 0)
      .map(c => ({ start: c.col, span: c.cs, width: c.widthHwp })),
  ).map(Metric.hwpToPt);

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
  gp.cellPadL = Metric.hwpToPt(tablePadding.left);
  gp.cellPadR = Metric.hwpToPt(tablePadding.right);
  gp.cellPadT = Metric.hwpToPt(tablePadding.top);
  gp.cellPadB = Metric.hwpToPt(tablePadding.bottom);
  const hasWidths = colWidthsPt.some(w => w > 0);
  if (hasWidths) gp.colWidths = colWidthsPt;
  const tableLayout = parseObjectLayout(recs[ctrlIdx].data);
  if (tableLayout && tableLayout.wrap !== 'inline') gp.layout = tableLayout;
  return { grid: buildGrid(rows, gp), next: i };
}

/* ── Cell record ────────────────────────────────────────────── */
/*  LIST_HEADER for cells (HWP 5.0/5.1):
    [0:2]  paraCount   [2:2] reserved   [4:4] attr (bits 5-6 = vertAlign)
    [8:2]  colAddr   [10:2] rowAddr
    [12:2] colSpan     [14:2] rowSpan
    [16:4] width(HWPUNIT)  [20:4] height(HWPUNIT)
    [24:8] padding[4]      [32:2] borderFillId                  */

function parseCellRec(
  d: Uint8Array, tag: number, recs: HwpRecord[], cStart: number, cEnd: number,
  di: DocInfo, shield: ShieldedParser, seqIdx: number, colCnt: number, gsoCtx: GsoCtx,
  tablePadding: HwpCellPadding,
) {
  let col: number, row: number, cs = 1, rs = 1;
  let widthHwp = 0;
  let heightHwp = 0;
  const props: CellProps = {};

  const attr = tag === TAG_LIST_HEADER
    ? (d.length >= 8 ? BinaryKit.readU32LE(d, 4) : 0)
    : (d.length >= 6 ? BinaryKit.readU32LE(d, 2) : 0);
  const va = (attr >> 5) & 0x3;
  if (va === 1) props.va = 'mid';
  else if (va === 2) props.va = 'bot';

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
      if (isCellPaddingOverride(pL, tablePadding.left)) props.padL = Metric.hwpToPt(pL);
      if (isCellPaddingOverride(pR, tablePadding.right)) props.padR = Metric.hwpToPt(pR);
      if (isCellPaddingOverride(pT, tablePadding.top)) props.padT = Metric.hwpToPt(pT);
      if (isCellPaddingOverride(pB, tablePadding.bottom)) props.padB = Metric.hwpToPt(pB);
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
      if (isCellPaddingOverride(pL, tablePadding.left)) props.padL = Metric.hwpToPt(pL);
      if (isCellPaddingOverride(pR, tablePadding.right)) props.padR = Metric.hwpToPt(pR);
      if (isCellPaddingOverride(pT, tablePadding.top)) props.padT = Metric.hwpToPt(pT);
      if (isCellPaddingOverride(pB, tablePadding.bottom)) props.padB = Metric.hwpToPt(pB);
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
          const ctrlHdrs: { ctrlId: number; imgId: number; wPt: number; hPt: number; layout?: ImgLayout }[] = [];
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
                  const rawH = recs[j].data.length >= 24 ? BinaryKit.readU32LE(recs[j].data, 20) : 0;
                  const wPt = rawW > 0 && rawW < MAX_HWP ? Metric.hwpToPt(rawW) : 0;
                  const hPt = rawH > 0 && rawH < MAX_HWP ? Metric.hwpToPt(rawH) : 0;
                  const layout = parseObjectLayout(recs[j].data);
                  const isPicture = ctrlId === CTRL_GSO || ctrlId === CTRL_PIC;
                  const imgId = isPicture
                    ? gsoCtx.count++
                    : (recs[j].data.length >= 6 ? BinaryKit.readU16LE(recs[j].data, 4) : 0);
                  const binIndex = isPicture ? pictureBinIndex(recs, j) : undefined;
                  ctrlHdrs.push({ ctrlId, imgId, wPt, hPt, layout });
                  const isImageCtrl =
                    ctrlId === CTRL_IMAGE || ctrlId === CTRL_PIC ||
                    ctrlId === CTRL_FIG || ctrlId === CTRL_OBJ || ctrlId === CTRL_GSO;
                  if (isImageCtrl) gsoCtx.objects.set(imgId, { wPt, hPt, layout, binIndex });
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
              const isImg = ch.ctrlId === CTRL_IMAGE || ch.ctrlId === CTRL_PIC || ch.ctrlId === CTRL_FIG || ch.ctrlId === CTRL_OBJ || ch.ctrlId === CTRL_GSO;
              if (!isImg) continue;
              paraContent.push(buildSpan(`__EXT_${ch.imgId}__`));
            }
          }
          const kids = paraContent.length > 0 ? paraContent as any : [buildSpan('')];
          // P6: innerGrids 먼저, 앵커 문단 나중 (P5와 동일한 순서)
          const hasPageBreakBefore = (cellDivide & 4) !== 0;
          const isPageBreakOnlyPara = hasPageBreakBefore && paraContent.length === 0 && innerGrids.length === 0;
          const items: (ParaNode | GridNode)[] = [...innerGrids];
          if (!isPageBreakOnlyPara) {
            const paraProps = buildParaProps(ps, cellStyleId, di);
            if (hasPageBreakBefore && innerGrids.length === 0)
              paraProps.pageBreakBefore = true;
            items.push(buildPara(kids, paraProps));
          }
          if (
            hasPageBreakBefore &&
            (innerGrids.length > 0 || paraContent.length === 0)
          )
            items.unshift(buildPara([{ tag: 'span', props: {}, kids: [buildPb()] } as SpanNode]));
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
      if (cellCtrlId === CTRL_GSO || cellCtrlId === CTRL_PIC) {
        const gsoId = gsoCtx.count++;
        const binIndex = pictureBinIndex(recs, k);
        const rawW = recs[k].data.length >= 24 ? BinaryKit.readU32LE(recs[k].data, 16) : 0;
        const rawH = recs[k].data.length >= 24 ? BinaryKit.readU32LE(recs[k].data, 20) : 0;
        const wPt = rawW > 0 && rawW < MAX_HWP ? Metric.hwpToPt(rawW) : 0;
        const hPt = rawH > 0 && rawH < MAX_HWP ? Metric.hwpToPt(rawH) : 0;
        const layout = parseObjectLayout(recs[k].data);
        gsoCtx.objects.set(gsoId, { wPt, hPt, layout, binIndex });
        cellChildren.push(buildPara([buildSpan(`__EXT_${gsoId}__`)]));
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

function inheritedHwpPadding(value: number, fallback: number): number {
  return value === 0xffff ? fallback : value;
}

function isCellPaddingOverride(value: number, inherited: number): boolean {
  return value !== 0xffff && value !== inherited;
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
    headerPt: Metric.hwpToPt(header),
    footerPt: Metric.hwpToPt(footer),
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

function headingFromStyle(style?: HwpStyle): Heading | undefined {
  if (!style) return undefined;
  for (const name of [style.name, style.engName]) {
    const match = name.match(/^(?:개요|outline|heading)\s*([1-6])$/i);
    if (match) return Number(match[1]) as Heading;
  }
  return undefined;
}

function buildParaProps(
  ps?: HwpParaShape,
  hwpStyleId?: number,
  di?: DocInfo,
): ParaProps {
  // P2: hwpStyleId를 초기값으로 포함 (undefined이면 빈 객체)
  const p: ParaProps = hwpStyleId !== undefined ? { hwpStyleId } : {};
  const heading = ps?.heading ?? headingFromStyle(di?.styles[hwpStyleId ?? -1]);
  if (heading !== undefined) p.heading = heading;
  if (!ps) return { ...p, spaceBefore: 0, spaceAfter: 0, lineHeight: 1.6 };
  if (ps.listOrd !== undefined) {
    p.listOrd = ps.listOrd;
    p.listLv = Math.max(0, Math.min(6, ps.listLevel ?? 0));
    if (ps.listOrd) {
      p.listMark = '1.';
    } else {
      const character = ps.listId && di
        ? di.bullets[ps.listId - 1]?.character
        : undefined;
      p.listMark = character || '-';
    }
  }
  if (ps.align && ps.align !== 'justify') p.align = ps.align;
  p.pageBreakBefore = ps.pageBreakBefore;
  p.keepWithNext = ps.keepWithNext;
  p.keepLines = ps.keepLines;
  p.widowControl = ps.widowControl;
  if (hwpStyleId === 18 && !p.align) p.align = 'justify';
  p.spaceBefore = Math.max(0, Metric.hwpToPt(ps.spaceBefore / 2));
  p.spaceAfter = Math.max(0, Metric.hwpToPt(ps.spaceAfter / 2));
  // 줄 간격: type=0(PERCENT) → lineHeight, type=1(FIXED)/3(AT_LEAST) → lineHeightFixed
  if (ps.lineSpacingType === 1 || ps.lineSpacingType === 3) {
    if (ps.lineSpacing > 0) {
      // PARA_SHAPE의 고정/최소 줄 높이도 여백 계열과 같이
      // HWPUNIT의 2배 값으로 저장된다.
      p.lineHeightFixed = Metric.hwpToPt(ps.lineSpacing / 2);
      p.lineHeightRule = ps.lineSpacingType === 3 ? 'atLeast' : 'exact';
    }
  } else {
    // P10: HWP 기본 줄 간격은 160%. 0/누락도 명시적으로 정규화해
    // Word의 1.15줄/문단 뒤 8pt 기본값이 셀 높이에 섞이지 않게 한다.
    p.lineHeight = ps.lineSpacing > 0 ? ps.lineSpacing / 100 : 1.6;
  }
  // HWP 5.0 ParaShape 여백 계열은 HWPUNIT의 2배 값으로 저장된다.
  // HWP leftMargin anchors the first line when indent is negative. The common
  // model stores the body margin, so include the hanging width at this boundary.
  const leftMarginPt = Metric.hwpToPt((ps.leftMargin - Math.min(0, ps.indent)) / 2);
  if (leftMarginPt !== 0) p.indentPt = leftMarginPt;
  // rightMargin (offset 8) = 문단 몸체 오른쪽 여백 → indentRightPt (pt)
  const rightMarginPt = Metric.hwpToPt(ps.rightMargin / 2);
  if (rightMarginPt !== 0) p.indentRightPt = rightMarginPt;
  // indent (offset 12) = 첫 줄 들여쓰기(양수) / 내어쓰기(음수) → firstLineIndentPt
  if (ps.indent !== 0) p.firstLineIndentPt = Metric.hwpToPt(ps.indent / 2);
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
      let di: DocInfo = {
        faceNames: [],
        charShapes: [],
        paraShapes: [],
        borderFills: [],
        styles: [],
        numberings: [],
        bullets: [],
      };
      if (diRaw) {
        di = shield.guard(() => parseDocInfo(diRaw, compressed), di, 'hwp:docInfo');
      }

      // Extract images from BinData streams.
      // HWP duplicates each BinData entry: once as "BinData/BIN0001.jpg" and once as "BIN0001.jpg".
      // Keep canonical streams and key them by their real BinData reference id.
      const binEntries: { binNum: number; ext: string; data: Uint8Array }[] = [];
      for (const [path, streamData] of streams) {
        // Match "BinData/BIN0001.jpg" style — the canonical form
        const m = path.match(/^BinData[/\\]BIN([0-9a-f]+)\.([a-z0-9]+)$/i);
        if (m) binEntries.push({ binNum: parseInt(m[1], 16), ext: m[2].toLowerCase(), data: streamData });
      }
      // Sort for deterministic processing; sparse BIN numbers remain sparse.
      binEntries.sort((a, b) => a.binNum - b.binNum);

      const objectMap = new Map<number, ImgNode>();
      for (const { binNum, ext, data: storedData } of binEntries) {
        let imgData = storedData;
        try {
          const inflated = pako.inflateRaw(storedData);
          if (looksLikeImageData(inflated, ext)) imgData = inflated;
        } catch {
          // EMBEDDING entries may explicitly opt out of BinData compression.
        }

        // Determine MIME type from binary signature first, then fall back to extension
        let mimeType: ImgNode['mime'] = 'image/jpeg';
        if (imgData[0] === 0x89 && imgData[1] === 0x50) mimeType = 'image/png';
        else if (imgData[0] === 0x47 && imgData[1] === 0x49) mimeType = 'image/gif';
        else if (imgData[0] === 0x42 && imgData[1] === 0x4D) mimeType = 'image/bmp';
        else if (ext === 'wmf') mimeType = 'image/x-wmf';
        else if (ext === 'emf') mimeType = 'image/x-emf';

        const base64 = TextKit.base64Encode(imgData);
        const { wPt, hPt } = getImageDimsPt(imgData, mimeType);
        objectMap.set(binNum - 1, buildImg(base64, mimeType, wPt, hPt));
      }

      // gsoCtx tracks sequential 'gso' encounter order — must be shared across all sections
      const gsoCtx: GsoCtx = { count: 0, objects: new Map() };

      // BodyText/SectionN is a logical HWP section, not just another chunk of
      // the same page stream. Preserve each one so the DOCX encoder can emit a
      // real section break with the matching page geometry.
      const parsedSections: Array<{
        content: ContentNode[];
        dims: PageDims;
      }> = [];
      let inheritedDims: PageDims = { ...A4 };

      for (let s = 0; s < 100; s++) {
        const sec = streams.get(`BodyText/Section${s}`) ?? streams.get(`Section${s}`);
        if (!sec) {
          if (s === 0) {
            const fb = findBodySection(streams);
            if (fb) {
              const r = parseBody(fb, compressed, di, shield, gsoCtx);
              inheritedDims = r.pageDims ?? inheritedDims;
              parsedSections.push({
                content: r.content,
                dims: inheritedDims,
              });
            }
          }
          break;
        }
        const r = shield.guard(
          () => parseBody(sec, compressed, di, shield, gsoCtx),
          { content: [], pageDims: undefined },
          `hwp:sec${s}`,
        );
        inheritedDims = r.pageDims ?? inheritedDims;
        parsedSections.push({
          content: r.content,
          dims: inheritedDims,
        });
      }

      const allContent = parsedSections.flatMap((section) => section.content);
      if (objectMap.size > 0) {
        injectImagesIntoContent(allContent, objectMap, gsoCtx.objects);
      }

      for (const section of parsedSections) {
        normalizeHancomParagraphAnchors(section.content, di);
      }

      warns.push(...shield.flush());
      if (parsedSections.length === 0) {
        parsedSections.push({
          content: [buildPara([buildSpan('')])],
          dims: inheritedDims,
        });
      }

      const sheets = parsedSections.map((section, index) =>
        buildSheet(
          section.content.length > 0
            ? section.content
            : [buildPara([buildSpan('')])],
          section.dims,
          // Header/footer controls are currently collected once and inherited
          // by later sections. Attach them to the first section explicitly.
          index === 0
            ? {
                headers: gsoCtx.headers
                  ? { default: gsoCtx.headers }
                  : undefined,
                footers: gsoCtx.footers
                  ? { default: gsoCtx.footers }
                  : undefined,
              }
            : undefined,
        ),
      );
      return succeed(buildRoot({}, sheets), warns);
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

function looksLikeImageData(data: Uint8Array, ext: string): boolean {
  if (data.length < 4) return false;
  if (data[0] === 0x89 && data[1] === 0x50 && data[2] === 0x4e && data[3] === 0x47) return true;
  if (data[0] === 0xff && data[1] === 0xd8 && data[2] === 0xff) return true;
  if (data[0] === 0x47 && data[1] === 0x49 && data[2] === 0x46) return true;
  if (data[0] === 0x42 && data[1] === 0x4d) return true;
  if (ext === 'wmf') {
    return (data[0] === 0xd7 && data[1] === 0xcd && data[2] === 0xc6 && data[3] === 0x9a) ||
      (data[0] === 0x01 && data[1] === 0x00 && data[2] === 0x09 && data[3] === 0x00);
  }
  return ext === 'emf' && data.length >= 44 &&
    data[40] === 0x20 && data[41] === 0x45 && data[42] === 0x4d && data[43] === 0x46;
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
  objectMap: Map<number, ImgNode>,
  objectInfo: Map<number, { wPt: number; hPt: number; layout?: ImgLayout; binIndex?: number }>,
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
          const info = objectInfo.get(objId);
          const base = objectMap.get(info?.binIndex ?? objId);
          if (base) {
            const wPt = match[2] ? parseInt(match[2], 10) : 0;
            const hPt = match[3] ? parseInt(match[3], 10) : 0;
            // Common-object dimensions are exact HWPUNIT values.  Pixel size is
            // only a fallback for malformed controls without a display box.
            kids[i] = {
              ...base,
              w: info?.wPt && info.wPt > 0 ? info.wPt : (wPt > 0 ? wPt : base.w),
              h: info?.hPt && info.hPt > 0 ? info.hPt : (hPt > 0 ? hPt : base.h),
              layout: info?.layout,
            };
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

function normalizeHancomParagraphAnchors(content: ContentNode[], di: DocInfo): void {
  normalizeContentList(content as any[], di);
}

function normalizeContentList(content: any[], di: DocInfo): void {
  for (const node of content) {
    if (node?.tag === 'grid') {
      for (const row of node.kids ?? []) {
        for (const cell of row.kids ?? []) normalizeContentList(cell.kids ?? [], di);
      }
    }
  }

  for (let i = 0; i < content.length; i++) {
    const node = content[i] as ContentNode;
    if (isEmptyCenterPara(node) && paraText(content[i + 1]).startsWith('※ 모든 서류')) {
      content.splice(i, 1);
      i--;
      continue;
    }
    if (
      paraText(node).startsWith('제출예시)') &&
      !isEmptyCenterPara(content[i - 1])
    ) {
      const font = firstAvailableFont(di, ['HCI Poppy']);
      content.splice(i, 0, buildPara([buildSpan('', font ? { font, pt: 13 } : {})], { hwpStyleId: 0, align: 'center' }));
      i++;
    }
  }
}

function isEmptyCenterPara(node: ContentNode | undefined): boolean {
  return !!node && node.tag === 'para' && paraText(node) === '' && node.props.align === 'center';
}

function paraText(node: ContentNode | undefined): string {
  if (!node || node.tag !== 'para') return '';
  let out = '';
  const collect = (kids: any[]) => {
    for (const kid of kids ?? []) {
      if (kid.tag === 'txt') out += kid.content ?? '';
      else if (kid.kids) collect(kid.kids);
    }
  };
  collect(node.kids as any[]);
  return out.trim();
}

registry.registerDecoder(new HwpScanner());
