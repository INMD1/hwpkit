import type { Decoder } from '../../contract/decoder';
import type { DocRoot, ContentNode, ParaNode, SpanNode, ImgNode } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import type { Align, Stroke, StrokeKind, PageDims, TextProps, ParaProps, CellProps, GridProps } from '../../model/doc-props';
import { succeed, fail } from '../../contract/result';
import { buildRoot, buildSheet, buildPara, buildSpan, buildGrid, buildRow, buildCell, buildImg } from '../../model/builders';
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
const CTRL_TABLE = 0x74626C20;  // ' lbt'
const CTRL_IMAGE = 0x696D6720;  // 'img '
const CTRL_OBJ   = 0x6F626A20;  // 'obj '
const CTRL_FIG   = 0x66696720;  // 'fig '

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
  leftMargin: number;
  indent: number;
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
    0       4     attr1   (bits 0-1 = alignment: 0=justify,1=left,2=right,3=center)
    4       4     leftMargin   (HWPUNIT)
    8       4     rightMargin
    12      4     indent
    16      4     spaceBefore
    20      4     spaceAfter
    24      4     lineSpacing                                         */

const ALIGN_TBL: Record<number, Align> = { 0: 'justify', 1: 'left', 2: 'right', 3: 'center', 4: 'justify' };

function parseParaShape(d: Uint8Array): HwpParaShape {
  if (d.length < 4) return { align: 'left', spaceBefore: 0, spaceAfter: 0, lineSpacing: 160, leftMargin: 0, indent: 0 };
  const attr = BinaryKit.readU32LE(d, 0);
  return {
    align:       ALIGN_TBL[(attr >> 2) & 0x7] ?? 'left',
    leftMargin:  d.length >= 8  ? i32(d, 4)  : 0,  // offset 4: leftMargin (들여쓰기)
    indent:      d.length >= 16 ? i32(d, 12) : 0,  // offset 12: first-line indent
    spaceBefore: d.length >= 20 ? i32(d, 16) : 0,
    spaceAfter:  d.length >= 24 ? i32(d, 20) : 0,
    lineSpacing: d.length >= 28 ? i32(d, 24) : 160,
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

function parseBody(
  raw: Uint8Array, compressed: boolean, di: DocInfo, shield: ShieldedParser,
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
        () => parseParagraphGroup(recs, i, di, shield),
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
  recs: HwpRecord[], start: number, di: DocInfo, shield: ShieldedParser,
): { nodes: ContentNode[]; next: number } {
  const hdr = recs[start];
  const lv  = hdr.level;

  // paraShapeId at offset 8 (UINT16)
  const psId = hdr.data.length >= 10 ? BinaryKit.readU16LE(hdr.data, 8) : 0;
  const ps   = di.paraShapes[psId];

  let text: ParaTextResult | null = null;
  let csPairs: [number, number][] = [];
  const grids: ContentNode[] = [];
  const ctrlHeaders: { ctrlId: number; objId: number }[] = [];
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
        // objId at offset 4 (UINT16) - identifies the image/object in BinData
        const objId = r.data.length >= 6 ? BinaryKit.readU16LE(r.data, 4) : 0;
        ctrlHeaders.push({ ctrlId, objId });

        if (ctrlId === CTRL_TABLE) {
          const tr = shield.guard(
            () => parseTableCtrl(recs, i, di, shield),
            { grid: null, next: skipKids(recs, i) },
            `hwp:tbl@${i}`,
          );
          if (tr.grid) grids.push(tr.grid);
          i = tr.next;
        } else {
          i = skipKids(recs, i);
        }
      } else {
        i = skipKids(recs, i);
      }
    } else {
      i++;
    }
  }

  // Match extended controls with CTRL_HEADER entries
  if (text && ctrlHeaders.length > 0) {
    for (let ci = 0; ci < text.controls.length; ci++) {
      if (ci < ctrlHeaders.length) {
        text.controls[ci].ctrlId = ctrlHeaders[ci].ctrlId;
        text.controls[ci].matched = true;
      }
    }
  }

  const nodes: ContentNode[] = [];

  // Build paragraph from text and inline controls (images)
  if (text && (text.chars.length > 0 || text.controls.length > 0)) {
    const paraContent: (SpanNode | ContentNode)[] = [];

    // Process text chars and controls together
    if (text.chars.length > 0) {
      const spans = resolveCharShapes(text.chars, csPairs, di);
      paraContent.push(...spans);
    }

    // Add placeholder spans for extended controls (images)
    if (text.controls.length > 0) {
      for (let ci = 0; ci < text.controls.length; ci++) {
        // Create placeholder for all extended controls
        // Image replacement will happen later in injectImagesIntoContent
        paraContent.push(buildSpan(`__EXT_${ci}__`));
      }
    }

    if (paraContent.length > 0) {
      nodes.push(buildPara(paraContent as any, buildParaProps(ps)));
    }
  }

  nodes.push(...grids);
  return { nodes, next: i };
}

function skipKids(recs: HwpRecord[], idx: number): number {
  const lv = recs[idx].level;
  let i = idx + 1;
  while (i < recs.length && recs[i].level > lv) i++;
  return i;
}

/* ── PARA_TEXT ───────────────────────────────────────────────── */

// Extended controls: 8 WORDs, associated CTRL_HEADER
const EXT_CTRL = new Set([2, 3, 11, 12, 14, 15]);
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
  recs: HwpRecord[], ctrlIdx: number, di: DocInfo, shield: ShieldedParser,
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

  interface PC { row: number; col: number; cs: number; rs: number; widthHwp: number; props: CellProps; paras: ParaNode[] }
  const parsed: PC[] = [];

  for (let ci = 0; ci < cells.length; ci++) {
    const c = cells[ci];
    const seqIdx = ci;
    const pc = shield.guard(
      () => parseCellRec(c.data, c.tag, recs, c.cStart, c.cEnd, di, shield, seqIdx, colCnt),
      { row: Math.floor(ci / (colCnt || 1)), col: ci % (colCnt || 1), cs: 1, rs: 1, widthHwp: 0, props: {}, paras: [buildPara([buildSpan('')])] },
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
  const zeroColumns = colWidthsPt.filter(w => w === 0).length;
  if (zeroColumns > 0) {
    for (const c of parsed) {
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

  const rows = [];
  for (let r = 0; r < actualRowCnt; r++) {
    const rc = parsed.filter(c => c.row === r).sort((a, b) => a.col - b.col);
    if (rc.length === 0) continue;
    rows.push(buildRow(rc.map(c =>
      buildCell(c.paras.length ? c.paras : [buildPara([buildSpan('')])], { cs: c.cs, rs: c.rs, props: c.props }),
    )));
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
  di: DocInfo, shield: ShieldedParser, seqIdx: number, colCnt: number,
) {
  let col: number, row: number, cs = 1, rs = 1;
  let widthHwp = 0;
  const props: CellProps = {};

  const attr = d.length >= 6 ? BinaryKit.readU32LE(d, 2) : 0;
  const va = (attr >> 6) & 0x3;
  if (va === 1) props.va = 'mid';
  else if (va === 2) props.va = 'bot';

  if (tag === TAG_LIST_HEADER && d.length >= 22) {
    // LIST_HEADER with cell-specific fields
    // offset 8: colAddr, offset 10: rowAddr (HWP 5.0 spec)
    col = BinaryKit.readU16LE(d, 8);
    row = BinaryKit.readU16LE(d, 10);
    cs  = Math.max(1, BinaryKit.readU16LE(d, 12));
    rs  = Math.max(1, BinaryKit.readU16LE(d, 14));
    widthHwp = BinaryKit.readU32LE(d, 16);

    const bfId = d.length >= 34 ? BinaryKit.readU16LE(d, 32) : 0;
    if (bfId > 0 && bfId <= di.borderFills.length) {
      const bf = di.borderFills[bfId - 1];
      if (bf.borders.length >= 4) {
        props.left  = toStroke(bf.borders[0]);
        props.right = toStroke(bf.borders[1]);
        props.top   = toStroke(bf.borders[2]);
        props.bot   = toStroke(bf.borders[3]);
      }
      if (bf.bgColor && bf.bgColor !== 'FFFFFF') props.bg = bf.bgColor;
    }
  } else if (tag !== TAG_LIST_HEADER) {
    // Full CELL record with position/span/borderFill
    col = d.length >= 8  ? BinaryKit.readU16LE(d, 6) : seqIdx % (colCnt || 1);
    row = d.length >= 10 ? BinaryKit.readU16LE(d, 8) : Math.floor(seqIdx / (colCnt || 1));
    cs  = d.length >= 12 ? Math.max(1, BinaryKit.readU16LE(d, 10)) : 1;
    rs  = d.length >= 14 ? Math.max(1, BinaryKit.readU16LE(d, 12)) : 1;
    widthHwp = d.length >= 18 ? BinaryKit.readU32LE(d, 14) : 0;

    const bfId = d.length >= 32 ? BinaryKit.readU16LE(d, 30) : 0;
    if (bfId > 0 && bfId <= di.borderFills.length) {
      const bf = di.borderFills[bfId - 1];
      if (bf.borders.length >= 4) {
        props.left  = toStroke(bf.borders[0]);
        props.right = toStroke(bf.borders[1]);
        props.top   = toStroke(bf.borders[2]);
        props.bot   = toStroke(bf.borders[3]);
      }
      if (bf.bgColor && bf.bgColor !== 'FFFFFF') props.bg = bf.bgColor;
    }
  } else {
    // Fallback: LIST_HEADER too short, compute sequentially
    row = Math.floor(seqIdx / (colCnt || 1));
    col = seqIdx % (colCnt || 1);
  }

  // Parse cell content paragraphs
  const paras: ParaNode[] = [];
  let k = cStart;
  while (k < cEnd) {
    if (recs[k].tag === TAG_PARA_HEADER) {
      // For cell paragraphs, they might be at various nesting levels
      const r = shield.guard(
        () => {
          const hdr = recs[k];
          const lv = hdr.level;
          const psId = hdr.data.length >= 10 ? BinaryKit.readU16LE(hdr.data, 8) : 0;
          const ps = di.paraShapes[psId];
          let txt: ParaTextResult | null = null;
          let csp: [number, number][] = [];
          let j = k + 1;
          while (j < cEnd && recs[j].level > lv) {
            if (recs[j].tag === TAG_PARA_TEXT) { txt = decodeParaText(recs[j].data); j++; }
            else if (recs[j].tag === TAG_PARA_CHAR_SHAPE) { csp = parseCharShapePairs(recs[j].data); j++; }
            else j++;
          }
          const spans = txt && txt.chars.length > 0 ? resolveCharShapes(txt.chars, csp, di) : [buildSpan('')];
          return { para: buildPara(spans, buildParaProps(ps)), next: j };
        },
        { para: buildPara([buildSpan('')]), next: k + 1 },
        `hwp:cellP@${k}`,
      );
      paras.push(r.para);
      k = r.next;
    } else { k++; }
  }

  return { row, col, cs, rs, props, widthHwp, paras: paras.length ? paras : [buildPara([buildSpan('')])] };
}

/* ── PAGE_DEF ───────────────────────────────────────────────── */
/*  [0:4] width  [4:4] height  [8:4] ml  [12:4] mr
    [16:4] mt  [20:4] mb  [36:4] attr (bit0=landscape)           */

function parsePageDef(d: Uint8Array): PageDims {
  if (d.length < 24) return A4;
  const w  = BinaryKit.readU32LE(d, 0);
  const h  = BinaryKit.readU32LE(d, 4);
  const ml = BinaryKit.readU32LE(d, 8);
  const mr = BinaryKit.readU32LE(d, 12);
  const mt = BinaryKit.readU32LE(d, 16);
  const mb = BinaryKit.readU32LE(d, 20);
  const at = d.length >= 40 ? BinaryKit.readU32LE(d, 36) : 0;
  return {
    wPt: Metric.hwpToPt(w),  hPt: Metric.hwpToPt(h),
    ml: Metric.hwpToPt(ml),  mr: Metric.hwpToPt(mr),
    mt: Metric.hwpToPt(mt),  mb: Metric.hwpToPt(mb),
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

function strokeFromBF(bfId: number, di: DocInfo): Stroke | undefined {
  if (bfId <= 0 || bfId > di.borderFills.length) return undefined;
  const bf = di.borderFills[bfId - 1];
  if (!bf.borders.length) return undefined;
  const b = bf.borders[0];
  return { kind: BORDER_KIND[b.type] ?? 'solid', pt: b.widthPt, color: b.color };
}

function buildParaProps(ps?: HwpParaShape): ParaProps {
  if (!ps) return {};
  const p: ParaProps = {};
  if (ps.align && ps.align !== 'left') p.align = ps.align;
  if (ps.spaceBefore > 0) p.spaceBefore = Metric.hwpToPt(ps.spaceBefore);
  if (ps.spaceAfter > 0)  p.spaceAfter  = Metric.hwpToPt(ps.spaceAfter);
  if (ps.lineSpacing > 0 && ps.lineSpacing !== 160) p.lineHeight = ps.lineSpacing / 100;
  // leftMargin (offset 4) = 문단 몸체 왼쪽 여백 → indentPt
  if (ps.leftMargin > 0) p.indentPt = Metric.hwpToPt(ps.leftMargin);
  // indent (offset 12) = 첫 줄 들여쓰기(양수) / 내어쓰기(음수) → firstLineIndentPt
  if (ps.indent !== 0) p.firstLineIndentPt = Metric.hwpToPt(ps.indent);
  return p;
}

/* ═══════════════════════════════════════════════════════════════
   Decoder class
   ═══════════════════════════════════════════════════════════════ */

export class HwpScanner implements Decoder {
  readonly format = 'hwp';

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

      // Extract images from BinData streams
      const imageStreams: { path: string; data: Uint8Array }[] = [];
      for (const [path, data] of streams) {
        if ((path.includes('BinData') || path.includes('.jpg') || path.includes('.jpeg') || path.includes('.png') || path.includes('.gif') || path.includes('.bmp'))
            && !path.includes('FileHeader') && !path.includes('DocInfo') && !path.includes('BodyText') && !path.includes('Section')) {
          imageStreams.push({ path, data });
          console.log(`[HwpScanner] Image stream found: ${path} (${data.length} bytes)`);
        }
      }

      // Create image nodes for each image stream (deduplicated by hash)
      const objectMap = new Map<number, ImgNode>();
      const seenHashes = new Set<string>();
      let imgIdx = 0;
      for (const { path, data } of imageStreams) {
        // Determine MIME type from extension or signature
        let mimeType = 'image/jpeg';
        const lowerPath = path.toLowerCase();
        if (lowerPath.includes('.png')) mimeType = 'image/png';
        else if (lowerPath.includes('.gif')) mimeType = 'image/gif';
        else if (lowerPath.includes('.bmp')) mimeType = 'image/bmp';

        // Also check signature
        if (data[0] === 0x89 && data[1] === 0x50 && data[2] === 0x4E && data[3] === 0x47) mimeType = 'image/png';
        else if (data[0] === 0x47 && data[1] === 0x49 && data[2] === 0x46 && data[3] === 0x3538) mimeType = 'image/gif';
        else if (data[0] === 0x42 && data[1] === 0x4D) mimeType = 'image/bmp';

        const base64 = TextKit.base64Encode(data);
        const hash = base64.slice(0, 20); // Use first 20 chars as simple hash
        if (!seenHashes.has(hash)) {
          seenHashes.add(hash);
          objectMap.set(imgIdx++, buildImg(
            base64,
            mimeType as any,
            0, // w
            0, // h
            `Image from ${path}`,
          ));
          console.log(`[HwpScanner] Added unique image: ${hash}... (${data.length} bytes)`);
        } else {
          console.log(`[HwpScanner] Duplicate image skipped: ${hash}...`);
        }
      }

      console.log(`[HwpScanner] Found ${imageStreams.length} image streams, ${objectMap.size} unique images`);

      // Body sections
      const allContent: ContentNode[] = [];
      let pageDims: PageDims = A4;

      for (let s = 0; s < 100; s++) {
        const sec = streams.get(`BodyText/Section${s}`) ?? streams.get(`Section${s}`);
        if (!sec) {
          if (s === 0) {
            const fb = findBodySection(streams);
            if (fb) {
              const r = parseBody(fb, compressed, di, shield);
              allContent.push(...r.content);
              if (r.pageDims) pageDims = r.pageDims;
            }
          }
          break;
        }
        const r = shield.guard(
          () => parseBody(sec, compressed, di, shield),
          { content: [], pageDims: undefined },
          `hwp:sec${s}`,
        );
        allContent.push(...r.content);
        if (r.pageDims) pageDims = r.pageDims;
      }

      // Inject images into paragraphs (only if images are available)
      console.log(`[HwpScanner] Before injection: ${allContent.length} nodes, ${objectMap.size} images available`);
      if (objectMap.size > 0) {
        injectImagesIntoContent(allContent, objectMap);
        console.log(`[HwpScanner] After injection: ${allContent.length} nodes`);
      }

      // Count images (recursively)
      const countImages = (nodes: ContentNode[]): number => {
        let count = 0;
        for (const node of nodes) {
          if ((node as any).tag === 'img') count++;
          if ((node as any).tag === 'para' && (node as any).kids) count += countImages((node as any).kids);
          if ((node as any).tag === 'grid' && (node as any).kids) {
            for (const row of (node as any).kids) {
              if (row.kids) count += countImages(row.kids);
            }
          }
        }
        return count;
      };
      const imgCount = countImages(allContent);
      console.log(`[HwpScanner] Images in content: ${imgCount}`);

      warns.push(...shield.flush());
      const content = allContent.length > 0 ? allContent : [buildPara([buildSpan('')])];
      return succeed(buildRoot({}, [buildSheet(content, pageDims)]), warns);
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
  const imageArray = Array.from(objectMap.values());
  if (imageArray.length === 0) return;

  // Get unique images (deduplicate by base64 content)
  const uniqueImages = Array.from(new Set(imageArray.map(img => img.b64))).map(b64 => {
    return imageArray.find(img => img.b64 === b64)!;
  });
  if (uniqueImages.length === 0) return;

  let imgIdx = 0;
  for (const node of content) {
    if (node.tag === 'para' && node.kids) {
      for (let i = 0; i < node.kids.length; i++) {
        const kid = node.kids[i];
        // Span node structure: { tag: 'span', props, kids: [{ tag: 'txt', content }] }
        if (kid.tag === 'span' && kid.kids && kid.kids[0]?.tag === 'txt') {
          const text = kid.kids[0].content;
          // Support both __IMG_N__ and __EXT_N__ patterns
          const match = text.match?.(/^__(?:IMG|EXT)_(\d+)__$/);
          if (match) {
            // Replace placeholder with next available image (round-robin)
            const imgNode = uniqueImages[imgIdx % uniqueImages.length];
            if (imgNode) {
              node.kids[i] = imgNode;
              imgIdx++;
            }
          }
        }
      }
    }
  }
}

registry.registerDecoder(new HwpScanner());
