/**
 * HWP 5.0 encoder — DocRoot → HWP binary (OLE2/CFB container)
 *
 * OLE2 layout:
 *   FileHeader (stream)  — 256-byte HWP signature + flags
 *   DocInfo    (stream)  — compressed FACE_NAME / CHAR_SHAPE / PARA_SHAPE records
 *   BodyText   (storage)
 *     Section0 (stream) — compressed PAGE_DEF + paragraph/table records
 */

import type { Encoder }       from '../../contract/encoder';
import type { DocRoot, ContentNode, ParaNode, SpanNode, GridNode } from '../../model/doc-tree';
import type { Outcome }       from '../../contract/result';
import type { TextProps, ParaProps, Stroke, PageDims } from '../../model/doc-props';
import { succeed, fail }      from '../../contract/result';
import { Metric, safeFontToKr } from '../../safety/StyleBridge';
import { registry }           from '../../pipeline/registry';
import { A4 }                 from '../../model/doc-props';
import pako                   from 'pako';
import { TextKit }            from '../../toolkit/TextKit';

/* ═══════════════════════════════════════════════════════════════
   HWP 5.0 tag IDs (HWP 5.0 spec 표 13, 표 57)
   HWPTAG_BEGIN = 16 (0x10)
   ═══════════════════════════════════════════════════════════════ */

const T = 16; // HWPTAG_BEGIN

// DocInfo tags (표 13)
const TAG_DOCUMENT_PROPERTIES = T + 0;  // 16 - HWPTAG_DOCUMENT_PROPERTIES
const TAG_ID_MAPPINGS          = T + 1;  // 17 - HWPTAG_ID_MAPPINGS
const TAG_FACE_NAME            = T + 3;  // 19 - HWPTAG_FACE_NAME
const TAG_BORDER_FILL          = T + 4;  // 20 - HWPTAG_BORDER_FILL
const TAG_CHAR_SHAPE           = T + 5;  // 21 - HWPTAG_CHAR_SHAPE
const TAG_PARA_SHAPE           = T + 9;  // 25 - HWPTAG_PARA_SHAPE

// DocInfo tags (additional)
const TAG_BIN_DATA             = T + 2;  // 18 - HWPTAG_BIN_DATA

// BodyText tags (표 57)
const TAG_PARA_HEADER          = T + 50; // 66 - HWPTAG_PARA_HEADER
const TAG_PARA_TEXT            = T + 51; // 67 - HWPTAG_PARA_TEXT
const TAG_PARA_CHAR_SHAPE      = T + 52; // 68 - HWPTAG_PARA_CHAR_SHAPE
const TAG_PARA_LINE_SEG        = T + 53; // 69 - HWPTAG_PARA_LINE_SEG
const TAG_CTRL_HEADER          = T + 55; // 71 - HWPTAG_CTRL_HEADER
const TAG_LIST_HEADER          = T + 56; // 72 - HWPTAG_LIST_HEADER
const TAG_PAGE_DEF             = T + 57; // 73 - HWPTAG_PAGE_DEF
const TAG_FOOTNOTE_SHAPE       = T + 58; // 74 - HWPTAG_FOOTNOTE_SHAPE
const TAG_TABLE                = T + 61; // 77 - HWPTAG_TABLE (표 개체)
const TAG_SHAPE_COMPONENT_PICTURE = T + 69; // 85 - HWPTAG_SHAPE_COMPONENT_PICTURE

// Control IDs (stored as LE UINT32, MAKE_4CHID order)
const CTRL_TABLE = 0x74626C20; // 'tbl ' table
const CTRL_SECD  = 0x73656364; // 'secd' section definition
const CTRL_PIC   = 0x24706963; // '$pic' picture/image

/** Border width index table (points) — 표 26 테두리선 굵기 (mm→pt conversion) */
const BORDER_W_PT = [
  0.28, 0.34, 0.43, 0.57, 0.71, 0.85,
  1.13, 1.42, 1.70, 1.98, 2.84, 4.25,
  5.67, 8.50, 11.34, 14.17,
];

/** 표 25 테두리선 종류: 0=실선, 1=점선, 2=긴점선, 3=dash-dot, 4=dash-dot-dot, 7=2중선, 8=3중선 등 */
const BORDER_KIND_IDX: Record<string, number> = {
  solid: 0,
  dot: 1,
  dash: 2,
  double: 7,
  triple: 8,
  none: 0,
};

/**
 * 표 44 문단 모양 속성1:
 * bits 2-4 = 정렬 방식 (0=양쪽, 1=왼쪽, 2=오른쪽, 3=가운데, 4=배분)
 */
const ALIGN_CODE: Record<string, number> = {
  justify: 0,
  left: 1,
  right: 2,
  center: 3,
  distribute: 4,
};

/* ═══════════════════════════════════════════════════════════════
   Binary buffer writer
   ═══════════════════════════════════════════════════════════════ */

class BufWriter {
  private chunks: Uint8Array[] = [];
  private _sz = 0;

  get size() { return this._sz; }

  u8(v: number): this {
    this.chunks.push(new Uint8Array([v & 0xFF]));
    this._sz++;
    return this;
  }

  u16(v: number): this {
    this.chunks.push(new Uint8Array([v & 0xFF, (v >> 8) & 0xFF]));
    this._sz += 2;
    return this;
  }

  u32(v: number): this {
    const b = new Uint8Array(4);
    b[0] = v & 0xFF;
    b[1] = (v >>> 8) & 0xFF;
    b[2] = (v >>> 16) & 0xFF;
    b[3] = (v >>> 24) & 0xFF;
    this.chunks.push(b);
    this._sz += 4;
    return this;
  }

  i32(v: number): this { return this.u32(v < 0 ? v + 0x100000000 : v); }

  i16(v: number): this { return this.u16(v < 0 ? v + 0x10000 : v); }

  bytes(d: Uint8Array): this { this.chunks.push(d); this._sz += d.length; return this; }
  zeros(n: number): this    { this.chunks.push(new Uint8Array(n)); this._sz += n; return this; }

  /** Write each char as UTF-16LE UINT16 (BMP only) */
  utf16(s: string): this {
    for (let i = 0; i < s.length; i++) this.u16(s.charCodeAt(i));
    return this;
  }

  /** Write 4-byte COLORREF (R, G, B, 0) from 6-hex string */
  colorRef(hex: string): this {
    const h = (hex || '000000').replace('#', '').padStart(6, '0');
    return this
      .u8(parseInt(h.slice(0, 2), 16))
      .u8(parseInt(h.slice(2, 4), 16))
      .u8(parseInt(h.slice(4, 6), 16))
      .u8(0);
  }

  build(): Uint8Array {
    const out = new Uint8Array(this._sz);
    let off = 0;
    for (const c of this.chunks) { out.set(c, off); off += c.length; }
    return out;
  }
}

/* ═══════════════════════════════════════════════════════════════
   HWP record builder
   Format: 32-bit header = size(12)|level(10)|tag(10)
   If size >= 0xFFF, append UINT32 with actual size.
   ═══════════════════════════════════════════════════════════════ */

function mkRec(tag: number, level: number, data: Uint8Array): Uint8Array {
  const sz = data.length;
  const enc = Math.min(sz, 0xFFF);
  const hdr = (enc << 20) | ((level & 0x3FF) << 10) | (tag & 0x3FF);
  const w = new BufWriter().u32(hdr);
  if (enc >= 0xFFF) w.u32(sz);
  w.bytes(data);
  return w.build();
}

/* ═══════════════════════════════════════════════════════════════
   Style collector (first pass — deduplicates fonts/shapes)
   ═══════════════════════════════════════════════════════════════ */

function csKey(p: TextProps): string {
  return [p.font ?? '', p.pt ?? 10, p.b ? 1 : 0, p.i ? 1 : 0, p.u ? 1 : 0,
    p.s ? 1 : 0, p.sup ? 1 : 0, p.sub ? 1 : 0, p.color ?? '000000'].join('|');
}
function psKey(p: ParaProps): string {
  return [p.align ?? 'left', p.indentPt ?? 0, p.firstLineIndentPt ?? 0,
    p.spaceBefore ?? 0, p.spaceAfter ?? 0, p.lineHeight ?? 1].join('|');
}
function bfKey(s: Stroke, bg?: string): string {
  return `${s.kind}|${s.pt}|${s.color}|${bg ?? ''}`;
}
function bfPerSideKey(l: Stroke, r: Stroke, t: Stroke, b: Stroke, bg?: string): string {
  return `${bfKey(l)}/${bfKey(r)}/${bfKey(t)}/${bfKey(b)}/${bg ?? ''}`;
}

type BfEntry =
  | { uniform: true; s: Stroke; bg?: string }
  | { uniform: false; l: Stroke; r: Stroke; t: Stroke; b: Stroke; bg?: string };

class StyleCollector {
  readonly DEF_STROKE: Stroke = { kind: 'solid', pt: 0.5, color: '000000' };

  fonts: string[] = ['맑은 고딕'];
  private fontIdx = new Map<string, number>([['맑은 고딕', 0]]);

  csProps: TextProps[] = [{}];
  private csIdx = new Map<string, number>([[csKey({}), 0]]);

  psProps: ParaProps[] = [{}];
  private psIdx = new Map<string, number>([[psKey({}), 0]]);

  bfData: BfEntry[] = [];
  private bfIdx = new Map<string, number>();

  constructor() {
    this.addBorderFill(this.DEF_STROKE); // bfId=1
  }

  font(name: string): number {
    const n = safeFontToKr(name) || '맑은 고딕';
    if (this.fontIdx.has(n)) return this.fontIdx.get(n)!;
    const id = this.fonts.length;
    this.fonts.push(n);
    this.fontIdx.set(n, id);
    return id;
  }

  addCharShape(p: TextProps): number {
    const k = csKey(p);
    if (this.csIdx.has(k)) return this.csIdx.get(k)!;
    const id = this.csProps.length;
    this.csProps.push(p);
    this.csIdx.set(k, id);
    if (p.font) this.font(p.font);
    return id;
  }

  addParaShape(p: ParaProps): number {
    const k = psKey(p);
    if (this.psIdx.has(k)) return this.psIdx.get(k)!;
    const id = this.psProps.length;
    this.psProps.push(p);
    this.psIdx.set(k, id);
    return id;
  }

  /** Returns 1-based border fill ID */
  addBorderFill(s: Stroke, bg?: string): number {
    const k = bfKey(s, bg);
    if (this.bfIdx.has(k)) return this.bfIdx.get(k)!;
    const id = this.bfData.length + 1;
    this.bfData.push({ uniform: true, s, bg });
    this.bfIdx.set(k, id);
    return id;
  }

  addBorderFillPerSide(l: Stroke, r: Stroke, t: Stroke, b: Stroke, bg?: string): number {
    const k = bfPerSideKey(l, r, t, b, bg);
    if (this.bfIdx.has(k)) return this.bfIdx.get(k)!;
    const id = this.bfData.length + 1;
    this.bfData.push({ uniform: false, l, r, t, b, bg });
    this.bfIdx.set(k, id);
    return id;
  }
}

function collectNode(node: ContentNode, col: StyleCollector): void {
  if (node.tag === 'para') {
    col.addParaShape(node.props);
    for (const kid of node.kids) {
      if (kid.tag === 'span') col.addCharShape((kid as SpanNode).props);
    }
  } else if (node.tag === 'grid') {
    if (node.props.defaultStroke) col.addBorderFill(node.props.defaultStroke);
    for (const row of node.kids) {
      for (const cell of row.kids) {
        const defStroke = node.props.defaultStroke ?? col.DEF_STROKE;
        const cp = cell.props;
        if (cp.top || cp.bot || cp.left || cp.right) {
          col.addBorderFillPerSide(
            cp.left ?? defStroke, cp.right ?? defStroke,
            cp.top ?? defStroke, cp.bot ?? defStroke,
            cp.bg,
          );
        } else {
          col.addBorderFill(defStroke, cp.bg);
        }
        for (const para of cell.kids) collectNode(para, col);
      }
    }
  }
}

/* ═══════════════════════════════════════════════════════════════
   DocInfo record builders
   ═══════════════════════════════════════════════════════════════ */

/**
 * HWPTAG_DOCUMENT_PROPERTIES (표 14 문서 속성) — 26 bytes
 * level 0 in DocInfo
 */
function mkDocumentProperties(): Uint8Array {
  return new BufWriter()
    .u16(1)   // nSection (구역 개수)
    .u16(1)   // pageStart (페이지 시작 번호)
    .u16(1)   // footnoteStart (각주 시작 번호)
    .u16(1)   // endnoteStart (미주 시작 번호)
    .u16(1)   // pictureStart (그림 시작 번호)
    .u16(1)   // tableStart (표 시작 번호)
    .u16(1)   // formulaStart (수식 시작 번호)
    .u32(0)   // listId (캐럿 위치: 리스트 아이디)
    .u32(0)   // paraId (캐럿 위치: 문단 아이디)
    .u32(0)   // charUnitPos (캐럿 위치: 글자 단위 위치)
    .build(); // 26 bytes
}

/**
 * HWPTAG_ID_MAPPINGS (표 15 아이디 매핑 헤더) — 72 bytes (INT32 array[18])
 * 표 16 아이디 매핑 개수 인덱스:
 *   [0]=binData, [1]=한글글꼴, [2]=영어글꼴, [3]=한자글꼴, [4]=일어글꼴,
 *   [5]=기타글꼴, [6]=기호글꼴, [7]=사용자글꼴,
 *   [8]=테두리/배경, [9]=글자모양, [10]=탭정의, [11]=문단번호,
 *   [12]=글머리표, [13]=문단모양, [14]=스타일, [15]=메모모양,
 *   [16]=변경추적, [17]=변경추적사용자
 * level 1 in DocInfo
 */
function mkIdMappings(col: StyleCollector, nBinData = 0): Uint8Array {
  const w = new BufWriter();
  w.u32(nBinData);             // [0] nBinData
  // [1-7]: 7 font language categories — all use same font list
  for (let i = 0; i < 7; i++) w.u32(col.fonts.length);
  w.u32(col.bfData.length);    // [8] nBorderFill
  w.u32(col.csProps.length);   // [9] nCharShape
  w.u32(0);                    // [10] nTabDef
  w.u32(0);                    // [11] nNumbering
  w.u32(0);                    // [12] nBullet
  w.u32(col.psProps.length);   // [13] nParaShape
  w.u32(0);                    // [14] nStyle
  w.u32(0);                    // [15] nMemoShape (5.0.2.1+)
  w.u32(0);                    // [16] nTrackChange (5.0.3.2+)
  w.u32(0);                    // [17] nTrackChangeAuthor (5.0.3.2+)
  return w.build(); // 18 × 4 = 72 bytes
}

/**
 * HWPTAG_FACE_NAME (표 19 글꼴) — variable length
 * Complete structure: attr + name + altFontType + altLen + fontTypeInfo[10] + basicLen
 * level 1 in DocInfo
 */
function mkFaceName(name: string): Uint8Array {
  return new BufWriter()
    .u8(0)              // attr (0 = no alt/base font available)
    .u16(name.length)   // len1: 글꼴 이름 길이
    .utf16(name)        // 글꼴 이름 (UTF-16LE)
    .u8(0)              // altFontType: 대체 글꼴 유형 (0)
    .u16(0)             // altFontLen (0 = no alt font name)
    .zeros(10)          // fontTypeInfo[10]: 글꼴 유형 정보 (표 22)
    .u16(0)             // basicFontLen (0 = no basic font name)
    .build();
}

function borderWidthIdx(pt: number): number {
  let best = 0;
  for (let i = 0; i < BORDER_W_PT.length; i++) {
    if (Math.abs(BORDER_W_PT[i] - pt) < Math.abs(BORDER_W_PT[best] - pt)) best = i;
  }
  return best;
}

/**
 * HWPTAG_BORDER_FILL (표 23 테두리/배경) — 32+n bytes
 * Field layout (GROUPED, not interleaved):
 *   offset 0:  UINT16 attr
 *   offset 2:  UINT8[4] border types  [left, right, top, bottom]
 *   offset 6:  UINT8[4] border widths [left, right, top, bottom]
 *   offset 10: COLORREF[4] border colors [left, right, top, bottom]
 *   offset 26: UINT8 diagonal type
 *   offset 27: UINT8 diagonal width
 *   offset 28: COLORREF diagonal color
 *   offset 32: fill info (채우기 정보, 표 28)
 * level 1 in DocInfo
 */
function mkBorderFill(s: Stroke, bg?: string): Uint8Array {
  const w = new BufWriter();
  const t  = BORDER_KIND_IDX[s.kind] ?? 0;
  const wi = borderWidthIdx(s.pt);
  const col = s.color || '000000';

  w.u16(0); // attr (UINT16)
  // 4 border types: [left, right, top, bottom]
  for (let i = 0; i < 4; i++) w.u8(t);
  // 4 border widths: [left, right, top, bottom]
  for (let i = 0; i < 4; i++) w.u8(wi);
  // 4 border colors: [left, right, top, bottom]
  for (let i = 0; i < 4; i++) w.colorRef(col);
  // diagonal: type, width, color
  w.u8(0).u8(0).colorRef('000000');
  // fill info (채우기 정보, 표 28): type + optional color data
  if (bg) {
    w.u32(1);              // type = 0x01 (단색 채우기)
    w.colorRef(bg);        // 배경색
    w.colorRef('FFFFFF');  // 무늬색 (no pattern = white)
    w.u32(0);              // 무늬 종류 (0 = none)
  } else {
    w.u32(0);              // type = 0x00 (채우기 없음)
  }
  return w.build(); // 36 bytes (no fill) or 48 bytes (solid fill)
}

function mkBorderFillPerSide(
  left: Stroke, right: Stroke, top: Stroke, bottom: Stroke, bg?: string,
): Uint8Array {
  const w = new BufWriter();
  w.u16(0); // attr
  // 4 border types [left, right, top, bottom]
  w.u8(BORDER_KIND_IDX[left.kind]   ?? 0);
  w.u8(BORDER_KIND_IDX[right.kind]  ?? 0);
  w.u8(BORDER_KIND_IDX[top.kind]    ?? 0);
  w.u8(BORDER_KIND_IDX[bottom.kind] ?? 0);
  // 4 border widths
  w.u8(borderWidthIdx(left.pt));
  w.u8(borderWidthIdx(right.pt));
  w.u8(borderWidthIdx(top.pt));
  w.u8(borderWidthIdx(bottom.pt));
  // 4 border colors
  w.colorRef(left.color   || '000000');
  w.colorRef(right.color  || '000000');
  w.colorRef(top.color    || '000000');
  w.colorRef(bottom.color || '000000');
  // diagonal
  w.u8(0).u8(0).colorRef('000000');
  // fill info
  if (bg) {
    w.u32(1).colorRef(bg).colorRef('FFFFFF').u32(0);
  } else {
    w.u32(0);
  }
  return w.build();
}

/**
 * HWPTAG_CHAR_SHAPE (표 33 글자 모양) — 72+ bytes
 * Field layout:
 *   offset  0: WORD[7]  faceId (언어별 글꼴 ID) — 14 bytes
 *   offset 14: UINT8[7] ratio  (장평 50~200%)   —  7 bytes
 *   offset 21: INT8[7]  spacing(자간 -50~50%)   —  7 bytes
 *   offset 28: UINT8[7] relSize(상대크기 10~250%)—  7 bytes
 *   offset 35: INT8[7]  offset (글자위치)       —  7 bytes
 *   offset 42: INT32    height (기준크기, HWP단위 = pt×100)
 *   offset 46: UINT32   attr   (표 35 글자 모양 속성)
 *   offset 50: INT8     shadowX
 *   offset 51: INT8     shadowY
 *   offset 52: COLORREF textColor (글자 색)
 *   offset 56: COLORREF underlineColor (밑줄 색)
 *   offset 60: COLORREF shadeColor (음영 색)
 *   offset 64: COLORREF shadowColor (그림자 색)
 *   offset 68: UINT16   borderFillId (5.0.2.1+)
 *   offset 70: COLORREF strikeColor (취소선 색, 5.0.3.0+)
 * level 1 in DocInfo
 */
function mkCharShape(p: TextProps, col: StyleCollector): Uint8Array {
  const fontId = p.font ? col.font(p.font) : 0;
  const height = Math.round((p.pt ?? 10) * 100); // HWP단위: pt×100

  // 표 35 글자 모양 속성:
  //  bit 0: 기울임(italic), bit 1: 진하게(bold)
  //  bits 2-4: 밑줄 종류 (0=없음, 1=글자아래, 2=글자위)
  //  bits 5-8: 밑줄 모양 (0=실선, 1=2중선 등)
  //  bits 16-17: 위/아래 첨자 (0=없음, 1=위첨자, 2=아래첨자)
  //  bits 18-20: 취소선 종류 (1=실선)
  //  bits 21-24: 취소선 모양
  let attr = 0;
  if (p.i)   attr |= (1 << 0);
  if (p.b)   attr |= (1 << 1);
  if (p.u)   {
    attr |= (1 << 2);   // 밑줄 종류 = 1 (글자 아래)
    attr |= (0 << 5);   // 밑줄 모양 = 0 (실선)
  }
  if (p.s)   {
    attr |= (1 << 18);  // 취소선 종류 = 1
    attr |= (0 << 21);  // 취소선 모양 = 0 (실선)
  }
  if (p.sup) attr |= (1 << 16);  // 위 첨자
  if (p.sub) attr |= (2 << 16);  // 아래 첨자

  const textColor = p.color ?? '000000';

  const w = new BufWriter();
  // faceId[7]: WORD[7] — 언어별 글꼴 ID (한글, 영어, 한자, 일어, 기타, 기호, 사용자)
  for (let i = 0; i < 7; i++) w.u16(fontId);
  // ratio[7]: UINT8[7] = 100 (100%)
  for (let i = 0; i < 7; i++) w.u8(100);
  // spacing[7]: INT8[7] = 0 (자간 0%)
  for (let i = 0; i < 7; i++) w.u8(0);
  // relSize[7]: UINT8[7] = 100 (100%)
  for (let i = 0; i < 7; i++) w.u8(100);
  // offset[7]: INT8[7] = 0 (글자 위치 0%)
  for (let i = 0; i < 7; i++) w.u8(0);
  // height: INT32
  w.i32(height);
  // attr: UINT32
  w.u32(attr);
  // shadowX, shadowY: INT8
  w.u8(0).u8(0);
  // textColor: COLORREF
  w.colorRef(textColor);
  // underlineColor: COLORREF (밑줄 색)
  w.colorRef('000000');
  // shadeColor: COLORREF (음영 색) — FFFFFF = no shade
  w.colorRef('FFFFFF');
  // shadowColor: COLORREF (그림자 색)
  w.colorRef('000000');
  // borderFillId: UINT16 (글자 테두리/배경 ID, 5.0.2.1+)
  w.u16(0);
  // strikeColor: COLORREF (취소선 색, 5.0.3.0+)
  w.colorRef('000000');
  return w.build(); // 42 + 4 + 4 + 1 + 1 + 4×4 + 2 + 4 = 74 bytes
}

/**
 * HWPTAG_PARA_SHAPE (표 43 문단 모양) — 54 bytes
 * Field layout:
 *   offset  0: UINT32 attr1   (표 44 문단 모양 속성1, bits 2-4 = 정렬 방식)
 *   offset  4: INT32  leftMargin
 *   offset  8: INT32  rightMargin
 *   offset 12: INT32  indent (들여/내어 쓰기)
 *   offset 16: INT32  spaceBefore (문단 간격 위)
 *   offset 20: INT32  spaceAfter  (문단 간격 아래)
 *   offset 24: INT32  lineSpacing (줄 간격, 한글 2007 이하, 5.0.2.5 미만)
 *   offset 28: UINT16 tabDefId
 *   offset 30: UINT16 numberingId/bulletId
 *   offset 32: UINT16 borderFillId
 *   offset 34: INT16  borderLeft
 *   offset 36: INT16  borderRight
 *   offset 38: INT16  borderTop
 *   offset 40: INT16  borderBottom
 *   offset 42: UINT32 attr2 (5.0.1.7+)
 *   offset 46: UINT32 attr3 (5.0.2.5+, bits 0-4 = 줄 간격 종류)
 *   offset 50: UINT32 lineSpacing2 (5.0.2.5+, 실제 줄 간격)
 * level 1 in DocInfo
 */
function mkParaShape(p: ParaProps): Uint8Array {
  // 표 44 bits 2-4: 정렬 방식 (0=양쪽, 1=왼쪽, 2=오른쪽, 3=가운데)
  const alignVal = ALIGN_CODE[p.align ?? 'left'] ?? 1;
  const attr1 = (alignVal & 0x7) << 2; // alignment at bits 2-4

  // 줄 간격: 비율(%) 표현, 160 = 160%
  const lineSpacePct = p.lineHeight ? Math.round(p.lineHeight * 100) : 160;

  return new BufWriter()
    .u32(attr1)                                    //  0: attr1
    .i32(Metric.ptToHwp(p.indentPt ?? 0))          //  4: leftMargin (HWPUNIT)
    .i32(0)                                        //  8: rightMargin
    .i32(Metric.ptToHwp(p.firstLineIndentPt ?? 0)) // 12: indent (첫 줄 들여/내어쓰기)
    .i32(Metric.ptToHwp(p.spaceBefore ?? 0))       // 16: spaceBefore
    .i32(Metric.ptToHwp(p.spaceAfter ?? 0))        // 20: spaceAfter
    .i32(lineSpacePct)                             // 24: lineSpacing (old format)
    .u16(0)                                        // 28: tabDefId
    .u16(0)                                        // 30: numberingId
    .u16(0)                                        // 32: borderFillId
    .i16(0)                                        // 34: borderLeft
    .i16(0)                                        // 36: borderRight
    .i16(0)                                        // 38: borderTop
    .i16(0)                                        // 40: borderBottom
    .u32(0)                                        // 42: attr2 (5.0.1.7+)
    .u32(4)                                        // 46: attr3 (bits 0-4 = 4: 백분율 줄 간격)
    .u32(lineSpacePct)                             // 50: lineSpacing2 (5.0.2.5+)
    .build(); // 54 bytes
}

/**
 * HWPTAG_BIN_DATA (표 17 바이너리 데이터) — variable
 * Storage type 2 = embedded in BinData storage
 * level 1 in DocInfo
 */
function mkBinData(id: number, ext: string): Uint8Array {
  const w = new BufWriter();
  w.u16(0x0002);        // attr: storage type = 2 (내장 파일)
  w.u16(id);            // binDataId (1-based)
  w.u16(ext.length);    // extLen (number of UTF-16LE chars)
  w.utf16(ext);         // extension string (e.g. "jpg", "png")
  return w.build();
}

interface BinImage {
  id: number;       // 1-based BIN ID
  ext: string;      // file extension without dot
  data: Uint8Array; // raw image bytes
}

function buildDocInfoStream(col: StyleCollector, images: BinImage[] = []): Uint8Array {
  const chunks: Uint8Array[] = [];

  // HWPTAG_DOCUMENT_PROPERTIES at level 0 (required first record)
  chunks.push(mkRec(TAG_DOCUMENT_PROPERTIES, 0, mkDocumentProperties()));

  // HWPTAG_ID_MAPPINGS at level 1 (child of DOCUMENT_PROPERTIES)
  chunks.push(mkRec(TAG_ID_MAPPINGS, 1, mkIdMappings(col, images.length)));

  // HWPTAG_BIN_DATA at level 1: one record per image
  for (const img of images) {
    chunks.push(mkRec(TAG_BIN_DATA, 1, mkBinData(img.id, img.ext)));
  }

  // HWPTAG_FACE_NAME at level 1: 7 language categories × nFonts records
  // Order: all hangul fonts, then all english fonts, ..., then all user fonts
  for (let cat = 0; cat < 7; cat++) {
    for (const name of col.fonts) {
      chunks.push(mkRec(TAG_FACE_NAME, 1, mkFaceName(name)));
    }
  }

  // HWPTAG_BORDER_FILL at level 1
  for (const entry of col.bfData) {
    chunks.push(mkRec(TAG_BORDER_FILL, 1,
      entry.uniform
        ? mkBorderFill(entry.s, entry.bg)
        : mkBorderFillPerSide(entry.l, entry.r, entry.t, entry.b, entry.bg)));
  }

  // HWPTAG_CHAR_SHAPE at level 1
  for (const p of col.csProps) {
    chunks.push(mkRec(TAG_CHAR_SHAPE, 1, mkCharShape(p, col)));
  }

  // HWPTAG_PARA_SHAPE at level 1
  for (const p of col.psProps) {
    chunks.push(mkRec(TAG_PARA_SHAPE, 1, mkParaShape(p)));
  }

  return concatU8(chunks);
}

/* ═══════════════════════════════════════════════════════════════
   BodyText record builders
   ═══════════════════════════════════════════════════════════════ */

function mkPageDef(dims: PageDims): Uint8Array {
  return new BufWriter()
    .u32(Metric.ptToHwp(dims.wPt))
    .u32(Metric.ptToHwp(dims.hPt))
    .u32(Metric.ptToHwp(dims.ml))
    .u32(Metric.ptToHwp(dims.mr))
    .u32(Metric.ptToHwp(dims.mt))
    .u32(Metric.ptToHwp(dims.mb))
    .zeros(12) // header/footer/gutter margins (3 × INT32)
    .u32(dims.orient === 'landscape' ? 1 : 0) // attr @ offset 36
    .build(); // 40 bytes
}

/**
 * HWPTAG_PARA_HEADER (표 58 문단 헤더) — 24 bytes
 *   offset  0: UINT32 nchars       (문단 내 글자 수, 컨트롤 포함)
 *   offset  4: UINT32 ctrlMask     (컨트롤 마스크: 1<<ctrlCode)
 *   offset  8: UINT16 paraShapeId  (문단 모양 아이디)
 *   offset 10: UINT8  styleId      (문단 스타일 아이디, 0=기본)
 *   offset 11: UINT8  columnBreak  (단 나누기 종류)
 *   offset 12: UINT16 csCount      (글자 모양 정보 수)
 *   offset 14: UINT16 rangeTagCount(range tag 정보 수)
 *   offset 16: UINT16 lineAlignCount(각 줄 align 정보 수)
 *   offset 18: UINT32 instanceId   (문단 Instance ID, unique)
 *   offset 22: UINT16 trackChange  (변경추적 병합 여부, 5.0.3.2+)
 */
function mkParaHeader(
  nchars: number,
  ctrlMask: number,
  psId: number,
  csCount: number,
  lineAlignCount: number = 0,
  instanceId: number = 0,
): Uint8Array {
  return new BufWriter()
    .u32(nchars)        //  0: nchars
    .u32(ctrlMask)      //  4: ctrlMask
    .u16(psId)          //  8: paraShapeId
    .u8(0)              // 10: styleId (0 = default)
    .u8(0)              // 11: columnBreak (0 = none)
    .u16(csCount)       // 12: charShapeCount
    .u16(0)             // 14: rangeTagCount
    .u16(lineAlignCount)// 16: lineAlignCount
    .u32(instanceId)    // 18: instanceId (unique)
    .u16(0)             // 22: trackChange
    .build(); // 24 bytes
}

function mkParaText(text: string): Uint8Array {
  const w = new BufWriter();
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    // 수정 이유(수정 2): HWP 5.0 스펙에서 TAB(0x09)·LF(0x0A)는 특수 WCHAR으로
    // 한컴오피스가 직접 해석하므로 원값 그대로 기록.
    // 나머지 0x01~0x1F 제어문자는 0으로 치환 (이전 동작 유지).
    if (c === 0x09 || c === 0x0A) {
      w.u16(c); // TAB·LF는 그대로 유지
    } else {
      w.u16(c < 32 ? 0 : c); // 기타 제어문자 → 0 치환
    }
  }
  w.u16(13); // paragraph terminator (0x000D = para break)
  return w.build();
}

function mkParaCharShape(pairs: [pos: number, id: number][]): Uint8Array {
  const w = new BufWriter();
  for (const [pos, id] of pairs) w.u32(pos).u32(id);
  return w.build();
}

function mkLineSeg(availWidthHwp: number, fontHwp = 1000): Uint8Array {
  // HWP 5.0 스펙: PARA_LINE_SEG 32bytes (8 × INT32 LE)
  // vertsize = fontHwp * 1.6 (160% 줄 간격 기본값)
  const vertSize = Math.round(fontHwp * 1.6);
  const spacing  = vertSize - fontHwp;
  const baseline = Math.round(fontHwp * 0.85);
  return new BufWriter()
    .u32(0)             //  0: textpos  (텍스트 시작 위치)
    .i32(0)             //  4: vertpos  (수직 위치)
    .i32(vertSize)      //  8: vertsize (줄 세로 크기)
    .i32(fontHwp)       // 12: textheight
    .i32(baseline)      // 16: baseline
    .i32(spacing)       // 20: spacing  (줄 간격 여분)
    .i32(0)             // 24: horzpos
    .i32(availWidthHwp) // 28: horzsize (가용 너비)
    .build();
}

/**
 * PARA_TEXT for section definition paragraph.
 * 확장 컨트롤(size=8): ctrl_code(1) + ctrlId_lo(1) + ctrlId_hi(1) + ptr[4] + ctrl_code(1)
 * + para terminator = 9 WCHAR total → nchars = 9
 */
function mkSecdParaText(): Uint8Array {
  const lo = CTRL_SECD & 0xFFFF;
  const hi = (CTRL_SECD >>> 16) & 0xFFFF;
  return new BufWriter()
    .u16(0x0002).u16(lo).u16(hi).u16(0).u16(0).u16(0).u16(0).u16(0x0002) // 8 WCHAR ctrl ref
    .u16(0x000D)  // para terminator
    .build(); // 9 WCHAR = 18 bytes
}

/**
 * PARA_TEXT for table container paragraph.
 * ctrl code for 그리기 개체/표 = 11 (0x000B)
 * 확장 컨트롤(size=8): 0x000B + tbl_lo + tbl_hi + 4×0 + 0x000B + terminator = 9 WCHAR
 */
function mkTableParaText(): Uint8Array {
  const lo = CTRL_TABLE & 0xFFFF;
  const hi = (CTRL_TABLE >>> 16) & 0xFFFF;
  return new BufWriter()
    .u16(0x000B).u16(lo).u16(hi).u16(0).u16(0).u16(0).u16(0).u16(0x000B) // 8 WCHAR ctrl ref
    .u16(0x000D)  // para terminator
    .build(); // 9 WCHAR = 18 bytes
}

/**
 * PARA_TEXT for picture container paragraph.
 * ctrl code for 그리기 개체/표/수식 등 = 11 (0x000B)
 * 확장 컨트롤(size=8): 0x000B + pic_lo + pic_hi + 4×0 + 0x000B + terminator = 9 WCHAR
 */
function mkPicParaText(): Uint8Array {
  const lo = CTRL_PIC & 0xFFFF;
  const hi = (CTRL_PIC >>> 16) & 0xFFFF;
  return new BufWriter()
    .u16(0x000B).u16(lo).u16(hi).u16(0).u16(0).u16(0).u16(0).u16(0x000B)
    .u16(0x000D)  // para terminator
    .build(); // 9 WCHAR = 18 bytes
}

/**
 * HWPTAG_SHAPE_COMPONENT_PICTURE (tag 85) — 97 bytes
 * Contains drawing common attrs + picture-specific attrs (표 107)
 * Critical field: BINDATA_ID at offset 87 (2 bytes)
 */
function mkShapeComponentPicture(binDataId: number, wHwp: number, hHwp: number): Uint8Array {
  const w = new BufWriter();
  // Drawing common attributes (19 bytes)
  w.u32(CTRL_PIC);  // ctrlId: '$pic' (4)
  w.zeros(15);      // border style, fill type (all default/none)
  // Picture attributes (표 107, 78 bytes total = up to offset 97)
  // Geometric region (HWPUNIT bounds)
  w.u32(0);         // left (4)
  w.u32(0);         // top (4)
  w.u32(wHwp);      // right = width (4)
  w.u32(hHwp);      // bottom = height (4)
  // Drawing bounding box
  w.u32(0);         // left (4)
  w.u32(0);         // top (4)
  w.u32(wHwp);      // right (4)
  w.u32(hHwp);      // bottom (4)
  // Rotation / effect (zeros = no rotation, no effect)
  w.zeros(20);      // rotate angle, rotation center X/Y, effect flags (20)
  // imageInfo (5 bytes at offset 87 from start)
  // Currently at offset: 19 + 4+4+4+4 + 4+4+4+4 + 20 = 19+16+16+20 = 71... need 87
  // Gap: 87 - 71 = 16 more bytes of zeros before binDataId
  w.zeros(16);      // padding to reach offset 87
  w.u16(binDataId); // BINDATA_ID (2) — which BIN stream to use
  w.u8(0);          // image type (0 = stored in BinData storage)
  w.u8(0);          // flags
  w.u8(0);          // padding
  w.zeros(5);       // remaining (to reach 97 bytes total)
  // Total: 19+16+16+20+16+2+1+1+1+5 = 97
  return w.build(); // 97 bytes
}

/** Encode an image node as a picture paragraph (level lv) */
function encodePicPara(
  imgNode: { b64: string; mime: string; w: number; h: number },
  binDataId: number,
  col: StyleCollector,
  lv: number,
  idGen: () => number,
  availWidthHwp: number,
): Uint8Array[] {
  // imgNode.w / imgNode.h are in pt (set by all decoders)
  const wHwp = Metric.ptToHwp(Math.max(imgNode.w, 10));
  const hHwp = Metric.ptToHwp(Math.max(imgNode.h, 10));

  const TABLE_CTRL_MASK = 1 << 11;
  const instanceId = idGen();
  const psId = col.addParaShape({});

  return [
    mkRec(TAG_PARA_HEADER,     lv,     mkParaHeader(9, TABLE_CTRL_MASK, psId, 1, 1, instanceId)),
    mkRec(TAG_PARA_TEXT,       lv + 1, mkPicParaText()),
    mkRec(TAG_PARA_CHAR_SHAPE, lv + 1, mkParaCharShape([[0, 0]])),
    mkRec(TAG_PARA_LINE_SEG,   lv + 1, mkLineSeg(availWidthHwp, Metric.ptToHwp(imgNode.h))),
    // $pic CTRL_HEADER (GHDR) at level lv+1
    mkRec(TAG_CTRL_HEADER, lv + 1, mkPicCtrl(wHwp, hHwp, idGen())),
    // SHAPE_COMPONENT_PICTURE at level lv+2
    mkRec(TAG_SHAPE_COMPONENT_PICTURE, lv + 2, mkShapeComponentPicture(binDataId, wHwp, hHwp)),
  ];
}

function encodePara(
  para: ParaNode, col: StyleCollector, lv: number,
  instanceId: number, availWidthHwp: number,
): Uint8Array[] {
  let text = '';
  const csPairs: [number, number][] = [];
  let pos = 0;

  // 첫 번째 span의 폰트 크기를 fontHwp로 계산
  let fontHwp = 1000;
  for (const kid of para.kids) {
    if (kid.tag === 'span') {
      const pt = (kid as SpanNode).props.pt;
      if (pt && pt > 0) { fontHwp = Metric.ptToHwp(pt); break; }
    }
  }

  for (const kid of para.kids) {
    if (kid.tag === 'span') {
      const span = kid as SpanNode;
      const csId = col.addCharShape(span.props);
      if (csPairs.length === 0 || csPairs[csPairs.length - 1][1] !== csId) {
        csPairs.push([pos, csId]);
      }
      for (const t of span.kids) {
        if (t.tag === 'txt') { text += t.content; pos += t.content.length; }
      }
    }
  }

  if (csPairs.length === 0) csPairs.push([0, 0]);

  const psId  = col.addParaShape(para.props);
  const nchars = text.length + 1; // text + para terminator

  return [
    mkRec(TAG_PARA_HEADER,     lv,     mkParaHeader(nchars, 0, psId, csPairs.length, 1, instanceId)),
    mkRec(TAG_PARA_TEXT,       lv + 1, mkParaText(text)),
    mkRec(TAG_PARA_CHAR_SHAPE, lv + 1, mkParaCharShape(csPairs)),
    mkRec(TAG_PARA_LINE_SEG,   lv + 1, mkLineSeg(availWidthHwp, fontHwp)),
  ];
}

/* ── Table encoding ─────────────────────────────────────────── */

/**
 * CTRL_HEADER for table ctrl — 46-byte GHDR (개체 공통 속성, 표 68)
 * Layout: ctrlId(4)+attr(4)+vOff(4)+hOff(4)+w(4)+h(4)+z(4)+margins(4×u16=8)+instanceId(4)+pageBreak(4)+captionLen(2)
 * 수정 이유(수정 5): 바이트 수 검증 결과 이미 정확히 46바이트임을 확인.
 *   4+4+4+4+4+4+4+8+4+4+2 = 46바이트 (HWP 5.0 스펙 표 68 일치, 패딩 없음)
 */
function mkTableCtrl(wHwp: number, hHwp: number, instanceId: number): Uint8Array {
  return new BufWriter()
    .u32(CTRL_TABLE)   // ctrlId: 'tbl ' (+4 → 4)
    .u32(0x082a2211)   // attr — 실 파일 참조값 (+4 → 8)
    .i32(0)            // vOff (+4 → 12)
    .i32(0)            // hOff (+4 → 16)
    .u32(wHwp)         // width in HWPUNIT (+4 → 20)
    .u32(hHwp)         // height in HWPUNIT (+4 → 24)
    .i32(7)            // z-order — 실 파일 참조값 (+4 → 28)
    .u16(140).u16(140).u16(140).u16(140) // outer margins L/R/T/B (+8 → 36)
    .u32(instanceId)   // instanceId (+4 → 40)
    .i32(0)            // pageBreak (+4 → 44)
    .u16(0)            // captionLen (+2 → 46) ← 스펙 정확, 추가 패딩 없음
    .build();          // 정확히 46 bytes
}

/**
 * CTRL_HEADER for picture ctrl — 46-byte GHDR ($pic)
 */
function mkPicCtrl(wHwp: number, hHwp: number, instanceId: number): Uint8Array {
  return new BufWriter()
    .u32(CTRL_PIC)     // ctrlId: '$pic' (4)
    .u32(0x082a2211)   // attr — same flags as table (4)
    .i32(0)            // vOff (4)
    .i32(0)            // hOff (4)
    .u32(wHwp)         // width in HWPUNIT (4)
    .u32(hHwp)         // height in HWPUNIT (4)
    .i32(0)            // z-order (4)
    .u16(0).u16(0).u16(0).u16(0) // outer margins (8)
    .u32(instanceId)   // instanceId (4)
    .i32(0)            // pageBreak (4)
    .u16(0)            // captionLen (2)
    .build();          // 46 bytes
}

/**
 * HWPTAG_TABLE (표 75) — variable
 * Layout (verified from real HWP file):
 *   UINT32 attr (0x04000006)
 *   UINT16 rowCnt
 *   UINT16 colCnt
 *   UINT16 cellSpacing
 *   UINT16[4] innerMargins: left=510, right=510, top=141, bottom=141 (HWPUNIT16)
 *   UINT16[rowCnt] rowSizes (nominal, actual height from cell LIST_HEADER)
 *   UINT16 borderFillId
 *   UINT16 validZoneCount
 */
function mkTableRecord(rowCnt: number, colCnt: number, rowHwp: number[], bfId: number): Uint8Array {
  const w = new BufWriter();
  w.u32(0x04000006); // attr — from real file
  w.u16(rowCnt);
  w.u16(colCnt);
  w.u16(0);          // cellSpacing
  w.u16(510);        // inner margin left  (0x01fe HWPUNIT16) — from real file
  w.u16(510);        // inner margin right
  w.u16(141);        // inner margin top   (0x008d HWPUNIT16) — from real file
  w.u16(141);        // inner margin bottom
  for (const h of rowHwp) w.u16(Math.max(1, h & 0xFFFF)); // row sizes as UINT16
  w.u16(bfId);       // borderFillId
  w.u16(0);          // validZoneCount
  return w.build();
}

function mkCellListHeader(
  paraCount: number,
  row: number, col: number,
  rs: number,  cs: number,
  wHwp: number, hHwp: number,
  bfId: number,
): Uint8Array {
  // 표 65 문단 리스트 헤더 (cell variant) — 47 bytes from real file:
  // offset 0:  UINT16 paraCount
  // offset 2:  UINT32 attr
  // offset 6:  UINT16 (extended attr)
  // offset 8:  UINT16 colAddr
  // offset 10: UINT16 rowAddr
  // offset 12: UINT16 rowSpan
  // offset 14: UINT16 colSpan
  // offset 16: UINT32 width
  // offset 20: UINT32 height
  // offset 24: UINT16 margin left  (510 = 0x01fe HWPUNIT16)
  // offset 26: UINT16 margin right
  // offset 28: UINT16 margin top   (141 = 0x008d HWPUNIT16)
  // offset 30: UINT16 margin bottom
  // offset 32: UINT16 borderFillId
  // offset 34-46: extended (13 bytes, zeros)
  return new BufWriter()
    .u16(paraCount) //  0: paraCount
    .u32(0)         //  2: attr
    .u16(0)         //  6: extended attr
    .u16(col)       //  8: colAddr
    .u16(row)       // 10: rowAddr
    .u16(rs)        // 12: rowSpan
    .u16(cs)        // 14: colSpan
    .u32(wHwp)      // 16: width
    .u32(hHwp)      // 20: height
    .u16(510)       // 24: margin left  (from real file)
    .u16(510)       // 26: margin right
    .u16(141)       // 28: margin top
    .u16(141)       // 30: margin bottom
    .u16(bfId)      // 32: borderFillId
    .zeros(13)      // 34-46: extended fields
    .build();       // 47 bytes
}

const DEFAULT_ROW_HEIGHT_PT = 14;

function encodeGrid(grid: GridNode, col: StyleCollector, lv: number, idGen: () => number, availWidthHwp: number): Uint8Array[] {
  const records: Uint8Array[] = [];
  const rowCnt = grid.kids.length;
  const colCnt = Math.max(1, grid.kids[0]?.kids.length ?? 1);

  const cwPt = (grid.props as any).colWidths ?? [];
  const totalPt = cwPt.reduce((s: number, w: number) => s + w, 0) || 453;
  const defColPt = totalPt / colCnt;

  const defStroke = grid.props.defaultStroke ?? col.DEF_STROKE;
  const defBfId   = col.addBorderFill(defStroke);

  const rowHwp = grid.kids.map((row: any) =>
    row.heightPt != null && row.heightPt > 0
      ? Metric.ptToHwp(row.heightPt)
      : Metric.ptToHwp(DEFAULT_ROW_HEIGHT_PT));

  // Compute total table dimensions for GHDR
  const totalWPt = cwPt.length > 0 ? cwPt.reduce((s: number, w: number) => s + w, 0) : totalPt;
  const totalHPt = grid.kids.reduce((s: number, row: any) =>
    s + (row.heightPt != null && row.heightPt > 0 ? row.heightPt : DEFAULT_ROW_HEIGHT_PT), 0);
  const tblWHwp = Metric.ptToHwp(totalWPt);
  const tblHHwp = Metric.ptToHwp(totalHPt);
  const tblInstanceId = idGen();

  // Table ctrl header at level lv (CTRL_HEADER for table — full 46-byte GHDR)
  records.push(mkRec(TAG_CTRL_HEADER, lv, mkTableCtrl(tblWHwp, tblHHwp, tblInstanceId)));

  // HWPTAG_TABLE at level lv+1 (child of ctrl header)
  records.push(mkRec(TAG_TABLE, lv + 1, mkTableRecord(rowCnt, colCnt, rowHwp, defBfId)));

  for (let r = 0; r < grid.kids.length; r++) {
    for (let c = 0; c < grid.kids[r].kids.length; c++) {
      const cell   = grid.kids[r].kids[c];
      const wHwp   = Metric.ptToHwp(cwPt[c] ?? defColPt);
      const hHwp   = rowHwp[r];
      const cp     = cell.props;
      const hasPerSide = cp.top || cp.bot || cp.left || cp.right;
      const bfId = hasPerSide
        ? col.addBorderFillPerSide(
            cp.left ?? defStroke, cp.right ?? defStroke,
            cp.top ?? defStroke, cp.bot ?? defStroke,
            cp.bg,
          )
        : col.addBorderFill(defStroke, cp.bg);

      const paras = cell.kids.length > 0
        ? cell.kids
        : [{ tag: 'para' as const, props: {}, kids: [] }];

      // LIST_HEADER at level lv+1 (sibling of TABLE)
      records.push(mkRec(TAG_LIST_HEADER, lv + 1,
        mkCellListHeader(paras.length, r, c, cell.rs, cell.cs, wHwp, hHwp, bfId)));

      // Cell paragraphs at level lv+2 (children of LIST_HEADER)
      const cellWidthHwp = Metric.ptToHwp(cwPt[c] ?? defColPt);
      for (const para of paras) {
        records.push(...encodePara(para as ParaNode, col, lv + 2, idGen(), cellWidthHwp));
      }
    }
  }

  return records;
}

/**
 * CTRL_HEADER for section definition — 47 bytes (matched from real HWP files)
 * Layout: ctrlId(4) + attr(4) + colGap(4) + u16+u16(4) + zeros(31)
 */
function mkSectionCtrl(): Uint8Array {
  return new BufWriter()
    .u32(CTRL_SECD)  // ctrlId: 'secd' (4)
    .u32(0)           // attr (4)
    .u32(1134)        // column gap in HWPUNIT — value from real file (4)
    .u16(0x4000)      // (2) — from real file
    .u16(0x001f)      // (2) — from real file
    .zeros(31)        // padding to reach 47 bytes total
    .build();         // 4+4+4+2+2+31 = 47 bytes
}

/**
 * Section definition paragraph structure:
 * Level 0: PARA_HEADER (nchars=9, ctrlMask=1<<2, lineAlignCount=1)
 * Level 1:   PARA_TEXT (secd ctrl reference + terminator, 9 WCHAR)
 * Level 1:   PARA_CHAR_SHAPE
 * Level 1:   PARA_LINE_SEG (32 bytes for 1 line)
 * Level 1:   CTRL_HEADER ('secd', 4 bytes)
 * Level 2:     PAGE_DEF (40 bytes)
 * Level 2:     FOOTNOTE_SHAPE (28 bytes, for footnotes)
 * Level 2:     FOOTNOTE_SHAPE (28 bytes, for endnotes)
 */
function buildSectionParagraph(dims: PageDims, instanceId: number): Uint8Array[] {
  const SECD_CTRL_MASK = 1 << 2; // code 2 = 구역 정의/단 정의
  const nchars = 9; // 8 ctrl wchars + 1 terminator
  const availWidthHwp = Math.max(1000,
    Metric.ptToHwp(dims.wPt) - Metric.ptToHwp(dims.ml) - Metric.ptToHwp(dims.mr),
  );

  return [
    mkRec(TAG_PARA_HEADER,     0, mkParaHeader(nchars, SECD_CTRL_MASK, 0, 1, 1, instanceId)),
    mkRec(TAG_PARA_TEXT,       1, mkSecdParaText()),
    mkRec(TAG_PARA_CHAR_SHAPE, 1, mkParaCharShape([[0, 0]])),
    mkRec(TAG_PARA_LINE_SEG,   1, mkLineSeg(availWidthHwp, 1000)),
    mkRec(TAG_CTRL_HEADER,     1, mkSectionCtrl()),    // 'secd' at level 1
    mkRec(TAG_PAGE_DEF,        2, mkPageDef(dims)),    // page def at level 2
    mkRec(TAG_FOOTNOTE_SHAPE,  2, new Uint8Array(28)), // footnote shape (defaults)
    mkRec(TAG_FOOTNOTE_SHAPE,  2, new Uint8Array(28)), // endnote shape (defaults)
  ];
}

function buildBodyTextStream(
  doc: DocRoot,
  col: StyleCollector,
  images: BinImage[],
): Uint8Array {
  const chunks: Uint8Array[] = [];
  const dims = doc.kids[0]?.dims ?? A4;
  let instanceIdCounter = 1;
  const idGen = () => instanceIdCounter++;
  const availWidthHwp = Math.max(1000,
    Metric.ptToHwp(dims.wPt) - Metric.ptToHwp(dims.ml) - Metric.ptToHwp(dims.mr),
  );

  // Section definition paragraph
  for (const r of buildSectionParagraph(dims, idGen())) chunks.push(r);

  const TABLE_CTRL_MASK = 1 << 11; // code 11 = 그리기 개체/표 (확장 컨트롤)

  for (const sheet of doc.kids) {
    for (const node of sheet.kids) {
      if (node.tag === 'para') {
        const para = node as ParaNode;
        // Check if the paragraph contains images
        const hasImg = para.kids.some((k: any) => k.tag === 'img');
        if (hasImg) {
          // Emit one image paragraph per img node found
          for (const kid of para.kids) {
            if ((kid as any).tag === 'img') {
              const img = kid as any;
              // Find matching BinImage
              const binImg = images.find(b => b64Matches(b, img.b64));
              if (binImg) {
                for (const r of encodePicPara(img, binImg.id, col, 0, idGen, availWidthHwp)) chunks.push(r);
              }
            }
          }
          // Also emit any text in the paragraph (spans)
          const textKids = para.kids.filter((k: any) => k.tag !== 'img');
          if (textKids.length > 0) {
            const textPara: ParaNode = { tag: 'para', props: para.props, kids: textKids as any };
            for (const r of encodePara(textPara, col, 0, idGen(), availWidthHwp)) chunks.push(r);
          }
        } else {
          for (const r of encodePara(para, col, 0, idGen(), availWidthHwp)) chunks.push(r);
        }
      } else if (node.tag === 'grid') {
        // Table container paragraph at level 0
        // nchars=9 (8 ctrl wchars + terminator), ctrlMask has bit 11 set for table ctrl
        chunks.push(mkRec(TAG_PARA_HEADER,     0, mkParaHeader(9, TABLE_CTRL_MASK, 0, 1, 1, idGen())));
        chunks.push(mkRec(TAG_PARA_TEXT,       1, mkTableParaText()));
        chunks.push(mkRec(TAG_PARA_CHAR_SHAPE, 1, mkParaCharShape([[0, 0]])));
        chunks.push(mkRec(TAG_PARA_LINE_SEG,   1, mkLineSeg(availWidthHwp, 1000)));
        // Table ctrl at level 1, TABLE record and cells at levels 2/3
        for (const r of encodeGrid(node as GridNode, col, 1, idGen, availWidthHwp)) chunks.push(r);
      }
    }
  }

  return concatU8(chunks);
}

function b64Matches(binImg: BinImage, b64: string): boolean {
  const a = TextKit.base64Encode(binImg.data).replace(/\s/g, '');
  const b = b64.replace(/\s/g, '');
  return a === b;
}

/* ═══════════════════════════════════════════════════════════════
   HWP FileHeader stream (256 bytes)
   ═══════════════════════════════════════════════════════════════ */

function buildHwpFileHeader(): Uint8Array {
  const buf = new Uint8Array(256);
  // Signature: "HWP Document File\x00" padded to 32 bytes
  const sig = 'HWP Document File';
  for (let i = 0; i < sig.length; i++) buf[i] = sig.charCodeAt(i);
  const dv = new DataView(buf.buffer);
  dv.setUint32(32, 0x05000300, true); // version 5.0.3.0
  dv.setUint32(36, 0x00000001, true); // flags: bit 0 = compressed (zlib)
  return buf;
}

/* ═══════════════════════════════════════════════════════════════
   OLE2 / CFB container builder
   ═══════════════════════════════════════════════════════════════ */

function buildHwpOle2(
  fileHeaderData: Uint8Array,
  docInfoData:    Uint8Array,
  section0Data:   Uint8Array,
  binImages:      BinImage[] = [],
): Uint8Array {
  const SS = 512;
  const ENDOFCHAIN = 0xFFFFFFFE;
  const FREESECT   = 0xFFFFFFFF;
  const FATSECT    = 0xFFFFFFFD;

  function padSector(d: Uint8Array): Uint8Array {
    const n = Math.ceil(Math.max(d.length, 1) / SS) * SS;
    if (d.length === n) return d;
    const out = new Uint8Array(n);
    out.set(d);
    return out;
  }

  // Pad all data to sector boundaries
  const fhPad  = padSector(fileHeaderData);
  const diPad  = padSector(docInfoData);
  const s0Pad  = padSector(section0Data);
  const imgPads = binImages.map(img => padSector(img.data));

  const fhN  = fhPad.length / SS;
  const diN  = diPad.length / SS;
  const s0N  = s0Pad.length / SS;
  const imgNs = imgPads.map(p => p.length / SS);
  const totalImgN = imgNs.reduce((s, n) => s + n, 0);

  // Directory sectors: 4 entries per sector
  // Entries: Root, FileHeader, DocInfo, BodyText, Section0 [+ BinData + images]
  const numDirEntries = 5 + (binImages.length > 0 ? 1 + binImages.length : 0);
  const dirN = Math.max(2, Math.ceil(numDirEntries / 4));

  // Compute FAT sector count iteratively
  let fatN = 1;
  for (let iter = 0; iter < 10; iter++) {
    const total  = fatN + dirN + fhN + diN + s0N + totalImgN;
    const needed = Math.ceil(total / 128);
    if (needed <= fatN) break;
    fatN = needed;
  }

  const dir1Sec = fatN;
  const fhSec   = fatN + dirN;
  const diSec   = fhSec + fhN;
  const s0Sec   = diSec + diN;

  // Image sectors come after section0
  const imgSecs: number[] = [];
  let curSec = s0Sec + s0N;
  for (const n of imgNs) { imgSecs.push(curSec); curSec += n; }
  const totalSec = curSec;

  // Build FAT
  const fatBuf = new Uint8Array(fatN * SS).fill(0xFF);
  const setFat = (i: number, v: number) => {
    fatBuf[i * 4]     = v & 0xFF;
    fatBuf[i * 4 + 1] = (v >>> 8) & 0xFF;
    fatBuf[i * 4 + 2] = (v >>> 16) & 0xFF;
    fatBuf[i * 4 + 3] = (v >>> 24) & 0xFF;
  };

  for (let i = 0; i < fatN; i++) setFat(i, FATSECT);
  for (let i = 0; i < dirN; i++) setFat(dir1Sec + i, i + 1 < dirN ? dir1Sec + i + 1 : ENDOFCHAIN);
  for (let i = 0; i < fhN;  i++) setFat(fhSec  + i, i + 1 < fhN  ? fhSec  + i + 1 : ENDOFCHAIN);
  for (let i = 0; i < diN;  i++) setFat(diSec  + i, i + 1 < diN  ? diSec  + i + 1 : ENDOFCHAIN);
  for (let i = 0; i < s0N;  i++) setFat(s0Sec  + i, i + 1 < s0N  ? s0Sec  + i + 1 : ENDOFCHAIN);
  for (let ii = 0; ii < imgNs.length; ii++) {
    const start = imgSecs[ii];
    const n = imgNs[ii];
    for (let i = 0; i < n; i++) setFat(start + i, i + 1 < n ? start + i + 1 : ENDOFCHAIN);
  }

  // Build directory
  const dirBuf = new Uint8Array(dirN * SS);
  const dv     = new DataView(dirBuf.buffer);

  function writeDirEntry(
    idx: number, name: string, type: number,
    left: number, right: number, child: number,
    startSec: number, size: number,
  ) {
    const base = idx * 128;
    const nl   = Math.min(name.length, 31);
    for (let i = 0; i < nl; i++) dv.setUint16(base + i * 2, name.charCodeAt(i), true);
    dv.setUint16(base + 64, (nl + 1) * 2, true);
    dirBuf[base + 66] = type;
    dirBuf[base + 67] = 1; // color = black
    dv.setInt32(base + 68, left,  true);
    dv.setInt32(base + 72, right, true);
    dv.setInt32(base + 76, child, true);
    dv.setUint32(base + 116, startSec >>> 0, true);
    dv.setUint32(base + 120, size >>> 0,     true);
  }

  // Prefill all unused dir entry slots with NOSTREAM (-1) for left/right/child
  // so parsers don't follow zero-valued pointers to Root Entry.
  for (let i = 0; i < dirN * 4; i++) {
    const base = i * 128;
    dv.setInt32(base + 68, -1, true); // left sibling = NOSTREAM
    dv.setInt32(base + 72, -1, true); // right sibling = NOSTREAM
    dv.setInt32(base + 76, -1, true); // child = NOSTREAM
  }

  if (binImages.length > 0) {
    // Valid BST for Root Entry children (ascending alphabetical order):
    //   BinData(5) < BodyText(3) < DocInfo(2) < FileHeader(1)
    // Chain: BinData→right=BodyText→right=DocInfo→right=FileHeader
    // Root Entry child = 5 (BinData, alphabetically smallest)
    writeDirEntry(0, 'Root Entry', 5, -1, -1,  5, ENDOFCHAIN, 0);
    writeDirEntry(1, 'FileHeader', 2, -1, -1, -1, fhSec, fileHeaderData.length); // leaf
    writeDirEntry(2, 'DocInfo',    2, -1,  1, -1, diSec, docInfoData.length);    // right=FileHeader(1)
    writeDirEntry(3, 'BodyText',   1, -1,  2,  4, ENDOFCHAIN, 0); // right=DocInfo(2), child=Section0(4)
    writeDirEntry(4, 'Section0',   2, -1, -1, -1, s0Sec, section0Data.length);   // leaf
    writeDirEntry(5, 'BinData',    1, -1,  3,  6, ENDOFCHAIN, 0); // right=BodyText(3), child=first image(6)
    for (let ii = 0; ii < binImages.length; ii++) {
      const img = binImages[ii];
      const streamName = `BIN${String(img.id).padStart(4, '0')}.${img.ext}`;
      // Images named BIN0001, BIN0002, ... are ascending alphabetically — chain via right sibling
      const sibling = ii + 1 < binImages.length ? 7 + ii : -1;
      writeDirEntry(6 + ii, streamName, 2, -1, sibling, -1, imgSecs[ii], img.data.length);
    }
  } else {
    // Valid BST for Root Entry children (ascending alphabetical order):
    //   BodyText(3) < DocInfo(2) < FileHeader(1)
    // Root Entry child = 3 (BodyText, alphabetically smallest)
    writeDirEntry(0, 'Root Entry', 5, -1, -1,  3, ENDOFCHAIN, 0);
    writeDirEntry(1, 'FileHeader', 2, -1, -1, -1, fhSec, fileHeaderData.length); // leaf
    writeDirEntry(2, 'DocInfo',    2, -1,  1, -1, diSec, docInfoData.length);    // right=FileHeader(1)
    writeDirEntry(3, 'BodyText',   1, -1,  2,  4, ENDOFCHAIN, 0); // right=DocInfo(2), child=Section0(4)
    writeDirEntry(4, 'Section0',   2, -1, -1, -1, s0Sec, section0Data.length);   // leaf
  }

  // Set HWP Root Entry CLSID: {C0E3E920-3546-11CF-8D81-00AA00389B71}
  // Required for Hangul Office to recognise this as an HWP document.
  // GUID wire format: Data1(LE u32), Data2(LE u16), Data3(LE u16), Data4(8 bytes BE)
  const HWP_CLSID = [
    0x20, 0xE9, 0xE3, 0xC0, // C0E3E920 little-endian
    0x46, 0x35,             // 3546     little-endian
    0xCF, 0x11,             // 11CF     little-endian
    0x8D, 0x81, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71, // 8D81-00AA00389B71 as-is
  ];
  for (let i = 0; i < 16; i++) dirBuf[80 + i] = HWP_CLSID[i];

  // Build OLE2 file header (512 bytes)
  const hdr  = new Uint8Array(SS);
  const hdv  = new DataView(hdr.buffer);
  const MAGIC = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
  MAGIC.forEach((b, i) => { hdr[i] = b; });
  hdv.setUint16(24, 0x003E, true); // minor version
  hdv.setUint16(26, 0x0003, true); // major version
  hdv.setUint16(28, 0xFFFE, true); // byte order (LE)
  hdv.setUint16(30, 9,      true); // sector size = 2^9 = 512
  hdv.setUint16(32, 6,      true); // mini sector size = 2^6 = 64
  hdv.setUint32(40, 0,          true); // num dir sectors (0 for v3)
  hdv.setUint32(44, fatN,       true); // num FAT sectors
  hdv.setUint32(48, dir1Sec,    true); // first directory sector
  hdv.setUint32(52, 0,          true); // transaction signature (0)
  hdv.setUint32(56, 0x1000,     true); // mini stream cutoff = 4096
  hdv.setUint32(60, ENDOFCHAIN, true); // first mini FAT sector (none)
  hdv.setUint32(64, 0,          true); // num mini FAT sectors (0)
  hdv.setUint32(68, ENDOFCHAIN, true); // first DIFAT extension (none)
  hdv.setUint32(72, 0,          true); // num DIFAT extensions (0)
  for (let i = 0; i < 109; i++) {
    hdv.setUint32(76 + i * 4, i < fatN ? i : FREESECT, true);
  }

  // Assemble output
  const out = new Uint8Array(SS + totalSec * SS);
  out.set(hdr, 0);
  for (let i = 0; i < fatN; i++) {
    out.set(fatBuf.subarray(i * SS, (i + 1) * SS), SS + i * SS);
  }
  for (let i = 0; i < dirN; i++) {
    out.set(dirBuf.subarray(i * SS, (i + 1) * SS), SS + (dir1Sec + i) * SS);
  }
  out.set(fhPad, SS + fhSec * SS);
  out.set(diPad, SS + diSec * SS);
  out.set(s0Pad, SS + s0Sec * SS);
  for (let ii = 0; ii < imgPads.length; ii++) {
    out.set(imgPads[ii], SS + imgSecs[ii] * SS);
  }
  return out;
}

/* ═══════════════════════════════════════════════════════════════
   Utility
   ═══════════════════════════════════════════════════════════════ */

function concatU8(arrays: Uint8Array[]): Uint8Array {
  const total = arrays.reduce((s, a) => s + a.length, 0);
  const out   = new Uint8Array(total);
  let off = 0;
  for (const a of arrays) { out.set(a, off); off += a.length; }
  return out;
}

/* ═══════════════════════════════════════════════════════════════
   Encoder entry point
   ═══════════════════════════════════════════════════════════════ */

export class HwpEncoder implements Encoder {
  readonly format = 'hwp';

  async encode(doc: DocRoot): Promise<Outcome<Uint8Array>> {
    try {
      // First pass: collect unique styles
      const col = new StyleCollector();
      for (const sheet of doc.kids) {
        for (const node of sheet.kids) collectNode(node, col);
      }

      // Collect images from all paragraphs
      const images: BinImage[] = [];
      const seenB64 = new Set<string>();
      let binIdCounter = 1;

      function collectImages(node: any): void {
        if (node.tag === 'para') {
          for (const kid of node.kids) {
            if (kid.tag === 'img') {
              const key = kid.b64.substring(0, 50);
              if (!seenB64.has(key)) {
                seenB64.add(key);
                const raw = TextKit.base64Decode(kid.b64);
                let ext = 'jpg';
                if (kid.mime === 'image/png') ext = 'png';
                else if (kid.mime === 'image/gif') ext = 'gif';
                else if (kid.mime === 'image/bmp') ext = 'bmp';
                images.push({ id: binIdCounter++, ext, data: new Uint8Array(raw) });
              }
            }
          }
        } else if (node.tag === 'grid') {
          for (const row of node.kids) {
            for (const cell of row.kids) {
              for (const para of cell.kids) collectImages(para);
            }
          }
        }
      }
      for (const sheet of doc.kids) {
        for (const node of sheet.kids) collectImages(node);
      }

      // Build streams (raw, uncompressed)
      const docInfoRaw = buildDocInfoStream(col, images);
      const bodyRaw    = buildBodyTextStream(doc, col, images);

      // Compress with zlib (HWP uses zlib deflate for DocInfo and BodyText streams)
      const docInfoCmp = pako.deflate(docInfoRaw);
      const bodyCmp    = pako.deflate(bodyRaw);

      // Assemble OLE2 file (images go as raw streams in BinData storage, NOT compressed)
      const fileHdr = buildHwpFileHeader();
      const hwp     = buildHwpOle2(fileHdr, docInfoCmp, bodyCmp, images);

      // 수정 이유(수정 1): OLE2 구조 검증 — "인코딩 설정" 대화상자 원인 진단용.
      // 한컴오피스가 HWP OLE2를 인식 못하면 텍스트 파일로 간주하여 해당 대화상자가 뜸.
      // magicOk가 false이거나 dirFirstSector가 이상한 값이면 buildHwpOle2 버그 탐색.
      const OLE2_MAGIC = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
      const magicOk = OLE2_MAGIC.every((b, i) => hwp[i] === b);
      const dv2 = new DataView(hwp.buffer, hwp.byteOffset);
      const fatN        = dv2.getUint32(44, true);
      const dirFirstSec = dv2.getUint32(48, true);
      const sectorSize  = 1 << dv2.getUint16(30, true);
      console.log('[HwpEncoder] 검증:', {
        magicOk,
        fileSize:      hwp.length,
        sectorSize,
        fatSectors:    fatN,
        dirFirstSector: dirFirstSec,
        docInfoSize:   docInfoCmp.length,
        bodyTextSize:  bodyCmp.length,
        imageCount:    images.length,
      });
      // Root Entry 디렉토리 확인 (offset = 512 + dirFirstSec * 512)
      const dirOff      = 512 + dirFirstSec * 512;
      const rootNameLen = dv2.getUint16(dirOff + 64, true);
      const rootChild   = dv2.getInt32(dirOff + 76, true);
      console.log('[HwpEncoder] Root Entry:', { nameLen: rootNameLen, childId: rootChild });

      return succeed(hwp);
    } catch (e: any) {
      return fail(`HwpEncoder: ${e instanceof Error ? e.message : String(e)}`);
    }
  }
}

registry.registerEncoder(new HwpEncoder());
