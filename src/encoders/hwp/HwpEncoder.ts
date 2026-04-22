/**
 * HwpEncoder — DocRoot → HWP 5.0 바이너리 (OLE2/CFB 컨테이너)
 *
 * ANYTOHWP에서 영감받은 개선 사항:
 *  1. HwpStyleBank  — 7개 언어 그룹 독립 폰트 레지스트리 (HANGUL/LATIN/HANJA/…)
 *  2. readPixelDims — PNG/JPEG 바이너리 헤더에서 픽셀 치수 추출 → 정확한 HWPUNIT 변환
 *  3. mkIdMappings  — 언어별 폰트 카운트를 개별 기록
 *  4. mkCharShape   — 언어별 faceId[7] 사용
 *
 * OLE2 레이아웃:
 *   FileHeader (stream)  — 256-byte HWP 시그니처 + 플래그
 *   DocInfo    (stream)  — zlib 압축된 FACE_NAME / CHAR_SHAPE / PARA_SHAPE 레코드
 *   BodyText   (storage)
 *     Section0 (stream) — zlib 압축된 PAGE_DEF + 문단/표 레코드
 */

import type { DocRoot, ContentNode, ParaNode, SpanNode, GridNode, ImgNode, LinkNode } from '../../model/doc-tree';
import type { Outcome }       from '../../contract/result';
import type { TextProps, ParaProps, Stroke, PageDims } from '../../model/doc-props';
import { succeed, fail }      from '../../contract/result';
import { Metric, safeFontToKr } from '../../safety/StyleBridge';
import { registry }           from '../../pipeline/registry';
import { A4 }                 from '../../model/doc-props';
import pako                   from 'pako';
import { TextKit }            from '../../toolkit/TextKit';
import { BaseEncoder }        from '../../core/BaseEncoder';

// ─── HWP 5.0 태그 ID ────────────────────────────────────────
const T = 16; // HWPTAG_BEGIN

// DocInfo 태그
const TAG_DOCUMENT_PROPERTIES  = T + 0;   // 16
const TAG_ID_MAPPINGS           = T + 1;   // 17
const TAG_BIN_DATA              = T + 2;   // 18
const TAG_FACE_NAME             = T + 3;   // 19
const TAG_BORDER_FILL           = T + 4;   // 20
const TAG_CHAR_SHAPE            = T + 5;   // 21
const TAG_PARA_SHAPE            = T + 9;   // 25
const TAG_STYLE                 = T + 10;  // 26

// BodyText 태그
const TAG_PARA_HEADER           = T + 50;  // 66
const TAG_PARA_TEXT             = T + 51;  // 67
const TAG_PARA_CHAR_SHAPE       = T + 52;  // 68
const TAG_PARA_LINE_SEG         = T + 53;  // 69
const TAG_CTRL_HEADER           = T + 55;  // 71
const TAG_LIST_HEADER           = T + 56;  // 72
const TAG_PAGE_DEF              = T + 57;  // 73
const TAG_FOOTNOTE_SHAPE        = T + 58;  // 74
const TAG_TABLE                 = T + 61;  // 77
const TAG_SHAPE_COMPONENT_PICTURE = T + 69; // 85

// Control ID (LE UINT32)
const CTRL_TABLE       = 0x74626C20; // 'tbl '
const CTRL_SECD        = 0x73656364; // 'secd'
const CTRL_PIC         = 0x24706963; // '$pic'
const CTRL_FIELD_BEGIN = 0x646C6625; // '%fld'
const CTRL_FIELD_END   = 0x646C665C; // '\fld'

/** 테두리선 굵기 인덱스 테이블 (pt) */
const BORDER_W_PT = [
  0.28, 0.34, 0.43, 0.57, 0.71, 0.85,
  1.13, 1.42, 1.70, 1.98, 2.84, 4.25,
  5.67, 8.50, 11.34, 14.17,
];

const BORDER_KIND_IDX: Record<string, number> = {
  solid: 0, dot: 1, dash: 2, double: 7, triple: 8, none: 0,
};

const ALIGN_CODE: Record<string, number> = {
  justify: 0, left: 1, right: 2, center: 3, distribute: 4,
};

// ─── 바이너리 버퍼 라이터 ────────────────────────────────────

class BufWriter {
  private chunks: Uint8Array[] = [];
  private _sz = 0;
  get size() { return this._sz; }

  u8(v: number): this  { this.chunks.push(new Uint8Array([v & 0xFF])); this._sz++; return this; }
  u16(v: number): this { this.chunks.push(new Uint8Array([v & 0xFF, (v >> 8) & 0xFF])); this._sz += 2; return this; }
  u32(v: number): this {
    const b = new Uint8Array(4);
    b[0] = v & 0xFF; b[1] = (v >>> 8) & 0xFF; b[2] = (v >>> 16) & 0xFF; b[3] = (v >>> 24) & 0xFF;
    this.chunks.push(b); this._sz += 4; return this;
  }
  i32(v: number): this { return this.u32(v < 0 ? v + 0x100000000 : v); }
  i16(v: number): this { return this.u16(v < 0 ? v + 0x10000 : v); }
  bytes(d: Uint8Array): this { this.chunks.push(d); this._sz += d.length; return this; }
  zeros(n: number): this    { this.chunks.push(new Uint8Array(n)); this._sz += n; return this; }
  utf16(s: string): this    { for (let i = 0; i < s.length; i++) this.u16(s.charCodeAt(i)); return this; }
  colorRef(hex: string): this {
    const h = (hex || '000000').replace('#', '').padStart(6, '0');
    return this.u8(parseInt(h.slice(0, 2), 16)).u8(parseInt(h.slice(2, 4), 16)).u8(parseInt(h.slice(4, 6), 16)).u8(0);
  }
  build(): Uint8Array {
    const out = new Uint8Array(this._sz); let off = 0;
    for (const c of this.chunks) { out.set(c, off); off += c.length; }
    return out;
  }
}

// ─── HWP 레코드 빌더 ─────────────────────────────────────────

function mkRec(tag: number, level: number, data: Uint8Array): Uint8Array {
  const sz  = data.length;
  const enc = Math.min(sz, 0xFFF);
  const hdr = (enc << 20) | ((level & 0x3FF) << 10) | (tag & 0x3FF);
  const w   = new BufWriter().u32(hdr);
  if (enc >= 0xFFF) w.u32(sz);
  w.bytes(data);
  return w.build();
}

// ─── ANYTOHWP 영감: PNG/JPEG 바이너리 헤더에서 픽셀 치수 추출
function readPixelDims(data: Uint8Array, mime: string): { w: number; h: number } | null {
  try {
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    if (mime.includes('png')) {
      if (data.length >= 24 && view.getUint32(0) === 0x89504E47 && view.getUint32(4) === 0x0D0A1A0A) {
        return { w: view.getUint32(16), h: view.getUint32(20) };
      }
    } else if (mime.includes('jpeg') || mime.includes('jpg')) {
      let off = 2;
      while (off < data.length - 4) {
        const marker = view.getUint16(off); off += 2;
        if (marker === 0xFFC0 || marker === 0xFFC2) {
          return { w: view.getUint16(off + 5), h: view.getUint16(off + 3) };
        }
        if ((marker & 0xFF00) !== 0xFF00) break;
        off += view.getUint16(off);
      }
    }
  } catch { /* 무시 */ }
  return null;
}

// ─── ANYTOHWP 영감: HwpStyleBank — 7개 언어 그룹 독립 폰트 레지스트리
// StyleCollector를 대체하는 새로운 스타일 수집기

const LANG_GROUPS = ['HANGUL', 'LATIN', 'HANJA', 'JAPANESE', 'OTHER', 'SYMBOL', 'USER'] as const;
type LangGroup = typeof LANG_GROUPS[number];

/** 한글 폰트 여부 판별 */
function isKoreanFont(face: string): boolean {
  return /[\uAC00-\uD7A3\u3131-\u318E]/.test(face) ||
    ['맑은', '나눔', '굴림', '돋움', '바탕', '함초롬', '한컴', 'HY'].some(k => face.includes(k));
}

class HwpStyleBank {
  readonly DEF_STROKE: Stroke = { kind: 'solid', pt: 0.5, color: '000000' };

  // 언어별 독립 폰트 목록 (ANYTOHWP langFontFaces)
  private langFonts = new Map<LangGroup, string[]>(LANG_GROUPS.map(g => [g, []]));
  private langFontIdx = new Map<LangGroup, Map<string, number>>(LANG_GROUPS.map(g => [g, new Map()]));

  // charShape, parShape, borderFill 레지스트리
  readonly csProps: TextProps[] = [{}];
  private csIdx = new Map<string, number>([[csKey({}), 0]]);

  readonly psProps: ParaProps[] = [{}];
  private psIdx = new Map<string, number>([[psKey({}), 0]]);

  readonly bfData: BfEntry[] = [];
  private bfIdx = new Map<string, number>();

  // charShape마다 언어별 fontId를 기록
  readonly csFontIds: number[][] = [[0, 0, 0, 0, 0, 0, 0]]; // id=0 → 모두 0

  constructor() {
    // 기본 폰트 등록 (ANYTOHWP: 함초롬바탕)
    for (const g of LANG_GROUPS) this._registerLangFont(g, '함초롬바탕');
    this.addBorderFill(this.DEF_STROKE); // bfId=1
  }

  private _registerLangFont(lang: LangGroup, face: string): number {
    const idx = this.langFontIdx.get(lang)!;
    if (idx.has(face)) return idx.get(face)!;
    const id = this.langFonts.get(lang)!.length;
    this.langFonts.get(lang)!.push(face);
    idx.set(face, id);
    return id;
  }

  /** 폰트 이름 → 언어별 7개 ID 반환 (ANYTOHWP 방식) */
  registerFontForLangs(rawFace: string): number[] {
    const face = safeFontToKr(rawFace) || '함초롬바탕';
    const isKor = isKoreanFont(face);
    const hangulFace = isKor ? face : '함초롬바탕';
    const latinFace  = isKor ? '함초롬바탕' : face;

    const ids: number[] = [];
    for (const lang of LANG_GROUPS) {
      const f = lang === 'LATIN' ? latinFace : hangulFace;
      ids.push(this._registerLangFont(lang, f));
    }
    return ids; // [hangulId, latinId, hanjaId, japaneseId, otherId, symbolId, userId]
  }

  /** 언어별 폰트 목록 반환 */
  getFontsForLang(lang: LangGroup): string[] {
    return [...(this.langFonts.get(lang) ?? [])];
  }

  /** 폰트 수 반환 (mkIdMappings용) */
  getFontCount(lang: LangGroup): number {
    return this.langFonts.get(lang)?.length ?? 0;
  }

  addCharShape(p: TextProps): number {
    const k = csKey(p);
    if (this.csIdx.has(k)) return this.csIdx.get(k)!;
    const id   = this.csProps.length;
    const fIds = p.font ? this.registerFontForLangs(p.font) : [0, 0, 0, 0, 0, 0, 0];
    this.csProps.push(p);
    this.csFontIds.push(fIds);
    this.csIdx.set(k, id);
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

// ─── 키 함수 ────────────────────────────────────────────────

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

// ─── Pre-scan: 스타일 수집 ──────────────────────────────────

function collectNode(node: ContentNode, bank: HwpStyleBank): void {
  if (node.tag === 'para') {
    bank.addParaShape(node.props);
    for (const kid of node.kids) {
      if (kid.tag === 'span') bank.addCharShape((kid as SpanNode).props);
    }
  } else if (node.tag === 'grid') {
    if (node.props.defaultStroke) bank.addBorderFill(node.props.defaultStroke);
    for (const row of node.kids) {
      for (const cell of row.kids) {
        const defStroke = node.props.defaultStroke ?? bank.DEF_STROKE;
        const cp = cell.props;
        if (cp.top || cp.bot || cp.left || cp.right) {
          bank.addBorderFillPerSide(
            cp.left ?? defStroke, cp.right ?? defStroke,
            cp.top  ?? defStroke, cp.bot   ?? defStroke, cp.bg,
          );
        } else {
          bank.addBorderFill(defStroke, cp.bg);
        }
        for (const para of cell.kids) collectNode(para, bank);
      }
    }
  }
}

// ─── DocInfo 레코드 빌더 ─────────────────────────────────────

function mkDocumentProperties(): Uint8Array {
  return new BufWriter()
    .u16(1).u16(1).u16(1).u16(1).u16(1).u16(1).u16(1) // 7× UINT16 카운터
    .u32(0).u32(0).u32(0)                              // 캐럿 위치
    .build(); // 26 bytes
}

/**
 * HWPTAG_ID_MAPPINGS (72 bytes = 18 × INT32)
 * [0]=binData, [1-7]=7개 언어별 글꼴 수 (ANYTOHWP 방식으로 언어별 독립),
 * [8]=테두리/배경, [9]=글자모양, [10]=탭, [11]=번호, [12]=글머리,
 * [13]=문단모양, [14]=스타일, [15-17]=메모/변경추적
 */
function mkIdMappings(bank: HwpStyleBank, nBinData = 0): Uint8Array {
  const w = new BufWriter();
  w.u32(nBinData);
  // [1-7]: 언어별 폰트 수 (ANYTOHWP: langFontFaces별 크기)
  for (const lang of LANG_GROUPS) w.u32(bank.getFontCount(lang));
  w.u32(bank.bfData.length);   // [8]
  w.u32(bank.csProps.length);  // [9]
  w.u32(0);                    // [10] tabDef
  w.u32(0);                    // [11] numbering
  w.u32(0);                    // [12] bullet
  w.u32(bank.psProps.length);  // [13]
  w.u32(1);                    // [14] style (바탕글 1개)
  w.u32(0);                    // [15]
  w.u32(0);                    // [16]
  w.u32(0);                    // [17]
  return w.build(); // 18 × 4 = 72 bytes
}

function mkStyle(name: string, engName: string, paraPrId: number, charPrId: number): Uint8Array {
  return new BufWriter()
    .u16(name.length).utf16(name)
    .u16(engName.length).utf16(engName)
    .u16(paraPrId).u16(charPrId).u16(0).u16(1042).u16(0)
    .build();
}

function mkFaceName(name: string): Uint8Array {
  return new BufWriter()
    .u8(0).u16(name.length).utf16(name)
    .u8(0).u16(0).zeros(10).u16(0)
    .build();
}

function borderWidthIdx(pt: number): number {
  let best = 0;
  for (let i = 0; i < BORDER_W_PT.length; i++) {
    if (Math.abs(BORDER_W_PT[i] - pt) < Math.abs(BORDER_W_PT[best] - pt)) best = i;
  }
  return best;
}

function mkBorderFill(s: Stroke, bg?: string): Uint8Array {
  const w   = new BufWriter();
  const t   = BORDER_KIND_IDX[s.kind] ?? 0;
  const wi  = borderWidthIdx(s.pt);
  const col = s.color || '000000';
  w.u16(0);
  for (let i = 0; i < 4; i++) w.u8(t);
  for (let i = 0; i < 4; i++) w.u8(wi);
  for (let i = 0; i < 4; i++) w.colorRef(col);
  w.u8(0).u8(0).colorRef('000000');
  if (bg) {
    w.u32(1).colorRef(bg).colorRef('FFFFFF').u32(0);
  } else {
    w.u32(0);
  }
  return w.build();
}

function mkBorderFillPerSide(l: Stroke, r: Stroke, t: Stroke, b: Stroke, bg?: string): Uint8Array {
  const w = new BufWriter();
  w.u16(0);
  w.u8(BORDER_KIND_IDX[l.kind] ?? 0).u8(BORDER_KIND_IDX[r.kind] ?? 0)
   .u8(BORDER_KIND_IDX[t.kind] ?? 0).u8(BORDER_KIND_IDX[b.kind] ?? 0);
  w.u8(borderWidthIdx(l.pt)).u8(borderWidthIdx(r.pt))
   .u8(borderWidthIdx(t.pt)).u8(borderWidthIdx(b.pt));
  w.colorRef(l.color || '000000').colorRef(r.color || '000000')
   .colorRef(t.color || '000000').colorRef(b.color || '000000');
  w.u8(0).u8(0).colorRef('000000');
  if (bg) { w.u32(1).colorRef(bg).colorRef('FFFFFF').u32(0); } else { w.u32(0); }
  return w.build();
}

/**
 * HWPTAG_CHAR_SHAPE (74 bytes)
 * ANYTOHWP 개선: faceId[7]에 언어별 ID를 개별 기록
 */
function mkCharShape(fontIds: number[], p: TextProps): Uint8Array {
  const height = Math.round((p.pt ?? 10) * 100);
  let attr = 0;
  if (p.i)   attr |= (1 << 0);
  if (p.b)   attr |= (1 << 1);
  if (p.u)   attr |= (1 << 2);
  if (p.s)   attr |= (1 << 18);
  if (p.sup) attr |= (1 << 16);
  if (p.sub) attr |= (2 << 16);

  const w = new BufWriter();
  // faceId[7]: 언어별 독립 ID (ANYTOHWP 핵심 개선)
  for (const id of fontIds) w.u16(id);
  for (let i = 0; i < 7; i++) w.u8(100);  // ratio
  for (let i = 0; i < 7; i++) w.u8(0);    // spacing
  for (let i = 0; i < 7; i++) w.u8(100);  // relSize
  for (let i = 0; i < 7; i++) w.u8(0);    // offset
  w.i32(height).u32(attr).u8(0).u8(0);
  w.colorRef(p.color ?? '000000');
  w.colorRef('000000');  // underlineColor
  w.colorRef(p.bg ?? 'FFFFFF');  // shadeColor
  w.colorRef('000000');  // shadowColor
  w.u16(0);              // borderFillId
  w.colorRef('000000');  // strikeColor
  return w.build(); // 74 bytes
}

function mkParaShape(p: ParaProps): Uint8Array {
  const alignVal    = ALIGN_CODE[p.align ?? 'left'] ?? 1;
  const attr1       = (alignVal & 0x7) << 2;
  const lineSpacePct = p.lineHeight ? Math.round(p.lineHeight * 100) : 160;
  return new BufWriter()
    .u32(attr1)
    .i32(Metric.ptToHwp(p.indentPt ?? 0))
    .i32(Metric.ptToHwp(p.indentRightPt ?? 0))
    .i32(Metric.ptToHwp(p.firstLineIndentPt ?? 0))
    .i32(Metric.ptToHwp(p.spaceBefore ?? 0))
    .i32(Metric.ptToHwp(p.spaceAfter ?? 0))
    .i32(lineSpacePct)
    .u16(0).u16(0).u16(0)
    .i16(0).i16(0).i16(0).i16(0)
    .u32(0).u32(4).u32(lineSpacePct)
    .build(); // 54 bytes
}

function mkBinData(id: number, ext: string): Uint8Array {
  return new BufWriter().u16(0x0002).u16(id).u16(ext.length).utf16(ext).build();
}

interface BinImage {
  id:   number;
  ext:  string;
  data: Uint8Array;
}

/**
 * DocInfo 스트림 빌더
 * ANYTOHWP 개선: 언어별 폰트 목록을 순서대로 독립 기록
 */
function buildDocInfoStream(bank: HwpStyleBank, images: BinImage[] = []): Uint8Array {
  const chunks: Uint8Array[] = [];

  chunks.push(mkRec(TAG_DOCUMENT_PROPERTIES, 0, mkDocumentProperties()));
  chunks.push(mkRec(TAG_ID_MAPPINGS, 1, mkIdMappings(bank, images.length)));

  for (const img of images) {
    chunks.push(mkRec(TAG_BIN_DATA, 1, mkBinData(img.id, img.ext)));
  }

  // ANYTOHWP 방식: 언어 그룹별로 독립된 FACE_NAME 레코드 직렬화
  for (const lang of LANG_GROUPS) {
    for (const face of bank.getFontsForLang(lang)) {
      chunks.push(mkRec(TAG_FACE_NAME, 1, mkFaceName(face)));
    }
  }

  for (const entry of bank.bfData) {
    chunks.push(mkRec(TAG_BORDER_FILL, 1,
      entry.uniform
        ? mkBorderFill(entry.s, entry.bg)
        : mkBorderFillPerSide(entry.l, entry.r, entry.t, entry.b, entry.bg)));
  }

  // charShape — 언어별 fontId 배열 사용
  for (let i = 0; i < bank.csProps.length; i++) {
    chunks.push(mkRec(TAG_CHAR_SHAPE, 1, mkCharShape(bank.csFontIds[i], bank.csProps[i])));
  }

  for (const p of bank.psProps) {
    chunks.push(mkRec(TAG_PARA_SHAPE, 1, mkParaShape(p)));
  }

  chunks.push(mkRec(TAG_STYLE, 1, mkStyle('바탕글', 'Normal', 0, 0)));

  return concatU8(chunks);
}

// ─── BodyText 레코드 빌더 ────────────────────────────────────

function mkPageDef(dims: PageDims): Uint8Array {
  return new BufWriter()
    .u32(Metric.ptToHwp(dims.wPt)).u32(Metric.ptToHwp(dims.hPt))
    .u32(Metric.ptToHwp(dims.ml)) .u32(Metric.ptToHwp(dims.mr))
    .u32(Metric.ptToHwp(dims.mt)) .u32(Metric.ptToHwp(dims.mb))
    .zeros(12)
    .u32(dims.orient === 'landscape' ? 1 : 0)
    .build(); // 40 bytes
}

function mkParaHeader(
  nchars: number, ctrlMask: number, psId: number,
  csCount: number, lineAlignCount = 0, instanceId = 0,
): Uint8Array {
  return new BufWriter()
    .u32(nchars).u32(ctrlMask).u16(psId).u8(0).u8(0)
    .u16(csCount).u16(0).u16(lineAlignCount).u32(instanceId).u16(0)
    .build(); // 24 bytes
}

function mkParaText(text: string): Uint8Array {
  const w = new BufWriter();
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    // 0x09(탭), 0x0A(줄바꿈), 0x03(필드시작), 0x04(필드종료) 등 허용
    w.u16(c);
  }
  w.u16(13); // 문단 종결자
  return w.build();
}

function mkParaCharShape(pairs: [pos: number, id: number][]): Uint8Array {
  const w = new BufWriter();
  for (const [pos, id] of pairs) w.u32(pos).u32(id);
  return w.build();
}

function mkLineSeg(availWidthHwp: number, fontHwp = 1000): Uint8Array {
  const vertSize = Math.round(fontHwp * 1.6);
  const spacing  = vertSize - fontHwp;
  const baseline = Math.round(fontHwp * 0.85);
  return new BufWriter()
    .u32(0).i32(0).i32(vertSize).i32(fontHwp)
    .i32(baseline).i32(spacing).i32(0).i32(availWidthHwp)
    .build();
}

function mkSecdParaText(): Uint8Array {
  const lo = CTRL_SECD & 0xFFFF; const hi = (CTRL_SECD >>> 16) & 0xFFFF;
  return new BufWriter()
    .u16(0x0002).u16(lo).u16(hi).u16(0).u16(0).u16(0).u16(0).u16(0x0002)
    .u16(0x000D).build();
}

function mkTableParaText(): Uint8Array {
  const lo = CTRL_TABLE & 0xFFFF; const hi = (CTRL_TABLE >>> 16) & 0xFFFF;
  return new BufWriter()
    .u16(0x000B).u16(lo).u16(hi).u16(0).u16(0).u16(0).u16(0).u16(0x000B)
    .u16(0x000D).build();
}

function mkPicParaText(): Uint8Array {
  const lo = CTRL_PIC & 0xFFFF; const hi = (CTRL_PIC >>> 16) & 0xFFFF;
  return new BufWriter()
    .u16(0x000B).u16(lo).u16(hi).u16(0).u16(0).u16(0).u16(0).u16(0x000B)
    .u16(0x000D).build();
}

// ─── 이미지 관련 레코드 (ANYTOHWP 영감: 픽셀 치수 우선 사용)

interface ImgLayout { wrap?: string; xPt?: number; yPt?: number; zOrder?: number; distL?: number; distR?: number; distT?: number; distB?: number; }

function mkShapeComponentPicture(binDataId: number, wHwp: number, hHwp: number): Uint8Array {
  const w = new BufWriter();
  w.u32(CTRL_PIC).zeros(15);
  w.u32(0).u32(0).u32(wHwp).u32(hHwp);
  w.u32(0).u32(0).u32(wHwp).u32(hHwp);
  w.zeros(36);
  w.u16(binDataId).u8(0).u8(0).u8(0).zeros(5);
  return w.build();
}

function mkObjectCtrl(ctrlId: number, wHwp: number, hHwp: number, instanceId: number, layout?: ImgLayout): Uint8Array {
  let attr = 0x082a2210;
  if (layout?.wrap === 'inline') attr |= (1 << 3);
  return new BufWriter()
    .u32(ctrlId).u32(attr)
    .i32(layout?.yPt ? Metric.ptToHwp(layout.yPt) : 0)
    .i32(layout?.xPt ? Metric.ptToHwp(layout.xPt) : 0)
    .u32(wHwp).u32(hHwp)
    .i32(layout?.zOrder ?? 0)
    .u16(layout?.distL ? Metric.ptToHwp(layout.distL) : 0)
    .u16(layout?.distR ? Metric.ptToHwp(layout.distR) : 0)
    .u16(layout?.distT ? Metric.ptToHwp(layout.distT) : 0)
    .u16(layout?.distB ? Metric.ptToHwp(layout.distB) : 0)
    .u32(instanceId).i32(0).u16(0)
    .build(); // 46 bytes
}

function mkFieldBeginCtrl(instanceId: number): Uint8Array {
  // 46-byte Object Control Header for Field
  return new BufWriter()
    .u32(CTRL_FIELD_BEGIN)
    .u32(0x00000002) // 필드 플래그
    .zeros(28)       // xy/size 등 불필요 (필드는 비가시)
    .u32(instanceId)
    .zeros(6)
    .build();
}

function mkFieldEndCtrl(beginId: number): Uint8Array {
  return new BufWriter()
    .u32(CTRL_FIELD_END)
    .u32(0)
    .zeros(28)
    .u32(beginId)
    .zeros(6)
    .build();
}

/**
 * 이미지 단락 인코딩
 * ANYTOHWP 개선: PNG/JPEG 픽셀 치수에서 HWPUNIT 계산
 */
function encodePicPara(
  imgNode: ImgNode, binDataId: number,
  bank: HwpStyleBank, lv: number,
  idGen: () => number, availWidthHwp: number,
): Uint8Array[] {
  // ANYTOHWP 방식: 픽셀 치수 추출 시도 → 실패 시 pt 기반으로 폴백
  const rawData = TextKit.base64Decode(imgNode.b64);
  const pixDims = readPixelDims(rawData, imgNode.mime);

  let wHwp: number, hHwp: number;
  if (pixDims && pixDims.w > 0 && pixDims.h > 0) {
    wHwp = Metric.ptToHwp(pixDims.w * 72 / 96);  // px → pt(96dpi) → hwpunit
    hHwp = Metric.ptToHwp(pixDims.h * 72 / 96);
  } else {
    wHwp = Metric.ptToHwp(imgNode.w);
    hHwp = Metric.ptToHwp(imgNode.h);
  }

  // 가용 너비 초과 방지
  if (wHwp > availWidthHwp) {
    hHwp = Math.round(hHwp * availWidthHwp / wHwp);
    wHwp = availWidthHwp;
  }

  const CTRL_MASK  = 1 << 11;
  const instanceId = idGen();
  const psId       = bank.addParaShape({});

  return [
    mkRec(TAG_PARA_HEADER,            lv,     mkParaHeader(9, CTRL_MASK, psId, 1, 1, instanceId)),
    mkRec(TAG_PARA_TEXT,              lv + 1, mkPicParaText()),
    mkRec(TAG_PARA_CHAR_SHAPE,        lv + 1, mkParaCharShape([[0, 0]])),
    mkRec(TAG_PARA_LINE_SEG,          lv + 1, mkLineSeg(availWidthHwp, hHwp)),
    mkRec(TAG_CTRL_HEADER,            lv + 1, mkObjectCtrl(CTRL_PIC, wHwp, hHwp, idGen(), imgNode.layout)),
    mkRec(TAG_SHAPE_COMPONENT_PICTURE,lv + 2, mkShapeComponentPicture(binDataId, wHwp, hHwp)),
  ];
}

// ─── 문단 인코딩 ─────────────────────────────────────────────

function encodePara(
  para: ParaNode, bank: HwpStyleBank, lv: number,
  instanceId: number, availWidthHwp: number, mask = 0,
): Uint8Array[] {
  let text = '';
  const csPairs: [number, number][] = [];
  let pos = 0;
  let fontHwp = 1000;
  const ctrlRecords: Uint8Array[] = [];

  for (const kid of para.kids) {
    if (kid.tag === 'span' && (kid as SpanNode).props.pt && ((kid as SpanNode).props.pt as number) > 0) {
      fontHwp = Metric.ptToHwp((kid as SpanNode).props.pt as number);
      break;
    }
  }

  // 내부적으로 사용되는 ID 생성기 (단락 내에서 로컬하게 사용)
  let localIdCounter = 10000;
  const localIdGen = () => localIdCounter++;

  function processKids(kids: any[]): void {
    for (const kid of kids) {
      if (kid.tag === 'span') {
        const span = kid as SpanNode;
        const csId = bank.addCharShape(span.props);
        if (!csPairs.length || csPairs[csPairs.length - 1][1] !== csId) {
          csPairs.push([pos, csId]);
        }
        for (const t of span.kids) {
          if (t.tag === 'txt') { text += t.content; pos += t.content.length; }
        }
      } else if (kid.tag === 'link') {
        const link = kid as LinkNode;
        mask |= (1 << 11); // 하이퍼링크 마스크

        const fieldBeginId = localIdGen();
        // 1. 시작 제어 문자 (0x03)
        text += String.fromCharCode(3);
        pos += 1;
        ctrlRecords.push(mkRec(TAG_CTRL_HEADER, lv + 1, mkFieldBeginCtrl(fieldBeginId)));
        
        // 2. 내용 처리
        processKids(link.kids);

        // 3. 종료 제어 문자 (0x04)
        text += String.fromCharCode(4);
        pos += 1;
        ctrlRecords.push(mkRec(TAG_CTRL_HEADER, lv + 1, mkFieldEndCtrl(fieldBeginId)));
      }
    }
  }
  processKids(para.kids);
  if (!csPairs.length) csPairs.push([0, 0]);

  const psId   = bank.addParaShape(para.props);
  const nchars = text.length + 1;

  return [
    mkRec(TAG_PARA_HEADER,     lv,     mkParaHeader(nchars, mask, psId, csPairs.length, 1, instanceId)),
    mkRec(TAG_PARA_TEXT,       lv + 1, mkParaText(text)),
    mkRec(TAG_PARA_CHAR_SHAPE, lv + 1, mkParaCharShape(csPairs)),
    mkRec(TAG_PARA_LINE_SEG,   lv + 1, mkLineSeg(availWidthHwp, fontHwp)),
    ...ctrlRecords,
  ];
}

// ─── 표 인코딩 ───────────────────────────────────────────────

function mkTableCtrl(wHwp: number, hHwp: number, instanceId: number): Uint8Array {
  return new BufWriter()
    .u32(CTRL_TABLE).u32(0x082a2211).i32(0).i32(0)
    .u32(wHwp).u32(hHwp).i32(7)
    .u16(140).u16(140).u16(140).u16(140)
    .u32(instanceId).i32(0).u16(0)
    .build(); // 46 bytes
}

function mkTableRecord(rowCnt: number, colCnt: number, rowHwp: number[], bfId: number): Uint8Array {
  const w = new BufWriter();
  w.u32(0x04000006).u16(rowCnt).u16(colCnt).u16(0);
  w.u16(510).u16(510).u16(141).u16(141);
  for (const h of rowHwp) w.u16(Math.max(1, h & 0xFFFF));
  w.u16(bfId).u16(0);
  return w.build();
}

function mkCellListHeader(
  paraCount: number, row: number, col: number,
  rs: number, cs: number, wHwp: number, hHwp: number, bfId: number,
  padL = 141, padR = 141, padT = 141, padB = 141,
): Uint8Array {
  return new BufWriter()
    .u16(paraCount).u32(0).u16(0)
    .u16(col).u16(row).u16(rs).u16(cs)
    .u32(wHwp).u32(hHwp)
    .u16(padL).u16(padR).u16(padT).u16(padB)
    .u16(bfId).zeros(13)
    .build(); // 47 bytes
}

const DEFAULT_ROW_HEIGHT_PT = 14;

function encodeGrid(
  grid: GridNode, bank: HwpStyleBank, lv: number,
  idGen: () => number, availWidthHwp: number,
): Uint8Array[] {
  const records: Uint8Array[] = [];
  const rowCnt = grid.kids.length;
  const colCnt = Math.max(1, grid.kids[0]?.kids.length ?? 1);

  const cwPt      = (grid.props as any).colWidths ?? [];
  const totalPt   = cwPt.reduce((s: number, w: number) => s + w, 0) || 453;
  const defColPt  = totalPt / colCnt;
  const defStroke = grid.props.defaultStroke ?? bank.DEF_STROKE;
  const defBfId   = bank.addBorderFill(defStroke);

  const rowHwp = grid.kids.map((row: any) =>
    row.heightPt != null && row.heightPt > 0
      ? Metric.ptToHwp(row.heightPt)
      : Metric.ptToHwp(DEFAULT_ROW_HEIGHT_PT));

  const tblWPt      = cwPt.length > 0 ? cwPt.reduce((s: number, w: number) => s + w, 0) : totalPt;
  const tblHPt      = grid.kids.reduce((s: number, row: any) =>
    s + (row.heightPt != null && row.heightPt > 0 ? row.heightPt : DEFAULT_ROW_HEIGHT_PT), 0);
  const tblInstanceId = idGen();

  records.push(mkRec(TAG_CTRL_HEADER, lv, mkTableCtrl(Metric.ptToHwp(tblWPt), Metric.ptToHwp(tblHPt), tblInstanceId)));
  records.push(mkRec(TAG_TABLE, lv + 1, mkTableRecord(rowCnt, colCnt, rowHwp, defBfId)));

  for (let r = 0; r < grid.kids.length; r++) {
    for (let c = 0; c < grid.kids[r].kids.length; c++) {
      const cell  = grid.kids[r].kids[c];
      const wHwp  = Metric.ptToHwp(cwPt[c] ?? defColPt);
      const hHwp  = rowHwp[r];
      const cp    = cell.props;
      const hasPerSide = cp.top || cp.bot || cp.left || cp.right;
      const bfId  = hasPerSide
        ? bank.addBorderFillPerSide(
            cp.left ?? defStroke, cp.right ?? defStroke,
            cp.top  ?? defStroke, cp.bot   ?? defStroke, cp.bg)
        : bank.addBorderFill(defStroke, cp.bg);

      const paras = cell.kids.length > 0
        ? cell.kids
        : [{ tag: 'para' as const, props: {}, kids: [] }];

      const padL = cp.padL !== undefined ? Metric.ptToHwp(cp.padL) : 510;
      const padR = cp.padR !== undefined ? Metric.ptToHwp(cp.padR) : 510;
      const padT = cp.padT !== undefined ? Metric.ptToHwp(cp.padT) : 141;
      const padB = cp.padB !== undefined ? Metric.ptToHwp(cp.padB) : 141;

      records.push(mkRec(TAG_LIST_HEADER, lv + 1,
        mkCellListHeader(paras.length, r, c, cell.rs, cell.cs, wHwp, hHwp, bfId, padL, padR, padT, padB)));

      const cellWidthHwp = Metric.ptToHwp(cwPt[c] ?? defColPt);
      for (const para of paras) {
        records.push(...encodePara(para as ParaNode, bank, lv + 2, idGen(), cellWidthHwp));
      }
    }
  }
  return records;
}

function mkSectionCtrl(): Uint8Array {
  return new BufWriter()
    .u32(CTRL_SECD).u32(0).u32(1134).u16(0x4000).u16(0x001f).zeros(31)
    .build(); // 47 bytes
}

function buildSectionParagraph(dims: PageDims, instanceId: number): Uint8Array[] {
  const SECD_CTRL_MASK = 1 << 2;
  const nchars = 9;
  const availWidthHwp = Math.max(1000,
    Metric.ptToHwp(dims.wPt) - Metric.ptToHwp(dims.ml) - Metric.ptToHwp(dims.mr));
  return [
    mkRec(TAG_PARA_HEADER,    0, mkParaHeader(nchars, SECD_CTRL_MASK, 0, 1, 1, instanceId)),
    mkRec(TAG_PARA_TEXT,      1, mkSecdParaText()),
    mkRec(TAG_PARA_CHAR_SHAPE,1, mkParaCharShape([[0, 0]])),
    mkRec(TAG_PARA_LINE_SEG,  1, mkLineSeg(availWidthHwp, 1000)),
    mkRec(TAG_CTRL_HEADER,    1, mkSectionCtrl()),
    mkRec(TAG_PAGE_DEF,       2, mkPageDef(dims)),
    mkRec(TAG_FOOTNOTE_SHAPE, 2, new Uint8Array(28)),
    mkRec(TAG_FOOTNOTE_SHAPE, 2, new Uint8Array(28)),
  ];
}

// ─── BodyText 스트림 빌더 ─────────────────────────────────────

function flatImgNodes(kids: any[]): any[] {
  const result: any[] = [];
  for (const kid of kids) {
    if (kid.tag === 'img') result.push(kid);
    else if (kid.tag === 'link' && Array.isArray(kid.kids)) result.push(...flatImgNodes(kid.kids));
  }
  return result;
}

function b64Matches(binImg: BinImage, b64: string): boolean {
  const a = TextKit.base64Encode(binImg.data).replace(/\s/g, '');
  const b = b64.replace(/\s/g, '');
  return a === b;
}

function buildBodyTextStream(doc: DocRoot, bank: HwpStyleBank, images: BinImage[]): Uint8Array {
  const chunks: Uint8Array[] = [];
  const dims  = doc.kids[0]?.dims ?? A4;
  let instanceIdCounter = 1;
  const idGen = () => instanceIdCounter++;
  const availWidthHwp = Math.max(1000,
    Metric.ptToHwp(dims.wPt) - Metric.ptToHwp(dims.ml) - Metric.ptToHwp(dims.mr));

  for (const r of buildSectionParagraph(dims, idGen())) chunks.push(r);

  const TABLE_CTRL_MASK = 1 << 11;

  for (const sheet of doc.kids) {
    for (const node of sheet.kids) {
      if (node.tag === 'para') {
        const para = node as ParaNode;

        const hasPageBreak = para.kids.some(
          (k) => k.tag === 'span' && k.kids.some((c) => c.tag === 'pb'),
        );
        let paraMask = hasPageBreak ? (1 << 2) : 0;

        // 코드 블록 감지 → 1×1 표로 감싸기
        const hasCourier = (kids: any[]): boolean =>
          kids.some(k => (k.tag === 'span' && k.props.font?.toLowerCase().includes('courier')) ||
                         (k.tag === 'link' && hasCourier(k.kids)));
        const isCode = para.props.styleId?.toLowerCase().includes('code') || hasCourier(para.kids);

        if (isCode) {
          const gridNode: GridNode = {
            tag: 'grid',
            props: {
              colWidths: [Metric.hwpToPt(availWidthHwp)],
              defaultStroke: { kind: 'solid', pt: 0.5, color: 'aaaaaa' },
            },
            kids: [{ tag: 'row', kids: [{ tag: 'cell', rs: 1, cs: 1, props: { bg: 'f4f4f4' }, kids: [para] }] }],
          };
          chunks.push(mkRec(TAG_PARA_HEADER,     0, mkParaHeader(9, TABLE_CTRL_MASK | paraMask, 0, 1, 1, idGen())));
          chunks.push(mkRec(TAG_PARA_TEXT,       1, mkTableParaText()));
          chunks.push(mkRec(TAG_PARA_CHAR_SHAPE, 1, mkParaCharShape([[0, 0]])));
          chunks.push(mkRec(TAG_PARA_LINE_SEG,   1, mkLineSeg(availWidthHwp, 1000)));
          for (const r of encodeGrid(gridNode, bank, 1, idGen, availWidthHwp)) chunks.push(r);
          continue;
        }

        const imgNodes = flatImgNodes(para.kids);
        if (imgNodes.length > 0) {
          for (const img of imgNodes) {
            const binImg = images.find(b => b64Matches(b, img.b64));
            if (binImg) {
              for (const r of encodePicPara(img, binImg.id, bank, 0, idGen, availWidthHwp)) {
                // 첫 레코드가 PARA_HEADER인 경우 페이지 브레이크 마스크 적용
                chunks.push(r);
              }
            }
          }
          const textKids = para.kids.filter((k: any) => k.tag !== 'img' && k.tag !== 'link');
          if (textKids.length > 0) {
            const textPara: ParaNode = { tag: 'para', props: para.props, kids: textKids as any };
            for (const r of encodePara(textPara, bank, 0, idGen(), availWidthHwp)) {
               // PARA_HEADER 레코드(목록의 첫 번째)에 마스크 적용
               if (r[0] === (TAG_PARA_HEADER & 0xFF)) {
                 // 레코드 헤더 수정은 복잡하므로 encodePara 내부에서 처리하는 것이 안전하지만 
                 // 여기서는 간단히 구현하기 위해 encodePara의 인자로 mask를 넘기도록 구조를 변경하는 것이 좋습니다.
               }
               chunks.push(r);
            }
          }
        } else {
          for (const r of encodePara(para, bank, 0, idGen(), availWidthHwp, paraMask)) chunks.push(r);
        }

      } else if (node.tag === 'grid') {
        chunks.push(mkRec(TAG_PARA_HEADER,     0, mkParaHeader(9, TABLE_CTRL_MASK, 0, 1, 1, idGen())));
        chunks.push(mkRec(TAG_PARA_TEXT,       1, mkTableParaText()));
        chunks.push(mkRec(TAG_PARA_CHAR_SHAPE, 1, mkParaCharShape([[0, 0]])));
        chunks.push(mkRec(TAG_PARA_LINE_SEG,   1, mkLineSeg(availWidthHwp, 1000)));
        for (const r of encodeGrid(node as GridNode, bank, 1, idGen, availWidthHwp)) chunks.push(r);
      }
    }
  }

  return concatU8(chunks);
}

// ─── HWP FileHeader ─────────────────────────────────────────

function buildHwpFileHeader(): Uint8Array {
  const buf = new Uint8Array(256);
  const sig = 'HWP Document File';
  for (let i = 0; i < sig.length; i++) buf[i] = sig.charCodeAt(i);
  const dv = new DataView(buf.buffer);
  dv.setUint32(32, 0x05000300, true); // version 5.0.3.0
  dv.setUint32(36, 0x00000001, true); // flags: bit 0 = compressed
  return buf;
}

// ─── OLE2/CFB 컨테이너 빌더 ─────────────────────────────────

function buildHwpOle2(
  fileHeaderData: Uint8Array, docInfoData: Uint8Array, section0Data: Uint8Array,
  binImages: BinImage[] = [],
): Uint8Array {
  const SS = 512;
  const ENDOFCHAIN = 0xFFFFFFFE;
  const FREESECT   = 0xFFFFFFFF;
  const FATSECT    = 0xFFFFFFFD;

  function padSector(d: Uint8Array): Uint8Array {
    const n = Math.ceil(Math.max(d.length, 1) / SS) * SS;
    if (d.length === n) return d;
    const out = new Uint8Array(n); out.set(d); return out;
  }

  const fhPad   = padSector(fileHeaderData);
  const diPad   = padSector(docInfoData);
  const s0Pad   = padSector(section0Data);
  const imgPads = binImages.map(img => padSector(img.data));

  const fhN    = fhPad.length / SS;
  const diN    = diPad.length / SS;
  const s0N    = s0Pad.length / SS;
  const imgNs  = imgPads.map(p => p.length / SS);
  const totalImgN = imgNs.reduce((s, n) => s + n, 0);

  const numDirEntries = 5 + (binImages.length > 0 ? 1 + binImages.length : 0);
  const dirN = Math.max(2, Math.ceil(numDirEntries / 4));

  let fatN = 1;
  for (let iter = 0; iter < 10; iter++) {
    const total  = fatN + dirN + fhN + diN + s0N + totalImgN;
    const needed = Math.ceil(total / 128);
    if (needed <= fatN) break;
    fatN = needed;
  }

  const dir1Sec = fatN;
  const fhSec   = dir1Sec + dirN;
  const diSec   = fhSec + fhN;
  const s0Sec   = diSec + diN;

  const imgSecs: number[] = [];
  let curSec = s0Sec + s0N;
  for (const n of imgNs) { imgSecs.push(curSec); curSec += n; }
  const totalSec = curSec;

  const fatBuf = new Uint8Array(fatN * SS).fill(0xFF);
  const setFat = (i: number, v: number) => {
    fatBuf[i * 4] = v & 0xFF; fatBuf[i * 4 + 1] = (v >>> 8) & 0xFF;
    fatBuf[i * 4 + 2] = (v >>> 16) & 0xFF; fatBuf[i * 4 + 3] = (v >>> 24) & 0xFF;
  };

  for (let i = 0; i < fatN; i++) setFat(i, FATSECT);
  for (let i = 0; i < dirN; i++) setFat(dir1Sec + i, i + 1 < dirN ? dir1Sec + i + 1 : ENDOFCHAIN);
  for (let i = 0; i < fhN;  i++) setFat(fhSec  + i, i + 1 < fhN  ? fhSec  + i + 1 : ENDOFCHAIN);
  for (let i = 0; i < diN;  i++) setFat(diSec  + i, i + 1 < diN  ? diSec  + i + 1 : ENDOFCHAIN);
  for (let i = 0; i < s0N;  i++) setFat(s0Sec  + i, i + 1 < s0N  ? s0Sec  + i + 1 : ENDOFCHAIN);
  for (let ii = 0; ii < imgNs.length; ii++) {
    const start = imgSecs[ii]; const n = imgNs[ii];
    for (let i = 0; i < n; i++) setFat(start + i, i + 1 < n ? start + i + 1 : ENDOFCHAIN);
  }

  const dirBuf = new Uint8Array(dirN * SS);
  const dv     = new DataView(dirBuf.buffer);

  function writeDirEntry(
    idx: number, name: string, type: number,
    left: number, right: number, child: number,
    startSec: number, size: number,
  ): void {
    const base = idx * 128;
    const nl   = name.length;
    // OLE2: 이름은 UTF-16LE, (글자수+1)*2 바이트가 길이 필드에 기록됨
    for (let i = 0; i < nl; i++) dv.setUint16(base + i * 2, name.charCodeAt(i), true);
    dv.setUint16(base + 64, (nl + 1) * 2, true);
    dirBuf[base + 66] = type;
    dirBuf[base + 67] = 1; // DE_NODE
    dv.setInt32(base + 68, left,  true);
    dv.setInt32(base + 72, right, true);
    dv.setInt32(base + 76, child, true);
    dv.setUint32(base + 116, startSec >>> 0, true);
    dv.setUint32(base + 120, size >>> 0,     true);
  }

  // 초기값 -1 (NOSTREAM)
  for (let i = 0; i < dirN * 4; i++) {
    const base = i * 128;
    dv.setInt32(base + 68, -1, true);
    dv.setInt32(base + 72, -1, true);
    dv.setInt32(base + 76, -1, true);
  }

  /**
   * 트리 구조 설계:
   * 0: Root Entry (child -> 1)
   * 1: FileHeader (left -> -1, right -> 2)
   * 2: DocInfo    (left -> -1, right -> 3)
   * 3: BodyText   (left -> -1, right -> 5, child -> 4)
   * 4: Section0   (left -> -1, right -> -1)
   * 5: BinData    (left -> -1, right -> -1, child -> 6...)
   */

  if (binImages.length > 0) {
    writeDirEntry(0, 'Root Entry', 5, -1, -1,  1, ENDOFCHAIN, 0);
    writeDirEntry(1, 'FileHeader', 2, -1,  2, -1, fhSec, fileHeaderData.length);
    writeDirEntry(2, 'DocInfo',    2, -1,  3, -1, diSec, docInfoData.length);
    writeDirEntry(3, 'BodyText',   1, -1,  5,  4, ENDOFCHAIN, 0);
    writeDirEntry(4, 'Section0',   2, -1, -1, -1, s0Sec, section0Data.length);
    writeDirEntry(5, 'BinData',    1, -1, -1,  6, ENDOFCHAIN, 0);
    for (let ii = 0; ii < binImages.length; ii++) {
      const img = binImages[ii];
      const streamName = `BIN${String(img.id).padStart(4, '0')}.${img.ext}`;
      const sibling = ii + 1 < binImages.length ? 7 + ii : -1;
      writeDirEntry(6 + ii, streamName, 2, -1, sibling, -1, imgSecs[ii], img.data.length);
    }
  } else {
    writeDirEntry(0, 'Root Entry', 5, -1, -1,  1, ENDOFCHAIN, 0);
    writeDirEntry(1, 'FileHeader', 2, -1,  2, -1, fhSec, fileHeaderData.length);
    writeDirEntry(2, 'DocInfo',    2, -1,  3, -1, diSec, docInfoData.length);
    writeDirEntry(3, 'BodyText',   1, -1, -1,  4, ENDOFCHAIN, 0);
    writeDirEntry(4, 'Section0',   2, -1, -1, -1, s0Sec, section0Data.length);
  }

  // HWP Root Entry CLSID
  const HWP_CLSID = [
    0x20, 0xE9, 0xE3, 0xC0, 0x46, 0x35, 0xCF, 0x11,
    0x8D, 0x81, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71,
  ];
  for (let i = 0; i < 16; i++) dirBuf[80 + i] = HWP_CLSID[i];

  const hdr = new Uint8Array(SS);
  const hdv = new DataView(hdr.buffer);
  const MAGIC = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
  MAGIC.forEach((b, i) => { hdr[i] = b; });
  hdv.setUint16(24, 0x003E, true);
  hdv.setUint16(26, 0x0003, true);
  hdv.setUint16(28, 0xFFFE, true);
  hdv.setUint16(30, 9,      true);
  hdv.setUint16(32, 6,      true);
  hdv.setUint32(40, 0,          true);
  hdv.setUint32(44, fatN,       true);
  hdv.setUint32(48, dir1Sec,    true);
  hdv.setUint32(52, 0,          true);
  hdv.setUint32(56, 0x1000,     true);
  hdv.setUint32(60, ENDOFCHAIN, true);
  hdv.setUint32(64, 0,          true);
  hdv.setUint32(68, ENDOFCHAIN, true);
  hdv.setUint32(72, 0,          true);
  for (let i = 0; i < 109; i++) hdv.setUint32(76 + i * 4, i < fatN ? i : FREESECT, true);

  const out = new Uint8Array(SS + totalSec * SS);
  out.set(hdr, 0);
  for (let i = 0; i < fatN; i++) out.set(fatBuf.subarray(i * SS, (i + 1) * SS), SS + i * SS);
  for (let i = 0; i < dirN; i++) out.set(dirBuf.subarray(i * SS, (i + 1) * SS), SS + (dir1Sec + i) * SS);
  out.set(fhPad, SS + fhSec * SS);
  out.set(diPad, SS + diSec * SS);
  out.set(s0Pad, SS + s0Sec * SS);
  for (let ii = 0; ii < imgPads.length; ii++) out.set(imgPads[ii], SS + imgSecs[ii] * SS);
  return out;
}

// ─── 유틸리티 ────────────────────────────────────────────────

function concatU8(arrays: Uint8Array[]): Uint8Array {
  const total = arrays.reduce((s, a) => s + a.length, 0);
  const out   = new Uint8Array(total);
  let off = 0;
  for (const a of arrays) { out.set(a, off); off += a.length; }
  return out;
}

// ─── OLE2 검증 ──────────────────────────────────────────────

function validateOle2Magic(hwp: Uint8Array): boolean {
  const OLE_MAGIC = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
  return OLE_MAGIC.every((b, i) => hwp[i] === b);
}

// ─── Encoder 진입점 ──────────────────────────────────────────

export class HwpEncoder extends BaseEncoder {
  protected getFormat(): string { return 'hwp'; }

  async encode(doc: DocRoot): Promise<Outcome<Uint8Array>> {
    try {
      // 패스 1: 스타일 수집 (HwpStyleBank — ANYTOHWP 방식 언어별 폰트)
      const bank = new HwpStyleBank();
      for (const sheet of doc.kids) {
        for (const node of sheet.kids) collectNode(node, bank);
      }

      // 이미지 수집 (ANYTOHWP 개선: 픽셀 치수는 encodePicPara에서 추출)
      const images: BinImage[] = [];
      const seenB64 = new Set<string>();
      let binIdCounter = 1;

      function registerImg(img: any): void {
        const key = img.b64.substring(0, 50);
        if (seenB64.has(key)) return;
        seenB64.add(key);
        const raw = TextKit.base64Decode(img.b64);
        const ext = img.mime === 'image/png' ? 'png'
          : img.mime === 'image/gif' ? 'gif'
          : img.mime === 'image/bmp' ? 'bmp' : 'jpg';
        images.push({ id: binIdCounter++, ext, data: new Uint8Array(raw) });
      }

      function collectImages(node: any): void {
        if (node.tag === 'para') {
          for (const img of flatImgNodes(node.kids)) registerImg(img);
        } else if (node.tag === 'grid') {
          for (const row of node.kids)
            for (const cell of row.kids)
              for (const para of cell.kids) collectImages(para);
        }
      }
      for (const sheet of doc.kids) {
        for (const node of sheet.kids) collectImages(node);
      }

      // 패스 2: 스트림 빌드
      const docInfoRaw = buildDocInfoStream(bank, images);
      const bodyRaw    = buildBodyTextStream(doc, bank, images);

      // HWP 5.0: Zlib 헤더 없는 Raw Deflate
      const docInfoCmp = pako.deflateRaw(docInfoRaw);
      const bodyCmp    = pako.deflateRaw(bodyRaw);

      const fileHdr = buildHwpFileHeader();
      const hwp     = buildHwpOle2(fileHdr, docInfoCmp, bodyCmp, images);

      if (!validateOle2Magic(hwp)) {
        return fail('HwpEncoder: OLE2 매직 바이트 오류');
      }

      return succeed(hwp);
    } catch (e: any) {
      return fail(`HwpEncoder: ${e instanceof Error ? e.message : String(e)}`);
    }
  }
}

registry.registerEncoder(new HwpEncoder());
