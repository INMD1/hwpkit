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
import { Metric }             from '../../safety/StyleBridge';
import { registry }           from '../../pipeline/registry';
import { A4 }                 from '../../model/doc-props';
import pako                   from 'pako';

/* ═══════════════════════════════════════════════════════════════
   HWP 5.0 tag IDs
   ═══════════════════════════════════════════════════════════════ */

const T = 16; // HWPTAG_BEGIN
const TAG_ID_MAPPINGS    = T + 8;   // 24
const TAG_FACE_NAME      = T + 3;   // 19
const TAG_BORDER_FILL    = T + 4;   // 20
const TAG_CHAR_SHAPE     = T + 5;   // 21
const TAG_PARA_SHAPE     = T + 9;   // 25
const TAG_PARA_HEADER    = T + 50;  // 66
const TAG_PARA_TEXT      = T + 51;  // 67
const TAG_PARA_CHAR_SHAPE = T + 52; // 68
const TAG_CTRL_HEADER    = T + 55;  // 71
const TAG_LIST_HEADER    = T + 56;  // 72
const TAG_PAGE_DEF       = T + 57;  // 73
const TAG_TABLE_B        = T + 64;  // 80

const CTRL_TABLE = 0x74626C20; // ' lbt' as LE uint32

/** Border width index table (points) — matches BORDER_W_PT in HwpScanner */
const BORDER_W_PT = [
  0.28, 0.34, 0.43, 0.57, 0.71, 0.85,
  1.13, 1.42, 1.70, 1.98, 2.84, 4.25,
  5.67, 8.50, 11.34, 14.17,
];

const BORDER_KIND_IDX: Record<string, number> = {
  none: 0, solid: 1, dash: 2, dot: 3, double: 8,
};
const ALIGN_CODE: Record<string, number> = {
  justify: 0, left: 1, right: 2, center: 3,
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
  return [p.align ?? 'left', p.indentPt ?? 0, p.spaceBefore ?? 0,
    p.spaceAfter ?? 0, p.lineHeight ?? 1].join('|');
}
function bfKey(s: Stroke, bg?: string): string {
  return `${s.kind}|${s.pt}|${s.color}|${bg ?? ''}`;
}

class StyleCollector {
  readonly DEF_STROKE: Stroke = { kind: 'solid', pt: 0.5, color: '000000' };

  fonts: string[] = ['Malgun Gothic'];
  private fontIdx = new Map<string, number>([['Malgun Gothic', 0]]);

  csProps: TextProps[] = [{}];
  private csIdx = new Map<string, number>([[csKey({}), 0]]);

  psProps: ParaProps[] = [{}];
  private psIdx = new Map<string, number>([[psKey({}), 0]]);

  bfData: { s: Stroke; bg?: string }[] = [];
  private bfIdx = new Map<string, number>();

  constructor() {
    this.addBorderFill(this.DEF_STROKE); // bfId=1
  }

  font(name: string): number {
    const n = name || 'Malgun Gothic';
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

  /** Returns 1-based border fill ID (HWP uses 1-based IDs for border fills) */
  addBorderFill(s: Stroke, bg?: string): number {
    const k = bfKey(s, bg);
    if (this.bfIdx.has(k)) return this.bfIdx.get(k)!;
    const id = this.bfData.length + 1;
    this.bfData.push({ s, bg });
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
        col.addBorderFill(cell.props.top ?? node.props.defaultStroke ?? col.DEF_STROKE, cell.props.bg);
        for (const para of cell.kids) collectNode(para, col);
      }
    }
  }
}

/* ═══════════════════════════════════════════════════════════════
   DocInfo record builders
   ═══════════════════════════════════════════════════════════════ */

function mkIdMappings(col: StyleCollector): Uint8Array {
  return new BufWriter()
    .u32(col.fonts.length)
    .u32(col.bfData.length)
    .u32(col.csProps.length)
    .u32(0) // tabDef count
    .u32(0) // numbering count
    .u32(0) // bullet count
    .u32(col.psProps.length)
    .u32(0) // style count
    .build();
}

function mkFaceName(name: string): Uint8Array {
  return new BufWriter()
    .u8(0)          // substType
    .u16(name.length)
    .utf16(name)
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
  const w = new BufWriter();
  w.u16(0); // attr
  const t  = BORDER_KIND_IDX[s.kind] ?? 1;
  const wi = borderWidthIdx(s.pt);
  // 5 borders: left, right, top, bottom, diagonal
  for (let i = 0; i < 5; i++) w.u8(t).u8(wi).colorRef(s.color || '000000');
  // fill: type(4) + faceColor(4) + reserved(4)
  if (bg) { w.u32(1).colorRef(bg).u32(0); }
  else    { w.u32(0).u32(0).u32(0); }
  return w.build(); // 40 bytes
}

function mkCharShape(p: TextProps, col: StyleCollector): Uint8Array {
  const fontId = p.font ? col.font(p.font) : 0;
  const w = new BufWriter();
  for (let i = 0; i < 7; i++) w.u16(fontId);  // faceId[7]
  for (let i = 0; i < 7; i++) w.u8(100);       // ratio[7]
  for (let i = 0; i < 7; i++) w.u8(0);         // spacing[7]
  for (let i = 0; i < 7; i++) w.u8(100);       // relSize[7]
  for (let i = 0; i < 7; i++) w.u8(0);         // offset[7]
  // height @ offset 42 (HWPUNIT: pt × 100)
  w.u32(Math.round((p.pt ?? 10) * 100));
  // attr @ offset 46
  let attr = 0;
  if (p.i)   attr |= 1;          // italic  = bit 0
  if (p.b)   attr |= 2;          // bold    = bit 1
  if (p.u)   attr |= (1 << 2);   // ulType  = bits 2-4, set to 1
  if (p.s)   attr |= (1 << 18);  // skType  = bits 18-20, set to 1
  if (p.sup) attr |= (1 << 16);  // suType  = bits 16-17, value 1
  if (p.sub) attr |= (2 << 16);  // suType  = bits 16-17, value 2
  w.u32(attr);
  w.u8(0).u8(0); // shadowX, shadowY @ 50-51
  w.colorRef(p.color ?? '000000'); // textColor @ 52 (4 bytes)
  return w.build(); // 56 bytes
}

function mkParaShape(p: ParaProps): Uint8Array {
  return new BufWriter()
    .u32(ALIGN_CODE[p.align ?? 'left'] ?? 1) // attr (bits 0-2 = align)
    .i32(Metric.ptToHwp(p.indentPt ?? 0))    // leftMargin
    .i32(0)                                    // rightMargin
    .i32(0)                                    // indent (first-line)
    .i32(Metric.ptToHwp(p.spaceBefore ?? 0))
    .i32(Metric.ptToHwp(p.spaceAfter ?? 0))
    .i32(p.lineHeight ? Math.round(p.lineHeight * 100) : 160) // lineSpacing
    .build(); // 28 bytes
}

function buildDocInfoStream(col: StyleCollector): Uint8Array {
  const chunks: Uint8Array[] = [
    mkRec(TAG_ID_MAPPINGS, 0, mkIdMappings(col)),
    ...col.fonts.map(n => mkRec(TAG_FACE_NAME, 0, mkFaceName(n))),
    ...col.bfData.map(({ s, bg }) => mkRec(TAG_BORDER_FILL, 0, mkBorderFill(s, bg))),
    ...col.csProps.map(p => mkRec(TAG_CHAR_SHAPE, 0, mkCharShape(p, col))),
    ...col.psProps.map(p => mkRec(TAG_PARA_SHAPE, 0, mkParaShape(p))),
  ];
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

function mkParaHeader(psId: number, csCount: number): Uint8Array {
  return new BufWriter()
    .u32(0)       // paragraphControlMask
    .u16(0)       // styleId
    .u8(0)        // divideAttr
    .u8(0)
    .u16(psId)    // paraShapeId  @ offset 8
    .u16(csCount) // charShapeCount @ offset 10
    .u16(0)       // rangeTagCount
    .u16(0)       // memoCount
    .i32(0)       // paraChangeId
    .build(); // 20 bytes
}

function mkParaText(text: string): Uint8Array {
  const w = new BufWriter();
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    w.u16(c < 32 ? 0 : c); // replace control chars
  }
  w.u16(13); // paragraph terminator (0x000D)
  return w.build();
}

function mkParaCharShape(pairs: [pos: number, id: number][]): Uint8Array {
  const w = new BufWriter();
  for (const [pos, id] of pairs) w.u32(pos).u32(id);
  return w.build();
}

function encodePara(para: ParaNode, col: StyleCollector, lv: number): Uint8Array[] {
  let text = '';
  const csPairs: [number, number][] = [];
  let pos = 0;

  for (const kid of para.kids) {
    if (kid.tag !== 'span') continue;
    const span = kid as SpanNode;
    const csId = col.addCharShape(span.props);
    // Only add a new pair when shape changes
    if (csPairs.length === 0 || csPairs[csPairs.length - 1][1] !== csId) {
      csPairs.push([pos, csId]);
    }
    for (const t of span.kids) {
      if (t.tag === 'txt') { text += t.content; pos += t.content.length; }
    }
  }

  if (csPairs.length === 0) csPairs.push([0, 0]);

  const psId = col.addParaShape(para.props);
  return [
    mkRec(TAG_PARA_HEADER,    lv,     mkParaHeader(psId, csPairs.length)),
    mkRec(TAG_PARA_TEXT,      lv + 1, mkParaText(text)),
    mkRec(TAG_PARA_CHAR_SHAPE, lv + 1, mkParaCharShape(csPairs)),
  ];
}

/* ── Table encoding ─────────────────────────────────────────── */

function mkTableCtrl(): Uint8Array {
  return new BufWriter().u32(CTRL_TABLE).zeros(12).build();
}

function mkTableB(rowCnt: number, colCnt: number, rowHwp: number[], bfId: number): Uint8Array {
  const w = new BufWriter();
  w.u32(0);         // attr
  w.u16(rowCnt);
  w.u16(colCnt);
  w.zeros(10);      // bytes 8-17: cell spacing / zone info
  for (const h of rowHwp) w.u16(h);
  w.u16(bfId);
  return w.build();
}

function mkCellListHeader(
  paraCount: number,
  row: number, col: number,
  rs: number,  cs: number,
  wHwp: number, hHwp: number,
  bfId: number,
): Uint8Array {
  // Scanner reads: col = readU16LE(d, 8), row = readU16LE(d, 10)
  // (HWP 5.0 spec: offset 8 = colAddr, offset 10 = rowAddr)
  return new BufWriter()
    .u16(paraCount) // 0-1: paraCount
    .u32(0)         // 2-5: attr
    .u16(0)         // 6-7: unknown
    .u16(col)       // 8-9:  colAddr  ← col first!
    .u16(row)       // 10-11: rowAddr ← then row
    .u16(rs)        // 12-13: rowSpan
    .u16(cs)        // 14-15: colSpan
    .u32(wHwp)      // 16-19: width
    .u32(hHwp)      // 20-23: height
    .zeros(8)       // 24-31: padding[4]
    .u16(bfId)      // 32-33: borderFillId
    .build(); // 34 bytes
}

const DEFAULT_ROW_HEIGHT_PT = 14; // reasonable row height

function encodeGrid(grid: GridNode, col: StyleCollector, lv: number): Uint8Array[] {
  const records: Uint8Array[] = [];
  const rowCnt = grid.kids.length;
  const colCnt = Math.max(1, grid.kids[0]?.kids.length ?? 1);

  // Column widths
  const cwPt = grid.props.colWidths ?? [];
  const totalPt = cwPt.reduce((s, w) => s + w, 0) || 453; // ~A4 content width
  const defColPt = totalPt / colCnt;

  const defStroke = grid.props.defaultStroke ?? col.DEF_STROKE;
  const defBfId   = col.addBorderFill(defStroke);
  const rowHwp    = Array.from({ length: rowCnt }, () => Metric.ptToHwp(DEFAULT_ROW_HEIGHT_PT));

  records.push(mkRec(TAG_CTRL_HEADER, lv,     mkTableCtrl()));
  records.push(mkRec(TAG_TABLE_B,     lv + 1, mkTableB(rowCnt, colCnt, rowHwp, defBfId)));

  for (let r = 0; r < grid.kids.length; r++) {
    for (let c = 0; c < grid.kids[r].kids.length; c++) {
      const cell   = grid.kids[r].kids[c];
      const wHwp   = Metric.ptToHwp(cwPt[c] ?? defColPt);
      const hHwp   = rowHwp[r];
      const stroke = cell.props.top ?? defStroke;
      const bfId   = col.addBorderFill(stroke, cell.props.bg);
      const paras  = cell.kids.length > 0 ? cell.kids : [{ tag: 'para' as const, props: {}, kids: [] }];

      records.push(mkRec(TAG_LIST_HEADER, lv + 1,
        mkCellListHeader(paras.length, r, c, cell.rs, cell.cs, wHwp, hHwp, bfId)));

      // Cell paragraphs are at same level as LIST_HEADER (lv+1);
      // their children (PARA_TEXT, PARA_CHAR_SHAPE) go to lv+2.
      for (const para of paras) records.push(...encodePara(para, col, lv + 1));
    }
  }

  return records;
}

function buildBodyTextStream(doc: DocRoot, col: StyleCollector): Uint8Array {
  const chunks: Uint8Array[] = [];
  const dims = doc.kids[0]?.dims ?? A4;
  chunks.push(mkRec(TAG_PAGE_DEF, 0, mkPageDef(dims)));

  for (const sheet of doc.kids) {
    for (const node of sheet.kids) {
      if (node.tag === 'para') {
        for (const r of encodePara(node as ParaNode, col, 0)) chunks.push(r);
      } else if (node.tag === 'grid') {
        // In HWP, a table is embedded inside a "container paragraph" at level 0.
        // CTRL_HEADER goes at level 1 (child of that paragraph).
        // TABLE_B / LIST_HEADER / cell PARA_HEADERs go at level 2.
        // Cell PARA_TEXT / PARA_CHAR_SHAPE go at level 3.
        chunks.push(mkRec(TAG_PARA_HEADER, 0, mkParaHeader(0, 1)));
        chunks.push(mkRec(TAG_PARA_TEXT, 1, mkParaText('')));
        chunks.push(mkRec(TAG_PARA_CHAR_SHAPE, 1, mkParaCharShape([[0, 0]])));
        for (const r of encodeGrid(node as GridNode, col, 1)) chunks.push(r);
      }
    }
  }

  return concatU8(chunks);
}

/* ═══════════════════════════════════════════════════════════════
   HWP FileHeader stream (256 bytes)
   ═══════════════════════════════════════════════════════════════ */

function buildHwpFileHeader(): Uint8Array {
  const buf = new Uint8Array(256);
  const sig = 'HWP Document File';
  for (let i = 0; i < sig.length; i++) buf[i] = sig.charCodeAt(i);
  const dv = new DataView(buf.buffer);
  dv.setUint32(32, 0x05000300, true); // version 5.0.3.0
  dv.setUint32(36, 0x00000001, true); // flags: bit 0 = compressed
  return buf;
}

/* ═══════════════════════════════════════════════════════════════
   OLE2 / CFB container builder
   Structure:
     OLE2 header (512 bytes, not a sector)
     Sector 0..fatN-1    : FAT sectors
     Sector fatN         : Directory sector 1 (entries 0-3)
     Sector fatN+1       : Directory sector 2 (entries 4-7)
     Sector fatN+2 ..    : FileHeader data
     then DocInfo data, then Section0 data
   ═══════════════════════════════════════════════════════════════ */

function buildHwpOle2(
  fileHeaderData: Uint8Array,
  docInfoData:    Uint8Array,
  section0Data:   Uint8Array,
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

  const fhPad = padSector(fileHeaderData);
  const diPad = padSector(docInfoData);
  const s0Pad = padSector(section0Data);
  const fhN   = fhPad.length / SS;
  const diN   = diPad.length / SS;
  const s0N   = s0Pad.length / SS;
  const dirN  = 2; // always 2 dir sectors (holds 8 dir entries)

  // Compute FAT sector count iteratively
  let fatN = 1;
  for (let iter = 0; iter < 10; iter++) {
    const total  = fatN + dirN + fhN + diN + s0N;
    const needed = Math.ceil(total / 128);
    if (needed <= fatN) break;
    fatN = needed;
  }

  // Assign sector indices
  const dir1Sec  = fatN;
  const dir2Sec  = fatN + 1;
  const fhSec    = fatN + dirN;
  const diSec    = fhSec + fhN;
  const s0Sec    = diSec + diN;
  const totalSec = s0Sec + s0N;

  // Build FAT (fatN × 128 entries × 4 bytes = fatN × 512 bytes)
  const fatBuf = new Uint8Array(fatN * SS).fill(0xFF); // FREESECT
  const setFat = (i: number, v: number) => {
    fatBuf[i * 4]     = v & 0xFF;
    fatBuf[i * 4 + 1] = (v >>> 8) & 0xFF;
    fatBuf[i * 4 + 2] = (v >>> 16) & 0xFF;
    fatBuf[i * 4 + 3] = (v >>> 24) & 0xFF;
  };

  for (let i = 0; i < fatN; i++) setFat(i, FATSECT);
  setFat(dir1Sec, dir2Sec);
  setFat(dir2Sec, ENDOFCHAIN);
  for (let i = 0; i < fhN; i++) setFat(fhSec + i, i + 1 < fhN ? fhSec + i + 1 : ENDOFCHAIN);
  for (let i = 0; i < diN; i++) setFat(diSec + i, i + 1 < diN ? diSec + i + 1 : ENDOFCHAIN);
  for (let i = 0; i < s0N; i++) setFat(s0Sec + i, i + 1 < s0N ? s0Sec + i + 1 : ENDOFCHAIN);

  // Build directory (8 entries × 128 bytes = dirN × SS)
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
    dv.setUint16(base + 64, (nl + 1) * 2, true); // name size (incl. null)
    dirBuf[base + 66] = type;
    dirBuf[base + 67] = 1; // color = black
    dv.setInt32(base + 68, left,  true); // left sibling
    dv.setInt32(base + 72, right, true); // right sibling
    dv.setInt32(base + 76, child, true); // child
    dv.setUint32(base + 116, startSec >>> 0, true);
    dv.setUint32(base + 120, size >>> 0,     true);
  }

  // Use right-skewed sibling chain (no left siblings) to avoid cycles in CFB parsers.
  // Root.child → FileHeader → DocInfo → BodyText (via sibRight).
  // BodyText.child → Section0.
  writeDirEntry(0, 'Root Entry', 5, -1, -1,  1, ENDOFCHAIN, 0);
  writeDirEntry(1, 'FileHeader', 2, -1,  2, -1, fhSec, fileHeaderData.length);
  writeDirEntry(2, 'DocInfo',   2,  -1,  3, -1, diSec, docInfoData.length);
  writeDirEntry(3, 'BodyText',  1,  -1, -1,  4, ENDOFCHAIN, 0);
  writeDirEntry(4, 'Section0',  2,  -1, -1, -1, s0Sec, section0Data.length);
  // Entries 5-7: type=0 (empty), everything else zeroed

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
  // OLE2 v3 header field layout (see ECMA-376 or MS-CFB spec):
  //   40-43: num dir sectors    (must be 0 for v3)
  //   44-47: num FAT sectors
  //   48-51: first dir sector
  //   52-55: transaction sig    (0)
  //   56-59: mini stream cutoff (4096)
  //   60-63: first mini FAT     (ENDOFCHAIN if none)
  //   64-67: num mini FAT       (0)
  //   68-71: first DIFAT ext    (ENDOFCHAIN if none)
  //   72-75: num DIFAT ext      (0)
  hdv.setUint32(40, 0,          true); // num dir sectors (0 for v3)
  hdv.setUint32(44, fatN,       true); // num FAT sectors
  hdv.setUint32(48, dir1Sec,    true); // first directory sector
  hdv.setUint32(52, 0,          true); // transaction signature (0)
  hdv.setUint32(56, 0x1000,     true); // mini stream cutoff = 4096
  hdv.setUint32(60, ENDOFCHAIN, true); // first mini FAT sector (none)
  hdv.setUint32(64, 0,          true); // num mini FAT sectors (0)
  hdv.setUint32(68, ENDOFCHAIN, true); // first DIFAT extension (none)
  hdv.setUint32(72, 0,          true); // num DIFAT extensions (0)
  // DIFAT[0..108]: first fatN entries = FAT sector numbers
  for (let i = 0; i < 109; i++) {
    hdv.setUint32(76 + i * 4, i < fatN ? i : FREESECT, true);
  }

  // Assemble output
  const out = new Uint8Array(SS + totalSec * SS);
  out.set(hdr, 0);
  // FAT sectors
  for (let i = 0; i < fatN; i++) {
    out.set(fatBuf.subarray(i * SS, (i + 1) * SS), SS + i * SS);
  }
  // Directory sectors
  out.set(dirBuf.subarray(0, SS),    SS + dir1Sec * SS);
  out.set(dirBuf.subarray(SS, 2*SS), SS + dir2Sec * SS);
  // Stream data
  out.set(fhPad, SS + fhSec * SS);
  out.set(diPad, SS + diSec * SS);
  out.set(s0Pad, SS + s0Sec * SS);
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

      // Build streams
      const docInfoRaw  = buildDocInfoStream(col);
      const bodyRaw     = buildBodyTextStream(doc, col);

      // Compress (HWP flags bit 0 = compressed)
      const docInfoCmp = pako.deflateRaw(docInfoRaw);
      const bodyCmp    = pako.deflateRaw(bodyRaw);

      // Assemble OLE2 file
      const fileHdr = buildHwpFileHeader();
      const hwp     = buildHwpOle2(fileHdr, docInfoCmp, bodyCmp);

      return succeed(hwp);
    } catch (e: any) {
      return fail(`HwpEncoder: ${e instanceof Error ? e.message : String(e)}`);
    }
  }
}

registry.registerEncoder(new HwpEncoder());
