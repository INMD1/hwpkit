/**
 * HwpxEncoder — DocRoot → HWPX (ZIP + XML)
 *
 * ANYTOHWP에서 영감받은 개선 사항:
 *  1. LangFontBank  — 7개 언어 그룹 독립 폰트 레지스트리 (HANGUL/LATIN/HANJA/…)
 *  2. BorderFillBank — 정확한 ID 관리 (하드코딩 "1" 제거)
 *  3. readPixelDims  — PNG/JPEG 바이너리 헤더에서 실제 픽셀 치수 추출
 *  4. 두 패스 구조   — Pre-scan(등록) → Encode(생성)
 */

import type {
  DocRoot,
  ParaNode,
  SpanNode,
  GridNode,
  ContentNode,
  ImgNode,
  SheetNode,
  CellNode,
  LinkNode,
} from "../../model/doc-tree";
import type { Outcome } from "../../contract/result";
import { BaseEncoder } from "../../core/BaseEncoder";
import type {
  DocMeta,
  PageDims,
  TextProps,
  ParaProps,
  CellProps,
  Stroke,
  ImgLayout,
} from "../../model/doc-props";
import { A4, DEFAULT_STROKE, normalizeDims } from "../../model/doc-props";
import { succeed, fail } from "../../contract/result";
import { Metric, safeFontToKr } from "../../safety/StyleBridge";
import { ArchiveKit } from "../../toolkit/ArchiveKit";
import { TextKit } from "../../toolkit/TextKit";
import { fitColumnWidths } from "../../toolkit/TableGeometry";
import { registry } from "../../pipeline/registry";
import { HWPX_MIME_TYPE } from "./constants";

// ─── HWPX 네임스페이스 ──────────────────────────────────────
const NS = [
  'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app"',
  'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"',
  'xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph"',
  'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"',
  'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"',
  'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"',
  'xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history"',
  'xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page"',
  'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"',
  'xmlns:dc="http://purl.org/dc/elements/1.1/"',
  'xmlns:opf="http://www.idpf.org/2007/opf/"',
  'xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart"',
  'xmlns:epub="http://www.idpf.org/2007/ops"',
  'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"',
  'xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"',
].join(" ");

// ─── LinesegArray Flags 상수 (HWPX 스펙) ─────────────────
// HWP 5.0 §4.3.4: bits 17/18 mark the first/last segment of a line.
// Bit 20 (0x100000) means indentation was applied; it does not mean "first line".
const LINESEG_FLAGS = 0x60000;
const LINESEG_FLAG_INDENT = 0x100000;
const LINESEG_FLAG_PAGE_FIRST = 0x1;
const LINESEG_FLAG_COLUMN_FIRST = 0x2;

// ─── ANYTOHWP 영감: 언어별 폰트 레지스트리 ─────────────────
// 7개 언어 그룹을 독립적으로 관리 — charPr fontRef의 정확한 ID 생성
const LANG_GROUPS = [
  "HANGUL",
  "LATIN",
  "HANJA",
  "JAPANESE",
  "OTHER",
  "SYMBOL",
  "USER",
] as const;
type LangGroup = (typeof LANG_GROUPS)[number];

class LangFontBank {
  // 언어 그룹별 독립 폰트 맵: face → localId (0-based)
  private maps = new Map<LangGroup, Map<string, number>>(
    LANG_GROUPS.map((g) => [g, new Map<string, number>()]),
  );

  constructor() {
    // ANYTOHWP 기본값: 모든 그룹에 한컴 기본 폰트 등록 (id=0)
    this.registerAll("함초롬바탕");
  }

  /** 모든 언어 그룹에 동일 폰트 등록 */
  registerAll(face: string): void {
    for (const g of LANG_GROUPS) {
      const m = this.maps.get(g)!;
      if (!m.has(face)) m.set(face, m.size);
    }
  }

  /** 특정 언어 그룹에 폰트 등록, 이미 있으면 기존 ID 반환 */
  register(lang: LangGroup, face: string): number {
    const m = this.maps.get(lang)!;
    if (m.has(face)) return m.get(face)!;
    const id = m.size;
    m.set(face, id);
    return id;
  }

  /** 폰트 이름 → 한글 폰트 여부 판별 (ANYTOHWP 방식) */
  private isKorean(face: string): boolean {
    return (
      /[\uAC00-\uD7A3\u3131-\u318E]/.test(face) ||
      ["맑은", "나눔", "굴림", "돋움", "바탕", "함초롬", "한컴", "HY"].some(
        (k) => face.includes(k),
      )
    );
  }

  /** Register a face in every language bank and return bank-local IDs. */
  registerFont(rawFace: string): Record<LangGroup, number> {
    const face = safeFontToKr(rawFace) || "함초롬바탕";
    const isKor = this.isKorean(face);
    const ids = {} as Record<LangGroup, number>;
    for (const group of LANG_GROUPS) {
      const useFace = group === "LATIN"
        ? (isKor ? "함초롬바탕" : face)
        : (isKor ? face : "함초롬바탕");
      ids[group] = this.register(group, useFace);
    }
    return ids;
  }

  /** 언어 그룹별 폰트 목록 반환 */
  getFaces(lang: LangGroup): string[] {
    return [...this.maps.get(lang)!.keys()];
  }

  getId(lang: LangGroup, face: string): number {
    return this.maps.get(lang)!.get(face) ?? 0;
  }

  /** hh:fontfaces XML 생성 */
  toXml(): string {
    let xml = `<hh:fontfaces itemCnt="${LANG_GROUPS.length}">`;
    for (const lang of LANG_GROUPS) {
      const faces = this.getFaces(lang);
      xml += `<hh:fontface lang="${lang}" fontCnt="${faces.length}">`;
      faces.forEach((face, i) => {
        xml +=
          `<hh:font id="${i}" face="${esc(face)}" type="TTF" isEmbedded="0">` +
          `<hh:typeInfo familyType="FCAT_UNKNOWN" weight="0" proportion="0" contrast="0" strokeVariation="0" armStyle="0" letterform="0" midline="252" xHeight="255"/>` +
          `</hh:font>`;
      });
      xml += `</hh:fontface>`;
    }
    return xml + `</hh:fontfaces>`;
  }
}

// ─── ANYTOHWP 영감: BorderFill 레지스트리 ───────────────────
// 하드코딩 "1" 제거 — 모든 셀/표의 실제 테두리를 추적

const KIND_MAP: Record<string, string> = {
  solid: "SOLID",
  dash: "DASH",
  dot: "DOT",
  double: "DOUBLE",
  none: "NONE",
  dash_dot: "DASH_DOT",
  dash_dot_dot: "DASH_DOT_DOT",
};

/**
 * 테두리 선 굵기 mm 값을 한글 표준 규격 리스트 중 가장 가까운 값으로 매핑(양자화)합니다.
 */
function quantizeBorderWidth(pt: number): string {
  const mm = pt * 0.3528;
  const standardWidths = [0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0];
  let closest = standardWidths[0];
  let minDiff = Math.abs(mm - closest);
  for (let i = 1; i < standardWidths.length; i++) {
    const diff = Math.abs(mm - standardWidths[i]);
    if (diff < minDiff) {
      minDiff = diff;
      closest = standardWidths[i];
    }
  }
  let str = closest.toFixed(2);
  if (str.endsWith("0")) {
    str = str.slice(0, -1);
  }
  if (str.endsWith(".0")) {
    str = str.slice(0, -2);
  }
  return `${str} mm`;
}

class BorderFillBank {
  private fills: { id: number; xml: string }[] = [];
  private keyMap = new Map<string, number>();

  constructor() {
    // id=1: 기본 (테두리 없음) — ANYTOHWP의 기본 초기화 방식
    this._addXml(
      this._buildXml(undefined, undefined, undefined, undefined, undefined),
    );
    // id=2: 표 기본 테두리 (solid 0.5pt black)
    const defS: Stroke = { kind: "solid", pt: 0.5, color: "000000" };
    this._addXml(this._buildXml(defS, defS, defS, defS, undefined));
  }

  private _strokeXml(tag: string, s?: Stroke): string {
    const type =
      s && s.kind !== "none" ? (KIND_MAP[s.kind] ?? "SOLID") : "NONE";
    const w =
      s && s.kind !== "none" ? quantizeBorderWidth(s.pt) : "0.12 mm";
    const c = s
      ? s.color.startsWith("#")
        ? s.color
        : `#${s.color}`
      : "#000000";
    return `<hh:${tag} type="${type}" width="${w}" color="${c}"/>`;
  }

  private _buildXml(
    top?: Stroke,
    right?: Stroke,
    bottom?: Stroke,
    left?: Stroke,
    bg?: string,
  ): string {
    const fill = bg
      ? `<hc:fillBrush><hc:winBrush faceColor="${bg.startsWith("#") ? bg : "#" + bg}" hatchColor="none" alpha="0"/></hc:fillBrush>`
      : "";
    return (
      `<hh:borderFill id="__ID__" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">` +
      `<hh:slash type="NONE" Crooked="0" isCounter="0"/>` +
      `<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>` +
      this._strokeXml("leftBorder", left) +
      this._strokeXml("rightBorder", right) +
      this._strokeXml("topBorder", top) +
      this._strokeXml("bottomBorder", bottom) +
      `<hh:diagonal type="NONE" width="0.12 mm" color="#000000"/>` +
      fill +
      `</hh:borderFill>`
    );
  }

  private _addXml(xml: string): number {
    const id = this.fills.length + 1;
    this.fills.push({ id, xml: xml.replace("__ID__", String(id)) });
    return id;
  }

  private _key(
    top?: Stroke,
    right?: Stroke,
    bottom?: Stroke,
    left?: Stroke,
    bg?: string,
  ): string {
    const sk = (s?: Stroke) =>
      s ? `${s.kind}:${s.pt.toFixed(2)}:${s.color}` : "none";
    return `${sk(top)}|${sk(right)}|${sk(bottom)}|${sk(left)}|${bg ?? ""}`;
  }

  /** 균일 테두리 등록 */
  addUniform(s?: Stroke, bg?: string): number {
    const key = this._key(s, s, s, s, bg);
    if (this.keyMap.has(key)) return this.keyMap.get(key)!;
    const id = this._addXml(this._buildXml(s, s, s, s, bg));
    this.keyMap.set(key, id);
    return id;
  }

  /** 방향별 테두리 등록 */
  addPerSide(
    top?: Stroke,
    right?: Stroke,
    bottom?: Stroke,
    left?: Stroke,
    bg?: string,
  ): number {
    const key = this._key(top, right, bottom, left, bg);
    if (this.keyMap.has(key)) return this.keyMap.get(key)!;
    const id = this._addXml(this._buildXml(top, right, bottom, left, bg));
    this.keyMap.set(key, id);
    return id;
  }

  /** CellProps에서 적절한 borderFill ID 계산 (하드코딩 "1" 완전 제거) */
  addFromCellProps(cp: CellProps, defStroke?: Stroke): number {
    const d = defStroke ?? DEFAULT_STROKE;
    const top = cp.top ?? d;
    const right = cp.right ?? d;
    const bottom = cp.bot ?? d;
    const left = cp.left ?? d;
    const bg = cp.bg;
    const uniform =
      top.kind === right.kind &&
      top.kind === bottom.kind &&
      top.kind === left.kind &&
      top.pt === right.pt &&
      top.pt === bottom.pt &&
      top.pt === left.pt &&
      top.color === right.color &&
      top.color === bottom.color &&
      top.color === left.color;
    return uniform
      ? this.addUniform(top, bg)
      : this.addPerSide(top, right, bottom, left, bg);
  }

  toXml(): string {
    return `<hh:borderFills itemCnt="${this.fills.length}">${this.fills.map((f) => f.xml).join("")}</hh:borderFills>`;
  }
}

// ─── ANYTOHWP 영감: PNG/JPEG 바이너리 헤더에서 픽셀 치수 추출
function readPixelDims(
  b64: string,
  mime: string,
): { w: number; h: number } | null {
  try {
    const raw = TextKit.base64Decode(b64);
    const view = new DataView(raw.buffer, raw.byteOffset, raw.byteLength);

    if (mime.includes("png")) {
      // PNG: 시그니처 8바이트 + IHDR 청크 길이(4) + 타입(4) + 너비(4) + 높이(4)
      if (
        raw.length >= 24 &&
        view.getUint32(0) === 0x89504e47 &&
        view.getUint32(4) === 0x0d0a1a0a
      ) {
        return { w: view.getUint32(16), h: view.getUint32(20) };
      }
    } else if (mime.includes("jpeg") || mime.includes("jpg")) {
      // JPEG: SOI(FF D8) 후 SOF0(FF C0) 또는 SOF2(FF C2) 마커 탐색
      let off = 2;
      while (off < raw.length - 4) {
        const marker = view.getUint16(off);
        off += 2;
        if (marker === 0xffc0 || marker === 0xffc2) {
          // SOF: length(2) + precision(1) + height(2) + width(2)
          return { w: view.getUint16(off + 5), h: view.getUint16(off + 3) };
        }
        if ((marker & 0xff00) !== 0xff00) break;
        const segLen = view.getUint16(off);
        off += segLen;
      }
    }
  } catch {
    /* 무시 */
  }
  return null;
}

// ─── charPr / paraPr 레지스트리 ─────────────────────────────

interface CharPrDef {
  id: number;
  height: number; // HWPX height 단위 (1000 = 10pt)
  bold: boolean;
  italic: boolean;
  underline: string; // "NONE" | "BOTTOM"
  strikeout: string; // "NONE" | "SOLID"
  textColor: string; // "#RRGGBB"
  hangulId: number; // HANGUL 그룹 폰트 ID
  latinId: number; // LATIN 그룹 폰트 ID
  hanjaId: number;
  japaneseId: number;
  otherId: number;
  symbolId: number;
  userId: number;
  bg?: string;
}
interface ParaPrDef {
  pageBreakBefore?: boolean;
  keepWithNext?: boolean;
  keepLines?: boolean;
  widowControl?: boolean;
  id: number;
  align: string;
  leftHwp: number;
  rightHwp: number;
  intentHwp: number;
  prevHwp: number;
  nextHwp: number;
  lineSpacing: number;
  lineSpacingFixed?: number; // paraPr 배율값(물리 HWPUNIT × 2), emitted as AT_LEAST
  listType?: string;
  listLevel?: number;
  verAlign?: string;
  lineWrap?: string;
}
interface StyleEntry {
  id: number;
  name: string;
  engName: string;
  paraPrIDRef: number;
  charPrIDRef: number;
}

interface BinEntry {
  id: string; // "BIN0001"
  name: string; // "BIN0001.png"
  data: Uint8Array;
}

function charPrKey(p: TextProps): string {
  return `${p.b ? 1 : 0}|${p.i ? 1 : 0}|${p.u ? 1 : 0}|${p.s ? 1 : 0}|${p.pt ?? 10}|${p.color ?? "000000"}|${p.font ?? ""}|${p.bg ?? ""}`;
}

/** paraPr의 배율 저장값을 linesegarray가 사용하는 물리 HWPUNIT로 환산한다. */
function paraShapeHwpToLayoutHwp(value: number): number {
  return Math.round(value / 2);
}

/**
 * ParaProps 를 해시 키로 변환 (동일 포맷팅 감지용)
 * null/undefined는 0 으로 처리하여 일관성 유지
 */
function paraPrKey(p: ParaProps): string {
  return `${p.pageBreakBefore ?? false}|${p.keepWithNext ?? false}|${p.keepLines ?? false}|${p.widowControl ?? false}|${p.align ?? "left"}|${p.verAlign ?? "baseline"}|${p.lineWrap ?? "break"}|${p.listOrd ?? ""}|${p.listLv ?? 0}|${p.indentPt ?? 0}|${p.indentRightPt ?? 0}|${p.firstLineIndentPt ?? 0}|${p.spaceBefore ?? 0}|${p.spaceAfter ?? 0}|${p.lineHeight ?? 0}|${p.lineHeightFixed ?? 0}|${p.styleId ?? ""}`;
}

// ─── 인코딩 컨텍스트 ─────────────────────────────────────────

interface HwpxCtx {
  fontBank: LangFontBank;
  borderFillBank: BorderFillBank;
  charPrs: CharPrDef[];
  charPrMap: Map<string, number>;
  paraPrs: ParaPrDef[];
  paraPrMap: Map<string, number>;
  bins: BinEntry[];
  nextBinNum: number;
  nextElementId: number;
  availableWidth: number; // HWPUNIT
  imgMap: WeakMap<ImgNode, string>;
  nextZOrder: number;
  styleIdToHwpxId: Map<string, number>;
  hwpxStyles: StyleEntry[];
}

function registerCharPr(props: TextProps, ctx: HwpxCtx): number {
  const key = charPrKey(props);
  const existing = ctx.charPrMap.get(key);
  if (existing !== undefined) return existing;

  const rawFont = props.font ?? "함초롬바탕";
  const fontIds = ctx.fontBank.registerFont(rawFont);
  const id = ctx.charPrs.length;

  ctx.charPrs.push({
    id,
    height: Metric.ptToHHeight(props.pt ?? 10),
    bold: !!props.b,
    italic: !!props.i,
    underline: props.u ? "BOTTOM" : "NONE",
    strikeout: props.s ? "SOLID" : "NONE",
    textColor: props.color ? `#${props.color}` : "#000000",
    hangulId: fontIds.HANGUL,
    latinId: fontIds.LATIN,
    hanjaId: fontIds.HANJA,
    japaneseId: fontIds.JAPANESE,
    otherId: fontIds.OTHER,
    symbolId: fontIds.SYMBOL,
    userId: fontIds.USER,
    bg: props.bg,
  });
  ctx.charPrMap.set(key, id);
  return id;
}

const ALIGN_MAP: Record<string, string> = {
  left: "LEFT",
  center: "CENTER",
  right: "RIGHT",
  justify: "JUSTIFY",
  distribute: "DISTRIBUTE",
  distribute_space: "DISTRIBUTE_SPACE",
};

const V_ALIGN_MAP: Record<string, string> = {
  baseline: "BASELINE",
  top: "TOP",
  center: "CENTER",
  bottom: "BOTTOM",
};

const LINE_WRAP_MAP: Record<string, string> = {
  break: "BREAK",
  squeeze: "SQUEEZE",
  keep: "KEEP",
};

function registerParaPr(props: ParaProps, ctx: HwpxCtx): number {
  const key = paraPrKey(props);
  const existing = ctx.paraPrMap.get(key);
  if (existing !== undefined) return existing;

  const id = ctx.paraPrs.length;

  const alignStr = props.align ? (ALIGN_MAP[props.align] ?? "LEFT") : "LEFT";
  const verAlignStr = props.verAlign ? (V_ALIGN_MAP[props.verAlign] ?? "BASELINE") : "BASELINE";
  const lineWrapStr = props.lineWrap ? (LINE_WRAP_MAP[props.lineWrap] ?? "BREAK") : "BREAK";

  const def: ParaPrDef = {
    pageBreakBefore: props.pageBreakBefore,
    keepWithNext: props.keepWithNext,
    keepLines: props.keepLines,
    widowControl: props.widowControl,
    id,
    align: alignStr,
    verAlign: verAlignStr,
    lineWrap: lineWrapStr,
    // HWPX stores the margin before applying a negative first-line indent.
    leftHwp: Metric.ptToHwp((props.indentPt ?? 0) + Math.min(0, props.firstLineIndentPt ?? 0)) * 2,
    rightHwp: Metric.ptToHwp(props.indentRightPt ?? 0) * 2,
    intentHwp: Metric.ptToHwp(props.firstLineIndentPt ?? 0) * 2,
    prevHwp: Metric.ptToHwp(props.spaceBefore ?? 0) * 2,
    nextHwp: Metric.ptToHwp(props.spaceAfter ?? 0) * 2,
    lineSpacing: props.lineHeightFixed
      ? 0
      : (props.lineHeight ? Math.round(props.lineHeight * 100) : 160),
    lineSpacingFixed: props.lineHeightFixed
      ? Math.max(
          Metric.ptToHwp(props.lineHeightFixed),
          Math.ceil(1000 * 1.15),
        ) * 2
      : undefined,
  };
  if (props.listOrd !== undefined) {
    def.listType = props.listOrd ? "DIGIT" : "BULLET";
    def.listLevel = props.listLv ?? 0;
  }
  ctx.paraPrs.push(def);
  ctx.paraPrMap.set(key, id);
  return id;
}

// ─── 이미지 등록 ─────────────────────────────────────────────

function mimeToExt(mime: string): string {
  if (mime.includes("jpeg")) return "jpg";
  if (mime.includes("gif")) return "gif";
  if (mime.includes("bmp")) return "bmp";
  if (mime.includes("wmf")) return "wmf";
  if (mime.includes("emf")) return "emf";
  return "png";
}

function registerImage(img: ImgNode, ctx: HwpxCtx): void {
  if (ctx.imgMap.has(img)) return;
  const ext = mimeToExt(img.mime);
  const id = `BIN${String(ctx.nextBinNum).padStart(4, "0")}`;
  const name = `${id}.${ext}`;
  ctx.nextBinNum++;
  const data = TextKit.base64Decode(img.b64);
  ctx.bins.push({ id, name, data });
  ctx.imgMap.set(img, id);
}

// ─── 스타일 등록 ─────────────────────────────────────────────

const STYLE_NAME_MAP: Record<string, string> = {
  Normal: "바탕글",
  "Heading 1": "개요 1",
  "Heading 2": "개요 2",
  "Heading 3": "개요 3",
  "Heading 4": "개요 4",
  "Heading 5": "개요 5",
  "Heading 6": "개요 6",
  "Body Text": "본문",
};

function registerStyle(
  styleId: string,
  paraPrId: number,
  charPrId: number,
  ctx: HwpxCtx,
): void {
  if (!styleId || ctx.styleIdToHwpxId.has(styleId)) return;
  if (styleId === "Normal" || styleId === "0") {
    ctx.styleIdToHwpxId.set(styleId, 0);
    return;
  }

  const usedIds = new Set(ctx.hwpxStyles.map((s) => s.id));
  const numericId = Number(styleId);
  let hwpxId =
    Number.isInteger(numericId) && numericId > 0 && !usedIds.has(numericId)
      ? numericId
      : nextStyleId(usedIds);
  ctx.styleIdToHwpxId.set(styleId, hwpxId);
  ctx.hwpxStyles.push({
    id: hwpxId,
    name: STYLE_NAME_MAP[styleId] ?? styleId,
    engName: "",
    paraPrIDRef: paraPrId,
    charPrIDRef: charPrId,
  });
}

function nextStyleId(usedIds: Set<number>): number {
  let id = 0;
  while (usedIds.has(id)) id++;
  return id;
}

function materializeContiguousStyles(styles: StyleEntry[]): StyleEntry[] {
  const byId = new Map(styles.map((style) => [style.id, style]));
  const maxId = Math.max(0, ...byId.keys());
  const dense: StyleEntry[] = [];
  for (let id = 0; id <= maxId; id++) {
    dense.push(byId.get(id) ?? {
      id,
      name: `사용자 스타일 ${id}`,
      engName: `User Style ${id}`,
      paraPrIDRef: 0,
      charPrIDRef: 0,
    });
  }
  return dense;
}

function paraStyleKey(props: ParaNode["props"]): string | undefined {
  if (props.hwpStyleId !== undefined) {
    const id = Math.trunc(props.hwpStyleId);
    if (id >= 0 && id <= 255) return String(id);
  }
  return props.styleId;
}

// ─── Pre-scan: 콘텐츠 순회하며 모든 ID 사전 등록 ─────────────

function scanPara(para: ParaNode, ctx: HwpxCtx): void {
  const paraPrId = registerParaPr(para.props, ctx);
  let firstCharPrId = 0;
  let hasFirstSpan = false;

  function scanKids(kids: ParaNode["kids"]): void {
    for (const kid of kids) {
      if (kid.tag === "span") {
        const cId = registerCharPr(kid.props, ctx);
        if (!hasFirstSpan) {
          firstCharPrId = cId;
          hasFirstSpan = true;
        }
      } else if (kid.tag === "img") {
        registerImage(kid, ctx);
      } else if (kid.tag === "link") {
        scanKids((kid as LinkNode).kids as ParaNode["kids"]);
      }
    }
  }
  scanKids(para.kids);
  const styleKey = paraStyleKey(para.props);
  if (styleKey) registerStyle(styleKey, paraPrId, firstCharPrId, ctx);
}

function scanGrid(grid: GridNode, ctx: HwpxCtx): void {
  const defStroke = grid.props.defaultStroke ?? DEFAULT_STROKE;
  // 표 기본 테두리 사전 등록
  ctx.borderFillBank.addUniform(defStroke);
  for (const row of grid.kids) {
    for (const cell of row.kids) {
      ctx.borderFillBank.addFromCellProps(cell.props, defStroke);
      for (const p of cell.kids) {
        if (p.tag === 'grid') scanGrid(p, ctx);
        else scanPara(p, ctx);
      }
    }
  }
}

function scanContent(kids: ContentNode[], ctx: HwpxCtx): void {
  for (const kid of kids) {
    if (kid.tag === "para") scanPara(kid, ctx);
    else if (kid.tag === "grid") scanGrid(kid, ctx);
  }
}

// ─── Encoder 클래스 ──────────────────────────────────────────

export class HwpxEncoder extends BaseEncoder {
  protected getFormat(): string { return "hwpx"; }
  protected getAliases(): string[] { return [HWPX_MIME_TYPE, "application/hwp+zip"]; }

  async encode(doc: DocRoot): Promise<Outcome<Uint8Array>> {
    try {
      const sheets: SheetNode[] = doc.kids.length ? doc.kids : [{ tag: "sheet", dims: A4, kids: [] }];
      const sheet = sheets[0];
      const dims = normalizeDims(sheet?.dims ?? A4);

      const safeML = (dims.ml !== undefined && dims.ml >= 0) ? dims.ml : 70.87;
      const safeMR = (dims.mr !== undefined && dims.mr >= 0) ? dims.mr : 70.87;
      const availableWidth = Math.round(
        Metric.ptToHwp(dims.wPt) -
          Metric.ptToHwp(safeML) -
          Metric.ptToHwp(safeMR),
      );

      // 컨텍스트 초기화
      const ctx: HwpxCtx = {
        fontBank: new LangFontBank(), // ANYTOHWP 방식 언어별 폰트
        borderFillBank: new BorderFillBank(), // 하드코딩 없는 테두리 관리
        charPrs: [],
        charPrMap: new Map(),
        paraPrs: [],
        paraPrMap: new Map(),
        bins: [],
        nextBinNum: 1,
        nextElementId: 10000,
        availableWidth,
        imgMap: new WeakMap(),
        nextZOrder: 0,
        styleIdToHwpxId: new Map(),
        hwpxStyles: [],
      };

      // id=0 기본 charPr/paraPr 등록
      registerCharPr({}, ctx);
      registerParaPr({}, ctx);

      // 바탕글(Normal) 스타일 id=0으로 고정
      ctx.hwpxStyles.push({
        id: 0,
        name: "바탕글",
        engName: "Normal",
        paraPrIDRef: 0,
        charPrIDRef: 0,
      });
      ctx.styleIdToHwpxId.set("Normal", 0);

      // 패스 1: Pre-scan — 모든 charPr/paraPr/이미지/테두리 사전 등록
      for (const section of sheets) {
        scanContent(section.kids, ctx);
        for (const paras of Object.values(section.headers ?? {})) for (const p of paras ?? []) scanPara(p, ctx);
        for (const paras of Object.values(section.footers ?? {})) for (const p of paras ?? []) scanPara(p, ctx);
      }

      // Build every section before writing the shared style/media catalog.
      const sectionData = sheets.map(section => this.stringToBytes(
        buildSectionXml(section, normalizeDims(section.dims), ctx)));
      const headerData = this.stringToBytes(buildHeaderXml(dims, doc.meta, ctx, sheets.length));
      const previewText = sheets.map(extractPreviewText).join("\n");

      const entries: { name: string; data: Uint8Array; mime: string; compression?: 'STORE' | 'DEFLATE' }[] = [
        {
          name: "mimetype",
          data: new TextEncoder().encode(HWPX_MIME_TYPE),
          compression: "STORE",
          mime: "",
        },
        {
          name: "version.xml",
          data: this.stringToBytes(VERSION_XML),
          mime: "application/xml",
        },
        {
          name: "META-INF/container.xml",
          data: this.stringToBytes(CONTAINER_XML),
          mime: "application/xml",
        },
        {
          name: "META-INF/manifest.xml",
          data: this.stringToBytes(MANIFEST_XML),
          mime: "application/xml",
        },
        {
          name: "META-INF/container.rdf",
          data: this.stringToBytes(buildContainerRdf(sheets.length)),
          mime: "application/rdf+xml",
        },
        {
          name: "Contents/content.hpf",
          data: this.stringToBytes(buildContentHpf(ctx, doc.meta, sheets.length)),
          mime: "application/hwpml-package+xml",
        },
        {
          name: "Contents/header.xml",
          data: headerData,
          mime: "application/xml",
        },
        ...sectionData.map((data, index) => ({
          name: `Contents/section${index}.xml`, data, mime: "application/xml",
        })),
        {
          name: "Preview/PrvText.txt",
          data: this.stringToBytes(previewText),
          mime: "text/plain",
        },
        {
          name: "settings.xml",
          data: this.stringToBytes(buildSettingsXml()),
          mime: "application/xml",
        },
      ];

      for (const bin of ctx.bins) {
        const ext = bin.name.split(".").pop()?.toLowerCase() ?? "png";
        const ct =
          ext === "png"
            ? "image/png"
            : ext === "jpg" || ext === "jpeg"
              ? "image/jpeg"
              : ext === "gif"
                ? "image/gif"
                : ext === "wmf"
                  ? "image/x-wmf"
                  : ext === "emf"
                    ? "image/x-emf"
                : "image/bmp";
        entries.push({ name: `BinData/${bin.name}`, data: bin.data, mime: ct });
      }

      return succeed(await this.zip(entries));
    } catch (e: any) {
      return fail(`HWPX 인코딩 오류: ${e?.message ?? String(e)}`);
    }
  }
}

// ─── 상수 XML ────────────────────────────────────────────────

// namespace: 실제 HWP가 기대하는 2011 버전 네임스페이스 사용 (owpml.org/2024는 열리지 않음)
const VERSION_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>` +
  `<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" ` +
  `tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="0" buildNumber="1" ` +
  `os="1" xmlVersion="1.4" application="Hancom Office Hangul" appVersion="11, 0, 0, 8227 WIN32LEWindows_10"/>`;

const CONTAINER_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>` +
  `<ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" ` +
  `xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf">` +
  `<ocf:rootfiles>` +
  `<ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/>` +
  `<ocf:rootfile full-path="Preview/PrvText.txt" media-type="text/plain"/>` +
  `<ocf:rootfile full-path="META-INF/container.rdf" media-type="application/rdf+xml"/>` +
  `</ocf:rootfiles></ocf:container>`;

const MANIFEST_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>` +
  `<odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>`;

function buildContainerRdf(sectionCount: number): string {
  return (
  `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>` +
  `<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">` +
  `<rdf:Description rdf:about=""><pkg:hasPart xmlns:pkg="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" rdf:resource="Contents/header.xml"/></rdf:Description>` +
  `<rdf:Description rdf:about="Contents/header.xml"><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#HeaderFile"/></rdf:Description>` +
  Array.from({ length: sectionCount }, (_, i) =>
    `<rdf:Description rdf:about=""><pkg:hasPart xmlns:pkg="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" rdf:resource="Contents/section${i}.xml"/></rdf:Description>` +
    `<rdf:Description rdf:about="Contents/section${i}.xml"><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#SectionFile"/></rdf:Description>`).join("") +
  `<rdf:Description rdf:about=""><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#Document"/></rdf:Description>` +
  `</rdf:RDF>`);
}

// ─── content.hpf ─────────────────────────────────────────────

function buildContentHpf(ctx: HwpxCtx, meta?: DocMeta, sectionCount = 1): string {
  const title = esc(meta?.title ?? "");
  const creator = esc(meta?.author ?? "text");
  const subject = esc(meta?.subject ?? "text");
  const desc = esc(meta?.desc ?? "text");
  const keyword = esc(meta?.keywords ?? "text");
  const created =
    meta?.created ?? new Date().toISOString().replace(/\.\d{3}Z$/, "Z");
  const modified = meta?.modified ?? created;

  let items =
    `<opf:item id="header"   href="Contents/header.xml"   media-type="application/xml"/>` +
    Array.from({ length: sectionCount }, (_, i) => `<opf:item id="section${i}" href="Contents/section${i}.xml" media-type="application/xml"/>`).join("") +
    `<opf:item id="settings" href="settings.xml"          media-type="application/xml"/>`;

  for (const bin of ctx.bins) {
    const ext = bin.name.split(".").pop()?.toLowerCase() ?? "png";
    const ct =
      ext === "png"
        ? "image/png"
        : ext === "jpg" || ext === "jpeg"
          ? "image/jpeg"
          : ext === "gif"
            ? "image/gif"
            : ext === "wmf"
              ? "image/x-wmf"
              : ext === "emf"
                ? "image/x-emf"
            : "image/bmp";
    items += `<opf:item id="${bin.id}" href="BinData/${bin.name}" media-type="${ct}" isEmbeded="1"/>`;
  }

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>` +
    `<opf:package ${NS} version="" unique-identifier="" id="">` +
    `<opf:metadata>` +
    `<opf:title>${title}</opf:title><opf:language>ko</opf:language>` +
    `<opf:meta name="creator"      content="text">${creator}</opf:meta>` +
    `<opf:meta name="subject"      content="text">${subject}</opf:meta>` +
    `<opf:meta name="description"  content="text">${desc}</opf:meta>` +
    `<opf:meta name="CreatedDate"  content="text">${created}</opf:meta>` +
    `<opf:meta name="ModifiedDate" content="text">${modified}</opf:meta>` +
    `<opf:meta name="keyword"      content="text">${keyword}</opf:meta>` +
    `<opf:meta name="trackchageConfig" content="text">0</opf:meta>` +
    `</opf:metadata>` +
    `<opf:manifest>${items}</opf:manifest>` +
    `<opf:spine><opf:itemref idref="header" linear="yes"/>${Array.from({ length: sectionCount }, (_, i) => `<opf:itemref idref="section${i}" linear="yes"/>`).join("")}</opf:spine>` +
    `</opf:package>`
  );
}

// ─── settings.xml ────────────────────────────────────────────

function buildSettingsXml(): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>` +
    `<ha:HWPApplicationSetting xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" ` +
    `xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">` +
    `<ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/>` +
    `<config:config-item-set name="PrintInfo">` +
    `<config:config-item name="PrintAutoFootNote" type="boolean">false</config:config-item>` +
    `<config:config-item name="PrintAutoHeadNote" type="boolean">false</config:config-item>` +
    `<config:config-item name="PrintMethod" type="short">4</config:config-item>` +
    `<config:config-item name="OverlapSize" type="short">0</config:config-item>` +
    `<config:config-item name="PrintCropMark" type="short">0</config:config-item>` +
    `<config:config-item name="BinderHoleType" type="short">0</config:config-item>` +
    `<config:config-item name="ZoomX" type="short">100</config:config-item>` +
    `<config:config-item name="ZoomY" type="short">100</config:config-item>` +
    `</config:config-item-set>` +
    `</ha:HWPApplicationSetting>`
  );
}

function buildNumberingsXml(): string {
  return (
    `<hh:numberings itemCnt="1">` +
    `<hh:numbering id="1" start="0">` +
    `<hh:paraHead start="1" level="1" align="LEFT" ` +
    `useInstWidth="1" autoIndent="0" widthAdjust="0" ` +
    `textOffsetType="PERCENT" textOffset="50" ` +
    `numFormat="DIGIT" charPrIDRef="0" checkable="0">^1.</hh:paraHead>` +
    `</hh:numbering></hh:numberings>`
  );
}

function buildBulletsXml(): string {
  return (
    `<hh:bullets itemCnt="1">` +
    `<hh:bullet id="1" char="&#x2022;" useImage="0">` +
    `<hh:paraHead level="0" align="LEFT" useInstWidth="0" autoIndent="1" widthAdjust="0" ` +
    `textOffsetType="PERCENT" textOffset="50" numFormat="DIGIT" charPrIDRef="0" checkable="0"/>` +
    `</hh:bullet></hh:bullets>`
  );
}

// ─── header.xml ──────────────────────────────────────────────

/**
 * Contents/header.xml 용 전역 구역 설정 리스트(secPrList)를 생성합니다.
 */
/** Convert body-edge distances to Hancom's outer margins and header/footer areas. */
function hancomVerticalMargins(dims: PageDims) {
  const top = Math.max(0, Math.min(dims.mt, dims.headerPt ?? dims.mt));
  const bottom = Math.max(0, Math.min(dims.mb, dims.footerPt ?? dims.mb));
  return { top: Metric.ptToHwp(top), bottom: Metric.ptToHwp(bottom),
    header: Metric.ptToHwp(Math.max(0, dims.mt - top)),
    footer: Metric.ptToHwp(Math.max(0, dims.mb - bottom)) };
}

function buildHeaderSecPrListXml(dims: PageDims): string {
  const wHwp = Metric.ptToHwp(dims.wPt);
  const hHwp = Metric.ptToHwp(dims.hPt);
  const ml = Metric.ptToHwp(dims.ml);
  const mr = Metric.ptToHwp(dims.mr);
  const { top: mt, bottom: mb, header: headerZone, footer: footerZone } = hancomVerticalMargins(dims);

  const pageBorderFill =
    `<hh:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">` +
    `<hh:offset left="1417" right="1417" top="1417" bottom="1417"/>` +
    `</hh:pageBorderFill>` +
    `<hh:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">` +
    `<hh:offset left="1417" right="1417" top="1417" bottom="1417"/>` +
    `</hh:pageBorderFill>` +
    `<hh:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">` +
    `<hh:offset left="1417" right="1417" top="1417" bottom="1417"/>` +
    `</hh:pageBorderFill>`;

  return (
    `<hh:secPrList itemCnt="1">` +
    `<hh:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="0" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0">` +
    `<hh:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>` +
    `<hh:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>` +
    `<hh:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/>` +
    `<hh:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>` +
    `<hh:pagePr landscape="WIDELY" width="${wHwp}" height="${hHwp}" gutterType="LEFT_ONLY">` +
    `<hh:margin header="${headerZone}" footer="${footerZone}" gutter="0" left="${ml}" right="${mr}" top="${mt}" bottom="${mb}"/>` +
    `</hh:pagePr>` +
    `<hh:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/>` +
    `<hh:footNotePr><hh:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supscript="1"/>` +
    `<hh:noteLine length="-1" type="SOLID" width="0.25 mm" color="#000000"/>` +
    `<hh:noteSpacing betweenNotes="283" belowLine="0" aboveLine="1000"/>` +
    `<hh:numbering type="CONTINUOUS" newNum="1"/>` +
    `<hh:placement place="EACH_COLUMN" beneathText="0"/>` +
    `</hh:footNotePr>` +
    `<hh:endNotePr><hh:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supscript="1"/>` +
    `<hh:noteLine length="-1" type="SOLID" width="0.25 mm" color="#000000"/>` +
    `<hh:noteSpacing betweenNotes="0" belowLine="0" aboveLine="1000"/>` +
    `<hh:numbering type="CONTINUOUS" newNum="1"/>` +
    `<hh:placement place="END_OF_DOCUMENT" beneathText="0"/>` +
    `</hh:endNotePr>` +
    pageBorderFill +
    `</hh:secPr>` +
    `</hh:secPrList>`
  );
}

function buildHeaderXml(dims: PageDims, meta: DocMeta, ctx: HwpxCtx, sectionCount = 1): string {
  // 언어별 폰트 (LangFontBank → XML)
  const fontFacesXml = ctx.fontBank.toXml();

  // charPr 목록 — 언어별 폰트 ID를 fontRef에 반영 (ANYTOHWP 핵심 개선)
  let charPrXml = "";
  for (const cp of ctx.charPrs) {
    const bold = cp.bold ? "<hh:bold/>" : "";
    const italic = cp.italic ? "<hh:italic/>" : "";
    const hid = cp.hangulId;
    const lid = cp.latinId;
    const shadeColor = cp.bg ? (cp.bg.startsWith("#") ? cp.bg : `#${cp.bg}`) : "none";
    charPrXml +=
      `<hh:charPr id="${cp.id}" height="${cp.height}" textColor="${cp.textColor}" ` +
      `shadeColor="${shadeColor}" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="1">` +
      `<hh:fontRef hangul="${hid}" latin="${lid}" hanja="${cp.hanjaId}" japanese="${cp.japaneseId}" ` +
      `other="${cp.otherId}" symbol="${cp.symbolId}" user="${cp.userId}"/>` +
      `<hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>` +
      `<hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>` +
      `<hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>` +
      `<hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>` +
      bold +
      italic +
      `<hh:underline type="${cp.underline}" shape="SOLID" color="#000000"/>` +
      `<hh:strikeout shape="${cp.strikeout}" color="#000000"/>` +
      `<hh:outline type="NONE"/>` +
      `<hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/>` +
      `</hh:charPr>`;
  }

  // paraPr 목록 (동적 정렬 및 줄바꿈 적용)
  let paraPrXml = "";
  for (const pp of ctx.paraPrs) {
    const ver = pp.verAlign ?? "BASELINE";
    const wrap = pp.lineWrap ?? "BREAK";
    const lsType = pp.lineSpacingFixed !== undefined ? "AT_LEAST" : "PERCENT";
    const lsValue = pp.lineSpacingFixed !== undefined ? pp.lineSpacingFixed : pp.lineSpacing;
    paraPrXml +=
      `<hh:paraPr id="${pp.id}" tabPrIDRef="0" condense="0" fontLineHeight="0" snapToGrid="0" suppressLineNumbers="0" checked="0">` +
      `<hh:align horizontal="${pp.align}" vertical="${ver}"/>` +
      `<hh:heading type="NONE" idRef="0" level="0"/>` +
      `<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="KEEP_WORD" widowOrphan="${pp.widowControl ? 1 : 0}" keepWithNext="${pp.keepWithNext ? 1 : 0}" keepLines="${pp.keepLines ? 1 : 0}" pageBreakBefore="${pp.pageBreakBefore ? 1 : 0}" lineWrap="${wrap}"/>` +
      `<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>` +
      `<hp:switch>` +
      `<hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar">` +
      `<hh:margin>` +
      `<hc:intent value="${pp.intentHwp}" unit="HWPUNIT"/>` +
      `<hc:left value="${pp.leftHwp}" unit="HWPUNIT"/>` +
      `<hc:right value="${pp.rightHwp}" unit="HWPUNIT"/>` +
      `<hc:prev value="${pp.prevHwp}" unit="HWPUNIT"/>` +
      `<hc:next value="${pp.nextHwp}" unit="HWPUNIT"/>` +
      `</hh:margin>` +
      `<hh:lineSpacing type="${lsType}" value="${lsValue}" unit="HWPUNIT"/>` +
      `</hp:case>` +
      `<hp:default>` +
      `<hh:margin>` +
      `<hc:intent value="${pp.intentHwp}" unit="HWPUNIT"/>` +
      `<hc:left value="${pp.leftHwp}" unit="HWPUNIT"/>` +
      `<hc:right value="${pp.rightHwp}" unit="HWPUNIT"/>` +
      `<hc:prev value="${pp.prevHwp}" unit="HWPUNIT"/>` +
      `<hc:next value="${pp.nextHwp}" unit="HWPUNIT"/>` +
      `</hh:margin>` +
      `<hh:lineSpacing type="${lsType}" value="${lsValue}" unit="HWPUNIT"/>` +
      `</hp:default>` +
      `</hp:switch>` +
      `<hh:border borderFillIDRef="1" offsetLeft="0" offsetRight="0" offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>` +
      `</hh:paraPr>`;
  }

  // borderFill 목록 (BorderFillBank → XML)
  const borderFillXml = ctx.borderFillBank.toXml();

  // 스타일 목록
  const denseStyles = materializeContiguousStyles(ctx.hwpxStyles);
  const stylesXml =
    `<hh:styles itemCnt="${denseStyles.length}">` +
    denseStyles
      .map(
        (s) =>
          `<hh:style id="${s.id}" type="PARA" name="${esc(s.name)}" engName="${esc(s.engName)}" ` +
          `paraPrIDRef="${s.paraPrIDRef}" charPrIDRef="${s.charPrIDRef}" nextStyleIDRef="0" langID="1042" lockForm="0"/>`,
      )
      .join("") +
    `</hh:styles>`;

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>` +
    `<hh:head ${NS} version="1.4" secCnt="${sectionCount}">` +
    `<hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>` +
    `<hh:refList>` +
    fontFacesXml +
    borderFillXml +
    `<hh:charProperties itemCnt="${ctx.charPrs.length}">${charPrXml}</hh:charProperties>` +
    `<hh:tabProperties itemCnt="1"><hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/></hh:tabProperties>` +
    buildNumberingsXml() +
    buildBulletsXml() +
    `<hh:paraProperties itemCnt="${ctx.paraPrs.length}">${paraPrXml}</hh:paraProperties>` +
    stylesXml +
    `</hh:refList>` +
    `<hh:compatibleDocument targetProgram="HWP201X"><hh:layoutCompatibility/></hh:compatibleDocument>` +
    `<hh:docOption><hh:linkinfo path="" pageInherit="0" footnoteInherit="0"/></hh:docOption>` +
    `<hh:trackchageConfig flags="56"/>` +
    `</hh:head>`
  );
}

// ─── section0.xml ────────────────────────────────────────────

function buildHeaderFooterRunXml(
  sheet: SheetNode,
  dims: PageDims,
  ctx: HwpxCtx,
): string {
  const headers = sheet.headers || {};
  const footers = sheet.footers || {};
  const hasAny = Object.keys(headers).length > 0 || Object.keys(footers).length > 0;
  if (!hasAny) return "";

  const availW = ctx.availableWidth;
  const mtHwp = Metric.ptToHwp(dims.mt);
  const mbHwp = Metric.ptToHwp(dims.mb);
  const headerZoneH = hancomVerticalMargins(dims).header;
  const footerZoneH = hancomVerticalMargins(dims).footer;

  let inner = "";

  // 1. 첫 페이지 숨김 설정 (first 헤더/푸터가 있으면 활성화)
  const hideFirst = !!(headers.first || footers.first);
  inner += `<hp:ctrl><hp:pageHiding hideHeader="${hideFirst ? 1 : 0}" hideFooter="${hideFirst ? 1 : 0}" hideMasterPage="0" hideBorder="0" hideFill="0" hidePageNum="0"/></hp:ctrl>`;

  // 2. 헤더들 생성
  for (const [type, paras] of Object.entries(headers)) {
    if (!Array.isArray(paras) || paras.length === 0) continue;
    const applyPageType = type === "even" ? "EVEN" : (type === "default" || type === "first" ? "BOTH" : "ODD");
    const savedId = ctx.nextElementId;
    ctx.nextElementId = 0;
    const parasXml = paras.map((p) => encodeParaPositioned(p, ctx, 0, "", availW).xml).join("");
    ctx.nextElementId = savedId;
    inner +=
      `<hp:ctrl>` +
      `<hp:header id="1" applyPageType="${applyPageType}">` +
      `<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" ` +
      `linkListIDRef="0" linkListNextIDRef="0" textWidth="${availW}" textHeight="${headerZoneH}" ` +
      `hasTextRef="0" hasNumRef="0">` +
      parasXml +
      `</hp:subList>` +
      `</hp:header>` +
      `</hp:ctrl>`;
  }

  // 3. 푸터들 생성
  for (const [type, paras] of Object.entries(footers)) {
    if (!Array.isArray(paras) || paras.length === 0) continue;
    const applyPageType = type === "even" ? "EVEN" : (type === "default" || type === "first" ? "BOTH" : "ODD");
    const savedId = ctx.nextElementId;
    ctx.nextElementId = 0;
    const parasXml = paras.map((p) => encodeParaPositioned(p, ctx, 0, "", availW).xml).join("");
    ctx.nextElementId = savedId;
    inner +=
      `<hp:ctrl>` +
      `<hp:footer id="2" applyPageType="${applyPageType}">` +
      `<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="BOTTOM" ` +
      `linkListIDRef="0" linkListNextIDRef="0" textWidth="${availW}" textHeight="${footerZoneH}" ` +
      `hasTextRef="0" hasNumRef="0">` +
      parasXml +
      `</hp:subList>` +
      `</hp:footer>` +
      `</hp:ctrl>`;
  }

  return `<hp:run charPrIDRef="0" charTcId="0">${inner}</hp:run>`;
}

function buildSectionXml(
  sheet: SheetNode | undefined,
  dims: PageDims,
  ctx: HwpxCtx,
): string {
  const secPrXml = buildSecPrXml(dims);
  const sectionControlRunXml =
    `<hp:run charPrIDRef="0" charTcId="0">` +
    secPrXml +
    `<hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/></hp:ctrl>` +
    `</hp:run>`;
  const kids = sheet?.kids ?? [];
  const hfRunXml = sheet ? buildHeaderFooterRunXml(sheet, dims, ctx) : "";

  // 가용 너비 계산 (HWPUNIT)
  const availWidth = Math.max(
    1000,
    Metric.ptToHwp(dims.wPt) - Metric.ptToHwp(dims.ml) - Metric.ptToHwp(dims.mr),
  );
  const bodyHeight = Math.max(
    1000,
    Metric.ptToHwp(dims.hPt) - Metric.ptToHwp(dims.mt) - Metric.ptToHwp(dims.mb),
  );
  ctx.availableWidth = availWidth;

  let contentXml = "";
  let vertPos = 0;
  let pageFirst = true;

  for (let i = 0; i < kids.length; i++) {
    const kid = kids[i];
    const isFirst = i === 0;
    const curSecPr = isFirst ? sectionControlRunXml : "";
    const curHfRun = isFirst ? hfRunXml : "";

    if (kid.tag === "para") {
      if (paraHasPageBreak(kid)) {
        vertPos = 0;
        pageFirst = true;
      }
      const { xml, nextVertPos, hasPageBreak } = encodeParaPositioned(
        kid,
        ctx,
        vertPos,
        curSecPr,
        availWidth,
        curHfRun,
        pageFirst,
      );
      contentXml += xml;
      if (nextVertPos >= bodyHeight) {
        vertPos = 0;
        pageFirst = true;
      } else {
        vertPos = nextVertPos;
        pageFirst = false;
      }
    } else if (kid.tag === "grid") {
      const { xml, nextVertPos, hasPageBreak } = encodeGridPositioned(
        kid,
        ctx,
        vertPos,
        curSecPr,
        curHfRun,
        pageFirst,
      );
      contentXml += xml;
      if (nextVertPos >= bodyHeight) {
        vertPos = 0;
        pageFirst = true;
      } else {
        vertPos = nextVertPos;
        pageFirst = false;
      }
    }
  }

  if (kids.length === 0) {
    // 빈 문서 — 최소 단락 1개 필수
    const fs = 1000;
    const vs = 1600;
    const { xml: linesegXml } = buildLinesegarray(
      " ", 0, fs, vs / (fs / 100), availWidth, undefined, { pageFirst: true },
    );
    contentXml =
      `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0" paraTcId="0">` +
      sectionControlRunXml +
      hfRunXml +
      `<hp:run charPrIDRef="0" charTcId="0"><hp:t xml:space="preserve"> </hp:t></hp:run>` +
      linesegXml +
      `</hp:p>`;
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec ${NS}>${contentXml}</hs:sec>`;
}

function buildSecPrXml(dims: PageDims): string {
  const wHwp = Metric.ptToHwp(dims.wPt);
  const hHwp = Metric.ptToHwp(dims.hPt);
  const ml = Metric.ptToHwp(dims.ml);
  const mr = Metric.ptToHwp(dims.mr);
  const { top: mt, bottom: mb, header: headerZone, footer: footerZone } = hancomVerticalMargins(dims);

  const pageBorderFill =
    `<hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">` +
    `<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>` +
    `</hp:pageBorderFill>` +
    `<hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">` +
    `<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>` +
    `</hp:pageBorderFill>` +
    `<hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER">` +
    `<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>` +
    `</hp:pageBorderFill>`;

  return (
    `<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="0" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0">` +
    `<hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>` +
    `<hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>` +
    `<hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/>` +
    `<hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>` +
    `<hp:pagePr landscape="WIDELY" width="${wHwp}" height="${hHwp}" gutterType="LEFT_ONLY">` +
    `<hp:margin header="${headerZone}" footer="${footerZone}" gutter="0" left="${ml}" right="${mr}" top="${mt}" bottom="${mb}"/>` +
    `</hp:pagePr>` +
    `<hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supscript="1"/>` +
    `<hp:noteLine length="-1" type="SOLID" width="0.25 mm" color="#000000"/>` +
    `<hp:noteSpacing betweenNotes="283" belowLine="0" aboveLine="1000"/>` +
    `<hp:numbering type="CONTINUOUS" newNum="1"/>` +
    `<hp:placement place="EACH_COLUMN" beneathText="0"/>` +
    `</hp:footNotePr>` +
    `<hp:endNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supscript="1"/>` +
    `<hp:noteLine length="-1" type="SOLID" width="0.25 mm" color="#000000"/>` +
    `<hp:noteSpacing betweenNotes="0" belowLine="0" aboveLine="1000"/>` +
    `<hp:numbering type="CONTINUOUS" newNum="1"/>` +
    `<hp:placement place="END_OF_DOCUMENT" beneathText="0"/>` +
    `</hp:endNotePr>` +
    pageBorderFill +
    `</hp:secPr>`
  );
}

// ─── 줄 정보 XML (linesegarray) ──────────────────────────────
// 가이드 준수: 실제 시각적 줄 단위로 lineseg 생성

interface LinesegLayout {
  textHeight?: number;
  firstHorzPos?: number;
  restHorzPos?: number;
  rightMargin?: number;
  indentFirst?: boolean;
  pageFirst?: boolean;
}

function buildLinesegarray(
  text: string,
  vertPosStart: number,
  fontSize: number,
  lineSpacingPct: number,
  horzSize: number,
  lineHeightHwp?: number,
  layout: LinesegLayout = {},
): { xml: string; totalHeight: number } {
  const textHeight = Math.max(fontSize, layout.textHeight ?? fontSize);
  const lineAdvance = Math.max(
    textHeight,
    lineHeightHwp ?? Math.round((fontSize * Math.max(100, lineSpacingPct)) / 100),
  );
  const spacing = Math.max(0, lineAdvance - textHeight);
  const baseline = Math.round(textHeight * 0.85);
  const firstHorzPos = Math.max(0, layout.firstHorzPos ?? 0);
  const restHorzPos = Math.max(0, layout.restHorzPos ?? firstHorzPos);
  const rightMargin = Math.max(0, layout.rightMargin ?? 0);
  const lineHorzPos = (index: number) => index === 0 ? firstHorzPos : restHorzPos;
  const lineHorzSize = (index: number) =>
    Math.max(100, horzSize - lineHorzPos(index) - rightMargin);

  if (text.length === 0) {
    const xml = `<hp:linesegarray>` +
      `<hp:lineseg textpos="0" vertpos="${vertPosStart}" vertsize="${textHeight}" ` +
      `textheight="${textHeight}" baseline="${baseline}" spacing="${spacing}" ` +
      `horzpos="${firstHorzPos}" horzsize="${lineHorzSize(0)}" ` +
      `flags="${LINESEG_FLAGS |
        (layout.indentFirst ? LINESEG_FLAG_INDENT : 0) |
        (layout.pageFirst ? LINESEG_FLAG_PAGE_FIRST | LINESEG_FLAG_COLUMN_FIRST : 0)}"/>` +
      `</hp:linesegarray>`;
    return { xml, totalHeight: lineAdvance };
  }

  // 문자 단위 정밀 가로 폭 계산 및 자동 줄바꿈 알고리즘 (개행 문자 지원)
  const lines: { startPos: number; width: number }[] = [];
  let currentLineWidth = 0;
  let lineStartIdx = 0;

  for (let i = 0; i < text.length; i++) {
    const charCode = text.charCodeAt(i);

    // 개행 문자 (\n 또는 \r) 감지 시, 현재 줄 세그먼트를 무조건 마감하고 새로운 줄 시작
    if (charCode === 10 || charCode === 13) {
      lines.push({ startPos: lineStartIdx, width: currentLineWidth });
      if (charCode === 13 && text.charCodeAt(i + 1) === 10) i++;
      lineStartIdx = i + 1;
      currentLineWidth = 0;
      continue;
    }

    let charW = fontSize * 0.55; // 기본값 (영문 소문자, 숫자 등)

    if (charCode >= 0xac00 && charCode <= 0xd7a3) {
      charW = fontSize; // 한글
    } else if (charCode >= 0x3130 && charCode <= 0x318f) {
      charW = fontSize; // 한글 자모
    } else if (charCode >= 0x4e00 && charCode <= 0x9fff) {
      charW = fontSize; // 한자
    } else if (charCode >= 65 && charCode <= 90) {
      charW = fontSize * 0.65; // 영문 대문자
    } else if (charCode === 32) {
      charW = fontSize * 0.32; // 공백
    } else if (charCode > 255) {
      charW = fontSize; // 기타 전각 문자
    } else {
      charW = fontSize * 0.42; // 기타 특수기호
    }

    if (currentLineWidth + charW > lineHorzSize(lines.length) && i > lineStartIdx) {
      lines.push({ startPos: lineStartIdx, width: currentLineWidth });
      lineStartIdx = i;
      currentLineWidth = charW;
    } else {
      currentLineWidth += charW;
    }
  }
  lines.push({ startPos: lineStartIdx, width: currentLineWidth });

  const lineCount = lines.length;
  const linesegParts: string[] = [];

  for (let i = 0; i < lineCount; i++) {
    const flags = LINESEG_FLAGS |
      (i === 0 && layout.indentFirst ? LINESEG_FLAG_INDENT : 0) |
      (i === 0 && layout.pageFirst
        ? LINESEG_FLAG_PAGE_FIRST | LINESEG_FLAG_COLUMN_FIRST
        : 0);
    const textpos = lines[i].startPos;
    linesegParts.push(
      `<hp:lineseg textpos="${textpos}" ` +
      `vertpos="${vertPosStart + i * lineAdvance}" ` +
      `vertsize="${textHeight}" textheight="${textHeight}" ` +
      `baseline="${baseline}" spacing="${spacing}" ` +
      `horzpos="${lineHorzPos(i)}" horzsize="${lineHorzSize(i)}" ` +
      `flags="${flags}"/>`,
    );
  }

  return {
    xml: `<hp:linesegarray>${linesegParts.join("")}</hp:linesegarray>`,
    totalHeight: lineCount * lineAdvance,
  };
}

/** 단락에서 순수 텍스트 추출 (줄바꿈 계산용) */
function extractParaText(para: ParaNode): string {
  let text = "";
  const walk = (kids: any[]) => {
    for (const k of kids) {
      if (k.tag === "span") {
        for (const c of k.kids) {
          if (c.tag === "txt") {
            text += c.content;
          } else if (c.tag === "br") {
            text += "\n";
          }
        }
      } else if (k.tag === "link") {
        walk(k.kids);
      }
    }
  };
  walk(para.kids);
  return text;
}

function paraHasPageBreak(para: ParaNode): boolean {
  const visit = (kids: ParaNode["kids"]): boolean =>
    kids.some((kid) => {
      if (kid.tag === "span") return kid.kids.some((child) => child.tag === "pb");
      if (kid.tag === "link") {
        return visit((kid as LinkNode).kids as ParaNode["kids"]);
      }
      return false;
    });
  return visit(para.kids);
}

function fontSizeForPara(para: ParaNode, ctx: HwpxCtx): number {
  let maxSize = 1000; // 기본 10pt
  const visit = (kids: any[]) => {
    for (const kid of kids ?? []) {
      if (kid.tag === "span") {
        const id = ctx.charPrMap.get(charPrKey(kid.props));
        if (id !== undefined && ctx.charPrs[id]) {
          maxSize = Math.max(maxSize, ctx.charPrs[id].height);
        }
      } else if (kid.tag === "link") {
        visit(kid.kids ?? []);
      }
    }
  };
  visit(para.kids as any[]);
  return maxSize;
}

function inlineObjectHeightForPara(para: ParaNode, ctx: HwpxCtx): number {
  let maxHeight = 0;
  for (const kid of para.kids) {
    if (kid.tag !== "img" || (kid.layout && kid.layout.wrap !== "inline")) continue;
    const dims = getImageDisplayDims(kid, ctx);
    maxHeight = Math.max(maxHeight, dims.h);
  }
  return maxHeight;
}

// ─── 단락 인코딩 ─────────────────────────────────────────────

function encodeParaPositioned(
  para: ParaNode,
  ctx: HwpxCtx,
  vertPos: number,
  secPr = "",
  availWidth?: number,
  hfRun = "",
  pageFirst = false,
): { xml: string; nextVertPos: number; hasPageBreak: boolean } {
  // ✅ 표(Grid)를 포함하는 단락인지 확인
  const gridKid = para.kids.find((k): k is GridNode => k.tag === "grid");
  if (gridKid) {
    return encodeTablePara(para, gridKid, ctx, vertPos, secPr, hfRun, pageFirst);
  }

  const paraPrId = ctx.paraPrMap.get(paraPrKey(para.props)) ?? 0;
  const styleKey = paraStyleKey(para.props);
  const styleIDRef = styleKey ? (ctx.styleIdToHwpxId.get(styleKey) ?? 0) : 0;
  const fontSize = fontSizeForPara(para, ctx);
  const paraPr = ctx.paraPrs[paraPrId];
  const lineSpacing = paraPr?.lineSpacing ?? 160;
  const lineHeightHwp = paraPr?.lineSpacingFixed !== undefined
    ? Math.max(
        paraShapeHwpToLayoutHwp(paraPr.lineSpacingFixed),
        Math.ceil(fontSize * 1.15),
      )
    : Math.max(fontSize, Math.round((fontSize * Math.max(100, lineSpacing)) / 100));
  const textHeight = Math.max(fontSize, inlineObjectHeightForPara(para, ctx));
  const effectiveLineHeight = textHeight + Math.max(0, lineHeightHwp - fontSize);
  const spacing = Math.max(0, effectiveLineHeight - textHeight);
  let vertSize = effectiveLineHeight;
  const horzSize = availWidth ?? ctx.availableWidth;

  // 코드 블록 감지 (Courier 폰트 또는 styleId "code")
  const isCourierFont = (kids: ParaNode["kids"]): boolean =>
    kids.some(
      (k) =>
        (k.tag === "span" && k.props.font?.toLowerCase().includes("courier")) ||
        (k.tag === "link" &&
          isCourierFont((k as LinkNode).kids as ParaNode["kids"])),
    );
  const isCode =
    availWidth === undefined &&
    (para.props.styleId?.toLowerCase().includes("code") ||
      isCourierFont(para.kids));

  if (isCode)
    return encodeCodeBlockPositioned(
      para,
      ctx,
      vertPos,
      secPr,
      fontSize,
      spacing,
      vertSize,
      pageFirst,
    );

  let runsXml = encodeParaKids(para.kids, ctx);
  if (!runsXml) runsXml = `<hp:run charPrIDRef="0" charTcId="0"><hp:t xml:space="preserve"> </hp:t></hp:run>`;

  const paraText = extractParaText(para);
  const paraStart =
    vertPos + Math.max(0, paraShapeHwpToLayoutHwp(paraPr?.prevHwp ?? 0));
  const firstHorzPos = Math.max(
    0,
    paraShapeHwpToLayoutHwp(
      (paraPr?.leftHwp ?? 0) + Math.max(0, paraPr?.intentHwp ?? 0),
    ),
  );
  const restHorzPos = Math.max(
    0,
    paraShapeHwpToLayoutHwp((paraPr?.leftHwp ?? 0) - Math.min(0, paraPr?.intentHwp ?? 0)),
  );
  const { xml: linesegXml, totalHeight } = buildLinesegarray(
    paraText,
    paraStart,
    fontSize,
    lineSpacing,
    horzSize,
    effectiveLineHeight,
    {
      textHeight,
      firstHorzPos,
      restHorzPos,
      rightMargin: Math.max(
        0,
        paraShapeHwpToLayoutHwp(paraPr?.rightHwp ?? 0),
      ),
      indentFirst: (paraPr?.intentHwp ?? 0) !== 0,
      pageFirst,
    },
  );

  const hasPageBreak = paraHasPageBreak(para);

  const xml =
    `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="${paraPrId}" styleIDRef="${styleIDRef}" ` +
    `pageBreak="${hasPageBreak ? 1 : 0}" columnBreak="0" merged="0" paraTcId="0">` +
    secPr +
    hfRun +
    runsXml +
    linesegXml +
    `</hp:p>`;

  return {
    xml,
    nextVertPos:
      paraStart +
      totalHeight +
      Math.max(0, paraShapeHwpToLayoutHwp(paraPr?.nextHwp ?? 0)),
    hasPageBreak,
  };
}

/** ✅ 가이드 준수: 표를 포함하는 단락 인코딩 */
function encodeTablePara(
  para: ParaNode,
  grid: GridNode,
  ctx: HwpxCtx,
  vertPos: number,
  secPr: string,
  hfRun: string,
  pageFirst: boolean,
): { xml: string; nextVertPos: number; hasPageBreak: boolean } {
  const paraPrId = ctx.paraPrMap.get(paraPrKey(para.props)) ?? 0;
  
  // 표 알맹이 생성 (기존 로직 재사용)
  const { xml: gridXml, height: tblHeight } = buildGridXml(grid, ctx);
  
  // A table occupies one cached line whose text box is the table height.
  const totalHeight = Math.max(1600, tblHeight);
  const baseline = Math.round(totalHeight * 0.85);

  const linesegXml =
    `<hp:linesegarray>` +
    `<hp:lineseg textpos="0" vertpos="${vertPos}" vertsize="${totalHeight}" ` +
    `textheight="${totalHeight}" baseline="${baseline}" spacing="0" ` +
    `horzpos="0" horzsize="${ctx.availableWidth}" ` +
    `flags="${LINESEG_FLAGS | (pageFirst ? LINESEG_FLAG_PAGE_FIRST | LINESEG_FLAG_COLUMN_FIRST : 0)}"/>` +
    `</hp:linesegarray>`;

  const hasPageBreak = paraHasPageBreak(para);

  const xml =
    `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="${paraPrId}" styleIDRef="0" ` +
    `pageBreak="${hasPageBreak ? 1 : 0}" columnBreak="0" merged="0" paraTcId="0">` +
    secPr +
    `<hp:run charPrIDRef="0" charTcId="0">` +
    gridXml +
    `</hp:run>` +
    hfRun +
    linesegXml +
    `</hp:p>`;

  return { xml, nextVertPos: vertPos + totalHeight, hasPageBreak };
}

function encodeCodeBlockPositioned(
  para: ParaNode,
  ctx: HwpxCtx,
  vertPos: number,
  secPr: string,
  fontSize: number,
  spacing: number,
  vertSize: number,
  pageFirst: boolean,
): { xml: string; nextVertPos: number; hasPageBreak: boolean } {
  const codeBfId = ctx.borderFillBank.addUniform(
    { kind: "solid", pt: 0.5, color: "aaaaaa" },
    "f4f4f4",
  );
  const cellW = ctx.availableWidth;
  const innerW = Math.max(cellW - 510, 100);
  const subListId = ctx.nextElementId++;
  const { xml: innerXml } = encodeParaPositioned(para, ctx, 0, "", innerW);

  const paraText = extractParaText(para);
  const { xml: linesegXml, totalHeight } = buildLinesegarray(
    paraText,
    vertPos,
    fontSize,
    160, // 코드 블록 기본 줄간격 160%
    ctx.availableWidth,
    undefined,
    { pageFirst },
  );

  const xml =
    `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0" paraTcId="0">` +
    secPr +
    `<hp:run charPrIDRef="0" charTcId="0">` +
    `<hp:tbl id="${ctx.nextElementId++}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="0" rowCnt="1" colCnt="1" cellSpacing="0" borderFillIDRef="${codeBfId}" noAdjust="0">` +
    `<hp:sz width="${cellW}" widthRelTo="ABSOLUTE" height="0" heightRelTo="ABSOLUTE" protect="0"/>` +
    `<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>` +
    `<hp:outMargin left="138" right="138" top="138" bottom="138"/>` +
    `<hp:inMargin left="138" right="138" top="138" bottom="138"/>` +
    `<hp:tr><hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="${codeBfId}">` +
    `<hp:subList id="${subListId}" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">` +
    innerXml +
    `</hp:subList>` +
    `<hp:cellAddr colAddr="0" rowAddr="0"/>` +
    `<hp:cellSpan colSpan="1" rowSpan="1"/>` +
    `<hp:cellSz width="${cellW}" height="0"/>` +
    `<hp:cellMargin left="283" right="283" top="141" bottom="141"/>` +
    `</hp:tc></hp:tr></hp:tbl><hp:t xml:space="preserve"> </hp:t></hp:run>` +
    linesegXml +
    `</hp:p>`;

  return { xml, nextVertPos: vertPos + totalHeight, hasPageBreak: false };
}

function encodeParaKids(kids: ParaNode["kids"], ctx: HwpxCtx): string {
  let xml = "";
  let currentRunCharPrId: number | null = null;
  let currentRunContent = "";

  const flushRun = () => {
    if (currentRunCharPrId !== null) {
      // 내용이 없더라도 빈 hp:t를 생성하여 '텍스트 없음' 오류 방지
      const content = currentRunContent || `<hp:t xml:space="preserve"> </hp:t>`;
      xml += `<hp:run charPrIDRef="${currentRunCharPrId}" charTcId="0">${content}</hp:run>`;
    }
    currentRunCharPrId = null;
    currentRunContent = "";
  };

  for (const kid of kids) {
    if (kid.tag === "span") {
      const span = kid as SpanNode;
      const charPrId = ctx.charPrMap.get(charPrKey(span.props)) ?? 0;
      
      if (currentRunCharPrId !== null && currentRunCharPrId !== charPrId) {
        flushRun();
      }
      
      currentRunCharPrId = charPrId;
      currentRunContent += encodeRunInner(span);
    } 
    else if (kid.tag === "link") {
      const link = kid as LinkNode;
      // 링크의 첫 번째 span 스타일을 기준으로 함
      let charPrId = 0;
      if (link.kids.length > 0 && link.kids[0].tag === "span") {
        charPrId = ctx.charPrMap.get(charPrKey((link.kids[0] as SpanNode).props)) ?? 0;
      }

      if (currentRunCharPrId !== null && currentRunCharPrId !== charPrId) {
        flushRun();
      }

      currentRunCharPrId = charPrId;
      currentRunContent += encodeLinkInner(link, ctx);
    }
    else if (kid.tag === "img") {
      flushRun();
      xml += encodeImgWrapped(kid, ctx);
    }
  }

  flushRun();
  return xml;
}

/** hp:run 내부의 태그들만 생성 (span용) */
function encodeRunInner(span: SpanNode): string {
  let xml = "";
  for (const kid of span.kids) {
    if (kid.tag === "txt") {
      const raw = kid.content.replace(/__EXT_\d+(?:_W\d+_H\d+)?__/g, "");
      if (!raw) continue;
      // Hancom RunType accepts text; lineBreak/tab belong INSIDE hp:t.
      const lines = TextKit.splitLines(raw);
      for (let li = 0; li < lines.length; li++) {
        if (lines[li] !== "") xml += `<hp:t xml:space="preserve">${esc(lines[li]).replace(/\t/g, '<hp:tab width="4000" leader="0" type="1"/>')}</hp:t>`;
        if (li < lines.length - 1) xml += `<hp:t><hp:lineBreak/></hp:t>`;
      }
    } else if (kid.tag === "br") {
      xml += `<hp:t><hp:lineBreak/></hp:t>`;
    } else if (kid.tag === "pagenum") {
      const fmt = (kid as any).format === "roman" ? "ROMAN_LOWER" 
                : (kid as any).format === "romanCaps" ? "ROMAN_UPPER" : "DIGIT";
      const numType = (kid as any).format === "total" ? "TOTAL_PAGE" : "PAGE";
      xml +=
        `<hp:ctrl><hp:autoNum num="1" numType="${numType}">` +
        `<hp:autoNumFormat type="${fmt}" userChar="" prefixChar="" suffixChar="" supscript="0"/>` +
        `</hp:autoNum></hp:ctrl>`;
    }
  }
  return xml;
}

/** hp:run 내부의 태그들만 생성 (link용) */
function encodeLinkInner(link: LinkNode, ctx: HwpxCtx): string {
  const fieldId = 600000000 + (ctx.nextElementId++ % 100000000);
  const instanceId = 2100000000 + (ctx.nextElementId++ % 100000000);
  const url = link.href;

  let xml = `<hp:ctrl>` +
    `<hp:fieldBegin id="${instanceId}" type="HYPERLINK" name="" editable="0" dirty="1" zorder="-1" fieldid="${fieldId}">` +
    `<hp:parameters cnt="6" name="">` +
    `<hp:integerParam name="Prop">0</hp:integerParam>` +
    `<hp:stringParam name="Command">${esc(url.replace(/:/g, "\\:"))};1;5;-1;</hp:stringParam>` +
    `<hp:stringParam name="Path">${esc(url)}</hp:stringParam>` +
    `<hp:stringParam name="Category">HWPHYPERLINK_TYPE_URL</hp:stringParam>` +
    `<hp:stringParam name="TargetType">HWPHYPERLINK_TARGET_HYPERLINK</hp:stringParam>` +
    `<hp:stringParam name="DocOpenType">HWPHYPERLINK_JUMP_DONTCARE</hp:stringParam>` +
    `</hp:parameters>` +
    `</hp:fieldBegin>` +
    `</hp:ctrl>`;

  for (const kid of link.kids) {
    if (kid.tag === "span") {
      xml += encodeRunInner(kid as SpanNode);
    }
  }

  xml += `<hp:ctrl><hp:fieldEnd beginIDRef="${instanceId}"/></hp:ctrl>`;
  return xml;
}

// ─── 이미지 인코딩 (ANYTOHWP 영감: 픽셀 치수 추출) ──────────

const WRAP_MAP: Record<string, string> = {
  inline: "TOP_AND_BOTTOM",
  square: "SQUARE",
  tight: "BOTH_SIDES",
  through: "BOTH_SIDES",
  none: "FRONT_TEXT",
  behind: "BEHIND_TEXT",
  front: "FRONT_TEXT",
};
const FLOW_MAP: Record<string, string> = {
  inline: "BOTH_SIDES",
  square: "LARGEST_ONLY",
  tight: "BOTH_SIDES",
  through: "BOTH_SIDES",
  none: "BOTH_SIDES",
  behind: "BOTH_SIDES",
  front: "BOTH_SIDES",
};

function getImageSourceDims(img: ImgNode): { w: number; h: number } {
  const pixelDims = img.b64 ? readPixelDims(img.b64, img.mime) : null;
  if (pixelDims && pixelDims.w > 0 && pixelDims.h > 0) {
    return {
      w: Metric.ptToHwp((pixelDims.w * 72) / 96),
      h: Metric.ptToHwp((pixelDims.h * 72) / 96),
    };
  }
  return {
    w: Math.max(1, Metric.ptToHwp(img.w || 1)),
    h: Math.max(1, Metric.ptToHwp(img.h || 1)),
  };
}

function getImageDisplayDims(img: ImgNode, ctx: HwpxCtx): { w: number; h: number } {
  const source = getImageSourceDims(img);
  let w = img.w > 0 ? Metric.ptToHwp(img.w) : source.w;
  let h = img.h > 0 ? Metric.ptToHwp(img.h) : source.h;
  if (w > ctx.availableWidth) {
    h = Math.round((h * ctx.availableWidth) / w);
    w = ctx.availableWidth;
  }
  return { w: Math.max(1, w), h: Math.max(1, h) };
}

function encodeImage(img: ImgNode, ctx: HwpxCtx): string {
  // 0. 플레이스홀더 처리 (차트 등 b64가 없는 경우)
  if (!img.b64) {
    return `<hp:t xml:space="preserve">${esc(img.alt || "[개체]")}</hp:t>`;
  }

  const binId = ctx.imgMap.get(img);
  if (!binId) return "";

  // Display size comes from the document model. Pixel dimensions describe the
  // source bitmap and are used only for the clip/source coordinate space.
  const { w: wHwp, h: hHwp } = getImageDisplayDims(img, ctx);
  const sourceDims = getImageSourceDims(img);

  // 회전 중심점 (rotation center) 계산: 이미지 중앙을 기준으로 회전
  const rotationCenterX = Math.round(wHwp / 2);
  const rotationCenterY = Math.round(hHwp / 2);

  const layout = img.layout;
  const isInline = !layout || layout.wrap === "inline";
  const textWrap = layout ? (WRAP_MAP[layout.wrap] ?? "SQUARE") : "SQUARE";
  const textFlow = layout
    ? (FLOW_MAP[layout.wrap] ?? "BOTH_SIDES")
    : "BOTH_SIDES";
  const zOrder = ctx.nextZOrder++;

  return (
    `<hp:pic id="${ctx.nextElementId++}" zOrder="${zOrder}" numberingType="PICTURE" ` +
    `textWrap="${textWrap}" textFlow="${textFlow}" lock="0" dropcapstyle="None" href="" groupLevel="0" instid="0" reverse="0">` +
    `<hp:offset x="0" y="0"/>` +
    `<hp:orgSz width="${wHwp}" height="${hHwp}"/>` +
    `<hp:curSz width="${wHwp}" height="${hHwp}"/>` +
    `<hp:flip horizontal="0" vertical="0"/>` +
    `<hp:rotationInfo angle="0" centerX="${rotationCenterX}" centerY="${rotationCenterY}" rotateimage="1"/>` +
    `<hp:renderingInfo>` +
    `<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>` +
    `<hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>` +
    `<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>` +
    `</hp:renderingInfo>` +
    `<hp:imgRect>` +
    `<hc:pt0 x="0" y="0"/><hc:pt1 x="${wHwp}" y="0"/>` +
    `<hc:pt2 x="${wHwp}" y="${hHwp}"/><hc:pt3 x="0" y="${hHwp}"/>` +
    `</hp:imgRect>` +
    `<hp:imgClip left="0" right="${sourceDims.w}" top="0" bottom="${sourceDims.h}"/>` +
    `<hp:inMargin left="0" right="0" top="0" bottom="0"/>` +
    `<hp:imgDim dimwidth="${sourceDims.w}" dimheight="${sourceDims.h}"/>` +
    `<hc:img binaryItemIDRef="${binId}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/>` +
    `<hp:effects/>` +
    `<hp:sz width="${wHwp}" widthRelTo="ABSOLUTE" height="${hHwp}" heightRelTo="ABSOLUTE" protect="0"/>` +
    `<hp:pos treatAsChar="${isInline ? 1 : 0}" affectLSpacing="0" flowWithText="1" ` +
    `allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" ` +
    `vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>` +
    `<hp:outMargin left="0" right="0" top="0" bottom="0"/>` +
    `</hp:pic>`
  );
}

function encodeImgWrapped(img: ImgNode, ctx: HwpxCtx): string {
  const content = encodeImage(img, ctx);
  if (!img.b64) {
    return `<hp:run charPrIDRef="0" charTcId="0">${content}</hp:run>`;
  }
  return `<hp:run charPrIDRef="0" charTcId="0">${content}<hp:t xml:space="preserve"> </hp:t></hp:run>`;
}

// ─── 표(Grid) 인코딩 ─────────────────────────────────────────

function encodeGridPositioned(
  grid: GridNode,
  ctx: HwpxCtx,
  vertPos: number,
  secPr = "",
  hfRun = "",
  pageFirst = false,
): { xml: string; nextVertPos: number; hasPageBreak: boolean } {
  const { xml: gridXml, height: tblHeight } = buildGridXml(grid, ctx);
  const floats =
    grid.props.layout !== undefined && grid.props.layout.wrap !== "inline";
  const totalHeight = floats ? 1000 : Math.max(1600, tblHeight);
  const baseline = Math.round(totalHeight * 0.85);

  const linesegXml =
    `<hp:linesegarray>` +
    `<hp:lineseg textpos="0" vertpos="${vertPos}" vertsize="${totalHeight}" ` +
    `textheight="${totalHeight}" baseline="${baseline}" spacing="0" ` +
    `horzpos="0" horzsize="${ctx.availableWidth}" ` +
    `flags="${LINESEG_FLAGS | (pageFirst ? LINESEG_FLAG_PAGE_FIRST | LINESEG_FLAG_COLUMN_FIRST : 0)}"/>` +
    `</hp:linesegarray>`;

  const xml =
    `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0" paraTcId="0">` +
    secPr +
    hfRun +
    `<hp:run charPrIDRef="0" charTcId="0">` +
    gridXml +
    `</hp:run>` +
    linesegXml +
    `</hp:p>`;

  return { xml, nextVertPos: vertPos + totalHeight, hasPageBreak: false };
}

function buildGridLayoutAttrs(
  layout: ImgLayout | undefined,
  fallbackHorzAlign: string,
): { textWrap: string; zOrder: number; noAdjust: string; posXml: string; outMarginXml: string } {
  const floats = layout !== undefined && layout.wrap !== "inline";
  const textWrapMap: Record<string, string> = {
    inline: "TOP_AND_BOTTOM",
    topAndBottom: "TOP_AND_BOTTOM",
    square: "SQUARE",
    tight: "BOTH_SIDES",
    through: "BOTH_SIDES",
    none: "TOP_AND_BOTTOM",
    behind: "BEHIND_TEXT",
    front: "FRONT_TEXT",
  };
  const horzRelMap: Record<string, string> = {
    para: "PARA",
    margin: "MARGIN",
    page: "PAGE",
    column: "COLUMN",
  };
  const vertRelMap: Record<string, string> = {
    para: "PARA",
    margin: "MARGIN",
    page: "PAGE",
    line: "LINE",
  };
  const horzAlignMap: Record<string, string> = {
    left: "LEFT",
    center: "CENTER",
    right: "RIGHT",
  };
  const vertAlignMap: Record<string, string> = {
    top: "TOP",
    center: "CENTER",
    bottom: "BOTTOM",
  };

  const horzAlign =
    (layout?.horzAlign ? horzAlignMap[layout.horzAlign] : undefined) ??
    fallbackHorzAlign;
  const vertAlign =
    (layout?.vertAlign ? vertAlignMap[layout.vertAlign] : undefined) ?? "TOP";
  const horzRelTo =
    (layout?.horzRelTo ? horzRelMap[layout.horzRelTo] : undefined) ?? "PARA";
  const vertRelTo =
    (layout?.vertRelTo ? vertRelMap[layout.vertRelTo] : undefined) ?? "PARA";
  const horzOffset = layout?.xPt != null ? Metric.ptToHwp(layout.xPt) : 0;
  const vertOffset = layout?.yPt != null ? Metric.ptToHwp(layout.yPt) : 0;

  const posXml =
    `<hp:pos treatAsChar="${floats ? "0" : "1"}" affectLSpacing="0" ` +
    `flowWithText="${floats && (layout?.vertRelTo === "page" || layout?.horzRelTo === "page") ? "0" : "1"}" ` +
    `allowOverlap="${floats ? "1" : "0"}" holdAnchorAndSO="0" ` +
    `vertRelTo="${vertRelTo}" horzRelTo="${horzRelTo}" ` +
    `vertAlign="${vertAlign}" horzAlign="${horzAlign}" ` +
    `vertOffset="${vertOffset}" horzOffset="${horzOffset}"/>`;

  const dist = (value: number | undefined) =>
    value != null ? Math.max(0, Metric.ptToHwp(value)) : 138;
  const outMarginXml =
    `<hp:outMargin left="${dist(layout?.distL)}" right="${dist(layout?.distR)}" ` +
    `top="${dist(layout?.distT)}" bottom="${dist(layout?.distB)}"/>`;

  return {
    textWrap: textWrapMap[layout?.wrap ?? "inline"] ?? "TOP_AND_BOTTOM",
    zOrder: Math.round(layout?.zOrder ?? 0),
    noAdjust: floats ? "1" : "0",
    posXml,
    outMarginXml,
  };
}

function buildGridXml(
  grid: GridNode,
  ctx: HwpxCtx,
  maxWidth = ctx.availableWidth,
): { xml: string; height: number } {
  const rowCount = grid.kids.length;
  // ... (기존 tableMap 생성 로직 동일)

  // 가상 2D 맵 — 병합 셀 처리
  interface CellEntry {
    type: "real" | "absorbed";
    cell?: CellNode;
  }
  const tableMap: CellEntry[][] = Array.from({ length: rowCount }, () => []);

  for (let ri = 0; ri < rowCount; ri++) {
    let ci = 0;
    for (const cell of grid.kids[ri].kids) {
      while (tableMap[ri][ci]) ci++;
      tableMap[ri][ci] = { type: "real", cell };
      for (let rr = 0; rr < cell.rs; rr++) {
        const tri = ri + rr;
        if (tri >= rowCount) break;
        for (let cc = 0; cc < cell.cs; cc++) {
          if (rr === 0 && cc === 0) continue;
          tableMap[tri][ci + cc] = { type: "absorbed" };
        }
      }
      ci += cell.cs;
    }
  }

  let colCount = 0;
  for (let ri = 0; ri < rowCount; ri++)
    colCount = Math.max(colCount, tableMap[ri].length);
  if (colCount === 0) colCount = 1;

  // 컬럼 너비 계산 (Bug 6: 균등 배분 금지, 원본 보존)
  const totalW = Math.max(1, Math.min(ctx.availableWidth, maxWidth));
  const sourceWidths = (grid.props.colWidths ?? []).map((width) =>
    width > 0 ? Metric.ptToHwp(width) : 0,
  );
  const colWidths = fitColumnWidths(
    sourceWidths,
    colCount,
    totalW,
    Math.min(100, Math.floor(totalW / colCount)),
  );
  const actualTotal = colWidths.reduce((s, w) => s + w, 0);
  const tablePadL = Metric.ptToHwp(grid.props.cellPadL ?? 1.41);
  const tablePadR = Metric.ptToHwp(grid.props.cellPadR ?? 1.41);
  const tablePadT = Metric.ptToHwp(grid.props.cellPadT ?? 1.41);
  const tablePadB = Metric.ptToHwp(grid.props.cellPadB ?? 1.41);

  // 행 높이 계산
  const rowHeights: number[] = [];
  for (let ri = 0; ri < rowCount; ri++) {
    let minRowH = 0;
    for (let ci = 0; ci < colCount; ci++) {
      const entry = tableMap[ri][ci];
      if (entry?.type === "real") {
        const cell = entry.cell!;
        const cp = cell.props ?? {};
        let cellW = 0;
        for (let sc = ci; sc < ci + cell.cs && sc < colWidths.length; sc++)
          cellW += colWidths[sc];
        if (!cellW) cellW = Math.round(totalW / colCount) * cell.cs;
        const padL = cp.padL !== undefined ? Metric.ptToHwp(cp.padL) : tablePadL;
        const padR = cp.padR !== undefined ? Metric.ptToHwp(cp.padR) : tablePadR;
        const innerW = Math.max(cellW - padL - padR, 100);
        const span = Math.max(1, cell.rs ?? 1);
        const h = estimateCellHeight(cell, ctx, innerW);
        minRowH = Math.max(minRowH, Math.ceil(h / span));
      }
    }
    const baseH =
      grid.kids[ri].heightPt != null &&
      (grid.kids[ri].heightPt as number) > 0
        ? Metric.ptToHwp(grid.kids[ri].heightPt as number)
        : Math.round(1000 * 1.6);
    if (
      grid.kids[ri].heightPt != null &&
      (grid.kids[ri].heightPt as number) > 0
    ) {
      rowHeights.push(Math.max(baseH, minRowH));
    } else {
      rowHeights.push(Math.max(baseH, minRowH));
    }
  }
  const totalH = rowHeights.reduce((s, h) => s + h, 0);

  const defStroke = grid.props.defaultStroke ?? DEFAULT_STROKE;
  // 표 기본 테두리 — BorderFillBank에서 실제 ID 조회
  const tblBfId = ctx.borderFillBank.addUniform(defStroke);

  let rowsXml = "";
  for (let ri = 0; ri < rowCount; ri++) {
    let cellsXml = "";
    for (let ci = 0; ci < colCount; ci++) {
      const entry = tableMap[ri][ci];
      if (!entry || entry.type === "absorbed") continue;
      const cell = entry.cell!;
      const cp = cell.props;

      // 셀 테두리 — BorderFillBank에서 실제 ID 조회 (하드코딩 제거)
      const cellBfId = ctx.borderFillBank.addFromCellProps(cp, defStroke);

      let cellW = 0;
      for (let sc = ci; sc < ci + cell.cs && sc < colWidths.length; sc++)
        cellW += colWidths[sc];
      if (!cellW) cellW = Math.round(totalW / colCount) * cell.cs;

      const subListId = ctx.nextElementId++;

      const padL = cp.padL !== undefined ? Metric.ptToHwp(cp.padL) : tablePadL;
      const padR = cp.padR !== undefined ? Metric.ptToHwp(cp.padR) : tablePadR;
      const padT = cp.padT !== undefined ? Metric.ptToHwp(cp.padT) : tablePadT;
      const padB = cp.padB !== undefined ? Metric.ptToHwp(cp.padB) : tablePadB;

      const innerW = Math.max(cellW - padL - padR, 100);
      let parasXml = '';
      let localVertPos = 0;
      if (cell.kids.length > 0) {
        for (const kid of cell.kids) {
          if (kid.tag === 'grid') {
            const { xml: tblXml, height: nestedHeight } = buildGridXml(kid, ctx, innerW);
            const pid = ctx.nextElementId++;
            const objectHeight = Math.max(1600, nestedHeight);
            const baseline = Math.round(objectHeight * 0.85);
            parasXml +=
              `<hp:p id="${pid}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0" paraTcId="0">` +
              `<hp:run charPrIDRef="0" charTcId="0">${tblXml}</hp:run>` +
              `<hp:linesegarray><hp:lineseg textpos="0" vertpos="${localVertPos}" ` +
              `vertsize="${objectHeight}" textheight="${objectHeight}" baseline="${baseline}" spacing="0" ` +
              `horzpos="0" horzsize="${innerW}" flags="${LINESEG_FLAGS}"/></hp:linesegarray>` +
              `</hp:p>`;
            localVertPos += objectHeight;
          } else {
            const encoded = encodeParaPositioned(kid, ctx, localVertPos, '', innerW);
            parasXml += encoded.xml;
            localVertPos = encoded.nextVertPos;
          }
        }
      } else {
        const { xml: emptyLineseg } = buildLinesegarray(" ", 0, 1000, 160, innerW);
        parasXml = `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0" paraTcId="0"><hp:run charPrIDRef="0" charTcId="0"><hp:t xml:space="preserve"> </hp:t></hp:run>${emptyLineseg}</hp:p>`;
      }

      const vAlign =
        cp.va === "mid" ? "CENTER" : cp.va === "bot" ? "BOTTOM" : "TOP";

      const cellHeight = rowHeights
        .slice(ri, Math.min(rowHeights.length, ri + Math.max(1, cell.rs)))
        .reduce((sum, height) => sum + height, 0);

      cellsXml +=
        `<hp:tc name="" header="${cp.isHeader || (grid.props.headerRow && ri === 0) ? 1 : 0}" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="${cellBfId}">` +
        `<hp:subList id="${subListId}" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="${vAlign}" ` +
        `linkListIDRef="0" linkListNextIDRef="0" textWidth="${innerW}" textHeight="${Math.max(100, cellHeight - padT - padB)}" hasTextRef="0" hasNumRef="0">` +
        parasXml +
        `</hp:subList>` +
        `<hp:cellAddr colAddr="${ci}" rowAddr="${ri}"/>` +
        `<hp:cellSpan colSpan="${cell.cs}" rowSpan="${cell.rs}"/>` +
        `<hp:cellSz width="${cellW}" height="${cellHeight}"/>` +
        `<hp:cellMargin left="${padL}" right="${padR}" top="${padT}" bottom="${padB}"/>` +
        `</hp:tc>`;
    }
    rowsXml += `<hp:tr>${cellsXml}</hp:tr>`;
  }

  // 표 정렬 처리
  const alignMap: Record<string, string> = {
    left: 'LEFT', right: 'RIGHT', center: 'CENTER', justify: 'JUSTIFY',
  };
  const horzAlign = alignMap[grid.props.align ?? 'left'] ?? 'LEFT';
  const layoutAttrs = buildGridLayoutAttrs(grid.props.layout, horzAlign);

  const repeatHeader = grid.props.headerRow ? 1 : 0;
  const xml =
    `<hp:tbl id="${ctx.nextElementId++}" zOrder="${layoutAttrs.zOrder}" numberingType="TABLE" textWrap="${layoutAttrs.textWrap}" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="${repeatHeader}" rowCnt="${rowCount}" colCnt="${colCount}" cellSpacing="0" borderFillIDRef="${tblBfId}" noAdjust="${layoutAttrs.noAdjust}">` +
    `<hp:sz width="${actualTotal}" widthRelTo="ABSOLUTE" height="${totalH}" heightRelTo="ABSOLUTE" protect="0"/>` +
    layoutAttrs.posXml +
    layoutAttrs.outMarginXml +
    `<hp:inMargin left="${tablePadL}" right="${tablePadR}" top="${tablePadT}" bottom="${tablePadB}"/>` +
    rowsXml +
    `</hp:tbl>`;

  return { xml, height: totalH };
}

function estimateLineCountForWidth(
  text: string,
  fontSize: number,
  horzSize?: number,
): number {
  if (!text) return 2;
  const maxWidth = Math.max(1, horzSize ?? 0);
  if (!horzSize || horzSize <= 0) return text.split(/\r\n|\r|\n/).length;
  let lines = 1;
  let currentLineWidth = 0;
  for (let i = 0; i < text.length; i++) {
    const charCode = text.charCodeAt(i);
    if (charCode === 10 || charCode === 13) {
      if (charCode === 13 && text.charCodeAt(i + 1) === 10) i++;
      lines++;
      currentLineWidth = 0;
      continue;
    }

    let charW = fontSize * 0.55;
    if (charCode >= 0xac00 && charCode <= 0xd7a3) charW = fontSize;
    else if (charCode >= 0x3130 && charCode <= 0x318f) charW = fontSize;
    else if (charCode >= 0x4e00 && charCode <= 0x9fff) charW = fontSize;
    else if (charCode >= 65 && charCode <= 90) charW = fontSize * 0.65;
    else if (charCode === 32) charW = fontSize * 0.32;
    else if (charCode > 255) charW = fontSize;
    else charW = fontSize * 0.42;

    if (currentLineWidth > 0 && currentLineWidth + charW > maxWidth) {
      lines++;
      currentLineWidth = charW;
    } else {
      currentLineWidth += charW;
    }
  }
  return Math.max(1, lines);
}

function estimateGridHeight(grid: GridNode, ctx: HwpxCtx): number {
  let total = 0;
  for (const row of grid.kids) {
    const base =
      row.heightPt != null && row.heightPt > 0
        ? Metric.ptToHwp(row.heightPt)
        : Math.round(1000 * 1.6);
    let minRow = 0;
    for (const cell of row.kids) {
      const span = Math.max(1, cell.rs ?? 1);
      minRow = Math.max(minRow, Math.ceil(estimateCellHeight(cell, ctx) / span));
    }
    total += Math.max(base, minRow);
  }
  return total;
}

function estimateCellHeight(
  cell: CellNode,
  ctx: HwpxCtx,
  innerWidth?: number,
): number {
  const cp = cell.props ?? {};
  const topPad = cp.padT !== undefined ? Metric.ptToHwp(cp.padT) : 141;
  const botPad = cp.padB !== undefined ? Metric.ptToHwp(cp.padB) : 141;
  let h = 0;
  for (const kid of cell.kids) {
    if (kid.tag === 'grid') {
      h += estimateGridHeight(kid, ctx);
      continue;
    }
    const para = kid;
    const fs = fontSizeForPara(para, ctx);
    const ppId = ctx.paraPrMap.get(paraPrKey(para.props));
    const pp = ppId !== undefined ? ctx.paraPrs[ppId] : null;
    const lineHeight = pp?.lineSpacingFixed !== undefined
      ? Math.max(
          paraShapeHwpToLayoutHwp(pp.lineSpacingFixed),
          Math.ceil(fs * 1.15),
        )
      : Math.max(fs, Math.round((fs * Math.max(100, pp?.lineSpacing ?? 160)) / 100));
    const textHeight = Math.max(fs, inlineObjectHeightForPara(para, ctx));
    const lineAdvance = textHeight + Math.max(0, lineHeight - fs);
    const lineCount = estimateLineCountForWidth(
      extractParaText(para),
      fs,
      innerWidth,
    );
    const before = paraShapeHwpToLayoutHwp(pp?.prevHwp ?? 0);
    const after = paraShapeHwpToLayoutHwp(pp?.nextHwp ?? 0);
    h += lineAdvance * lineCount + before + after;
  }
  if (!h) h = Math.round(1000 * 1.6);
  return h + topPad + botPad;
}

// ─── 미리보기 텍스트 추출 ────────────────────────────────────

function extractPreviewText(sheet?: SheetNode): string {
  if (!sheet) return "";
  const lines: string[] = [];
  for (const kid of sheet.kids) {
    if (kid.tag === "para") {
      const text = kid.kids
        .flatMap((k) =>
          k.tag === "span"
            ? k.kids.flatMap((c) => (c.tag === "txt" ? [c.content] : []))
            : [],
        )
        .join("");
      if (text) lines.push(text);
    } else if (kid.tag === "grid") {
      for (const row of kid.kids) {
        const cells = row.kids.map((cell) =>
          cell.kids
            .flatMap((p) =>
              p.tag === 'para'
                ? p.kids.flatMap((k) =>
                    k.tag === "span"
                      ? k.kids.flatMap((c) => (c.tag === "txt" ? [c.content] : []))
                      : [],
                  )
                : [],
            )
            .join(""),
        );
        lines.push(cells.join("\t"));
      }
    }
  }
  return lines.join("\r\n");
}

// ─── XML 이스케이프 ──────────────────────────────────────────

function esc(s: string): string {
  if (!s) return "";
  s = s.replace(/__EXT_\d+(?:_W\d+_H\d+)?__/g, "");
  s = s.replace(/湰灧/g, "").replace(/\uFEFF/g, "");
  // XML 1.0 비허용 제어문자 제거
  // eslint-disable-next-line no-control-regex
  s = s.replace(
    /[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/gu,
    "",
  );
  return TextKit.escapeXml(s);
}

registry.registerEncoder(new HwpxEncoder());
