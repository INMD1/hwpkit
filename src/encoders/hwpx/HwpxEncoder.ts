/**
 * HwpxEncoder — DocRoot → HWPX (ZIP + XML)
 *
 * ANYTOHWP에서 영감받은 개선 사항:
 *  1. LangFontBank  — 7개 언어 그룹 독립 폰트 레지스트리 (HANGUL/LATIN/HANJA/…)
 *  2. BorderFillBank — 정확한 ID 관리 (하드코딩 "1" 제거)
 *  3. readPixelDims  — PNG/JPEG 바이너리 헤더에서 실제 픽셀 치수 추출
 *  4. 두 패스 구조   — Pre-scan(등록) → Encode(생성)
 */

import type { Encoder } from "../../contract/encoder";
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
import type {
  DocMeta,
  PageDims,
  TextProps,
  ParaProps,
  CellProps,
  Stroke,
} from "../../model/doc-props";
import { A4, DEFAULT_STROKE, normalizeDims } from "../../model/doc-props";
import { succeed, fail } from "../../contract/result";
import { Metric, safeFontToKr } from "../../safety/StyleBridge";
import { ArchiveKit } from "../../toolkit/ArchiveKit";
import { TextKit } from "../../toolkit/TextKit";
import { registry } from "../../pipeline/registry";

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
].join(" ");

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

  /** TextProps.font 문자열에서 적절한 HANGUL/LATIN 그룹에 등록 */
  registerFont(rawFace: string): { hangulId: number; latinId: number } {
    const face = safeFontToKr(rawFace) || "함초롬바탕";
    const isKor = this.isKorean(face);
    // 한글 폰트: HANGUL/HANJA/JAPANESE/OTHER/SYMBOL/USER에 등록
    // 라틴 폰트: LATIN에 등록, 나머지는 기본값(0) 유지
    const hangulId = this.register("HANGUL", isKor ? face : "함초롬바탕");
    const latinId = this.register("LATIN", isKor ? "함초롬바탕" : face);
    for (const g of [
      "HANJA",
      "JAPANESE",
      "OTHER",
      "SYMBOL",
      "USER",
    ] as LangGroup[]) {
      this.register(g, isKor ? face : "함초롬바탕");
    }
    return { hangulId, latinId };
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
      s && s.kind !== "none" ? `${(s.pt * 0.3528).toFixed(2)} mm` : "0.12 mm";
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
  bg?: string;
}

interface ParaPrDef {
  id: number;
  align: string;
  leftHwp: number;
  intentHwp: number;
  prevHwp: number;
  nextHwp: number;
  lineSpacing: number;
  listType?: string;
  listLevel?: number;
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

function paraPrKey(p: ParaProps): string {
  return `${p.align ?? "left"}|${p.listOrd ?? ""}|${p.listLv ?? 0}|${p.indentPt ?? 0}|${p.firstLineIndentPt ?? 0}|${p.spaceBefore ?? 0}|${p.spaceAfter ?? 0}|${p.lineHeight ?? 0}|${p.styleId ?? ""}`;
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
  const { hangulId, latinId } = ctx.fontBank.registerFont(rawFont);
  const id = ctx.charPrs.length;

  ctx.charPrs.push({
    id,
    height: Metric.ptToHHeight(props.pt ?? 10),
    bold: !!props.b,
    italic: !!props.i,
    underline: props.u ? "BOTTOM" : "NONE",
    strikeout: props.s ? "SOLID" : "NONE",
    textColor: props.color ? `#${props.color}` : "#000000",
    hangulId,
    latinId,
    bg: props.bg,
  });
  ctx.charPrMap.set(key, id);
  return id;
}

function registerParaPr(props: ParaProps, ctx: HwpxCtx): number {
  const key = paraPrKey(props);
  const existing = ctx.paraPrMap.get(key);
  if (existing !== undefined) return existing;

  const id = ctx.paraPrs.length;
  const def: ParaPrDef = {
    id,
    align: (props.align ?? "left").toUpperCase(),
    leftHwp: props.indentPt ? Metric.ptToHwp(props.indentPt) : 0,
    intentHwp: props.firstLineIndentPt
      ? Metric.ptToHwp(props.firstLineIndentPt)
      : 0,
    prevHwp: props.spaceBefore ? Metric.ptToHwp(props.spaceBefore) : 0,
    nextHwp: props.spaceAfter ? Metric.ptToHwp(props.spaceAfter) : 0,
    lineSpacing: props.lineHeight ? Math.round(props.lineHeight * 100) : 160,
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
  if (styleId === "Normal") {
    ctx.styleIdToHwpxId.set(styleId, 0);
    return;
  }
  const hwpxId = ctx.hwpxStyles.length;
  ctx.styleIdToHwpxId.set(styleId, hwpxId);
  ctx.hwpxStyles.push({
    id: hwpxId,
    name: STYLE_NAME_MAP[styleId] ?? styleId,
    engName: "",
    paraPrIDRef: paraPrId,
    charPrIDRef: charPrId,
  });
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
  if (para.props.styleId)
    registerStyle(para.props.styleId, paraPrId, firstCharPrId, ctx);
}

function scanGrid(grid: GridNode, ctx: HwpxCtx): void {
  const defStroke = grid.props.defaultStroke ?? DEFAULT_STROKE;
  // 표 기본 테두리 사전 등록
  ctx.borderFillBank.addUniform(defStroke);
  for (const row of grid.kids) {
    for (const cell of row.kids) {
      ctx.borderFillBank.addFromCellProps(cell.props, defStroke);
      for (const p of cell.kids) scanPara(p, ctx);
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

export class HwpxEncoder implements Encoder {
  readonly format = "hwpx";

  async encode(doc: DocRoot): Promise<Outcome<Uint8Array>> {
    try {
      const sheet = doc.kids[0];
      const dims = normalizeDims(sheet?.dims ?? A4);

      const safeML = dims.ml > 0 ? dims.ml : 70.87;
      const safeMR = dims.mr > 0 ? dims.mr : 70.87;
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
      scanContent(sheet?.kids ?? [], ctx);
      if (sheet?.header) for (const p of sheet.header) scanPara(p, ctx);
      if (sheet?.footer) for (const p of sheet.footer) scanPara(p, ctx);

      // 패스 2: Encode — section 먼저 (borderFill 동적 등록 완료 후 header 생성)
      const sectionData = TextKit.encode(buildSectionXml(sheet, dims, ctx));
      const headerData = TextKit.encode(buildHeaderXml(dims, doc.meta, ctx));
      const previewText = extractPreviewText(sheet);

      const entries: { name: string; data: Uint8Array; mime: string }[] = [
        {
          name: "mimetype",
          data: TextKit.encode("application/hwp+zip"),
          mime: "",
        },
        {
          name: "version.xml",
          data: TextKit.encode(VERSION_XML),
          mime: "application/xml",
        },
        {
          name: "META-INF/container.xml",
          data: TextKit.encode(CONTAINER_XML),
          mime: "application/xml",
        },
        {
          name: "META-INF/container.rdf",
          data: TextKit.encode(CONTAINER_RDF),
          mime: "application/rdf+xml",
        },
        {
          name: "Contents/content.hpf",
          data: TextKit.encode(buildContentHpf(ctx, doc.meta)),
          mime: "application/hwpml-package+xml",
        },
        {
          name: "Contents/header.xml",
          data: headerData,
          mime: "application/xml",
        },
        {
          name: "Contents/section0.xml",
          data: sectionData,
          mime: "application/xml",
        },
        {
          name: "Preview/PrvText.txt",
          data: TextKit.encode(previewText),
          mime: "text/plain",
        },
        {
          name: "settings.xml",
          data: TextKit.encode(buildSettingsXml()),
          mime: "application/xml",
        },
        {
          name: "META-INF/manifest.xml",
          data: TextKit.encode(MANIFEST_XML),
          mime: "text/xml",
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
                : "image/bmp";
        entries.push({ name: `BinData/${bin.name}`, data: bin.data, mime: ct });
      }

      return succeed(await ArchiveKit.zip(entries));
    } catch (e: any) {
      return fail(`HWPX 인코딩 오류: ${e?.message ?? String(e)}`);
    }
  }
}

// ─── 상수 XML ────────────────────────────────────────────────

const VERSION_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>` +
  `<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" ` +
  `targetApplication="WORDPROCESSOR" major="5" minor="1" micro="0" buildNumber="1" ` +
  `os="1" xmlVersion="1.4" application="Hancom Office Hangul" appVersion="11, 0, 0, 0"/>`;

const CONTAINER_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>` +
  `<ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" ` +
  `xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf">` +
  `<ocf:rootfiles>` +
  `<ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/>` +
  `<ocf:rootfile full-path="Preview/PrvText.txt" media-type="text/plain"/>` +
  `<ocf:rootfile full-path="META-INF/container.rdf" media-type="application/rdf+xml"/>` +
  `</ocf:rootfiles></ocf:container>`;

const CONTAINER_RDF =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>` +
  `<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">` +
  `<rdf:Description rdf:about=""><ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" rdf:resource="Contents/header.xml"/></rdf:Description>` +
  `<rdf:Description rdf:about="Contents/header.xml"><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#HeaderFile"/></rdf:Description>` +
  `<rdf:Description rdf:about=""><ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" rdf:resource="Contents/section0.xml"/></rdf:Description>` +
  `<rdf:Description rdf:about="Contents/section0.xml"><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#SectionFile"/></rdf:Description>` +
  `<rdf:Description rdf:about=""><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#Document"/></rdf:Description>` +
  `</rdf:RDF>`;

const MANIFEST_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>` +
  `<odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>`;

// ─── content.hpf ─────────────────────────────────────────────

function buildContentHpf(ctx: HwpxCtx, meta?: DocMeta): string {
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
    `<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>` +
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
    `<opf:spine><opf:itemref idref="header"/><opf:itemref idref="section0"/></opf:spine>` +
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
    `</ha:HWPApplicationSetting>`
  );
}

// ─── header.xml ──────────────────────────────────────────────

function buildHeaderXml(dims: PageDims, meta: DocMeta, ctx: HwpxCtx): string {
  // 언어별 폰트 (LangFontBank → XML)
  const fontFacesXml = ctx.fontBank.toXml();

  // charPr 목록 — 언어별 폰트 ID를 fontRef에 반영 (ANYTOHWP 핵심 개선)
  let charPrXml = "";
  for (const cp of ctx.charPrs) {
    const bold = cp.bold ? "<hh:bold/>" : "";
    const italic = cp.italic ? "<hh:italic/>" : "";
    const hid = cp.hangulId;
    const lid = cp.latinId;
    charPrXml +=
      `<hh:charPr id="${cp.id}" height="${cp.height}" textColor="${cp.textColor}" ` +
      `shadeColor="none" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="1">` +
      `<hh:fontRef hangul="${hid}" latin="${lid}" hanja="${hid}" japanese="${hid}" other="${lid}" symbol="${lid}" user="${lid}"/>` +
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

  // paraPr 목록
  let paraPrXml = "";
  for (const pp of ctx.paraPrs) {
    paraPrXml +=
      `<hh:paraPr id="${pp.id}" tabPrIDRef="0" condense="0" fontLineHeight="0" snapToGrid="0" suppressLineNumbers="0" checked="0">` +
      `<hh:align horizontal="${pp.align}" vertical="BASELINE"/>` +
      `<hh:heading type="NONE" idRef="0" level="0"/>` +
      `<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="BREAK_WORD" widowOrphan="0" keepWithNext="0" keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/>` +
      `<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>` +
      `<hh:margin>` +
      `<hc:intent value="${pp.intentHwp}" unit="HWPUNIT"/>` +
      `<hc:left value="${pp.leftHwp}" unit="HWPUNIT"/>` +
      `<hc:right value="0" unit="HWPUNIT"/>` +
      `<hc:prev value="${pp.prevHwp}" unit="HWPUNIT"/>` +
      `<hc:next value="${pp.nextHwp}" unit="HWPUNIT"/>` +
      `</hh:margin>` +
      `<hh:lineSpacing type="PERCENT" value="${pp.lineSpacing}" unit="HWPUNIT"/>` +
      `<hh:border borderFillIDRef="1" offsetLeft="0" offsetRight="0" offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>` +
      `</hh:paraPr>`;
  }

  // borderFill 목록 (BorderFillBank → XML)
  const borderFillXml = ctx.borderFillBank.toXml();

  // 스타일 목록
  const stylesXml =
    `<hh:styles itemCnt="${ctx.hwpxStyles.length}">` +
    ctx.hwpxStyles
      .map(
        (s) =>
          `<hh:style id="${s.id}" type="PARA" name="${esc(s.name)}" engName="${esc(s.engName)}" ` +
          `paraPrIDRef="${s.paraPrIDRef}" charPrIDRef="${s.charPrIDRef}" nextStyleIDRef="0" langID="1042" lockForm="0"/>`,
      )
      .join("") +
    `</hh:styles>`;

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>` +
    `<hh:head ${NS} version="1.2" secCnt="1">` +
    `<hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>` +
    `<hh:refList>` +
    fontFacesXml +
    borderFillXml +
    `<hh:charProperties itemCnt="${ctx.charPrs.length}">${charPrXml}</hh:charProperties>` +
    `<hh:tabProperties itemCnt="1"><hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/></hh:tabProperties>` +
    `<hh:paraProperties itemCnt="${ctx.paraPrs.length}">${paraPrXml}</hh:paraProperties>` +
    stylesXml +
    `</hh:refList>` +
    `<hh:compatibleDocument targetProgram="MS_WORD">        
          <hh:layoutCompatibility>
            <hh:applyFontWeightToBold />
            <hh:useInnerUnderline />
            <hh:useLowercaseStrikeout />
            <hh:extendLineheightToOffset />
            <hh:treatQuotationAsLatin />
            <hh:doNotAlignWhitespaceOnRight />
            <hh:doNotAdjustWordInJustify />
            <hh:baseCharUnitOnEAsian />
            <hh:baseCharUnitOfIndentOnFirstChar />
            <hh:adjustLineheightToFont />
            <hh:adjustBaselineInFixedLinespacing />
            <hh:applyPrevspacingBeneathObject />
            <hh:applyNextspacingOfLastPara />
            <hh:adjustParaBorderfillToSpacing />
            <hh:connectParaBorderfillOfEqualBorder />
            <hh:adjustParaBorderOffsetWithBorder />
            <hh:extendLineheightToParaBorderOffset />
            <hh:applyParaBorderToOutside />
            <hh:applyMinColumnWidthTo1mm />
            <hh:applyTabPosBasedOnSegment />
            <hh:breakTabOverLine />
            <hh:adjustVertPosOfLine />
            <hh:doNotAlignLastForbidden />
            <hh:adjustMarginFromAdjustLineheight />
            <hh:baseLineSpacingOnLineGrid />
            <hh:applyCharSpacingToCharGrid />
            <hh:doNotApplyGridInHeaderFooter />
            <hh:applyExtendHeaderFooterEachSection />
            <hh:doNotApplyLinegridAtNoLinespacing />
            <hh:doNotAdjustEmptyAnchorLine />
            <hh:overlapBothAllowOverlap />
            <hh:extendVertLimitToPageMargins />
            <hh:doNotHoldAnchorOfTable />
            <hh:doNotFormattingAtBeneathAnchor />
            <hh:adjustBaselineOfObjectToBottom />
        </hh:layoutCompatibility>
    </hh:compatibleDocument>` +
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
  const hasHeader = sheet.header && sheet.header.length > 0;
  const hasFooter = sheet.footer && sheet.footer.length > 0;
  if (!hasHeader && !hasFooter) return "";

  const availW = ctx.availableWidth;
  const mtHwp = Metric.ptToHwp(dims.mt);
  const mbHwp = Metric.ptToHwp(dims.mb);
  const headerZoneH = dims.headerPt
    ? Math.max(100, mtHwp - Metric.ptToHwp(dims.headerPt))
    : 2125;
  const footerZoneH = dims.footerPt
    ? Math.max(100, mbHwp - Metric.ptToHwp(dims.footerPt))
    : 4255;

  let inner = "";

  // pageHiding: first-page hides header/footer when they exist
  inner += `<hp:ctrl><hp:pageHiding hideHeader="0" hideFooter="0" hideMasterPage="0" hideBorder="0" hideFill="0" hidePageNum="0"/></hp:ctrl>`;

  if (hasHeader) {
    // header/footer 내부 단락 ID는 원본처럼 0부터 시작하는 별도 공간 사용
    const savedId = ctx.nextElementId;
    ctx.nextElementId = 0;
    const headerParasXml = sheet.header!
      .map((p) => encodeParaPositioned(p, ctx, 0, "", availW).xml)
      .join("");
    ctx.nextElementId = savedId;
    inner +=
      `<hp:ctrl>` +
      `<hp:header id="1" applyPageType="BOTH">` +
      `<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" ` +
      `linkListIDRef="0" linkListNextIDRef="0" textWidth="${availW}" textHeight="${headerZoneH}" ` +
      `hasTextRef="0" hasNumRef="0">` +
      headerParasXml +
      `</hp:subList>` +
      `</hp:header>` +
      `</hp:ctrl>`;
  }

  if (hasFooter) {
    const savedId = ctx.nextElementId;
    ctx.nextElementId = 0;
    const footerParasXml = sheet.footer!
      .map((p) => encodeParaPositioned(p, ctx, 0, "", availW).xml)
      .join("");
    ctx.nextElementId = savedId;
    inner +=
      `<hp:ctrl>` +
      `<hp:footer id="2" applyPageType="BOTH">` +
      `<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="BOTTOM" ` +
      `linkListIDRef="0" linkListNextIDRef="0" textWidth="${availW}" textHeight="${footerZoneH}" ` +
      `hasTextRef="0" hasNumRef="0">` +
      footerParasXml +
      `</hp:subList>` +
      `</hp:footer>` +
      `</hp:ctrl>`;
  }

  return `<hp:run charPrIDRef="0">${inner}</hp:run>`;
}

function buildSectionXml(
  sheet: SheetNode | undefined,
  dims: PageDims,
  ctx: HwpxCtx,
): string {
  const secPrXml = buildSecPrXml(dims);
  const kids = sheet?.kids ?? [];
  const hfRunXml = sheet ? buildHeaderFooterRunXml(sheet, dims, ctx) : "";

  let contentXml = "";
  let vertPos = 0;

  for (let i = 0; i < kids.length; i++) {
    const kid = kids[i];
    const isFirst = i === 0;
    const curSecPr = isFirst ? secPrXml : "";
    const curHfRun = isFirst ? hfRunXml : "";

    if (kid.tag === "para") {
      const { xml, nextVertPos } = encodeParaPositioned(
        kid,
        ctx,
        vertPos,
        curSecPr,
        undefined,
        curHfRun,
      );
      contentXml += xml;
      vertPos = nextVertPos;
    } else if (kid.tag === "grid") {
      const { xml, nextVertPos } = encodeGridPositioned(
        kid,
        ctx,
        vertPos,
        curSecPr,
        curHfRun,
      );
      contentXml += xml;
      vertPos = nextVertPos;
    }
  }

  if (!contentXml) {
    // 빈 문서 — 최소 단락 1개 필수
    const fs = 1000;
    const sp = 600;
    contentXml =
      `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">` +
      secPrXml +
      hfRunXml +
      `<hp:run charPrIDRef="0"><hp:t> </hp:t></hp:run>` +
      buildLineSeg(0, fs + sp, fs, ctx.availableWidth) +
      `</hp:p>`;
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec ${NS}>${contentXml}</hs:sec>`;
}

function buildSecPrXml(dims: PageDims): string {
  const wHwp = Metric.ptToHwp(dims.wPt);
  const hHwp = Metric.ptToHwp(dims.hPt);
  const ml = Metric.ptToHwp(dims.ml);
  const mr = Metric.ptToHwp(dims.mr);
  const mt = Metric.ptToHwp(dims.mt);
  const mb = Metric.ptToHwp(dims.mb);
  // HWPX margin header/footer = header/footer ZONE HEIGHT (not distance from paper edge)
  // = top_hwp - header_from_top_hwp  (and  bottom_hwp - footer_from_bottom_hwp)
  const headerZone = dims.headerPt
    ? Math.max(0, mt - Metric.ptToHwp(dims.headerPt))
    : 0;
  const footerZone = dims.footerPt
    ? Math.max(0, mb - Metric.ptToHwp(dims.footerPt))
    : 0;

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
    `<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" outlineShapeIDRef="0" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0">` +
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
    `</hp:secPr>` +
    `<hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/></hp:ctrl>`
  );
}

// ─── 줄 정보 XML (linesegarray) ──────────────────────────────
// ANYTOHWP 방식: vertPos를 정확히 추적하며 lineseg 생성

function buildLineSeg(
  vertPos: number,
  vertSize: number,
  textHeight: number,
  horzSize: number,
): string {
  const baseline = Math.round(textHeight * 0.85);
  const spacing = vertSize - textHeight;
  return (
    `<hp:linesegarray>` +
    `<hp:lineseg textpos="0" vertpos="${vertPos}" vertsize="${vertSize}" ` +
    `textheight="${textHeight}" baseline="${baseline}" spacing="${spacing}" ` +
    `horzpos="0" horzsize="${Math.max(1, horzSize)}" flags="393216"/>` +
    `</hp:linesegarray>`
  );
}

function fontSizeForPara(para: ParaNode, ctx: HwpxCtx): number {
  for (const kid of para.kids) {
    if (kid.tag === "span") {
      const id = ctx.charPrMap.get(charPrKey(kid.props));
      if (id !== undefined && ctx.charPrs[id]) return ctx.charPrs[id].height;
    }
  }
  return 1000; // 기본 10pt
}

// ─── 단락 인코딩 ─────────────────────────────────────────────

function encodeParaPositioned(
  para: ParaNode,
  ctx: HwpxCtx,
  vertPos: number,
  secPr = "",
  availWidth?: number,
  hfRun = "",
): { xml: string; nextVertPos: number } {
  const paraPrId = ctx.paraPrMap.get(paraPrKey(para.props)) ?? 0;
  const styleIDRef = para.props.styleId
    ? (ctx.styleIdToHwpxId.get(para.props.styleId) ?? 0)
    : 0;
  const fontSize = fontSizeForPara(para, ctx);
  const paraPr = ctx.paraPrs[paraPrId];
  const lineSpacing = paraPr?.lineSpacing ?? 160;
  const spacing = Math.max(0, Math.round(fontSize * (lineSpacing / 100 - 1)));
  const vertSize = fontSize + spacing;
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
    );

  let runsXml = encodeParaKids(para.kids, ctx);
  if (!runsXml && !secPr)
    runsXml = `<hp:run charPrIDRef="0"><hp:t> </hp:t></hp:run>`;

  const hasPageBreak = para.kids.some(
    (k) => k.tag === "span" && k.kids.some((c) => c.tag === "pb"),
  );
  const linesegXml = buildLineSeg(vertPos, vertSize, fontSize, horzSize);

  const xml =
    `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="${paraPrId}" styleIDRef="${styleIDRef}" ` +
    `pageBreak="${hasPageBreak ? 1 : 0}" columnBreak="0" merged="0">` +
    secPr +
    hfRun +
    runsXml +
    linesegXml +
    `</hp:p>`;

  return { xml, nextVertPos: vertPos + vertSize };
}

function encodeCodeBlockPositioned(
  para: ParaNode,
  ctx: HwpxCtx,
  vertPos: number,
  secPr: string,
  fontSize: number,
  spacing: number,
  vertSize: number,
): { xml: string; nextVertPos: number } {
  const codeBfId = ctx.borderFillBank.addUniform(
    { kind: "solid", pt: 0.5, color: "aaaaaa" },
    "f4f4f4",
  );
  const cellW = ctx.availableWidth;
  const innerW = Math.max(cellW - 510, 100);
  const subListId = ctx.nextElementId++;
  const { xml: innerXml } = encodeParaPositioned(para, ctx, 0, "", innerW);

  const linesegXml = buildLineSeg(
    vertPos,
    vertSize,
    fontSize,
    ctx.availableWidth,
  );

  const xml =
    `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="0" styleIDRef="0">` +
    secPr +
    `<hp:run charPrIDRef="0">` +
    `<hp:tbl id="${ctx.nextElementId++}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" rowCnt="1" colCnt="1" cellSpacing="0" borderFillIDRef="${codeBfId}" noAdjust="0">` +
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
    `</hp:tc></hp:tr></hp:tbl><hp:t> </hp:t></hp:run>` +
    linesegXml +
    `</hp:p>`;

  return { xml, nextVertPos: vertPos + vertSize };
}

function encodeParaKids(kids: ParaNode["kids"], ctx: HwpxCtx): string {
  let xml = "";
  for (const kid of kids) {
    if (kid.tag === "span") xml += encodeRun(kid, ctx);
    else if (kid.tag === "img") xml += encodeImgWrapped(kid, ctx);
    else if (kid.tag === "link")
      xml += encodeParaKids((kid as LinkNode).kids as ParaNode["kids"], ctx);
  }
  return xml;
}

function encodeRun(span: SpanNode, ctx: HwpxCtx): string {
  const charPrId = ctx.charPrMap.get(charPrKey(span.props)) ?? 0;
  const parts: string[] = [];

  for (const kid of span.kids) {
    if (kid.tag === "txt") {
      const content = esc(kid.content);
      if (content) parts.push(`<hp:t xml:space="preserve">${content}</hp:t>`);
    } else if (kid.tag === "br") {
      parts.push(`<hp:t xml:space="preserve">\n</hp:t>`);
    } else if (kid.tag === "pagenum") {
      const fmt =
        (kid as any).format === "roman"
          ? "ROMAN_LOWER"
          : (kid as any).format === "romanCaps"
            ? "ROMAN_UPPER"
            : "DIGIT";
      parts.push(`<hp:pageNum pageStartsOn="BOTH" formatType="${fmt}"/>`);
    }
  }

  if (!parts.length) return "";
  return `<hp:run charPrIDRef="${charPrId}">${parts.join("")}</hp:run>`;
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

function encodeImage(img: ImgNode, ctx: HwpxCtx): string {
  const binId = ctx.imgMap.get(img);
  if (!binId) return "";

  // ANYTOHWP 영감: PNG/JPEG 바이너리 헤더에서 실제 픽셀 치수 추출
  // img.w / img.h는 pt 단위이지만, 이미지 실제 픽셀을 HWPUNIT으로 변환하면 더 정확
  const pixelDims = readPixelDims(img.b64, img.mime);
  let wHwp: number, hHwp: number;

  if (pixelDims && pixelDims.w > 0 && pixelDims.h > 0) {
    // 픽셀 → pt (96dpi 기준) → HWPUNIT
    wHwp = Metric.ptToHwp((pixelDims.w * 72) / 96);
    hHwp = Metric.ptToHwp((pixelDims.h * 72) / 96);
  } else {
    wHwp = Metric.ptToHwp(img.w);
    hHwp = Metric.ptToHwp(img.h);
  }

  // 가용 너비 초과 방지 (비율 유지)
  if (wHwp > ctx.availableWidth) {
    hHwp = Math.round((hHwp * ctx.availableWidth) / wHwp);
    wHwp = ctx.availableWidth;
  }

  const cx = Math.round(wHwp / 2);
  const cy = Math.round(hHwp / 2);

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
    `<hp:rotationInfo angle="0" centerX="${cx}" centerY="${cy}" rotateimage="1"/>` +
    `<hp:renderingInfo>` +
    `<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>` +
    `<hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>` +
    `<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>` +
    `</hp:renderingInfo>` +
    `<hp:imgRect>` +
    `<hc:pt0 x="0" y="0"/><hc:pt1 x="${wHwp}" y="0"/>` +
    `<hc:pt2 x="${wHwp}" y="${hHwp}"/><hc:pt3 x="0" y="${hHwp}"/>` +
    `</hp:imgRect>` +
    `<hp:imgClip left="0" right="0" top="0" bottom="0"/>` +
    `<hp:inMargin left="0" right="0" top="0" bottom="0"/>` +
    `<hp:imgDim dimwidth="${wHwp}" dimheight="${hHwp}"/>` +
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
  return `<hp:run charPrIDRef="0">${encodeImage(img, ctx)}<hp:t> </hp:t></hp:run>`;
}

// ─── 표(Grid) 인코딩 ─────────────────────────────────────────

function encodeGridPositioned(
  grid: GridNode,
  ctx: HwpxCtx,
  vertPos: number,
  secPr = "",
  hfRun = "",
): { xml: string; nextVertPos: number } {
  const gridXml = buildGridXml(grid, ctx);
  const fs = 1000;
  const sp = 600;
  const vs = fs + sp;
  const linesegXml = buildLineSeg(vertPos, vs, fs, ctx.availableWidth);

  const xml =
    `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">` +
    secPr +
    hfRun +
    `<hp:run charPrIDRef="0">${gridXml}<hp:t> </hp:t></hp:run>` +
    linesegXml +
    `</hp:p>`;

  return { xml, nextVertPos: vertPos + vs };
}

function buildGridXml(grid: GridNode, ctx: HwpxCtx): string {
  const rowCount = grid.kids.length;

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
  const totalW = ctx.availableWidth;
  const colWidths: number[] = [];

  if (grid.props.colWidths && grid.props.colWidths.length === colCount) {
    // pt -> HWPUNIT 변환하여 원본 값 보존
    for (const wPt of grid.props.colWidths) {
      colWidths.push(Metric.ptToHwp(wPt));
    }
  } else {
    // 너비 정보가 없을 때만 균등 배분
    const defW = Math.round(totalW / colCount);
    for (let c = 0; c < colCount; c++) colWidths.push(defW);
  }

  // 본문 너비 초과 시에만 비율 축소 보정
  const rawTotal = colWidths.reduce((s, w) => s + w, 0);
  if (rawTotal > totalW && rawTotal > 0) {
    const scale = totalW / rawTotal;
    for (let i = 0; i < colWidths.length; i++) {
      colWidths[i] = Math.max(100, Math.round(colWidths[i] * scale));
    }
  }
  const actualTotal = colWidths.reduce((s, w) => s + w, 0);

  // 행 높이 계산
  const rowHeights: number[] = [];
  for (let ri = 0; ri < rowCount; ri++) {
    if (
      grid.kids[ri].heightPt != null &&
      (grid.kids[ri].heightPt as number) > 0
    ) {
      rowHeights.push(Metric.ptToHwp(grid.kids[ri].heightPt as number));
    } else {
      let maxH = 0;
      for (let ci = 0; ci < colCount; ci++) {
        const entry = tableMap[ri][ci];
        if (entry?.type === "real") {
          const h = estimateCellHeight(entry.cell!, ctx);
          if (h > maxH) maxH = h;
        }
      }
      rowHeights.push(maxH || Math.round(1000 * 1.6));
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

      const innerW = Math.max(cellW - 282, 100);
      const parasXml =
        cell.kids.length > 0
          ? cell.kids
              .map((p) => encodeParaPositioned(p, ctx, 0, "", innerW).xml)
              .join("")
          : `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="0" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t> </hp:t></hp:run></hp:p>`;

      const subListId = ctx.nextElementId++;
      const vAlign =
        cp.va === "mid" ? "CENTER" : cp.va === "bot" ? "BOTTOM" : "TOP";

      cellsXml +=
        `<hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="${cellBfId}">` +
        `<hp:subList id="${subListId}" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="${vAlign}" ` +
        `linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">` +
        parasXml +
        `</hp:subList>` +
        `<hp:cellAddr colAddr="${ci}" rowAddr="${ri}"/>` +
        `<hp:cellSpan colSpan="${cell.cs}" rowSpan="${cell.rs}"/>` +
        `<hp:cellSz width="${cellW}" height="${rowHeights[ri]}"/>` +
        `<hp:cellMargin left="141" right="141" top="141" bottom="141"/>` +
        `</hp:tc>`;
    }
    rowsXml += `<hp:tr>${cellsXml}</hp:tr>`;
  }

  const headerRow = grid.props.headerRow ? ' repeatHeader="1"' : "";
  return (
    `<hp:tbl id="${ctx.nextElementId++}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE"${headerRow} rowCnt="${rowCount}" colCnt="${colCount}" cellSpacing="0" borderFillIDRef="${tblBfId}" noAdjust="0">` +
    `<hp:sz width="${actualTotal}" widthRelTo="ABSOLUTE" height="${totalH}" heightRelTo="ABSOLUTE" protect="0"/>` +
    `<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>` +
    `<hp:outMargin left="138" right="138" top="138" bottom="138"/>` +
    `<hp:inMargin left="138" right="138" top="138" bottom="138"/>` +
    rowsXml +
    `</hp:tbl>`
  );
}

function estimateCellHeight(cell: CellNode, ctx: HwpxCtx): number {
  const topPad = 141;
  const botPad = 141;
  let h = 0;
  for (const para of cell.kids) {
    const fs = fontSizeForPara(para, ctx);
    const ppId = ctx.paraPrMap.get(paraPrKey(para.props));
    const pp = ppId !== undefined ? ctx.paraPrs[ppId] : null;
    const ls = pp?.lineSpacing ?? 160;
    const before = pp?.prevHwp ?? 0;
    const after = pp?.nextHwp ?? 0;
    h += Math.round((fs * ls) / 100) + before + after;
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
              p.kids.flatMap((k) =>
                k.tag === "span"
                  ? k.kids.flatMap((c) => (c.tag === "txt" ? [c.content] : []))
                  : [],
              ),
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
  s = s.replace(/__EXT_\d+__/g, "");
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
