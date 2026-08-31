import type {
  DocRoot,
  ContentNode,
  ParaNode,
  SpanNode,
  GridNode,
  ImgNode,
  PageNumNode,
} from "../../model/doc-tree";
import type { Outcome } from "../../contract/result";
import type {
  DocMeta,
  PageDims,
  TextProps,
  ParaProps,
  CellProps,
  GridProps,
  Stroke,
  ImgLayout,
  ImgWrap,
  ImgHorzAlign,
  ImgVertAlign,
  ImgHorzRelTo,
  ImgVertRelTo,
} from "../../model/doc-props";
import { A4 } from "../../model/doc-props";
import { succeed, fail } from "../../contract/result";
import {
  buildRoot,
  buildSheet,
  buildPara,
  buildSpan,
  buildImg,
  buildGrid,
  buildRow,
  buildCell,
  buildPb,
} from "../../model/builders";
import { ShieldedParser } from "../../safety/ShieldedParser";
import {
  Metric,
  safeAlign,
  safeFont,
  safeHex,
  safeStrokeHwpx,
} from "../../safety/StyleBridge";
import { ArchiveKit } from "../../toolkit/ArchiveKit";
import { XmlKit } from "../../toolkit/XmlKit";
import { TextKit } from "../../toolkit/TextKit";
import { inferColumnWidths } from "../../toolkit/TableGeometry";
import { registry } from "../../pipeline/registry";
import { BaseDecoder } from "../../core/BaseDecoder";
import { HWPX_MIME_TYPE } from "../../encoders/hwpx/constants";

interface BorderFillInfo {
  stroke?: Stroke; // uniform fallback (used when all sides are the same)
  top?: Stroke;
  right?: Stroke;
  bottom?: Stroke;
  left?: Stroke;
  bgColor?: string;
}

interface CharPrInfo {
  b?: boolean;
  i?: boolean;
  u?: boolean;
  s?: boolean;
  pt?: number;
  color?: string;
  font?: string;
  bg?: string;
}

interface ParaPrInfo {
  align?: string;
  indentPt?: number; // hc:left → 문단 전체 왼쪽 여백
  indentRightPt?: number; // hc:right → 문단 전체 오른쪽 여백
  firstLineIndentPt?: number; // hc:indent → 첫 줄 들여쓰기 (양수=들여쓰기, 음수=내어쓰기)
  spaceBefore?: number;
  spaceAfter?: number;
  lineHeight?: number;
  lineHeightFixed?: number; // FIXED 행 높이 (pt)
  lineHeightRule?: "exact" | "atLeast";
}

interface DecCtx {
  files: Map<string, Uint8Array>;
  shield: ShieldedParser;
  borderFills: Map<number, BorderFillInfo>;
  charPrs: Map<number, CharPrInfo>;
  paraPrs: Map<number, ParaPrInfo>;
  warns: string[];
}

export class HwpxDecoder extends BaseDecoder {
  protected getFormat(): string {
    return "hwpx";
  }
  protected getAliases(): string[] {
    return [HWPX_MIME_TYPE, "application/hwp+zip"];
  }

  async decode(data: Uint8Array): Promise<Outcome<DocRoot>> {
    const shield = new ShieldedParser();
    const warns: string[] = [];

    try {
      const files = await ArchiveKit.unzip(data);

      const sectionFiles: Uint8Array[] = [];
      for (let i = 0; ; i++) {
        const sec =
          files.get(`Contents/section${i}.xml`) ?? files.get(`section${i}.xml`);
        if (!sec) break;
        sectionFiles.push(sec);
      }
      if (sectionFiles.length === 0) {
        const fallback = findSectionFile(files);
        if (fallback) sectionFiles.push(fallback);
      }

      if (sectionFiles.length === 0)
        return fail("HWPX: No section files found");

      const headXml =
        files.get("Contents/header.xml") ?? files.get("header.xml");

      let meta: DocMeta = {};
      let dims: PageDims = { ...A4 };
      let borderFills = new Map<number, BorderFillInfo>();
      let charPrs = new Map<number, CharPrInfo>();
      let paraPrs = new Map<number, ParaPrInfo>();

      if (headXml) {
        try {
          const headStr = TextKit.decode(headXml);
          const headObj: any = await XmlKit.parseStrict(headStr);
          if (headObj) {
            meta = extractMeta(headObj);
            dims = extractDims(headObj) ?? dims;
            borderFills = extractBorderFills(headObj);
            charPrs = extractCharPrs(headObj);
            paraPrs = extractParaPrs(headObj);
          }
        } catch {
          // header parse failure is non-fatal
        }
      }

      const ctx: DecCtx = {
        files,
        shield,
        borderFills,
        charPrs,
        paraPrs,
        warns,
      };

      const allSections: any[] = [];
      for (const secFile of sectionFiles) {
        const bodyStr = TextKit.decode(secFile);
        const bodyObj: any = await XmlKit.parseStrict(bodyStr);
        allSections.push(...normalizeSections(bodyObj));
      }

      const kids = shield.guardAll(
        allSections,
        (sec: any) => decodeSection(sec, dims, ctx),
        () => buildSheet([buildPara([buildSpan("[섹션 파싱 실패]")])], dims),
        "hwpx:section",
      );

      warns.push(...shield.flush());
      return succeed(buildRoot(meta, kids), warns);
    } catch (e: any) {
      warns.push(...shield.flush());
      return fail(`HWPX decode error: ${e?.message ?? String(e)}`, warns);
    }
  }
}

// ─── helpers ────────────────────────────────────────────────

function findSectionFile(
  files: Map<string, Uint8Array>,
): Uint8Array | undefined {
  for (const [key, val] of files) {
    if (key.toLowerCase().includes("section") && key.endsWith(".xml"))
      return val;
  }
  return undefined;
}

function normalizeSections(bodyObj: any): any[] {
  // <hs:sec> (real HWPX), <hp:SEC> (legacy)
  if (bodyObj?.["hs:sec"]) return toArr(bodyObj["hs:sec"]);
  if (bodyObj?.["hp:SEC"]) return toArr(bodyObj["hp:SEC"]);

  const root = bodyObj?.["hp:HWPML"] ?? bodyObj?.HWPML ?? bodyObj;
  const body =
    root?.["hp:BODY"]?.[0] ??
    root?.BODY?.[0] ??
    root?.["hp:BODY"] ??
    root?.BODY;
  if (!body) return [bodyObj];
  const sections = body?.["hp:SECTION"] ?? body?.SECTION ?? [];
  return Array.isArray(sections) ? sections : [sections];
}

// Get a tag regardless of namespace/case variations
function getTag(obj: any, ...names: string[]): any[] {
  for (const n of names) {
    const v = obj?.[n];
    if (v != null) return toArr(v);
  }
  return [];
}

function extractMeta(headObj: any): DocMeta {
  try {
    // Support both <hh:HEAD> and <hh:head>
    const root =
      headObj?.["hh:head"]?.[0] ??
      headObj?.["hh:HEAD"]?.[0] ??
      headObj?.HEAD?.[0] ??
      headObj;
    const info = root?.["hh:DOCSUMMARY"]?.[0] ?? root?.DOCSUMMARY?.[0];
    if (!info) return {};
    const a = (k: string) =>
      info?.[`hh:${k}`]?.[0]?._text ?? info?.[k]?.[0]?._text ?? "";
    return {
      title: a("TITLE") || undefined,
      author: a("AUTHOR") || undefined,
      subject: a("SUBJECT") || undefined,
    };
  } catch {
    return {};
  }
}

function extractDims(headObj: any): PageDims | null {
  try {
    const root =
      headObj?.["hh:head"]?.[0] ??
      headObj?.["hh:HEAD"]?.[0] ??
      headObj?.HEAD?.[0] ??
      headObj;

    const modernSecPr =
      root?.["hh:secPrList"]?.[0]?.["hh:secPr"]?.[0] ??
      root?.["hh:SECPRLST"]?.[0]?.["hh:SECPR"]?.[0];
    const modernPagePr =
      modernSecPr?.["hh:pagePr"]?.[0]?._attr ??
      modernSecPr?.["hh:PAGEPR"]?.[0]?._attr;
    if (modernPagePr) {
      const margin =
        modernSecPr?.["hh:pagePr"]?.[0]?.["hh:margin"]?.[0]?._attr ??
        modernSecPr?.["hh:PAGEPR"]?.[0]?.["hh:MARGIN"]?.[0]?._attr ??
        {};
      let ew = Number(modernPagePr.width ?? modernPagePr.Width ?? 59528);
      let eh = Number(modernPagePr.height ?? modernPagePr.Height ?? 84188);
      const landscape = String(modernPagePr.landscape ?? "").toUpperCase();
      if ((landscape === "NARROWLY" || landscape === "LANDSCAPE") && ew < eh) {
        [ew, eh] = [eh, ew];
      }
      const mt = Number(margin.top ?? margin.TopMargin ?? 5670);
      const mb = Number(margin.bottom ?? margin.BottomMargin ?? 4252);
      const ml = Number(margin.left ?? margin.LeftMargin ?? 8504);
      const mr = Number(margin.right ?? margin.RightMargin ?? 8504);
      const header = Number(margin.header ?? margin.HeaderMargin ?? 0);
      const footer = Number(margin.footer ?? margin.FooterMargin ?? 0);
      return {
        wPt: Metric.hwpToPt(ew),
        hPt: Metric.hwpToPt(eh),
        mt: Metric.hwpToPt(mt),
        mb: Metric.hwpToPt(mb),
        ml: Metric.hwpToPt(ml),
        mr: Metric.hwpToPt(mr),
        headerPt: Metric.hwpToPt(Math.max(0, header)),
        footerPt: Metric.hwpToPt(Math.max(0, footer)),
        orient: ew > eh ? "landscape" : "portrait",
      };
    }

    const refList =
      root?.["hh:refList"]?.[0] ??
      root?.["hh:REFLIST"]?.[0] ??
      root?.REFLIST?.[0];
    if (!refList) return null;

    const secPrList =
      refList?.["hh:SECPRLST"]?.[0]?.["hh:SECPR"] ??
      refList?.SECPRLST?.[0]?.SECPR;
    const sec = Array.isArray(secPrList) ? secPrList[0] : secPrList;
    if (!sec) return null;

    const pa =
      sec?.["hh:PAGEPROPERTY"]?.[0]?._attr ?? sec?.PAGEPROPERTY?.[0]?._attr;
    if (!pa) return null;

    const ew = Number(pa.Width ?? 59528);
    const eh = Number(pa.Height ?? 84188);
    const mt = Number(pa.TopMargin ?? 5670);
    const mb = Number(pa.BottomMargin ?? 4252);
    const ml = Number(pa.LeftMargin ?? 8504);
    const mr = Number(pa.RightMargin ?? 8504);
    const header = Number(pa.HeaderMargin ?? 0);
    const footer = Number(pa.FooterMargin ?? 0);
    return {
      wPt: Metric.hwpToPt(ew),
      hPt: Metric.hwpToPt(eh),
      mt: Metric.hwpToPt(mt),
      mb: Metric.hwpToPt(mb),
      ml: Metric.hwpToPt(ml),
      mr: Metric.hwpToPt(mr),
      headerPt: Metric.hwpToPt(Math.max(0, header)),
      footerPt: Metric.hwpToPt(Math.max(0, footer)),
      orient: ew > eh ? "landscape" : "portrait",
    };
  } catch {
    return null;
  }
}

function extractBorderFills(headObj: any): Map<number, BorderFillInfo> {
  const map = new Map<number, BorderFillInfo>();
  try {
    const root =
      headObj?.["hh:head"]?.[0] ??
      headObj?.["hh:HEAD"]?.[0] ??
      headObj?.HEAD?.[0] ??
      headObj;
    const refList =
      root?.["hh:refList"]?.[0] ??
      root?.["hh:REFLIST"]?.[0] ??
      root?.REFLIST?.[0];
    if (!refList) return map;

    const bfList =
      refList?.["hh:borderFills"]?.[0] ??
      refList?.["hh:BORDERFILLLIST"]?.[0] ??
      refList?.BORDERFILLLIST?.[0];
    if (!bfList) return map;

    const bfs = getTag(bfList, "hh:borderFill", "hh:BORDERFILL");
    for (const bf of bfs) {
      const attr = bf?._attr ?? {};
      const id = Number(attr.id ?? 0);
      if (id === 0) continue;

      const info: BorderFillInfo = {};

      // Helper: parse a border element into a Stroke
      const parseBorderEl = (el: any): Stroke | undefined => {
        if (!el) return undefined;
        const a = el?._attr ?? {};
        const mmVal = parseFloat(a.width) || undefined;
        const hwpVal = mmVal != null ? mmVal * 2.835 * 100 : undefined;
        return safeStrokeHwpx(a.type, hwpVal, a.color);
      };

      // Parse all four sides
      const topEl =
        bf?.["hh:topBorder"]?.[0] ?? bf?.["hh:top"]?.[0] ?? bf?.top?.[0];
      const rightEl =
        bf?.["hh:rightBorder"]?.[0] ?? bf?.["hh:right"]?.[0] ?? bf?.right?.[0];
      const bottomEl =
        bf?.["hh:bottomBorder"]?.[0] ??
        bf?.["hh:bottom"]?.[0] ??
        bf?.bottom?.[0];
      const leftEl =
        bf?.["hh:leftBorder"]?.[0] ?? bf?.["hh:left"]?.[0] ?? bf?.left?.[0];

      info.top = parseBorderEl(topEl);
      info.right = parseBorderEl(rightEl);
      info.bottom = parseBorderEl(bottomEl);
      info.left = parseBorderEl(leftEl);

      // Set uniform stroke fallback = top border (for defaultStroke etc.)
      info.stroke = info.top ?? info.left ?? info.right ?? info.bottom;

      // Parse fill (real HWPX uses hc:fillBrush, not hh:fillBrush)
      const fillBrush =
        bf?.["hc:fillBrush"]?.[0] ??
        bf?.["hh:fillBrush"]?.[0] ??
        bf?.["hh:fill"]?.[0] ??
        bf?.fill?.[0] ??
        bf?.fillBrush?.[0];
      if (fillBrush) {
        const winBrush =
          fillBrush?.["hc:winBrush"]?.[0]?._attr ??
          fillBrush?.["hh:winBrush"]?.[0]?._attr ??
          fillBrush?.winBrush?.[0]?._attr;
        if (winBrush?.faceColor && winBrush.faceColor !== "none") {
          info.bgColor = safeHex(winBrush.faceColor);
        }
      }

      map.set(id, info);
    }
  } catch {
    /* non-fatal */
  }
  return map;
}

function buildFontIdMap(headObj: any): Map<number, string> {
  const fontMap = new Map<number, string>();
  try {
    const root =
      headObj?.["hh:head"]?.[0] ??
      headObj?.["hh:HEAD"]?.[0] ??
      headObj?.HEAD?.[0] ??
      headObj;
    const refList =
      root?.["hh:refList"]?.[0] ??
      root?.["hh:REFLIST"]?.[0] ??
      root?.REFLIST?.[0];
    if (!refList) return fontMap;

    const fontfaces =
      refList?.["hh:fontfaces"]?.[0] ?? refList?.["hh:FONTFACES"]?.[0];
    if (!fontfaces) return fontMap;

    // Try each fontface group (HANGUL, LATIN, etc.) — use the first group that has entries
    const ffGroups = getTag(fontfaces, "hh:fontface", "hh:FONTFACE");
    for (const ff of ffGroups) {
      const fonts = getTag(ff, "hh:font", "hh:FONT");
      for (const font of fonts) {
        const fa = font?._attr ?? {};
        const fid = Number(fa.id ?? -1);
        const name = fa.face ?? fa.name ?? fa.Face ?? "";
        if (fid >= 0 && name && !fontMap.has(fid)) fontMap.set(fid, name);
      }
      if (fontMap.size > 0) break; // use first group (usually HANGUL)
    }
  } catch {
    /* non-fatal */
  }
  return fontMap;
}

function extractCharPrs(headObj: any): Map<number, CharPrInfo> {
  const map = new Map<number, CharPrInfo>();
  try {
    const root =
      headObj?.["hh:head"]?.[0] ??
      headObj?.["hh:HEAD"]?.[0] ??
      headObj?.HEAD?.[0] ??
      headObj;
    const refList =
      root?.["hh:refList"]?.[0] ??
      root?.["hh:REFLIST"]?.[0] ??
      root?.REFLIST?.[0];
    if (!refList) return map;

    // Build font id → name map from fontfaces
    const fontIdMap = buildFontIdMap(headObj);

    const cpList =
      refList?.["hh:charProperties"]?.[0] ??
      refList?.["hh:CHARPROPERTIES"]?.[0];
    if (!cpList) return map;

    const cps = getTag(cpList, "hh:charPr", "hh:CHARPR");
    for (const cp of cps) {
      const attr = cp?._attr ?? {};
      const id = Number(attr.id ?? -1);
      if (id < 0) continue;

      const info: CharPrInfo = {};

      // height → pt
      if (attr.height) info.pt = Metric.hHeightToPt(Number(attr.height));

      // textColor
      if (attr.textColor) info.color = normalizeHwpxTextColor(attr.textColor);

      // bold
      if (cp?.["hh:bold"]?.[0] != null) info.b = true;

      // italic
      if (cp?.["hh:italic"]?.[0] != null) info.i = true;

      // underline
      const ulAttr = cp?.["hh:underline"]?.[0]?._attr;
      if (ulAttr?.type && ulAttr.type !== "NONE") info.u = true;

      // strikeout — shape="3D" is default "no strikeout" in real HWPX; only SOLID/etc means active
      const stAttr = cp?.["hh:strikeout"]?.[0]?._attr;
      if (stAttr?.shape && stAttr.shape !== "NONE" && stAttr.shape !== "3D")
        info.s = true;

      // font name — resolve from fontRef.hangul → fontfaces
      const fontRefAttr =
        cp?.["hh:fontRef"]?.[0]?._attr ?? cp?.["hh:FONTREF"]?.[0]?._attr;
      if (fontRefAttr) {
        const fid = Number(
          fontRefAttr.hangul ?? fontRefAttr.latin ?? fontRefAttr.Hangul ?? 0,
        );
        const name = fontIdMap.get(fid);
        if (name) info.font = safeFont(name);
      }

      map.set(id, info);
    }
  } catch {
    /* non-fatal */
  }
  return map;
}

function extractParaPrs(headObj: any): Map<number, ParaPrInfo> {
  const map = new Map<number, ParaPrInfo>();
  try {
    const root =
      headObj?.["hh:head"]?.[0] ??
      headObj?.["hh:HEAD"]?.[0] ??
      headObj?.HEAD?.[0] ??
      headObj;
    const refList =
      root?.["hh:refList"]?.[0] ??
      root?.["hh:REFLIST"]?.[0] ??
      root?.REFLIST?.[0];
    if (!refList) return map;

    const ppList =
      refList?.["hh:paraProperties"]?.[0] ??
      refList?.["hh:PARAPROPERTIES"]?.[0];
    if (!ppList) return map;

    const pps = getTag(ppList, "hh:paraPr", "hh:PARAPR");
    for (const pp of pps) {
      const attr = pp?._attr ?? {};
      const id = Number(attr.id ?? -1);
      if (id < 0) continue;

      const alignNode =
        pp?.["hh:align"]?.[0]?._attr ?? pp?.["hh:ALIGN"]?.[0]?._attr;
      const align = alignNode?.horizontal ?? alignNode?.Horizontal;

      // Read margin and lineSpacing from direct child OR hp:switch > hp:default/hp:case
      let marginEl = pp?.["hh:margin"]?.[0] ?? null;
      let lineSpEl = pp?.["hh:lineSpacing"]?.[0] ?? null;
      if (!marginEl) {
        const sw = pp?.["hp:switch"]?.[0];
        const container = sw?.["hp:default"]?.[0] ?? sw?.["hp:case"]?.[0];
        marginEl = container?.["hh:margin"]?.[0] ?? null;
        lineSpEl = lineSpEl ?? container?.["hh:lineSpacing"]?.[0] ?? null;
      }

      let indentPt: number | undefined;
      let indentRightPt: number | undefined;
      let firstLineIndentPt: number | undefined;
      let spaceBefore: number | undefined;
      let spaceAfter: number | undefined;
      let lineHeight: number | undefined;
      let lineHeightFixed: number | undefined;

      if (marginEl) {
        // OWPML §7.5.4.4: hc:left=전체왼쪽여백, hc:right=전체오른쪽여백,
        // hc:indent=첫줄들여쓰기(양수)/내어쓰기(음수)
        // hc:intent는 자사 인코더가 생성하는 오기 표기로, hc:indent와 동일하게 처리
        const leftEl = marginEl?.["hc:left"]?.[0];
        const rightEl = marginEl?.["hc:right"]?.[0];
        const indentEl =
          marginEl?.["hc:intent"]?.[0] ?? marginEl?.["hc:indent"]?.[0];
        const prevEl = marginEl?.["hc:prev"]?.[0];
        const nextEl = marginEl?.["hc:next"]?.[0];

        const leftVal = Number(leftEl?._attr?.value ?? 0);
        const rightVal = Number(rightEl?._attr?.value ?? 0);
        const indentVal = Number(indentEl?._attr?.value ?? 0);
        const prevVal = Number(prevEl?._attr?.value ?? 0);
        const nextVal = Number(nextEl?._attr?.value ?? 0);

        // HWPX paraPr 여백 값은 HWP 바이너리 PARA_SHAPE와 동일하게
        // 물리 HWPUNIT의 2배로 기록된다.
        if (leftVal !== 0) indentPt = Metric.hwpToPt(leftVal / 2);
        if (rightVal !== 0) indentRightPt = Metric.hwpToPt(rightVal / 2);
        if (indentVal !== 0) firstLineIndentPt = Metric.hwpToPt(indentVal / 2);
        if (prevVal > 0) spaceBefore = Metric.hwpToPt(prevVal / 2);
        if (nextVal > 0) spaceAfter = Metric.hwpToPt(nextVal / 2);
      }

      if (lineSpEl) {
        const lsAttr = lineSpEl._attr ?? {};
        const lsType = lsAttr.type ?? "PERCENT";
        const lsVal = Number(lsAttr.value ?? 160);
        // OWPML §7.5.4.6: PERCENT(비율), FIXED(고정), BETWEEN_LINES(줄간격), AT_LEAST(최소)
        if (lsType === "PERCENT" && lsVal > 0) {
          lineHeight = lsVal / 100;
        } else if ((lsType === "FIXED" || lsType === "AT_LEAST") && lsVal > 0) {
          // FIXED/AT_LEAST도 paraPr 저장 단위(물리 HWPUNIT의 2배)를 사용한다.
          lineHeightFixed = Metric.hwpToPt(lsVal / 2);
        }
      }

      map.set(id, {
        align,
        indentPt,
        indentRightPt,
        firstLineIndentPt,
        spaceBefore: spaceBefore ?? 0,
        spaceAfter: spaceAfter ?? 0,
        lineHeight: lineHeightFixed === undefined ? (lineHeight ?? 1.6) : undefined,
        lineHeightFixed,
        lineHeightRule:
          lineHeightFixed !== undefined
            ? (lineSpEl?._attr?.type === "AT_LEAST" ? "atLeast" : "exact")
            : undefined,
      });
    }
  } catch {
    /* non-fatal */
  }
  return map;
}

// ─── Section decoding ──────────────────────────────────────

function addParaItems(p: any, items: { type: string; node: any }[]): void {
  // Check if this paragraph contains a table in its runs
  const runs = getTag(p, "hp:run", "hp:RUN");
  for (const run of runs) {
    const tbls = getTag(run, "hp:tbl", "hp:TABLE");
    if (tbls.length > 0) {
      for (const tbl of tbls) {
        items.push({ type: "table", node: tbl });
      }
    }
  }
  // A table-only HWPX paragraph is still its positioning/flow anchor.  Hancom
  // exports the table first and keeps this paragraph immediately afterwards.
  items.push({ type: "para", node: p });
}

function decodeSection(sec: any, dims: PageDims, ctx: DecCtx) {
  // Try to extract dims from first paragraph's secPr
  const firstParas = getTag(sec, "hp:p", "hp:P");
  const pageDims = extractSectionDims(sec) ?? extractSecPrDims(firstParas[0]) ?? dims;

  // Build items list preserving document order via _childOrder
  const items: { type: string; node: any }[] = [];
  const paras = getTag(sec, "hp:p", "hp:P");
  const tbls = getTag(sec, "hp:tbl", "hp:TABLE");

  const childOrder = sec?.["_childOrder"] as string[] | undefined;

  if (Array.isArray(childOrder)) {
    let pi = 0;
    let ti = 0;
    for (const tag of childOrder) {
      if ((tag === "hp:p" || tag === "hp:P") && pi < paras.length) {
        addParaItems(paras[pi++], items);
      } else if ((tag === "hp:tbl" || tag === "hp:TABLE") && ti < tbls.length) {
        items.push({ type: "table", node: tbls[ti++] });
      }
    }
    // Append any remaining (fallback)
    while (pi < paras.length) addParaItems(paras[pi++], items);
    while (ti < tbls.length) items.push({ type: "table", node: tbls[ti++] });
  } else {
    // No order info — process paragraphs sequentially (fallback to previous logic)
    for (const p of paras) addParaItems(p, items);
    // Note: direct tables are appended after paras in this fallback
    for (const t of tbls) items.push({ type: "table", node: t });
  }

  const kids: ContentNode[] = ctx.shield.guardAll(
    items,
    (item: any) => {
      if (item.type === "table") {
        try {
          const { value } = ctx.shield.guardGrid(
            item.node,
            (n) => decodeGrid(n, ctx),
            (n) => decodeGridSimple(n, ctx),
            (n) => decodeGridFlat(n),
            (n) => decodeGridText(n) as unknown as GridNode,
            "hwpx:table",
          );
          return value;
        } catch {
          return buildPara([buildSpan("[표 파싱 실패]")]);
        }
      }
      return decodePara(item.node, ctx);
    },
    () => buildPara([buildSpan("[파싱 실패]")]),
    "hwpx:content",
  );

  // Decode header/footer
  const headerParas = decodeHeaderFooter(sec, "header", ctx);
  const footerParas = decodeHeaderFooter(sec, "footer", ctx);

  return buildSheet(kids.filter(Boolean) as ContentNode[], pageDims, {
    headers: { default: headerParas },
    footers: { default: footerParas },
  });
}

function parseSecPrDims(secPr: any): PageDims | null {
  const pagePr =
    secPr?.["hp:pagePr"]?.[0]?._attr ?? secPr?.["hp:PAGEPR"]?.[0]?._attr;
  if (!pagePr) return null;
  const margin =
    secPr?.["hp:pagePr"]?.[0]?.["hp:margin"]?.[0]?._attr ??
    secPr?.["hp:PAGEPR"]?.[0]?.["hp:MARGIN"]?.[0]?._attr ??
    {};
  let pw = Number(pagePr.width ?? 59528);
  let ph = Number(pagePr.height ?? 84188);
  const landscape = String(pagePr.landscape ?? "").toUpperCase();
  if ((landscape === "NARROWLY" || landscape === "LANDSCAPE") && pw < ph) {
    [pw, ph] = [ph, pw];
  }
  const mt = Number(margin.top ?? 5670);
  const mb = Number(margin.bottom ?? 4252);
  const ml = Number(margin.left ?? 8504);
  const mr = Number(margin.right ?? 8504);
  const header = Number(margin.header ?? 0);
  const footer = Number(margin.footer ?? 0);
  return {
    wPt: Metric.hwpToPt(pw),
    hPt: Metric.hwpToPt(ph),
    mt: Metric.hwpToPt(mt),
    mb: Metric.hwpToPt(mb),
    ml: Metric.hwpToPt(ml),
    mr: Metric.hwpToPt(mr),
    headerPt: Metric.hwpToPt(Math.max(0, header)),
    footerPt: Metric.hwpToPt(Math.max(0, footer)),
    orient: pw > ph ? "landscape" : "portrait",
  };
}

function extractSectionDims(sec: any): PageDims | null {
  try {
    const secPr = sec?.["hp:secPr"]?.[0] ?? sec?.["hp:SECPR"]?.[0];
    return secPr ? parseSecPrDims(secPr) : null;
  } catch {
    return null;
  }
}

function extractSecPrDims(p: any): PageDims | null {
  if (!p) return null;
  try {
    // Primary: hp:secPr is a DIRECT child of hp:p (as generated by HwpxEncoder)
    const secPrDirect = p?.["hp:secPr"]?.[0] ?? p?.["hp:SECPR"]?.[0];
    if (secPrDirect) {
      const dims = parseSecPrDims(secPrDirect);
      if (dims) return dims;
    }
    // Fallback: legacy format may nest hp:secPr inside hp:run
    const runs = getTag(p, "hp:run", "hp:RUN");
    for (const run of runs) {
      const secPr = run?.["hp:secPr"]?.[0] ?? run?.["hp:SECPR"]?.[0];
      if (!secPr) continue;
      const dims = parseSecPrDims(secPr);
      if (dims) return dims;
    }
  } catch {
    /* ignore */
  }
  return null;
}

function decodeHeaderFooter(
  sec: any,
  kind: "header" | "footer",
  ctx: DecCtx,
): ParaNode[] | undefined {
  try {
    const hf =
      sec?.["hp:headerFooter"]?.[0] ??
      sec?.["hp:HEADERFOOTER"]?.[0] ??
      sec?.headerFooter?.[0] ??
      sec?.HEADERFOOTER?.[0];
    if (!hf) return undefined;

    const part =
      hf?.["hp:" + kind]?.[0] ??
      hf?.["hp:" + kind.toUpperCase()]?.[0] ??
      hf?.[kind]?.[0] ??
      hf?.[kind.toUpperCase()]?.[0];
    if (!part) return undefined;

    const paras = getTag(part, "hp:p", "hp:P");
    if (paras.length === 0) return undefined;

    return paras.map((p: any) => decodePara(p, ctx));
  } catch {
    return undefined;
  }
}

// ─── Paragraph & run decoding ──────────────────────────────

function decodePara(p: any, ctx: DecCtx): ParaNode {
  const pAttr = p?._attr ?? {};
  const paraPrIdRef = Number(pAttr.paraPrIDRef ?? -1);
  const styleIdRef = Number(pAttr.styleIDRef ?? pAttr.styleIdRef ?? pAttr.styleID ?? pAttr.styleId);

  // Resolve paraPr from IDRef or inline
  let align: string | undefined;
  const paraPrDef = ctx.paraPrs.get(paraPrIdRef);
  if (paraPrDef?.align) align = paraPrDef.align;

  // Check inline PARAPR too
  const inlineParaPr =
    p?.["hp:PARAPR"]?.[0] ?? p?.["hp:paraPr"]?.[0] ?? p?.PARAPR?.[0];
  if (inlineParaPr) {
    const alignNode =
      inlineParaPr?.["hp:ALIGN"]?.[0]?._attr ??
      inlineParaPr?.["hp:align"]?.[0]?._attr ??
      inlineParaPr?.ALIGN?.[0]?._attr;
    if (alignNode?.Type) align = alignNode.Type;
    if (alignNode?.horizontal) align = alignNode.horizontal;
  }

  const inlineAttr = inlineParaPr?._attr ?? {};
  const props: ParaProps = {
    align: safeAlign(align === "JUSTIFY" ? "LEFT" : align),
    spaceBefore: 0,
    spaceAfter: 0,
    lineHeight: 1.6,
  };
  if (Number.isFinite(styleIdRef) && styleIdRef >= 0) props.hwpStyleId = styleIdRef;

  // Apply spacing/indent/lineHeight from paraPr definition
  if (paraPrDef) {
    if (paraPrDef.indentPt !== undefined) props.indentPt = paraPrDef.indentPt;
    if (paraPrDef.indentRightPt !== undefined)
      props.indentRightPt = paraPrDef.indentRightPt;
    if (paraPrDef.firstLineIndentPt !== undefined)
      props.firstLineIndentPt = paraPrDef.firstLineIndentPt;
    if (paraPrDef.spaceBefore !== undefined)
      props.spaceBefore = paraPrDef.spaceBefore;
    if (paraPrDef.spaceAfter !== undefined)
      props.spaceAfter = paraPrDef.spaceAfter;
    if (paraPrDef.lineHeight !== undefined)
      props.lineHeight = paraPrDef.lineHeight;
    if (paraPrDef.lineHeightFixed !== undefined) {
      props.lineHeight = undefined;
      props.lineHeightFixed = paraPrDef.lineHeightFixed;
    }
    if (paraPrDef.lineHeightRule !== undefined)
      props.lineHeightRule = paraPrDef.lineHeightRule;
  }

  // List support (from inline attr)
  if (inlineAttr.listType) {
    props.listOrd =
      inlineAttr.listType === "DIGIT" || inlineAttr.listType === "DECIMAL";
    props.listLv = Number(inlineAttr.listLevel ?? 0);
  }

  const runs = getTag(p, "hp:run", "hp:RUN");
  const kids: (SpanNode | ImgNode)[] = [];

  // Helper: collect hp:pic elements from a container (direct child OR inside hp:ctrl)
  const collectPics = (container: any): any[] => {
    const direct = getTag(container, "hp:pic", "hp:PIC");
    const ctrls = getTag(container, "hp:ctrl", "hp:CTRL");
    const nested = ctrls.flatMap((c: any) => getTag(c, "hp:pic", "hp:PIC"));
    return [...direct, ...nested];
  };

  // Images that are direct children of <hp:p> (common in table cells and floats)
  for (const pic of collectPics(p)) {
    const img = decodePic(pic, ctx);
    if (img) kids.push(img);
  }

  for (const run of runs) {
    // Images: directly in run OR in run→ctrl (both patterns appear in practice)
    for (const pic of collectPics(run)) {
      const img = decodePic(pic, ctx);
      if (img) kids.push(img);
    }

    // Page number
    const pageNums = getTag(run, "hp:pageNum", "hp:PAGENUM");
    if (pageNums.length > 0) {
      const pn = pageNums[0]?._attr ?? {};
      const fmt =
        pn.formatType === "ROMAN_LOWER"
          ? ("roman" as const)
          : pn.formatType === "ROMAN_UPPER"
            ? ("romanCaps" as const)
            : ("decimal" as const);
      const pageNumNode: PageNumNode = { tag: "pagenum", format: fmt };
      const spanProps = resolveCharPr(run, ctx);
      kids.push({ tag: "span", props: spanProps, kids: [pageNumNode] });
      continue;
    }

    // Text
    const runPics = collectPics(run);
    const textNodes = getTag(run, "hp:t", "hp:T", "hp:CHAR");
    const content = textNodes
      .map((t: any) => {
        const val =
          typeof t === "string" ? t : (t?._text ?? t?._ ?? t?.["#text"] ?? "");
        return val.replace(/__EXT_\d+(?:_W\d+_H\d+)?__/g, "");
      })
      .join("");

    // Skip empty secPr-only runs that produced no images
    if (
      content === "" &&
      (run?.["hp:secPr"]?.[0] || run?.["hp:SECPR"]?.[0]) &&
      runPics.length === 0 &&
      pageNums.length === 0
    )
      continue;

    // Only push text span when there's actual content and no image already pushed for this run
    if (content !== "" || (runPics.length === 0 && pageNums.length === 0)) {
      const spanProps = content === "" ? {} : resolveCharPr(run, ctx);
      kids.push(buildSpan(content, spanProps));
    }
  }

  // pageBreak="1" → prepend a pb node in its own span
  if (pAttr.pageBreak === "1") {
    kids.unshift({ tag: "span", props: {}, kids: [buildPb()] });
  }

  return buildPara(kids.filter(Boolean) as ParaNode["kids"], props);
}

function resolveCharPr(run: any, ctx: DecCtx): TextProps {
  const runAttr = run?._attr ?? {};
  const charPrIdRef = Number(runAttr.charPrIDRef ?? runAttr.CharPrIDRef ?? -1);

  // IDRef로 먼저 조회
  const def = ctx.charPrs.get(charPrIdRef);
  if (def) {
    return {
      b: def.b,
      i: def.i,
      u: def.u,
      s: def.s,
      pt: def.pt,
      color: def.color,
      font: def.font,
      bg: def.bg,
    };
  }

  // 인라인 CHARPR fallback — 대소문자 모두 시도
  const inlinePr =
    run?.["hp:CHARPR"]?.[0] ??
    run?.["hp:charPr"]?.[0] ??
    run?.CHARPR?.[0] ??
    run?.charPr?.[0];
  const ca = inlinePr?._attr ?? {};

  const bVal = ca.Bold ?? ca.bold ?? ca.B ?? "";
  const iVal = ca.Italic ?? ca.italic ?? ca.I ?? "";
  const uVal = ca.Underline ?? ca.underline ?? "";
  const sVal = ca.Strikeout ?? ca.strikeout ?? "";
  const fontName =
    ca.FontName ?? ca.fontName ?? ca.FaceNameHangul ?? ca.faceNameHangul ?? "";
  const heightVal = ca.Height ?? ca.height ?? "";

  return {
    b: bVal === "1" || bVal === "true" || bVal === "True" || undefined,
    i: iVal === "1" || iVal === "true" || iVal === "True" || undefined,
    u: uVal && uVal !== "NONE" ? true : undefined,
    s: sVal && sVal !== "NONE" && sVal !== "3D" ? true : undefined,
    font: fontName ? safeFont(fontName) : undefined,
    pt: heightVal ? Metric.hHeightToPt(Number(heightVal)) : undefined,
    color: normalizeHwpxTextColor(ca.TextColor ?? ca.textColor),
    bg: safeHex(ca.BgColor ?? ca.bgColor),
  };
}

function normalizeHwpxTextColor(raw: string | number | null | undefined): string | undefined {
  const color = safeHex(raw);
  return color === "000000" ? undefined : color;
}

// ─── Image decoding ────────────────────────────────────────

function decodePic(pic: any, ctx: DecCtx): ImgNode | null {
  try {
    const szAttr = pic?.["hp:sz"]?.[0]?._attr ?? pic?.sz?.[0]?._attr ?? {};
    const w = Metric.hwpToPt(Number(szAttr.width ?? 0));
    const h = Metric.hwpToPt(Number(szAttr.height ?? 0));

    // Try multiple tag patterns for image reference
    const imgNode =
      pic?.["hp:img"]?.[0]?._attr ??
      pic?.["hc:img"]?.[0]?._attr ??
      pic?.img?.[0]?._attr ??
      {};
    const binRef = imgNode.binaryItemIDRef ?? imgNode.BinaryItemIDRef;
    if (!binRef) return null;

    // Find binary data
    let imgData: Uint8Array | undefined;
    for (const [key, val] of ctx.files) {
      if (
        key.includes(binRef) ||
        key.toLowerCase().includes(binRef.toLowerCase())
      ) {
        imgData = val;
        break;
      }
    }
    if (!imgData) return null;

    const ext = binRef.split(".").pop()?.toLowerCase() ?? "png";
    const mimeMap: Record<string, ImgNode["mime"]> = {
      png: "image/png",
      jpg: "image/jpeg",
      jpeg: "image/jpeg",
      gif: "image/gif",
      bmp: "image/bmp",
      wmf: "image/x-wmf",
      emf: "image/x-emf",
    };

    // ── hp:pos에서 layout 추출 ───────────────────────────────
    const posAttr = pic?.["hp:pos"]?.[0]?._attr ?? pic?.pos?.[0]?._attr ?? {};
    const layout = extractHwpxLayout(posAttr, pic);

    return buildImg(
      TextKit.base64Encode(imgData),
      mimeMap[ext] ?? "image/png",
      w,
      h,
      undefined,
      layout,
    );
  } catch {
    return null;
  }
}

function extractHwpxLayout(posAttr: any, pic: any): ImgLayout {
  const textWrap: string =
    pic?._attr?.textWrap ?? pic?.pic?.[0]?._attr?.textWrap ?? "TOP_AND_BOTTOM";
  const layout = extractHwpxObjectLayout(posAttr, textWrap);
  applyHwpxOutMargin(layout, pic);
  return layout;
}

function extractHwpxTableLayout(tbl: any): ImgLayout | undefined {
  const posAttr = tbl?.["hp:pos"]?.[0]?._attr ?? tbl?.pos?.[0]?._attr ?? {};
  const textWrap: string = tbl?._attr?.textWrap ?? "TOP_AND_BOTTOM";
  const layout = extractHwpxObjectLayout(posAttr, textWrap);
  if (layout.wrap === "inline") return undefined;
  applyHwpxOutMargin(layout, tbl);
  return layout;
}

function extractHwpxObjectLayout(posAttr: any, textWrap: string): ImgLayout {
  const treatAsChar =
    posAttr.treatAsChar === "1" || posAttr.treatAsChar === "true";
  if (treatAsChar) return { wrap: "inline" };

  // OWPML §7.5.8.1 textWrap → ImgWrap 매핑
  // TOP_AND_BOTTOM: 텍스트가 이미지 위아래로만 흐름 → DOCX wrapTopAndBottom (float anchor)
  const wrapMap: Record<string, ImgWrap> = {
    TOP_AND_BOTTOM: "topAndBottom", // float, 위아래 텍스트 흐름
    SQUARE: "square",
    BOTH_SIDES: "tight",
    LEFT: "tight",
    RIGHT: "tight",
    LARGER_ONLY: "tight",
    SMALLER_ONLY: "tight",
    LARGEST_ONLY: "tight",
    BEHIND_TEXT: "behind",
    FRONT_TEXT: "front",
    IN_FRONT_OF_TEXT: "front",
  };
  const wrap: ImgWrap = wrapMap[textWrap] ?? "square";

  // 기준점
  const horzRelToMap: Record<string, ImgHorzRelTo> = {
    PARA: "para",
    MARGIN: "margin",
    PAGE: "page",
    COLUMN: "column",
  };
  const vertRelToMap: Record<string, ImgVertRelTo> = {
    PARA: "para",
    MARGIN: "margin",
    PAGE: "page",
    PAPER: "page",
    LINE: "line",
  };
  const horzRelTo = horzRelToMap[posAttr.horzRelTo ?? ""] ?? "para";
  const vertRelTo = vertRelToMap[posAttr.vertRelTo ?? ""] ?? "para";

  // 정렬
  const horzAlignMap: Record<string, ImgHorzAlign> = {
    LEFT: "left",
    CENTER: "center",
    RIGHT: "right",
  };
  const vertAlignMap: Record<string, ImgVertAlign> = {
    TOP: "top",
    CENTER: "center",
    BOTTOM: "bottom",
  };
  const horzAlign = horzAlignMap[posAttr.horzAlign ?? ""];
  const vertAlign = vertAlignMap[posAttr.vertAlign ?? ""];

  // 오프셋
  const horzOffset = Number(posAttr.horzOffset ?? 0);
  const vertOffset = Number(posAttr.vertOffset ?? 0);
  const xPt = horzOffset !== 0 ? Metric.hwpToPt(horzOffset) : undefined;
  const yPt = vertOffset !== 0 ? Metric.hwpToPt(vertOffset) : undefined;

  return { wrap, horzAlign, vertAlign, horzRelTo, vertRelTo, xPt, yPt };
}

function applyHwpxOutMargin(layout: ImgLayout, obj: any): void {
  const outMargin = obj?.["hp:outMargin"]?.[0]?._attr ?? obj?.outMargin?.[0]?._attr;
  if (!outMargin) return;
  const assign = (
    attr: string,
    key: "distT" | "distB" | "distL" | "distR",
  ) => {
    const raw = outMargin[attr];
    if (raw === undefined) return;
    const value = Number(raw);
    if (Number.isFinite(value) && value >= 0) layout[key] = Metric.hwpToPt(value);
  };
  assign("top", "distT");
  assign("bottom", "distB");
  assign("left", "distL");
  assign("right", "distR");
}

// ─── Table decoding ────────────────────────────────────────

function validHwpxCellPadding(value: unknown): number | undefined {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed >= 0 && parsed < 0xffff
    ? parsed
    : undefined;
}

function decodeGrid(tbl: any, ctx: DecCtx): GridNode {
  const tblAttr = tbl?._attr ?? {};
  const borderFillId = Number(tblAttr.borderFillIDRef ?? 0);
  const borderFill = ctx.borderFills.get(borderFillId);
  const headerRow = tblAttr.repeatHeader === "1";
  const inMarginAttr = tbl?.["hp:inMargin"]?.[0]?._attr ?? {};
  const tablePadding = {
    left: validHwpxCellPadding(inMarginAttr.left) ?? 510,
    right: validHwpxCellPadding(inMarginAttr.right) ?? 510,
    top: validHwpxCellPadding(inMarginAttr.top) ?? 141,
    bottom: validHwpxCellPadding(inMarginAttr.bottom) ?? 141,
  };

  const gridProps: GridProps = {
    headerRow: headerRow || undefined,
    cellPadL: Metric.hwpToPt(tablePadding.left),
    cellPadR: Metric.hwpToPt(tablePadding.right),
    cellPadT: Metric.hwpToPt(tablePadding.top),
    cellPadB: Metric.hwpToPt(tablePadding.bottom),
  };
  if (borderFill?.stroke) gridProps.defaultStroke = borderFill.stroke;
  const layout = extractHwpxTableLayout(tbl);
  if (layout) gridProps.layout = layout;

  // 표 정렬 — <hp:pos horzAlign="..."> 에서 읽음 (OWPML CAbstractShapeObjectType)
  const posAttr = tbl?.["hp:pos"]?.[0]?._attr ?? {};
  if (posAttr.horzAlign) {
    const alignMap: Record<string, "left" | "right" | "center" | "justify"> = {
      LEFT: "left", RIGHT: "right", CENTER: "center", JUSTIFY: "justify",
    };
    const a = alignMap[posAttr.horzAlign];
    if (a) gridProps.align = a;
  }

  const rowArr = getTag(tbl, "hp:tr", "hp:ROW");
  // Every merged cell contributes a linear width constraint.  Solving all of
  // them together recovers the underlying columns; equal splitting loses the
  // asymmetric widths used by Hancom forms.
  let detectedCols = Math.max(0, Number(tblAttr.colCnt ?? tblAttr.ColCnt ?? 0));
  const widthConstraints: { start: number; span: number; width: number }[] = [];
  for (const row of rowArr) {
    let sequentialCol = 0;
    for (const cell of getTag(row, "hp:tc", "hp:CELL")) {
      const spanAttr = cell?.["hp:cellSpan"]?.[0]?._attr ?? {};
      const addrAttr = cell?.["hp:cellAddr"]?.[0]?._attr ?? {};
      const span = Math.max(1, Number(spanAttr.colSpan ?? cell?._attr?.ColSpan ?? 1));
      const address = Number(addrAttr.colAddr ?? sequentialCol);
      const start = Number.isFinite(address) && address >= 0 ? address : sequentialCol;
      const width = Number(cell?.["hp:cellSz"]?.[0]?._attr?.width ?? 0);
      if (width > 0) widthConstraints.push({ start, span, width });
      detectedCols = Math.max(detectedCols, start + span);
      sequentialCol = start + span;
    }
  }
  if (detectedCols > 0) {
    const inferred = inferColumnWidths(detectedCols, widthConstraints);
    if (inferred.some((width) => width > 0)) {
      gridProps.colWidths = inferred.map(Metric.hwpToPt);
    }
  }
  const rowNodes = rowArr.map((row: any) => {
    const cellArr = [...getTag(row, "hp:tc", "hp:CELL")].sort((a, b) => {
      const aa = Number(a?.["hp:cellAddr"]?.[0]?._attr?.colAddr ?? 0);
      const ba = Number(b?.["hp:cellAddr"]?.[0]?._attr?.colAddr ?? 0);
      return aa - ba;
    });
    const cellNodes = cellArr.map((cell: any) => {
      const ca = cell?._attr ?? {};

      // Cell borderFill
      const cellBfId = Number(ca.borderFillIDRef ?? 0);
      const cellBf = ctx.borderFills.get(cellBfId);

      const cellProps: CellProps = {
        bg: cellBf?.bgColor ?? safeHex(ca.BgColor),
      };

      if (cellBf) {
        // Preserve explicit NONE so it overrides table-level defaultStroke in DOCX tcBorders.
        // Only skip when the side is truly undefined (not specified in borderFill).
        cellProps.top = cellBf.top ?? cellBf.stroke;
        cellProps.bot = cellBf.bottom ?? cellBf.stroke;
        cellProps.left = cellBf.left ?? cellBf.stroke;
        cellProps.right = cellBf.right ?? cellBf.stroke;
      }

      // 수직 정렬 — <hp:subList vertAlign="..."> (OWPML CParaListType::GetVertAlign)
      const subList = cell?.["hp:subList"]?.[0] ?? cell?.subList?.[0];
      const subAttr = subList?._attr ?? {};
      if (subAttr.vertAlign) {
        const vaMap: Record<string, "top" | "mid" | "bot"> = {
          TOP: "top",
          CENTER: "mid",
          BOTTOM: "bot",
        };
        cellProps.va = vaMap[subAttr.vertAlign];
      }
      // 셀 여백 — <hp:cellMargin left/right/top/bottom> (OWPML CTc::SetcellMargin)
      // subList 속성이 아닌 <hp:cellMargin> 자식 요소에서 읽어야 함
      const cellMarginAttr =
        cell?.["hp:cellMargin"]?.[0]?._attr ?? {};
      const mL = validHwpxCellPadding(cellMarginAttr.left);
      const mR = validHwpxCellPadding(cellMarginAttr.right);
      const mT = validHwpxCellPadding(cellMarginAttr.top);
      const mB = validHwpxCellPadding(cellMarginAttr.bottom);
      if (mL !== undefined && mL !== tablePadding.left) cellProps.padL = Metric.hwpToPt(mL);
      if (mR !== undefined && mR !== tablePadding.right) cellProps.padR = Metric.hwpToPt(mR);
      if (mT !== undefined && mT !== tablePadding.top) cellProps.padT = Metric.hwpToPt(mT);
      if (mB !== undefined && mB !== tablePadding.bottom) cellProps.padB = Metric.hwpToPt(mB);

      // Colspan/rowspan from cellSpan element or attributes
      const cellSpan = cell?.["hp:cellSpan"]?.[0]?._attr ?? {};
      const cs = Number(cellSpan.colSpan ?? ca.ColSpan ?? 1);
      const rs = Number(cellSpan.rowSpan ?? ca.RowSpan ?? 1);

      // Parse cell content — paragraphs and nested tables (중첩 표)
      const cellKids: (ParaNode | GridNode)[] = [];
      const source = subList ?? cell;
      const sourcePSource = getTag(source, "hp:p", "hp:P");
      for (const sp of sourcePSource) {
        try {
          // Check if this paragraph contains a nested table in its runs
          const runs = getTag(sp, "hp:run", "hp:RUN");
          for (const run of runs) {
            const nestedTbls = getTag(run, "hp:tbl", "hp:TABLE");
            for (const nestedTbl of nestedTbls) {
              try {
                cellKids.push(decodeGrid(nestedTbl, ctx));
              } catch {
                /* skip malformed nested table */
              }
            }
          }
          // Preserve the source anchor paragraph after nested tables as well;
          // the DOCX cell requires it and its spacing/style affects cell size.
          cellKids.push(decodePara(sp, ctx));
        } catch {
          /* skip corrupted para in cell */
        }
      }

      return buildCell(
        cellKids.length > 0 ? cellKids : [buildPara([buildSpan("")])],
        { cs, rs, props: cellProps },
      );
    });
    // Row height: prefer a non-merged cell (rs=1) for accuracy.
    // For merged cells, divide total height by rowSpan to get per-row height.
    let rowHeightPt: number | undefined;
    for (const cell of cellArr) {
      const ca = cell?._attr ?? {};
      const cellSpan = cell?.["hp:cellSpan"]?.[0]?._attr ?? {};
      const cellRs = Math.max(1, Number(cellSpan.rowSpan ?? ca.RowSpan ?? 1));
      const hSz = cell?.["hp:cellSz"]?.[0]?._attr ?? {};
      const hVal = Number(hSz.height ?? 0);
      if (hVal > 0) {
        rowHeightPt = Metric.hwpToPt(hVal) / cellRs;
        if (cellRs === 1) break; // exact match — stop searching
      }
    }
    return buildRow(cellNodes, rowHeightPt);
  });
  return buildGrid(rowNodes, gridProps);
}

function decodeGridSimple(tbl: any, ctx: DecCtx): GridNode {
  const rowArr = getTag(tbl, "hp:tr", "hp:ROW");
  const rowNodes = rowArr.map((row: any) => {
    const cellArr = getTag(row, "hp:tc", "hp:CELL");
    return buildRow(
      cellArr.map((cell: any) =>
        buildCell([buildPara([buildSpan(cellText(cell))])]),
      ),
    );
  });
  return buildGrid(rowNodes);
}

function decodeGridFlat(tbl: any): GridNode {
  return buildGrid([
    buildRow([buildCell([buildPara([buildSpan(tableText(tbl))])])]),
  ]);
}

function decodeGridText(tbl: any): ParaNode {
  return buildPara([buildSpan(tableText(tbl))]);
}

function cellText(cell: any): string {
  const subList = cell?.["hp:subList"]?.[0] ?? cell?.subList?.[0];
  const source = subList ?? cell;
  return getTag(source, "hp:p", "hp:P")
    .map((p: any) =>
      getTag(p, "hp:run", "hp:RUN")
        .map((r: any) =>
          getTag(r, "hp:t", "hp:T")
            .map((t: any) => {
              const val =
                typeof t === "string"
                  ? t
                  : (t?._text ?? t?._ ?? t?.["#text"] ?? "");
              return val.replace(/__EXT_\d+(?:_W\d+_H\d+)?__/g, "");
            })
            .join(""),
        )
        .join(""),
    )
    .join(" ");
}

function tableText(tbl: any): string {
  return getTag(tbl, "hp:tr", "hp:ROW")
    .map((row: any) =>
      getTag(row, "hp:tc", "hp:CELL")
        .map((c: any) => cellText(c))
        .join("\t"),
    )
    .join("\n");
}

function toArr(v: any): any[] {
  return v == null ? [] : Array.isArray(v) ? v : [v];
}

// Auto-register
registry.registerDecoder(new HwpxDecoder());
