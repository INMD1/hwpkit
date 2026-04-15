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

// ─── All HWPX namespaces ────────────────────────────────────
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
  'xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"',
  'xmlns:epub="http://www.idpf.org/2007/ops"',
  'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"',
].join(" ");

// ─── Registries for IDRef system ────────────────────────────

interface CharPrDef {
  id: number;
  height: number;
  bold: boolean;
  italic: boolean;
  underline: string;
  strikeout: string;
  textColor: string;
  fontName: string;
  fontId: number;
  bg?: string;
}

interface ParaPrDef {
  id: number;
  align: string;
  listType?: string;
  listLevel?: number;
  leftHwp: number;    // 왼쪽 여백 (body left indent) in HWPUNIT
  intentHwp: number; // 첫 줄 들여쓰기 (first-line indent) in HWPUNIT
  prevHwp: number;   // space before paragraph in HWPUNIT
  nextHwp: number;   // space after paragraph in HWPUNIT
  lineSpacing: number; // line spacing percentage (e.g., 160 = 160%)
}

interface BinEntry {
  id: string;
  name: string;
  data: Uint8Array;
  originalName?: string;
}

interface HwpxCtx {
  charPrs: CharPrDef[];
  charPrMap: Map<string, number>;
  paraPrs: ParaPrDef[];
  paraPrMap: Map<string, number>;
  borderFills: { id: number; xml: string }[];
  bins: BinEntry[];
  nextBinNum: number;
  nextElementId: number;
  availableWidth: number; // HWPUNIT — page width minus margins
  fonts: string[];
  fontMap: Map<string, number>;
  imgMap: WeakMap<ImgNode, string>; // ImgNode → binId (no mutation)
  nextZOrder: number; // monotonically increasing z-order for images/objects
}

function charPrKey(p: TextProps): string {
  return `${p.b ? 1 : 0}|${p.i ? 1 : 0}|${p.u ? 1 : 0}|${p.s ? 1 : 0}|${p.pt ?? 10}|${p.color ?? "000000"}|${p.font ?? ""}|${p.bg ?? ""}`;
}

function paraPrKey(p: ParaProps): string {
  return `${p.align ?? "left"}|${p.listOrd ?? ""}|${p.listLv ?? 0}|${p.indentPt ?? 0}|${p.firstLineIndentPt ?? 0}|${p.spaceBefore ?? 0}|${p.spaceAfter ?? 0}|${p.lineHeight ?? 0}`;
}

function registerFont(name: string, ctx: HwpxCtx): number {
  const n = name || "굴림체";
  const existing = ctx.fontMap.get(n);
  if (existing !== undefined) return existing;
  const id = ctx.fonts.length;
  ctx.fonts.push(n);
  ctx.fontMap.set(n, id);
  return id;
}

function registerCharPr(props: TextProps, ctx: HwpxCtx): number {
  const key = charPrKey(props);
  const existing = ctx.charPrMap.get(key);
  if (existing !== undefined) return existing;

  const fontName = safeFontToKr(props.font) || "굴림체";
  const fontId = registerFont(fontName, ctx);

  const id = ctx.charPrs.length;
  const def: CharPrDef = {
    id,
    height: Metric.ptToHHeight(props.pt ?? 10),
    bold: !!props.b,
    italic: !!props.i,
    underline: props.u ? "BOTTOM" : "NONE",
    strikeout: props.s ? "SOLID" : "NONE",
    textColor: props.color ? `#${props.color}` : "#000000",
    fontName,
    fontId,
    bg: props.bg,
  };
  ctx.charPrs.push(def);
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
    intentHwp: props.firstLineIndentPt ? Metric.ptToHwp(props.firstLineIndentPt) : 0,
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

// ─── Pre-scan: collect all charPr/paraPr used ───────────────

function scanContent(kids: ContentNode[], ctx: HwpxCtx): void {
  for (const kid of kids) {
    if (kid.tag === "para") scanPara(kid, ctx);
    else if (kid.tag === "grid") scanGrid(kid, ctx);
  }
}

function scanPara(para: ParaNode, ctx: HwpxCtx): void {
  registerParaPr(para.props, ctx);
  for (const kid of para.kids) {
    if (kid.tag === "span") registerCharPr(kid.props, ctx);
    else if (kid.tag === "img") registerImage(kid, ctx);
  }
}

function scanGrid(grid: GridNode, ctx: HwpxCtx): void {
  for (const row of grid.kids)
    for (const cell of row.kids) for (const p of cell.kids) scanPara(p, ctx);
}

function scanParas(paras: ParaNode[], ctx: HwpxCtx): void {
  for (const p of paras) scanPara(p, ctx);
}

// ─── Image handling ─────────────────────────────────────────

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

// ─── BorderFill ─────────────────────────────────────────────

function addBorderFill(
  ctx: HwpxCtx,
  stroke?: Stroke,
  bgColor?: string,
): number {
  const id = ctx.borderFills.length + 1;
  const s = stroke ?? DEFAULT_STROKE;
  const kindMap: Record<string, string> = {
    solid: "SOLID",
    dash: "DASH",
    dot: "DOT",
    double: "DOUBLE",
    none: "NONE",
  };
  const type = kindMap[s.kind] ?? "SOLID";
  const w = `${(s.pt * 0.3528).toFixed(2)} mm`;
  const c = s.color.startsWith("#") ? s.color : `#${s.color}`;

  let fill = "";
  if (bgColor) {
    const bc = bgColor.startsWith("#") ? bgColor : `#${bgColor}`;
    fill = `<hc:fillBrush><hc:winBrush faceColor="${bc}" hatchColor="none" alpha="0"/></hc:fillBrush>`;
  }

  const xml = `<hh:borderFill id="${id}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0"><hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/><hh:leftBorder type="${type}" width="${w}" color="${c}"/><hh:rightBorder type="${type}" width="${w}" color="${c}"/><hh:topBorder type="${type}" width="${w}" color="${c}"/><hh:bottomBorder type="${type}" width="${w}" color="${c}"/><hh:diagonal type="NONE" width="0.12 mm" color="#000000"/>${fill}</hh:borderFill>`;
  ctx.borderFills.push({ id, xml });
  return id;
}

function addBorderFillPerSide(
  ctx: HwpxCtx,
  top?: Stroke,
  right?: Stroke,
  bottom?: Stroke,
  left?: Stroke,
  bgColor?: string,
): number {
  const id = ctx.borderFills.length + 1;
  const kindMap: Record<string, string> = {
    solid: "SOLID", dash: "DASH", dot: "DOT", double: "DOUBLE", none: "NONE",
  };
  function sideXml(tag: string, s?: Stroke): string {
    const type = s ? (kindMap[s.kind] ?? "SOLID") : "NONE";
    const w = s ? `${(s.pt * 0.3528).toFixed(2)} mm` : "0.12 mm";
    const c = s ? (s.color.startsWith("#") ? s.color : `#${s.color}`) : "#000000";
    return `<hh:${tag} type="${type}" width="${w}" color="${c}"/>`;
  }

  let fill = "";
  if (bgColor) {
    const bc = bgColor.startsWith("#") ? bgColor : `#${bgColor}`;
    fill = `<hc:fillBrush><hc:winBrush faceColor="${bc}" hatchColor="none" alpha="0"/></hc:fillBrush>`;
  }

  const xml = `<hh:borderFill id="${id}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0"><hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/>${sideXml("leftBorder", left)}${sideXml("rightBorder", right)}${sideXml("topBorder", top)}${sideXml("bottomBorder", bottom)}<hh:diagonal type="NONE" width="0.12 mm" color="#000000"/>${fill}</hh:borderFill>`;
  ctx.borderFills.push({ id, xml });
  return id;
}

// ─── Encoder class ──────────────────────────────────────────

export class HwpxEncoder implements Encoder {
  readonly format = "hwpx";

  async encode(doc: DocRoot): Promise<Outcome<Uint8Array>> {
    try {
      const sheet = doc.kids[0];
      const dims = normalizeDims(sheet?.dims ?? A4);

      // Available width = page width - left margin - right margin (in HWPUNIT)
      const availableWidth =
        Metric.ptToHwp(dims.wPt) -
        Metric.ptToHwp(dims.ml) -
        Metric.ptToHwp(dims.mr);

      const ctx: HwpxCtx = {
        charPrs: [],
        charPrMap: new Map(),
        paraPrs: [],
        paraPrMap: new Map(),
        borderFills: [],
        bins: [],
        nextBinNum: 1,
        nextElementId: 10000,
        availableWidth,
        fonts: [],
        fontMap: new Map(),
        imgMap: new WeakMap(),
        nextZOrder: 0,
      };

      // Default borderFill (id=1, no border)
      addBorderFill(ctx, { kind: "none", pt: 0.1, color: "000000" });
      // Table border borderFill (id=2)
      addBorderFill(ctx, DEFAULT_STROKE);
      // Default no-border for text areas (id=3)
      addBorderFill(ctx, { kind: "none", pt: 0.1, color: "000000" });

      // Register default charPr (id=0) and paraPr (id=0)
      registerCharPr({}, ctx);
      registerParaPr({}, ctx);

      // Pre-scan all content to collect charPr/paraPr/images
      scanContent(sheet?.kids ?? [], ctx);
      if (sheet?.header) scanParas(sheet.header, ctx);
      if (sheet?.footer) scanParas(sheet.footer, ctx);

      // Extract plain text preview from document
      const previewText = extractPreviewText(sheet);

      // IMPORTANT: Generate section XML FIRST so that borderFills created
      // during table encoding are registered in ctx before headerXml runs.
      const sectionData = TextKit.encode(sectionXml(sheet, dims, ctx));
      const headerData = TextKit.encode(headerXml(dims, doc.meta, ctx));

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
          data: TextKit.encode(settingsXml(doc.meta)),
          mime: "application/xml",
        },
        {
          name: "META-INF/container.rdf",
          data: TextKit.encode(CONTAINER_RDF),
          mime: "application/rdf+xml",
        },
        {
          name: "Contents/content.hpf",
          data: TextKit.encode(contentHpf(ctx, doc.meta)),
          mime: "application/hwpml-package+xml",
        },
        {
          name: "META-INF/container.xml",
          data: TextKit.encode(CONTAINER_XML),
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
                : "image/bmp";
        entries.push({ name: `BinData/${bin.name}`, data: bin.data, mime: ct });
      }

      // Generate manifest.xml listing all files (Required for v1.4)
      entries.push({
        name: "META-INF/manifest.xml",
        data: TextKit.encode(manifestXml(ctx)),
        mime: "text/xml",
      });

      return succeed(await ArchiveKit.zip(entries));
    } catch (e: any) {
      return fail(`HWPX encode error: ${e?.message ?? String(e)}`);
    }
  }
}

// ─── Constants ──────────────────────────────────────────────

const VERSION_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" targetApplication="WORDPROCESSOR" major="5" minor="1" micro="0" buildNumber="1" os="1" xmlVersion="1.4" application="Hancom Office Hangul" appVersion="11, 0, 0, 0"/>`;

const CONTAINER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"><ocf:rootfiles><ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/><ocf:rootfile full-path="Preview/PrvText.txt" media-type="text/plain"/><ocf:rootfile full-path="META-INF/container.rdf" media-type="application/rdf+xml"/></ocf:rootfiles></ocf:container>`;

const CONTAINER_RDF = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"><rdf:Description rdf:about=""><ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" rdf:resource="Contents/header.xml"/></rdf:Description><rdf:Description rdf:about="Contents/header.xml"><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#HeaderFile"/></rdf:Description><rdf:Description rdf:about=""><ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" rdf:resource="Contents/section0.xml"/></rdf:Description><rdf:Description rdf:about="Contents/section0.xml"><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#SectionFile"/></rdf:Description><rdf:Description rdf:about=""><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#Document"/></rdf:Description></rdf:RDF>`;

function manifestXml(ctx: HwpxCtx): string {
  let items = `<odf:file-entry odf:media-type="application/hwp+zip" odf:full-path="/"/><odf:file-entry odf:media-type="application/hwpml-package+xml" odf:full-path="Contents/content.hpf"/><odf:file-entry odf:media-type="application/xml" odf:full-path="Contents/header.xml"/><odf:file-entry odf:media-type="application/xml" odf:full-path="Contents/section0.xml"/><odf:file-entry odf:media-type="application/xml" odf:full-path="settings.xml"/><odf:file-entry odf:media-type="application/xml" odf:full-path="version.xml"/><odf:file-entry odf:media-type="text/plain" odf:full-path="Preview/PrvText.txt"/><odf:file-entry odf:media-type="application/rdf+xml" odf:full-path="META-INF/container.rdf"/><odf:file-entry odf:media-type="application/xml" odf:full-path="META-INF/container.xml"/>`;
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
    items += `<odf:file-entry odf:media-type="${ct}" odf:full-path="BinData/${bin.name}"/>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">${items}</odf:manifest>`;
}

function settingsXml(meta?: DocMeta): string {
  const zoom = meta?.zoom ?? 100;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ha:HWPApplicationSetting xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"><ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/><config:config-item-set name="PrintInfo"><config:config-item name="PrintMethod" type="short">0</config:config-item><config:config-item name="ZoomX" type="short">${zoom}</config:config-item><config:config-item name="ZoomY" type="short">${zoom}</config:config-item></config:config-item-set></ha:HWPApplicationSetting>`;
}

function contentHpf(ctx: HwpxCtx, meta?: DocMeta): string {
  const title = esc(meta?.title ?? "");
  const creator = esc(meta?.author ?? "text");
  const subject = esc(meta?.subject ?? "text");
  const description = esc(meta?.desc ?? "text");
  const keyword = esc(meta?.keywords ?? "text");
  const created =
    meta?.created ?? new Date().toISOString().replace(/\.\d{3}Z$/, "Z");
  const modified = meta?.modified ?? created;

  let items =
    `<opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>` +
    `<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>` +
    `<opf:item id="settings" href="settings.xml" media-type="application/xml"/>`;
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
    const comment = bin.originalName
      ? `<!-- original: ${esc(bin.originalName)} -->`
      : "";
    items += `${comment}<opf:item id="${bin.id}" href="BinData/${bin.name}" media-type="${ct}" isEmbeded="1"/>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><opf:package xmlns:opf="http://www.idpf.org/2007/opf/" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" version="2.0" unique-identifier="hwpx" id="hwpx"><opf:metadata><opf:title>${title}</opf:title><opf:language>ko</opf:language><opf:meta name="creator" content="text">${creator}</opf:meta><opf:meta name="subject" content="text">${subject}</opf:meta><opf:meta name="description" content="text">${description}</opf:meta><opf:meta name="CreatedDate" content="text">${created}</opf:meta><opf:meta name="ModifiedDate" content="text">${modified}</opf:meta><opf:meta name="keyword" content="text">${keyword}</opf:meta></opf:metadata><opf:manifest>${items}</opf:manifest><opf:spine><opf:itemref idref="section0" linear="yes"/></opf:spine></opf:package>`;
}

// ─── header.xml ─────────────────────────────────────────────

function headerXml(dims: PageDims, meta: DocMeta, ctx: HwpxCtx): string {
  // Font face definitions — register all unique fonts per language group
  const fontCount = ctx.fonts.length || 1;
  const langs = [
    "HANGUL",
    "LATIN",
    "HANJA",
    "JAPANESE",
    "OTHER",
    "SYMBOL",
    "USER",
  ];
  let fontFaces = `<hh:fontfaces itemCnt="${langs.length}">`;
  for (const lang of langs) {
    fontFaces += `<hh:fontface lang="${lang}" fontCnt="${fontCount}">`;
    for (let fi = 0; fi < ctx.fonts.length; fi++) {
      fontFaces += `<hh:font id="${fi}" face="${esc(ctx.fonts[fi])}" type="TTF" isEmbedded="0"><hh:typeInfo familyType="FCAT_GOTHIC" weight="6" proportion="0" contrast="0" strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/></hh:font>`;
    }
    if (ctx.fonts.length === 0) {
      fontFaces += `<hh:font id="0" face="맑은 고딕" type="TTF" isEmbedded="0"><hh:typeInfo familyType="FCAT_GOTHIC" weight="6" proportion="0" contrast="0" strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/></hh:font>`;
    }
    fontFaces += `</hh:fontface>`;
  }
  fontFaces += `</hh:fontfaces>`;

  // CharPr definitions
  let charPrXml = "";
  for (const cp of ctx.charPrs) {
    const bold = cp.bold ? "<hh:bold/>" : "";
    const italic = cp.italic ? "<hh:italic/>" : "";
    const fid = cp.fontId ?? 0;
    charPrXml += `<hh:charPr id="${cp.id}" height="${cp.height}" textColor="${cp.textColor}" shadeColor="none" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="3"><hh:fontRef hangul="${fid}" latin="${fid}" hanja="${fid}" japanese="${fid}" other="${fid}" symbol="${fid}" user="${fid}"/><hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/><hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/><hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/><hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>${bold}${italic}<hh:underline type="${cp.underline}" shape="SOLID" color="#000000"/><hh:strikeout shape="${cp.strikeout}" color="#000000"/><hh:outline type="NONE"/><hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/></hh:charPr>`;
  }

  // ParaPr definitions
  let paraPrXml = "";
  for (const pp of ctx.paraPrs) {
    const marginSwitch = `<hp:switch><hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"><hh:margin><hc:intent value="${pp.intentHwp}" unit="HWPUNIT"/><hc:left value="${pp.leftHwp}" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/><hc:prev value="${pp.prevHwp}" unit="HWPUNIT"/><hc:next value="${pp.nextHwp}" unit="HWPUNIT"/></hh:margin><hh:lineSpacing type="PERCENT" value="${pp.lineSpacing}" unit="HWPUNIT"/></hp:case><hp:default><hh:margin><hc:intent value="${pp.intentHwp}" unit="HWPUNIT"/><hc:left value="${pp.leftHwp}" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/><hc:prev value="${pp.prevHwp}" unit="HWPUNIT"/><hc:next value="${pp.nextHwp}" unit="HWPUNIT"/></hh:margin><hh:lineSpacing type="PERCENT" value="${pp.lineSpacing}" unit="HWPUNIT"/></hp:default></hp:switch>`;
    paraPrXml += `<hh:paraPr id="${pp.id}" tabPrIDRef="0" condense="0" fontLineHeight="0" snapToGrid="0" suppressLineNumbers="0" checked="0"><hh:align horizontal="${pp.align}" vertical="BASELINE"/><hh:heading type="NONE" idRef="0" level="0"/><hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="BREAK_WORD" widowOrphan="0" keepWithNext="0" keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/><hh:autoSpacing eAsianEng="0" eAsianNum="0"/>${marginSwitch}<hh:border borderFillIDRef="1" offsetLeft="0" offsetRight="0" offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/></hh:paraPr>`;
  }

  // BorderFill definitions
  const borderFillXml = ctx.borderFills.map((bf) => bf.xml).join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hh:head ${NS} version="1.4" secCnt="1"><hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/><hh:refList>${fontFaces}<hh:borderFills itemCnt="${ctx.borderFills.length}">${borderFillXml}</hh:borderFills><hh:charProperties itemCnt="${ctx.charPrs.length}">${charPrXml}</hh:charProperties><hh:paraProperties itemCnt="${ctx.paraPrs.length}">${paraPrXml}</hh:paraProperties><hh:tabProperties itemCnt="1"><hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/></hh:tabProperties><hh:styles itemCnt="1"><hh:style id="0" type="PARA" name="바탕글" engName="Normal" paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0" langID="1042" lockForm="0"/></hh:styles></hh:refList><hh:compatibleDocument targetProgram="HWP201X"><hh:layoutCompatibility/></hh:compatibleDocument><hh:docOption><hh:linkinfo path="" pageInherit="1" footnoteInherit="0"/></hh:docOption><hh:trackchageConfig flags="56"/></hh:head>`;
}

// ─── section0.xml ───────────────────────────────────────────

function sectionXml(
  sheet: SheetNode | undefined,
  dims: PageDims,
  ctx: HwpxCtx,
): string {
  const kids = sheet?.kids ?? [];

  // First paragraph includes secPr
  // WIDELY = portrait (standard), NARROWLY = landscape
  const secPr = `<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="0" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0"><hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/><hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/><hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/><hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/><hp:pagePr landscape="${dims.orient === "landscape" ? "NARROWLY" : "WIDELY"}" width="${Metric.ptToHwp(dims.wPt)}" height="${Metric.ptToHwp(dims.hPt)}" gutterType="LEFT_ONLY"><hp:margin header="2834" footer="2834" gutter="0" left="${Metric.ptToHwp(dims.ml)}" right="${Metric.ptToHwp(dims.mr)}" top="${Metric.ptToHwp(dims.mt)}" bottom="${Metric.ptToHwp(dims.mb)}"/></hp:pagePr><hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/><hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="284" belowLine="568" aboveLine="852"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" beneathText="0"/></hp:footNotePr><hp:endNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/><hp:noteLine length="0" type="NONE" width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="0" belowLine="576" aboveLine="864"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr><hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill><hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill><hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill></hp:secPr><hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/></hp:ctrl>`;

  let contentXml = "";
  let isFirst = true;

  const defaultLineseg = `<hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="600" horzpos="0" horzsize="${ctx.availableWidth}" flags="393216"/></hp:linesegarray>`;

  for (const kid of kids) {
    if (kid.tag === "para") {
      contentXml += encodePara(kid, ctx, isFirst ? secPr : "");
      isFirst = false;
    } else if (kid.tag === "grid") {
      const gridXml = encodeGrid(kid, ctx);
      const runsXml = `<hp:run charPrIDRef="0">${gridXml}<hp:t></hp:t></hp:run>`;
      contentXml += `<hp:p paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">${isFirst ? secPr : ""}${runsXml}${defaultLineseg}</hp:p>`;
      isFirst = false;
    }
  }

  // If empty, add one empty paragraph with secPr
  if (contentXml === "") {
    contentXml = `<hp:p paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">${secPr}<hp:run charPrIDRef="0"><hp:t></hp:t></hp:run>${defaultLineseg}</hp:p>`;
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec ${NS}>${contentXml}</hs:sec>`;
}

function estimateCellHeight(cell: CellNode, ctx: HwpxCtx): number {
  const topPad = 141;
  const botPad = 141;
  let contentHeight = 0;
  for (const para of cell.kids) {
    const fontSize = getFontSizeForPara(para, ctx);
    const paraPrId = ctx.paraPrMap.get(paraPrKey(para.props));
    const paraPr = paraPrId !== undefined ? ctx.paraPrs[paraPrId] : null;
    const lineSpacing = paraPr ? paraPr.lineSpacing : 160;
    const spaceBefore = paraPr ? paraPr.prevHwp : 0;
    const spaceAfter = paraPr ? paraPr.nextHwp : 0;
    const lineHeight = Math.round((fontSize * lineSpacing) / 100);
    contentHeight += lineHeight + spaceBefore + spaceAfter;
  }
  if (contentHeight === 0) contentHeight = Math.round(1000 * 1.6); // 10pt @ 160%
  return contentHeight + topPad + botPad;
}

function encodePara(
  para: ParaNode,
  ctx: HwpxCtx,
  secPr: string = "",
  availWidth?: number,
): string {
  const paraPrId = registerParaPr(para.props, ctx);

  // Detect page break: paragraph starts on new page if any span contains a pb node
  const hasPageBreak = para.kids.some(
    (k) => k.tag === "span" && k.kids.some((c) => c.tag === "pb"),
  );

  let runs = "";
  function encodeContentRecursive(kids: ParaNode["kids"]): string {
    let r = "";
    for (const kid of kids) {
      if (kid.tag === "span") {
        r += encodeRun(kid, ctx);
      } else if (kid.tag === "img") {
        r += encodeImage(kid, ctx);
      } else if (kid.tag === "link") {
        r += encodeContentRecursive((kid as LinkNode).kids);
      }
    }
    return r;
  }

  runs = encodeContentRecursive(para.kids);

  // If no runs, add default empty run
  if (runs === "") {
    runs = `<hp:run charPrIDRef="0"><hp:t></hp:t></hp:run>`;
  }

  // Use font height, but expand to accommodate inline images
  let fontSize = getFontSizeForPara(para, ctx);
  for (const kid of para.kids) {
    if (kid.tag === "img" && (!kid.layout || kid.layout.wrap === "inline")) {
      const imgH = Metric.ptToHHeight(kid.h);
      if (imgH > fontSize) fontSize = imgH;
    }
  }
  const paraPr = ctx.paraPrs[paraPrId];
  const lineSpacing = paraPr?.lineSpacing ?? 160;
  const spacing = Math.max(0, Math.round(fontSize * (lineSpacing / 100 - 1)));
  const horzsize = availWidth ?? ctx.availableWidth;
  const linesegarray = `<hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="${fontSize}" textheight="${fontSize}" baseline="${Math.round(fontSize * 0.85)}" spacing="${spacing}" horzpos="0" horzsize="${horzsize}" flags="393216"/></hp:linesegarray>`;

  return `<hp:p paraPrIDRef="${paraPrId}" styleIDRef="0" pageBreak="${hasPageBreak ? "1" : "0"}" columnBreak="0" merged="0">${secPr}${runs}${linesegarray}</hp:p>`;
}

/** Get the font size (in HWPX height units, 1000=10pt) for lineseg computation */
function getFontSizeForPara(para: ParaNode, ctx: HwpxCtx): number {
  for (const kid of para.kids) {
    if (kid.tag === "span") {
      const charPrId = ctx.charPrMap.get(charPrKey(kid.props));
      if (charPrId !== undefined && ctx.charPrs[charPrId]) {
        return ctx.charPrs[charPrId].height;
      }
    }
  }
  return 1000; // default 10pt
}

function encodeRun(span: SpanNode, ctx: HwpxCtx): string {
  const charPrId = registerCharPr(span.props, ctx);

  const parts: string[] = [];
  for (const kid of span.kids) {
    if (kid.tag === "txt") {
      if (kid.content) {
        parts.push(`<hp:t>${esc(kid.content)}</hp:t>`);
      } else {
        parts.push(`<hp:t></hp:t>`);
      }
    } else if (kid.tag === "pagenum") {
      const fmt =
        kid.format === "roman"
          ? "ROMAN_LOWER"
          : kid.format === "romanCaps"
            ? "ROMAN_UPPER"
            : "DIGIT";
      parts.push(`<hp:pageNum pageStartsOn="BOTH" formatType="${fmt}"/>`);
    } else if (kid.tag === "br") {
      parts.push(`<hp:t>\n</hp:t>`);
    }
    // pb is handled at paragraph level (pageBreak attribute), skip here
  }

  return `<hp:run charPrIDRef="${charPrId}">${parts.join("")}</hp:run>`;
}

// HWPX textWrap 매핑
const WRAP_HWPX: Record<string, string> = {
  inline: 'TOP_AND_BOTTOM',
  square: 'SQUARE',
  tight: 'BOTH_SIDES',
  through: 'BOTH_SIDES',
  none: 'FRONT_TEXT',
  behind: 'BEHIND_TEXT',
  front: 'FRONT_TEXT',
};
// textFlow 매핑 (wrap 타입별)
const TEXT_FLOW_HWPX: Record<string, string> = {
  inline: 'BOTH_SIDES',
  square: 'LARGEST_ONLY',
  tight: 'BOTH_SIDES',
  through: 'BOTH_SIDES',
  none: 'BOTH_SIDES',
  behind: 'BOTH_SIDES',
  front: 'BOTH_SIDES',
};
const HORZ_RELTO_HWPX: Record<string, string> = {
  para: 'PARA', margin: 'MARGIN', page: 'PAPER', column: 'COLUMN',
};
const VERT_RELTO_HWPX: Record<string, string> = {
  para: 'PARA', margin: 'MARGIN', page: 'PAPER', line: 'LINE',
};

function encodeImage(img: ImgNode, ctx: HwpxCtx): string {
  const binId = ctx.imgMap.get(img);
  if (!binId) return "";

  const charPrId = registerCharPr({}, ctx);
  const w = Metric.ptToHwp(img.w);
  const h = Metric.ptToHwp(img.h);
  const cx = Math.round(w / 2);
  const cy = Math.round(h / 2);

  const layout = img.layout;
  const isInline = !layout || layout.wrap === 'inline';

  const textWrap = layout ? (WRAP_HWPX[layout.wrap] ?? 'TOP_AND_BOTTOM') : 'TOP_AND_BOTTOM';
  const textFlow = layout ? (TEXT_FLOW_HWPX[layout.wrap] ?? 'BOTH_SIDES') : 'BOTH_SIDES';
  const treatAsChar = isInline ? '1' : '0';
  const flowWithText = '1';
  // behind/front/inline 이미지는 다른 객체와 겹침 허용 불필요; square/tight는 허용
  const allowOverlap = (!isInline && layout?.wrap !== 'behind' && layout?.wrap !== 'front') ? '1' : '0';

  const horzRelTo = layout?.horzRelTo ? (HORZ_RELTO_HWPX[layout.horzRelTo] ?? 'PARA') : 'PARA';
  const vertRelTo = layout?.vertRelTo ? (VERT_RELTO_HWPX[layout.vertRelTo] ?? 'PARA') : 'PARA';

  const ALIGN_H: Record<string, string> = { left: 'LEFT', center: 'CENTER', right: 'RIGHT' };
  const ALIGN_V: Record<string, string> = { top: 'TOP', center: 'CENTER', bottom: 'BOTTOM' };
  const horzAlign = layout?.horzAlign ? (ALIGN_H[layout.horzAlign] ?? 'LEFT') : 'LEFT';
  const vertAlign = layout?.vertAlign ? (ALIGN_V[layout.vertAlign] ?? 'TOP') : 'TOP';
  const horzOffset = layout?.xPt != null ? Metric.ptToHwp(layout.xPt) : 0;
  const vertOffset = layout?.yPt != null ? Metric.ptToHwp(layout.yPt) : 0;

  // hp:pic children must follow the exact HWPX spec order.
  const zOrder = ctx.nextZOrder++;
  return `<hp:run charPrIDRef="${charPrId}"><hp:pic id="${ctx.nextElementId++}" zOrder="${zOrder}" numberingType="PICTURE" textWrap="${textWrap}" textFlow="${textFlow}" lock="0" dropcapstyle="None" href="" groupLevel="0" instid="0" reverse="0"><hp:offset x="0" y="0"/><hp:orgSz width="${w}" height="${h}"/><hp:curSz width="${w}" height="${h}"/><hp:flip horizontal="0" vertical="0"/><hp:rotationInfo angle="0" centerX="${cx}" centerY="${cy}" rotateimage="1"/><hp:renderingInfo><hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/><hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/></hp:renderingInfo><hp:imgRect><hc:pt0 x="0" y="0"/><hc:pt1 x="${w}" y="0"/><hc:pt2 x="${w}" y="${h}"/><hc:pt3 x="0" y="${h}"/></hp:imgRect><hp:imgClip left="0" right="0" top="0" bottom="0"/><hp:inMargin left="0" right="0" top="0" bottom="0"/><hp:imgDim dimwidth="${w}" dimheight="${h}"/><hc:img binaryItemIDRef="${binId}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/><hp:effects/><hp:sz width="${w}" widthRelTo="ABSOLUTE" height="${h}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="${treatAsChar}" affectLSpacing="0" flowWithText="${flowWithText}" allowOverlap="${allowOverlap}" holdAnchorAndSO="0" vertRelTo="${vertRelTo}" horzRelTo="${horzRelTo}" vertAlign="${vertAlign}" horzAlign="${horzAlign}" vertOffset="${vertOffset}" horzOffset="${horzOffset}"/><hp:outMargin left="0" right="0" top="0" bottom="0"/></hp:pic><hp:t></hp:t></hp:run>`;
}

function encodeGrid(grid: GridNode, ctx: HwpxCtx): string {
  const rowCount = grid.kids.length;

  // 1단계: 가상 2D 맵핑 (Virtual Table Map) 생성
  interface CellMap {
    type: 'real' | 'absorbed';
    cell?: CellNode;
  }
  const tableMap: CellMap[][] = Array.from({ length: rowCount }, () => []);

  for (let ri = 0; ri < rowCount; ri++) {
    let ci = 0;
    for (const cell of grid.kids[ri].kids) {
      while (tableMap[ri][ci]) ci++; // 이미 점유된 자리 건너뜀

      tableMap[ri][ci] = { type: 'real', cell };

      // 병합 영역 예약
      for (let rr = 0; rr < cell.rs; rr++) {
        const targetRi = ri + rr;
        if (targetRi >= rowCount) break;
        if (!tableMap[targetRi]) tableMap[targetRi] = [];
        for (let cc = 0; cc < cell.cs; cc++) {
          if (rr === 0 && cc === 0) continue;
          tableMap[targetRi][ci + cc] = { type: 'absorbed' };
        }
      }
      ci += cell.cs;
    }
  }

  // 정확한 전체 열 개수 계산
  let colCount = 0;
  for (let ri = 0; ri < rowCount; ri++) {
    colCount = Math.max(colCount, tableMap[ri].length);
  }
  if (colCount === 0) colCount = 1;

  // 2단계: 컬럼 너비 계산
  const totalWidth = ctx.availableWidth;
  const defaultColW = Math.round(totalWidth / colCount);
  const colWidths: number[] = [];
  if (grid.props.colWidths && grid.props.colWidths.length === colCount) {
    const srcPt = [...grid.props.colWidths];
    const knownTotal = srcPt.filter((w) => w > 0).reduce((s, w) => s + w, 0);
    const zeroCount = srcPt.filter((w) => w <= 0).length;
    const availPt = Metric.hwpToPt(totalWidth);
    const remaining = Math.max(0, availPt - knownTotal);
    const zeroFill = zeroCount > 0 ? remaining / zeroCount : 0;
    for (let i = 0; i < srcPt.length; i++) {
      if (srcPt[i] <= 0) srcPt[i] = zeroFill > 0 ? zeroFill : Metric.hwpToPt(defaultColW);
    }
    for (const wPt of srcPt) colWidths.push(Metric.ptToHwp(wPt));
  } else {
    for (let c = 0; c < colCount; c++) colWidths.push(defaultColW);
  }

  const rawTotal = colWidths.reduce((s, w) => s + w, 0);
  if (rawTotal > totalWidth * 1.05) {
    const scale = totalWidth / rawTotal;
    for (let i = 0; i < colWidths.length; i++) colWidths[i] = Math.round(colWidths[i] * scale);
  }
  const actualTotal = colWidths.reduce((s, w) => s + w, 0);

  // 3단계: 행 높이 계산
  const rowHeights: number[] = [];
  for (let ri = 0; ri < rowCount; ri++) {
    const row = grid.kids[ri];
    if (row.heightPt != null && row.heightPt > 0) {
      rowHeights.push(Metric.ptToHwp(row.heightPt));
    } else {
      let maxH = 0;
      for (let ci = 0; ci < colCount; ci++) {
        const entry = tableMap[ri][ci];
        if (entry?.type === 'real') {
          const h = estimateCellHeight(entry.cell!, ctx);
          if (h > maxH) maxH = h;
        }
      }
      rowHeights.push(maxH || Math.round(1000 * 1.6));
    }
  }
  const totalTableHeight = rowHeights.reduce((s, h) => s + h, 0);

  // 4단계: XML 조립
  const tblBfId = grid.props.defaultStroke ? addBorderFill(ctx, grid.props.defaultStroke) : 2;
  let rowsXml = "";

  for (let ri = 0; ri < rowCount; ri++) {
    let cellsXml = "";
    for (let ci = 0; ci < colCount; ci++) {
      const entry = tableMap[ri][ci];
      if (!entry || entry.type === 'absorbed') continue;

      const cell = entry.cell!;
      const cp = cell.props;
      let cellBfId = tblBfId;

      const hasPerSideBorder = cp.top || cp.bot || cp.left || cp.right;
      if (hasPerSideBorder || cp.bg) {
        const defStroke = grid.props.defaultStroke ?? DEFAULT_STROKE;
        cellBfId = hasPerSideBorder 
          ? addBorderFillPerSide(ctx, cp.top ?? defStroke, cp.right ?? defStroke, cp.bot ?? defStroke, cp.left ?? defStroke, cp.bg)
          : addBorderFill(ctx, defStroke, cp.bg);
      }

      let cellW = 0;
      for (let sc = ci; sc < ci + cell.cs && sc < colWidths.length; sc++) cellW += colWidths[sc];
      if (cellW === 0) cellW = defaultColW * cell.cs;

      const cellInnerW = Math.max(cellW - 282, 100);
      const parasXml = cell.kids.map((p) => encodePara(p, ctx, "", cellInnerW)).join("");

      cellsXml += `<hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="${cellBfId}">` +
                  `<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="${cp.va === "mid" ? "CENTER" : cp.va === "bot" ? "BOTTOM" : "TOP"}" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">${parasXml}</hp:subList>` +
                  `<hp:cellAddr colAddr="${ci}" rowAddr="${ri}"/><hp:cellSpan colSpan="${cell.cs}" rowSpan="${cell.rs}"/><hp:cellSz width="${cellW}" height="${rowHeights[ri]}"/><hp:cellMargin left="141" right="141" top="141" bottom="141"/></hp:tc>`;
    }
    rowsXml += `<hp:tr>${cellsXml}</hp:tr>`;
  }

  const headerRow = grid.props.headerRow ? ' repeatHeader="1"' : "";
  return `<hp:tbl id="${ctx.nextElementId++}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE"${headerRow} rowCnt="${rowCount}" colCnt="${colCount}" cellSpacing="0" borderFillIDRef="${tblBfId}" noAdjust="0"><hp:sz width="${actualTotal}" widthRelTo="ABSOLUTE" height="${totalTableHeight}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="138" right="138" top="138" bottom="138"/><hp:inMargin left="138" right="138" top="138" bottom="138"/>${rowsXml}</hp:tbl>`;
}

function extractPreviewText(sheet?: SheetNode): string {
  if (!sheet) return "";
  const lines: string[] = [];
  for (const kid of sheet.kids) {
    if (kid.tag === "para") {
      const text = kid.kids
        .map((k) => {
          if (k.tag === "span")
            return k.kids
              .map((c) => (c.tag === "txt" ? c.content : ""))
              .join("");
          return "";
        })
        .join("");
      if (text) lines.push(text);
    } else if (kid.tag === "grid") {
      for (const row of kid.kids) {
        const cells = row.kids.map((cell) =>
          cell.kids
            .map((p) =>
              p.kids
                .map((k) => {
                  if (k.tag === "span")
                    return k.kids
                      .map((c) => (c.tag === "txt" ? c.content : ""))
                      .join("");
                  return "";
                })
                .join(""),
            )
            .join(""),
        );
        lines.push(cells.join("\t"));
      }
    }
  }
  return lines.join("\r\n");
}

function esc(s: string): string {
  if (!s) return "";
  // 1. 내부 처리용 플레이스홀더(__EXT_0__ 등) 제거
  s = s.replace(/__EXT_\d+__/g, "");
  // 2. 글자 깨짐을 유발하는 쓰레기값 및 BOM 기호 명시적 제거
  s = s.replace(/湰灧/g, "");
  s = s.replace(/\uFEFF/g, "");
  // 3. XML 1.0에서 허용하지 않는 보이지 않는 제어문자 모두 제거
  s = s.replace(/[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD]/g, "");

  return TextKit.escapeXml(s);
}

registry.registerEncoder(new HwpxEncoder());
