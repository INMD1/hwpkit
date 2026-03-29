import type { Decoder } from "../../contract/decoder";
import type {
  DocRoot,
  ContentNode,
  ParaNode,
  SpanNode,
  GridNode,
  ImgNode,
  PageNumNode,
  CellNode,
} from "../../model/doc-tree";
import type { Outcome } from "../../contract/result";
import type {
  DocMeta,
  PageDims,
  TextProps,
  ParaProps,
  CellProps,
  GridProps,
  TableLook,
  ImgLayout,
  ImgHorzAlign,
  ImgVertAlign,
  ImgHorzRelTo,
  ImgVertRelTo,
  ImgWrap,
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
  safeStrokeDocx,
} from "../../safety/StyleBridge";
import { ArchiveKit } from "../../toolkit/ArchiveKit";
import { XmlKit } from "../../toolkit/XmlKit";
import { TextKit } from "../../toolkit/TextKit";
import { registry } from "../../pipeline/registry";

export class DocxDecoder implements Decoder {
  readonly format = "docx";

  async decode(data: Uint8Array): Promise<Outcome<DocRoot>> {
    const shield = new ShieldedParser();
    const warns: string[] = [];

    try {
      const files = await ArchiveKit.unzip(data);

      const docXml = files.get("word/document.xml");
      if (!docXml) return fail("DOCX: word/document.xml not found");

      const relsXml = files.get("word/_rels/document.xml.rels");
      const relsMap = relsXml
        ? await parseRels(TextKit.decode(relsXml))
        : new Map<string, string>();

      const coreXml = files.get("docProps/core.xml");
      let meta: DocMeta = {};
      if (coreXml) {
        try {
          meta = await parseCoreProps(TextKit.decode(coreXml));
        } catch {
          // ignore — meta is optional
        }
      }

      // Parse numbering.xml for list support
      const numXml = files.get("word/numbering.xml");
      let numMap: NumMap = new Map();
      if (numXml) {
        try {
          numMap = await parseNumbering(TextKit.decode(numXml));
        } catch {
          /* non-fatal */
        }
      }

      const docStr = TextKit.decode(docXml);
      const docObj: any = await XmlKit.parseStrict(docStr);

      const body = getBody(docObj);
      const dims = extractDims(body) ?? { ...A4 };
      const elements = getBodyElements(body);

      const decCtx: DecCtx = { relsMap, files, shield, numMap, warns };

      const kids: ContentNode[] = [];
      for (const el of elements) {
        const node = shield.guard(
          () => decodeElement(el, decCtx),
          buildPara([buildSpan("[요소 파싱 실패]")]),
          "docx:bodyElement",
        );
        kids.push(node);

        // Inline sectPr in pPr = section break → insert page-break paragraph after
        if (el.type === 'para') {
          const pPr = el.node?.["w:pPr"]?.[0] ?? el.node?.pPr?.[0] ?? {};
          const inlineSectPr = pPr?.["w:sectPr"]?.[0] ?? pPr?.sectPr?.[0];
          if (inlineSectPr) {
            const typeAttr = inlineSectPr?.["w:type"]?.[0]?._attr;
            const sectType = typeAttr?.["w:val"] ?? typeAttr?.val ?? 'nextPage';
            if (sectType !== 'continuous') {
              kids.push(buildPara([{ tag: 'span', props: {}, kids: [buildPb()] }]));
            }
          }
        }
      }

      // Decode header/footer
      const headerParas = await decodeHeaderFooter(
        "header",
        body,
        relsMap,
        files,
        decCtx,
      );
      const footerParas = await decodeHeaderFooter(
        "footer",
        body,
        relsMap,
        files,
        decCtx,
      );

      warns.push(...shield.flush());
      const sheet = buildSheet(kids.filter(Boolean) as ContentNode[], dims, {
        header: headerParas,
        footer: footerParas,
      });
      return succeed(buildRoot(meta, [sheet]), warns);
    } catch (e: any) {
      warns.push(...shield.flush());
      return fail(`DOCX decode error: ${e?.message ?? String(e)}`, warns);
    }
  }
}

// ─── types ─────────────────────────────────────────────────

interface DecCtx {
  relsMap: Map<string, string>;
  files: Map<string, Uint8Array>;
  shield: ShieldedParser;
  numMap: NumMap;
  warns: string[];
}

// numId → { abstractNumId, levels: Map<ilvl, { fmt, isOrdered }> }
type NumMap = Map<
  number,
  { levels: Map<number, { fmt: string; isOrdered: boolean }> }
>;

// ─── helpers ────────────────────────────────────────────────

function toArr(v: any): any[] {
  return v == null ? [] : Array.isArray(v) ? v : [v];
}

/** Resolve DOCX relative paths. e.g. ("word", "../media/image1.png") → "word/media/image1.png" */
function resolveDocxPath(baseDir: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = (baseDir + "/" + target).split("/");
  const stack: string[] = [];
  for (const p of parts) {
    if (p === "..") {
      stack.pop();
    } else if (p !== ".") {
      stack.push(p);
    }
  }
  return stack.join("/");
}

async function parseRels(xml: string): Promise<Map<string, string>> {
  const map = new Map<string, string>();
  try {
    const obj: any = await XmlKit.parseStrict(xml);
    for (const rel of toArr(obj?.Relationships?.[0]?.Relationship)) {
      const a = rel?._attr ?? {};
      if (a.Id && a.Target) map.set(a.Id, a.Target);
    }
  } catch {
    /* ignore */
  }
  return map;
}

async function parseCoreProps(xml: string): Promise<DocMeta> {
  try {
    const obj: any = await XmlKit.parseStrict(xml);
    const c = obj?.["cp:coreProperties"]?.[0] ?? obj?.coreProperties?.[0] ?? {};
    return {
      title: c?.["dc:title"]?.[0]?._text ?? undefined,
      author: c?.["dc:creator"]?.[0]?._text ?? undefined,
      subject: c?.["dc:subject"]?.[0]?._text ?? undefined,
      created: c?.["dcterms:created"]?.[0]?._text ?? undefined,
      modified: c?.["dcterms:modified"]?.[0]?._text ?? undefined,
    };
  } catch {
    return {};
  }
}

async function parseNumbering(xml: string): Promise<NumMap> {
  const map: NumMap = new Map();
  try {
    const obj: any = await XmlKit.parseStrict(xml);
    const root = obj?.["w:numbering"]?.[0] ?? obj?.numbering?.[0] ?? obj;

    // Parse abstractNums
    const absMap = new Map<
      number,
      Map<number, { fmt: string; isOrdered: boolean }>
    >();
    for (const abs of toArr(root?.["w:abstractNum"] ?? root?.abstractNum)) {
      const absId = Number(
        abs?._attr?.["w:abstractNumId"] ?? abs?._attr?.abstractNumId ?? 0,
      );
      const levels = new Map<number, { fmt: string; isOrdered: boolean }>();
      for (const lvl of toArr(abs?.["w:lvl"] ?? abs?.lvl)) {
        const ilvl = Number(lvl?._attr?.["w:ilvl"] ?? lvl?._attr?.ilvl ?? 0);
        const fmtNode =
          lvl?.["w:numFmt"]?.[0]?._attr ?? lvl?.numFmt?.[0]?._attr ?? {};
        const fmt = fmtNode?.["w:val"] ?? fmtNode?.val ?? "decimal";
        levels.set(ilvl, { fmt, isOrdered: fmt !== "bullet" });
      }
      absMap.set(absId, levels);
    }

    // Parse nums
    for (const num of toArr(root?.["w:num"] ?? root?.num)) {
      const numId = Number(num?._attr?.["w:numId"] ?? num?._attr?.numId ?? 0);
      const absRef =
        num?.["w:abstractNumId"]?.[0]?._attr ??
        num?.abstractNumId?.[0]?._attr ??
        {};
      const absId = Number(absRef?.["w:val"] ?? absRef?.val ?? 0);
      const levels = absMap.get(absId) ?? new Map();
      map.set(numId, { levels });
    }
  } catch {
    /* non-fatal */
  }
  return map;
}

function getBody(obj: any): any {
  return (
    obj?.["w:document"]?.[0]?.["w:body"]?.[0] ??
    obj?.document?.[0]?.body?.[0] ??
    obj
  );
}

function extractDims(body: any): PageDims | null {
  try {
    const sp = body?.["w:sectPr"]?.[0] ?? body?.sectPr?.[0];
    if (!sp) return null;
    const sz = sp?.["w:pgSz"]?.[0]?._attr ?? sp?.pgSz?.[0]?._attr;
    const mar = sp?.["w:pgMar"]?.[0]?._attr ?? sp?.pgMar?.[0]?._attr;
    if (!sz) return null;
    return {
      wPt: Metric.dxaToPt(Number(sz["w:w"] ?? sz.w ?? 11906)),
      hPt: Metric.dxaToPt(Number(sz["w:h"] ?? sz.h ?? 16838)),
      mt: Metric.dxaToPt(Number(mar?.["w:top"] ?? mar?.top ?? 1440)),
      mb: Metric.dxaToPt(Number(mar?.["w:bottom"] ?? mar?.bottom ?? 1440)),
      ml: Metric.dxaToPt(Number(mar?.["w:left"] ?? mar?.left ?? 1800)),
      mr: Metric.dxaToPt(Number(mar?.["w:right"] ?? mar?.right ?? 1800)),
      orient:
        (sz["w:orient"] ?? sz.orient) === "landscape"
          ? "landscape"
          : "portrait",
    };
  } catch {
    return null;
  }
}

function getBodyElements(body: any): { type: string; node: any }[] {
  const paras = toArr(body?.["w:p"] ?? body?.p);
  const tables = toArr(body?.["w:tbl"] ?? body?.tbl);

  if (tables.length === 0)
    return paras.map((n: any) => ({ type: "para", node: n }));
  if (paras.length === 0)
    return tables.map((n: any) => ({ type: "table", node: n }));

  // Use _childOrder from XmlKit to preserve document order
  const childOrder = body?.["_childOrder"] as string[] | undefined;
  if (Array.isArray(childOrder)) {
    const items: { type: string; node: any }[] = [];
    let pi = 0,
      ti = 0;
    for (const tag of childOrder) {
      if ((tag === "w:p" || tag === "p") && pi < paras.length) {
        items.push({ type: "para", node: paras[pi++] });
      } else if ((tag === "w:tbl" || tag === "tbl") && ti < tables.length) {
        items.push({ type: "table", node: tables[ti++] });
      }
    }
    while (pi < paras.length) items.push({ type: "para", node: paras[pi++] });
    while (ti < tables.length)
      items.push({ type: "table", node: tables[ti++] });
    return items;
  }

  // Fallback: paragraphs first, then tables
  return [
    ...paras.map((n: any) => ({ type: "para", node: n })),
    ...tables.map((n: any) => ({ type: "table", node: n })),
  ];
}

// ─── Header/Footer decoding ────────────────────────────────

async function decodeHeaderFooter(
  kind: "header" | "footer",
  body: any,
  relsMap: Map<string, string>,
  files: Map<string, Uint8Array>,
  ctx: DecCtx,
): Promise<ParaNode[] | undefined> {
  try {
    const sp = body?.["w:sectPr"]?.[0] ?? body?.sectPr?.[0];
    if (!sp) return undefined;

    const refTag =
      kind === "header" ? "w:headerReference" : "w:footerReference";
    const refs = toArr(sp?.[refTag] ?? sp?.[refTag.replace("w:", "")]);
    if (refs.length === 0) return undefined;

    const rId =
      refs[0]?._attr?.["r:id"] ??
      refs[0]?._attr?.["r:Id"] ??
      refs[0]?._attr?.id;
    if (!rId) return undefined;

    const target = relsMap.get(rId);
    if (!target) return undefined;

    const filePath = resolveDocxPath("word", target);
    const fileData = files.get(filePath);
    if (!fileData) return undefined;

    const xmlStr = TextKit.decode(fileData);
    const obj: any = await XmlKit.parseStrict(xmlStr);

    const rootTag = kind === "header" ? "w:hdr" : "w:ftr";
    const root =
      obj?.[rootTag]?.[0] ?? obj?.[rootTag.replace("w:", "")]?.[0] ?? obj;

    const paras = toArr(root?.["w:p"] ?? root?.p);
    if (paras.length === 0) return undefined;

    return paras.map((p: any) => decodePara(p, ctx));
  } catch {
    return undefined;
  }
}

// ─── Element decoding ──────────────────────────────────────

//만약에 drawing 태그가 안에 있으면 true 반환
function hasDrawingDeep(node: any): boolean {
  if (!node || typeof node !== "object") return false;

  if (node["w:drawing"] || node["w:pict"]) return true;

  return Object.values(node).some((v) => {
    if (Array.isArray(v)) return v.some(hasDrawingDeep);
    return hasDrawingDeep(v);
  });
}

function decodeElement(
  el: { type: string; node: any },
  ctx: DecCtx,
): ContentNode {
  if (el.type === "table") {
    const { value } = ctx.shield.guardGrid(
      el.node,
      (n) => decodeGrid(n as any, ctx),
      (n) => decodeGridSimple(n as any),
      (n) => decodeGridFlat(n as any),
      (n) => decodeGridText(n as any) as unknown as GridNode,
      "docx:table",
    );
    return value;
  }
  return decodePara(el.node, ctx);
}

function decodePara(p: any, ctx: DecCtx): ParaNode {
  const pPr = p?.["w:pPr"]?.[0] ?? {};
  const alignVal =
    pPr?.["w:jc"]?.[0]?._attr?.["w:val"] ?? pPr?.["w:jc"]?.[0]?._attr?.val;
  const headStyle =
    pPr?.["w:pStyle"]?.[0]?._attr?.["w:val"] ??
    pPr?.["w:pStyle"]?.[0]?._attr?.val ??
    "";

  const props: ParaProps = {
    align: safeAlign(alignVal),
    heading: parseHeading(headStyle),
  };

  // Spacing (before/after/line height)
  const spacingAttr =
    pPr?.["w:spacing"]?.[0]?._attr ?? pPr?.spacing?.[0]?._attr ?? {};
  const beforeVal = Number(
    spacingAttr?.["w:before"] ?? spacingAttr?.before ?? 0,
  );
  const afterVal = Number(spacingAttr?.["w:after"] ?? spacingAttr?.after ?? 0);
  const lineVal = Number(spacingAttr?.["w:line"] ?? spacingAttr?.line ?? 0);
  const lineRule =
    spacingAttr?.["w:lineRule"] ?? spacingAttr?.lineRule ?? "auto";
  if (beforeVal > 0) props.spaceBefore = Metric.dxaToPt(beforeVal);
  if (afterVal > 0) props.spaceAfter = Metric.dxaToPt(afterVal);
  if (lineVal > 0 && lineRule === "auto") props.lineHeight = lineVal / 240;

  // Indentation
  const indAttr = pPr?.["w:ind"]?.[0]?._attr ?? pPr?.ind?.[0]?._attr ?? {};
  const leftVal = Number(indAttr?.["w:left"] ?? indAttr?.left ?? 0);
  if (leftVal > 0) props.indentPt = Metric.dxaToPt(leftVal);

  // List/numbering
  const numPr = pPr?.["w:numPr"]?.[0] ?? pPr?.numPr?.[0];
  if (numPr) {
    const ilvlNode =
      numPr?.["w:ilvl"]?.[0]?._attr ?? numPr?.ilvl?.[0]?._attr ?? {};
    const numIdNode =
      numPr?.["w:numId"]?.[0]?._attr ?? numPr?.numId?.[0]?._attr ?? {};
    const ilvl = Number(ilvlNode?.["w:val"] ?? ilvlNode?.val ?? 0);
    const numId = Number(numIdNode?.["w:val"] ?? numIdNode?.val ?? 0);

    props.listLv = ilvl;
    const numEntry = ctx.numMap.get(numId);
    if (numEntry) {
      const lvlInfo = numEntry.levels.get(ilvl) ?? numEntry.levels.get(0);
      props.listOrd = lvlInfo?.isOrdered ?? false;
    } else {
      // Fallback: numId=1 is typically bullet, numId=2 is numbered
      props.listOrd = numId >= 2;
    }
  }

  // pageBreakBefore: paragraph always starts on a new page
  const pbBeforeNode = pPr?.["w:pageBreakBefore"]?.[0] ?? pPr?.pageBreakBefore?.[0];
  const hasPageBreakBefore = pbBeforeNode != null &&
    (pbBeforeNode?._attr?.["w:val"] ?? pbBeforeNode?._attr?.val ?? "1") !== "0";

  const runs = toArr(p?.["w:r"] ?? p?.r);

  // 3/28 이미지 태크를 찾을수 있기 때문에 별도 함수 구현
  const kids: (SpanNode | ImgNode)[] = ctx.shield.guardAll(
    runs,
    (run: any) =>
      hasDrawingDeep(run) ? decodeRunOrImage(run, ctx) : decodeRun(run, ctx),
    () => buildSpan(""),
    "docx:run",
  );

  const filteredKids = kids.filter(Boolean) as ParaNode["kids"];

  // Prepend pb span when pageBreakBefore is set
  if (hasPageBreakBefore) {
    filteredKids.unshift({ tag: 'span', props: {}, kids: [buildPb()] });
  }

  return buildPara(filteredKids, props);
}

// 3/28 코드 수정
function decodeRunOrImage(run: any, ctx: DecCtx): SpanNode | ImgNode {
  function findFirstDrawing(node: any): any | null {
    if (!node || typeof node !== "object") return null;

    if (node["w:drawing"]) return node["w:drawing"][0];
    if (node["w:pict"]) return node["w:pict"][0];

    for (const value of Object.values(node)) {
      if (Array.isArray(value)) {
        for (const v of value) {
          const found = findFirstDrawing(v);
          if (found) return found;
        }
      } else {
        const found = findFirstDrawing(value);
        if (found) return found;
      }
    }

    return null;
  }

  const drawing = findFirstDrawing(run);

  if (drawing) {
    const img = decodeDrawing(drawing, ctx);
    if (img) return img;
  }

  return decodeRun(run, ctx);
}
function decodeDrawing(drawing: any, ctx: DecCtx): ImgNode | null {
  try {
    const inline = drawing?.["wp:inline"]?.[0] ?? drawing?.inline?.[0];
    const anchor = drawing?.["wp:anchor"]?.[0] ?? drawing?.anchor?.[0];
    const container = inline ?? anchor;
    if (!container) return null;

    // Get dimensions
    const extent =
      container?.["wp:extent"]?.[0]?._attr ??
      container?.extent?.[0]?._attr ??
      {};
    const cx = Number(extent?.cx ?? 0);
    const cy = Number(extent?.cy ?? 0);
    const wPt = Metric.emuToPt(cx);
    const hPt = Metric.emuToPt(cy);

    // Get alt text
    const docPr =
      container?.["wp:docPr"]?.[0]?._attr ?? container?.docPr?.[0]?._attr ?? {};
    const alt = docPr?.descr ?? docPr?.name ?? "";

    // Navigate to blip
    const graphic = container?.["a:graphic"]?.[0] ?? container?.graphic?.[0];
    const graphicData =
      graphic?.["a:graphicData"]?.[0] ?? graphic?.graphicData?.[0];
    const pic = graphicData?.["pic:pic"]?.[0] ?? graphicData?.pic?.[0];
    const blipFill = pic?.["pic:blipFill"]?.[0] ?? pic?.blipFill?.[0];
    const blip =
      blipFill?.["a:blip"]?.[0]?._attr ?? blipFill?.blip?.[0]?._attr ?? {};
    const rId = blip?.["r:embed"] ?? blip?.embed;

    if (!rId) return null;

    const target = ctx.relsMap.get(rId);
    if (!target) return null;

    const filePath = resolveDocxPath("word", target);
    const fileData = ctx.files.get(filePath);
    if (!fileData) {
      console.warn(
        `[DocxDecoder] image not found in ZIP: "${filePath}" (rId=${rId}, target=${target})`,
      );
      return null;
    }

    const ext = target.split(".").pop()?.toLowerCase() ?? "png";
    const mimeMap: Record<string, ImgNode["mime"]> = {
      png: "image/png",
      jpg: "image/jpeg",
      jpeg: "image/jpeg",
      gif: "image/gif",
      bmp: "image/bmp",
    };
    const mime = mimeMap[ext] ?? "image/png";
    console.log(
      `[DocxDecoder] image loaded: ${filePath} (${mime}, ${fileData.length} bytes)`,
    );

    // ── layout 추출 ──────────────────────────────────────────
    const layout: ImgLayout = inline
      ? { wrap: 'inline' }
      : extractAnchorLayout(anchor);

    return buildImg(TextKit.base64Encode(fileData), mime, wPt, hPt, alt || undefined, layout);
  } catch {
    return null;
  }
}

function decodeRun(run: any, ctx: DecCtx): SpanNode {
  const rPr = run?.["w:rPr"]?.[0] ?? run?.rPr?.[0] ?? {};

  const szAttr = rPr?.["w:sz"]?.[0]?._attr ?? rPr?.sz?.[0]?._attr ?? {};
  const szVal = szAttr?.["w:val"] ?? szAttr?.val;

  const colorAttr =
    rPr?.["w:color"]?.[0]?._attr ?? rPr?.color?.[0]?._attr ?? {};
  const colorVal = colorAttr?.["w:val"] ?? colorAttr?.val;

  const fontAttr =
    rPr?.["w:rFonts"]?.[0]?._attr ?? rPr?.rFonts?.[0]?._attr ?? {};
  const fontName =
    fontAttr?.["w:ascii"] ??
    fontAttr?.ascii ??
    fontAttr?.["w:hAnsi"] ??
    fontAttr?.hAnsi ??
    fontAttr?.["w:eastAsia"] ??
    fontAttr?.eastAsia;

  const underVal =
    rPr?.["w:u"]?.[0]?._attr?.["w:val"] ?? rPr?.["w:u"]?.[0]?._attr?.val;

  // Background/highlight
  const shdAttr = rPr?.["w:shd"]?.[0]?._attr ?? rPr?.shd?.[0]?._attr ?? {};
  const bgVal = safeHex(shdAttr?.["w:fill"] ?? shdAttr?.fill);

  // Superscript/subscript
  const vertAlignVal =
    rPr?.["w:vertAlign"]?.[0]?._attr?.["w:val"] ??
    rPr?.["w:vertAlign"]?.[0]?._attr?.val;

  // Check bold/italic/strike — val="0" means explicitly OFF
  const bNode = rPr?.["w:b"]?.[0] ?? rPr?.b?.[0];
  const isBold =
    bNode != null &&
    (bNode?._attr?.["w:val"] ?? bNode?._attr?.val ?? "1") !== "0";
  const iNode = rPr?.["w:i"]?.[0] ?? rPr?.i?.[0];
  const isItalic =
    iNode != null &&
    (iNode?._attr?.["w:val"] ?? iNode?._attr?.val ?? "1") !== "0";
  const sNode = rPr?.["w:strike"]?.[0] ?? rPr?.strike?.[0];
  const isStrike =
    sNode != null &&
    (sNode?._attr?.["w:val"] ?? sNode?._attr?.val ?? "1") !== "0";

  const props: TextProps = {
    b: isBold || undefined,
    i: isItalic || undefined,
    u: underVal && underVal !== "none" ? true : undefined,
    s: isStrike || undefined,
    sup: vertAlignVal === "superscript" || undefined,
    sub: vertAlignVal === "subscript" || undefined,
    pt: szVal ? Metric.halfPtToPt(Number(szVal)) : undefined,
    color: safeHex(colorVal),
    font: fontName ? safeFont(fontName) : undefined,
    bg: bgVal,
  };

  // Check for field codes (PAGE number)
  const fldChar = run?.["w:fldChar"]?.[0]?._attr ?? run?.fldChar?.[0]?._attr;
  const instrText = run?.["w:instrText"]?.[0];

  // Page break: <w:br w:type="page"/>
  const brNodes = toArr(run?.["w:br"] ?? run?.br ?? []);
  for (const br of brNodes) {
    const brType = br?._attr?.["w:type"] ?? br?._attr?.type;
    if (brType === "page") {
      return { tag: "span", props, kids: [buildPb()] };
    }
  }

  const textNodes = toArr(run?.["w:t"] ?? run?.t);
  const content = textNodes
    .map((t: any) => (typeof t === "string" ? t : (t?._ ?? t?._text ?? "")))
    .join("");

  // Handle page number field in instrText
  if (instrText) {
    const instrStr =
      typeof instrText === "string" ? instrText : (instrText?._text ?? "");
    if (instrStr.trim().toUpperCase() === "PAGE") {
      const pageNum: PageNumNode = { tag: "pagenum", format: "decimal" };
      return { tag: "span", props, kids: [pageNum] };
    }
  }

  return buildSpan(content, props);
}

function decodeGrid(tbl: any, ctx: DecCtx): GridNode {
  // Parse tblPr for table styles
  const tblPr = tbl?.["w:tblPr"]?.[0] ?? tbl?.tblPr?.[0] ?? {};
  const tblLookAttr =
    tblPr?.["w:tblLook"]?.[0]?._attr ?? tblPr?.tblLook?.[0]?._attr ?? {};

  const look: TableLook = {
    firstRow: tblLookAttr?.["w:firstRow"] === "1" || undefined,
    lastRow: tblLookAttr?.["w:lastRow"] === "1" || undefined,
    firstCol:
      tblLookAttr?.["w:firstColumn"] === "1" ||
      tblLookAttr?.["w:firstCol"] === "1" ||
      undefined,
    lastCol:
      tblLookAttr?.["w:lastColumn"] === "1" ||
      tblLookAttr?.["w:lastCol"] === "1" ||
      undefined,
    bandedRows: tblLookAttr?.["w:noHBand"] === "0" || undefined,
    bandedCols: tblLookAttr?.["w:noVBand"] === "0" || undefined,
  };

  // Parse table borders for defaultStroke
  const tblBorders = tblPr?.["w:tblBorders"]?.[0] ?? tblPr?.tblBorders?.[0];
  let defaultStroke = undefined;
  if (tblBorders) {
    const top =
      tblBorders?.["w:top"]?.[0]?._attr ?? tblBorders?.top?.[0]?._attr;
    if (top) {
      defaultStroke = safeStrokeDocx(
        top?.["w:val"] ?? top?.val,
        Number(top?.["w:sz"] ?? top?.sz ?? 4),
        top?.["w:color"] ?? top?.color,
      );
    }
  }

  const gridProps: GridProps = { look, defaultStroke };

  // Read column widths from w:tblGrid
  const tblGrid = tbl?.["w:tblGrid"]?.[0] ?? tbl?.tblGrid?.[0];
  if (tblGrid) {
    const gridCols = toArr(tblGrid?.["w:gridCol"] ?? tblGrid?.gridCol ?? []);
    const colWidthsPt = gridCols
      .map((gc: any) =>
        Metric.dxaToPt(Number(gc?._attr?.["w:w"] ?? gc?._attr?.w ?? 0)),
      )
      .filter((w: number) => w > 0);
    if (colWidthsPt.length > 0) gridProps.colWidths = colWidthsPt;
  }

  const rowArr = toArr(tbl?.["w:tr"] ?? tbl?.tr);

  // ── Pass 1: parse raw cells with vMerge info ──
  interface RawCell {
    cell: any;
    gridSpan: number;
    vMergeRestart: boolean;
    vMergeContinue: boolean;
  }
  const rawGrid: RawCell[][] = rowArr.map((row: any) => {
    const cellArr = toArr(row?.["w:tc"] ?? row?.tc);
    return cellArr.map((cell: any): RawCell => {
      const tcPr = cell?.["w:tcPr"]?.[0] ?? {};
      const gridSpan = Number(tcPr?.["w:gridSpan"]?.[0]?._attr?.["w:val"] ?? 1);
      const vMergeNode = tcPr?.["w:vMerge"]?.[0];
      const vMergeVal = vMergeNode?._attr?.["w:val"] ?? vMergeNode?._attr?.val;
      const vMergeRestart = vMergeVal === "restart";
      // vMerge present but val is not "restart" → continuation cell
      const vMergeContinue = vMergeNode != null && !vMergeRestart;
      return { cell, gridSpan, vMergeRestart, vMergeContinue };
    });
  });

  // ── Pass 2: compute rowSpan for restart cells ──
  // rsMap[ri][ci] = computed rowSpan (only set for restart cells)
  const rsMap: Map<string, number> = new Map();
  for (let ri = 0; ri < rawGrid.length; ri++) {
    let gridCol = 0;
    for (let ci = 0; ci < rawGrid[ri].length; ci++) {
      const rc = rawGrid[ri][ci];
      if (rc.vMergeRestart) {
        let span = 1;
        for (let nr = ri + 1; nr < rawGrid.length; nr++) {
          // Find the cell at the same grid column in the next row
          let col = 0;
          let found = false;
          for (const nc of rawGrid[nr]) {
            if (col === gridCol && nc.vMergeContinue) {
              span++;
              found = true;
              break;
            }
            col += nc.gridSpan;
          }
          if (!found) break;
        }
        rsMap.set(`${ri},${ci}`, span);
      }
      gridCol += rc.gridSpan;
    }
  }

  // ── Pass 3: build CellNodes, skip continuation cells ──
  const rowNodes = rawGrid.map((rawRow, ri) => {
    // Check for header row
    const row = rowArr[ri];
    const trPr = row?.["w:trPr"]?.[0] ?? row?.trPr?.[0] ?? {};
    const isHeaderRow =
      trPr?.["w:tblHeader"]?.[0] != null || trPr?.tblHeader?.[0] != null;
    if (ri === 0 && isHeaderRow) gridProps.headerRow = true;

    // Row height from w:trHeight
    let rowHeightPt: number | undefined;
    const trHAttr = trPr?.["w:trHeight"]?.[0]?._attr ?? trPr?.trHeight?.[0]?._attr;
    if (trHAttr) {
      const hDxa = Number(trHAttr?.["w:val"] ?? trHAttr?.val ?? 0);
      if (hDxa > 0) rowHeightPt = Metric.dxaToPt(hDxa);
    }

    const cellNodes: CellNode[] = [];
    for (let ci = 0; ci < rawRow.length; ci++) {
      const rc = rawRow[ci];
      // Skip continuation cells — they are part of a vertical merge
      if (rc.vMergeContinue) continue;

      const cell = rc.cell;
      const tcPr = cell?.["w:tcPr"]?.[0] ?? {};

      // Cell background
      const bgAttr = tcPr?.["w:shd"]?.[0]?._attr ?? {};
      const bg = safeHex(bgAttr?.["w:fill"] ?? bgAttr?.fill);

      // Cell borders
      const tcBorders = tcPr?.["w:tcBorders"]?.[0] ?? tcPr?.tcBorders?.[0];
      const cp: CellProps = { bg, isHeader: isHeaderRow || undefined };

      if (tcBorders) {
        const dirs: Array<[string, "top" | "bot" | "left" | "right"]> = [
          ["top", "top"],
          ["bottom", "bot"],
          ["left", "left"],
          ["right", "right"],
        ];
        for (const [xmlTag, propKey] of dirs) {
          const bdr =
            tcBorders?.["w:" + xmlTag]?.[0]?._attr ??
            tcBorders?.[xmlTag]?.[0]?._attr;
          if (bdr) {
            cp[propKey] = safeStrokeDocx(
              bdr?.["w:val"] ?? bdr?.val,
              Number(bdr?.["w:sz"] ?? bdr?.sz ?? 4),
              bdr?.["w:color"] ?? bdr?.color,
            );
          }
        }
      }

      // Vertical alignment
      const vaAttr =
        tcPr?.["w:vAlign"]?.[0]?._attr ?? tcPr?.vAlign?.[0]?._attr ?? {};
      const vaVal = vaAttr?.["w:val"] ?? vaAttr?.val;
      if (vaVal) {
        const vaMap: Record<string, "top" | "mid" | "bot"> = {
          top: "top",
          center: "mid",
          bottom: "bot",
        };
        cp.va = vaMap[vaVal];
      }

      const rs = rsMap.get(`${ri},${ci}`) ?? 1;

      const paras = toArr(cell?.["w:p"] ?? cell?.p).map((p: any) =>
        decodePara(p, ctx),
      );
      cellNodes.push(
        buildCell(paras.length > 0 ? paras : [buildPara([buildSpan("")])], {
          cs: rc.gridSpan,
          rs,
          props: cp,
        }),
      );
    }
    return buildRow(cellNodes, rowHeightPt);
  });
  return buildGrid(rowNodes, gridProps);
}

function decodeGridSimple(tbl: any): GridNode {
  const rowArr = toArr(tbl?.["w:tr"] ?? tbl?.tr);
  const rowNodes = rowArr.map((row: any) => {
    const cellArr = toArr(row?.["w:tc"] ?? row?.tc);
    return buildRow(
      cellArr.map((c: any) => buildCell([buildPara([buildSpan(cellText(c))])])),
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
  return toArr(cell?.["w:p"] ?? cell?.p)
    .map((p: any) =>
      toArr(p?.["w:r"] ?? p?.r)
        .map((r: any) =>
          toArr(r?.["w:t"] ?? r?.t)
            .map((t: any) => (typeof t === "string" ? t : (t?._ ?? "")))
            .join(""),
        )
        .join(""),
    )
    .join(" ");
}

function tableText(tbl: any): string {
  return toArr(tbl?.["w:tr"] ?? tbl?.tr)
    .map((row: any) =>
      toArr(row?.["w:tc"] ?? row?.tc)
        .map((c: any) => cellText(c))
        .join("\t"),
    )
    .join("\n");
}

function parseHeading(style?: string): 1 | 2 | 3 | 4 | 5 | 6 | undefined {
  if (!style) return undefined;
  const m = style.match(/[Hh]eading(\d)/);
  if (m) {
    const n = Number(m[1]);
    if (n >= 1 && n <= 6) return n as any;
  }
  return undefined;
}

registry.registerDecoder(new DocxDecoder());

// ─── Anchor layout 추출 ─────────────────────────────────────

function extractAnchorLayout(anchor: any): ImgLayout {
  const attr = anchor?._attr ?? {};
  const behindDoc = attr.behindDoc === '1';

  // 텍스트 감싸기 타입
  let wrap: ImgWrap = 'square';
  if (anchor?.['wp:wrapNone']?.[0] != null)         wrap = behindDoc ? 'behind' : 'none';
  else if (anchor?.['wp:wrapTight']?.[0] != null)   wrap = 'tight';
  else if (anchor?.['wp:wrapThrough']?.[0] != null) wrap = 'through';
  else if (anchor?.['wp:wrapSquare']?.[0] != null)  wrap = 'square';
  else if (anchor?.['wp:wrapTopAndBottom']?.[0] != null) wrap = 'square';
  else if (anchor?.['wp:wrapBehind']?.[0] != null || behindDoc) wrap = 'behind';

  // 가로 위치
  const posH = anchor?.['wp:positionH']?.[0];
  const horzRelTo = parseHorzRelTo(posH?._attr?.relativeFrom);
  const horzAlignTxt = posH?.['wp:align']?.[0]?._text;
  const horzOffsetTxt = posH?.['wp:posOffset']?.[0]?._text;
  const horzAlign = horzAlignTxt ? parseHorzAlign(horzAlignTxt) : undefined;
  const xPt = horzOffsetTxt && !horzAlignTxt
    ? Metric.emuToPt(Number(horzOffsetTxt))
    : undefined;

  // 세로 위치
  const posV = anchor?.['wp:positionV']?.[0];
  const vertRelTo = parseVertRelTo(posV?._attr?.relativeFrom);
  const vertAlignTxt = posV?.['wp:align']?.[0]?._text;
  const vertOffsetTxt = posV?.['wp:posOffset']?.[0]?._text;
  const vertAlign = vertAlignTxt ? parseVertAlign(vertAlignTxt) : undefined;
  const yPt = vertOffsetTxt && !vertAlignTxt
    ? Metric.emuToPt(Number(vertOffsetTxt))
    : undefined;

  // 텍스트와의 거리
  const distT = attr.distT ? Metric.emuToPt(Number(attr.distT)) : undefined;
  const distB = attr.distB ? Metric.emuToPt(Number(attr.distB)) : undefined;
  const distL = attr.distL ? Metric.emuToPt(Number(attr.distL)) : undefined;
  const distR = attr.distR ? Metric.emuToPt(Number(attr.distR)) : undefined;
  const zOrder = attr.relativeHeight ? Number(attr.relativeHeight) : undefined;

  return { wrap, horzAlign, vertAlign, horzRelTo, vertRelTo, xPt, yPt, distT, distB, distL, distR, behindDoc, zOrder };
}

const HORZ_RELTO_MAP: Record<string, ImgHorzRelTo> = {
  margin: 'margin', leftMargin: 'margin', rightMargin: 'margin',
  insideMargin: 'margin', outsideMargin: 'margin',
  column: 'column', page: 'page', character: 'para', paragraph: 'para',
};
const VERT_RELTO_MAP: Record<string, ImgVertRelTo> = {
  margin: 'margin', topMargin: 'margin', bottomMargin: 'margin',
  insideMargin: 'margin', outsideMargin: 'margin',
  line: 'line', page: 'page', paragraph: 'para',
};
const HORZ_ALIGN_MAP: Record<string, ImgHorzAlign> = {
  left: 'left', center: 'center', right: 'right',
  inside: 'left', outside: 'right',
};
const VERT_ALIGN_MAP: Record<string, ImgVertAlign> = {
  top: 'top', center: 'center', bottom: 'bottom',
  inside: 'top', outside: 'bottom',
};

function parseHorzRelTo(v?: string): ImgHorzRelTo { return HORZ_RELTO_MAP[v ?? ''] ?? 'column'; }
function parseVertRelTo(v?: string): ImgVertRelTo { return VERT_RELTO_MAP[v ?? ''] ?? 'para'; }
function parseHorzAlign(v?: string): ImgHorzAlign | undefined { return HORZ_ALIGN_MAP[v ?? '']; }
function parseVertAlign(v?: string): ImgVertAlign | undefined { return VERT_ALIGN_MAP[v ?? '']; }
