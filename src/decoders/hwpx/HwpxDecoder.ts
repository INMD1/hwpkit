import type { Decoder } from '../../contract/decoder';
import type { DocRoot, ContentNode, ParaNode, SpanNode, GridNode, ImgNode, PageNumNode } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import type { DocMeta, PageDims, TextProps, ParaProps, CellProps, GridProps, Stroke, ImgLayout, ImgWrap, ImgHorzAlign, ImgVertAlign, ImgHorzRelTo, ImgVertRelTo } from '../../model/doc-props';
import { A4 } from '../../model/doc-props';
import { succeed, fail } from '../../contract/result';
import { buildRoot, buildSheet, buildPara, buildSpan, buildImg, buildGrid, buildRow, buildCell, buildPb } from '../../model/builders';
import { ShieldedParser } from '../../safety/ShieldedParser';
import { Metric, safeAlign, safeFont, safeHex, safeStrokeHwpx } from '../../safety/StyleBridge';
import { ArchiveKit } from '../../toolkit/ArchiveKit';
import { XmlKit } from '../../toolkit/XmlKit';
import { TextKit } from '../../toolkit/TextKit';
import { registry } from '../../pipeline/registry';

interface BorderFillInfo {
  stroke?: Stroke;
  bgColor?: string;
}

interface CharPrInfo {
  b?: boolean; i?: boolean; u?: boolean; s?: boolean;
  pt?: number; color?: string; font?: string; bg?: string;
}

interface ParaPrInfo {
  align?: string;
  indentPt?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  lineHeight?: number;
}

interface DecCtx {
  files: Map<string, Uint8Array>;
  shield: ShieldedParser;
  borderFills: Map<number, BorderFillInfo>;
  charPrs: Map<number, CharPrInfo>;
  paraPrs: Map<number, ParaPrInfo>;
  warns: string[];
}

export class HwpxDecoder implements Decoder {
  readonly format = 'hwpx';

  async decode(data: Uint8Array): Promise<Outcome<DocRoot>> {
    const shield = new ShieldedParser();
    const warns: string[] = [];

    try {
      const files = await ArchiveKit.unzip(data);

      const bodyXml = files.get('Contents/section0.xml')
        ?? files.get('section0.xml')
        ?? findSectionFile(files);

      if (!bodyXml) return fail('HWPX: section0.xml not found in archive');

      const headXml = files.get('Contents/header.xml') ?? files.get('header.xml');

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

      const ctx: DecCtx = { files, shield, borderFills, charPrs, paraPrs, warns };

      const bodyStr = TextKit.decode(bodyXml);
      const bodyObj: any = await XmlKit.parseStrict(bodyStr);

      const sections = normalizeSections(bodyObj);
      const kids = shield.guardAll(
        sections,
        (sec: any) => decodeSection(sec, dims, ctx),
        () => buildSheet([buildPara([buildSpan('[섹션 파싱 실패]')])], dims),
        'hwpx:section',
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

function findSectionFile(files: Map<string, Uint8Array>): Uint8Array | undefined {
  for (const [key, val] of files) {
    if (key.toLowerCase().includes('section') && key.endsWith('.xml')) return val;
  }
  return undefined;
}

function normalizeSections(bodyObj: any): any[] {
  // <hs:sec> (real HWPX), <hp:SEC> (legacy)
  if (bodyObj?.['hs:sec']) return toArr(bodyObj['hs:sec']);
  if (bodyObj?.['hp:SEC']) return toArr(bodyObj['hp:SEC']);

  const root = bodyObj?.['hp:HWPML'] ?? bodyObj?.HWPML ?? bodyObj;
  const body = root?.['hp:BODY']?.[0] ?? root?.BODY?.[0] ?? root?.['hp:BODY'] ?? root?.BODY;
  if (!body) return [bodyObj];
  const sections = body?.['hp:SECTION'] ?? body?.SECTION ?? [];
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
    const root = headObj?.['hh:head']?.[0] ?? headObj?.['hh:HEAD']?.[0] ?? headObj?.HEAD?.[0] ?? headObj;
    const info = root?.['hh:DOCSUMMARY']?.[0] ?? root?.DOCSUMMARY?.[0];
    if (!info) return {};
    const a = (k: string) => info?.[`hh:${k}`]?.[0]?._text ?? info?.[k]?.[0]?._text ?? '';
    return { title: a('TITLE') || undefined, author: a('AUTHOR') || undefined, subject: a('SUBJECT') || undefined };
  } catch { return {}; }
}

function extractDims(headObj: any): PageDims | null {
  try {
    const root = headObj?.['hh:head']?.[0] ?? headObj?.['hh:HEAD']?.[0] ?? headObj?.HEAD?.[0] ?? headObj;
    const refList = root?.['hh:refList']?.[0] ?? root?.['hh:REFLIST']?.[0] ?? root?.REFLIST?.[0];
    if (!refList) return null;

    const secPrList = refList?.['hh:SECPRLST']?.[0]?.['hh:SECPR']
      ?? refList?.SECPRLST?.[0]?.SECPR;
    const sec = Array.isArray(secPrList) ? secPrList[0] : secPrList;
    if (!sec) return null;

    const pa = sec?.['hh:PAGEPROPERTY']?.[0]?._attr ?? sec?.PAGEPROPERTY?.[0]?._attr;
    if (!pa) return null;

    return {
      wPt:    Metric.hwpToPt(Number(pa.Width ?? 59528)),
      hPt:    Metric.hwpToPt(Number(pa.Height ?? 84188)),
      mt:     Metric.hwpToPt(Number(pa.TopMargin ?? 5670)),
      mb:     Metric.hwpToPt(Number(pa.BottomMargin ?? 4252)),
      ml:     Metric.hwpToPt(Number(pa.LeftMargin ?? 8504)),
      mr:     Metric.hwpToPt(Number(pa.RightMargin ?? 8504)),
      orient: Number(pa.Landscape) === 1 ? 'landscape' : 'portrait',
    };
  } catch { return null; }
}

function extractBorderFills(headObj: any): Map<number, BorderFillInfo> {
  const map = new Map<number, BorderFillInfo>();
  try {
    const root = headObj?.['hh:head']?.[0] ?? headObj?.['hh:HEAD']?.[0] ?? headObj?.HEAD?.[0] ?? headObj;
    const refList = root?.['hh:refList']?.[0] ?? root?.['hh:REFLIST']?.[0] ?? root?.REFLIST?.[0];
    if (!refList) return map;

    const bfList = refList?.['hh:borderFills']?.[0] ?? refList?.['hh:BORDERFILLLIST']?.[0] ?? refList?.BORDERFILLLIST?.[0];
    if (!bfList) return map;

    const bfs = getTag(bfList, 'hh:borderFill', 'hh:BORDERFILL');
    for (const bf of bfs) {
      const attr = bf?._attr ?? {};
      const id = Number(attr.id ?? 0);
      if (id === 0) continue;

      const info: BorderFillInfo = {};

      // Parse border (take top as representative)
      const top = bf?.['hh:topBorder']?.[0]?._attr ?? bf?.['hh:top']?.[0]?._attr ?? bf?.top?.[0]?._attr;
      if (top) {
        // width is in mm (e.g. "0.18 mm"), convert mm → pt (1mm ≈ 2.835pt), then pt → hwp (*100) for safeStrokeHwpx
        const mmVal = parseFloat(top.width) || undefined;
        const hwpVal = mmVal != null ? mmVal * 2.835 * 100 : undefined;
        info.stroke = safeStrokeHwpx(top.type, hwpVal, top.color);
      }

      // Parse fill (real HWPX uses hc:fillBrush, not hh:fillBrush)
      const fillBrush = bf?.['hc:fillBrush']?.[0] ?? bf?.['hh:fillBrush']?.[0] ?? bf?.['hh:fill']?.[0] ?? bf?.fill?.[0] ?? bf?.fillBrush?.[0];
      if (fillBrush) {
        const winBrush = fillBrush?.['hc:winBrush']?.[0]?._attr ?? fillBrush?.['hh:winBrush']?.[0]?._attr ?? fillBrush?.winBrush?.[0]?._attr;
        if (winBrush?.faceColor && winBrush.faceColor !== 'none') {
          info.bgColor = safeHex(winBrush.faceColor);
        }
      }

      map.set(id, info);
    }
  } catch { /* non-fatal */ }
  return map;
}

function extractCharPrs(headObj: any): Map<number, CharPrInfo> {
  const map = new Map<number, CharPrInfo>();
  try {
    const root = headObj?.['hh:head']?.[0] ?? headObj?.['hh:HEAD']?.[0] ?? headObj?.HEAD?.[0] ?? headObj;
    const refList = root?.['hh:refList']?.[0] ?? root?.['hh:REFLIST']?.[0] ?? root?.REFLIST?.[0];
    if (!refList) return map;

    const cpList = refList?.['hh:charProperties']?.[0] ?? refList?.['hh:CHARPROPERTIES']?.[0];
    if (!cpList) return map;

    const cps = getTag(cpList, 'hh:charPr', 'hh:CHARPR');
    for (const cp of cps) {
      const attr = cp?._attr ?? {};
      const id = Number(attr.id ?? -1);
      if (id < 0) continue;

      const info: CharPrInfo = {};

      // height → pt
      if (attr.height) info.pt = Metric.hHeightToPt(Number(attr.height));

      // textColor
      if (attr.textColor) info.color = safeHex(attr.textColor);

      // bold
      if (cp?.['hh:bold']?.[0] != null) info.b = true;

      // italic
      if (cp?.['hh:italic']?.[0] != null) info.i = true;

      // underline
      const ulAttr = cp?.['hh:underline']?.[0]?._attr;
      if (ulAttr?.type && ulAttr.type !== 'NONE') info.u = true;

      // strikeout — shape="3D" is default "no strikeout" in real HWPX; only SOLID/etc means active
      const stAttr = cp?.['hh:strikeout']?.[0]?._attr;
      if (stAttr?.shape && stAttr.shape !== 'NONE' && stAttr.shape !== '3D') info.s = true;

      // font — from fontRef + fontface
      // (simplified: just store what we find)

      map.set(id, info);
    }
  } catch { /* non-fatal */ }
  return map;
}

function extractParaPrs(headObj: any): Map<number, ParaPrInfo> {
  const map = new Map<number, ParaPrInfo>();
  try {
    const root = headObj?.['hh:head']?.[0] ?? headObj?.['hh:HEAD']?.[0] ?? headObj?.HEAD?.[0] ?? headObj;
    const refList = root?.['hh:refList']?.[0] ?? root?.['hh:REFLIST']?.[0] ?? root?.REFLIST?.[0];
    if (!refList) return map;

    const ppList = refList?.['hh:paraProperties']?.[0] ?? refList?.['hh:PARAPROPERTIES']?.[0];
    if (!ppList) return map;

    const pps = getTag(ppList, 'hh:paraPr', 'hh:PARAPR');
    for (const pp of pps) {
      const attr = pp?._attr ?? {};
      const id = Number(attr.id ?? -1);
      if (id < 0) continue;

      const alignNode = pp?.['hh:align']?.[0]?._attr ?? pp?.['hh:ALIGN']?.[0]?._attr;
      const align = alignNode?.horizontal ?? alignNode?.Horizontal;

      // Read margin and lineSpacing from direct child OR hp:switch > hp:default/hp:case
      let marginEl = pp?.['hh:margin']?.[0] ?? null;
      let lineSpEl = pp?.['hh:lineSpacing']?.[0] ?? null;
      if (!marginEl) {
        const sw = pp?.['hp:switch']?.[0];
        const container = sw?.['hp:default']?.[0] ?? sw?.['hp:case']?.[0];
        marginEl = container?.['hh:margin']?.[0] ?? null;
        lineSpEl = lineSpEl ?? container?.['hh:lineSpacing']?.[0] ?? null;
      }

      let indentPt: number | undefined;
      let spaceBefore: number | undefined;
      let spaceAfter: number | undefined;
      let lineHeight: number | undefined;

      if (marginEl) {
        // Handle both hc:intent (our encoder) and hc:indent (Hancom standard)
        const intentEl = marginEl?.['hc:intent']?.[0] ?? marginEl?.['hc:indent']?.[0];
        const prevEl   = marginEl?.['hc:prev']?.[0];
        const nextEl   = marginEl?.['hc:next']?.[0];
        const intentVal = Number(intentEl?._attr?.value ?? 0);
        const prevVal   = Number(prevEl?._attr?.value ?? 0);
        const nextVal   = Number(nextEl?._attr?.value ?? 0);
        if (intentVal !== 0) indentPt    = Metric.hwpToPt(intentVal);
        if (prevVal   >  0)  spaceBefore = Metric.hwpToPt(prevVal);
        if (nextVal   >  0)  spaceAfter  = Metric.hwpToPt(nextVal);
      }

      if (lineSpEl) {
        const lsAttr = lineSpEl._attr ?? {};
        const lsType = lsAttr.type ?? 'PERCENT';
        const lsVal  = Number(lsAttr.value ?? 160);
        if (lsType === 'PERCENT' && lsVal > 0) lineHeight = lsVal / 100;
      }

      map.set(id, { align, indentPt, spaceBefore, spaceAfter, lineHeight });
    }
  } catch { /* non-fatal */ }
  return map;
}

// ─── Section decoding ──────────────────────────────────────

function addParaItems(p: any, items: { type: string; node: any }[]): void {
  // Check if this paragraph contains a table in its runs
  const runs = getTag(p, 'hp:run', 'hp:RUN');
  let hasTable = false;
  for (const run of runs) {
    const tbls = getTag(run, 'hp:tbl', 'hp:TABLE');
    for (const tbl of tbls) {
      items.push({ type: 'table', node: tbl });
      hasTable = true;
    }
  }
  // Also add as paragraph unless it's just a table container
  const hasText = runs.some((run: any) => {
    const ts = getTag(run, 'hp:t', 'hp:T', 'hp:CHAR');
    return ts.some((t: any) => {
      const text = typeof t === 'string' ? t : t?._text ?? '';
      return text.trim().length > 0;
    });
  });
  if (hasText || !hasTable) {
    items.push({ type: 'para', node: p });
  }
}

function decodeSection(sec: any, dims: PageDims, ctx: DecCtx) {
  // Try to extract dims from first paragraph's secPr
  const firstParas = getTag(sec, 'hp:p', 'hp:P');
  const pageDims = extractSecPrDims(firstParas[0]) ?? dims;

  // Build items list preserving document order via _childOrder
  const items: { type: string; node: any }[] = [];
  const paras = getTag(sec, 'hp:p', 'hp:P');
  const childOrder = sec?.['_childOrder'] as string[] | undefined;

  if (Array.isArray(childOrder)) {
    let pi = 0;
    for (const tag of childOrder) {
      if ((tag === 'hp:p' || tag === 'hp:P') && pi < paras.length) {
        const p = paras[pi++];
        addParaItems(p, items);
      }
    }
    // Append any remaining
    while (pi < paras.length) addParaItems(paras[pi++], items);
  } else {
    // No order info — process paragraphs sequentially
    for (const p of paras) addParaItems(p, items);
  }

  const kids: ContentNode[] = ctx.shield.guardAll(
    items,
    (item: any) => {
      if (item.type === 'table') {
        const { value } = ctx.shield.guardGrid(
          item.node,
          (n) => decodeGrid(n, ctx),
          (n) => decodeGridSimple(n, ctx),
          (n) => decodeGridFlat(n),
          (n) => decodeGridText(n) as unknown as GridNode,
          'hwpx:table',
        );
        return value;
      }
      return decodePara(item.node, ctx);
    },
    () => buildPara([buildSpan('[파싱 실패]')]),
    'hwpx:content',
  );

  // Decode header/footer
  const headerParas = decodeHeaderFooter(sec, 'header', ctx);
  const footerParas = decodeHeaderFooter(sec, 'footer', ctx);

  return buildSheet(
    kids.filter(Boolean) as ContentNode[],
    pageDims,
    { header: headerParas, footer: footerParas },
  );
}

function extractSecPrDims(p: any): PageDims | null {
  if (!p) return null;
  try {
    const runs = getTag(p, 'hp:run', 'hp:RUN');
    for (const run of runs) {
      const secPr = run?.['hp:secPr']?.[0] ?? run?.['hp:SECPR']?.[0];
      if (!secPr) continue;
      const pagePr = secPr?.['hp:pagePr']?.[0]?._attr ?? secPr?.['hp:PAGEPR']?.[0]?._attr;
      if (!pagePr) continue;
      const margin = secPr?.['hp:pagePr']?.[0]?.['hp:margin']?.[0]?._attr
        ?? secPr?.['hp:PAGEPR']?.[0]?.['hp:MARGIN']?.[0]?._attr ?? {};
      return {
        wPt:    Metric.hwpToPt(Number(pagePr.width ?? 59528)),
        hPt:    Metric.hwpToPt(Number(pagePr.height ?? 84188)),
        mt:     Metric.hwpToPt(Number(margin.top ?? 5670)),
        mb:     Metric.hwpToPt(Number(margin.bottom ?? 4252)),
        ml:     Metric.hwpToPt(Number(margin.left ?? 8504)),
        mr:     Metric.hwpToPt(Number(margin.right ?? 8504)),
        orient: pagePr.landscape === 'NARROWLY' ? 'landscape' : 'portrait',
      };
    }
  } catch { /* ignore */ }
  return null;
}

function decodeHeaderFooter(sec: any, kind: 'header' | 'footer', ctx: DecCtx): ParaNode[] | undefined {
  try {
    const hf = sec?.['hp:headerFooter']?.[0] ?? sec?.['hp:HEADERFOOTER']?.[0]
      ?? sec?.headerFooter?.[0] ?? sec?.HEADERFOOTER?.[0];
    if (!hf) return undefined;

    const part = hf?.['hp:' + kind]?.[0] ?? hf?.['hp:' + kind.toUpperCase()]?.[0]
      ?? hf?.[kind]?.[0] ?? hf?.[kind.toUpperCase()]?.[0];
    if (!part) return undefined;

    const paras = getTag(part, 'hp:p', 'hp:P');
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

  // Resolve paraPr from IDRef or inline
  let align: string | undefined;
  const paraPrDef = ctx.paraPrs.get(paraPrIdRef);
  if (paraPrDef?.align) align = paraPrDef.align;

  // Check inline PARAPR too
  const inlineParaPr = p?.['hp:PARAPR']?.[0] ?? p?.['hp:paraPr']?.[0] ?? p?.PARAPR?.[0];
  if (inlineParaPr) {
    const alignNode = inlineParaPr?.['hp:ALIGN']?.[0]?._attr ?? inlineParaPr?.['hp:align']?.[0]?._attr
      ?? inlineParaPr?.ALIGN?.[0]?._attr;
    if (alignNode?.Type) align = alignNode.Type;
    if (alignNode?.horizontal) align = alignNode.horizontal;
  }

  const inlineAttr = inlineParaPr?._attr ?? {};
  const props: ParaProps = { align: safeAlign(align) };

  // Apply spacing/indent/lineHeight from paraPr definition
  if (paraPrDef) {
    if (paraPrDef.indentPt    !== undefined) props.indentPt    = paraPrDef.indentPt;
    if (paraPrDef.spaceBefore !== undefined) props.spaceBefore = paraPrDef.spaceBefore;
    if (paraPrDef.spaceAfter  !== undefined) props.spaceAfter  = paraPrDef.spaceAfter;
    if (paraPrDef.lineHeight  !== undefined) props.lineHeight  = paraPrDef.lineHeight;
  }

  // List support (from inline attr)
  if (inlineAttr.listType) {
    props.listOrd = inlineAttr.listType === 'DIGIT' || inlineAttr.listType === 'DECIMAL';
    props.listLv = Number(inlineAttr.listLevel ?? 0);
  }

  const runs = getTag(p, 'hp:run', 'hp:RUN');
  const kids: (SpanNode | ImgNode)[] = [];

  for (const run of runs) {
    // Images inside run
    const pics = getTag(run, 'hp:pic', 'hp:PIC');
    for (const pic of pics) {
      const img = decodePic(pic, ctx);
      if (img) kids.push(img);
    }

    // Page number
    const pageNums = getTag(run, 'hp:pageNum', 'hp:PAGENUM');
    if (pageNums.length > 0) {
      const pn = pageNums[0]?._attr ?? {};
      const fmt = pn.formatType === 'ROMAN_LOWER' ? 'roman' as const
        : pn.formatType === 'ROMAN_UPPER' ? 'romanCaps' as const
        : 'decimal' as const;
      const pageNumNode: PageNumNode = { tag: 'pagenum', format: fmt };
      const spanProps = resolveCharPr(run, ctx);
      kids.push({ tag: 'span', props: spanProps, kids: [pageNumNode] });
      continue;
    }

    // Text
    const textNodes = getTag(run, 'hp:t', 'hp:T', 'hp:CHAR');
    const content = textNodes.map((t: any) => typeof t === 'string' ? t : t?._text ?? t?._ ?? '').join('');

    // Skip empty secPr-only runs
    if (content === '' && (run?.['hp:secPr']?.[0] || run?.['hp:SECPR']?.[0]) && pics.length === 0) continue;

    const spanProps = resolveCharPr(run, ctx);
    kids.push(buildSpan(content, spanProps));
  }

  // pageBreak="1" → prepend a pb node in its own span
  if (pAttr.pageBreak === '1') {
    kids.unshift({ tag: 'span', props: {}, kids: [buildPb()] });
  }

  return buildPara(kids.filter(Boolean) as ParaNode['kids'], props);
}

function resolveCharPr(run: any, ctx: DecCtx): TextProps {
  const runAttr = run?._attr ?? {};
  const charPrIdRef = Number(runAttr.charPrIDRef ?? -1);

  // Try IDRef first
  const def = ctx.charPrs.get(charPrIdRef);
  if (def) {
    return {
      b: def.b, i: def.i, u: def.u, s: def.s,
      pt: def.pt, color: def.color, font: def.font, bg: def.bg,
    };
  }

  // Fallback: inline CHARPR
  const ca = run?.['hp:CHARPR']?.[0]?._attr ?? run?.['hp:charPr']?.[0]?._attr ?? run?.CHARPR?.[0]?._attr ?? {};
  return {
    b:     ca.Bold === '1' || ca.Bold === 'true' || undefined,
    i:     ca.Italic === '1' || ca.Italic === 'true' || undefined,
    u:     ca.Underline ? ca.Underline !== 'NONE' : undefined,
    s:     ca.Strikeout ? ca.Strikeout !== 'NONE' : undefined,
    font:  safeFont(ca.FontName ?? ca.FaceNameHangul),
    pt:    ca.Height ? Metric.hHeightToPt(Number(ca.Height)) : undefined,
    color: safeHex(ca.TextColor),
    bg:    safeHex(ca.BgColor),
  };
}

// ─── Image decoding ────────────────────────────────────────

function decodePic(pic: any, ctx: DecCtx): ImgNode | null {
  try {
    const szAttr = pic?.['hp:sz']?.[0]?._attr ?? pic?.sz?.[0]?._attr ?? {};
    const w = Metric.hwpToPt(Number(szAttr.width ?? 0));
    const h = Metric.hwpToPt(Number(szAttr.height ?? 0));

    // Try multiple tag patterns for image reference
    const imgNode = pic?.['hp:img']?.[0]?._attr ?? pic?.['hc:img']?.[0]?._attr
      ?? pic?.img?.[0]?._attr ?? {};
    const binRef = imgNode.binaryItemIDRef ?? imgNode.BinaryItemIDRef;
    if (!binRef) return null;

    // Find binary data
    let imgData: Uint8Array | undefined;
    for (const [key, val] of ctx.files) {
      if (key.includes(binRef) || key.toLowerCase().includes(binRef.toLowerCase())) {
        imgData = val;
        break;
      }
    }
    if (!imgData) return null;

    const ext = binRef.split('.').pop()?.toLowerCase() ?? 'png';
    const mimeMap: Record<string, ImgNode['mime']> = {
      png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg',
      gif: 'image/gif', bmp: 'image/bmp',
    };

    // ── hp:pos에서 layout 추출 ───────────────────────────────
    const posAttr = pic?.['hp:pos']?.[0]?._attr ?? pic?.pos?.[0]?._attr ?? {};
    const layout = extractHwpxLayout(posAttr, pic);

    return buildImg(TextKit.base64Encode(imgData), mimeMap[ext] ?? 'image/png', w, h, undefined, layout);
  } catch {
    return null;
  }
}

function extractHwpxLayout(posAttr: any, pic: any): ImgLayout {
  const treatAsChar = posAttr.treatAsChar === '1' || posAttr.treatAsChar === 'true';
  if (treatAsChar) return { wrap: 'inline' };

  // textWrap → wrap
  const textWrap: string = (pic?._attr?.textWrap ?? pic?.['hp:pic']?.[0]?._attr?.textWrap ?? 'TOP_AND_BOTTOM');
  const wrapMap: Record<string, ImgWrap> = {
    TOP_AND_BOTTOM: 'square', BOTH_SIDES: 'tight',
    LARGER_ONLY: 'tight', SMALLER_ONLY: 'tight',
  };
  const wrap: ImgWrap = wrapMap[textWrap] ?? 'square';

  // 기준점
  const horzRelToMap: Record<string, ImgHorzRelTo> = {
    PARA: 'para', MARGIN: 'margin', PAGE: 'page', COLUMN: 'column',
  };
  const vertRelToMap: Record<string, ImgVertRelTo> = {
    PARA: 'para', MARGIN: 'margin', PAGE: 'page', LINE: 'line',
  };
  const horzRelTo = horzRelToMap[posAttr.horzRelTo ?? ''] ?? 'para';
  const vertRelTo = vertRelToMap[posAttr.vertRelTo ?? ''] ?? 'para';

  // 정렬
  const horzAlignMap: Record<string, ImgHorzAlign> = { LEFT: 'left', CENTER: 'center', RIGHT: 'right' };
  const vertAlignMap: Record<string, ImgVertAlign> = { TOP: 'top', CENTER: 'center', BOTTOM: 'bottom' };
  const horzAlign = horzAlignMap[posAttr.horzAlign ?? ''];
  const vertAlign = vertAlignMap[posAttr.vertAlign ?? ''];

  // 오프셋
  const horzOffset = Number(posAttr.horzOffset ?? 0);
  const vertOffset = Number(posAttr.vertOffset ?? 0);
  const xPt = horzOffset !== 0 ? Metric.hwpToPt(horzOffset) : undefined;
  const yPt = vertOffset !== 0 ? Metric.hwpToPt(vertOffset) : undefined;

  return { wrap, horzAlign, vertAlign, horzRelTo, vertRelTo, xPt, yPt };
}

// ─── Table decoding ────────────────────────────────────────

function decodeGrid(tbl: any, ctx: DecCtx): GridNode {
  const tblAttr = tbl?._attr ?? {};
  const borderFillId = Number(tblAttr.borderFillIDRef ?? 0);
  const borderFill = ctx.borderFills.get(borderFillId);
  const headerRow = tblAttr.repeatHeader === '1';

  const gridProps: GridProps = { headerRow: headerRow || undefined };
  if (borderFill?.stroke) gridProps.defaultStroke = borderFill.stroke;

  const rowArr = getTag(tbl, 'hp:tr', 'hp:ROW');

  // Read column widths from the first row that has all cs=1 cells
  for (const row of rowArr) {
    const cells = getTag(row, 'hp:tc', 'hp:CELL');
    const rowWidths: number[] = [];
    let allSingle = true;
    for (const cell of cells) {
      const cellSpanAttr = cell?.['hp:cellSpan']?.[0]?._attr ?? {};
      const cs = Number(cellSpanAttr.colSpan ?? cell?._attr?.ColSpan ?? 1);
      if (cs > 1) { allSingle = false; break; }
      const szAttr = cell?.['hp:cellSz']?.[0]?._attr ?? {};
      const w = Number(szAttr.width ?? 0);
      rowWidths.push(Metric.hwpToPt(w));
    }
    if (allSingle && rowWidths.length > 0 && rowWidths.some(w => w > 0)) {
      gridProps.colWidths = rowWidths;
      break;
    }
  }
  const rowNodes = rowArr.map((row: any) => {
    const cellArr = getTag(row, 'hp:tc', 'hp:CELL');
    const cellNodes = cellArr.map((cell: any) => {
      const ca = cell?._attr ?? {};

      // Cell borderFill
      const cellBfId = Number(ca.borderFillIDRef ?? 0);
      const cellBf = ctx.borderFills.get(cellBfId);

      const cellProps: CellProps = {
        bg: cellBf?.bgColor ?? safeHex(ca.BgColor),
      };

      if (cellBf?.stroke) {
        cellProps.top = cellBf.stroke;
        cellProps.bot = cellBf.stroke;
        cellProps.left = cellBf.stroke;
        cellProps.right = cellBf.stroke;
      }

      // Vertical alignment from subList
      const subList = cell?.['hp:subList']?.[0] ?? cell?.subList?.[0];
      const subAttr = subList?._attr ?? {};
      if (subAttr.vertAlign) {
        const vaMap: Record<string, 'top' | 'mid' | 'bot'> = {
          TOP: 'top', CENTER: 'mid', BOTTOM: 'bot',
        };
        cellProps.va = vaMap[subAttr.vertAlign];
      }

      // Colspan/rowspan from cellSpan element or attributes
      const cellSpan = cell?.['hp:cellSpan']?.[0]?._attr ?? {};
      const cs = Number(cellSpan.colSpan ?? ca.ColSpan ?? 1);
      const rs = Number(cellSpan.rowSpan ?? ca.RowSpan ?? 1);

      // Parse paragraphs
      let paras: ParaNode[];
      if (subList) {
        const subParas = getTag(subList, 'hp:p', 'hp:P');
        paras = subParas.map((p: any) => decodePara(p, ctx));
      } else {
        paras = getTag(cell, 'hp:p', 'hp:P').map((p: any) => decodePara(p, ctx));
      }

      return buildCell(
        paras.length > 0 ? paras : [buildPara([buildSpan('')])],
        { cs, rs, props: cellProps },
      );
    });
    return buildRow(cellNodes);
  });
  return buildGrid(rowNodes, gridProps);
}

function decodeGridSimple(tbl: any, ctx: DecCtx): GridNode {
  const rowArr = getTag(tbl, 'hp:tr', 'hp:ROW');
  const rowNodes = rowArr.map((row: any) => {
    const cellArr = getTag(row, 'hp:tc', 'hp:CELL');
    return buildRow(cellArr.map((cell: any) => buildCell([buildPara([buildSpan(cellText(cell))])])));
  });
  return buildGrid(rowNodes);
}

function decodeGridFlat(tbl: any): GridNode {
  return buildGrid([buildRow([buildCell([buildPara([buildSpan(tableText(tbl))])])])]);
}

function decodeGridText(tbl: any): ParaNode {
  return buildPara([buildSpan(tableText(tbl))]);
}

function cellText(cell: any): string {
  const subList = cell?.['hp:subList']?.[0] ?? cell?.subList?.[0];
  const source = subList ?? cell;
  return getTag(source, 'hp:p', 'hp:P').map((p: any) =>
    getTag(p, 'hp:run', 'hp:RUN').map((r: any) =>
      getTag(r, 'hp:t', 'hp:T').map((t: any) => typeof t === 'string' ? t : t?._text ?? t?._ ?? '').join(''),
    ).join(''),
  ).join(' ');
}

function tableText(tbl: any): string {
  return getTag(tbl, 'hp:tr', 'hp:ROW').map((row: any) =>
    getTag(row, 'hp:tc', 'hp:CELL').map((c: any) => cellText(c)).join('\t'),
  ).join('\n');
}

function toArr(v: any): any[] { return v == null ? [] : Array.isArray(v) ? v : [v]; }

// Auto-register
registry.registerDecoder(new HwpxDecoder());
