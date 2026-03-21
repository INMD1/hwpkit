import type { Encoder } from '../../contract/encoder';
import type { DocRoot, ParaNode, SpanNode, GridNode, ContentNode, ImgNode, SheetNode } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import type { PageDims, TextProps, ParaProps, CellProps, Stroke } from '../../model/doc-props';
import { A4, DEFAULT_STROKE } from '../../model/doc-props';
import { succeed, fail } from '../../contract/result';
import { Metric } from '../../safety/StyleBridge';
import { ArchiveKit } from '../../toolkit/ArchiveKit';
import { TextKit } from '../../toolkit/TextKit';
import { registry } from '../../pipeline/registry';

// ─── All HWPX namespaces ────────────────────────────────────
const NS = [
  'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app"',
  'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"',
  'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"',
  'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"',
  'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"',
  'xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page"',
  'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"',
  'xmlns:dc="http://purl.org/dc/elements/1.1/"',
  'xmlns:opf="http://www.idpf.org/2007/opf/"',
  'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"',
].join(' ');

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
  bg?: string;
}

interface ParaPrDef {
  id: number;
  align: string;
  listType?: string;
  listLevel?: number;
}

interface BinEntry { id: string; name: string; data: Uint8Array }

interface HwpxCtx {
  charPrs: CharPrDef[];
  charPrMap: Map<string, number>;
  paraPrs: ParaPrDef[];
  paraPrMap: Map<string, number>;
  borderFills: { id: number; xml: string }[];
  bins: BinEntry[];
  nextBinNum: number;
  availableWidth: number; // HWPUNIT — page width minus margins
}

function charPrKey(p: TextProps): string {
  return `${p.b ? 1 : 0}|${p.i ? 1 : 0}|${p.u ? 1 : 0}|${p.s ? 1 : 0}|${p.pt ?? 10}|${p.color ?? '000000'}|${p.font ?? ''}|${p.bg ?? ''}`;
}

function paraPrKey(p: ParaProps): string {
  return `${p.align ?? 'left'}|${p.listOrd ?? ''}|${p.listLv ?? 0}`;
}

function registerCharPr(props: TextProps, ctx: HwpxCtx): number {
  const key = charPrKey(props);
  const existing = ctx.charPrMap.get(key);
  if (existing !== undefined) return existing;

  const id = ctx.charPrs.length;
  const def: CharPrDef = {
    id,
    height: Metric.ptToHHeight(props.pt ?? 10),
    bold: !!props.b,
    italic: !!props.i,
    underline: props.u ? 'BOTTOM' : 'NONE',
    strikeout: props.s ? 'SOLID' : 'NONE',
    textColor: props.color ? `#${props.color}` : '#000000',
    fontName: props.font ?? '굴림체',
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
    align: (props.align ?? 'left').toUpperCase(),
  };
  if (props.listOrd !== undefined) {
    def.listType = props.listOrd ? 'DIGIT' : 'BULLET';
    def.listLevel = props.listLv ?? 0;
  }
  ctx.paraPrs.push(def);
  ctx.paraPrMap.set(key, id);
  return id;
}

// ─── Pre-scan: collect all charPr/paraPr used ───────────────

function scanContent(kids: ContentNode[], ctx: HwpxCtx): void {
  for (const kid of kids) {
    if (kid.tag === 'para') scanPara(kid, ctx);
    else if (kid.tag === 'grid') scanGrid(kid, ctx);
  }
}

function scanPara(para: ParaNode, ctx: HwpxCtx): void {
  registerParaPr(para.props, ctx);
  for (const kid of para.kids) {
    if (kid.tag === 'span') registerCharPr(kid.props, ctx);
    else if (kid.tag === 'img') registerImage(kid, ctx);
  }
}

function scanGrid(grid: GridNode, ctx: HwpxCtx): void {
  for (const row of grid.kids)
    for (const cell of row.kids)
      for (const p of cell.kids) scanPara(p, ctx);
}

function scanParas(paras: ParaNode[], ctx: HwpxCtx): void {
  for (const p of paras) scanPara(p, ctx);
}

// ─── Image handling ─────────────────────────────────────────

function mimeToExt(mime: string): string {
  if (mime.includes('jpeg')) return 'jpg';
  if (mime.includes('gif')) return 'gif';
  if (mime.includes('bmp')) return 'bmp';
  return 'png';
}

function registerImage(img: ImgNode, ctx: HwpxCtx): void {
  if ((img as any)._binId) return;
  const ext = mimeToExt(img.mime);
  const id = `BIN${String(ctx.nextBinNum).padStart(4, '0')}`;
  const name = `${id}.${ext}`;
  ctx.nextBinNum++;
  const data = TextKit.base64Decode(img.b64);
  ctx.bins.push({ id, name, data });
  (img as any)._binId = id;
}

// ─── BorderFill ─────────────────────────────────────────────

function addBorderFill(ctx: HwpxCtx, stroke?: Stroke, bgColor?: string): number {
  const id = ctx.borderFills.length + 1;
  const s = stroke ?? DEFAULT_STROKE;
  const kindMap: Record<string, string> = { solid: 'SOLID', dash: 'DASH', dot: 'DOT', double: 'DOUBLE', none: 'NONE' };
  const type = kindMap[s.kind] ?? 'SOLID';
  const w = `${(s.pt * 0.3528).toFixed(2)} mm`;
  const c = s.color.startsWith('#') ? s.color : `#${s.color}`;

  let fill = '';
  if (bgColor) {
    const bc = bgColor.startsWith('#') ? bgColor : `#${bgColor}`;
    fill = `<hc:fillBrush><hc:winBrush faceColor="${bc}" hatchColor="none" alpha="0"/></hc:fillBrush>`;
  }

  const xml = `<hh:borderFill id="${id}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0"><hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/><hh:leftBorder type="${type}" width="${w}" color="${c}"/><hh:rightBorder type="${type}" width="${w}" color="${c}"/><hh:topBorder type="${type}" width="${w}" color="${c}"/><hh:bottomBorder type="${type}" width="${w}" color="${c}"/><hh:diagonal type="NONE" width="0.1 mm" color="#000000"/>${fill}</hh:borderFill>`;
  ctx.borderFills.push({ id, xml });
  return id;
}

// ─── Encoder class ──────────────────────────────────────────

export class HwpxEncoder implements Encoder {
  readonly format = 'hwpx';

  async encode(doc: DocRoot): Promise<Outcome<Uint8Array>> {
    try {
      const sheet = doc.kids[0];
      const dims = sheet?.dims ?? A4;

      // Available width = page width - left margin - right margin (in HWPUNIT)
      const availableWidth = Metric.ptToHwp(dims.wPt) - Metric.ptToHwp(dims.ml) - Metric.ptToHwp(dims.mr);

      const ctx: HwpxCtx = {
        charPrs: [], charPrMap: new Map(),
        paraPrs: [], paraPrMap: new Map(),
        borderFills: [], bins: [], nextBinNum: 1,
        availableWidth,
      };

      // Default borderFill (id=1, no border)
      addBorderFill(ctx, { kind: 'none', pt: 0.1, color: '000000' });
      // Table border borderFill (id=2)
      addBorderFill(ctx, DEFAULT_STROKE);
      // Default no-border for text areas (id=3)
      addBorderFill(ctx, { kind: 'none', pt: 0.1, color: '000000' });

      // Register default charPr (id=0) and paraPr (id=0)
      registerCharPr({}, ctx);
      registerParaPr({}, ctx);

      // Pre-scan all content to collect charPr/paraPr/images
      scanContent(sheet?.kids ?? [], ctx);
      if (sheet?.header) scanParas(sheet.header, ctx);
      if (sheet?.footer) scanParas(sheet.footer, ctx);

      const entries: { name: string; data: Uint8Array }[] = [
        { name: 'mimetype',                data: TextKit.encode('application/hwp+zip') },
        { name: 'version.xml',             data: TextKit.encode(VERSION_XML) },
        { name: 'META-INF/container.xml',   data: TextKit.encode(containerXml()) },
        { name: 'META-INF/manifest.xml',    data: TextKit.encode(MANIFEST_XML) },
        { name: 'settings.xml',            data: TextKit.encode(SETTINGS_XML) },
        { name: 'Contents/content.hpf',     data: TextKit.encode(contentHpf(ctx)) },
        { name: 'Contents/header.xml',      data: TextKit.encode(headerXml(dims, doc.meta, ctx)) },
        { name: 'Contents/section0.xml',    data: TextKit.encode(sectionXml(sheet, dims, ctx)) },
      ];

      for (const bin of ctx.bins) {
        entries.push({ name: `BinData/${bin.name}`, data: bin.data });
      }

      return succeed(await ArchiveKit.zip(entries));
    } catch (e: any) {
      return fail(`HWPX encode error: ${e?.message ?? String(e)}`);
    }
  }
}

// ─── Constants ──────────────────────────────────────────────

const VERSION_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="0" buildNumber="1" os="1" xmlVersion="1.4" application="hwpkit" appVersion="2, 0, 0, 0"/>`;

function containerXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"><ocf:rootfiles><ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/></ocf:rootfiles></ocf:container>`;
}

const MANIFEST_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>`;

const SETTINGS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ha:HWPApplicationSetting xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"/>`;

function contentHpf(ctx: HwpxCtx): string {
  let items = `<opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>
    <opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>`;
  for (const bin of ctx.bins) {
    const ext = bin.name.split('.').pop()?.toLowerCase() ?? 'png';
    const ct = ext === 'png' ? 'image/png' : ext === 'jpg' ? 'image/jpeg' : ext === 'gif' ? 'image/gif' : 'image/bmp';
    items += `\n    <opf:item id="${bin.id}" href="BinData/${bin.name}" media-type="${ct}" isEmbeded="1"/>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><opf:package ${NS} version="" unique-identifier="" id=""><opf:metadata><opf:title/><opf:language>ko</opf:language></opf:metadata><opf:manifest>
    ${items}
  </opf:manifest><opf:spine><opf:itemref idref="section0"/></opf:spine></opf:package>`;
}

// ─── header.xml ─────────────────────────────────────────────

function headerXml(dims: PageDims, meta: any, ctx: HwpxCtx): string {
  // Font face definitions
  const defaultFont = '굴림체';
  const fontFaces = `<hh:fontfaces itemCnt="1"><hh:fontface lang="HANGUL" fontCnt="1"><hh:font id="0" face="${defaultFont}" type="TTF" isEmbedded="0"><hh:typeInfo familyType="FCAT_GOTHIC" weight="6" proportion="0" contrast="0" strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/></hh:font></hh:fontface><hh:fontface lang="LATIN" fontCnt="1"><hh:font id="0" face="${defaultFont}" type="TTF" isEmbedded="0"><hh:typeInfo familyType="FCAT_GOTHIC" weight="6" proportion="0" contrast="0" strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/></hh:font></hh:fontface></hh:fontfaces>`;

  // CharPr definitions
  let charPrXml = '';
  for (const cp of ctx.charPrs) {
    const bold = cp.bold ? '<hh:bold/>' : '';
    const italic = cp.italic ? '<hh:italic/>' : '';
    charPrXml += `<hh:charPr id="${cp.id}" height="${cp.height}" textColor="${cp.textColor}" shadeColor="none" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="3"><hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/><hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/><hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/><hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/><hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>${bold}${italic}<hh:underline type="${cp.underline}" shape="SOLID" color="#000000"/><hh:strikeout shape="${cp.strikeout === 'NONE' ? 'NONE' : 'SOLID'}" color="#000000"/><hh:outline type="NONE"/><hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/></hh:charPr>`;
  }

  // ParaPr definitions
  let paraPrXml = '';
  for (const pp of ctx.paraPrs) {
    paraPrXml += `<hh:paraPr id="${pp.id}" tabPrIDRef="0" condense="0" fontLineHeight="0" snapToGrid="0" suppressLineNumbers="0" checked="0"><hh:align horizontal="${pp.align}" vertical="BASELINE"/><hh:heading type="NONE" idRef="0" level="0"/><hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="BREAK_WORD" widowOrphan="0" keepWithNext="0" keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/><hh:autoSpacing eAsianEng="0" eAsianNum="0"/><hh:margin><hc:intent value="0" unit="HWPUNIT"/><hc:left value="0" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/><hc:prev value="0" unit="HWPUNIT"/><hc:next value="0" unit="HWPUNIT"/></hh:margin><hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/><hh:border borderFillIDRef="1" offsetLeft="0" offsetRight="0" offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/></hh:paraPr>`;
  }

  // BorderFill definitions
  const borderFillXml = ctx.borderFills.map(bf => bf.xml).join('');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hh:head ${NS} version="1.4" secCnt="1"><hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/><hh:refList>${fontFaces}<hh:borderFills itemCnt="${ctx.borderFills.length}">${borderFillXml}</hh:borderFills><hh:charProperties itemCnt="${ctx.charPrs.length}">${charPrXml}</hh:charProperties><hh:paraProperties itemCnt="${ctx.paraPrs.length}">${paraPrXml}</hh:paraProperties><hh:tabProperties itemCnt="1"><hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/></hh:tabProperties><hh:styles itemCnt="1"><hh:style id="0" type="PARA" name="Normal" engName="Normal" paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0" langIDRef="0" lockForm="0"/></hh:styles></hh:refList><hh:compatibleDocument targetProgram="NONE"/><hh:docOption linkinfo="" gridMethod="0" memoOption="3" headFootFlags="0"/></hh:head>`;
}

// ─── section0.xml ───────────────────────────────────────────

function sectionXml(sheet: SheetNode | undefined, dims: PageDims, ctx: HwpxCtx): string {
  const kids = sheet?.kids ?? [];

  // First paragraph includes secPr
  // WIDELY = portrait (standard), NARROWLY = landscape
  const secPr = `<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="0" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0"><hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/><hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/><hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/><hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/><hp:pagePr landscape="${dims.orient === 'landscape' ? 'NARROWLY' : 'WIDELY'}" width="${Metric.ptToHwp(dims.wPt)}" height="${Metric.ptToHwp(dims.hPt)}" gutterType="LEFT_ONLY"><hp:margin header="0" footer="0" gutter="0" left="${Metric.ptToHwp(dims.ml)}" right="${Metric.ptToHwp(dims.mr)}" top="${Metric.ptToHwp(dims.mt)}" bottom="${Metric.ptToHwp(dims.mb)}"/></hp:pagePr><hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/><hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="284" belowLine="568" aboveLine="852"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" beneathText="0"/></hp:footNotePr><hp:endNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/><hp:noteLine length="0" type="NONE" width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="0" belowLine="576" aboveLine="864"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr><hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill></hp:secPr>`;

  let contentXml = '';
  let isFirst = true;

  for (const kid of kids) {
    if (kid.tag === 'para') {
      contentXml += encodePara(kid, ctx, isFirst ? secPr : '');
      isFirst = false;
    } else if (kid.tag === 'grid') {
      // Grid is embedded inside a paragraph's run
      if (isFirst) {
        contentXml += `<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0">${secPr}${encodeGrid(kid, ctx)}</hp:run></hp:p>`;
        isFirst = false;
      } else {
        contentXml += `<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0">${encodeGrid(kid, ctx)}</hp:run></hp:p>`;
      }
    }
  }

  // If empty, add one empty paragraph with secPr
  if (contentXml === '') {
    contentXml = `<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0">${secPr}</hp:run></hp:p>`;
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec ${NS}>${contentXml}</hs:sec>`;
}

function encodePara(para: ParaNode, ctx: HwpxCtx, secPr: string = ''): string {
  const paraPrId = registerParaPr(para.props, ctx);

  let runs = '';
  for (const kid of para.kids) {
    if (kid.tag === 'span') {
      runs += encodeRun(kid, ctx);
    } else if (kid.tag === 'img') {
      runs += encodeImage(kid, ctx);
    }
  }

  // If no runs, add default empty run
  if (runs === '') {
    runs = `<hp:run charPrIDRef="0"><hp:t></hp:t></hp:run>`;
  }

  // Inject secPr into first run
  if (secPr) {
    // Insert secPr after the first charPrIDRef attribute close
    const firstRunEnd = runs.indexOf('>');
    if (firstRunEnd > -1) {
      runs = runs.substring(0, firstRunEnd + 1) + secPr + runs.substring(firstRunEnd + 1);
    }
  }

  return `<hp:p id="0" paraPrIDRef="${paraPrId}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">${runs}</hp:p>`;
}

function encodeRun(span: SpanNode, ctx: HwpxCtx): string {
  const charPrId = registerCharPr(span.props, ctx);

  const parts: string[] = [];
  for (const kid of span.kids) {
    if (kid.tag === 'txt') {
      parts.push(`<hp:t>${esc(kid.content)}</hp:t>`);
    } else if (kid.tag === 'pagenum') {
      const fmt = kid.format === 'roman' ? 'ROMAN_LOWER' : kid.format === 'romanCaps' ? 'ROMAN_UPPER' : 'DIGIT';
      parts.push(`<hp:pageNum pageStartsOn="BOTH" formatType="${fmt}"/>`);
    } else if (kid.tag === 'br') {
      parts.push(`<hp:t>\n</hp:t>`);
    }
  }

  return `<hp:run charPrIDRef="${charPrId}">${parts.join('')}</hp:run>`;
}

function encodeImage(img: ImgNode, ctx: HwpxCtx): string {
  const binId = (img as any)._binId;
  if (!binId) return '';

  const charPrId = registerCharPr({}, ctx);
  const w = Metric.ptToHwp(img.w);
  const h = Metric.ptToHwp(img.h);

  return `<hp:run charPrIDRef="${charPrId}"><hp:pic id="0" zOrder="0" numberingType="PICTURE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" treatAsChar="1"><hp:sz width="${w}" widthRelTo="ABSOLUTE" height="${h}" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="0" right="0" top="0" bottom="0"/><hp:imgRect x="0" y="0" w="${w}" h="${h}"/><hp:imgClip left="0" right="0" top="0" bottom="0"/><hp:img binaryItemIDRef="${binId}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/></hp:pic></hp:run>`;
}

function encodeGrid(grid: GridNode, ctx: HwpxCtx): string {
  const colCount = grid.kids[0]?.kids.length ?? 0;
  const rowCount = grid.kids.length;

  // Calculate column widths from available page width
  const totalWidth = ctx.availableWidth;
  const colWidth = Math.round(totalWidth / (colCount || 1));

  // Table borderFillIDRef
  const tblBfId = grid.props.defaultStroke
    ? addBorderFill(ctx, grid.props.defaultStroke)
    : 2; // default table border

  // Rows
  let rowsXml = '';
  for (let ri = 0; ri < grid.kids.length; ri++) {
    const row = grid.kids[ri];
    let cellsXml = '';
    for (let ci = 0; ci < row.kids.length; ci++) {
      const cell = row.kids[ci];

      // Cell borderFill
      let cellBfId = tblBfId;
      if (cell.props.bg) {
        cellBfId = addBorderFill(ctx, grid.props.defaultStroke ?? DEFAULT_STROKE, cell.props.bg);
      }

      // Encode cell paragraphs
      const parasXml = cell.kids.map(p => encodePara(p, ctx)).join('');

      cellsXml += `<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="${cellBfId}"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="${cell.props.va === 'mid' ? 'CENTER' : cell.props.va === 'bot' ? 'BOTTOM' : 'TOP'}" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">${parasXml}</hp:subList><hp:cellAddr colAddr="${ci}" rowAddr="${ri}"/><hp:cellSpan colSpan="${cell.cs}" rowSpan="${cell.rs}"/><hp:cellSz width="${colWidth * cell.cs}" height="0"/><hp:cellMargin left="141" right="141" top="141" bottom="141"/></hp:tc>`;
    }
    rowsXml += `<hp:tr>${cellsXml}</hp:tr>`;
  }

  const headerRow = grid.props.headerRow ? ' repeatHeader="1"' : '';

  return `<hp:tbl id="0" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL"${headerRow} rowCnt="${rowCount}" colCnt="${colCount}" cellSpacing="0" borderFillIDRef="${tblBfId}" noAdjust="0"><hp:sz width="${totalWidth}" widthRelTo="ABSOLUTE" height="0" heightRelTo="ABSOLUTE" protect="0"/><hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/><hp:outMargin left="0" right="0" top="0" bottom="0"/><hp:inMargin left="138" right="138" top="138" bottom="138"/>${rowsXml}</hp:tbl>`;
}

function esc(s: string): string { return TextKit.escapeXml(s); }

registry.registerEncoder(new HwpxEncoder());
