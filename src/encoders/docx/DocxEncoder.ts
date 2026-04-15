import type { Encoder } from '../../contract/encoder';
import type { DocRoot, ParaNode, SpanNode, GridNode, ContentNode, ImgNode, SheetNode } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import type { PageDims, GridProps, CellProps } from '../../model/doc-props';
import { A4, normalizeDims } from '../../model/doc-props';
import { succeed, fail } from '../../contract/result';
import { Metric } from '../../safety/StyleBridge';
import { ArchiveKit } from '../../toolkit/ArchiveKit';
import { TextKit } from '../../toolkit/TextKit';
import { registry } from '../../pipeline/registry';

interface ImageEntry { rId: string; name: string; data: Uint8Array; ext: string }

export class DocxEncoder implements Encoder {
  readonly format = 'docx';

  async encode(doc: DocRoot): Promise<Outcome<Uint8Array>> {
    try {
      const sheet = doc.kids[0];
      const dims = normalizeDims(sheet?.dims ?? A4);
      const kids = sheet?.kids ?? [];

      const images: ImageEntry[] = [];
      const ctx: EncCtx = { images, nextId: 10, nextImgNum: 1, warns: [], imgMap: new WeakMap() };

      // Collect images from content
      collectImages(kids, ctx);

      // Header / footer
      let headerParas = sheet?.header;
      let footerParas = sheet?.footer;
      const hasHeader = headerParas && headerParas.length > 0;
      const hasFooter = footerParas && footerParas.length > 0;

      // Collect images from header/footer
      if (hasHeader) collectImagesFromParas(headerParas!, ctx);
      if (hasFooter) collectImagesFromParas(footerParas!, ctx);

      const headerRId = hasHeader ? `rId${ctx.nextId++}` : '';
      const footerRId = hasFooter ? `rId${ctx.nextId++}` : '';

      // Numbering: collect list info
      const numInfo = collectNumbering(kids);

      const entries: { name: string; data: Uint8Array }[] = [
        { name: '[Content_Types].xml', data: TextKit.encode(contentTypes(images, hasHeader, hasFooter)) },
        { name: '_rels/.rels', data: TextKit.encode(pkgRels()) },
        { name: 'word/document.xml', data: TextKit.encode(documentXml(kids, dims, ctx, headerRId, footerRId)) },
        { name: 'word/styles.xml', data: TextKit.encode(stylesXml()) },
        { name: 'word/settings.xml', data: TextKit.encode(settingsXml()) },
        { name: 'word/_rels/document.xml.rels', data: TextKit.encode(docRels(images, headerRId, footerRId, numInfo.hasLists)) },
        { name: 'docProps/app.xml', data: TextKit.encode(appXml()) },
        { name: 'docProps/core.xml', data: TextKit.encode(coreXml(doc.meta)) },
      ];

      // Add numbering.xml if needed
      if (numInfo.hasLists) {
        entries.push({ name: 'word/numbering.xml', data: TextKit.encode(numberingXml(numInfo)) });
      }

      // Add header/footer files
      if (hasHeader) {
        entries.push({ name: 'word/header1.xml', data: TextKit.encode(headerFooterXml('hdr', headerParas!, ctx)) });
      }
      if (hasFooter) {
        entries.push({ name: 'word/footer1.xml', data: TextKit.encode(headerFooterXml('ftr', footerParas!, ctx)) });
      }

      // Add image media files
      for (const img of images) {
        entries.push({ name: `word/media/${img.name}`, data: img.data });
      }

      return succeed(await ArchiveKit.zip(entries));
    } catch (e: any) {
      return fail(`DOCX encode error: ${e?.message ?? String(e)}`);
    }
  }
}

// ─── Context ────────────────────────────────────────────────

interface EncCtx {
  images: ImageEntry[];
  nextId: number;
  nextImgNum: number;
  warns: string[];
  imgMap: WeakMap<ImgNode, string>; // ImgNode → rId (no mutation)
}

// ─── Image collection ───────────────────────────────────────

function mimeToExt(mime: string): string {
  if (mime.includes('jpeg')) return 'jpeg';
  if (mime.includes('gif')) return 'gif';
  if (mime.includes('bmp')) return 'bmp';
  return 'png';
}

function collectImages(kids: ContentNode[], ctx: EncCtx): void {
  for (const kid of kids) {
    if (kid.tag === 'para') collectImagesFromPara(kid, ctx);
    else if (kid.tag === 'grid') {
      for (const row of kid.kids)
        for (const cell of row.kids)
          for (const p of cell.kids) collectImagesFromPara(p, ctx);
    }
  }
}

function collectImagesFromParas(paras: ParaNode[], ctx: EncCtx): void {
  for (const p of paras) collectImagesFromPara(p, ctx);
}

function collectImagesFromPara(para: ParaNode, ctx: EncCtx): void {
  for (const kid of para.kids) {
    if (kid.tag === 'img') registerImage(kid, ctx);
  }
}

function registerImage(img: ImgNode, ctx: EncCtx): void {
  if (ctx.imgMap.has(img)) return;
  const ext = mimeToExt(img.mime);
  const name = `image${ctx.nextImgNum++}.${ext}`;
  const rId = `rId${ctx.nextId++}`;
  const data = TextKit.base64Decode(img.b64);
  ctx.images.push({ rId, name, data, ext });
  ctx.imgMap.set(img, rId);
}

// ─── List/numbering collection ──────────────────────────────

interface NumInfo {
  hasLists: boolean;
  hasBullet: boolean;
  hasNumbered: boolean;
}

function collectNumbering(kids: ContentNode[]): NumInfo {
  let hasBullet = false;
  let hasNumbered = false;
  for (const kid of kids) {
    if (kid.tag === 'para') {
      if (kid.props.listOrd === true) hasNumbered = true;
      else if (kid.props.listOrd === false) hasBullet = true;
    }
  }
  return { hasLists: hasBullet || hasNumbered, hasBullet, hasNumbered };
}

// ─── OOXML boilerplate ──────────────────────────────────────

function contentTypes(images: ImageEntry[], hasHeader?: boolean, hasFooter?: boolean): string {
  const imgDefaults = new Set<string>();
  for (const img of images) imgDefaults.add(img.ext);

  let defaults = `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>`;

  for (const ext of imgDefaults) {
    const ct = ext === 'png' ? 'image/png' : ext === 'jpeg' ? 'image/jpeg' : ext === 'gif' ? 'image/gif' : 'image/bmp';
    defaults += `\n  <Default Extension="${ext}" ContentType="${ct}"/>`;
  }

  let overrides = `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`;

  if (hasHeader) overrides += `\n  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`;
  if (hasFooter) overrides += `\n  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  ${defaults}
  ${overrides}
</Types>`;
}

function pkgRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
}

function docRels(images: ImageEntry[], headerRId?: string, footerRId?: string, hasLists?: boolean): string {
  let rels = `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`;

  // Numbering relationship — only when lists exist
  if (hasLists) {
    rels += `\n  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`;
  }

  for (const img of images) {
    rels += `\n  <Relationship Id="${img.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${img.name}"/>`;
  }

  if (headerRId) {
    rels += `\n  <Relationship Id="${headerRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>`;
  }
  if (footerRId) {
    rels += `\n  <Relationship Id="${footerRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>`;
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${rels}
</Relationships>`;
}

function appXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>hwpkit</Application>
</Properties>`;
}

function coreXml(meta: any): string {
  const now = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${esc(meta.title ?? '')}</dc:title>
  <dc:creator>${esc(meta.author ?? '')}</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">${meta.created ?? now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`;
}

function stylesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="맑은 고딕" w:eastAsia="맑은 고딕" w:hAnsi="맑은 고딕"/>
      <w:sz w:val="20"/>
      <w:szCs w:val="20"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="0" w:line="384" w:lineRule="auto"/>
      <w:jc w:val="both"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>
  <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
</w:styles>`;
}

function settingsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:bordersDoNotSurroundHeader/>
  <w:bordersDoNotSurroundFooter/>
  <w:defaultTabStop w:val="800"/>
  <w:compat>
    <w:spaceForUL/>
    <w:balanceSingleByteDoubleByteWidth/>
    <w:doNotLeaveBackslashAlone/>
    <w:ulTrailSpace/>
    <w:doNotExpandShiftReturn/>
    <w:adjustLineHeightInTable/>
    <w:useFELayout/>
  </w:compat>
</w:settings>`;
}

// ─── numbering.xml ──────────────────────────────────────────

function numberingXml(info: NumInfo): string {
  let abstractNums = '';
  let nums = '';

  // Bullet list: abstractNumId=0, numId=1
  if (info.hasBullet) {
    abstractNums += `<w:abstractNum w:abstractNumId="0">`;
    for (let lvl = 0; lvl < 9; lvl++) {
      const marker = lvl === 0 ? '●' : lvl === 1 ? '○' : '■';
      const indent = (lvl + 1) * 720;
      abstractNums += `<w:lvl w:ilvl="${lvl}"><w:numFmt w:val="bullet"/><w:lvlText w:val="${marker}"/><w:pPr><w:ind w:left="${indent}" w:hanging="360"/></w:pPr></w:lvl>`;
    }
    abstractNums += `</w:abstractNum>`;
    nums += `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`;
  }

  // Numbered list: abstractNumId=1, numId=2
  if (info.hasNumbered) {
    abstractNums += `<w:abstractNum w:abstractNumId="1">`;
    for (let lvl = 0; lvl < 9; lvl++) {
      const fmt = lvl % 3 === 0 ? 'decimal' : lvl % 3 === 1 ? 'lowerLetter' : 'lowerRoman';
      const indent = (lvl + 1) * 720;
      abstractNums += `<w:lvl w:ilvl="${lvl}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="%${lvl + 1}."/><w:pPr><w:ind w:left="${indent}" w:hanging="360"/></w:pPr></w:lvl>`;
    }
    abstractNums += `</w:abstractNum>`;
    nums += `<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>`;
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  ${abstractNums}
  ${nums}
</w:numbering>`;
}

// ─── header / footer xml ────────────────────────────────────

function headerFooterXml(type: 'hdr' | 'ftr', paras: ParaNode[], ctx: EncCtx): string {
  const tag = type === 'hdr' ? 'w:hdr' : 'w:ftr';
  const body = paras.map(p => encodeParaInner(p, ctx)).join('\n');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<${tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
${body}
</${tag}>`;
}

// ─── document.xml ───────────────────────────────────────────

function documentXml(kids: ContentNode[], dims: PageDims, ctx: EncCtx, headerRId?: string, footerRId?: string): string {
  const body = kids.map(k => encodeContent(k, ctx, dims)).join('\n');

  let sectRefs = '';
  if (headerRId) sectRefs += `\n      <w:headerReference w:type="default" r:id="${headerRId}"/>`;
  if (footerRId) sectRefs += `\n      <w:footerReference w:type="default" r:id="${footerRId}"/>`;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
${body}
    <w:sectPr>${sectRefs}
      <w:pgSz w:w="${Metric.ptToDxa(dims.wPt)}" w:h="${Metric.ptToDxa(dims.hPt)}" w:orient="${dims.orient ?? 'portrait'}"/>
      <w:pgMar w:top="${Metric.ptToDxa(dims.mt)}" w:right="${Metric.ptToDxa(dims.mr)}" w:bottom="${Metric.ptToDxa(dims.mb)}" w:left="${Metric.ptToDxa(dims.ml)}" w:header="709" w:footer="709" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

function encodeContent(node: ContentNode, ctx: EncCtx, dims?: PageDims): string {
  return node.tag === 'grid' ? encodeGrid(node, ctx, dims) : encodeParaInner(node, ctx);
}

function encodeParaInner(para: ParaNode, ctx: EncCtx): string {
  const align = para.props.align ?? 'left';
  const headStyle = para.props.heading ? `<w:pStyle w:val="Heading${para.props.heading}"/>` : '';

  // List numbering
  let numPr = '';
  if (para.props.listOrd !== undefined) {
    const numId = para.props.listOrd ? 2 : 1;
    const ilvl = para.props.listLv ?? 0;
    numPr = `<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`;
  }

  // Spacing (before / after / line height)
  let spacingXml = '';
  const { spaceBefore, spaceAfter, lineHeight } = para.props;
  if (spaceBefore !== undefined || spaceAfter !== undefined || lineHeight !== undefined) {
    const parts: string[] = [];
    if (spaceBefore !== undefined) parts.push(`w:before="${Metric.ptToDxa(spaceBefore)}"`);
    if (spaceAfter !== undefined) parts.push(`w:after="${Metric.ptToDxa(spaceAfter)}"`);
    if (lineHeight !== undefined) parts.push(`w:line="${Math.round(lineHeight * 240)}" w:lineRule="auto"`);
    spacingXml = `<w:spacing ${parts.join(' ')}/>`;
  }

  // Indentation
  let indentXml = '';
  const leftDxa  = para.props.indentPt !== undefined ? Metric.ptToDxa(para.props.indentPt) : 0;
  const firstPt  = para.props.firstLineIndentPt;
  if (leftDxa > 0 || firstPt !== undefined) {
    const parts: string[] = [];
    if (leftDxa > 0) parts.push(`w:left="${leftDxa}"`);
    if (firstPt !== undefined) {
      const dxa = Metric.ptToDxa(Math.abs(firstPt));
      if (firstPt >= 0) parts.push(`w:firstLine="${dxa}"`);
      else              parts.push(`w:hanging="${dxa}"`);
    }
    if (parts.length > 0) indentXml = `<w:ind ${parts.join(' ')}/>`;
  }

  const runs = para.kids.map(k => {
    if (k.tag === 'span') return encodeRun(k, ctx);
    if (k.tag === 'img') return encodeImage(k, ctx);
    return '';
  }).join('');

  return `    <w:p>
      <w:pPr>${headStyle}${numPr}${spacingXml}${indentXml}<w:jc w:val="${align === 'justify' ? 'both' : align}"/></w:pPr>
      ${runs}
    </w:p>`;
}

function encodeRun(span: SpanNode, _ctx: EncCtx): string {
  const p = span.props;
  const rPr: string[] = [];
  if (p.b) rPr.push('<w:b/>');
  if (p.i) rPr.push('<w:i/>');
  if (p.u) rPr.push('<w:u w:val="single"/>');
  if (p.s) rPr.push('<w:strike/>');
  if (p.sup) rPr.push('<w:vertAlign w:val="superscript"/>');
  if (p.sub) rPr.push('<w:vertAlign w:val="subscript"/>');
  if (p.pt) rPr.push(`<w:sz w:val="${Metric.ptToHalfPt(p.pt)}"/><w:szCs w:val="${Metric.ptToHalfPt(p.pt)}"/>`);
  if (p.color) rPr.push(`<w:color w:val="${p.color}"/>`);
  if (p.font) rPr.push(`<w:rFonts w:ascii="${esc(p.font)}" w:hAnsi="${esc(p.font)}" w:eastAsia="${esc(p.font)}"/>`);
  if (p.bg) rPr.push(`<w:shd w:val="clear" w:color="auto" w:fill="${p.bg}"/>`);

  // Process kids — text and pagenum
  const parts: string[] = [];
  for (const kid of span.kids) {
    if (kid.tag === 'txt') {
      parts.push(`<w:r><w:rPr>${rPr.join('')}</w:rPr><w:t xml:space="preserve">${esc(kid.content)}</w:t></w:r>`);
    } else if (kid.tag === 'pagenum') {
      parts.push(`<w:r><w:rPr>${rPr.join('')}</w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr>${rPr.join('')}</w:rPr><w:instrText> PAGE </w:instrText></w:r><w:r><w:rPr>${rPr.join('')}</w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr>${rPr.join('')}</w:rPr><w:t>1</w:t></w:r><w:r><w:rPr>${rPr.join('')}</w:rPr><w:fldChar w:fldCharType="end"/></w:r>`);
    } else if (kid.tag === 'br') {
      parts.push(`<w:r><w:br/></w:r>`);
    } else if (kid.tag === 'pb') {
      parts.push(`<w:r><w:br w:type="page"/></w:r>`);
    }
  }

  return parts.join('');
}

function encodeImage(img: ImgNode, ctx: EncCtx): string {
  const rId = ctx.imgMap.get(img);
  if (!rId) return '';

  const cx = Metric.ptToEmu(img.w);
  const cy = Metric.ptToEmu(img.h);
  const alt = esc(img.alt ?? '');
  const docPrId = ctx.nextId++;

  const graphic = `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Image"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;

  const layout = img.layout;
  const isInline = !layout || layout.wrap === 'inline';

  if (isInline) {
    return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${docPrId}" name="Image" descr="${alt}"/>${graphic}</wp:inline></w:drawing></w:r>`;
  }

  return `<w:r><w:drawing>${encodeAnchor(img, cx, cy, alt, docPrId, graphic, layout)}</w:drawing></w:r>`;
}

function encodeAnchor(
  _img: ImgNode, cx: number, cy: number, alt: string, docPrId: number,
  graphic: string, layout: NonNullable<ImgNode['layout']>,
): string {
  const distT = Metric.ptToEmu(layout.distT ?? 0);
  const distB = Metric.ptToEmu(layout.distB ?? 0);
  const distL = Metric.ptToEmu(layout.distL ?? 9144);   // DOCX 기본 0.18pt
  const distR = Metric.ptToEmu(layout.distR ?? 9144);
  const behindDoc = (layout.behindDoc || layout.wrap === 'behind') ? '1' : '0';
  const relH = layout.zOrder ?? 251658240;

  // 가로 위치
  const horzRelFrom = HORZ_RELTO_DOCX[layout.horzRelTo ?? 'column'] ?? 'column';
  let posH: string;
  if (layout.xPt != null) {
    posH = `<wp:positionH relativeFrom="${horzRelFrom}"><wp:posOffset>${Metric.ptToEmu(layout.xPt)}</wp:posOffset></wp:positionH>`;
  } else {
    const ha = HORZ_ALIGN_DOCX[layout.horzAlign ?? 'left'] ?? 'left';
    posH = `<wp:positionH relativeFrom="${horzRelFrom}"><wp:align>${ha}</wp:align></wp:positionH>`;
  }

  // 세로 위치
  const vertRelFrom = VERT_RELTO_DOCX[layout.vertRelTo ?? 'para'] ?? 'paragraph';
  let posV: string;
  if (layout.yPt != null) {
    posV = `<wp:positionV relativeFrom="${vertRelFrom}"><wp:posOffset>${Metric.ptToEmu(layout.yPt)}</wp:posOffset></wp:positionV>`;
  } else {
    const va = VERT_ALIGN_DOCX[layout.vertAlign ?? 'top'] ?? 'top';
    posV = `<wp:positionV relativeFrom="${vertRelFrom}"><wp:align>${va}</wp:align></wp:positionV>`;
  }

  // 텍스트 감싸기
  const wrapXml = WRAP_DOCX[layout.wrap] ?? '<wp:wrapSquare wrapText="bothSides"/>';

  return `<wp:anchor distT="${distT}" distB="${distB}" distL="${distL}" distR="${distR}" simplePos="0" relativeHeight="${relH}" behindDoc="${behindDoc}" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>${posH}${posV}<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrapXml}<wp:docPr id="${docPrId}" name="Image" descr="${alt}"/>${graphic}</wp:anchor>`;
}

const HORZ_RELTO_DOCX: Record<string, string> = {
  margin: 'margin', column: 'column', page: 'page', para: 'paragraph',
};
const VERT_RELTO_DOCX: Record<string, string> = {
  margin: 'margin', line: 'line', page: 'page', para: 'paragraph',
};
const HORZ_ALIGN_DOCX: Record<string, string> = {
  left: 'left', center: 'center', right: 'right',
};
const VERT_ALIGN_DOCX: Record<string, string> = {
  top: 'top', center: 'center', bottom: 'bottom',
};
const WRAP_DOCX: Record<string, string> = {
  square: '<wp:wrapSquare wrapText="bothSides"/>',
  tight: '<wp:wrapTight><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>',
  through: '<wp:wrapThrough wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapThrough>',
  none: '<wp:wrapNone/>',
  behind: '<wp:wrapNone/>',
  front: '<wp:wrapNone/>',
};

function encodeGrid(grid: GridNode, ctx: EncCtx, dims?: PageDims): string {
  const gp = grid.props;
  const look = gp.look;

  // tblLook attributes
  const firstRow = look?.firstRow ? '1' : '0';
  const lastRow = look?.lastRow ? '1' : '0';
  const firstCol = look?.firstCol ? '1' : '0';
  const lastCol = look?.lastCol ? '1' : '0';
  const noHBand = look?.bandedRows ? '0' : '1';
  const noVBand = look?.bandedCols ? '0' : '1';

  const d = dims ?? A4;
  const availDxa = Metric.ptToDxa(d.wPt - d.ml - d.mr);

  // 1단계: 표의 가상 2D 맵핑 (Virtual Table Map) 생성
  // 'real': 데이터 셀, 'continue': 세로 병합 지속 셀, 'absorbed': 가로/세로 병합으로 흡수된 자리, 'void': 빈 공간
  interface CellMap {
    type: 'real' | 'continue' | 'absorbed' | 'void';
    cell?: any;
    width?: number;
  }
  const tableMap: CellMap[][] = Array.from({ length: grid.kids.length }, () => []);

  for (let ri = 0; ri < grid.kids.length; ri++) {
    let c = 0;
    for (const cell of grid.kids[ri].kids) {
      // 이미 이전 행의 rowspan이나 현재 행의 colspan으로 차지된 자리 건너뜀
      while (tableMap[ri][c]) c++;

      // 실제 데이터 셀 배치
      tableMap[ri][c] = { type: 'real', cell, width: cell.cs };

      // 병합 영역(colspan, rowspan) 예약 처리
      for (let rr = 0; rr < cell.rs; rr++) {
        const targetRi = ri + rr;
        if (targetRi >= grid.kids.length) break;
        if (!tableMap[targetRi]) tableMap[targetRi] = [];

        for (let cc = 0; cc < cell.cs; cc++) {
          if (rr === 0 && cc === 0) continue; // 시작 셀은 이미 'real'로 처리됨

          if (rr > 0 && cc === 0) {
            // 세로 병합이 시작된 이후 행의 첫 번째 칸
            tableMap[targetRi][c + cc] = { type: 'continue', width: cell.cs };
          } else {
            // 가로 병합으로 흡수된 칸 또는 세로 병합 중 가로 병합된 칸
            tableMap[targetRi][c + cc] = { type: 'absorbed' };
          }
        }
      }
      c += cell.cs;
    }
  }

  // 정확한 전체 열 개수(colCount) 계산 (모든 행 중 최대 길이)
  let colCount = 0;
  for (let ri = 0; ri < grid.kids.length; ri++) {
    colCount = Math.max(colCount, tableMap[ri].length);
  }
  if (colCount === 0) colCount = 1;

  // 빈 공간(void) 채우기 및 colCount에 맞춰 배열 길이 정규화
  for (let ri = 0; ri < grid.kids.length; ri++) {
    for (let c = 0; c < colCount; c++) {
      if (!tableMap[ri][c]) tableMap[ri][c] = { type: 'void' };
    }
  }

  // 2단계: 컬럼 너비(dxa) 계산
  const defaultColDxa = Math.round(availDxa / colCount);
  let colWidthsDxa: number[] = [];

  if (grid.props.colWidths && grid.props.colWidths.length > 0) {
    const srcPt = [...grid.props.colWidths];
    while (srcPt.length < colCount) srcPt.push(0);
    srcPt.length = colCount;

    const knownTotalPt = srcPt.filter(w => w > 0).reduce((s, w) => s + w, 0);
    const zeroCount = srcPt.filter(w => w <= 0).length;
    const availPt = Metric.dxaToPt(availDxa);

    const remainingPt = Math.max(0, availPt - knownTotalPt);
    const zeroFillPt = zeroCount > 0 ? remainingPt / zeroCount : 0;

    for (let i = 0; i < srcPt.length; i++) {
      if (srcPt[i] <= 0) {
        srcPt[i] = zeroFillPt > 0 ? zeroFillPt : (availPt / colCount);
      }
    }

    colWidthsDxa = srcPt.map(w => Math.round(Metric.ptToDxa(w)));
    const computedTotalDxa = colWidthsDxa.reduce((s, w) => s + w, 0);
    if (computedTotalDxa > availDxa) {
      const scale = availDxa / computedTotalDxa;
      colWidthsDxa = colWidthsDxa.map(w => Math.round(w * scale));
    }
  } else {
    for (let c = 0; c < colCount; c++) colWidthsDxa.push(defaultColDxa);
  }

  const totalDxa = colWidthsDxa.reduce((s, w) => s + w, 0);
  const gridCols = colWidthsDxa.map(w => `<w:gridCol w:w="${w}"/>`).join('');

  // 3단계: XML 렌더링
  const rows = grid.kids.map((row, ri) => {
    const cellXmls: string[] = [];

    for (let c = 0; c < colCount; c++) {
      const mapEntry = tableMap[ri][c];

      // 가로 병합으로 흡수된 칸은 렌더링하지 않음 (앞선 칸의 gridSpan이 차지)
      if (mapEntry.type === 'absorbed') continue;

      // 세로 병합 지속(continue), 실제 셀(real), 또는 빈 공간(void) 처리
      const isContinue = mapEntry.type === 'continue';
      const isReal = mapEntry.type === 'real';
      const isVoid = mapEntry.type === 'void';

      if (isContinue || isReal || isVoid) {
        let cw = 0;
        const cellWidth = mapEntry.width || 1;

        // 너비 계산: 현재 칸부터 colspan(width) 만큼의 컬럼 너비 합산
        for (let sc = c; sc < c + cellWidth && sc < colWidthsDxa.length; sc++) {
          cw += colWidthsDxa[sc];
        }
        if (cw === 0) cw = defaultColDxa * cellWidth;

        const tcPrParts: string[] = [];
        tcPrParts.push(`<w:tcW w:w="${Math.round(cw)}" w:type="dxa"/>`);

        if (cellWidth > 1) {
          tcPrParts.push(`<w:gridSpan w:val="${cellWidth}"/>`);
        }

        if (isContinue) {
          tcPrParts.push(`<w:vMerge/>`);
        }

        let cellContent = '';
        if (isReal) {
          const cell = mapEntry.cell!;
          const cp = cell.props;
          if (cell.rs > 1) tcPrParts.push(`<w:vMerge w:val="restart"/>`);

          const borders = encodeCellBorders(cp);
          if (borders) tcPrParts.push(borders);
          if (cp.bg) tcPrParts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${cp.bg}"/>`);
          if (cp.va) {
            const vaMap: Record<string, string> = { top: 'top', mid: 'center', bot: 'bottom' };
            tcPrParts.push(`<w:vAlign w:val="${vaMap[cp.va] ?? 'top'}"/>`);
          }
          cellContent = cell.kids.map((p: any) => encodeParaInner(p, ctx)).join('');
        } else {
          // continue 거나 void 인 경우 빈 단락 추가
          cellContent = `<w:p><w:pPr/></w:p>`;
        }

        const tcPr = `<w:tcPr>${tcPrParts.join('')}</w:tcPr>`;
        cellXmls.push(`      <w:tc>${tcPr}${cellContent}</w:tc>`);
      }
    }

    const trPrParts: string[] = [];
    if (ri === 0 && (gp.headerRow || look?.firstRow)) {
      trPrParts.push('<w:tblHeader/>');
    }
    if (row.heightPt != null && row.heightPt > 0) {
      const hDxa = Math.round(Metric.ptToDxa(row.heightPt));
      trPrParts.push(`<w:trHeight w:val="${hDxa}" w:hRule="exact"/>`);
    }
    const trPr = trPrParts.length > 0 ? `<w:trPr>${trPrParts.join('')}</w:trPr>` : '';

    return `    <w:tr>${trPr}\n${cellXmls.join('\n')}\n    </w:tr>`;
  }).join('\n');

  // 4단계: 테두리 및 최종 테이블 XML 조립
  let tblBorders = '';
  const strokeKindMap: Record<string, string> = {
    solid: 'single', dash: 'dash', dot: 'dot', double: 'double', none: 'none',
    dotDash: 'dotDash', dotDotDash: 'dotDotDash', triple: 'triple', thinThickSmallGap: 'thinThickSmallGap',
    thickThinSmallGap: 'thickThinSmallGap', thinThickThinSmallGap: 'thinThickThinSmallGap',
  };

  if (gp.defaultStroke) {
    const s = gp.defaultStroke;
    const val = strokeKindMap[s.kind] ?? 'single';

    if (val === 'none' || s.pt <= 0) {
      tblBorders = '<w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>';
    } else {
      // DOCX sz는 1/8pt 단위. 최소 굵기 2(0.25pt) 보장
      const sz = Math.max(2, Math.round(s.pt * 8));
      // 색상 '#' 제거 및 빈 값일 경우 auto 처리
      const clr = s.color ? s.color.replace('#', '') : 'auto';
      const bdr = `w:val="${val}" w:sz="${sz}" w:space="0" w:color="${clr}"`;
      tblBorders = `<w:tblBorders><w:top ${bdr}/><w:left ${bdr}/><w:bottom ${bdr}/><w:right ${bdr}/><w:insideH ${bdr}/><w:insideV ${bdr}/></w:tblBorders>`;
    }
  }

  return `    <w:tbl>
      <w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${Math.round(totalDxa)}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="${firstRow}" w:lastRow="${lastRow}" w:firstColumn="${firstCol}" w:lastColumn="${lastCol}" w:noHBand="${noHBand}" w:noVBand="${noVBand}"/>${tblBorders}<w:tblCellMar><w:top w:w="28" w:type="dxa"/><w:left w:w="102" w:type="dxa"/><w:bottom w:w="28" w:type="dxa"/><w:right w:w="102" w:type="dxa"/></w:tblCellMar></w:tblPr>
      <w:tblGrid>${gridCols}</w:tblGrid>
${rows}
    </w:tbl>`;
}
function encodeCellBorders(cp: CellProps): string {
  if (!cp.top && !cp.bot && !cp.left && !cp.right) return '';
  const strokeKindMap: Record<string, string> = {
    solid: 'single', dash: 'dash', dot: 'dot', double: 'double', none: 'none',
    dotDash: 'dotDash', dotDotDash: 'dotDotDash', triple: 'triple',
  };

  const encode = (s?: { kind: string; pt: number; color: string }, tag?: string) => {
    if (!s || !tag) return '';
    const val = strokeKindMap[s.kind] ?? 'single';

    // 선이 없거나 굵기가 0 이하인 경우 확실하게 제거 처리
    if (val === 'none' || s.pt <= 0) {
      return `<w:${tag} w:val="none" w:sz="0" w:space="0" w:color="auto"/>`;
    }

    // 최소 굵기 sz=2 (0.25pt) 보장
    const sz = Math.max(2, Math.round(s.pt * 8));
    // 색상 '#' 제거 및 빈 값일 경우 auto 처리
    const clr = s.color ? s.color.replace('#', '') : 'auto';

    return `<w:${tag} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${clr}"/>`;
  };

  return `<w:tcBorders>${encode(cp.top, 'top')}${encode(cp.bot, 'bottom')}${encode(cp.left, 'left')}${encode(cp.right, 'right')}</w:tcBorders>`;
}

function esc(s: string): string {
  if (!s) return '';
  // 1. 내부 처리용 플레이스홀더(__EXT_0__ 등) 제거
  s = s.replace(/__EXT_\d+__/g, '');

  // 2. 글자 깨짐을 유발하는 쓰레기값 및 BOM 기호 명시적 제거 (이 부분 추가!)
  s = s.replace(/湰灧/g, '');
  s = s.replace(/\uFEFF/g, '');

  // 3. DOCX(XML 1.0)에서 허용하지 않는 보이지 않는 제어문자 모두 제거
  s = s.replace(/[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD]/g, '');

  return TextKit.escapeXml(s);
}

registry.registerEncoder(new DocxEncoder());
