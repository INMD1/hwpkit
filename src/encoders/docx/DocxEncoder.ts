import type {
  DocRoot,
  ParaNode,
  SpanNode,
  GridNode,
  ContentNode,
  ImgNode,
  SheetNode,
} from "../../model/doc-tree";
import type { Outcome } from "../../contract/result";
import type { PageDims, GridProps, CellProps, ImgLayout } from "../../model/doc-props";
import { A4, normalizeDims } from "../../model/doc-props";
import { succeed, fail } from "../../contract/result";
import { Metric } from "../../safety/StyleBridge";
import { ArchiveKit } from "../../toolkit/ArchiveKit";
import { TextKit } from "../../toolkit/TextKit";
import { registry } from "../../pipeline/registry";
import { BaseEncoder } from "../../core/BaseEncoder";
import { fitColumnWidths } from "../../toolkit/TableGeometry";
import rawFontMapping from "./font-mapping.json";

interface ImageEntry {
  rId: string;
  name: string;
  data: Uint8Array;
  ext: string;
}

type FontMappingEntry =
  | string
  | {
      nearest?: string;
      altName?: string;
      candidates?: Array<string | { name?: string; font?: string; distance?: number }>;
    };

const FONT_MAPPING = rawFontMapping as Record<string, FontMappingEntry>;

export class DocxEncoder extends BaseEncoder {
  protected getFormat(): string {
    return "docx";
  }

  async encode(doc: DocRoot): Promise<Outcome<Uint8Array>> {
    try {
      // 모든 섹션(SheetNode)을 처리 — 첫 번째 섹션에서 헤더/푸터/페이지 설정을 가져오고
      // 이후 섹션들의 콘텐츠는 sectPr로 구분하여 단일 body에 병합한다.
      const sheets = doc.kids.length > 0 ? doc.kids : [];
      const firstSheet = sheets[0];
      const dims = normalizeDims(firstSheet?.dims ?? A4);

      // 모든 섹션의 콘텐츠를 합산
      const allKids: ContentNode[] = sheets.flatMap((s) => s?.kids ?? []);

      const images: ImageEntry[] = [];
      const ctx: EncCtx = {
        images,
        dims,
        nextId: 10,
        nextImgNum: 1,
        warns: [],
        imgMap: new WeakMap(),
      };

      // Collect images from all content
      collectImages(allKids, ctx);

      // Header / footer (첫 번째 섹션 기준)
      const headerContents: ContentNode[] = [...(firstSheet?.headers?.default ?? [])];
      const footerContents: ContentNode[] = [...(firstSheet?.footers?.default ?? [])];
      const hasHeader = headerContents.length > 0;
      const hasFooter = footerContents.length > 0;

      // Collect images from header/footer
      if (hasHeader) collectImages(headerContents, ctx);
      if (hasFooter) collectImages(footerContents, ctx);

      const fonts = collectFonts(allKids);
      if (hasHeader) collectFonts(headerContents, fonts);
      if (hasFooter) collectFonts(footerContents, fonts);
      fonts.add("함초롬바탕");
      fonts.add("맑은 고딕");
      const hasFontTable = fonts.size > 0;

      const headerRId = hasHeader ? `rId${ctx.nextId++}` : "";
      const footerRId = hasFooter ? `rId${ctx.nextId++}` : "";

      // Numbering: collect list info from all sections
      const numInfo = collectNumbering(allKids);

      // kids 참조를 allKids로 통일 (후속 코드 호환)
      const kids = allKids;
      const mainDocumentXml = documentXml(kids, dims, ctx, headerRId, footerRId);
      const headerXml = hasHeader
        ? headerFooterXml("hdr", headerContents, ctx, dims)
        : "";
      const footerXml = hasFooter
        ? headerFooterXml("ftr", footerContents, ctx, dims)
        : "";

      const entries: { name: string; data: Uint8Array }[] = [
        {
          name: "[Content_Types].xml",
          data: this.stringToBytes(contentTypes(images, hasHeader, hasFooter, hasFontTable)),
        },
        { name: "_rels/.rels", data: this.stringToBytes(pkgRels()) },
        {
          name: "word/document.xml",
          data: this.stringToBytes(mainDocumentXml),
        },
        { name: "word/styles.xml", data: this.stringToBytes(stylesXml()) },
        { name: "word/settings.xml", data: this.stringToBytes(settingsXml()) },
        {
          name: "word/_rels/document.xml.rels",
          data: this.stringToBytes(
            docRels(images, headerRId, footerRId, numInfo.hasLists, hasFontTable),
          ),
        },
        { name: "docProps/app.xml", data: this.stringToBytes(appXml()) },
        {
          name: "docProps/core.xml",
          data: this.stringToBytes(coreXml(doc.meta)),
        },
      ];

      // Add numbering.xml if needed
      if (numInfo.hasLists) {
        entries.push({
          name: "word/numbering.xml",
          data: this.stringToBytes(numberingXml(numInfo)),
        });
      }

      if (hasFontTable) {
        entries.push({
          name: "word/fontTable.xml",
          data: this.stringToBytes(fontTableXml(fonts)),
        });
      }

      // Add header/footer files
      if (hasHeader) {
        entries.push({
          name: "word/header1.xml",
          data: this.stringToBytes(headerXml),
        });
        entries.push({
          name: "word/_rels/header1.xml.rels",
          data: this.stringToBytes(imagePartRels(images)),
        });
      }
      if (hasFooter) {
        entries.push({
          name: "word/footer1.xml",
          data: this.stringToBytes(footerXml),
        });
        entries.push({
          name: "word/_rels/footer1.xml.rels",
          data: this.stringToBytes(imagePartRels(images)),
        });
      }

      // Add image media files
      for (const img of images) {
        entries.push({ name: `word/media/${img.name}`, data: img.data });
      }

      return succeed(await this.zip(entries), ctx.warns);
    } catch (e: any) {
      return fail(`DOCX encode error: ${e?.message ?? String(e)}`);
    }
  }
}

// ─── Context ────────────────────────────────────────────────

interface EncCtx {
  images: ImageEntry[];
  dims: PageDims;
  nextId: number;
  nextImgNum: number;
  warns: string[];
  imgMap: WeakMap<ImgNode, string>; // ImgNode → rId (no mutation)
}

// ─── Image collection ───────────────────────────────────────

function mimeToExt(mime: string): string {
  if (mime.includes("jpeg")) return "jpeg";
  if (mime.includes("gif")) return "gif";
  if (mime.includes("bmp")) return "bmp";
  if (mime.includes("wmf")) return "wmf";
  if (mime.includes("emf")) return "emf";
  return "png";
}

function collectImages(kids: ContentNode[], ctx: EncCtx): void {
  for (const kid of kids) {
    if (kid.tag === "para") collectImagesFromPara(kid, ctx);
    else if (kid.tag === "grid") {
      for (const row of kid.kids)
        for (const cell of row.kids)
          for (const p of cell.kids)
            if (p.tag === "para") collectImagesFromPara(p, ctx);
            else collectImages([p], ctx);
    }
  }
}

function collectImagesFromParas(paras: ParaNode[], ctx: EncCtx): void {
  for (const p of paras) collectImagesFromPara(p, ctx);
}

function collectImagesFromPara(para: ParaNode, ctx: EncCtx): void {
  for (const kid of para.kids) {
    if (kid.tag === "img") registerImage(kid, ctx);
  }
}

function registerImage(img: ImgNode, ctx: EncCtx): void {
  if (ctx.imgMap.has(img)) return;
  const data = TextKit.base64Decode(img.b64);
  const ext = imageExtFromBytes(data) ?? mimeToExt(img.mime);
  const name = `image${ctx.nextImgNum++}.${ext}`;
  const rId = `rId${ctx.nextId++}`;
  ctx.images.push({ rId, name, data, ext });
  ctx.imgMap.set(img, rId);
}

function imageExtFromBytes(data: Uint8Array): string | undefined {
  if (data.length >= 8 && data[0] === 0x89 && data[1] === 0x50 && data[2] === 0x4e && data[3] === 0x47) return "png";
  if (data.length >= 3 && data[0] === 0xff && data[1] === 0xd8 && data[2] === 0xff) return "jpeg";
  if (data.length >= 6 && data[0] === 0x47 && data[1] === 0x49 && data[2] === 0x46) return "gif";
  if (data.length >= 2 && data[0] === 0x42 && data[1] === 0x4d) return "bmp";
  if (data.length >= 4 && data[0] === 0xd7 && data[1] === 0xcd && data[2] === 0xc6 && data[3] === 0x9a) return "wmf";
  if (data.length >= 44 && data[40] === 0x20 && data[41] === 0x45 && data[42] === 0x4d && data[43] === 0x46) return "emf";
  return undefined;
}

function registerSvgImage(svg: string, ctx: EncCtx): ImageEntry {
  const entry: ImageEntry = {
    rId: `rId${ctx.nextId++}`,
    name: `image${ctx.nextImgNum++}.svg`,
    data: TextKit.encode(svg),
    ext: "svg",
  };
  ctx.images.push(entry);
  return entry;
}

// ─── Font collection / fallback metadata ────────────────────

function collectFonts(kids: ContentNode[], fonts: Set<string> = new Set()): Set<string> {
  for (const kid of kids) {
    if (kid.tag === "para") collectFontsFromPara(kid, fonts);
    else if (kid.tag === "grid") {
      for (const row of kid.kids) {
        for (const cell of row.kids) {
          for (const child of cell.kids) {
            if (child.tag === "para") collectFontsFromPara(child, fonts);
            else collectFonts([child], fonts);
          }
        }
      }
    }
  }
  return fonts;
}

function collectFontsFromParas(paras: ParaNode[], fonts: Set<string>): void {
  for (const para of paras) collectFontsFromPara(para, fonts);
}

function collectFontsFromPara(para: ParaNode, fonts: Set<string>): void {
  for (const kid of para.kids) {
    if (kid.tag === "span") collectFontsFromSpan(kid, fonts);
    else if (kid.tag === "link") {
      for (const span of kid.kids) collectFontsFromSpan(span, fonts);
    } else if (kid.tag === "grid") {
      collectFonts([kid], fonts);
    }
  }
}

function collectFontsFromSpan(span: SpanNode, fonts: Set<string>): void {
  const font = span.props.font?.trim();
  if (font) fonts.add(font);
}

function mappedFontName(font: string): string | undefined {
  const entry = FONT_MAPPING[font] ?? FONT_MAPPING[font.trim()];
  if (!entry) return undefined;
  if (typeof entry === "string") return entry;
  if (entry.altName) return entry.altName;
  if (entry.nearest) return entry.nearest;
  const first = entry.candidates?.[0];
  if (typeof first === "string") return first;
  return first?.name ?? first?.font;
}

function fontTableXml(fonts: Set<string>): string {
  const body = Array.from(fonts)
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b, "ko"))
    .map((font) => {
      const alt = mappedFontName(font);
      const altXml = alt && alt !== font ? `<w:altName w:val="${esc(alt)}"/>` : "";
      return `<w:font w:name="${esc(font)}">${altXml}<w:family w:val="auto"/><w:pitch w:val="variable"/></w:font>`;
    })
    .join("\n  ");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  ${body}
</w:fonts>`;
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
    if (kid.tag === "para") {
      if (kid.props.listOrd === true) hasNumbered = true;
      else if (kid.props.listOrd === false) hasBullet = true;
    }
  }
  return { hasLists: hasBullet || hasNumbered, hasBullet, hasNumbered };
}

// ─── OOXML boilerplate ──────────────────────────────────────

function contentTypes(
  images: ImageEntry[],
  hasHeader?: boolean,
  hasFooter?: boolean,
  hasFontTable?: boolean,
): string {
  const imgDefaults = new Set<string>();
  for (const img of images) imgDefaults.add(img.ext);

  let defaults = `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>`;

  for (const ext of imgDefaults) {
    const ct =
      ext === "png"
        ? "image/png"
        : ext === "jpeg"
          ? "image/jpeg"
          : ext === "gif"
            ? "image/gif"
            : ext === "svg"
              ? "image/svg+xml"
              : ext === "wmf"
                ? "image/x-wmf"
                : ext === "emf"
                  ? "image/x-emf"
              : "image/bmp";
    defaults += `\n  <Default Extension="${ext}" ContentType="${ct}"/>`;
  }

  let overrides = `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`;

  if (hasHeader)
    overrides += `\n  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`;
  if (hasFooter)
    overrides += `\n  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`;
  if (hasFontTable)
    overrides += `\n  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>`;

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

function docRels(
  images: ImageEntry[],
  headerRId?: string,
  footerRId?: string,
  hasLists?: boolean,
  hasFontTable?: boolean,
): string {
  let rels = `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`;

  // Numbering relationship — only when lists exist
  if (hasLists) {
    rels += `\n  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`;
  }

  if (hasFontTable) {
    rels += `\n  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>`;
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

/** Relationships are scoped to each OOXML part, so header/footer drawings
 * cannot reuse image relationships declared only by document.xml. */
function imagePartRels(images: ImageEntry[]): string {
  const rels = images
    .map(
      (img) =>
        `<Relationship Id="${img.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${img.name}"/>`,
    )
    .join("\n  ");
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
  <dc:title>${esc(meta.title ?? "")}</dc:title>
  <dc:creator>${esc(meta.author ?? "")}</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">${meta.created ?? now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`;
}

function stylesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="함초롬바탕" w:eastAsia="함초롬바탕" w:hAnsi="함초롬바탕" w:hint="eastAsia"/>
      <w:sz w:val="20"/>
      <w:szCs w:val="20"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="0" w:line="384" w:lineRule="auto"/>
      <w:jc w:val="both"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="0"><w:name w:val="바탕글"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="1"><w:name w:val="본문"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="2"><w:name w:val="개요 1"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="3"><w:name w:val="개요 2"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="4"><w:name w:val="개요 3"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="5"><w:name w:val="개요 4"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="6"><w:name w:val="개요 5"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="7"><w:name w:val="개요 6"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="8"><w:name w:val="개요 7"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="9"><w:name w:val="개요 8"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="10"><w:name w:val="개요 9"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="11"><w:name w:val="개요 10"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="12"><w:name w:val="쪽 번호"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="13"><w:name w:val="머리말"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="14"><w:name w:val="각주"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="15"><w:name w:val="미주"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="16"><w:name w:val="메모"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="17"><w:name w:val="차례 제목"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="18"><w:name w:val="차례 1"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="19"><w:name w:val="차례 2"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="20"><w:name w:val="차례 3"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="21"><w:name w:val="본문 제목"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="22"><w:name w:val="그림"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="23"><w:name w:val="표"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="24"><w:name w:val="수식"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="25"><w:name w:val="인용문"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="26"><w:name w:val="날짜"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="27"><w:name w:val="발신명의"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="28"><w:name w:val="제목"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="29"><w:name w:val="부제목"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="30"><w:name w:val="문단 제목"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="31"><w:name w:val="MEMO"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="32"><w:name w:val="개요"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="33"><w:name w:val="표 제목"/><w:basedOn w:val="Normal"/></w:style>
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
  let abstractNums = "";
  let nums = "";

  // Bullet list: abstractNumId=0, numId=1
  if (info.hasBullet) {
    abstractNums += `<w:abstractNum w:abstractNumId="0">`;
    for (let lvl = 0; lvl < 9; lvl++) {
      const marker = lvl === 0 ? "●" : lvl === 1 ? "○" : "■";
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
      const fmt =
        lvl % 3 === 0
          ? "decimal"
          : lvl % 3 === 1
            ? "lowerLetter"
            : "lowerRoman";
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

function headerFooterXml(
  type: "hdr" | "ftr",
  contents: ContentNode[],
  ctx: EncCtx,
  dims: PageDims,
): string {
  const tag = type === "hdr" ? "w:hdr" : "w:ftr";
  const bodyParts = contents.map((node) => encodeContent(node, ctx, dims));
  if (contents[contents.length - 1]?.tag === "grid") {
    bodyParts.push('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>');
  }
  const body = bodyParts.join("\n");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<${tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
${body}
</${tag}>`;
}

// ─── document.xml ───────────────────────────────────────────

function documentXml(
  kids: ContentNode[],
  dims: PageDims,
  ctx: EncCtx,
  headerRId?: string,
  footerRId?: string,
): string {
  const body = kids.map((k) => encodeContent(k, ctx, dims)).join("\n");

  let sectRefs = "";
  if (headerRId)
    sectRefs += `\n      <w:headerReference w:type="default" r:id="${headerRId}"/>`;
  if (footerRId)
    sectRefs += `\n      <w:footerReference w:type="default" r:id="${footerRId}"/>`;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
${body}
    <w:sectPr>${sectRefs}
      <w:pgSz w:w="${Metric.ptToDxa(dims.wPt)}" w:h="${Metric.ptToDxa(dims.hPt)}"${dims.orient === "landscape" ? ' w:orient="landscape"' : ""}/>
      <w:pgMar w:top="${Metric.ptToDxa(dims.mt)}" w:right="${Metric.ptToDxa(dims.mr)}" w:bottom="${Metric.ptToDxa(dims.mb)}" w:left="${Metric.ptToDxa(dims.ml)}" w:header="${Metric.ptToDxa(dims.headerPt ?? 42.52)}" w:footer="${Metric.ptToDxa(dims.footerPt ?? 42.52)}" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

function encodeContent(
  node: ContentNode,
  ctx: EncCtx,
  dims?: PageDims,
): string {
  return node.tag === "grid"
    ? encodeGrid(node, ctx, dims)
    : encodeParaInner(node, ctx);
}

function encodeParaInner(
  para: ParaNode,
  ctx: EncCtx,
  maxWidthPt?: number,
): string {
  const align = para.props.align;
  // P3: hwpStyleId(숫자 ID) 우선, 없으면 heading 스타일, 둘 다 없으면 빈 문자열
  let headStyle = "";
  if (para.props.hwpStyleId !== undefined) {
    headStyle = `<w:pStyle w:val="${para.props.hwpStyleId}"/>`;
  } else if (para.props.heading) {
    headStyle = `<w:pStyle w:val="Heading${para.props.heading}"/>`;
  }

  // List numbering
  let numPr = "";
  if (para.props.listOrd !== undefined) {
    const numId = para.props.listOrd ? 2 : 1;
    const ilvl = para.props.listLv ?? 0;
    numPr = `<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`;
  }

  // Spacing (before / after / line height) - ensure all values are non-negative
  // ECMA-376 §17.3.1.33 spacing: line은 1/240th of a line(auto) 또는 dxa(exact/atLeast)
  let spacingXml = "";
  const { spaceBefore, spaceAfter, lineHeight, lineHeightFixed, lineHeightRule } = para.props;
  if (
    spaceBefore !== undefined ||
    spaceAfter !== undefined ||
    lineHeight !== undefined ||
    lineHeightFixed !== undefined
  ) {
    const parts: string[] = [];
    if (spaceBefore !== undefined)
      parts.push(`w:before="${Math.max(0, Metric.ptToDxa(spaceBefore))}"`);
    if (spaceAfter !== undefined)
      parts.push(`w:after="${Math.max(0, Metric.ptToDxa(spaceAfter))}"`);
    if (lineHeightFixed !== undefined) {
      parts.push(
        `w:line="${Math.max(1, Metric.ptToDxa(lineHeightFixed))}" w:lineRule="${lineHeightRule ?? "exact"}"`,
      );
    } else if (lineHeight !== undefined) {
      const ratio = docxLineHeightRatio(lineHeight);
      parts.push(
        `w:line="${Math.max(1, Math.floor(ratio * 240))}" w:lineRule="auto"`,
      );
    }
    spacingXml = `<w:spacing ${parts.join(" ")}/>`;
  }

  // Indentation — ECMA-376 §17.3.1.12 ind
  // w:left = 전체 왼쪽 여백(dxa), w:right = 전체 오른쪽 여백(dxa)
  // w:firstLine = 첫 줄 추가 들여쓰기(dxa, 양수), w:hanging = 내어쓰기(dxa, 양수값으로 표현)
  let indentXml = "";
  const indParts: string[] = [];
  // indentPt is the body margin. hanging locates the first line relative to
  // that margin; adding it to w:left again causes cumulative round-trip drift.
  if (para.props.indentPt !== undefined || (para.props.firstLineIndentPt ?? 0) < 0)
    indParts.push(`w:left="${Math.round(Metric.ptToDxa(para.props.indentPt ?? 0))}"`);
  if (para.props.indentRightPt !== undefined)
    indParts.push(`w:right="${Math.round(Metric.ptToDxa(para.props.indentRightPt))}"`);
  if (para.props.firstLineIndentPt !== undefined) {
    const first = Math.round(Metric.ptToDxa(para.props.firstLineIndentPt));
    indParts.push(first < 0 ? `w:hanging="${-first}"` : `w:firstLine="${first}"`);
  }
  if (indParts.length > 0) indentXml = `<w:ind ${indParts.join(" ")}/>`;

  const cjkLineBreakXml = "<w:kinsoku/><w:wordWrap/><w:overflowPunct/>";
  const omitEmptyLeftAlign = align === "left" && paraTextContent(para) === "";
  const jcXml = align && !omitEmptyLeftAlign ? `<w:jc w:val="${docxJcValue(align)}"/>` : "";

  const runs = para.kids
    .map((k) => {
      if (k.tag === "span") return encodeRun(k, ctx);
      if (k.tag === "img") return encodeImage(k, ctx, maxWidthPt);
      // P9: PageNumNode가 para.kids에 직접 있는 경우 (머리말/꼬리말 등)
      if (k.tag === "pagenum") {
        const instr = k.format === "total" ? " NUMPAGES " : " PAGE ";
        return `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>${instr}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
      }
      return "";
    })
    .join("");

  return `    <w:p>
      <w:pPr>${headStyle}${numPr}${spacingXml}${indentXml}${cjkLineBreakXml}${jcXml}</w:pPr>
      ${runs}
    </w:p>`;
}

function docxJcValue(align: NonNullable<ParaNode["props"]["align"]>): string {
  if (align === "justify") return "both";
  if (align === "distribute_space") return "distribute";
  return align;
}

function paraTextContent(para: ParaNode): string {
  let text = "";
  const collect = (kids: any[]) => {
    for (const kid of kids ?? []) {
      if (kid.tag === "txt") text += kid.content ?? "";
      else if (kid.kids) collect(kid.kids);
    }
  };
  collect(para.kids as any[]);
  return text;
}

function paraHasNonTextContent(para: ParaNode): boolean {
  let found = false;
  const visit = (kids: any[]) => {
    for (const kid of kids ?? []) {
      if (
        kid.tag === "img" ||
        kid.tag === "br" ||
        kid.tag === "pb" ||
        kid.tag === "pagenum"
      ) {
        found = true;
        return;
      }
      if (kid.kids) visit(kid.kids);
      if (found) return;
    }
  };
  visit(para.kids as any[]);
  return found;
}

function maxFontPtInPara(para: ParaNode): number {
  let maxPt = 10;
  const visit = (kids: any[]) => {
    for (const kid of kids ?? []) {
      if (kid.tag === "span" && typeof kid.props?.pt === "number") {
        maxPt = Math.max(maxPt, kid.props.pt);
      }
      if (kid.kids) visit(kid.kids);
    }
  };
  visit(para.kids as any[]);
  return maxPt;
}

function safeFixedLineHeightPt(para: ParaNode, fixedPt: number): number {
  const fontPt = maxFontPtInPara(para);
  return Math.max(fixedPt, fontPt * 1.15);
}

function minParaHeightPt(para: ParaNode): number {
  const fontPt = maxFontPtInPara(para);
  const before = Math.max(0, para.props.spaceBefore ?? 0);
  const after = Math.max(0, para.props.spaceAfter ?? 0);
  const lineCount = paraLineCount(para);
  let linePt: number;
  if (para.props.lineHeightFixed !== undefined) {
    linePt = safeFixedLineHeightPt(para, para.props.lineHeightFixed);
  } else {
    const ratio = docxLineHeightRatio(para.props.lineHeight ?? 1.15);
    linePt = fontPt * ratio;
  }
  return linePt * lineCount + before + after;
}

function docxLineHeightRatio(lineHeight: number): number {
  return Math.max(0.01, lineHeight);
}

function paraLineCount(para: ParaNode): number {
  let lines = 1;
  const visit = (kids: any[]) => {
    for (const kid of kids ?? []) {
      if (kid.tag === "span") {
        for (const child of kid.kids ?? []) {
          if (child.tag === "br") lines++;
          else if (child.tag === "txt") {
            lines += String(child.content ?? "").split(/\r\n|\r|\n/).length - 1;
          }
        }
      }
      if (kid.kids) visit(kid.kids);
    }
  };
  visit(para.kids as any[]);
  return Math.max(1, lines);
}

function minGridHeightPt(grid: GridNode): number {
  return (grid.kids ?? []).reduce((sum: number, row: any) => {
    const base = row.heightPt != null && row.heightPt > 0 ? row.heightPt : 14;
    let minRow = 0;
    for (const cell of row.kids ?? []) {
      const span = Math.max(1, cell.rs ?? 1);
      minRow = Math.max(minRow, minCellHeightPt(cell) / span);
    }
    return sum + Math.max(base, minRow);
  }, 0);
}

function minCellHeightPt(cell: any): number {
  const cp = cell.props ?? {};
  const padT = cp.padT ?? 1.4;
  const padB = cp.padB ?? 1.4;
  if (!cellHasVisibleContent(cell)) return padT + padB;
  let content = 0;
  for (const kid of cell.kids ?? []) {
    if (kid.tag === "para") content += minParaHeightPt(kid);
    else if (kid.tag === "grid") content += minGridHeightPt(kid);
  }
  return Math.max(content, 10) + padT + padB;
}

function cellHasVisibleContent(cell: any): boolean {
  for (const kid of cell.kids ?? []) {
    if (kid.tag === "grid") return true;
    if (kid.tag === "para") {
      if (paraTextContent(kid).trim() !== "") return true;
      if (paraHasNonTextContent(kid)) return true;
    }
  }
  return false;
}

function encodeRun(span: SpanNode, _ctx: EncCtx): string {
  const p = span.props;
  const rPr: string[] = [];
  if (p.b) rPr.push("<w:b/>");
  if (p.i) rPr.push("<w:i/>");
  if (p.u) rPr.push('<w:u w:val="single"/>');
  if (p.s) rPr.push("<w:strike/>");
  if (p.sup) rPr.push('<w:vertAlign w:val="superscript"/>');
  if (p.sub) rPr.push('<w:vertAlign w:val="subscript"/>');
  if (p.pt)
    rPr.push(
      `<w:sz w:val="${Metric.ptToHalfPt(p.pt)}"/><w:szCs w:val="${Metric.ptToHalfPt(p.pt)}"/>`,
    );
  if (p.color) rPr.push(`<w:color w:val="${p.color}"/>`);
  if (p.font)
    rPr.push(
      `<w:rFonts w:ascii="${esc(p.font)}" w:hAnsi="${esc(p.font)}" w:eastAsia="${esc(p.font)}" w:hint="eastAsia"/>`,
    );
  if (p.bg) rPr.push(`<w:shd w:val="clear" w:color="auto" w:fill="${p.bg}"/>`);

  // Process kids — text and pagenum
  const parts: string[] = [];
  for (const kid of span.kids) {
    if (kid.tag === "txt") {
      // __EXT_N__ or __EXT_N_W<w>_H<h>__ 자리표시자 제거
      const content = kid.content.replace(/__EXT_\d+(?:_W\d+_H\d+)?__/g, "");
      if (content || rPr.length > 0) {
        parts.push(
          `<w:r><w:rPr>${rPr.join("")}</w:rPr><w:t xml:space="preserve">${esc(content)}</w:t></w:r>`,
        );
      }
    } else if (kid.tag === "pagenum") {
      // P9: format === 'total' → NUMPAGES 필드, 나머지 → PAGE 필드
      const instr = kid.format === "total" ? " NUMPAGES " : " PAGE ";
      parts.push(
        `<w:r><w:rPr>${rPr.join("")}</w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr>${rPr.join("")}</w:rPr><w:instrText>${instr}</w:instrText></w:r><w:r><w:rPr>${rPr.join("")}</w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr>${rPr.join("")}</w:rPr><w:t>1</w:t></w:r><w:r><w:rPr>${rPr.join("")}</w:rPr><w:fldChar w:fldCharType="end"/></w:r>`,
      );
    } else if (kid.tag === "br") {
      parts.push(`<w:r><w:br/></w:r>`);
    } else if (kid.tag === "pb") {
      parts.push(`<w:r><w:br w:type="page"/></w:r>`);
    }
  }

  return parts.join("");
}

function encodeImage(img: ImgNode, ctx: EncCtx, maxWidthPt?: number): string {
  const rId = ctx.imgMap.get(img);
  if (!rId) return "";

  // Keep drawings inside the page/cell while preserving their aspect ratio.
  const bodyWidthPt = Math.max(1, ctx.dims.wPt - ctx.dims.ml - ctx.dims.mr);
  const bodyHeightPt = Math.max(1, ctx.dims.hPt - ctx.dims.mt - ctx.dims.mb);
  let widthPt = Number.isFinite(img.w) && img.w > 0 ? img.w : 72;
  let heightPt = Number.isFinite(img.h) && img.h > 0 ? img.h : 72;
  const widthLimit = Math.max(1, Math.min(bodyWidthPt, maxWidthPt ?? bodyWidthPt));
  const scale = Math.min(1, widthLimit / widthPt, bodyHeightPt / heightPt);
  widthPt *= scale;
  heightPt *= scale;
  const cx = Metric.ptToEmu(widthPt);
  const cy = Metric.ptToEmu(heightPt);
  const alt = esc(img.alt ?? "");
  const docPrId = ctx.nextId++;

  const graphic = `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Image"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;

  const layout = img.layout;

  const isInline = !layout || layout.wrap === "inline";
  // topAndBottom은 반드시 anchor (float)로 처리
  const forceAnchor =
    layout?.wrap === "topAndBottom" ||
    layout?.wrap === "square" ||
    layout?.wrap === "tight" ||
    layout?.wrap === "behind" ||
    layout?.wrap === "front";
  if (isInline && !forceAnchor) {
    return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${docPrId}" name="Image" descr="${alt}"/>${graphic}</wp:inline></w:drawing></w:r>`;
  }
  return `<w:r><w:drawing>${encodeAnchor(img, cx, cy, alt, docPrId, graphic, layout!)}</w:drawing></w:r>`;
}

function encodeAnchor(
  _img: ImgNode,
  cx: number,
  cy: number,
  alt: string,
  docPrId: number,
  graphic: string,
  layout: NonNullable<ImgNode["layout"]>,
): string {
  const distT = Metric.ptToEmu(layout.distT ?? 0);
  const distB = Metric.ptToEmu(layout.distB ?? 0);
  const distL = Metric.ptToEmu(layout.distL ?? 0);
  const distR = Metric.ptToEmu(layout.distR ?? 0);
  const behindDoc = layout.behindDoc || layout.wrap === "behind" ? "1" : "0";
  const relH = layout.zOrder ?? 251658240;

  // 가로 위치
  const horzRelFrom = HORZ_RELTO_DOCX[layout.horzRelTo ?? "column"] ?? "column";
  let posH: string;
  if (layout.xPt != null) {
    posH = `<wp:positionH relativeFrom="${horzRelFrom}"><wp:posOffset>${Metric.ptToEmu(layout.xPt)}</wp:posOffset></wp:positionH>`;
  } else {
    const ha = HORZ_ALIGN_DOCX[layout.horzAlign ?? "left"] ?? "left";
    posH = `<wp:positionH relativeFrom="${horzRelFrom}"><wp:align>${ha}</wp:align></wp:positionH>`;
  }

  // 세로 위치
  const vertRelFrom =
    VERT_RELTO_DOCX[layout.vertRelTo ?? "para"] ?? "paragraph";
  let posV: string;
  if (layout.yPt != null) {
    posV = `<wp:positionV relativeFrom="${vertRelFrom}"><wp:posOffset>${Metric.ptToEmu(layout.yPt)}</wp:posOffset></wp:positionV>`;
  } else {
    const va = VERT_ALIGN_DOCX[layout.vertAlign ?? "top"] ?? "top";
    posV = `<wp:positionV relativeFrom="${vertRelFrom}"><wp:align>${va}</wp:align></wp:positionV>`;
  }

  // 텍스트 감싸기
  const wrapXml =
    WRAP_DOCX[layout.wrap] ?? '<wp:wrapSquare wrapText="bothSides"/>';

  return `<wp:anchor distT="${distT}" distB="${distB}" distL="${distL}" distR="${distR}" simplePos="0" relativeHeight="${relH}" behindDoc="${behindDoc}" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>${posH}${posV}<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrapXml}<wp:docPr id="${docPrId}" name="Image" descr="${alt}"/>${graphic}</wp:anchor>`;
}

const HORZ_RELTO_DOCX: Record<string, string> = {
  margin: "margin",
  column: "column",
  page: "page",
  para: "paragraph",
};
const VERT_RELTO_DOCX: Record<string, string> = {
  margin: "margin",
  line: "line",
  page: "page",
  para: "paragraph",
};
const HORZ_ALIGN_DOCX: Record<string, string> = {
  left: "left",
  center: "center",
  right: "right",
};
const VERT_ALIGN_DOCX: Record<string, string> = {
  top: "top",
  center: "center",
  bottom: "bottom",
};
// ECMA-376 §20.4.2 wordprocessingDrawing wrapping elements
const WRAP_DOCX: Record<string, string> = {
  square: '<wp:wrapSquare wrapText="bothSides"/>',
  tight:
    '<wp:wrapTight><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>',
  through:
    '<wp:wrapThrough wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapThrough>',
  // ECMA-376 §20.4.2.15: wrapTopAndBottom — 텍스트가 이미지 위아래로만 흐름
  topAndBottom: "<wp:wrapTopAndBottom/>",
  none: "<wp:wrapNone/>",
  behind: "<wp:wrapNone/>",
  front: "<wp:wrapNone/>",
};

function shouldEncodeGridAsSvgFallback(grid: GridNode, dims: PageDims): boolean {
  const layout = grid.props.layout;
  if (!layout || layout.wrap === "inline") return false;
  if (layout.vertRelTo !== "page" || layout.vertAlign !== "bottom") return false;

  const widthPt = gridWidthPt(grid, dims);
  const bodyWidthPt = Math.max(1, dims.wPt - dims.ml - dims.mr);
  return grid.kids.length >= 4 || widthPt >= bodyWidthPt * 0.55;
}

function encodeGridAsSvgPicture(
  grid: GridNode,
  ctx: EncCtx,
  dims: PageDims,
): string {
  const widthPt = gridWidthPt(grid, dims);
  const heightPt = svgFallbackHeightPt(grid);
  const svg = renderGridSvg(grid, widthPt, heightPt);
  const entry = registerSvgImage(svg, ctx);
  const cx = Metric.ptToEmu(widthPt);
  const cy = Metric.ptToEmu(heightPt);
  const docPrId = ctx.nextId++;
  const layout = grid.props.layout!;
  const bottomGapPt = Math.min(44, Math.max(32, dims.mb * 0.9));
  const pictureLayout: ImgLayout = {
    ...layout,
    wrap: "none",
    horzRelTo: layout.horzRelTo === "para" ? "page" : (layout.horzRelTo ?? "page"),
    vertRelTo: "page",
    horzAlign: layout.horzAlign ?? "center",
    vertAlign: "bottom",
    yPt: layout.yPt ?? Math.max(0, dims.hPt - heightPt - bottomGapPt),
  };
  const graphic =
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="${escAttr(entry.name)}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${entry.rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic>`;

  return (
    `    <w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/></w:pPr>` +
    `<w:r><w:drawing>${encodeAnchor({} as ImgNode, cx, cy, "Positioned table", docPrId, graphic, pictureLayout)}</w:drawing></w:r>` +
    `</w:p>`
  );
}

function gridWidthPt(grid: GridNode, dims: PageDims): number {
  const widths = grid.props.colWidths ?? [];
  const total = widths.reduce((sum, w) => sum + Math.max(0, w), 0);
  if (total > 0) return total;
  return Math.max(1, dims.wPt - dims.ml - dims.mr);
}

function svgFallbackHeightPt(grid: GridNode): number {
  const rowSum = grid.kids.reduce((sum, row) => sum + Math.max(0, row.heightPt ?? 0), 0);
  const source = rowSum > 0 ? rowSum : minGridHeightPt(grid);
  return Math.min(220, Math.max(140, source * 0.88));
}

function renderGridSvg(grid: GridNode, widthPt: number, heightPt: number): string {
  const colCount = gridColumnCount(grid);
  const sourceCols = [...(grid.props.colWidths ?? [])];
  while (sourceCols.length < colCount) sourceCols.push(0);
  sourceCols.length = colCount;
  const known = sourceCols.filter((w) => w > 0).reduce((s, w) => s + w, 0);
  const unknown = sourceCols.filter((w) => w <= 0).length;
  const fill = unknown > 0 ? Math.max(1, (widthPt - known) / unknown) : 0;
  const colWidths = sourceCols.map((w) => (w > 0 ? w : fill));
  const colScale = widthPt / Math.max(1, colWidths.reduce((s, w) => s + w, 0));
  for (let i = 0; i < colWidths.length; i++) colWidths[i] *= colScale;

  const sourceRows = grid.kids.map((row) =>
    Math.max(row.heightPt ?? 0, row.kids.length > 0 ? 10 : 0),
  );
  const rowTotal = Math.max(1, sourceRows.reduce((s, h) => s + h, 0));
  const drawableHeightPt = Math.max(1, heightPt - 34);
  const rowHeights = sourceRows.map((h) => (h / rowTotal) * drawableHeightPt);

  const colX = cumulativePositions(colWidths);
  const rowY = cumulativePositions(rowHeights);
  const occupied: boolean[][] = Array.from({ length: grid.kids.length }, () => []);
  const defs: string[] = [];
  const body: string[] = [];
  let clipId = 0;

  for (let ri = 0; ri < grid.kids.length; ri++) {
    let ci = 0;
    for (const cell of grid.kids[ri].kids) {
      while (occupied[ri]?.[ci]) ci++;
      const cs = Math.max(1, cell.cs ?? 1);
      const rs = Math.max(1, cell.rs ?? 1);
      for (let rr = 0; rr < rs; rr++) {
        for (let cc = 0; cc < cs; cc++) {
          if (!occupied[ri + rr]) occupied[ri + rr] = [];
          occupied[ri + rr][ci + cc] = true;
        }
      }

      const x = colX[ci] ?? 0;
      const y = rowY[ri] ?? 0;
      const w = sumSlice(colWidths, ci, cs);
      const h = sumSlice(rowHeights, ri, rs);
      const cp = cell.props ?? {};
      if (cp.bg) {
        body.push(`<rect x="${fmt(x)}" y="${fmt(y)}" width="${fmt(w)}" height="${fmt(h)}" fill="#${escAttr(cp.bg)}"/>`);
      }
      drawSvgBorder(body, x, y, w, h, "top", cp.top ?? grid.props.defaultStroke);
      drawSvgBorder(body, x, y, w, h, "bottom", cp.bot ?? grid.props.defaultStroke);
      drawSvgBorder(body, x, y, w, h, "left", cp.left ?? grid.props.defaultStroke);
      drawSvgBorder(body, x, y, w, h, "right", cp.right ?? grid.props.defaultStroke);

      const lines = cellTextLines(cell);
      if (lines.length > 0 && w > 1 && h > 1) {
        const id = `c${clipId++}`;
        defs.push(`<clipPath id="${id}"><rect x="${fmt(x + 0.5)}" y="${fmt(y + 0.5)}" width="${fmt(Math.max(0, w - 1))}" height="${fmt(Math.max(0, h - 1))}"/></clipPath>`);
        body.push(renderSvgCellText(cell, lines, x, y, w, h, id));
      }
      ci += cs;
    }
  }

  const defsXml = defs.length ? `<defs>${defs.join("")}</defs>` : "";
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" width="${fmt(widthPt)}pt" height="${fmt(heightPt)}pt" viewBox="0 0 ${fmt(widthPt)} ${fmt(heightPt)}">` +
    defsXml +
    body.join("") +
    `</svg>`
  );
}

function gridColumnCount(grid: GridNode): number {
  let max = grid.props.colWidths?.length ?? 0;
  for (const row of grid.kids) {
    let count = 0;
    for (const cell of row.kids) count += Math.max(1, cell.cs ?? 1);
    max = Math.max(max, count);
  }
  return Math.max(1, max);
}

function cumulativePositions(values: number[]): number[] {
  const positions: number[] = [];
  let acc = 0;
  for (const value of values) {
    positions.push(acc);
    acc += value;
  }
  return positions;
}

function sumSlice(values: number[], start: number, count: number): number {
  let sum = 0;
  for (let i = start; i < start + count && i < values.length; i++) sum += values[i];
  return sum;
}

function drawSvgBorder(
  out: string[],
  x: number,
  y: number,
  w: number,
  h: number,
  side: "top" | "bottom" | "left" | "right",
  stroke?: { kind: string; pt: number; color: string },
): void {
  if (!stroke || stroke.kind === "none" || stroke.pt <= 0) return;
  const color = escAttr((stroke.color ?? "000000").replace(/^#/, ""));
  const sw = Math.max(0.4, stroke.pt);
  const dash =
    stroke.kind === "dash"
      ? ` stroke-dasharray="${fmt(sw * 6)} ${fmt(sw * 3)}"`
      : stroke.kind === "dot"
        ? ` stroke-dasharray="${fmt(sw)} ${fmt(sw * 3)}"`
        : "";
  const attrs = `stroke="#${color}" stroke-width="${fmt(sw)}"${dash} fill="none"`;
  if (side === "top") out.push(`<line x1="${fmt(x)}" y1="${fmt(y)}" x2="${fmt(x + w)}" y2="${fmt(y)}" ${attrs}/>`);
  else if (side === "bottom") out.push(`<line x1="${fmt(x)}" y1="${fmt(y + h)}" x2="${fmt(x + w)}" y2="${fmt(y + h)}" ${attrs}/>`);
  else if (side === "left") out.push(`<line x1="${fmt(x)}" y1="${fmt(y)}" x2="${fmt(x)}" y2="${fmt(y + h)}" ${attrs}/>`);
  else out.push(`<line x1="${fmt(x + w)}" y1="${fmt(y)}" x2="${fmt(x + w)}" y2="${fmt(y + h)}" ${attrs}/>`);
}

function renderSvgCellText(
  cell: any,
  lines: string[],
  x: number,
  y: number,
  w: number,
  h: number,
  clipId: string,
): string {
  const fontPt = Math.max(4.8, Math.min(7.2, maxFontPtInCell(cell) * 0.75));
  const step = fontPt * 1.18;
  const totalTextH = lines.length * step;
  const startY = y + Math.max(fontPt, (h - totalTextH) / 2 + fontPt * 0.85);
  const pad = Math.min(3, Math.max(1.2, w * 0.04));
  const align = firstParaAlign(cell);
  const anchor = align === "center" ? "middle" : align === "right" ? "end" : "start";
  const tx = align === "center" ? x + w / 2 : align === "right" ? x + w - pad : x + pad;
  const text = lines
    .map((line, i) =>
      `<text x="${fmt(tx)}" y="${fmt(startY + i * step)}" font-family="Malgun Gothic, Apple SD Gothic Neo, sans-serif" font-size="${fmt(fontPt)}" text-anchor="${anchor}" fill="#000000">${TextKit.escapeXml(line)}</text>`,
    )
    .join("");
  return `<g clip-path="url(#${clipId})">${text}</g>`;
}

function cellTextLines(cell: any): string[] {
  const lines: string[] = [];
  for (const kid of cell.kids ?? []) {
    if (kid.tag !== "para") continue;
    const text = paraPlainText(kid).trim();
    if (text) lines.push(...text.split(/\r\n|\r|\n/).map((s) => s.trim()).filter(Boolean));
  }
  return lines;
}

function paraPlainText(para: ParaNode): string {
  let text = "";
  const visit = (kids: any[]) => {
    for (const kid of kids ?? []) {
      if (kid.tag === "txt") text += kid.content ?? "";
      else if (kid.tag === "br") text += "\n";
      if (kid.kids) visit(kid.kids);
    }
  };
  visit(para.kids as any[]);
  return text;
}

function maxFontPtInCell(cell: any): number {
  let max = 9;
  const visit = (kids: any[]) => {
    for (const kid of kids ?? []) {
      if (kid.tag === "span" && typeof kid.props?.pt === "number") max = Math.max(max, kid.props.pt);
      if (kid.kids) visit(kid.kids);
    }
  };
  for (const kid of cell.kids ?? []) {
    if (kid.tag === "para") visit(kid.kids as any[]);
  }
  return max;
}

function firstParaAlign(cell: any): string {
  for (const kid of cell.kids ?? []) {
    if (kid.tag === "para" && kid.props?.align) return kid.props.align;
  }
  return cell.props?.align ?? "left";
}

function fmt(value: number): string {
  return Number.isFinite(value)
    ? value.toFixed(2).replace(/\.?0+$/, "")
    : "0";
}

function escAttr(value: string): string {
  return TextKit.escapeXml(value);
}

function encodeGrid(
  grid: GridNode,
  ctx: EncCtx,
  dims: PageDims = A4,
  maxWidthDxa?: number,
): string {
  const gp = grid.props;
  const look = gp.look;

  // tblLook attributes
  const firstRow = look?.firstRow ? "1" : "0";
  const lastRow = look?.lastRow ? "1" : "0";
  const firstCol = look?.firstCol ? "1" : "0";
  const lastCol = look?.lastCol ? "1" : "0";
  const noHBand = look?.bandedRows ? "0" : "1";
  const noVBand = look?.bandedCols ? "0" : "1";

  const d = dims ?? A4;
  const bodyDxa = Math.max(1, Metric.ptToDxa(d.wPt - d.ml - d.mr));
  const availDxa = Math.max(1, Math.min(bodyDxa, maxWidthDxa ?? bodyDxa));
  const tablePadT = Math.max(0, Math.round(Metric.ptToDxa(gp.cellPadT ?? 1.41)));
  const tablePadB = Math.max(0, Math.round(Metric.ptToDxa(gp.cellPadB ?? 1.41)));
  const tablePadL = Math.max(0, Math.round(Metric.ptToDxa(gp.cellPadL ?? 5.1)));
  const tablePadR = Math.max(0, Math.round(Metric.ptToDxa(gp.cellPadR ?? 5.1)));

  // 1단계: 표의 가상 2D 맵핑 (Virtual Table Map) 생성
  // 'real': 데이터 셀, 'continue': 세로 병합 지속 셀, 'absorbed': 가로/세로 병합으로 흡수된 자리, 'void': 빈 공간
  interface CellMap {
    type: "real" | "continue" | "absorbed" | "void";
    cell?: any;
    width?: number;
  }
  const tableMap: CellMap[][] = Array.from(
    { length: grid.kids.length },
    () => [],
  );

  for (let ri = 0; ri < grid.kids.length; ri++) {
    let c = 0;
    for (const cell of grid.kids[ri].kids) {
      // 이미 이전 행의 rowspan이나 현재 행의 colspan으로 차지된 자리 건너뜀
      while (tableMap[ri][c]) c++;

      // 실제 데이터 셀 배치
      tableMap[ri][c] = { type: "real", cell, width: cell.cs };

      // 병합 영역(colspan, rowspan) 예약 처리
      for (let rr = 0; rr < cell.rs; rr++) {
        const targetRi = ri + rr;
        if (targetRi >= grid.kids.length) break;
        if (!tableMap[targetRi]) tableMap[targetRi] = [];

        for (let cc = 0; cc < cell.cs; cc++) {
          if (rr === 0 && cc === 0) continue; // 시작 셀은 이미 'real'로 처리됨

          if (rr > 0 && cc === 0) {
            // 세로 병합이 시작된 이후 행의 첫 번째 칸
            tableMap[targetRi][c + cc] = { type: "continue", width: cell.cs };
          } else {
            // 가로 병합으로 흡수된 칸 또는 세로 병합 중 가로 병합된 칸
            tableMap[targetRi][c + cc] = { type: "absorbed" };
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
      if (!tableMap[ri][c]) tableMap[ri][c] = { type: "void" };
    }
  }

  // 2단계: 컬럼 너비(dxa) 계산
  const defaultColDxa = Math.max(1, Math.floor(availDxa / colCount));
  const sourceWidthsDxa = (grid.props.colWidths ?? []).map((width) =>
    width > 0 ? Metric.ptToDxa(width) : 0,
  );
  const colWidthsDxa = fitColumnWidths(
    sourceWidthsDxa,
    colCount,
    availDxa,
    Math.min(100, defaultColDxa),
  );

  const totalDxa = colWidthsDxa.reduce((s, w) => s + w, 0);
  const gridCols = colWidthsDxa.map((w) => `<w:gridCol w:w="${w}"/>`).join("");

  // 3단계: XML 렌더링
  const rows = tableMap
    .map((rowMap, ri) => {
      const cellXmls: string[] = [];

      for (let c = 0; c < colCount; c++) {
        const mapEntry = rowMap[c];

        // 가로 병합으로 흡수된 칸은 렌더링하지 않음 (앞선 칸의 gridSpan이 차지)
        if (mapEntry.type === "absorbed") continue;

        // 세로 병합 지속(continue), 실제 셀(real), 또는 빈 공간(void) 처리
        const isContinue = mapEntry.type === "continue";
        const isReal = mapEntry.type === "real";
        const isVoid = mapEntry.type === "void";

        if (isContinue || isReal || isVoid) {
          let cw = 0;
          const cellWidth = mapEntry.width || 1;

          // 너비 계산: 현재 칸부터 colspan(width) 만큼의 컬럼 너비 합산
          // colWidthsDxa 가 colCount 보다 작은 경우를 대비하여 safe 하게 처리
          const safeColWidths =
            colWidthsDxa.length >= colCount
              ? colWidthsDxa
              : [
                  ...colWidthsDxa,
                  ...Array(colCount - colWidthsDxa.length).fill(defaultColDxa),
                ];
          for (
            let sc = c;
            sc < c + cellWidth && sc < safeColWidths.length;
            sc++
          ) {
            cw += safeColWidths[sc];
          }
          if (cw <= 0) cw = defaultColDxa * cellWidth;

          const tcPrParts: string[] = [];
          tcPrParts.push(`<w:tcW w:w="${Math.round(cw)}" w:type="dxa"/>`);

          if (cellWidth > 1) {
            tcPrParts.push(`<w:gridSpan w:val="${cellWidth}"/>`);
          }

          if (isContinue) {
            tcPrParts.push(`<w:vMerge/>`);
          }

          let cellContent = "";
          if (isReal) {
            const cell = mapEntry.cell!;
            const cp = cell.props;
            if (cell.rs > 1) tcPrParts.push(`<w:vMerge w:val="restart"/>`);

            const borders = encodeCellBorders(cp);
            if (borders) tcPrParts.push(borders);
            if (cp.bg)
              tcPrParts.push(
                `<w:shd w:val="clear" w:color="auto" w:fill="${cp.bg}"/>`,
              );
            if (cp.va) {
              const vaMap: Record<string, string> = {
                top: "top",
                mid: "center",
                bot: "bottom",
              };
              tcPrParts.push(`<w:vAlign w:val="${vaMap[cp.va] ?? "top"}"/>`);
            }
            // Per-cell margin override (only when explicitly set)
            const cPadT =
              cp.padT != null ? Math.round(Metric.ptToDxa(cp.padT)) : null;
            const cPadB =
              cp.padB != null ? Math.round(Metric.ptToDxa(cp.padB)) : null;
            const cPadL =
              cp.padL != null ? Math.round(Metric.ptToDxa(cp.padL)) : null;
            const cPadR =
              cp.padR != null ? Math.round(Metric.ptToDxa(cp.padR)) : null;
            if (
              cPadT != null ||
              cPadB != null ||
              cPadL != null ||
              cPadR != null
            ) {
              const t = cPadT ?? tablePadT;
              const b = cPadB ?? tablePadB;
              const l = cPadL ?? tablePadL;
              const r = cPadR ?? tablePadR;
              tcPrParts.push(
                `<w:tcMar><w:top w:w="${t}" w:type="dxa"/><w:left w:w="${l}" w:type="dxa"/><w:bottom w:w="${b}" w:type="dxa"/><w:right w:w="${r}" w:type="dxa"/></w:tcMar>`,
              );
            }
            // Encode cell content: paragraphs and nested grids (중첩 표)
            // DOCX cells must end with <w:p>; nested <w:tbl> goes between paras
            const parts: string[] = [];
            for (const kid of cell.kids) {
              if (kid.tag === "grid") {
                const nestedLimit = Math.max(
                  1,
                  cw - (cPadL ?? tablePadL) - (cPadR ?? tablePadR),
                );
                parts.push(encodeGrid(kid, ctx, dims, nestedLimit));
              } else if (kid.tag === "para") {
                const textLimit = Math.max(
                  1,
                  cw - (cPadL ?? tablePadL) - (cPadR ?? tablePadR),
                );
                parts.push(encodeParaInner(kid, ctx, Metric.dxaToPt(textLimit)));
              }
            }
            // DOCX 셀은 반드시 <w:p>로 끝나야 함
            const lastKid = cell.kids[cell.kids.length - 1];
            if (cell.kids.length === 0 || lastKid?.tag === "grid") {
              parts.push('<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>');
            }
            cellContent = parts.join("");
          } else {
            // continue 거나 void 인 경우 빈 단락 추가
            cellContent = `<w:p><w:pPr/></w:p>`;
          }

          const tcPr = `<w:tcPr>${tcPrParts.join("")}</w:tcPr>`;
          cellXmls.push(`      <w:tc>${tcPr}${cellContent}</w:tc>`);
        }
      }

      const trPrParts: string[] = [];
      if (ri === 0 && (gp.headerRow || look?.firstRow)) {
        trPrParts.push("<w:tblHeader/>");
      }
      // 원본 행 정보에서 높이 가져오기 (tableMap 과 grid.kids 는 같은 인덱스)
      const originalRow = grid.kids[ri];
      if (originalRow?.heightPt != null && originalRow.heightPt > 0) {
        const hDxa = Math.round(Metric.ptToDxa(originalRow.heightPt));
        // Omitted hRule is Word's atLeast default and matches Hancom's DOCX
        // export while retaining the exact source minimum height.
        trPrParts.push(`<w:trHeight w:val="${hDxa}"/>`);
      }
      const trPr =
        trPrParts.length > 0 ? `<w:trPr>${trPrParts.join("")}</w:trPr>` : "";

      return `    <w:tr>${trPr}\n${cellXmls.join("\n")}\n    </w:tr>`;
    })
    .join("\n");

  // 4단계: 테두리 및 최종 테이블 XML 조립
  let tblBorders = "";
  const strokeKindMap: Record<string, string> = {
    solid: "single",
    dash: "dashed",
    dot: "dotted",
    double: "double",
    none: "none",
    dashDot: "dotDash",
    dashDotDot: "dotDotDash",
    dotDash: "dotDash",
    dotDotDash: "dotDotDash",
    triple: "triple",
    thinThickSmallGap: "thinThickSmallGap",
    thickThinSmallGap: "thickThinSmallGap",
    thinThickThinSmallGap: "thinThickThinSmallGap",
  };

  if (gp.defaultStroke) {
    const s = gp.defaultStroke;
    const val = strokeKindMap[s.kind] ?? "single";

    if (val === "none" || s.pt <= 0) {
      tblBorders =
        '<w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>';
    } else {
      // DOCX sz는 1/8pt 단위. 최소 굵기 2(0.25pt) 보장
      const minSz = val === "dashed" || val === "dotted" ? 4 : 2;
      const sz = Math.max(minSz, Math.round(s.pt * 8));
      // 색상 '#' 제거 및 빈 값일 경우 auto 처리
      const clr = s.color ? s.color.replace("#", "") : "auto";
      const bdr = `w:val="${val}" w:sz="${sz}" w:space="0" w:color="${clr}"`;
      tblBorders = `<w:tblBorders><w:top ${bdr}/><w:left ${bdr}/><w:bottom ${bdr}/><w:right ${bdr}/><w:insideH ${bdr}/><w:insideV ${bdr}/></w:tblBorders>`;
    }
  }

  // ECMA-376 §17.4.58: tblJc — 표 가로 정렬 (start/center/end)
  const tblAlignMap: Record<string, string> = {
    left: "start",
    center: "center",
    right: "end",
    justify: "start",
  };
  const tblJc = gp.align
    ? `<w:jc w:val="${tblAlignMap[gp.align] ?? "start"}"/>`
    : "";
  const tblPosition = encodeFloatingTablePr(gp.layout);

  return `    <w:tbl>
      <w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${Math.round(totalDxa)}" w:type="dxa"/>${tblPosition}<w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="${firstRow}" w:lastRow="${lastRow}" w:firstColumn="${firstCol}" w:lastColumn="${lastCol}" w:noHBand="${noHBand}" w:noVBand="${noVBand}"/>${tblBorders}${tblJc}<w:tblCellMar><w:top w:w="${tablePadT}" w:type="dxa"/><w:left w:w="${tablePadL}" w:type="dxa"/><w:bottom w:w="${tablePadB}" w:type="dxa"/><w:right w:w="${tablePadR}" w:type="dxa"/></w:tblCellMar></w:tblPr>
      <w:tblGrid>${gridCols}</w:tblGrid>
${rows}
    </w:tbl>`;
}

function encodeFloatingTablePr(layout?: ImgLayout): string {
  if (!layout || layout.wrap === "inline") return "";

  const horzAnchorMap: Record<string, string> = {
    margin: "margin",
    page: "page",
    column: "text",
    para: "text",
  };
  const vertAnchorMap: Record<string, string> = {
    margin: "margin",
    page: "page",
    line: "text",
    para: "text",
  };
  const alignMap: Record<string, string> = {
    left: "left",
    center: "center",
    right: "right",
    top: "top",
    bottom: "bottom",
  };

  const attrs: string[] = [
    `w:leftFromText="${Math.max(0, Metric.ptToDxa(layout.distL ?? 0))}"`,
    `w:rightFromText="${Math.max(0, Metric.ptToDxa(layout.distR ?? 0))}"`,
    `w:topFromText="${Math.max(0, Metric.ptToDxa(layout.distT ?? 0))}"`,
    `w:bottomFromText="${Math.max(0, Metric.ptToDxa(layout.distB ?? 0))}"`,
    `w:vertAnchor="${vertAnchorMap[layout.vertRelTo ?? "para"] ?? "text"}"`,
    `w:horzAnchor="${horzAnchorMap[layout.horzRelTo ?? "para"] ?? "text"}"`,
  ];

  if (layout.xPt != null) {
    attrs.push(`w:tblpX="${Metric.ptToDxa(layout.xPt)}"`);
  } else {
    attrs.push(`w:tblpXSpec="${alignMap[layout.horzAlign ?? "left"] ?? "left"}"`);
  }

  if (layout.yPt != null) {
    attrs.push(`w:tblpY="${Metric.ptToDxa(layout.yPt)}"`);
  } else {
    attrs.push(`w:tblpYSpec="${alignMap[layout.vertAlign ?? "top"] ?? "top"}"`);
  }

  return `<w:tblpPr ${attrs.join(" ")}/><w:tblOverlap w:val="overlap"/>`;
}

function encodeCellBorders(cp: CellProps): string {
  if (!cp.top && !cp.bot && !cp.left && !cp.right) return "";
  const strokeKindMap: Record<string, string> = {
    solid: "single",
    dash: "dashed",
    dot: "dotted",
    double: "double",
    none: "none",
    dashDot: "dotDash",
    dashDotDot: "dotDotDash",
    dotDash: "dotDash",
    dotDotDash: "dotDotDash",
    triple: "triple",
  };

  const encode = (
    s?: { kind: string; pt: number; color: string },
    tag?: string,
  ) => {
    if (!s || !tag) return "";
    const val = strokeKindMap[s.kind] ?? "single";

    // 선이 없거나 굵기가 0 이하인 경우 확실하게 제거 처리
    if (val === "none" || s.pt <= 0) {
      return `<w:${tag} w:val="none" w:sz="0" w:space="0" w:color="auto"/>`;
    }

    // 최소 굵기 보장. dash/dot은 LibreOffice PDF에서 0.25pt가 거의 사라져 보인다.
    const minSz = val === "dashed" || val === "dotted" ? 4 : 2;
    const sz = Math.max(minSz, Math.round(s.pt * 8));
    // 색상 '#' 제거 및 빈 값일 경우 auto 처리
    const clr = s.color ? s.color.replace("#", "") : "auto";

    return `<w:${tag} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${clr}"/>`;
  };

  return `<w:tcBorders>${encode(cp.top, "top")}${encode(cp.bot, "bottom")}${encode(cp.left, "left")}${encode(cp.right, "right")}</w:tcBorders>`;
}

function esc(s: string): string {
  if (!s) return "";
  // 1. 내부 처리용 플레이스홀더(__EXT_0__ 또는 __EXT_0_W144_H108__ 등) 제거
  s = s.replace(/__EXT_\d+(?:_W\d+_H\d+)?__/g, "");

  // 2. 글자 깨짐을 유발하는 쓰레기값 및 BOM 기호 명시적 제거 (이 부분 추가!)
  s = s.replace(/湰灧/g, "");
  s = s.replace(/\uFEFF/g, "");

  // 3. DOCX(XML 1.0)에서 허용하지 않는 보이지 않는 제어문자 모두 제거.
  // JS 문자열의 보조 평면 문자는 surrogate pair라서 정규식 range로 필터링하면
  // 한컴 특수문자(U+F080F 등)가 삭제된다. code point 단위로 판정한다.
  let xmlSafe = "";
  for (const ch of s) {
    const cp = ch.codePointAt(0);
    if (
      cp === 0x09 ||
      cp === 0x0a ||
      cp === 0x0d ||
      (cp !== undefined && cp >= 0x20 && cp <= 0xd7ff) ||
      (cp !== undefined && cp >= 0xe000 && cp <= 0xfffd) ||
      (cp !== undefined && cp >= 0x10000 && cp <= 0x10ffff)
    ) {
      xmlSafe += ch;
    }
  }

  return TextKit.escapeXml(xmlSafe);
}

registry.registerEncoder(new DocxEncoder());
