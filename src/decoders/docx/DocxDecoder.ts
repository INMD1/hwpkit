import type { Decoder } from '../../contract/decoder';
import type { DocRoot, ContentNode, ParaNode, SpanNode, GridNode } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import type { DocMeta, PageDims, TextProps, ParaProps, CellProps } from '../../model/doc-props';
import { A4 } from '../../model/doc-props';
import { succeed, fail } from '../../contract/result';
import { buildRoot, buildSheet, buildPara, buildSpan, buildGrid, buildRow, buildCell } from '../../model/builders';
import { ShieldedParser } from '../../safety/ShieldedParser';
import { Metric, safeAlign, safeFont, safeHex } from '../../safety/StyleBridge';
import { ArchiveKit } from '../../toolkit/ArchiveKit';
import { XmlKit } from '../../toolkit/XmlKit';
import { TextKit } from '../../toolkit/TextKit';
import { registry } from '../../pipeline/registry';

export class DocxDecoder implements Decoder {
  readonly format = 'docx';

  async decode(data: Uint8Array): Promise<Outcome<DocRoot>> {
    const shield = new ShieldedParser();
    const warns: string[] = [];

    try {
      const files = await ArchiveKit.unzip(data);

      const docXml = files.get('word/document.xml');
      if (!docXml) return fail('DOCX: word/document.xml not found');

      const relsXml = files.get('word/_rels/document.xml.rels');
      const relsMap = relsXml ? await parseRels(TextKit.decode(relsXml)) : new Map<string, string>();

      const coreXml = files.get('docProps/core.xml');
      let meta: DocMeta = {};
      if (coreXml) {
        try {
          meta = await parseCoreProps(TextKit.decode(coreXml));
        } catch {
          // ignore — meta is optional
        }
      }

      const docStr = TextKit.decode(docXml);
      const docObj: any = await XmlKit.parseStrict(docStr);

      const body = getBody(docObj);
      const dims = extractDims(body) ?? { ...A4 };
      const elements = getBodyElements(body);

      const kids: ContentNode[] = shield.guardAll(
        elements,
        (el: any) => decodeElement(el, relsMap, files, shield),
        () => buildPara([buildSpan('[요소 파싱 실패]')]),
        'docx:bodyElement',
      );

      warns.push(...shield.flush());
      const sheet = buildSheet(kids.filter(Boolean) as ContentNode[], dims);
      return succeed(buildRoot(meta, [sheet]), warns);
    } catch (e: any) {
      warns.push(...shield.flush());
      return fail(`DOCX decode error: ${e?.message ?? String(e)}`, warns);
    }
  }
}

// ─── helpers ────────────────────────────────────────────────

function toArr(v: any): any[] { return v == null ? [] : Array.isArray(v) ? v : [v]; }

async function parseRels(xml: string): Promise<Map<string, string>> {
  const map = new Map<string, string>();
  try {
    const obj: any = await XmlKit.parseStrict(xml);
    for (const rel of toArr(obj?.Relationships?.[0]?.Relationship)) {
      const a = rel?._attr ?? {};
      if (a.Id && a.Target) map.set(a.Id, a.Target);
    }
  } catch { /* ignore */ }
  return map;
}

async function parseCoreProps(xml: string): Promise<DocMeta> {
  try {
    const obj: any = await XmlKit.parseStrict(xml);
    const c = obj?.['cp:coreProperties']?.[0] ?? obj?.coreProperties?.[0] ?? {};
    return {
      title:    c?.['dc:title']?.[0]?._ ?? c?.['dc:title']?.[0] ?? undefined,
      author:   c?.['dc:creator']?.[0]?._ ?? c?.['dc:creator']?.[0] ?? undefined,
      subject:  c?.['dc:subject']?.[0]?._ ?? c?.['dc:subject']?.[0] ?? undefined,
      created:  c?.['dcterms:created']?.[0]?._ ?? undefined,
      modified: c?.['dcterms:modified']?.[0]?._ ?? undefined,
    };
  } catch { return {}; }
}

function getBody(obj: any): any {
  return obj?.['w:document']?.[0]?.['w:body']?.[0] ?? obj?.document?.[0]?.body?.[0] ?? obj;
}

function extractDims(body: any): PageDims | null {
  try {
    const sp = body?.['w:sectPr']?.[0] ?? body?.sectPr?.[0];
    if (!sp) return null;
    const sz  = sp?.['w:pgSz']?.[0]?._attr ?? sp?.pgSz?.[0]?._attr;
    const mar = sp?.['w:pgMar']?.[0]?._attr ?? sp?.pgMar?.[0]?._attr;
    if (!sz) return null;
    return {
      wPt:    Metric.dxaToPt(Number(sz['w:w'] ?? sz.w ?? 11906)),
      hPt:    Metric.dxaToPt(Number(sz['w:h'] ?? sz.h ?? 16838)),
      mt:     Metric.dxaToPt(Number(mar?.['w:top'] ?? mar?.top ?? 1440)),
      mb:     Metric.dxaToPt(Number(mar?.['w:bottom'] ?? mar?.bottom ?? 1440)),
      ml:     Metric.dxaToPt(Number(mar?.['w:left'] ?? mar?.left ?? 1800)),
      mr:     Metric.dxaToPt(Number(mar?.['w:right'] ?? mar?.right ?? 1800)),
      orient: (sz['w:orient'] ?? sz.orient) === 'landscape' ? 'landscape' : 'portrait',
    };
  } catch { return null; }
}

function getBodyElements(body: any): { type: string; node: any }[] {
  return [
    ...toArr(body?.['w:p'] ?? body?.p).map((n: any) => ({ type: 'para', node: n })),
    ...toArr(body?.['w:tbl'] ?? body?.tbl).map((n: any) => ({ type: 'table', node: n })),
  ];
}

function decodeElement(
  el: { type: string; node: any },
  relsMap: Map<string, string>,
  files: Map<string, Uint8Array>,
  shield: ShieldedParser,
): ContentNode {
  if (el.type === 'table') {
    const { value } = shield.guardGrid(
      el.node,
      (n) => decodeGrid(n as any, relsMap, files, shield),
      (n) => decodeGridSimple(n as any),
      (n) => decodeGridFlat(n as any),
      (n) => decodeGridText(n as any) as unknown as GridNode,
      'docx:table',
    );
    return value;
  }
  return decodePara(el.node, relsMap, files, shield);
}

function decodePara(
  p: any,
  relsMap: Map<string, string>,
  files: Map<string, Uint8Array>,
  shield: ShieldedParser,
): ParaNode {
  const pPr = p?.['w:pPr']?.[0] ?? {};
  const alignVal = pPr?.['w:jc']?.[0]?._attr?.['w:val'] ?? pPr?.['w:jc']?.[0]?._attr?.val;
  const headStyle = pPr?.['w:pStyle']?.[0]?._attr?.['w:val'] ?? pPr?.['w:pStyle']?.[0]?._attr?.val ?? '';

  const props: ParaProps = {
    align: safeAlign(alignVal),
    heading: parseHeading(headStyle),
  };

  const runs = toArr(p?.['w:r'] ?? p?.r);
  const kids: SpanNode[] = shield.guardAll(
    runs,
    (run: any) => decodeRun(run),
    () => buildSpan(''),
    'docx:run',
  );

  return buildPara(kids.filter(Boolean) as SpanNode[], props);
}

function decodeRun(run: any): SpanNode {
  const rPr = run?.['w:rPr']?.[0] ?? run?.rPr?.[0] ?? {};

  const szAttr = rPr?.['w:sz']?.[0]?._attr ?? rPr?.sz?.[0]?._attr ?? {};
  const szVal  = szAttr?.['w:val'] ?? szAttr?.val;

  const colorAttr = rPr?.['w:color']?.[0]?._attr ?? rPr?.color?.[0]?._attr ?? {};
  const colorVal  = colorAttr?.['w:val'] ?? colorAttr?.val;

  const fontAttr = rPr?.['w:rFonts']?.[0]?._attr ?? rPr?.rFonts?.[0]?._attr ?? {};
  const fontName = fontAttr?.['w:ascii'] ?? fontAttr?.ascii
    ?? fontAttr?.['w:hAnsi'] ?? fontAttr?.hAnsi
    ?? fontAttr?.['w:eastAsia'] ?? fontAttr?.eastAsia;

  const underVal = rPr?.['w:u']?.[0]?._attr?.['w:val'] ?? rPr?.['w:u']?.[0]?._attr?.val;

  const props: TextProps = {
    b:     (rPr?.['w:b']?.[0] != null || rPr?.b?.[0] != null) || undefined,
    i:     (rPr?.['w:i']?.[0] != null || rPr?.i?.[0] != null) || undefined,
    u:     underVal && underVal !== 'none' ? true : undefined,
    s:     (rPr?.['w:strike']?.[0] != null) || undefined,
    pt:    szVal ? Metric.halfPtToPt(Number(szVal)) : undefined,
    color: safeHex(colorVal),
    font:  fontName ? safeFont(fontName) : undefined,
  };

  const textNodes = toArr(run?.['w:t'] ?? run?.t);
  const content = textNodes.map((t: any) => typeof t === 'string' ? t : t?._ ?? t?._text ?? '').join('');

  return buildSpan(content, props);
}

function decodeGrid(
  tbl: any,
  relsMap: Map<string, string>,
  files: Map<string, Uint8Array>,
  shield: ShieldedParser,
): GridNode {
  const rowArr = toArr(tbl?.['w:tr'] ?? tbl?.tr);
  const rowNodes = rowArr.map((row: any) => {
    const cellArr = toArr(row?.['w:tc'] ?? row?.tc);
    const cellNodes = cellArr.map((cell: any) => {
      const tcPr = cell?.['w:tcPr']?.[0] ?? {};
      const gridSpan = Number(tcPr?.['w:gridSpan']?.[0]?._attr?.['w:val'] ?? 1);
      const bgAttr = tcPr?.['w:shd']?.[0]?._attr ?? {};
      const bg = safeHex(bgAttr?.['w:fill'] ?? bgAttr?.fill);

      const paras = toArr(cell?.['w:p'] ?? cell?.p).map((p: any) => decodePara(p, relsMap, files, shield));
      return buildCell(
        paras.length > 0 ? paras : [buildPara([buildSpan('')])],
        { cs: gridSpan, rs: 1, props: { bg } as CellProps },
      );
    });
    return buildRow(cellNodes);
  });
  return buildGrid(rowNodes);
}

function decodeGridSimple(tbl: any): GridNode {
  const rowArr = toArr(tbl?.['w:tr'] ?? tbl?.tr);
  const rowNodes = rowArr.map((row: any) => {
    const cellArr = toArr(row?.['w:tc'] ?? row?.tc);
    return buildRow(cellArr.map((c: any) => buildCell([buildPara([buildSpan(cellText(c))])])));
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
  return toArr(cell?.['w:p'] ?? cell?.p).map((p: any) =>
    toArr(p?.['w:r'] ?? p?.r).map((r: any) =>
      toArr(r?.['w:t'] ?? r?.t).map((t: any) => typeof t === 'string' ? t : t?._ ?? '').join(''),
    ).join(''),
  ).join(' ');
}

function tableText(tbl: any): string {
  return toArr(tbl?.['w:tr'] ?? tbl?.tr).map((row: any) =>
    toArr(row?.['w:tc'] ?? row?.tc).map((c: any) => cellText(c)).join('\t'),
  ).join('\n');
}

function parseHeading(style?: string): 1 | 2 | 3 | 4 | 5 | 6 | undefined {
  if (!style) return undefined;
  const m = style.match(/[Hh]eading(\d)/);
  if (m) { const n = Number(m[1]); if (n >= 1 && n <= 6) return n as any; }
  return undefined;
}

registry.registerDecoder(new DocxDecoder());
