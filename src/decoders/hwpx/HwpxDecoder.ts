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

      if (headXml) {
        try {
          const headStr = TextKit.decode(headXml);
          const headObj: any = await XmlKit.parseStrict(headStr);
          if (headObj) {
            meta = extractMeta(headObj);
            dims = extractDims(headObj) ?? dims;
          }
        } catch {
          // header parse failure is non-fatal
        }
      }

      const bodyStr = TextKit.decode(bodyXml);
      const bodyObj: any = await XmlKit.parseStrict(bodyStr);

      const sections = normalizeSections(bodyObj);
      const kids = shield.guardAll(
        sections,
        (sec: any) => decodeSection(sec, dims, shield),
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
  const root = bodyObj?.['hp:HWPML'] ?? bodyObj?.HWPML ?? bodyObj;
  const body = root?.['hp:BODY']?.[0] ?? root?.BODY?.[0] ?? root?.['hp:BODY'] ?? root?.BODY;
  if (!body) return [bodyObj];
  const sections = body?.['hp:SECTION'] ?? body?.SECTION ?? [];
  return Array.isArray(sections) ? sections : [sections];
}

function extractMeta(headObj: any): DocMeta {
  try {
    const root = headObj?.['hh:HEAD']?.[0] ?? headObj?.HEAD?.[0] ?? headObj;
    const info = root?.['hh:DOCSUMMARY']?.[0] ?? root?.DOCSUMMARY?.[0];
    if (!info) return {};
    const a = (k: string) => info?.[`hh:${k}`]?.[0]?._ ?? info?.[k]?.[0]?._ ?? '';
    return { title: a('TITLE') || undefined, author: a('AUTHOR') || undefined, subject: a('SUBJECT') || undefined };
  } catch { return {}; }
}

function extractDims(headObj: any): PageDims | null {
  try {
    const root = headObj?.['hh:HEAD']?.[0] ?? headObj?.HEAD?.[0] ?? headObj;
    const secList = root?.['hh:REFLIST']?.[0]?.['hh:SECPRLST']?.[0]?.['hh:SECPR']
      ?? root?.REFLIST?.[0]?.SECPRLST?.[0]?.SECPR;
    const sec = Array.isArray(secList) ? secList[0] : secList;
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

function decodeSection(sec: any, dims: PageDims, shield: ShieldedParser) {
  const paras  = toArr(sec?.['hp:P'] ?? sec?.P);
  const tables = toArr(sec?.['hp:TABLE'] ?? sec?.TABLE);

  const items: { type: string; node: any }[] = [
    ...paras.map((n: any) => ({ type: 'para', node: n })),
    ...tables.map((n: any) => ({ type: 'table', node: n })),
  ];

  const kids: ContentNode[] = shield.guardAll(
    items,
    (item: any) => {
      if (item.type === 'table') {
        const { value } = shield.guardGrid(
          item.node,
          (n) => decodeGrid(n, shield),
          (n) => decodeGridSimple(n, shield),
          (n) => decodeGridFlat(n),
          (n) => decodeGridText(n) as unknown as GridNode,
          'hwpx:table',
        );
        return value;
      }
      return decodePara(item.node, shield);
    },
    () => buildPara([buildSpan('[파싱 실패]')]),
    'hwpx:content',
  );

  return buildSheet(kids.filter(Boolean) as ContentNode[], dims);
}

function decodePara(p: any, shield: ShieldedParser): ParaNode {
  const runs = toArr(p?.['hp:RUN'] ?? p?.RUN);
  const paraAttr = p?.['hp:PARAPR']?.[0]?._attr ?? p?.PARAPR?.[0]?._attr ?? {};
  const props: ParaProps = { align: safeAlign(paraAttr.Align) };

  const kids: SpanNode[] = shield.guardAll(
    runs,
    (run: any) => decodeRun(run),
    () => buildSpan(''),
    'hwpx:run',
  );

  return buildPara(kids.filter(Boolean) as SpanNode[], props);
}

function decodeRun(run: any): SpanNode {
  const ca = run?.['hp:CHARPR']?.[0]?._attr ?? run?.CHARPR?.[0]?._attr ?? {};
  const textNodes = toArr(run?.['hp:T'] ?? run?.T ?? run?.['hp:CHAR'] ?? run?.CHAR);
  const content = textNodes.map((t: any) => typeof t === 'string' ? t : t?._ ?? t?._text ?? '').join('');

  const props: TextProps = {
    b:     ca.Bold === '1' || ca.Bold === 'true' || undefined,
    i:     ca.Italic === '1' || ca.Italic === 'true' || undefined,
    u:     ca.Underline ? ca.Underline !== 'NONE' : undefined,
    font:  safeFont(ca.FontName ?? ca.FaceNameHangul),
    pt:    ca.Height ? Metric.hHeightToPt(Number(ca.Height)) : undefined,
    color: safeHex(ca.TextColor),
  };

  return buildSpan(content, props);
}

function decodeGrid(tbl: any, shield: ShieldedParser): GridNode {
  const rowArr = toArr(tbl?.['hp:ROW'] ?? tbl?.ROW);
  const rowNodes = rowArr.map((row: any) => {
    const cellArr = toArr(row?.['hp:CELL'] ?? row?.CELL);
    const cellNodes = cellArr.map((cell: any) => {
      const ca = cell?._attr ?? {};
      const paras = toArr(cell?.['hp:P'] ?? cell?.P).map((p: any) => decodePara(p, shield));
      return buildCell(
        paras.length > 0 ? paras : [buildPara([buildSpan('')])],
        { cs: Number(ca.ColSpan ?? 1), rs: Number(ca.RowSpan ?? 1), props: { bg: safeHex(ca.BgColor) } as CellProps },
      );
    });
    return buildRow(cellNodes);
  });
  return buildGrid(rowNodes);
}

function decodeGridSimple(tbl: any, shield: ShieldedParser): GridNode {
  const rowArr = toArr(tbl?.['hp:ROW'] ?? tbl?.ROW);
  const rowNodes = rowArr.map((row: any) => {
    const cellArr = toArr(row?.['hp:CELL'] ?? row?.CELL);
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
  return toArr(cell?.['hp:P'] ?? cell?.P).map((p: any) =>
    toArr(p?.['hp:RUN'] ?? p?.RUN).map((r: any) =>
      toArr(r?.['hp:T'] ?? r?.T).map((t: any) => typeof t === 'string' ? t : t?._ ?? '').join(''),
    ).join(''),
  ).join(' ');
}

function tableText(tbl: any): string {
  return toArr(tbl?.['hp:ROW'] ?? tbl?.ROW).map((row: any) =>
    toArr(row?.['hp:CELL'] ?? row?.CELL).map((c: any) => cellText(c)).join('\t'),
  ).join('\n');
}

function toArr(v: any): any[] { return v == null ? [] : Array.isArray(v) ? v : [v]; }

// Auto-register
registry.registerDecoder(new HwpxDecoder());
