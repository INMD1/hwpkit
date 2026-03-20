import type { Encoder } from '../../contract/encoder';
import type { DocRoot, ParaNode, SpanNode, GridNode, ContentNode } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import type { PageDims } from '../../model/doc-props';
import { A4 } from '../../model/doc-props';
import { succeed, fail } from '../../contract/result';
import { Metric } from '../../safety/StyleBridge';
import { ArchiveKit } from '../../toolkit/ArchiveKit';
import { TextKit } from '../../toolkit/TextKit';
import { registry } from '../../pipeline/registry';

export class HwpxEncoder implements Encoder {
  readonly format = 'hwpx';

  async encode(doc: DocRoot): Promise<Outcome<Uint8Array>> {
    try {
      const sheet = doc.kids[0];
      const dims  = sheet?.dims ?? A4;

      const entries = [
        { name: 'mimetype',               data: TextKit.encode('application/hwp+zip') },
        { name: 'META-INF/container.xml',  data: TextKit.encode(CONTAINER_XML) },
        { name: 'Contents/content.hpf',    data: TextKit.encode(contentHpf()) },
        { name: 'Contents/header.xml',     data: TextKit.encode(headerXml(dims, doc.meta)) },
        { name: 'Contents/section0.xml',   data: TextKit.encode(sectionXml(sheet?.kids ?? [], dims)) },
      ];

      return succeed(await ArchiveKit.zip(entries));
    } catch (e: any) {
      return fail(`HWPX encode error: ${e?.message ?? String(e)}`);
    }
  }
}

const CONTAINER_XML = `<?xml version="1.0" encoding="UTF-8"?>
<container>
  <rootfiles>
    <rootfile full-path="Contents/content.hpf" media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>`;

function contentHpf(): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<opf:package xmlns:opf="http://www.idpf.org/2007/opf">
  <opf:manifest>
    <opf:item id="header" href="header.xml" media-type="application/xml"/>
    <opf:item id="section0" href="section0.xml" media-type="application/xml"/>
  </opf:manifest>
  <opf:spine><opf:itemref idref="section0"/></opf:spine>
</opf:package>`;
}

function headerXml(dims: PageDims, meta: any): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<hh:HEAD xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" version="1.0">
  <hh:DOCSUMMARY>
    <hh:TITLE>${esc(meta.title ?? '')}</hh:TITLE>
    <hh:AUTHOR>${esc(meta.author ?? '')}</hh:AUTHOR>
  </hh:DOCSUMMARY>
  <hh:REFLIST>
    <hh:SECPRLST>
      <hh:SECPR id="0" name="바탕쪽">
        <hh:PAGEPROPERTY Width="${Metric.ptToHwp(dims.wPt)}" Height="${Metric.ptToHwp(dims.hPt)}"
          TopMargin="${Metric.ptToHwp(dims.mt)}" BottomMargin="${Metric.ptToHwp(dims.mb)}"
          LeftMargin="${Metric.ptToHwp(dims.ml)}" RightMargin="${Metric.ptToHwp(dims.mr)}"
          Landscape="${dims.orient === 'landscape' ? 1 : 0}"/>
      </hh:SECPR>
    </hh:SECPRLST>
  </hh:REFLIST>
</hh:HEAD>`;
}

function sectionXml(kids: ContentNode[], _dims: PageDims): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<hp:SEC xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
${kids.map(encodeContent).join('\n')}
</hp:SEC>`;
}

function encodeContent(node: ContentNode): string {
  return node.tag === 'grid' ? encodeGrid(node) : encodePara(node);
}

function encodePara(para: ParaNode): string {
  const runs = para.kids.map(k => k.tag === 'span' ? encodeRun(k) : '').join('');
  return `<hp:P>
  <hp:PARAPR><hp:ALIGN Type="${(para.props.align ?? 'left').toUpperCase()}"/></hp:PARAPR>
  <hp:RUN>${runs}</hp:RUN>
</hp:P>`;
}

function encodeRun(span: SpanNode): string {
  const p = span.props;
  const content = span.kids.filter(k => k.tag === 'txt').map(k => k.tag === 'txt' ? esc(k.content) : '').join('');
  return `<hp:CHARPR Height="${Metric.ptToHHeight(p.pt ?? 10)}" Bold="${p.b ? '1' : '0'}" Italic="${p.i ? '1' : '0'}" Underline="${p.u ? 'BOTTOM' : 'NONE'}" TextColor="${p.color ?? '000000'}"/><hp:T>${content}</hp:T>`;
}

function encodeGrid(grid: GridNode): string {
  const rows = grid.kids.map(row => {
    const cells = row.kids.map(cell => {
      const paras = cell.kids.map(p => encodePara(p)).join('');
      return `<hp:CELL ColSpan="${cell.cs}" RowSpan="${cell.rs}">${paras}</hp:CELL>`;
    }).join('');
    return `<hp:ROW>${cells}</hp:ROW>`;
  }).join('');
  return `<hp:TABLE>${rows}</hp:TABLE>`;
}

function esc(s: string): string { return TextKit.escapeXml(s); }

registry.registerEncoder(new HwpxEncoder());
