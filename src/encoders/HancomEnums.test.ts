import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import pako from 'pako';
import { SaxesParser } from 'saxes';
import { describe, expect, it } from 'vitest';
import type { DocRoot, GridNode, ImgNode, PageNumNode } from '../model/doc-tree';
import type { ImgLayout, ImgWrap, Stroke, StrokeKind } from '../model/doc-props';
import {
  buildCell, buildGrid, buildImg, buildPageNum, buildPara, buildRoot, buildRow, buildSheet, buildSpan,
} from '../model/builders';
import { Pipeline, registry } from '../index';
import { safeStrokeHwpx } from '../safety/StyleBridge';
import { BinaryKit } from '../toolkit/BinaryKit';
import { HANCOM_LINE_TYPES, HANCOM_LINE_WIDTHS, strokeKindToHwpxLineType } from '../toolkit/HancomEnums';
import { parseHwpRecords } from './hwp/verify';

// 한컴 OWPML 열거형 사전(속성 이름 → 허용 이름).
const DICT: Record<string, string[]> = JSON.parse(
  readFileSync('tests/fixtures/hancom-owpml-enums.json', 'utf8'),
).attributes;
const PNG_1PX =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const KINDS: StrokeKind[] = ['solid', 'dash', 'dot', 'double', 'dashDot', 'dashDotDot', 'wave', 'none'];
const WRAPS: ImgWrap[] = ['square', 'tight', 'through', 'topAndBottom', 'behind', 'front', 'none'];
const PAGE_FORMATS: NonNullable<PageNumNode['format']>[] = ['decimal', 'roman', 'romanCaps', 'total'];
// 한컴에는 tight/through가 없어 어울림으로, 그림의 wrapNone은 글 앞으로 간다.
const HANCOM_WRAP: Record<ImgWrap, ImgWrap> = {
  inline: 'inline', square: 'square', tight: 'square', through: 'square',
  topAndBottom: 'topAndBottom', behind: 'behind', front: 'front', none: 'front',
};

function sampleDoc(): DocRoot {
  const cell = (kind: StrokeKind, pt: number) => {
    const s: Stroke = { kind, pt, color: '000000' };
    return buildCell([buildPara([buildSpan(kind)])], { props: { top: s, bot: s, left: s, right: s } });
  };
  const layout = (wrap: ImgWrap, i: number): ImgLayout => ({
    wrap,
    horzRelTo: (['para', 'margin', 'page', 'column'] as const)[i % 4],
    vertRelTo: (['para', 'margin', 'page', 'line'] as const)[i % 4],
    horzAlign: 'center',
    vertAlign: 'top',
    xPt: 10 + i,
    yPt: 20 + i,
  });
  return buildRoot({}, [buildSheet([
    buildGrid([buildRow(KINDS.map((k, i) => cell(k, [0.5, 2.83, 8.5, 1][i % 4])))]),
    ...WRAPS.map((wrap, i) => buildPara([buildImg(PNG_1PX, 'image/png', 20, 20, `img-${wrap}`, layout(wrap, i))])),
    ...WRAPS.map((wrap, i) => buildGrid([buildRow([cell('solid', 0.5)])], { layout: layout(wrap, i) })),
    buildPara([{ tag: 'span', props: {}, kids: PAGE_FORMATS.map(f => buildPageNum(f)) }]),
  ])]);
}

async function encode(doc: DocRoot, format: 'hwpx' | 'hwp'): Promise<Uint8Array> {
  const result = await registry.getEncoder(format)!.encode(doc);
  if (!result.ok) throw new Error(result.error);
  return result.data;
}

function collect(node: unknown, out: { strokes: StrokeKind[]; imgWraps: ImgWrap[]; gridWraps: ImgWrap[]; pageNums: string[] }) {
  const n = node as { tag?: string; kids?: unknown[]; props?: Record<string, unknown>; layout?: ImgLayout; format?: string };
  if (n.tag === 'cell') out.strokes.push(((n.props as { top?: Stroke }).top ?? { kind: 'none' }).kind as StrokeKind);
  if (n.tag === 'img') out.imgWraps.push((n as ImgNode).layout?.wrap ?? 'inline');
  if (n.tag === 'grid') out.gridWraps.push((n as GridNode).props.layout?.wrap ?? 'inline');
  if (n.tag === 'pagenum') out.pageNums.push(n.format ?? 'decimal');
  for (const kid of n.kids ?? []) collect(kid, out);
}

async function roundTrip(format: 'hwpx' | 'hwp') {
  const parsed = await Pipeline.open(await encode(sampleDoc(), format), format).inspect();
  if (!parsed.ok) throw new Error(parsed.error);
  const out = { strokes: [] as StrokeKind[], imgWraps: [] as ImgWrap[], gridWraps: [] as ImgWrap[], pageNums: [] as string[] };
  collect(parsed.data, out);
  return out;
}

describe('Hancom OWPML enum names', () => {
  it('HWPX output uses only names present in the Hancom tables', async () => {
    const zip = await JSZip.loadAsync(await encode(sampleDoc(), 'hwpx'));
    const violations: string[] = [];
    let checked = 0;
    for (const entry of Object.values(zip.files)) {
      if (!/^Contents\/.*\.xml$/.test(entry.name)) continue;
      const parser = new SaxesParser({ xmlns: true });
      parser.on('opentag', tag => {
        for (const attr of Object.values(tag.attributes)) {
          const allowed = DICT[attr.local];
          if (!allowed || attr.local === 'value' || /^-?\d+$/.test(attr.value)) continue;
          checked++;
          if (!allowed.includes(attr.value)) violations.push(`<${tag.name} ${attr.local}="${attr.value}">`);
        }
      });
      parser.write(await entry.async('string')).close();
    }
    expect(checked).toBeGreaterThan(50);
    expect(violations).toEqual([]);
  });

  it('keeps line kinds, wraps and page-number formats through HWPX', async () => {
    const out = await roundTrip('hwpx');
    expect(out.strokes.slice(0, KINDS.length)).toEqual(KINDS);
    expect(out.imgWraps).toEqual(WRAPS.map(w => HANCOM_WRAP[w]));
    expect(out.pageNums).toEqual(PAGE_FORMATS);
  });

  it('keeps line kinds and image wraps through HWP 5', async () => {
    const out = await roundTrip('hwp');
    expect(out.strokes.slice(0, KINDS.length)).toEqual(KINDS);
    expect(out.imgWraps).toEqual(WRAPS.map(w => HANCOM_WRAP[w]));
  });

  it('writes HWP 5 BORDER_FILL and object wrap codes like Hancom-saved files', async () => {
    const streams = BinaryKit.parseCfb(await encode(sampleDoc(), 'hwp'));
    const inflate = (name: string) => new Uint8Array(pako.inflateRaw(streams.get(name)!));
    const types = new Set(parseHwpRecords(inflate('DocInfo')).filter(r => r.tag === 20).map(r => r.data[2]));
    // 2=DOT(긴 점선), 3=DASH(점선), 4/5=DASH_DOT(_DOT), 8=DOUBLE_SLIM, 12=WAVE.
    for (const code of [1, 2, 3, 4, 5, 8, 12]) expect(types, HANCOM_LINE_TYPES[code]).toContain(code);
    expect(types).not.toContain(7); // CIRCLE: 이전 이중선 오매핑

    const wraps = parseHwpRecords(inflate('BodyText/Section0'))
      .filter(r => r.tag === 71 && BinaryKit.readU32LE(r.data, 0) === 0x67736f20)
      .map(r => (BinaryKit.readU32LE(r.data, 4) >>> 21) & 7);
    expect(wraps).toEqual([0, 0, 0, 1, 2, 3, 3]);
  });

  it('reads Hancom-only line names and legacy hwpkit names', () => {
    expect(HANCOM_LINE_TYPES.map(t => safeStrokeHwpx(t).kind)).toEqual([
      'none', 'solid', 'dash', 'dot', 'dashDot', 'dashDotDot', 'dash', 'dot',
      'double', 'double', 'double', 'double', 'wave', 'wave', 'solid', 'solid', 'solid', 'solid',
    ]);
    expect(safeStrokeHwpx('DOUBLE').kind).toBe('double');
    // "DOT"(2)은 긴 점선, "DASH"(3)는 점선 모양이다.
    expect(strokeKindToHwpxLineType('dash')).toBe('DOT');
    expect(strokeKindToHwpxLineType('dot')).toBe('DASH');
    expect(HANCOM_LINE_WIDTHS).toContain('1.0 mm');
  });

  it('reads Hancom PAPER references and IN_FRONT_OF_TEXT pictures', async () => {
    const zip = await JSZip.loadAsync(await encode(sampleDoc(), 'hwpx'));
    const xml = await zip.file('Contents/section0.xml')!.async('string');
    zip.file('Contents/section0.xml', xml.replace(/horzRelTo="[A-Z]+"/g, 'horzRelTo="PAPER"'));
    const parsed = await Pipeline.open(await zip.generateAsync({ type: 'uint8array' }), 'hwpx').inspect();
    if (!parsed.ok) throw new Error(parsed.error);
    const imgs: ImgNode[] = [];
    const walk = (n: any) => { if (n.tag === 'img') imgs.push(n); for (const k of n.kids ?? []) walk(k); };
    walk(parsed.data);
    expect(imgs.map(i => i.layout?.horzRelTo)).toEqual(WRAPS.map(() => 'page'));
    expect(imgs.filter(i => i.layout?.wrap === 'front')).toHaveLength(2);
  });
});
