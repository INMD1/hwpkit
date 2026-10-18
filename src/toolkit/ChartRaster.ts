import pako from 'pako';
import { XmlKit } from './XmlKit';

// Charts are flattened to a PNG for formats without a native DrawingML chart
// model. Both the browser canvas and the portable renderer consume one scene.
type Color = string;
type Command = { kind: 'rect'; x: number; y: number; w: number; h: number; color: Color }
  | { kind: 'line'; x: number; y: number; x2: number; y2: number; color: Color }
  | { kind: 'text'; x: number; y: number; text: string; size: number; color: Color; align?: 'left' | 'center' | 'right' };
interface Series { name: string; categories: string[]; values: number[]; color: string }
const palette = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
const children = (node: any, tag: string): any[] => node?.[tag] ?? [];
const first = (node: any, tag: string): any => children(node, tag)[0];
const value = (node: any, tag: string): string | undefined => first(node, tag)?._attr?.val;
function cache(node: any): string[] {
  const ref = first(node, 'c:numRef') ?? first(node, 'c:strRef');
  const c = first(ref, 'c:numCache') ?? first(ref, 'c:strCache') ?? first(node, 'c:numLit') ?? first(node, 'c:strLit');
  const count = Math.min(1024, Math.max(0, Number(value(c, 'c:ptCount') ?? 0)));
  const out: string[] = Array(count).fill('');
  for (const point of children(c, 'c:pt').slice(0, 1024)) {
    const index = Number(point?._attr?.idx);
    if (Number.isInteger(index) && index >= 0 && index < 1024) out[index] = first(point, 'c:v')?._text ?? '';
  }
  return out;
}
function color(node: any, fallback: string): string {
  const raw = first(first(node, 'a:solidFill'), 'a:srgbClr')?._attr?.val;
  return /^[a-f0-9]{6}$/i.test(raw ?? '') ? raw : fallback;
}

export async function renderDocxChart(xml: string, wPt: number, hPt: number): Promise<Uint8Array> {
  const object: any = await XmlKit.parseStrict(xml);
  const chart = first(first(object, 'c:chartSpace'), 'c:chart');
  const plot = first(chart, 'c:plotArea');
  const bars = first(plot, 'c:barChart');
  const lines = first(plot, 'c:lineChart');
  const graph = bars ?? lines;
  const chartTypes = Object.keys(plot ?? {}).filter(k => /^c:.*Chart$/.test(k));
  if (!graph || chartTypes.length !== 1 || children(plot, bars ? 'c:barChart' : 'c:lineChart').length !== 1
    || (value(graph, 'c:grouping') ?? 'clustered') === 'percentStacked'
    || (lines && ['stacked', 'percentStacked'].includes(value(lines, 'c:grouping') ?? 'standard')))
    throw new Error('unsupported chart type (supported: 2D bar/column and line charts)');
  const series: Series[] = children(graph, 'c:ser').slice(0, 64).map((s, i) => ({
    name: cache(first(s, 'c:tx'))[0] ?? first(first(s, 'c:tx'), 'c:v')?._text ?? `Series ${i + 1}`,
    categories: cache(first(s, 'c:cat')),
    values: cache(first(s, 'c:val')).map(n => n.trim() ? Number(n) : NaN),
    color: color(first(s, 'c:spPr'), palette[i % palette.length]),
  }));
  if (!series.length || !series.some(s => s.values.some(Number.isFinite))) throw new Error('chart has no cached numeric data');
  const width = Math.max(72, Math.min(1200, Number.isFinite(wPt) ? wPt : 360));
  const height = Math.max(48, Math.min(1200, Number.isFinite(hPt) ? hPt : 216));
  const scene: Command[] = [];
  const rect = (x: number, y: number, w: number, h: number, c: Color) => scene.push({ kind: 'rect', x, y, w, h, color: c });
  const line = (x: number, y: number, x2: number, y2: number, c = 'B3B3B3') => scene.push({ kind: 'line', x, y, x2, y2, color: c });
  const text = (x: number, y: number, t: string, align: 'left' | 'center' | 'right' = 'left') =>
    scene.push({ kind: 'text', x, y, text: t.slice(0, 80), size: 7, color: '222222', align });
  const hasLegend = first(chart, 'c:legend') !== undefined;
  const legendWidth = hasLegend ? Math.min(width * .30, Math.max(55, ...series.map(s => s.name.length * 4 + 15))) : 0;
  const left = Math.min(28, width * .18), top = 10, bottom = height - 22, right = width - legendWidth - 7;
  const pw = Math.max(10, right - left), ph = Math.max(10, bottom - top);
  const count = Math.max(...series.map(s => Math.max(s.categories.length, s.values.length)));
  const stacked = value(graph, 'c:grouping') === 'stacked';
  const vals = series.flatMap(s => s.values).filter(Number.isFinite);
  if (stacked) for (let i = 0; i < count; i++) {
    vals.push(series.reduce((n, s) => n + Math.max(0, s.values[i] || 0), 0));
    vals.push(series.reduce((n, s) => n + Math.min(0, s.values[i] || 0), 0));
  }
  const rawMin = Math.min(0, ...vals), rawMax = Math.max(0, ...vals);
  const span = Math.max(1, rawMax - rawMin);
  const power = Math.pow(10, Math.floor(Math.log10(span / 5)));
  const tick = [1, 2, 5, 10].map(n => n * power).find(n => n >= span / 5) ?? power * 10;
  const axis = first(plot, 'c:valAx'), scaling = first(axis, 'c:scaling');
  const minimum = Number(value(scaling, 'c:min') ?? Math.floor(rawMin / tick) * tick);
  const maximum = Number(value(scaling, 'c:max') ?? Math.ceil((rawMax || 1) / tick) * tick);
  if (!Number.isFinite(minimum) || !Number.isFinite(maximum) || maximum <= minimum) throw new Error('invalid chart axis range');
  const horizontal = bars && value(bars, 'c:barDir') === 'bar';
  const py = (n: number) => bottom - (n - minimum) / (maximum - minimum) * ph;
  const px = (n: number) => left + (n - minimum) / (maximum - minimum) * pw;
  for (let n = minimum, i = 0; n <= maximum + tick * .001 && i < 50; n += tick, i++) {
    const label = Number(n.toPrecision(6)).toString();
    if (horizontal) { line(px(n), top, px(n), bottom); text(px(n), bottom + 5, label, 'center'); }
    else { line(left, py(n), right, py(n)); text(left - 5, py(n) - 3.5, label, 'right'); }
  }
  line(left, top, left, bottom); line(left, bottom, right, bottom);
  const gap = Math.min(5, Math.max(0, Number(value(graph, 'c:gapWidth') ?? 150) / 100));
  const groupSpan = (horizontal ? ph : pw) / Math.max(1, count);
  const band = groupSpan / (1 + gap / Math.max(1, series.length));
  for (let i = 0; i < count; i++) {
    const label = series.find(s => s.categories[i])?.categories[i] ?? String(i + 1);
    if (horizontal) text(left - 4, top + (i + .5) * groupSpan - 3.5, label, 'right');
    else text(left + (i + .5) * groupSpan, bottom + 5, label, 'center');
    let positive = 0, negative = 0;
    for (let j = 0; j < series.length; j++) {
      const v = series[j].values[i];
      if (!Number.isFinite(v)) continue;
      if (lines) {
        if (i > 0 && Number.isFinite(series[j].values[i - 1]))
          line(left + (i - .5) * groupSpan, py(series[j].values[i - 1]), left + (i + .5) * groupSpan, py(v), series[j].color);
        rect(left + (i + .5) * groupSpan - 1.5, py(v) - 1.5, 3, 3, series[j].color);
      } else {
        const start = stacked ? (v < 0 ? negative : positive) : 0;
        if (stacked) { if (v < 0) negative += v; else positive += v; }
        const size = band / (stacked ? 1 : series.length);
        const offset = (groupSpan - band) / 2 + (stacked ? 0 : j * size);
        if (horizontal) rect(Math.min(px(start), px(start + v)), top + i * groupSpan + offset, Math.abs(px(start + v) - px(start)), size, series[j].color);
        else rect(left + i * groupSpan + offset, Math.min(py(start), py(start + v)), size, Math.abs(py(start) - py(start + v)), series[j].color);
      }
    }
  }
  if (hasLegend) series.forEach((s, i) => {
    const y = Math.max(top, (height - series.length * 12) / 2) + i * 12;
    rect(right + 7, y + 1, 6, 6, s.color); text(right + 17, y, s.name);
  });
  return rasterize(scene, width, height);
}

async function rasterize(scene: Command[], width: number, height: number): Promise<Uint8Array> {
  const scale = 2, w = Math.ceil(width * scale), h = Math.ceil(height * scale);
  // Use real fonts where a canvas exists; PNG encoding and the scene geometry
  // are portable, so importing the decoder never requires DOM/Node bindings.
  const canvas = typeof OffscreenCanvas !== 'undefined' ? new OffscreenCanvas(w, h)
    : typeof document !== 'undefined' ? document.createElement('canvas') : undefined;
  if (canvas) {
    canvas.width = w; canvas.height = h;
    const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null;
    if (ctx) {
      ctx.scale(scale, scale); ctx.fillStyle = '#FFFFFF'; ctx.fillRect(0, 0, width, height);
      for (const c of scene) {
        ctx.fillStyle = ctx.strokeStyle = '#' + c.color;
        if (c.kind === 'rect') ctx.fillRect(c.x, c.y, c.w, c.h);
        else if (c.kind === 'line') { ctx.lineWidth = .5; ctx.beginPath(); ctx.moveTo(c.x, c.y); ctx.lineTo(c.x2, c.y2); ctx.stroke(); }
        else { ctx.font = `${c.size}px Arial, sans-serif`; ctx.textAlign = c.align ?? 'left'; ctx.textBaseline = 'top'; ctx.fillText(c.text, c.x, c.y); }
      }
      const blob = 'convertToBlob' in canvas ? await canvas.convertToBlob({ type: 'image/png' })
        : await new Promise<Blob>((resolve, reject) => canvas.toBlob(b => b ? resolve(b) : reject(Error('canvas PNG encoding failed')), 'image/png'));
      return new Uint8Array(await blob.arrayBuffer());
    }
  }
  if (scene.some(c => c.kind === 'text' && [...c.text].some(ch => !glyphs[ch] && !glyphs[ch.toUpperCase()])))
    throw new Error('chart labels require a browser canvas for characters outside the portable font');
  const pixels = new Uint8Array(w * h * 3).fill(255);
  function rect(x: number, y: number, rw: number, rh: number, color: string) {
    const rgb = [0, 2, 4].map(i => parseInt(color.slice(i, i + 2), 16));
    for (let yy = Math.max(0, Math.floor(y)); yy < Math.min(h, Math.ceil(y + rh)); yy++)
      for (let xx = Math.max(0, Math.floor(x)); xx < Math.min(w, Math.ceil(x + rw)); xx++) pixels.set(rgb, (yy * w + xx) * 3);
  }
  for (const c of scene) {
    if (c.kind === 'rect') rect(c.x * scale, c.y * scale, c.w * scale, c.h * scale, c.color);
    else if (c.kind === 'line') {
      const steps = Math.max(1, Math.ceil(Math.max(Math.abs(c.x2 - c.x), Math.abs(c.y2 - c.y)) * scale));
      for (let i = 0; i <= steps; i++) rect((c.x + (c.x2 - c.x) * i / steps) * scale, (c.y + (c.y2 - c.y) * i / steps) * scale, 1, 1, c.color);
    } else {
      const unit = c.size * scale / 8, tw = c.text.length * 6 * unit;
      const x = c.x * scale - (c.align === 'center' ? tw / 2 : c.align === 'right' ? tw : 0);
      [...c.text].forEach((ch, i) => {
        const rows = glyphs[ch] ?? glyphs[ch.toUpperCase()] ?? glyphs['?'];
        rows.forEach((bits, y) => { for (let col = 0; col < 5; col++) if (bits & (1 << (4 - col))) rect(x + (i * 6 + col) * unit, c.y * scale + y * unit, unit, unit, c.color); });
      });
    }
  }
  return png(pixels, w, h);
}

const glyphs: Record<string, number[]> = Object.fromEntries(Object.entries({
  ' ': '00000000000000', '?': '0e110102040004', '.': '00000000000c0c', '-': '0000001f000000',
  '0': '0e11131519110e', '1': '040c040404040e', '2': '0e11010204081f', '3': '1e01010e01011e',
  '4': '02060a121f0202', '5': '1f10101e01011e', '6': '0e10101e11110e', '7': '1f010204080808',
  '8': '0e11110e11110e', '9': '0e11110f01010e',
  A: '0e11111f111111', B: '1e11111e11111e', C: '0e11101010110e', D: '1e11111111111e',
  E: '1f10101e10101f', F: '1f10101e101010', G: '0e11101711110f', H: '1111111f111111',
  I: '0e04040404040e', J: '0702020212120c', K: '11121418141211', L: '1010101010101f',
  M: '111b1515111111', N: '11191513111111', O: '0e11111111110e', P: '1e11111e101010',
  Q: '0e11111115120d', R: '1e11111e141211', S: '0f10100e01011e', T: '1f040404040404',
  U: '1111111111110e', V: '11111111110a04', W: '1111111515150a', X: '11110a040a1111',
  Y: '11110a04040404', Z: '1f01020408101f',
  a: '00000e010f110f', b: '10101e1111111e', c: '00000e1010110e', d: '01010f1111110f',
  e: '00000e111f100e', f: '0609081c080808', g: '00000f110f010e', h: '10101e11111111',
  i: '04000c0404040e', j: '0200060202120c', k: '101011121c1211', l: '0c04040404040e',
  m: '00001a15151515', n: '00001e11111111', o: '00000e1111110e', p: '00001e111e1010',
  q: '00000f110f0101', r: '00001619101010', s: '00000f100e011e', t: '08081c08080906',
  u: '0000111111130d', v: '00001111110a04', w: '0000111115150a', x: '0000110a040a11',
  y: '000011110f010e', z: '00001f0204081f',
}).map(([key, data]) => [key, data.match(/../g)!.map(n => parseInt(n, 16))]));

function png(rgb: Uint8Array, width: number, height: number): Uint8Array {
  function chunk(name: string, data: Uint8Array): Uint8Array {
    const out = new Uint8Array(data.length + 12), view = new DataView(out.buffer);
    view.setUint32(0, data.length); [...name].forEach((c, i) => out[4 + i] = c.charCodeAt(0)); out.set(data, 8);
    let crc = 0xffffffff;
    for (const b of out.subarray(4, out.length - 4)) { crc ^= b; for (let i = 0; i < 8; i++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0); }
    view.setUint32(out.length - 4, (crc ^ 0xffffffff) >>> 0); return out;
  }
  const ihdr = new Uint8Array(13), view = new DataView(ihdr.buffer);
  view.setUint32(0, width); view.setUint32(4, height); ihdr[8] = 8; ihdr[9] = 2;
  const scanlines = new Uint8Array((width * 3 + 1) * height);
  for (let y = 0; y < height; y++) scanlines.set(rgb.subarray(y * width * 3, (y + 1) * width * 3), y * (width * 3 + 1) + 1);
  const chunks = [new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', pako.deflate(scanlines)), chunk('IEND', new Uint8Array())];
  const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0)); let offset = 0;
  for (const c of chunks) { out.set(c, offset); offset += c.length; } return out;
}
