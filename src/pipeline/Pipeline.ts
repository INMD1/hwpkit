import JSZip from 'jszip';
import { BinaryKit } from '../toolkit/BinaryKit';
import type { DocRoot } from '../model/doc-tree';
import type { Outcome } from '../contract/result';
import type { EncoderOptions } from '../contract/encoder';
import { succeed, fail } from '../contract/result';
import { registry } from './registry';

// Side-effect imports: auto-register all decoders and encoders
import '../decoders/hwpx/HwpxDecoder';
import '../decoders/hwp/HwpScanner';
import '../decoders/docx/DocxDecoder';
import '../decoders/doc/DocDecoder';
import '../decoders/md/MdDecoder';
import '../decoders/html/HtmlDecoder';
import '../encoders/hwpx/HwpxEncoder';
import '../encoders/docx/DocxEncoder';
import '../encoders/md/MdEncoder';
import '../encoders/html/HtmlEncoder';
import '../encoders/hwp/HwpEncoder';

export class Pipeline {
  private constructor(private raw: Uint8Array, private srcFmt?: string) {}

  /** 파일을 열고 포맷을 자동 감지하거나 명시 */
  static open(input: Uint8Array | string, fmt?: string): Pipeline {
    if (typeof input === 'string') {
      return new Pipeline(new TextEncoder().encode(input), normalizeFormat(fmt) ?? 'md');
    }
    return new Pipeline(input, normalizeFormat(fmt));
  }

  /** File/Blob 비동기 입력 */
  static async openAsync(input: File | Blob | Uint8Array | string, fmt?: string): Promise<Pipeline> {
    if (input instanceof Uint8Array || typeof input === 'string') {
      return Pipeline.open(input, fmt);
    }
    const buf = await input.arrayBuffer();
    const data = new Uint8Array(buf);
    const extension = typeof File !== 'undefined' && input instanceof File ? getExt(input.name) : undefined;
    // Office documents are often saved with a stale extension. Inspect their
    // container; keep explicit caller choices and text-file extension hints.
    const hint = extension === 'txt' ? 'md' : extension === 'htm' ? 'html' : extension;
    const detectedFmt = normalizeFormat(fmt) ?? (hint && ['hwp', 'hwpx', 'docx', 'doc'].includes(hint) ? undefined : hint);
    return new Pipeline(data, detectedFmt);
  }

  /** 목표 포맷으로 변환 */
  async to(targetFmt: string, options?: EncoderOptions): Promise<Outcome<Uint8Array>> {
    const srcFmt = this.srcFmt ?? await detectFormat(this.raw);
    const decoder = registry.getDecoder(srcFmt);
    const encoder = registry.getEncoder(normalizeFormat(targetFmt) ?? targetFmt);

    if (!decoder) return fail(`지원하지 않는 입력 포맷: ${srcFmt}`);
    if (!encoder) return fail(`지원하지 않는 출력 포맷: ${targetFmt}`);

    const docResult = await decoder.decode(this.raw);
    if (!docResult.ok) return docResult;

    const encResult = await encoder.encode(docResult.data, options);
    if (!encResult.ok) return { ...encResult, warns: [...docResult.warns, ...encResult.warns] };

    return { ...encResult, warns: [...docResult.warns, ...encResult.warns] };
  }

  /** DocRoot만 추출 (인코딩 없이) */
  async inspect(): Promise<Outcome<DocRoot>> {
    const srcFmt = this.srcFmt ?? await detectFormat(this.raw);
    const decoder = registry.getDecoder(srcFmt);
    if (!decoder) return fail(`디코더 없음: ${srcFmt}`);
    return decoder.decode(this.raw);
  }
}

async function detectFormat(data: Uint8Array): Promise<string> {
  if (BinaryKit.isOle2(data)) {
    try {
      const streams = BinaryKit.parseCfb(data);
      const header = streams.get('FileHeader');
      if (header && new TextDecoder().decode(header.subarray(0, 17)) === 'HWP Document File') return 'hwp';
      if (streams.has('WordDocument')) return 'doc';
    } catch { /* Malformed or unsupported OLE: do not treat it as Markdown. */ }
    return 'ole';
  }
  if (data[0] === 0x50 && data[1] === 0x4B) {
    try {
      // Package entry names remain readable regardless of compression/order/size.
      const zip = await JSZip.loadAsync(data);
      if (zip.file('word/document.xml')) return 'docx';
      if (zip.file('Contents/content.hpf') || zip.file('Contents/section0.xml')) return 'hwpx';
    } catch { /* Invalid package is an unsupported input. */ }
    return 'zip';
  }
  const prefix = new TextDecoder().decode(data.subarray(0, 4096)).trimStart();
  if (prefix.startsWith('%PDF-')) return 'pdf';
  if (prefix.startsWith('{\\rtf')) return 'rtf';
  if (/^(?:<!doctype\s+html\b|<html\b)/i.test(prefix)) return 'html';
  return 'md';
}

function normalizeFormat(fmt?: string): string | undefined {
  return fmt?.trim().toLowerCase().replace(/^\./, '') || undefined;
}

function getExt(name: string): string | undefined {
  const parts = name.split('.');
  return parts.length > 1 ? parts[parts.length - 1].toLowerCase() : undefined;
}
