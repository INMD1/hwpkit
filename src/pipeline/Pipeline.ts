import type { DocRoot } from '../model/doc-tree';
import type { Outcome } from '../contract/result';
import { succeed, fail } from '../contract/result';
import { registry } from './registry';

// Side-effect imports: auto-register all decoders and encoders
import '../decoders/hwpx/HwpxDecoder';
import '../decoders/hwp/HwpScanner';
import '../decoders/docx/DocxDecoder';
import '../decoders/md/MdDecoder';
import '../decoders/html/HtmlDecoder';
import '../encoders/hwpx/HwpxEncoder';
import '../encoders/docx/DocxEncoder';
import '../encoders/md/MdEncoder';
import '../encoders/html/HtmlEncoder';
import '../encoders/hwp/HwpEncoder';

export class Pipeline {
  private constructor(private raw: Uint8Array, private srcFmt: string) {}

  /** 파일을 열고 포맷을 자동 감지하거나 명시 */
  static open(input: Uint8Array | string, fmt?: string): Pipeline {
    if (typeof input === 'string') {
      return new Pipeline(new TextEncoder().encode(input), fmt ?? 'md');
    }
    return new Pipeline(input, fmt ?? detectFormat(input));
  }

  /** File/Blob 비동기 입력 */
  static async openAsync(input: File | Blob | Uint8Array | string, fmt?: string): Promise<Pipeline> {
    if (input instanceof Uint8Array || typeof input === 'string') {
      return Pipeline.open(input, fmt);
    }
    const buf = await input.arrayBuffer();
    const data = new Uint8Array(buf);
    const detectedFmt = fmt ?? (input instanceof File ? getExt(input.name) : undefined) ?? detectFormat(data);
    return new Pipeline(data, detectedFmt);
  }

  /** 목표 포맷으로 변환 */
  async to(targetFmt: string): Promise<Outcome<Uint8Array>> {
    const decoder = registry.getDecoder(this.srcFmt);
    const encoder = registry.getEncoder(targetFmt);

    if (!decoder) return fail(`지원하지 않는 입력 포맷: ${this.srcFmt}`);
    if (!encoder) return fail(`지원하지 않는 출력 포맷: ${targetFmt}`);

    const docResult = await decoder.decode(this.raw);
    if (!docResult.ok) return docResult;

    const encResult = await encoder.encode(docResult.data);
    if (!encResult.ok) return { ...encResult, warns: [...docResult.warns, ...encResult.warns] };

    return { ...encResult, warns: [...docResult.warns, ...encResult.warns] };
  }

  /** DocRoot만 추출 (인코딩 없이) */
  async inspect(): Promise<Outcome<DocRoot>> {
    const decoder = registry.getDecoder(this.srcFmt);
    if (!decoder) return fail(`디코더 없음: ${this.srcFmt}`);
    return decoder.decode(this.raw);
  }
}

function detectFormat(data: Uint8Array): string {
  // HWP 파일 (OLE Compound Document)
  if (data[0] === 0xD0 && data[1] === 0xCF && data[2] === 0x11 && data[3] === 0xE0) return 'hwp';

  // ZIP 기반 파일 (DOCX, HWPX)
  if (data[0] === 0x50 && data[1] === 0x4B) {
    // DOCX 는 [Content_Types].xml 에 application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml 이 있음
    // HWPX 는 application/ha-xml-core-document
    const str = new TextDecoder('utf-8', { fatal: false }).decode(data.slice(0, 4096));
    if (str.includes('wordprocessingml')) return 'docx';
    if (str.includes('ha-xml')) return 'hwpx';
    if (str.includes('hwpml/')) return 'hwpx';
    if (str.includes('word/')) return 'docx';
    return 'hwpx'; // 기본값
  }

  return 'md';
}

function getExt(name: string): string | undefined {
  const parts = name.split('.');
  return parts.length > 1 ? parts[parts.length - 1].toLowerCase() : undefined;
}
