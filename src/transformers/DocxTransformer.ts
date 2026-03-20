/**
 * DocxTransformer - DOCX 출력 변환기
 *
 * 공개 API: 정적 팩토리 메서드 패턴을 사용합니다.
 *
 * 사용 예:
 *   const blob = await DocxTransformer.fromHwpx(file);
 *   const blob = await DocxTransformer.fromHwp(file);
 *   const blob = await DocxTransformer.fromMd(markdownString);
 */

import { HwpxReader } from '../readers/hwpx';
import { HwpReader } from '../readers/hwp';
import { DocxWriter } from '../writers/docx';
import { MarkdownReader } from '../readers/markdown';
import type { IrDocumentNode } from '../core/ir';

export class DocxTransformer {
  private constructor(
    private readonly ir: IrDocumentNode,
  ) {}

  /** HWPX 파일 → DOCX */
  static async fromHwpx(source: File | Blob | Uint8Array): Promise<Blob> {
    const data = await toUint8Array(source);
    const ir = await new HwpxReader().read(data);
    return new DocxTransformer(ir).emit();
  }

  /** HWP 바이너리 파일 → DOCX */
  static async fromHwp(source: File | Blob | Uint8Array): Promise<Blob> {
    const data = await toUint8Array(source);
    const ir = await new HwpReader().read(data);
    return new DocxTransformer(ir).emit();
  }

  /** Markdown 문자열 → DOCX */
  static async fromMd(markdown: string): Promise<Blob> {
    const ir = await new MarkdownReader().readText(markdown);
    return new DocxTransformer(ir).emit();
  }

  /** 이미 파싱된 IR에서 DOCX 생성 */
  static fromIR(ir: IrDocumentNode): DocxTransformer {
    return new DocxTransformer(ir);
  }

  /** DOCX Blob 출력 */
  async emit(): Promise<Blob> {
    return new DocxWriter().write(this.ir);
  }

  getIR(): IrDocumentNode {
    return this.ir;
  }
}

async function toUint8Array(source: File | Blob | Uint8Array): Promise<Uint8Array> {
  if (source instanceof Uint8Array) return source;
  const buf = await source.arrayBuffer();
  return new Uint8Array(buf);
}
