/**
 * HwpxTransformer - HWPX 출력 변환기
 *
 * 공개 API: 정적 팩토리 메서드 패턴을 사용합니다.
 *
 * 사용 예:
 *   const blob = await HwpxTransformer.fromDocx(file);
 *   const blob = await HwpxTransformer.fromMd(markdownString);
 */

import { DocxReader } from '../readers/docx';
import { HwpxWriter } from '../writers/hwpx';
import type { IrDocumentNode } from '../core/ir';
import { makeDocument, makeSection, makeParagraph, makeRun } from '../core/ir';

// Markdown → IR 변환은 간단한 라인-기반 파서로 처리
import { MarkdownReader } from '../readers/markdown';

export class HwpxTransformer {
  private constructor(
    private readonly ir: IrDocumentNode,
  ) {}

  /** DOCX 파일 → HWPX */
  static async fromDocx(source: File | Blob | Uint8Array): Promise<Blob> {
    const data = await toUint8Array(source);
    const ir = await new DocxReader().read(data);
    return new HwpxTransformer(ir).emit();
  }

  /** Markdown 문자열 → HWPX */
  static async fromMd(markdown: string): Promise<Blob> {
    const ir = await new MarkdownReader().readText(markdown);
    return new HwpxTransformer(ir).emit();
  }

  /** 이미 파싱된 IR에서 HWPX 생성 */
  static fromIR(ir: IrDocumentNode): HwpxTransformer {
    return new HwpxTransformer(ir);
  }

  /** HWPX Blob 출력 */
  async emit(): Promise<Blob> {
    return new HwpxWriter().write(this.ir);
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
