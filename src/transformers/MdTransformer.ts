/**
 * MdTransformer - Markdown 출력 변환기
 *
 * 공개 API: 정적 팩토리 메서드 패턴을 사용합니다.
 * 내부적으로 적절한 ReadStrategy + MdWriter 를 조합합니다.
 *
 * 사용 예:
 *   const md = await MdTransformer.fromHwpx(file);
 *   const md = await MdTransformer.fromHwp(file);
 *   const md = await MdTransformer.fromDocx(file);
 */

import { HwpxReader } from '../readers/hwpx';
import { HwpReader } from '../readers/hwp';
import { DocxReader } from '../readers/docx';
import { MdWriter } from '../writers/md';
import type { IrDocumentNode } from '../core/ir';

export class MdTransformer {
  private constructor(
    private readonly ir: IrDocumentNode,
  ) {}

  /** HWPX 파일 → Markdown */
  static async fromHwpx(source: File | Blob | Uint8Array): Promise<string> {
    const data = await toUint8Array(source);
    const ir = await new HwpxReader().read(data);
    return new MdTransformer(ir).emit();
  }

  /** HWP 바이너리 파일 → Markdown */
  static async fromHwp(source: File | Blob | Uint8Array): Promise<string> {
    const data = await toUint8Array(source);
    const ir = await new HwpReader().read(data);
    return new MdTransformer(ir).emit();
  }

  /** DOCX 파일 → Markdown */
  static async fromDocx(source: File | Blob | Uint8Array): Promise<string> {
    const data = await toUint8Array(source);
    const ir = await new DocxReader().read(data);
    return new MdTransformer(ir).emit();
  }

  /** 이미 파싱된 IR에서 Markdown 생성 */
  static fromIR(ir: IrDocumentNode): MdTransformer {
    return new MdTransformer(ir);
  }

  /** Markdown 문자열 출력 */
  async emit(): Promise<string> {
    return new MdWriter().write(this.ir);
  }

  /** IR 접근 (다른 Transformer로 재사용 가능) */
  getIR(): IrDocumentNode {
    return this.ir;
  }
}

async function toUint8Array(source: File | Blob | Uint8Array): Promise<Uint8Array> {
  if (source instanceof Uint8Array) return source;
  const buf = await source.arrayBuffer();
  return new Uint8Array(buf);
}
