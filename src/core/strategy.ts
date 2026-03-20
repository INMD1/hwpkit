/**
 * Strategy 패턴
 *
 * ReadStrategy: 원본 바이너리 → IrDocumentNode
 * WriteStrategy: IrDocumentNode → 목표 포맷 데이터
 *
 * 각 포맷별 Reader/Writer는 이 인터페이스를 구현합니다.
 * Transformer 클래스는 런타임에 적절한 Strategy를 선택하여 주입합니다.
 */

import type { IrDocumentNode } from './ir';

// ─── Read Strategy ────────────────────────────────────────────────────────────

export interface ReadStrategy {
  /** 파일 바이너리를 받아 IR 트리를 반환합니다 */
  read(data: Uint8Array): Promise<IrDocumentNode>;
}

// ─── Write Strategy ───────────────────────────────────────────────────────────

export interface WriteStrategy<Output> {
  /** IR 트리를 받아 목표 포맷 데이터를 반환합니다 */
  write(document: IrDocumentNode): Promise<Output>;
}

// ─── Strategy Registry ────────────────────────────────────────────────────────

export type SupportedInputFormat = 'hwpx' | 'hwp' | 'docx' | 'md';
export type SupportedOutputFormat = 'hwpx' | 'docx' | 'md';

export type FormatPair = `${SupportedInputFormat}->${SupportedOutputFormat}`;

const SUPPORTED_PAIRS: FormatPair[] = [
  'hwpx->md',
  'hwpx->docx',
  'hwp->md',
  'hwp->docx',
  'docx->hwpx',
  'docx->md',
  'md->hwpx',
  'md->docx',
];

export function isSupportedPair(from: string, to: string): from is SupportedInputFormat {
  return SUPPORTED_PAIRS.includes(`${from}->${to}` as FormatPair);
}
