import type { DocRoot } from '../model/doc-tree';
import type { Outcome } from './result';

export interface EncoderOptions {
  [key: string]: any;  // 포맷별 옵션 확장 가능
}

export interface Encoder {
  readonly format: string;
  readonly aliases?: string[];
  encode(doc: DocRoot, options?: EncoderOptions): Promise<Outcome<Uint8Array>>;
}
