import type { DocRoot } from '../model/doc-tree';
import type { Outcome } from './result';

export interface Decoder {
  readonly format: string;
  readonly aliases?: string[];
  decode(data: Uint8Array): Promise<Outcome<DocRoot>>;
}
