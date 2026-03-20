import type { DocRoot } from '../model/doc-tree';
import type { Outcome } from './result';

export interface Decoder {
  readonly format: string;
  decode(data: Uint8Array): Promise<Outcome<DocRoot>>;
}
