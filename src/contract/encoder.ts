import type { DocRoot } from '../model/doc-tree';
import type { Outcome } from './result';

export interface Encoder {
  readonly format: string;
  readonly aliases?: string[];
  encode(doc: DocRoot): Promise<Outcome<Uint8Array>>;
}
