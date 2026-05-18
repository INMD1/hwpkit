import type { Decoder } from '../contract/decoder';
import type { Encoder } from '../contract/encoder';

class FormatRegistry {
  private decoders = new Map<string, Decoder>();
  private encoders = new Map<string, Encoder>();

  registerDecoder(d: Decoder): void {
    this.decoders.set(d.format, d);
    if (d.aliases) {
      for (const alias of d.aliases) this.decoders.set(alias, d);
    }
  }

  registerEncoder(e: Encoder): void {
    this.encoders.set(e.format, e);
    if (e.aliases) {
      for (const alias of e.aliases) this.encoders.set(alias, e);
    }
  }

  getDecoder(fmt: string): Decoder | undefined { return this.decoders.get(fmt); }
  getEncoder(fmt: string): Encoder | undefined { return this.encoders.get(fmt); }

  supportedInputs(): string[]  { return [...this.decoders.keys()]; }
  supportedOutputs(): string[] { return [...this.encoders.keys()]; }
}

export const registry = new FormatRegistry();
