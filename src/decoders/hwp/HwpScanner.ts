import type { Decoder } from '../../contract/decoder';
import type { DocRoot } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import { succeed, fail } from '../../contract/result';
import { buildRoot, buildSheet, buildPara, buildSpan } from '../../model/builders';
import { ShieldedParser } from '../../safety/ShieldedParser';
import { BinaryKit } from '../../toolkit/BinaryKit';
import { registry } from '../../pipeline/registry';
import { A4 } from '../../model/doc-props';
import pako from 'pako';

const HWP_TAG_PARA_TEXT = 67;

export class HwpScanner implements Decoder {
  readonly format = 'hwp';

  async decode(data: Uint8Array): Promise<Outcome<DocRoot>> {
    const shield = new ShieldedParser();
    const warns: string[] = [];

    try {
      if (!BinaryKit.isOle2(data)) return fail('HWP: Invalid OLE2 signature');

      const streams = BinaryKit.parseCfb(data);
      const sectionStream = streams.get('BodyText/Section0')
        ?? streams.get('Section0')
        ?? findBodySection(streams);

      if (!sectionStream) return fail('HWP: BodyText section not found');

      // HWP body text is typically zlib-compressed
      let sectionData: Uint8Array;
      try {
        sectionData = pako.inflateRaw(sectionStream);
      } catch {
        sectionData = sectionStream;
      }

      const paragraphs = extractParagraphs(sectionData, shield);
      warns.push(...shield.flush());

      const paras = paragraphs.map(text => buildPara([buildSpan(text)]));
      const sheet = buildSheet(paras.length > 0 ? paras : [buildPara([buildSpan('')])], A4);
      return succeed(buildRoot({}, [sheet]), warns);
    } catch (e: any) {
      warns.push(...shield.flush());
      return fail(`HWP decode error: ${e?.message ?? String(e)}`, warns);
    }
  }
}

function findBodySection(streams: Map<string, Uint8Array>): Uint8Array | undefined {
  for (const [key, val] of streams) {
    if (key.includes('Section') && !key.includes('Header') && !key.includes('Info')) return val;
  }
  return undefined;
}

function extractParagraphs(data: Uint8Array, shield: ShieldedParser): string[] {
  const texts: string[] = [];
  let offset = 0;

  while (offset + 4 <= data.length) {
    const header = BinaryKit.readU32LE(data, offset);
    const tag = header & 0x3FF;
    let size = (header >> 20) & 0xFFF;
    offset += 4;

    if (size === 0xFFF) {
      if (offset + 4 > data.length) break;
      size = BinaryKit.readU32LE(data, offset);
      offset += 4;
    }

    if (tag === HWP_TAG_PARA_TEXT) {
      const chunk = data.subarray(offset, offset + size);
      const text = shield.guard(() => decodeParaText(chunk), '', `hwp:paraText@${offset}`);
      if (text.trim()) texts.push(text);
    }

    offset += size;
  }

  return texts;
}

function decodeParaText(data: Uint8Array): string {
  const chars: string[] = [];
  let i = 0;
  while (i + 1 < data.length) {
    const code = data[i] | (data[i + 1] << 8);
    i += 2;

    if (code === 0x000D) break;
    if (code >= 1 && code <= 31) {
      if (code === 9) chars.push('\t');
      else if (code === 10 || code === 13) chars.push('\n');
    } else if (code > 0) {
      chars.push(String.fromCharCode(code));
    }
  }
  return chars.join('');
}

registry.registerDecoder(new HwpScanner());
