import { BaseDecoder } from '../../core/BaseDecoder';
import { BinaryKit } from '../../toolkit/BinaryKit';
import { buildRoot, buildSheet, buildPara, buildSpan, buildPb } from '../../model/builders';
import type { DocRoot, ParaNode } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import { succeed, fail } from '../../contract/result';
import { DocxDecoder } from '../docx/DocxDecoder';
import { registry } from '../../pipeline/registry';

/** Optional local/server converter for preserving DOC formatting and objects. */
export type DocToDocx = (data: Uint8Array) => Promise<Uint8Array>;
let docConverter: DocToDocx | undefined;

export function configureDocConverter(converter?: DocToDocx): void {
  docConverter = converter;
}

/** Word 97–2003 binary DOC. Never interpret arbitrary OLE bytes as text. */
export class DocDecoder extends BaseDecoder {
  protected getFormat(): string { return 'doc'; }
  protected getAliases(): string[] { return ['application/msword']; }

  async decode(data: Uint8Array): Promise<Outcome<DocRoot>> {
    try {
      const streams = BinaryKit.parseCfb(data);
      const word = streams.get('WordDocument');
      if (!word || word.length < 32) throw new Error('WordDocument 스트림이 없거나 손상되었습니다.');
      const view = new DataView(word.buffer, word.byteOffset, word.byteLength);
      if (view.getUint16(0, true) !== 0xa5ec) throw new Error('Word DOC 서명이 아닙니다.');
      const flags = view.getUint16(10, true);
      if (flags & 0x8100) throw new Error('암호화된 DOC는 지원하지 않습니다. 암호를 해제한 뒤 다시 변환하세요.');

      if (docConverter) {
        const docx = await docConverter(data);
        return new DocxDecoder().decode(docx);
      }

      const version = view.getUint16(2, true);
      if (version < 0x00c1) throw new Error('Word 97 이전 DOC는 DOCX 변환기를 설정해야 합니다.');
      const table = streams.get(flags & 0x0200 ? '1Table' : '0Table');
      if (!table) throw new Error('DOC Table 스트림이 없습니다.');

      // MS-DOC §2.5 Fib: variable-sized rgw/rglw precede FibRgFcLcb97.
      let offset = 32;
      const csw = view.getUint16(offset, true);
      offset += 2 + csw * 2;
      const cslw = view.getUint16(offset, true);
      offset += 2;
      if (cslw < 4) throw new Error('잘못된 DOC FibRgLw입니다.');
      const bodyLength = view.getUint32(offset + 3 * 4, true); // ccpText
      offset += cslw * 4;
      const pairCount = view.getUint16(offset, true);
      offset += 2;
      if (pairCount <= 33) throw new Error('DOC CLX 참조가 없습니다.');
      const clxOffset = view.getUint32(offset + 33 * 8, true);
      const clxLength = view.getUint32(offset + 33 * 8 + 4, true);
      if (clxLength === 0 || clxOffset + clxLength > table.length) throw new Error('DOC CLX 범위가 잘못되었습니다.');
      const clx = table.subarray(clxOffset, clxOffset + clxLength);
      const cv = new DataView(clx.buffer, clx.byteOffset, clx.byteLength);
      let pos = 0;
      // CLX = zero or more Prc records, followed by one Pcdt.
      while (pos < clx.length && clx[pos] === 1) pos += 3 + cv.getUint16(pos + 1, true);
      if (clx[pos] !== 2) throw new Error('DOC piece table이 없습니다.');
      const size = cv.getUint32(pos + 1, true);
      pos += 5;
      if (size < 4 || (size - 4) % 12 !== 0 || pos + size > clx.length) throw new Error('손상된 DOC piece table입니다.');
      const count = (size - 4) / 12;
      const pcdOffset = pos + (count + 1) * 4;
      const chunks: string[] = [];
      let covered = 0;
      for (let i = 0; i < count; i++) {
        const start = cv.getUint32(pos + i * 4, true);
        const end = cv.getUint32(pos + (i + 1) * 4, true);
        if ((i === 0 && start !== 0) || end < start) throw new Error('DOC 문자 위치가 잘못되었습니다.');
        if (start >= bodyLength) break; // Exclude headers, comments and footnotes.
        const length = Math.min(end, bodyLength) - start;
        const fc = cv.getUint32(pcdOffset + i * 8 + 2, true);
        const compressed = (fc & 0x40000000) !== 0;
        const byteOffset = (fc & 0x3fffffff) / (compressed ? 2 : 1);
        const byteLength = length * (compressed ? 1 : 2);
        if (!Number.isInteger(byteOffset) || byteOffset + byteLength > word.length) throw new Error('DOC 텍스트 범위가 잘못되었습니다.');
        const bytes = word.subarray(byteOffset, byteOffset + byteLength);
        chunks.push(new TextDecoder(compressed ? 'windows-1252' : 'utf-16le').decode(bytes));
        covered += length;
      }
      if (covered !== bodyLength) throw new Error('DOC 본문이 잘렸습니다.');
      const text = fieldResults(chunks.join(''));
      const paragraphs: ParaNode[] = [];
      // Cell/row markers preserve cell text in reading order in basic mode.
      for (const line of text.replace(/\x07/g, '\r').split('\r')) {
        const pages = line.split('\x0c');
        const para = buildPara([]);
        pages.forEach((page, i) => {
          if (i) { const span = buildSpan(''); span.kids = [buildPb()]; para.kids.push(span); }
          para.kids.push(buildSpan(page.replace(/\x0b/g, '\n').replace(/[\x00-\x08\x0e-\x1f]/g, '')));
        });
        paragraphs.push(para);
      }
      // The final CR terminates the last paragraph; it is not another paragraph.
      if (text.endsWith('\r') && paragraphs.length > 1) paragraphs.pop();
      return succeed(buildRoot({}, [buildSheet(paragraphs)]), [
        'DOC 기본 읽기 모드: 본문 텍스트·문단·줄바꿈을 변환합니다. 서식, 표 구조, 그림, 머리말/꼬리말은 보존하지 않습니다. 전체 변환은 configureDocConverter 또는 hwpkit-dev/node의 createLibreOfficeDocConverter를 사용하세요.',
      ]);
    } catch (e) {
      return fail(`DOC 디코딩 오류: ${e instanceof Error ? e.message : String(e)}`);
    }
  }
}

/** Keep displayed field results, excluding instructions (including nested fields). */
function fieldResults(text: string): string {
  const fields: boolean[] = [];
  let result = '';
  for (const ch of text) {
    if (ch === '\x13') fields.push(false);
    else if (ch === '\x14' && fields.length) fields[fields.length - 1] = true;
    else if (ch === '\x15' && fields.length) fields.pop();
    else if (fields.every(Boolean)) result += ch;
  }
  return result;
}

registry.registerDecoder(new DocDecoder());
