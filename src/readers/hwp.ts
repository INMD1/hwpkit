/**
 * HwpReader - HWP 바이너리 → IrDocumentNode
 *
 * Strategy: ReadStrategy 구현체
 * HWP는 OLE Compound File Binary(CFB) 포맷입니다.
 * cfb 라이브러리로 구조를 추출하고, pako/fflate로 압축을 해제합니다.
 *
 * 데이터 흐름:
 *   Uint8Array (HWP binary)
 *     → CFB.read()              (OLE 스트림 분리)
 *     → decompressStream()      (zlib 압축 해제 - fflate 사용)
 *     → RecordScanner           (HWP 레코드 순회, Visitor 패턴 적용)
 *     → IrDocumentNode
 */

import * as CFB from 'cfb';
import { inflateSync } from 'fflate';
import type { ReadStrategy } from '../core/strategy';
import type {
  IrDocumentNode,
  IrSectionNode,
  IrBlockNode,
  IrParagraphNode,
  IrRunNode,
  RunStyle,
  ParagraphStyle,
  PageLayout,
  DocumentMeta,
} from '../core/ir';
import {
  makeDocument,
  makeSection,
  makeParagraph,
  makeRun,
  DEFAULT_PAGE_LAYOUT,
} from '../core/ir';
import { decodeBuffer } from '../services/encoding';

// ─── HWP 레코드 태그 상수 ────────────────────────────────────────────────────

const TAG = {
  DOCUMENT_PROPERTIES: 0x01,
  ID_MAPPINGS: 0x05,
  BIN_DATA: 0x06,
  FACE_NAME: 0x08,
  BORDER_FILL: 0x09,
  CHAR_SHAPE: 0x0A,
  TAB_DEF: 0x0B,
  NUMBERING: 0x0C,
  BULLET: 0x0D,
  PARA_SHAPE: 0x0E,
  STYLE: 0x0F,
  DOC_DATA: 0x11,
  MEMO_SHAPE: 0x1F,
  TRACK_CHANGE: 0x20,
  PARA_TEXT: 0x32,
  PARA_CHAR_SHAPE: 0x33,
  PARA_LINE_SEG: 0x34,
  PARA_RANGE_TAG: 0x35,
  CTRL_HEADER: 0x36,
  LIST_HEADER: 0x37,
  PAGE_DEF: 0x38,
  FOOTNOTE_SHAPE: 0x39,
  PAGE_BORDER_FILL: 0x3A,
  SHAPE_COMPONENT: 0x3C,
  TABLE: 0x3D,
} as const;

// ─── 레코드 구조 ──────────────────────────────────────────────────────────────

interface HwpRecord {
  tag: number;
  level: number;
  size: number;
  data: DataView;
}

/** HWP 레코드 스캐너 (Visitor 패턴의 순회 담당) */
class RecordScanner {
  private pos = 0;
  private view: DataView;

  constructor(private buffer: ArrayBuffer) {
    this.view = new DataView(buffer);
  }

  hasMore(): boolean {
    return this.pos < this.view.byteLength - 4;
  }

  next(): HwpRecord | null {
    if (!this.hasMore()) return null;
    const header = this.view.getUint32(this.pos, true);
    this.pos += 4;

    const tag = header & 0x3FF;
    const level = (header >> 10) & 0x3FF;
    let size = (header >> 20) & 0xFFF;

    if (size === 0xFFF) {
      size = this.view.getUint32(this.pos, true);
      this.pos += 4;
    }

    const data = new DataView(this.buffer, this.pos, Math.min(size, this.buffer.byteLength - this.pos));
    this.pos += size;

    return { tag, level, size, data };
  }

  /** 특정 태그의 레코드들만 수집 */
  collectTag(tag: number): HwpRecord[] {
    const saved = this.pos;
    this.pos = 0;
    const results: HwpRecord[] = [];
    while (this.hasMore()) {
      const rec = this.next();
      if (rec && rec.tag === tag) results.push(rec);
    }
    this.pos = saved;
    return results;
  }
}

// ─── CFB 스트림 처리 ──────────────────────────────────────────────────────────

function readCfbStream(cfb: CFB.CFB$Container, path: string): Uint8Array | null {
  try {
    const entry = CFB.find(cfb, path);
    if (!entry) return null;
    const content = entry.content;
    if (!content) return null;
    if (content instanceof Uint8Array) return content;
    if (Buffer.isBuffer(content)) return new Uint8Array(content.buffer as ArrayBuffer, content.byteOffset, content.byteLength);
    // cfb 라이브러리의 content 타입이 불명확하므로 unknown을 통해 변환
    return new Uint8Array(content as unknown as ArrayBuffer);
  } catch {
    return null;
  }
}

/** HWP DocInfo 스트림에서 압축 여부 확인 후 압축 해제 */
function decompressIfNeeded(data: Uint8Array, compressed: boolean): ArrayBuffer {
  const sliceBuffer = (d: Uint8Array): ArrayBuffer =>
    d.buffer.slice(d.byteOffset, d.byteOffset + d.byteLength) as ArrayBuffer;
  if (!compressed) return sliceBuffer(data);
  try {
    return sliceBuffer(inflateSync(data));
  } catch {
    return sliceBuffer(data);
  }
}

// ─── CharShape 파싱 ───────────────────────────────────────────────────────────

interface HwpCharShape {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  fontSizePt: number;
  colorHex?: string;
}

function parseCharShape(rec: HwpRecord): HwpCharShape {
  const v = rec.data;
  if (v.byteLength < 40) return { bold: false, italic: false, underline: false, strikethrough: false, fontSizePt: 10 };

  const attr = v.getUint32(8, true);
  const bold = Boolean(attr & (1 << 0));
  const italic = Boolean(attr & (1 << 1));
  const underline = Boolean(attr & (1 << 2));
  const strikethrough = Boolean(attr & (1 << 9));
  const sizeRaw = v.getUint32(0, true);
  const fontSizePt = sizeRaw / 100;

  const r = v.getUint8(28);
  const g = v.getUint8(29);
  const b = v.getUint8(30);
  const colorHex = r === 0 && g === 0 && b === 0
    ? undefined
    : `${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;

  return { bold, italic, underline, strikethrough, fontSizePt, colorHex };
}

// ─── 텍스트 추출 ──────────────────────────────────────────────────────────────

/** HWP BodyText 스트림에서 단락 텍스트와 서식 정보를 추출 */
interface ParsedParagraph {
  text: string;
  charShapeId: number;
}

function extractBodyText(scanner: RecordScanner, charShapes: HwpCharShape[]): ParsedParagraph[] {
  const paragraphs: ParsedParagraph[] = [];
  let currentText = '';
  let currentCharShapeId = 0;

  scanner['pos'] = 0; // reset

  while (scanner.hasMore()) {
    const rec = scanner.next();
    if (!rec) break;

    if (rec.tag === TAG.PARA_TEXT) {
      const bytes = new Uint8Array(rec.data.buffer, rec.data.byteOffset, rec.data.byteLength);
      try {
        const text = decodeBuffer(bytes, 'utf-16le');
        currentText += text.replace(/\x00/g, '').replace(/\r/g, '');
      } catch {
        // 인코딩 실패 시 무시
      }
    } else if (rec.tag === TAG.PARA_CHAR_SHAPE) {
      if (rec.data.byteLength >= 8) {
        currentCharShapeId = rec.data.getUint32(4, true);
      }
    } else if (rec.tag === TAG.CTRL_HEADER) {
      // 단락 경계 - 현재 단락 저장
      if (currentText.trim()) {
        paragraphs.push({ text: currentText, charShapeId: currentCharShapeId });
      }
      currentText = '';
      currentCharShapeId = 0;
    }
  }

  if (currentText.trim()) {
    paragraphs.push({ text: currentText, charShapeId: currentCharShapeId });
  }

  return paragraphs;
}

// ─── 페이지 레이아웃 ──────────────────────────────────────────────────────────

function parsePageDef(rec: HwpRecord): Partial<PageLayout> {
  if (rec.data.byteLength < 32) return {};
  const v = rec.data;
  const hw2pt = (x: number) => (x / 7200) * 72;
  return {
    widthPt: hw2pt(v.getUint32(0, true)),
    heightPt: hw2pt(v.getUint32(4, true)),
    marginTopPt: hw2pt(v.getUint32(8, true)),
    marginBottomPt: hw2pt(v.getUint32(12, true)),
    marginLeftPt: hw2pt(v.getUint32(16, true)),
    marginRightPt: hw2pt(v.getUint32(20, true)),
  };
}

// ─── HwpReader ────────────────────────────────────────────────────────────────

export class HwpReader implements ReadStrategy {
  async read(data: Uint8Array): Promise<IrDocumentNode> {
    const cfb = CFB.read(data, { type: 'buffer' });

    // FileHeader에서 압축 여부 확인
    const fileHeaderData = readCfbStream(cfb, 'FileHeader');
    const compressed = fileHeaderData
      ? Boolean(new DataView(fileHeaderData.buffer).getUint32(36, true) & 1)
      : true;

    // DocInfo 스트림 파싱 (CharShape, PageDef 등)
    const docInfoRaw = readCfbStream(cfb, 'DocInfo');
    const charShapes: HwpCharShape[] = [];
    let pageLayout: Partial<PageLayout> = {};

    if (docInfoRaw) {
      const docInfoBuf = decompressIfNeeded(docInfoRaw, compressed);
      const docScanner = new RecordScanner(docInfoBuf);
      while (docScanner.hasMore()) {
        const rec = docScanner.next();
        if (!rec) break;
        if (rec.tag === TAG.CHAR_SHAPE) {
          charShapes.push(parseCharShape(rec));
        } else if (rec.tag === TAG.PAGE_DEF) {
          pageLayout = parsePageDef(rec);
        }
      }
    }

    // 섹션별 BodyText 파싱
    const sections: IrSectionNode[] = [];
    let sectionIndex = 0;

    while (true) {
      const streamName = `BodyText/Section${sectionIndex}`;
      const rawSection = readCfbStream(cfb, streamName);
      if (!rawSection) break;

      const sectionBuf = decompressIfNeeded(rawSection, compressed);
      const bodyScanner = new RecordScanner(sectionBuf);
      const parsedParas = extractBodyText(bodyScanner, charShapes);

      const blocks: IrBlockNode[] = parsedParas.map(({ text, charShapeId }) => {
        const cs = charShapes[charShapeId];
        const runStyle: RunStyle = cs
          ? {
              bold: cs.bold || undefined,
              italic: cs.italic || undefined,
              underline: cs.underline || undefined,
              strikethrough: cs.strikethrough || undefined,
              fontSizePt: cs.fontSizePt > 0 ? cs.fontSizePt : undefined,
              colorHex: cs.colorHex,
            }
          : {};

        const run = makeRun(text, runStyle);
        return makeParagraph({}, [run]);
      });

      sections.push(makeSection({ ...DEFAULT_PAGE_LAYOUT, ...pageLayout }, blocks));
      sectionIndex++;
    }

    if (sections.length === 0) {
      sections.push(makeSection({ ...DEFAULT_PAGE_LAYOUT, ...pageLayout }, []));
    }

    return makeDocument({}, sections);
  }
}
