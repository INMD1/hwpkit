/**
 * BaseDecoder - 모든 디코더의 기본 클래스
 *
 * 공통 로직을 제공하여 각 포맷별 디코더가 포맷 특유의 로직만 구현하도록 합니다.
 */

import type { Decoder } from '../contract/decoder';
import type { Outcome } from '../contract/result';
import type { DocRoot } from '../model/doc-tree';
import { TextKit } from '../toolkit/TextKit';
import { ArchiveKit } from '../toolkit/ArchiveKit';

// ─── BaseDecoder ───────────────────────────────────

export abstract class BaseDecoder implements Decoder {
  readonly format: string = this.getFormat();

  /** 포맷 이름 반환 (하위 클래스에서 오버라이드) */
  protected abstract getFormat(): string;

  /** 데이터 디코딩 (하위 클래스에서 오버라이드) */
  abstract decode(data: Uint8Array): Promise<Outcome<DocRoot>>;

  // ─── 공통 유틸리티 메서드 ──────────────────────────

  /** 바이트를 UTF-8 문자열로 변환 */
  protected bytesToString(data: Uint8Array): string {
    return TextKit.decode(data);
  }

  /** 문자열을 UTF-8 바이트로 변환 */
  protected stringToBytes(s: string): Uint8Array {
    return TextKit.encode(s);
  }

  /** XML 이스케이프 */
  protected escapeXml(s: string): string {
    return TextKit.escapeXml(s);
  }

  /** XML 언이스케이프 */
  protected unescapeXml(s: string): string {
    return TextKit.unescapeXml(s);
  }

  /** base64 문자열을 Uint8Array 로 변환 */
  protected base64ToBytes(b64: string): Uint8Array {
    return TextKit.base64Decode(b64);
  }

  /** Uint8Array 를 base64 문자열로 변환 */
  protected bytesToBase64(data: Uint8Array): string {
    return TextKit.base64Encode(data);
  }

  /** ZIP 해제 */
  protected async unzip(data: Uint8Array): Promise<Map<string, Uint8Array>> {
    return ArchiveKit.unzip(data);
  }

  /** ZIP 압축 */
  protected async zip(entries: { name: string; data: Uint8Array }[]): Promise<Uint8Array> {
    return ArchiveKit.zip(entries);
  }

  /** inflate 해제 */
  protected async inflate(data: Uint8Array): Promise<Uint8Array> {
    return ArchiveKit.inflate(data);
  }

  /** deflate 압축 */
  protected async deflate(data: Uint8Array): Promise<Uint8Array> {
    return ArchiveKit.deflate(data);
  }

  /** 제어 문자 제거 */
  protected stripControl(s: string): string {
    return TextKit.stripControl(s);
  }

  /** 공백 정규화 */
  protected normalizeWhitespace(s: string): string {
    return TextKit.normalizeWhitespace(s);
  }

  /** XML 파싱 (DOMParser 사용) */
  protected parseXml(xmlString: string): Document {
    const parser = new DOMParser();
    return parser.parseFromString(xmlString, 'text/xml');
  }

  /** XML 요소에서 텍스트 내용 추출 */
  protected getTextContent(element: Element | null | undefined): string {
    if (!element) return '';
    return element.textContent ?? '';
  }

  /** XML 요소의 속성 값 추출 */
  protected getAttr(element: Element | null | undefined, name: string): string | null {
    return element?.getAttribute(name) ?? null;
  }

  /** XML 요소의 자식 요소 찾기 */
  protected getChild(element: Element | null | undefined, tagName: string): Element | null {
    if (!element) return null;
    return element.querySelector(`>${tagName}`) ?? null;
  }

  /** XML 요소의 모든 자식 요소 찾기 */
  protected getChildren(element: Element | null | undefined, tagName: string): Element[] {
    if (!element) return [];
    return Array.from(element.querySelectorAll(`>${tagName}`));
  }
}
