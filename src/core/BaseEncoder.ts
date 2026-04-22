/**
 * BaseEncoder - 모든 인코더의 기본 클래스
 *
 * 공통 로직을 제공하여 각 포맷별 인코더가 포맷 특유의 로직만 구현하도록 합니다.
 */

import type { Encoder } from '../contract/encoder';
import type { Outcome } from '../contract/result';
import type { DocRoot, AnyNode, ImgNode } from '../model/doc-tree';
import { TextKit } from '../toolkit/TextKit';
import { ArchiveKit } from '../toolkit/ArchiveKit';

// ─── 타입 정의 ───────────────────────────────────

export interface ImageData {
  node: ImgNode;
  path: string;
  data: Uint8Array;
}

export interface ZipEntry {
  name: string;
  data: Uint8Array;
}

// ─── BaseEncoder ───────────────────────────────

export abstract class BaseEncoder implements Encoder {
  readonly format: string = this.getFormat();

  /** 포맷 이름 반환 (하위 클래스에서 오버라이드) */
  protected abstract getFormat(): string;

  /** 문서 인코딩 (하위 클래스에서 오버라이드) */
  abstract encode(doc: DocRoot): Promise<Outcome<Uint8Array>>;

  // ─── 공통 유틸리티 메서드 ───────────────────────────

  /** 문서 내 모든 이미지 노드 수집 */
  protected collectImages(doc: DocRoot): ImgNode[] {
    const images: ImgNode[] = [];
    this.collectImagesRecursive(doc, images);
    return images;
  }

  /** 재귀적으로 이미지 수집 */
  private collectImagesRecursive(node: AnyNode, images: ImgNode[]): void {
    if (node.tag === 'img') {
      images.push(node);
      return;
    }

    const children = this.getChildren(node);
    for (const child of children) {
      this.collectImagesRecursive(child, images);
    }
  }

  /** 노드의 자식 노드 반환 */
  private getChildren(node: AnyNode): AnyNode[] {
    switch (node.tag) {
      case 'root':
        return node.kids;
      case 'sheet':
        return [...node.kids, ...(node.headers?.default ?? []), ...(node.footers?.default ?? [])];
      case 'para':
        return node.kids;
      case 'span':
        return node.kids;
      case 'link':
        return node.kids;
      case 'row':
        return node.kids;
      case 'cell':
        return node.kids;
      case 'grid':
        return node.kids;
      default:
        return [];
    }
  }

  /** base64 문자열을 Uint8Array 로 변환 */
  protected base64ToBytes(b64: string): Uint8Array {
    return TextKit.base64Decode(b64);
  }

  /** Uint8Array 를 base64 문자열로 변환 */
  protected bytesToBase64(data: Uint8Array): string {
    return TextKit.base64Encode(data);
  }

  /** XML 이스케이프 */
  protected escapeXml(s: string): string {
    return TextKit.escapeXml(s);
  }

  /** XML 언이스케이프 */
  protected unescapeXml(s: string): string {
    return TextKit.unescapeXml(s);
  }

  /** 문자열을 UTF-8 바이트로 변환 */
  protected stringToBytes(s: string): Uint8Array {
    return TextKit.encode(s);
  }

  /** 바이트를 UTF-8 문자열로 변환 */
  protected bytesToString(data: Uint8Array): string {
    return TextKit.decode(data);
  }

  /** ZIP 압축 */
  protected async zip(entries: ZipEntry[]): Promise<Uint8Array> {
    return ArchiveKit.zip(entries);
  }

  /** ZIP 해제 */
  protected async unzip(data: Uint8Array): Promise<Map<string, Uint8Array>> {
    return ArchiveKit.unzip(data);
  }

  /** deflate 압축 */
  protected async deflate(data: Uint8Array): Promise<Uint8Array> {
    return ArchiveKit.deflate(data);
  }

  /** inflate 해제 */
  protected async inflate(data: Uint8Array): Promise<Uint8Array> {
    return ArchiveKit.inflate(data);
  }

  /** 제어 문자 제거 */
  protected stripControl(s: string): string {
    return TextKit.stripControl(s);
  }

  /** 공백 정규화 */
  protected normalizeWhitespace(s: string): string {
    return TextKit.normalizeWhitespace(s);
  }
}
